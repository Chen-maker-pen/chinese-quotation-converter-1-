import express from "express";
import path from "path";
import dotenv from "dotenv";
import ExcelJS from "exceljs";
import { GoogleGenAI, Type } from "@google/genai";

dotenv.config();

// Initialize the Express App
const app = express();
const PORT = 3000;

// Set maximum body size for handling large Base64 Excel attachments and images
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Initialize Gemini Client
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

// Robust Gemini API call with exponential backoff and fallback model support
async function generateContentWithRetryAndFallback(params: any): Promise<any> {
  const config = { ...params };
  const modelsToTry = [config.model, "gemini-3.1-flash-lite"];
  
  // Clean up duplicates
  const uniqueModels = Array.from(new Set(modelsToTry.filter(Boolean)));
  
  for (const currentModel of uniqueModels) {
    config.model = currentModel;
    let delay = 1000;
    const maxRetries = 3;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        console.log(`Executing Gemini request using model: ${currentModel} (Attempt ${attempt}/${maxRetries})...`);
        return await ai.models.generateContent(config);
      } catch (err: any) {
        const errMsg = err?.message || String(err);
        const isRateLimit = errMsg.includes("429") || errMsg.includes("RESOURCE_EXHAUSTED") || errMsg.includes("quota");
        const isUnavailable = errMsg.includes("503") || errMsg.includes("UNAVAILABLE") || errMsg.includes("high demand") || errMsg.includes("temporary");
        
        console.warn(`Gemini API [${currentModel}] attempt ${attempt} failed with error:`, errMsg);
        
        // Fast-pivot: If the model is temporary overloaded (503), allow only 2 attempts on it before falling back to next model
        const actualAttemptsAllowed = isUnavailable ? 2 : maxRetries;
        
        if (attempt < actualAttemptsAllowed) {
          let waitTime = delay;
          
          // Parse retry duration from Google API response if available
          const retryMatch = errMsg.match(/Please retry in ([\d\.]+)s/i);
          if (retryMatch) {
            const waitSeconds = parseFloat(retryMatch[1]);
            waitTime = Math.ceil(waitSeconds * 1000) + 1500; // adding 1.5s safety margin
            console.log(`Rate limit detected. Custom sleep of ${waitTime}ms requested by Google API.`);
          } else if (isRateLimit) {
            waitTime = Math.max(delay, 5000);
          } else if (isUnavailable) {
            waitTime = 1000;
          }
          
          console.log(`Waiting ${waitTime}ms before retry attempt ${attempt + 1}...`);
          await new Promise((resolve) => setTimeout(resolve, waitTime));
          delay = waitTime * 2;
        } else {
          console.warn(`Model ${currentModel} reached its error threshold of ${actualAttemptsAllowed} attempt(s).`);
          if (currentModel === uniqueModels[uniqueModels.length - 1]) {
            throw err;
          }
          console.log(`Switching fallback to next available model...`);
          break; // break the inner loop to pivot to the next model in the outer loop
        }
      }
    }
  }
}

// Helper: Serialize sheets data into a compact, token-efficient CSV-like text representation
function serializeSheetsDataCompact(sheetsData: any[]): string {
  let output = "";
  for (const sheet of sheetsData) {
    output += `Sheet: ${sheet.sheetName}\n`;
    for (const row of sheet.rows) {
      const cellParts: string[] = [];
      const colIndices = Object.keys(row.cells).map(Number).sort((a, b) => a - b);
      for (const col of colIndices) {
        cellParts.push(`Col${col}: "${row.cells[col]}"`);
      }
      output += `  Row ${row.rowIndex}: ${cellParts.join(" | ")}\n`;
    }
    output += "\n";
  }
  return output.trim();
}

// Helper: Fetch Live CNY -> MYR exchange rate with fallback
async function fetchExchangeRate(): Promise<number> {
  try {
    const response = await fetch("https://open.er-api.com/v6/latest/CNY");
    if (response.ok) {
      const data = await response.json();
      if (data && data.rates && data.rates.MYR) {
        return Number(data.rates.MYR);
      }
    }
  } catch (error) {
    console.error("Failed to fetch live exchange rate from API, using fallback:", error);
  }
  return 0.60; // Standard fallback (1 CNY = 0.60 MYR)
}

// 1. API: Get Live Exchange Rate
app.get("/api/exchange-rate", async (req, res) => {
  const rate = await fetchExchangeRate();
  res.json({ rate });
});

// 2. API: Convert Quotation
app.post("/api/convert-quotation", async (req, res) => {
  try {
    const { excelBase64, manualRate, useManualRate, customInstruction } = req.body;

    if (!excelBase64) {
      return res.status(400).json({ error: "No Excel file provided in request." });
    }

    // Determine exchange rate
    let exchangeRate = 0.60;
    if (useManualRate && manualRate) {
      exchangeRate = Number(manualRate);
    } else {
      exchangeRate = await fetchExchangeRate();
    }

    // Decode and load Excel file using ExcelJS
    const buffer = Buffer.from(excelBase64, "base64");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // Image Extraction Mapping
    const extractedImages: { [sheetName: string]: Array<{ row: number; col: number; base64: string }> } = {};

    workbook.worksheets.forEach((worksheet) => {
      const imagesInSheet: Array<{ row: number; col: number; base64: string }> = [];
      worksheet.getImages().forEach((img) => {
        try {
          const imageObj = workbook.getImage(Number(img.imageId));
          if (!imageObj) return;

          let base64Img = "";
          if (Buffer.isBuffer(imageObj.buffer)) {
            base64Img = imageObj.buffer.toString("base64");
          } else if (imageObj.base64) {
            base64Img = imageObj.base64;
          } else {
            return;
          }

          const range = img.range;
          if (range && range.tl) {
            imagesInSheet.push({
              row: Math.floor(range.tl.row),
              col: Math.floor(range.tl.col),
              base64: `data:image/${imageObj.extension || "png"};base64,${base64Img}`
            });
          }
        } catch (err) {
          console.error("Error reading image metadata on sheet:", worksheet.name, err);
        }
      });
      extractedImages[worksheet.name] = imagesInSheet;
    });

    // Content Extraction (grid representation)
    const sheetsData: Array<{ 
      sheetName: string; 
      rows: Array<{ 
        rowIndex: number; 
        cells: { [col: number]: string };
        cellStyles?: { [col: number]: { bg?: string; color?: string; bold?: boolean; align?: string } };
      }>;
      images?: Array<{ row: number; col: number; base64: string }>;
    }> = [];

    workbook.worksheets.forEach((worksheet) => {
      const wsRows: Array<{ 
        rowIndex: number; 
        cells: { [colIndex: number]: string };
        cellStyles?: { [colIndex: number]: { bg?: string; color?: string; bold?: boolean; align?: string } };
      }> = [];
      
      let maxRowWithContent = 1;
      
      // Update maxRow based on rows with cell values
      worksheet.eachRow({ includeEmpty: false }, (row, rowIndex) => {
        let rowHasVal = false;
        row.eachCell({ includeEmpty: true }, (cell) => {
          if (cell.value !== null && cell.value !== undefined && String(cell.value).trim()) {
            rowHasVal = true;
          }
        });
        if (rowHasVal) {
          maxRowWithContent = Math.max(maxRowWithContent, rowIndex);
        }
      });

      // Update maxRow based on images in the sheet
      const wsImages = extractedImages[worksheet.name] || [];
      wsImages.forEach((img) => {
        maxRowWithContent = Math.max(maxRowWithContent, img.row + 1);
      });

      // Read contiguous rows from row 1 up to maxRowWithContent
      for (let r = 1; r <= maxRowWithContent; r++) {
        const row = worksheet.getRow(r);
        const rowVals: { [col: number]: string } = {};
        const stylesToSave: { [col: number]: { bg?: string; color?: string; bold?: boolean; align?: string } } = {};

        row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
          let val = "";
          if (cell.value !== null && cell.value !== undefined) {
            if (typeof cell.value === "object") {
              const cellObj = cell.value as any;
              if ("result" in cellObj && cellObj.result !== undefined && cellObj.result !== null) {
                val = String(cellObj.result);
              } else if ("richText" in cellObj && Array.isArray(cellObj.richText)) {
                val = cellObj.richText.map((rt: any) => rt.text).join("");
              } else if ("text" in cellObj && cellObj.text) {
                val = String(cellObj.text);
              } else {
                val = JSON.stringify(cellObj);
              }
            } else {
              val = String(cell.value);
            }
          }

          if (val && val.trim()) {
            rowVals[colIndex] = val.trim();
          }

          // Extract style information
          const style: any = {};
          
          // 1. Fill/background style
          if (cell.fill && cell.fill.type === "pattern" && cell.fill.pattern === "solid") {
            const fgColor = cell.fill.fgColor;
            if (fgColor && fgColor.argb) {
              const argb = String(fgColor.argb);
              style.bg = argb.length === 8 ? `#${argb.substring(2)}` : `#${argb}`;
            }
          }

          // 2. Font Bold & Color
          if (cell.font) {
            if (cell.font.bold) {
              style.bold = true;
            }
            if (cell.font.color && cell.font.color.argb) {
              const argb = String(cell.font.color.argb);
              style.color = argb.length === 8 ? `#${argb.substring(2)}` : `#${argb}`;
            }
          }

          // 3. Alignment
          if (cell.alignment && cell.alignment.horizontal) {
            style.align = cell.alignment.horizontal; // e.g. 'left' | 'center' | 'right'
          }

          if (Object.keys(style).length > 0) {
            stylesToSave[colIndex] = style;
          }
        });

        wsRows.push({
          rowIndex: r - 1,
          cells: rowVals,
          cellStyles: stylesToSave
        });
      }

      if (wsRows.length > 0) {
        sheetsData.push({
          sheetName: worksheet.name,
          rows: wsRows,
          images: wsImages
        });
      }
    });

    if (sheetsData.length === 0) {
      return res.status(400).json({ error: "The uploaded Excel file appears to contain no text details." });
    }

    // AI Instruction and Prompt (using gemini-3.5-flash as the standard text task choice)
    const prompt = `
You are a highly detailed and intelligent furniture supplier invoice and quotation specialist.
You are given a raw structure of rows from worksheets of a Chinese supplier’s quotation Excel file.

Your output must be a clean, parsed JSON array of objects representing distinct PRODUCT ITEMS only.

${customInstruction && customInstruction.trim() ? `CRITICAL - OVERRIDE USER CUSTOM REQUEST TO BE METICULOUSLY APPLIED:
"${customInstruction}"

Be sure to obey this instruction strictly when identify, filter, translate, and compute prices!` : ""}

Rules for identification:
- Correctly locate product rows. Do NOT extract metadata rows, header labels, contact details, totals, notes, template structures, or blank records.
- For each product line item:
  - "sheetName": the name of the worksheet where the item resides.
  - "excelRowIndex": the provided integer row index (0-indexed). Keep it exactly correct.
  - "originalName": the raw Chinese product title (e.g. "实木餐桌" or "胡桃木大床").
  - "originalSpecs": dimensions or specific specs (e.g. "1800*900*750mm" or "W150*D50*H200cm").
  - "originalColor": color descriptions (e.g. "胡桃色", "象牙白", "天蓝色").
  - "originalMaterial": materials details (e.g. "高密度海绵", "胡桃木实木框", "PU皮面").
  - "originalDescription": other remarks or raw details.
  - "itemCode": supplier product codes/IDs (e.g. "MC-04A" or "WF1014") if any.
  - "modelNum": model numbers or references.
  - "quantity": extracting numbers. Default to 1 if empty.
  - "unitPriceCNY": extract unit cost in Chinese Yuan.
  - "totalPriceCNY": extract total cost in Chinese Yuan. If empty, compute quantity * unitPriceCNY.
  - "remarks": any special notes of product specifications. DO NOT contain raw factory cost margin values or production secrets.

Translation rules:
- Translate all Chinese content to professional natural business English designed specifically for furniture customers (e.g. for MOCOF showroom).
  - "实木餐桌" -> "Solid Wood Dining Table"
  - "高密度海绵" -> "High Density Foam"
  - "胡桃木饰面" -> "Walnut Veneer Finish"
  - "胡桃色" -> "Walnut Color"
- Keep product codes, model numbers, and measurements (e.g. '1800*900*750mm') completely unchanged.
  - Provide translation inside: "translatedName", "translatedSpecs", "translatedColor", "translatedMaterial", "translatedDescription".

Be extremely smart in extracting details even if the supplier layout contains custom column formatting or merge cells. Remove all hidden/factory cost calculations and supplier-only references, showing only final polished descriptions.

Spreadsheet Grid Data (Compact View):
${serializeSheetsDataCompact(sheetsData)}
`;

    // Call Gemini API with response schema via robust fallback/retry engine
    const response = await generateContentWithRetryAndFallback({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              sheetName: { type: Type.STRING },
              excelRowIndex: { type: Type.INTEGER },
              originalName: { type: Type.STRING },
              originalSpecs: { type: Type.STRING },
              originalColor: { type: Type.STRING },
              originalMaterial: { type: Type.STRING },
              originalDescription: { type: Type.STRING },
              translatedName: { type: Type.STRING },
              translatedSpecs: { type: Type.STRING },
              translatedColor: { type: Type.STRING },
              translatedMaterial: { type: Type.STRING },
              translatedDescription: { type: Type.STRING },
              quantity: { type: Type.NUMBER },
              unitPriceCNY: { type: Type.NUMBER },
              totalPriceCNY: { type: Type.NUMBER },
              itemCode: { type: Type.STRING },
              modelNum: { type: Type.STRING },
              remarks: { type: Type.STRING }
            },
            required: ["sheetName", "excelRowIndex", "originalName", "translatedName", "quantity", "unitPriceCNY"]
          }
        }
      }
    });

    const aiText = response.text || "[]";
    const parsedItems = JSON.parse(aiText.trim());

    // Map extracted Images to items based on Sheet Name and Row Index proximity
    const itemsWithImages = parsedItems.map((item: any) => {
      const sheet = item.sheetName;
      const rowIdx = item.excelRowIndex;
      const imagesList = extractedImages[sheet] || [];

      // Find an image on the exact same row, or row offset of +/- 1
      // Preferred column is lower indices (like Column A or B, which usually hold product thumbnails)
      let matchedImage = "";
      
      const exactMatch = imagesList.find(img => img.row === rowIdx);
      if (exactMatch) {
        matchedImage = exactMatch.base64;
      } else {
        const nearMatch = imagesList.find(img => Math.abs(img.row - rowIdx) <= 1);
        if (nearMatch) {
          matchedImage = nearMatch.base64;
        }
      }

      // Convert pricing factors to MYR
      const unitCNY = Number(item.unitPriceCNY) || 0;
      const qty = Number(item.quantity) || 1;
      const totalPriceCNY = qty * unitCNY;

      const unitMYR = Math.round((unitCNY * exchangeRate) * 100) / 100;
      const totalMYR = Math.round((totalPriceCNY * exchangeRate) * 100) / 100;

      return {
        id: `col_${sheet}_${rowIdx}_${Math.random().toString(36).substr(2, 5)}`,
        originalName: item.originalName || "",
        originalSpecs: item.originalSpecs || "",
        originalColor: item.originalColor || "",
        originalMaterial: item.originalMaterial || "",
        originalDescription: item.originalDescription || "",
        translatedName: item.translatedName || "",
        translatedSpecs: item.translatedSpecs || "",
        translatedColor: item.translatedColor || "",
        translatedMaterial: item.translatedMaterial || "",
        translatedDescription: item.translatedDescription || "",
        quantity: qty,
        unitPriceCNY: unitCNY,
        totalPriceCNY: totalPriceCNY,
        unitPriceMYR: unitMYR,
        totalPriceMYR: totalMYR,
        image: matchedImage, // Holds the extracted base64 data URL
        remarks: item.remarks || "",
        itemCode: item.itemCode || "",
        modelNum: item.modelNum || ""
      };
    });

    // Run Verification / Alert Auditing Engine
    const alerts: Array<{ id: string; type: "info" | "warning" | "error"; message: string; itemId?: string }> = [];
    const nameMap = new Map<string, string>();

    itemsWithImages.forEach((item: any) => {
      // 1. Missing quantity
      if (item.quantity <= 0) {
        alerts.push({
          id: `alert_qty_${item.id}`,
          type: "warning",
          message: `Item "${item.translatedName || item.originalName}" has a missing or invalid quantity (${item.quantity}). Setting to 1.`,
          itemId: item.id
        });
        item.quantity = 1;
        item.totalPriceCNY = item.unitPriceCNY;
        item.totalPriceMYR = item.unitPriceMYR;
      }

      // 2. Missing price
      if (item.unitPriceCNY <= 0) {
        alerts.push({
          id: `alert_val_${item.id}`,
          type: "warning",
          message: `Item "${item.translatedName || item.originalName}" is marked with a unit price of zero CNY. Please verify.`,
          itemId: item.id
        });
      }

      // 3. Formula alignment issues
      const expectedTotalCNY = Math.round(item.quantity * item.unitPriceCNY * 100) / 100;
      if (Math.abs(item.totalPriceCNY - expectedTotalCNY) > 0.1) {
        alerts.push({
          id: `alert_formula_${item.id}`,
          type: "info",
          message: `Item "${item.translatedName || item.originalName}" total CNY mismatched spreadsheet. Recalculated using standard math.`,
          itemId: item.id
        });
        item.totalPriceCNY = expectedTotalCNY;
        item.totalPriceMYR = Math.round((expectedTotalCNY * exchangeRate) * 100) / 100;
      }

      // 4. Duplicate Item check
      const uniqueKey = `${item.translatedName.toLowerCase().trim()}_spec_${item.translatedSpecs.toLowerCase().trim()}`;
      if (nameMap.has(uniqueKey)) {
        alerts.push({
          id: `alert_dup_${item.id}`,
          type: "info",
          message: `Item "${item.translatedName}" with equivalent specification appears multiple times. Verification advised.`,
          itemId: item.id
        });
      } else {
        nameMap.set(uniqueKey, item.id);
      }

      // 5. Image status check
      if (!item.image) {
        alerts.push({
          id: `alert_img_${item.id}`,
          type: "info",
          message: `No image automatically detected in spreadsheet for "${item.translatedName}". You can upload one manually in the table.`,
          itemId: item.id
        });
      }
    });

    // Compute Quote metrics
    const subtotalMYR = itemsWithImages.reduce((sum: number, it: any) => sum + it.totalPriceMYR, 0);

    // Terms and Conditions fallbacks
    const defaultTerms = [
      "Prices are in Malaysian Ringgit (MYR).",
      "Quotations are valid for 30 days from issued date.",
      "Delivery and installation lead time is 6 to 8 weeks upon order placement and deposit payment.",
      "50% deposit payment upon order confirmation, 50% remainder upon delivery.",
      "Warranties: 1-year manufacturing defect warranty on structural elements."
    ];

    res.json({
      items: itemsWithImages,
      exchangeRateUsed: exchangeRate,
      subtotalMYR: Math.round(subtotalMYR * 100) / 100,
      alerts,
      defaultTerms,
      rawSheets: sheetsData,
      date: new Date().toISOString().split("T")[0]
    });

  } catch (error: any) {
    console.error("Quotation Conversion Server Error:", error);
    res.status(500).json({ error: error?.message || "An internal error occurred during the conversion process." });
  }
});

// 3. API: Curate & Correct Quotation Items via Chat Command
app.post("/api/curate-quotation", async (req, res) => {
  try {
    const { items, instruction, exchangeRate } = req.body;

    if (!instruction || !instruction.trim()) {
      return res.status(400).json({ error: "Instruction cannot be blank." });
    }

    const rate = Number(exchangeRate) || 0.60;

    // Prompt Gemini with exact instruction and list with indices, asking it to detect if it's an edit vs. simple question
    const prompt = `
You are an expert furniture supply chains analyst and data curation supervisor.
Your job is to read the user's natural language input, look at the active list of quotation items, and determine if the user is asking a QUESTION/QUERY about the data (e.g., "why are there warnings?", "what is the total quantity of chairs?", "which row is the most expensive?") or giving a modification COMMAND/EDIT (e.g., "remove row 3", "cut row 3", "add bed for RM 1500", "change Row 1 quantity to 5").

User request: "${instruction}"

Current active list of items:
${JSON.stringify(items.map((item: any, idx: number) => ({
  rowNumber: idx + 1, // 1-indexed for easy user mapping
  id: item.id,
  originalName: item.originalName,
  translatedName: item.translatedName,
  translatedSpecs: item.translatedSpecs,
  translatedColor: item.translatedColor,
  translatedMaterial: item.translatedMaterial,
  translatedDescription: item.translatedDescription,
  quantity: item.quantity,
  unitPriceCNY: item.unitPriceCNY,
  itemCode: item.itemCode,
  modelNum: item.modelNum,
  remarks: item.remarks,
  image: item.image ? "[has_image_base64_data]" : ""
})))}

Rules for response:
1. Classify the user instruction strictly.
   - If it is a QUESTION (e.g. requesting calculations, comparison, list query, explanation or inquiry of why certain values exist): set "isEdit" to false. Do NOT modify any items! Preserve the identical "items" list exactly, and craft an elegant, detailed, and beautifully structured query answer in "changeDescription". Use Markdown or simple layouts (e.g., bullets) for easy reading.
   - If it is an EDIT or deletion or addition command (e.g. "remove row X", "change row Y price to RM Z", "add new row"): set "isEdit" to true. Apply the edit meticulously:
     - "remove row X" or "cut row X" means delete the item at rowNumber = X.
     - "add a row" or "add a table" means append a new quotation item. Generate appropriate English translations, specifications, and fill standard fields intelligently.
     - "change price of column Y in row X" or "change row X's Qty to Y" means update that item.
2. If instructions specify a price in MYR (e.g., "change row 2 price to RM 300"), convert it to CNY first by dividing by the exchange rate (Current exchange rate used: 1 CNY = RM ${rate}). For example: 300 / ${rate} = ${Math.round(300/rate * 100)/100} CNY.
3. Preserve existing item IDs, image data (for existing items, keep the same Base64 image payload if present), and unedited fields. If adding a new item, generate a random ID.
4. Give a clear, friendly summary of your action or your exact conversational answer in "changeDescription".

IMPORTANT: Respond in exact JSON matching the schema format.
`;

    const response = await generateContentWithRetryAndFallback({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            isEdit: { type: Type.BOOLEAN },
            items: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  id: { type: Type.STRING },
                  originalName: { type: Type.STRING },
                  translatedName: { type: Type.STRING },
                  translatedSpecs: { type: Type.STRING },
                  translatedColor: { type: Type.STRING },
                  translatedMaterial: { type: Type.STRING },
                  translatedDescription: { type: Type.STRING },
                  quantity: { type: Type.NUMBER },
                  unitPriceCNY: { type: Type.NUMBER },
                  itemCode: { type: Type.STRING },
                  modelNum: { type: Type.STRING },
                  remarks: { type: Type.STRING },
                  image: { type: Type.STRING }
                },
                required: ["translatedName", "quantity", "unitPriceCNY"]
              }
            },
            changeDescription: {
              type: Type.STRING
            }
          },
          required: ["isEdit", "items", "changeDescription"]
        }
      }
    });

    const bodyText = response.text || "{}";
    const parsed = JSON.parse(bodyText.trim());

    // If it's not an edit, we keep client-side items exactly as is to preserve all local hand-edited fields
    if (parsed.isEdit === false) {
      return res.json({
        isEdit: false,
        items: items, // No-op, preserve original items list
        changeDescription: parsed.changeDescription || "No changes requested for active quotation worksheet."
      });
    }

    // Reinforce values on backend for correct math and data formats
    const updatedItems = parsed.items.map((item: any) => {
      // Find original item to preserve base64 image if exists
      const originalItem = items.find((o: any) => o.id === item.id);
      let finalImg = item.image || "";
      if (originalItem && (finalImg === "[has_image_base64_data]" || !finalImg)) {
        finalImg = originalItem.image || "";
      }

      const q = Math.max(1, Number(item.quantity) || 1);
      const cnyPrice = Math.max(0, Number(item.unitPriceCNY) || 0);
      const totPriceCNY = q * cnyPrice;

      const myrPrice = Math.round((cnyPrice * rate) * 100) / 100;
      const totPriceMYR = Math.round((totPriceCNY * rate) * 100) / 100;

      return {
        id: item.id || `col_curated_${Math.random().toString(36).substr(2, 5)}`,
        originalName: item.originalName || "",
        originalSpecs: item.originalSpecs || "",
        originalColor: item.originalColor || "",
        originalMaterial: item.originalMaterial || "",
        originalDescription: item.originalDescription || "",
        translatedName: item.translatedName || "",
        translatedSpecs: item.translatedSpecs || "",
        translatedColor: item.translatedColor || "",
        translatedMaterial: item.translatedMaterial || "",
        translatedDescription: item.translatedDescription || "",
        quantity: q,
        unitPriceCNY: cnyPrice,
        totalPriceCNY: totPriceCNY,
        unitPriceMYR: myrPrice,
        totalPriceMYR: totPriceMYR,
        image: finalImg,
        remarks: item.remarks || "",
        itemCode: item.itemCode || "",
        modelNum: item.modelNum || ""
      };
    });

    res.json({
      isEdit: true,
      items: updatedItems,
      changeDescription: parsed.changeDescription || "Quotation items updated successfully."
    });

  } catch (error: any) {
    console.error("Quotation Curation Error:", error);
    res.status(500).json({ error: error?.message || "An error occurred while curating the quotation." });
  }
});

// Configure Vite or Static Assets depending on Environment
async function startApp() {
  if (process.env.NODE_ENV !== "production") {
    // Development mode - dynamic import to avoid production bundling issues on environments like Vercel
    const { createServer } = await import("vite");
    const vite = await createServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
    console.log("Vite development middleware loaded.");
  } else {
    // Production mode
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
    console.log("Production static files server loaded.");
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Application started and running at host 0.0.0.0:3000`);
  });
}

if (!process.env.VERCEL) {
  startApp().catch((err) => {
    console.error("Failed to start server:", err);
  });
}

export default app;
