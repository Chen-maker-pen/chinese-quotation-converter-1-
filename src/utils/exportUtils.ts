import ExcelJS from "exceljs";
import { jsPDF } from "jspdf";
import "jspdf-autotable";
import { QuotationData, QuotationItem } from "../types";

// Helper: Parse Base64 Image to clean base64 data and extension
interface ParsedImg {
  extension: "png" | "jpeg" | "gif";
  base64: string;
}

function parseBase64Image(dataUrl: string): ParsedImg | null {
  if (!dataUrl || !dataUrl.startsWith("data:image/")) return null;
  const matches = dataUrl.match(/^data:image\/([a-zA-Z+]+);base64,(.+)$/);
  if (!matches || matches.length !== 3) return null;
  
  let ext = matches[1].toLowerCase();
  if (ext === "jpg") ext = "jpeg";
  if (ext !== "png" && ext !== "jpeg" && ext !== "gif") return null;

  return {
    extension: ext as "png" | "jpeg" | "gif",
    base64: matches[2]
  };
}

// 1. Export Quotation to EXCEL
export async function exportToExcel(quotation: QuotationData) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Customer Quotation");

  // Gridlines enabled
  worksheet.views = [{ showGridLines: true }];

  // 1. Corporate Branding Title Block
  worksheet.mergeCells("A2:K2");
  const titleCell = worksheet.getCell("A2");
  titleCell.value = "MOCOF CUSTOMER QUOTATION";
  titleCell.font = { name: "Inter", size: 16, bold: true, color: { argb: "FFFFFF" } };
  titleCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "0F172A" } // Dark Slate #0f172a
  };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  worksheet.getRow(2).height = 40;

  // 2. Metadata / Information Section
  worksheet.getCell("A4").value = "CUSTOMER DETAILS";
  worksheet.getCell("A4").font = { name: "Inter", size: 11, bold: true };
  worksheet.getCell("A5").value = "Name:";
  worksheet.getCell("A5").font = { name: "Inter", bold: true };
  worksheet.getCell("B5").value = quotation.customerName;
  worksheet.getCell("B5").font = { name: "Inter" };

  worksheet.getCell("G4").value = "QUOTATION METADATA";
  worksheet.getCell("G4").font = { name: "Inter", size: 11, bold: true };
  worksheet.getCell("G5").value = "Quotation No:";
  worksheet.getCell("G5").font = { name: "Inter", bold: true };
  worksheet.getCell("H5").value = quotation.quotationNumber;
  worksheet.getCell("H5").font = { name: "Inter" };

  worksheet.getCell("G6").value = "Date:";
  worksheet.getCell("G6").font = { name: "Inter", bold: true };
  worksheet.getCell("H6").value = quotation.date;
  worksheet.getCell("H6").font = { name: "Inter" };

  worksheet.getCell("G7").value = "Prepared By:";
  worksheet.getCell("G7").font = { name: "Inter", bold: true };
  worksheet.getCell("H7").value = quotation.preparedBy;
  worksheet.getCell("H7").font = { name: "Inter" };

  worksheet.getCell("G8").value = "Exchange Rate:";
  worksheet.getCell("G8").font = { name: "Inter", bold: true };
  worksheet.getCell("H8").value = `1 CNY = ${quotation.exchangeRate} MYR`;
  worksheet.getCell("H8").font = { name: "Inter", italic: true };

  // Set widths for nice spacing
  worksheet.columns = [
    { key: "no", width: 6 },
    { key: "image", width: 16 },
    { key: "code", width: 14 },
    { key: "name", width: 26 },
    { key: "specs", width: 22 },
    { key: "material", width: 22 },
    { key: "color", width: 14 },
    { key: "qty", width: 10 },
    { key: "unit", width: 16 },
    { key: "total", width: 18 },
    { key: "remarks", width: 24 }
  ];

  // 3. Main Data Table Headers (Starting Row 10)
  const headerRowIdx = 10;
  const headers = [
    "No",
    "Product Item", // Image cell
    "Item Code",
    "Product Name",
    "Specifications",
    "Material & Finish",
    "Color",
    "Quantity",
    "Unit Price (MYR)",
    "Total (MYR)",
    "Remarks / Notes"
  ];

  const headerRow = worksheet.getRow(headerRowIdx);
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = header;
    cell.font = { name: "Inter", size: 10, bold: true, color: { argb: "FFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "334155" } // Slate color #334155
    };
    cell.alignment = { horizontal: index === 7 || index === 8 || index === 9 ? "right" : "left", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "medium" },
      right: { style: "thin" }
    };
  });
  headerRow.height = 25;

  let startRow = 11;
  quotation.items.forEach((item, index) => {
    const row = worksheet.getRow(startRow);
    row.height = 60; // Make row tall to comfortably host a thumbnail image

    // Set cells
    row.getCell(1).value = index + 1;
    row.getCell(3).value = item.itemCode || item.modelNum || "N/A";
    row.getCell(4).value = item.translatedName;
    row.getCell(5).value = item.translatedSpecs || "Standard Size";
    row.getCell(6).value = item.translatedMaterial || "-";
    row.getCell(7).value = item.translatedColor || "-";
    row.getCell(8).value = item.quantity;
    row.getCell(9).value = item.unitPriceMYR;
    
    // Use an Excel Formula for Total Price so the sheet remains dynamic and editable
    row.getCell(10).value = { formula: `H${startRow}*I${startRow}` };
    row.getCell(11).value = item.remarks || "";

    // Style and borders
    for (let c = 1; c <= 11; c++) {
      const cell = row.getCell(c);
      cell.font = { name: "Inter", size: 9.5 };
      cell.alignment = {
        horizontal: c === 8 || c === 9 || c === 10 ? "right" : "left",
        vertical: "middle",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "D1D5DB" } },
        left: { style: "thin", color: { argb: "D1D5DB" } },
        bottom: { style: "thin", color: { argb: "D1D5DB" } },
        right: { style: "thin", color: { argb: "D1D5DB" } }
      };

      // Number formatting
      if (c === 9 || c === 10) {
        cell.numFmt = '"RM "#,##0.00';
      }
    }

    // Embed Image safely if available in item
    if (item.image) {
      const prs = parseBase64Image(item.image);
      if (prs) {
        try {
          const imageId = workbook.addImage({
            base64: prs.base64,
            extension: prs.extension
          });
          worksheet.addImage(imageId, {
            tl: { col: 1, row: startRow - 1 } as any,
            br: { col: 2, row: startRow } as any,
            editAs: "oneCell"
          });
        } catch (imgErr) {
          console.error("Error embedding product image in Excel rendering:", imgErr);
        }
      }
    }

    startRow++;
  });

  // 4. Summaries & Totals Grid (Start dynamic rows)
  const lastProductRow = startRow - 1;
  const subtotalRowIdx = startRow;
  const taxRowIdx = startRow + 1;
  const grandTotalRowIdx = startRow + 2;

  // Subtotal row
  const subtotalRow = worksheet.getRow(subtotalRowIdx);
  subtotalRow.getCell(9).value = "Subtotal:";
  subtotalRow.getCell(9).font = { name: "Inter", bold: true };
  subtotalRow.getCell(10).value = { formula: `SUM(J11:J${lastProductRow})` };
  subtotalRow.getCell(10).font = { name: "Inter", bold: true };
  subtotalRow.getCell(10).numFmt = '"RM "#,##0.00';
  subtotalRow.getCell(10).alignment = { horizontal: "right" };

  // SST tax row
  const taxRow = worksheet.getRow(taxRowIdx);
  taxRow.getCell(9).value = `SST Service Tax (${quotation.sstPercentage}%):`;
  taxRow.getCell(9).font = { name: "Inter" };
  taxRow.getCell(10).value = { formula: `J${subtotalRowIdx}*${quotation.sstPercentage / 100}` };
  taxRow.getCell(10).font = { name: "Inter" };
  taxRow.getCell(10).numFmt = '"RM "#,##0.00';
  taxRow.getCell(10).alignment = { horizontal: "right" };

  // Grand Total row
  const grandTotalRow = worksheet.getRow(grandTotalRowIdx);
  grandTotalRow.getCell(9).value = "GRAND TOTAL (MYR):";
  grandTotalRow.getCell(9).font = { name: "Inter", bold: true, color: { argb: "FFFFFF" } };
  grandTotalRow.getCell(9).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F172A" } };
  
  grandTotalRow.getCell(10).value = { formula: `J${subtotalRowIdx}+J${taxRowIdx}` };
  grandTotalRow.getCell(10).font = { name: "Inter", size: 11, bold: true, color: { argb: "FFFFFF" } };
  grandTotalRow.getCell(10).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F172A" } };
  grandTotalRow.getCell(10).numFmt = '"RM "#,##0.00';
  grandTotalRow.getCell(10).alignment = { horizontal: "right" };

  worksheet.getRow(grandTotalRowIdx).height = 24;

  // 5. Terms and Conditions
  let termsStartRow = grandTotalRowIdx + 3;
  worksheet.getCell(`A${termsStartRow}`).value = "TERMS & CONDITIONS";
  worksheet.getCell(`A${termsStartRow}`).font = { name: "Inter", size: 11, bold: true };
  
  quotation.termsAndConditions.forEach((term, idx) => {
    termsStartRow++;
    worksheet.mergeCells(`A${termsStartRow}:F${termsStartRow}`);
    const cell = worksheet.getCell(`A${termsStartRow}`);
    cell.value = `•  ${term}`;
    cell.font = { name: "Inter", size: 9, italic: true };
  });

  // Save/Download Excel file using window Blob URL
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const downloadUrl = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = downloadUrl;
  anchor.download = `MOCOF_Quotation_${quotation.quotationNumber || "MYR"}.xlsx`;
  anchor.click();
  window.URL.revokeObjectURL(downloadUrl);
}

// 2. Export Quotation to A4 Portrait PDF
export async function exportToPDF(quotation: QuotationData) {
  const doc = new jsPDF("p", "mm", "a4");
  const pageWidth = doc.internal.pageSize.width;
  const pageHeight = doc.internal.pageSize.height;

  // Styling palette
  const mainColor = [15, 23, 42]; // Slate 900
  const secondaryColor = [51, 65, 85]; // Slate 700
  const lightGrey = [241, 245, 249]; // Slate 100

  // Title branding
  doc.setFillColor(mainColor[0], mainColor[1], mainColor[2]);
  doc.rect(0, 0, pageWidth, 28, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("MOCOF QUOTATION SUMMARY", 15, 12);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text("Bespoke Furniture Importers & Design specialists", 15, 18);

  // Logo text placeholder on the header right
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("MOCOF", pageWidth - 42, 16);

  // 1. Transaction metadata
  doc.setTextColor(51, 51, 51);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("CLIENT SUMMARY:", 15, 38);
  doc.setFont("helvetica", "normal");
  doc.text(`Name:   ${quotation.customerName || "Valued Client"}`, 15, 43);

  doc.setFont("helvetica", "bold");
  doc.text("QUOTATION DETAILS:", pageWidth - 90, 38);
  doc.setFont("helvetica", "normal");
  doc.text(`Quotation No:  ${quotation.quotationNumber}`, pageWidth - 90, 43);
  doc.text(`Date:               ${quotation.date}`, pageWidth - 90, 48);
  doc.text(`Prepared By:    ${quotation.preparedBy}`, pageWidth - 90, 53);
  doc.text(`Exchange Rate:  1 CNY = ${quotation.exchangeRate} MYR`, pageWidth - 90, 58);

  // Divider
  doc.setDrawColor(226, 232, 240);
  doc.line(15, 63, pageWidth - 15, 63);

  // 2. Prepare tabular data for jsPDF AutoTable
  const headers = ["No", "Image", "Details", "Qty", "Unit Price", "Total (MYR)"];
  
  const body = quotation.items.map((item, index) => {
    // Compile translated details into a nicely structured visual block
    const specsStr = item.translatedSpecs ? `Dims: ${item.translatedSpecs}` : "Standard specifications";
    const detailBlock = [
      `${index + 1}.  ${item.translatedName}`,
      `Code: ${item.itemCode || item.modelNum || "N/A"}`,
      specsStr,
      item.translatedMaterial ? `Material: ${item.translatedMaterial}` : null,
      item.translatedColor ? `Color: ${item.translatedColor}` : null,
      item.remarks ? `Remarks: ${item.remarks}` : null
    ].filter(Boolean).join("\n");

    return [
      index + 1,
      item.image, // Passed to didDrawCell custom drawer
      detailBlock,
      item.quantity,
      `RM ${item.unitPriceMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      `RM ${item.totalPriceMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
    ];
  });

  // Calculate table positions dynamically
  (doc as any).autoTable({
    startY: 68,
    head: [headers],
    body: body,
    theme: "striped",
    headStyles: {
      fillColor: mainColor,
      textColor: [255, 255, 255],
      fontSize: 9,
      fontStyle: "bold",
      halign: "left"
    },
    columnStyles: {
      0: { cellWidth: 10, halign: "center" },
      1: { cellWidth: 26, halign: "center" }, // Reserve spacing for images
      2: { cellWidth: 80, fontSize: 8.5 },
      3: { cellWidth: 12, halign: "center" },
      4: { cellWidth: 28, halign: "right" },
      5: { cellWidth: 28, halign: "right" }
    },
    styles: {
      valign: "middle"
    },
    // Custom drawer so we draw base64 image perfectly inside the reserved cell
    didDrawCell: (data: any) => {
      if (data.column.index === 1 && data.cell.section === "body" && data.cell.raw) {
        try {
          const base64Url = data.cell.raw;
          const prs = parseBase64Image(base64Url);
          if (prs) {
            // Draw thumbnail image nicely centered in cell box
            doc.addImage(
              base64Url,
              prs.extension.toUpperCase(),
              data.cell.x + 3,
              data.cell.y + 2,
              20,
              20
            );
          }
        } catch (err) {
          console.error("Failed to render item thumbnail image in PDF output", err);
        }
      }
    },
    willDrawCell: (data: any) => {
      // Clear image cell value before printing text inside it
      if (data.column.index === 1 && data.cell.section === "body") {
        data.cell.text = "";
      }
    },
    rowPageBreak: "avoid",
    margin: { left: 15, right: 15 }
  });

  // 3. Summaries & totals positioning directly underneath the table
  let finalY = (doc as any).lastAutoTable.finalY + 10;

  // Add page break if total box will overflow on the page bottom
  if (finalY > pageHeight - 65) {
    doc.addPage();
    finalY = 20;
  }

  // Draw Totals Box
  const summaryBoxWidth = 80;
  const summaryBoxX = pageWidth - 15 - summaryBoxWidth;

  doc.setFillColor(lightGrey[0], lightGrey[1], lightGrey[2]);
  doc.rect(summaryBoxX, finalY, summaryBoxWidth, 34, "F");
  
  doc.setFontSize(9);
  doc.setTextColor(51, 51, 51);
  doc.setFont("helvetica", "normal");
  doc.text("Subtotal:", summaryBoxX + 5, finalY + 7);
  doc.text(`SST Service Tax (${quotation.sstPercentage}%):`, summaryBoxX + 5, finalY + 16);
  
  doc.setFont("helvetica", "bold");
  doc.text("Grand Total (MYR):", summaryBoxX + 5, finalY + 27);

  doc.setFont("helvetica", "normal");
  const subTotalStr = `RM ${quotation.subtotalMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const sstStr = `RM ${quotation.sstAmountMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const grandTotalStr = `RM ${quotation.grandTotalMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  doc.text(subTotalStr, pageWidth - 20, finalY + 7, { align: "right" });
  doc.text(sstStr, pageWidth - 20, finalY + 16, { align: "right" });
  doc.setFont("helvetica", "bold");
  doc.text(grandTotalStr, pageWidth - 20, finalY + 27, { align: "right" });

  // 4. Terms and signature block
  let termsY = finalY + 45;
  if (termsY > pageHeight - 50) {
    doc.addPage();
    termsY = 20;
  }

  // Terms label
  doc.setFontSize(9.5);
  doc.setFont("helvetica", "bold");
  doc.text("TERMS AND CONDITIONS", 15, termsY);
  doc.setFont("helvetica", "italic");
  doc.setFontSize(8);
  doc.setTextColor(102, 102, 102);
  
  let offset = 5;
  quotation.termsAndConditions.forEach((term, index) => {
    doc.text(`•  ${term}`, 15, termsY + offset);
    offset += 4.5;
  });

  // Footer / Signature Section
  const signatureY = termsY + offset + 15;
  if (signatureY < pageHeight - 30) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(51, 51, 51);
    
    // Prepared by side
    doc.line(15, signatureY + 12, 65, signatureY + 12);
    doc.text("Prepared By (MOCOF)", 15, signatureY + 16);
    doc.setFontSize(7.5);
    doc.text(quotation.preparedBy, 15, signatureY + 20);

    // Client approval side
    doc.setFontSize(9);
    doc.line(pageWidth - 65, signatureY + 12, pageWidth - 15, signatureY + 12);
    doc.text("Accepted By (Customer Sign)", pageWidth - 65, signatureY + 16);
  }

  // Print/Download PDF
  doc.save(`MOCOF_Quotation_${quotation.quotationNumber}.pdf`);
}
