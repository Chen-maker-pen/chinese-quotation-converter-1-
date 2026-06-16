import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import {
  Upload,
  FileSpreadsheet,
  ArrowRightLeft,
  FileText,
  AlertTriangle,
  Info,
  RefreshCw,
  Plus,
  Trash2,
  Check,
  CheckCircle2,
  ShieldAlert,
  Download,
  Percent,
  User,
  Settings,
  HelpCircle,
  Hash,
  Sparkles,
  ChevronRight,
  Eye,
  Camera,
  Layers,
  ArrowLeft,
  Mic,
  Send,
  MessageSquare,
  Search,
  TrendingUp
} from "lucide-react";
import { QuotationData, QuotationItem, VerificationAlert } from "./types";
import { exportToExcel, exportToPDF } from "./utils/exportUtils";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip as RechartsTooltip,
  Legend as RechartsLegend,
  PieChart,
  Pie,
  Cell
} from "recharts";

// Helper to get category for each item dynamically based on its English description/name
const getCategoryForItem = (name: string): string => {
  const n = name.toLowerCase();
  if (n.includes("sofa") || n.includes("couch") || n.includes("settee") || n.includes("daybed")) return "Sofa & Seating";
  if (n.includes("chair") || n.includes("stool") || n.includes("armchair") || n.includes("lounger") || n.includes("bench")) return "Chairs & Barstools";
  if (n.includes("table") || n.includes("desk") || n.includes("console") || n.includes("island") || n.includes("board")) return "Tables & Desks";
  if (n.includes("cabinet") || n.includes("wardrobe") || n.includes("cupboard") || n.includes("shelf") || n.includes("shelves") || n.includes("sideboard") || n.includes("chest") || n.includes("drawer") || n.includes("rack") || n.includes("stand") || n.includes("credenza")) return "Cabinets & Storage";
  if (n.includes("bed") || n.includes("mattress") || n.includes("headboard")) return "Beds & Mattresses";
  if (n.includes("mirror")) return "Mirrors & Glass";
  if (n.includes("lamp") || n.includes("light") || n.includes("chandelier")) return "Lighting & Luminaires";
  if (n.includes("curtain") || n.includes("blind") || n.includes("rug") || n.includes("carpet") || n.includes("cushion") || n.includes("pillow")) return "Soft Furnishings & Rugs";
  return "Bespoke & Other Furniture";
};

// Convert column position (1, 2, 3...) to standard Excel alphabetical symbols (A, B, C... Z, AA, AB...)
const getColumnLetter = (colIndex: number): string => {
  let temp = colIndex;
  let letter = "";
  while (temp > 0) {
    let modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }
  return letter || `Col ${colIndex}`;
};

// Colors for the chart sectors & bars representing furniture components
const CHART_COLORS = [
  "#4F46E5", // Indigo 600
  "#06B6D4", // Cyan 500
  "#F59E0B", // Amber 500
  "#10B981", // Emerald 500
  "#EC4899", // Pink 500
  "#8B5CF6", // Violet 500
  "#14B8A6", // Teal 500
  "#FF6B6B", // Coral Red
  "#94A3B8"  // Slate Gray
];

export default function App() {
  // Setup fields
  const [customerName, setCustomerName] = useState("MOCOF Valued Client");
  const [preparedBy, setPreparedBy] = useState("MOCOF Supply Team");
  const [quotationNumber, setQuotationNumber] = useState("");
  const [sstPercentage, setSstPercentage] = useState(6);

  // File Upload states
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [excelBase64, setExcelBase64] = useState<string>("");
  const [isDragOver, setIsDragOver] = useState(false);

  // Exchange rate states
  const [liveRate, setLiveRate] = useState<number>(0.60);
  const [useManualRate, setUseManualRate] = useState(false);
  const [manualExchangeRate, setManualExchangeRate] = useState<number>(0.60);
  const [fetchingRate, setFetchingRate] = useState(false);

  // App running states
  const [isConverting, setIsConverting] = useState(false);
  const [convertError, setConvertError] = useState<string | null>(null);
  const [convertingStep, setConvertingStep] = useState(0);
  const [activeStepText, setActiveStepText] = useState("");

  // Output Quotation Data State
  const [quotation, setQuotation] = useState<QuotationData | null>(null);
  const [verificationConfirmed, setVerificationConfirmed] = useState(false);
  
  // Modal for individual image inspection
  const [activePreviewImage, setActivePreviewImage] = useState<string | null>(null);

  // Workspace sub-tabs and raw excel visualizer configuration states
  const [workspaceActiveTab, setWorkspaceActiveTab] = useState<"quotation" | "rawExcel" | "insights">("quotation");
  const [selectedRawSheetIndex, setSelectedRawSheetIndex] = useState<number>(0);
  const [rawSheetSearch, setRawSheetSearch] = useState("");

  // Row selection or insertion state
  const [newItemName, setNewItemName] = useState("");
  const [newItemPriceCNY, setNewItemPriceCNY] = useState("");
  const [newItemQty, setNewItemQty] = useState("1");
  const [newItemSpecs, setNewItemSpecs] = useState("");

  // AI Curation Assistant States
  const [preCurationInstruction, setPreCurationInstruction] = useState("");
  const [curationInstruction, setCurationInstruction] = useState("");
  const [isCurating, setIsCurating] = useState(false);
  const [aiChangeSummary, setAiChangeSummary] = useState<string | null>(null);
  const [curationError, setCurationError] = useState<string | null>(null);
  const [chatHistory, setChatHistory] = useState<Array<{
    role: "user" | "assistant";
    content: string;
    isCommand?: boolean;
    timestamp: string;
  }>>([
    {
      role: "assistant",
      content: "Hello! I am your furniture sheet curation assistant. You can specify edits (e.g. 'remove row 3', 'add 10% to row 2 prices') or ask questions about the worksheet data (e.g. 'how many total tables?', 'which furniture items are wooden?'). Try typing below or click a chip suggestion!",
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    }
  ]);

  // Dynamic memoized charts data calculations
  const categoryChartData = React.useMemo(() => {
    if (!quotation) return [];
    const map: Record<string, { name: string; value: number; count: number }> = {};
    
    quotation.items.forEach(item => {
      const cat = getCategoryForItem(item.translatedName);
      const val = item.totalPriceMYR || 0;
      if (!map[cat]) {
        map[cat] = { name: cat, value: 0, count: 0 };
      }
      map[cat].value += val;
      map[cat].count += item.quantity || 1;
    });

    return Object.values(map)
      .filter(item => item.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [quotation]);

  const itemChartData = React.useMemo(() => {
    if (!quotation) return [];
    return quotation.items
      .map((item, index) => ({
        shortName: `Row ${index + 1}`,
        fullName: item.translatedName,
        rowNumber: index + 1,
        value: item.totalPriceMYR || 0,
        unitPrice: item.unitPriceMYR || 0,
        quantity: item.quantity || 0,
      }))
      .filter(item => item.value > 0)
      .sort((a, b) => b.value - a.value)
      .slice(0, 8); // show top 8 items max to avoid cluttering, with full transparency in tooltip
  }, [quotation]);

  // Total sum of item value for percentage calculations
  const chartTotalCost = React.useMemo(() => {
    return categoryChartData.reduce((acc, curr) => acc + curr.value, 0);
  }, [categoryChartData]);

  const CustomRechartsTooltip = ({ active, payload }: any) => {
    if (active && payload && payload.length) {
      const data = payload[0].payload;
      return (
        <div className="bg-slate-900 border border-slate-700 p-3 rounded-xl shadow-xl max-w-xs text-left">
          <p className="text-[11px] font-black text-indigo-300 uppercase tracking-wider mb-1 leading-normal break-words font-sans">
            {data.fullName || data.name}
          </p>
          <div className="space-y-1 font-mono text-[10px] text-zinc-300 leading-tight">
            <p>Total: <strong className="text-white">RM {data.value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></p>
            {chartTotalCost > 0 && (
              <p>Contribution: <strong className="text-emerald-400 font-bold">{((data.value / chartTotalCost) * 100).toFixed(1)}%</strong></p>
            )}
            {data.unitPrice !== undefined && <p>Unit Price: RM {data.unitPrice.toFixed(2)}</p>}
            {data.quantity !== undefined && <p>Quantity: {data.quantity} pcs</p>}
            {data.count !== undefined && <p>Unique Lines: {data.count}</p>}
          </div>
        </div>
      );
    }
    return null;
  };

  const stepsList = [
    "Reading and parsing uploaded Chinese quotation sheet...",
    "Extracting product layout coordinates and embedded drawings...",
    "Querying Gemini 3.5-flash AI engine for precise spatial matching...",
    "Translating Chinese names, colors, and raw materials into business English...",
    "Converting Chinese Yuan (CNY) into local Malaysian Ringgit (MYR)...",
    "Auditing product quantities, pricing formulas, and completing checks..."
  ];

  // 1. Generate Quote Number and fetch rate on mount
  useEffect(() => {
    const randomized = Math.floor(1000 + Math.random() * 9000);
    const yr = new Date().getFullYear();
    setQuotationNumber(`MOCOF-QT-${yr}-${randomized}`);
    getLiveExchangeRate();
  }, []);

  const getLiveExchangeRate = async () => {
    setFetchingRate(true);
    try {
      const res = await fetch("/api/exchange-rate");
      if (res.ok) {
        const data = await res.json();
        if (data && data.rate) {
          setLiveRate(data.rate);
          if (!useManualRate) {
            setManualExchangeRate(data.rate);
          }
        }
      }
    } catch (e) {
      console.error("Failed to query live rate endpoint, using standard fallback", e);
    } finally {
      setFetchingRate(false);
    }
  };

  // Keep manualExchangeRate input synchronized with live rate if checkbox is off
  useEffect(() => {
    if (!useManualRate) {
      setManualExchangeRate(liveRate);
    }
  }, [liveRate, useManualRate]);

  // 2. Drag and Drop file listeners
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processSelectedFile(files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processSelectedFile(files[0]);
    }
  };

  const processSelectedFile = (file: File) => {
    const fileExt = file.name.split(".").pop()?.toLowerCase();
    if (fileExt !== "xlsx" && fileExt !== "xls") {
      alert("Invalid file format. Please upload a spreadsheet with .xlsx or .xls extension.");
      return;
    }

    setSelectedFile(file);
    setConvertError(null);

    const reader = new FileReader();
    reader.onload = () => {
      const resultStr = reader.result as string;
      const base64 = resultStr.split(",")[1];
      setExcelBase64(base64);
    };
    reader.readAsDataURL(file);
  };

  // Reset converter state
  const handleReset = () => {
    setSelectedFile(null);
    setExcelBase64("");
    setQuotation(null);
    setVerificationConfirmed(false);
    setConvertError(null);
    setWorkspaceActiveTab("quotation");
    setSelectedRawSheetIndex(0);
    setRawSheetSearch("");
  };

  // 3. Trigger Quotation Transformation REST call
  const handleConvertQuotation = async () => {
    if (!excelBase64) {
      alert("Please upload a Chinese supplier raw quotation Excel file first!");
      return;
    }

    setIsConverting(true);
    setConvertError(null);
    setConvertingStep(0);
    setActiveStepText(stepsList[0]);

    // Setup visual intervals to mimic process pipeline blocks
    const progressTimer = setInterval(() => {
      setConvertingStep((prev) => {
        if (prev < stepsList.length - 1) {
          const nextStep = prev + 1;
          setActiveStepText(stepsList[nextStep]);
          return nextStep;
        }
        return prev;
      });
    }, 3800);

    try {
      const response = await fetch("/api/convert-quotation", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          excelBase64,
          manualRate: manualExchangeRate,
          useManualRate: useManualRate,
          customInstruction: preCurationInstruction
        })
      });

      clearInterval(progressTimer);

      let resData;
      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("application/json")) {
        try {
          resData = await response.json();
        } catch (jsonErr) {
          throw new Error("Failed to parse supplier quotation conversion response. Gemini might be experiencing high load.");
        }
      } else {
        const text = await response.text();
        if (text.includes("<!doctype html>") || text.includes("<html")) {
          throw new Error(`AI Conversion service returned HTML response (Status ${response.status}). The server might be restarting or overloaded. Please try again shortly.`);
        }
        throw new Error(text || `AI Conversion service failed with Status ${response.status}`);
      }

      if (!response.ok) {
        throw new Error(resData?.error || "AI Conversion service encountered a runtime failure.");
      }

      // Recalculate local subtotal details
      const rate = useManualRate ? manualExchangeRate : resData.exchangeRateUsed;
      const items = resData.items;

      const subtotal = items.reduce((sum: number, it: any) => sum + it.totalPriceMYR, 0);
      const subTotalRounded = Math.round(subtotal * 100) / 100;
      const sstAmountRounded = Math.round((subTotalRounded * (sstPercentage / 100)) * 100) / 100;
      const grandTotalRounded = Math.round((subTotalRounded + sstAmountRounded) * 100) / 100;

      // Map Quotation Struct
      const completeQuotation: QuotationData = {
        items,
        customerName: customerName || "MOCOF Valued Client",
        companyName: "MOCOF Sdn Bhd",
        preparedBy: preparedBy || "MOCOF Supply Team",
        quotationNumber: quotationNumber || "MOCOF-QT-GEN",
        date: resData.date || new Date().toISOString().split("T")[0],
        exchangeRate: rate,
        useManualRate,
        manualExchangeRate: useManualRate ? manualExchangeRate : undefined,
        sstPercentage,
        alerts: resData.alerts || [],
        termsAndConditions: resData.defaultTerms || [],
        subtotalMYR: subTotalRounded,
        sstAmountMYR: sstAmountRounded,
        grandTotalMYR: grandTotalRounded,
        rawSheets: resData.rawSheets
      };

      setQuotation(completeQuotation);
    } catch (e: any) {
      console.error(e);
      setConvertError(e?.message || "An unexpected error occurred during model inference parsing.");
    } finally {
      setIsConverting(false);
    }
  };

  // 4. Live spreadsheet mutation edits listeners
  const handleCellEdit = (
    itemId: string,
    field: keyof QuotationItem,
    value: string | number
  ) => {
    if (!quotation) return;

    const rate = useManualRate ? manualExchangeRate : quotation.exchangeRate;

    const updatedItems = quotation.items.map((item) => {
      if (item.id === itemId) {
        const updated = { ...item, [field]: value };

        // Auto mathematically synchronize fields based on specific edit types
        if (field === "quantity" || field === "unitPriceCNY") {
          const qty = Number(updated.quantity) || 0;
          const uPriceCNY = Number(updated.unitPriceCNY) || 0;
          updated.totalPriceCNY = Math.round((qty * uPriceCNY) * 100) / 100;

          // Compute MYR conversions
          updated.unitPriceMYR = Math.round((uPriceCNY * rate) * 100) / 100;
          updated.totalPriceMYR = Math.round((updated.totalPriceCNY * rate) * 100) / 100;
        } else if (field === "unitPriceMYR") {
          const qty = Number(updated.quantity) || 0;
          const uPriceValue = Number(value) || 0;
          updated.totalPriceMYR = Math.round((qty * uPriceValue) * 100) / 100;

          // Backport CNY roughly for coherence
          updated.unitPriceCNY = Math.round((uPriceValue / rate) * 100) / 100;
          updated.totalPriceCNY = Math.round((updated.totalPriceMYR / rate) * 100) / 100;
        }

        return updated;
      }
      return item;
    });

    // Recalculate totals
    const subtotal = updatedItems.reduce((sum, it) => sum + it.totalPriceMYR, 0);
    const subTotalRounded = Math.round(subtotal * 100) / 100;
    const sstAmountRounded = Math.round((subTotalRounded * (sstPercentage / 100)) * 100) / 100;
    const grandTotalRounded = Math.round((subTotalRounded + sstAmountRounded) * 100) / 100;

    setQuotation({
      ...quotation,
      items: updatedItems,
      subtotalMYR: subTotalRounded,
      sstAmountMYR: sstAmountRounded,
      grandTotalMYR: grandTotalRounded
    });
  };

  // Adjust SST tax rates with live updating
  const handleSstChange = (newSstPct: number) => {
    setSstPercentage(newSstPct);

    if (!quotation) return;

    const sstAmountRounded = Math.round((quotation.subtotalMYR * (newSstPct / 100)) * 100) / 100;
    const grandTotalRounded = Math.round((quotation.subtotalMYR + sstAmountRounded) * 100) / 100;

    setQuotation({
      ...quotation,
      sstPercentage: newSstPct,
      sstAmountMYR: sstAmountRounded,
      grandTotalMYR: grandTotalRounded
    });
  };

  // Delete product row
  const handleDeleteItem = (itemId: string) => {
    if (!quotation) return;

    const filtered = quotation.items.filter(it => it.id !== itemId);

    const subtotal = filtered.reduce((sum, it) => sum + it.totalPriceMYR, 0);
    const subTotalRounded = Math.round(subtotal * 100) / 100;
    const sstAmountRounded = Math.round((subTotalRounded * (sstPercentage / 100)) * 100) / 100;
    const grandTotalRounded = Math.round((subTotalRounded + sstAmountRounded) * 100) / 100;

    setQuotation({
      ...quotation,
      items: filtered,
      subtotalMYR: subTotalRounded,
      sstAmountMYR: sstAmountRounded,
      grandTotalMYR: grandTotalRounded
    });
  };

  // Add custom manual item to table
  const handleAddCustomItem = () => {
    if (!quotation || !newItemName.trim()) {
      alert("Please provide at least a Product Name to manually insert a row.");
      return;
    }

    const rate = useManualRate ? manualExchangeRate : quotation.exchangeRate;
    const qty = Number(newItemQty) || 1;
    const uPriceCNY = Number(newItemPriceCNY) || 0;
    
    const totalPriceCNY = qty * uPriceCNY;

    const unitPriceMYR = Math.round((uPriceCNY * rate) * 100) / 100;
    const totalPriceMYR = Math.round((totalPriceCNY * rate) * 100) / 100;

    const appendedItem: QuotationItem = {
      id: `manual_${Date.now()}`,
      originalName: newItemName,
      originalSpecs: newItemSpecs,
      originalColor: "",
      originalMaterial: "",
      originalDescription: "",
      
      translatedName: newItemName,
      translatedSpecs: newItemSpecs,
      translatedColor: "",
      translatedMaterial: "",
      translatedDescription: "",
      
      quantity: qty,
      unitPriceCNY: uPriceCNY,
      totalPriceCNY: totalPriceCNY,
      unitPriceMYR,
      totalPriceMYR,
      image: "",
      remarks: "",
      itemCode: "",
      modelNum: ""
    };

    const updatedList = [...quotation.items, appendedItem];

    // Compute Quote metrics
    const subtotal = updatedList.reduce((sum, it) => sum + it.totalPriceMYR, 0);
    const subTotalRounded = Math.round(subtotal * 100) / 100;
    const sstAmountRounded = Math.round((subTotalRounded * (sstPercentage / 100)) * 100) / 100;
    const grandTotalRounded = Math.round((subTotalRounded + sstAmountRounded) * 100) / 100;

    // Check duplicate or add fresh alert if item has zero price
    const updatedAlerts = [...quotation.alerts];
    if (uPriceCNY <= 0) {
      updatedAlerts.push({
        id: `manual_alert_${appendedItem.id}`,
        type: "warning",
        message: `Manually added item "${newItemName}" is initialized with a zero price.`,
        itemId: appendedItem.id
      });
    }

    setQuotation({
      ...quotation,
      items: updatedList,
      alerts: updatedAlerts,
      subtotalMYR: subTotalRounded,
      sstAmountMYR: sstAmountRounded,
      grandTotalMYR: grandTotalRounded
    });

    // Clear inserts inputs
    setNewItemName("");
    setNewItemPriceCNY("");
    setNewItemQty("1");
    setNewItemSpecs("");
  };

  // 4b. AI-Assisted smart curation of quotation worksheet items using Gemini
  const handleCurateQuotation = async (explicitInstruction?: string) => {
    if (!quotation) return;
    const finalInstruction = explicitInstruction || curationInstruction;
    if (!finalInstruction.trim()) {
      alert("Please enter a curation instruction first.");
      return;
    }

    setIsCurating(true);
    setAiChangeSummary(null);
    setCurationError(null);

    // Append user request to conversational history thread
    const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setChatHistory((prev) => [
      ...prev,
      {
        role: "user",
        content: finalInstruction,
        timestamp: timeStr
      }
    ]);

    if (!explicitInstruction) {
      setCurationInstruction(""); // Reset active input field on manual submit
    }

    try {
      const rate = useManualRate ? manualExchangeRate : quotation.exchangeRate;
      const res = await fetch("/api/curate-quotation", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          items: quotation.items,
          instruction: finalInstruction,
          exchangeRate: rate
        })
      });

      let data;
      const contentType = res.headers.get("content-type");
      if (contentType && contentType.includes("application/json")) {
        try {
          data = await res.json();
        } catch (jsonErr) {
          throw new Error("Failed to parse curation response. Gemini might be experiencing high load.");
        }
      } else {
        const text = await res.text();
        if (text.includes("<!doctype html>") || text.includes("<html")) {
          throw new Error(`Curation service returned HTML response (Status ${res.status}). The server might be restarting or overloaded. Please try again.`);
        }
        throw new Error(text || `Curation service failed with Status ${res.status}`);
      }

      if (!res.ok) {
        throw new Error(data?.error || "Failed to curate quotation items.");
      }
      if (data && data.items) {
        const updatedItems = data.items;
        const subtotal = updatedItems.reduce((sum: number, it: any) => sum + it.totalPriceMYR, 0);
        const subTotalRounded = Math.round(subtotal * 100) / 100;
        const sstAmountRounded = Math.round((subTotalRounded * (sstPercentage / 100)) * 100) / 100;
        const grandTotalRounded = Math.round((subTotalRounded + sstAmountRounded) * 100) / 100;

        // Perform Verification & Auditing Checks on current database items list
        const updatedAlerts: VerificationAlert[] = [];
        const nameMap = new Map<string, string>();

        updatedItems.forEach((item: any) => {
          // Check invalid quantity
          if (item.quantity <= 0) {
            updatedAlerts.push({
              id: `alert_qty_${item.id}`,
              type: "warning",
              message: `Item "${item.translatedName || item.originalName}" has a missing or invalid quantity.`,
              itemId: item.id
            });
          }
          // Check zero price
          if (item.unitPriceCNY <= 0) {
            updatedAlerts.push({
              id: `alert_val_${item.id}`,
              type: "warning",
              message: `Item "${item.translatedName || item.originalName}" has a unit price of zero CNY.`,
              itemId: item.id
            });
          }
          // Duplicate items check
          const uniqueKey = `${(item.translatedName || "").toLowerCase().trim()}_spec_${(item.translatedSpecs || "").toLowerCase().trim()}`;
          if (nameMap.has(uniqueKey)) {
            updatedAlerts.push({
              id: `alert_dup_${item.id}`,
              type: "info",
              message: `Item "${item.translatedName}" with equivalent specifications appears multiple times.`,
              itemId: item.id
            });
          } else {
            nameMap.set(uniqueKey, item.id);
          }
          // Missing picture status indicator
          if (!item.image) {
            updatedAlerts.push({
              id: `alert_img_${item.id}`,
              type: "info",
              message: `No drawing or picture detected for "${item.translatedName}".`,
              itemId: item.id
            });
          }
        });

        setQuotation({
          ...quotation,
          items: updatedItems,
          alerts: updatedAlerts,
          subtotalMYR: subTotalRounded,
          sstAmountMYR: sstAmountRounded,
          grandTotalMYR: grandTotalRounded
        });

        // Add assistant reply message
        setChatHistory((prev) => [
          ...prev,
          {
            role: "assistant",
            content: data.changeDescription || "Quotation details recalculated and successfully updated.",
            timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
          }
        ]);

        setAiChangeSummary(data.changeDescription);
      }
    } catch (e: any) {
      console.error(e);
      setCurationError(e.message || "An error occurred with the curation assistant.");
      
      // Add error response to chat
      setChatHistory((prev) => [
        ...prev,
        {
          role: "assistant",
          content: `⚠️ Error processing requesting curation: ${e.message || "Model timeout. Please retry."}`,
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        }
      ]);
    } finally {
      setIsCurating(false);
    }
  };

  // Image upload listener to replace thumbnails on cell
  const handleCellImageReplacement = (itemId: string, event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0 && quotation) {
      const reader = new FileReader();
      reader.onload = () => {
        const updatedList = quotation.items.map((item) => {
          if (item.id === itemId) {
            return { ...item, image: reader.result as string };
          }
          return item;
        });

        // Remove the "missing image" alerts for this item if resolved
        const filteredAlerts = quotation.alerts.filter(
          (alert) => !(alert.itemId === itemId && alert.id.includes("img"))
        );

        setQuotation({
          ...quotation,
          items: updatedList,
          alerts: filteredAlerts
        });
      };
      reader.readAsDataURL(files[0]);
    }
  };

  // Modify manual terms line
  const handleTermEdit = (index: number, newTermText: string) => {
    if (!quotation) return;
    const updatedTerms = [...quotation.termsAndConditions];
    updatedTerms[index] = newTermText;
    setQuotation({
      ...quotation,
      termsAndConditions: updatedTerms
    });
  };

  // 5. Handlers for output generation
  const handleDownloadExcel = () => {
    if (!quotation) return;
    exportToExcel(quotation);
  };

  const handleDownloadPDF = () => {
    if (!quotation) return;
    exportToPDF(quotation);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 antialiased selection:bg-indigo-600 selection:text-white">
      {/* Upper Navigation Rail */}
      <header className="sticky top-0 z-40 border-b border-slate-200 bg-white py-4 shadow-sm backdrop-blur-md">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-lg flex items-center justify-center text-white shadow-md shadow-indigo-100 shrink-0">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight text-slate-900 sm:text-xl">
                MOCOF <span className="text-indigo-600">AI Quoter</span>
              </h1>
              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest leading-3 sm:block">
                CNY to MYR Furniture Intelligence
              </p>
            </div>
          </div>
          
          <div className="flex items-center gap-4 sm:gap-6">
            <div className="hidden sm:flex flex-col items-end">
              <span className="text-[9px] uppercase tracking-widest text-slate-400 font-bold">Live Market Rate</span>
              <span className="text-sm font-mono font-bold text-emerald-600">1.00 CNY = {liveRate.toFixed(4)} MYR</span>
            </div>
            <div className="hidden sm:block h-8 w-[1px] bg-slate-200"></div>
            
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center gap-1.5 rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-semibold text-emerald-800 border border-emerald-200">
                <span className="size-1.5 rounded-full bg-emerald-500 animate-pulse" />
                Live Engine
              </span>
              <div className="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center border border-slate-200 text-xs font-bold text-slate-600">
                MC
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Primary Application Layout */}
      <main className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
        
        {/* Banner Section */}
        <div className="mb-8 flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-200 pb-6 gap-4 animate-fade-in">
          <div>
            <h1 className="font-display text-3xl font-extrabold tracking-tight text-slate-900 sm:text-4xl">
              AI Chinese Quotation Converter
            </h1>
            <p className="mt-2 text-sm text-slate-500 max-w-2xl">
              Instantly transform Chinese supplier quotation sheets into professional English proposals formatted in Malaysian Ringgit (MYR). Optimized for immediate MOCOF customer review.
            </p>
          </div>
          
          {quotation && (
            <button
              onClick={handleReset}
              className="inline-flex items-center justify-center gap-2 rounded-full border border-slate-200 bg-white px-5 py-2 text-sm font-bold text-slate-700 shadow-sm hover:bg-slate-50 cursor-pointer transition-all active:scale-95"
            >
              <ArrowLeft className="size-4 text-indigo-600" />
              Upload New Template
            </button>
          )}
        </div>

        <AnimatePresence mode="wait">
          {/* STATE A: SETTING AND CONVERT (UPLOAD GRID) */}
          {!quotation && !isConverting && (
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 gap-8 lg:grid-cols-12"
            >
              
              {/* Left Column: Form Settings */}
              <div className="lg:col-span-4 space-y-6">
                
                {/* 1. Quote Configuration Card */}
                <div className="rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600"></div>
                  <div className="mb-4 flex items-center gap-2.5">
                    <span className="flex size-8 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600">
                      <Settings className="size-4" />
                    </span>
                    <h2 className="font-display font-bold text-slate-900 text-base">Quotation Options</h2>
                  </div>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 flex items-center gap-1">
                        <User className="size-3 text-indigo-500" /> Customer Name
                      </label>
                      <input
                        type="text"
                        value={customerName}
                        onChange={(e) => setCustomerName(e.target.value)}
                        placeholder="e.g. MOCOF Showroom Client"
                        className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-sm focus:bg-white focus:border-indigo-500 focus:outline-none transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 flex items-center gap-1">
                        <Hash className="size-3 text-indigo-500" /> Quote Reference Number
                      </label>
                      <input
                        type="text"
                        value={quotationNumber}
                        onChange={(e) => setQuotationNumber(e.target.value)}
                        placeholder="e.g. MOCOF-QT-2026-003"
                        className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-sm focus:bg-white focus:border-indigo-500 focus:outline-none transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 flex items-center gap-1">
                        Prepared By
                      </label>
                      <input
                        type="text"
                        value={preparedBy}
                        onChange={(e) => setPreparedBy(e.target.value)}
                        className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-sm focus:bg-white focus:border-indigo-500 focus:outline-none transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 flex items-center gap-1">
                        <Percent className="size-3 text-indigo-500" /> SST Service Tax (%)
                      </label>
                      <select
                        value={sstPercentage}
                        onChange={(e) => setSstPercentage(Number(e.target.value))}
                        className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-sm focus:bg-white focus:border-indigo-500 focus:outline-none transition-all cursor-pointer"
                      >
                        <option value={0}>0% - Exempt / Tax Free</option>
                        <option value={6}>6% - Standard Malaysia SST (Services)</option>
                        <option value={10}>10% - Sales Tax (B2C Goods)</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* 2. Live Currency Hub Card */}
                <div className="rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600"></div>
                  <div className="mb-4 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <span className="flex size-8 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600">
                        <ArrowRightLeft className="size-4" />
                      </span>
                      <h2 className="font-display font-bold text-slate-900 text-base">Exchange Rate</h2>
                    </div>
                    <button
                      type="button"
                      onClick={getLiveExchangeRate}
                      disabled={fetchingRate}
                      className="rounded-full p-1.5 text-indigo-400 hover:bg-indigo-50 hover:text-indigo-600 disabled:opacity-50 transition-all cursor-pointer"
                    >
                      <RefreshCw className={`size-4 ${fetchingRate ? "animate-spin" : ""}`} />
                    </button>
                  </div>

                  <div className="space-y-4">
                    {/* Live exchange display */}
                    <div className="rounded-2xl bg-indigo-50/50 p-4 border border-indigo-100">
                      <div className="text-[10px] text-indigo-500/80 font-bold uppercase tracking-wider mb-0.5">Live Bank Rate</div>
                      <div className="font-mono text-xl font-extrabold text-indigo-950 flex items-baseline gap-1.5">
                        <span>1.00 CNY = {liveRate.toFixed(4)} MYR</span>
                      </div>
                      <div className="text-[10px] text-indigo-500 mt-1 flex items-center gap-1.5 font-medium">
                        <span className="inline-block size-1.5 rounded-full bg-emerald-500 animate-pulse" />
                        Integrated Live Exchange Feed
                      </div>
                    </div>

                    {/* Manual override option */}
                    <div className="border-t border-slate-100 pt-3">
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={useManualRate}
                          onChange={(e) => setUseManualRate(e.target.checked)}
                          className="size-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                        />
                        <span className="text-xs font-semibold text-slate-600">Apply custom override</span>
                      </label>

                      {useManualRate && (
                        <div className="mt-2.5">
                          <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Set Ratio (CNY to MYR)</label>
                          <div className="relative rounded-xl shadow-sm">
                            <input
                              type="number"
                              step="0.0001"
                              value={manualExchangeRate}
                              onChange={(e) => setManualExchangeRate(Number(e.target.value))}
                              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 pl-10 text-sm focus:bg-white focus:border-indigo-500 focus:outline-none focus:ring-0"
                            />
                            <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400 text-xs font-mono font-bold">
                              RM
                            </div>
                          </div>
                          <p className="text-[9px] text-amber-600 mt-1 font-semibold leading-tight">* Custom override rate will apply to all calculations.</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

              </div>

              {/* Right Column: Upload Box */}
              <div className="lg:col-span-8 flex flex-col justify-between">
                <div
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  className={`relative flex flex-col items-center justify-center min-h-[380px] rounded-[24px] border-2 border-dashed p-10 text-center transition-all ${
                    isDragOver
                      ? "border-indigo-600 bg-indigo-50/20"
                      : "border-slate-200 bg-white hover:border-indigo-300 hover:bg-slate-50/50"
                  }`}
                >
                  <input
                    type="file"
                    id="excel-file-uploader"
                    accept=".xlsx, .xls"
                    onChange={handleFileChange}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                  />

                  <div className="absolute top-0 left-0 w-full h-1.5 bg-indigo-600" />

                  {/* Icon Block */}
                  <div className="w-24 h-24 bg-indigo-50 rounded-full flex items-center justify-center mb-6 text-indigo-600 shrink-0 shadow-inner">
                    <Upload className="size-10 text-indigo-600" />
                  </div>

                  {selectedFile ? (
                    <div>
                      <h3 className="font-display font-extrabold text-slate-950 text-xl tracking-tight">
                        {selectedFile.name}
                      </h3>
                      <p className="text-xs text-indigo-600 font-mono font-semibold mt-1 bg-indigo-50 px-3 py-1 rounded-full inline-block">
                        {(selectedFile.size / 1024).toFixed(1)} KB • Microsoft Excel Spreadsheet
                      </p>

                      <div className="mt-8 flex justify-center gap-3">
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedFile(null);
                            setExcelBase64("");
                          }}
                          className="px-6 py-2.5 rounded-full border border-slate-200 bg-white text-xs font-bold text-slate-500 hover:bg-slate-100 cursor-pointer transition-all active:scale-95 shadow-sm"
                        >
                          Clear Selection
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <h2 className="font-display font-extrabold text-slate-900 text-2xl tracking-tight">
                        Drop Chinese Quotation Here
                      </h2>
                      <p className="mt-2 text-sm text-slate-500 max-w-md mx-auto leading-relaxed">
                        Drag and drop your supplier Excel file here, or click to browse files. Our system will automatically translate terms, pair currency rates, and clean records.
                      </p>
                      
                      <div className="mt-6 inline-flex items-center gap-1.5 rounded-full bg-slate-100 px-4 py-1.5 text-xs font-bold text-slate-600 shadow-inner">
                        <Layers className="size-3.5 text-indigo-600" /> Smart Column Autodetect Configured
                      </div>
                    </div>
                  )}

                  {/* In-place checks beneath */}
                  <div className="mt-10 flex flex-wrap justify-center gap-8 border-t border-slate-100 pt-6 w-full max-w-xl">
                    <div className="flex items-center gap-2">
                      <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center">
                        <svg className="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7"/></svg>
                      </div>
                      <span className="text-xs font-bold text-slate-600">AI Translation</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center">
                        <svg className="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7"/></svg>
                      </div>
                      <span className="text-xs font-bold text-slate-600">MYR Conversion</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center">
                        <svg className="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7"/></svg>
                      </div>
                      <span className="text-xs font-bold text-slate-600">Data Scrubbing</span>
                    </div>
                  </div>
                </div>

                {selectedFile && (
                  <div className="mt-6 p-5 rounded-[24px] border border-slate-205 bg-slate-50/50 space-y-3.5 text-left transition-all">
                    <div className="flex items-center gap-2 text-xs font-extrabold text-slate-800 uppercase tracking-wider">
                      <Sparkles className="size-4 text-indigo-600 animate-pulse animate-duration-1000" />
                      <span>Extraction Guidance (Optional)</span>
                    </div>
                    
                    {/* Gemini-Style Chat Capsule Pill */}
                    <div className="w-full bg-[#131314] rounded-full border border-neutral-800/80 pl-4 pr-2.5 py-1.5 flex items-center justify-between shadow-xl transition-all hover:border-neutral-700/90 focus-within:border-indigo-500 focus-within:ring-1 focus-within:ring-indigo-500/20">
                      {/* Left + Icon */}
                      <button type="button" className="text-neutral-400 hover:text-white p-1 hover:bg-neutral-800 rounded-full transition-all flex items-center justify-center">
                        <Plus className="size-4" />
                      </button>

                      {/* Text Input */}
                      <input
                        type="text"
                        value={preCurationInstruction}
                        onChange={(e) => setPreCurationInstruction(e.target.value)}
                        placeholder="e.g. 'Ignore factory codes', 'Translate all materials carefully', 'Apply 15% luxury buffer'..."
                        className="flex-1 bg-transparent border-none outline-none focus:outline-none text-[11px] text-white placeholder-neutral-500 px-3 py-2 font-sans"
                      />

                      {/* Right elements */}
                      <div className="flex items-center gap-3">
                        <Mic className="size-4 text-neutral-400 hover:text-white transition-all cursor-pointer" />
                        
                        {/* Interactive Waveform / Badge */}
                        <div className="w-8 h-8 rounded-full bg-white flex items-center justify-center shrink-0 shadow-md">
                          <div className="flex items-center gap-[2px]">
                            <span className="w-[2px] h-2 bg-neutral-950 rounded-full" />
                            <span className="w-[2px] h-3.5 bg-neutral-950 rounded-full" />
                            <span className="w-[2px] h-2.5 bg-neutral-950 rounded-full" />
                          </div>
                        </div>
                      </div>
                    </div>
                    <p className="text-[10px] text-slate-400 font-bold tracking-tight">💡 This instruction will be natively integrated into Gemini's translation guidelines model during extraction.</p>
                  </div>
                )}

                {/* Conversion Trigger Block */}
                <div className="mt-6 flex flex-col sm:flex-row items-center justify-between gap-4 bg-white border border-slate-200 p-5 rounded-[24px] shadow-sm relative overflow-hidden">
                  <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-indigo-600" />
                  <div className="text-left pl-2">
                    <span className="text-[9px] text-slate-400 font-bold block uppercase tracking-wider">Target Output Spec</span>
                    <span className="text-sm text-slate-700 font-bold">Client-Ready MYR Proposal • Sanitized Dimensions & Terms</span>
                  </div>

                  <button
                    type="button"
                    onClick={handleConvertQuotation}
                    disabled={!excelBase64}
                    className={`w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-full px-10 py-4 text-sm font-bold text-white shadow-lg transition-all cursor-pointer active:scale-95 ${
                      excelBase64
                        ? "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-100"
                        : "bg-slate-200 text-slate-400 shadow-none cursor-not-allowed"
                    }`}
                  >
                    Convert Quotation
                    <ChevronRight className="size-4" />
                  </button>
                </div>

                {convertError && (
                  <div className="mt-4 rounded-xl border border-rose-200 bg-rose-50/50 p-4 flex gap-3 text-left">
                    <AlertTriangle className="size-5 text-rose-600 shrink-0 mt-0.5" />
                    <div>
                      <h4 className="font-semibold text-rose-800 text-sm">Conversion Failed</h4>
                      <p className="text-xs text-rose-700 mt-1">{convertError}</p>
                    </div>
                  </div>
                )}

              </div>
            </motion.div>
          )}

          {/* STATE B: PIPELINE RUNNING TRANSITION */}
          {isConverting && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="flex flex-col items-center justify-center py-20 text-center max-w-xl mx-auto"
            >
              {/* Spinner animation */}
              <div className="relative mb-8">
                <div className="size-16 rounded-full border-4 border-slate-100 border-t-indigo-600 animate-spin" />
                <div className="absolute inset-0 flex items-center justify-center">
                  <Sparkles className="size-5 text-indigo-600 animate-pulse" />
                </div>
              </div>

              <h2 className="font-display font-extrabold text-slate-900 text-xl tracking-tight">
                Converting Quotation Sheet...
              </h2>
              
              <div className="mt-4 w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
                <motion.div
                  className="bg-indigo-600 h-full rounded-full"
                  initial={{ width: "0%" }}
                  animate={{ width: `${((convertingStep + 1) / stepsList.length) * 100}%` }}
                  transition={{ duration: 0.4 }}
                />
              </div>

              {/* Progress step text */}
              <div className="mt-6 rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm w-full text-left relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600" />
                <div className="text-[10px] text-indigo-500 font-extrabold uppercase tracking-wider mb-2">
                  Completed Phase {convertingStep + 1} of {stepsList.length}
                </div>
                <div className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <CheckCircle2 className="size-4 text-indigo-600 shrink-0" />
                  {activeStepText}
                </div>
              </div>

              <p className="text-xs text-slate-400 mt-8 leading-relaxed">
                Typically takes 10 to 15 seconds. Please do not close or reload this browser tab while our model translates the sheets.
              </p>
            </motion.div>
          )}

          {/* STATE C: RESULTS PREVIEW & VERIFICATION HUB */}
          {quotation && (
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-8"
            >
              
              {/* Info summary header strip */}
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 p-6 bg-white border border-slate-200 rounded-[24px] shadow-sm relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-1.5 bg-indigo-600" />
                <div className="p-1">
                  <span className="text-[10px] text-slate-600 font-black uppercase tracking-widest block leading-3">Customer</span>
                  <span className="text-sm font-extrabold text-indigo-950 mt-2 inline-block truncate max-w-full">
                    {quotation.customerName}
                  </span>
                </div>
                <div className="p-1 border-l border-slate-105 pl-4">
                  <span className="text-[10px] text-slate-600 font-black uppercase tracking-widest block leading-3">Quotation No</span>
                  <span className="text-sm font-mono font-extrabold text-indigo-950 mt-2 inline-block">
                    {quotation.quotationNumber}
                  </span>
                </div>
                <div className="p-1 border-l border-slate-105 pl-4">
                  <span className="text-[10px] text-slate-600 font-black uppercase tracking-widest block leading-3">Exchange Rate</span>
                  <span className="text-sm font-bold text-indigo-950 mt-2 inline-block">
                    1 CNY = RM {quotation.exchangeRate.toFixed(4)}
                  </span>
                </div>
                <div className="p-1 border-l border-slate-105 pl-4">
                  <span className="text-[10px] text-slate-600 font-black uppercase tracking-widest block leading-3">Date Converted</span>
                  <span className="text-sm font-bold text-indigo-950 mt-2 inline-block">
                    {quotation.date}
                  </span>
                </div>
              </div>

              {/* Premium Choice Workspace Sub-Tabs */}
              <div className="flex flex-wrap p-1.5 bg-slate-100 rounded-[20px] max-w-4xl gap-1.5 border border-slate-200">
                <button
                  type="button"
                  onClick={() => setWorkspaceActiveTab("quotation")}
                  className={`flex-1 min-w-[200px] flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${
                    workspaceActiveTab === "quotation"
                      ? "bg-white text-indigo-950 shadow-sm border border-slate-200"
                      : "text-slate-600 hover:text-indigo-650 hover:bg-slate-50/50 cursor-pointer"
                  }`}
                >
                  <FileText className="size-4 text-indigo-600" />
                  <span>1. Translated Quotation</span>
                </button>
                <button
                  type="button"
                  onClick={() => setWorkspaceActiveTab("rawExcel")}
                  className={`flex-1 min-w-[220px] flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${
                    workspaceActiveTab === "rawExcel"
                      ? "bg-white text-indigo-950 shadow-sm border border-slate-200"
                      : "text-slate-600 hover:text-indigo-650 hover:bg-slate-50/50 cursor-pointer"
                  }`}
                >
                  <Eye className="size-4 text-cyan-600" />
                  <span>2. Original Chinese Excel</span>
                  {quotation.rawSheets && (
                    <span className="text-[9px] bg-slate-200 text-slate-700 px-1.5 py-0.5 rounded-full font-extrabold font-mono shrink-0">
                      {quotation.rawSheets.length} sheet{quotation.rawSheets.length > 1 ? 's' : ''}
                    </span>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setWorkspaceActiveTab("insights")}
                  className={`flex-1 min-w-[200px] flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${
                    workspaceActiveTab === "insights"
                      ? "bg-white text-indigo-950 shadow-sm border border-slate-200"
                      : "text-slate-600 hover:text-indigo-650 hover:bg-slate-50/50 cursor-pointer"
                  }`}
                >
                  <TrendingUp className="size-4 text-violet-600" />
                  <span>3. Cost Charts Analytics</span>
                </button>
              </div>

              {/* TAB CONTAINER 1: TRANSLATED QUOTATION (VERIFICATION DESK) */}
              {workspaceActiveTab === "quotation" && (
                <>
                  {/* Accuracy audit feedback indicators (Verification Desk) */}
                  <div className="rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600" />
                <div className="mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex items-center gap-3">
                    <span className="flex size-9 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600">
                      <ShieldAlert className="size-5" />
                    </span>
                    <div>
                      <h3 className="font-display font-bold text-slate-900 text-base leading-tight">Quotation Verification Engine</h3>
                      <p className="text-xs text-slate-600 font-medium mt-0.5">Warnings and diagnostics flagged during extraction</p>
                    </div>
                  </div>

                  {quotation.alerts.length > 0 && (
                    <span className="rounded-full bg-amber-50 px-3 py-1 text-xs font-bold text-amber-700 border border-amber-200 shadow-inner">
                      {quotation.alerts.length} Audit Notices
                    </span>
                  )}
                </div>

                {quotation.alerts.length > 0 ? (
                  <div className="space-y-2 max-h-[220px] overflow-y-auto mb-6 pr-2">
                    {quotation.alerts.map((alert) => (
                      <div
                        key={alert.id}
                        className={`rounded-xl border p-3.5 flex gap-3 text-left items-start transition-all ${
                          alert.type === "error"
                            ? "border-rose-100 bg-rose-50/40 text-rose-800"
                            : alert.type === "warning"
                            ? "border-amber-100 bg-amber-50/30 text-amber-900"
                            : "border-indigo-100 bg-indigo-50/20 text-indigo-900"
                        }`}
                      >
                        {alert.type === "error" || alert.type === "warning" ? (
                          <AlertTriangle className={`size-4 shrink-0 mt-0.5 ${alert.type === "error" ? "text-rose-600" : "text-amber-600"}`} />
                        ) : (
                          <Info className="size-4 shrink-0 mt-0.5 text-indigo-600" />
                        )}
                        <span className="text-xs font-medium">{alert.message}</span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="mb-6 rounded-2xl border border-emerald-100 bg-emerald-50/20 p-5 text-center">
                    <div className="flex justify-center mb-1.5">
                      <CheckCircle2 className="size-6 text-emerald-600" />
                    </div>
                    <h4 className="font-semibold text-emerald-800 text-sm">Perfect Score Validation Passed</h4>
                    <p className="text-xs text-emerald-700 mt-1 max-w-lg mx-auto">Mathematical counts match, no empty price points, and all embedded images aligned perfectly.</p>
                  </div>
                )}

                {/* Confirm before export block */}
                <div className="border-t border-slate-100 pt-5 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                  <label className="flex items-start gap-2.5 cursor-pointer select-none max-w-xl">
                    <input
                      type="checkbox"
                      checked={verificationConfirmed}
                      onChange={(e) => setVerificationConfirmed(e.target.checked)}
                      className="size-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 mt-0.5 cursor-pointer"
                    />
                    <div className="text-left">
                      <span className="text-xs font-bold text-slate-800 block">I confirm this quotation is ready for client dispatch</span>
                      <span className="text-[10px] text-slate-400">Verifies that translations match specifications and pricing structures. Required to export PDF/Excel.</span>
                    </div>
                  </label>

                  <div className="flex gap-2">
                    <button
                      type="button"
                      disabled={!verificationConfirmed}
                      onClick={handleDownloadExcel}
                      className={`inline-flex items-center gap-1.5 rounded-full px-6 py-2.5 text-xs font-bold transition-all shadow-sm ${
                        verificationConfirmed
                          ? "bg-slate-100 text-slate-800 hover:bg-slate-200 hover:text-slate-900 cursor-pointer active:scale-95"
                          : "bg-slate-50 text-slate-300 border border-slate-150 cursor-not-allowed"
                      }`}
                    >
                      <FileSpreadsheet className="size-3.5" />
                      Export Excel
                    </button>

                    <button
                      type="button"
                      disabled={!verificationConfirmed}
                      onClick={handleDownloadPDF}
                      className={`inline-flex items-center gap-1.5 rounded-full px-7 py-2.5 text-xs font-bold tracking-tight text-white transition-all shadow-md active:scale-95 ${
                        verificationConfirmed
                          ? "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-100 cursor-pointer"
                          : "bg-slate-200 text-slate-400 shadow-none cursor-not-allowed"
                      }`}
                    >
                      <Download className="size-3.5" />
                      Export A4 PDF
                    </button>
                  </div>
                </div>
                
                {!verificationConfirmed && (
                  <p className="text-[10px] text-amber-600 text-right mt-2 font-semibold">
                    * Check client confirmation box to unlock export features.
                  </p>
                )}
              </div>
                </>
              )}

              {/* Cost Analytical Visualisation Dashboard */}
              {workspaceActiveTab === "insights" && (
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                {/* Category breakdown (Donut Pie Chart) */}
                <div className="lg:col-span-5 rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm relative overflow-hidden flex flex-col justify-between">
                  <div className="absolute top-0 left-0 w-full h-1 bg-cyan-500" />
                  <div>
                    <div className="flex items-center justify-between mb-4">
                      <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest">
                        Category Expense Distribution
                      </h4>
                      <span className="text-[10px] bg-cyan-50 text-cyan-700 px-2.5 py-0.5 rounded-full font-bold border border-cyan-100">
                        MYR Cost
                      </span>
                    </div>
                    
                    {categoryChartData.length === 0 ? (
                      <div className="h-[220px] flex flex-col items-center justify-center text-slate-400 text-xs">
                        No product costs available for category analysis.
                      </div>
                    ) : (
                      <div className="relative h-[220px] flex items-center justify-center">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={categoryChartData}
                              cx="50%"
                              cy="50%"
                              innerRadius={60}
                              outerRadius={85}
                              paddingAngle={3}
                              dataKey="value"
                            >
                              {categoryChartData.map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                              ))}
                            </Pie>
                            <RechartsTooltip content={<CustomRechartsTooltip />} />
                          </PieChart>
                        </ResponsiveContainer>
                        {/* Centered Total Indicator */}
                        <div className="absolute flex flex-col items-center justify-center pointer-events-none">
                          <span className="text-[9px] text-slate-400 font-extrabold uppercase tracking-wider">Total Value</span>
                          <span className="text-sm font-black text-slate-900 font-mono">
                            RM {chartTotalCost.toLocaleString("en-US", { maximumFractionDigits: 0 })}
                          </span>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* List of categories with dynamic percentages */}
                  {categoryChartData.length > 0 && (
                    <div className="mt-4 pt-4 border-t border-slate-100 space-y-1.5 max-h-[140px] overflow-y-auto pr-1">
                      {categoryChartData.map((cat, i) => {
                        const pct = chartTotalCost > 0 ? (cat.value / chartTotalCost) * 100 : 0;
                        return (
                          <div key={cat.name} className="flex justify-between items-center text-xs">
                            <div className="flex items-center gap-2 font-semibold text-slate-800">
                              <span
                                className="size-2 rounded-full shrink-0"
                                style={{ backgroundColor: CHART_COLORS[i % CHART_COLORS.length] }}
                              />
                              <span className="truncate max-w-[170px]">{cat.name}</span>
                            </div>
                            <div className="flex items-center gap-2 font-mono text-[11px]">
                              <span className="font-extrabold text-slate-900">
                                RM {cat.value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </span>
                              <span className="text-slate-400 text-[10px] font-bold">
                                ({pct.toFixed(0)}%)
                              </span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Individual item comparison (Bar Chart) */}
                <div className="lg:col-span-7 rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm relative overflow-hidden flex flex-col justify-between">
                  <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600" />
                  <div>
                    <div className="flex items-center justify-between mb-4">
                      <div>
                        <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest">
                          Item Cost Ranking (Top 8)
                        </h4>
                        <p className="text-[10px] text-slate-500 font-medium mt-0.5">Top quotation lines sorted by maximum total MYR price</p>
                      </div>
                      <span className="text-[10px] bg-indigo-50 text-indigo-700 px-2.5 py-0.5 rounded-full font-bold border border-indigo-100">
                        Itemized Compare
                      </span>
                    </div>

                    {itemChartData.length === 0 ? (
                      <div className="h-[220px] flex flex-col items-center justify-center text-slate-400 text-xs">
                        No product cost lines to rank.
                      </div>
                    ) : (
                      <div className="h-[230px] mt-2">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart
                            data={itemChartData}
                            margin={{ top: 10, right: 10, left: -25, bottom: 5 }}
                          >
                            <XAxis
                              dataKey="shortName"
                              tick={{ fill: '#475569', fontSize: 10, fontWeight: 700 }}
                              axisLine={false}
                              tickLine={false}
                            />
                            <YAxis
                              tick={{ fill: '#475569', fontSize: 9, fontWeight: 600 }}
                              axisLine={false}
                              tickLine={false}
                            />
                            <RechartsTooltip content={<CustomRechartsTooltip />} cursor={{ fill: 'rgba(79, 70, 229, 0.05)' }} />
                            <Bar
                              dataKey="value"
                              radius={[6, 6, 0, 0]}
                              maxBarSize={45}
                            >
                              {itemChartData.map((entry, index) => (
                                <Cell key={`bar-cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                  </div>

                  {/* Visual Insight Summary strip */}
                  {itemChartData.length > 0 && (
                    <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-2 gap-4">
                      <div className="bg-indigo-50/20 p-2.5 rounded-xl border border-indigo-100/50">
                        <span className="text-[8px] text-indigo-800 uppercase tracking-wider block font-bold leading-tight">Highest Cost Line</span>
                        <span className="text-[11px] font-black text-slate-900 truncate block mt-1" title={itemChartData[0]?.fullName}>
                          {itemChartData[0]?.fullName || "N/A"}
                        </span>
                      </div>
                      <div className="bg-emerald-50/20 p-2.5 rounded-xl border border-emerald-100/50">
                        <span className="text-[8px] text-emerald-800 uppercase tracking-wider block font-bold leading-tight">Max Category Share</span>
                        <span className="text-[11px] font-black text-slate-900 truncate block mt-1">
                          {categoryChartData[0]?.name || "N/A"} ({((categoryChartData[0]?.value / chartTotalCost) * 100 || 0).toFixed(0)}%)
                        </span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
              )}

              {/* TAB CONTAINER 2: ORIGINAL EXCEL SHEET INSPECTOR */}
              {workspaceActiveTab === "rawExcel" && (
                <div className="rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm overflow-hidden relative">
                  <div className="absolute top-0 left-0 w-full h-1 bg-cyan-500" />
                  
                  {/* Title and stats bar */}
                  <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-5">
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="font-display font-extrabold text-slate-900 text-lg tracking-tight">Original Chinese Excel Inspector</h3>
                        <span className="text-[10px] bg-cyan-100 text-cyan-800 border border-cyan-200 px-2.5 py-0.5 rounded-full font-extrabold flex items-center gap-1">
                          <Eye className="size-3" /> Source File
                        </span>
                      </div>
                      <p className="text-xs text-slate-500 font-medium mt-1">
                        Row-by-row structure parsed natively from your uploaded sheet. Review original cells to trace missing details or prepare curation instructions.
                      </p>
                    </div>

                    {/* Quick help notice */}
                    <div className="bg-slate-50 border border-slate-200/60 rounded-xl p-3 max-w-sm text-left">
                      <p className="text-[10px] font-bold text-slate-705 leading-tight flex items-center gap-1">
                        <span className="inline-block w-2 h-2 rounded-full bg-cyan-500 animate-pulse" /> 
                        💡 Visual Curation Guide
                      </p>
                      <ul className="list-disc pl-3 text-[9px] text-slate-500 mt-1 space-y-0.5 font-medium">
                        <li><span className="font-bold text-emerald-600">Highlighted rows</span> were successfully extracted into your quotation.</li>
                        <li>Non-highlighted rows represent headers, metadata, or skipped lines.</li>
                        <li>To manually edit cells, use the editor below or type requests in the <strong className="text-zinc-650">Curation Desk</strong>.</li>
                      </ul>
                    </div>
                  </div>

                  {/* Sheet switcher & search controllers */}
                  <div className="flex flex-col sm:flex-row gap-4 items-center justify-between mb-5 bg-slate-50/50 p-4 rounded-2xl border border-slate-150">
                    {/* Switcher pills */}
                    <div className="flex flex-wrap gap-2.5 items-center w-full sm:w-auto">
                      <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Worksheets:</span>
                      {quotation.rawSheets?.map((sheet, index) => (
                        <button
                          key={sheet.sheetName}
                          type="button"
                          onClick={() => setSelectedRawSheetIndex(index)}
                          className={`px-3.5 py-1.5 rounded-full text-xs font-bold transition-all ${
                            selectedRawSheetIndex === index
                              ? "bg-cyan-600 text-white shadow-md border border-cyan-600"
                              : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-50 cursor-pointer"
                          }`}
                        >
                          {sheet.sheetName}
                        </button>
                      ))}
                    </div>

                    {/* Live Search filter */}
                    <div className="relative w-full sm:w-72">
                      <input
                        type="text"
                        placeholder="Search cells (e.g. 'sofa', '餐桌')..."
                        value={rawSheetSearch}
                        onChange={(e) => setRawSheetSearch(e.target.value)}
                        className="w-full pl-8 pr-3 py-2 text-xs rounded-full border border-slate-200 bg-white placeholder-slate-450 font-medium text-slate-800 focus:outline-none focus:border-cyan-500 focus:ring-1 focus:ring-cyan-150 transition-all shadow-sm"
                      />
                      <Search className="absolute left-3 top-2.5 size-3.5 text-slate-400 font-bold" />
                      {rawSheetSearch && (
                        <button
                          type="button"
                          onClick={() => setRawSheetSearch("")}
                          className="absolute right-3 top-2.5 text-[10px] text-slate-400 hover:text-slate-600 font-bold"
                        >
                          Clear
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Grid Table Workspace */}
                  {quotation.rawSheets?.[selectedRawSheetIndex] ? (
                    (() => {
                      const activeSheet = quotation.rawSheets[selectedRawSheetIndex];
                      const maxColIndex = Math.max(
                        ...activeSheet.rows.flatMap(row => Object.keys(row.cells).map(Number)),
                        6
                      );
                      const filteredRows = activeSheet.rows.filter(row => {
                        if (!rawSheetSearch.trim()) return true;
                        const searchLower = rawSheetSearch.toLowerCase();
                        return Object.values(row.cells).some(val => 
                          String(val).toLowerCase().includes(searchLower)
                        );
                      });

                      if (filteredRows.length === 0) {
                        return (
                          <div className="text-center py-12 bg-slate-50 border border-dashed border-slate-200 rounded-2xl">
                            <Layers className="size-8 mx-auto text-slate-300 mb-2" />
                            <p className="text-xs font-bold text-slate-600">No matching spreadsheet rows found</p>
                            <p className="text-[10px] text-slate-400 mt-1">Try adjusting your search query or switching worksheets.</p>
                          </div>
                        );
                      }

                      return (
                        <div className="overflow-x-auto -mx-6 border-t border-slate-100">
                          <div className="inline-block min-w-full align-middle px-6">
                            <table className="min-w-full text-xs font-sans border-collapse border border-slate-350">
                              <thead>
                                <tr className="bg-slate-100 text-[10px] font-black uppercase tracking-wider text-slate-550 border-b border-slate-350">
                                  <th className="py-2.5 px-3 text-center w-16 bg-slate-50 border-r border-b border-slate-300 font-mono font-bold text-slate-400">Row</th>
                                  <th className="py-2.5 px-3 text-left w-44 bg-slate-50 border-r border-b border-slate-300 text-slate-600">Sync Status</th>
                                  {Array.from({ length: maxColIndex }).map((_, colIdx) => (
                                    <th key={colIdx} className="py-2.5 px-3 text-left min-w-[120px] font-mono font-bold text-slate-550 border-r border-b border-slate-300">
                                      {getColumnLetter(colIdx + 1)}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-300 bg-white">
                                {filteredRows.map((row) => {
                                  // Locate item by matching ExcelRowIndex (remember that row.rowIndex is 0-indexed on backend)
                                  const matchedItem = quotation.items.find(
                                    item => item.sheetName === activeSheet.sheetName && item.excelRowIndex === (row.rowIndex + 1)
                                  );
                                  return (
                                    <tr
                                      key={row.rowIndex}
                                      className="hover:bg-slate-50/40 transition-colors"
                                    >
                                      {/* Row Index */}
                                      <td className="py-2 px-3 text-center border-r border-b border-slate-300 font-mono font-bold text-slate-400 bg-slate-50/50 text-[10px]">
                                        {row.rowIndex + 1}
                                      </td>

                                      {/* Matched Details cell */}
                                      <td className="py-2 px-3 border-r border-b border-slate-300 leading-snug bg-slate-50/30">
                                        {matchedItem ? (
                                          <div className="flex flex-col gap-0.5">
                                            <span className="inline-flex items-center gap-1 rounded bg-emerald-100 px-1.5 py-0.5 text-[9px] font-black text-emerald-800 w-fit">
                                              <span className="size-1 rounded-full bg-emerald-550 shrink-0" />
                                              Synchronized
                                            </span>
                                            <span className="text-[10px] font-black text-slate-900 truncate max-w-[150px] leading-tight mt-0.5 block" title={matchedItem.translatedName}>
                                              {matchedItem.translatedName}
                                            </span>
                                            <div className="text-[9px] font-mono text-indigo-700 font-black mt-0.5">
                                              RM {matchedItem.totalPriceMYR.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                            </div>
                                          </div>
                                        ) : (
                                          <span className="inline-flex items-center gap-1 rounded bg-slate-100 px-1.5 py-0.5 text-[9px] font-bold text-slate-405 w-fit">
                                            Not Extracted
                                          </span>
                                        )}
                                      </td>

                                      {/* Cells content iteration with Style + Image rendering support */}
                                      {Array.from({ length: maxColIndex }).map((_, colIdx) => {
                                        const colNum = colIdx + 1;
                                        const cellVal = row.cells[colNum] || "";
                                        const cellStyleData = row.cellStyles?.[colNum];
                                        const cellImg = activeSheet.images?.find(
                                          (img) => img.row === row.rowIndex && img.col === colIdx
                                        );
                                        const isCellContainingChinese = /[\u4e00-\u9fa5]/.test(cellVal);
                                        
                                        // Inline Style building for excel replicates
                                        const styleObj: React.CSSProperties = {};
                                        if (cellStyleData?.bg) {
                                          styleObj.backgroundColor = cellStyleData.bg;
                                        }
                                        if (cellStyleData?.color) {
                                          styleObj.color = cellStyleData.color;
                                        }
                                        if (cellStyleData?.bold) {
                                          styleObj.fontWeight = "bold";
                                        }
                                        if (cellStyleData?.align) {
                                          const alignStr = cellStyleData.align.toLowerCase();
                                          if (alignStr === "center" || alignStr === "right" || alignStr === "left") {
                                            styleObj.textAlign = alignStr as any;
                                          }
                                        }

                                        // Inverse contrast option for very dark background colors like classic dark navy blue sheets
                                        const isDarkBg = cellStyleData?.bg && (
                                          cellStyleData.bg.toLowerCase().includes("#00206") || 
                                          cellStyleData.bg.toLowerCase().includes("#1f497") ||
                                          cellStyleData.bg.toLowerCase().includes("#00000") ||
                                          cellStyleData.bg.toLowerCase().includes("#1f4e7") ||
                                          cellStyleData.bg.toLowerCase() === "#0b0c10"
                                        );
                                        if (isDarkBg && !cellStyleData?.color) {
                                          styleObj.color = "#ffffff";
                                        }

                                        return (
                                          <td
                                            key={colIdx}
                                            style={styleObj}
                                            className={`py-2 px-3 max-w-xs break-words border-r border-b border-slate-350 hover:bg-slate-50/10 font-semibold text-[11px] vertical-align-middle ${
                                              matchedItem && !cellStyleData?.bg ? "bg-emerald-50/10" : ""
                                            } ${
                                              cellStyleData?.bg ? "" : "text-slate-700 bg-white"
                                            }`}
                                          >
                                            <div className="flex flex-col gap-1 justify-center min-h-[1.5rem]">
                                              {/* Embedded Spreadsheet Image render */}
                                              {cellImg && (
                                                <div className="flex justify-center my-1 bg-white p-1 rounded border border-slate-205 shadow-sm max-w-[120px] mx-auto group">
                                                  <img
                                                    src={cellImg.base64}
                                                    alt={`Row ${row.rowIndex + 1} Col ${getColumnLetter(colNum)}`}
                                                    referrerPolicy="no-referrer"
                                                    className="max-h-24 object-contain rounded hover:scale-110 cursor-zoom-in transition-transform duration-200"
                                                    onClick={() => {
                                                      setActivePreviewImage(cellImg.base64);
                                                    }}
                                                  />
                                                </div>
                                              )}

                                              {cellVal ? (
                                                <div className="space-y-0.5 whitespace-pre-wrap leading-normal">
                                                  <span>{cellVal}</span>
                                                  {isCellContainingChinese && !cellStyleData?.bg && (
                                                    <span className="block text-[8px] text-indigo-500 font-bold italic tracking-tight">
                                                      Chinese detail
                                                    </span>
                                                  )}
                                                </div>
                                              ) : (
                                                !cellImg && (
                                                  <span className="text-slate-300 font-mono select-none">-</span>
                                                )
                                              )}
                                            </div>
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );
                    })()
                  ) : (
                    <div className="text-center py-12 text-slate-400 text-xs">
                      No spreadsheet data loaded inside raw sheets container.
                    </div>
                  )}
                </div>
              )}

              {/* TAB CONTAINER 3: QUOTATION TABLE WORKSPACE & CURATION DESK */}
              {workspaceActiveTab === "quotation" && (
                <div className="rounded-[24px] border border-slate-200 bg-white p-6 shadow-sm overflow-hidden relative">
                <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600" />
                <div className="mb-5 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
                  <div>
                    <h3 className="font-display font-extrabold text-slate-900 text-lg tracking-tight">Client Quotation Worksheet</h3>
                    <p className="text-xs text-slate-705 font-bold mt-0.5">Edit names, specifications, colors, quantity and RM prices directly inside cells</p>
                  </div>

                  <span className="rounded-full bg-indigo-50 px-3 py-1 text-xs font-bold text-indigo-700 font-mono border border-indigo-100">
                    Total Items: {quotation.items.length}
                  </span>
                </div>

                {/* Table wrapper */}
                <div className="overflow-x-auto -mx-6">
                  <div className="inline-block min-w-full align-middle px-6">
                    <table className="min-w-full divide-y divide-slate-200">
                      <thead>
                        <tr className="bg-slate-105 text-slate-800 text-[10px] font-black uppercase tracking-widest border-y border-slate-200">
                          <th className="py-3.5 px-3 text-center w-14 text-slate-700 cursor-help" title="Click any row number below to instantly load it into the Curation Desk input bar!">No 🖱️</th>
                          <th className="py-3.5 px-3 text-center w-24 text-slate-700">Image</th>
                          <th className="py-3.5 px-3 text-left text-slate-700">Product / Code Details</th>
                          <th className="py-3.5 px-3 text-left text-slate-700">Specs & Materials</th>
                          <th className="py-3.5 px-3 text-center w-16 text-slate-700">Qty</th>
                          <th className="py-3.5 px-3 text-right w-40 text-slate-700">Unit Price (MYR / CNY)</th>
                          <th className="py-3.5 px-3 text-right w-36 text-indigo-950 font-black">Total (MYR)</th>
                          <th className="py-3.5 px-3 text-center w-12 text-slate-700">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 bg-white text-xs">
                        {quotation.items.map((item, index) => (
                          <tr key={item.id} className="hover:bg-indigo-50/15 transition-colors border-b border-slate-100">
                            {/* index column */}
                            <td 
                              onClick={() => {
                                setCurationInstruction(`correct row ${index + 1}: `);
                                const inputEl = document.querySelector('input[placeholder="Ask anything or request sheet corrections..."]') as HTMLInputElement;
                                if (inputEl) {
                                  inputEl.focus();
                                  inputEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                }
                              }}
                              className="p-3 text-center text-slate-800 font-black font-mono cursor-pointer hover:bg-indigo-600 hover:text-white rounded-xl transition-all duration-150 active:scale-95 group border border-transparent hover:border-indigo-500 shadow-sm"
                              title="Click to automatically load this row into the AI Curation template bar"
                            >
                              <span className="underline decoration-indigo-400 decoration-2 underline-offset-2 group-hover:no-underline group-hover:text-white">
                                {index + 1}
                              </span>
                            </td>

                            {/* image slot column */}
                            <td className="p-3 text-center">
                              <div className="relative group size-16 mx-auto rounded-xl border border-slate-200 bg-slate-50 overflow-hidden flex items-center justify-center shadow-xs">
                                {item.image ? (
                                  <>
                                    <img
                                      src={item.image}
                                      alt="thumbnail"
                                      className="size-full object-cover transition-transform duration-300 group-hover:scale-105"
                                    />
                                    {/* Action overlays */}
                                    <div className="absolute inset-0 bg-slate-950/60 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-1.5">
                                      <button
                                        type="button"
                                        onClick={() => setActivePreviewImage(item.image)}
                                        className="rounded-full bg-white p-1 text-slate-700 hover:text-black focus:outline-none shadow-sm cursor-pointer transition-transform hover:scale-110"
                                      >
                                        <Eye className="size-3.5" />
                                      </button>
                                      <label className="rounded-full bg-white p-1 text-slate-700 hover:text-black focus:outline-none cursor-pointer shadow-sm transition-transform hover:scale-110">
                                        <Camera className="size-3.5" />
                                        <input
                                          type="file"
                                          accept="image/*"
                                          className="hidden"
                                          onChange={(e) => handleCellImageReplacement(item.id, e)}
                                        />
                                      </label>
                                    </div>
                                  </>
                                ) : (
                                  <label className="flex flex-col items-center justify-center size-full cursor-pointer hover:bg-slate-100 transition-colors">
                                    <Upload className="size-4 text-slate-400" />
                                    <span className="text-[8px] text-slate-450 font-bold mt-1 uppercase tracking-wider">Upload</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      className="hidden"
                                      onChange={(e) => handleCellImageReplacement(item.id, e)}
                                    />
                                  </label>
                                )}
                              </div>
                            </td>

                            {/* code and title product block */}
                            <td className="p-3 text-left space-y-1.5">
                              <div>
                                <label className="text-[10px] text-indigo-805 block font-extrabold uppercase leading-tight tracking-wider">English Product Title</label>
                                <input
                                  type="text"
                                  value={item.translatedName}
                                  onChange={(e) => handleCellEdit(item.id, "translatedName", e.target.value)}
                                  className="w-full text-xs font-black text-indigo-950 border-b border-indigo-200 hover:border-indigo-500 focus:border-indigo-600 focus:outline-none px-1 rounded transition-colors"
                                />
                              </div>

                              <div className="grid grid-cols-2 gap-2">
                                <div>
                                  <span className="text-[8px] text-slate-600 block font-extrabold uppercase leading-tight tracking-wider">Product Code</span>
                                  <input
                                    type="text"
                                    value={item.itemCode || ""}
                                    onChange={(e) => handleCellEdit(item.id, "itemCode", e.target.value)}
                                    placeholder="N/A"
                                    className="w-full text-xs text-slate-800 border-b border-transparent hover:border-slate-350 focus:border-indigo-400 focus:outline-none px-1 rounded transition-colors"
                                  />
                                </div>
                                <div>
                                  <span className="text-[8px] text-slate-600 block font-extrabold uppercase leading-tight tracking-wider">Supplier Code</span>
                                  <span className="text-xs text-slate-700 truncate block px-1 font-mono font-bold" title={item.originalName}>
                                    {item.originalName || "-"}
                                  </span>
                                </div>
                              </div>
                            </td>

                            {/* specs and materials details translation column */}
                            <td className="p-3 text-left space-y-1.5">
                              <div>
                                <label className="text-[10px] text-indigo-805 block font-extrabold uppercase leading-tight tracking-wider">Dimensions / Specs</label>
                                <input
                                  type="text"
                                  value={item.translatedSpecs}
                                  onChange={(e) => handleCellEdit(item.id, "translatedSpecs", e.target.value)}
                                  className="w-full text-xs text-slate-900 border-b border-indigo-200 hover:border-indigo-500 focus:border-indigo-650 focus:outline-none px-1 rounded transition-colors"
                                />
                              </div>

                              <div className="grid grid-cols-2 gap-2">
                                <div>
                                  <span className="text-[8px] text-slate-600 block font-extrabold uppercase leading-tight tracking-wider">Material</span>
                                  <input
                                    type="text"
                                    value={item.translatedMaterial}
                                    onChange={(e) => handleCellEdit(item.id, "translatedMaterial", e.target.value)}
                                    placeholder="Materials config"
                                    className="w-full text-xs text-slate-800 border-b border-transparent hover:border-slate-350 focus:border-indigo-400 focus:outline-none px-1 rounded transition-colors"
                                  />
                                </div>
                                <div>
                                  <span className="text-[8px] text-slate-600 block font-extrabold uppercase leading-tight tracking-wider">Color Finish</span>
                                  <input
                                    type="text"
                                    value={item.translatedColor}
                                    onChange={(e) => handleCellEdit(item.id, "translatedColor", e.target.value)}
                                    placeholder="Color details"
                                    className="w-full text-xs text-slate-800 border-b border-transparent hover:border-slate-350 focus:border-indigo-400 focus:outline-none px-1 rounded transition-colors"
                                  />
                                </div>
                              </div>
                            </td>

                            {/* quantity cell editor column */}
                            <td className="p-3 text-center">
                              <input
                                type="number"
                                min="1"
                                value={item.quantity}
                                onChange={(e) => handleCellEdit(item.id, "quantity", Number(e.target.value))}
                                className="w-14 text-center text-xs font-black text-slate-900 bg-slate-50 border border-slate-300 rounded-lg px-2 py-1.5 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-100 transition-all font-mono"
                              />
                            </td>

                            {/* myr/cny pricing cost editor column */}
                            <td className="p-3 text-right space-y-2">
                              <div>
                                <label className="text-[8px] text-indigo-700 block font-black uppercase text-right leading-none mb-1 tracking-wider">MYR Price</label>
                                <div className="relative rounded-lg max-w-[140px] ml-auto">
                                  <input
                                    type="number"
                                    step="0.01"
                                    value={item.unitPriceMYR}
                                    onChange={(e) => handleCellEdit(item.id, "unitPriceMYR", Number(e.target.value))}
                                    className="w-full text-right text-xs font-mono font-black text-slate-900 bg-slate-50 border border-slate-300 rounded-lg px-2 py-1.5 pl-7 hover:border-slate-400 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-150 transition-all shadow-sm"
                                  />
                                  <div className="absolute inset-y-0 left-0 pl-2.5 flex items-center pointer-events-none text-indigo-750 text-[9px] font-black">
                                    RM
                                  </div>
                                </div>
                              </div>

                              <div>
                                <label className="text-[8px] text-amber-700 block font-black uppercase text-right leading-none mb-1 tracking-wider">CNY Cost</label>
                                <div className="relative rounded-lg max-w-[140px] ml-auto">
                                  <input
                                    type="number"
                                    step="0.01"
                                    value={item.unitPriceCNY}
                                    onChange={(e) => handleCellEdit(item.id, "unitPriceCNY", Number(e.target.value))}
                                    className="w-full text-right text-xs font-mono font-bold text-amber-950 bg-amber-50/20 border border-amber-300/60 rounded-lg px-2 py-1.5 pl-7 hover:border-amber-400 focus:outline-none focus:border-amber-500 transition-all shadow-xs"
                                  />
                                  <div className="absolute inset-y-0 left-0 pl-2.5 flex items-center pointer-events-none text-amber-800 text-[9px] font-black font-mono">
                                    ¥
                                  </div>
                                </div>
                              </div>
                            </td>

                            {/* total calculated price column */}
                            <td className="p-3 text-right text-indigo-950 font-mono font-black text-sm">
                              RM {item.totalPriceMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </td>

                            {/* erase row column action */}
                            <td className="p-3 text-center">
                              <button
                                type="button"
                                onClick={() => handleDeleteItem(item.id)}
                                className="rounded-xl p-2 text-slate-300 hover:text-rose-600 hover:bg-rose-50 transition-colors cursor-pointer"
                              >
                                <Trash2 className="size-4.5" />
                              </button>
                            </td>
                          </tr>
                        ))}

                        {/* Blank checklist screen if rows deletion */}
                        {quotation.items.length === 0 && (
                          <tr>
                            <td colSpan={8} className="py-12 text-center text-slate-400 font-medium">
                              No product lines in active quotation list. Create one below.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Insertion Workspace controls */}
                <div className="mt-6 rounded-[20px] border border-dashed border-indigo-100 bg-indigo-50/20 p-5 space-y-4">
                  <span className="text-xs font-extrabold text-indigo-950 block gap-2 flex items-center uppercase tracking-wider">
                    <Plus className="size-4 text-indigo-600" /> Append Custom Product Row
                  </span>

                  <div className="grid grid-cols-1 md:grid-cols-12 gap-3">
                    <div className="md:col-span-5">
                      <input
                        type="text"
                        placeholder="Product English Title / Name"
                        value={newItemName}
                        onChange={(e) => setNewItemName(e.target.value)}
                        className="w-full text-xs rounded-full border border-slate-200 bg-white px-4 py-2.5 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-150 transition-all font-medium text-slate-800"
                      />
                    </div>
                    <div className="md:col-span-3">
                      <input
                        type="text"
                        placeholder="Specifications / Dimension Size"
                        value={newItemSpecs}
                        onChange={(e) => setNewItemSpecs(e.target.value)}
                        className="w-full text-xs rounded-full border border-slate-200 bg-white px-4 py-2.5 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-150 transition-all font-medium text-slate-800"
                      />
                    </div>
                    <div className="md:col-span-1">
                      <input
                        type="number"
                        placeholder="Qty"
                        value={newItemQty}
                        onChange={(e) => setNewItemQty(e.target.value)}
                        className="w-full text-xs text-center rounded-full border border-slate-200 bg-white px-2 py-2.5 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-150 transition-all font-bold text-slate-850 font-mono"
                      />
                    </div>
                    <div className="md:col-span-2">
                      <div className="relative">
                        <input
                          type="number"
                          placeholder="Price (CNY)"
                          value={newItemPriceCNY}
                          onChange={(e) => setNewItemPriceCNY(e.target.value)}
                          className="w-full text-xs rounded-full border border-slate-200 bg-white pl-8 pr-4 py-2.5 text-right focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-150 transition-all font-bold text-slate-850 font-mono"
                        />
                        <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400 text-[10px] font-extrabold">
                          ¥
                        </div>
                      </div>
                    </div>
                    <div className="md:col-span-1">
                      <button
                        type="button"
                        onClick={handleAddCustomItem}
                        className="w-full text-xs font-bold bg-indigo-600 text-white rounded-full px-4 py-2.5 hover:bg-indigo-700 active:scale-95 transition-all text-center cursor-pointer shadow-sm"
                      >
                        Add
                      </button>
                    </div>
                  </div>
                </div>

                {/* AI Smart Curation Assistant Workspace (Gemini-style Chat Console) */}
                <div id="ai-curation-desk" className="mt-6 rounded-[24px] border border-zinc-800 bg-[#131314] p-5 space-y-4 relative overflow-hidden shadow-2xl">
                  <div className="absolute top-0 right-0 p-4 opacity-5 pointer-events-none">
                    <Sparkles className="size-20 text-indigo-400" />
                  </div>
                  
                  {/* Workspace Header */}
                  <div className="flex items-center justify-between border-b border-zinc-900 pb-3">
                    <div className="flex items-center gap-2.5">
                      <div className="rounded-xl bg-indigo-600/10 p-2 text-indigo-400 shrink-0 border border-indigo-500/20">
                        <Sparkles className="size-4 animate-pulse" />
                      </div>
                      <div>
                        <h4 className="text-xs font-black text-white uppercase tracking-widest flex items-center gap-1.5 leading-none">
                          Gemini AI Curation Desk
                          <span className="rounded-full bg-indigo-500/10 border border-indigo-400/25 text-indigo-300 text-[9px] px-2 py-0.5 normal-case font-extrabold">
                            Live Conversation
                          </span>
                        </h4>
                        <p className="text-[10px] text-zinc-400 mt-1">
                          Directly write requests to delete lines, edit prices, insert products, or ask questions.
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* Real-time Assistant chat logs message thread */}
                  <div className="h-[210px] overflow-y-auto mb-4 border border-zinc-900/80 bg-[#0e0f10] rounded-2xl p-4 space-y-3.5 scrollbar-thin scrollbar-thumb-zinc-800">
                    {chatHistory.map((msg, index) => (
                      <div key={index} className={`flex gap-2.5 ${msg.role === "user" ? "justify-end" : "justify-start"}`}>
                        {msg.role === "assistant" && (
                          <div className="size-6 rounded-lg bg-[#18191b] border border-zinc-800 flex items-center justify-center shrink-0">
                            <Sparkles className="size-3.5 text-indigo-400" />
                          </div>
                        )}
                        <div className={`max-w-[85%] rounded-[18px] px-3.5 py-2 text-[11px] font-medium leading-relaxed shadow-sm ${
                          msg.role === "user"
                            ? "bg-indigo-600/15 text-indigo-200 border border-indigo-500/25"
                            : "bg-[#18191b] text-zinc-300 border border-zinc-800/60"
                        }`}>
                          <p className="whitespace-pre-line">{msg.content}</p>
                          <span className="text-[8px] text-zinc-500 font-bold block mt-1 tracking-wide uppercase text-right leading-none">
                            {msg.timestamp}
                          </span>
                        </div>
                      </div>
                    ))}
                    {isCurating && (
                      <div className="flex gap-2.5 justify-start">
                        <div className="size-6 rounded-lg bg-[#18191b] border border-zinc-800 flex items-center justify-center shrink-0">
                          <RefreshCw className="size-3.5 text-indigo-400 animate-spin" />
                        </div>
                        <div className="rounded-[18px] px-3.5 py-2 text-[11px] font-semibold text-zinc-400 bg-zinc-900/40 border border-dashed border-zinc-800 animate-pulse">
                          Curation model analyzing columns & rows...
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Gemini-Style Chat Capsule Pill */}
                  <div className="w-full bg-[#1e1f20] rounded-full border border-neutral-800 pl-4 pr-2.5 py-1.5 flex items-center justify-between shadow-xl transition-all hover:border-neutral-700 focus-within:border-indigo-500 focus-within:ring-1 focus-within:ring-indigo-500/20">
                    {/* Left + Icon */}
                    <button type="button" className="text-zinc-400 hover:text-white p-1 hover:bg-zinc-800 rounded-full transition-all flex items-center justify-center">
                      <Plus className="size-4" />
                    </button>

                    {/* Text Input */}
                    <input
                      type="text"
                      placeholder="Ask anything or request sheet corrections..."
                      value={curationInstruction}
                      onChange={(e) => setCurationInstruction(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          handleCurateQuotation();
                        }
                      }}
                      disabled={isCurating}
                      className="flex-1 bg-transparent border-none outline-none focus:outline-none text-[11px] text-white placeholder-zinc-500 px-3 py-2 font-sans"
                    />

                    {/* Right Elements Block */}
                    <div className="flex items-center gap-3">
                      <Mic className="size-4 text-zinc-400 hover:text-white transition-all cursor-pointer" />
                      
                      {/* Interactive Submission Waveform Button */}
                      <button
                        type="button"
                        onClick={() => handleCurateQuotation()}
                        disabled={isCurating || !curationInstruction.trim()}
                        className={`w-8 h-8 rounded-full flex items-center justify-center shrink-0 shadow-md transition-all active:scale-95 ${
                          isCurating
                            ? "bg-zinc-800 text-zinc-600 cursor-not-allowed"
                            : curationInstruction.trim()
                            ? "bg-white text-zinc-950 hover:scale-105 cursor-pointer"
                            : "bg-zinc-800 text-zinc-500 cursor-not-allowed shadow-none"
                        }`}
                      >
                        <div className="flex items-center gap-[2.5px]">
                          <span className={`w-[2.5px] h-2 bg-current rounded-full ${isCurating ? "animate-bounce" : ""}`} style={{ animationDelay: '0.1s' }} />
                          <span className={`w-[2.5px] h-3.5 bg-current rounded-full ${isCurating ? "animate-bounce" : ""}`} style={{ animationDelay: '0.3s' }} />
                          <span className={`w-[2.5px] h-2 bg-current rounded-full ${isCurating ? "animate-bounce" : ""}`} style={{ animationDelay: '0.2s' }} />
                        </div>
                      </button>
                    </div>
                  </div>

                  {/* Command Suggestion Chips Block */}
                  <div className="space-y-2 pt-1">
                    <span className="text-[9px] font-black text-indigo-400 uppercase tracking-widest block">Quick Action Presets (Click to execute):</span>
                    <div className="flex flex-wrap gap-1.5">
                      {[
                        { label: "✂️ Remove row 3", cmd: "remove row 3" },
                        { label: "➕ Add solid walnut dining table", cmd: "add a solid walnut dining table for 2800 CNY with specifications 1600*800*750mm and quantity 1" },
                        { label: "✏️ Change Row 1 Qty to 5", cmd: "change row 1 quantity to 5" },
                        { label: "📐 Update Row 2 spec to 1.8m", cmd: "update row 2 specifications to 1.8m size" },
                        { label: "💰 Decrease all prices by 5%", cmd: "please decrease all unit price CNY by 5 percent" },
                        { label: "❓ How many items are there?", cmd: "how many total furniture items did you extract?" }
                      ].map((chip, idx) => (
                        <button
                          key={idx}
                          type="button"
                          onClick={() => handleCurateQuotation(chip.cmd)}
                          className="rounded-full bg-zinc-900 border border-zinc-800/60 text-zinc-300 text-[10px] font-bold px-3 py-1.5 hover:bg-zinc-805 hover:border-zinc-700 transition-colors cursor-pointer active:scale-95 text-left leading-none"
                        >
                          {chip.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Realtime Action Error state */}
                  {curationError && (
                    <div className="rounded-xl border border-rose-950/40 bg-rose-950/20 p-3.5 text-xs text-rose-300 flex items-start gap-2.5 mt-4">
                      <AlertTriangle className="size-4.5 text-rose-500 shrink-0 mt-0.5" />
                      <div className="space-y-1">
                        <span className="font-extrabold text-rose-400 uppercase text-[9px] tracking-wider block">Curation Refused:</span>
                        <p className="font-medium text-[11px] leading-relaxed">{curationError}</p>
                      </div>
                    </div>
                  )}
                </div>

                {/* Subtotals summaries sheet table */}
                <div className="mt-8 border-t border-slate-200 pt-6 flex flex-col md:flex-row md:justify-between items-start gap-6">
                  {/* Terms editor sidebar */}
                  <div className="w-full md:max-w-md space-y-3 shrink-0">
                    <span className="text-xs font-black text-slate-900 tracking-wider uppercase block">Company Terms & Contract Notes (Editable)</span>
                    <div className="space-y-1 bg-slate-50/50 p-4 rounded-2xl border border-slate-200">
                      {quotation.termsAndConditions.map((term, i) => (
                        <div key={i} className="flex gap-2 items-center">
                          <span className="size-1.5 rounded-full bg-indigo-550 shrink-0" />
                          <input
                            type="text"
                            value={term}
                            onChange={(e) => handleTermEdit(i, e.target.value)}
                            className="w-full text-xs text-slate-850 font-semibold border-b border-transparent bg-transparent hover:border-slate-300 hover:bg-white focus:border-indigo-500 focus:bg-white focus:outline-none px-2 py-1 rounded-lg transition-all"
                          />
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Calculations breakdown list */}
                  <div className="w-full md:max-w-xs space-y-3.5 bg-slate-50/50 p-6 rounded-[24px] border border-slate-250 relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600" />
                    <div className="flex justify-between items-center text-xs">
                      <span className="text-slate-800 font-black">Quotations Subtotal:</span>
                      <span className="font-mono text-slate-955 font-extrabold text-right">
                        RM {quotation.subtotalMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </span>
                    </div>

                    <div className="flex justify-between items-center text-xs">
                      <span className="text-slate-800 font-black">Service Tax (SST {quotation.sstPercentage}%):</span>
                      <div className="flex items-center gap-1">
                        <span className="font-mono text-slate-955 font-extrabold text-right">
                          RM {quotation.sstAmountMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                      </div>
                    </div>

                    <div className="border-t border-slate-250 my-2" />

                    <div className="flex justify-between items-baseline pt-1">
                      <span className="text-xs font-black text-slate-950 uppercase tracking-wide">Grand Total:</span>
                      <span className="font-mono font-display text-lg font-black text-indigo-950 text-right">
                        RM {quotation.grandTotalMYR.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </span>
                    </div>
                  </div>

                </div>

              </div>
              )}

            </motion.div>
          )}
        </AnimatePresence>

      </main>

      {/* FOOTER METRICS INFO */}
      <footer className="border-t border-slate-200 bg-white py-8 mt-20 text-xs text-slate-400">
        <div className="mx-auto max-w-7xl px-4 text-center space-y-2.5">
          <p className="font-semibold text-slate-500">© 2026 MOCOF Luxury Furnishings. Supply Chain AI translation client utility. All rates integrated with real-time exchange banks.</p>
          <p className="text-[10px] text-slate-400">Secure sandboxed execution environment. Developed using TypeScript & Gemini-2.5-flash AI core representation.</p>
        </div>
      </footer>

      {/* POPUP LIGHTBOX FOR FULL IMAGE INSPECTION */}
      <AnimatePresence>
        {activePreviewImage && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setActivePreviewImage(null)}
            className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4 cursor-zoom-out"
          >
            <motion.div
              initial={{ scale: 0.95 }}
              animate={{ scale: 1 }}
              exit={{ scale: 0.95 }}
              onClick={(e) => e.stopPropagation()}
              className="relative max-w-2xl bg-white rounded-[24px] overflow-hidden shadow-2xl p-4"
            >
              <img
                src={activePreviewImage}
                alt="Product Spec Preview"
                className="max-h-[65vh] w-auto max-w-full rounded-[16px] object-contain"
              />
              <div className="flex items-center justify-between text-xs text-slate-400 mt-3 px-1">
                <span>MOCOF Supply Chain Inspection</span>
                <span className="font-semibold text-slate-700">Matched Product Drawing</span>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
