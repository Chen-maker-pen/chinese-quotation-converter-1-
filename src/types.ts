export interface QuotationItem {
  id: string;
  originalName: string;
  originalSpecs: string; // Dimensions, spec, etc.
  originalColor: string;
  originalMaterial: string;
  originalDescription: string;
  
  translatedName: string;
  translatedSpecs: string;
  translatedColor: string;
  translatedMaterial: string;
  translatedDescription: string;
  
  quantity: number;
  unitPriceCNY: number;
  totalPriceCNY: number;
  unitPriceMYR: number;
  totalPriceMYR: number;
  image: string; // base64 data url or empty string
  imageIndex?: number;
  remarks: string;
  itemCode: string;
  modelNum: string;
}

export interface VerificationAlert {
  id: string;
  type: 'info' | 'warning' | 'error';
  message: string;
  itemId?: string;
}

export interface QuotationData {
  items: QuotationItem[];
  customerName: string;
  companyName: string; // e.g. MOCOF Sdn Bhd or custom
  preparedBy: string;
  quotationNumber: string;
  date: string;
  exchangeRate: number; // 1 CNY = X MYR
  useManualRate: boolean;
  manualExchangeRate?: number;
  sstPercentage: number; // e.g. 6 or 0
  termsAndConditions: string[];
  alerts: VerificationAlert[];
  
  subtotalMYR: number;
  sstAmountMYR: number;
  grandTotalMYR: number;
  rawSheets?: Array<{
    sheetName: string;
    rows: Array<{
      rowIndex: number;
      cells: { [colIndex: number]: string };
      cellStyles?: {
        [colIndex: number]: {
          bg?: string;
          color?: string;
          bold?: boolean;
          align?: "left" | "center" | "right" | string;
        };
      };
    }>;
    images?: Array<{
      row: number;
      col: number;
      base64: string;
    }>;
  }>;
}
