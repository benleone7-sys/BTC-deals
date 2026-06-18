export interface TransactionItem {
  id: string;
  machineName: string;
  phone: string;
  rawRow: Record<string, any>;
  [key: string]: any;
}

export interface FileData {
  fileName: string;
  columns: string[];
  rows: Record<string, any>[];
}

export interface ComparisonResult {
  totalFile1: number;
  totalFile2: number;
  matchedCount: number;
  filteredCount: number; // rows kept (only TLV 1 or TLV 2)
  skippedCount: number;  // rows discarded (not TLV 1/2)
  missingItems: TransactionItem[];
}

export interface ColumnMapping {
  // File A mapping
  file1IdKey: string;
  file1MachineKey: string;
  file1PhoneKey: string;
  // File B mapping
  file2SessionIdKey: string;
}
