export const KNOWN_STATES = ["CA", "MA", "NJ", "IL", "CT", "FL", "TX", "NY", "MD", "RI", "NH"];

export const COL = {
  A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9,
} as const;

export const COLUMN_LABELS: Record<number, string> = {
  [COL.B]: "Phone",
  [COL.D]: "Customer Name",
  [COL.E]: "Lead ID",
  [COL.F]: "Address",
  [COL.G]: "Rep",
  [COL.I]: "Email",
  [COL.J]: "Utility",
};

export const VISIBLE_COLS = [COL.B, COL.D, COL.E, COL.F, COL.G, COL.I, COL.J];

export type CellValue = string | number | null;

export interface RowData {
  cells: CellValue[];
  isRed: boolean;
  isDivider: boolean;
  originalRow: number; // 1-based source row
}

export interface SheetData {
  name: string;
  headers: string[];
  rows: RowData[];
  isStateSheet: boolean;
}

export interface WorkbookData {
  sheets: SheetData[];
  fileName: string;
}

export type MacroOp =
  | "sort-asc"
  | "sort-desc"
  | "clean-rep"
  | "clean-proper"
  | "format-phone"
  | "capital-states"
  | "dedupe"
  | "remove-outliers"
  | "move-red-top"
  | "run-all";

export interface MacroResult {
  op: MacroOp;
  changed: number;
  message: string;
}
