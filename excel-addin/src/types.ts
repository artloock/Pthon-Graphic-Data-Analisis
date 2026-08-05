export type CellValue = string | number | boolean | null;
export type Matrix = CellValue[][];

export interface Dataset {
  headers: string[];
  rows: CellValue[][];
  sheetName: string;
}

export interface Suggestion {
  id: string;
  row: number;
  column: number;
  header: string;
  original: CellValue;
  replacement: CellValue;
  reason: string;
  confidence: "alta" | "média";
}

export interface AppliedChange extends Suggestion {}

export interface ComparisonRow {
  category: string;
  current: number;
  previous: number;
  delta: number;
}
