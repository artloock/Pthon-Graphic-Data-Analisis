import * as XLSX from "xlsx";
import type { Dataset } from "./types";

export async function importDataset(file: File): Promise<Dataset> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstName = workbook.SheetNames[0];
  if (!firstName) throw new Error("O arquivo não contém planilhas.");
  const sheet = workbook.Sheets[firstName];
  if (!sheet) throw new Error("Não foi possível ler a primeira planilha.");
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: false });
  if (matrix.length < 2) throw new Error("O arquivo precisa ter cabeçalhos e dados.");
  return {
    sheetName: firstName,
    headers: matrix[0]!.map((value, index) => String(value || `Coluna ${index + 1}`)),
    rows: matrix.slice(1) as Dataset["rows"]
  };
}
