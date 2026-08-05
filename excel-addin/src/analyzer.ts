import type { CellValue, Dataset, Suggestion } from "./types";

export function normalizedKey(value: CellValue): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("pt-BR");
}

export function cleanLabel(value: CellValue): string {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function canonicalLabels(values: CellValue[]): Map<string, string> {
  const grouped = new Map<string, Map<string, number>>();
  for (const value of values) {
    if (typeof value !== "string" || !value.trim()) continue;
    const clean = cleanLabel(value);
    const key = normalizedKey(clean);
    const variants = grouped.get(key) ?? new Map<string, number>();
    variants.set(clean, (variants.get(clean) ?? 0) + 1);
    grouped.set(key, variants);
  }

  const result = new Map<string, string>();
  for (const [key, variants] of grouped) {
    const winner = [...variants.entries()].sort((a, b) => b[1] - a[1] || a[0].length - b[0].length)[0];
    if (winner) result.set(key, winner[0]);
  }
  return result;
}

export function analyzeDataset(dataset: Dataset): Suggestion[] {
  const suggestions: Suggestion[] = [];

  dataset.headers.forEach((header, column) => {
    const cleanedHeader = cleanLabel(header);
    if (cleanedHeader !== header) {
      suggestions.push({
        id: `header-${column}`,
        row: 0,
        column,
        header: header || `Coluna ${column + 1}`,
        original: header,
        replacement: cleanedHeader,
        reason: "Espaços inconsistentes no cabeçalho",
        confidence: "alta"
      });
    }

    const values = dataset.rows.map((row) => row[column] ?? null);
    const canonical = canonicalLabels(values);
    values.forEach((value, rowIndex) => {
      if (typeof value !== "string" || !value.trim()) return;
      const clean = cleanLabel(value);
      const replacement = canonical.get(normalizedKey(clean)) ?? clean;
      if (replacement !== value) {
        suggestions.push({
          id: `${rowIndex + 1}-${column}`,
          row: rowIndex + 1,
          column,
          header: cleanedHeader || `Coluna ${column + 1}`,
          original: value,
          replacement,
          reason: clean !== value ? "Espaçamento padronizado" : "Variação de texto padronizada",
          confidence: "alta"
        });
      }
    });
  });

  return suggestions;
}

export function aggregate(dataset: Dataset, categoryColumn: number, valueColumn: number): Map<string, number> {
  const grouped = new Map<string, { label: string; total: number }>();
  for (const row of dataset.rows) {
    const category = cleanLabel(row[categoryColumn] ?? "Sem categoria") || "Sem categoria";
    const key = normalizedKey(category);
    const rawValue = row[valueColumn];
    const value = typeof rawValue === "number" ? rawValue : Number(String(rawValue ?? "").replace(",", "."));
    if (Number.isFinite(value)) {
      const existing = grouped.get(key);
      grouped.set(key, { label: existing?.label ?? category, total: (existing?.total ?? 0) + value });
    }
  }
  return new Map([...grouped.values()].map(({ label, total }) => [label, total]));
}

export function compareDatasets(current: Dataset, previous: Dataset, categoryColumn: number, valueColumn: number) {
  const currentTotals = aggregate(current, categoryColumn, valueColumn);
  const categoryHeader = current.headers[categoryColumn] ?? "";
  const valueHeader = current.headers[valueColumn] ?? "";
  const previousCategory = previous.headers.findIndex((h) => normalizedKey(h) === normalizedKey(categoryHeader));
  const previousValue = previous.headers.findIndex((h) => normalizedKey(h) === normalizedKey(valueHeader));
  if (previousCategory < 0 || previousValue < 0) throw new Error("O arquivo anterior não possui as colunas selecionadas.");
  const previousTotals = aggregate(previous, previousCategory, previousValue);
  const aligned = new Map<string, { category: string; current: number; previous: number }>();
  for (const [category, total] of currentTotals) aligned.set(normalizedKey(category), { category, current: total, previous: 0 });
  for (const [category, total] of previousTotals) {
    const key = normalizedKey(category);
    const row = aligned.get(key) ?? { category, current: 0, previous: 0 };
    row.previous = total;
    aligned.set(key, row);
  }
  return [...aligned.values()].sort((a, b) => a.category.localeCompare(b.category, "pt-BR")).map((row) => ({
    ...row,
    delta: row.current - row.previous
  }));
}
