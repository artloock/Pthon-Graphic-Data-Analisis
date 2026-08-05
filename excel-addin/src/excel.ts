import type { AppliedChange, Dataset } from "./types";

export async function readActiveDataset(): Promise<Dataset> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRangeOrNullObject(true);
    sheet.load("name");
    range.load(["values", "isNullObject"]);
    await context.sync();
    if (range.isNullObject || range.values.length < 2) throw new Error("A planilha precisa ter cabeçalhos e pelo menos uma linha de dados.");
    const values = range.values as unknown[][];
    return {
      sheetName: sheet.name,
      headers: values[0]!.map((value, index) => String(value || `Coluna ${index + 1}`)),
      rows: values.slice(1).map((row) => row.map((value) => (value === undefined ? null : value))) as Dataset["rows"]
    };
  });
}

export async function applyChanges(changes: AppliedChange[]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    for (const change of changes) sheet.getCell(change.row, change.column).values = [[change.replacement]];
    await context.sync();
  });
}

export async function undoChanges(changes: AppliedChange[]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    for (const change of changes) sheet.getCell(change.row, change.column).values = [[change.original]];
    await context.sync();
  });
}

export async function writeReport(title: string, headers: string[], rows: (string | number)[][]): Promise<void> {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const existing = workbook.worksheets.getItemOrNullObject("DataAlign Report");
    existing.load("isNullObject");
    await context.sync();
    if (!existing.isNullObject) existing.delete();
    const sheet = workbook.worksheets.add("DataAlign Report");
    sheet.getRange("A1").values = [[title]];
    sheet.getRangeByIndexes(2, 0, 1, headers.length).values = [headers];
    if (rows.length) sheet.getRangeByIndexes(3, 0, rows.length, headers.length).values = rows;
    const dataRange = sheet.getRangeByIndexes(2, 0, Math.max(rows.length + 1, 2), headers.length);
    dataRange.format.autofitColumns();
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.columns);
    chart.title.text = title;
    chart.setPosition("F3", "N20");
    sheet.activate();
    await context.sync();
  });
}
