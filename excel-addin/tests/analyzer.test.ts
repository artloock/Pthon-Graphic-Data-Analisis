import { describe, expect, it } from "vitest";
import { analyzeDataset, compareDatasets, normalizedKey } from "../src/analyzer";
import type { Dataset } from "../src/types";

const dataset: Dataset = {
  sheetName: "Vendas",
  headers: [" loja ", "valor"],
  rows: [["Centro", 100], [" Centro ", 50], ["NORTE", 25], ["norte", 75]]
};

describe("analyzer", () => {
  it("normaliza acentos, caixa e espaços", () => expect(normalizedKey("  São   Paulo ")).toBe("sao paulo"));
  it("propõe correções sem alterar os dados", () => {
    const result = analyzeDataset(dataset);
    expect(result.some((item) => item.row === 0 && item.replacement === "loja")).toBe(true);
    expect(result.some((item) => item.original === " Centro " && item.replacement === "Centro")).toBe(true);
    expect(dataset.rows[1]?.[0]).toBe(" Centro ");
  });
  it("compara períodos alinhando cabeçalhos", () => {
    const previous: Dataset = { sheetName: "Antigo", headers: ["Loja", "Valor"], rows: [["Centro", 80], ["norte", 50]] };
    expect(compareDatasets(dataset, previous, 0, 1)).toEqual([
      { category: "Centro", current: 150, previous: 80, delta: 70 },
      { category: "NORTE", current: 100, previous: 50, delta: 50 }
    ]);
  });
});
