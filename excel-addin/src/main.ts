import "./styles.css";
import { aggregate, compareDatasets } from "./analyzer";
import { LocalRulesProvider } from "./ai";
import { applyChanges, readActiveDataset, undoChanges, writeReport } from "./excel";
import { importDataset } from "./importer";
import type { Dataset, Suggestion } from "./types";

const $ = <T extends HTMLElement>(selector: string) => document.querySelector<T>(selector)!;
let current: Dataset | null = null;
let previous: Dataset | null = null;
let suggestions: Suggestion[] = [];
let lastApplied: Suggestion[] = [];
const provider = new LocalRulesProvider();

function setStatus(message: string, error = false) {
  const element = $("#status");
  element.textContent = message;
  element.classList.toggle("error", error);
}

function populateColumns(dataset: Dataset) {
  for (const selector of ["#category", "#value"]) {
    const select = $<HTMLSelectElement>(selector);
    select.innerHTML = dataset.headers.map((header, index) => `<option value="${index}">${header}</option>`).join("");
  }
  const numeric = dataset.headers.findIndex((_, index) => dataset.rows.some((row) => typeof row[index] === "number"));
  if (numeric >= 0) $("#value") .setAttribute("value", String(numeric));
  $<HTMLSelectElement>("#value").value = String(Math.max(numeric, 0));
  $<HTMLButtonElement>("#chart").disabled = false;
}

function renderSuggestions() {
  const container = $("#suggestions");
  $("#issue-count").textContent = `${suggestions.length} ajuste${suggestions.length === 1 ? "" : "s"}`;
  if (!suggestions.length) {
    container.innerHTML = `<div class="empty">Nenhuma inconsistência automática encontrada.</div>`;
    return;
  }
  container.innerHTML = suggestions.slice(0, 100).map((item) => `
    <label class="suggestion">
      <input type="checkbox" data-id="${item.id}" checked />
      <span><strong>${item.header} · linha ${item.row + 1}</strong><small>${String(item.original)} → ${String(item.replacement)}</small><em>${item.reason}</em></span>
    </label>`).join("") + (suggestions.length > 100 ? `<div class="empty">Mostrando os primeiros 100 ajustes.</div>` : "");
}

async function analyze() {
  try {
    setStatus("Lendo a planilha…");
    current = await readActiveDataset();
    suggestions = await provider.analyze(current);
    $("#sheet-name").textContent = current.sheetName;
    $("#dataset-meta").textContent = `${current.rows.length.toLocaleString("pt-BR")} linhas · ${current.headers.length} colunas`;
    $("#results").classList.remove("hidden");
    populateColumns(current);
    renderSuggestions();
    setStatus(`Análise concluída com ${provider.name}.`);
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha na análise.", true); }
}

$("#analyze").addEventListener("click", analyze);
$("#apply").addEventListener("click", async () => {
  const selected = new Set([...document.querySelectorAll<HTMLInputElement>("[data-id]:checked")].map((input) => input.dataset.id));
  lastApplied = suggestions.filter((item) => selected.has(item.id));
  if (!lastApplied.length) return setStatus("Selecione pelo menos um ajuste.", true);
  try {
    await applyChanges(lastApplied);
    $<HTMLButtonElement>("#undo").disabled = false;
    setStatus(`${lastApplied.length} ajuste(s) aplicado(s).`);
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha ao aplicar.", true); }
});

$("#undo").addEventListener("click", async () => {
  try {
    await undoChanges(lastApplied);
    $<HTMLButtonElement>("#undo").disabled = true;
    setStatus("Última aplicação desfeita.");
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha ao desfazer.", true); }
});

$("#chart").addEventListener("click", async () => {
  if (!current) return;
  try {
    const category = Number($<HTMLSelectElement>("#category").value);
    const value = Number($<HTMLSelectElement>("#value").value);
    const rows = [...aggregate(current, category, value)].map(([label, total]) => [label, total]);
    await writeReport(`Total de ${current.headers[value]} por ${current.headers[category]}`, [current.headers[category]!, current.headers[value]!], rows);
    setStatus("Relatório e gráfico criados em “DataAlign Report”.");
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha ao gerar gráfico.", true); }
});

$("#previous-file").addEventListener("change", async (event) => {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (!file) return;
  try {
    previous = await importDataset(file);
    $("#file-name").textContent = file.name;
    $<HTMLButtonElement>("#compare").disabled = !current;
    setStatus("Arquivo histórico carregado.");
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha na importação.", true); }
});

$("#compare").addEventListener("click", async () => {
  if (!current || !previous) return;
  try {
    const category = Number($<HTMLSelectElement>("#category").value);
    const value = Number($<HTMLSelectElement>("#value").value);
    const result = compareDatasets(current, previous, category, value);
    await writeReport("Comparação histórica", [current.headers[category]!, "Atual", "Anterior", "Variação"], result.map((row) => [row.category, row.current, row.previous, row.delta]));
    setStatus("Comparação criada em “DataAlign Report”.");
  } catch (error) { setStatus(error instanceof Error ? error.message : "Falha na comparação.", true); }
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) setStatus("Excel conectado. Selecione uma tabela para começar.");
});
