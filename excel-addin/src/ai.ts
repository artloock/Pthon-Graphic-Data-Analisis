import { analyzeDataset } from "./analyzer";
import type { Dataset, Suggestion } from "./types";

export interface AnalysisProvider {
  readonly name: string;
  analyze(dataset: Dataset): Promise<Suggestion[]>;
}

export class LocalRulesProvider implements AnalysisProvider {
  readonly name = "Motor local de regras";
  async analyze(dataset: Dataset): Promise<Suggestion[]> {
    return analyzeDataset(dataset);
  }
}

// Ponto de extensão para um agente do Microsoft 365 Copilot. A integração exige
// tenant, autenticação Entra ID e publicação administrativa; não há API fictícia aqui.
export class CopilotProvider implements AnalysisProvider {
  readonly name = "Microsoft 365 Copilot";
  async analyze(_dataset: Dataset): Promise<Suggestion[]> {
    throw new Error("O provedor Copilot ainda não foi configurado para este tenant.");
  }
}
