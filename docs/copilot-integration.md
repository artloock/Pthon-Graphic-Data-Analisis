# Integração com Microsoft 365 Copilot

## Decisão de arquitetura

O DataAlign não depende do Copilot para operações essenciais. Leitura, validação, padronização, comparação e gráficos funcionam com regras locais reproduzíveis. Isso mantém o suplemento utilizável em ambientes sem licença avançada e evita que uma resposta probabilística altere células sem revisão.

A interface `AnalysisProvider`, em `excel-addin/src/ai.ts`, permite trocar o motor local por um provedor corporativo. O retorno sempre deve ser uma lista estruturada de sugestões; a aplicação continua sendo feita pela Office JavaScript API somente após aprovação do usuário.

## Caminho recomendado

1. Registrar uma aplicação no Microsoft Entra ID.
2. Definir permissões mínimas e o fluxo de consentimento do tenant.
3. Criar um agente declarativo do Microsoft 365 Copilot com ação baseada em API.
4. Usar APIs locais do Office para fornecer contexto da pasta de trabalho quando disponível.
5. Validar o retorno contra um schema estrito de sugestões.
6. Exibir origem, justificativa e confiança antes de escrever na planilha.
7. Publicar o suplemento e o agente pelo catálogo administrativo da organização.

## Limites atuais

- A integração de agentes do Copilot com APIs JavaScript locais do Office está documentada pela Microsoft como recurso de prévia e pode mudar.
- Uma assinatura Microsoft 365 com Copilot não cria automaticamente um endpoint genérico para o suplemento chamar.
- Autenticação, disponibilidade de APIs e custos dependem do tenant e do produto contratado.
- Dados sensíveis não devem ser enviados a um modelo sem política organizacional, consentimento e auditoria.

## Regra de segurança

O Copilot sugere; o usuário aprova; o DataAlign executa. Nenhum provedor de IA recebe autorização direta para alterar a pasta de trabalho.

## Referências oficiais

- [Visão geral dos suplementos do Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Plugins com APIs JavaScript locais do Office](https://learn.microsoft.com/microsoft-365/copilot/extensibility/build-api-plugins-local-office-api)
- [Segurança e autenticação das APIs do Copilot](https://learn.microsoft.com/microsoft-365/copilot/extensibility/copilot-apis-security-authentication)
