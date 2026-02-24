import pandas as pd
import plotly.express as px
import os

# 1. Configuração de Pastas
output_dir = "relatorios_html"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 2. Carga de Dados
try:
    df = pd.read_excel('vendas.xlsx')
except FileNotFoundError:
    print("Erro: O arquivo 'vendas.xlsx' não foi encontrado!")
    exit()

# 3. Processamento de Dados
# Agrupamento para análise detalhada
df_grouped = df.groupby(["estado", "cidade", "loja", "forma_pagamento"])["preco"].sum().reset_index()
df_grouped.to_excel("GroupedData.xlsx", index=False)

# 4. Geração de Gráficos Automatizada
columns_to_analyze = ["loja", "cidade", "estado", "tamanho", "local_consumo"]

for column in columns_to_analyze:
    if column in df.columns:
        fig = px.histogram(
            df, 
            x=column,
            y="preco",
            text_auto='.2s', # Formata números grandes (ex: 1.2k)
            title=f"Análise de Faturamento por {column.capitalize()}",
            color="forma_pagamento",
            template="plotly_white" # Visual mais limpo e profissional
        )
        
        # Atualiza o layout para ficar mais elegante
        fig.update_layout(yaxis_title="Faturamento Total (R$)", xaxis_title=column.capitalize())
        
        # Guarda na pasta de relatórios
        file_path = os.path.join(output_dir, f"faturamento_{column}.html")
        fig.write_html(file_path)
        print(f"✅ Gráfico para '{column}' gerado em: {file_path}")
    else:
        print(f"⚠️ Aviso: A coluna '{column}' não existe no arquivo.")