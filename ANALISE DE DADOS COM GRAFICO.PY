import pandas as pd
import plotly.express as px

#load Exel data
data = pd.read_excel('vendas.xlsx')

grouped_data = data.groupby(["estado", "cidade", "loja", "forma_pagamento"])["preco"].sum().to_frame()
grouped_data.to_excel("GroupedData.xlsx")

#colun lists to graphics
colun_list = ["loja", "cidade", "estado", "tamanho", "local_consumo"]

for colun in colun_list:
    if colun in data.columns:
        greaphic = px.histogram(data, x=colun,
                                y= "preco",
                                text_auto=True,
                                title="Faturamento",
                                color="forma_pagamento")
        greaphic.show()
        greaphic.write_html(f"Faturamento-{colun}")
    else:
        print(f"column {colun} is not in DataFrame")