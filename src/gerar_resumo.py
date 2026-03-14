import pandas as pd
from datetime import datetime


def read_base(df, caminho_do_arquivo):

    tabela_resumo = df.groupby(['filial', 'box']).agg(
        carga=("carga", "count"),
        pedidos=("pedidos", "sum"),
        cubagem=("cubagem", "sum")
    ).reset_index()

    tabela_resumo['timestamp_leitura'] = datetime.now().strftime("%d/%m/%Y %H:%M")



    ordem_colunas = ["filial", "carga", "pedidos", "cubagem", "box", "timestamp_leitura"]

    tabela_resumo = tabela_resumo[ordem_colunas]

    with pd.ExcelWriter(caminho_do_arquivo, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        tabela_resumo.to_excel(writer, sheet_name="resumo", index=False)

if __name__ == "__main__":
    read_base(df, r"C:\Users\minhaMatricula\Downloads\minhaRota\minhaRota.xlsx")
