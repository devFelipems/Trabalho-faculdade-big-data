# main.py
import pandas as pd
import plotly.express as px
from conectar import conectar

EXCEL_PATH = "C:/bigdata/Planejamento (1).xlsx"

# Função robusta para ler abas
def read_sheet_robust(path, sheet_name, try_headers=(0,1,2,3,4)):
    last_df = None
    for h in try_headers:
        df = pd.read_excel(path, sheet_name=sheet_name, header=h)
        last_df = df
        cols = [str(c) for c in df.columns]
        if not all(c.startswith("Unnamed") for c in cols):
            return df
    return last_df

# Ler abas
df_resumo = read_sheet_robust(EXCEL_PATH, "resumo")
df_assert = read_sheet_robust(EXCEL_PATH, "assertividade")
df_ciclo = read_sheet_robust(EXCEL_PATH, "ciclo")
df_reserva_parcial = read_sheet_robust(EXCEL_PATH, "Reserva Parcial")
df_reserva_integral = read_sheet_robust(EXCEL_PATH, "Reserva integral")
df_reserva_credito = read_sheet_robust(EXCEL_PATH, "Reservado-Credito")

# Padronizar colunas
df_resumo = df_resumo.iloc[:, :4]
df_resumo.columns = ["Divisão", "Vlr de Pedido", "Valor Faturado", "%"]

df_assert = df_assert.iloc[:, :2]
df_assert.columns = ["Data", "Assertividade"]
df_assert["Data"] = pd.to_datetime(df_assert["Data"], errors="coerce")
df_assert = df_assert.dropna(subset=["Data"])

# Usar colunas por posição para evitar erro
df_ciclo_std = pd.DataFrame({
    "Produto": df_ciclo.iloc[:, 0],
    "Liberada": pd.to_numeric(df_ciclo.iloc[:, 1], errors="coerce"),
    "Bloqueada": pd.to_numeric(df_ciclo.iloc[:, 2], errors="coerce"),
    "Produção": pd.to_numeric(df_ciclo.iloc[:, 3], errors="coerce"),
})

def guess_value_column(df):
    for col in df.columns:
        if "Vlr" in str(col) and "Pedido" in str(col):
            return col
    return df.columns[1]

val_parcial = guess_value_column(df_reserva_parcial)
val_integral = guess_value_column(df_reserva_integral)
val_credito = guess_value_column(df_reserva_credito)

df_reservas = pd.concat([
    df_reserva_parcial[[df_reserva_parcial.columns[0], val_parcial]].assign(Status="Parcial").rename(columns={df_reserva_parcial.columns[0]: "Pedido", val_parcial: "Valor"}),
    df_reserva_integral[[df_reserva_integral.columns[0], val_integral]].assign(Status="Integral").rename(columns={df_reserva_integral.columns[0]: "Pedido", val_integral: "Valor"}),
    df_reserva_credito[[df_reserva_credito.columns[0], val_credito]].assign(Status="Crédito").rename(columns={df_reserva_credito.columns[0]: "Pedido", val_credito: "Valor"}),
], ignore_index=True)

df_reservas["Valor"] = pd.to_numeric(df_reservas["Valor"], errors="coerce")
df_reservas = df_reservas.dropna(subset=["Valor"])

# Criar gráficos
fig1 = px.bar(df_resumo, x="Divisão", y=["Vlr de Pedido", "Valor Faturado"],
              barmode="group", title="Pedidos vs Faturamento por Divisão")

fig2 = px.line(df_assert, x="Data", y="Assertividade", markers=True,
               title="Evolução da Assertividade")

fig3 = px.pie(df_reservas, names="Status", values="Valor",
              title="Distribuição dos Pedidos por Status")

fig4 = px.bar(df_ciclo_std, x="Produto",
              y=["Liberada", "Bloqueada", "Produção"],
              barmode="stack", title="Status dos Pedidos por Produto")

# Mostrar gráficos (cada um abre em uma aba)
fig1.show()
fig2.show()
fig3.show()
fig4.show()

#Chamar a função conectar
cursor, conn = conectar()
cursor.execute("SELECT * FROM pedido")
pedidos = cursor.fetchall()
print(pedidos)

conn.close()