# dash_app.py
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html

EXCEL_PATH = r"C:\Trabalho big data python\Trabalho-faculdade-big-data\Planejamento (1).xlsx"

def read_sheet_robust(path, sheet_name, try_headers=(0,1,2,3,4)):
    last_df = None
    for h in try_headers:
        df = pd.read_excel(path, sheet_name=sheet_name, header=h)
        last_df = df
        cols = [str(c) for c in df.columns]
        if not all(c.startswith("Unnamed") for c in cols):
            return df
    return last_df

# Ler e tratar dados (igual ao main, sem mexer nele)
df_resumo = read_sheet_robust(EXCEL_PATH, "resumo")
df_assert = read_sheet_robust(EXCEL_PATH, "assertividade")
df_ciclo = read_sheet_robust(EXCEL_PATH, "ciclo")
df_reserva_parcial = read_sheet_robust(EXCEL_PATH, "Reserva Parcial")
df_reserva_integral = read_sheet_robust(EXCEL_PATH, "Reserva integral")
df_reserva_credito = read_sheet_robust(EXCEL_PATH, "Reservado-Credito")

df_resumo = df_resumo.iloc[:, :4]
df_resumo.columns = ["Divisão", "Vlr de Pedido", "Valor Faturado", "%"]

df_assert = df_assert.iloc[:, :2]
df_assert.columns = ["Data", "Assertividade"]
df_assert["Data"] = pd.to_datetime(df_assert["Data"], errors="coerce")
df_assert = df_assert.dropna(subset=["Data"])

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

# ========================
# GRÁFICOS
# ========================

# 1 — Pizza Pedidos vs Faturamento
df_resumo_pizza = pd.melt(df_resumo, id_vars="Divisão",
                          value_vars=["Vlr de Pedido", "Valor Faturado"],
                          var_name="Tipo", value_name="Valor")
fig1 = px.pie(df_resumo_pizza, names="Tipo", values="Valor",
              title="Pedidos vs Faturamento por Divisão (Pizza)")

# 2 — Linha Assertividade
fig2 = px.line(df_assert, x="Data", y="Assertividade", markers=True,
               title="Evolução da Assertividade")
fig2.update_layout(yaxis_tickformat=".0%")

# 3 — Pizza Reservas
fig3 = px.pie(df_reservas, names="Status", values="Valor",
              title="Distribuição dos Pedidos por Status", hole=0.35)

# 4 — Linha Status por Produto
fig4 = px.line(df_ciclo_std, x="Produto", y=["Liberada", "Bloqueada", "Produção"],
               markers=True, title="Status dos Pedidos por Produto (Linha)")

# 5 — Novo: Barras Reservas por Status
df_reservas_sum = df_reservas.groupby("Status", as_index=False)["Valor"].sum()
fig5 = px.bar(df_reservas_sum, x="Status", y="Valor", color="Status",
              title="Total de Reservas por Status (Barras)")

# ===========================
# APP DASH COM 5 ABAS
# ===========================
app = Dash(__name__)
app.layout = html.Div(
    style={"backgroundColor": "#f4f6f9", "minHeight": "100vh", "padding": "20px", "fontFamily": "Arial"},
    children=[
        html.H1("Dashboard Big Data", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "40px"}),
        dcc.Tabs(children=[
            dcc.Tab(label="Resumo", children=[dcc.Graph(figure=fig1)]),
            dcc.Tab(label="Assertividade", children=[dcc.Graph(figure=fig2)]),
            dcc.Tab(label="Reservas", children=[dcc.Graph(figure=fig3)]),
            dcc.Tab(label="Ciclo", children=[dcc.Graph(figure=fig4)]),
            dcc.Tab(label="Total de Reservas", children=[dcc.Graph(figure=fig5)]),
        ])
    ]
)

if __name__ == "__main__":
    app.run(debug=True)
