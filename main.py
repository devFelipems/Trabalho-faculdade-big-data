# main.py
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from conectar import conectar

EXCEL_PATH = r"C:\Trabalho big data python\Trabalho-faculdade-big-data\Planejamento (1).xlsx"

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

# 
# GRÁFICOS
# 
COLOR1 = "#4A90E2"
COLOR2 = "#FF6F61"
COLOR3 = "#50C878"
COLOR4 = "#F5A623"
FONT = "Arial"

# 1 — Pizza Pedidos vs Faturamento
df_resumo_pizza = pd.melt(df_resumo, id_vars="Divisão",
                          value_vars=["Vlr de Pedido", "Valor Faturado"],
                          var_name="Tipo", value_name="Valor")
fig1 = px.pie(df_resumo_pizza, names="Tipo", values="Valor",
              title="Pedidos vs Faturamento por Divisão (Pizza)",
              color="Tipo",
              color_discrete_map={"Vlr de Pedido": COLOR1, "Valor Faturado": COLOR2})

# 2 — Linha Assertividade
fig2 = px.line(df_assert, x="Data", y="Assertividade", markers=True,
               title="Evolução da Assertividade")
fig2.update_layout(yaxis_tickformat=".0%")

# 3 — Pizza Reservas
fig3 = px.pie(df_reservas, names="Status", values="Valor",
              title="Distribuição dos Pedidos por Status", hole=0.35,
              color="Status",
              color_discrete_map={"Parcial": COLOR1, "Integral": COLOR3, "Crédito": COLOR2})

# 4 — Linha Status por Produto
fig4 = px.line(df_ciclo_std, x="Produto", y=["Liberada", "Bloqueada", "Produção"],
               markers=True, title="Status dos Pedidos por Produto (Linha)")

# 5 — Barras Reservas por Status
df_reservas_sum = df_reservas.groupby("Status", as_index=False)["Valor"].sum()
fig5 = px.bar(df_reservas_sum, x="Status", y="Valor", color="Status",
              title="Total de Reservas por Status (Barras)",
              color_discrete_map={"Parcial": COLOR1, "Integral": COLOR3, "Crédito": COLOR2})

#
# DASHBOARD COM 5 GRÁFICOS
# 
from plotly.subplots import make_subplots

fig_total = make_subplots(
    rows=5, cols=1,
    subplot_titles=[
        "Pedidos vs Faturamento (Pizza)",
        "Evolução da Assertividade",
        "Distribuição dos Pedidos por Status",
        "Status dos Pedidos por Produto (Linha)",
        "Total de Reservas por Status (Barras)"
    ],
    specs=[[{"type": "domain"}],
           [{"type": "xy"}],
           [{"type": "domain"}],
           [{"type": "xy"}],
           [{"type": "xy"}]],
    vertical_spacing=0.08
)

for trace in fig1.data: fig_total.add_trace(trace, row=1, col=1)
for trace in fig2.data: fig_total.add_trace(trace, row=2, col=1)
for trace in fig3.data: fig_total.add_trace(trace, row=3, col=1)
for trace in fig4.data: fig_total.add_trace(trace, row=4, col=1)
for trace in fig5.data: fig_total.add_trace(trace, row=5, col=1)

fig_total.update_layout(height=4200, width=1100, template="plotly_white",
                        title_text="Dashboard Big Data", title_x=0.5,
                        font=dict(family=FONT, size=16, color="#333"))

fig_total.show()

# Conexão com banco
cursor, conn = conectar()
cursor.execute("SELECT * FROM pedido")
pedidos = cursor.fetchall()
print(pedidos)
conn.close()

