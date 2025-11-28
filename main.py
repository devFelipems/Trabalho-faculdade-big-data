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
import plotly.express as px

# ======================================================
#  FIG 1 — Pedidos vs Faturamento por Divisão (AGORA PIZZA)
# ======================================================

df_resumo_pizza = pd.melt(
    df_resumo,
    id_vars="Divisão",
    value_vars=["Vlr de Pedido", "Valor Faturado"],
    var_name="Tipo",
    value_name="Valor"
)
# ========================
# MELHORIAS VISUAIS GERAIS
# ========================
COLOR1 = "#4A90E2"   # Azul
COLOR2 = "#FF6F61"   # Vermelho suave
COLOR3 = "#50C878"   # Verde
COLOR4 = "#F5A623"   # Amarelo
FONT = "Arial"

fig1 = px.pie(
    df_resumo_pizza,
    names="Tipo",         # Pizza com 2 fatias
    values="Valor",       # Soma dos valores
    title="Pedidos vs Faturamento por Divisão (Pizza)",
    width=700, 
    height=700, # QUADRADO
    color="Tipo",
    color_discrete_map={
        "Vlr de Pedido": COLOR1,
        "Valor Faturado": COLOR2
    }
)
fig1.update_traces(textinfo="percent+label", pull=[0.03, 0.05])

# Gráfico 2 – Evolução da Assertividade
fig2 = px.line(
    df_assert,
    x="Data", 
    y="Assertividade",
    markers=True,
    title="Evolução da Assertividade",
    width=700, 
    height=700,  # QUADRADO
)
fig2.update_traces(line=dict(width=3, color=COLOR3))
fig2.update_layout(yaxis_tickformat=".0%")

# Gráfico 3 – Pizza (maior)
fig3 = px.pie(
    df_reservas,
    names="Status", values="Valor",
    title="Distribuição dos Pedidos por Status",
    hole=0.35,
    width=900, 
    height=900,  # MAIOR
    color="Status",
    color_discrete_map={
        "Parcial": COLOR1,
        "Integral": COLOR3,
        "Crédito": COLOR2
    }
)

# Gráfico 4 – Status por Produto (AGORA LINE)
fig4 = px.line(
    df_ciclo_std,
    x="Produto",
    y=["Liberada", "Bloqueada", "Produção"],
    markers=True,
    title="Status dos Pedidos por Produto (Linha)",
    width=700, 
    height=700,  # QUADRADO
)
fig4.update_traces(line=dict(width=2))

from plotly.subplots import make_subplots
import plotly.graph_objects as go

# === CRIAR SUBPLOTS VERTICAIS (4 linhas x 1 coluna) ===
fig_total = make_subplots(
    rows=4, cols=1,
    subplot_titles=[
        "Pedidos vs Faturamento (Pizza)",
        "Evolução da Assertividade",
        "Distribuição dos Pedidos por Status",
        "Status dos Pedidos por Produto (Linha)"
    ],
    specs=[
        [{"type": "domain"}],  # Pizza
        [{"type": "xy"}],      # Linha
        [{"type": "domain"}],  # Pizza
        [{"type": "xy"}]       # Linha
    ],
    vertical_spacing=0.1

)

# === ADICIONAR FIG1 (PIZZA) ===
for trace in fig1.data:
    fig_total.add_trace(trace, row=1, col=1)

# === ADICIONAR FIG2 (LINE) ===
for trace in fig2.data:
    fig_total.add_trace(trace, row=2, col=1)

# === ADICIONAR FIG3 (PIZZA) ===
for trace in fig3.data:
    fig_total.add_trace(trace, row=3, col=1)

# === ADICIONAR FIG4 (LINE) ===
for trace in fig4.data:
    fig_total.add_trace(trace, row=4, col=1)


# === TAMANHO TOTAL DA PÁGINA (4 gráficos empilhados) ===
fig_total.update_layout(
    height=3500,       # aumenta a altura total
    width=1100,        # largura maior, mas ajustável
    template="plotly_white",   # tema claro
    title_text="Dashboard Big Data",
    title_x=0.5,  # centraliza título
    title_font=dict(size=26),
    paper_bgcolor="#ffffff",   # fundo branco da página
    plot_bgcolor="#fafafa",    # fundo branco dos gráficos
    margin=dict(l=40, r=40, t=60, b=40),
    font=dict(family=FONT, size=16, color="#333"),
)

fig_total.update_yaxes(
    showgrid=True, gridcolor="lightgray", zeroline=False
)
fig_total.update_xaxes(
    showgrid=False, tickangle=45
)
# === MOSTRAR TUDO EM UMA ÚNICA ABA ===
fig_total.show()


#Chamar a função conectar
cursor, conn = conectar()
cursor.execute("SELECT * FROM pedido")
pedidos = cursor.fetchall()
print(pedidos)

conn.close()
