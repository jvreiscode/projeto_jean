import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings

warnings.filterwarnings("ignore")

# tema escuro
plt.style.use("dark_background")

COR_FUNDO = "#0B132B"
COR_PAINEL = "#1C2541"
COR_TEXTO = "#FFFFFF"
COR_AZUL = "#5BC0EB"
COR_VERDE = "#00C853"
COR_VERMELHO = "#FF5252"

# caminho do excel
arquivo = r"C:\Users\joao.soares\Documents\João\PYTHON\PROJETOS\JEAN\Excel\Cidade Inteligente.xlsx"

# lendo planilha
df = pd.read_excel(
    arquivo,
    sheet_name="DESSUST",
    header=[1, 2]
)

# ajusta nomes das colunas
df.columns = [
    col[1] if "Unnamed" in str(col[0]) else col[1]
    for col in df.columns
]

# colunas principais
INFO_COLS = [
    "Cidade",
    "Estado",
    "População total estimada do município",
    "PIB per capita do município",
    "Índice de desenvolvimento humano do município (IDH-M)"
]

# categorias do projeto
CATEGORIAS = {
    "ENERGIA": [
        "Soluções inteligentes para gestão do consumo de energia elétrica",
        "Soluções para telegestão da iluminação pública"
    ],
    "QUALIDADE DO AR": [
        "Soluções em monitoramento de gases de efeito estufa e qualidade do ar",
        "Monitoramento da qualidade do ar"
    ],
    "ESTRATÉGIA": [
        "Governança colaborativa",
        "Governança tecnológica",
        "Planejamento",
        "Seguimento de políticas públicas municipais",
        "Visão e conceito de cidade"
    ],
    "CULTURA": [
        "Estrutura de equipamentos culturais e esportivos",
        "Proteção do patrimônio cultural material e imaterial",
        "Serviços on-line para promoção de cultura",
        "Serviços culturais on-line oferecidos para a população"
    ],
    "SAÚDE": [
        "Serviços de telemedicina ou telessaúde",
        "Leitos hospitalares na rede pública municipal",
        "Médicos disponíveis na rede pública municipal",
        "Prontuário eletrônico",
        "Serviços on-line de saúde oferecidos aos pacientes",
        "Índice de risco e proteção à saúde dos nascidos vivos"
    ],
    "EDUCAÇÃO": [
        "Índice de equipamentos de tecnologia disponíveis nas escolas públicas municipais",
        "Taxa de analfabetismo",
        "Índice de desenvolvimento da educação básica (IDEB) - anos finais",
        "Vagas no ensino superior",
        "Centros de educação tecnológica",
        "Ações de educação para comunidades específicas",
        "Taxas de distorção idade-série",
        "Percentual de escolas municipais com acesso à internet",
        "Computadores para uso dos alunos"
    ],
    "INCLUSÃO SOCIAL": [
        "Políticas públicas para mulheres",
        "Inclusão social para grupos específicos"
    ],
    "INCLUSÃO DIGITAL": [
        "Promoção de inclusão digital",
        "Cursos de capacitação tecnológica"
    ],
    "SEGURANÇA PÚBLICA": [
        "Soluções em monitoramento para a segurança pública",
        "Taxa de homicídios",
        "Políticas públicas e ações para segurança pública"
    ],
    "INFRAESTRUTURA": [
        " Abrangência e Qualidade",
        "Governança de TI",
        "Infraestrutura de hardware e software",
        "Institucionalização da gestão de TI",
        "Planejamento para infraestrutura urbana",
        "Planejamento para infraestrutura de TI"
    ]
}

# colunas onde menor valor é melhor
INVERSE_COLS = [
    "Taxa de analfabetismo",
    "Taxas de distorção idade-série",
    "Taxa de homicídios"
]

# cria base final
df_s = df[INFO_COLS].copy()

# pega somente as colunas que existem
todas_colunas = []

for categoria, colunas in CATEGORIAS.items():
    for coluna in colunas:
        if coluna in df.columns:
            todas_colunas.append(coluna)

# limpa dados numéricos
for coluna in todas_colunas:
    df[coluna] = (
        df[coluna]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .replace("ND", np.nan)
        .replace("nan", np.nan)
        .replace("-", np.nan)
    )

    df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

# normaliza os indicadores de 0 a 10
for coluna in todas_colunas:
    serie = df[coluna].copy()

    # reduz outliers
    q1 = serie.quantile(0.25)
    q3 = serie.quantile(0.75)
    iqr = q3 - q1

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    serie = serie.clip(
        lower=limite_inferior,
        upper=limite_superior
    )

    minimo = serie.min()
    maximo = serie.max()

    if pd.notna(minimo) and pd.notna(maximo) and maximo != minimo:
        nota = ((serie - minimo) / (maximo - minimo)) * 10
    else:
        nota = pd.Series(5, index=df.index)

    if coluna in INVERSE_COLS:
        nota = 10 - nota

    df_s[coluna] = nota

# calcula nota por categoria
for categoria, colunas in CATEGORIAS.items():
    validas = [
        c for c in colunas
        if c in df_s.columns
    ]

    if len(validas) > 0:
        df_s[categoria] = df_s[validas].mean(axis=1).round(2)

# categorias encontradas
categorias_validas = [
    c for c in CATEGORIAS.keys()
    if c in df_s.columns
]

# calcula nota final
df_s["media"] = df_s[categorias_validas].mean(axis=1).round(2)
df_s["total"] = df_s[categorias_validas].sum(axis=1).round(2)

# cria ranking
ranking = df_s.sort_values(
    by="media",
    ascending=False
).reset_index(drop=True)

ranking["Ranking"] = ranking.index + 1

top10 = ranking.head(10)

# salva resultado no excel
ranking.to_excel(
    "ranking_cidades.xlsx",
    index=False
)

# ==================================================
# DASHBOARD EM UMA TELA
# ==================================================

fig = plt.figure(
    figsize=(21, 11),
    facecolor=COR_FUNDO
)

fig.suptitle(
    "CIDADES INTELIGENTES - BRASIL",
    fontsize=22,
    fontweight="bold",
    color=COR_TEXTO
)

gs = fig.add_gridspec(
    2,
    2,
    height_ratios=[1, 1],
    width_ratios=[1.1, 1.4],
    hspace=0.55,
    wspace=0.45
)

# ==================================================
# TABELA - TOP 10
# ==================================================

ax1 = fig.add_subplot(gs[0, 0])
ax1.set_facecolor(COR_PAINEL)
ax1.axis("off")

tabela = top10[
    ["Ranking", "Cidade", "Estado", "media"]
].copy()

tabela["media"] = tabela["media"].map(lambda x: f"{x:.2f}")

ax1.set_title(
    "Tabela - Top 10",
    pad=10,
    color=COR_TEXTO
)

table = ax1.table(
    cellText=tabela.values,
    colLabels=tabela.columns,
    cellLoc="center",
    loc="center"
)

table.auto_set_font_size(False)
table.set_fontsize(9)
table.scale(1.1, 1.45)

# cor da tabela
for (row, col), cell in table.get_celld().items():

    if row == 0:
        cell.set_facecolor("#3A506B")
        cell.set_text_props(color="white", weight="bold")
    else:
        cell.set_facecolor(COR_PAINEL)
        cell.set_text_props(color="white")

    cell.set_edgecolor(COR_AZUL)

# ==================================================
# GRÁFICO 2 - TOP 5 X PIORES 5
# ==================================================

ax2 = fig.add_subplot(gs[0, 1])
ax2.set_facecolor(COR_PAINEL)

top5 = ranking.head(5)
piores5 = ranking.tail(5)

comparacao = pd.concat([top5, piores5])

nomes_comparacao = comparacao["Cidade"] + " - " + comparacao["Estado"]

cores = [COR_VERDE] * 5 + [COR_VERMELHO] * 5

ax2.barh(
    nomes_comparacao[::-1],
    comparacao["media"][::-1],
    color=cores[::-1]
)

ax2.set_title(
    "Top 5 x Piores 5",
    color=COR_TEXTO
)

ax2.set_xlabel(
    "Nota média",
    color=COR_TEXTO
)

ax2.set_xlim(0, 10)

ax2.tick_params(
    axis="y",
    labelsize=9,
    colors=COR_TEXTO
)

ax2.tick_params(
    axis="x",
    colors=COR_TEXTO
)

for i, valor in enumerate(comparacao["media"][::-1]):
    ax2.text(
        valor + 0.05,
        i,
        f"{valor:.2f}",
        va="center",
        fontsize=9,
        color=COR_TEXTO
    )

# ==================================================
# GRÁFICO 3 - MÉDIA POR CATEGORIA
# ==================================================

ax3 = fig.add_subplot(gs[1, 0])
ax3.set_facecolor(COR_PAINEL)

medias_categorias = {}

for categoria in categorias_validas:
    medias_categorias[categoria] = df_s[categoria].mean()

nomes_categorias = [
    c.replace("QUALIDADE DO AR", "QUAL. AR")
     .replace("INCLUSÃO SOCIAL", "INC. SOCIAL")
     .replace("INCLUSÃO DIGITAL", "INC. DIGITAL")
     .replace("SEGURANÇA PÚBLICA", "SEG. PÚBLICA")
     .replace("INFRAESTRUTURA", "INFRA")
    for c in medias_categorias.keys()
]

ax3.bar(
    nomes_categorias,
    medias_categorias.values(),
    color=COR_AZUL
)

ax3.set_title(
    "Média Geral por Categoria",
    color=COR_TEXTO
)

ax3.set_ylabel(
    "Nota média",
    color=COR_TEXTO
)

ax3.set_ylim(0, 10)

ax3.tick_params(
    axis="x",
    rotation=25,
    labelsize=8,
    colors=COR_TEXTO
)

ax3.tick_params(
    axis="y",
    colors=COR_TEXTO
)

for i, valor in enumerate(medias_categorias.values()):
    ax3.text(
        i,
        valor + 0.10,
        f"{valor:.2f}",
        ha="center",
        fontsize=8,
        color=COR_TEXTO
    )

# ==================================================
# GRÁFICO 4 - TOP 10 CIDADES
# ==================================================

ax4 = fig.add_subplot(gs[1, 1])
ax4.set_facecolor(COR_PAINEL)

nomes_cidades = top10["Cidade"] + " - " + top10["Estado"]

ax4.barh(
    nomes_cidades[::-1],
    top10["media"][::-1],
    color=COR_AZUL
)

ax4.set_title(
    "Top 10 Cidades",
    color=COR_TEXTO
)

ax4.set_xlabel(
    "Nota média",
    color=COR_TEXTO
)

ax4.set_xlim(0, 10)

ax4.tick_params(
    axis="y",
    labelsize=9,
    colors=COR_TEXTO
)

ax4.tick_params(
    axis="x",
    colors=COR_TEXTO
)

for i, valor in enumerate(top10["media"][::-1]):
    ax4.text(
        valor + 0.05,
        i,
        f"{valor:.2f}",
        va="center",
        fontsize=9,
        color=COR_TEXTO
    )

# ajuste final
plt.subplots_adjust(
    top=0.88,
    bottom=0.10,
    left=0.08,
    right=0.97,
    hspace=0.55,
    wspace=0.45
)

plt.savefig(
    "dashboard_cidades_inteligentes.png",
    dpi=300,
    bbox_inches="tight",
    facecolor=COR_FUNDO
)

plt.show()