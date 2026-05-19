import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings

# evita aviso chato
warnings.filterwarnings("ignore")

# caminho do excel
arquivo = r"C:\Users\joao.soares\Documents\João\PYTHON\PROJETOS\JEAN\Excel\Cidade Inteligente.xlsx"

# lendo planilha
df = pd.read_excel(
    arquivo,
    sheet_name="DESSUST",
    header=[1, 2]
)

# ajeita os nomes das colunas
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

# colunas onde menor valor = melhor
INVERSE_COLS = [
    "Taxa de analfabetismo",
    "Taxas de distorção idade-série",
    "Taxa de homicídios"
]

# copia base principal
df_s = df[INFO_COLS].copy()

# pega todas colunas usadas
todas_colunas = []

for categoria, colunas in CATEGORIAS.items():

    for coluna in colunas:

        if coluna in df.columns:
            todas_colunas.append(coluna)

# limpa dados
for coluna in todas_colunas:

    df[coluna] = (
        df[coluna]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .replace("ND", np.nan)
        .replace("nan", np.nan)
        .replace("-", np.nan)
    )

    df[coluna] = pd.to_numeric(
        df[coluna],
        errors="coerce"
    )

# normaliza de 0 a 10
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

    # transforma em nota
    if pd.notna(minimo) and pd.notna(maximo) and maximo != minimo:

        nota = (
            (serie - minimo)
            / (maximo - minimo)
        ) * 10

    else:
        nota = pd.Series(5, index=df.index)

    # inverte coluna
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

        df_s[categoria] = (
            df_s[validas]
            .mean(axis=1)
            .round(2)
        )

# categorias válidas
categorias_validas = [
    c for c in CATEGORIAS.keys()
    if c in df_s.columns
]

# nota final
df_s["media"] = (
    df_s[categorias_validas]
    .mean(axis=1)
    .round(2)
)

df_s["total"] = (
    df_s[categorias_validas]
    .sum(axis=1)
    .round(2)
)

# cria ranking
ranking = df_s.sort_values(
    by="media",
    ascending=False
).reset_index(drop=True)

ranking["Ranking"] = ranking.index + 1

# top 10
top10 = ranking.head(10)

# mostra ranking
print("\nTOP 10 CIDADES:\n")

print(
    top10[
        ["Ranking", "Cidade", "Estado", "media"]
    ]
)

# salva excel
ranking.to_excel(
    "ranking_cidades.xlsx",
    index=False
)

# ==================================================
# GRAFICO TOP 10
# ==================================================

plt.figure(figsize=(14, 7))

nomes_cidades = (
    top10["Cidade"]
    + " - "
    + top10["Estado"]
)

plt.barh(
    nomes_cidades[::-1],
    top10["media"][::-1]
)

plt.title("Top 10 Cidades Inteligentes")
plt.xlabel("Nota média")
plt.ylabel("Cidade")

plt.xlim(0, 10)

# coloca nota do lado
for i, valor in enumerate(top10["media"][::-1]):

    plt.text(
        valor + 0.05,
        i,
        f"{valor:.2f}",
        va="center"
    )

plt.tight_layout()

plt.savefig(
    "grafico_top10.png",
    dpi=300
)

plt.show()

# ==================================================
# GRAFICO MEDIA DAS CATEGORIAS
# ==================================================

medias_categorias = {}

for categoria in categorias_validas:

    medias_categorias[categoria] = (
        df_s[categoria].mean()
    )

plt.figure(figsize=(14, 6))

plt.bar(
    medias_categorias.keys(),
    medias_categorias.values()
)

plt.title("Média Geral por Categoria")

plt.ylabel("Nota média")

plt.ylim(0, 10)

plt.xticks(
    rotation=45,
    ha="right"
)

# coloca valor encima
for i, valor in enumerate(medias_categorias.values()):

    plt.text(
        i,
        valor + 0.1,
        f"{valor:.2f}",
        ha="center"
    )

plt.tight_layout()

plt.savefig(
    "grafico_categorias.png",
    dpi=300
)

plt.show()

# ==================================================
# GRAFICO COMPARAÇÃO TOP 10
# ==================================================

top10_categorias = top10[
    ["Cidade"] + categorias_validas
].set_index("Cidade")

top10_categorias.plot(
    kind="bar",
    figsize=(15, 7)
)

plt.title("Notas por Categoria - Top 10")

plt.ylabel("Nota")

plt.xlabel("Cidade")

plt.ylim(0, 10)

plt.xticks(
    rotation=45,
    ha="right"
)

plt.legend(
    title="Categorias",
    bbox_to_anchor=(1.05, 1),
    loc="upper left"
)

plt.tight_layout()

plt.savefig(
    "grafico_comparacao.png",
    dpi=300
)

plt.show()

print("\nArquivos gerados:")
print("ranking_cidades.xlsx")
print("grafico_top10.png")
print("grafico_categorias.png")
print("grafico_comparacao.png")