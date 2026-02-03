import pandas as pd
import os


# Configurações
ARQUIVO = "CAMINHO/DO/ARQUIVO.xlsx"
ANO = 2026

PASTA_SAIDA = os.path.join(os.path.dirname(ARQUIVO), "output")
os.makedirs(PASTA_SAIDA, exist_ok=True)


# Manager -> Cost Center
MANAGERS_CC = {
    "Manager A": "CC001",
    "Manager B": "CC002",
    "Manager C": "CC003",
    "Manager D": "CC004",
    "Manager E": "CC005",
}


# =====================
# TABLE 1
# =====================

estrutura = pd.read_excel(ARQUIVO, sheet_name="estrutura")
estrutura.columns = estrutura.columns.str.strip()

linhas = []

for _, row in estrutura.iterrows():
    manager = row["GERENTE DTM"]
    researcher = row["DTM"]

    if manager in MANAGERS_CC and pd.notna(researcher):
        linhas.append({
            "Year": ANO,
            "Cost Center": MANAGERS_CC[manager],
            "Researcher": researcher
        })

# manager as researcher
for manager, cc in MANAGERS_CC.items():
    linhas.append({
        "Year": ANO,
        "Cost Center": cc,
        "Researcher": manager
    })

df_1 = pd.DataFrame(linhas)
df_1.to_excel(
    os.path.join(PASTA_SAIDA, f"CostCenter_Researcher_{ANO}.xlsx"),
    index=False
)

print("Table 1 generated")


# =====================
# TABLE 2
# =====================

cost_centers = list(MANAGERS_CC.values())

depara = pd.read_excel(ARQUIVO, sheet_name="de-para")
depara.columns = depara.columns.str.strip()

classes = depara["Descricao classe de custo"].dropna().unique()

months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

linhas = []

for cc in cost_centers:
    for classe in classes:
        for month in months:
            linhas.append({
                "Cost Center": cc,
                "Month": month,
                "Cost Class": classe,
                "Year": ANO
            })

df_2 = pd.DataFrame(linhas)
df_2.to_excel(
    os.path.join(PASTA_SAIDA, f"CostCenter_Month_CostClass_{ANO}.xlsx"),
    index=False
)

print("Table 2 generated")


# =====================
# TABLE 3
# =====================

estrutura = estrutura.drop_duplicates(subset=["GERENTE DTM", "DTM"])

linhas = []

for _, row in estrutura.iterrows():
    manager = row["GERENTE DTM"]
    researcher = row["DTM"]

    if manager in MANAGERS_CC:
        cc = MANAGERS_CC[manager]

        for month in months:
            for classe in classes:
                linhas.append({
                    "Cost Center": cc,
                    "Month": month,
                    "Cost Class": classe,
                    "Year": ANO,
                    "Researcher": researcher
                })

df_3 = pd.DataFrame(linhas).drop_duplicates()
df_3.to_excel(
    os.path.join(PASTA_SAIDA, f"Budget_Researcher_Month_CostClass_{ANO}.xlsx"),
    index=False
)

print("Table 3 generated")
