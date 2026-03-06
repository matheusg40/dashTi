import pandas as pd
import os
from datetime import datetime, date, timedelta

ARQUIVO         = "data/projetos.xlsx"
ARQUIVO_REUNIOES= "data/reunioes.xlsx"
ARQUIVO_SPRINT  = "data/sprints.xlsx"

# ── Checkboxes de progresso ───────────────────────────────────────────────────
ETAPAS_PROJETO = [
    "Levantamento de requisitos",
    "Aprovação do escopo",
    "Desenvolvimento / Execução",
    "Testes e validação",
    "Homologação com cliente",
    "Documentação",
    "Deploy / Entrega",
    "Encerramento e lições aprendidas",
]

def calcular_progresso(etapas_concluidas):
    total = len(ETAPAS_PROJETO)
    feitas = sum(etapas_concluidas)
    return round((feitas / total) * 100)

# ── Projetos ──────────────────────────────────────────────────────────────────
def carregar_dados():
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(ARQUIVO):
        df = pd.DataFrame(columns=[
            "Projeto", "Responsável", "Prioridade", "Status", "Progresso (%)",
            "Etapas", "Início", "Prazo", "Horas Gastas", "Descrição"
        ])
        df.to_excel(ARQUIVO, index=False)
        return df

    df = pd.read_excel(ARQUIVO)
    for col in ["Início", "Prazo"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Progresso (%)" in df.columns:
        df["Progresso (%)"] = pd.to_numeric(df["Progresso (%)"], errors="coerce").fillna(0)
    if "Prioridade" not in df.columns:
        df["Prioridade"] = "Média"
    if "Etapas" not in df.columns:
        df["Etapas"] = ""
    return df


def salvar_projeto(novo):
    df = carregar_dados()
    df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
    df.to_excel(ARQUIVO, index=False)


def atualizar_etapas(idx, etapas):
    df = carregar_dados()
    etapas_str = ",".join(["1" if e else "0" for e in etapas])
    progresso  = calcular_progresso(etapas)
    df.at[idx, "Etapas"]       = etapas_str
    df.at[idx, "Progresso (%)"] = progresso
    if progresso == 100:
        df.at[idx, "Status"] = "Concluído"
    df.to_excel(ARQUIVO, index=False)


def get_etapas(row):
    val = str(row.get("Etapas", ""))
    if val and val != "nan":
        bits   = val.split(",")
        result = [b.strip() == "1" for b in bits]
        while len(result) < len(ETAPAS_PROJETO):
            result.append(False)
        return result[:len(ETAPAS_PROJETO)]
    return [False] * len(ETAPAS_PROJETO)


def projetos_atrasados(df):
    hoje = pd.Timestamp(datetime.today().date())
    mask = (df["Prazo"] < hoje) & (df["Status"] != "Concluído")
    return df[mask]


# ── Reuniões ──────────────────────────────────────────────────────────────────
def carregar_reunioes():
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(ARQUIVO_REUNIOES):
        df = pd.DataFrame(columns=[
            "Título", "Responsável", "Participantes", "Empresa",
            "Data", "Horário", "Local", "Observações"
        ])
        df.to_excel(ARQUIVO_REUNIOES, index=False)
        return df
    df = pd.read_excel(ARQUIVO_REUNIOES)
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    return df


def salvar_reuniao(nova):
    df = carregar_reunioes()
    df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
    df.to_excel(ARQUIVO_REUNIOES, index=False)


def deletar_reuniao(index):
    df = carregar_reunioes()
    df = df.drop(index=index).reset_index(drop=True)
    df.to_excel(ARQUIVO_REUNIOES, index=False)


# ── Sprints ───────────────────────────────────────────────────────────────────
def segunda_da_semana():
    hoje = date.today()
    return hoje - timedelta(days=hoje.weekday())


def proxima_segunda():
    hoje = date.today()
    dias = (7 - hoje.weekday()) % 7
    if dias == 0:
        dias = 7
    return hoje + timedelta(days=dias)


def carregar_sprints():
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(ARQUIVO_SPRINT):
        df = pd.DataFrame(columns=[
            "Semana", "BU", "Progressos", "Desafios",
            "Próxima Sprint", "Meta", "Realizado", "Responsável"
        ])
        df.to_excel(ARQUIVO_SPRINT, index=False)
        return df
    df = pd.read_excel(ARQUIVO_SPRINT)
    if "Semana" in df.columns:
        df["Semana"] = pd.to_datetime(df["Semana"], errors="coerce")
    return df


def salvar_sprint(nova):
    df = carregar_sprints()
    df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
    df.to_excel(ARQUIVO_SPRINT, index=False)