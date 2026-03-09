import pandas as pd
import os
from datetime import datetime, date, timedelta

ARQUIVO         = "data/projetos.xlsx"
ARQUIVO_REUNIOES= "data/reunioes.xlsx"
ARQUIVO_SPRINT  = "data/sprints_db.xlsx"

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

# BUs válidas — fonte única da verdade
BUS_VALIDAS = [
    "Estratégia & Projetos",
    "Governança & Sustentação",
]

_BU_MAPA = {
    "projetos":      "Estratégia & Projetos",
    "estrategia":    "Estratégia & Projetos",
    "estratégia":    "Estratégia & Projetos",
    "governanca":    "Governança & Sustentação",
    "governança":    "Governança & Sustentação",
    "sustentacao":   "Governança & Sustentação",
    "sustentação":   "Governança & Sustentação",
}

def _normalizar_bu(bu):
    if pd.isna(bu):
        return bu
    bu_str = str(bu).strip()
    if bu_str in BUS_VALIDAS:
        return bu_str
    bu_lower = bu_str.lower()
    for chave, valor in _BU_MAPA.items():
        if chave in bu_lower:
            return valor
    return bu_str

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
    try:
        df.to_excel(ARQUIVO, index=False)
    except PermissionError:
        raise PermissionError("❌ Feche o arquivo 'projetos.xlsx' no Excel e tente novamente.")

def atualizar_etapas(idx, etapas):
    df = carregar_dados()
    etapas_str = ",".join(["1" if e else "0" for e in etapas])
    progresso  = calcular_progresso(etapas)
    df.at[idx, "Etapas"]        = etapas_str
    df.at[idx, "Progresso (%)"] = progresso
    if progresso == 100:
        df.at[idx, "Status"] = "Concluído"
    try:
        df.to_excel(ARQUIVO, index=False)
    except PermissionError:
        raise PermissionError("❌ Feche o arquivo 'projetos.xlsx' no Excel e tente novamente.")

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
    df = df.reset_index(drop=True)
    return df

def salvar_reuniao(nova):
    df = carregar_reunioes()
    df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
    try:
        df.to_excel(ARQUIVO_REUNIOES, index=False)
    except PermissionError:
        raise PermissionError("❌ Feche o arquivo 'reunioes.xlsx' no Excel e tente novamente.")

def deletar_reuniao(index):
    df = carregar_reunioes()
    df = df.drop(index=index).reset_index(drop=True)
    try:
        df.to_excel(ARQUIVO_REUNIOES, index=False)
    except PermissionError:
        raise PermissionError("❌ Feche o arquivo 'reunioes.xlsx' no Excel e tente novamente.")

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
            "Semana", "BU", "Responsável", "Progressos", "Desafios",
            "Próxima Sprint", "Meta", "Realizado"
        ])
        df.to_excel(ARQUIVO_SPRINT, index=False)
        return df

    df = pd.read_excel(ARQUIVO_SPRINT)
    if "Semana" in df.columns:
        df["Semana"] = pd.to_datetime(df["Semana"], errors="coerce")
    if "BU" in df.columns:
        bu_original = df["BU"].copy()
        df["BU"] = df["BU"].apply(_normalizar_bu)
        if not bu_original.equals(df["BU"]):
            try:
                df.to_excel(ARQUIVO_SPRINT, index=False)
            except PermissionError:
                pass
    return df

def salvar_sprint(nova):
    if "BU" in nova:
        nova["BU"] = _normalizar_bu(nova["BU"])
    df = carregar_sprints()
    df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
    try:
        df.to_excel(ARQUIVO_SPRINT, index=False)
    except PermissionError:
        raise PermissionError("❌ Feche o arquivo 'sprints_db.xlsx' no Excel e tente novamente.")