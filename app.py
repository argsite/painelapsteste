import io
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple
from datetime import datetime
from openpyxl.utils import get_column_letter

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="APS 360 - Painel de Indicadores",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# Utilidades
# =========================


def strip_accents(text: str) -> str:
    text = str(text)
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch)
    )


def normalize_col(name: str) -> str:
    name = strip_accents(str(name).strip().lower())
    name = re.sub(r"[^a-z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name


def to_bool(series: pd.Series) -> pd.Series:
    if series.dtype == bool:
        return series.fillna(False)
    vals = series.astype(str).str.strip().str.lower()
    return vals.isin(["1", "true", "sim", "s", "x", "ok", "yes"])


def parse_count(series: pd.Series) -> pd.Series:
    vals = series.astype(str).str.strip().str.lower()
    vals = vals.replace({"": np.nan, "nan": np.nan, "none": np.nan, "n/a": np.nan})
    vals = vals.str.replace("+", "", regex=False)
    return pd.to_numeric(vals, errors="coerce")


def classificar_score(score: float) -> str:
    if score >= 75:
        return "Ótimo"
    if score >= 50:
        return "Bom"
    if score >= 25:
        return "Suficiente"
    return "Regular"


def faixa_etaria(idade: float) -> str:
    if pd.isna(idade):
        return "Sem idade"
    idade = int(idade)
    if idade < 1:
        return "<1"
    if idade <= 4:
        return "1-4"
    if idade <= 9:
        return "5-9"
    if idade <= 14:
        return "10-14"
    if idade <= 19:
        return "15-19"
    if idade <= 39:
        return "20-39"
    if idade <= 59:
        return "40-59"
    return "60+"


def ensure_column(df: pd.DataFrame, col: str, default=None):
    if col not in df.columns:
        df[col] = default


def first_existing(df: pd.DataFrame, cols: List[str]) -> Optional[str]:
    for c in cols:
        if c in df.columns:
            return c
    return None


def map_first(df: pd.DataFrame, target: str, candidates: List[str], default=""):
    src = first_existing(df, candidates)
    if src and target not in df.columns:
        df[target] = df[src]
    elif target not in df.columns:
        df[target] = default


def infer_tipo_equipe_from_text(series: pd.Series) -> pd.Series:
    vals = series.astype(str).str.upper()
    out = np.where(vals.str.contains(" 76") | vals.str.contains("TIPO 76"), "76", "")
    out = np.where(
        (pd.Series(out, index=series.index) == "")
        & (vals.str.contains(" 70") | vals.str.contains("TIPO 70")),
        "70",
        out,
    )
    return pd.Series(out, index=series.index)

# Nomes arquivos exportados
def slugify_filename(text: str) -> str:
    text = strip_accents(str(text)).lower().strip()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "arquivo"

def clean_team_name(value: str) -> str:
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"^\s*\d+\s*[-–—]\s*", "", text)
    return text.strip()

def friendly_indicator_name(spec: IndicatorSpec) -> str:
    return slugify_filename(spec.name)

def friendly_pendencia_name(letra: str) -> str:
    return slugify_filename(f"pendencia_{letra}")

def friendly_team_name(df: pd.DataFrame) -> str:
    if "equipe_area" in df.columns and df["equipe_area"].notna().any():
        teams = sorted({
            clean_team_name(v)
            for v in df["equipe_area"].dropna().astype(str)
            if clean_team_name(v)
        })
        if len(teams) == 1:
            return slugify_filename(teams[0])
        if len(teams) > 1:
            return slugify_filename("_".join(teams))

    if "equipe" in df.columns and df["equipe"].notna().any():
        teams = sorted({
            clean_team_name(v)
            for v in df["equipe"].dropna().astype(str)
            if clean_team_name(v)
        })
        if len(teams) == 1:
            return slugify_filename(teams[0])
        if len(teams) > 1:
            return slugify_filename("_".join(teams))

    if "equipe_vinculo" in df.columns and df["equipe_vinculo"].notna().any():
        teams = sorted({
            clean_team_name(v)
            for v in df["equipe_vinculo"].dropna().astype(str)
            if clean_team_name(v)
        })
        if len(teams) == 1:
            return slugify_filename(teams[0])
        if len(teams) > 1:
            return slugify_filename("_".join(teams))

    return "todas_as_equipes"



TAB_SHORT_LABELS = {
    "C2": {
        "A": "Consulta precoce",
        "B": "Consultas",
        "C": "Peso e altura",
        "D": "Visita domiciliar",
        "E": "Vacinas",
    },
    "C3": {
        "A": "Consulta inicial",
        "B": "Consultas pré-natal",
        "C": "Pressão arterial",
        "D": "Peso e altura",
        "E": "Visitas domiciliares",
        "F": "dTpa",
        "G": "Exames 1º tri",
        "H": "Exames 3º tri",
        "I": "Puerpério",
        "J": "Visita puerpério",
        "K": "Saúde bucal",
    },
    "C4": {
        "A": "Consulta",
        "B": "Pressão arterial",
        "C": "Peso e altura",
        "D": "Visitas domiciliares",
        "E": "Hemoglobina glicada",
        "F": "Avaliação dos pés",
    },
    "C5": {
        "A": "Consulta",
        "B": "Pressão arterial",
        "C": "Peso e altura",
        "D": "Visitas domiciliares",
    },
    "C6": {
        "A": "Consulta",
        "B": "Peso e altura",
        "C": "Visitas domiciliares",
        "D": "Influenza",
    },
    "C7": {
        "A": "Exame citopatológico",
        "B": "Vacina HPV",
        "C": "Saúde reprodutiva",
        "D": "Mamografia",
    },
}

# =========================
# Especificações
# =========================


@dataclass
class IndicatorSpec:
    code: str
    name: str
    type: str
    description: str
    weights: Dict[str, int] = field(default_factory=dict)
    non_conditionals: Dict[str, Callable[[pd.DataFrame], pd.Series]] = field(
        default_factory=dict
    )
    numerator_col: Optional[str] = None
    denominator_col: Optional[str] = None
    entity_label: str = "pessoas"
    applicable_age_rule: Optional[Callable[[pd.DataFrame], pd.Series]] = None


BOA_PRATICA_LABELS = {
    "C2": {
        "c2_a_ok": "A - Ter a 1ª consulta presencial realizada por médica(o) ou enfermeira(o), até o 30º dia de vida",
        "c2_b_ok": "B - Ter pelo menos 09 (nove) consultas presenciais ou remotas realizadas por médica(o) ou enfermeira(o) até dois anos de vida",
        "c2_c_ok": "C - Ter pelo menos 09 (nove) registros simultâneos de peso e altura até os dois anos de vida",
        "c2_d_ok": "D - Ter pelo menos 02 (duas) visitas domiciliares realizadas por ACS/TACS, sendo a primeira até os primeiros 30 (trinta) dias de vida e a segunda até os 06 (seis) meses de vida",
        "c2_e_ok": "E - Ter vacinas registradas com todas as doses recomendadas até os 2 anos",
    },
    "C3": {
        "c3_a_ok": "A - Ter a 1ª consulta presencial ou remota realizada por médica(o) ou enfermeira(o), até a 12ª semana de gestação.",
        "c3_b_ok": "B - Ter pelo menos 07 (sete) consultas presenciais ou remotas realizadas por médica(o) ou enfermeira(o) durante o período da gestação.",
        "c3_c_ok": "C - Ter pelo menos 07 (sete) registro de aferição de pressão arterial realizados durante o período da gestação.",
        "c3_d_ok": "D - Ter pelo menos 07 (sete) registros simultâneos de peso e altura durante o período da gestação.",
        "c3_e_ok": "E - Ter pelo menos 03 (três) visitas domiciliares realizadas por ACS/TACS, após a primeira consulta do pré-natal.",
        "c3_f_ok": "F - Ter vacina acelular contra difteria, tétano, coqueluche (dTpa) registrada a partir da 20ª semana de cada gestação.",
        "c3_g_ok": "G - Ter registro dos testes rápidos ou dos exames avaliados para sífilis, HIV e hepatites B e C realizados no 1º trimestre de cada gestação.",
        "c3_h_ok": "H - Ter registro dos testes rápidos ou dos exames avaliados para sífilis e HIV realizados no 3º trimestre de cada gestação.",
        "c3_i_ok": "I - Ter pelo menos 01 registro de consulta presencial ou remota realizada por médica(o) ou enfermeira(o) durante o puerpério.",
        "c3_j_ok": "J - Ter pelo menos 01 visita domiciliar realizada por ACS/TACS durante o puerpério.",
        "c3_k_ok": "K - Ter pelo menos 01 atividade em saúde bucal realizada por cirurgiã(ão) dentista ou técnica(o) de saúde bucal durante o período da gestação.",
    },
    "C4": {
        "c4_a_ok": "A - Ter pelo menos 01 (uma) consulta presencial ou remota realizadas por médica(o) ou enfermeira(o), nos últimos 06 (seis) meses",
        "c4_b_ok": "B - Ter pelo menos 01 (um) registro de aferição de pressão arterial realizado nos últimos 06 (seis) meses",
        "c4_c_ok": "C - Ter pelo menos 01 (um) registro simultâneos de peso e altura realizado nos últimos 12 (doze) meses",
        "c4_d_ok": "D - Ter pelo menos 02 (duas) visitas domiciliares realizadas por ACS/TACS, com intervalo mínimo de 30 (trinta) dias, nos últimos 12 (doze) meses",
        "c4_e_ok": "E - Ter pelo menos 01 (um) registro de solicitação de hemoglobina glicada realizada ou avaliada, nos últimos 12 (doze) meses",
        "c4_f_ok": "F - Ter pelo menos 01 (uma) avaliação dos pés realizada nos últimos 12 (doze) meses",
    },
    "C5": {
        "c5_a_ok": "A - Ter pelo menos 01 (uma) consulta presencial ou remota realizadas por médica(o) ou enfermeira(o), nos últimos 06 (seis) meses",
        "c5_b_ok": "B - Ter pelo menos 01 (um) registro de aferição de pressão arterial realizado nos últimos 06 (seis) meses",
        "c5_c_ok": "C - Ter pelo menos 01 (um) registro simultâneos de peso e altura realizado nos últimos 12 (doze) meses",
        "c5_d_ok": "D - Ter pelo menos 02 (duas) visitas domiciliares realizadas por ACS/TACS, com intervalo mínimo de 30 (trinta) dias, nos últimos 12 (doze) meses",
    },
    "C6": {
        "consulta_ok": "A - Ter registro de pelo menos 01 consulta presencial ou remota por profissional médica(o) ou enfermeira(o) realizada nos últimos 12 meses",
        "antropometria_ok": "B - Ter realizado pelo menos 01 (um) registro simultâneo (no mesmo dia) de peso e altura para avaliação antropométrica nos últimos 12 meses",
        "visitas_ok": "C - Ter registro de pelo menos 02 visitas domiciliares por ACS/TACS, com intervalo mínimo de 30 dias, realizadas nos últimos 12 meses",
        "influenza_ok": "D - Ter registro de 1 dose da vacina contra influenza realizada nos últimos 12 meses",
    },
    "C7": {
        "c7_a_ok": "A - Exame citopatológico (25-64 anos) ou molecular de HPV (até 60 meses)",
        "c7_b_ok": "B - Pelo menos 1 dose da vacina HPV (9-14 anos)",
        "c7_c_ok": "C - Atendimento em saúde sexual e reprodutiva nos últimos 12 meses",
        "c7_d_ok": "D - Mamografia de rastreamento (50-69 anos) realizada ou avaliada em 24 meses",
    },
}


def label_boa_pratica(indicator_code: str, col: str) -> str:
    return BOA_PRATICA_LABELS.get(indicator_code, {}).get(col, col.replace("_", " ").capitalize())


INDICATORS: Dict[str, IndicatorSpec] = {
    "C1": IndicatorSpec(
        code="C1",
        name="Mais acesso",
        type="percentual",
        description="Indicador operacional local de acesso/vínculo a partir do relatório importado.",
        numerator_col="numerador_c1",
        denominator_col="denominador_c1",
        entity_label="pessoas cadastradas",
    ),
    "C2": IndicatorSpec(
        code="C2",
        name="Cuidado no desenvolvimento infantil",
        type="score",
        description="Monitoramento da puericultura de crianças até 2 anos com base nas práticas A–E.",
        weights={
            "c2_a_ok": 20,
            "c2_b_ok": 20,
            "c2_c_ok": 20,
            "c2_d_ok": 20,
            "c2_e_ok": 20,
        },
        entity_label="crianças acompanhadas",
    ),
    "C3": IndicatorSpec(
        code="C3",
        name="Cuidado na gestação e puerpério",
        type="score",
        description="Painel operacional local para gestantes e puérperas com base nas práticas A–K.",
        weights={
            "c3_a_ok": 10,
            "c3_b_ok": 9,
            "c3_c_ok": 9,
            "c3_d_ok": 9,
            "c3_e_ok": 9,
            "c3_f_ok": 9,
            "c3_g_ok": 9,
            "c3_h_ok": 9,
            "c3_i_ok": 9,
            "c3_j_ok": 9,
            "c3_k_ok": 9,
        },
        entity_label="gestantes/puérperas",
    ),
    "C4": IndicatorSpec(
        code="C4",
        name="Cuidado da pessoa com diabetes",
        type="score",
        description="Pontuação por pessoa com diabetes até 100 pontos a partir das práticas A–F.",
        weights={
            "c4_a_ok": 20,
            "c4_b_ok": 15,
            "c4_c_ok": 15,
            "c4_d_ok": 20,
            "c4_e_ok": 15,
            "c4_f_ok": 15,
        },
        entity_label="pessoas com diabetes",
    ),
    "C5": IndicatorSpec(
        code="C5",
        name="Cuidado da pessoa com hipertensão",
        type="score",
        description="Pontuação por pessoa com hipertensão até 100 pontos a partir das práticas A–D.",
        weights={
            "c5_a_ok": 25,
            "c5_b_ok": 25,
            "c5_c_ok": 25,
            "c5_d_ok": 25,
        },
        entity_label="pessoas com hipertensão",
    ),
    "C6": IndicatorSpec(
        code="C6",
        name="Cuidado da pessoa idosa",
        type="score",
        description="Pontuação por pessoa idosa até 100 pontos.",
        weights={
            "consulta_ok": 25,
            "antropometria_ok": 25,
            "visitas_ok": 25,
            "influenza_ok": 25,
        },
        entity_label="pessoas idosas",
    ),
    "C7": IndicatorSpec(
        code="C7",
        name="Cuidado da mulher na prevenção do câncer",
        type="score",
        description="Painel operacional local para prevenção do câncer da mulher com base nas práticas A–D.",
        weights={
            "c7_a_ok": 20,
            "c7_b_ok": 30,
            "c7_c_ok": 30,
            "c7_d_ok": 20,
        },
        entity_label="mulheres acompanhadas",
    ),
}

# =========================
# Leitura e identificação
# =========================


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    suffix = uploaded_file.name.lower()
    if suffix.endswith(".csv"):
        for enc in ["utf-8", "latin1", "cp1252"]:
            uploaded_file.seek(0)
            try:
                return pd.read_csv(uploaded_file, encoding=enc, dtype=str)
            except Exception:
                pass
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, dtype=str)
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, dtype=str)


def detect_indicator_from_columns(df: pd.DataFrame, filename: str) -> Optional[str]:
    cols = set(df.columns)
    name = normalize_col(filename)

    if {"hemoglobina_glicada", "avaliacao_dos_pes"}.issubset(cols) or "diabetes" in name:
        return "C4"
    if "hipertensao" in name or {
        "afericao_de_pa",
        "qtd_registros_de_peso_altura",
    }.issubset(cols):
        if "hemoglobina_glicada" not in cols and "avaliacao_dos_pes" not in cols:
            return "C5"
    if "idosa" in name or "idoso" in name or "vacina_influenza" in cols:
        return "C6"
    if "gestante" in name or "puerpera" in name or "gestacao" in name:
        return "C3"
    if "desenvolvimento_infantil" in name or "infantil" in name or "crianca" in name:
        return "C2"
    if "cancer" in name or "mulher" in name or "mamografia" in cols or "citopatologico" in cols:
        return "C7"
    if "acesso" in name:
        return "C1"
    return None

# =========================
# Pré-processamento
# =========================


def preprocess_df(df: pd.DataFrame, indicator_code: Optional[str] = None) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_col(c) for c in df.columns]

    map_first(df, "nome", ["nome", "nome_completo", "cidadao", "usuario", "paciente"])
    map_first(df, "cpf", ["cpf"])
    map_first(df, "cns", ["cns", "cns_cidadao", "cartao_sus"])
    map_first(df, "data_nascimento", ["data_nascimento", "dt_nascimento", "nascimento", "data_nasc", "data_de_nascimento"])
    map_first(df, "idade", ["idade"])
    map_first(df, "endereco", ["endereco", "logradouro"])
    map_first(df, "equipe", ["equipe_area", "equipe", "equipe_de_area"])
    map_first(df, "micro_area", ["micro_area", "microarea"])
    map_first(df, "equipe_vinculo", ["equipe_vinculo", "equipe_de_vinculo"])
    map_first(df, "cadastro_atualizado", ["cadastro_atualizado"])
    map_first(df, "data_atualizacao_cadastro", ["data_atualizacao_cadastro"])
    map_first(df, "acompanhado", ["acompanhado"])

    if "idade" in df.columns:
        df["idade"] = pd.to_numeric(df["idade"], errors="coerce")
    else:
        df["idade"] = np.nan
    df["faixa_etaria"] = df["idade"].apply(faixa_etaria)

    if "tipo_equipe" not in df.columns:
        if "equipe_vinculo" in df.columns:
            df["tipo_equipe"] = infer_tipo_equipe_from_text(df["equipe_vinculo"])
        else:
            df["tipo_equipe"] = ""

    if "consulta_medica_enfermagem" in df.columns:
        df["consulta_ok"] = to_bool(df["consulta_medica_enfermagem"])
    elif "consulta" in df.columns:
        df["consulta_ok"] = to_bool(df["consulta"])
    else:
        df["consulta_ok"] = False

    if "afericao_de_pressao_arterial" in df.columns:
        df["pa_ok"] = to_bool(df["afericao_de_pressao_arterial"])
        df["c5_b_ok"] = df["pa_ok"]
    elif "afericao_de_pa" in df.columns:
        df["pa_ok"] = to_bool(df["afericao_de_pa"])
        df["c5_b_ok"] = df["pa_ok"]
    else:
        df["pa_ok"] = False
        df["c5_b_ok"] = False

    if "hemoglobina_glicada" in df.columns:
        df["hba1c_ok"] = to_bool(df["hemoglobina_glicada"])
    else:
        df["hba1c_ok"] = False

    if "avaliacao_dos_pes" in df.columns:
        df["pes_ok"] = to_bool(df["avaliacao_dos_pes"])
    else:
        df["pes_ok"] = False

    if "qtd_registros_de_peso_altura" in df.columns:
        qtd = parse_count(df["qtd_registros_de_peso_altura"])
        df["antropometria_ok"] = qtd.fillna(0).ge(1)
    elif "peso_altura" in df.columns:
        df["antropometria_ok"] = to_bool(df["peso_altura"])
    else:
        df["antropometria_ok"] = False

    if "qtd_visitas_domiciliares" in df.columns:
        qtd_vis = parse_count(df["qtd_visitas_domiciliares"])
        df["visita_ok"] = qtd_vis.fillna(0).ge(2)
        df["visitas_ok"] = df["visita_ok"]
    elif "visita_domiciliar" in df.columns:
        df["visita_ok"] = to_bool(df["visita_domiciliar"])
        df["visitas_ok"] = df["visita_ok"]
    else:
        df["visita_ok"] = False
        df["visitas_ok"] = False

    if "vacina_influenza" in df.columns:
        df["influenza_ok"] = to_bool(df["vacina_influenza"])
    else:
        df["influenza_ok"] = False

    df["cadastro_ok"] = to_bool(df["cadastro_atualizado"]) if "cadastro_atualizado" in df.columns else False
    df["atendimento_ok"] = to_bool(df["acompanhado"]) if "acompanhado" in df.columns else df["consulta_ok"]
    df["numerador_c1"] = (df["cadastro_ok"] | df["atendimento_ok"]).astype(int)
    df["denominador_c1"] = 1

    if indicator_code == "C2" or (indicator_code is None and df["idade"].notna().any()):
        df["vacina_ok"] = to_bool(df["vacina_influenza"]) if "vacina_influenza" in df.columns else False

    df["exame_ok"] = False
    possible_exam_cols = [
        c for c in df.columns if any(k in c for k in ["exame", "teste", "hemoglobina", "citopatologico", "mamografia"])
    ]
    if possible_exam_cols:
        temp = pd.Series(False, index=df.index)
        for c in possible_exam_cols:
            temp = temp | to_bool(df[c])
        df["exame_ok"] = temp

    df["citopatologico_ok"] = False
    df["mamografia_ok"] = False
    if "citopatologico" in df.columns:
        df["citopatologico_ok"] = to_bool(df["citopatologico"])
    elif "acompanhado" in df.columns and indicator_code == "C7":
        df["citopatologico_ok"] = to_bool(df["acompanhado"])

    if "mamografia" in df.columns:
        df["mamografia_ok"] = to_bool(df["mamografia"])

    if indicator_code == "C2":
        df = preprocess_c2_visits(df)

    if indicator_code == "C3":
        df = preprocess_c3_puerperio_visits(df)

    # C2
    if indicator_code == "C2":
        # A - 1ª consulta até 30 dias
        # Procurar coluna da consulta 1º mês por padrão no nome
        consulta_1m_col = None
        for c in df.columns:
            if (
                "consulta" in c
                and "medica" in c
                and "enfermagem" in c
                and "1" in c
                and "mes" in c
            ):
                consulta_1m_col = c
                break
    
        if consulta_1m_col:
            # Usar essa coluna para marcar a boa prática A
            df["c2_a_ok"] = to_bool(df[consulta_1m_col])
    
        # B, C, D, E seguem como você já deixou...
        if "nr_consultas" in df.columns:
            df["c2_b_ok"] = parse_count(df["nr_consultas"]).fillna(0).ge(9)
    
        if "qtd_registros_de_peso_altura" in df.columns:
            df["c2_c_ok"] = parse_count(df["qtd_registros_de_peso_altura"]).fillna(0).ge(9)
    
        if "visita_domiciliar_1_mes" in df.columns and "visita_domiciliar_6_mes" in df.columns:
            v1 = to_bool(df["visita_domiciliar_1_mes"])
            v6 = to_bool(df["visita_domiciliar_6_mes"])
            df["c2_d_ok"] = v1 & v6
    
        if "esquema_vacinal_completo" in df.columns:
            df["c2_e_ok"] = to_bool(df["esquema_vacinal_completo"])

    # C3
    if indicator_code == "C3":
        # A - Consulta inicial até 12 semanas
        # Coluna: "Consulta de pré-natal até 12 semanas"
        if "consulta_de_pre_natal_ate_12_semanas" in df.columns:
            df["c3_a_ok"] = to_bool(df["consulta_de_pre_natal_ate_12_semanas"])
    
        # B - Pelo menos 7 consultas de pré-natal
        # Coluna: "Consulta médica/enfermagem Gestação" (número de consultas)
        if "consulta_medica_enfermagem_gestacao" in df.columns:
            df["c3_b_ok"] = parse_count(
                df["consulta_medica_enfermagem_gestacao"]
            ).fillna(0).ge(7)
    
        # C - Pelo menos 7 aferições de PA
        # Coluna: "Aferição de pressão arterial"
        if "afericao_de_pressao_arterial" in df.columns:
            df["c3_c_ok"] = parse_count(
                df["afericao_de_pressao_arterial"]
            ).fillna(0).ge(7)
    
        # D - Pelo menos 7 registros simultâneos de peso/altura
        # Coluna: "Registro de peso/altura"
        if "registro_de_peso_altura" in df.columns:
            df["c3_d_ok"] = parse_count(
                df["registro_de_peso_altura"]
            ).fillna(0).ge(7)
    
        # E - Pelo menos 3 visitas domiciliares na gestação
        # Coluna: "Visitas domiciliares (ACS/TACS) Gestação"
        if "visitas_domiciliares_acs_tacs_gestacao" in df.columns:
            df["c3_e_ok"] = parse_count(
                df["visitas_domiciliares_acs_tacs_gestacao"]
            ).fillna(0).ge(3)
            
    
        # F - Vacina dTpa registrada
        # Coluna: "Vacina dTpa"
        if "vacina_dtpa" in df.columns:
            df["c3_f_ok"] = to_bool(df["vacina_dtpa"])
    
        # G - Exames 1º trimestre (sífilis, HIV, hepatites B e C)
        # Usar todas as colunas de testes do 1º trimestre combinadas
        cols_1t = [
            "teste_rapido_sifilis_primeiro_trimestre",
            "teste_rapido_hiv_primeiro_trimestre",
            "teste_rapido_hepatite_b_primeiro_trimestre",
            "teste_rapido_hepatite_c_primeiro_trimestre",
        ]
        present_1t = [c for c in cols_1t if c in df.columns]
        if present_1t:
            temp = pd.Series(False, index=df.index)
            for c in present_1t:
                temp = temp | to_bool(df[c])
            df["c3_g_ok"] = temp
    
        # H - Exames 3º trimestre (sífilis e HIV)
        cols_3t = [
            "teste_rapido_sifilis_terceiro_trimestre",
            "teste_rapido_hiv_terceiro_trimestre",
        ]
        present_3t = [c for c in cols_3t if c in df.columns]
        if present_3t:
            temp = pd.Series(False, index=df.index)
            for c in present_3t:
                temp = temp | to_bool(df[c])
            df["c3_h_ok"] = temp
    
        # I - Consulta médica/enfermagem no puerpério
        if "consulta_medica_enfermagem_puerperio" in df.columns:
            df["c3_i_ok"] = to_bool(df["consulta_medica_enfermagem_puerperio"])
    
        # J - Pelo menos 1 visita domiciliar no puerpério
        if "visitas_domiciliares_acs_tacs_puerperio" in df.columns:
            df["c3_j_ok"] = parse_count(
                df["visitas_domiciliares_acs_tacs_puerperio"]
            ).fillna(0).ge(1)
    
        # K - Pelo menos 1 atividade em saúde bucal na gestação
        if "avaliacao_odontologica_gestacao" in df.columns:
            df["c3_k_ok"] = to_bool(df["avaliacao_odontologica_gestacao"])

         
    

    # C4
    if indicator_code == "C4":
        df["c4_a_ok"] = to_bool(df["consulta_medica_enfermagem"]) if "consulta_medica_enfermagem" in df.columns else df.get("consulta_ok", False)
        df["c4_b_ok"] = to_bool(df["afericao_de_pa"]) if "afericao_de_pa" in df.columns else df.get("pa_ok", False)
        if "qtd_registros_de_peso_altura" in df.columns:
            df["c4_c_ok"] = parse_count(df["qtd_registros_de_peso_altura"]).fillna(0).ge(1)
        else:
            df["c4_c_ok"] = df.get("antropometria_ok", False)
        if "qtd_visitas_domiciliares" in df.columns:
            df["c4_d_ok"] = parse_count(df["qtd_visitas_domiciliares"]).fillna(0).ge(2)
        else:
            df["c4_d_ok"] = df.get("visita_ok", False)
        if "hemoglobina_glicada" in df.columns:
            df["c4_e_ok"] = to_bool(df["hemoglobina_glicada"])
        else:
            df["c4_e_ok"] = df.get("hba1c_ok", False)
        if "avaliacao_dos_pes" in df.columns:
            df["c4_f_ok"] = to_bool(df["avaliacao_dos_pes"])
        else:
            df["c4_f_ok"] = df.get("pes_ok", False)

    # C5
    if indicator_code == "C5":
        df["c5_a_ok"] = to_bool(df["consulta_medica_enfermagem"]) if "consulta_medica_enfermagem" in df.columns else df.get("consulta_ok", False)
        df["c5_b_ok"] = to_bool(df["afericao_de_pa"]) if "afericao_de_pa" in df.columns else df.get("pa_ok", False)
        if "qtd_registros_de_peso_altura" in df.columns:
            df["c5_c_ok"] = parse_count(df["qtd_registros_de_peso_altura"]).fillna(0).ge(1)
        else:
            df["c5_c_ok"] = df.get("antropometria_ok", False)
        if "qtd_visitas_domiciliares" in df.columns:
            df["c5_d_ok"] = parse_count(df["qtd_visitas_domiciliares"]).fillna(0).ge(2)
        else:
            df["c5_d_ok"] = df.get("visita_ok", False)

    # C6
    if indicator_code == "C6":
        if "consulta_medica_enfermagem" in df.columns:
            df["consulta_ok"] = to_bool(df["consulta_medica_enfermagem"])
        if "qtd_registros_de_peso_altura" in df.columns:
            df["antropometria_ok"] = parse_count(df["qtd_registros_de_peso_altura"]).fillna(0).ge(1)
        if "qtd_visitas_domiciliares" in df.columns:
            df["visitas_ok"] = parse_count(df["qtd_visitas_domiciliares"]).fillna(0).ge(2)
        if "vacina_influenza" in df.columns:
            df["influenza_ok"] = to_bool(df["vacina_influenza"])

    # C7
    if indicator_code == "C7":
        c7_map = {
            "c7_a_ok": ["rast_cancer_do_colo_do_utero", "rast_cancer_do_colo_do_tero", "rast_cancer_do_colodo_utero", "c7_a_ok"],
            "c7_b_ok": ["vacina_hpv_entre_9_e_14_anos", "vacina_hpv", "c7_b_ok"],
            "c7_c_ok": ["atend_saude_reprodutiva", "atendimento_saude_reprodutiva", "saude_sexual_reprodutiva", "c7_c_ok"],
            "c7_d_ok": ["rast_cancer_de_mama", "rast_cancer_da_mama", "mamografia", "c7_d_ok"],
        }
        for target, candidates in c7_map.items():
            src = first_existing(df, candidates)
            if src is not None:
                df[target] = to_bool(df[src])
            elif target not in df.columns:
                df[target] = False

        age = df["idade"]
        df["c7_a_applicable"] = age.between(25, 64, inclusive="both")
        df["c7_b_applicable"] = age.between(9, 14, inclusive="both")
        df["c7_c_applicable"] = age.between(14, 69, inclusive="both")
        df["c7_d_applicable"] = age.between(50, 69, inclusive="both")

        df["c7_a_ok"] = df["c7_a_ok"] & df["c7_a_applicable"]
        df["c7_b_ok"] = df["c7_b_ok"] & df["c7_b_applicable"]
        df["c7_c_ok"] = df["c7_c_ok"] & df["c7_c_applicable"]
        df["c7_d_ok"] = df["c7_d_ok"] & df["c7_d_applicable"]

    return df

def preprocess_c2_visits(df: pd.DataFrame) -> pd.DataFrame:
    # Garantir que colunas de visita domiciliar do C2 existam com nomes padrão
    # após a normalização de cabeçalhos.

    cols = list(df.columns)

    def is_c2_visit_1m(col: str) -> bool:
        return (
            "visita" in col
            and "domiciliar" in col
            and "mes" in col
            and (
                "1_mes" in col
                or "1o_mes" in col
                or "1_mes_de_vida" in col
                or "primeiro_mes" in col
            )
        )

    def is_c2_visit_6m(col: str) -> bool:
        return (
            "visita" in col
            and "domiciliar" in col
            and "mes" in col
            and (
                "6_mes" in col
                or "6o_mes" in col
                or "6_mes_de_vida" in col
                or "sexto_mes" in col
            )
        )

    v1_candidates = [c for c in cols if is_c2_visit_1m(c)]
    v6_candidates = [c for c in cols if is_c2_visit_6m(c)]

    if not v1_candidates:
        v1_candidates = [
            c for c in cols
            if "visita" in c and "domiciliar" in c and "1" in c and "mes" in c
        ]

    if not v6_candidates:
        v6_candidates = [
            c for c in cols
            if "visita" in c and "domiciliar" in c and "6" in c and "mes" in c
        ]

    if v1_candidates and "visita_domiciliar_1_mes" not in df.columns:
        df["visita_domiciliar_1_mes"] = df[v1_candidates[0]]

    if v6_candidates and "visita_domiciliar_6_mes" not in df.columns:
        df["visita_domiciliar_6_mes"] = df[v6_candidates[0]]

    return df


def preprocess_c3_puerperio_visits(df: pd.DataFrame) -> pd.DataFrame:
    """Garantir coluna padronizada para visitas domiciliares no puerpério (C3 - J).

    Depois de normalize_col, vamos procurar qualquer coluna que pareça
    "visitas domiciliares (ACS/TACS) Puerpério" e copiar para
    "visitas_domiciliares_acs_tacs_puerperio" se ainda não existir.
    """

    cols = list(df.columns)

    # Candidatos: nome contendo "visita", "domiciliar", "acs"/"tacs" e "puerperio"
    candidates = [
        c
        for c in cols
        if "visita" in c
        and "domiciliar" in c
        and ("acs" in c or "tacs" in c)
        and "puerperio" in c
    ]

    if candidates and "visitas_domiciliares_acs_tacs_puerperio" not in df.columns:
        src = candidates[0]
        df["visitas_domiciliares_acs_tacs_puerperio"] = df[src]

    return df



# =========================
# Cálculos
# =========================


def calculate_score_indicator(df: pd.DataFrame, spec: IndicatorSpec) -> pd.DataFrame:
    df = df.copy()
    weights = spec.weights or {}
    for c in list(weights.keys()):
        ensure_column(df, c, False)

    total_score = np.zeros(len(df), dtype=float)
    total_pendencias = np.zeros(len(df), dtype=int)

    for col, weight in weights.items():
        pratica_ok = to_bool(df[col])
        total_score += np.where(pratica_ok, weight, 0)
        total_pendencias += np.where(~pratica_ok, 1, 0)

    df["score"] = total_score
    df["pendencias"] = total_pendencias
    df["classificacao"] = df["score"].apply(classificar_score)
    return df


def calculate_percentual_indicator(df: pd.DataFrame, spec: IndicatorSpec) -> Tuple[pd.DataFrame, float]:
    df = df.copy()
    num = pd.to_numeric(df[spec.numerator_col], errors="coerce").fillna(0) if spec.numerator_col else pd.Series(0, index=df.index)
    den = pd.to_numeric(df[spec.denominator_col], errors="coerce").fillna(0) if spec.denominator_col else pd.Series(0, index=df.index)
    df["numerador"] = num
    df["denominador"] = den
    total_num = num.sum()
    total_den = den.sum()
    indicador = (total_num / total_den * 100) if total_den > 0 else 0
    df["score"] = np.where(den > 0, (num / den) * 100, 0)
    df["classificacao"] = df["score"].apply(classificar_score)
    df["pendencias"] = np.where(num > 0, 0, 1)
    return df, indicador


def build_good_practices_df(df: pd.DataFrame, spec: IndicatorSpec) -> pd.DataFrame:
    rows = []
    weights = spec.weights or {}
    age_rules = {"A": (25, 64), "B": (9, 14), "C": (14, 69), "D": (50, 69)} if spec.code == "C7" else {}

    for col, peso in weights.items():
        if col not in df.columns:
            continue
        subset = df
        letra = label_boa_pratica(spec.code, col)[:1].upper()
        if spec.code == "C7" and letra in age_rules and "idade" in df.columns:
            lo, hi = age_rules[letra]
            subset = df[df["idade"].between(lo, hi, inclusive="both")].copy()
        total = len(subset)
        realizados = int(to_bool(subset[col]).sum())
        nao_realizados = max(total - realizados, 0)
        perc = round((realizados / total) * 100, 1) if total else 0.0
        rows.append(
            {
                "Boa prática": label_boa_pratica(spec.code, col),
                "coluna": col,
                "Peso": peso,
                "Realizados": realizados,
                "% Realizado": perc,
                "Não realizado": nao_realizados,
            }
        )
    return pd.DataFrame(rows)

# =========================
# Filtros
# =========================


def apply_global_filters(df: pd.DataFrame, spec: IndicatorSpec) -> Tuple[pd.DataFrame, Optional[str]]:
    with st.sidebar:
        st.header("Filtros do painel")
        equipes = sorted([str(e) for e in df.get("equipe", pd.Series(dtype=str)).dropna().unique() if str(e).strip()])
        microareas = sorted([str(m) for m in df.get("micro_area", pd.Series(dtype=str)).dropna().unique() if str(m).strip()])
        faixas = sorted([str(f) for f in df.get("faixa_etaria", pd.Series(dtype=str)).dropna().unique() if str(f).strip()])

        eq_sel = st.multiselect("Por equipe", equipes)
        ma_sel = st.multiselect("Por microárea", microareas)
        fx_sel = st.multiselect("Por faixa etária", faixas)

    out = df.copy()
    if eq_sel:
        out = out[out["equipe"].astype(str).isin(eq_sel)]
    if ma_sel:
        out = out[out["micro_area"].astype(str).isin(ma_sel)]
    if fx_sel:
        out = out[out["faixa_etaria"].astype(str).isin(fx_sel)]

    return out, None

# =========================
# Renderização
# =========================


def render_good_practices(df: pd.DataFrame, spec: IndicatorSpec):
    bp_df = build_good_practices_df(df, spec)
    st.markdown("### Cumprimento das boas práticas")
    if bp_df.empty:
        st.info("Não foi possível identificar boas práticas estruturadas para este relatório.")
        return

    bp_df_display = bp_df.copy()
    if "% Realizado" in bp_df_display.columns:
        bp_df_display["% Realizado"] = bp_df_display["% Realizado"].map(
            lambda v: f"{v:.1f}%" if pd.notna(v) else ""
        )

    st.dataframe(
        bp_df_display[["Boa prática", "Peso", "Realizados", "% Realizado", "Não realizado"]],
        use_container_width=True,
    )

    team_display = "não identificada"

    if "equipe_area" in df.columns and df["equipe_area"].notna().any():
        teams = df["equipe_area"].astype(str).map(clean_team_name)
        teams = teams[teams != ""]
        if not teams.empty:
            team_display = teams.value_counts().idxmax()
    
    elif "equipe" in df.columns and df["equipe"].notna().any():
        teams = df["equipe"].astype(str).map(clean_team_name)
        teams = teams[teams != ""]
        if not teams.empty:
            team_display = teams.value_counts().idxmax()



    data_exportacao = datetime.now().strftime("%d/%m/%Y")
    titulo_export = f"Cumprimento das boas práticas - {team_display} - {data_exportacao}"

    st.download_button(
        "Baixar Relatório das Boas Práticas",
        data=export_excel_bytes(
            bp_df[["Boa prática", "Peso", "Realizados", "% Realizado", "Não realizado"]],
            title=titulo_export,
        ),
        file_name=(
            f"cumprimento_boas_praticas_"
            f"{friendly_indicator_name(spec)}_"
            f"{friendly_team_name(df)}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{spec.code}_boas_praticas_xlsx",
    )

def export_excel_bytes(df: pd.DataFrame, title: Optional[str] = None) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        startrow = 0
        if title:
            pd.DataFrame([[title]]).to_excel(
                writer,
                index=False,
                header=False,
                sheet_name="dados",
                startrow=0
            )
            startrow = 2

        df.to_excel(writer, index=False, sheet_name="dados", startrow=startrow)

        ws = writer.sheets["dados"]

        for idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(idx)
            max_len = len(str(col_name))

            for value in df[col_name].astype(str).fillna(""):
                max_len = max(max_len, len(value))

            adjusted_width = min(max_len + 2, 60)
            ws.column_dimensions[col_letter].width = adjusted_width

    buffer.seek(0)
    return buffer.read()

# =========================
# Vacinação infantil (C2)
# =========================

VACCINE_COL_MAP = {
    "Vacina Pentavalente": "vacina_pentavalente",
    "Vacina Pólio Injetável": "vacina_polio_injetavel",
    "Vacina Sarampo, Caxumba e Rubéola": "vacina_sarampo_caxumba_e_rubeola",
    "Vacina Pneumocócica": "vacina_pneumococica",
}


def build_vaccination_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for label, raw_col in VACCINE_COL_MAP.items():
        if raw_col not in df.columns:
            continue
        series = df[raw_col].astype(str).str.strip().str.lower()
        realizados = int(
            series.isin(["s", "sim", "1", "true", "ok", "x", "yes"]).sum()
        )
        pendentes = int(
            series.isin(["n", "nao", "não", "0", "false"]).sum()
        )
        total = realizados + pendentes
        perc = (realizados / total * 100) if total > 0 else 0.0
        rows.append(
            {
                "Vacina": label,
                "Realizados": realizados,
                "Pendentes": pendentes,
                "% realizado": round(perc, 1),
            }
        )
    return pd.DataFrame(rows)


def build_vaccination_pending_df(df: pd.DataFrame) -> pd.DataFrame:
    # Pacientes com pelo menos uma vacina pendente
    mask_any_pending = pd.Series(False, index=df.index)
    for raw_col in VACCINE_COL_MAP.values():
        if raw_col not in df.columns:
            continue
        series = df[raw_col].astype(str).str.strip().str.lower()
        mask_any_pending = mask_any_pending | series.isin(
            ["n", "nao", "não", "0", "false"]
        )

    base_cols = [
        "nome",
        "cpf",
        "cns",
        "idade",
        "faixa_etaria",
        "endereco",
        "equipe",
        "micro_area",
    ]
    cols_present = [c for c in base_cols if c in df.columns]
    vaccine_cols_present = [
        c for c in VACCINE_COL_MAP.values() if c in df.columns
    ]

    pending_df = df[mask_any_pending].copy()
    pending_df = pending_df[cols_present + vaccine_cols_present]

    # Renomear cabeçalhos das vacinas para rótulos amigáveis
    rename_map = {
        raw: label
        for label, raw in VACCINE_COL_MAP.items()
        if raw in pending_df.columns
    }
    pending_df = pending_df.rename(columns=rename_map)
    return pending_df


def render_vaccination_section(df: pd.DataFrame):
    st.markdown("### Vacinação infantil - pendências e cobertura")

    summary_df = build_vaccination_summary(df)
    pending_df = build_vaccination_pending_df(df)

    if summary_df.empty:
        st.info(
            "Não foi possível identificar colunas de vacinação infantil neste relatório."
        )
        return

    # Tabela de resumo por vacina
    st.subheader("Resumo por vacina")
    display_summary = summary_df.copy()
    display_summary["% realizado"] = display_summary["% realizado"].map(
        lambda v: f"{v:.1f}%" if pd.notna(v) else ""
    )
    st.dataframe(display_summary, use_container_width=True)

    # Gráfico de barras de cobertura
    fig = px.bar(
        summary_df,
        x="Vacina",
        y="% realizado",
        text="% realizado",
        title="Percentual de crianças com esquema realizado por vacina",
    )
    fig.update_layout(xaxis_title="Vacina", yaxis_title="% realizado")
    st.plotly_chart(fig, use_container_width=True)

    # Lista nominal de pacientes com alguma vacina pendente
    st.subheader("Lista de pacientes com vacinas pendentes")
    st.dataframe(pending_df, use_container_width=True, height=360)
    st.caption(
        f"Total de pacientes com alguma vacina pendente: {len(pending_df)}"
    )

    # Exportar CSV/Excel da lista de pacientes com vacinas pendentes (geral)
    csv_bytes = pending_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar CSV - pendências de vacinação (geral)",
        data=csv_bytes,
        file_name="pendencias_vacinacao_geral.csv",
        mime="text/csv",
        key="c2_vacinas_csv_geral",
    )

    st.download_button(
        "Baixar Excel - pendências de vacinação (geral)",
        data=export_excel_bytes(pending_df),
        file_name="pendencias_vacinacao_geral.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="c2_vacinas_xlsx_geral",
    )

    # Exportar por vacina (pendentes específicos)
    st.markdown("#### Exportar pendências específicas por vacina")
    for label, raw_col in VACCINE_COL_MAP.items():
        if raw_col not in df.columns:
            continue
        series = df[raw_col].astype(str).str.strip().str.lower()
        mask_pending = series.isin(["n", "nao", "não", "0", "false"])
        pend_df_vac = df[mask_pending].copy()
        if pend_df_vac.empty:
            continue

        # Usar as mesmas colunas da lista geral, quando existirem
        cols_for_vac = [
            c for c in pending_df.columns if c in pend_df_vac.columns
        ]
        pend_df_vac = pend_df_vac[cols_for_vac]

        csv_bytes_v = pend_df_vac.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            f"Baixar CSV - pendências de {label}",
            data=csv_bytes_v,
            file_name=f"pendencias_{raw_col}.csv",
            mime="text/csv",
            key=f"c2_vacinas_csv_{raw_col}",
        )
        st.download_button(
            f"Baixar Excel - pendências de {label}",
            data=export_excel_bytes(pend_df_vac),
            file_name=f"pendencias_{raw_col}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"c2_vacinas_xlsx_{raw_col}",
        )

def render_c7_age_dashboard(df: pd.DataFrame):
    age_rows = []
    rules = [
        ("A - 25-64", "c7_a_ok", (25, 64)),
        ("B - 9-14", "c7_b_ok", (9, 14)),
        ("C - 14-69", "c7_c_ok", (14, 69)),
        ("D - 50-69", "c7_d_ok", (50, 69)),
    ]
    for label, col, (lo, hi) in rules:
        if "idade" in df.columns:
            subset = df[df["idade"].between(lo, hi, inclusive="both")].copy()
        else:
            subset = df.iloc[0:0].copy()
        elegiveis = len(subset)
        positivos = int(to_bool(subset[col]).sum()) if col in subset.columns else 0
        age_rows.append(
            {
                "Faixa etária": label,
                "Elegíveis": elegiveis,
                "Boas práticas positivas": positivos,
            }
        )

    age_df = pd.DataFrame(age_rows)
    fig = px.bar(
        age_df,
        x="Faixa etária",
        y=["Elegíveis", "Boas práticas positivas"],
        barmode="group",
        title="Distribuição de pacientes e boas práticas por faixa etária",
        labels={"value": "Quantidade", "variable": "Série"},
    )
    fig.update_layout(xaxis_title="Faixa etária", yaxis_title="Quantidade")
    st.plotly_chart(fig, use_container_width=True)



def render_score_dashboard(df: pd.DataFrame, spec: IndicatorSpec):
    df_scored = calculate_score_indicator(df, spec)

    total = len(df_scored)
    media_score = df_scored["score"].mean() if total > 0 else 0
    desempenho = classificar_score(media_score)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Pacientes", total)
    c2.metric("Score", f"{media_score:.1f}")
    c3.metric("Desempenho", desempenho)

    colg1, colg2 = st.columns(2)

    with colg1:
        if total > 0:
            bp_df = build_good_practices_df(df_scored, spec)
            if not bp_df.empty:
                bp_df = bp_df.copy()
                bp_df["Letra"] = bp_df["Boa prática"].str.extract(r"^([A-Z])", expand=False).fillna("")
                fig_bp = px.bar(
                    bp_df,
                    x="Letra",
                    y="% Realizado",
                    text="% Realizado",
                    title="Percentual de realização por boa prática",
                )
                fig_bp.update_layout(xaxis_title="Boa prática", yaxis_title="%")
                st.plotly_chart(fig_bp, use_container_width=True)

    with colg2:
        class_df = df_scored["classificacao"].value_counts().reset_index()
        class_df.columns = ["Classificação", "Quantidade"]
        fig_class = px.pie(
            class_df,
            names="Classificação",
            values="Quantidade",
            title="Distribuição dos pacientes por faixa de desempenho",
        )
        st.plotly_chart(fig_class, use_container_width=True)

    if spec.code == "C7":
        render_c7_age_dashboard(df_scored)

    render_good_practices(df_scored, spec)
    render_nominal(df_scored, spec)

    # Abaixo da lista nominal, apenas para C2
    if spec.code == "C2":
        render_vaccination_section(df_scored)

def render_percentual_dashboard(df: pd.DataFrame, spec: IndicatorSpec):
    df_calc, indicador = calculate_percentual_indicator(df, spec)
    total = len(df_calc)
    desempenho = classificar_score(indicador)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Pacientes", total)
    c2.metric("Score", f"{indicador:.1f}")
    c3.metric("Desempenho", desempenho)

    if "equipe" in df_calc.columns:
        by_team = (
            df_calc.groupby("equipe", dropna=False)
            .agg(numerador=("numerador", "sum"), denominador=("denominador", "sum"))
            .reset_index()
        )
        by_team["percentual"] = np.where(
            by_team["denominador"] > 0,
            by_team["numerador"] / by_team["denominador"] * 100,
            0,
        )
        st.dataframe(by_team, use_container_width=True)
        fig = px.bar(by_team, x="equipe", y="percentual", title="Indicador por equipe")
        st.plotly_chart(fig, use_container_width=True)

    render_nominal(df_calc, spec)


def render_nominal(df: pd.DataFrame, spec: IndicatorSpec):
    st.markdown("### Lista nominal")

    base_cols = [
        "nome",
        "cpf",
        "cns",
        "data_nascimento",
        "idade",
        "faixa_etaria",
        "endereco",
        "equipe",
        "micro_area",
        "equipe_vinculo",
        "score",
        "classificacao",
        "pendencias",
        "cadastro_ok",
    ]

    indicator_cols_map = {
        "C1": ["numerador", "denominador"],
        "C2": ["c2_a_ok", "c2_b_ok", "c2_c_ok", "c2_d_ok", "c2_e_ok"],
        "C3": ["c3_a_ok", "c3_b_ok", "c3_c_ok", "c3_d_ok", "c3_e_ok", "c3_f_ok", "c3_g_ok", "c3_h_ok", "c3_i_ok", "c3_j_ok", "c3_k_ok"],
        "C4": ["c4_a_ok", "c4_b_ok", "c4_c_ok", "c4_d_ok", "c4_e_ok", "c4_f_ok"],
        "C5": ["c5_a_ok", "c5_b_ok", "c5_c_ok", "c5_d_ok"],
        "C6": ["consulta_ok", "antropometria_ok", "visitas_ok", "influenza_ok"],
        "C7": ["c7_a_ok", "c7_b_ok", "c7_c_ok", "c7_d_ok"],
    }

    cols = [c for c in base_cols if c in df.columns]
    cols += [c for c in indicator_cols_map.get(spec.code, []) if c in df.columns]

    if not cols:
        cols = list(df.columns)

    col_labels = {
        "nome": "Nome",
        "cpf": "CPF",
        "cns": "CNS",
        "data_nascimento": "Data nascimento",
        "idade": "Idade",
        "faixa_etaria": "Faixa etária",
        "endereco": "Endereço",
        "equipe": "Equipe",
        "micro_area": "Microárea",
        "equipe_vinculo": "Equipe vínculo",
        "score": "Score",
        "classificacao": "Classificação",
        "pendencias": "Pendências",
        "cadastro_ok": "Cadastro OK",
        "numerador": "Numerador",
        "denominador": "Denominador",
        "c2_a_ok": "C2 - A",
        "c2_b_ok": "C2 - B",
        "c2_c_ok": "C2 - C",
        "c2_d_ok": "C2 - D",
        "c2_e_ok": "C2 - E",
        "c3_a_ok": "C3 - A",
        "c3_b_ok": "C3 - B",
        "c3_c_ok": "C3 - C",
        "c3_d_ok": "C3 - D",
        "c3_e_ok": "C3 - E",
        "c3_f_ok": "C3 - F",
        "c3_g_ok": "C3 - G",
        "c3_h_ok": "C3 - H",
        "c3_i_ok": "C3 - I",
        "c3_j_ok": "C3 - J",
        "c3_k_ok": "C3 - K",
        "c4_a_ok": "C4 - A",
        "c4_b_ok": "C4 - B",
        "c4_c_ok": "C4 - C",
        "c4_d_ok": "C4 - D",
        "c4_e_ok": "C4 - E",
        "c4_f_ok": "C4 - F",
        "c5_a_ok": "C5 - A",
        "c5_b_ok": "C5 - B",
        "c5_c_ok": "C5 - C",
        "c5_d_ok": "C5 - D",
        "consulta_ok": "Consulta OK",
        "antropometria_ok": "Antropometria OK",
        "visitas_ok": "Visitas OK",
        "influenza_ok": "Influenza OK",
        "c7_a_ok": "C7 - A",
        "c7_b_ok": "C7 - B",
        "c7_c_ok": "C7 - C",
        "c7_d_ok": "C7 - D",
    }

    df_display = df[cols].rename(
        columns={c: col_labels.get(c, c.replace("_", " ").title()) for c in cols}
    )

    bp_df = build_good_practices_df(df, spec)

    if bp_df.empty:
        st.dataframe(df_display, use_container_width=True, height=420)
        st.caption(f"Total de pacientes exibidos: {len(df_display)}")

        csv_bytes = df_display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar CSV filtrado",
            data=csv_bytes,
            file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_team_name(df)}.csv",
            mime="text/csv",
        )

        st.download_button(
            "Baixar Excel filtrado",
            data=export_excel_bytes(df_display),
            file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_team_name(df)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    label_to_col = {}
    letters = []
    for _, row in bp_df.iterrows():
        label = str(row["Boa prática"])
        col = str(row["coluna"])
        letra = label[:1].upper()
        label_to_col[letra] = col
        if letra and letra not in letters:
            letters.append(letra)

    tab_labels = ["Lista nominal"] + [
    f"Pendência {l} - {TAB_SHORT_LABELS.get(spec.code, {}).get(l, l)}"
    for l in letters
    ]

    tabs = st.tabs(tab_labels)

    with tabs[0]:
        st.dataframe(df_display, use_container_width=True, height=420)
        st.caption(f"Total de pacientes exibidos: {len(df_display)}")

        csv_bytes = df_display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar CSV filtrado",
            data=csv_bytes,
            file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_team_name(df)}.csv",
            mime="text/csv",
            key=f"{spec.code}_csv_all",
        )

        st.download_button(
            "Baixar Excel filtrado",
            data=export_excel_bytes(df_display),
            file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_team_name(df)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{spec.code}_xlsx_all",
        )

    c7_age_rules = {"A": (25, 64), "B": (9, 14), "C": (14, 69), "D": (50, 69)} if spec.code == "C7" else {}

    for i, letra in enumerate(letters, start=1):
        col_bp = label_to_col.get(letra)
        if col_bp not in df.columns:
            filtered = df.iloc[0:0].copy()
        else:
            filtered = df[~to_bool(df[col_bp])].copy()
            if spec.code == "C7" and letra in c7_age_rules and "idade" in filtered.columns:
                lo, hi = c7_age_rules[letra]
                filtered = filtered[filtered["idade"].between(lo, hi, inclusive="both")].copy()

        filtered_display = filtered[cols].rename(
            columns={c: col_labels.get(c, c.replace("_", " ").title()) for c in cols}
        )

        with tabs[i]:
            st.dataframe(filtered_display, use_container_width=True, height=420)
            st.caption(f"Total de pacientes exibidos: {len(filtered_display)}")

            csv_bytes = filtered_display.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Baixar CSV filtrado",
                data=csv_bytes,
                file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_pendencia_name(letra)}_{friendly_team_name(df)}.csv",
                mime="text/csv",
                key=f"{spec.code}_csv_{letra}",
            )

            st.download_button(
                "Baixar Excel filtrado",
                data=export_excel_bytes(filtered_display),
                file_name=f"lista_nominal_{friendly_indicator_name(spec)}_{friendly_pendencia_name(letra)}_{friendly_team_name(df)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{spec.code}_xlsx_{letra}",
            )
# =========================
# Aplicação
# =========================


def main():
    st.title("APS 360 - Painel de Indicadores")
    st.caption(
        "Ferramenta de apoio às equipes e à gestão no monitoramento dos indicadores e do cuidado na APS.."
    )

    st.sidebar.header("Importação")
    uploaded_file = st.sidebar.file_uploader(
        "Envie um relatório CSV/XLS/XLSX", type=["csv", "xls", "xlsx"]
    )

    st.sidebar.header("Indicador")
    manual_indicator = st.sidebar.selectbox(
        "Selecionar manualmente (opcional)",
        ["Automático"] + [f"{k} - {v.name}" for k, v in INDICATORS.items()],
    )

    if uploaded_file is None:
        st.info(
            "Envie um relatório para começar."
        )
        st.stop()

    try:
        df_raw = read_uploaded_file(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()

    detected = detect_indicator_from_columns(
        pd.DataFrame(columns=[normalize_col(c) for c in df_raw.columns]),
        uploaded_file.name,
    )
    selected_code = (
        manual_indicator.split(" ")[0]
        if manual_indicator != "Automático"
        else detected
    )

    if selected_code is None:
        st.warning(
            "Não foi possível identificar automaticamente o indicador. "
            "Escolha manualmente na barra lateral."
        )
        st.stop()

    spec = INDICATORS[selected_code]
    df = preprocess_df(df_raw, selected_code)

    df_filtered, _ = apply_global_filters(df, spec)

    team_display = None

    if "equipe_area" in df_filtered.columns and df_filtered["equipe_area"].notna().any():
        teams = (
            df_filtered["equipe_area"]
            .astype(str)
            .map(clean_team_name)
        )
        teams = teams[teams != ""]
    
        if not teams.empty:
            team_display = teams.value_counts().idxmax()
    
    elif "equipe" in df_filtered.columns and df_filtered["equipe"].notna().any():
        teams = (
            df_filtered["equipe"]
            .astype(str)
            .map(clean_team_name)
        )
        teams = teams[teams != ""]
    
        if not teams.empty:
            team_display = teams.value_counts().idxmax()
    
    if team_display:
        st.success(f"Equipe em análise: {team_display}")
    else:
        st.success("Equipe em análise: não identificada")
    
    st.markdown(f"## {spec.code} - {spec.name}")
    st.write(spec.description)
    if spec.type == "score":
        render_score_dashboard(df_filtered, spec)
    else:
        render_percentual_dashboard(df_filtered, spec)


if __name__ == "__main__":
    main()
