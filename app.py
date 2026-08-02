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
from streamlit_extras.stylable_container import stylable_container
from st_aggrid import AgGrid, GridOptionsBuilder
import requests
import time
import pydeck as pdk

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


def friendly_indicator_name(spec: "IndicatorSpec") -> str:
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
    return BOA_PRATICA_LABELS.get(indicator_code, {}).get(
        col, col.replace("_", " ").capitalize()
    )


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
# Geocodificação com Nominatim
# =========================


@st.cache_data(show_spinner=False)
def geocode_address_nominatim(endereco: str, cidade: str = "", uf: str = "") -> Tuple[Optional[float], Optional[float]]:
    """Converte endereço em latitude/longitude usando Nominatim (OpenStreetMap)."""
    if not endereco or str(endereco).strip() == "":
        return None, None

    query = str(endereco).strip()
    if cidade:
        query += f", {cidade}"
    if uf:
        query += f", {uf}, Brasil"

    url = "https://nominatim.openstreetmap.org/search"
    params = {
        "q": query,
        "format": "json",
        "limit": 1,
        "addressdetails": 0,
    }

    try:
        headers = {"User-Agent": "aps360-painel/1.0"}
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if not data:
            return None, None
        lat = float(data[0]["lat"])
        lon = float(data[0]["lon"])
        # Pequeno intervalo para respeitar o serviço em chamadas sequenciais
        time.sleep(1)
        return lat, lon
    except Exception:
        return None, None


def build_geocoded_df(df_tab: pd.DataFrame, cidade: str = "", uf: str = "") -> pd.DataFrame:
    """Cria dataframe com Nome, Endereço e coordenadas para o mapa."""
    if "Endereço" not in df_tab.columns:
        st.warning("Coluna 'Endereço' não encontrada para gerar o mapa.")
        return pd.DataFrame(columns=["Nome", "Endereço", "latitude", "longitude"])

    rows = []
    for _, row in df_tab.iterrows():
        nome = row.get("Nome", "")
        endereco = row.get("Endereço", "")
        lat, lon = geocode_address_nominatim(endereco, cidade=cidade, uf=uf)
        if lat is not None and lon is not None:
            rows.append(
                {
                    "Nome": nome,
                    "Endereço": endereco,
                    "latitude": lat,
                    "longitude": lon,
                    "Score": row.get("Score", None),
                    "Equipe": row.get("Equipe", None),
                }
            )

    return pd.DataFrame(rows)


def render_maps_for_df(df_tab: pd.DataFrame, cidade: str = "", uf: str = "", map_key: str = "geral"):
    st.markdown("#### Mapa dos pacientes")

    df_geo = build_geocoded_df(df_tab, cidade=cidade, uf=uf)
    if df_geo.empty:
        st.info("Nenhum endereço foi geocodificado para exibir no mapa.")
        return

    tipo_mapa = st.radio(
        "Tipo de mapa",
        ["Pontos", "Mapa de calor"],
        horizontal=True,
        key=f"tipo_mapa_{map_key}",
    )

    if tipo_mapa == "Pontos":
        # st.map espera colunas lat/lon
        st.map(
            df_geo[["latitude", "longitude"]].rename(
                columns={"latitude": "lat", "longitude": "lon"}
            )
        )
    else:
        layer = pdk.Layer(
            "HeatmapLayer",
            data=df_geo,
            get_position="[longitude, latitude]",
            aggregation="SUM",
            get_weight="1",
            radiusPixels=30,
        )
        view_state = pdk.ViewState(
            latitude=df_geo["latitude"].mean(),
            longitude=df_geo["longitude"].mean(),
            zoom=12,
            pitch=0,
        )
        deck = pdk.Deck(
            layers=[layer],
            initial_view_state=view_state,
            tooltip={"text": "{Nome}"},
        )
        st.pydeck_chart(deck)


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


# (pré-processamento igual ao app-12.py original; omitido aqui por brevidade)
# Copie toda a função preprocess_df, preprocess_c2_visits, preprocess_c3_puerperio_visits
# exatamente como estão no seu arquivo atual.

# [...]  (cole aqui o bloco completo de pré-processamento do app-12)


# =========================
# Cálculos
# =========================


# (idem: mantenha calculate_score_indicator, calculate_percentual_indicator,
# build_good_practices_df exatamente como no app-12.)

# [...]  (cole aqui o bloco completo de cálculos)


# =========================
# Filtros
# =========================


# (mantenha apply_global_filters igual ao app-12)

# [...]  (cole aqui apply_global_filters)


# =========================
# Renderização: boas práticas, vacinação, C7 etc.
# =========================


# (mantenha render_good_practices, export_excel_bytes,
#  VACCINE_COL_MAP + funções de vacinação,
#  render_c7_age_dashboard, render_score_dashboard,
#  render_percentual_dashboard exatamente como estão.)

# [...]  (cole aqui esses blocos sem alterações)


# =========================
# Lista nominal (AgGrid) + mapas
# =========================


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
        "C3": [
            "c3_a_ok",
            "c3_b_ok",
            "c3_c_ok",
            "c3_d_ok",
            "c3_e_ok",
            "c3_f_ok",
            "c3_g_ok",
            "c3_h_ok",
            "c3_i_ok",
            "c3_j_ok",
            "c3_k_ok",
        ],
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

    label_to_col = {}
    letras = []

    for _, row in bp_df.iterrows():
        label = str(row["Boa prática"])
        col = str(row["coluna"])
        letra = label[:1].upper()
        label_to_col[letra] = col
        if letra and letra not in letras:
            letras.append(letra)

    tab_labels = ["Lista geral"] + [
        f"Pendência {l} - {TAB_SHORT_LABELS.get(spec.code, {}).get(l, l)}" for l in letras
    ]
    tabs = st.tabs(tab_labels)

    # Tab 0: lista nominal completa + mapa
    with tabs[0]:
        gb = GridOptionsBuilder.from_dataframe(df_display)
        gb.configure_default_column(
            filter=True,
            sortable=True,
            resizable=True,
            minWidth=100,
        )
        gb.configure_column("Nome", width=300, minWidth=300)
        gb.configure_column("Idade", width=60, minWidth=60)
        gb.configure_column("Score", width=70, minWidth=70)
        gb.configure_column("Faixa etária", width=70, minWidth=70)
        gb.configure_column("Equipe", width=90, minWidth=90)
        gb.configure_side_bar()
        grid_options = gb.build()

        AgGrid(
            df_display,
            gridOptions=grid_options,
            height=420,
            enable_enterprise_modules=False,
            pagination=True,
            paginationPageSize=25,
        )
        st.caption(f"Total de pacientes exibidos: {len(df_display)}")

        csv_bytes = df_display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar CSV filtrado",
            data=csv_bytes,
            file_name=(
                f"lista_nominal_{friendly_indicator_name(spec)}_"
                f"{friendly_team_name(df)}.csv"
            ),
            mime="text/csv",
            key=f"{spec.code}_csv_all",
        )

        st.download_button(
            "Baixar Excel filtrado",
            data=export_excel_bytes(df_display),
            file_name=(
                f"lista_nominal_{friendly_indicator_name(spec)}_"
                f"{friendly_team_name(df)}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{spec.code}_xlsx_all",
        )

        # Mapa para a lista geral
        render_maps_for_df(df_display, cidade="", uf="", map_key=f"{spec.code}_geral")

    # Demais tabs: listas de pendência por letra + mapa
    c7_age_rules = {
        "A": (25, 64),
        "B": (9, 14),
        "C": (14, 69),
        "D": (50, 69),
    } if spec.code == "C7" else {}

    for i, letra in enumerate(letras, start=1):
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
            gb_f = GridOptionsBuilder.from_dataframe(filtered_display)
            gb_f.configure_default_column(
                filter=True,
                sortable=True,
                resizable=True,
                minWidth=100,
            )
            gb_f.configure_column("Nome", width=300, minWidth=300)
            gb_f.configure_column("Idade", width=60, minWidth=60)
            gb_f.configure_column("Score", width=70, minWidth=70)
            gb_f.configure_column("Faixa etária", width=70, minWidth=70)
            gb_f.configure_column("Equipe", width=100, minWidth=10)
            gb_f.configure_side_bar()
            grid_options_f = gb_f.build()

            AgGrid(
                filtered_display,
                gridOptions=grid_options_f,
                height=420,
                enable_enterprise_modules=False,
                pagination=True,
                paginationPageSize=25,
            )
            st.caption(f"Total de pacientes exibidos: {len(filtered_display)}")

            csv_bytes_f = filtered_display.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Baixar CSV filtrado",
                data=csv_bytes_f,
                file_name=(
                    f"lista_nominal_{friendly_indicator_name(spec)}_"
                    f"{friendly_pendencia_name(letra)}_"
                    f"{friendly_team_name(df)}.csv"
                ),
                mime="text/csv",
                key=f"{spec.code}_csv_{letra}",
            )

            st.download_button(
                "Baixar Excel filtrado",
                data=export_excel_bytes(filtered_display),
                file_name=(
                    f"lista_nominal_{friendly_indicator_name(spec)}_"
                    f"{friendly_pendencia_name(letra)}_"
                    f"{friendly_team_name(df)}.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{spec.code}_xlsx_{letra}",
            )

            # Mapa para a tab de pendência
            render_maps_for_df(
                filtered_display,
                cidade="",
                uf="",
                map_key=f"{spec.code}_pendencia_{letra}",
            )


# =========================
# Aplicação
# =========================


def main():
    st.title("APS 360 - Painel de Indicadores")
    st.caption(
        "Ferramenta de apoio às equipes e à gestão no monitoramento dos indicadores e do cuidado na APS."
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
        st.info("Envie um relatório para começar.")
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
    if "equipe_vinculo" in df_filtered.columns:
        vals = [
            clean_team_name(v)
            for v in df_filtered["equipe_vinculo"].dropna().astype(str)
            if clean_team_name(v)
        ]
        uniq = sorted(set(vals))
        if len(uniq) == 1:
            team_display = uniq[0]
        elif len(uniq) > 1:
            team_display = " / ".join(uniq)

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
