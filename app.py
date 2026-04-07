# Garimpeiro — código-fonte único: faça todas as alterações neste ficheiro (app2.py).
# app.py só reencaminha para aqui quando a Cloud usa app.py como entrada.
import os
import sys


def _garimpeiro_forcar_temp_na_pasta_do_projeto():
    """
    TEMP/TMP e PIP_CACHE_DIR ficam na pasta do app2.py (ex.: disco D:).
    Assim Excel/xlsxwriter e ferramentas Python não dependem de C:\\Users\\…\\Temp
    (crítico quando C: está sem espaço). Corre antes dos restantes imports.
    """
    try:
        root = os.path.dirname(os.path.abspath(__file__))
        tmp = os.path.join(root, "temp_windows_ambiente")
        os.makedirs(tmp, exist_ok=True)
        os.environ["TEMP"] = tmp
        os.environ["TMP"] = tmp
        pipc = os.path.join(root, "pip_cache_local")
        os.makedirs(pipc, exist_ok=True)
        os.environ["PIP_CACHE_DIR"] = pipc
    except OSError:
        pass


_garimpeiro_forcar_temp_na_pasta_do_projeto()

import streamlit as st
import zipfile
import io
from contextlib import contextmanager
import hashlib
import re
import pandas as pd
import random
import gc
import shutil
from collections import Counter, defaultdict
from calendar import monthrange
from datetime import date, datetime
import unicodedata
import json
from pathlib import Path


def _erro_sem_espaco_disco(exc: BaseException) -> bool:
    """errno 28 / temp ou disco cheio (xlsxwriter FileCreateError envolve OSError)."""
    if exc is None:
        return False
    o: BaseException | None = exc
    for _ in range(5):
        if isinstance(o, OSError) and getattr(o, "errno", None) == 28:
            return True
        if type(o).__name__ == "FileCreateError" and getattr(o, "args", None):
            o = o.args[0] if isinstance(o.args[0], BaseException) else None
            if o is None:
                break
            continue
        break
    s = str(exc).lower()
    return "errno 28" in s or "no space left" in s or "no space" in s


def _msg_sem_espaco_disco_garimpeiro() -> str:
    return (
        "ERR:Sem espaço em disco (errno 28). O Garimpeiro força **TEMP/TMP** na pasta do projeto "
        "**temp_windows_ambiente** (mesmo disco que o app2.py). Liberte espaço **nesse disco** e apague ficheiros antigos nessa pasta se precisar. "
        "Se ainda aparecer C: no erro, confirme que abriu a app com **streamlit run app2.py** (o arranque do Python tem de carregar o início do ficheiro)."
    )


_AGGRID_LOCALE_PT_BR_CACHE = None


def _aggrid_locale_pt_br():
    """Textos do AG Grid em pt-BR (filtros, menus) — ficheiro gerado a partir do locale oficial MIT."""
    global _AGGRID_LOCALE_PT_BR_CACHE
    if _AGGRID_LOCALE_PT_BR_CACHE is not None:
        return _AGGRID_LOCALE_PT_BR_CACHE
    p = Path(__file__).resolve().parent / "ag_grid_locale_pt_br.json"
    if p.is_file():
        with p.open(encoding="utf-8") as fp:
            _AGGRID_LOCALE_PT_BR_CACHE = json.load(fp)
    else:
        _AGGRID_LOCALE_PT_BR_CACHE = {}
    return _AGGRID_LOCALE_PT_BR_CACHE


def _instrucoes_instalar_fpdf2_markdown():
    """Streamlit corre com o interpretador em sys.executable — fpdf2 tem de estar aí."""
    exe = sys.executable or "python"
    cmd = f'"{exe}" -m pip install fpdf2'
    cloud = "/home/adminuser/" in exe or "/mount/src/" in exe
    if cloud:
        return (
            "No **Streamlit Community Cloud** não é possível instalar pacotes a partir da app — o ambiente "
            "é montado só com o que está no repositório.\n\n"
            "1. Confirme que na **raiz do repositório** existe `requirements.txt` com a linha **`fpdf2>=2.7.0`** "
            "(o projeto Garimpeiro já a inclui).\n"
            "2. Faça **commit** e **push** desse ficheiro para o ramo que a Cloud usa.\n"
            "3. No painel [share.streamlit.io](https://share.streamlit.io), abra a app → **⋮** → "
            "**Reboot app** (ou desligue e volte a publicar) para reinstalar dependências.\n\n"
            "Se a app aponta para uma **pasta** dentro do repo, a Cloud continua a ler `requirements.txt` "
            "só da **raiz** — não pode haver outro `requirements.txt` sem `fpdf2` a substituir."
        )
    return (
        "O pacote **fpdf2** não está instalado no **mesmo Python** que está a executar o Streamlit.\n\n"
        f"No terminal, corra:\n\n`{cmd}`\n\n"
        "Se usa **ambiente virtual**, ative-o antes desse comando, instale e **volte a iniciar** a app "
        "(`streamlit run …`). O ficheiro **requirements.txt** já inclui `fpdf2` — também pode usar "
        "`pip install -r requirements.txt` nesse ambiente."
    )


def _df_resumo_para_exibicao_sem_separador_milhar(df):
    """Streamlit formata inteiros grandes com separador de milhares; nºs de nota/série não devem ter vírgula."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    out = df.copy()

    def _int_str(x):
        if pd.isna(x):
            return ""
        try:
            return str(int(round(float(x))))
        except (TypeError, ValueError):
            return str(x)

    for col in ("Início", "Fim", "Quantidade"):
        if col in out.columns:
            out[col] = out[col].map(_int_str)
    if "Série" in out.columns:
        def _serie_plain(x):
            if pd.isna(x):
                return ""
            s = str(x).strip()
            if s.endswith(".0") and len(s) > 2 and s[:-2].replace("-", "").isdigit():
                return s[:-2]
            return s

        out["Série"] = out["Série"].map(_serie_plain)
    return out


def _df_terceiros_por_tipo_para_exibicao_sem_separador_milhar(df):
    """Contagem por tipo de documento (só inteiros); evita vírgula/ponto de milhar no st.dataframe."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    out = df.copy()
    if "Quantidade" not in out.columns:
        return out

    def _q_str(x):
        if pd.isna(x):
            return ""
        try:
            return str(int(round(float(x))))
        except (TypeError, ValueError):
            return str(x)

    out["Quantidade"] = out["Quantidade"].map(_q_str)
    return out


def _df_relatorio_leitura_abas_para_exibicao_sem_sep_milhar(df):
    """
    Abas do relatório da leitura: nº de nota e «número em falta» não devem aparecer com separador de milhares
    no st.dataframe (são identificadores, não quantias).
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    out = df.copy()

    def _nota_ou_gap_str(x):
        if pd.isna(x):
            return ""
        try:
            return str(int(round(float(x))))
        except (TypeError, ValueError):
            s = str(x).strip()
            return s if s and s.lower() != "nan" else ""

    for col in ("Nota", "Num_Faltante"):
        if col in out.columns:
            out[col] = out[col].map(_nota_ou_gap_str)

    if "Série" in out.columns:

        def _serie_plain_sl(x):
            if pd.isna(x):
                return ""
            s = str(x).strip()
            if s.endswith(".0") and len(s) > 2 and s[:-2].replace("-", "").isdigit():
                return s[:-2]
            return s

        out["Série"] = out["Série"].map(_serie_plain_sl)

    if "Data Emissão" in out.columns:
        out["Data Emissão"] = out["Data Emissão"].map(_valor_data_emissao_dd_mm_yyyy)

    return out


def _valor_data_emissao_dd_mm_yyyy(x):
    """Valor de data → texto dd/mm/aaaa (exibição Streamlit, Excel e PDF)."""
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(x, pd.Timestamp):
        if pd.isna(x):
            return ""
        return x.strftime("%d/%m/%Y")
    if isinstance(x, datetime):
        return x.strftime("%d/%m/%Y")
    if isinstance(x, date):
        return x.strftime("%d/%m/%Y")
    s = str(x).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return ""
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s
    ts = pd.to_datetime(s, errors="coerce")
    if pd.notna(ts):
        return ts.strftime("%d/%m/%Y")
    return s


def _df_com_data_emissao_dd_mm_yyyy(df):
    """Cópia com coluna «Data Emissão» em dd/mm/aaaa, se existir."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    if "Data Emissão" not in df.columns:
        return df
    out = df.copy()
    out["Data Emissão"] = out["Data Emissão"].map(_valor_data_emissao_dd_mm_yyyy)
    return out


def _garim_emoji(grapheme: str) -> str:
    """Envolve um emoji em span — só o ícone fica maior (.garim-emoji); usar com unsafe_allow_html=True."""
    if not grapheme:
        return ""
    return f'<span class="garim-emoji" aria-hidden="true">{grapheme}</span>'


# --- CONFIGURAÇÃO E ESTILO (CLONE ABSOLUTO DO DIAMOND TAX) ---
st.set_page_config(page_title="Garimpeiro", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');

        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { 
            background: radial-gradient(circle at top right, #FFDEEF 0%, #F8F9FA 100%) !important;
            /* Primária do tema: Streamlit usa isto em multiselect, checkbox, focos, etc. */
            --primary-color: #ff69b4 !important;
            --accent-color: #ff69b4 !important;
        }
        /* Reforço no contentor principal (algumas versões lêem a variável aqui) */
        section.main .block-container {
            --primary-color: #ff69b4 !important;
            padding-top: 1.1rem !important;
            padding-left: clamp(1rem, 2.5vw, 2rem) !important;
            padding-right: clamp(1rem, 2.5vw, 2rem) !important;
            max-width: min(1680px, 100%) !important;
        }
        /* Chips / etiquetas dos multiselects (Base Web) — continuavam vermelhos com o tema por defeito */
        span[data-baseweb="tag"] {
            background-color: #ff69b4 !important;
            color: #ffffff !important;
            border-color: #f06292 !important;
        }
        span[data-baseweb="tag"] svg,
        span[data-baseweb="tag"] path {
            fill: #ffffff !important;
        }
        /* Opções assinaladas ao abrir o multiselect */
        li[role="option"][aria-selected="true"],
        [role="listbox"] [aria-selected="true"] {
            background-color: rgba(255, 105, 180, 0.2) !important;
        }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #FFDEEF !important;
            min-width: min(272px, 100vw) !important;
            max-width: min(380px, 100vw) !important;
        }
        /* Colunas na lateral: sem min-width por defeito do flex = conteúdo cortado */
        [data-testid="stSidebar"] [data-testid="column"] {
            min-width: 0 !important;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div {
            min-width: 0 !important;
        }
        /* Tabela do editor na lateral: evita cortar conteúdo */
        [data-testid="stSidebar"] [data-testid="stDataFrame"],
        [data-testid="stSidebar"] [data-testid="stDataEditor"] {
            overflow-x: auto !important;
        }
        /* Último nº por série: cartões; scroll horizontal se ainda faltar espaço */
        /* Caixas com borda na lateral (se existirem): traço mínimo, sem sombra */
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] {
            background: transparent !important;
            border: none !important;
            border-left: 2px solid rgba(255, 105, 180, 0.22) !important;
            border-radius: 0 !important;
            padding: 0.2rem 0 0.45rem 0.45rem !important;
            margin-bottom: 0.35rem !important;
            box-shadow: none !important;
            overflow-x: auto !important;
            overflow-y: visible !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
        }
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="select"] > div {
            border-radius: 8px !important;
            border-color: #f0e0e8 !important;
            min-height: 2.05rem !important;
            max-width: 100% !important;
        }
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] input {
            border-radius: 8px !important;
            border-color: #ece8ea !important;
            min-height: 2.05rem !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
        }
        /* Lateral: menos peso visual global */
        [data-testid="stSidebar"] hr {
            margin: 0.65rem 0 !important;
            border: none !important;
            border-top: 1px solid rgba(0, 0, 0, 0.06) !important;
        }
        [data-testid="stSidebar"] div.stButton > button {
            height: auto !important;
            min-height: 2.35rem !important;
            padding: 0.45rem 0.7rem !important;
            font-size: 0.82rem !important;
            font-weight: 600 !important;
            border-radius: 10px !important;
            text-transform: none !important;
            box-shadow: none !important;
        }
        [data-testid="stSidebar"] div.stButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 2px 8px rgba(255, 105, 180, 0.12) !important;
        }
        [data-testid="stSidebar"] div.stDownloadButton > button {
            border-radius: 10px !important;
            text-transform: none !important;
            padding: 0.45rem 0.7rem !important;
            min-height: 2.35rem !important;
            height: auto !important;
            box-shadow: none !important;
            border-width: 1px !important;
            font-size: 0.82rem !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            border: 1px solid rgba(0, 0, 0, 0.06) !important;
            border-radius: 10px !important;
            background: rgba(255, 255, 255, 0.55) !important;
            box-shadow: none !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] details {
            border: none !important;
        }
        /* Séries empilhadas: separador suave entre linhas (sem caixa por série) */
        [data-testid="stSidebar"] .garim-seq-row-spacer {
            border: none !important;
            border-top: 1px solid rgba(0, 0, 0, 0.05) !important;
            margin: 0.2rem 0 0.25rem 0 !important;
        }
        /* Referência de séries (dentro do expander): linhas mais baixas e menos espaço */
        [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="column"] {
            padding-top: 0.1rem !important;
            padding-bottom: 0.1rem !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] > div {
            margin-bottom: 0.2rem !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] [data-baseweb="select"] > div {
            min-height: 1.8rem !important;
            font-size: 0.78rem !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] input {
            min-height: 1.8rem !important;
            font-size: 0.78rem !important;
            padding: 0.22rem 0.35rem !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stWidgetLabel"] p {
            font-size: 0.72rem !important;
            margin-bottom: 0.1rem !important;
        }

        /* Etapa 3: dois painéis (emissão própria | terceiros) */
        div.garim-etapa3-bloco {
            border-radius: 14px !important;
            padding: 0.85rem 1rem 1rem !important;
            margin-bottom: 0.6rem !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
        }
        div.garim-etapa3-propria {
            background: linear-gradient(160deg, #fffafd 0%, #ffffff 55%, #fff8fc 100%) !important;
            border: 1px solid rgba(255, 105, 180, 0.35) !important;
            box-shadow: 0 2px 12px rgba(255, 105, 180, 0.06) !important;
        }
        div.garim-etapa3-terc {
            background: linear-gradient(160deg, #f8fbff 0%, #ffffff 55%, #f3f8ff 100%) !important;
            border: 1px solid rgba(100, 149, 237, 0.35) !important;
            box-shadow: 0 2px 12px rgba(100, 149, 237, 0.08) !important;
        }
        p.garim-etapa3-titulo {
            font-weight: 700 !important;
            font-size: 0.95rem !important;
            margin: 0 0 0.4rem 0 !important;
            color: #5D1B36 !important;
        }
        div.garim-etapa3-terc p.garim-etapa3-titulo {
            color: #1a3a5c !important;
        }
        p.garim-etapa3-sub {
            font-size: 0.78rem !important;
            line-height: 1.45 !important;
            color: #6d5c63 !important;
            margin: 0 !important;
        }
        div.garim-etapa3-terc p.garim-etapa3-sub {
            color: #5a6570 !important;
        }

        div.stButton > button {
            color: #6C757D !important; 
            background-color: #FFFFFF !important; 
            border: 1px solid #DEE2E6 !important;
            border-radius: 15px !important;
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 800 !important;
            height: 60px !important;
            text-transform: uppercase;
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important;
            width: 100% !important;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important;
        }

        div.stButton > button:hover {
            transform: translateY(-5px) !important;
            box-shadow: 0 10px 20px rgba(255,105,180,0.2) !important;
            border-color: #FF69B4 !important;
            color: #FF69B4 !important;
        }

        /* Dentro de expanders: botões mais baixos (lateral, etc.) */
        [data-testid="stExpander"] [data-testid="stButton"] button,
        [data-testid="stExpander"] div.stButton > button {
            min-height: 2rem !important;
            height: auto !important;
            padding: 0.25rem 0.5rem !important;
            font-size: 0.78rem !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            text-transform: none !important;
            line-height: 1.25 !important;
        }
        [data-testid="stExpander"] [data-testid="stButton"] button:hover,
        [data-testid="stExpander"] div.stButton > button:hover {
            transform: translateY(-2px) !important;
        }
        /* Primary do Streamlit (vermelho por defeito) → rosa Garimpeiro */
        [data-testid="stButton"] button[kind="primary"],
        div.stButton > button[kind="primary"] {
            background: linear-gradient(180deg, #ff8cc8, #ff69b4) !important;
            color: #ffffff !important;
            border: 1px solid #f06292 !important;
            box-shadow: 0 2px 12px rgba(255, 105, 180, 0.4) !important;
        }
        [data-testid="stButton"] button[kind="primary"]:hover,
        div.stButton > button[kind="primary"]:hover {
            filter: brightness(1.06) !important;
            color: #ffffff !important;
            border-color: #ec407a !important;
        }

        [data-testid="stFileUploader"] { 
            border: 2px dashed #FF69B4 !important; 
            border-radius: 20px !important;
            background: #FFFFFF !important;
            padding: 20px !important;
        }

        div.stDownloadButton > button {
            background-color: #FF69B4 !important; 
            color: white !important; 
            border: 2px solid #FFFFFF !important;
            font-weight: 700 !important;
            border-radius: 15px !important;
            box-shadow: 0 0 15px rgba(255, 105, 180, 0.3) !important;
            text-transform: uppercase;
            width: 100% !important;
        }

        h1, h2, h3 {
            font-family: 'Montserrat', sans-serif;
            font-weight: 800;
            color: #FF69B4 !important;
            text-align: center;
        }
        /* Só emojis com span.garim-emoji — texto ao lado mantém tamanho normal */
        span.garim-emoji {
            font-size: 1.38em;
            line-height: 1;
            display: inline-block;
            vertical-align: -0.1em;
        }
        h3.garim-sec span.garim-emoji {
            vertical-align: -0.05em;
        }
        h1 span.garim-emoji {
            vertical-align: -0.06em;
        }
        /* Ajuda dos buracos: <details> dentro do expander «Emissão própria» (evita expander aninhado) */
        details.garim-detalhe-ajuda {
            margin: 0.35rem 0 0.5rem 0 !important;
            padding: 0.35rem 0.5rem !important;
            border-radius: 8px !important;
            border: 1px solid rgba(161, 134, 158, 0.28) !important;
            background: rgba(255, 255, 255, 0.55) !important;
        }
        details.garim-detalhe-ajuda summary {
            cursor: pointer !important;
            font-size: 0.88rem !important;
            color: #5D1B36 !important;
        }
        /* h4/h5 em HTML próprio: espaçamento semelhante ao markdown ##### da app */
        [data-testid="stMarkdown"] h4 {
            margin: 0.5rem 0 0.35rem 0 !important;
            font-size: 1rem !important;
            font-weight: 600;
        }
        [data-testid="stMarkdown"] h5 {
            margin: 0.85rem 0 0.4rem 0 !important;
            font-size: 0.875rem !important;
            font-weight: 600;
        }

        .instrucoes-card {
            background-color: rgba(255, 255, 255, 0.7);
            border-radius: 15px;
            padding: 20px;
            border-left: 5px solid #FF69B4;
            margin-bottom: 20px;
            min-height: 280px;
        }
        .instrucoes-card.manual-compacto {
            min-height: 0;
            margin-bottom: 12px;
        }

        [data-testid="stMetric"] {
            background: white !important;
            border-radius: 20px !important;
            border: 1px solid #FFDEEF !important;
            padding: 15px !important;
        }

        h3.garim-sec {
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 800 !important;
            font-size: 1.08rem !important;
            color: #5D1B36 !important;
            text-align: left !important;
            margin: 1.35rem 0 0.55rem 0 !important;
            letter-spacing: 0.02em;
            border-left: 4px solid #A1869E;
            padding: 0.2rem 0 0.2rem 0.75rem;
            line-height: 1.35 !important;
        }
        section.main [data-testid="stDataFrame"] {
            border-radius: 12px;
            border: 1px solid rgba(161, 134, 158, 0.28);
            overflow: hidden;
        }
        /* Ag-Grid (relatório da leitura): borda + scroll horizontal no contentor (grelha larga) */
        section.main div[data-testid="stIFrame"] {
            border-radius: 12px !important;
            overflow-x: auto !important;
            max-width: 100% !important;
        }
        section.main iframe[title="streamlit_aggrid.agGrid"] {
            border-radius: 12px !important;
            min-width: 720px !important;
        }
        /* Abas do relatório: mais ar e menos texto espremido */
        section.main [data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: 0.4rem !important;
            flex-wrap: wrap !important;
            padding-bottom: 0.35rem !important;
        }
        section.main [data-testid="stTabs"] button[data-baseweb="tab"] {
            padding: 0.5rem 0.85rem !important;
            font-size: 0.88rem !important;
            font-weight: 600 !important;
            border-radius: 10px !important;
            white-space: nowrap !important;
        }
        section.main [data-testid="stTabs"] [data-testid="stVerticalBlock"] {
            padding-top: 0.35rem !important;
        }
        /* Painel com borda (uploads à direita): respiro interno */
        section.main [data-testid="stVerticalBlockBorderWrapper"] {
            padding: 0.9rem 1.1rem !important;
        }
        /* Segunda coluna (rail direita): largura mínima em ecrã largo */
        @media (min-width: 1100px) {
            section.main [data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-of-type(2) {
                min-width: 300px !important;
            }
        }
        /* Painel direito (uploads): alinhar ao topo — o contentor com borda é o nativo do Streamlit */
        @media (min-width: 1100px) {
            section.main div[data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- VARIÁVEIS DE SISTEMA DE ARQUIVOS (PREVENÇÃO DE QUEDA DE MEMÓRIA) ---
TEMP_EXTRACT_DIR = "temp_garimpo_zips"
TEMP_UPLOADS_DIR = "temp_garimpo_uploads"
MAX_XML_PER_ZIP = 10000  # Máx. XMLs por ficheiro ZIP (lista específica e Etapa 3); reparte em vários lotes
ZIP_EXPORT_COMPRESSLEVEL = 9  # 1–9: 9 = .zip menores na exportação (mais CPU ao gravar)

# Exportação dedicada: ZIPs ~50 MB, só NF-e/NFC-e série 4 autorizadas (emissão própria).
CNPJ_CLIENTE_EXPORT_SERIE4_ZIP50MB = "45785442000390"
ZIP50_SERIE4_MAX_BYTES = 50 * 1024 * 1024
SESSION_KEY_SERIE4_ZIP50_PARTS = "_garim_serie4_emitidas_zip50_parts"


def _zipfile_open_write_export(path: str):
    """ZIP de download/pacotes: DEFLATE com compressão máxima."""
    return zipfile.ZipFile(
        str(path), "w", zipfile.ZIP_DEFLATED, compresslevel=ZIP_EXPORT_COMPRESSLEVEL
    )


def _lista_ficheiros_pasta_uploads():
    """
    Só ficheiros em TEMP_UPLOADS_DIR (ignora subpastas).
    Abrir uma pasta com open(..., 'rb') no Windows gera erro — o garimpo e os ZIPs devem iterar só ficheiros.
    """
    if not os.path.isdir(TEMP_UPLOADS_DIR):
        return []
    try:
        return [
            f
            for f in os.listdir(TEMP_UPLOADS_DIR)
            if os.path.isfile(os.path.join(TEMP_UPLOADS_DIR, f))
        ]
    except OSError:
        return []


# Bytes dos XML/ZIP do lote atual (garimpo + «incluir mais»), quando não se grava em disco local.
SESSION_KEY_FONTES_XML_MEMORIA = "_garimpo_fontes_xml_memoria"
# SHA-256 de ficheiros já incorporados ao lote via «Incluir mais» (evita duplicar a cada rerun / duplo Processar).
SESSION_KEY_EXTRA_DIGESTS = "_garimpo_extra_sha256_vistos"


def _session_state_get_garimpo(key, default=None):
    """
    Lê session_state sem rebentar com «SessionInfo before initialization».
    O get_script_run_ctx() sozinho não basta em algumas versões; só try/except é fiável.
    """
    try:
        return st.session_state.get(key, default)
    except Exception:
        return default


def _session_state_pop_garimpo(key):
    try:
        st.session_state.pop(key, None)
    except Exception:
        pass


def _garimpo_analise_sem_pasta_local_projeto() -> bool:
    """
    True (omissão): durante a análise **não** grava uploads em `temp_garimpo_uploads` na pasta do projeto —
    mantém os bytes na sessão Streamlit até nova análise. Exportações ZIP/Excel continuam a ir para a pasta **que escolher**.
    Defina `GARIMPEIRO_ANALISE_SEM_DISCO_LOCAL=0` para voltar a gravar no disco (lotes muito grandes e pouca RAM).
    """
    v = os.environ.get("GARIMPEIRO_ANALISE_SEM_DISCO_LOCAL", "1").strip().lower()
    return v not in ("0", "false", "no", "off")


def _garimpo_limpar_fontes_xml_memoria_sessao():
    _session_state_pop_garimpo(SESSION_KEY_FONTES_XML_MEMORIA)


def _garimpo_nome_chave_upload(indice: int, nome_original: str) -> str:
    base = os.path.basename(str(nome_original or "ficheiro"))
    safe = "".join(c if (c.isalnum() or c in "._- ") else "_" for c in base).strip() or "ficheiro"
    return f"{int(indice):05d}_{safe}"[:220]


def _lista_nomes_fontes_xml_garimpo():
    """
    Nomes-chave de todas as fontes do lote. Com lote em memória, inclui também ficheiros em
    temp_garimpo_uploads (ex.: «Incluir mais» gravado em disco por fallback) — antes só a
    memória era lida e esses XML eram ignorados no «Processar Dados».
    """
    disk = _lista_ficheiros_pasta_uploads()
    if _garimpo_analise_sem_pasta_local_projeto():
        mem = _session_state_get_garimpo(SESSION_KEY_FONTES_XML_MEMORIA)
        if isinstance(mem, dict) and mem:
            return sorted(set(mem.keys()) | set(disk))
    return sorted(disk)


def _garimpo_existem_fontes_xml_lote():
    return bool(_lista_nomes_fontes_xml_garimpo())


@contextmanager
def _abrir_fonte_xml_garimpo_stream(f_name: str):
    """Abre um ficheiro do lote: da memória da sessão ou de TEMP_UPLOADS_DIR."""
    mem = _session_state_get_garimpo(SESSION_KEY_FONTES_XML_MEMORIA)
    if (
        _garimpo_analise_sem_pasta_local_projeto()
        and isinstance(mem, dict)
        and f_name in mem
    ):
        bio = io.BytesIO(mem[f_name])
        try:
            yield bio
        finally:
            bio.close()
    else:
        f_path = os.path.join(TEMP_UPLOADS_DIR, f_name)
        with open(f_path, "rb") as f:
            yield f


def _garimpo_absorver_uploads_extra_no_lote(extra_files) -> int:
    """
    Incorpora ficheiros do «Incluir mais» ao lote (memória ou pasta temp), deduplicando por SHA-256
    do conteúdo. Assim os XML entram no lote logo após a escolha (cada rerun) e não dependem só
    do retorno do widget no mesmo instante do clique em «Processar Dados».
    """
    if not extra_files:
        return 0
    files = list(extra_files) if isinstance(extra_files, (list, tuple)) else [extra_files]
    try:
        seen = st.session_state.get(SESSION_KEY_EXTRA_DIGESTS)
        if not isinstance(seen, set):
            seen = set()
        n_new = 0
        if _garimpo_analise_sem_pasta_local_projeto():
            mem = _session_state_get_garimpo(SESSION_KEY_FONTES_XML_MEMORIA)
            if not isinstance(mem, dict):
                mem = {}
            for f in files:
                try:
                    raw = f.getvalue()
                except Exception:
                    continue
                d = hashlib.sha256(raw).hexdigest()
                if d in seen:
                    continue
                seen.add(d)
                key = _garimpo_nome_chave_upload(len(mem), getattr(f, "name", None) or "extra")
                mem[key] = raw
                n_new += 1
            if n_new:
                st.session_state[SESSION_KEY_FONTES_XML_MEMORIA] = mem
        else:
            os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
            start_i = len(_lista_ficheiros_pasta_uploads())
            for f in files:
                try:
                    raw = f.getvalue()
                except Exception:
                    continue
                d = hashlib.sha256(raw).hexdigest()
                if d in seen:
                    continue
                seen.add(d)
                key = _garimpo_nome_chave_upload(start_i + n_new, getattr(f, "name", None) or "extra")
                caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, key)
                with open(caminho_salvo, "wb") as out_f:
                    out_f.write(raw)
                n_new += 1
        st.session_state[SESSION_KEY_EXTRA_DIGESTS] = seen
        return n_new
    except Exception:
        return 0


def _streamlit_likely_community_cloud() -> bool:
    """Heurística para Streamlit Community Cloud / ambiente de deploy sem disco persistente."""
    exe = (sys.executable or "").lower()
    return "/mount/src/" in exe or "/home/adminuser/" in exe


def _mariana_zip_default_dir() -> Path:
    """Pasta por omissão dos ZIP do pacote apuração (junto a app2.py)."""
    return Path(__file__).resolve().parent


def _mariana_destino_zip_para_gravar():
    """
    Resolve o caminho em session_state['mariana_zip_save_dir'].
    Pasta em branco → erro (não usa pasta por omissão).
    Devolve (Path, None) em caso de sucesso ou (None, mensagem de erro).
    """
    raw = st.session_state.get("mariana_zip_save_dir")
    s = (str(raw).strip().strip('"').strip("'") if raw is not None else "")
    if not s:
        return (
            None,
            "Indique a **pasta completa** onde gravar (o campo não pode ficar em branco).",
        )
    try:
        p = Path(s).expanduser().resolve()
    except (OSError, ValueError):
        return None, "Caminho inválido. Use um caminho completo (ex.: D:\\Exportacoes\\Contabilidade)."
    if p.exists() and not p.is_dir():
        return (
            None,
            "Esse caminho já existe e não é uma pasta (é um ficheiro). Escolha outra pasta.",
        )
    try:
        p.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return None, f"Não foi possível criar ou aceder à pasta: {e}"
    return p, None


def _v2_destino_zip_etapa3_para_gravar():
    """
    Pasta para ZIP «com pastas» / «XML soltos» (Gerar sua empresa / terceiros).
    Em branco → erro (sem pasta por omissão). Devolve (Path, None) ou (None, mensagem).
    """
    raw = st.session_state.get("v2_etapa3_zip_save_dir")
    s = (str(raw).strip().strip('"').strip("'") if raw is not None else "")
    if not s:
        return (
            None,
            "Indique a **pasta completa** para os ZIP (o campo não pode ficar em branco).",
        )
    try:
        p = Path(s).expanduser().resolve()
    except (OSError, ValueError):
        return None, "Caminho inválido para ZIP (Etapa 3). Use um caminho completo (ex.: D:\\Exportacoes\\Garimpeiro)."
    if p.exists() and not p.is_dir():
        return (
            None,
            "Esse caminho já existe e não é uma pasta (é um ficheiro). Escolha outra pasta para os ZIP.",
        )
    try:
        p.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return None, f"Não foi possível criar ou aceder à pasta dos ZIP: {e}"
    return p, None


def _v2_sanitize_nome_export(s, max_len=72):
    """Trecho seguro para nome de ficheiro (sem caminhos nem caracteres proibidos no Windows)."""
    if s is None:
        return ""
    t = str(s).strip().strip('"').strip("'")
    if not t:
        return ""
    proib = '\\/:*?"<>|'
    out = []
    for c in t:
        if c in proib:
            continue
        if c.isspace():
            out.append("_")
        elif c.isalnum() or c in "._-":
            out.append(c)
    r = "".join(out).strip("._-")
    while "__" in r:
        r = r.replace("__", "_")
    if not r:
        return ""
    return r[:max_len]


def _v2_stems_zip_nome_ficheiro_etapa3(nome_raw: str, zip_tag):
    """
    Nomes dos ZIP no disco. Sem nome → z_org_propria_pt1… Com nome → Nome_org_propria_pt1…
    zip_tag: «propria», «terceiros» ou None (lote completo → sufixo final).
    """
    org_def = "z_org_final" if not zip_tag else f"z_org_{zip_tag}"
    todos_def = "z_todos_final" if not zip_tag else f"z_todos_{zip_tag}"
    u = _v2_sanitize_nome_export(nome_raw, max_len=64)
    if not u:
        return org_def, todos_def
    t = zip_tag if zip_tag else "final"
    return f"{u}_org_{t}", f"{u}_todos_{t}"


_PACOTE_CONTAB_NOME_EXCEL_RAIZ = "relatorio_garimpeiro_completo.xlsx"
# Pasta “mãe” dentro de cada ZIP de contabilidade; dentro dela vêm Lote_001, Lote_002, … (até MAX_XML_PER_ZIP XML por pasta).
PACOTE_CONTAB_PASTA_MAE_XML = "XML"


def _nome_excel_pacote_contab_dentro_zip(slug: str) -> str:
    """
    Nome do .xlsx **dentro** de cada ZIP do pacote contabilidade — inclui série / mês / grupo (o mesmo critério do nome do .zip),
    para extrair vários ZIPs para a mesma pasta sem o Excel ser sempre substituído por «relatorio_garimpeiro_completo.xlsx».
    """
    tail = _v2_sanitize_nome_export(slug, max_len=150) or "lote"
    base = _v2_sanitize_nome_export(f"relatorio_garimpeiro_{tail}", max_len=200) or f"relatorio_garimpeiro_{tail}"
    if not str(base).lower().endswith(".xlsx"):
        base = f"{base}.xlsx"
    return base

_LEIAME_ESTRUTURA_CONTABILIDADE = """Pacote Garimpeiro — contabilidade / matriz
========================================

Na **pasta que escolheu** para guardar o pacote:
  • **Um Excel solto** — ex.: «nome_do_pacote_relatorio_garimpeiro_completo.xlsx» (Geral, Buracos, status, Terceiros, Dashboard, etc.; **sem** folha Painel Fiscal).
  • Vários **ficheiros .zip** (grupos Emitidas / Terceiros).

Cada ficheiro ZIP contém:
  • **Pasta XML** — dentro dela, subpastas **Lote_001**, **Lote_002**, … com até 10 000 XML cada. O nome do .xml costuma ser a chave de 44 dígitos (ou `…_cancelamento.xml` / `…_denegada.xml` / `…_rejeitada.xml` quando for evento distinto da mesma chave), ou INUT_….
  • **relatorio_garimpeiro_…xlsx** — na **raiz** do ZIP (nome inclui série/mês/grupo como o .zip; sem folha Painel Fiscal).

Nome do ficheiro .zip no disco (quando há notas no relatório): inclui **nota inicial e nota final** daquele grupo, ex.: …_Emitidas_Autorizadas_Serie_1_notas_1500_96842.zip

Vários ZIP (só são criados se existir XML naquele grupo):
  • **Emitidas** (sua empresa), por série **e mês de emissão** — ex.: Emitidas_Autorizadas_Serie_1_Mes_2024_03, …
    (evita misturar no mesmo ZIP notas de meses diferentes). Dentro: pasta XML/Lote_001… (10 000 XML por pasta).
  • **Terceiros**, por modelo e status (sem separação por mês) — ex.: Terceiros_NFe_Autorizadas, …

O Excel em cada ZIP é o **mesmo conteúdo** (todo o lote lido), mas o **nome do ficheiro** muda por grupo para não se sobrepor ao extrair.
"""


def _leiame_contab_zipfile(zf):
    if zf is None:
        return
    zf.writestr(
        "LEIAME_pacote_contabilidade.txt",
        _LEIAME_ESTRUTURA_CONTABILIDADE.encode("utf-8"),
    )


def _origem_row_e_propria(origem) -> bool:
    """True = emissão própria (cliente é emitente), a partir da coluna Origem do relatório."""
    o = str(origem or "").upper()
    if "TERCEIROS" in o:
        return False
    return "PRÓPRIA" in o or "PROPRIA" in o


def _pacote_contab_status_curto(status_text) -> str:
    """Nome curto de status para o stem do ZIP (Autorizadas, Canceladas, …)."""
    st = str(status_text or "").strip().upper()
    if st in ("NORMAIS", "AUTORIZADA", "AUTORIZADAS"):
        return "Autorizadas"
    if st in ("CANCELADOS", "CANCELADA", "CANCELADAS"):
        return "Canceladas"
    if st in ("INUTILIZADOS", "INUTILIZADA", "INUTILIZADAS"):
        return "Inutilizadas"
    if st in ("DENEGADOS", "DENEGADA", "DENEGADAS"):
        return "Denegadas"
    if st in ("REJEITADOS", "REJEITADA", "REJEITADAS"):
        return "Rejeitadas"
    return _v2_sanitize_nome_export(st, max_len=24) or "Outros_status"


def _pacote_contab_tipo_zip_terceiros(modelo) -> str:
    """Grupo do nome do ZIP para terceiros (um par modelo×status por ficheiro)."""
    m = str(modelo or "").strip()
    return {
        "NF-e": "NFe",
        "NFC-e": "NFCe",
        "CT-e": "CTe",
        "MDF-e": "MDFe",
        "NFS-e": "NFSe",
        "CT-e OS": "CTeOS",
    }.get(m, "Outros")


def _pacote_contab_slug_zip(is_propria: bool, status_text, serie, modelo) -> str:
    """
    Stem base do ficheiro ZIP: Emitidas por série + status; terceiros por modelo (NFe, CTe, …) + status.
    Nas emitidas o sufixo de mês é acrescentado à parte (ver _pacote_contab_slug_emitidas_com_mes).
    status_text: «Status Final» ou res['Status']; modelo: «Modelo» no DF ou res['Tipo'].
    """
    st_short = _pacote_contab_status_curto(status_text)
    if is_propria:
        ser_raw = str(serie if serie is not None else "").strip() or "0"
        ser = _v2_sanitize_nome_export(ser_raw, max_len=16) or "0"
        return f"Emitidas_{st_short}_Serie_{ser}"
    g = _pacote_contab_tipo_zip_terceiros(modelo)
    return f"Terceiros_{g}_{st_short}"


def _pacote_contab_sufixo_mes_emissao_emitidas(ano, mes) -> str:
    """Sufixo para ZIPs de emissão própria: _Mes_AAAA_MM (separa ficheiros de meses diferentes no pacote)."""
    try:
        a = int(ano)
        m = int(mes)
        if a <= 0 or m < 1 or m > 12:
            return "_Mes_desconhecido"
        return f"_Mes_{a:04d}_{m:02d}"
    except (TypeError, ValueError):
        return "_Mes_desconhecido"


def _pacote_contab_slug_emitidas_com_mes(is_propria: bool, status_text, serie, modelo, ano, mes) -> str:
    base = _pacote_contab_slug_zip(is_propria, status_text, serie, modelo)
    if not is_propria:
        return base
    return base + _pacote_contab_sufixo_mes_emissao_emitidas(ano, mes)


def _montar_mapa_chave_slug_contab(df: pd.DataFrame, chaves_permitidas: set) -> dict:
    """Chave normalizada → slug do lote (só chaves que entram no pacote). Emitidas: um ZIP por mês de emissão."""
    out = {}
    if df is None or df.empty or "Chave" not in df.columns or not chaves_permitidas:
        return out
    for _, row in df.iterrows():
        k = _chave_para_conjunto_export(row.get("Chave"))
        if not k or k not in chaves_permitidas:
            continue
        is_pr = _origem_row_e_propria(row.get("Origem", ""))
        mod = row.get("Modelo", "") or row.get("Tipo", "")
        ano = row["Ano"] if "Ano" in row.index else None
        mes = row["Mes"] if "Mes" in row.index else None
        slug = _pacote_contab_slug_emitidas_com_mes(
            is_pr,
            row.get("Status Final", ""),
            row.get("Série", "0"),
            mod,
            ano,
            mes,
        )
        out[k] = slug
    return out


def _nota_int_de_linha_relatorio(row):
    """Número da nota a partir de uma linha do relatório geral (Número, Num_Faltante ou posições 25–33 da chave 44)."""
    if row is None:
        return None
    for col in ("Número", "Num_Faltante"):
        if col not in row.index:
            continue
        n = row[col]
        if pd.isna(n):
            continue
        try:
            return int(float(n))
        except (TypeError, ValueError):
            continue
    if "Chave" not in row.index:
        return None
    ch = _chave_para_conjunto_export(row["Chave"])
    if len(ch) == 44 and ch.isdigit():
        try:
            return int(ch[25:34])
        except ValueError:
            pass
    return None


def _pacote_contab_notas_min_max_por_slug(df_ref, mapa_slug: dict, filtro_chaves: set) -> dict:
    """slug → (n_min, n_max) com base no DataFrame de referência (para sufixo _notas_ no nome do .zip)."""
    mm = defaultdict(lambda: [None, None])
    if df_ref is None or getattr(df_ref, "empty", True) or not filtro_chaves:
        return {}
    if "Chave" not in df_ref.columns:
        return {}
    for _, row in df_ref.iterrows():
        k = _chave_para_conjunto_export(row["Chave"] if "Chave" in row.index else None)
        if not k or k not in filtro_chaves:
            continue
        slug = mapa_slug.get(k)
        if not slug:
            continue
        n = _nota_int_de_linha_relatorio(row)
        if n is None:
            continue
        lo, hi = mm[slug]
        mm[slug][0] = n if lo is None else min(lo, n)
        mm[slug][1] = n if hi is None else max(hi, n)
    return {s: (lo, hi) for s, (lo, hi) in mm.items() if lo is not None and hi is not None}


def _v2_export_pacote_contab_por_dimensoes(
    out_dir: Path,
    stem_org: str,
    filtro_chaves: set,
    cnpj_limpo: str,
    xb_completo,
    excel_fn_completo: str,
    df_ref: pd.DataFrame,
):
    """
    ZIP por emitida (status × série × **mês**) ou por terceiros (modelo × status).
    Dentro de cada ZIP: pasta XML/Lote_001, Lote_002, … (até MAX_XML_PER_ZIP XML por pasta)
    + Excel na raiz com nome único por grupo (`relatorio_garimpeiro_<slug>.xlsx`). Nome do .zip pode incluir _notas_min_max.
    Grava também o Excel completo **solto** em out_dir (prefixo = stem_org).
    Devolve (paths, [], matched, aviso, caminho_excel_solta|None).
    """
    mapa_slug = _montar_mapa_chave_slug_contab(df_ref, filtro_chaves)
    slug_ranges = _pacote_contab_notas_min_max_por_slug(df_ref, mapa_slug, filtro_chaves)
    zips_abertos = {}
    paths_ordered = {}
    slug_lote_idx = {}
    slug_in_lote = {}

    def _ensure_zip(slug: str):
        if slug not in zips_abertos:
            lo_hi = slug_ranges.get(slug)
            suf = ""
            if lo_hi and lo_hi[0] is not None and lo_hi[1] is not None:
                suf = f"_notas_{int(lo_hi[0])}_{int(lo_hi[1])}"
            combo = _v2_sanitize_nome_export(f"{stem_org}__{slug}{suf}", max_len=200) or "pacote"
            path = out_dir / f"{combo}.zip"
            zf = _zipfile_open_write_export(path)
            zips_abertos[slug] = zf
            paths_ordered[slug] = str(path)
            _leiame_contab_zipfile(zf)
        return zips_abertos[slug]

    def _prefixo_lote_xml(slug: str) -> str:
        if slug not in slug_lote_idx:
            slug_lote_idx[slug] = 1
            slug_in_lote[slug] = 0
        if slug_in_lote[slug] >= MAX_XML_PER_ZIP:
            slug_lote_idx[slug] += 1
            slug_in_lote[slug] = 0
        li = slug_lote_idx[slug]
        slug_in_lote[slug] += 1
        return f"{PACOTE_CONTAB_PASTA_MAE_XML}/Lote_{li:03d}"

    xml_matched = 0
    chaves_ja_gravadas_pacote = set()
    for f_name in _lista_nomes_fontes_xml_garimpo():
        with _abrir_fonte_xml_garimpo_stream(f_name) as f_temp:
            for name, xml_data in extrair_recursivo(f_temp, f_name):
                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                ck = _chave_para_conjunto_export(res["Chave"]) if res else None
                if res and ck and ck in filtro_chaves:
                    td = _tupla_dedupe_export_xml(res, ck)
                    if td is None or td in chaves_ja_gravadas_pacote:
                        del xml_data
                        continue
                    chaves_ja_gravadas_pacote.add(td)
                    xml_matched += 1
                    slug = mapa_slug.get(ck) or _pacote_contab_slug_emitidas_com_mes(
                        is_p,
                        str(res.get("Status") or "NORMAIS"),
                        res.get("Série", "0"),
                        res.get("Tipo", "Outros"),
                        res.get("Ano"),
                        res.get("Mes"),
                    )
                    zf = _ensure_zip(slug)
                    _pfx = _prefixo_lote_xml(slug)
                    _inner = f"{_pfx}/{_nome_arquivo_xml_contabilidade(res, name)}"
                    zf.writestr(_inner, xml_data)
                del xml_data

    for slug, zf in zips_abertos.items():
        try:
            if xb_completo:
                zf.writestr(_nome_excel_pacote_contab_dentro_zip(slug), xb_completo)
        except OSError:
            pass
        try:
            zf.close()
        except OSError:
            pass

    excel_solta_path = None
    if xb_completo:
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            stem_safe = _v2_sanitize_nome_export(stem_org, max_len=80) or "pacote_apuracao"
            nome_xlsx = _v2_sanitize_nome_export(
                f"{stem_safe}_{_PACOTE_CONTAB_NOME_EXCEL_RAIZ}", max_len=200
            ) or _PACOTE_CONTAB_NOME_EXCEL_RAIZ
            if not str(nome_xlsx).lower().endswith(".xlsx"):
                nome_xlsx = f"{nome_xlsx}.xlsx"
            p_x = out_dir / nome_xlsx
            _raw = (
                xb_completo
                if isinstance(xb_completo, (bytes, bytearray))
                else bytes(xb_completo)
            )
            with open(p_x, "wb") as xf:
                xf.write(_raw)
            excel_solta_path = str(p_x.resolve())
        except OSError:
            excel_solta_path = None

    aviso = None
    if xml_matched == 0:
        aviso = (
            "Nenhum XML em disco correspondeu às chaves. "
            "Causas frequentes: pasta do garimpo apagada, ou chaves na tabela que não batem com os ficheiros."
        )
    _paths_list = sorted(paths_ordered.values(), key=lambda p: os.path.basename(p).lower())
    return _paths_list, [], xml_matched, aviso, excel_solta_path


def _tupla_dedupe_export_xml(res: dict | None, ck) -> tuple | None:
    """
    Uma entrada por (chave lógica, «tipo» de XML) dentro do mesmo ZIP.
    Permite **dois** ficheiros com a mesma chave 44 quando um é a NF (ou doc «normal»)
    e outro é cancelamento / denegação / rejeição (evento ou proc distinto).
    Duplicar **o mesmo** tipo (ex.: dois proc autorizados) continua a ser ignorado.
    """
    if ck is None:
        return None
    if not res:
        return (ck, "UNK")
    if str(ck).startswith("INUT_"):
        return (ck, "INUT")
    st = str(res.get("Status") or "").strip().upper()
    if st == "CANCELADOS":
        return (ck, "CANCELADOS")
    if st == "DENEGADOS":
        return (ck, "DENEGADOS")
    if st == "REJEITADOS":
        return (ck, "REJEITADOS")
    return (ck, "DOC")


def _nome_arquivo_xml_contabilidade(res: dict, nome_original: str) -> str:
    ch = str(res.get("Chave") or "").strip()
    st = str(res.get("Status") or "").strip().upper()
    if len(ch) == 44 and ch.isdigit():
        if st == "CANCELADOS":
            return f"{ch}_cancelamento.xml"
        if st == "DENEGADOS":
            return f"{ch}_denegada.xml"
        if st == "REJEITADOS":
            return f"{ch}_rejeitada.xml"
        return f"{ch}.xml"
    if ch.startswith("INUT_"):
        safe = "".join(c if (c.isalnum() or c in "_-") else "_" for c in ch)
        return f"{safe}.xml"
    base = os.path.basename(nome_original)
    if not base.lower().endswith(".xml"):
        base = f"{base}.xml"
    return base


def _caminho_xml_pacote_contab_raiz(res: dict, nome_xml: str) -> str:
    """Nome base do .xml (chave 44 ou INUT_…); no pacote contabilidade completo o caminho inclui XML/Lote_NNN/."""
    return _nome_arquivo_xml_contabilidade(res, nome_xml)


def _is_mariana_pc_bundle() -> bool:
    """
    Mostra o bloco «Pacote para contabilidade / matriz» (todo o lote lido, sem filtros da Etapa 3).
    • Local (streamlit run): ligado por omissão.
    • Streamlit Cloud: desligado por omissão (ligar com GARIMPEIRO_MARIANA_PC=1 no deploy).
    • Desligar em qualquer sítio: GARIMPEIRO_MARIANA_PC=0
    """
    env = os.environ.get("GARIMPEIRO_MARIANA_PC", "").strip().lower()
    if env in ("0", "false", "no"):
        return False
    if env in ("1", "true", "yes"):
        return True
    if (Path(__file__).resolve().parent / ".mariana_pc").is_file():
        return True
    s = str(Path(__file__).resolve().parent).lower()
    if "servidor pc" in s or "garimpeiro servidor pc" in s:
        return True
    try:
        if Path(r"N:\MARIANA").is_dir():
            return True
    except OSError:
        pass
    if _streamlit_likely_community_cloud():
        return False
    return True


# Se dois números emitidos consecutivos (ordenados) diferem mais que isto, tratamos como outra faixa.
# Assim evitamos milhões de "buracos" falsos (ex.: uma chave/XML errado com nº gigante ou duas séries distantes misturadas).
MAX_SALTO_ENTRE_NOTAS_CONSECUTIVAS = 25000


def format_cnpj_visual(digits: str) -> str:
    """Máscara CNPJ (00.000.000/0000-00) a partir apenas de dígitos, até 14."""
    d = "".join(c for c in str(digits) if c.isdigit())[:14]
    if not d:
        return ""
    n = len(d)
    if n <= 2:
        return d
    if n <= 5:
        return f"{d[:2]}.{d[2:]}"
    if n <= 8:
        return f"{d[:2]}.{d[2:5]}.{d[5:]}"
    if n <= 12:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:]}"
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


# --- MOTOR DE IDENTIFICAÇÃO ---
def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj))) if client_cnpj else ""
    nome_puro = os.path.basename(file_name)
    if nome_puro.startswith('.') or nome_puro.startswith('~') or not nome_puro.lower().endswith('.xml'):
        return None, False
    
    resumo = {
        "Arquivo": nome_puro, 
        "Chave": "", 
        "Tipo": "Outros", 
        "Série": "0",
        "Número": 0, 
        "Status": "NORMAIS", 
        "Pasta": "",
        "Valor": 0.0, 
        "Conteúdo": b"", 
        "Ano": "0000", 
        "Mes": "00",
        "Operacao": "SAIDA", 
        "Data_Emissao": "",
        "CNPJ_Emit": "", 
        "Nome_Emit": "", 
        "Doc_Dest": "", 
        "Nome_Dest": "",
        "UF_Dest": "",
    }
    
    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        if '<?xml' not in tag_l and '<inf' not in tag_l and '<inut' not in tag_l and '<retinut' not in tag_l: 
            return None, False
        
        # Identificação de tpNF (0=Entrada, 1=Saída)
        tp_nf_match = re.search(r'<tpnf>([01])</tpnf>', tag_l)
        if tp_nf_match:
            if tp_nf_match.group(1) == "0":
                resumo["Operacao"] = "ENTRADA"
            else:
                resumo["Operacao"] = "SAIDA"

        # Extração de Dados das Partes
        resumo["CNPJ_Emit"] = re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S).group(1) if re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S) else ""
        resumo["Nome_Emit"] = re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""
        resumo["Doc_Dest"] = re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S).group(1) if re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S) else ""
        resumo["Nome_Dest"] = re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""
        _uf_m = re.search(r'<dest>.*?<uf>([A-Za-z]{2})</uf>', tag_l, re.S)
        resumo["UF_Dest"] = _uf_m.group(1).upper() if _uf_m else ""

        # Data de Emissão Genérica
        data_match = re.search(r'<(?:dhemi|demi|dhregevento|dhrecbto)>(\d{4})-(\d{2})-(\d{2})', tag_l)
        if data_match: 
            resumo["Data_Emissao"] = f"{data_match.group(1)}-{data_match.group(2)}-{data_match.group(3)}"
            resumo["Ano"] = data_match.group(1)
            resumo["Mes"] = data_match.group(2)

        # 1. IDENTIFICAÇÃO DE INUTILIZADAS
        if '<inutnfe' in tag_l or '<retinutnfe' in tag_l or '<procinut' in tag_l:
            resumo["Status"] = "INUTILIZADOS"
            resumo["Tipo"] = "NF-e"
            
            if '<mod>65</mod>' in tag_l: 
                resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l: 
                resumo["Tipo"] = "CT-e"
            
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            fin = re.search(r'<nnffin>(\d+)</', tag_l).group(1) if re.search(r'<nnffin>(\d+)</', tag_l) else ini
            
            resumo["Número"] = int(ini)
            resumo["Range"] = (int(ini), int(fin))
            
            if resumo["Ano"] == "0000":
                ano_match = re.search(r'<ano>(\d+)</', tag_l)
                if ano_match: 
                    resumo["Ano"] = "20" + ano_match.group(1)[-2:]
                    
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"

            # Inutilização: o emitente está em infInut / inutNFe, não em <emit> — sem isto CNPJ_Emit fica vazio
            # e is_p fica sempre falso (pacote/ZIP aparece como «terceiros» mesmo sendo o mesmo CNPJ).
            if not resumo["CNPJ_Emit"]:
                for _pat in (
                    r"<[^>]*infinut[^>]*>[\s\S]{0,12000}?<cnpj>(\d{11,14})</cnpj>",
                    r"<inutnfe[^>]*>[\s\S]{0,12000}?<cnpj>(\d{11,14})</cnpj>",
                ):
                    _mc = re.search(_pat, tag_l, re.I)
                    if _mc:
                        resumo["CNPJ_Emit"] = _mc.group(1)
                        break

        else:
            match_ch = re.search(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</', tag_l)
            if not match_ch:
                match_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
                if match_ch:
                    resumo["Chave"] = match_ch.group(1)
                else:
                    resumo["Chave"] = ""
            else:
                resumo["Chave"] = match_ch.group(1)

            if resumo["Chave"] and len(resumo["Chave"]) == 44:
                resumo["Ano"] = "20" + resumo["Chave"][2:4]
                resumo["Mes"] = resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])
                
                if not resumo["Data_Emissao"]: 
                    resumo["Data_Emissao"] = f"{resumo['Ano']}-{resumo['Mes']}-01"

            # Modelo: 1) dígitos 21–22 da chave de acesso (44) = código Sefaz; 2) <mod> / raiz XML;
            # 3) heurística NFS-e por último — evita CT-e (57) com <servico> no XML ser lido como NFS-e.
            tipo = "NF-e"
            _ch44 = resumo.get("Chave") or ""
            if len(_ch44) == 44 and _ch44.isdigit():
                _mk = _ch44[20:22]
                if _mk == "57":
                    tipo = "CT-e"
                elif _mk == "58":
                    tipo = "MDF-e"
                elif _mk == "65":
                    tipo = "NFC-e"
                elif _mk == "67":
                    tipo = "CT-e OS"
                elif _mk == "55":
                    tipo = "NF-e"
            if tipo == "NF-e" and not (len(_ch44) == 44 and _ch44.isdigit() and _ch44[20:22] in ("55", "57", "58", "65", "67")):
                if '<mod>57</mod>' in tag_l or '<infcte' in tag_l:
                    tipo = "CT-e"
                elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l:
                    tipo = "MDF-e"
                elif '<mod>65</mod>' in tag_l:
                    tipo = "NFC-e"
                elif '<mod>67</mod>' in tag_l or "<dadcte" in tag_l or "dacte" in tag_l or "<cteos" in tag_l or "cte-os" in tag_l:
                    tipo = "CT-e OS"
                elif '<mod>55</mod>' in tag_l:
                    tipo = "NF-e"
                elif re.search(r'<[^>]*nfse', tag_l) or "<nfse" in tag_l or ("servico" in tag_l and "nfse" in tag_l):
                    tipo = "NFS-e"
            
            status = "NORMAIS"
            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: 
                status = "CANCELADOS"
            elif '110110' in tag_l: 
                status = "CARTA_CORRECAO"
            elif re.search(r"<cstat>110</cstat>", tag_l) or "deneg" in tag_l:
                status = "DENEGADOS"
            elif re.search(r"<cstat>30[1-9]</cstat>", tag_l) or re.search(
                r"<cstat>3[1-4]\d</cstat>", tag_l
            ):
                status = "REJEITADOS"
                
            resumo["Tipo"] = tipo
            resumo["Status"] = status
            # Carta de correção (evento 110110): não entra no lote nem no dashboard/PDF.
            if status == "CARTA_CORRECAO":
                return None, False

            if status == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                if v_match:
                    resumo["Valor"] = float(v_match.group(1))
                else:
                    resumo["Valor"] = 0.0
            
        if not resumo["CNPJ_Emit"] and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"): 
            resumo["CNPJ_Emit"] = resumo["Chave"][6:20]
        
        if resumo["Mes"] == "00": 
            resumo["Mes"] = "01"
            
        if resumo["Ano"] == "0000": 
            resumo["Ano"] = "2000"

        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        
        if is_p:
            resumo["Pasta"] = f"EMITIDOS_CLIENTE/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
        else:
            resumo["Pasta"] = f"RECEBIDOS_TERCEIROS/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Ano']}/{resumo['Mes']}"
            
        return resumo, is_p
        
    except Exception as e: 
        return None, False


_TIPOS_RESUMO_POR_SERIE = frozenset({"NF-e", "NFC-e", "NFS-e"})


def _incluir_em_resumo_por_serie(res, is_p, client_cnpj_clean: str) -> bool:
    """
    «Resumo por série» e buracos na mesma base: NF-e, NFC-e e NFS-e com emitente = CNPJ da barra lateral.
    """
    if not is_p:
        return False
    if str(res.get("Tipo", "")).strip() not in _TIPOS_RESUMO_POR_SERIE:
        return False
    emit = "".join(c for c in str(res.get("CNPJ_Emit", "")) if c.isdigit())[:14]
    cli = "".join(c for c in str(client_cnpj_clean or "") if c.isdigit())[:14]
    return len(cli) == 14 and emit == cli


# --- FUNÇÃO RECURSIVA OTIMIZADA PARA DISCO ---
def extrair_recursivo(conteudo_ou_file, nome_arquivo):
    """
    Abre .zip e devolve (nome_base, bytes) para cada .xml (incluindo dentro de ZIPs «matriosca»).
    ZIP dentro de ZIP: lê para memória (BytesIO) em vez de extract para disco — evita falhas no Windows
    (caminho longo, antivírus, pastas aninhadas no nome da entrada).
    """
    if not os.path.exists(TEMP_EXTRACT_DIR):
        os.makedirs(TEMP_EXTRACT_DIR)

    if nome_arquivo.lower().endswith(".zip"):
        try:
            if hasattr(conteudo_ou_file, "read"):
                file_obj = conteudo_ou_file
                if hasattr(file_obj, "seek"):
                    try:
                        file_obj.seek(0)
                    except (OSError, io.UnsupportedOperation):
                        pass
            else:
                file_obj = io.BytesIO(conteudo_ou_file)

            with zipfile.ZipFile(file_obj) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith("__MACOSX") or os.path.basename(sub_nome).startswith("."):
                        continue
                    if sub_nome.endswith("/") or sub_nome.endswith("\\"):
                        continue
                    base_sub = os.path.basename(sub_nome)
                    if not base_sub:
                        continue
                    if sub_nome.lower().endswith(".zip"):
                        try:
                            _zip_inner = z.read(sub_nome)
                        except (zipfile.BadZipFile, OSError, RuntimeError):
                            continue
                        if not _zip_inner:
                            continue
                        yield from extrair_recursivo(
                            io.BytesIO(_zip_inner),
                            base_sub,
                        )
                    elif sub_nome.lower().endswith(".xml"):
                        try:
                            yield (base_sub, z.read(sub_nome))
                        except (OSError, KeyError, RuntimeError):
                            continue
        except (zipfile.BadZipFile, OSError, RuntimeError):
            pass

    elif nome_arquivo.lower().endswith(".xml"):
        if hasattr(conteudo_ou_file, "read"):
            yield (os.path.basename(nome_arquivo), conteudo_ou_file.read())
        else:
            yield (os.path.basename(nome_arquivo), conteudo_ou_file)

# --- LIMPEZA DE PASTAS TEMPORÁRIAS ---
def limpar_arquivos_temp():
    """
    Só limpa subpastas dedicadas da app — **nunca** apaga ficheiros na pasta de trabalho (cwd),
    para não remover ZIP que o utilizador descarregou para a mesma pasta onde corre o Streamlit.
    """
    try:
        if os.path.exists(TEMP_EXTRACT_DIR):
            shutil.rmtree(TEMP_EXTRACT_DIR, ignore_errors=True)

        if os.path.exists(TEMP_UPLOADS_DIR):
            shutil.rmtree(TEMP_UPLOADS_DIR, ignore_errors=True)
        _garimpo_limpar_fontes_xml_memoria_sessao()
        try:
            st.session_state.pop(SESSION_KEY_EXTRA_DIGESTS, None)
        except Exception:
            pass
    except Exception:
        pass

# --- DIVISOR DE LOTES HTML (Para deixar botões organizados) ---
def chunk_list(lst, n):
    for i in range(0, len(lst), n): 
        yield lst[i:i + n]


def compactar_dataframe_memoria(df):
    """Reduz uso de RAM (categorias + downcast); seguro para filtros .str / .isin."""
    if df is None or df.empty:
        return df
    out = df.copy()
    n = len(out)
    for col in out.columns:
        if out[col].dtype != object and not str(out[col].dtype).startswith("string"):
            continue
        nu = out[col].nunique(dropna=False)
        if nu <= 1 or nu > min(4096, max(48, n // 2)):
            continue
        try:
            out[col] = out[col].astype("category")
        except (TypeError, ValueError):
            pass
    for col in out.select_dtypes(include=["float64"]).columns:
        out[col] = pd.to_numeric(out[col], downcast="float")
    for col in out.select_dtypes(include=["int64"]).columns:
        out[col] = pd.to_numeric(out[col], downcast="integer")
    return out


def dataframe_para_excel_bytes(df, sheet_name="Dados"):
    """
    Excel com as mesmas colunas do DataFrame (para download alinhado à tabela na tela).
    Em disco/temp cheio (errno 28) tenta openpyxl; em falha total devolve None.
    """
    if df is None or df.empty:
        return None
    sn = (sheet_name or "Dados")[:31]
    dfx = df.reset_index(drop=True)
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            dfx.to_excel(writer, sheet_name=sn, index=False)
        return buf.getvalue()
    except Exception as exc:
        if not _erro_sem_espaco_disco(exc):
            raise
    try:
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
            dfx.to_excel(writer, sheet_name=sn, index=False)
        return buf2.getvalue()
    except Exception:
        return None


# Limites de linhas por tabela no PDF do dashboard (evita ficheiros gigantes).
_DASH_PDF_MAX = {"resumo": 100, "tabela": 90, "geral": 75}
# Colunas preferidas no relatório geral no PDF (espelho legível da página da app).
_DASH_PDF_GERAL_COLS = [
    "Modelo",
    "Série",
    "Nota",
    "Data Emissão",
    "Status Final",
    "Valor",
    "Origem",
    "Chave",
]


def _format_celula_pdf_col(nome_col, val):
    if val is None:
        return "-"
    try:
        if pd.isna(val):
            return "-"
    except (TypeError, ValueError):
        pass
    nome = str(nome_col).strip().lower()
    nome_sem_acento = "".join(
        ch for ch in unicodedata.normalize("NFD", nome) if unicodedata.category(ch) != "Mn"
    )
    if "data" in nome_sem_acento and "emiss" in nome_sem_acento:
        f = _valor_data_emissao_dd_mm_yyyy(val)
        return f if f else "-"
    s = str(val).strip()
    if nome == "chave" and len(s) > 18:
        return f"...{s[-14:]}"
    if len(s) > 48:
        return s[:45] + "..."
    return s


def _preview_df_para_pdf(df, max_rows, colunas_preferidas=None, msg_se_vazio=None):
    """
    Prepara cabeçalhos e linhas para desenhar tabela no PDF.
    Retorno: cols, rows (listas de str), total, truncated, empty_msg opcional.
    """
    if df is None or df.empty:
        return {
            "cols": [],
            "rows": [],
            "total": 0,
            "truncated": False,
            "empty_msg": msg_se_vazio,
        }
    d = df.reset_index(drop=True)
    if colunas_preferidas:
        existentes = [c for c in colunas_preferidas if c in d.columns]
        if existentes:
            d = d[existentes]
        # se nenhuma coluna preferida existir, usa todas
        elif not existentes:
            pass
    cols = [str(c) for c in d.columns]
    total = len(d)
    truncated = total > max_rows
    sub = d.head(max_rows)
    rows = []
    for _, r in sub.iterrows():
        rows.append([_format_celula_pdf_col(c, r[c]) for c in d.columns])
    return {"cols": cols, "rows": rows, "total": total, "truncated": truncated, "empty_msg": None}


def _preview_terceiros_para_pdf(terc_cnt):
    if not terc_cnt:
        return {
            "cols": [],
            "rows": [],
            "total": 0,
            "truncated": False,
            "empty_msg": "Nenhum XML de terceiros no lote.",
        }
    rows = [[str(m), str(int(q))] for m, q in sorted(terc_cnt.items(), key=lambda x: x[0])]
    return {
        "cols": ["Modelo", "Quantidade"],
        "rows": rows,
        "total": len(rows),
        "truncated": False,
        "empty_msg": None,
    }


def coletar_kpis_dashboard():
    """Indicadores agregados para dashboard na app, Excel (folha Dashboard) e PDF."""
    rel = st.session_state.get("relatorio") or []
    sc = st.session_state.get("st_counts") or {}
    df_g = st.session_state.get("df_geral")
    df_r = st.session_state.get("df_resumo")
    df_f = st.session_state.get("df_faltantes")
    n_geral = len(df_g) if df_g is not None and not df_g.empty else 0
    n_bur = len(df_f) if df_f is not None and not df_f.empty else 0
    n_proprios = sum(1 for x in rel if "EMITIDOS_CLIENTE" in (x.get("Pasta") or ""))
    n_terc = sum(1 for x in rel if "RECEBIDOS_TERCEIROS" in (x.get("Pasta") or ""))
    terc_cnt = Counter()
    for x in rel:
        if "RECEBIDOS_TERCEIROS" in (x.get("Pasta") or ""):
            terc_cnt[x.get("Tipo") or "Outros"] += 1
    valor = 0.0
    if df_r is not None and not df_r.empty and "Valor Contábil (R$)" in df_r.columns:
        try:
            valor = float(df_r["Valor Contábil (R$)"].sum())
        except (TypeError, ValueError):
            valor = 0.0
    status_dist = {}
    if df_g is not None and not df_g.empty and "Status Final" in df_g.columns:
        vc = df_g["Status Final"].value_counts()
        status_dist = {str(k): int(v) for k, v in vc.items()}
    terc_status_dist = {}
    if (
        df_g is not None
        and not df_g.empty
        and "Origem" in df_g.columns
        and "Status Final" in df_g.columns
    ):
        _m_terc = df_g["Origem"].astype(str).str.contains("TERCEIROS", case=False, na=False)
        if _m_terc.any():
            _vc_terc = df_g.loc[_m_terc, "Status Final"].value_counts()
            terc_status_dist = {str(k): int(v) for k, v in _vc_terc.items()}
    ref_ok = bool(st.session_state.get("seq_ref_ultimos"))
    val_ok = bool(st.session_state.get("validation_done"))
    pares = [
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Linhas no relatório geral", n_geral),
        ("Itens no lote (relatório bruto)", len(rel)),
        ("Autorizadas (emissão própria)", int(sc.get("AUTORIZADAS", 0) or 0)),
        ("Canceladas (emissão própria)", int(sc.get("CANCELADOS", 0) or 0)),
        ("Denegadas (emissão própria)", int(sc.get("DENEGADOS", 0) or 0)),
        ("Rejeitadas (emissão própria)", int(sc.get("REJEITADOS", 0) or 0)),
        ("Inutilizadas (emissão própria)", int(sc.get("INUTILIZADOS", 0) or 0)),
        ("Buracos na sequência", n_bur),
        ("XML emissão própria (itens)", n_proprios),
        ("XML terceiros (itens)", n_terc),
        ("Valor contábil — resumo séries (R$)", round(valor, 2)),
        ("Referência último nº guardada", "Sim" if ref_ok else "Não"),
        ("Validação autenticidade", "Sim" if val_ok else "Não"),
    ]
    df_inu = st.session_state.get("df_inutilizadas")
    df_can = st.session_state.get("df_canceladas")
    df_aut = st.session_state.get("df_autorizadas")
    df_den = st.session_state.get("df_denegadas")
    df_rej = st.session_state.get("df_rejeitadas")
    df_g_pdf_geral = df_g
    if df_g is not None and not df_g.empty and "Origem" in df_g.columns:
        df_g_pdf_geral = df_g.loc[_mask_emissao_propria_df(df_g)].reset_index(drop=True)
    pdf_previews = {
        "resumo": _preview_df_para_pdf(
            df_r,
            _DASH_PDF_MAX["resumo"],
            msg_se_vazio="Sem linhas no resumo por série.",
        ),
        "terceiros": _preview_terceiros_para_pdf(dict(terc_cnt)),
        "buracos": _preview_df_para_pdf(
            df_f,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Tudo em ordem — nenhum buraco na auditoria.",
        ),
        "inutilizadas": _preview_df_para_pdf(
            df_inu,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Nenhuma inutilizada listada neste detalhe.",
        ),
        "canceladas": _preview_df_para_pdf(
            df_can,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Nenhuma cancelada listada neste detalhe.",
        ),
        "autorizadas": _preview_df_para_pdf(
            df_aut,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Nenhuma autorizada listada neste detalhe.",
        ),
        "denegadas": _preview_df_para_pdf(
            df_den,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Nenhuma denegada listada neste detalhe.",
        ),
        "rejeitadas": _preview_df_para_pdf(
            df_rej,
            _DASH_PDF_MAX["tabela"],
            msg_se_vazio="Nenhuma rejeitada listada neste detalhe.",
        ),
        "geral": _preview_df_para_pdf(
            df_g_pdf_geral,
            _DASH_PDF_MAX["geral"],
            _DASH_PDF_GERAL_COLS,
            msg_se_vazio="Relatório geral vazio (emissão própria).",
        ),
    }
    return {
        "pares": pares,
        "n_geral": n_geral,
        "n_bur": n_bur,
        "n_terc": n_terc,
        "n_docs": len(rel),
        "valor": valor,
        "status_dist": status_dist,
        "terc_cnt": dict(terc_cnt),
        "terc_status_dist": terc_status_dist,
        "sc": sc,
        "pdf_previews": pdf_previews,
    }


def _excel_nome_folha_seguro(nome, usados):
    """Nomes de folha Excel: máx. 31 caracteres; sem \\ / * ? : [ ]."""
    inv = frozenset('[]:*?/\\')
    base = "".join(c for c in str(nome) if c not in inv).strip()[:31] or "Sheet"
    out = base
    k = 2
    while out in usados:
        suf = f" ({k})"
        out = (base[: max(1, 31 - len(suf))] + suf).strip()
        k += 1
    usados.add(out)
    return out


def _excel_escrever_folha_df(writer, df, nome_desejado, usados):
    """Escreve um DataFrame na folha; se vazio, cabeçalhos ou nota curta."""
    sn = _excel_nome_folha_seguro(nome_desejado, usados)
    if df is None:
        pd.DataFrame({"Nota": ["Sem dados nesta vista."]}).to_excel(
            writer, sheet_name=sn, index=False
        )
        return
    d = df.reset_index(drop=True)
    if d.empty:
        if len(d.columns) > 0:
            d.to_excel(writer, sheet_name=sn, index=False)
        else:
            pd.DataFrame({"Nota": ["Sem registos nesta vista."]}).to_excel(
                writer, sheet_name=sn, index=False
            )
    else:
        d.to_excel(writer, sheet_name=sn, index=False)


def _excel_df_conta_par_modelo_serie(df, col_mod, col_ser):
    if df is None or df.empty or col_mod not in df.columns or col_ser not in df.columns:
        return {}
    d = df[[col_mod, col_ser]].copy()
    d["_k"] = d[col_mod].astype(str).str.strip() + "|" + d[col_ser].astype(str).str.strip()
    return d.groupby("_k").size().to_dict()


def _excel_fmt_milhar_pt(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "-"


def _excel_fmt_reais_pt_str(valor):
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return "R$ 0,00"
    neg = v < 0
    v = abs(v)
    cent = int(round(v * 100 + 1e-9))
    inte, frac = divmod(cent, 100)
    s = str(inte)
    parts = []
    while len(s) > 3:
        parts.insert(0, s[-3:])
        s = s[:-3]
    if s:
        parts.insert(0, s)
    body = ".".join(parts)
    out = f"R$ {body},{frac:02d}"
    return f"- {out}" if neg else out


def _excel_escrever_painel_fiscal(wb, kpi, usados_nomes):
    """
    Folha estilo dashboard (grelha tipo Excel + gráficos nativos), alinhada ao mock «Painel Fiscal».
    Dados reais do kpi / dataframes em sessão.
    """
    sn = _excel_nome_folha_seguro("Painel Fiscal", usados_nomes)
    ws = wb.add_worksheet(sn)
    # Parece mais “painel” / menos folha de cálculo (ideia de mock com Excel visual).
    try:
        ws.hide_gridlines(2)
    except Exception:
        pass

    pm = dict(kpi.get("pares") or [])
    sc = kpi.get("sc") or {}
    n_docs = int(kpi.get("n_docs") or 0)
    try:
        n_prop = int(pm.get("XML emissão própria (itens)", 0) or 0)
    except (TypeError, ValueError):
        n_prop = 0
    try:
        n_terc_xml = int(pm.get("XML terceiros (itens)", 0) or 0)
    except (TypeError, ValueError):
        n_terc_xml = 0
    aut = int(sc.get("AUTORIZADAS", 0) or 0)
    can = int(sc.get("CANCELADOS", 0) or 0)
    inu = int(sc.get("INUTILIZADOS", 0) or 0)
    den = int(sc.get("DENEGADOS", 0) or 0)
    rej = int(sc.get("REJEITADOS", 0) or 0)
    valor = float(kpi.get("valor") or 0)
    n_bur = int(kpi.get("n_bur") or 0)
    terc_cnt = kpi.get("terc_cnt") or {}
    if not isinstance(terc_cnt, dict):
        terc_cnt = {}

    df_r = st.session_state.get("df_resumo")
    df_aut = st.session_state.get("df_autorizadas")
    df_can = st.session_state.get("df_canceladas")
    df_inu = st.session_state.get("df_inutilizadas")
    df_bur = st.session_state.get("df_faltantes")
    if df_bur is not None and not df_bur.empty:
        if "Serie" in df_bur.columns and "Série" not in df_bur.columns:
            df_bur = df_bur.rename(columns={"Serie": "Série"})

    c_aut = _excel_df_conta_par_modelo_serie(df_aut, "Modelo", "Série")
    c_can = _excel_df_conta_par_modelo_serie(df_can, "Modelo", "Série")
    c_inu = _excel_df_conta_par_modelo_serie(df_inu, "Modelo", "Série")
    c_bur = _excel_df_conta_par_modelo_serie(df_bur, "Tipo", "Série")

    mes_ref = datetime.now().strftime("%m/%Y")
    ref_txt = "Sim" if pm.get("Referência último nº guardada") == "Sim" else "Não"
    val_txt = "Sim" if pm.get("Validação autenticidade") == "Sim" else "Não"
    _df_div = st.session_state.get("df_divergencias")
    if _df_div is not None and isinstance(_df_div, pd.DataFrame) and not _df_div.empty:
        n_div = len(_df_div)
    else:
        n_div = 0

    cor_fundo = "#FDFBF7"
    cor_vinho = "#5D1B36"
    cor_borda = "#A1869E"
    cor_texto = "#20232A"
    marrom = cor_vinho

    fmt_fundo = wb.add_format({"bg_color": cor_fundo})
    fmt_titulo_dash = wb.add_format(
        {
            "bold": True,
            "font_color": cor_vinho,
            "bg_color": cor_fundo,
            "font_size": 14,
            "valign": "vcenter",
            "align": "center",
        }
    )
    fmt_titulo_card = wb.add_format(
        {
            "bold": True,
            "font_color": cor_texto,
            "bg_color": "#FFFFFF",
            "top": 2,
            "left": 2,
            "right": 2,
            "top_color": cor_borda,
            "left_color": cor_borda,
            "right_color": cor_borda,
            "font_size": 9,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    fmt_valor_card = wb.add_format(
        {
            "bold": True,
            "font_color": cor_texto,
            "bg_color": "#FFFFFF",
            "left": 2,
            "right": 2,
            "left_color": cor_borda,
            "right_color": cor_borda,
            "font_size": 16,
            "align": "center",
            "valign": "vcenter",
        }
    )
    fmt_rodape_card = wb.add_format(
        {
            "font_color": cor_texto,
            "bg_color": "#FFFFFF",
            "bottom": 2,
            "left": 2,
            "right": 2,
            "bottom_color": cor_borda,
            "left_color": cor_borda,
            "right_color": cor_borda,
            "font_size": 9,
            "align": "center",
            "valign": "top",
            "text_wrap": True,
        }
    )
    fmt_kpi_t = wb.add_format(
        {
            "bold": True,
            "font_size": 10,
            "font_color": cor_vinho,
            "bg_color": cor_fundo,
            "border": 1,
            "border_color": cor_borda,
            "text_wrap": True,
            "valign": "vcenter",
        }
    )
    fmt_kpi_l = wb.add_format(
        {
            "font_size": 9,
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
            "bg_color": "#FFFFFF",
        }
    )
    fmt_tab_h = wb.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": marrom,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    fmt_tab_c = wb.add_format({"border": 1, "valign": "vcenter"})
    fmt_tab_warn = wb.add_format({"border": 1, "font_color": "#C62828", "bold": True})
    fmt_tab_ok = wb.add_format({"border": 1, "font_color": "#2E7D32"})
    fmt_foot = wb.add_format({"italic": True, "font_size": 9, "font_color": "#777777"})

    for row in range(0, 42):
        ws.set_row(row, 20, fmt_fundo)
    ws.set_row(0, 26)
    ws.set_row(1, 22)
    ws.set_column(0, 0, 4, fmt_fundo)
    ws.set_column(1, 12, 13, fmt_fundo)

    ws.merge_range(0, 1, 0, 12, "GARIMPEIRO | PAINEL FISCAL", fmt_titulo_dash)
    ws.merge_range(1, 1, 1, 12, "Olá! Bem-vinda à sua boutique de dados.", fmt_titulo_dash)

    terc_rodape = []
    try:
        itens = sorted(terc_cnt.items(), key=lambda x: int(x[1] or 0), reverse=True)
        for mod, q in itens[:2]:
            terc_rodape.append(f"{str(mod).strip()}: {_excel_fmt_milhar_pt(int(q or 0))}")
    except Exception:
        terc_rodape = []
    if not terc_rodape:
        terc_rodape = ["—"]

    ws.merge_range(3, 1, 3, 3, "TOTAL DE DOCUMENTOS\nLIDOS (NUM GERAL)", fmt_titulo_card)
    ws.merge_range(4, 1, 5, 3, _excel_fmt_milhar_pt(n_docs), fmt_valor_card)
    ws.merge_range(
        6,
        1,
        7,
        3,
        f"Próprios: {_excel_fmt_milhar_pt(n_prop)}\nTerceiros: {_excel_fmt_milhar_pt(n_terc_xml)}",
        fmt_rodape_card,
    )

    ws.merge_range(3, 4, 3, 6, "DETALHAMENTO\nEMISSÃO PRÓPRIA", fmt_titulo_card)
    ws.merge_range(4, 4, 5, 6, f"{_excel_fmt_milhar_pt(aut)} Aut.", fmt_valor_card)
    ws.merge_range(
        6,
        4,
        7,
        6,
        f"Canceladas: {_excel_fmt_milhar_pt(can)}\nInutilizadas: {_excel_fmt_milhar_pt(inu)}\n"
        f"Denegadas: {_excel_fmt_milhar_pt(den)}\nRejeitadas: {_excel_fmt_milhar_pt(rej)}",
        fmt_rodape_card,
    )

    ws.merge_range(3, 7, 3, 9, "DETALHAMENTO\nDOCUMENTOS TERCEIROS", fmt_titulo_card)
    ws.merge_range(4, 7, 5, 9, f"{_excel_fmt_milhar_pt(n_terc_xml)} Terc.", fmt_valor_card)
    ws.merge_range(6, 7, 7, 9, "\n".join(terc_rodape), fmt_rodape_card)

    ws.merge_range(3, 10, 3, 12, "VOLUME\nFINANCEIRO", fmt_titulo_card)
    ws.merge_range(4, 10, 5, 12, _excel_fmt_reais_pt_str(valor), fmt_valor_card)
    ws.merge_range(6, 10, 7, 12, "", fmt_rodape_card)

    chart_row = 10

    # Dados ocultos para gráficos (colunas Q-S = 16+)
    hid_c = 16
    r0 = 40
    ws.write(r0, hid_c, "Categoria", fmt_kpi_l)
    ws.write(r0, hid_c + 1, "Valor", fmt_kpi_l)
    ws.write(r0 + 1, hid_c, "Autorizadas")
    ws.write_number(r0 + 1, hid_c + 1, max(0, aut))
    ws.write(r0 + 2, hid_c, "Canceladas")
    ws.write_number(r0 + 2, hid_c + 1, max(0, can))
    ws.write(r0 + 3, hid_c, "Inutilizadas")
    ws.write_number(r0 + 3, hid_c + 1, max(0, inu))
    ws.write(r0 + 4, hid_c, "Denegadas")
    ws.write_number(r0 + 4, hid_c + 1, max(0, den))
    ws.write(r0 + 5, hid_c, "Rejeitadas")
    ws.write_number(r0 + 5, hid_c + 1, max(0, rej))

    r1 = r0 + 8
    terc_items = sorted(terc_cnt.items(), key=lambda x: -x[1])[:8]
    if not terc_items:
        ws.write(r1, hid_c, "—")
        ws.write_number(r1, hid_c + 1, 1)
        n_trows = 1
    else:
        n_trows = 0
        for mod, q in terc_items:
            ws.write(r1 + n_trows, hid_c, str(mod))
            ws.write_number(r1 + n_trows, hid_c + 1, int(q))
            n_trows += 1

    r2 = r1 + max(n_trows, 1) + 2
    ws.write(r2, hid_c, "Própria (aut.)")
    ws.write_number(r2, hid_c + 1, max(0, aut))
    ws.write(r2 + 1, hid_c, "Terceiros (XML)")
    ws.write_number(r2 + 1, hid_c + 1, max(0, n_terc_xml))

    ws.set_column(hid_c, hid_c + 1, None, None, {"hidden": True})

    ch1 = wb.add_chart({"type": "doughnut"})
    ch1.add_series(
        {
            "name": "Emissão própria",
            "categories": [sn, r0 + 1, hid_c, r0 + 5, hid_c],
            "values": [sn, r0 + 1, hid_c + 1, r0 + 5, hid_c + 1],
        }
    )
    ch1.set_title({"name": f"STATUS DE EMISSÃO PRÓPRIA ({mes_ref})"})
    ch1.set_style(10)
    ws.insert_chart(chart_row, 0, ch1, {"x_scale": 0.95, "y_scale": 0.95})

    ch2 = wb.add_chart({"type": "doughnut"})
    last_tr = r1 + n_trows - 1
    ch2.add_series(
        {
            "name": "Terceiros",
            "categories": [sn, r1, hid_c, last_tr, hid_c],
            "values": [sn, r1, hid_c + 1, last_tr, hid_c + 1],
        }
    )
    ch2.set_title({"name": f"DISTRIBUIÇÃO TERCEIROS POR MODELO ({mes_ref})"})
    ch2.set_style(10)
    ws.insert_chart(chart_row, 5, ch2, {"x_scale": 0.95, "y_scale": 0.95})

    ch3 = wb.add_chart({"type": "area"})
    ch3.add_series(
        {
            "name": "No lote",
            "categories": [sn, r2, hid_c, r2 + 1, hid_c],
            "values": [sn, r2, hid_c + 1, r2 + 1, hid_c + 1],
            "fill": {"color": cor_borda},
            "line": {"color": cor_vinho},
        }
    )
    ch3.set_title({"name": "RESUMO RÁPIDO (autorizadas própria vs XML terceiros)"})
    ch3.set_legend({"none": True})
    ws.insert_chart(chart_row, 10, ch3, {"x_scale": 0.95, "y_scale": 0.95})

    # Alertas de série (séries com buracos) — à direita, acima dos gráficos
    alert_row = 3
    ws.write(alert_row, 14, "ALERTAS DE SÉRIE", fmt_kpi_t)
    if df_bur is not None and not df_bur.empty and "Tipo" in df_bur.columns and "Série" in df_bur.columns:
        gb = df_bur.groupby(["Tipo", "Série"]).size().reset_index(name="n")
        gb = gb.sort_values("n", ascending=False).head(8)
        ar = alert_row + 1
        for _, rr in gb.iterrows():
            ws.write(ar, 14, f"{rr['Tipo']} sér. {rr['Série']}: {int(rr['n'])} buraco(s)", fmt_tab_warn)
            ar += 1
            if ar > alert_row + 6:
                break
        if ar == alert_row + 1:
            ws.write(ar, 14, "Nenhum buraco listado.", fmt_tab_ok)
    else:
        ws.write(alert_row + 1, 14, "Sem buracos ou dados indisponíveis.", fmt_tab_ok)

    ctx_r = alert_row + 10
    ws.write(ctx_r, 14, f"Último nº guardado: {ref_txt}", fmt_kpi_l)
    ws.write(ctx_r + 1, 14, f"Autenticidade: {val_txt}", fmt_kpi_l)
    if n_div:
        ws.write(ctx_r + 2, 14, f"Divergências XML×Sefaz: {n_div}", fmt_tab_warn)

    # Tabela totalizador (abaixo dos gráficos)
    t_row = chart_row + 16
    ws.merge_range(
        t_row,
        0,
        t_row,
        8,
        f"TOTALIZADOR DE SÉRIE E FAIXAS ({mes_ref})",
        fmt_kpi_t,
    )
    t_row += 1
    headers = [
        "Modelo",
        "Série",
        "Faixa (início–fim)",
        "Qtd lidos",
        "Autorizadas",
        "Canceladas",
        "Inutilizadas",
        "Buracos",
        "OK?",
    ]
    for c, h in enumerate(headers):
        ws.write(t_row, c, h, fmt_tab_h)
    t_row += 1

    if df_r is not None and not df_r.empty:
        doc_col = "Documento" if "Documento" in df_r.columns else None
        ser_col = "Série" if "Série" in df_r.columns else None
        if doc_col and ser_col:
            for _, row in df_r.iterrows():
                mod = str(row.get(doc_col, "")).strip()
                ser = str(row.get(ser_col, "")).strip()
                k = f"{mod}|{ser}"
                ini = row.get("Início", "")
                fim = row.get("Fim", "")
                faixa = f"{ini} – {fim}"
                qtd = row.get("Quantidade", "")
                na = int(c_aut.get(k, 0))
                nc = int(c_can.get(k, 0))
                ni = int(c_inu.get(k, 0))
                nb = int(c_bur.get(k, 0))
                ok = "✓" if nb == 0 else "▲"
                fmt_end = fmt_tab_ok if nb == 0 else fmt_tab_warn
                ws.write(t_row, 0, mod, fmt_tab_c)
                ws.write(t_row, 1, ser, fmt_tab_c)
                ws.write(t_row, 2, faixa, fmt_tab_c)
                ws.write(t_row, 3, qtd, fmt_tab_c)
                ws.write_number(t_row, 4, na, fmt_tab_c)
                ws.write_number(t_row, 5, nc, fmt_tab_c)
                ws.write_number(t_row, 6, ni, fmt_tab_c)
                ws.write_number(t_row, 7, nb, fmt_tab_c)
                ws.write(t_row, 8, ok, fmt_end)
                t_row += 1
        else:
            ws.merge_range(t_row, 0, t_row, 8, "Resumo por série sem colunas esperadas.", fmt_tab_c)
            t_row += 1
    else:
        ws.merge_range(t_row, 0, t_row, 8, "Sem resumo por série neste lote.", fmt_tab_c)
        t_row += 1

    t_row += 1
    ws.merge_range(t_row, 0, t_row, 8, "Garimpeiro · Painel gerado a partir do lote atual (mesmos dados que na app).", fmt_foot)

    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 2, 24)
    ws.set_column(3, 8, 11)
    ws.set_column(14, 15, 20)


def _excel_relatorio_geral_openpyxl_fallback_bytes(
    df_g,
    df_bur,
    df_inu,
    df_can,
    df_aut,
    df_den,
    df_rej,
    df_cte,
    df_terc_rows,
    *,
    omit_bur_inu: bool,
    kpi: dict,
    incluir_painel_fiscal: bool,
):
    """
    Mesmas folhas de dados que o livro principal, sem layout xlsxwriter (dashboard formatado / painel fiscal).
    Usado quando xlsxwriter falha por disco ou temp cheio (errno 28).
    """
    buf = io.BytesIO()
    usados = set()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            _excel_escrever_folha_df(writer, df_g, "Geral", usados)
            if not omit_bur_inu:
                _excel_escrever_folha_df(writer, df_bur, "Buracos", usados)
                _excel_escrever_folha_df(writer, df_inu, "Inutilizadas", usados)
            _excel_escrever_folha_df(writer, df_can, "Canceladas", usados)
            _excel_escrever_folha_df(writer, df_aut, "Autorizadas", usados)
            _excel_escrever_folha_df(writer, df_den, "Denegadas", usados)
            _excel_escrever_folha_df(writer, df_rej, "Rejeitadas", usados)
            _excel_escrever_folha_df(writer, df_cte, "CT-e e CT-e OS", usados)
            _excel_escrever_folha_df(writer, df_terc_rows, "Terceiros lidas", usados)
            par = kpi.get("pares") or []
            if par:
                df_d = pd.DataFrame(list(par), columns=["Indicador", "Valor"])
            else:
                df_d = pd.DataFrame(
                    {
                        "Aviso": [
                            "Export alternativo (sem espaço para xlsxwriter em disco/Temp). "
                            "Liberte espaço em C: ou defina TEMP/TMP noutro disco."
                        ]
                    }
                )
            _excel_escrever_folha_df(writer, df_d, "Dashboard", usados)
            df_r = st.session_state.get("df_resumo")
            if df_r is not None and isinstance(df_r, pd.DataFrame) and not df_r.empty:
                _excel_escrever_folha_df(writer, df_r, "Resumo por série", usados)
            tc = kpi.get("terc_cnt") or {}
            if tc:
                df_tc = pd.DataFrame(
                    [
                        {"Modelo": m, "Quantidade": int(q)}
                        for m, q in sorted(tc.items(), key=lambda x: str(x[0]))
                    ]
                )
                _excel_escrever_folha_df(writer, df_tc, "Terceiros qtd", usados)
            if incluir_painel_fiscal:
                df_pf = pd.DataFrame(
                    {
                        "Nota": [
                            "Painel fiscal: só no export completo (xlsxwriter). "
                            "Liberte espaço em C: (Temp) e exporte de novo."
                        ]
                    }
                )
                _excel_escrever_folha_df(writer, df_pf, "Painel fiscal", usados)
        return buf.getvalue()
    except Exception:
        return None


def excel_relatorio_geral_com_dashboard_bytes(
    df_geral, *, incluir_painel_fiscal: bool = True, folhas_detalhe: dict | None = None
):
    """
    Excel com várias folhas alinhadas às abas da página da app:
    Geral, Buracos, Inutilizadas, Canceladas, Autorizadas, Denegadas, Rejeitadas, CT-e e CT-e OS, Terceiros lidas, Dashboard;
    opcionalmente Painel Fiscal (incluir_painel_fiscal=False para pacote contabilidade).
    Se ``folhas_detalhe`` for um dict com chaves df_bur, df_inu, df_can, df_aut, df_den, df_rej,
    usa esses DataFrames em vez dos da sessão (ex.: livro só «terceiros»). Se df_bur e df_inu
    vierem vazios (caso típico terceiros), as folhas Buracos e Inutilizadas não são criadas.
    """
    if df_geral is None or df_geral.empty:
        return None
    kpi = coletar_kpis_dashboard()
    usados_nomes = set()

    if folhas_detalhe is None:
        df_bur = st.session_state.get("df_faltantes")
        df_inu = st.session_state.get("df_inutilizadas")
        df_can = st.session_state.get("df_canceladas")
        df_aut = st.session_state.get("df_autorizadas")
        df_den = st.session_state.get("df_denegadas")
        df_rej = st.session_state.get("df_rejeitadas")
    else:
        df_bur = folhas_detalhe.get("df_bur")
        df_inu = folhas_detalhe.get("df_inu")
        df_can = folhas_detalhe.get("df_can")
        df_aut = folhas_detalhe.get("df_aut")
        df_den = folhas_detalhe.get("df_den")
        df_rej = folhas_detalhe.get("df_rej")
        if df_bur is None or not isinstance(df_bur, pd.DataFrame):
            df_bur = pd.DataFrame()
        if df_inu is None or not isinstance(df_inu, pd.DataFrame):
            df_inu = pd.DataFrame()
        if df_can is None or not isinstance(df_can, pd.DataFrame):
            df_can = pd.DataFrame()
        if df_aut is None or not isinstance(df_aut, pd.DataFrame):
            df_aut = pd.DataFrame()
        if df_den is None or not isinstance(df_den, pd.DataFrame):
            df_den = pd.DataFrame()
        if df_rej is None or not isinstance(df_rej, pd.DataFrame):
            df_rej = pd.DataFrame()

    df_g = df_geral.reset_index(drop=True)
    df_g = _df_com_data_emissao_dd_mm_yyyy(df_g)
    if "Modelo" in df_g.columns:
        _m = df_g["Modelo"].astype(str).str.strip()
        df_cte = df_g[_m.isin(("CT-e", "CT-e OS"))].copy()
    else:
        df_cte = pd.DataFrame()
    if "Origem" in df_g.columns:
        df_terc_rows = df_g[
            df_g["Origem"].astype(str).str.contains("TERCEIROS", case=False, na=False)
        ].copy()
    else:
        df_terc_rows = pd.DataFrame()

    df_bur = _df_com_data_emissao_dd_mm_yyyy(df_bur)
    df_inu = _df_com_data_emissao_dd_mm_yyyy(df_inu)
    df_can = _df_com_data_emissao_dd_mm_yyyy(df_can)
    df_aut = _df_com_data_emissao_dd_mm_yyyy(df_aut)
    df_den = _df_com_data_emissao_dd_mm_yyyy(df_den)
    df_rej = _df_com_data_emissao_dd_mm_yyyy(df_rej)

    # Livro só terceiros: buracos/inutilizadas não existem — não criar folhas vazias.
    _omit_bur_inu = (
        folhas_detalhe is not None
        and df_bur.empty
        and df_inu.empty
    )

    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            _excel_escrever_folha_df(writer, df_g, "Geral", usados_nomes)
            if not _omit_bur_inu:
                _excel_escrever_folha_df(writer, df_bur, "Buracos", usados_nomes)
                _excel_escrever_folha_df(writer, df_inu, "Inutilizadas", usados_nomes)
            _excel_escrever_folha_df(writer, df_can, "Canceladas", usados_nomes)
            _excel_escrever_folha_df(writer, df_aut, "Autorizadas", usados_nomes)
            _excel_escrever_folha_df(writer, df_den, "Denegadas", usados_nomes)
            _excel_escrever_folha_df(writer, df_rej, "Rejeitadas", usados_nomes)
            _excel_escrever_folha_df(writer, df_cte, "CT-e e CT-e OS", usados_nomes)
            _excel_escrever_folha_df(writer, df_terc_rows, "Terceiros lidas", usados_nomes)

            wb = writer.book
            dash_sn = _excel_nome_folha_seguro("Dashboard", usados_nomes)
            ws = wb.add_worksheet(dash_sn)
            title_f = wb.add_format(
                {"bold": True, "font_size": 16, "font_color": "#AD1457", "valign": "vcenter"}
            )
            hdr_f = wb.add_format(
                {"bold": True, "bg_color": "#F8BBD0", "border": 1, "valign": "vcenter"}
            )
            cell_f = wb.add_format({"border": 1, "valign": "vcenter"})
            sub_f = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#FCE4EC", "border": 1})

            ws.merge_range(0, 0, 0, 3, "Garimpeiro — Dashboard", title_f)
            ws.set_row(0, 26)
            row = 2
            ws.write(row, 0, "Indicador", hdr_f)
            ws.write(row, 1, "Valor", hdr_f)
            row += 1
            for lab, val in kpi["pares"]:
                ws.write(row, 0, lab, cell_f)
                ws.write(row, 1, val, cell_f)
                row += 1
            row += 1

            df_r = st.session_state.get("df_resumo")
            if df_r is not None and not df_r.empty:
                last_c = max(5, len(df_r.columns) - 1)
                ws.merge_range(
                    row,
                    0,
                    row,
                    last_c,
                    "Resumo por série (NF-e / NFC-e / NFS-e, emitente = CNPJ da barra lateral)",
                    sub_f,
                )
                row += 1
                for c, colname in enumerate(df_r.columns):
                    ws.write(row, c, str(colname), hdr_f)
                row += 1
                for _, rr in df_r.iterrows():
                    for c, colname in enumerate(df_r.columns):
                        v = rr[colname]
                        ws.write(row, c, v, cell_f)
                    row += 1
                row += 1

            tc = kpi.get("terc_cnt") or {}
            if tc:
                ws.merge_range(row, 0, row, 2, "Terceiros — quantidade por modelo", sub_f)
                row += 1
                ws.write(row, 0, "Modelo", hdr_f)
                ws.write(row, 1, "Quantidade", hdr_f)
                row += 1
                for mod, q in sorted(tc.items(), key=lambda x: x[0]):
                    ws.write(row, 0, mod, cell_f)
                    ws.write(row, 1, int(q), cell_f)
                    row += 1

            ws.set_column(0, 0, 42)
            ws.set_column(1, 1, 22)

            if incluir_painel_fiscal:
                _excel_escrever_painel_fiscal(wb, kpi, usados_nomes)

        return buf.getvalue()
    except Exception as exc:
        if not _erro_sem_espaco_disco(exc):
            raise
    return _excel_relatorio_geral_openpyxl_fallback_bytes(
        df_g,
        df_bur,
        df_inu,
        df_can,
        df_aut,
        df_den,
        df_rej,
        df_cte,
        df_terc_rows,
        omit_bur_inu=_omit_bur_inu,
        kpi=kpi,
        incluir_painel_fiscal=incluir_painel_fiscal,
    )


def _pdf_ascii_seguro(txt):
    if txt is None:
        return ""
    s = str(txt)
    return (
        unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii") or "-"
    )


def _pdf_txt(pdf, s, use_dejavu):
    if use_dejavu:
        return str(s)
    return _pdf_ascii_seguro(s)


def _pdf_font(pdf, use_dejavu, style="", size=10):
    fam = "DejaVu" if use_dejavu else "Helvetica"
    try:
        pdf.set_font(fam, style, size)
    except Exception:
        if use_dejavu and style == "B":
            try:
                pdf.set_font("DejaVu", "", min(size + 1.4, 16))
            except Exception:
                pdf.set_font("Helvetica", "B", size)
        else:
            pdf.set_font("Helvetica", "" if not style else "B", size)


def _pdf_multi_texto_largura_total(pdf, altura_linha, texto, use_dejavu):
    """
    Texto em bloco na largura útil da página.
    Evita FPDFException (fpdf2): multi_cell(0, ...) com get_x() à direita — largura útil ~0.
    """
    pdf.set_x(pdf.l_margin)
    w = max(30.0, pdf.w - pdf.l_margin - pdf.r_margin)
    pdf.multi_cell(w, altura_linha, _pdf_txt(pdf, texto, use_dejavu))


def _pdf_cabecalho_executivo_painel(pdf, use_dejavu, linha_extra=None):
    """Cabeçalho alinhado ao painel da app: off-white, texto vinho."""
    y = pdf.get_y()
    w = pdf.w - pdf.l_margin - pdf.r_margin
    h = 20.0
    pdf.set_fill_color(253, 251, 247)
    pdf.set_draw_color(220, 210, 200)
    pdf.set_line_width(0.25)
    pdf.rect(pdf.l_margin, y, w, h, "DF")
    _pdf_font(pdf, use_dejavu, "B", 12)
    pdf.set_text_color(93, 27, 54)
    pdf.set_xy(pdf.l_margin + 4, y + 3.5)
    pdf.cell(w - 8, 6, _pdf_txt(pdf, "GARIMPEIRO · PAINEL DO LOTE", use_dejavu), ln=False)
    _pdf_font(pdf, use_dejavu, "", 7.5)
    pdf.set_text_color(90, 75, 85)
    sub = "Boutique de dados · mesmo desenho que na página principal do Garimpeiro."
    if linha_extra:
        sub = f"{linha_extra} · {sub}"
    pdf.set_xy(pdf.l_margin + 4, y + 10)
    pdf.cell(w - 8, 4.5, _pdf_txt(pdf, sub, use_dejavu), ln=False)
    dt = datetime.now().strftime("%d/%m/%Y · %H:%M")
    _pdf_font(pdf, use_dejavu, "", 6.5)
    pdf.set_text_color(130, 115, 125)
    pdf.set_xy(pdf.l_margin + 4, y + 15)
    pdf.cell(w - 8, 4, _pdf_txt(pdf, dt, use_dejavu), ln=False)
    pdf.set_text_color(32, 35, 42)
    pdf.set_xy(pdf.l_margin, y + h + 3)


def _pdf_quatro_kpi_cards_executivo(pdf, kpi, use_dejavu):
    """Quatro cartões KPI como no Streamlit / Excel (borda #A1869E, fundo branco)."""
    pm = dict(kpi.get("pares") or [])
    try:
        n_prop = int(pm.get("XML emissão própria (itens)", 0) or 0)
    except (TypeError, ValueError):
        n_prop = 0
    try:
        n_terc_xml = int(pm.get("XML terceiros (itens)", 0) or 0)
    except (TypeError, ValueError):
        n_terc_xml = 0
    n_docs = int(kpi.get("n_docs") or 0)
    sc = kpi.get("sc") or {}
    aut = int(sc.get("AUTORIZADAS", 0) or 0)
    can = int(sc.get("CANCELADOS", 0) or 0)
    inu = int(sc.get("INUTILIZADOS", 0) or 0)
    valor = float(kpi.get("valor") or 0.0)
    terc_cnt = kpi.get("terc_cnt") or {}
    if not isinstance(terc_cnt, dict):
        terc_cnt = {}
    terc_linhas = []
    try:
        itens = sorted(terc_cnt.items(), key=lambda x: int(x[1] or 0), reverse=True)
        for mod, q in itens[:2]:
            terc_linhas.append(f"{str(mod).strip()}: {_excel_fmt_milhar_pt(int(q or 0))}")
    except Exception:
        terc_linhas = []
    if not terc_linhas:
        terc_linhas = ["—"]

    specs = [
        (
            "TOTAL DE DOCUMENTOS\nLIDOS (N.º GERAL)",
            _excel_fmt_milhar_pt(n_docs),
            f"Próprios: {_excel_fmt_milhar_pt(n_prop)}\nTerceiros: {_excel_fmt_milhar_pt(n_terc_xml)}",
        ),
        (
            "DETALHAMENTO\nEMISSÃO PRÓPRIA",
            f"{_excel_fmt_milhar_pt(aut)} Aut.",
            f"Canceladas: {_excel_fmt_milhar_pt(can)}\nInutilizadas: {_excel_fmt_milhar_pt(inu)}",
        ),
        (
            "DETALHAMENTO\nDOCUMENTOS TERCEIROS",
            f"{_excel_fmt_milhar_pt(n_terc_xml)} Terc.",
            "\n".join(terc_linhas),
        ),
        (
            "VOLUME\nFINANCEIRO",
            _excel_fmt_reais_pt_str(valor),
            "Soma do valor contábil no resumo por série",
        ),
    ]

    margin = pdf.l_margin
    full = pdf.w - margin - pdf.r_margin
    gap = 2.4
    ncols = 4
    cw = (full - gap * (ncols - 1)) / ncols
    y0 = pdf.get_y()
    card_h = 38.0
    title_h = 10.5
    foot_h = 12.0
    val_h = card_h - title_h - foot_h
    bor = (161, 134, 158)
    sep = (210, 195, 205)

    pdf.set_line_width(0.45)
    for i, (tit, val, foot) in enumerate(specs):
        x = margin + i * (cw + gap)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(*bor)
        pdf.rect(x, y0, cw, card_h, "D")
        pdf.set_draw_color(*sep)
        pdf.set_line_width(0.15)
        pdf.line(x, y0 + title_h, x + cw, y0 + title_h)
        pdf.line(x, y0 + title_h + val_h, x + cw, y0 + title_h + val_h)
        pdf.set_line_width(0.45)

        _pdf_font(pdf, use_dejavu, "B", 5.8)
        pdf.set_text_color(32, 35, 42)
        yl = y0 + 2.2
        for part in str(tit).split("\n")[:2]:
            pdf.set_xy(x + 1.5, yl)
            pdf.cell(cw - 3, 3.4, _pdf_txt(pdf, part.strip(), use_dejavu), align="C", ln=False)
            yl += 3.5

        _pdf_font(pdf, use_dejavu, "B", 9.5 if len(val) < 14 else 8.0)
        pdf.set_text_color(32, 35, 42)
        pdf.set_xy(x + 1.5, y0 + title_h + 3.5)
        pdf.cell(cw - 3, val_h - 4, _pdf_txt(pdf, val, use_dejavu), align="C", ln=False)

        _pdf_font(pdf, use_dejavu, "", 5.6)
        pdf.set_text_color(61, 53, 64)
        yf = y0 + title_h + val_h + 2.0
        for fl in str(foot).split("\n")[:3]:
            pdf.set_xy(x + 1.5, yf)
            pdf.cell(cw - 3, 3.2, _pdf_txt(pdf, fl.strip(), use_dejavu), align="C", ln=False)
            yf += 3.25

    pdf.set_xy(pdf.l_margin, y0 + card_h + 4)


def _pdf_faixa_buracos_executivo(pdf, n_bur, use_dejavu):
    """Uma linha de contexto: buracos (saiu do antigo quadro 2×2)."""
    margin = pdf.l_margin
    full = pdf.w - margin - pdf.r_margin
    y = pdf.get_y()
    h = 7.5
    pdf.set_fill_color(255, 255, 255)
    pdf.set_draw_color(161, 134, 158)
    pdf.set_line_width(0.2)
    pdf.rect(margin, y, full, h, "D")
    _pdf_font(pdf, use_dejavu, "", 7)
    pdf.set_text_color(93, 27, 54)
    pdf.set_xy(margin + 3, y + 2)
    pdf.cell(
        full - 6,
        4,
        _pdf_txt(
            pdf,
            f"Buracos na sequência (emissão própria): {_excel_fmt_milhar_pt(int(n_bur or 0))}",
            use_dejavu,
        ),
        ln=False,
    )
    pdf.set_xy(pdf.l_margin, y + h + 3)


def _pdf_serie_cards_emissao_propria(pdf, df_resumo, use_dejavu):
    """Cartões no mesmo padrão dos KPI: cada série da emissão própria com faixa inicial–final."""
    pdf.ln(1)
    _pdf_font(pdf, use_dejavu, "B", 9)
    pdf.set_text_color(93, 27, 54)
    pdf.set_x(pdf.l_margin)
    pdf.cell(0, 5, _pdf_txt(pdf, "Emissão própria — séries e faixas de numeração", use_dejavu), ln=True)
    _pdf_font(pdf, use_dejavu, "", 7)
    pdf.set_text_color(90, 75, 85)
    pdf.set_x(pdf.l_margin)
    pdf.cell(
        0,
        4,
        _pdf_txt(
            pdf,
            "Um cartão por modelo e série: menor e maior número de nota encontrados nos ficheiros XML do lote.",
            use_dejavu,
        ),
        ln=True,
    )
    pdf.ln(1.5)

    if df_resumo is None or not isinstance(df_resumo, pd.DataFrame) or df_resumo.empty:
        _pdf_font(pdf, use_dejavu, "", 8)
        pdf.set_text_color(130, 115, 125)
        pdf.set_x(pdf.l_margin)
        pdf.cell(0, 5, _pdf_txt(pdf, "Sem linhas no resumo por série.", use_dejavu), ln=True)
        pdf.ln(2)
        return

    doc_col = "Documento" if "Documento" in df_resumo.columns else None
    ser_col = "Série" if "Série" in df_resumo.columns else ("Serie" if "Serie" in df_resumo.columns else None)
    if not doc_col or not ser_col:
        _pdf_font(pdf, use_dejavu, "", 8)
        pdf.set_text_color(180, 90, 90)
        pdf.set_x(pdf.l_margin)
        pdf.cell(
            0,
            5,
            _pdf_txt(pdf, "Resumo sem colunas Documento/Série — cartões não gerados.", use_dejavu),
            ln=True,
        )
        pdf.ln(2)
        return

    ini_col = "Início" if "Início" in df_resumo.columns else ("Inicio" if "Inicio" in df_resumo.columns else None)
    fim_col = "Fim" if "Fim" in df_resumo.columns else None
    qtd_col = "Quantidade" if "Quantidade" in df_resumo.columns else None
    val_col = "Valor Contábil (R$)" if "Valor Contábil (R$)" in df_resumo.columns else None

    margin = pdf.l_margin
    full = pdf.w - margin - pdf.r_margin
    gap_h = 2.4
    gap_v = 2.8
    ncols = 2
    cw = (full - gap_h * (ncols - 1)) / ncols
    card_h = 35.0
    title_h = 10.0
    foot_h = 10.5
    val_h = card_h - title_h - foot_h
    bor = (161, 134, 158)
    sep = (210, 195, 205)

    recs = [row for _, row in df_resumo.iterrows()]
    idx = 0
    while idx < len(recs):
        y0 = pdf.get_y()
        if y0 + card_h > 275:
            pdf.add_page()
            y0 = pdf.get_y()
        for col in range(ncols):
            if idx >= len(recs):
                break
            row = recs[idx]
            x = margin + col * (cw + gap_h)
            doc = str(row.get(doc_col, "") or "").strip()
            ser = str(row.get(ser_col, "") or "").strip()
            if len(doc) > 24:
                doc = doc[:22] + "…"
            tit_a = doc if doc else "—"
            tit_b = f"Série {ser}" if ser else "Série —"
            ini = row.get(ini_col, "") if ini_col else ""
            fim = row.get(fim_col, "") if fim_col else ""
            foot_lines = []
            if qtd_col is not None and row.get(qtd_col, "") != "" and str(row.get(qtd_col)).strip() != "":
                try:
                    qv = int(row[qtd_col])
                    foot_lines.append(f"Quantidade: {_excel_fmt_milhar_pt(qv)}")
                except (TypeError, ValueError):
                    foot_lines.append(f"Quantidade: {row.get(qtd_col)}")
            if val_col is not None:
                try:
                    vv = float(row[val_col])
                    foot_lines.append(_excel_fmt_reais_pt_str(vv))
                except (TypeError, ValueError):
                    pass

            pdf.set_line_width(0.45)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(*bor)
            pdf.rect(x, y0, cw, card_h, "D")
            pdf.set_draw_color(*sep)
            pdf.set_line_width(0.15)
            pdf.line(x, y0 + title_h, x + cw, y0 + title_h)
            pdf.line(x, y0 + title_h + val_h, x + cw, y0 + title_h + val_h)
            pdf.set_line_width(0.45)

            _pdf_font(pdf, use_dejavu, "B", 6.5)
            pdf.set_text_color(32, 35, 42)
            yl = y0 + 2.0
            pdf.set_xy(x + 1.5, yl)
            pdf.cell(cw - 3, 3.3, _pdf_txt(pdf, tit_a, use_dejavu), align="C", ln=False)
            yl += 3.5
            pdf.set_xy(x + 1.5, yl)
            pdf.cell(cw - 3, 3.3, _pdf_txt(pdf, tit_b, use_dejavu), align="C", ln=False)

            inner_w = cw - 3.0
            gutter = 1.2
            half_w = (inner_w - gutter) / 2.0
            x_left = x + 1.5
            x_right = x + 1.5 + half_w + gutter
            def _pdf_val_nota_intervalo(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                s = str(v).strip()
                return "" if not s or s.lower() == "nan" else s

            ini_s = _pdf_val_nota_intervalo(ini)
            fim_s = _pdf_val_nota_intervalo(fim)
            if not ini_s and not fim_s:
                ini_s, fim_s = "—", "—"
            elif not ini_s:
                ini_s = "—"
            elif not fim_s:
                fim_s = "—"
            fs_ini = 11.0 if len(ini_s) < 11 else (9.0 if len(ini_s) < 16 else 7.5)
            fs_fim = 11.0 if len(fim_s) < 11 else (9.0 if len(fim_s) < 16 else 7.5)
            y_mid = y0 + title_h
            y_num = y_mid + 2.8
            h_num_cell = 6.8
            _pdf_font(pdf, use_dejavu, "B", fs_ini)
            pdf.set_text_color(32, 35, 42)
            pdf.set_xy(x_left, y_num)
            pdf.cell(half_w, h_num_cell, _pdf_txt(pdf, ini_s, use_dejavu), align="C", ln=False)
            _pdf_font(pdf, use_dejavu, "B", fs_fim)
            pdf.set_xy(x_right, y_num)
            pdf.cell(half_w, h_num_cell, _pdf_txt(pdf, fim_s, use_dejavu), align="C", ln=False)
            _pdf_font(pdf, use_dejavu, "", 5.5)
            pdf.set_text_color(90, 75, 85)
            y_lbl = y_num + h_num_cell + 0.5
            pdf.set_xy(x_left, y_lbl)
            pdf.cell(half_w, 3.2, _pdf_txt(pdf, "Inicial", use_dejavu), align="C", ln=False)
            pdf.set_xy(x_right, y_lbl)
            pdf.cell(half_w, 3.2, _pdf_txt(pdf, "Final", use_dejavu), align="C", ln=False)

            _pdf_font(pdf, use_dejavu, "", 5.8)
            pdf.set_text_color(61, 53, 64)
            yf = y0 + title_h + val_h + 1.8
            for fl in foot_lines[:2]:
                if not str(fl).strip():
                    continue
                pdf.set_xy(x + 1.5, yf)
                pdf.cell(cw - 3, 3.0, _pdf_txt(pdf, fl, use_dejavu), align="C", ln=False)
                yf += 3.05

            idx += 1
        pdf.set_xy(margin, y0 + card_h + gap_v)

    pdf.ln(1)


def _pdf_lista_rosa(pdf, titulo, pares, use_dejavu, subtitulo=None):
    """Lista rótulo → número com linhas em grelha suave (folha 1, como no PDF de referência)."""
    if not pares:
        return
    pdf.ln(2)
    _pdf_font(pdf, use_dejavu, "B", 9)
    pdf.set_text_color(173, 20, 87)
    pdf.set_x(pdf.l_margin)
    pdf.cell(0, 5, _pdf_txt(pdf, titulo, use_dejavu), ln=True)
    if subtitulo:
        _pdf_font(pdf, use_dejavu, "", 7)
        pdf.set_text_color(130, 85, 115)
        _pdf_multi_texto_largura_total(pdf, 3.5, subtitulo, use_dejavu)
    pdf.ln(1)
    full = pdf.w - pdf.l_margin - pdf.r_margin
    row_h = 6
    for lab, val in pares:
        if pdf.get_y() > 268:
            pdf.add_page()
        y = pdf.get_y()
        pdf.set_fill_color(255, 252, 254)
        pdf.set_draw_color(255, 192, 216)
        pdf.rect(pdf.l_margin, y, full, row_h, "D")
        _pdf_font(pdf, use_dejavu, "", 8)
        pdf.set_text_color(95, 55, 80)
        pdf.set_xy(pdf.l_margin + 3, y + 1.3)
        pdf.cell(full * 0.62, row_h - 2, _pdf_txt(pdf, str(lab), use_dejavu), ln=False)
        _pdf_font(pdf, use_dejavu, "B", 9.5)
        pdf.set_text_color(199, 21, 133)
        pdf.set_xy(pdf.l_margin + full * 0.65, y + 1)
        pdf.cell(full * 0.32, row_h - 2, _pdf_txt(pdf, str(val), use_dejavu), ln=False, align="R")
        pdf.set_xy(pdf.l_margin, y + row_h)
    pdf.ln(0.5)


def _pdf_extras_lote_lista_rosa(pdf, kpi, use_dejavu):
    """Linhas do relatório geral, valor, XML próprio/terceiro — lista simples."""
    pm = dict(kpi.get("pares") or [])
    val = float(kpi.get("valor") or 0)
    val_txt = f"{val:,.2f}".replace(",", " ").replace(".", ",") + " R$"
    itens = []
    if "Linhas no relatório geral" in pm:
        itens.append(("Linhas no relatório geral (expandido)", str(pm["Linhas no relatório geral"])))
    itens.append(("Valor contábil no resumo por série", val_txt))
    if "XML emissão própria (itens)" in pm:
        itens.append(("XML emissão própria", str(pm["XML emissão própria (itens)"])))
    if "XML terceiros (itens)" in pm:
        itens.append(("XML terceiros", str(pm["XML terceiros (itens)"])))
    _pdf_lista_rosa(
        pdf,
        "Mais números do lote",
        itens,
        use_dejavu,
        subtitulo="Complemento aos totalizadores acima (sem gráficos — só contagem neste garimpo).",
    )


def _pdf_contexto_lista_rosa(pdf, pares_lista, use_dejavu):
    """Referência lateral, autenticidade, gerado em — estilo suave."""
    pm = dict(pares_lista or [])
    rotulos = [
        ("Gerado em", "Gerado em"),
        ("Referência último nº guardada", "Último nº por série (lateral)"),
        ("Validação autenticidade", "Autenticidade"),
    ]
    itens = [(c, pm[k]) for k, c in rotulos if k in pm]
    if not itens:
        return
    _pdf_lista_rosa(
        pdf,
        "Contexto na app (opcional)",
        itens,
        use_dejavu,
        subtitulo="Opções que usou na barra lateral.",
    )


def _pdf_secao_resumo_folha(pdf, titulo, use_dejavu, texto_explicativo=None):
    """Título de secção + texto explicativo (folhas descritivas 2+)."""
    pdf.ln(3)
    pdf.set_draw_color(255, 105, 180)
    pdf.set_line_width(0.35)
    yl = pdf.get_y()
    pdf.line(pdf.l_margin, yl, pdf.l_margin + 28, yl)
    pdf.set_line_width(0.2)
    pdf.set_draw_color(255, 192, 216)
    pdf.line(pdf.l_margin + 30, yl, pdf.w - pdf.r_margin, yl)
    pdf.ln(2)
    _pdf_font(pdf, use_dejavu, "B", 9.5)
    pdf.set_text_color(136, 14, 79)
    pdf.set_x(pdf.l_margin)
    pdf.cell(0, 5, _pdf_txt(pdf, titulo, use_dejavu), ln=True)
    pdf.set_text_color(95, 55, 80)
    pdf.ln(0.5)
    if texto_explicativo:
        _pdf_font(pdf, use_dejavu, "", 7.5)
        pdf.set_text_color(130, 85, 115)
        _pdf_multi_texto_largura_total(pdf, 3.6, texto_explicativo, use_dejavu)
        pdf.set_text_color(60, 40, 55)
        pdf.ln(0.5)


def _pdf_tabela_preview(pdf, preview, use_dejavu, y_max=276, estilo_moderno=False):
    cols = preview.get("cols") or []
    rows = preview.get("rows") or []
    em = preview.get("empty_msg")
    if em and not cols:
        _pdf_font(pdf, use_dejavu, "", 8.5)
        pdf.set_text_color(148, 163, 184)
        _pdf_multi_texto_largura_total(pdf, 4.5, em, use_dejavu)
        pdf.set_text_color(30, 41, 59)
        return
    if not cols:
        return
    max_w = pdf.w - pdf.l_margin - pdf.r_margin
    n = len(cols)
    fs = 6.0 if n >= 8 else 7.0
    row_h = 3.8
    cw = max_w / n

    def _cabecalho():
        if estilo_moderno:
            pdf.set_fill_color(252, 210, 228)
            pdf.set_draw_color(244, 143, 177)
            pdf.set_text_color(99, 17, 58)
        else:
            pdf.set_fill_color(248, 187, 208)
            pdf.set_draw_color(236, 160, 188)
            pdf.set_text_color(55, 55, 60)
        _pdf_font(pdf, use_dejavu, "B", fs - 0.2)
        for c in cols:
            t = str(c)[:16] + ("…" if len(str(c)) > 16 else "")
            pdf.cell(cw, row_h + 0.7, _pdf_txt(pdf, t, use_dejavu), border=1, align="C", fill=True)
        pdf.ln()

    _cabecalho()
    _pdf_font(pdf, use_dejavu, "", fs)
    for ri, row in enumerate(rows):
        if pdf.get_y() > y_max:
            pdf.add_page()
            _cabecalho()
            _pdf_font(pdf, use_dejavu, "", fs)
        if estilo_moderno and ri % 2 == 0:
            pdf.set_fill_color(255, 248, 252)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(255, 205, 220)
        for j, cell in enumerate(row):
            s = str(cell)
            lim = 14 if cw < 22 else 22
            if len(s) > lim:
                s = s[: max(1, lim - 2)] + "…"
            pdf.set_text_color(75, 40, 60)
            pdf.cell(cw, row_h, _pdf_txt(pdf, s, use_dejavu), border=1, align="L", fill=True)
        pdf.ln()
    pdf.set_text_color(30, 41, 59)
    if preview.get("truncated"):
        pdf.ln(1)
        _pdf_font(pdf, use_dejavu, "", 6.5)
        pdf.set_text_color(148, 163, 184)
        tot = preview.get("total", 0)
        most = len(rows)
        msg = f"Amostra: {most}/{tot} linhas — Excel na app para tudo."
        _pdf_multi_texto_largura_total(pdf, 3.5, msg, use_dejavu)
        pdf.set_text_color(30, 41, 59)


def pdf_dashboard_garimpeiro_bytes(kpi, cnpj_fmt="", df_resumo=None):
    """
    PDF: folha 1 = cabeçalho executivo + quatro cartões KPI, faixa de buracos, cartões por série (emissão própria),
    listas terceiros / extras; folhas seguintes = indicadores detalhados em tabelas com bordas.
    """
    try:
        from fpdf import FPDF
        import fpdf as _fpdf_mod
    except ImportError:
        return None
    if not kpi:
        return None

    font_path = None
    font_bold_path = None
    try:
        _root = Path(_fpdf_mod.__file__).resolve().parent / "font"
        _p = _root / "DejaVuSans.ttf"
        if _p.is_file():
            font_path = str(_p)
        _pb = _root / "DejaVuSans-Bold.ttf"
        if _pb.is_file():
            font_bold_path = str(_pb)
    except Exception:
        font_path = None
        font_bold_path = None

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.set_margins(12, 12, 12)
    pdf.add_page()

    use_dejavu = bool(font_path)
    if use_dejavu:
        pdf.add_font("DejaVu", "", font_path)
        if font_bold_path:
            pdf.add_font("DejaVu", "B", font_bold_path)

    linha_cnpj = f"CNPJ {cnpj_fmt}" if cnpj_fmt else None
    _pdf_cabecalho_executivo_painel(pdf, use_dejavu, linha_cnpj)
    _pdf_quatro_kpi_cards_executivo(pdf, kpi, use_dejavu)
    _pdf_faixa_buracos_executivo(pdf, int(kpi.get("n_bur") or 0), use_dejavu)
    _pdf_serie_cards_emissao_propria(pdf, df_resumo, use_dejavu)

    tc = kpi.get("terc_cnt") or {}
    if tc:
        pares_tc = sorted(tc.items(), key=lambda x: x[0])
        _pdf_lista_rosa(
            pdf,
            "Terceiros · por tipo de documento",
            [(str(m), str(int(q))) for m, q in pares_tc],
            use_dejavu,
            subtitulo="Contagem de XML recebidos de terceiros (NF-e, NFC-e, CT-e, MDF-e…).",
        )

    tsd = kpi.get("terc_status_dist") or {}
    if tsd:
        pares_st = sorted(tsd.items(), key=lambda x: -x[1])
        _pdf_lista_rosa(
            pdf,
            "Terceiros · por status (como na emissão própria)",
            [(str(k), str(int(v))) for k, v in pares_st],
            use_dejavu,
            subtitulo="Situação das linhas de terceiros no relatório geral: normal, cancelada, denegada, rejeitada, etc. (sem buracos nem inutilizadas.)",
        )

    _pdf_extras_lote_lista_rosa(pdf, kpi, use_dejavu)
    _pdf_contexto_lista_rosa(pdf, kpi.get("pares", []), use_dejavu)

    pv = kpi.get("pdf_previews") or {}

    pdf.add_page()
    _pdf_font(pdf, use_dejavu, "B", 11)
    pdf.set_text_color(136, 14, 79)
    pdf.set_x(pdf.l_margin)
    pdf.cell(0, 7, _pdf_txt(pdf, "Indicadores detalhados", use_dejavu), ln=True)
    _pdf_font(pdf, use_dejavu, "", 7.5)
    pdf.set_text_color(130, 85, 115)
    _pdf_multi_texto_largura_total(
        pdf,
        3.8,
        "As secções abaixo repetem a mesma ordem e o mesmo significado que na página do Garimpeiro "
        "(resumo por série, terceiros e o relatório da leitura em **dois blocos recolhíveis** — emissão própria e, abaixo, terceiros). Cada bloco inclui uma "
        "nota curta sobre o que a tabela representa.",
        use_dejavu,
    )
    pdf.ln(1)
    pdf.set_text_color(60, 40, 55)

    def _folha_se_cheio(ymin=238):
        if pdf.get_y() > ymin:
            pdf.add_page()

    _ex = {
        "serie": (
            "Corresponde ao quadro «Resumo por série»: NF-e, NFC-e e NFS-e com emitente igual ao CNPJ da barra lateral; "
            "por modelo e série mostra o intervalo de numeração nos ficheiros, a quantidade e o valor contábil somado."
        ),
        "terc": (
            "Corresponde a «Terceiros — total por tipo»: soma de documentos recebidos de terceiros "
            "(por exemplo NF-e, NFC-e, CT-e, MDF-e) contados neste lote."
        ),
        "bur": (
            "Corresponde à aba «Buracos» (mesma base do resumo por série): faltas de numeração para NF-e, NFC-e e NFS-e "
            "da sua empresa (emitente = CNPJ da barra lateral). Se guardou «último nº por série» na lateral, só essas séries entram na lista de buracos, a partir desse mês e desse último número."
        ),
        "inut": (
            "Corresponde à aba «Inutilizadas»: notas inutilizadas na Sefaz, incluindo as que declarou "
            "manualmente sem XML (quando usou essa opção na app)."
        ),
        "canc": (
            "Corresponde à aba «Canceladas»: notas canceladas no conjunto analisado, conforme o interpretado nos XML do lote."
        ),
        "aut": (
            "Corresponde à aba «Autorizadas»: notas com situação normal/autorizada na emissão própria, neste lote."
        ),
        "deneg": (
            "Corresponde à aba «Denegadas» (emissão própria): uso denegado de numeração / situação tratada como denegação nos XML."
        ),
        "rej": (
            "Corresponde à aba «Rejeitadas» (emissão própria): rejeição Sefaz (códigos agregados como rejeitadas na app)."
        ),
        "geral": (
            "Corresponde à aba «Relatório geral» do painel emissão própria, com as colunas principais. A chave de acesso aparece "
            "abreviada neste PDF; use «Baixar Excel» na app para todas as linhas e colunas completas. O painel **Terceiros** na página tem abas para canceladas, autorizadas, denegadas, rejeitadas e relatório geral (XML recebidos)."
        ),
    }

    _pdf_secao_resumo_folha(pdf, "Resumo por série (NF-e / NFC-e / NFS-e)", use_dejavu, _ex["serie"])
    _pdf_tabela_preview(pdf, pv.get("resumo") or {}, use_dejavu, estilo_moderno=True)

    _pdf_secao_resumo_folha(pdf, "Terceiros — total por tipo", use_dejavu, _ex["terc"])
    _pdf_tabela_preview(pdf, pv.get("terceiros") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Buracos", use_dejavu, _ex["bur"])
    _pdf_tabela_preview(pdf, pv.get("buracos") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Inutilizadas", use_dejavu, _ex["inut"])
    _pdf_tabela_preview(pdf, pv.get("inutilizadas") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Canceladas", use_dejavu, _ex["canc"])
    _pdf_tabela_preview(pdf, pv.get("canceladas") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Autorizadas", use_dejavu, _ex["aut"])
    _pdf_tabela_preview(pdf, pv.get("autorizadas") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Denegadas", use_dejavu, _ex["deneg"])
    _pdf_tabela_preview(pdf, pv.get("denegadas") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio()
    _pdf_secao_resumo_folha(pdf, "Rejeitadas", use_dejavu, _ex["rej"])
    _pdf_tabela_preview(pdf, pv.get("rejeitadas") or {}, use_dejavu, estilo_moderno=True)

    _folha_se_cheio(220)
    _pdf_secao_resumo_folha(pdf, "Relatório geral (colunas principais)", use_dejavu, _ex["geral"])
    _pdf_tabela_preview(pdf, pv.get("geral") or {}, use_dejavu, estilo_moderno=True)

    pdf.ln(4)
    _pdf_font(pdf, use_dejavu, "", 6.5)
    pdf.set_text_color(148, 163, 184)
    _pdf_multi_texto_largura_total(
        pdf,
        3.5,
        "Garimpeiro · Chaves abreviadas no PDF. Lista completa: Excel na app.",
        use_dejavu,
    )

    _pdf_out = io.BytesIO()
    pdf.output(_pdf_out)
    raw = _pdf_out.getvalue()
    return raw if raw else None


def aplicar_compactacao_dfs_sessao():
    """Compacta DataFrames grandes na sessão (útil no Streamlit Cloud)."""
    for k in (
        "df_geral",
        "df_resumo",
        "df_faltantes",
        "df_canceladas",
        "df_inutilizadas",
        "df_autorizadas",
        "df_denegadas",
        "df_rejeitadas",
        "df_divergencias",
    ):
        v = st.session_state.get(k)
        if v is not None and isinstance(v, pd.DataFrame) and not v.empty:
            st.session_state[k] = compactar_dataframe_memoria(v)
    gc.collect()


def _ym_tuple(ano, mes):
    try:
        return (int(ano), int(mes))
    except (ValueError, TypeError):
        return None


def _ym_gt(ano_a, mes_a, ano_b, mes_b):
    ta, tb = _ym_tuple(ano_a, mes_a), _ym_tuple(ano_b, mes_b)
    if not ta or not tb:
        return False
    return ta > tb


def _ym_eq(ano_a, mes_a, ano_b, mes_b):
    ta, tb = _ym_tuple(ano_a, mes_a), _ym_tuple(ano_b, mes_b)
    if not ta or not tb:
        return False
    return ta == tb


def _ym_lt(ano_a, mes_a, ano_b, mes_b):
    ta, tb = _ym_tuple(ano_a, mes_a), _ym_tuple(ano_b, mes_b)
    if not ta or not tb:
        return False
    return ta < tb


def buraco_ctx_sessao():
    """
    Só activa regras especiais de buracos quando existe **pelo menos um** último nº guardado
    (Guardar referência com linhas válidas). Nesse caso **só** as séries que estão na grelha da lateral
    entram na detecção de buracos (a partir do último nº e do mês indicados). Sem referência guardada,
    buracos usam toda a numeração lida em emissão própria — como antes.
    """
    try:
        rmap = st.session_state.get("seq_ref_ultimos")
        rm = dict(rmap) if isinstance(rmap, dict) and rmap else {}
        if not rm:
            return None, None, {}
        ar = st.session_state.get("seq_ref_ano")
        mr = st.session_state.get("seq_ref_mes")
        if ar is None or mr is None:
            return None, None, {}
        return int(ar), int(mr), rm
    except Exception:
        return None, None, {}


def incluir_numero_no_conjunto_buraco(ano, mes, n, ref_ar, ref_mr, ultimo_u):
    """
    Sem referência guardada na lateral (ref_ar/ref_mr ausentes): toda a numeração lida pode gerar buraco.
    Com referência activa: **só** séries com linha na grelha (último nº definido); as restantes ignoram-se
    na análise de buracos. Numa série referenciada: mês anterior ao de referência ignorado; no mês de
    referência só entram notas com nº **maior** que o último informado; meses seguintes seguem a sequência.
    """
    if ref_ar is None or ref_mr is None:
        return True
    if ultimo_u is None:
        return False
    return numero_entra_conjunto_buraco(ano, mes, n, ref_ar, ref_mr, ultimo_u)


def ultimo_ref_lookup(ref_map, tipo, serie):
    """
    Último nº guardado na lateral para modelo|série.
    Aceita o mesmo par com série «4» ou «004» (chave na grelha vs chave na NF-e) para não excluir notas de nums_buraco por engano.
    """
    if not ref_map:
        return None
    t = str(tipo).strip()
    s_raw = str(serie).strip()
    candidates = [f"{t}|{s_raw}"]
    if s_raw.isdigit():
        n = int(s_raw)
        candidates.append(f"{t}|{n}")
        candidates.append(f"{t}|{str(n).zfill(3)}")
    for k in dict.fromkeys(candidates):
        if k in ref_map:
            return ref_map[k]
    return None


def numero_entra_conjunto_buraco(ano, mes, n, ref_ar, ref_mr, ultimo_u):
    """
    Com mês de referência na sessão: ignora competências anteriores; no próprio mês só conta n > último informado;
    meses posteriores contam. Sem ref_ar/ref_mr na chamada: o chamador trata (incluir_numero… devolve True antes).
    Se ref está activo mas falta último nº da série, não contar.
    """
    if ref_ar is None or ref_mr is None:
        return True
    if str(ano) == "0000":
        return False
    if _ym_lt(ano, mes, ref_ar, ref_mr):
        return False
    if ultimo_u is None:
        return False
    if _ym_eq(ano, mes, ref_ar, ref_mr):
        try:
            return int(n) > int(ultimo_u)
        except (TypeError, ValueError):
            return False
    return True


def falhas_buraco_por_serie(nums_buraco, tipo_doc, serie_str, ultimo_u, gap_max=MAX_SALTO_ENTRE_NOTAS_CONSECUTIVAS):
    """
    Buracos a partir do último nº informado (se houver): preenche o intervalo até ao primeiro nº relevante nos XMLs
    e mantém a lógica de trechos (saltos grandes) no restante.
    """
    ns = sorted(nums_buraco)
    if not ns:
        return []
    U = None
    if ultimo_u is not None:
        try:
            U = int(ultimo_u)
        except (TypeError, ValueError):
            U = None
    out = []
    if U is not None:
        ns_eff = [x for x in ns if x > U]
        if not ns_eff:
            return []
        for b in range(U + 1, ns_eff[0]):
            out.append({"Tipo": tipo_doc, "Série": serie_str, "Num_Faltante": b})
        out.extend(enumerar_buracos_por_segmento(ns_eff, tipo_doc, serie_str, gap_max))
    else:
        out.extend(enumerar_buracos_por_segmento(ns, tipo_doc, serie_str, gap_max))
    return out


def ultimos_dict_para_dataframe(ultimos_dict):
    if not ultimos_dict:
        return pd.DataFrame(columns=["Modelo", "Série", "Último número"])
    rows = []
    for kstr, v in ultimos_dict.items():
        if "|" not in kstr:
            continue
        a, b = kstr.split("|", 1)
        rows.append({"Modelo": a.strip(), "Série": b.strip(), "Último número": int(v)})
    return pd.DataFrame(rows)


def ref_map_from_dataframe(df):
    """Monta o mapa 'Modelo|Série' -> último a partir da tabela do editor."""
    out = {}
    if df is None or df.empty:
        return out
    for _, row in df.iterrows():
        modelo = row.get("Modelo")
        if modelo is None or pd.isna(modelo):
            continue
        modelo = str(modelo).strip()
        if not modelo or modelo.lower() == "nan":
            continue
        serie = row.get("Série")
        if serie is None or pd.isna(serie):
            serie = ""
        else:
            serie = str(serie).strip()
        if not serie:
            continue
        ult = row.get("Último número")
        if ult is None or pd.isna(ult):
            continue
        if isinstance(ult, str):
            d = "".join(filter(str.isdigit, ult.strip()))
            if not d:
                continue
            try:
                u = int(d)
            except ValueError:
                continue
        else:
            try:
                u = int(float(ult))
            except (TypeError, ValueError):
                continue
        if u <= 0:
            continue
        out[f"{modelo}|{serie}"] = u
    return out


def normalize_seq_ref_editor_df(df):
    """Prepara a grelha: último nº em texto (evita float/NaN do NumberColumn que some ao recarregar)."""
    cols = ["Modelo", "Série", "Último número"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    def _str_model(x):
        if x is None or pd.isna(x):
            return None
        s = str(x).strip()
        return s if s and s.lower() != "nan" else None

    def _str_ser(x):
        if x is None or pd.isna(x):
            return ""
        return str(x).strip()

    def _ult_txt(x):
        if x is None or pd.isna(x):
            return ""
        if isinstance(x, bool):
            return ""
        if isinstance(x, int):
            return str(x) if x >= 0 else ""
        if isinstance(x, float):
            try:
                if pd.isna(x):
                    return ""
                return str(int(x))
            except (ValueError, OverflowError):
                return ""
        return "".join(filter(str.isdigit, str(x)))

    out = df.reindex(columns=cols).copy()
    out["Modelo"] = out["Modelo"].map(_str_model)
    out["Série"] = out["Série"].map(_str_ser)
    out["Último número"] = out["Último número"].map(_ult_txt)
    return out


def collect_seq_ref_from_widgets(struct_v: int, n_rows: int, default_modelo: str = "NF-e") -> pd.DataFrame:
    """Lê select/text da sidebar (chaves sr_{v}_{i}_*) e devolve DataFrame normalizado."""
    recs = []
    for i in range(n_rows):
        recs.append(
            {
                "Modelo": st.session_state.get(f"sr_{struct_v}_{i}_m", default_modelo),
                "Série": str(st.session_state.get(f"sr_{struct_v}_{i}_s", "") or ""),
                "Último número": str(st.session_state.get(f"sr_{struct_v}_{i}_u", "") or ""),
            }
        )
    return normalize_seq_ref_editor_df(pd.DataFrame(recs))


def item_registro_manual_inutilizado(cnpj_limpo, tipo_man, serie_man, nota_man):
    serie_str = str(serie_man).strip()
    return {
        "Arquivo": "REGISTRO_MANUAL",
        "Chave": f"MANUAL_INUT_{tipo_man}_{serie_str}_{nota_man}",
        "Tipo": tipo_man,
        "Série": serie_str,
        "Número": int(nota_man),
        "Status": "INUTILIZADOS",
        "Pasta": f"EMITIDOS_CLIENTE/SAIDA/{tipo_man}/INUTILIZADOS/0000/01/Serie_{serie_str}",
        "Valor": 0.0,
        "Conteúdo": b"",
        "Ano": "0000",
        "Mes": "01",
        "Operacao": "SAIDA",
        "Data_Emissao": "",
        "CNPJ_Emit": cnpj_limpo,
        "Nome_Emit": "INSERÇÃO MANUAL",
        "Doc_Dest": "",
        "Nome_Dest": "",
    }


def item_registro_manual_cancelado(cnpj_limpo, tipo_man, serie_man, nota_man):
    """Cancelamento declarado sem XML de evento (mesmas regras de buraco que inutilizada manual)."""
    serie_str = str(serie_man).strip()
    return {
        "Arquivo": "REGISTRO_MANUAL_CANCELADO",
        "Chave": f"MANUAL_CANC_{tipo_man}_{serie_str}_{nota_man}",
        "Tipo": tipo_man,
        "Série": serie_str,
        "Número": int(nota_man),
        "Status": "CANCELADOS",
        "Pasta": f"EMITIDOS_CLIENTE/SAIDA/{tipo_man}/CANCELADOS/0000/01/Serie_{serie_str}",
        "Valor": 0.0,
        "Conteúdo": b"",
        "Ano": "0000",
        "Mes": "01",
        "Operacao": "SAIDA",
        "Data_Emissao": "",
        "CNPJ_Emit": cnpj_limpo,
        "Nome_Emit": "INSERÇÃO MANUAL",
        "Doc_Dest": "",
        "Nome_Dest": "",
    }


def _inutil_sem_xml_manual(res):
    """Inutilização declarada manualmente (sem XML de inutilização)."""
    if str(res.get("Chave") or "").startswith("MANUAL_INUT_"):
        return True
    return res.get("Arquivo") == "REGISTRO_MANUAL" and res.get("Status") == "INUTILIZADOS"


def _cancel_sem_xml_manual(res):
    """Cancelamento declarado manualmente (sem XML de cancelamento)."""
    if str(res.get("Chave") or "").startswith("MANUAL_CANC_"):
        return True
    return res.get("Arquivo") == "REGISTRO_MANUAL_CANCELADO" and res.get("Status") == "CANCELADOS"


def _registro_manual_inutil_duplicada(relatorio, tipo, serie_str, nota_int):
    """Já existe **inutilização** manual (sem XML) para o mesmo modelo/série/nota — não misturar com canceladas."""
    ser = str(serie_str).strip()
    for r in relatorio or []:
        if not _inutil_sem_xml_manual(r):
            continue
        if str(r.get("Tipo") or "").strip() != str(tipo).strip():
            continue
        if str(r.get("Série") or "").strip() != ser:
            continue
        try:
            if int(r.get("Número") or 0) != int(nota_int):
                continue
        except (TypeError, ValueError):
            continue
        return True
    return False


def _registro_manual_cancel_duplicada(relatorio, tipo, serie_str, nota_int):
    """Já existe **cancelamento** manual (sem XML) para o mesmo modelo/série/nota — independente das inutilizadas."""
    ser = str(serie_str).strip()
    for r in relatorio or []:
        if not _cancel_sem_xml_manual(r):
            continue
        if str(r.get("Tipo") or "").strip() != str(tipo).strip():
            continue
        if str(r.get("Série") or "").strip() != ser:
            continue
        try:
            if int(r.get("Número") or 0) != int(nota_int):
                continue
        except (TypeError, ValueError):
            continue
        return True
    return False


def _norm_cab_inutil_col(c):
    s = unicodedata.normalize("NFD", str(c).strip().lower())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.replace(" ", "_")


def triplas_inutil_de_dataframe(df):
    """
    Cabeçalhos flexíveis → lista (modelo, série, nota).
    Modelo: modelo, documento, tipo, doc… | Série: série, serie, ser… | Nota: nota, número, num, num_faltante…
    """
    if df is None or df.empty:
        return None, "A planilha está vazia."
    df = df.dropna(how="all")
    if df.empty:
        return None, "A planilha está vazia."
    ren = {c: _norm_cab_inutil_col(c) for c in df.columns}
    d2 = df.rename(columns=ren)

    def col(*aliases):
        for a in aliases:
            if a in d2.columns:
                return a
        return None

    cm = col("modelo", "documento", "tipo", "doc", "document", "mod", "cod_mod", "codigo")
    cs = col("serie", "ser")
    cn = col("nota", "numero", "num", "num_faltante", "n", "numeracao", "no")
    if not cm or not cs or not cn:
        return (
            None,
            "Faltam colunas reconhecíveis. Use **Modelo** (código Sefaz **55**, **65**, **57**, **67**, **58** ou NF-e…), "
            "**Série** e **Nota** (ou Número / Num_Faltante).",
        )
    out = []
    for _, row in d2.iterrows():
        m = row.get(cm)
        s = row.get(cs)
        nraw = row.get(cn)
        if (m is None or pd.isna(m)) and (s is None or pd.isna(s)) and (nraw is None or pd.isna(nraw)):
            continue
        if m is None or pd.isna(m) or s is None or pd.isna(s) or nraw is None or pd.isna(nraw):
            continue
        mod = _normaliza_modelo_filtro(m)
        ser = _normaliza_serie_filtro(s)
        if not mod or not ser:
            continue
        if isinstance(nraw, (int, float)) and not pd.isna(nraw):
            try:
                n = int(float(nraw))
            except (TypeError, ValueError):
                continue
        else:
            d = "".join(filter(str.isdigit, str(nraw)))
            if not d:
                continue
            try:
                n = int(d)
            except ValueError:
                continue
        if n <= 0:
            continue
        out.append((mod, ser, n))
    if not out:
        return None, "Nenhuma linha válida (modelo, série e nota preenchidos)."
    return out, None


def dataframe_de_upload_inutil(uploaded_file, max_linhas=50000):
    """Lê CSV ou Excel enviado pelo utilizador."""
    if uploaded_file is None:
        return None, None
    nome = (getattr(uploaded_file, "name", None) or "").lower()
    raw = uploaded_file.read()
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    if nome.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=enc)
                break
            except Exception:
                df = None
        if df is None:
            return None, "Não foi possível ler o CSV (tente UTF-8 ou separador `;` / `,`)."
    elif nome.endswith((".xlsx", ".xls")):
        try:
            df = pd.read_excel(io.BytesIO(raw))
        except Exception as e:
            return None, f"Erro ao ler Excel: {e}"
    else:
        return None, "Use ficheiro **.csv**, **.xlsx** ou **.xls**."
    if len(df) > max_linhas:
        return None, f"No máximo {max_linhas} linhas por ficheiro."
    return df, None


def _computar_df_divergencias_autenticidade(df_geral: pd.DataFrame, triplas_sefaz: list):
    """
    Compara (Modelo, Série, Nota) do relatório Sefaz com NF **autorizadas** (Status NORMAIS)
    de **emissão própria** no relatório geral do XML.
    """
    if df_geral is None or df_geral.empty or not triplas_sefaz:
        return pd.DataFrame()
    need = {"Modelo", "Série", "Nota", "Status Final", "Origem", "Chave"}
    if not need.issubset(set(df_geral.columns)):
        return pd.DataFrame()
    mp = _mask_emissao_propria_df(df_geral)
    st_col = df_geral["Status Final"].astype(str).str.upper()
    m_aut = mp & st_col.eq("NORMAIS")
    dfp = df_geral.loc[m_aut]
    xml_set = set()
    chave_por_tripla = {}
    for _, row in dfp.iterrows():
        mod = _normaliza_modelo_filtro(row.get("Modelo"))
        ser = _normaliza_serie_filtro(row.get("Série"))
        ni = _nota_int_linha(row)
        if not mod or ser == "" or ni is None:
            continue
        t = (mod, ser, ni)
        xml_set.add(t)
        ck = str(row.get("Chave") or "").strip()
        if ck and t not in chave_por_tripla:
            chave_por_tripla[t] = ck
    sefaz_set = set()
    for mod, ser, n in triplas_sefaz:
        try:
            mn = _normaliza_modelo_filtro(mod)
            sn = _normaliza_serie_filtro(ser)
            if not mn or sn == "":
                continue
            sefaz_set.add((mn, sn, int(n)))
        except (TypeError, ValueError):
            continue
    only_sefaz = sefaz_set - xml_set
    only_xml = xml_set - sefaz_set
    out_rows = []
    for t in sorted(only_sefaz, key=lambda x: (x[0], x[1], x[2])):
        out_rows.append(
            {
                "Tipo divergência": "Na Sefaz / relatório, ausente no XML do lote",
                "Modelo": t[0],
                "Série": t[1],
                "Nota": t[2],
                "Chave (XML)": "",
            }
        )
    for t in sorted(only_xml, key=lambda x: (x[0], x[1], x[2])):
        out_rows.append(
            {
                "Tipo divergência": "No XML do lote, ausente na Sefaz / relatório",
                "Modelo": t[0],
                "Série": t[1],
                "Nota": t[2],
                "Chave (XML)": chave_por_tripla.get(t, ""),
            }
        )
    return pd.DataFrame(out_rows)


def _tem_upload_autenticidade(up) -> bool:
    """True se há pelo menos um ficheiro no uploader (lista vazia ou None → False)."""
    if up is None:
        return False
    if isinstance(up, (list, tuple)):
        return len(up) > 0
    return True


def _aplicar_comparacao_autenticidade(upload_files) -> tuple:
    """Lê uma ou várias planilhas Sefaz, agrega linhas e cruza com df_geral. Devolve (ok, mensagem)."""
    st.session_state.pop("df_divergencias", None)
    st.session_state["validation_done"] = False
    if upload_files is None:
        return True, ""
    files = list(upload_files) if isinstance(upload_files, (list, tuple)) else [upload_files]
    files = [f for f in files if f is not None]
    if not files:
        return True, ""
    dfs_valid = []
    for f in files:
        nome = getattr(f, "name", "") or "sem_nome"
        df, err = dataframe_de_upload_inutil(f)
        if err:
            return False, f"**Autenticidade** (`{nome}`): {err}"
        if df is not None and not df.empty:
            dfs_valid.append(df)
    if not dfs_valid:
        return False, "**Autenticidade:** nenhum ficheiro tinha linhas de dados (planilhas vazias?)."
    merged = pd.concat(dfs_valid, ignore_index=True) if len(dfs_valid) > 1 else dfs_valid[0]
    triplas, err2 = triplas_inutil_de_dataframe(merged)
    if err2:
        return False, f"**Autenticidade** ({len(files)} ficheiro(s) combinados): {err2}"
    df_g = st.session_state.get("df_geral")
    if df_g is None or (isinstance(df_g, pd.DataFrame) and df_g.empty):
        return False, "**Autenticidade:** relatório geral vazio — faça o garimpo primeiro."
    df_div = _computar_df_divergencias_autenticidade(df_g, triplas)
    if df_div is not None and not df_div.empty:
        st.session_state["df_divergencias"] = df_div
    else:
        st.session_state.pop("df_divergencias", None)
    st.session_state["validation_done"] = True
    n_lin = len(triplas)
    n_div = len(df_div) if df_div is not None else 0
    n_f = len(files)
    suf = f"{n_f} ficheiro(s), " if n_f > 1 else ""
    return (
        True,
        f"**Autenticidade:** {suf}**{n_lin}** linha(s) na planilha Sefaz (agregadas); **{n_div}** divergência(s) "
        f"(autorizadas próprias NORMAIS no XML vs relatório).",
    )


def conjunto_triplas_buracos(df_faltantes):
    """{(Tipo, série_str, Num_Faltante)} a partir da tabela de buracos do garimpeiro."""
    if df_faltantes is None or df_faltantes.empty:
        return set()
    d = df_faltantes.copy()
    if "Serie" in d.columns and "Série" not in d.columns:
        d = d.rename(columns={"Serie": "Série"})
    if not {"Tipo", "Série", "Num_Faltante"}.issubset(d.columns):
        return set()
    out = set()
    for _, row in d.iterrows():
        try:
            out.add(
                (
                    str(row["Tipo"]).strip(),
                    str(row["Série"]).strip(),
                    int(row["Num_Faltante"]),
                )
            )
        except (TypeError, ValueError):
            continue
    return out


def _dataframe_modelo_planilha_inutil_sem_xml():
    """Linhas de exemplo com código numérico Sefaz (como na página da Sefaz) — também aceita NF-e, NFC-e…"""
    return pd.DataFrame(
        [
            {"Modelo": 55, "Série": 1, "Nota": 1520},
            {"Modelo": 55, "Série": 1, "Nota": 1521},
            {"Modelo": 65, "Série": 2, "Nota": 100},
            {"Modelo": 57, "Série": 1, "Nota": 500},
            {"Modelo": 67, "Série": 1, "Nota": 10},
        ]
    )


def _bytes_modelo_planilha_exemplo_xlsx(sheet_name: str):
    """
    Gera bytes do .xlsx modelo (exemplo Modelo/Série/Nota).
    Se xlsxwriter falhar por disco/temp cheio, tenta openpyxl; em falha total devolve None.
    """
    df = _dataframe_modelo_planilha_inutil_sem_xml()
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.set_column(0, 0, 14)
            ws.set_column(1, 1, 10)
            ws.set_column(2, 2, 12)
        return buf.getvalue()
    except Exception as exc:
        if not _erro_sem_espaco_disco(exc):
            raise
    try:
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return buf2.getvalue()
    except Exception:
        return None


def bytes_modelo_planilha_inutil_sem_xml_xlsx():
    """Bytes do modelo ou None se não houver espaço para gerar o .xlsx."""
    return _bytes_modelo_planilha_exemplo_xlsx("Inutil_sem_XML")


def bytes_modelo_planilha_cancel_sem_xml_xlsx():
    """Mesmo layout que inutilizadas (Modelo, Série, Nota) — só canceladas declaradas manualmente."""
    return _bytes_modelo_planilha_exemplo_xlsx("Cancel_sem_XML")


def _item_inutil_manual_sem_xml(res):
    """Inutilização «sem XML» inserida pelo utilizador (não vem de ficheiro)."""
    return (
        res.get("Status") == "INUTILIZADOS"
        and _inutil_sem_xml_manual(res)
        and "EMITIDOS_CLIENTE" in res.get("Pasta", "")
    )


def _lote_recalc_de_relatorio(relatorio_list):
    """Mesma deduplicação por Chave que reconstruir_dataframes_relatorio_simples."""
    lote = {}
    for item in relatorio_list:
        key = item["Chave"]
        is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
        if key in lote:
            if item["Status"] in [
                "CANCELADOS",
                "INUTILIZADOS",
                "DENEGADOS",
                "REJEITADOS",
            ]:
                lote[key] = (item, is_p)
        else:
            lote[key] = (item, is_p)
    return lote


def _conjunto_buracos_sem_inutil_manual(lote_recalc, ref_ar, ref_mr, ref_map):
    """
    Buracos atuais ignorando inutilizações manuais «sem XML».
    Tuplas (Tipo, série_str, número) para cruzar com o que o utilizador declara.
    Mantém todos os modelos em emissão própria (H) para não apagar inutil. manual de NFC-e/CT-e por engano.
    """
    audit_map = {}
    for k, (res, is_p) in lote_recalc.items():
        if not is_p:
            continue
        if _item_inutil_manual_sem_xml(res):
            continue
        sk = (res["Tipo"], res["Série"])
        if sk not in audit_map:
            audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
        ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
        if res["Status"] == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                audit_map[sk]["nums"].add(n)
                if incluir_numero_no_conjunto_buraco(
                    res["Ano"], res["Mes"], n, ref_ar, ref_mr, ult_u
                ):
                    audit_map[sk]["nums_buraco"].add(n)
        else:
            if res["Número"] > 0:
                audit_map[sk]["nums"].add(res["Número"])
                if incluir_numero_no_conjunto_buraco(
                    res["Ano"],
                    res["Mes"],
                    res["Número"],
                    ref_ar,
                    ref_mr,
                    ult_u,
                ):
                    audit_map[sk]["nums_buraco"].add(res["Número"])
                audit_map[sk]["valor"] += res["Valor"]

    H = set()
    for (t, s), dados in audit_map.items():
        ult_lookup = ultimo_ref_lookup(ref_map, t, s) if ref_ar is not None else None
        for row in falhas_buraco_por_serie(dados["nums_buraco"], t, s, ult_lookup):
            H.add((row["Tipo"], str(row["Série"]).strip(), int(row["Num_Faltante"])))
    return H


def reconstruir_dataframes_relatorio_simples():
    """Recalcula tabelas a partir de st.session_state['relatorio'] (status no próprio item)."""
    rel_list = list(st.session_state["relatorio"])
    ref_ar, ref_mr, ref_map = buraco_ctx_sessao()
    _cnpj_cli = "".join(c for c in str(st.session_state.get("cnpj_widget", "")) if c.isdigit())[:14]

    lote_full = _lote_recalc_de_relatorio(rel_list)
    lote_sem_manual = {
        k: v
        for k, v in lote_full.items()
        if not _item_inutil_manual_sem_xml(v[0])
    }
    H = _conjunto_buracos_sem_inutil_manual(lote_sem_manual, ref_ar, ref_mr, ref_map)

    drop_ch = set()
    for k, (res, is_p) in lote_full.items():
        if not _item_inutil_manual_sem_xml(res):
            continue
        r = res.get("Range", (res["Número"], res["Número"]))
        ra, rb = int(r[0]), int(r[1])
        ser_s = str(res["Série"]).strip()
        if not any((res["Tipo"], ser_s, n) in H for n in range(ra, rb + 1)):
            drop_ch.add(k)
    if drop_ch:
        st.session_state["relatorio"] = [x for x in rel_list if x["Chave"] not in drop_ch]
        rel_list = list(st.session_state["relatorio"])
        lote_full = _lote_recalc_de_relatorio(rel_list)

    audit_map = {}
    canc_list = []
    inut_list = []
    aut_list = []
    deneg_list = []
    rej_list = []
    geral_list = []

    for k, (res, is_p) in lote_full.items():
        if is_p:
            origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})"
        else:
            origem_label = f"TERCEIROS ({res['Operacao']})"

        registro_detalhado = {
            "Origem": origem_label,
            "Operação": res["Operacao"],
            "Modelo": res["Tipo"],
            "Série": res["Série"],
            "Nota": res["Número"],
            "Data Emissão": res["Data_Emissao"],
            "CNPJ Emitente": res["CNPJ_Emit"],
            "Nome Emitente": res["Nome_Emit"],
            "Doc Destinatário": res["Doc_Dest"],
            "Nome Destinatário": res["Nome_Dest"],
            "UF Destino": res.get("UF_Dest") or "",
            "Chave": res["Chave"],
            "Status Final": res["Status"],
            "Valor": res["Valor"],
            "Ano": res["Ano"],
            "Mes": res["Mes"],
        }

        if res["Status"] == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            ra, rb = int(r[0]), int(r[1])
            _man_inut = _inutil_sem_xml_manual(res)
            for n in range(ra, rb + 1):
                if _man_inut:
                    if (res["Tipo"], str(res["Série"]).strip(), n) not in H:
                        continue
                item_inut = registro_detalhado.copy()
                item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                geral_list.append(item_inut)
                if is_p:
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                if _incluir_em_resumo_por_serie(res, is_p, _cnpj_cli):
                    sk = (res["Tipo"], res["Série"])
                    if sk not in audit_map:
                        audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                    ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
                    audit_map[sk]["nums"].add(n)
                    if _man_inut:
                        audit_map[sk]["nums_buraco"].add(n)
                    else:
                        if incluir_numero_no_conjunto_buraco(
                            res["Ano"], res["Mes"], n, ref_ar, ref_mr, ult_u
                        ):
                            audit_map[sk]["nums_buraco"].add(n)
        else:
            geral_list.append(registro_detalhado)
            if is_p:
                if res["Status"] == "DENEGADOS":
                    deneg_list.append(registro_detalhado)
                elif res["Status"] == "REJEITADOS":
                    rej_list.append(registro_detalhado)
                elif res["Número"] > 0:
                    if res["Status"] == "CANCELADOS":
                        canc_list.append(registro_detalhado)
                    elif res["Status"] == "NORMAIS":
                        aut_list.append(registro_detalhado)
            if _incluir_em_resumo_por_serie(res, is_p, _cnpj_cli) and res["Número"] > 0:
                sk = (res["Tipo"], res["Série"])
                if sk not in audit_map:
                    audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
                audit_map[sk]["nums"].add(res["Número"])
                if _cancel_sem_xml_manual(res):
                    audit_map[sk]["nums_buraco"].add(res["Número"])
                elif incluir_numero_no_conjunto_buraco(
                    res["Ano"],
                    res["Mes"],
                    res["Número"],
                    ref_ar,
                    ref_mr,
                    ult_u,
                ):
                    audit_map[sk]["nums_buraco"].add(res["Número"])
                audit_map[sk]["valor"] += res["Valor"]

    res_final = []
    fal_final = []

    for (t, s), dados in audit_map.items():
        ns = sorted(list(dados["nums"]))
        if ns:
            n_min = ns[0]
            n_max = ns[-1]
            res_final.append(
                {
                    "Documento": t,
                    "Série": s,
                    "Início": n_min,
                    "Fim": n_max,
                    "Quantidade": len(ns),
                    "Valor Contábil (R$)": round(dados["valor"], 2),
                }
            )
        ult_lookup = ultimo_ref_lookup(ref_map, t, s) if ref_ar is not None else None
        fal_final.extend(
            falhas_buraco_por_serie(dados["nums_buraco"], t, s, ult_lookup)
        )

    st.session_state.update(
        {
            "df_resumo": pd.DataFrame(res_final),
            "df_faltantes": pd.DataFrame(fal_final),
            "df_canceladas": pd.DataFrame(canc_list),
            "df_inutilizadas": pd.DataFrame(inut_list),
            "df_autorizadas": pd.DataFrame(aut_list),
            "df_denegadas": pd.DataFrame(deneg_list),
            "df_rejeitadas": pd.DataFrame(rej_list),
            "df_geral": pd.DataFrame(geral_list),
            "st_counts": {
                "CANCELADOS": len(canc_list),
                "INUTILIZADOS": len(inut_list),
                "AUTORIZADAS": len(aut_list),
                "DENEGADOS": len(deneg_list),
                "REJEITADOS": len(rej_list),
            },
        }
    )
    try:
        st.session_state.pop("_v2_cascade_cache_v1", None)
    except Exception:
        pass
    aplicar_compactacao_dfs_sessao()


def reprocessar_garimpeiro_a_partir_do_disco(cnpj_limpo: str):
    """
    Relê todos os XML/ZIP do lote (memória da sessão ou TEMP_UPLOADS_DIR), mesmas regras de fusão por chave,
    mantém registos manuais de inutilização e de cancelamento «sem XML» e recalcula os dataframes.
    """
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    if len(cnpj) != 14:
        return False, "CNPJ inválido — confira a barra lateral."

    if not _garimpo_existem_fontes_xml_lote():
        if _garimpo_analise_sem_pasta_local_projeto():
            return (
                False,
                "Não há ficheiros do lote em memória. Faça **Iniciar grande garimpo** ou **Incluir mais XML**.",
            )
        return False, "Pasta de uploads não existe."

    nomes = _lista_nomes_fontes_xml_garimpo()
    if not nomes:
        return (
            False,
            "Nenhum ficheiro no lote. Use «Incluir mais XML / ZIP» ou inicie um novo garimpo.",
        )

    lote_dict = {}
    for f_name in nomes:
        try:
            with _abrir_fonte_xml_garimpo_stream(f_name) as file_obj:
                todos_xmls = extrair_recursivo(file_obj, f_name)
                for name, xml_data in todos_xmls:
                    res, is_p = identify_xml_info(xml_data, cnpj, name)
                    if res:
                        key = res["Chave"]
                        if key in lote_dict:
                            if res["Status"] in [
                                "CANCELADOS",
                                "INUTILIZADOS",
                                "DENEGADOS",
                                "REJEITADOS",
                            ]:
                                lote_dict[key] = (res, is_p)
                        else:
                            lote_dict[key] = (res, is_p)
                    del xml_data
        except Exception:
            continue

    rel_disk = [t[0] for t in lote_dict.values()]
    if not rel_disk:
        return (
            False,
            "Nenhum documento reconhecido ao reler a pasta (verifique CNPJ e ficheiros). O relatório não foi alterado.",
        )

    rel_atual = list(st.session_state.get("relatorio") or [])
    manuais = [r for r in rel_atual if _inutil_sem_xml_manual(r) or _cancel_sem_xml_manual(r)]

    st.session_state["relatorio"] = rel_disk + manuais
    st.session_state["export_ready"] = False
    st.session_state["excel_buffer"] = None
    if st.session_state.get("validation_done"):
        st.session_state["validation_done"] = False
    st.session_state.pop("df_divergencias", None)
    st.session_state.pop("_xlsx_mem_geral_workbook", None)
    st.session_state.pop("_xlsx_mem_geral_workbook_terc", None)
    for _pfx in (
        "rep_bur",
        "rep_inu",
        "rep_canc",
        "rep_aut",
        "rep_den",
        "rep_rej",
        "rep_ger",
        "rep_canc_t",
        "rep_aut_t",
        "rep_den_t",
        "rep_rej_t",
        "rep_ger_t",
    ):
        st.session_state.pop(f"{_pfx}_zip_parts_ready", None)
        st.session_state.pop(f"{_pfx}_zip_src_sig", None)

    reconstruir_dataframes_relatorio_simples()

    _n_inm = sum(1 for r in manuais if _inutil_sem_xml_manual(r))
    _n_cam = sum(1 for r in manuais if _cancel_sem_xml_manual(r))
    return (
        True,
        f"Concluído: {len(nomes)} ficheiro(s) lidos, {len(rel_disk)} documento(s) únicos; "
        f"{len(manuais)} registo(s) manual(is) mantido(s) ({_n_inm} inutil., {_n_cam} cancel.).",
    )


def _relatorio_ja_tem_chave(chave: str) -> bool:
    return any(r.get("Chave") == chave for r in (st.session_state.get("relatorio") or []))


def processar_painel_lateral_direito(
    cnpj_limpo: str,
    extra_files,
    pick_bur_inut,
    mb_bur,
    sb_bur,
    up_inut_planilha,
    mf_faixa,
    sf_faixa,
    n0_faixa,
    n1_faixa,
    pick_bur_canc=None,
    mb_canc=None,
    sb_canc=None,
    up_canc_planilha=None,
    mf_canc_f=None,
    sf_canc_f=None,
    n0_canc_f=None,
    n1_canc_f=None,
    up_autent_sefaz=None,
):
    """
    Um único passo: grava XML/ZIP extra em disco, aplica inutilizações e canceladas manuais
    (buracos / planilha / faixa, mesmas regras), recalcula a partir do disco (ou só reconstrói)
    e, se anexar planilha(s), compara autenticidade Sefaz × XML (vários Excel/CSV são agregados).
    """
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    linhas = []
    n_extra = 0
    if extra_files:
        n_extra = _garimpo_absorver_uploads_extra_no_lote(extra_files)
        if n_extra:
            _lbl = (
                "**em memória** (e/ou pasta temp, conforme configuração)"
                if _garimpo_analise_sem_pasta_local_projeto()
                else "**na pasta de uploads**"
            )
            linhas.append(f"**{n_extra}** ficheiro(s) extra incorporado(s) ao lote {_lbl}.")

    n_b = 0
    if pick_bur_inut:
        if mb_bur is None or sb_bur is None or str(sb_bur).strip() == "":
            return (
                False,
                "Para aplicar notas em «Dos buracos», escolha **modelo** e **série** nesse separador.",
                linhas,
            )
        for _nb in pick_bur_inut:
            _it = item_registro_manual_inutilizado(cnpj_limpo, mb_bur, sb_bur, _nb)
            if not _relatorio_ja_tem_chave(_it["Chave"]) and not _registro_manual_inutil_duplicada(
                st.session_state["relatorio"], mb_bur, sb_bur, _nb
            ):
                st.session_state["relatorio"].append(_it)
                n_b += 1
        if n_b:
            linhas.append(f"**{n_b}** inutilização(ões) a partir dos buracos.")

    n_p = 0
    err_p = None
    if up_inut_planilha is not None:
        _df_up, _err_up = dataframe_de_upload_inutil(up_inut_planilha)
        if _err_up:
            err_p = _err_up
        else:
            _tri, _err_tr = triplas_inutil_de_dataframe(_df_up)
            if _err_tr:
                err_p = _err_tr
            else:
                _df_bu = st.session_state["df_faltantes"].copy()
                _bur_t = conjunto_triplas_buracos(_df_bu)
                if not _bur_t:
                    err_p = "Não há buracos na auditoria — nada importado da planilha."
                else:
                    _aplic_rows = []
                    _ign = 0
                    _seen = set()
                    for _mod, _ser, _nota in _tri:
                        _k = (_mod.strip(), str(_ser).strip(), int(_nota))
                        if _k not in _bur_t:
                            _ign += 1
                            continue
                        if _k in _seen:
                            continue
                        _seen.add(_k)
                        _aplic_rows.append((_mod.strip(), str(_ser).strip(), int(_nota)))
                    if not _aplic_rows:
                        err_p = "Nenhuma linha da planilha coincide com buraco listado."
                    else:
                        n_p = 0
                        for _mod, _ser, _nota in _aplic_rows:
                            _itp = item_registro_manual_inutilizado(cnpj_limpo, _mod, _ser, _nota)
                            if not _relatorio_ja_tem_chave(_itp["Chave"]) and not _registro_manual_inutil_duplicada(
                                st.session_state["relatorio"], _mod, _ser, _nota
                            ):
                                st.session_state["relatorio"].append(_itp)
                                n_p += 1
                        if n_p:
                            linhas.append(f"**{n_p}** linha(s) da planilha (buracos).")
                        elif _aplic_rows:
                            linhas.append("Planilha: linhas já presentes no relatório (sem duplicar).")
                        if _ign > 0:
                            linhas.append(f"({_ign} linha(s) ignoradas — não eram buraco.)")

    _MAX_FAIXA_INUT = 5000
    n_f = 0
    err_f = None
    df_fb = st.session_state["df_faltantes"].copy()
    if not df_fb.empty and "Serie" in df_fb.columns and "Série" not in df_fb.columns:
        df_fb = df_fb.rename(columns={"Serie": "Série"})
    _bur_f = set()
    if (
        not df_fb.empty
        and {"Tipo", "Série", "Num_Faltante"}.issubset(df_fb.columns)
    ):
        _subf = df_fb[
            (df_fb["Tipo"].astype(str) == mf_faixa)
            & (df_fb["Série"].astype(str) == str(sf_faixa).strip())
        ]
        _bur_f = set(_subf["Num_Faltante"].astype(int).unique())
    if not sf_faixa or not str(sf_faixa).strip():
        pass
    elif n0_faixa > n1_faixa:
        err_f = "Faixa: a nota inicial não pode ser maior que a final."
    elif (n1_faixa - n0_faixa + 1) > _MAX_FAIXA_INUT:
        err_f = f"Faixa: máximo {_MAX_FAIXA_INUT} notas."
    elif not _bur_f:
        err_f = "Faixa: não há buracos para este modelo/série."
    else:
        _aplic = [n for n in range(int(n0_faixa), int(n1_faixa) + 1) if n in _bur_f]
        if _aplic:
            n_f = 0
            for _nn in _aplic:
                _itf = item_registro_manual_inutilizado(
                    cnpj_limpo, mf_faixa, str(sf_faixa).strip(), _nn
                )
                if not _relatorio_ja_tem_chave(_itf["Chave"]) and not _registro_manual_inutil_duplicada(
                    st.session_state["relatorio"], mf_faixa, str(sf_faixa).strip(), _nn
                ):
                    st.session_state["relatorio"].append(_itf)
                    n_f += 1
            if n_f:
                linhas.append(f"**{n_f}** nota(s) da faixa (buracos).")

    pick_bc = list(pick_bur_canc or [])
    n_b_c = 0
    if pick_bc:
        if mb_canc is None or sb_canc is None or str(sb_canc).strip() == "":
            return (
                False,
                "Para aplicar canceladas em «Dos buracos», escolha **modelo** e **série** nesse separador.",
                linhas,
            )
        for _nb in pick_bc:
            _itc = item_registro_manual_cancelado(cnpj_limpo, mb_canc, sb_canc, _nb)
            if not _relatorio_ja_tem_chave(_itc["Chave"]) and not _registro_manual_cancel_duplicada(
                st.session_state["relatorio"], mb_canc, sb_canc, _nb
            ):
                st.session_state["relatorio"].append(_itc)
                n_b_c += 1
        if n_b_c:
            linhas.append(f"**{n_b_c}** cancelada(s) a partir dos buracos.")

    n_p_c = 0
    err_p_c = None
    if up_canc_planilha is not None:
        _df_uc, _err_uc = dataframe_de_upload_inutil(up_canc_planilha)
        if _err_uc:
            err_p_c = _err_uc
        else:
            _trc, _err_trc = triplas_inutil_de_dataframe(_df_uc)
            if _err_trc:
                err_p_c = _err_trc
            else:
                _df_buc = st.session_state["df_faltantes"].copy()
                _bur_tc = conjunto_triplas_buracos(_df_buc)
                if not _bur_tc:
                    err_p_c = "Canceladas — não há buracos na auditoria — nada importado da planilha."
                else:
                    _aplic_c = []
                    _ign_c = 0
                    _seen_c = set()
                    for _mod, _ser, _nota in _trc:
                        _k = (_mod.strip(), str(_ser).strip(), int(_nota))
                        if _k not in _bur_tc:
                            _ign_c += 1
                            continue
                        if _k in _seen_c:
                            continue
                        _seen_c.add(_k)
                        _aplic_c.append((_mod.strip(), str(_ser).strip(), int(_nota)))
                    if not _aplic_c:
                        err_p_c = "Canceladas — nenhuma linha da planilha coincide com buraco listado."
                    else:
                        for _mod, _ser, _nota in _aplic_c:
                            _itpc = item_registro_manual_cancelado(cnpj_limpo, _mod, _ser, _nota)
                            if not _relatorio_ja_tem_chave(_itpc["Chave"]) and not _registro_manual_cancel_duplicada(
                                st.session_state["relatorio"], _mod, _ser, _nota
                            ):
                                st.session_state["relatorio"].append(_itpc)
                                n_p_c += 1
                        if n_p_c:
                            linhas.append(f"**{n_p_c}** linha(s) da planilha de **canceladas** (buracos).")
                        elif _aplic_c:
                            linhas.append("Planilha canceladas: linhas já presentes no relatório (sem duplicar).")
                        if _ign_c > 0:
                            linhas.append(f"({_ign_c} linha(s) ignoradas nas canceladas — não eram buraco.)")

    _MAX_FAIXA_CANC = 5000
    n_f_c = 0
    err_f_c = None
    df_fbc = st.session_state["df_faltantes"].copy()
    if not df_fbc.empty and "Serie" in df_fbc.columns and "Série" not in df_fbc.columns:
        df_fbc = df_fbc.rename(columns={"Serie": "Série"})
    _bur_fc = set()
    _mfc = mf_canc_f if mf_canc_f is not None else "NF-e"
    _sfc = (sf_canc_f if sf_canc_f is not None else "1") or "1"
    _n0c = int(n0_canc_f if n0_canc_f is not None else 1)
    _n1c = int(n1_canc_f if n1_canc_f is not None else 1)
    if (
        not df_fbc.empty
        and {"Tipo", "Série", "Num_Faltante"}.issubset(df_fbc.columns)
    ):
        _subfc = df_fbc[
            (df_fbc["Tipo"].astype(str) == _mfc)
            & (df_fbc["Série"].astype(str) == str(_sfc).strip())
        ]
        _bur_fc = set(_subfc["Num_Faltante"].astype(int).unique())
    if not str(_sfc).strip():
        pass
    elif _n0c > _n1c:
        err_f_c = "Canceladas — faixa: a nota inicial não pode ser maior que a final."
    elif (_n1c - _n0c + 1) > _MAX_FAIXA_CANC:
        err_f_c = f"Canceladas — faixa: máximo {_MAX_FAIXA_CANC} notas."
    elif not _bur_fc:
        err_f_c = "Canceladas — faixa: não há buracos para este modelo/série."
    else:
        _aplic_fc = [n for n in range(int(_n0c), int(_n1c) + 1) if n in _bur_fc]
        if _aplic_fc:
            for _nn in _aplic_fc:
                _itfc = item_registro_manual_cancelado(cnpj_limpo, _mfc, str(_sfc).strip(), _nn)
                if not _relatorio_ja_tem_chave(_itfc["Chave"]) and not _registro_manual_cancel_duplicada(
                    st.session_state["relatorio"], _mfc, str(_sfc).strip(), _nn
                ):
                    st.session_state["relatorio"].append(_itfc)
                    n_f_c += 1
            if n_f_c:
                linhas.append(f"**{n_f_c}** nota(s) da faixa de **canceladas** (buracos).")

    tem_ficheiros = _garimpo_existem_fontes_xml_lote()
    fez_algo = (
        n_extra > 0
        or n_b > 0
        or n_p > 0
        or n_f > 0
        or n_b_c > 0
        or n_p_c > 0
        or n_f_c > 0
    )
    quer_compare_autent = _tem_upload_autenticidade(up_autent_sefaz)

    if err_p:
        return False, err_p, linhas
    if err_p_c:
        return False, err_p_c, linhas
    if (err_f or err_f_c) and not fez_algo and not tem_ficheiros and not quer_compare_autent:
        _msg_ff = err_f or err_f_c
        return False, _msg_ff, linhas
    if err_f:
        linhas.append(f"Aviso (faixa inutil.): {err_f}")
    if err_f_c:
        linhas.append(f"Aviso (faixa cancel.): {err_f_c}")

    if not fez_algo and not tem_ficheiros and not quer_compare_autent:
        return (
            False,
            "Nada a processar: inclua XML/ZIP, inutilizações / canceladas manuais, "
            "o **relatório de autenticidade** (Sefaz) ou faça o garimpo para haver ficheiros em disco.",
            linhas,
        )

    if not quer_compare_autent and (fez_algo or tem_ficheiros):
        st.session_state.pop("df_divergencias", None)
        st.session_state["validation_done"] = False

    st.session_state["export_ready"] = False
    st.session_state["excel_buffer"] = None

    if tem_ficheiros:
        if len(cnpj) != 14:
            return False, "CNPJ inválido na barra lateral — necessário para reler os XML.", linhas
        ok, msg_rep = reprocessar_garimpeiro_a_partir_do_disco(cnpj_limpo)
        if not ok:
            return False, msg_rep, linhas
        linhas.append(msg_rep)
        if quer_compare_autent:
            ok_a, msg_a = _aplicar_comparacao_autenticidade(up_autent_sefaz)
            if not ok_a:
                return False, msg_a, linhas
            linhas.append(msg_a)
        return True, "\n\n".join(linhas), linhas

    reconstruir_dataframes_relatorio_simples()
    linhas.append(
        "Relatório recalculado (sem ficheiros na pasta de uploads — só registos manuais de inutil. / cancel.)."
    )
    if quer_compare_autent:
        ok_a, msg_a = _aplicar_comparacao_autenticidade(up_autent_sefaz)
        if not ok_a:
            return False, msg_a, linhas
        linhas.append(msg_a)
    return True, "\n\n".join(linhas), linhas


_V2_STATUS_UI_PARA_DF = {
    "Autorizadas": ["NORMAIS"],
    "Canceladas": ["CANCELADOS"],
    "Inutilizadas": ["INUTILIZADOS", "INUTILIZADA"],
    "Denegadas": ["DENEGADOS"],
    "Rejeitadas": ["REJEITADOS"],
}
_V2_OP_UI_PARA_INTERNO = {"Entrada": "ENTRADA", "Saída": "SAIDA"}
_V2_DATA_MODELO_LABELS = ("Qualquer", "Maior ou igual a", "Menor ou igual a", "Igual a", "Entre")
_V2_FAIXA_MODELO_LABELS = _V2_DATA_MODELO_LABELS


def v2_status_labels_para_valores(labels):
    out = []
    for L in labels or []:
        out.extend(_V2_STATUS_UI_PARA_DF.get(L, []))
    return list(dict.fromkeys(out))


def v2_op_labels_para_interno(labels):
    return [_V2_OP_UI_PARA_INTERNO[x] for x in (labels or []) if x in _V2_OP_UI_PARA_INTERNO]


def _mask_emissao_propria_df(df: pd.DataFrame) -> pd.Series:
    if df is None or df.empty or "Origem" not in df.columns:
        return pd.Series(dtype=bool)
    o = df["Origem"].astype(str)
    return o.str.contains("PRÓPRIA|PROPRIA", case=False, na=False, regex=True)


def _v2_parse_modo_data(label: str) -> str:
    return {
        "Qualquer": "",
        "Maior ou igual a": ">=",
        "Menor ou igual a": "<=",
        "Igual a": "=",
        "Entre": "entre",
    }.get(label or "", "")


def _v2_parse_modo_faixa(label: str) -> str:
    return _v2_parse_modo_data(label)


def _v2_aplicar_filtro_data_emissao(df_p: pd.DataFrame, modo: str, d1, d2) -> pd.DataFrame:
    if df_p.empty or not modo or "Data Emissão" not in df_p.columns:
        return df_p
    dt = pd.to_datetime(df_p["Data Emissão"], errors="coerce").dt.normalize()
    if d1 is not None:
        t1 = pd.Timestamp(d1).normalize()
    else:
        t1 = None
    if d2 is not None:
        t2 = pd.Timestamp(d2).normalize()
    else:
        t2 = None
    if modo == ">=" and t1 is not None:
        return df_p.loc[dt >= t1]
    if modo == "<=" and t1 is not None:
        return df_p.loc[dt <= t1]
    if modo == "=" and t1 is not None:
        return df_p.loc[dt == t1]
    if modo == "entre" and t1 is not None and t2 is not None:
        lo, hi = min(t1, t2), max(t1, t2)
        return df_p.loc[(dt >= lo) & (dt <= hi)]
    return df_p


def _v2_aplicar_filtro_faixa_nota(df_p: pd.DataFrame, modo: str, n1: int, n2: int) -> pd.DataFrame:
    if df_p.empty or not modo or "Nota" not in df_p.columns:
        return df_p
    nn = pd.to_numeric(df_p["Nota"], errors="coerce")
    if modo == ">=":
        return df_p.loc[nn >= n1]
    if modo == "<=":
        return df_p.loc[nn <= n1]
    if modo == "=":
        return df_p.loc[nn == n1]
    if modo == "entre":
        lo, hi = min(n1, n2), max(n1, n2)
        return df_p.loc[(nn >= lo) & (nn <= hi)]
    return df_p


def filtrar_df_geral_para_exportacao(
    df_base,
    filtro_origem,
    filtro_tipos,
    filtro_series,
    filtro_status_labels,
    filtro_operacao_labels,
    filtro_data_modo_label,
    filtro_data_d1,
    filtro_data_d2,
    filtro_faixa_modo_label,
    filtro_faixa_n1,
    filtro_faixa_n2,
    filtro_ufs,
    nota_esp_chave="",
    nota_esp_num=0,
    nota_esp_serie="",
    terceiros_status_labels=None,
    terceiros_tipos=None,
    terceiros_operacao_labels=None,
    terceiros_data_modo_label="Qualquer",
    terceiros_data_d1=None,
    terceiros_data_d2=None,
    *,
    skip_filtro_serie=False,
    skip_filtro_uf=False,
    skip_nota_especifica=False,
):
    """
    filtro_origem: aplica-se a todas as linhas (própria e/ou terceiros), antes do resto.
    Critérios v2_f_*: emissão própria. Nota específica: só própria.
    Critérios terceiros_*: só linhas de terceiros (XML recebidos).
    """
    if df_base is None or df_base.empty:
        return df_base
    out = df_base.copy()
    if len(filtro_origem) > 0:
        pat = "|".join([re.escape(o.split()[0]) for o in filtro_origem])
        out = out[out["Origem"].str.contains(pat, regex=True, na=False)]

    mask_p = _mask_emissao_propria_df(out)
    out_p = out.loc[mask_p].copy()
    out_t = out.loc[~mask_p].copy()

    st_vals = v2_status_labels_para_valores(filtro_status_labels)
    if st_vals:
        out_p = out_p[out_p["Status Final"].isin(st_vals)]

    if len(filtro_tipos) > 0:
        out_p = out_p[out_p["Modelo"].isin(filtro_tipos)]

    op_int = v2_op_labels_para_interno(filtro_operacao_labels)
    if op_int and "Operação" in out_p.columns:
        out_p = out_p[out_p["Operação"].isin(op_int)]

    modo_d = _v2_parse_modo_data(filtro_data_modo_label or "Qualquer")
    out_p = _v2_aplicar_filtro_data_emissao(out_p, modo_d, filtro_data_d1, filtro_data_d2)

    if not skip_filtro_serie and len(filtro_series) > 0:
        ser_set = {str(x) for x in filtro_series}
        out_p = out_p[out_p["Série"].astype(str).isin(ser_set)]

    modo_f = _v2_parse_modo_faixa(filtro_faixa_modo_label or "Qualquer")
    if modo_f and len(filtro_series) > 0:
        n1 = int(filtro_faixa_n1) if filtro_faixa_n1 is not None else 0
        n2 = int(filtro_faixa_n2) if filtro_faixa_n2 is not None else n1
        out_p = _v2_aplicar_filtro_faixa_nota(out_p, modo_f, n1, n2)

    if not skip_filtro_uf and len(filtro_ufs) > 0 and "UF Destino" in out_p.columns:
        ufs = {str(u).strip().upper() for u in filtro_ufs}
        out_p = out_p[out_p["UF Destino"].astype(str).str.upper().isin(ufs)]

    out_p = _v2_aplicar_nota_especifica_propria(
        out_p,
        nota_esp_chave,
        nota_esp_num,
        nota_esp_serie,
        skip=skip_nota_especifica,
    )

    _t_st_lab = list(terceiros_status_labels or [])
    _t_tip = list(terceiros_tipos or [])
    _t_op_lab = list(terceiros_operacao_labels or [])
    st_t_vals = v2_status_labels_para_valores(_t_st_lab)
    if st_t_vals and not out_t.empty:
        out_t = out_t[out_t["Status Final"].isin(st_t_vals)]
    if len(_t_tip) > 0 and not out_t.empty:
        out_t = out_t[out_t["Modelo"].isin(_t_tip)]
    op_t_int = v2_op_labels_para_interno(_t_op_lab)
    if op_t_int and not out_t.empty and "Operação" in out_t.columns:
        out_t = out_t[out_t["Operação"].isin(op_t_int)]
    modo_td = _v2_parse_modo_data(terceiros_data_modo_label or "Qualquer")
    if not out_t.empty:
        out_t = _v2_aplicar_filtro_data_emissao(
            out_t, modo_td, terceiros_data_d1, terceiros_data_d2
        )

    if out_p.empty and out_t.empty:
        return out_p
    if out_p.empty:
        return out_t.reset_index(drop=True)
    if out_t.empty:
        return out_p.reset_index(drop=True)
    return pd.concat([out_p, out_t], ignore_index=True)


def _mask_terceiros_df(df: pd.DataFrame) -> pd.Series:
    """Linhas cujo relatório indica documento recebido de terceiros."""
    if df is None or df.empty or "Origem" not in df.columns:
        return pd.Series(dtype=bool)
    return df["Origem"].astype(str).str.contains("TERCEIROS", case=False, na=False)


def _df_apenas_emissao_propria(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    if "Origem" not in df.columns:
        return df.reset_index(drop=True)
    mp = _mask_emissao_propria_df(df)
    mt = _mask_terceiros_df(df)
    # Própria explícita, ou linhas sem etiqueta clara (não ficam só de fora dos dois lotes)
    m = mp | (~mp & ~mt)
    return df.loc[m].reset_index(drop=True)


def _df_apenas_terceiros(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    if "Origem" not in df.columns:
        return pd.DataFrame()
    mt = _mask_terceiros_df(df)
    return df.loc[mt].reset_index(drop=True)


def _folhas_detalhe_terceiros_do_subset(df: pd.DataFrame) -> dict:
    """
    Folhas do livro Excel alinhadas ao painel «terceiros»: canceladas, autorizadas,
    denegadas e rejeitadas a partir de linhas TERCEIROS. Buracos e inutilizadas
    não se aplicam a XML recebido de terceiros — df_bur e df_inu ficam sempre vazios.
    """
    empty = pd.DataFrame()
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return {
            "df_bur": empty,
            "df_inu": empty,
            "df_can": empty,
            "df_aut": empty,
            "df_den": empty,
            "df_rej": empty,
        }
    df_t = df.loc[_mask_terceiros_df(df)].copy()
    if df_t.empty or "Status Final" not in df_t.columns:
        return {
            "df_bur": empty,
            "df_inu": empty,
            "df_can": empty,
            "df_aut": empty,
            "df_den": empty,
            "df_rej": empty,
        }
    stu = df_t["Status Final"].astype(str).str.upper()
    return {
        "df_bur": empty,
        "df_inu": empty,
        "df_can": df_t.loc[stu.eq("CANCELADOS")].reset_index(drop=True),
        "df_aut": df_t.loc[stu.eq("NORMAIS")].reset_index(drop=True),
        "df_den": df_t.loc[stu.eq("DENEGADOS")].reset_index(drop=True),
        "df_rej": df_t.loc[stu.eq("REJEITADOS")].reset_index(drop=True),
    }


def _v2_limpar_zips_gerados_etapa3(output_dir=None):
    """Política: não apagar ZIP já gravados no disco — só gravar novos por cima se o nome repetir."""
    return


def _v2_excel_bytes_filtrado_etapa3(df: pd.DataFrame):
    """Um .xlsx com folhas Filtrado e Resumo_status (igual ao ramo excel_filtro da Etapa 3)."""
    if df is None or df.empty:
        return None
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
        _df_xls_f = _df_com_data_emissao_dd_mm_yyyy(df.reset_index(drop=True))
        _df_xls_f.to_excel(writer, sheet_name="Filtrado", index=False)
        rs = (
            _df_xls_f.groupby("Status Final", dropna=False)
            .size()
            .reset_index(name="Quantidade")
        )
        rs.to_excel(writer, sheet_name="Resumo_status", index=False)
    return buffer_excel.getvalue()


def excel_bytes_relatorio_bloco(df_filtrado: pd.DataFrame, chaves_bloco: set):
    """Bytes de um .xlsx só com as linhas cujas Chave aparecem no bloco de XML (máx. 10k ficheiros)."""
    if df_filtrado is None or df_filtrado.empty or not chaves_bloco:
        return None
    if "Chave" not in df_filtrado.columns:
        return None
    _norm = df_filtrado["Chave"].map(_chave_para_conjunto_export)
    dfp = df_filtrado.loc[_norm.isin(chaves_bloco)]
    if dfp.empty:
        return None
    dfp = _df_com_data_emissao_dd_mm_yyyy(dfp.reset_index(drop=True))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        dfp.to_excel(writer, sheet_name="Filtrado", index=False)
        rs = (
            dfp.groupby("Status Final", dropna=False)
            .size()
            .reset_index(name="Quantidade")
        )
        rs.to_excel(writer, sheet_name="Resumo_status", index=False)
    return buf.getvalue()


def _excel_bytes_pacote_contabilidade(
    df_todas_lidas: pd.DataFrame, df_xml_pacote: pd.DataFrame
):
    """
    Excel para envio à contabilidade: folha «Todas_lidas» + resumo; se df_xml_pacote for outro
    DataFrame (não o mesmo objeto que df_todas_lidas), acrescenta XML_neste_pacote e resumo.
    Folhas com datas em dd/mm/aaaa.
    Devolve bytes ou None se falhar por disco/temp cheio (errno 28).
    """
    has_t = df_todas_lidas is not None and not df_todas_lidas.empty
    has_x = df_xml_pacote is not None and not df_xml_pacote.empty
    mesmo_lote = has_t and has_x and (df_xml_pacote is df_todas_lidas)
    if not has_t and not has_x:
        try:
            buf0 = io.BytesIO()
            with pd.ExcelWriter(buf0, engine="xlsxwriter") as w0:
                pd.DataFrame({"Info": ["Sem linhas para exportar no Excel do pacote."]}).to_excel(
                    w0, sheet_name="Aviso", index=False
                )
            return buf0.getvalue()
        except Exception as e:
            if _erro_sem_espaco_disco(e):
                return None
            raise
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            if has_t:
                d1 = _df_com_data_emissao_dd_mm_yyyy(df_todas_lidas.reset_index(drop=True))
                d1.to_excel(writer, sheet_name="Todas_lidas", index=False)
                if "Status Final" in d1.columns:
                    rs = (
                        d1.groupby("Status Final", dropna=False)
                        .size()
                        .reset_index(name="Quantidade")
                    )
                    rs.to_excel(writer, sheet_name="Resumo_todas", index=False)
            if has_x and not mesmo_lote:
                d2 = _df_com_data_emissao_dd_mm_yyyy(df_xml_pacote.reset_index(drop=True))
                d2.to_excel(writer, sheet_name="XML_neste_pacote", index=False)
                if "Status Final" in d2.columns:
                    r2 = (
                        d2.groupby("Status Final", dropna=False)
                        .size()
                        .reset_index(name="Quantidade")
                    )
                    r2.to_excel(writer, sheet_name="Resumo_XML_pacote", index=False)
        return buf.getvalue()
    except Exception as e:
        if _erro_sem_espaco_disco(e):
            return None
        raise


def _excel_bytes_geral_e_resumo_status(df: pd.DataFrame) -> bytes:
    """Um .xlsx com folhas Geral e Resumo_status (datas em dd/mm/aaaa)."""
    buf = io.BytesIO()
    _ddf = _df_com_data_emissao_dd_mm_yyyy(df.reset_index(drop=True))
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        _ddf.to_excel(writer, sheet_name="Geral", index=False)
        rs = (
            _ddf.groupby("Status Final", dropna=False)
            .size()
            .reset_index(name="Quantidade")
        )
        rs.to_excel(writer, sheet_name="Resumo_status", index=False)
    return buf.getvalue()


def _v2_export_zip_etapa3(
    df_g_base: pd.DataFrame,
    *,
    xml_respeita_filtro: bool,
    df_filtrado_para_excel_bloco=None,
    excel_um_só_completo: bool,
    df_excel_completo: pd.DataFrame,
    v2_zip_org: bool,
    v2_zip_plano: bool,
    cnpj_limpo: str,
    zip_tag=None,
    zip_output_dir=None,
    zip_nome_ficheiro="",
    pacote_pastas_garimpo: bool = False,
    pacote_pastas_contabilidade: bool = False,
    df_excel_todas_notas=None,
):
    """
    Escreve ZIP(s) em disco (org / plano).
    Devolve (org_parts, todos_parts, xml_matched, aviso_sem_xml|None, excel_pacote_matriz_solta|None).
    xml_respeita_filtro=False → todos os XML cujo resumo está em df_g_base.
    excel_um_só_completo=True → em cada parte, o mesmo Excel com df_excel_completo (relatório inteiro).
    zip_tag: ex. «propria» / «terceiros» para nomes z_org_propria_pt1.zip (dois lotes independentes).
    zip_output_dir: pasta onde gravar os .zip (None = diretório de trabalho atual).
    zip_nome_ficheiro: opcional — base do nome (ex.: GarimpoMarco → GarimpoMarco_org_propria_pt1.zip).
    pacote_pastas_garimpo=True → um ZIP só «com pastas»; stem = zip_nome_ficheiro ou pacote_apuracao_ptN.zip.
    pacote_pastas_contabilidade=True → ZIP Emitidas_*_Serie_*_*Mes_AAAA_MM* ou Terceiros_*; XML em XML/Lote_001…;
    Excel na raiz do ZIP;
    Excel em cada ZIP: mesmo conteúdo (todo o lido), nome `relatorio_garimpeiro_<grupo>.xlsx` (sem Painel Fiscal).
    df_excel_todas_notas: DataFrame do relatório geral (df_geral) para a folha «Todas_lidas».
    """
    out_dir = (
        Path(zip_output_dir).resolve()
        if zip_output_dir is not None
        else Path(".").resolve()
    )
    if pacote_pastas_garimpo:
        u = _v2_sanitize_nome_export(zip_nome_ficheiro or "", max_len=80)
        stem_org = u if u else "pacote_apuracao"
        stem_todos = "z_todos_unused"
    else:
        stem_org, stem_todos = _v2_stems_zip_nome_ficheiro_etapa3(
            zip_nome_ficheiro or "", zip_tag
        )
    if df_g_base is None or df_g_base.empty or "Chave" not in df_g_base.columns:
        return [], [], 0, "ERR:Relatório geral vazio ou sem coluna Chave.", None
    if xml_respeita_filtro:
        if df_filtrado_para_excel_bloco is None or df_filtrado_para_excel_bloco.empty:
            return [], [], 0, "ERR:Resultado filtrado: 0 linhas. Ajuste os filtros.", None
        _df_ch = df_filtrado_para_excel_bloco
    else:
        _df_ch = df_g_base
    filtro_chaves = {
        k
        for k in (_chave_para_conjunto_export(x) for x in _df_ch["Chave"].tolist())
        if k
    }
    if not filtro_chaves:
        return [], [], 0, "ERR:Nenhuma chave válida para exportar XML.", None

    excel_fn_completo = (
        _PACOTE_CONTAB_NOME_EXCEL_RAIZ
        if pacote_pastas_contabilidade and excel_um_só_completo
        else "RELATORIO_GARIMPEIRO/relatorio_geral_completo.xlsx"
    )

    xb_completo = None
    aviso_sem_espaco_excel = None
    if excel_um_só_completo:
        if pacote_pastas_contabilidade:
            df_xml_plan = (
                df_filtrado_para_excel_bloco
                if xml_respeita_filtro
                else df_g_base
            )
            df_todas = df_excel_todas_notas
            if df_todas is None or getattr(df_todas, "empty", True):
                df_todas = df_g_base
            try:
                xb_completo = excel_relatorio_geral_com_dashboard_bytes(
                    df_todas, incluir_painel_fiscal=False
                )
            except Exception as ex_dash:
                if _erro_sem_espaco_disco(ex_dash):
                    aviso_sem_espaco_excel = _msg_sem_espaco_disco_garimpeiro()
                xb_completo = None
            if not xb_completo:
                xb_completo = _excel_bytes_pacote_contabilidade(df_todas, df_xml_plan)
                if xb_completo is None and not aviso_sem_espaco_excel:
                    aviso_sem_espaco_excel = _msg_sem_espaco_disco_garimpeiro()
        else:
            xb_completo = _excel_bytes_geral_e_resumo_status(df_excel_completo)

    if pacote_pastas_contabilidade and v2_zip_org and not v2_zip_plano:
        paths, empty, xm, av, xls = _v2_export_pacote_contab_por_dimensoes(
            out_dir,
            stem_org,
            filtro_chaves,
            cnpj_limpo,
            xb_completo,
            excel_fn_completo,
            _df_ch,
        )
        if aviso_sem_espaco_excel:
            av = (
                f"{aviso_sem_espaco_excel} {av}"
                if av
                else aviso_sem_espaco_excel
            )
        return paths, empty, xm, av, xls

    Z = {
        "z_org": None,
        "z_todos": None,
        "org_parts": [],
        "todos_parts": [],
        "org_count": 0,
        "todos_count": 0,
        "curr_org_part": 1,
        "curr_todos_part": 1,
        "chaves_bloco": set(),
        "seq_bloco": 1,
        "xml_matched": 0,
    }

    def _fechar_bloco_zip():
        if excel_um_só_completo:
            excel_fn = excel_fn_completo
            xb = xb_completo
        else:
            excel_fn = f"RELATORIO_GARIMPEIRO/relatorio_filtrado_pt{Z['seq_bloco']:03d}.xlsx"
            xb = excel_bytes_relatorio_bloco(
                df_filtrado_para_excel_bloco, Z["chaves_bloco"]
            )
        if xb:
            if v2_zip_org and Z["z_org"] is not None:
                Z["z_org"].writestr(excel_fn, xb)
            if v2_zip_plano and Z["z_todos"] is not None:
                Z["z_todos"].writestr(excel_fn, xb)
        Z["chaves_bloco"].clear()
        Z["seq_bloco"] += 1
        if (
            v2_zip_org
            and Z["z_org"] is not None
            and Z["org_count"] >= MAX_XML_PER_ZIP
        ):
            try:
                Z["z_org"].close()
            except OSError:
                pass
            Z["curr_org_part"] += 1
            oname = str(out_dir / f"{stem_org}_pt{Z['curr_org_part']}.zip")
            Z["z_org"] = _zipfile_open_write_export(oname)
            Z["org_parts"].append(oname)
            Z["org_count"] = 0
        if (
            v2_zip_plano
            and Z["z_todos"] is not None
            and Z["todos_count"] >= MAX_XML_PER_ZIP
        ):
            try:
                Z["z_todos"].close()
            except OSError:
                pass
            Z["curr_todos_part"] += 1
            tname = str(out_dir / f"{stem_todos}_pt{Z['curr_todos_part']}.zip")
            Z["z_todos"] = _zipfile_open_write_export(tname)
            Z["todos_parts"].append(tname)
            Z["todos_count"] = 0

    if v2_zip_org:
        org_name = str(out_dir / f"{stem_org}_pt{Z['curr_org_part']}.zip")
        Z["z_org"] = _zipfile_open_write_export(org_name)
        Z["org_parts"].append(org_name)
    if v2_zip_plano:
        todos_name = str(out_dir / f"{stem_todos}_pt{Z['curr_todos_part']}.zip")
        Z["z_todos"] = _zipfile_open_write_export(todos_name)
        Z["todos_parts"].append(todos_name)

    if v2_zip_org or v2_zip_plano:
        chaves_ja_org = set()
        chaves_ja_todos = set()
        for f_name in _lista_nomes_fontes_xml_garimpo():
            with _abrir_fonte_xml_garimpo_stream(f_name) as f_temp:
                for name, xml_data in extrair_recursivo(f_temp, f_name):
                    res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                    ck = _chave_para_conjunto_export(res["Chave"]) if res else None
                    if res and ck and ck in filtro_chaves:
                        td = _tupla_dedupe_export_xml(res, ck)
                        if td is None:
                            del xml_data
                            continue
                        org_ok = (
                            v2_zip_org
                            and Z["z_org"] is not None
                            and td not in chaves_ja_org
                        )
                        todos_ok = (
                            v2_zip_plano
                            and Z["z_todos"] is not None
                            and td not in chaves_ja_todos
                        )
                        if not org_ok and not todos_ok:
                            del xml_data
                            continue
                        if org_ok:
                            chaves_ja_org.add(td)
                            if pacote_pastas_contabilidade:
                                inner = _caminho_xml_pacote_contab_raiz(res, name)
                            else:
                                inner = f"{res['Pasta']}/{name}"
                            Z["z_org"].writestr(inner, xml_data)
                            Z["org_count"] += 1
                        if todos_ok:
                            chaves_ja_todos.add(td)
                            Z["z_todos"].writestr(name, xml_data)
                            Z["todos_count"] += 1
                        Z["xml_matched"] += int(org_ok) + int(todos_ok)
                        Z["chaves_bloco"].add(ck)
                        limite = (
                            v2_zip_org and Z["org_count"] >= MAX_XML_PER_ZIP
                        ) or (v2_zip_plano and Z["todos_count"] >= MAX_XML_PER_ZIP)
                        if limite:
                            _fechar_bloco_zip()
                    del xml_data

    if Z["chaves_bloco"] and (
        (v2_zip_org and Z["org_count"] > 0) or (v2_zip_plano and Z["todos_count"] > 0)
    ):
        excel_fn_last = (
            excel_fn_completo
            if excel_um_só_completo
            else f"RELATORIO_GARIMPEIRO/relatorio_filtrado_pt{Z['seq_bloco']:03d}.xlsx"
        )
        if excel_um_só_completo:
            xb_last = xb_completo
        else:
            xb_last = excel_bytes_relatorio_bloco(
                df_filtrado_para_excel_bloco, Z["chaves_bloco"]
            )
        if xb_last:
            if v2_zip_org and Z["z_org"] is not None and Z["org_count"] > 0:
                Z["z_org"].writestr(excel_fn_last, xb_last)
            if v2_zip_plano and Z["z_todos"] is not None and Z["todos_count"] > 0:
                Z["z_todos"].writestr(excel_fn_last, xb_last)

    if Z["z_org"] is not None:
        try:
            Z["z_org"].close()
        except OSError:
            pass
    if Z["z_todos"] is not None:
        try:
            Z["z_todos"].close()
        except OSError:
            pass

    org_parts = Z["org_parts"]
    todos_parts = Z["todos_parts"]
    if v2_zip_org and Z["org_count"] == 0 and org_parts:
        org_parts = []
    if v2_zip_plano and Z["todos_count"] == 0 and todos_parts:
        todos_parts = []

    aviso = None
    if (v2_zip_org or v2_zip_plano) and Z.get("xml_matched", 0) == 0:
        aviso = (
            "Nenhum XML em disco correspondeu às chaves. "
            "Causas frequentes: pasta do garimpo apagada, ou chaves na tabela que não batem com os ficheiros."
        )
    return org_parts, todos_parts, Z["xml_matched"], aviso, None


def _v2_export_zip_mariana(
    df_geral: pd.DataFrame,
    cnpj_limpo: str,
    *,
    zip_output_dir: Path,
    zip_file_stem=None,
):
    """
    Pacote contabilidade/matriz: **todo** o lote lido no garimpo (df_geral), **sem** filtros da Etapa 3.
    ZIPs: **Emitidas** por série, status e mês (`…_Mes_AAAA_MM`), **Terceiros** por modelo×status;
    sufixo _notas_min_max quando aplicável; dentro: XML/Lote_001… (10k/pasta) + Excel na raiz.
    """
    nome = str(zip_file_stem or "").strip()
    o, _t, xm, av, xls = _v2_export_zip_etapa3(
        df_geral,
        xml_respeita_filtro=False,
        df_filtrado_para_excel_bloco=df_geral,
        excel_um_só_completo=True,
        df_excel_completo=df_geral,
        v2_zip_org=True,
        v2_zip_plano=False,
        cnpj_limpo=cnpj_limpo,
        zip_tag=None,
        zip_output_dir=zip_output_dir,
        zip_nome_ficheiro=nome,
        pacote_pastas_garimpo=True,
        pacote_pastas_contabilidade=True,
        df_excel_todas_notas=df_geral,
    )
    return o, xm, av, xls


def _v2_uniq_sorted_str_series_vals(s: pd.Series) -> list:
    return sorted(
        {str(x) for x in s.tolist() if str(x) not in ("", "nan", "None")},
        key=lambda x: (len(x), x),
    )


def v2_opcoes_cascata_etapa3(
    df_base: pd.DataFrame,
    filtro_origem: list,
    filtro_tipos: list,
    filtro_series: list,
    filtro_status_labels: list,
    filtro_operacao_labels: list,
    filtro_data_modo_label: str,
    filtro_data_d1,
    filtro_data_d2,
    filtro_faixa_modo_label: str,
    filtro_faixa_n1: int,
    filtro_faixa_n2: int,
    filtro_ufs: list,
    nota_esp_chave: str,
    nota_esp_num: int,
    nota_esp_serie: str,
    terceiros_status_labels: list,
    terceiros_tipos: list,
    terceiros_operacao_labels: list,
    terceiros_data_modo_label: str,
    terceiros_data_d1,
    terceiros_data_d2,
) -> dict:
    """Listas dependentes para Série e UF (só emissão própria), dados os outros filtros."""
    empty = {"series": [], "ufs": []}
    if df_base is None or df_base.empty:
        return empty

    def _series_em_propria(df):
        if df is None or df.empty or "Série" not in df.columns:
            return []
        m = _mask_emissao_propria_df(df)
        return _v2_uniq_sorted_str_series_vals(df.loc[m, "Série"].astype(str))

    def _ufs_em_propria(df):
        if df is None or df.empty or "UF Destino" not in df.columns:
            return []
        m = _mask_emissao_propria_df(df)
        ser = df.loc[m, "UF Destino"].astype(str).str.upper().str.strip()
        return sorted({x for x in ser.tolist() if x and x not in ("NAN", "NONE", "")})

    d_ser = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_tipos,
        filtro_series,
        filtro_status_labels,
        filtro_operacao_labels,
        filtro_data_modo_label,
        filtro_data_d1,
        filtro_data_d2,
        filtro_faixa_modo_label,
        filtro_faixa_n1,
        filtro_faixa_n2,
        filtro_ufs,
        nota_esp_chave,
        nota_esp_num,
        nota_esp_serie,
        terceiros_status_labels,
        terceiros_tipos,
        terceiros_operacao_labels,
        terceiros_data_modo_label,
        terceiros_data_d1,
        terceiros_data_d2,
        skip_filtro_serie=True,
        skip_filtro_uf=False,
        skip_nota_especifica=True,
    )
    d_uf = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_tipos,
        filtro_series,
        filtro_status_labels,
        filtro_operacao_labels,
        filtro_data_modo_label,
        filtro_data_d1,
        filtro_data_d2,
        filtro_faixa_modo_label,
        filtro_faixa_n1,
        filtro_faixa_n2,
        filtro_ufs,
        nota_esp_chave,
        nota_esp_num,
        nota_esp_serie,
        terceiros_status_labels,
        terceiros_tipos,
        terceiros_operacao_labels,
        terceiros_data_modo_label,
        terceiros_data_d1,
        terceiros_data_d2,
        skip_filtro_serie=False,
        skip_filtro_uf=True,
        skip_nota_especifica=True,
    )
    return {
        "series": _series_em_propria(d_ser),
        "ufs": _ufs_em_propria(d_uf),
    }


def v2_sanear_selecoes_contra_opcoes(series_opts: list, ufs_opts: list) -> None:
    """Remove da sessão valores que deixaram de existir nas listas em cascata (série / UF)."""
    for key, permitidos in (
        ("v2_f_ser", set(series_opts)),
        ("v2_f_uf", set(ufs_opts)),
    ):
        cur = list(st.session_state.get(key) or [])
        novo = [x for x in cur if x in permitidos]
        if novo != cur:
            st.session_state[key] = novo


def _v2_remover_zips_em_disco_listados():
    """Política: nunca apagar ficheiros exportados que o utilizador guardou em pasta."""
    return


def _v2_limpar_estado_exportacao_etapa3(remover_zip_do_disco=True):
    """
    Reinicia sessão de downloads da Etapa 3.
    Os ZIP no disco do utilizador **não são apagados** (política: só gravar).
    remover_zip_do_disco mantido por compatibilidade; a remoção em disco está desativada.
    """
    if remover_zip_do_disco:
        _v2_remover_zips_em_disco_listados()
    st.session_state["export_ready"] = False
    st.session_state["org_zip_parts"] = []
    st.session_state["todos_zip_parts"] = []
    st.session_state["excel_buffer"] = None
    st.session_state.pop("excel_buffer_propria", None)
    st.session_state.pop("excel_buffer_terceiros", None)
    st.session_state.pop("export_excel_name_propria", None)
    st.session_state.pop("export_excel_name_terceiros", None)
    st.session_state.pop("v2_export_sig", None)
    st.session_state.pop("v2_export_sem_xml", None)
    st.session_state["v2_etapa3_dual_export"] = False
    st.session_state.pop("v2_export_lados", None)
    st.session_state.pop("mariana_zip_parts", None)
    st.session_state["mariana_export_ready"] = False
    st.session_state.pop("mariana_export_sig", None)
    st.session_state.pop("mariana_sem_xml_msg", None)
    st.session_state.pop("mariana_excel_completo_path", None)


def v2_assinatura_exportacao_sessao():
    """
    Identifica filtros + confirmação «exportar tudo» + formato (ZIP/Excel).
    Se mudar após uma geração, os downloads guardados deixam de ser válidos.
    """
    return (
        tuple(st.session_state.get("v2_f_orig") or []),
        tuple(st.session_state.get("v2_f_tip") or []),
        tuple(st.session_state.get("v2_f_ser") or []),
        tuple(st.session_state.get("v2_f_stat") or []),
        tuple(st.session_state.get("v2_f_op") or []),
        tuple(st.session_state.get("v2_f_uf") or []),
        str(st.session_state.get("v2_f_data_modo", "Qualquer")),
        str(st.session_state.get("v2_f_data_d1", "")),
        str(st.session_state.get("v2_f_data_d2", "")),
        str(st.session_state.get("v2_f_faixa_modo", "Qualquer")),
        int(st.session_state.get("v2_f_faixa_n1") or 0),
        int(st.session_state.get("v2_f_faixa_n2") or 0),
        str(st.session_state.get("v2_f_esp_chave", "")),
        int(st.session_state.get("v2_f_esp_num") or 0),
        str(st.session_state.get("v2_f_esp_ser", "")),
        tuple(st.session_state.get("v2_t_stat") or []),
        tuple(st.session_state.get("v2_t_tip") or []),
        tuple(st.session_state.get("v2_t_op") or []),
        str(st.session_state.get("v2_t_data_modo", "Qualquer")),
        str(st.session_state.get("v2_t_data_d1", "")),
        str(st.session_state.get("v2_t_data_d2", "")),
        bool(st.session_state.get("v2_confirm_full", True)),
        str(st.session_state.get("v2_export_format", "zip_tudo_pastas")),
    )


def v2_assinatura_pacote_matriz_sessao(df_geral):
    """
    Pacote contabilidade / matriz: **não** inclui filtros da Etapa 3.
    Invalida ao mudar o relatório geral (novo garimpo), pasta/nome do ZIP ou CNPJ na sessão.
    """
    return (
        id(df_geral) if df_geral is not None else 0,
        str(st.session_state.get("mariana_zip_save_dir") or ""),
        str(st.session_state.get("mariana_zip_basename") or ""),
        str(st.session_state.get("cnpj_widget") or ""),
    )


def v2_callback_repor_filtros():
    """Limpa multiselects da Etapa 3. Deve ser usado com on_click (antes dos widgets na mesma corrida)."""
    for _kx in (
        "v2_f_orig",
        "v2_f_tip",
        "v2_f_ser",
        "v2_f_stat",
        "v2_f_op",
        "v2_f_uf",
        "v2_t_stat",
        "v2_t_tip",
        "v2_t_op",
    ):
        st.session_state[_kx] = []
    st.session_state["v2_f_data_modo"] = "Qualquer"
    st.session_state["v2_f_faixa_modo"] = "Qualquer"
    st.session_state["v2_f_faixa_n1"] = 0
    st.session_state["v2_f_faixa_n2"] = 0
    st.session_state.pop("v2_f_data_d1", None)
    st.session_state.pop("v2_f_data_d2", None)
    st.session_state["v2_f_esp_chave"] = ""
    st.session_state["v2_f_esp_num"] = 0
    st.session_state["v2_f_esp_ser"] = ""
    st.session_state["v2_t_data_modo"] = "Qualquer"
    st.session_state.pop("v2_t_data_d1", None)
    st.session_state.pop("v2_t_data_d2", None)
    st.session_state["v2_export_format"] = "zip_tudo_pastas"
    _v2_limpar_estado_exportacao_etapa3()
    st.session_state.pop("_v2_e3_preview_dis_cache", None)


def rotulo_download_zip_parte(caminho_ficheiro):
    base = os.path.basename(caminho_ficheiro)
    m = re.search(r"pt(\d+)\.zip$", base, re.I)
    if m:
        return f"ZIP — parte {m.group(1)}"
    if "__" in base and base.lower().endswith(".zip"):
        return base[:-4].split("__", 1)[-1]
    if base.lower().endswith(".zip"):
        return base[:-4]
    return base


def enumerar_buracos_por_segmento(nums_sorted, tipo_doc, serie_str, gap_max=MAX_SALTO_ENTRE_NOTAS_CONSECUTIVAS):
    """Buracos só dentro de cada trecho; saltos grandes quebram o trecho (não preenche o intervalo entre faixas)."""
    out = []
    if not nums_sorted:
        return out
    segmentos = [[nums_sorted[0]]]
    for i in range(1, len(nums_sorted)):
        if nums_sorted[i] - nums_sorted[i - 1] > gap_max:
            segmentos.append([nums_sorted[i]])
        else:
            segmentos[-1].append(nums_sorted[i])
    for seg in segmentos:
        lo, hi = seg[0], seg[-1]
        seg_set = set(seg)
        for b in range(lo, hi + 1):
            if b not in seg_set:
                out.append({"Tipo": tipo_doc, "Série": serie_str, "Num_Faltante": b})
    return out


def extrair_chaves_de_excel(arquivo_excel):
    chaves = []
    try:
        df_keys = pd.read_excel(arquivo_excel, header=None)
        for _, row in df_keys.iterrows():
            raw = row.iloc[0]
            if pd.isna(raw):
                continue
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                try:
                    f = float(raw)
                    s = str(int(f)) if f.is_integer() else str(raw).strip()
                except (ValueError, OverflowError):
                    s = str(raw).strip()
            else:
                s = str(raw).strip()
            digitos = "".join(filter(str.isdigit, s))
            if len(digitos) >= 44:
                chaves.append(digitos[:44])
    except Exception:
        pass
    return list(dict.fromkeys(chaves))


_MAX_FAIXA_EXPORT_DOM = 5000  # Máx. largura de faixa por linha (lista específica / inutilizadas)


def _excel_celula_int(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return None
        try:
            return int(float(s.replace(",", ".")))
        except ValueError:
            return None
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None


def _excel_celula_serie(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        try:
            f = float(val)
            if f.is_integer():
                return str(int(f))
        except (ValueError, OverflowError):
            pass
    return str(val).strip()


def _coluna_por_palavras(nomes_cols, palavras, ja_usados):
    """Índice da primeira coluna cujo nome contém alguma palavra-chave (não em ja_usados)."""
    for i, nome in enumerate(nomes_cols):
        if i in ja_usados:
            continue
        c = str(nome).strip().lower().replace("_", " ")
        comp = c.replace(" ", "")
        for p in palavras:
            p2 = p.lower().replace(" ", "")
            if p2 in comp or p.lower() in c:
                return i
    return None


def extrair_faixas_ini_fim_serie_excel(arquivo_excel):
    """
    Planilha com numeração inicial, final e série (3 colunas).
    Aceita cabeçalhos em português ou dados nas colunas A, B, C sem título.
    Retorno: (lista de dicts n_ini, n_fim, serie, linhas_ignoradas, mensagem_erro).
    """
    try:
        df = pd.read_excel(arquivo_excel)
    except Exception:
        return [], 0, "Não foi possível ler o ficheiro Excel."

    if df is None or df.empty:
        return [], 0, "Planilha vazia."

    nomes = list(df.columns)
    lowered = [str(x).strip().lower() for x in nomes]

    i_ini = _coluna_por_palavras(
        lowered,
        [
            "numeracao inicial",
            "numeração inicial",
            "nota inicial",
            "n inicial",
            "inicial",
            "inicio",
            "início",
        ],
        set(),
    )
    i_fim = _coluna_por_palavras(
        lowered,
        [
            "numeracao final",
            "numeração final",
            "nota final",
            "n final",
            "final",
            "fim",
            "até",
            "ate",
        ],
        {i_ini} if i_ini is not None else set(),
    )
    _us_ser = set()
    if i_ini is not None:
        _us_ser.add(i_ini)
    if i_fim is not None:
        _us_ser.add(i_fim)
    i_ser = _coluna_por_palavras(
        lowered,
        ["serie", "série", "ser"],
        _us_ser,
    )

    if i_ini is None or i_fim is None or i_ser is None:
        if len(nomes) >= 3:
            i_ini, i_fim, i_ser = 0, 1, 2
        else:
            return (
                [],
                0,
                "Indique 3 colunas (inicial, final, série) ou use cabeçalhos reconhecíveis.",
            )

    c_ini, c_fim, c_ser = nomes[i_ini], nomes[i_fim], nomes[i_ser]
    faixas = []
    ignoradas = 0

    for _, row in df.iterrows():
        n0 = _excel_celula_int(row[c_ini])
        n1 = _excel_celula_int(row[c_fim])
        ser = _excel_celula_serie(row[c_ser])
        if n0 is None or n1 is None or not ser:
            ignoradas += 1
            continue
        if n0 > n1:
            n0, n1 = n1, n0
        if (n1 - n0 + 1) > _MAX_FAIXA_EXPORT_DOM:
            ignoradas += 1
            continue
        faixas.append({"n_ini": n0, "n_fim": n1, "serie": ser})

    if not faixas:
        msg = "Nenhuma linha válida."
        if ignoradas:
            msg += f" ({ignoradas} linha(s) ignorada(s): vazias, série em falta ou faixa acima de {_MAX_FAIXA_EXPORT_DOM} notas.)"
        return [], ignoradas, msg

    return faixas, ignoradas, None


_MAX_CHAVES_EXCEL_FAIXAS = 75000  # Limite de chaves agregadas por planilha (várias linhas)


def chaves_agregadas_de_excel_faixas(df_geral, faixas_lista, modelo):
    """Cruza cada faixa com o relatório geral; devolve (chaves únicas, cortado_por_limite)."""
    if df_geral is None or df_geral.empty or not faixas_lista:
        return [], False
    vistos = set()
    ordenadas = []
    for fx in faixas_lista:
        ch_sub = chaves_por_faixa_numeracao(
            df_geral,
            modelo,
            fx["serie"],
            fx["n_ini"],
            fx["n_fim"],
        )
        for ch in ch_sub:
            if ch not in vistos:
                vistos.add(ch)
                ordenadas.append(ch)
                if len(ordenadas) >= _MAX_CHAVES_EXCEL_FAIXAS:
                    return ordenadas, True
    return ordenadas, False


def _nome_xml_raiz_zip_unico(usados, nome_arquivo):
    """Nome dentro do ZIP só na raiz; evita colisão se houver ficheiros homónimos."""
    base = os.path.basename(str(nome_arquivo).replace("\\", "/"))
    if not base or base in (".", ".."):
        base = "documento.xml"
    if base not in usados:
        usados.add(base)
        return base
    stem, ext = os.path.splitext(base)
    if not ext:
        ext = ".xml"
    k = 2
    while True:
        cand = f"{stem}_{k}{ext}"
        if cand not in usados:
            usados.add(cand)
            return cand
        k += 1


def _chave44_digitos(ch):
    d = "".join(filter(str.isdigit, str(ch or "")))
    if len(d) >= 44:
        return d[:44]
    return None


def _chaves_lista_do_df(df):
    """Chaves NFe 44 dígitos únicas, por ordem de aparição."""
    if df is None or df.empty or "Chave" not in df.columns:
        return []
    out = []
    for v in df["Chave"].tolist():
        k = _chave44_digitos(v)
        if k:
            out.append(k)
    return list(dict.fromkeys(out))


def _df_sig_hash_memo(df):
    if df is None or df.empty:
        return "empty"
    try:
        h = pd.util.hash_pandas_object(df.reset_index(drop=True), index=True)
        return hashlib.md5(h.values.tobytes()).hexdigest()
    except Exception:
        return hashlib.md5(
            f"{len(df)}|{list(df.columns)}".encode("utf-8", errors="replace")
        ).hexdigest()


def _excel_bytes_memo(prefix, df_work, sheet_name):
    """
    Gera bytes Excel só quando o DataFrame filtrado muda — evita recalcular a cada clique
    noutros widgets (menos RAM e menos tempo na zona de relatório / Etapa 3).
    """
    if df_work is None or df_work.empty:
        st.session_state.pop(f"_xlsx_mem_{prefix}", None)
        return None
    sig = _df_sig_hash_memo(df_work)
    sk = f"_xlsx_mem_{prefix}"
    prev = st.session_state.get(sk)
    if isinstance(prev, tuple) and prev[0] == sig:
        return prev[1]
    df_ex = (
        _df_com_data_emissao_dd_mm_yyyy(df_work.copy())
        if "Data Emissão" in df_work.columns
        else df_work.copy()
    )
    b = dataframe_para_excel_bytes(df_ex, sheet_name)
    if b is not None:
        st.session_state[sk] = (sig, b)
    else:
        st.session_state.pop(sk, None)
    return b


def _relatorio_leitura_tabela_aggrid(df_raw: pd.DataFrame, grid_key: str, height: int = 420):
    """
    Grelha Ag-Grid: filtro e ordenação no cabeçalho de cada coluna (comportamento próximo do Excel).
    Devolve o DataFrame original (`df_raw`) restrito às linhas visíveis após filtro/ordenação,
    para Excel e ZIP XML.
    """
    if df_raw is None or df_raw.empty:
        return df_raw

    try:
        from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode
    except ImportError:
        st.error(
            "Instale **streamlit-aggrid** no mesmo Python do Streamlit: "
            "`pip install streamlit-aggrid`"
        )
        st.dataframe(
            _df_relatorio_leitura_abas_para_exibicao_sem_sep_milhar(df_raw),
            width="stretch",
            hide_index=True,
            height=height,
        )
        return df_raw

    df_base = df_raw.reset_index(drop=True).copy()
    df_show = _df_relatorio_leitura_abas_para_exibicao_sem_sep_milhar(df_base.copy())
    # Tabela nativa sempre visível (nºs de nota, série, etc.) — o AgGrid por vezes não desenha com Streamlit novo.
    st.markdown("**Lista com numeração** — ordenação nas colunas; deslize para ver todas as linhas.")
    st.dataframe(
        df_show,
        width="stretch",
        hide_index=True,
        height=min(max(260, height), 520),
        key=f"tbl_num_{grid_key}",
    )
    st.caption("**Grelha com filtros** (tipo Excel) na secção abaixo — o Excel/ZIP usam as linhas visíveis na grelha.")

    df_grid = df_show.copy()
    df_grid.insert(0, "__garim_idx", range(len(df_grid)))

    gb = GridOptionsBuilder.from_dataframe(
        df_grid,
        editable=False,
        filter=True,
        sortable=True,
        resizable=True,
    )
    gb.configure_column(
        "__garim_idx",
        header_name="",
        hide=True,
        filter=False,
        sortable=False,
        resizable=False,
    )
    gb.configure_grid_options(rowHeight=28, headerHeight=36)
    grid_options = gb.build()
    # from_dataframe força fitGridWidth e esmaga colunas; removemos para permitir scroll horizontal.
    grid_options.pop("autoSizeStrategy", None)
    _dc = dict(grid_options.get("defaultColDef") or {})
    _dc.setdefault("minWidth", 140)
    grid_options["defaultColDef"] = _dc
    grid_options["suppressHorizontalScroll"] = False
    _loc = _aggrid_locale_pt_br()
    if _loc:
        grid_options["localeText"] = _loc

    resp = AgGrid(
        df_grid,
        gridOptions=grid_options,
        height=height,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        theme="streamlit",
        key=grid_key,
        show_download_button=False,
        show_search=False,
        enable_enterprise_modules=False,
        update_on=["filterChanged", "sortChanged"],
    )

    df_vis = resp.data
    if df_vis is None:
        return df_base
    if df_vis.empty:
        return df_base.iloc[0:0].copy()
    if "__garim_idx" not in df_vis.columns:
        return df_base
    try:
        idx = (
            pd.to_numeric(df_vis["__garim_idx"], errors="coerce")
            .dropna()
            .astype(int)
            .tolist()
        )
    except (TypeError, ValueError, KeyError):
        return df_base
    if not idx:
        return df_base.iloc[0:0].copy()
    _n = len(df_base)
    idx_ok = []
    _seen = set()
    for _i in idx:
        try:
            _j = int(_i)
        except (TypeError, ValueError):
            continue
        if 0 <= _j < _n and _j not in _seen:
            _seen.add(_j)
            idx_ok.append(_j)
    if not idx_ok:
        return df_base.iloc[0:0].copy()
    return df_base.iloc[idx_ok].reset_index(drop=True)


def _painel_zip_xml_filtrado(prefix, df_filtrado, cnpj_limpo, df_geral_full):
    """ZIP com XMLs das chaves visíveis após filtros (lê de temp_garimpo_uploads)."""
    cur_sig = _df_sig_hash_memo(df_filtrado)
    kzip = f"{prefix}_zip_parts_ready"
    ksig = f"{prefix}_zip_src_sig"
    if st.session_state.get(ksig) != cur_sig:
        st.session_state.pop(kzip, None)
    st.session_state[ksig] = cur_sig
    chaves = _chaves_lista_do_df(df_filtrado)
    if not chaves:
        st.caption(
            "ZIP com XML: use abas com coluna **Chave** (44 dígitos). Em **Buracos** só há Excel."
        )
        return
    if st.button(
        "Gerar ZIP com XML (linhas filtradas)",
        key=f"{prefix}_btn_zip",
        width="stretch",
    ):
        with st.spinner("A montar ZIP a partir do lote (memória ou pasta)…"):
            parts, tot = escrever_zip_dominio_por_chaves(
                cnpj_limpo, chaves, df_geral_full
            )
        st.session_state[kzip] = parts
        if tot == 0:
            st.warning(
                "Nenhum XML do lote correspondeu a estas chaves (lote vazio ou chaves externas ao lote)."
            )
        else:
            st.success(
                f"Incluídos **{tot}** XML(s) em **{len(parts)}** parte(s). Descarregue abaixo."
            )
    for idx, part in enumerate(st.session_state.get(kzip) or []):
        if os.path.isfile(part):
            with open(part, "rb") as fp:
                st.download_button(
                    rotulo_download_zip_parte(part),
                    fp.read(),
                    file_name=os.path.basename(part),
                    key=f"{prefix}_dlz_{idx}_{hashlib.md5(part.encode()).hexdigest()[:8]}",
                    width="stretch",
                )


def _excel_bytes_lista_especifica(df_geral, chaves_ordem_unicas):
    """
    Excel com número, série, chave, status, etc., alinhado ao relatório geral.
    chaves_ordem_unicas: ordem estável das chaves 44 dígitos neste lote ZIP.
    """
    if not chaves_ordem_unicas:
        return None
    cols_pref = [
        "Modelo",
        "Série",
        "Nota",
        "Chave",
        "Status Final",
        "Origem",
        "Operação",
        "Data Emissão",
        "CNPJ Emitente",
        "Nome Emitente",
        "Valor",
        "Ano",
        "Mes",
    ]
    por_chave = {}
    if df_geral is not None and not df_geral.empty and "Chave" in df_geral.columns:
        dfc = df_geral.copy()
        dfc["_k44"] = dfc["Chave"].map(_chave44_digitos)
        dfc = dfc[dfc["_k44"].notna()]
        dfc = dfc.drop_duplicates(subset=["_k44"], keep="first")
        for _, row in dfc.iterrows():
            k44 = row["_k44"]
            por_chave[k44] = row.drop(labels=["_k44"], errors="ignore")

    rows = []
    vistos = set()
    for ch in chaves_ordem_unicas:
        if not ch or ch in vistos:
            continue
        vistos.add(ch)
        if ch in por_chave:
            rows.append(por_chave[ch].to_dict())
        else:
            rows.append({"Chave": ch})

    out = pd.DataFrame(rows)
    cols = [c for c in cols_pref if c in out.columns] + [
        c for c in out.columns if c not in cols_pref
    ]
    out = out[[c for c in cols if c in out.columns]]
    out = _df_com_data_emissao_dd_mm_yyyy(out)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        out.to_excel(writer, sheet_name="Lista", index=False)
    return buf.getvalue()


def _zip_anexar_excel_lista_especifica(zf, df_geral, chaves_ordem, idx_parte):
    if not chaves_ordem:
        return
    xb = _excel_bytes_lista_especifica(df_geral, chaves_ordem)
    if xb:
        zf.writestr(
            f"RELATORIO_GARIMPEIRO/lista_especifica_pt{idx_parte:03d}.xlsx",
            xb,
        )


def escrever_zip_dominio_por_chaves(cnpj_limpo, chaves_lista, df_geral=None):
    """Gera um ou mais ZIPs (máx. MAX_XML_PER_ZIP XMLs cada); XMLs na raiz; Excel do lote em RELATORIO_GARIMPEIRO/. Retorna (lista_caminhos, total_xml)."""
    if not chaves_lista or not _garimpo_existem_fontes_xml_lote():
        return [], 0
    ch_set = set()
    for c in chaves_lista:
        k = _chave44_digitos(c)
        if k:
            ch_set.add(k)
    if not ch_set:
        return [], 0

    parts = []
    part_idx = 1
    count_xml = 0
    no_lote = 0
    nome = f"faltantes_dominio_final_pt{part_idx}.zip"
    zf = _zipfile_open_write_export(nome)
    parts.append(nome)
    usados_nomes_parte = set()
    chaves_excel_ordem = []
    vistos_chave_excel = set()
    ch44_ja_gravado = set()

    try:
        for fn in _lista_nomes_fontes_xml_garimpo():
            with _abrir_fonte_xml_garimpo_stream(fn) as ft:
                for name, data in extrair_recursivo(ft, fn):
                    res, _ = identify_xml_info(data, cnpj_limpo, name)
                    ch44 = _chave44_digitos(res.get("Chave")) if res else None
                    if res and ch44 and ch44 in ch_set:
                        td = _tupla_dedupe_export_xml(res, ch44)
                        if td is None or td in ch44_ja_gravado:
                            continue
                        ch44_ja_gravado.add(td)
                        if no_lote >= MAX_XML_PER_ZIP:
                            _zip_anexar_excel_lista_especifica(
                                zf, df_geral, chaves_excel_ordem, part_idx
                            )
                            zf.close()
                            part_idx += 1
                            nome = f"faltantes_dominio_final_pt{part_idx}.zip"
                            zf = _zipfile_open_write_export(nome)
                            parts.append(nome)
                            no_lote = 0
                            usados_nomes_parte = set()
                            chaves_excel_ordem = []
                            vistos_chave_excel = set()
                        arc = _nome_xml_raiz_zip_unico(usados_nomes_parte, name)
                        zf.writestr(arc, data)
                        if ch44 not in vistos_chave_excel:
                            vistos_chave_excel.add(ch44)
                            chaves_excel_ordem.append(ch44)
                        count_xml += 1
                        no_lote += 1
        if count_xml > 0:
            _zip_anexar_excel_lista_especifica(
                zf, df_geral, chaves_excel_ordem, part_idx
            )
    finally:
        try:
            zf.close()
        except Exception:
            pass

    if count_xml == 0:
        return [], 0

    return parts, count_xml


def _zip_bytes_from_arc_pairs(pairs: list) -> bytes:
    """Um único .zip em memória a partir de [(nome_interno, bytes), …]."""
    if not pairs:
        return b""
    buf = io.BytesIO()
    with zipfile.ZipFile(
        buf,
        "w",
        zipfile.ZIP_DEFLATED,
        compresslevel=ZIP_EXPORT_COMPRESSLEVEL,
    ) as zf:
        for arc, raw in pairs:
            zf.writestr(arc, raw if isinstance(raw, (bytes, bytearray)) else bytes(raw))
    return buf.getvalue()


def _split_arc_pairs_into_zips_max_bytes(pairs: list, max_bytes: int) -> list:
    """
    Divide a lista de (arcname, xml_bytes) em vários ZIPs, cada um ≤ max_bytes.
    Recursão por metades se um ZIP ainda exceder o limite (ficheiro único enorme fica num ZIP só).
    """
    if not pairs:
        return []
    blob = _zip_bytes_from_arc_pairs(pairs)
    if len(blob) <= max_bytes or len(pairs) == 1:
        return [blob]
    mid = max(1, len(pairs) // 2)
    return _split_arc_pairs_into_zips_max_bytes(
        pairs[:mid], max_bytes
    ) + _split_arc_pairs_into_zips_max_bytes(pairs[mid:], max_bytes)


def _coletar_xml_serie4_emitidas_propria(cnpj_limpo: str) -> list:
    """
    NF-e / NFC-e, emissão própria, status NORMAIS, série 4. Deduplicação igual à exportação ZIP.
    Ordem estável por nome interno.
    """
    out = []
    seen = set()
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    if len(cnpj) != 14:
        return []
    for f_name in _lista_nomes_fontes_xml_garimpo():
        with _abrir_fonte_xml_garimpo_stream(f_name) as f_temp:
            for name, xml_data in extrair_recursivo(f_temp, f_name):
                res, is_p = identify_xml_info(xml_data, cnpj, name)
                if not res or not is_p:
                    del xml_data
                    continue
                if str(res.get("Status") or "").strip().upper() != "NORMAIS":
                    del xml_data
                    continue
                tp = str(res.get("Tipo") or "").strip()
                if tp not in ("NF-e", "NFC-e"):
                    del xml_data
                    continue
                ser = _normaliza_serie_filtro(res.get("Série"))
                if ser != "4":
                    del xml_data
                    continue
                ck = _chave_para_conjunto_export(res.get("Chave"))
                td = _tupla_dedupe_export_xml(res, ck)
                if td is None or td in seen:
                    del xml_data
                    continue
                seen.add(td)
                arc = _nome_arquivo_xml_contabilidade(res, name)
                out.append((arc, bytes(xml_data)))
                del xml_data
    out.sort(key=lambda x: x[0])
    return out


def _gerar_lista_zips_serie4_emitidas_50mb(cnpj_limpo: str) -> tuple:
    """
    Devolve (lista de (nome_ficheiro, bytes_zip), mensagem_erro_ou_None).
    Só para CNPJ_CLIENTE_EXPORT_SERIE4_ZIP50MB.
    """
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    if cnpj != CNPJ_CLIENTE_EXPORT_SERIE4_ZIP50MB:
        return [], "Esta ferramenta não está disponível para este CNPJ."
    if not _garimpo_existem_fontes_xml_lote():
        return [], "Sem ficheiros no lote — faça o garimpo ou **Incluir mais XML** primeiro."
    pairs = _coletar_xml_serie4_emitidas_propria(cnpj)
    if not pairs:
        return (
            [],
            "Nenhum XML encontrado: **série 4**, **NF-e / NFC-e**, **autorizadas** (NORMAIS), **emissão própria**.",
        )
    blobs = _split_arc_pairs_into_zips_max_bytes(pairs, ZIP50_SERIE4_MAX_BYTES)
    nomes = []
    for i, blob in enumerate(blobs, start=1):
        nomes.append(
            (f"Garim_serie4_emitidas_pt{i:03d}.zip", blob)
        )
    return nomes, None


def _intervalo_mes_relatorio(ano, mes):
    try:
        a, m = int(ano), int(mes)
        if a < 1900 or not (1 <= m <= 12):
            return None, None
        d1 = date(a, m, 1)
        d2 = date(a, m, monthrange(a, m)[1])
        return d1, d2
    except (TypeError, ValueError):
        return None, None


def _data_emissao_linha(row):
    de = row.get("Data Emissão")
    if de is not None and not (isinstance(de, float) and pd.isna(de)):
        s = str(de).strip()[:10]
        if len(s) >= 10 and s[4] == "-" and s[7] == "-":
            try:
                return date.fromisoformat(s[:10])
            except ValueError:
                pass
    return None


def _linha_no_periodo(row, d_ini, d_fim):
    d0 = _data_emissao_linha(row)
    if d0 is not None:
        return d_ini <= d0 <= d_fim
    lo, hi = _intervalo_mes_relatorio(row.get("Ano"), row.get("Mes"))
    if lo is None:
        return False
    return not (hi < d_ini or lo > d_fim)


def _chave44_de_linha(row):
    ch = row.get("Chave")
    if ch is None or (isinstance(ch, float) and pd.isna(ch)):
        return None
    s = "".join(filter(str.isdigit, str(ch)))
    if len(s) >= 44:
        return s[:44]
    return None


def _chave_para_conjunto_export(ch):
    """
    Normaliza Chave para cruzar df_geral com identify_xml_info ao gerar ZIPs (Etapa 3).
    Evita falhas por float/.0 na coluna, notação científica parcial ou espaços — casos em que
    `res['Chave'] in set(df['Chave'])` nunca era verdadeiro.
    """
    if ch is None:
        return None
    try:
        if pd.isna(ch):
            return None
    except (TypeError, ValueError):
        pass
    t = str(ch).strip()
    if not t or t.lower() == "nan":
        return None
    if t.startswith("INUT_"):
        return t
    d = "".join(filter(str.isdigit, t))
    if len(d) >= 44:
        return d[:44]
    return None


def _v2_aplicar_nota_especifica_propria(
    out_p: pd.DataFrame,
    chave_raw,
    numero_esp,
    serie_raw,
    *,
    skip: bool = False,
) -> pd.DataFrame:
    """Restringe emissão própria a uma linha: chave (44 dígitos ou INUT_…) ou par n.º + série."""
    if skip or out_p is None or out_p.empty or "Chave" not in out_p.columns:
        return out_p
    ch_raw = str(chave_raw or "").strip()
    ch_u = ch_raw.upper()
    if ch_u.startswith("INUT_"):
        return out_p.loc[out_p["Chave"].astype(str).str.strip().str.upper() == ch_u]
    ch_d = "".join(c for c in ch_raw if c.isdigit())
    if len(ch_d) >= 44:
        ch_d = ch_d[:44]
        ck_norm = out_p["Chave"].map(_chave_para_conjunto_export)
        return out_p.loc[ck_norm == ch_d]
    try:
        n = int(numero_esp)
    except (TypeError, ValueError):
        n = 0
    ser = str(serie_raw or "").strip()
    if ser and n > 0 and "Nota" in out_p.columns and "Série" in out_p.columns:
        nn = pd.to_numeric(out_p["Nota"], errors="coerce")
        ser_col = out_p["Série"].astype(str).str.strip()
        mask = (nn == n) & (ser_col == ser)
        if not mask.any() and ser.isdigit():
            mask = (nn == n) & (pd.to_numeric(ser_col, errors="coerce") == int(ser))
        return out_p.loc[mask]
    return out_p


def _nota_int_linha(row):
    n = row.get("Nota")
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return None
    try:
        return int(n)
    except (ValueError, TypeError):
        try:
            return int(float(n))
        except (ValueError, TypeError):
            return None


# Código <mod> da Sefaz → rótulo usado na coluna Modelo do relatório geral
_MODELO_SEFAZ_PARA_RELATORIO = {
    "55": "NF-e",
    "65": "NFC-e",
    "57": "CT-e",
    "58": "MDF-e",
    "67": "CT-e OS",
}


def _normaliza_modelo_filtro(modelo):
    """Aceita NF-e, NFC-e… ou 55, 65… (código Sefaz) para cruzar com df_geral."""
    if modelo is None:
        return ""
    s = str(modelo).strip()
    if not s:
        return ""
    if s in _MODELO_SEFAZ_PARA_RELATORIO:
        return _MODELO_SEFAZ_PARA_RELATORIO[s]
    try:
        k = str(int(float(s.replace(",", "."))))
        if k in _MODELO_SEFAZ_PARA_RELATORIO:
            return _MODELO_SEFAZ_PARA_RELATORIO[k]
    except (ValueError, TypeError):
        pass
    return s


def _normaliza_serie_filtro(serie):
    """Alinha com a chave: série na app costuma vir sem zeros à esquerda (ex. 1 em vez de 001)."""
    if serie is None or (isinstance(serie, float) and pd.isna(serie)):
        return ""
    if isinstance(serie, (int, float)) and not isinstance(serie, bool):
        try:
            f = float(serie)
            if f.is_integer():
                return str(int(f))
        except (ValueError, OverflowError):
            pass
    t = str(serie).strip()
    if not t:
        return ""
    try:
        f = float(t.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return t


def _modelo_serie_coincidem(row, modelo, serie):
    m = row.get("Modelo")
    s = row.get("Série")
    if m is None or s is None:
        return False
    mn = _normaliza_modelo_filtro(modelo)
    sn = _normaliza_serie_filtro(serie)
    return (
        str(m).strip() == mn
        and _normaliza_serie_filtro(s) == sn
    )


def chaves_por_periodo_data(df_geral, d_ini, d_fim):
    if df_geral is None or df_geral.empty:
        return []
    df = df_geral
    out = []
    for _, row in df.iterrows():
        if _linha_no_periodo(row, d_ini, d_fim):
            ch = _chave44_de_linha(row)
            if ch:
                out.append(ch)
    return list(dict.fromkeys(out))


def chaves_por_faixa_numeracao(df_geral, modelo, serie, n_ini, n_fim):
    if df_geral is None or df_geral.empty:
        return []
    df = df_geral
    out = []
    for _, row in df.iterrows():
        if not _modelo_serie_coincidem(row, modelo, serie):
            continue
        ni = _nota_int_linha(row)
        if ni is None or not (n_ini <= ni <= n_fim):
            continue
        ch = _chave44_de_linha(row)
        if ch:
            out.append(ch)
    return list(dict.fromkeys(out))


def chaves_por_nota_serie(df_geral, modelo, serie, nota):
    if df_geral is None or df_geral.empty:
        return []
    df = df_geral
    out = []
    for _, row in df.iterrows():
        if not _modelo_serie_coincidem(row, modelo, serie):
            continue
        ni = _nota_int_linha(row)
        if ni is None or ni != nota:
            continue
        ch = _chave44_de_linha(row)
        if ch:
            out.append(ch)
    return list(dict.fromkeys(out))


# Texto espelhado na área “copiar guia” (alinhar ao fluxo real da app)
TEXTO_GUIA_GARIMPEIRO = """
Garimpeiro — Manual em texto simples (para copiar)

=== O QUE É LIDO E O QUE RECEBE ===
Dados extraídos dos XML (e cruzamentos opcionais):
• Identificação do documento: chave de 44 dígitos, modelo/tipo, série, número, datas de emissão, valores, operação, UF, emitente e destinatário quando constam no ficheiro.
• Classificação por linha: emissão própria (seu CNPJ como emitente) vs. terceiros (documentos recebidos).
• Agregados: resumo por série (totais), lista geral, canceladas, inutilizadas, autorizadas, buracos na numeração (com ou sem referência “último nº + mês” na lateral).
• Um mesmo número de nota pode ter vários ficheiros XML (ex.: NF-e + evento) — várias linhas com a mesma chave ou lógica equivalente.

Ficheiros e descargas que pode obter:
• Nesta página: tabelas e indicadores (KPI) sobre o lote atual.
• Excel: relatório geral completo; em Etapa 3 — “só Excel” (lote inteiro ou só linhas filtradas); em “lista específica” — planilhas por chaves, faixa, período, etc. (ver botões dessa zona).
• ZIP (Etapa 3): até 10 000 XML por parte; modos “tudo” (filtros não cortam XML) ou “filtrado” (só XML das linhas filtradas); estrutura em raiz ou por pastas. Em cada parte: pasta RELATORIO_GARIMPEIRO/ com Excel desse envio.
• PDF: resumo visual do painel (se a biblioteca fpdf2 estiver disponível — mensagem na app se faltar).
• Modelo Excel para inutilizadas declaradas manualmente (quando usar essa funcionalidade).

=== MANUAL PASSO A PASSO ===
1. Lateral: CNPJ do emitente (cliente) — só os 14 dígitos ou cole já mascarado; clique em Liberar operação.
2. Envie ZIP ou XML soltos (grandes volumes são suportados).
3. Iniciar grande garimpo e aguardar o processamento.
4. Depois do primeiro resultado, pode acrescentar mais XML/ZIP no topo da página sem reiniciar o garimpo.
5. (Opcional) Lateral “Último nº por série”: afeta só o cálculo de buracos — **só as séries que preencher e guardar**; âncora por último nº e mês. O garimpo e o resumo por série continuam totais. Sem Guardar referência válida, buracos usam todo o intervalo lido em emissão própria.
6. Inutilizadas: a partir dos buracos, por planilha (Excel/CSV) ou faixa — só números que já forem buraco listado (não alarga intervalos).
7. Painel à direita: **Processar Dados** grava XML/ZIP extra, aplica inutilizações «sem XML» (se configurou) e recalcula a partir da pasta de uploads.
8. Etapa 3: filtros em cascata (emissão própria e terceiros em colunas separadas). Escolha um dos seis modos: ZIP tudo (raiz ou pastas), ZIP filtrado (pastas ou tudo na raiz), Excel só lote completo, Excel só filtrado — e gere as partes quando aplicável.
9. Lista específica (secção própria): exporte subconjuntos por chaves, faixa, período, série ou uma nota — em Excel e/ou ZIP conforme os botões apresentados.

=== DICAS ===
• Resetar sistema: limpa sessão e temporários ao mudar de cliente ou recomeçar.
• Nos filtros, lista vazia = esse critério não restringe. Opções inválidas após mudar outro filtro são limpas automaticamente.
• Nomes de modelos na app: NF-e, NFS-e, NFC-e, CT-e, CT-e OS (mod. 67), MDF-e, Outros (cartas de correção não são lidas).
""".strip()


# --- INTERFACE ---
st.markdown(f"<h1>{_garim_emoji('⛏️')} Garimpeiro</h1>", unsafe_allow_html=True)

with st.container():
    with st.expander(
        "📘 Manual do Garimpeiro — clique para abrir",
        expanded=False,
    ):
        st.caption(
            "Este painel funciona como um manual embutido: abra quando quiser consultar; pode voltar a fechar a qualquer momento."
        )
        st.markdown(
            '<h3 class="garim-sec">Resumo: que dados são tratados e o que obtém</h3>',
            unsafe_allow_html=True,
        )
        st.markdown(
            """
        <div class="instrucoes-card manual-compacto">
            <p style="margin:0 0 10px 0;font-size:0.95rem;line-height:1.55;color:#333;">
            <b>Dados que o sistema extrai ou cruza</b> (a partir dos XML e, se usar, de planilhas Sefaz / inutilização manual):
            </p>
            <ul style="margin:0;padding-left:1.25rem;line-height:1.55;font-size:0.93rem;color:#333;">
                <li><b>Documento:</b> chave de 44 dígitos, tipo/modelo, série, número, datas, valores, operação, UF, emitente e destinatário quando existirem no XML.</li>
                <li><b>Origem:</b> linhas de <b>emissão própria</b> (seu CNPJ como emitente) e de <b>terceiros</b> (documentos recebidos), com totais por tipo.</li>
                <li><b>Estado fiscal interpretado:</b> autorizadas, canceladas e inutilizadas a partir dos <b>XML</b>; <b>buracos</b> na numeração (com ou sem referência “último nº + mês” na lateral).</li>
                <li><b>Varios XML por mesma nota:</b> por exemplo nota + evento — pode haver mais do que um ficheiro por chave ou lógica equivalente.</li>
            </ul>
            <p style="margin:14px 0 8px 0;font-size:0.95rem;line-height:1.55;color:#333;">
            <b>O que recebe na prática</b> (nesta página e em ficheiros):
            </p>
            <ul style="margin:0;padding-left:1.25rem;line-height:1.55;font-size:0.93rem;color:#333;">
                <li><b>Nesta página:</b> painel com tabelas, filtros e indicadores sobre o lote atual.</li>
                <li><b>Excel:</b> relatório geral completo; na Etapa 3 — exportação “só Excel” (lote inteiro ou só o filtrado); na lista específica — planilhas por critério (chaves, faixa, período, etc.).</li>
                <li><b>ZIP:</b> até 10&nbsp;000 XML por parte — modo “tudo” (filtros não cortam ficheiros) ou “filtrado” (só XML das linhas filtradas), em raiz ou pastas; cada parte inclui <code>RELATORIO_GARIMPEIRO/</code> com Excel.</li>
                <li><b>PDF:</b> resumo do dashboard para arquivo ou impressão (se <code>fpdf2</code> estiver instalado).</li>
                <li><b>Modelo</b> para inutilizadas declaradas manualmente, quando usar essa opção.</li>
            </ul>
        </div>
        """,
            unsafe_allow_html=True,
        )
        st.markdown(
            '<h3 class="garim-sec">Como usar — passo a passo</h3>',
            unsafe_allow_html=True,
        )
        st.markdown(
            """
        <div class="instrucoes-card manual-compacto">
            <ol style="margin:0;padding-left:1.25rem;line-height:1.6;font-size:0.93rem;color:#333;">
                <li style="margin-bottom:8px;"><b>CNPJ na lateral:</b> introduza <b>só os 14 dígitos</b> do emitente (cliente) ou cole já mascarado; clique em <b>Liberar operação</b>.</li>
                <li style="margin-bottom:8px;"><b>Lote:</b> envie ficheiros <b>ZIP</b> ou <b>XML</b> soltos — volumes grandes são suportados.</li>
                <li style="margin-bottom:8px;"><b>Garimpo:</b> <b>Iniciar grande garimpo</b> e aguarde até aparecerem os resultados.</li>
                <li style="margin-bottom:8px;"><b>Mais ficheiros:</b> no <b>topo da área de resultados</b>, pode acrescentar ZIP/XML <b>sem reiniciar</b> o processamento anterior.</li>
                <li style="margin-bottom:8px;"><b>(Opcional)</b> <b>Último nº por série</b> (lateral): altera apenas o cálculo de <b>buracos</b> — <b>só para as séries que indicar e guardar</b> (âncora por último número e mês). O garimpo e o resumo por série mantêm-se <b>totais</b>. Sem <b>Guardar referência</b> com linhas válidas, os buracos consideram toda a numeração lida em emissão própria.</li>
                <li style="margin-bottom:8px;"><b>Processar Dados</b> (painel à direita): grava uploads extra, aplica inutilizações configuradas e recalcula o relatório a partir da pasta em disco.</li>
                <li style="margin-bottom:8px;"><b>Inutilizadas:</b> a partir dos <b>buracos</b>, use <b>planilha</b> (Excel/CSV) ou <b>faixa</b> — só entram números que já forem buraco listado (não alarga intervalos).</li>
                <li style="margin-bottom:8px;"><b>Etapa 3 — filtros e exportação:</b> refine por emissão própria e/ou terceiros; escolha um dos <b>seis</b> modos (ZIP tudo raiz/pastas, ZIP filtrado pastas/raiz, Excel lote completo, Excel só filtrado) e gere as partes quando a app repartir o lote.</li>
                <li style="margin-bottom:0;"><b>Lista específica:</b> use a secção dedicada para extrair subconjuntos por chaves, faixa, período, série ou uma nota — nos formatos indicados pelos botões dessa zona.</li>
            </ol>
        </div>
        """,
            unsafe_allow_html=True,
        )
        st.markdown(
            """
        <div style="font-size:0.88rem;line-height:1.5;color:#555;margin:6px 0 12px 0;padding:10px 12px;background:rgba(255,240,248,0.6);border-radius:10px;border-left:3px solid #FF69B4;">
        <b>Dicas rápidas:</b> <b>Resetar sistema</b> limpa sessão e temporários ao mudar de cliente.
        Nos filtros, lista vazia = esse critério não aplica. Modelos usados na app incluem NF-e, NFS-e, NFC-e, CT-e, CT-e OS, MDF-e e Outros (cartas de correção ignoradas).
        </div>
        """,
            unsafe_allow_html=True,
        )
        # Streamlit não permite expander dentro de expander — usar checkbox para recolher o texto.
        if st.checkbox(
            "📋 Mostrar texto do manual para copiar (Ctrl+A, Ctrl+C)",
            value=False,
            key="garimpeiro_mostrar_guia_txt",
        ):
            st.caption("Clique na caixa, Ctrl+A (Cmd+A no Mac) e Ctrl+C para copiar tudo.")
            st.text_area(
                "Guia",
                value=TEXTO_GUIA_GARIMPEIRO,
                height=340,
                key="garimpeiro_guia_copiar_v2",
                label_visibility="collapsed",
            )

st.markdown("---")

keys_to_init = [
    'garimpo_ok', 
    'confirmado', 
    'relatorio', 
    'df_resumo', 
    'df_faltantes', 
    'df_canceladas', 
    'df_inutilizadas', 
    'df_autorizadas',
    'df_denegadas',
    'df_rejeitadas',
    'df_geral', 
    'df_divergencias', 
    'st_counts', 
    'validation_done', 
    'export_ready',
    'org_zip_parts',
    'todos_zip_parts',
    'ch_falt_dom',
    'zip_dom_parts',
]

for k in keys_to_init:
    if k not in st.session_state:
        if 'df' in k: 
            st.session_state[k] = pd.DataFrame()
        elif k in ['relatorio', 'org_zip_parts', 'todos_zip_parts', 'ch_falt_dom', 'zip_dom_parts']: 
            st.session_state[k] = []
        elif k == 'st_counts': 
            st.session_state[k] = {
                "CANCELADOS": 0,
                "INUTILIZADOS": 0,
                "AUTORIZADAS": 0,
                "DENEGADOS": 0,
                "REJEITADOS": 0,
            }
        else: 
            st.session_state[k] = False

if "excel_buffer" not in st.session_state:
    st.session_state["excel_buffer"] = None
if "export_excel_name" not in st.session_state:
    st.session_state["export_excel_name"] = "relatorio.xlsx"
if "seq_ref_ultimos" not in st.session_state:
    st.session_state["seq_ref_ultimos"] = None
if "seq_ref_ano" not in st.session_state:
    st.session_state["seq_ref_ano"] = None
if "seq_ref_mes" not in st.session_state:
    st.session_state["seq_ref_mes"] = None
if "seq_ref_rows" not in st.session_state:
    if st.session_state.get("seq_ref_ultimos"):
        st.session_state["seq_ref_rows"] = ultimos_dict_para_dataframe(st.session_state["seq_ref_ultimos"])
    else:
        st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(
            pd.DataFrame([{"Modelo": "NF-e", "Série": "1", "Último número": ""}])
        )
if "seq_struct_v" not in st.session_state:
    st.session_state["seq_struct_v"] = 0
if "cnpj_widget" not in st.session_state:
    st.session_state["cnpj_widget"] = ""
if SESSION_KEY_SERIE4_ZIP50_PARTS not in st.session_state:
    st.session_state[SESSION_KEY_SERIE4_ZIP50_PARTS] = None

with st.sidebar:
    st.markdown(
        f'<h4>{_garim_emoji("🔍")} Configuração</h4>',
        unsafe_allow_html=True,
    )
    # Normalizar máscara *antes* do text_input: depois do widget o Streamlit bloqueia
    # `session_state["cnpj_widget"] = ...` na mesma execução (StreamlitAPIException).
    _cnpj_key = "cnpj_widget"
    _cnpj_raw = st.session_state.get(_cnpj_key, "")
    cnpj_limpo = "".join(c for c in str(_cnpj_raw) if c.isdigit())[:14]
    _cnpj_fmt = format_cnpj_visual(cnpj_limpo)
    if _cnpj_fmt != str(_cnpj_raw):
        st.session_state[_cnpj_key] = _cnpj_fmt
    st.text_input(
        "CNPJ do cliente",
        key=_cnpj_key,
        placeholder="Somente números",
        help="14 dígitos; pode colar com ou sem máscara. A formatação é aplicada ao digitar.",
    )

    if cnpj_limpo and len(cnpj_limpo) != 14:
        st.error("⚠️ CNPJ Inválido.")
        
    if len(cnpj_limpo) == 14:
        if st.button("✅ Liberar operação"): 
            st.session_state['confirmado'] = True

        if cnpj_limpo == CNPJ_CLIENTE_EXPORT_SERIE4_ZIP50MB:
            with st.expander("📦 ZIP série 4 — emitidas (~50 MB)", expanded=False):
                st.caption(
                    "Gera ZIPs só com **NF-e / NFC-e autorizadas** (NORMAIS), **série 4**, **emissão própria**, "
                    f"a partir do lote atual. Cada ficheiro tem no máximo **~{ZIP50_SERIE4_MAX_BYTES // (1024 * 1024)} MB**."
                )
                if not _garimpo_existem_fontes_xml_lote():
                    st.info("Sem lote XML — carregue ficheiros e processe o garimpo.")
                else:
                    if st.button(
                        "Gerar ZIPs (série 4, ~50 MB cada)",
                        key="btn_serie4_zip50_gerar",
                        width="stretch",
                    ):
                        _parts, _err = _gerar_lista_zips_serie4_emitidas_50mb(cnpj_limpo)
                        if _err:
                            st.session_state[SESSION_KEY_SERIE4_ZIP50_PARTS] = None
                            st.error(_err)
                        else:
                            st.session_state[SESSION_KEY_SERIE4_ZIP50_PARTS] = _parts
                            st.success(
                                f"**{len(_parts)}** ZIP(s) gerado(s). Use os botões abaixo para descarregar."
                            )
                            st.rerun()
                    _zp = st.session_state.get(SESSION_KEY_SERIE4_ZIP50_PARTS)
                    if _zp:
                        st.caption(f"{len(_zp)} ficheiro(s) na última geração.")
                        for _zi, (_zfn, _zblob) in enumerate(_zp):
                            st.download_button(
                                f"⬇️ {_zfn}",
                                data=_zblob,
                                file_name=_zfn,
                                mime="application/zip",
                                key=f"dl_serie4zip50_{_zi}_{hashlib.md5(_zblob[: min(4096, len(_zblob))]).hexdigest()[:10]}",
                                width="stretch",
                            )
                        if st.button("Limpar lista de ZIPs", key="btn_serie4_zip50_limpar"):
                            st.session_state[SESSION_KEY_SERIE4_ZIP50_PARTS] = None
                            st.rerun()

        with st.expander("📌 Últimos nº / séries (mês ref.)", expanded=False):
            st.caption(
                "Mês de referência + por linha modelo, série e último nº. **Só essas séries** entram no cálculo de **buracos** (a partir desse último nº); as outras séries do lote ignoram-se nos buracos."
            )
            d = date.today()
            def_ano = d.year - 1 if d.month == 1 else d.year
            def_mes = 12 if d.month == 1 else d.month - 1
            a0 = st.session_state["seq_ref_ano"] if st.session_state.get("seq_ref_ano") is not None else def_ano
            m0 = st.session_state["seq_ref_mes"] if st.session_state.get("seq_ref_mes") is not None else def_mes
            if st.session_state.get("garimpo_ok"):
                if st.button("Puxar séries do resumo", key="seq_btn_puxar", width="stretch"):
                    dfr = st.session_state.get("df_resumo")
                    if dfr is not None and not dfr.empty:
                        novas = []
                        for _, r in dfr.iterrows():
                            novas.append(
                                {
                                    "Modelo": r["Documento"],
                                    "Série": str(r["Série"]),
                                    "Último número": "",
                                }
                            )
                        st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(pd.DataFrame(novas))
                        st.session_state["seq_struct_v"] = int(st.session_state.get("seq_struct_v", 0)) + 1
                        st.success("Preencha **Últ. nº** em cada cartão e carregue em **Guardar referência**.")
                        st.rerun()
                    else:
                        st.warning("Resumo por série ainda vazio.")

            _opts = ["NF-e", "NFC-e", "CT-e", "CT-e OS"]
            _df_base = normalize_seq_ref_editor_df(st.session_state["seq_ref_rows"])
            _recs = (
                _df_base.to_dict("records")
                if not _df_base.empty
                else [{"Modelo": "NF-e", "Série": "1", "Último número": ""}]
            )
            n_rows = len(_recs)
            v = int(st.session_state.get("seq_struct_v", 0))

            ca, cm = st.columns(2)
            with ca:
                sr_ano = st.number_input(
                    "Ano",
                    min_value=2000,
                    max_value=2100,
                    value=int(a0),
                    key="seq_sidebar_ano",
                )
            with cm:
                sr_mes = st.number_input(
                    "Mês",
                    min_value=1,
                    max_value=12,
                    value=int(m0),
                    key="seq_sidebar_mes",
                )

            h0, h1, h2 = st.columns([1.25, 0.62, 0.95], gap="small")
            with h0:
                st.caption("Modelo")
            with h1:
                st.caption("Sér.")
            with h2:
                st.caption("Últ. nº")

            for i, row in enumerate(_recs):
                if i > 0:
                    st.markdown(
                        '<hr class="garim-seq-row-spacer" />',
                        unsafe_allow_html=True,
                    )
                modelo_raw = row.get("Modelo")
                if modelo_raw is None or pd.isna(modelo_raw):
                    modelo_cur = "NF-e"
                else:
                    modelo_cur = str(modelo_raw).strip()
                if not modelo_cur or modelo_cur.lower() == "nan":
                    modelo_cur = "NF-e"
                if modelo_cur not in _opts:
                    opts_row = _opts + [modelo_cur]
                    idx = len(_opts)
                else:
                    opts_row = _opts
                    idx = _opts.index(modelo_cur)
                ser_raw = row.get("Série")
                if ser_raw is None or (isinstance(ser_raw, float) and pd.isna(ser_raw)):
                    ser_cur = ""
                else:
                    ser_cur = str(ser_raw).strip()
                ult_raw = row.get("Último número")
                if ult_raw is None or (isinstance(ult_raw, float) and pd.isna(ult_raw)):
                    ult_cur = ""
                else:
                    ult_cur = str(ult_raw).strip()

                c_m, c_s, c_u = st.columns([1.25, 0.62, 0.95], gap="small")
                with c_m:
                    st.selectbox(
                        "Modelo",
                        opts_row,
                        index=idx,
                        key=f"sr_{v}_{i}_m",
                        label_visibility="collapsed",
                    )
                with c_s:
                    st.text_input(
                        "Série",
                        value=ser_cur,
                        key=f"sr_{v}_{i}_s",
                        label_visibility="collapsed",
                        max_chars=10,
                        placeholder="1",
                    )
                with c_u:
                    st.text_input(
                        "Último nº",
                        value=ult_cur,
                        key=f"sr_{v}_{i}_u",
                        label_visibility="collapsed",
                        max_chars=18,
                        placeholder="nº",
                    )

            b1, b2 = st.columns(2)
            with b1:
                if st.button("➕ Série", key="seq_add_row", width="stretch"):
                    cur_df = collect_seq_ref_from_widgets(v, n_rows)
                    novo = pd.DataFrame([{"Modelo": "NF-e", "Série": "", "Último número": ""}])
                    st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(
                        pd.concat([cur_df, novo], ignore_index=True)
                    )
                    st.session_state["seq_struct_v"] = v + 1
                    st.rerun()
            with b2:
                if n_rows > 1 and st.button("➖ Última", key="seq_rem_row", width="stretch"):
                    cur_df = collect_seq_ref_from_widgets(v, n_rows)
                    st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(cur_df.iloc[:-1])
                    st.session_state["seq_struct_v"] = v + 1
                    st.rerun()

            if st.button(
                "Guardar referência",
                type="primary",
                width="stretch",
                key="seq_btn_guardar",
                help="Grava ano, mês e séries na sessão.",
            ):
                cur_df = collect_seq_ref_from_widgets(v, n_rows)
                st.session_state["seq_ref_rows"] = cur_df
                st.session_state["seq_ref_ano"] = int(sr_ano)
                st.session_state["seq_ref_mes"] = int(sr_mes)
                parsed = ref_map_from_dataframe(cur_df)
                if parsed:
                    st.session_state["seq_ref_ultimos"] = parsed
                    st.success(f"{len(parsed)} série(s) guardada(s).")
                else:
                    st.warning(
                        "Preencha **documento**, **série** e **últ. nº** (> 0) em pelo menos um cartão e volte a guardar."
                    )
                if st.session_state.get("garimpo_ok") and st.session_state.get("relatorio"):
                    reconstruir_dataframes_relatorio_simples()

            if st.session_state.get("seq_ref_ultimos"):
                st.caption(
                    f"Referência ativa: **{st.session_state['seq_ref_ano']}/"
                    f"{int(st.session_state['seq_ref_mes']):02d}** — "
                    f"**{len(st.session_state['seq_ref_ultimos'])}** série(s)."
                )

        if st.session_state.get("garimpo_ok") and st.session_state.get("relatorio"):
            st.markdown("---")
            st.markdown(
                f'<h5>{_garim_emoji("📄")} PDF do dashboard</h5>',
                unsafe_allow_html=True,
            )
            st.caption(
                "Resumo do lote em PDF (métricas e amostras). Só descarrega o ficheiro — não altera o que vê nesta página."
            )
            _kpi_sb = coletar_kpis_dashboard()
            _cnpj_sb = format_cnpj_visual(cnpj_limpo) if len(cnpj_limpo) == 14 else ""
            _df_res_pdf = st.session_state.get("df_resumo")
            _pdf_sb = pdf_dashboard_garimpeiro_bytes(_kpi_sb, _cnpj_sb, _df_res_pdf)
            if _pdf_sb:
                st.download_button(
                    "⬇️ Baixar PDF do dashboard",
                    data=_pdf_sb,
                    file_name="dashboard_garimpeiro.pdf",
                    mime="application/pdf",
                    key="dl_dash_pdf_sidebar",
                    width="stretch",
                )
            else:
                st.markdown(_instrucoes_instalar_fpdf2_markdown())

    st.markdown("---")

    if st.button("🗑️ Resetar sistema"):
        limpar_arquivos_temp()
        st.session_state.clear()
        st.rerun()


def _garim_etapa3_corpo(cnpj_limpo):
    """Etapa 3 — filtros e exportação (isolado para st.fragment)."""
    st.caption(
        "Opcional: ajuste filtros · escolha o tipo · **Gerar só sua empresa**, só **terceiros**, ou **os dois** · descarregue em baixo."
    )
    if hasattr(st, "fragment"):
        st.caption(
            "Esta secção corre em **modo fragmento**: ao mudar filtros aqui, a página **não** recalcula tabelas e o painel de uploads acima — resposta mais rápida."
        )
    # Não usar st.expander aqui: esta função corre dentro do expander «Etapa 3» (Streamlit proíbe expanders aninhados).
    if st.checkbox("📖 Como isto funciona", value=False, key="garim_e3_como_funciona"):
        st.markdown(
            """
    <div style="background:#fff8fc;border:1px solid #f8bbd9;border-radius:10px;padding:14px 16px;margin-bottom:14px;font-size:0.93rem;line-height:1.55;color:#333;">
    <b>Filtros</b><br/>
    • <b>Emissão própria (esquerda):</b> só a sua empresa — status (inclui <b>Denegadas</b> e <b>Rejeitadas</b> em separado de canceladas/inutilizadas), tipo, operação, datas, série, n.º, UF destino, nota por chave ou n.º+série. Vazio = não filtra por estes campos.<br/>
    • <b>Terceiros (direita):</b> só documentos recebidos — os mesmos tipos de status quando aplicável, tipo, operação, período.<br/><br/>
    <b>Exportar</b><br/>
    • <b>ZIP todo o lote</b> — ignora filtros para escolher XML; Excel completo dentro de cada ZIP.<br/>
    • <b>ZIP filtrado</b> — só XML que passam nos filtros; Excel coerente com esse conjunto.<br/>
    • <b>Só Excel</b> — sem XML.<br/><br/>
    <b>Descargas</b> — sempre em <b>dois ficheiros</b> (própria e terceiros). Nos ZIP, o Excel está em <code>RELATORIO_GARIMPEIRO/</code> (até 10 000 XML por parte).<br/><br/>
    <b>Pacote contabilidade / matriz</b> (se aparecer em baixo) — exporta o lote lido; filtros da Etapa 3 <b>não</b> cortam. Na pasta escolhida: <b>Excel solto</b> (<code>…_relatorio_garimpeiro_completo.xlsx</code>, sem folha Painel Fiscal) + ZIPs. Em cada ZIP: pasta <code>XML/Lote_001</code>, <code>Lote_002</code>, … (até 10&nbsp;000 XML por pasta) + Excel na raiz com nome <code>relatorio_garimpeiro_…</code> (inclui série/mês/grupo, para não se sobrepor ao extrair); o .zip pode incluir <code>_notas_</code> inicial–final. <b>Emitidas</b>: um ZIP por série, status e <b>mês de emissão</b> (<code>Mes_AAAA_MM</code>); <b>Terceiros</b>: por modelo e status (sem mês).
    </div>
            """,
            unsafe_allow_html=True,
        )
    
    st.session_state.pop("v2_f_mes", None)
    st.session_state.pop("v2_f_mod", None)
    st.session_state.pop("v2_zip_org", None)
    st.session_state.pop("v2_zip_plano", None)
    st.session_state.pop("v2_excel_completo", None)
    for _k_v2 in ("v2_f_tip", "v2_f_ser", "v2_f_stat", "v2_f_op", "v2_f_uf"):
        if _k_v2 not in st.session_state:
            st.session_state[_k_v2] = []
    st.session_state["v2_f_orig"] = []
    if "v2_f_data_modo" not in st.session_state:
        st.session_state["v2_f_data_modo"] = "Qualquer"
    if "v2_f_faixa_modo" not in st.session_state:
        st.session_state["v2_f_faixa_modo"] = "Qualquer"
    if "v2_f_faixa_n1" not in st.session_state:
        st.session_state["v2_f_faixa_n1"] = 0
    if "v2_f_faixa_n2" not in st.session_state:
        st.session_state["v2_f_faixa_n2"] = 0
    _v2_fmt_validos = frozenset(
        {
            "zip_tudo_raiz",
            "zip_tudo_pastas",
            "zip_filt_pastas",
            "zip_filt_raiz",
            "excel_todo_lote",
            "excel_filtro",
        }
    )
    if "v2_export_format" not in st.session_state:
        st.session_state["v2_export_format"] = "zip_tudo_pastas"
    else:
        _vf_m = st.session_state.get("v2_export_format")
        if _vf_m == "zip_pastas":
            st.session_state["v2_export_format"] = "zip_tudo_pastas"
        elif _vf_m == "zip_raiz":
            st.session_state["v2_export_format"] = "zip_tudo_raiz"
        elif _vf_m not in _v2_fmt_validos:
            st.session_state["v2_export_format"] = "zip_tudo_pastas"
    if "v2_f_esp_chave" not in st.session_state:
        st.session_state["v2_f_esp_chave"] = ""
    if "v2_f_esp_num" not in st.session_state:
        st.session_state["v2_f_esp_num"] = 0
    if "v2_f_esp_ser" not in st.session_state:
        st.session_state["v2_f_esp_ser"] = ""
    for _kt in ("v2_t_stat", "v2_t_tip", "v2_t_op"):
        if _kt not in st.session_state:
            st.session_state[_kt] = []
    if "v2_t_data_modo" not in st.session_state:
        st.session_state["v2_t_data_modo"] = "Qualquer"
    
    df_g_base = st.session_state["df_geral"]
    _v2_tipos_ui = ["NF-e", "NFS-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e", "Outros"]
    _v2_stat_ui = list(_V2_STATUS_UI_PARA_DF.keys())
    _v2_op_ui = list(_V2_OP_UI_PARA_INTERNO.keys())
    
    _fo = list(st.session_state.get("v2_f_orig") or [])
    _ftip = list(st.session_state.get("v2_f_tip") or [])
    _fser = list(st.session_state.get("v2_f_ser") or [])
    _fst = list(st.session_state.get("v2_f_stat") or [])
    _fop = list(st.session_state.get("v2_f_op") or [])
    _fuf = list(st.session_state.get("v2_f_uf") or [])
    _fdm = st.session_state.get("v2_f_data_modo", "Qualquer")
    _fd1 = st.session_state.get("v2_f_data_d1")
    _fd2 = st.session_state.get("v2_f_data_d2")
    _ffm = st.session_state.get("v2_f_faixa_modo", "Qualquer")
    _fn1 = int(st.session_state.get("v2_f_faixa_n1") or 0)
    _fn2 = int(st.session_state.get("v2_f_faixa_n2") or 0)
    _ech = str(st.session_state.get("v2_f_esp_chave", "") or "")
    _en = int(st.session_state.get("v2_f_esp_num") or 0)
    _es = str(st.session_state.get("v2_f_esp_ser", "") or "").strip()
    _tst = list(st.session_state.get("v2_t_stat") or [])
    _ttp = list(st.session_state.get("v2_t_tip") or [])
    _top = list(st.session_state.get("v2_t_op") or [])
    _tdm_t = st.session_state.get("v2_t_data_modo", "Qualquer")
    _td1_t = st.session_state.get("v2_t_data_d1")
    _td2_t = st.session_state.get("v2_t_data_d2")
    
    _v2_c_sig = (id(df_g_base), v2_assinatura_exportacao_sessao())
    _v2_cc = st.session_state.get("_v2_cascade_cache_v1")
    if _v2_cc and _v2_cc.get("sig") == _v2_c_sig:
        series = _v2_cc["series"]
        ufs_opts = _v2_cc["ufs"]
    else:
        _opts = v2_opcoes_cascata_etapa3(
            df_g_base,
            _fo,
            _ftip,
            _fser,
            _fst,
            _fop,
            _fdm,
            _fd1,
            _fd2,
            _ffm,
            _fn1,
            _fn2,
            _fuf,
            _ech,
            _en,
            _es,
            _tst,
            _ttp,
            _top,
            _tdm_t,
            _td1_t,
            _td2_t,
        )
        series = _opts["series"]
        ufs_opts = _opts["ufs"]
        st.session_state["_v2_cascade_cache_v1"] = {
            "sig": _v2_c_sig,
            "series": series,
            "ufs": ufs_opts,
        }
    v2_sanear_selecoes_contra_opcoes(series, ufs_opts)
    
    _wp = st.session_state.pop("v2_preset_warn", None)
    if _wp:
        st.warning(_wp)
    
    if st.session_state.get("export_ready"):
        _sig_now = v2_assinatura_exportacao_sessao()
        if st.session_state.get("v2_export_sig") != _sig_now:
            _v2_limpar_estado_exportacao_etapa3(remover_zip_do_disco=False)
            st.session_state["v2_show_regen_hint"] = True
    
    if st.session_state.get("mariana_export_ready"):
        _sig_mar = v2_assinatura_pacote_matriz_sessao(df_g_base)
        if st.session_state.get("mariana_export_sig") != _sig_mar:
            st.session_state.pop("mariana_zip_parts", None)
            st.session_state["mariana_export_ready"] = False
            st.session_state.pop("mariana_export_sig", None)
            st.session_state.pop("mariana_sem_xml_msg", None)
            st.session_state.pop("mariana_excel_completo_path", None)
    
    if st.session_state.pop("v2_show_regen_hint", False):
        st.caption(
            "Mudou o tipo de exportação ou os filtros — clique **Gerar ficheiros** outra vez para atualizar."
        )
    
    _parts_o_pre = st.session_state.get("org_zip_parts") or []
    _parts_t_pre = st.session_state.get("todos_zip_parts") or []
    
    def _parts_com_tag_pre(parts, sub):
        sub_l = sub.lower()
        return [p for p in parts if sub_l in os.path.basename(p).lower()]
    
    _dual_nomes_zip_pre = (_parts_o_pre or _parts_t_pre) and any(
        "propria" in os.path.basename(p).lower()
        or "terceiros" in os.path.basename(p).lower()
        for p in (_parts_o_pre + _parts_t_pre)
    )
    _po_pr_pre = _parts_com_tag_pre(_parts_o_pre, "propria")
    _po_tc_pre = _parts_com_tag_pre(_parts_o_pre, "terceiros")
    _pt_pr_pre = _parts_com_tag_pre(_parts_t_pre, "propria")
    _pt_tc_pre = _parts_com_tag_pre(_parts_t_pre, "terceiros")
    _xbp_pre = st.session_state.get("excel_buffer_propria")
    _xbt_pre = st.session_state.get("excel_buffer_terceiros")
    _dual_ui_pre = (
        bool(st.session_state.get("v2_etapa3_dual_export"))
        or _dual_nomes_zip_pre
        or bool(_xbp_pre or _xbt_pre)
    )
    _dl_sem_pre = None
    if st.session_state.get("export_ready"):
        _dl_sem_pre = st.session_state.pop("v2_export_sem_xml", None)
    
    _dl_k_zip = [0]
    
    col_f_prop, col_f_terc = st.columns(2, gap="large")
    
    with col_f_prop:
        st.markdown(
            """
    <div class="garim-etapa3-bloco garim-etapa3-propria">
    <p class="garim-etapa3-titulo">Sua empresa (emissão própria)</p>
    </div>
            """,
            unsafe_allow_html=True,
        )
        with st.container(border=True):
            filtro_status = st.multiselect(
                "Status (vazio = todos)",
                _v2_stat_ui,
                key="v2_f_stat",
                help="Autorizadas = NORMAIS; Inutilizadas inclui faixas e linhas INUTILIZADA; Rejeitadas = denegação / cStat 3xx quando detectado no XML.",
            )
            filtro_tipos = st.multiselect(
                "Tipo de documento (vazio = todos)",
                _v2_tipos_ui,
                key="v2_f_tip",
                help="Alinhado à coluna Modelo. NFS-e / CT-e OS dependem do conteúdo do XML.",
            )
            filtro_operacao = st.multiselect(
                "Operação (vazio = entrada e saída)",
                _v2_op_ui,
                key="v2_f_op",
                help="Entrada / Saída (tpNF no XML).",
            )
            st.selectbox(
                "Data de emissão",
                _V2_DATA_MODELO_LABELS,
                key="v2_f_data_modo",
                help="Compara com a coluna «Data Emissão» do relatório (nas tabelas aqui e nos Excel: dd/mm/aaaa). «Qualquer» ignora datas.",
            )
            _dm_ui = st.session_state.get("v2_f_data_modo", "Qualquer")
            if _dm_ui == "Entre":
                _dc1, _dc2 = st.columns(2)
                with _dc1:
                    st.date_input("De", key="v2_f_data_d1")
                with _dc2:
                    st.date_input("Até", key="v2_f_data_d2")
            elif _dm_ui != "Qualquer":
                st.date_input("Data", key="v2_f_data_d1")
    
            if series:
                filtro_series = st.multiselect(
                    "Série (vazio = todas)",
                    series,
                    key="v2_f_ser",
                    help="Listagem em cascata: só séries da emissão própria compatíveis com o resto.",
                )
            else:
                st.caption("Série: nenhuma emissão própria disponível com os filtros atuais.")
                st.session_state["v2_f_ser"] = []
                filtro_series = []
            _tem_serie = bool(st.session_state.get("v2_f_ser"))
            if _tem_serie:
                st.selectbox(
                    "Faixa do n.º da nota (requer série acima)",
                    _V2_FAIXA_MODELO_LABELS,
                    key="v2_f_faixa_modo",
                )
                _fm_ui = st.session_state.get("v2_f_faixa_modo", "Qualquer")
                if _fm_ui == "Entre":
                    _fc1, _fc2 = st.columns(2)
                    with _fc1:
                        st.number_input("N.º ≥ / início", min_value=0, step=1, key="v2_f_faixa_n1")
                    with _fc2:
                        st.number_input("N.º ≤ / fim", min_value=0, step=1, key="v2_f_faixa_n2")
                elif _fm_ui != "Qualquer":
                    st.number_input("N.º da nota", min_value=0, step=1, key="v2_f_faixa_n1")
            else:
                st.caption("Escolha **série(s)** para ativar filtro por faixa do número da nota.")
    
            if ufs_opts:
                filtro_ufs = st.multiselect(
                    "UF destino (destinatário na NF)",
                    ufs_opts,
                    key="v2_f_uf",
                    help="UF extraída do bloco dest do XML. Vazio = todos os estados.",
                )
            else:
                st.caption(
                    "UF destino: ainda sem valores no lote (coluna «UF Destino» após novo garimpo) ou nenhuma emissão própria com UF no XML."
                )
                st.session_state["v2_f_uf"] = []
                filtro_ufs = []
    
            if st.checkbox(
                "Nota específica (chave **ou** n.º + série)",
                value=False,
                key="garim_e3_nota_esp",
            ):
                st.text_input(
                    "Chave da NF (44 dígitos; pode colar com espaços)",
                    key="v2_f_esp_chave",
                    placeholder="Ex.: 3525… (44 números)",
                    help="Se preencher 44 dígitos, usa só esta chave (ignora n.º/série abaixo). Inutilização: chave tipo INUT_…",
                )
                _ec1, _ec2 = st.columns(2)
                with _ec1:
                    st.number_input(
                        "N.º da nota",
                        min_value=0,
                        step=1,
                        key="v2_f_esp_num",
                        help="Use com **Série** se não usar chave completa.",
                    )
                with _ec2:
                    st.text_input(
                        "Série",
                        key="v2_f_esp_ser",
                        placeholder="Ex.: 1",
                        help="Obrigatória com n.º se não usar chave de 44 dígitos.",
                    )
                st.caption(
                    "Prioridade: **chave** com 44 dígitos (ou INUT_…) vence; senão filtra **n.º + série** na emissão própria."
                )
    
    with col_f_terc:
        st.markdown(
            """
    <div class="garim-etapa3-bloco garim-etapa3-terc">
    <p class="garim-etapa3-titulo">Terceiros (documentos recebidos)</p>
    </div>
            """,
            unsafe_allow_html=True,
        )
        with st.container(border=True):
            st.multiselect(
                "Status (vazio = todos)",
                _v2_stat_ui,
                key="v2_t_stat",
                help="Aplica-se apenas a linhas TERCEIROS no relatório.",
            )
            st.multiselect(
                "Tipo de documento (vazio = todos)",
                _v2_tipos_ui,
                key="v2_t_tip",
                help="Coluna Modelo — só terceiros.",
            )
            st.multiselect(
                "Operação (vazio = entrada e saída)",
                _v2_op_ui,
                key="v2_t_op",
                help="Entrada / Saída — só terceiros.",
            )
            st.selectbox(
                "Período (data de emissão)",
                _V2_DATA_MODELO_LABELS,
                key="v2_t_data_modo",
                help="Compara com «Data Emissão» (nas tabelas: dd/mm/aaaa). «Qualquer» ignora.",
            )
            _tdm_terc = st.session_state.get("v2_t_data_modo", "Qualquer")
            if _tdm_terc == "Entre":
                _tc1, _tc2 = st.columns(2)
                with _tc1:
                    st.date_input("De", key="v2_t_data_d1")
                with _tc2:
                    st.date_input("Até", key="v2_t_data_d2")
            elif _tdm_terc != "Qualquer":
                st.date_input("Data", key="v2_t_data_d1")
    
    with st.container(border=True):
        st.button(
            "Repor filtros e limpar exportação gerada",
            key="v2_pre_clr",
            on_click=v2_callback_repor_filtros,
            width="stretch",
        )
    
    _fmt_labels = {
        "zip_tudo_raiz": "ZIP de todo o lote (XML soltos, sem pastas)",
        "zip_tudo_pastas": "ZIP de todo o lote (com pastas)",
        "zip_filt_pastas": "ZIP só do filtrado (com pastas)",
        "zip_filt_raiz": "ZIP só do filtrado (XML soltos)",
        "excel_todo_lote": "Apenas Excel — relatório completo",
        "excel_filtro": "Apenas Excel — linhas filtradas",
    }
    _fmt_ordem = (
        "zip_tudo_pastas",
        "zip_tudo_raiz",
        "zip_filt_pastas",
        "zip_filt_raiz",
        "excel_todo_lote",
        "excel_filtro",
    )
    st.selectbox(
        "Tipo de exportação",
        _fmt_ordem,
        format_func=lambda k: _fmt_labels[k],
        key="v2_export_format",
        help="Depois de gerar: use os botões de descarga que aparecem abaixo. Vários ZIP = várias partes; guarde todas.",
    )
    if "v2_etapa3_nome_ficheiro" not in st.session_state:
        st.session_state["v2_etapa3_nome_ficheiro"] = str(
            st.session_state.pop("v2_etapa3_export_prefix", "") or ""
        )
    st.text_input(
        "Nome do ficheiro (opcional, sem extensão)",
        key="v2_etapa3_nome_ficheiro",
        placeholder="Ex.: Garimpo_Marco — vazio = nomes automáticos da app",
        help=(
            "Escreva só o **nome base** (sem .zip nem .xlsx). "
            "ZIP: Nome_org_propria_pt1.zip e Nome_todos_propria_pt1.zip (várias partes: _pt2, _pt3…). "
            "Excel: Nome_todo_propria_data.xlsx, Nome_filt_terceiros_data.xlsx, etc. "
            "Espaços viram _; caracteres inválidos são removidos."
        ),
    )

    fmt_e3 = st.session_state.get("v2_export_format", "zip_tudo_pastas")
    export_ignora_filtros = fmt_e3 in (
        "excel_todo_lote",
        "zip_tudo_raiz",
        "zip_tudo_pastas",
    )
    
    _ech_nf = str(st.session_state.get("v2_f_esp_chave", "") or "").strip()
    _nota_esp_ativa = (
        (_ech_nf.upper().startswith("INUT_") and len(_ech_nf) > 5)
        or len("".join(c for c in _ech_nf if c.isdigit())) >= 44
        or (
            int(st.session_state.get("v2_f_esp_num") or 0) > 0
            and str(st.session_state.get("v2_f_esp_ser", "") or "").strip() != ""
        )
    )
    
    filtro_origem = list(st.session_state.get("v2_f_orig") or [])
    nenhum_filtro = (
        len(filtro_origem) == 0
        and len(filtro_tipos) == 0
        and len(filtro_series) == 0
        and len(filtro_status) == 0
        and len(filtro_operacao) == 0
        and len(filtro_ufs) == 0
        and (st.session_state.get("v2_f_data_modo") or "Qualquer") == "Qualquer"
        and (
            not st.session_state.get("v2_f_ser")
            or (st.session_state.get("v2_f_faixa_modo") or "Qualquer") == "Qualquer"
        )
        and not _nota_esp_ativa
        and len(st.session_state.get("v2_t_stat") or []) == 0
        and len(st.session_state.get("v2_t_tip") or []) == 0
        and len(st.session_state.get("v2_t_op") or []) == 0
        and (st.session_state.get("v2_t_data_modo") or "Qualquer") == "Qualquer"
    )
    precisa_confirm_filtro = nenhum_filtro and not export_ignora_filtros
    if precisa_confirm_filtro:
        confirm_export_total = st.checkbox(
            "Incluir todo o relatório (não há filtros por coluna)",
            value=True,
            key="v2_confirm_full",
        )
    else:
        confirm_export_total = True
    
    _btn_dis = (precisa_confirm_filtro and not confirm_export_total) or df_g_base.empty
    
    def _v2_montar_filtrado_export():
        _fo = list(st.session_state.get("v2_f_orig") or [])
        _ftip = list(st.session_state.get("v2_f_tip") or [])
        _fser = list(st.session_state.get("v2_f_ser") or [])
        _fst = list(st.session_state.get("v2_f_stat") or [])
        _fop = list(st.session_state.get("v2_f_op") or [])
        _fuf = list(st.session_state.get("v2_f_uf") or [])
        _fdm = st.session_state.get("v2_f_data_modo", "Qualquer")
        _fd1 = st.session_state.get("v2_f_data_d1")
        _fd2 = st.session_state.get("v2_f_data_d2")
        _ffm = st.session_state.get("v2_f_faixa_modo", "Qualquer")
        _fn1 = int(st.session_state.get("v2_f_faixa_n1") or 0)
        _fn2 = int(st.session_state.get("v2_f_faixa_n2") or 0)
        _ech_btn = str(st.session_state.get("v2_f_esp_chave", "") or "")
        _en_btn = int(st.session_state.get("v2_f_esp_num") or 0)
        _es_btn = str(st.session_state.get("v2_f_esp_ser", "") or "").strip()
        _tst_btn = list(st.session_state.get("v2_t_stat") or [])
        _ttp_btn = list(st.session_state.get("v2_t_tip") or [])
        _top_btn = list(st.session_state.get("v2_t_op") or [])
        _tdm_btn = st.session_state.get("v2_t_data_modo", "Qualquer")
        _td1_btn = st.session_state.get("v2_t_data_d1")
        _td2_btn = st.session_state.get("v2_t_data_d2")
        return filtrar_df_geral_para_exportacao(
            df_g_base,
            _fo,
            _ftip,
            _fser,
            _fst,
            _fop,
            _fdm,
            _fd1,
            _fd2,
            _ffm,
            _fn1,
            _fn2,
            _fuf,
            _ech_btn,
            _en_btn,
            _es_btn,
            _tst_btn,
            _ttp_btn,
            _top_btn,
            _tdm_btn,
            _td1_btn,
            _td2_btn,
        )
    
    _prev_k = (
        id(df_g_base),
        fmt_e3,
        v2_assinatura_exportacao_sessao(),
        bool(confirm_export_total),
    )
    _pc = st.session_state.get("_v2_e3_preview_dis_cache")
    if _pc and _pc.get("k") == _prev_k:
        _dis_pr = _pc["dis_pr"]
        _dis_tc = _pc["dis_tc"]
        _dis_ambos = _pc["dis_ambos"]
    else:
        _dfp_h = _df_apenas_emissao_propria(df_g_base)
        _dft_h = _df_apenas_terceiros(df_g_base)
        if fmt_e3.startswith("zip_filt") or fmt_e3 == "excel_filtro":
            _daf_h = _v2_montar_filtrado_export()
            if _daf_h is None or _daf_h.empty:
                _dfp_h = pd.DataFrame()
                _dft_h = pd.DataFrame()
            else:
                _dfp_h = _df_apenas_emissao_propria(_daf_h)
                _dft_h = _df_apenas_terceiros(_daf_h)

        _dis_pr = _btn_dis or _dfp_h.empty
        _dis_tc = _btn_dis or _dft_h.empty
        _dis_ambos = _btn_dis or (_dfp_h.empty and _dft_h.empty)
        st.session_state["_v2_e3_preview_dis_cache"] = {
            "k": _prev_k,
            "dis_pr": _dis_pr,
            "dis_tc": _dis_tc,
            "dis_ambos": _dis_ambos,
        }
    
    _pacote_matriz_btn_dis = df_g_base.empty
    gen_pacote_matriz = False

    if str(st.session_state.get("v2_export_format", "")).startswith("zip_"):
        if "v2_etapa3_zip_save_dir" not in st.session_state:
            st.session_state["v2_etapa3_zip_save_dir"] = ""
        st.text_input(
            "Pasta onde guardar os ZIP — sua empresa / terceiros",
            key="v2_etapa3_zip_save_dir",
            help="Obrigatório antes de gerar ZIP. Caminho completo (ex.: D:\\Exportacoes). A pasta é criada se não existir.",
        )

    col_g_pr, col_g_tc = st.columns(2, gap="large")
    with col_g_pr:
        gen_pr = st.button(
            "Gerar — sua empresa",
            type="primary",
            key="v2_btn_export_pr",
            disabled=_dis_pr,
            width="stretch",
        )
    with col_g_tc:
        gen_tc = st.button(
            "Gerar — terceiros",
            type="primary",
            key="v2_btn_export_tc",
            disabled=_dis_tc,
            width="stretch",
        )
    gen_ambos = st.button(
        "Gerar os dois lados",
        key="v2_btn_export_ambos",
        disabled=_dis_ambos,
        width="stretch",
    )
    
    _lados_run = None
    if gen_pr:
        _lados_run = frozenset({"propria"})
    elif gen_tc:
        _lados_run = frozenset({"terceiros"})
    elif gen_ambos:
        _lados_run = frozenset({"propria", "terceiros"})
    
    if _lados_run:
        _fmt_run = st.session_state.get("v2_export_format", "zip_tudo_pastas")
        _lados_tuple = tuple(sorted(_lados_run))
        _nome_arq = _v2_sanitize_nome_export(
            st.session_state.get("v2_etapa3_nome_ficheiro")
        )
        _zip_nome_raw = st.session_state.get("v2_etapa3_nome_ficheiro", "")

        def _montar_filtrado(sem_origem=False):
            _fo = list(st.session_state.get("v2_f_orig") or [])
            _ftip = list(st.session_state.get("v2_f_tip") or [])
            _fser = list(st.session_state.get("v2_f_ser") or [])
            _fst = list(st.session_state.get("v2_f_stat") or [])
            _fop = list(st.session_state.get("v2_f_op") or [])
            _fuf = list(st.session_state.get("v2_f_uf") or [])
            _fdm = st.session_state.get("v2_f_data_modo", "Qualquer")
            _fd1 = st.session_state.get("v2_f_data_d1")
            _fd2 = st.session_state.get("v2_f_data_d2")
            _ffm = st.session_state.get("v2_f_faixa_modo", "Qualquer")
            _fn1 = int(st.session_state.get("v2_f_faixa_n1") or 0)
            _fn2 = int(st.session_state.get("v2_f_faixa_n2") or 0)
            _ech_btn = str(st.session_state.get("v2_f_esp_chave", "") or "")
            _en_btn = int(st.session_state.get("v2_f_esp_num") or 0)
            _es_btn = str(st.session_state.get("v2_f_esp_ser", "") or "").strip()
            _tst_btn = list(st.session_state.get("v2_t_stat") or [])
            _ttp_btn = list(st.session_state.get("v2_t_tip") or [])
            _top_btn = list(st.session_state.get("v2_t_op") or [])
            _tdm_btn = st.session_state.get("v2_t_data_modo", "Qualquer")
            _td1_btn = st.session_state.get("v2_t_data_d1")
            _td2_btn = st.session_state.get("v2_t_data_d2")
            fo_eff = [] if sem_origem else _fo
            return filtrar_df_geral_para_exportacao(
                df_g_base,
                fo_eff,
                _ftip,
                _fser,
                _fst,
                _fop,
                _fdm,
                _fd1,
                _fd2,
                _ffm,
                _fn1,
                _fn2,
                _fuf,
                _ech_btn,
                _en_btn,
                _es_btn,
                _tst_btn,
                _ttp_btn,
                _top_btn,
                _tdm_btn,
                _td1_btn,
                _td2_btn,
            )
    
        ts = datetime.now().strftime("%Y%m%d_%H%M")

        if _fmt_run == "excel_todo_lote":
            if df_g_base.empty:
                st.warning("Relatório geral vazio.")
            else:
                with st.spinner("A gerar Excel…"):
                    st.session_state["org_zip_parts"] = []
                    st.session_state["todos_zip_parts"] = []
                    st.session_state["excel_buffer"] = None
                    st.session_state.pop("excel_buffer_propria", None)
                    st.session_state.pop("excel_buffer_terceiros", None)
                    st.session_state.pop("export_excel_name_propria", None)
                    st.session_state.pop("export_excel_name_terceiros", None)
                    _dp = _df_apenas_emissao_propria(df_g_base)
                    _dt = _df_apenas_terceiros(df_g_base)
                    _bp = None
                    _bt = None
                    if "propria" in _lados_run and not _dp.empty:
                        _bp = _excel_bytes_geral_e_resumo_status(_dp)
                    if "terceiros" in _lados_run and not _dt.empty:
                        _bt = _excel_bytes_geral_e_resumo_status(_dt)
                    if not _bp and not _bt:
                        st.warning(
                            "Sem linhas para o lado escolhido no relatório geral."
                        )
                        st.session_state["export_ready"] = False
                        st.session_state.pop("v2_export_lados", None)
                    else:
                        if _bp:
                            st.session_state["excel_buffer_propria"] = _bp
                            st.session_state["export_excel_name_propria"] = (
                                f"{_nome_arq}_todo_propria_{ts}.xlsx"
                                if _nome_arq
                                else f"relatorio_todo_lote_emissao_propria_{ts}.xlsx"
                            )
                        if _bt:
                            st.session_state["excel_buffer_terceiros"] = _bt
                            st.session_state["export_excel_name_terceiros"] = (
                                f"{_nome_arq}_todo_terceiros_{ts}.xlsx"
                                if _nome_arq
                                else f"relatorio_todo_lote_terceiros_{ts}.xlsx"
                            )
                        st.session_state["export_ready"] = True
                        st.session_state["v2_etapa3_dual_export"] = True
                        st.session_state["v2_export_lados"] = _lados_tuple
                        st.session_state["v2_export_sig"] = (
                            v2_assinatura_exportacao_sessao()
                        )
                        st.session_state.pop("v2_export_sem_xml", None)
                    gc.collect()
                st.rerun()
    
        elif _fmt_run == "excel_filtro":
            df_geral_filtrado = _montar_filtrado(sem_origem=False)
            if df_geral_filtrado is None or df_geral_filtrado.empty:
                st.warning("Resultado filtrado: 0 linhas. Ajuste os filtros.")
            else:
                with st.spinner("A gerar Excel…"):
                    st.session_state["org_zip_parts"] = []
                    st.session_state["todos_zip_parts"] = []
                    st.session_state["excel_buffer"] = None
                    st.session_state.pop("excel_buffer_propria", None)
                    st.session_state.pop("excel_buffer_terceiros", None)
                    st.session_state.pop("export_excel_name_propria", None)
                    st.session_state.pop("export_excel_name_terceiros", None)
                    _dp = _df_apenas_emissao_propria(df_geral_filtrado)
                    _dt = _df_apenas_terceiros(df_geral_filtrado)
                    _bp = None
                    _bt = None
                    if "propria" in _lados_run and not _dp.empty:
                        _bp = _v2_excel_bytes_filtrado_etapa3(_dp)
                    if "terceiros" in _lados_run and not _dt.empty:
                        _bt = _v2_excel_bytes_filtrado_etapa3(_dt)
                    if not _bp and not _bt:
                        st.warning(
                            "Com os filtros atuais não há linhas do lado escolhido."
                        )
                        st.session_state["export_ready"] = False
                        st.session_state.pop("v2_export_lados", None)
                    else:
                        if _bp:
                            st.session_state["excel_buffer_propria"] = _bp
                            st.session_state["export_excel_name_propria"] = (
                                f"{_nome_arq}_filt_propria_{ts}.xlsx"
                                if _nome_arq
                                else f"relatorio_filtrado_emissao_propria_{ts}.xlsx"
                            )
                        if _bt:
                            st.session_state["excel_buffer_terceiros"] = _bt
                            st.session_state["export_excel_name_terceiros"] = (
                                f"{_nome_arq}_filt_terceiros_{ts}.xlsx"
                                if _nome_arq
                                else f"relatorio_filtrado_terceiros_{ts}.xlsx"
                            )
                        st.session_state["export_ready"] = True
                        st.session_state["v2_etapa3_dual_export"] = True
                        st.session_state["v2_export_lados"] = _lados_tuple
                        st.session_state["v2_export_sig"] = (
                            v2_assinatura_exportacao_sessao()
                        )
                        st.session_state.pop("v2_export_sem_xml", None)
                    gc.collect()
                st.rerun()
    
        elif _fmt_run.startswith("zip_"):
            _xml_filt = _fmt_run.startswith("zip_filt")
            dfp = pd.DataFrame()
            dft = pd.DataFrame()
            df_all_f = None
            if _xml_filt:
                df_all_f = _montar_filtrado(sem_origem=False)
                if df_all_f is None or df_all_f.empty:
                    st.warning("Resultado filtrado: 0 linhas. Ajuste os filtros.")
                else:
                    dfp = _df_apenas_emissao_propria(df_all_f)
                    dft = _df_apenas_terceiros(df_all_f)
            elif df_g_base.empty:
                st.warning("Relatório geral vazio.")
            else:
                dfp = _df_apenas_emissao_propria(df_g_base)
                dft = _df_apenas_terceiros(df_g_base)
    
            _pares_zip = []
            if "propria" in _lados_run and dfp is not None and not dfp.empty:
                _pares_zip.append((dfp, "propria"))
            if "terceiros" in _lados_run and dft is not None and not dft.empty:
                _pares_zip.append((dft, "terceiros"))
    
            _pode_zip = False
            if _xml_filt:
                _pode_zip = (
                    df_all_f is not None
                    and not df_all_f.empty
                    and bool(_pares_zip)
                )
            else:
                _pode_zip = not df_g_base.empty and bool(_pares_zip)
    
            if _pode_zip and _pares_zip:
                _path_zip_out, _err_zip_dir = _v2_destino_zip_etapa3_para_gravar()
                if _err_zip_dir:
                    st.error(_err_zip_dir)
                else:
                    with st.spinner("A gerar ZIP…"):
                        st.session_state["excel_buffer"] = None
                        st.session_state.pop("excel_buffer_propria", None)
                        st.session_state.pop("excel_buffer_terceiros", None)
                        st.session_state.pop("export_excel_name_propria", None)
                        st.session_state.pop("export_excel_name_terceiros", None)
                        gc.collect()
                        st.session_state["export_excel_name"] = (
                            f"{_nome_arq}_relatorio_zip_{ts}.xlsx"
                            if _nome_arq
                            else f"relatorio_completo_{ts}.xlsx"
                        )
                        org_all = []
                        todos_all = []
                        _err_zip = None
                        for df_sl, ztag in _pares_zip:
                            o, t, _xm, av, _ = _v2_export_zip_etapa3(
                                df_sl,
                                xml_respeita_filtro=_xml_filt,
                                df_filtrado_para_excel_bloco=(
                                    df_sl if _xml_filt else None
                                ),
                                excel_um_só_completo=_fmt_run.startswith(
                                    "zip_tudo"
                                ),
                                df_excel_completo=df_sl,
                                v2_zip_org=_fmt_run.endswith("_pastas"),
                                v2_zip_plano=_fmt_run.endswith("_raiz"),
                                cnpj_limpo=cnpj_limpo,
                                zip_tag=ztag,
                                zip_output_dir=_path_zip_out,
                                zip_nome_ficheiro=_zip_nome_raw,
                            )
                            if av and str(av).startswith("ERR:"):
                                _err_zip = str(av)[4:].strip()
                                break
                            org_all.extend(o)
                            todos_all.extend(t)
                    if _err_zip:
                        st.warning(_err_zip)
                        st.session_state.pop("v2_export_sem_xml", None)
                        st.session_state["org_zip_parts"] = []
                        st.session_state["todos_zip_parts"] = []
                        st.session_state["export_ready"] = False
                        st.session_state["v2_etapa3_dual_export"] = False
                        st.session_state.pop("v2_export_lados", None)
                    else:
                        st.session_state.pop("v2_export_sem_xml", None)
                        st.session_state.update(
                            {
                                "org_zip_parts": org_all
                                if _fmt_run.endswith("_pastas")
                                else [],
                                "todos_zip_parts": todos_all
                                if _fmt_run.endswith("_raiz")
                                else [],
                                "export_ready": True,
                                "v2_etapa3_dual_export": True,
                                "v2_export_lados": _lados_tuple,
                                "v2_export_sig": v2_assinatura_exportacao_sessao(),
                            }
                        )
                    gc.collect()
                st.rerun()
            elif _xml_filt and df_all_f is not None and not df_all_f.empty:
                if not _pares_zip:
                    st.warning(
                        "Não há linhas do lado escolhido para exportar em ZIP."
                    )
            elif not _xml_filt and not df_g_base.empty and not _pares_zip:
                st.warning(
                    "Não há linhas do lado escolhido para exportar em ZIP."
                )
    
    if _dl_sem_pre:
        st.warning(_dl_sem_pre)
    
    if st.session_state.get("export_ready") and _dual_ui_pre:
        st.markdown("**Descarregar**")
        st.caption(
            "Só aparecem botões do lado que gerou. Vários ZIP = várias partes — guarde todos."
        )
        _lados_ger = st.session_state.get("v2_export_lados")
        if _lados_ger is None:
            _lados_ger = ("propria", "terceiros")
        d_dl_pr, d_dl_tc = st.columns(2, gap="large")
        with d_dl_pr:
            with st.container(border=True):
                st.caption("Sua empresa")
                if (
                    "propria" not in _lados_ger
                    and not _xbp_pre
                    and not (_po_pr_pre or _pt_pr_pre)
                ):
                    st.caption(
                        "Não gerado nesta sessão — use **Gerar — sua empresa** ou **Gerar os dois**."
                    )
                if _dual_nomes_zip_pre:
                    if _po_pr_pre or _pt_pr_pre:
                        for part in _po_pr_pre:
                            _dl_k_zip[0] += 1
                            if os.path.exists(part):
                                with open(part, "rb") as fp:
                                    st.download_button(
                                        rotulo_download_zip_parte(part),
                                        fp.read(),
                                        file_name=os.path.basename(part),
                                        key=f"v2_dlo_p_{_dl_k_zip[0]}",
                                        width="stretch",
                                    )
                        for part in _pt_pr_pre:
                            _dl_k_zip[0] += 1
                            if os.path.exists(part):
                                with open(part, "rb") as fp:
                                    st.download_button(
                                        rotulo_download_zip_parte(part),
                                        fp.read(),
                                        file_name=os.path.basename(part),
                                        key=f"v2_dlt_p_{_dl_k_zip[0]}",
                                        width="stretch",
                                    )
                    elif "propria" in _lados_ger:
                        st.caption("Nada a descarregar deste lado.")
                if _xbp_pre:
                    st.download_button(
                        "Excel — sua empresa",
                        _xbp_pre,
                        file_name=st.session_state.get(
                            "export_excel_name_propria",
                            "relatorio_emissao_propria.xlsx",
                        ),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="v2_dl_xlsx_propria",
                        width="stretch",
                    )
        with d_dl_tc:
            with st.container(border=True):
                st.caption("Terceiros")
                if (
                    "terceiros" not in _lados_ger
                    and not _xbt_pre
                    and not (_po_tc_pre or _pt_tc_pre)
                ):
                    st.caption(
                        "Não gerado nesta sessão — use **Gerar — terceiros** ou **Gerar os dois**."
                    )
                if _dual_nomes_zip_pre:
                    if _po_tc_pre or _pt_tc_pre:
                        for part in _po_tc_pre:
                            _dl_k_zip[0] += 1
                            if os.path.exists(part):
                                with open(part, "rb") as fp:
                                    st.download_button(
                                        rotulo_download_zip_parte(part),
                                        fp.read(),
                                        file_name=os.path.basename(part),
                                        key=f"v2_dlo_t_{_dl_k_zip[0]}",
                                        width="stretch",
                                    )
                        for part in _pt_tc_pre:
                            _dl_k_zip[0] += 1
                            if os.path.exists(part):
                                with open(part, "rb") as fp:
                                    st.download_button(
                                        rotulo_download_zip_parte(part),
                                        fp.read(),
                                        file_name=os.path.basename(part),
                                        key=f"v2_dlt_t_{_dl_k_zip[0]}",
                                        width="stretch",
                                    )
                    elif "terceiros" in _lados_ger:
                        st.caption("Nada a descarregar deste lado.")
                if _xbt_pre:
                    st.download_button(
                        "Excel — terceiros",
                        _xbt_pre,
                        file_name=st.session_state.get(
                            "export_excel_name_terceiros",
                            "relatorio_terceiros.xlsx",
                        ),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="v2_dl_xlsx_terceiros",
                        width="stretch",
                    )
    
    if st.session_state.get("export_ready"):
        _parts_o = st.session_state.get("org_zip_parts") or []
        _parts_t = st.session_state.get("todos_zip_parts") or []
        _xbuf = st.session_state.get("excel_buffer")
        _xbp = st.session_state.get("excel_buffer_propria")
        _xbt = st.session_state.get("excel_buffer_terceiros")
        _dual_nomes_zip_b = (_parts_o or _parts_t) and any(
            "propria" in os.path.basename(p).lower()
            or "terceiros" in os.path.basename(p).lower()
            for p in (_parts_o + _parts_t)
        )
        _dual_ui_b = (
            bool(st.session_state.get("v2_etapa3_dual_export"))
            or _dual_nomes_zip_b
            or bool(_xbp or _xbt)
        )
        if not _dual_ui_b:
            if _parts_o or _parts_t:
                st.caption(
                    "ZIP (formato antigo). Prefira gerar de novo para ter própria e terceiros separados."
                )
                _dl_i = 0
                if _parts_o:
                    for part in _parts_o:
                        _dl_i += 1
                        if os.path.exists(part):
                            with open(part, "rb") as fp:
                                st.download_button(
                                    rotulo_download_zip_parte(part),
                                    fp.read(),
                                    file_name=os.path.basename(part),
                                    key=f"v2_dlo_{_dl_i}",
                                    width="stretch",
                                )
                if _parts_t:
                    for part in _parts_t:
                        _dl_i += 1
                        if os.path.exists(part):
                            with open(part, "rb") as fp:
                                st.download_button(
                                    rotulo_download_zip_parte(part),
                                    fp.read(),
                                    file_name=os.path.basename(part),
                                    key=f"v2_dlt_{_dl_i}",
                                    width="stretch",
                                )
            elif _xbuf or _xbp or _xbt:
                st.caption("Ficheiros prontos abaixo.")
    
            if _xbuf:
                st.download_button(
                    "Excel",
                    _xbuf,
                    file_name=st.session_state.get(
                        "export_excel_name", "relatorio_completo.xlsx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="v2_dl_xlsx",
                    width="stretch",
                )
            if _xbp:
                st.download_button(
                    "Excel — própria",
                    _xbp,
                    file_name=st.session_state.get(
                        "export_excel_name_propria",
                        "relatorio_emissao_propria.xlsx",
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="v2_dl_xlsx_propria_lo",
                    width="stretch",
                )
            if _xbt:
                st.download_button(
                    "Excel — terceiros",
                    _xbt,
                    file_name=st.session_state.get(
                        "export_excel_name_terceiros",
                        "relatorio_terceiros.xlsx",
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="v2_dl_xlsx_terceiros_lo",
                    width="stretch",
                )

    if _is_mariana_pc_bundle():
        st.markdown("---")
        st.markdown("##### Pacote para **contabilidade / matriz**")
        if "mariana_zip_save_dir" not in st.session_state:
            st.session_state["mariana_zip_save_dir"] = ""
        st.text_input(
            "Pasta onde guardar o pacote ZIP (caminho completo)",
            key="mariana_zip_save_dir",
            help="Obrigatório antes de gerar. Caminho completo — a pasta é criada se não existir.",
        )
        st.text_input(
            "Nome base do ZIP (opcional, sem .zip)",
            key="mariana_zip_basename",
            placeholder="Opcional",
            help="Prefixo do nome dos ficheiros ZIP (opcional).",
        )
        gen_pacote_matriz = st.button(
            "Gerar pacote ZIP (apuração / contabilidade)",
            key="v2_btn_mariana_zip",
            disabled=_pacote_matriz_btn_dis,
            width="stretch",
        )
        if gen_pacote_matriz:
            if df_g_base is None or df_g_base.empty:
                st.warning("Relatório geral vazio. Conclua o garimpo antes de gerar o pacote.")
            else:
                _pacote_matriz_rerun = False
                try:
                    with st.spinner("A gerar pacote ZIP…"):
                        _out_m, _err_m = _mariana_destino_zip_para_gravar()
                        if _err_m:
                            st.error(_err_m)
                        else:
                            st.session_state.pop("mariana_zip_parts", None)
                            st.session_state.pop("mariana_excel_completo_path", None)
                            _mar_bn = str(
                                st.session_state.get("mariana_zip_basename") or ""
                            ).strip()
                            parts_m, _xm_m, av_m, excel_m = _v2_export_zip_mariana(
                                df_g_base,
                                cnpj_limpo,
                                zip_output_dir=_out_m,
                                zip_file_stem=_mar_bn if _mar_bn else None,
                            )
                            if av_m and str(av_m).startswith("ERR:"):
                                st.warning(str(av_m)[4:].strip())
                                st.session_state["mariana_export_ready"] = False
                                st.session_state.pop("mariana_export_sig", None)
                                st.session_state.pop("mariana_sem_xml_msg", None)
                                st.session_state.pop("mariana_excel_completo_path", None)
                            elif not parts_m and not excel_m:
                                st.warning(av_m or "Nada a exportar.")
                                st.session_state["mariana_export_ready"] = False
                                st.session_state.pop("mariana_export_sig", None)
                                st.session_state.pop("mariana_sem_xml_msg", None)
                                st.session_state.pop("mariana_excel_completo_path", None)
                            else:
                                st.session_state["mariana_zip_parts"] = parts_m or []
                                if excel_m:
                                    st.session_state["mariana_excel_completo_path"] = (
                                        excel_m
                                    )
                                else:
                                    st.session_state.pop(
                                        "mariana_excel_completo_path", None
                                    )
                                st.session_state["mariana_export_ready"] = True
                                st.session_state["mariana_export_sig"] = (
                                    v2_assinatura_pacote_matriz_sessao(df_g_base)
                                )
                                if av_m:
                                    st.session_state["mariana_sem_xml_msg"] = av_m
                                else:
                                    st.session_state.pop("mariana_sem_xml_msg", None)
                                _pacote_matriz_rerun = True
                        gc.collect()
                except Exception as ex:
                    if _erro_sem_espaco_disco(ex):
                        st.error(
                            "**Sem espaço em disco** ao gerar o Excel do pacote (errno 28). "
                            "Liberte espaço em **C:** (pasta Temp do utilizador) e no disco onde grava o pacote; "
                            "ou defina **TEMP** e **TMP** para uma pasta noutro disco (ex.: `D:\\Temp`) e reinicie o Garimpeiro."
                        )
                    else:
                        st.error(f"**Erro ao gerar o pacote:** {ex}")
                        st.exception(ex)
                    st.session_state["mariana_export_ready"] = False
                    st.session_state.pop("mariana_export_sig", None)
                    st.session_state.pop("mariana_sem_xml_msg", None)
                    st.session_state.pop("mariana_excel_completo_path", None)
                if _pacote_matriz_rerun:
                    st.rerun()
        if st.session_state.get("mariana_export_ready"):
            _mmsg = st.session_state.get("mariana_sem_xml_msg")
            if _mmsg:
                st.warning(_mmsg)
            _mxp = st.session_state.get("mariana_excel_completo_path")
            if _mxp and os.path.isfile(_mxp):
                with open(_mxp, "rb") as _xfe:
                    st.download_button(
                        "Excel completo (sem Painel Fiscal) — mesmo ficheiro da pasta",
                        _xfe.read(),
                        file_name=os.path.basename(_mxp),
                        key="v2_dl_mariana_excel_completo",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
            st.markdown("**Descarregar pacote (ZIP)**")
            for _i_m, part_m in enumerate(
                st.session_state.get("mariana_zip_parts") or []
            ):
                if os.path.exists(part_m):
                    with open(part_m, "rb") as fm:
                        st.download_button(
                            f"Contabilidade · {rotulo_download_zip_parte(part_m)}",
                            fm.read(),
                            file_name=os.path.basename(part_m),
                            key=f"v2_dl_mariana_{_i_m}",
                            width="stretch",
                        )


def _garim_etapa3_fragment_entry():
    """Ponto de entrada estável para `st.fragment` (evita redefinir a função a cada rerun)."""
    _cx = "".join(c for c in str(st.session_state.get("cnpj_widget", "")) if c.isdigit())[:14]
    _garim_etapa3_corpo(_cx)


if st.session_state.get("confirmado"):
    if not st.session_state.get("garimpo_ok"):
        st.markdown(
            f'<h5>{_garim_emoji("📎")} Documentos XML / ZIP para ler</h5>',
            unsafe_allow_html=True,
        )
        st.caption(
            "Carregue abaixo os ficheiros do lote; depois use **Iniciar grande garimpo** para ler e montar o relatório."
        )
        uploaded_files = st.file_uploader("📂 Escolha os XML e/ou ZIP (suporta grandes volumes):", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            limpar_arquivos_temp()
            _mem_lote = {}
            if not _garimpo_analise_sem_pasta_local_projeto():
                os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)

            lote_dict = {}
            progresso_bar = st.progress(0)
            status_text = st.empty()
            total_arquivos = len(uploaded_files)

            _lbl_status = (
                "⛏️ A carregar o lote em memória (sem pasta no projeto)…"
                if _garimpo_analise_sem_pasta_local_projeto()
                else "⛏️ Minerando e salvando fisicamente…"
            )
            with st.status(_lbl_status, expanded=True) as status_box:

                for i, f in enumerate(uploaded_files):
                    raw = f.getvalue()
                    key = _garimpo_nome_chave_upload(i, getattr(f, "name", None))
                    if _garimpo_analise_sem_pasta_local_projeto():
                        _mem_lote[key] = raw
                    else:
                        caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, key)
                        with open(caminho_salvo, "wb") as out_f:
                            out_f.write(raw)

                if _garimpo_analise_sem_pasta_local_projeto() and _mem_lote:
                    try:
                        st.session_state[SESSION_KEY_FONTES_XML_MEMORIA] = _mem_lote
                    except Exception:
                        os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
                        for _mk, _mraw in _mem_lote.items():
                            with open(os.path.join(TEMP_UPLOADS_DIR, _mk), "wb") as out_f:
                                out_f.write(_mraw)
                
                lista_salvos = _lista_nomes_fontes_xml_garimpo()
                total_salvos = len(lista_salvos)
                
                for i, f_name in enumerate(lista_salvos):
                    if i % 50 == 0: 
                        gc.collect()
                        
                    progresso_bar.progress((i + 1) / max(total_salvos, 1))
                    status_text.text(f"⛏️ Lendo conteúdo: {f_name}")
                    
                    try:
                        with _abrir_fonte_xml_garimpo_stream(f_name) as file_obj:
                            todos_xmls = extrair_recursivo(file_obj, f_name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in [
                                            "CANCELADOS",
                                            "INUTILIZADOS",
                                            "DENEGADOS",
                                            "REJEITADOS",
                                        ]:
                                            lote_dict[key] = (res, is_p)
                                    else:
                                        lote_dict[key] = (res, is_p)
                                del xml_data 
                    except Exception as e: 
                        continue
                
                status_box.update(label="✅ Leitura Concluída!", state="complete", expanded=False)
                progresso_bar.empty()
                status_text.empty()

            rel_list = []
            ref_ar, ref_mr, ref_map = buraco_ctx_sessao()
            audit_map = {}
            canc_list = []
            inut_list = []
            aut_list = []
            deneg_list = []
            rej_list = []
            geral_list = []
            
            for k, (res, is_p) in lote_dict.items():
                rel_list.append(res)
                
                if is_p:
                    origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})"
                else:
                    origem_label = f"TERCEIROS ({res['Operacao']})"
                
                registro_base = {
                    "Origem": origem_label, 
                    "Operação": res["Operacao"], 
                    "Modelo": res["Tipo"], 
                    "Série": res["Série"], 
                    "Nota": res["Número"], 
                    "Data Emissão": res["Data_Emissao"],
                    "CNPJ Emitente": res["CNPJ_Emit"], 
                    "Nome Emitente": res["Nome_Emit"],
                    "Doc Destinatário": res["Doc_Dest"], 
                    "Nome Destinatário": res["Nome_Dest"],
                    "UF Destino": res.get("UF_Dest") or "",
                    "Chave": res["Chave"], 
                    "Status Final": res["Status"], 
                    "Valor": res["Valor"],
                    "Ano": res["Ano"], 
                    "Mes": res["Mes"]
                }

                if res["Status"] == "INUTILIZADOS":
                    r = res.get("Range", (res["Número"], res["Número"]))
                    for n in range(r[0], r[1] + 1):
                        item_inut = registro_base.copy()
                        item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                        geral_list.append(item_inut)
                else:
                    geral_list.append(registro_base)

                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])

                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        _man_inut = _inutil_sem_xml_manual(res)
                        for n in range(r[0], r[1] + 1):
                            inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                            if _incluir_em_resumo_por_serie(res, is_p, cnpj_limpo):
                                if sk not in audit_map:
                                    audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                                audit_map[sk]["nums"].add(n)
                                if _man_inut or incluir_numero_no_conjunto_buraco(
                                    res["Ano"], res["Mes"], n, ref_ar, ref_mr, ult_u
                                ):
                                    audit_map[sk]["nums_buraco"].add(n)
                    else:
                        if res["Status"] == "DENEGADOS":
                            deneg_list.append(registro_base)
                        elif res["Status"] == "REJEITADOS":
                            rej_list.append(registro_base)
                        elif res["Número"] > 0:
                            if res["Status"] == "CANCELADOS":
                                canc_list.append(registro_base)
                            elif res["Status"] == "NORMAIS":
                                aut_list.append(registro_base)
                            if _incluir_em_resumo_por_serie(res, is_p, cnpj_limpo):
                                if sk not in audit_map:
                                    audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                                audit_map[sk]["nums"].add(res["Número"])
                                if _cancel_sem_xml_manual(res):
                                    audit_map[sk]["nums_buraco"].add(res["Número"])
                                elif incluir_numero_no_conjunto_buraco(
                                    res["Ano"],
                                    res["Mes"],
                                    res["Número"],
                                    ref_ar,
                                    ref_mr,
                                    ult_u,
                                ):
                                    audit_map[sk]["nums_buraco"].add(res["Número"])
                                audit_map[sk]["valor"] += res["Valor"]

            res_final = []
            fal_final = []
            
            for (t, s), dados in audit_map.items():
                ns = sorted(list(dados["nums"]))
                if ns:
                    n_min = ns[0]
                    n_max = ns[-1]
                    res_final.append({
                        "Documento": t, 
                        "Série": s, 
                        "Início": n_min, 
                        "Fim": n_max, 
                        "Quantidade": len(ns), 
                        "Valor Contábil (R$)": round(dados["valor"], 2)
                    })
                ult_lookup = ultimo_ref_lookup(ref_map, t, s) if ref_ar is not None else None
                fal_final.extend(
                    falhas_buraco_por_serie(dados["nums_buraco"], t, s, ult_lookup)
                )

            st.session_state.update({
                'relatorio': rel_list,
                'df_resumo': pd.DataFrame(res_final), 
                'df_faltantes': pd.DataFrame(fal_final), 
                'df_canceladas': pd.DataFrame(canc_list), 
                'df_inutilizadas': pd.DataFrame(inut_list), 
                'df_autorizadas': pd.DataFrame(aut_list), 
                'df_denegadas': pd.DataFrame(deneg_list),
                'df_rejeitadas': pd.DataFrame(rej_list),
                'df_geral': pd.DataFrame(geral_list),
                'st_counts': {
                    "CANCELADOS": len(canc_list), 
                    "INUTILIZADOS": len(inut_list), 
                    "AUTORIZADAS": len(aut_list),
                    "DENEGADOS": len(deneg_list),
                    "REJEITADOS": len(rej_list),
                }, 
                'garimpo_ok': True, 
                'export_ready': False,
                'excel_buffer': None,
            })
            aplicar_compactacao_dfs_sessao()
            st.rerun()
    else:
        # --- RESULTADOS TELA INICIAL (cartões por série só no PDF; aqui só tabela de resumo e abas) ---
        _gcm, _gcr = st.columns([2.95, 1.55], gap="large")
        with _gcm:
            st.markdown(
                f'<h3 class="garim-sec">{_garim_emoji("📤")} Emissões próprias — total por tipo</h3>',
                unsafe_allow_html=True,
            )
            _df_rp = st.session_state.get("df_resumo")
            if isinstance(_df_rp, pd.DataFrame) and not _df_rp.empty and "Quantidade" in _df_rp.columns:
                _tot_prop = int(round(float(pd.to_numeric(_df_rp["Quantidade"], errors="coerce").fillna(0).sum())))
            else:
                _tot_prop = 0
            st.caption(
                f"Total de documentos de emissão própria lidos (soma por tipo): {_tot_prop}"
            )
            st.dataframe(
                _df_resumo_para_exibicao_sem_separador_milhar(st.session_state["df_resumo"]),
                width="stretch",
                hide_index=True,
            )

            st.markdown(
                f'<h3 class="garim-sec">{_garim_emoji("📥")} Terceiros — total por tipo</h3>',
                unsafe_allow_html=True,
            )
            _rels_terc = [
                r
                for r in st.session_state["relatorio"]
                if "RECEBIDOS_TERCEIROS" in r.get("Pasta", "")
            ]
            if not _rels_terc:
                st.info("Nenhum XML de terceiros no lote.")
            else:
                _cnt_terc = Counter((r.get("Tipo") or "Outros") for r in _rels_terc)
                _df_terc = pd.DataFrame(
                    [{"Modelo": t, "Quantidade": n} for t, n in sorted(_cnt_terc.items(), key=lambda x: x[0])]
                )
                _tot_terc = int(round(float(_df_terc["Quantidade"].sum())))
                st.caption(
                    f"Total de documentos de terceiros lidos (soma por tipo): {_tot_terc}"
                )
                st.dataframe(
                    _df_terceiros_por_tipo_para_exibicao_sem_separador_milhar(_df_terc),
                    width="stretch",
                    hide_index=True,
                )

            st.markdown("---")
            st.markdown(
                f'<h3 class="garim-sec">{_garim_emoji("📊")} Relatório da leitura</h3>',
                unsafe_allow_html=True,
            )
            df_fal = st.session_state["df_faltantes"]
            df_inu = st.session_state["df_inutilizadas"]
            df_can = st.session_state["df_canceladas"]
            df_aut = st.session_state["df_autorizadas"]
            df_den = st.session_state.get("df_denegadas")
            if not isinstance(df_den, pd.DataFrame):
                df_den = pd.DataFrame()
            df_rej = st.session_state.get("df_rejeitadas")
            if not isinstance(df_rej, pd.DataFrame):
                df_rej = pd.DataFrame()
            df_ger = st.session_state.get("df_geral")
            if not isinstance(df_ger, pd.DataFrame):
                df_ger = pd.DataFrame()
            if df_ger.empty or "Origem" not in df_ger.columns:
                df_ger_p = df_ger.copy()
            else:
                df_ger_p = df_ger.loc[_mask_emissao_propria_df(df_ger)].reset_index(drop=True)
            df_ger_t = _df_apenas_terceiros(df_ger)
            _folhas_t = _folhas_detalhe_terceiros_do_subset(df_ger)

            with st.expander(
                "🏢 Emissão própria — buracos, inutilizadas, canceladas, autorizadas, denegadas, rejeitadas e relatório geral",
                expanded=True,
            ):
                st.caption(
                    "Só linhas da **sua empresa** (emitente = CNPJ da barra lateral). Use a setinha para recolher este bloco."
                )
                _n_bur = len(df_fal) if not df_fal.empty else 0
                _n_inu = len(df_inu) if not df_inu.empty else 0
                _n_can = len(df_can) if not df_can.empty else 0
                _n_aut = len(df_aut) if not df_aut.empty else 0
                _n_den = len(df_den) if not df_den.empty else 0
                _n_rej = len(df_rej) if not df_rej.empty else 0
                _n_ger_p = len(df_ger_p) if not df_ger_p.empty else 0
                tab_bur, tab_inut, tab_canc, tab_aut, tab_den, tab_rej, tab_geral = st.tabs(
                    [
                        f"⚠️ Buracos ({_n_bur})",
                        f"🚫 Inutilizadas ({_n_inu})",
                        f"❌ Canceladas ({_n_can})",
                        f"✅ Autorizadas ({_n_aut})",
                        f"⛔ Denegadas ({_n_den})",
                        f"📛 Rejeitadas ({_n_rej})",
                        f"📋 Relatório geral ({_n_ger_p})",
                    ]
                )

                with tab_bur:
                    if not df_fal.empty:
                        if {"Tipo", "Série"}.issubset(df_fal.columns):
                            _df_rb = (
                                df_fal.groupby(["Tipo", "Série"], dropna=False)
                                .size()
                                .reset_index(name="Qtd. buracos")
                                .sort_values(["Tipo", "Série"], kind="stable")
                                .reset_index(drop=True)
                            )
                            _df_rb["Qtd. buracos"] = _df_rb["Qtd. buracos"].map(
                                lambda x: str(int(x)) if pd.notna(x) else ""
                            )
                            st.markdown("**Resumo nos buracos** — quantidade de números em falta por modelo e série")
                            st.dataframe(
                                _df_rb,
                                width="stretch",
                                hide_index=True,
                                height=min(260, 44 + 28 * max(1, len(_df_rb))),
                            )
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna** (menu do filtro no título). "
                            "**Excel** e **ZIP XML** usam só as linhas visíveis na grelha."
                        )
                        df_b_f = _relatorio_leitura_tabela_aggrid(df_fal, "aggrid_rep_bur", height=420)
                        st.caption(f"**{len(df_b_f)}** linha(s) na vista (total na aba: {len(df_fal)}).")
                        xlsx_b = _excel_bytes_memo("rep_bur", df_b_f, "Buracos")
                        if xlsx_b:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_b,
                                file_name="relatorio_buracos.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_buracos_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_bur", df_b_f, cnpj_limpo, df_ger)
                    else:
                        st.info("✅ Tudo em ordem.")
                    # <details> em vez de st.expander: o Streamlit não permite expander dentro de expander.
                    st.markdown(
                        """
<details class="garim-detalhe-ajuda">
<summary><strong>Como funcionam os buracos e a referência na lateral (último nº / mês)</strong></summary>
<div style="font-size:0.9rem;line-height:1.55;color:#444;padding:0.45rem 0 0.15rem 0;">
<p style="margin:0;">Só <b>emissão própria</b>. O <b>Resumo por série</b> e esta lista de <b>buracos</b> incluem <b>NF-e</b>, <b>NFC-e</b> e <b>NFS-e</b> em que o emitente é o <b>CNPJ da barra lateral</b> (outros modelos não entram nestes quadros). Aqui: <b>números em falta</b> na sequência.
Com <b>Guardar referência</b> na lateral (mês + último nº por série), cada série indicada ignora XMLs de <b>meses antes</b> desse mês e
lista buracos <b>a partir do último nº + 1</b> — evita buraco gigante se aparecer uma nota fora da ordem (ex. janeiro no meio de março).
Séries <b>não</b> listadas na referência: buracos em <b>todo</b> o intervalo dos XMLs. <b>Sem</b> referência guardada: mesmo comportamento antigo (intervalo completo; pode ser enorme).
Na <b>Etapa 3</b> escolhe o que exportar.</p>
</div>
</details>
""",
                        unsafe_allow_html=True,
                    )

                with tab_inut:
                    if not df_inu.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_i_f = _relatorio_leitura_tabela_aggrid(df_inu, "aggrid_rep_inu", height=420)
                        st.caption(f"**{len(df_i_f)}** linha(s) na vista (total: {len(df_inu)}).")
                        xlsx_i = _excel_bytes_memo("rep_inu", df_i_f, "Inutilizadas")
                        if xlsx_i:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_i,
                                file_name="relatorio_inutilizadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_inut_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_inu", df_i_f, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma nota.")

                with tab_canc:
                    if not df_can.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_c_f = _relatorio_leitura_tabela_aggrid(df_can, "aggrid_rep_canc", height=420)
                        st.caption(f"**{len(df_c_f)}** linha(s) na vista (total: {len(df_can)}).")
                        xlsx_c = _excel_bytes_memo("rep_canc", df_c_f, "Canceladas")
                        if xlsx_c:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_c,
                                file_name="relatorio_canceladas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_canc_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_canc", df_c_f, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma nota.")

                with tab_aut:
                    if not df_aut.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_a_f = _relatorio_leitura_tabela_aggrid(df_aut, "aggrid_rep_aut", height=420)
                        st.caption(f"**{len(df_a_f)}** linha(s) na vista (total: {len(df_aut)}).")
                        xlsx_a = _excel_bytes_memo("rep_aut", df_a_f, "Autorizadas")
                        if xlsx_a:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_a,
                                file_name="relatorio_autorizadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_aut_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_aut", df_a_f, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma nota autorizada na amostra.")

                with tab_den:
                    if not df_den.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_d_f = _relatorio_leitura_tabela_aggrid(df_den, "aggrid_rep_den", height=420)
                        st.caption(f"**{len(df_d_f)}** linha(s) na vista (total: {len(df_den)}).")
                        xlsx_d = _excel_bytes_memo("rep_den", df_d_f, "Denegadas")
                        if xlsx_d:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_d,
                                file_name="relatorio_denegadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_den_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_den", df_d_f, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma denegada neste lote.")

                with tab_rej:
                    if not df_rej.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_r_f = _relatorio_leitura_tabela_aggrid(df_rej, "aggrid_rep_rej", height=420)
                        st.caption(f"**{len(df_r_f)}** linha(s) na vista (total: {len(df_rej)}).")
                        xlsx_r = _excel_bytes_memo("rep_rej", df_r_f, "Rejeitadas")
                        if xlsx_r:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_r,
                                file_name="relatorio_rejeitadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_rej_xlsx",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_rej", df_r_f, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma rejeitada neste lote.")

                with tab_geral:
                    if not df_ger_p.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. Com **todas** as linhas visíveis, o **Excel** pode trazer o livro completo + dashboard "
                            "(folhas alinhadas a estas abas, **só emissão própria** no «Geral»); com filtro na grelha, só a folha **Filtrado**."
                        )
                        df_g_f = _relatorio_leitura_tabela_aggrid(df_ger_p, "aggrid_rep_ger", height=480)
                        _sig_f = _df_sig_hash_memo(df_g_f)
                        _sig_full = _df_sig_hash_memo(df_ger_p)
                        _full_vista = _sig_f == _sig_full and len(df_g_f) == len(df_ger_p)
                        st.caption(f"**{len(df_g_f)}** linha(s) na vista (total: {len(df_ger_p)}).")
                        if _full_vista:
                            sk_wb = "_xlsx_mem_geral_workbook"
                            prev_wb = st.session_state.get(sk_wb)
                            if isinstance(prev_wb, tuple) and prev_wb[0] == _sig_full:
                                xlsx_g = prev_wb[1]
                            else:
                                xlsx_g = excel_relatorio_geral_com_dashboard_bytes(df_ger_p)
                                if xlsx_g:
                                    st.session_state[sk_wb] = (_sig_full, xlsx_g)
                        else:
                            xlsx_g = _excel_bytes_memo("rep_ger_filt", df_g_f, "Filtrado")
                        if xlsx_g:
                            st.download_button(
                                "Baixar Excel (completo + dashboard)" if _full_vista else "Baixar Excel (só filtrado)",
                                data=xlsx_g,
                                file_name="relatorio_geral.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_geral_xlsx",
                                width="stretch",
                            )
                        elif _full_vista:
                            st.warning(_msg_sem_espaco_disco_garimpeiro())
                        _painel_zip_xml_filtrado("rep_ger", df_g_f, cnpj_limpo, df_ger)
                    else:
                        st.info("Sem linhas de emissão própria no relatório geral.")

            with st.expander(
                "🤝 Terceiros — XML recebidos (canceladas, autorizadas, denegadas, rejeitadas e relatório geral)",
                expanded=True,
            ):
                st.caption(
                    "Só linhas **recebidas de terceiros**. Buracos e inutilizadas não se aplicam aqui (só **emissão própria**, no bloco acima). Use a setinha para recolher."
                )
                _tca = _folhas_t["df_can"]
                _tau = _folhas_t["df_aut"]
                _tde = _folhas_t["df_den"]
                _tre = _folhas_t["df_rej"]
                _n_ger_t = len(df_ger_t) if not df_ger_t.empty else 0
                tab_canc_t, tab_aut_t, tab_den_t, tab_rej_t, tab_ger_t = st.tabs(
                    [
                        f"❌ Canceladas ({len(_tca) if not _tca.empty else 0})",
                        f"✅ Autorizadas ({len(_tau) if not _tau.empty else 0})",
                        f"⛔ Denegadas ({len(_tde) if not _tde.empty else 0})",
                        f"📛 Rejeitadas ({len(_tre) if not _tre.empty else 0})",
                        f"📋 Relatório geral ({_n_ger_t})",
                    ]
                )

                with tab_canc_t:
                    if not _tca.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_c_ft = _relatorio_leitura_tabela_aggrid(_tca, "aggrid_rep_canc_t", height=420)
                        st.caption(f"**{len(df_c_ft)}** linha(s) na vista (total: {len(_tca)}).")
                        xlsx_ct = _excel_bytes_memo("rep_canc_t", df_c_ft, "Canceladas")
                        if xlsx_ct:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_ct,
                                file_name="relatorio_terceiros_canceladas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_canc_xlsx_t",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_canc_t", df_c_ft, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma cancelada de terceiros.")

                with tab_aut_t:
                    if not _tau.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_a_ft = _relatorio_leitura_tabela_aggrid(_tau, "aggrid_rep_aut_t", height=420)
                        st.caption(f"**{len(df_a_ft)}** linha(s) na vista (total: {len(_tau)}).")
                        xlsx_at = _excel_bytes_memo("rep_aut_t", df_a_ft, "Autorizadas")
                        if xlsx_at:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_at,
                                file_name="relatorio_terceiros_autorizadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_aut_xlsx_t",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_aut_t", df_a_ft, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma autorizada (normal) de terceiros na amostra.")

                with tab_den_t:
                    if not _tde.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_d_ft = _relatorio_leitura_tabela_aggrid(_tde, "aggrid_rep_den_t", height=420)
                        st.caption(f"**{len(df_d_ft)}** linha(s) na vista (total: {len(_tde)}).")
                        xlsx_dt = _excel_bytes_memo("rep_den_t", df_d_ft, "Denegadas")
                        if xlsx_dt:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_dt,
                                file_name="relatorio_terceiros_denegadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_den_xlsx_t",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_den_t", df_d_ft, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma denegada de terceiros.")

                with tab_rej_t:
                    if not _tre.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. **Excel** e **ZIP XML** — só linhas visíveis."
                        )
                        df_r_ft = _relatorio_leitura_tabela_aggrid(_tre, "aggrid_rep_rej_t", height=420)
                        st.caption(f"**{len(df_r_ft)}** linha(s) na vista (total: {len(_tre)}).")
                        xlsx_rt = _excel_bytes_memo("rep_rej_t", df_r_ft, "Rejeitadas")
                        if xlsx_rt:
                            st.download_button(
                                "Baixar Excel (vista filtrada)",
                                data=xlsx_rt,
                                file_name="relatorio_terceiros_rejeitadas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_rej_xlsx_t",
                                width="stretch",
                            )
                        _painel_zip_xml_filtrado("rep_rej_t", df_r_ft, cnpj_limpo, df_ger)
                    else:
                        st.info("ℹ️ Nenhuma rejeitada de terceiros.")

                with tab_ger_t:
                    if not df_ger_t.empty:
                        st.caption(
                            "Filtre e ordene **no cabeçalho de cada coluna**. Com **todas** as linhas visíveis, o **Excel** traz livro + dashboard com folhas **só terceiros**; "
                            "com filtro na grelha, só a folha **Filtrado**."
                        )
                        df_gt_f = _relatorio_leitura_tabela_aggrid(df_ger_t, "aggrid_rep_ger_t", height=480)
                        _sig_ft = _df_sig_hash_memo(df_gt_f)
                        _sig_full_t = _df_sig_hash_memo(df_ger_t)
                        _full_vista_t = _sig_ft == _sig_full_t and len(df_gt_f) == len(df_ger_t)
                        st.caption(f"**{len(df_gt_f)}** linha(s) na vista (total: {len(df_ger_t)}).")
                        if _full_vista_t:
                            sk_wb_t = "_xlsx_mem_geral_workbook_terc"
                            prev_wb_t = st.session_state.get(sk_wb_t)
                            _fd_t = _folhas_detalhe_terceiros_do_subset(df_ger)
                            if isinstance(prev_wb_t, tuple) and prev_wb_t[0] == _sig_full_t:
                                xlsx_gt = prev_wb_t[1]
                            else:
                                xlsx_gt = excel_relatorio_geral_com_dashboard_bytes(
                                    df_ger_t, folhas_detalhe=_fd_t
                                )
                                if xlsx_gt:
                                    st.session_state[sk_wb_t] = (_sig_full_t, xlsx_gt)
                        else:
                            xlsx_gt = _excel_bytes_memo("rep_ger_filt_t", df_gt_f, "Filtrado")
                        if xlsx_gt:
                            st.download_button(
                                "Baixar Excel (completo + dashboard)" if _full_vista_t else "Baixar Excel (só filtrado)",
                                data=xlsx_gt,
                                file_name="relatorio_geral_terceiros.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rep_geral_xlsx_t",
                                width="stretch",
                            )
                        elif _full_vista_t:
                            st.warning(_msg_sem_espaco_disco_garimpeiro())
                        _painel_zip_xml_filtrado("rep_ger_t", df_gt_f, cnpj_limpo, df_ger)
                    else:
                        st.info("Nenhum documento de terceiros no relatório geral.")

        with _gcr:
            with st.container(border=True):
                st.markdown(
                    f'<h5>{_garim_emoji("📤")} Uploads e validação</h5>',
                    unsafe_allow_html=True,
                )
                st.caption(
                    "Configure **abaixo** os ficheiros, **inutilizações** e **canceladas** manuais; **um único botão** em baixo (**Processar Dados**) grava, aplica e recalcula tudo."
                )
                # MÓDULO: DOCUMENTOS XML/ZIP (carga incremental)
                # =====================================================================
                st.markdown(
                    f'<h5>{_garim_emoji("📎")} Documentos XML / ZIP para ler</h5>',
                    unsafe_allow_html=True,
                )
                with st.expander("➕ Incluir mais ficheiros no lote (sem resetar)", expanded=False):
                    extra_files = st.file_uploader(
                        "Escolha os XML ou ZIP a acrescentar ao lote atual:",
                        accept_multiple_files=True,
                        key="extra_files",
                    )
                    _garimpo_absorver_uploads_extra_no_lote(extra_files)
                    st.caption(
                        "Os ficheiros passam **para o lote ao serem escolhidos** (e de novo em **Processar Dados** se ainda estiverem no campo); o relatório é recalculado a partir do lote completo."
                    )

                # =====================================================================
                # MÓDULO: AUTENTICIDADE — mesmo nível visual que Inutilizadas
                # =====================================================================
                st.markdown(
                    f'<h5>{_garim_emoji("🔐")} Validação de autenticidade</h5>',
                    unsafe_allow_html=True,
                )
                with st.expander(
                    "Relatório exportado da Sefaz para confrontar com o lote de XML (opcional).",
                    expanded=False,
                ):
                    st.caption(
                        "Use o Excel ou CSV que o **portal da Sefaz** disponibiliza (lista de **autorizadas** / emitidas). "
                        "Pode anexar **vários ficheiros** (ex.: uma planilha por série ou por período) — as linhas são **juntadas** antes da comparação. "
                        "A app compara **Modelo**, **Série** e **Nota** com as NF **autorizadas** (NORMAIS) de **emissão própria** nos XML. "
                        "As **divergências** aparecem aqui em baixo, no **Excel** do relatório geral e no **PDF** do dashboard. "
                        "Depois de anexar, carregue em **Processar Dados** (pode ser só isto ou juntamente com XML / inutilizações)."
                    )
                    st.file_uploader(
                        "Anexar relatório(is) Sefaz (.csv, .xlsx ou .xls)",
                        type=["csv", "xlsx", "xls"],
                        accept_multiple_files=True,
                        key="autent_sefaz_up",
                    )
                    st.caption(
                        "Colunas reconhecidas (cabeçalho), iguais em todos os ficheiros: **Modelo** (55, 65, 57, 67, 58 ou NF-e…), **Série**, **Nota** / Número — "
                        "o mesmo critério da planilha de inutilizadas. Linhas repetidas em ficheiros diferentes contam uma vez na comparação."
                    )
                    _bytes_m_aut = bytes_modelo_planilha_inutil_sem_xml_xlsx()
                    if _bytes_m_aut:
                        st.download_button(
                            "Baixar modelo de colunas (Excel)",
                            data=_bytes_m_aut,
                            file_name="MODELO_autenticidade_Sefaz_garimpeiro.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_modelo_autent_xlsx",
                            width="stretch",
                        )
                    else:
                        st.warning(_msg_sem_espaco_disco_garimpeiro())
                    _df_div_ui = st.session_state.get("df_divergencias")
                    if _df_div_ui is not None and isinstance(_df_div_ui, pd.DataFrame) and not _df_div_ui.empty:
                        st.caption(f"**{len(_df_div_ui)}** divergência(s) na última comparação.")
                        st.dataframe(_df_div_ui, width="stretch", height=min(260, 40 + 24 * len(_df_div_ui)))
                    elif st.session_state.get("validation_done"):
                        st.success("Última comparação de autenticidade: **sem divergências** (ou planilha vazia após cruzamento).")

                # =====================================================================
                # MÓDULO: DECLARAR INUTILIZADAS MANUAIS
                # =====================================================================
                st.markdown(
                    f'<h5>{_garim_emoji("🛠️")} Inutilizadas</h5>',
                    unsafe_allow_html=True,
                )
                with st.expander(
                    "Inclua notas que a Sefaz mostra inutilizadas mas que não estão no lote de ficheiros.",
                    expanded=False,
                ):
                    st.caption(
                        "Só vale para **buracos** já listados. **Dos buracos**, **planilha** ou **faixa**: preencha o separador pretendido — tudo é aplicado ao carregar em **Processar Dados** em baixo."
                    )
                    tab_b, tab_p, tab_f = st.tabs(["Dos buracos", "Planilha (Excel/CSV)", "Faixa de números"])

                    with tab_b:
                        df_b = st.session_state["df_faltantes"].copy()
                        if not df_b.empty and "Serie" in df_b.columns and "Série" not in df_b.columns:
                            df_b = df_b.rename(columns={"Serie": "Série"})
                        if df_b.empty:
                            st.info("Sem buracos na auditoria — faça o garimpo primeiro ou verifique a referência de último nº.")
                        elif not {"Tipo", "Série", "Num_Faltante"}.issubset(df_b.columns):
                            st.warning(
                                "A tabela de buracos não tem o formato esperado (Tipo, Série, Num_Faltante). "
                                "Faça **Novo garimpo** para recalcular."
                            )
                        else:
                            _mods_b = sorted(df_b["Tipo"].astype(str).unique())
                            _mb = st.selectbox("Modelo", _mods_b, key="inut_b_mod")
                            _sub_b = df_b[df_b["Tipo"].astype(str) == _mb]
                            _sers_b = sorted(_sub_b["Série"].astype(str).unique())
                            _sb = st.selectbox("Série", _sers_b, key="inut_b_ser")
                            _sub2_b = _sub_b[_sub_b["Série"].astype(str) == _sb]
                            _nums_b = sorted(_sub2_b["Num_Faltante"].astype(int).unique())
                            _set_buracos = set(_nums_b)
                            st.caption(f"{len(_nums_b)} buraco(s) neste modelo/série — só estes podem ser declarados aqui.")
                            _pick_b = st.multiselect(
                                "Marque os que quer tratar como inutilizados:",
                                options=_nums_b,
                                format_func=lambda x: f"Nota n.º {x}",
                                key="inut_b_pick",
                            )

                    with tab_p:
                        st.markdown("**Subir tabela** com inutilizadas a declarar")
                        st.caption(
                            "Colunas (1.ª linha = cabeçalho): **Modelo** = código Sefaz (**55** NF-e, **65** NFC-e, **57** CT-e, **67** CT-e OS, **58** MDF-e) "
                            "ou nome NF-e / NFC-e…; **Série**; **Nota** (ou Número / Num_Faltante). "
                            "Ideal para copiar/colar da Sefaz. Só entram linhas que já forem **buraco** no garimpeiro."
                        )
                        _bytes_m_inut = bytes_modelo_planilha_inutil_sem_xml_xlsx()
                        if _bytes_m_inut:
                            st.download_button(
                                "Baixar Excel",
                                data=_bytes_m_inut,
                                file_name="MODELO_inutilizadas_sem_XML_garimpeiro.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_modelo_inut_xlsx",
                                width="stretch",
                            )
                        else:
                            st.warning(_msg_sem_espaco_disco_garimpeiro())
                        st.caption(
                            "No modelo: **Modelo** em número (55, 65, 57, 67, 58) como na Sefaz; **Série** e **Nota**. "
                            "Substitua ou apague as linhas de exemplo e guarde antes de importar."
                        )
                        _up_inut = st.file_uploader(
                            "Ficheiro .csv, .xlsx ou .xls",
                            type=["csv", "xlsx", "xls"],
                            key="inut_planilha_up",
                        )

                    with tab_f:
                        _mf = st.selectbox(
                            "Modelo",
                            ["NF-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e"],
                            key="inut_f_mod",
                        )
                        _sf = st.text_input("Série", value="1", key="inut_f_ser").strip()
                        _c1f, _c2f = st.columns(2)
                        _n0 = _c1f.number_input("Nota inicial", min_value=1, value=1, step=1, key="inut_f_i")
                        _n1 = _c2f.number_input("Nota final", min_value=1, value=1, step=1, key="inut_f_f")
                        _MAX_FAIXA_INUT = 5000
                        df_fb = st.session_state["df_faltantes"].copy()
                        if not df_fb.empty and "Serie" in df_fb.columns and "Série" not in df_fb.columns:
                            df_fb = df_fb.rename(columns={"Serie": "Série"})
                        _bur_f = set()
                        if (
                            not df_fb.empty
                            and {"Tipo", "Série", "Num_Faltante"}.issubset(df_fb.columns)
                        ):
                            _subf = df_fb[
                                (df_fb["Tipo"].astype(str) == _mf)
                                & (df_fb["Série"].astype(str) == str(_sf).strip())
                            ]
                            _bur_f = set(_subf["Num_Faltante"].astype(int).unique())
                        st.caption(
                            f"No máximo {_MAX_FAIXA_INUT} notas analisadas por vez. "
                            f"Só entram na inutilização manual as que forem **buraco** neste modelo/série "
                            f"({len(_bur_f)} buraco(s) conhecidos). A faixa é aplicada em **Processar Dados**."
                        )

                # =====================================================================
                # MÓDULO: DECLARAR CANCELADAS MANUAIS (sem XML de cancelamento)
                # =====================================================================
                st.markdown(
                    f'<h5>{_garim_emoji("❌")} Canceladas</h5>',
                    unsafe_allow_html=True,
                )
                with st.expander(
                    "Inclua notas que sabe estar canceladas na Sefaz mas não tem o XML de cancelamento no lote.",
                    expanded=False,
                ):
                    st.caption(
                        "Mesmas regras que **Inutilizadas**: só **buracos** já listados. "
                        "**Dos buracos**, **planilha** ou **faixa** — aplicado em **Processar Dados**."
                    )
                    tab_cb, tab_cp, tab_cf = st.tabs(["Dos buracos", "Planilha (Excel/CSV)", "Faixa de números"])

                    with tab_cb:
                        df_cb = st.session_state["df_faltantes"].copy()
                        if not df_cb.empty and "Serie" in df_cb.columns and "Série" not in df_cb.columns:
                            df_cb = df_cb.rename(columns={"Serie": "Série"})
                        if df_cb.empty:
                            st.info("Sem buracos na auditoria — faça o garimpo primeiro ou verifique a referência de último nº.")
                        elif not {"Tipo", "Série", "Num_Faltante"}.issubset(df_cb.columns):
                            st.warning(
                                "A tabela de buracos não tem o formato esperado (Tipo, Série, Num_Faltante). "
                                "Faça **Novo garimpo** para recalcular."
                            )
                        else:
                            _mods_c = sorted(df_cb["Tipo"].astype(str).unique())
                            _mbc = st.selectbox("Modelo", _mods_c, key="canc_b_mod")
                            _sub_c = df_cb[df_cb["Tipo"].astype(str) == _mbc]
                            _sers_c = sorted(_sub_c["Série"].astype(str).unique())
                            _sbc = st.selectbox("Série", _sers_c, key="canc_b_ser")
                            _sub2_c = _sub_c[_sub_c["Série"].astype(str) == _sbc]
                            _nums_c = sorted(_sub2_c["Num_Faltante"].astype(int).unique())
                            st.caption(f"{len(_nums_c)} buraco(s) neste modelo/série — só estes podem ser declarados como cancelados.")
                            st.multiselect(
                                "Marque os que quer tratar como cancelados:",
                                options=_nums_c,
                                format_func=lambda x: f"Nota n.º {x}",
                                key="canc_b_pick",
                            )

                    with tab_cp:
                        st.markdown("**Subir tabela** com canceladas a declarar")
                        st.caption(
                            "Colunas: **Modelo** (55, 65, 57, 67, 58 ou NF-e…), **Série**, **Nota**. "
                            "Só entram linhas que já forem **buraco** no garimpeiro."
                        )
                        _bytes_m_canc = bytes_modelo_planilha_cancel_sem_xml_xlsx()
                        if _bytes_m_canc:
                            st.download_button(
                                "Baixar modelo Excel",
                                data=_bytes_m_canc,
                                file_name="MODELO_canceladas_sem_XML_garimpeiro.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_modelo_canc_xlsx",
                                width="stretch",
                            )
                        else:
                            st.warning(_msg_sem_espaco_disco_garimpeiro())
                        st.file_uploader(
                            "Ficheiro .csv, .xlsx ou .xls",
                            type=["csv", "xlsx", "xls"],
                            key="canc_planilha_up",
                        )

                    with tab_cf:
                        _mfc_ui = st.selectbox(
                            "Modelo",
                            ["NF-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e"],
                            key="canc_f_mod",
                        )
                        _sfc_ui = st.text_input("Série", value="1", key="canc_f_ser").strip()
                        _c1fc, _c2fc = st.columns(2)
                        _c1fc.number_input("Nota inicial", min_value=1, value=1, step=1, key="canc_f_i")
                        _c2fc.number_input("Nota final", min_value=1, value=1, step=1, key="canc_f_f")
                        _MAX_FAIXA_CANC_UI = 5000
                        df_cfb = st.session_state["df_faltantes"].copy()
                        if not df_cfb.empty and "Serie" in df_cfb.columns and "Série" not in df_cfb.columns:
                            df_cfb = df_cfb.rename(columns={"Serie": "Série"})
                        _bur_cfu = set()
                        if (
                            not df_cfb.empty
                            and {"Tipo", "Série", "Num_Faltante"}.issubset(df_cfb.columns)
                        ):
                            _subcf = df_cfb[
                                (df_cfb["Tipo"].astype(str) == _mfc_ui)
                                & (df_cfb["Série"].astype(str) == str(_sfc_ui).strip())
                            ]
                            _bur_cfu = set(_subcf["Num_Faltante"].astype(int).unique())
                        st.caption(
                            f"No máximo {_MAX_FAIXA_CANC_UI} notas. "
                            f"Só entram como canceladas manuais as que forem **buraco** neste modelo/série "
                            f"({len(_bur_cfu)} buraco(s) conhecidos)."
                        )

                # =====================================================================
                # MÓDULO: DESFAZER REGISTO MANUAL (inutil. / cancel.)
                # =====================================================================
                _manuais_undo = [
                    item
                    for item in st.session_state["relatorio"]
                    if item.get("Arquivo") in ("REGISTRO_MANUAL", "REGISTRO_MANUAL_CANCELADO")
                ]
                if _manuais_undo:
                    with st.expander("🔙 Desfazer registo manual (inutil. / cancel.)", expanded=False):
                        _df_m = pd.DataFrame(
                            [
                                {
                                    "Chave": i["Chave"],
                                    "Tipo": i["Tipo"],
                                    "Série": str(i["Série"]),
                                    "Nota": i["Número"],
                                    "Estado": "Inutilizada"
                                    if i.get("Status") == "INUTILIZADOS"
                                    else "Cancelada",
                                }
                                for i in _manuais_undo
                            ]
                        )
                        _dm = sorted(_df_m["Tipo"].astype(str).unique())
                        _mdes = st.selectbox("Modelo", _dm, key="desf_man_mod")
                        _sub_d = _df_m[_df_m["Tipo"].astype(str) == _mdes]
                        _dsers = sorted(_sub_d["Série"].astype(str).unique())
                        _sdes = st.selectbox("Série", _dsers, key="desf_man_ser")
                        _sub2_d = _sub_d[_sub_d["Série"].astype(str) == _sdes].sort_values("Nota")
                        _rotulos = {
                            row["Chave"]: f"[{row['Estado']}] Nota n.º {int(row['Nota'])}"
                            for _, row in _sub2_d.iterrows()
                        }
                        _chaves_sel = st.multiselect(
                            "Remover da lista de registos manuais:",
                            options=list(_rotulos.keys()),
                            format_func=lambda k: _rotulos.get(k, k),
                            key="desf_man_pick",
                        )
                        if st.button("Remover seleção e atualizar tabelas", key="desf_man_btn"):
                            if _chaves_sel:
                                with st.spinner("A remover…"):
                                    _set_rem = set(_chaves_sel)
                                    st.session_state["relatorio"] = [
                                        i for i in st.session_state["relatorio"] if i["Chave"] not in _set_rem
                                    ]
                                    reconstruir_dataframes_relatorio_simples()
                                st.rerun()
                            else:
                                st.warning("Selecione pelo menos um registo.")

                st.divider()

                st.markdown(
                    f'<h5>{_garim_emoji("🔁")} Processar Dados</h5>',
                    unsafe_allow_html=True,
                )
                if st.button(
                    "Processar Dados",
                    key="btn_reprocessar_garimpo",
                    width="stretch",
                ):
                    # Usar o retorno do file_uploader (variável `extra_files`): em alguns Streamlit,
                    # `st.session_state["extra_files"]` fica vazio no mesmo run do clique e os XML extra não entram no lote.
                    _ef = extra_files if extra_files else st.session_state.get("extra_files")
                    if _ef is not None and not isinstance(_ef, (list, tuple)):
                        _ef = [_ef]
                    _pick = list(st.session_state.get("inut_b_pick") or [])
                    _mb = st.session_state.get("inut_b_mod")
                    _sb = st.session_state.get("inut_b_ser")
                    _up_pl = st.session_state.get("inut_planilha_up")
                    _mf = st.session_state.get("inut_f_mod", "NF-e")
                    _sf = st.session_state.get("inut_f_ser", "1")
                    _n0 = int(st.session_state.get("inut_f_i", 1))
                    _n1 = int(st.session_state.get("inut_f_f", 1))
                    _pick_c = list(st.session_state.get("canc_b_pick") or [])
                    _mbc = st.session_state.get("canc_b_mod")
                    _sbc = st.session_state.get("canc_b_ser")
                    _up_canc = st.session_state.get("canc_planilha_up")
                    _mfc = st.session_state.get("canc_f_mod", "NF-e")
                    _sfc = st.session_state.get("canc_f_ser", "1")
                    _n0c = int(st.session_state.get("canc_f_i", 1))
                    _n1c = int(st.session_state.get("canc_f_f", 1))
                    _up_aut = st.session_state.get("autent_sefaz_up")
                    if _up_aut is not None and not isinstance(_up_aut, (list, tuple)):
                        _up_aut = [_up_aut]
                    with st.spinner("A gravar, aplicar inutilizações / canceladas e a recalcular…"):
                        _ok_all, _msg_all, _ = processar_painel_lateral_direito(
                            cnpj_limpo,
                            _ef,
                            _pick,
                            _mb,
                            _sb,
                            _up_pl,
                            _mf,
                            _sf,
                            _n0,
                            _n1,
                            pick_bur_canc=_pick_c,
                            mb_canc=_mbc,
                            sb_canc=_sbc,
                            up_canc_planilha=_up_canc,
                            mf_canc_f=_mfc,
                            sf_canc_f=_sfc,
                            n0_canc_f=_n0c,
                            n1_canc_f=_n1c,
                            up_autent_sefaz=_up_aut,
                        )
                    if _ok_all:
                        st.success(_msg_all)
                    else:
                        st.warning(_msg_all)
                    st.rerun()

        st.divider()
        with st.expander(
            "📦 Etapa 3 — Filtros e exportação (ZIP / Excel)",
            expanded=False,
        ):
            if hasattr(st, "fragment"):
                st.fragment(_garim_etapa3_fragment_entry)()
            else:
                _garim_etapa3_fragment_entry()

        if st.button("⛏️ NOVO GARIMPO / LIMPAR TUDO"):
            limpar_arquivos_temp(); st.session_state.clear(); st.rerun()

        # =====================================================================
        # BLOCO 4: EXPORTAR LISTA ESPECÍFICA
        # =====================================================================
        st.divider()
        st.markdown(
            f'<h3>{_garim_emoji("🔎")} EXPORTAR LISTA ESPECÍFICA</h3>',
            unsafe_allow_html=True,
        )
        with st.expander(
            "Excel (chaves ou inicial/final/série), período, faixa ou uma nota — gera ZIP(s) com XML do lote"
        ):
            tab_xlsx, tab_xlsx_ns, tab_periodo, tab_faixa, tab_unica = st.tabs(
                [
                    "📊 Excel (chaves)",
                    "📋 Excel (nº e série)",
                    "📅 Período",
                    "🔢 Faixa de notas",
                    "1️⃣ Nota única",
                ]
            )

            with tab_xlsx:
                xlsx_dom = st.file_uploader(
                    "Planilha (.xlsx ou .xls): coluna A = chave de 44 dígitos",
                    type=["xlsx", "xls"],
                    key="xlsx_dom_final",
                )
                if xlsx_dom and st.button("🔎 BUSCAR XMLS NO LOTE (EXCEL)", key="btn_run_dom_xlsx"):
                    with st.spinner("Lendo chaves e organizando arquivos..."):
                        chaves_lidas = extrair_chaves_de_excel(xlsx_dom)
                        if not chaves_lidas:
                            st.warning("⚠️ Nenhuma chave válida (44 dígitos) na primeira coluna.")
                        else:
                            partes, n_xml = escrever_zip_dominio_por_chaves(
                                cnpj_limpo,
                                chaves_lidas,
                                st.session_state.get("df_geral"),
                            )
                            if partes and n_xml > 0:
                                st.session_state["ch_falt_dom"] = chaves_lidas
                                st.session_state["zip_dom_parts"] = partes
                                nl = len(partes)
                                st.success(
                                    f"✅ Sucesso! {len(chaves_lidas)} chave(s) na planilha; {n_xml} XML(s) em "
                                    f"{nl} ZIP(s) (até {MAX_XML_PER_ZIP} XMLs por lote)."
                                )
                            else:
                                st.warning("⚠️ Nenhum XML encontrado no lote para essas chaves.")

            with tab_xlsx_ns:
                xlsx_ns = st.file_uploader(
                    "Planilha (.xlsx ou .xls): **inicial**, **final** e **série** (uma faixa por linha)",
                    type=["xlsx", "xls"],
                    key="xlsx_dom_ns",
                )
                mod_ns = st.selectbox(
                    "Modelo no relatório geral",
                    ["NF-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e"],
                    index=0,
                    key="dom_ns_modelo",
                    help="Deve coincidir com o tipo das linhas no garimpo (coluna Modelo).",
                )
                st.caption(
                    f"Reconhece cabeçalhos como *Inicial*, *Final*, *Série* (ou sem cabeçalho: colunas A, B, C). "
                    f"Até **{_MAX_FAIXA_EXPORT_DOM}** notas por linha; até **{_MAX_CHAVES_EXCEL_FAIXAS}** chaves no total da planilha."
                )
                if xlsx_ns and st.button(
                    "🔎 BUSCAR XMLS NO LOTE (EXCEL Nº E SÉRIE)", key="btn_run_dom_xlsx_ns"
                ):
                    with st.spinner("A ler faixas e cruzar com o relatório geral..."):
                        faixas_ns, ign_ns, err_ns = extrair_faixas_ini_fim_serie_excel(xlsx_ns)
                        if err_ns and not faixas_ns:
                            st.warning(err_ns)
                        else:
                            if ign_ns:
                                st.caption(f"ℹ️ {ign_ns} linha(s) da planilha ignorada(s) (vazias, série em falta ou faixa larga demais).")
                            df_base = st.session_state.get("df_geral")
                            if df_base is None or df_base.empty:
                                st.warning("Relatório geral vazio — faça o garimpo primeiro.")
                            else:
                                ch_ns, cortado_ns = chaves_agregadas_de_excel_faixas(
                                    df_base, faixas_ns, mod_ns
                                )
                                if cortado_ns:
                                    st.warning(
                                        f"⚠️ Limite de {_MAX_CHAVES_EXCEL_FAIXAS} chaves atingido — divida a planilha ou refine as faixas."
                                    )
                                if not ch_ns:
                                    st.warning(
                                        "Nenhuma chave encontrada no relatório geral para essas faixas/modelo/série."
                                    )
                                elif not _garimpo_existem_fontes_xml_lote():
                                    st.error(
                                        "Não há ficheiros do lote (memória ou pasta). Volte a correr o garimpo ou **Incluir mais XML**."
                                    )
                                else:
                                    partes, n_xml = escrever_zip_dominio_por_chaves(
                                        cnpj_limpo, ch_ns, df_base
                                    )
                                    if partes and n_xml > 0:
                                        st.session_state["ch_falt_dom"] = ch_ns
                                        st.session_state["zip_dom_parts"] = partes
                                        nl = len(partes)
                                        st.success(
                                            f"✅ {len(faixas_ns)} linha(s) na planilha → **{len(ch_ns)}** chave(s); "
                                            f"{n_xml} XML(s) em {nl} ZIP(s) (até {MAX_XML_PER_ZIP} por ficheiro)."
                                        )
                                    else:
                                        st.warning(
                                            "⚠️ Há chaves no relatório, mas **nenhum XML** correspondente no lote atual."
                                        )

            with tab_periodo:
                c_p1, c_p2 = st.columns(2)
                d_ini_dom = c_p1.date_input(
                    "Data inicial (emissão)",
                    value=date.today().replace(day=1),
                    key="dom_per_dini",
                )
                d_fim_dom = c_p2.date_input(
                    "Data final (emissão)",
                    value=date.today(),
                    key="dom_per_dfim",
                )
                if st.button("🔎 BUSCAR XMLS NO LOTE (PERÍODO)", key="btn_run_dom_periodo"):
                    di, dfim = d_ini_dom, d_fim_dom
                    if di > dfim:
                        di, dfim = dfim, di
                    df_base = st.session_state.get("df_geral")
                    if df_base is None or df_base.empty:
                        st.warning("Relatório geral vazio — faça o garimpo primeiro.")
                    else:
                        ch_per = chaves_por_periodo_data(df_base, di, dfim)
                        if not ch_per:
                            st.warning(
                                "Nenhuma chave de 44 dígitos no relatório geral para esse intervalo de datas."
                            )
                        elif not _garimpo_existem_fontes_xml_lote():
                            st.error(
                                "Não há ficheiros do lote (memória ou pasta). Volte a correr o garimpo ou **Incluir mais XML**."
                            )
                        else:
                            partes, n_xml = escrever_zip_dominio_por_chaves(
                                cnpj_limpo, ch_per, df_base
                            )
                            if partes and n_xml > 0:
                                st.session_state["ch_falt_dom"] = ch_per
                                st.session_state["zip_dom_parts"] = partes
                                nl = len(partes)
                                st.success(
                                    f"✅ {len(ch_per)} chave(s) no período; {n_xml} XML(s) em "
                                    f"{nl} ZIP(s) (até {MAX_XML_PER_ZIP} por ficheiro)."
                                )
                            else:
                                st.warning(
                                    "⚠️ Há chaves no relatório, mas **nenhum XML** no lote atual "
                                    f"para esse período. Confira se o lote contém esses ficheiros."
                                )

            with tab_faixa:
                mod_f = st.selectbox(
                    "Modelo",
                    ["NF-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e"],
                    index=0,
                    key="dom_faixa_modelo",
                    help="Igual à coluna Modelo do relatório geral (não use 55/65 — isso é o código Sefaz).",
                )
                ser_f = st.text_input("Série", key="dom_faixa_serie")
                cf1, cf2 = st.columns(2)
                n0_f = int(cf1.number_input("Nota inicial", min_value=1, value=1, step=1, key="dom_faixa_n0"))
                n1_f = int(cf2.number_input("Nota final", min_value=1, value=1, step=1, key="dom_faixa_n1"))
                st.caption(f"No máximo {_MAX_FAIXA_EXPORT_DOM} notas por pedido (proteção do sistema).")
                if st.button("🔎 BUSCAR XMLS NO LOTE (FAIXA)", key="btn_run_dom_faixa"):
                    if not str(ser_f).strip():
                        st.warning("Informe a **série**.")
                    else:
                        a, b = n0_f, n1_f
                        if a > b:
                            a, b = b, a
                        if (b - a + 1) > _MAX_FAIXA_EXPORT_DOM:
                            st.warning(
                                f"Reduza a faixa (máximo {_MAX_FAIXA_EXPORT_DOM} notas de uma vez)."
                            )
                        else:
                            df_base = st.session_state.get("df_geral")
                            if df_base is None or df_base.empty:
                                st.warning("Relatório geral vazio — faça o garimpo primeiro.")
                            else:
                                ch_f = chaves_por_faixa_numeracao(
                                    df_base,
                                    mod_f,
                                    str(ser_f).strip(),
                                    a,
                                    b,
                                )
                                if not ch_f:
                                    st.warning(
                                        "Nenhuma nota nessa faixa/modelo/série no relatório geral."
                                    )
                                elif not _garimpo_existem_fontes_xml_lote():
                                    st.error(
                                        "Não há ficheiros do lote (memória ou pasta). Volte a correr o garimpo."
                                    )
                                else:
                                    partes, n_xml = escrever_zip_dominio_por_chaves(
                                        cnpj_limpo, ch_f, df_base
                                    )
                                    if partes and n_xml > 0:
                                        st.session_state["ch_falt_dom"] = ch_f
                                        st.session_state["zip_dom_parts"] = partes
                                        nl = len(partes)
                                        st.success(
                                            f"✅ {len(ch_f)} chave(s); {n_xml} XML(s) em "
                                            f"{nl} ZIP(s) (até {MAX_XML_PER_ZIP} por ficheiro)."
                                        )
                                    else:
                                        st.warning(
                                            "⚠️ Chaves encontradas no relatório, mas **nenhum XML** no lote atual."
                                        )

            with tab_unica:
                mod_u = st.selectbox(
                    "Modelo",
                    ["NF-e", "NFC-e", "CT-e", "CT-e OS", "MDF-e"],
                    index=0,
                    key="dom_unica_modelo",
                    help="Igual à coluna Modelo do relatório geral.",
                )
                ser_u = st.text_input("Série", key="dom_unica_serie")
                nu = int(
                    st.number_input(
                        "Número da nota",
                        min_value=1,
                        value=1,
                        step=1,
                        key="dom_unica_nota",
                    )
                )
                if st.button("🔎 BUSCAR XML NO LOTE (NOTA ÚNICA)", key="btn_run_dom_unica"):
                    if not str(ser_u).strip():
                        st.warning("Informe a **série**.")
                    else:
                        df_base = st.session_state.get("df_geral")
                        if df_base is None or df_base.empty:
                            st.warning("Relatório geral vazio — faça o garimpo primeiro.")
                        else:
                            ch_u = chaves_por_nota_serie(
                                df_base,
                                mod_u,
                                str(ser_u).strip(),
                                nu,
                            )
                            if not ch_u:
                                st.warning(
                                    "Nenhuma linha com esse modelo/série/número no relatório geral. "
                                    "Confirme **Modelo** (NF-e, NFC-e…) como na tabela do garimpo, **série** e **número**; "
                                    "a série no relatório vem sem zeros à esquerda (ex. **1**, não 001)."
                                )
                            elif not _garimpo_existem_fontes_xml_lote():
                                st.error(
                                    "Não há ficheiros do lote (memória ou pasta). Volte a correr o garimpo."
                                )
                            else:
                                partes, n_xml = escrever_zip_dominio_por_chaves(
                                    cnpj_limpo, ch_u, df_base
                                )
                                if partes and n_xml > 0:
                                    st.session_state["ch_falt_dom"] = ch_u
                                    st.session_state["zip_dom_parts"] = partes
                                    nl = len(partes)
                                    st.success(
                                        f"✅ {len(ch_u)} chave(s); {n_xml} XML(s) em "
                                        f"{nl} ZIP(s) (até {MAX_XML_PER_ZIP} por ficheiro)."
                                    )
                                else:
                                    st.warning(
                                        "⚠️ Chave no relatório, mas **nenhum XML** correspondente no lote atual."
                                    )

            if st.session_state.get("zip_dom_parts"):
                st.caption(
                    f"Cada ZIP tem no máximo {MAX_XML_PER_ZIP} XMLs na raiz e um Excel em "
                    "**RELATORIO_GARIMPEIRO/** (`lista_especifica_ptXXX.xlsx`) com modelo, série, nota, chave, status, etc."
                )
                for row in chunk_list(st.session_state["zip_dom_parts"], 3):
                    cols = st.columns(len(row))
                    for idx, part in enumerate(row):
                        if os.path.exists(part):
                            with open(part, "rb") as f_final:
                                cols[idx].download_button(
                                    label=rotulo_download_zip_parte(part),
                                    data=f_final.read(),
                                    file_name=os.path.basename(part),
                                    mime="application/zip",
                                    key=f"btn_dl_dom_{part}",
                                    width="stretch",
                                )
else:
    st.warning("👈 Insira o CNPJ lateral para começar.")

