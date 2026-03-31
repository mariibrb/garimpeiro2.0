import streamlit as st
import zipfile
import io
import os
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
import sys
import json
from pathlib import Path

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

            tipo = "NF-e"
            if re.search(r'<[^>]*nfse', tag_l) or "<nfse" in tag_l or ("servico" in tag_l and "nfse" in tag_l):
                tipo = "NFS-e"
            elif '<mod>65</mod>' in tag_l: 
                tipo = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: 
                tipo = "CT-e"
            elif '<mod>67</mod>' in tag_l or "<dadcte" in tag_l or "dacte" in tag_l:
                tipo = "DACT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: 
                tipo = "MDF-e"
            
            status = "NORMAIS"
            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: 
                status = "CANCELADOS"
            elif '110110' in tag_l: 
                status = "CARTA_CORRECAO"
            elif re.search(r'<cstat>110</cstat>', tag_l) or "deneg" in tag_l:
                status = "REJEITADOS"
            elif re.search(r'<cstat>30[1-9]</cstat>', tag_l) or re.search(r'<cstat>3[1-4]\d</cstat>', tag_l):
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
    if not os.path.exists(TEMP_EXTRACT_DIR): 
        os.makedirs(TEMP_EXTRACT_DIR)
        
    if nome_arquivo.lower().endswith('.zip'):
        try:
            if hasattr(conteudo_ou_file, 'read'):
                file_obj = conteudo_ou_file
            else:
                file_obj = io.BytesIO(conteudo_ou_file)
                
            with zipfile.ZipFile(file_obj) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith('__MACOSX') or os.path.basename(sub_nome).startswith('.'): 
                        continue
                        
                    if sub_nome.lower().endswith('.zip'):
                        temp_path = z.extract(sub_nome, path=TEMP_EXTRACT_DIR)
                        with open(temp_path, 'rb') as f_temp:
                            yield from extrair_recursivo(f_temp, sub_nome)
                        try: 
                            os.remove(temp_path)
                        except: 
                            pass
                    elif sub_nome.lower().endswith('.xml'):
                        yield (os.path.basename(sub_nome), z.read(sub_nome))
        except: 
            pass
            
    elif nome_arquivo.lower().endswith('.xml'):
        if hasattr(conteudo_ou_file, 'read'): 
            yield (os.path.basename(nome_arquivo), conteudo_ou_file.read())
        else: 
            yield (os.path.basename(nome_arquivo), conteudo_ou_file)

# --- LIMPEZA DE PASTAS TEMPORÁRIAS ---
def limpar_arquivos_temp():
    try:
        for f in os.listdir('.'):
            if f.endswith('.zip') and ('z_org_final' in f or 'z_todos_final' in f or 'faltantes_dominio_final' in f):
                try: os.remove(f)
                except: pass
            
        if os.path.exists(TEMP_EXTRACT_DIR): 
            shutil.rmtree(TEMP_EXTRACT_DIR, ignore_errors=True)
            
        if os.path.exists(TEMP_UPLOADS_DIR): 
            shutil.rmtree(TEMP_UPLOADS_DIR, ignore_errors=True)
    except: 
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
    """Excel com as mesmas colunas do DataFrame (para download alinhado à tabela na tela)."""
    if df is None or df.empty:
        return None
    buf = io.BytesIO()
    sn = (sheet_name or "Dados")[:31]
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.reset_index(drop=True).to_excel(writer, sheet_name=sn, index=False)
    return buf.getvalue()


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
        "geral": _preview_df_para_pdf(
            df_g,
            _DASH_PDF_MAX["geral"],
            _DASH_PDF_GERAL_COLS,
            msg_se_vazio="Relatório geral vazio.",
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
        f"Canceladas: {_excel_fmt_milhar_pt(can)}\nInutilizadas: {_excel_fmt_milhar_pt(inu)}",
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

    r1 = r0 + 5
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
            "categories": [sn, r0 + 1, hid_c, r0 + 3, hid_c],
            "values": [sn, r0 + 1, hid_c + 1, r0 + 3, hid_c + 1],
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


def excel_relatorio_geral_com_dashboard_bytes(df_geral):
    """
    Excel com várias folhas alinhadas às abas da página da app:
    Geral, Buracos, Inutilizadas, Canceladas, Autorizadas, CT-e lidas, Terceiros lidas, Dashboard, Painel Fiscal.
    """
    if df_geral is None or df_geral.empty:
        return None
    kpi = coletar_kpis_dashboard()
    buf = io.BytesIO()
    usados_nomes = set()

    df_bur = st.session_state.get("df_faltantes")
    df_inu = st.session_state.get("df_inutilizadas")
    df_can = st.session_state.get("df_canceladas")
    df_aut = st.session_state.get("df_autorizadas")

    df_g = df_geral.reset_index(drop=True)
    df_g = _df_com_data_emissao_dd_mm_yyyy(df_g)
    if "Modelo" in df_g.columns:
        df_cte = df_g[df_g["Modelo"].astype(str).str.strip().eq("CT-e")].copy()
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

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        _excel_escrever_folha_df(writer, df_g, "Geral", usados_nomes)
        _excel_escrever_folha_df(writer, df_bur, "Buracos", usados_nomes)
        _excel_escrever_folha_df(writer, df_inu, "Inutilizadas", usados_nomes)
        _excel_escrever_folha_df(writer, df_can, "Canceladas", usados_nomes)
        _excel_escrever_folha_df(writer, df_aut, "Autorizadas", usados_nomes)
        _excel_escrever_folha_df(writer, df_cte, "CT-e lidas", usados_nomes)
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

        _excel_escrever_painel_fiscal(wb, kpi, usados_nomes)

    return buf.getvalue()


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
            subtitulo="Situação das linhas de terceiros no relatório geral: normal, cancelada, inutilizada, etc.",
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
        "(resumo por série, terceiros e cada aba do relatório da leitura). Cada bloco inclui uma "
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
            "da sua empresa (emitente = CNPJ da barra lateral). Se guardou «último nº por série» na lateral, a lista respeita esse mês de referência e o último número."
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
        "geral": (
            "Corresponde à aba «Relatório geral», com as colunas principais. A chave de acesso aparece "
            "abreviada neste PDF; use «Baixar Excel» na app para todas as linhas e colunas completas."
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

    raw = pdf.output(dest="S")
    if raw is None:
        return None
    if isinstance(raw, bytes):
        return raw
    if isinstance(raw, bytearray):
        return bytes(raw)
    return str(raw).encode("latin-1", "replace")


def aplicar_compactacao_dfs_sessao():
    """Compacta DataFrames grandes na sessão (útil no Streamlit Cloud)."""
    for k in (
        "df_geral",
        "df_resumo",
        "df_faltantes",
        "df_canceladas",
        "df_inutilizadas",
        "df_autorizadas",
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
    (Guardar referência com linhas válidas). Sem isso, buracos usam toda a numeração lida — como antes.
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
    Com referência activa: séries **com** último informado usam mês/âncora; séries **sem** linha na
    referência comportam-se como leitura total só nos buracos (não cortam meses anteriores).
    """
    if ref_ar is None or ref_mr is None:
        return True
    if ultimo_u is None:
        return True
    return numero_entra_conjunto_buraco(ano, mes, n, ref_ar, ref_mr, ultimo_u)


def ultimo_ref_lookup(ref_map, tipo, serie):
    if not ref_map:
        return None
    return ref_map.get(f"{tipo}|{str(serie).strip()}")


def numero_entra_conjunto_buraco(ano, mes, n, ref_ar, ref_mr, ultimo_u):
    """
    Se há mês de referência: ignora competências anteriores; no próprio mês só conta n > último informado.
    Sem referência na sessão: conta tudo (comportamento antigo).
    """
    if ref_ar is None or ref_mr is None:
        return True
    if str(ano) == "0000":
        return False
    if _ym_lt(ano, mes, ref_ar, ref_mr):
        return False
    if ultimo_u is None:
        return True
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


def _registro_manual_buraco_existe(relatorio, tipo, serie_str, nota_int):
    """Já existe inutil. ou cancel. manual para o mesmo modelo/série/número."""
    ser = str(serie_str).strip()
    for r in relatorio or []:
        if not (_inutil_sem_xml_manual(r) or _cancel_sem_xml_manual(r)):
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
            "Faltam colunas reconhecíveis. Use **Modelo** (código Sefaz **55**, **65**, **57**, **58** ou NF-e…), "
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
        ]
    )


def bytes_modelo_planilha_inutil_sem_xml_xlsx():
    df = _dataframe_modelo_planilha_inutil_sem_xml()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Inutil_sem_XML", index=False)
        ws = writer.sheets["Inutil_sem_XML"]
        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 2, 12)
    return buf.getvalue()


def bytes_modelo_planilha_cancel_sem_xml_xlsx():
    """Mesmo layout que inutilizadas (Modelo, Série, Nota) — só canceladas declaradas manualmente."""
    df = _dataframe_modelo_planilha_inutil_sem_xml()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Cancel_sem_XML", index=False)
        ws = writer.sheets["Cancel_sem_XML"]
        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 2, 12)
    return buf.getvalue()


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
            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
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
            if is_p and res["Número"] > 0:
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
            "df_geral": pd.DataFrame(geral_list),
            "st_counts": {
                "CANCELADOS": len(canc_list),
                "INUTILIZADOS": len(inut_list),
                "AUTORIZADAS": len(aut_list),
            },
        }
    )
    aplicar_compactacao_dfs_sessao()


def reprocessar_garimpeiro_a_partir_do_disco(cnpj_limpo: str):
    """
    Relê todos os XML/ZIP em TEMP_UPLOADS_DIR (mesmas regras de fusão por chave do 1.º garimpo),
    mantém registos manuais de inutilização e de cancelamento «sem XML» e recalcula os dataframes.
    """
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    if len(cnpj) != 14:
        return False, "CNPJ inválido — confira a barra lateral."

    if not os.path.isdir(TEMP_UPLOADS_DIR):
        return False, "Pasta de uploads não existe."

    nomes = [
        f
        for f in os.listdir(TEMP_UPLOADS_DIR)
        if os.path.isfile(os.path.join(TEMP_UPLOADS_DIR, f))
    ]
    if not nomes:
        return (
            False,
            "Nenhum ficheiro na pasta de uploads. Use «Incluir mais XML / ZIP» ou inicie um novo garimpo.",
        )

    lote_dict = {}
    for f_name in nomes:
        caminho_leitura = os.path.join(TEMP_UPLOADS_DIR, f_name)
        try:
            with open(caminho_leitura, "rb") as file_obj:
                todos_xmls = extrair_recursivo(file_obj, f_name)
                for name, xml_data in todos_xmls:
                    res, is_p = identify_xml_info(xml_data, cnpj, name)
                    if res:
                        key = res["Chave"]
                        if key in lote_dict:
                            if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
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
    for _pfx in ("rep_bur", "rep_inu", "rep_canc", "rep_aut", "rep_ger"):
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


def _lista_ficheiros_pasta_uploads():
    if not os.path.isdir(TEMP_UPLOADS_DIR):
        return []
    return [
        f
        for f in os.listdir(TEMP_UPLOADS_DIR)
        if os.path.isfile(os.path.join(TEMP_UPLOADS_DIR, f))
    ]


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
):
    """
    Um único passo: grava XML/ZIP extra em disco, aplica inutilizações e canceladas manuais
    (buracos / planilha / faixa, mesmas regras) e recalcula a partir do disco (ou só reconstrói).
    """
    cnpj = "".join(c for c in str(cnpj_limpo or "") if c.isdigit())[:14]
    linhas = []
    n_extra = 0
    if extra_files:
        os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
        for f in extra_files:
            caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, f.name)
            with open(caminho_salvo, "wb") as out_f:
                out_f.write(f.getvalue())
            n_extra += 1
        if n_extra:
            linhas.append(f"**{n_extra}** ficheiro(s) extra gravados na pasta de uploads.")

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
            if not _relatorio_ja_tem_chave(_it["Chave"]) and not _registro_manual_buraco_existe(
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
                            if not _relatorio_ja_tem_chave(_itp["Chave"]) and not _registro_manual_buraco_existe(
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
                if not _relatorio_ja_tem_chave(_itf["Chave"]) and not _registro_manual_buraco_existe(
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
            if not _relatorio_ja_tem_chave(_itc["Chave"]) and not _registro_manual_buraco_existe(
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
                            if not _relatorio_ja_tem_chave(_itpc["Chave"]) and not _registro_manual_buraco_existe(
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
                if not _relatorio_ja_tem_chave(_itfc["Chave"]) and not _registro_manual_buraco_existe(
                    st.session_state["relatorio"], _mfc, str(_sfc).strip(), _nn
                ):
                    st.session_state["relatorio"].append(_itfc)
                    n_f_c += 1
            if n_f_c:
                linhas.append(f"**{n_f_c}** nota(s) da faixa de **canceladas** (buracos).")

    tem_ficheiros = bool(_lista_ficheiros_pasta_uploads())
    fez_algo = (
        n_extra > 0
        or n_b > 0
        or n_p > 0
        or n_f > 0
        or n_b_c > 0
        or n_p_c > 0
        or n_f_c > 0
    )

    if err_p:
        return False, err_p, linhas
    if err_p_c:
        return False, err_p_c, linhas
    if (err_f or err_f_c) and not fez_algo and not tem_ficheiros:
        _msg_ff = err_f or err_f_c
        return False, _msg_ff, linhas
    if err_f:
        linhas.append(f"Aviso (faixa inutil.): {err_f}")
    if err_f_c:
        linhas.append(f"Aviso (faixa cancel.): {err_f_c}")

    if not fez_algo and not tem_ficheiros:
        return (
            False,
            "Nada a processar: inclua XML/ZIP, inutilizações / canceladas manuais ou faça primeiro o garimpo para haver ficheiros em disco.",
            linhas,
        )

    st.session_state["export_ready"] = False
    st.session_state["excel_buffer"] = None

    if tem_ficheiros:
        if len(cnpj) != 14:
            return False, "CNPJ inválido na barra lateral — necessário para reler os XML.", linhas
        ok, msg_rep = reprocessar_garimpeiro_a_partir_do_disco(cnpj_limpo)
        if not ok:
            return False, msg_rep, linhas
        linhas.append(msg_rep)
        return True, "\n\n".join(linhas), linhas

    reconstruir_dataframes_relatorio_simples()
    linhas.append(
        "Relatório recalculado (sem ficheiros na pasta de uploads — só registos manuais de inutil. / cancel.)."
    )
    return True, "\n\n".join(linhas), linhas


_V2_STATUS_UI_PARA_DF = {
    "Autorizadas": ["NORMAIS"],
    "Canceladas": ["CANCELADOS"],
    "Inutilizadas": ["INUTILIZADOS", "INUTILIZADA"],
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


def _v2_limpar_zips_gerados_etapa3_no_cwd():
    """Apaga ZIPs gerados pela Etapa 3 no diretório atual (nomes z_org_*_ptN / z_todos_*_ptN)."""
    rx = re.compile(r"^z_(org|todos)_.+_pt\d+\.zip$", re.I)
    for fn in os.listdir("."):
        if rx.match(fn):
            try:
                os.remove(fn)
            except OSError:
                pass


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
):
    """
    Escreve ZIP(s) em disco (org / plano). Devolve (org_parts, todos_parts, xml_matched, aviso_sem_xml|None).
    xml_respeita_filtro=False → todos os XML cujo resumo está em df_g_base.
    excel_um_só_completo=True → em cada parte, o mesmo Excel com df_excel_completo (relatório inteiro).
    zip_tag: ex. «propria» / «terceiros» para nomes z_org_propria_pt1.zip (dois lotes independentes).
    """
    stem_org = "z_org_final" if not zip_tag else f"z_org_{zip_tag}"
    stem_todos = "z_todos_final" if not zip_tag else f"z_todos_{zip_tag}"
    if df_g_base is None or df_g_base.empty or "Chave" not in df_g_base.columns:
        return [], [], 0, "ERR:Relatório geral vazio ou sem coluna Chave."
    if xml_respeita_filtro:
        if df_filtrado_para_excel_bloco is None or df_filtrado_para_excel_bloco.empty:
            return [], [], 0, "ERR:Resultado filtrado: 0 linhas. Ajuste os filtros."
        _df_ch = df_filtrado_para_excel_bloco
    else:
        _df_ch = df_g_base
    filtro_chaves = {
        k
        for k in (_chave_para_conjunto_export(x) for x in _df_ch["Chave"].tolist())
        if k
    }
    if not filtro_chaves:
        return [], [], 0, "ERR:Nenhuma chave válida para exportar XML."

    xb_completo = None
    if excel_um_só_completo:
        xb_completo = _excel_bytes_geral_e_resumo_status(df_excel_completo)

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
            excel_fn = "RELATORIO_GARIMPEIRO/relatorio_geral_completo.xlsx"
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
            oname = f"{stem_org}_pt{Z['curr_org_part']}.zip"
            Z["z_org"] = zipfile.ZipFile(oname, "w", zipfile.ZIP_DEFLATED)
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
            tname = f"{stem_todos}_pt{Z['curr_todos_part']}.zip"
            Z["z_todos"] = zipfile.ZipFile(tname, "w", zipfile.ZIP_DEFLATED)
            Z["todos_parts"].append(tname)
            Z["todos_count"] = 0

    if v2_zip_org:
        org_name = f"{stem_org}_pt{Z['curr_org_part']}.zip"
        Z["z_org"] = zipfile.ZipFile(org_name, "w", zipfile.ZIP_DEFLATED)
        Z["org_parts"].append(org_name)
    if v2_zip_plano:
        todos_name = f"{stem_todos}_pt{Z['curr_todos_part']}.zip"
        Z["z_todos"] = zipfile.ZipFile(todos_name, "w", zipfile.ZIP_DEFLATED)
        Z["todos_parts"].append(todos_name)

    if os.path.exists(TEMP_UPLOADS_DIR) and (v2_zip_org or v2_zip_plano):
        for f_name in os.listdir(TEMP_UPLOADS_DIR):
            f_path = os.path.join(TEMP_UPLOADS_DIR, f_name)
            with open(f_path, "rb") as f_temp:
                for name, xml_data in extrair_recursivo(f_temp, f_name):
                    res, _ = identify_xml_info(xml_data, cnpj_limpo, name)
                    ck = _chave_para_conjunto_export(res["Chave"]) if res else None
                    if res and ck and ck in filtro_chaves:
                        Z["xml_matched"] += 1
                        Z["chaves_bloco"].add(ck)
                        if v2_zip_org and Z["z_org"] is not None:
                            Z["z_org"].writestr(f"{res['Pasta']}/{name}", xml_data)
                            Z["org_count"] += 1
                        if v2_zip_plano and Z["z_todos"] is not None:
                            Z["z_todos"].writestr(name, xml_data)
                            Z["todos_count"] += 1
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
            "RELATORIO_GARIMPEIRO/relatorio_geral_completo.xlsx"
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
        try:
            os.remove(org_parts[-1])
        except OSError:
            pass
        org_parts = []
    if v2_zip_plano and Z["todos_count"] == 0 and todos_parts:
        try:
            os.remove(todos_parts[-1])
        except OSError:
            pass
        todos_parts = []

    aviso = None
    if (v2_zip_org or v2_zip_plano) and Z.get("xml_matched", 0) == 0:
        aviso = (
            "Nenhum XML em disco correspondeu às chaves. "
            "Causas frequentes: pasta do garimpo apagada, ou chaves na tabela que não batem com os ficheiros."
        )
    return org_parts, todos_parts, Z["xml_matched"], aviso


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
    for key in ("org_zip_parts", "todos_zip_parts"):
        for p in st.session_state.get(key) or []:
            try:
                if p and os.path.isfile(p):
                    os.remove(p)
            except OSError:
                pass


def _v2_limpar_estado_exportacao_etapa3():
    """Remove ficheiros ZIP gerados e reinicia sessão de downloads da Etapa 3."""
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
    m = re.search(r"pt(\d+)\.zip$", caminho_ficheiro, re.I)
    if m:
        return f"ZIP — parte {m.group(1)}"
    return "ZIP"


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
    st.session_state[sk] = (sig, b)
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
            use_container_width=True,
            hide_index=True,
            height=height,
        )
        return df_raw

    df_base = df_raw.reset_index(drop=True).copy()
    df_grid = _df_relatorio_leitura_abas_para_exibicao_sem_sep_milhar(df_base.copy())
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
    return df_base.iloc[idx].reset_index(drop=True)


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
        use_container_width=True,
    ):
        with st.spinner("A montar ZIP a partir dos ficheiros em disco…"):
            parts, tot = escrever_zip_dominio_por_chaves(
                cnpj_limpo, chaves, df_geral_full
            )
        st.session_state[kzip] = parts
        if tot == 0:
            st.warning(
                "Nenhum XML em disco correspondeu a estas chaves (pasta de upload limpa ou chaves externas ao lote)."
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
                    use_container_width=True,
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
    if not chaves_lista or not os.path.exists(TEMP_UPLOADS_DIR):
        return [], 0
    ch_set = set()
    for c in chaves_lista:
        k = _chave44_digitos(c)
        if k:
            ch_set.add(k)
    if not ch_set:
        return [], 0
    try:
        for f in os.listdir("."):
            if f.endswith(".zip") and "faltantes_dominio_final" in f:
                try:
                    os.remove(f)
                except Exception:
                    pass
    except Exception:
        pass

    parts = []
    part_idx = 1
    count_xml = 0
    no_lote = 0
    nome = f"faltantes_dominio_final_pt{part_idx}.zip"
    zf = zipfile.ZipFile(nome, "w", zipfile.ZIP_DEFLATED)
    parts.append(nome)
    usados_nomes_parte = set()
    chaves_excel_ordem = []
    vistos_chave_excel = set()

    try:
        for fn in os.listdir(TEMP_UPLOADS_DIR):
            f_path = os.path.join(TEMP_UPLOADS_DIR, fn)
            with open(f_path, "rb") as ft:
                for name, data in extrair_recursivo(ft, fn):
                    res, _ = identify_xml_info(data, cnpj_limpo, name)
                    ch44 = _chave44_digitos(res.get("Chave")) if res else None
                    if res and ch44 and ch44 in ch_set:
                        if no_lote >= MAX_XML_PER_ZIP:
                            _zip_anexar_excel_lista_especifica(
                                zf, df_geral, chaves_excel_ordem, part_idx
                            )
                            zf.close()
                            part_idx += 1
                            nome = f"faltantes_dominio_final_pt{part_idx}.zip"
                            zf = zipfile.ZipFile(nome, "w", zipfile.ZIP_DEFLATED)
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
        for p in parts:
            try:
                os.remove(p)
            except Exception:
                pass
        return [], 0

    return parts, count_xml


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
5. (Opcional) Lateral “Último nº por série”: afeta só o cálculo de buracos (âncora por último nº e mês). O garimpo e o resumo por série continuam totais. Sem Guardar referência válida, buracos usam todo o intervalo lido.
6. Inutilizadas: a partir dos buracos, por planilha (Excel/CSV) ou faixa — só números que já forem buraco listado (não alarga intervalos).
7. Painel à direita: **Processar Dados** grava XML/ZIP extra, aplica inutilizações «sem XML» (se configurou) e recalcula a partir da pasta de uploads.
8. Etapa 3: filtros em cascata (emissão própria e terceiros em colunas separadas). Escolha um dos seis modos: ZIP tudo (raiz ou pastas), ZIP filtrado (pastas ou tudo na raiz), Excel só lote completo, Excel só filtrado — e gere as partes quando aplicável.
9. Lista específica (secção própria): exporte subconjuntos por chaves, faixa, período, série ou uma nota — em Excel e/ou ZIP conforme os botões apresentados.

=== DICAS ===
• Resetar sistema: limpa sessão e temporários ao mudar de cliente ou recomeçar.
• Nos filtros, lista vazia = esse critério não restringe. Opções inválidas após mudar outro filtro são limpas automaticamente.
• Nomes de modelos na app: NF-e, NFS-e, NFC-e, CT-e, DACT-e, MDF-e, Outros (cartas de correção não são lidas).
""".strip()


# --- INTERFACE ---
st.markdown("<h1>⛏️ Garimpeiro</h1>", unsafe_allow_html=True)

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
                <li style="margin-bottom:8px;"><b>(Opcional)</b> <b>Último nº por série</b> (lateral): altera apenas o cálculo de <b>buracos</b> (âncora por último número e mês). O garimpo e o resumo por série mantêm-se <b>totais</b>. Sem <b>Guardar referência</b> com linhas válidas, os buracos consideram toda a numeração lida.</li>
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
        Nos filtros, lista vazia = esse critério não aplica. Modelos usados na app incluem NF-e, NFS-e, NFC-e, CT-e, DACT-e, MDF-e e Outros (cartas de correção ignoradas).
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
            st.session_state[k] = {"CANCELADOS": 0, "INUTILIZADOS": 0, "AUTORIZADAS": 0}
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

with st.sidebar:
    st.markdown("#### 🔍 Configuração")
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

        with st.expander("📌 Últimos nº / séries (mês ref.)", expanded=False):
            st.caption(
                "Define o **mês de referência** e, por linha, modelo + série + último nº (buracos no dashboard)."
            )
            d = date.today()
            def_ano = d.year - 1 if d.month == 1 else d.year
            def_mes = 12 if d.month == 1 else d.month - 1
            a0 = st.session_state["seq_ref_ano"] if st.session_state.get("seq_ref_ano") is not None else def_ano
            m0 = st.session_state["seq_ref_mes"] if st.session_state.get("seq_ref_mes") is not None else def_mes
            if st.session_state.get("garimpo_ok"):
                if st.button("Puxar séries do resumo", key="seq_btn_puxar", use_container_width=True):
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

            _opts = ["NF-e", "NFC-e", "CT-e"]
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
                if st.button("➕ Série", key="seq_add_row", use_container_width=True):
                    cur_df = collect_seq_ref_from_widgets(v, n_rows)
                    novo = pd.DataFrame([{"Modelo": "NF-e", "Série": "", "Último número": ""}])
                    st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(
                        pd.concat([cur_df, novo], ignore_index=True)
                    )
                    st.session_state["seq_struct_v"] = v + 1
                    st.rerun()
            with b2:
                if n_rows > 1 and st.button("➖ Última", key="seq_rem_row", use_container_width=True):
                    cur_df = collect_seq_ref_from_widgets(v, n_rows)
                    st.session_state["seq_ref_rows"] = normalize_seq_ref_editor_df(cur_df.iloc[:-1])
                    st.session_state["seq_struct_v"] = v + 1
                    st.rerun()

            if st.button(
                "Guardar referência",
                type="primary",
                use_container_width=True,
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
            st.markdown("##### 📄 PDF do dashboard")
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
                    use_container_width=True,
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
    • <b>Emissão própria (esquerda):</b> só a sua empresa — status, tipo, operação, datas, série, n.º, UF destino, nota por chave ou n.º+série. Vazio = não filtra por estes campos.<br/>
    • <b>Terceiros (direita):</b> só documentos recebidos — status, tipo, operação, período.<br/><br/>
    <b>Exportar</b><br/>
    • <b>ZIP todo o lote</b> — ignora filtros para escolher XML; Excel completo dentro de cada ZIP.<br/>
    • <b>ZIP filtrado</b> — só XML que passam nos filtros; Excel coerente com esse conjunto.<br/>
    • <b>Só Excel</b> — sem XML.<br/><br/>
    <b>Descargas</b> — sempre em <b>dois ficheiros</b> (própria e terceiros). Nos ZIP, o Excel está em <code>RELATORIO_GARIMPEIRO/</code> (até 10 000 XML por parte).
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
    _v2_tipos_ui = ["NF-e", "NFS-e", "NFC-e", "CT-e", "DACT-e", "MDF-e", "Outros"]
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
    
    _v2_c_sig = (id(df_g_base), v2_assinatura_exportacao_sessao()[:-2])
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
            _v2_limpar_estado_exportacao_etapa3()
            st.session_state["v2_show_regen_hint"] = True
    
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
                help="Alinhado à coluna Modelo. NFS-e / DACT-e dependem do conteúdo do XML.",
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
        st.caption("**Opções** — repor filtros e limpar exportação gerada")
        st.button(
            "Repor filtros e limpar exportação gerada",
            key="v2_pre_clr",
            on_click=v2_callback_repor_filtros,
            use_container_width=True,
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
    
    st.caption(
        "Cada botão gera **só** esse lado (menos trabalho e memória). **Os dois lados** gera os dois de uma vez."
    )
    col_g_pr, col_g_tc = st.columns(2, gap="large")
    with col_g_pr:
        gen_pr = st.button(
            "Gerar — sua empresa",
            type="primary",
            key="v2_btn_export_pr",
            disabled=_dis_pr,
            use_container_width=True,
        )
    with col_g_tc:
        gen_tc = st.button(
            "Gerar — terceiros",
            type="primary",
            key="v2_btn_export_tc",
            disabled=_dis_tc,
            use_container_width=True,
        )
    gen_ambos = st.button(
        "Gerar os dois lados",
        key="v2_btn_export_ambos",
        disabled=_dis_ambos,
        use_container_width=True,
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
                    _v2_limpar_zips_gerados_etapa3_no_cwd()
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
                                f"relatorio_todo_lote_emissao_propria_{ts}.xlsx"
                            )
                        if _bt:
                            st.session_state["excel_buffer_terceiros"] = _bt
                            st.session_state["export_excel_name_terceiros"] = (
                                f"relatorio_todo_lote_terceiros_{ts}.xlsx"
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
                    _v2_limpar_zips_gerados_etapa3_no_cwd()
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
                                f"relatorio_filtrado_emissao_propria_{ts}.xlsx"
                            )
                        if _bt:
                            st.session_state["excel_buffer_terceiros"] = _bt
                            st.session_state["export_excel_name_terceiros"] = (
                                f"relatorio_filtrado_terceiros_{ts}.xlsx"
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
                with st.spinner("A gerar ZIP…"):
                    st.session_state["excel_buffer"] = None
                    st.session_state.pop("excel_buffer_propria", None)
                    st.session_state.pop("excel_buffer_terceiros", None)
                    st.session_state.pop("export_excel_name_propria", None)
                    st.session_state.pop("export_excel_name_terceiros", None)
                    gc.collect()
                    st.session_state["export_excel_name"] = (
                        f"relatorio_completo_{ts}.xlsx"
                    )
                    _v2_limpar_zips_gerados_etapa3_no_cwd()
                    org_all = []
                    todos_all = []
                    _err_zip = None
                    for df_sl, ztag in _pares_zip:
                        o, t, _xm, av = _v2_export_zip_etapa3(
                            df_sl,
                            xml_respeita_filtro=_xml_filt,
                            df_filtrado_para_excel_bloco=(
                                df_sl if _xml_filt else None
                            ),
                            excel_um_só_completo=_fmt_run.startswith("zip_tudo"),
                            df_excel_completo=df_sl,
                            v2_zip_org=_fmt_run.endswith("_pastas"),
                            v2_zip_plano=_fmt_run.endswith("_raiz"),
                            cnpj_limpo=cnpj_limpo,
                            zip_tag=ztag,
                        )
                        if av and str(av).startswith("ERR:"):
                            _err_zip = str(av)[4:].strip()
                            break
                        org_all.extend(o)
                        todos_all.extend(t)
                    if _err_zip:
                        _v2_limpar_zips_gerados_etapa3_no_cwd()
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
                                        use_container_width=True,
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
                                        use_container_width=True,
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
                        use_container_width=True,
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
                                        use_container_width=True,
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
                                        use_container_width=True,
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
                        use_container_width=True,
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
                                    use_container_width=True,
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
                                    use_container_width=True,
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
                    use_container_width=True,
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
                    use_container_width=True,
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
                    use_container_width=True,
                )


def _garim_etapa3_fragment_entry():
    """Ponto de entrada estável para `st.fragment` (evita redefinir a função a cada rerun)."""
    _cx = "".join(c for c in str(st.session_state.get("cnpj_widget", "")) if c.isdigit())[:14]
    _garim_etapa3_corpo(_cx)


if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        st.markdown("##### 📎 Documentos XML / ZIP para ler")
        st.caption(
            "Carregue abaixo os ficheiros do lote; depois use **Iniciar grande garimpo** para ler e montar o relatório."
        )
        uploaded_files = st.file_uploader("📂 Escolha os XML e/ou ZIP (suporta grandes volumes):", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            limpar_arquivos_temp() 
            os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
            
            lote_dict = {}
            progresso_bar = st.progress(0)
            status_text = st.empty()
            total_arquivos = len(uploaded_files)
            
            with st.status("⛏️ Minerando e salvando fisicamente...", expanded=True) as status_box:
                
                # 1. Salva uploads fisicamente no disco para evitar estouro de RAM
                for i, f in enumerate(uploaded_files):
                    caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, f.name)
                    with open(caminho_salvo, "wb") as out_f:
                        out_f.write(f.read())
                
                # 2. Lê do disco e monta as tabelas em tempo real
                lista_salvos = os.listdir(TEMP_UPLOADS_DIR)
                total_salvos = len(lista_salvos)
                
                for i, f_name in enumerate(lista_salvos):
                    if i % 50 == 0: 
                        gc.collect()
                        
                    progresso_bar.progress((i + 1) / total_salvos)
                    status_text.text(f"⛏️ Lendo conteúdo: {f_name}")
                    
                    caminho_leitura = os.path.join(TEMP_UPLOADS_DIR, f_name)
                    try:
                        with open(caminho_leitura, "rb") as file_obj:
                            todos_xmls = extrair_recursivo(file_obj, f_name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: 
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
                        if res["Número"] > 0:
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
                'df_geral': pd.DataFrame(geral_list),
                'st_counts': {
                    "CANCELADOS": len(canc_list), 
                    "INUTILIZADOS": len(inut_list), 
                    "AUTORIZADAS": len(aut_list)
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
                '<h3 class="garim-sec">📤 Emissões próprias — total por tipo</h3>',
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
                use_container_width=True,
                hide_index=True,
            )

            st.markdown(
                '<h3 class="garim-sec">📥 Terceiros — total por tipo</h3>', unsafe_allow_html=True
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
                    use_container_width=True,
                    hide_index=True,
                )

            st.markdown("---")
            st.markdown(
                '<h3 class="garim-sec">📊 Relatório da leitura</h3>', unsafe_allow_html=True
            )
            df_fal = st.session_state["df_faltantes"]
            df_inu = st.session_state["df_inutilizadas"]
            df_can = st.session_state["df_canceladas"]
            df_aut = st.session_state["df_autorizadas"]
            df_ger = st.session_state["df_geral"]
            _n_bur = len(df_fal) if not df_fal.empty else 0
            _n_inu = len(df_inu) if not df_inu.empty else 0
            _n_can = len(df_can) if not df_can.empty else 0
            _n_aut = len(df_aut) if not df_aut.empty else 0
            _n_ger = len(df_ger) if not df_ger.empty else 0
            tab_bur, tab_inut, tab_canc, tab_aut, tab_geral = st.tabs(
                [
                    f"⚠️ Buracos ({_n_bur})",
                    f"🚫 Inutilizadas ({_n_inu})",
                    f"❌ Canceladas ({_n_can})",
                    f"✅ Autorizadas ({_n_aut})",
                    f"📋 Relatório geral ({_n_ger})",
                ]
            )

            with tab_bur:
                if not df_fal.empty:
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
                            use_container_width=True,
                        )
                    _painel_zip_xml_filtrado("rep_bur", df_b_f, cnpj_limpo, df_ger)
                else:
                    st.info("✅ Tudo em ordem.")
                with st.expander(
                    "Como funcionam os buracos e a referência na lateral (último nº / mês)",
                    expanded=False,
                ):
                    st.caption(
                        "Só **emissão própria**. O **Resumo por série** e esta lista de **buracos** incluem **NF-e**, **NFC-e** e **NFS-e** em que o emitente é o **CNPJ da barra lateral** (outros modelos não entram nestes quadros). Aqui: **números em falta** na sequência. "
                        "Com **Guardar referência** na lateral (mês + último nº por série), cada série indicada ignora XMLs de **meses antes** desse mês e "
                        "lista buracos **a partir do último nº + 1** — evita buraco gigante se aparecer uma nota fora da ordem (ex. janeiro no meio de março). "
                        "Séries **não** listadas na referência: buracos em **todo** o intervalo dos XMLs. **Sem** referência guardada: mesmo comportamento antigo (intervalo completo; pode ser enorme). "
                        "Na **Etapa 3** escolhe o que exportar."
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
                            use_container_width=True,
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
                            use_container_width=True,
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
                            use_container_width=True,
                        )
                    _painel_zip_xml_filtrado("rep_aut", df_a_f, cnpj_limpo, df_ger)
                else:
                    st.info("ℹ️ Nenhuma nota autorizada na amostra.")

            with tab_geral:
                if not df_ger.empty:
                    st.caption(
                        "Filtre e ordene **no cabeçalho de cada coluna**. Com **todas** as linhas visíveis, o **Excel** pode trazer o livro completo + dashboard; "
                        "com filtro na grelha, só a folha **Filtrado**."
                    )
                    df_g_f = _relatorio_leitura_tabela_aggrid(df_ger, "aggrid_rep_ger", height=480)
                    _sig_f = _df_sig_hash_memo(df_g_f)
                    _sig_full = _df_sig_hash_memo(df_ger)
                    _full_vista = _sig_f == _sig_full and len(df_g_f) == len(df_ger)
                    st.caption(f"**{len(df_g_f)}** linha(s) na vista (total: {len(df_ger)}).")
                    if _full_vista:
                        sk_wb = "_xlsx_mem_geral_workbook"
                        prev_wb = st.session_state.get(sk_wb)
                        if isinstance(prev_wb, tuple) and prev_wb[0] == _sig_full:
                            xlsx_g = prev_wb[1]
                        else:
                            xlsx_g = excel_relatorio_geral_com_dashboard_bytes(df_ger)
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
                            use_container_width=True,
                        )
                    _painel_zip_xml_filtrado("rep_ger", df_g_f, cnpj_limpo, df_ger)
                else:
                    st.info("Relatório geral vazio.")

        with _gcr:
            with st.container(border=True):
                st.markdown("##### 📤 Uploads e validação")
                st.caption(
                    "Configure **abaixo** os ficheiros, **inutilizações** e **canceladas** manuais; **um único botão** em baixo (**Processar Dados**) grava, aplica e recalcula tudo."
                )
                # MÓDULO: DOCUMENTOS XML/ZIP (carga incremental)
                # =====================================================================
                st.markdown("##### 📎 Documentos XML / ZIP para ler")
                with st.expander("➕ Incluir mais ficheiros no lote (sem resetar)", expanded=False):
                    extra_files = st.file_uploader(
                        "Escolha os XML ou ZIP a acrescentar ao lote atual:",
                        accept_multiple_files=True,
                        key="extra_files",
                    )
                    st.caption(
                        "Os ficheiros escolhidos são **gravados na pasta de uploads** ao carregar em **Processar Dados**; em seguida o relatório é **relido a partir do disco**."
                    )

                # =====================================================================
                # MÓDULO: AUTENTICIDADE — mesmo nível visual que Inutilizadas
                # =====================================================================
                st.markdown("##### 🔐 Validação de autenticidade")
                with st.expander(
                    "Relatório exportado da Sefaz para confrontar com o lote de XML (opcional).",
                    expanded=False,
                ):
                    st.caption(
                        "Use o Excel ou relatório que o **portal da Sefaz** disponibiliza (autorizadas / emitidas, conforme o caso). "
                        "O objetivo é cruzar essa lista com o que foi lido dos **XML** e assinalar **divergências** no painel fiscal (Excel/PDF). "
                        "Quando o upload e o processamento estiverem ativos nesta app, o passo integra-se no **Processar Dados** como as outras secções."
                    )

                # =====================================================================
                # MÓDULO: DECLARAR INUTILIZADAS MANUAIS
                # =====================================================================
                st.markdown("##### 🛠️ Inutilizadas")
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
                            "Colunas (1.ª linha = cabeçalho): **Modelo** = código Sefaz (**55** NF-e, **65** NFC-e, **57** CT-e, **58** MDF-e) "
                            "ou nome NF-e / NFC-e…; **Série**; **Nota** (ou Número / Num_Faltante). "
                            "Ideal para copiar/colar da Sefaz. Só entram linhas que já forem **buraco** no garimpeiro."
                        )
                        st.download_button(
                            "Baixar Excel",
                            data=bytes_modelo_planilha_inutil_sem_xml_xlsx(),
                            file_name="MODELO_inutilizadas_sem_XML_garimpeiro.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_modelo_inut_xlsx",
                            use_container_width=True,
                        )
                        st.caption(
                            "No modelo: **Modelo** em número (55, 65, 57, 58) como na Sefaz; **Série** e **Nota**. "
                            "Substitua ou apague as linhas de exemplo e guarde antes de importar."
                        )
                        _up_inut = st.file_uploader(
                            "Ficheiro .csv, .xlsx ou .xls",
                            type=["csv", "xlsx", "xls"],
                            key="inut_planilha_up",
                        )

                    with tab_f:
                        _mf = st.selectbox("Modelo", ["NF-e", "NFC-e", "CT-e", "MDF-e"], key="inut_f_mod")
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
                st.markdown("##### ❌ Canceladas")
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
                            "Colunas: **Modelo** (55, 65, 57, 58 ou NF-e…), **Série**, **Nota**. "
                            "Só entram linhas que já forem **buraco** no garimpeiro."
                        )
                        st.download_button(
                            "Baixar modelo Excel",
                            data=bytes_modelo_planilha_cancel_sem_xml_xlsx(),
                            file_name="MODELO_canceladas_sem_XML_garimpeiro.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_modelo_canc_xlsx",
                            use_container_width=True,
                        )
                        st.file_uploader(
                            "Ficheiro .csv, .xlsx ou .xls",
                            type=["csv", "xlsx", "xls"],
                            key="canc_planilha_up",
                        )

                    with tab_cf:
                        _mfc_ui = st.selectbox("Modelo", ["NF-e", "NFC-e", "CT-e", "MDF-e"], key="canc_f_mod")
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

                st.markdown("##### 🔁 Processar Dados")
                if st.button(
                    "Processar Dados",
                    key="btn_reprocessar_garimpo",
                    use_container_width=True,
                ):
                    _ef = st.session_state.get("extra_files")
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
        st.markdown("### 🔎 EXPORTAR LISTA ESPECÍFICA")
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
                    ["NF-e", "NFC-e", "CT-e", "MDF-e"],
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
                                elif not os.path.exists(TEMP_UPLOADS_DIR):
                                    st.error(
                                        "A pasta dos XML carregados não existe. Volte a correr o garimpo."
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
                                            "⚠️ Há chaves no relatório, mas **nenhum XML** correspondente no lote em disco."
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
                        elif not os.path.exists(TEMP_UPLOADS_DIR):
                            st.error(
                                "A pasta dos XML carregados não existe. Volte a correr o garimpo ou **Incluir mais XML**."
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
                                    "⚠️ Há chaves no relatório, mas **nenhum XML** foi encontrado em disco "
                                    f"para esse período. Confira se o lote contém esses ficheiros."
                                )

            with tab_faixa:
                mod_f = st.selectbox(
                    "Modelo",
                    ["NF-e", "NFC-e", "CT-e", "MDF-e"],
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
                                elif not os.path.exists(TEMP_UPLOADS_DIR):
                                    st.error(
                                        "A pasta dos XML carregados não existe. Volte a correr o garimpo."
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
                                            "⚠️ Chaves encontradas no relatório, mas **nenhum XML** no lote em disco."
                                        )

            with tab_unica:
                mod_u = st.selectbox(
                    "Modelo",
                    ["NF-e", "NFC-e", "CT-e", "MDF-e"],
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
                            elif not os.path.exists(TEMP_UPLOADS_DIR):
                                st.error(
                                    "A pasta dos XML carregados não existe. Volte a correr o garimpo."
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
                                        "⚠️ Chave no relatório, mas **nenhum XML** correspondente no lote em disco."
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
                                    use_container_width=True,
                                )
else:
    st.warning("👈 Insira o CNPJ lateral para começar.")


