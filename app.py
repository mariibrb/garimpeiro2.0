# Garimpeiro — código-fonte único: faça todas as alterações neste ficheiro (app2.py).
# app.py só reencaminha para aqui quando a Cloud usa app.py como entrada.
import streamlit as st
import zipfile
import io
import os
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
from pathlib import Path


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
            min-width: min(312px, 100vw) !important;
            max-width: min(460px, 100vw) !important;
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
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] {
            background: linear-gradient(160deg, #fffafd 0%, #ffffff 50%, #fff8fc 100%) !important;
            border: 1px solid rgba(255, 105, 180, 0.35) !important;
            border-radius: 14px !important;
            padding: 0.5rem 0.65rem 0.6rem !important;
            margin-bottom: 0.5rem !important;
            box-shadow: 0 2px 12px rgba(255, 105, 180, 0.07) !important;
            overflow-x: auto !important;
            overflow-y: visible !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
        }
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="select"] > div {
            border-radius: 10px !important;
            border-color: #f8bbd0 !important;
            min-height: 2.15rem !important;
            max-width: 100% !important;
        }
        [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] input {
            border-radius: 10px !important;
            border-color: #f5c6d8 !important;
            min-height: 2.15rem !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
        }
        p.garim-seq-head {
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            font-size: 0.82rem !important;
            font-weight: 700 !important;
            color: #c2185b !important;
            margin: 0.35rem 0 0.15rem 0 !important;
            letter-spacing: 0.02em !important;
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

        /* Dentro de expanders: botões mais baixos (Etapa 2, lateral, etc.) */
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

        [data-testid="stMetric"] {
            background: white !important;
            border-radius: 20px !important;
            border: 1px solid #FFDEEF !important;
            padding: 15px !important;
        }

        h3.garim-sec {
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 800 !important;
            font-size: 1.02rem !important;
            color: #5D1B36 !important;
            text-align: left !important;
            margin: 1.25rem 0 0.45rem 0 !important;
            letter-spacing: 0.02em;
            border-left: 4px solid #A1869E;
            padding: 0.15rem 0 0.15rem 0.65rem;
        }
        section.main [data-testid="stDataFrame"] {
            border-radius: 12px;
            border: 1px solid rgba(161, 134, 158, 0.28);
            overflow: hidden;
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
        "Nome_Dest": ""
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
            if '<mod>65</mod>' in tag_l: 
                tipo = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: 
                tipo = "CT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: 
                tipo = "MDF-e"
            
            status = "NORMAIS"
            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: 
                status = "CANCELADOS"
            elif '110110' in tag_l: 
                status = "CARTA_CORRECAO"
                
            resumo["Tipo"] = tipo
            resumo["Status"] = status

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
# Colunas preferidas no relatório geral no PDF (espelho legível do ecrã).
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
        ("Validação autenticidade (Etapa 2)", "Sim" if val_ok else "Não"),
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
    val_txt = "Sim" if pm.get("Validação autenticidade (Etapa 2)") == "Sim" else "Não"
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
    ws.write(ctx_r + 1, 14, f"Etapa 2: {val_txt}", fmt_kpi_l)
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
    Excel com várias folhas alinhadas às abas do ecrã:
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
            ws.merge_range(row, 0, row, last_c, "Resumo por série (emissão própria)", sub_f)
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
    sub = "Boutique de dados · mesmo desenho que no ecrã principal."
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
            "TOTAL DE DOCUMENTOS\nLIDOS (NUM GERAL)",
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
            "Soma no resumo por série",
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
            "Um cartão por modelo e série: primeiro e último número encontrados nos XML do lote.",
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
    card_h = 33.0
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
            val_mid = f"{ini} – {fim}" if (str(ini) != "" or str(fim) != "") else "—"
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

            vm = str(val_mid)
            fs_val = 11.0 if len(vm) < 20 else (9.0 if len(vm) < 28 else 7.5)
            _pdf_font(pdf, use_dejavu, "B", fs_val)
            pdf.set_text_color(32, 35, 42)
            pdf.set_xy(x + 1.5, y0 + title_h + 3.5)
            pdf.cell(cw - 3, val_h - 4, _pdf_txt(pdf, vm, use_dejavu), align="C", ln=False)

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
    """Referência lateral, Etapa 2, gerado em — estilo suave."""
    pm = dict(pares_lista or [])
    rotulos = [
        ("Gerado em", "Gerado em"),
        ("Referência último nº guardada", "Último nº por série (lateral)"),
        ("Validação autenticidade (Etapa 2)", "Autenticidade (Etapa 2)"),
    ]
    itens = [(c, pm[k]) for k, c in rotulos if k in pm]
    if not itens:
        return
    _pdf_lista_rosa(
        pdf,
        "Contexto na app (opcional)",
        itens,
        use_dejavu,
        subtitulo="Opções que usou na barra lateral ou na Etapa 2.",
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
        "As secções abaixo repetem a mesma ordem e o mesmo significado que no ecrã do Garimpeiro "
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
            "Corresponde ao quadro «Resumo por série» na página de resultados: para cada modelo e série "
            "da emissão própria, o intervalo de numeração encontrado nos ficheiros, a quantidade de notas "
            "e o valor contábil somado."
        ),
        "terc": (
            "Corresponde a «Terceiros — total por tipo»: soma de documentos recebidos de terceiros "
            "(por exemplo NF-e, NFC-e, CT-e, MDF-e) contados neste lote."
        ),
        "bur": (
            "Corresponde à aba «Buracos»: números em falta na sequência da numeração da emissão própria. "
            "Se guardou «último nº por série» na lateral, a lista respeita esse mês de referência e o último número."
        ),
        "inut": (
            "Corresponde à aba «Inutilizadas»: notas inutilizadas na Sefaz, incluindo as que declarou "
            "manualmente sem XML (quando usou essa opção na app)."
        ),
        "canc": (
            "Corresponde à aba «Canceladas»: notas canceladas no conjunto analisado. Se carregou o Excel "
            "da Etapa 2 (autenticidade), o status fica alinhado à Sefaz."
        ),
        "aut": (
            "Corresponde à aba «Autorizadas»: notas com situação normal/autorizada na emissão própria, neste lote."
        ),
        "geral": (
            "Corresponde à aba «Relatório geral», com as colunas principais. A chave de acesso aparece "
            "abreviada neste PDF; use «Baixar Excel» na app para todas as linhas e colunas completas."
        ),
    }

    _pdf_secao_resumo_folha(pdf, "Resumo por série", use_dejavu, _ex["serie"])
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


def _inutil_sem_xml_manual(res):
    """Inutilização declarada em «sem XML» — deve fechar buraco (não depender de Ano/Mês 0000)."""
    return res.get("Arquivo") == "REGISTRO_MANUAL" or str(
        res.get("Chave") or ""
    ).startswith("MANUAL_INUT_")


def parse_numeros_um_por_linha(text):
    """Um inteiro por linha (só dígitos na linha são usados)."""
    out = []
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        d = "".join(filter(str.isdigit, line))
        if not d:
            continue
        try:
            out.append(int(d))
        except ValueError:
            continue
    return out


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
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
        else:
            geral_list.append(registro_detalhado)
            if is_p:
                sk = (res["Tipo"], res["Série"])
                if sk not in audit_map:
                    audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
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
                    if res["Status"] == "CANCELADOS":
                        canc_list.append(registro_detalhado)
                    elif res["Status"] == "NORMAIS":
                        aut_list.append(registro_detalhado)
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


def filtrar_df_geral_para_exportacao(
    df_base,
    filtro_origem,
    filtro_meses,
    aplicar_mes_so_na_propria,
    filtro_modelos,
    filtro_series,
    filtro_status,
    filtro_operacao,
):
    """Mesma lógica da Etapa 3, reutilizada para pré-visualização e exportação."""
    if df_base is None or df_base.empty:
        return df_base
    out = df_base.copy()
    if len(filtro_origem) > 0:
        pat = "|".join([re.escape(o.split()[0]) for o in filtro_origem])
        out = out[out["Origem"].str.contains(pat, regex=True, na=False)]
    if len(filtro_meses) > 0:
        out = out.copy()
        out["Mes_Comp"] = out["Ano"].astype(str) + "/" + out["Mes"].astype(str)
        if aplicar_mes_so_na_propria:
            out = out[
                (out["Mes_Comp"].isin(filtro_meses))
                | (out["Origem"].str.contains("TERCEIROS", na=False))
            ]
        else:
            out = out[out["Mes_Comp"].isin(filtro_meses)]
    if len(filtro_modelos) > 0:
        out = out[out["Modelo"].isin(filtro_modelos)]
    if len(filtro_series) > 0:
        out = out[out["Série"].astype(str).isin(filtro_series)]
    if len(filtro_status) > 0:
        out = out[out["Status Final"].isin(filtro_status)]
    if len(filtro_operacao) > 0 and "Operação" in out.columns:
        out = out[out["Operação"].isin(filtro_operacao)]
    return out


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


def v2_opcoes_cascata_etapa3(
    df_base: pd.DataFrame,
    filtro_origem: list,
    filtro_meses: list,
    aplicar_mes_so_na_propria: bool,
    filtro_modelos: list,
    filtro_series: list,
    filtro_status: list,
    filtro_operacao: list,
) -> dict:
    """
    Opções de cada multiselect = valores distintos ainda possíveis no relatório
    depois de aplicar todos os filtros exceto o da própria dimensão.
    Assim, com só «EMISSÃO PRÓPRIA», Série lista apenas séries que existem nessa origem.
    """
    empty = {
        "anos_meses": [],
        "modelos": [],
        "series": [],
        "status": [],
        "operacoes": [],
    }
    if df_base is None or df_base.empty:
        return empty

    def uniq_sorted_series(s: pd.Series) -> list:
        return sorted(
            {str(x) for x in s.tolist() if str(x) not in ("", "nan", "None")},
            key=lambda x: (len(x), x),
        )

    # Ano/mês: ignora filtro de mês; mantém os restantes
    d_m = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        [],
        aplicar_mes_so_na_propria,
        filtro_modelos,
        filtro_series,
        filtro_status,
        filtro_operacao,
    )
    anos_meses: list = []
    if d_m is not None and not d_m.empty:
        mc = d_m["Ano"].astype(str) + "/" + d_m["Mes"].astype(str)
        anos_meses = sorted({x for x in mc.tolist() if x and not str(x).startswith("0000/")})

    d_mod = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_meses,
        aplicar_mes_so_na_propria,
        [],
        filtro_series,
        filtro_status,
        filtro_operacao,
    )
    modelos = uniq_sorted_series(d_mod["Modelo"]) if d_mod is not None and not d_mod.empty else []

    d_ser = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_meses,
        aplicar_mes_so_na_propria,
        filtro_modelos,
        [],
        filtro_status,
        filtro_operacao,
    )
    series = (
        uniq_sorted_series(d_ser["Série"].astype(str))
        if d_ser is not None and not d_ser.empty
        else []
    )

    d_st = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_meses,
        aplicar_mes_so_na_propria,
        filtro_modelos,
        filtro_series,
        [],
        filtro_operacao,
    )
    status = (
        uniq_sorted_series(d_st["Status Final"])
        if d_st is not None and not d_st.empty and "Status Final" in d_st.columns
        else []
    )

    d_op = filtrar_df_geral_para_exportacao(
        df_base,
        filtro_origem,
        filtro_meses,
        aplicar_mes_so_na_propria,
        filtro_modelos,
        filtro_series,
        filtro_status,
        [],
    )
    operacoes = (
        uniq_sorted_series(d_op["Operação"])
        if d_op is not None and not d_op.empty and "Operação" in d_op.columns
        else []
    )

    return {
        "anos_meses": anos_meses,
        "modelos": modelos,
        "series": series,
        "status": status,
        "operacoes": operacoes,
    }


def v2_sanear_selecoes_contra_opcoes(
    anos_meses: list,
    modelos: list,
    series: list,
    status_opcoes: list,
    operacoes_opts: list,
) -> None:
    """Remove da sessão valores que deixaram de existir nas listas em cascata."""
    pares = [
        ("v2_f_mes", set(anos_meses)),
        ("v2_f_mod", set(modelos)),
        ("v2_f_ser", set(series)),
        ("v2_f_stat", set(status_opcoes)),
        ("v2_f_op", set(operacoes_opts)),
    ]
    for key, permitidos in pares:
        cur = list(st.session_state.get(key) or [])
        novo = [x for x in cur if x in permitidos]
        if novo != cur:
            st.session_state[key] = novo


def v2_callback_repor_filtros():
    """Limpa multiselects da Etapa 3. Deve ser usado com on_click (antes dos widgets na mesma corrida)."""
    for _kx in ("v2_f_orig", "v2_f_mes", "v2_f_mod", "v2_f_ser", "v2_f_stat", "v2_f_op"):
        st.session_state[_kx] = []


def rotulo_download_zip_parte(caminho_ficheiro):
    m = re.search(r"pt(\d+)\.zip$", caminho_ficheiro, re.I)
    if m:
        return f"Baixar XML (parte {m.group(1)})"
    return f"Baixar XML — {os.path.basename(caminho_ficheiro)}"


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


# Texto espelhado nos cartões + área “copiar guia” (manter alinhado ao fluxo real da app)
TEXTO_GUIA_GARIMPEIRO = """
Garimpeiro — Guia rápido (para copiar)

PASSO A PASSO
1. Na barra lateral: digite **só os 14 números** do CNPJ do emitente (cliente); pontos, barra e traço são colocados automaticamente (também pode colar já formatado). Depois clique em Liberar operação.
2. Envie ZIP ou XML soltos (volumes grandes são suportados). Depois do primeiro resultado, pode incluir mais ficheiros no topo da página, sem reiniciar o garimpo.
3. Clique em Iniciar grande garimpo e aguarde a leitura.
4. (Opcional) Lateral “Último nº por série”: **só muda os buracos** (âncora a partir do último nº + mês; evita buraco gigante se vier nota velha no meio). O **garimpo e o resumo** continuam **totais**. Sem **Guardar referência** com linhas válidas, buracos usam toda a numeração lida.
5. (Opcional) Etapa 2: suba um ou mais Excel de autenticidade (coluna A = chave 44 dígitos; coluna F = status) para alinhar cancelamentos com a Sefaz.
6. Inutilizadas sem XML: **Dos buracos**, **Planilha** (Excel/CSV) ou **Faixa** — só vale para o que o garimpeiro já listou como buraco (não alarga intervalos).
7. Etapa 3: filtros em cascata; ZIPs em partes de até 10 mil XML, cada ZIP já traz Excel do bloco em RELATORIO_GARIMPEI/; Excel com o filtro completo é opcional à parte.
8. Exportar lista específica: Excel com **chaves**, Excel com **inicial/final/série**, **período**, **faixa** ou **uma nota**.

O QUE O SISTEMA FAZ
• Emissão própria: leitura e **resumo por série totais**; **buracos** com referência guardada ficam ancorados (séries indicadas); sem referência, buracos em todo o intervalo; canceladas/inutilizadas; trechos limitam saltos falsos.
• Terceiros: totalizador por tipo (NF-e, NFC-e, CT-e, MDF-e).
• Um mesmo documento pode gerar mais do que um XML no disco (ex.: nota e evento) — o mesmo número de chaves pode corresponder a vários ficheiros.

DICAS
• CNPJ na lateral: apenas dígitos ou colar com máscara — a app normaliza e mostra 00.000.000/0000-00.
• Resetar sistema limpa sessão e temporários; use se trocar de cliente ou quiser recomeçar do zero.
• Modelos na app: NF-e, NFC-e, CT-e, MDF-e (use estes nomes nas tabelas e colagens).
• Etapa 3: lista vazia num critério = esse critério não corta nada. Seleções que deixam de existir após mudar outro filtro são limpas automaticamente.
""".strip()


# --- INTERFACE ---
st.markdown("<h1>⛏️ Garimpeiro</h1>", unsafe_allow_html=True)

with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 Como usar (passo a passo)</h3>
            <ol>
                <li><b>CNPJ:</b> Na lateral, <b>só os 14 números</b> do emitente (cliente); a máscara preenche sozinha → Liberar operação.</li>
                <li><b>Lote:</b> Envie ZIP ou XML. Grandes volumes são suportados.</li>
                <li><b>Garimpo:</b> Iniciar grande garimpo e aguardar.</li>
                <li><b>Mais ficheiros:</b> No <b>topo dos resultados</b>, inclua XML/ZIP extra <b>sem reiniciar</b>.</li>
                <li><b>(Opcional)</b> Último nº + mês (lateral) → só **buracos** ancorados; leitura/resumo **sempre totais**.</li>
                <li><b>(Opcional)</b> Etapa 2 — um ou mais Excel de autenticidade (chave na col. A, status na col. F).</li>
                <li><b>Inutilizadas sem XML:</b> Dos buracos, planilha Excel/CSV ou Faixa — só buracos já detectados.</li>
                <li><b>Exportar:</b> Etapa 3 — ZIP em blocos de 10 000 XML (com Excel do bloco dentro); Excel do filtro completo opcional à parte; ou lista por chaves (col. A).</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O que o sistema faz</h3>
            <ul>
                <li><b>Emissão própria:</b> Resumo **total** por série; buracos (referência opcional; trechos); canceladas e inutilizadas.</li>
                <li><b>Terceiros:</b> Contagem por tipo de documento (NF-e, CT-e, etc.).</li>
                <li><b>Exportação:</b> Etapa 3 — ZIP(s) com até 10 000 XML + Excel por bloco; relatório completo em Excel opcional.</li>
                <li><b>Lista de chaves:</b> Planilha com chaves 44 dígitos → ZIP só com esses XMLs do lote.</li>
                <li><b>Eventos:</b> Uma chave pode corresponder a vários XMLs (ex.: NF-e + evento).</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with st.expander("📋 Guia em texto simples (para copiar)", expanded=False):
        st.caption("Clique na caixa, Ctrl+A (Cmd+A no Mac) e Ctrl+C para copiar tudo.")
        st.text_area(
            "Guia",
            value=TEXTO_GUIA_GARIMPEIRO,
            height=320,
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
    st.markdown("### 🔍 Configuração")
    # Normalizar máscara *antes* do text_input: depois do widget o Streamlit bloqueia
    # `session_state["cnpj_widget"] = ...` na mesma execução (StreamlitAPIException).
    _cnpj_key = "cnpj_widget"
    _cnpj_raw = st.session_state.get(_cnpj_key, "")
    cnpj_limpo = "".join(c for c in str(_cnpj_raw) if c.isdigit())[:14]
    _cnpj_fmt = format_cnpj_visual(cnpj_limpo)
    if _cnpj_fmt != str(_cnpj_raw):
        st.session_state[_cnpj_key] = _cnpj_fmt
    st.text_input(
        "CNPJ DO CLIENTE",
        key=_cnpj_key,
        placeholder="somente numeros",
        help="14 dígitos; pode colar com ou sem máscara. A formatação é aplicada ao digitar.",
    )

    if cnpj_limpo and len(cnpj_limpo) != 14:
        st.error("⚠️ CNPJ Inválido.")
        
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): 
            st.session_state['confirmado'] = True

        with st.expander("📌 Último nº por série (fim do mês de referência)"):
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

            st.markdown('<p class="garim-seq-head">Séries do cliente</p>', unsafe_allow_html=True)

            for i, row in enumerate(_recs):
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

                with st.container(border=True):
                    st.selectbox(
                        "Tipo de documento",
                        opts_row,
                        index=idx,
                        key=f"sr_{v}_{i}_m",
                        label_visibility="visible",
                    )
                    c_ser, c_ult = st.columns(2, gap="small")
                    with c_ser:
                        st.text_input(
                            "Série",
                            value=ser_cur,
                            key=f"sr_{v}_{i}_s",
                            label_visibility="visible",
                            max_chars=10,
                            placeholder="ex. 1",
                        )
                    with c_ult:
                        st.text_input(
                            "Último nº",
                            value=ult_cur,
                            key=f"sr_{v}_{i}_u",
                            label_visibility="visible",
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
                st.info(
                    f"Referência ativa: {st.session_state['seq_ref_ano']}/"
                    f"{int(st.session_state['seq_ref_mes']):02d} — "
                    f"{len(st.session_state['seq_ref_ultimos'])} série(s)."
                )

        if st.session_state.get("garimpo_ok") and st.session_state.get("relatorio"):
            st.divider()
            st.markdown("#### 📄 PDF do dashboard")
            st.caption(
                "Resumo do lote em PDF (métricas e amostras). Só descarrega o ficheiro — não altera o ecrã."
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

    st.divider()
    
    if st.button("🗑️ RESETAR SISTEMA"):
        limpar_arquivos_temp()
        st.session_state.clear()
        st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded_files = st.file_uploader("📂 ARQUIVOS XML/ZIP (Suporta grandes volumes):", accept_multiple_files=True)
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
                    
                    if sk not in audit_map: 
                        audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                    ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
                        
                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        _man_inut = _inutil_sem_xml_manual(res)
                        for n in range(r[0], r[1] + 1):
                            audit_map[sk]["nums"].add(n)
                            if _man_inut or incluir_numero_no_conjunto_buraco(
                                res["Ano"], res["Mes"], n, ref_ar, ref_mr, ult_u
                            ):
                                audit_map[sk]["nums_buraco"].add(n)
                            inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
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
                            
                            if res["Status"] == "CANCELADOS":
                                canc_list.append(registro_base)
                            elif res["Status"] == "NORMAIS":
                                aut_list.append(registro_base)
                                
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
        # --- RESULTADOS TELA INICIAL (painel em cartões só no PDF; aqui tabelas e abas) ---
        st.caption(
            "Se faltar XML ou ZIP, use o bloco abaixo sem reiniciar o garimpo: os totais e as tabelas atualizam na hora."
        )
        # =====================================================================
        # MÓDULO: ADICIONAR MAIS ARQUIVOS (CARGA INCREMENTAL) — no topo dos resultados
        # =====================================================================
        with st.expander("➕ Incluir mais XML / ZIP no lote (sem resetar)", expanded=False):
            extra_files = st.file_uploader(
                "Ficheiros a acrescentar ao lote actual:",
                accept_multiple_files=True,
                key="extra_files",
            )
            if extra_files and st.button("Processar e atualizar", key="extra_btn_proc", type="primary"):
                with st.spinner("A adicionar…"):
                    os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
                    for f in extra_files:
                        caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, f.name)
                        with open(caminho_salvo, "wb") as out_f:
                            out_f.write(f.read())

                        f.seek(0)
                        try:
                            todos_xmls = extrair_recursivo(f, f.name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    ja_existe = any(
                                        item["Chave"] == res["Chave"] for item in st.session_state["relatorio"]
                                    )
                                    if not ja_existe:
                                        st.session_state["relatorio"].append(res)
                                del xml_data
                        except Exception:
                            pass

                    st.session_state["export_ready"] = False
                    reconstruir_dataframes_relatorio_simples()
                st.rerun()

        st.markdown(
            '<h3 class="garim-sec">📊 Resumo por série</h3>', unsafe_allow_html=True
        )
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)

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
            st.caption(f"Somatório geral (documentos lidos): {_df_terc['Quantidade'].sum()}")
            st.dataframe(_df_terc, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown(
            '<h3 class="garim-sec">📊 Relatório da leitura (abas)</h3>', unsafe_allow_html=True
        )
        st.caption(
            "Buracos, inutilizadas, canceladas, autorizadas e relatório geral — mesmas colunas que nas tabelas; use **Baixar Excel** abaixo de cada tabela."
        )
        tab_bur, tab_inut, tab_canc, tab_aut, tab_geral = st.tabs(
            ["⚠️ Buracos", "🚫 Inutilizadas", "❌ Canceladas", "✅ Autorizadas", "📋 Relatório geral"]
        )

        df_fal = st.session_state["df_faltantes"]
        df_inu = st.session_state["df_inutilizadas"]
        df_can = st.session_state["df_canceladas"]
        df_aut = st.session_state["df_autorizadas"]
        df_ger = st.session_state["df_geral"]

        with tab_bur:
            qtd_buracos = len(df_fal) if not df_fal.empty else 0
            st.markdown(f"#### ⚠️ Buracos ({qtd_buracos})")
            if not df_fal.empty:
                st.dataframe(df_fal, use_container_width=True, hide_index=True)
                xlsx_b = dataframe_para_excel_bytes(df_fal, "Buracos")
                if xlsx_b:
                    st.download_button(
                        "Baixar Excel",
                        data=xlsx_b,
                        file_name="relatorio_buracos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rep_buracos_xlsx",
                        use_container_width=True,
                    )
            else:
                st.info("✅ Tudo em ordem.")
            with st.expander(
                "Como funcionam os buracos e a referência na lateral (último nº / mês)",
                expanded=False,
            ):
                st.caption(
                    "Só **emissão própria**. **Resumo** e totais acima = **tudo** o que foi lido. Aqui: **números em falta** na sequência. "
                    "Com **Guardar referência** na lateral (mês + último nº por série), cada série indicada ignora XMLs de **meses antes** desse mês e "
                    "lista buracos **a partir do último nº + 1** — evita buraco gigante se aparecer uma nota fora da ordem (ex. janeiro no meio de março). "
                    "Séries **não** listadas na referência: buracos em **todo** o intervalo dos XMLs. **Sem** referência guardada: mesmo comportamento antigo (intervalo completo; pode ser enorme). "
                    "Na **Etapa 3** escolhe o que exportar."
                )

        with tab_inut:
            _q_inut = len(df_inu) if not df_inu.empty else 0
            st.markdown(f"#### 🚫 Inutilizadas ({_q_inut})")
            if not df_inu.empty:
                st.dataframe(df_inu, use_container_width=True, hide_index=True)
                xlsx_i = dataframe_para_excel_bytes(df_inu, "Inutilizadas")
                if xlsx_i:
                    st.download_button(
                        "Baixar Excel",
                        data=xlsx_i,
                        file_name="relatorio_inutilizadas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rep_inut_xlsx",
                        use_container_width=True,
                    )
            else:
                st.info("ℹ️ Nenhuma nota.")

        with tab_canc:
            _q_canc = len(df_can) if not df_can.empty else 0
            st.markdown(f"#### ❌ Canceladas ({_q_canc})")
            if not df_can.empty:
                st.dataframe(df_can, use_container_width=True, hide_index=True)
                xlsx_c = dataframe_para_excel_bytes(df_can, "Canceladas")
                if xlsx_c:
                    st.download_button(
                        "Baixar Excel",
                        data=xlsx_c,
                        file_name="relatorio_canceladas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rep_canc_xlsx",
                        use_container_width=True,
                    )
            else:
                st.info("ℹ️ Nenhuma nota.")

        with tab_aut:
            _q_aut = len(df_aut) if not df_aut.empty else 0
            st.markdown(f"#### ✅ Autorizadas ({_q_aut})")
            if not df_aut.empty:
                st.dataframe(df_aut, use_container_width=True, hide_index=True)
                xlsx_a = dataframe_para_excel_bytes(df_aut, "Autorizadas")
                if xlsx_a:
                    st.download_button(
                        "Baixar Excel",
                        data=xlsx_a,
                        file_name="relatorio_autorizadas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rep_aut_xlsx",
                        use_container_width=True,
                    )
            else:
                st.info("ℹ️ Nenhuma nota autorizada na amostra.")

        with tab_geral:
            _q_ger = len(df_ger) if not df_ger.empty else 0
            st.markdown(f"#### 📋 Relatório geral ({_q_ger})")
            if not df_ger.empty:
                st.caption(
                    "O **Excel** inclui as folhas **Geral**, **Buracos**, **Inutilizadas**, **Canceladas**, **Autorizadas**, "
                    "**CT-e lidas**, **Terceiros lidas** e **Dashboard** (indicadores e resumo por série)."
                )
                st.dataframe(df_ger, use_container_width=True, hide_index=True, height=420)
                xlsx_g = excel_relatorio_geral_com_dashboard_bytes(df_ger)
                if xlsx_g:
                    st.download_button(
                        "Baixar Excel",
                        data=xlsx_g,
                        file_name="relatorio_geral.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rep_geral_xlsx",
                        use_container_width=True,
                    )
            else:
                st.info("Relatório geral vazio.")

        st.divider()

        # =====================================================================
        # MÓDULO: DECLARAR INUTILIZADAS MANUAIS
        # =====================================================================
        st.markdown("### 🛠️ Inutilizadas sem XML")
        with st.expander(
            "Inclua notas que a Sefaz mostra inutilizadas mas que não estão no lote de ficheiros.",
            expanded=False,
        ):
            st.caption(
                "Só vale para **buracos** que o garimpeiro já listou: não cria novos buracos nem alarga intervalos. "
                "**Dos buracos:** multiselect ou colar números. **Planilha:** Excel/CSV com modelo, série e nota (várias séries num ficheiro). "
                "**Faixa:** só aplica notas que forem buraco nesse intervalo."
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
                    if st.button("Aplicar seleção", type="primary", key="inut_b_go"):
                        if not _pick_b:
                            st.warning("Selecione pelo menos um número.")
                        else:
                            with st.spinner("A atualizar…"):
                                for _nb in _pick_b:
                                    st.session_state["relatorio"].append(
                                        item_registro_manual_inutilizado(cnpj_limpo, _mb, _sb, _nb)
                                    )
                                reconstruir_dataframes_relatorio_simples()
                            st.rerun()

                    st.divider()
                    st.markdown("**Colar lista de números** (mesmo modelo e série acima)")
                    st.caption(
                        "Um número por linha (só dígitos). Notas que **não** forem buraco aqui são **ignoradas**."
                    )
                    _lista_txt = st.text_area(
                        "Números",
                        height=110,
                        key="inut_b_lista_txt",
                        placeholder="1520\n1521\n1525",
                        label_visibility="collapsed",
                    )
                    if st.button("Aplicar lista (só buracos)", type="secondary", key="inut_b_lista_go"):
                        _parsed = parse_numeros_um_por_linha(_lista_txt)
                        if not _parsed:
                            st.warning("Cole pelo menos um número válido.")
                        else:
                            _ok = []
                            for _pn in _parsed:
                                if _pn in _set_buracos and _pn not in _ok:
                                    _ok.append(_pn)
                            _ign = sum(1 for _pn in _parsed if _pn not in _set_buracos)
                            if not _ok:
                                st.warning(
                                    "Nenhum número da lista coincide com buraco neste modelo/série — nada foi aplicado."
                                )
                            else:
                                with st.spinner("A atualizar…"):
                                    for _nb in sorted(_ok):
                                        st.session_state["relatorio"].append(
                                            item_registro_manual_inutilizado(cnpj_limpo, _mb, _sb, _nb)
                                        )
                                    reconstruir_dataframes_relatorio_simples()
                                if _ign > 0:
                                    st.info(
                                        f"Aplicados **{len(_ok)}** número(s) que eram buraco. "
                                        f"**{_ign}** linha(s) ignoradas (não eram buraco neste modelo/série)."
                                    )
                                else:
                                    st.success(f"Incluídos **{len(_ok)}** registo(s).")
                                st.rerun()

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
                if st.button("Importar planilha (só buracos)", type="primary", key="inut_planilha_go"):
                    if _up_inut is None:
                        st.warning("Escolha um ficheiro primeiro.")
                    else:
                        _df_up, _err_up = dataframe_de_upload_inutil(_up_inut)
                        if _err_up:
                            st.error(_err_up)
                        else:
                            _tri, _err_tr = triplas_inutil_de_dataframe(_df_up)
                            if _err_tr:
                                st.warning(_err_tr)
                            else:
                                _df_bu = st.session_state["df_faltantes"].copy()
                                _bur_t = conjunto_triplas_buracos(_df_bu)
                                if not _bur_t:
                                    st.warning(
                                        "Não há buracos na auditoria — nada a aplicar. Faça o garimpeiro primeiro."
                                    )
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
                                        st.warning(
                                            "Nenhuma linha da planilha coincide com buraco listado pelo garimpeiro."
                                        )
                                    else:
                                        with st.spinner("A atualizar…"):
                                            for _mod, _ser, _nota in _aplic_rows:
                                                st.session_state["relatorio"].append(
                                                    item_registro_manual_inutilizado(
                                                        cnpj_limpo, _mod, _ser, _nota
                                                    )
                                                )
                                            reconstruir_dataframes_relatorio_simples()
                                        st.success(
                                            f"Incluídos **{len(_aplic_rows)}** registo(s) que eram buraco."
                                        )
                                        if _ign > 0:
                                            st.info(
                                                f"**{_ign}** linha(s) ignoradas (não eram buraco neste garimpo)."
                                            )
                                        st.rerun()

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
                    f"({len(_bur_f)} buraco(s) conhecidos)."
                )
                if st.button("Marcar buracos na faixa", type="primary", key="inut_f_go"):
                    if not _sf:
                        st.warning("Indique a série.")
                    elif _n0 > _n1:
                        st.warning("A nota inicial não pode ser maior que a final.")
                    elif (_n1 - _n0 + 1) > _MAX_FAIXA_INUT:
                        st.warning(f"Reduza a faixa (máximo {_MAX_FAIXA_INUT} notas).")
                    elif not _bur_f:
                        st.warning(
                            "Não há buracos listados para este modelo/série — verifique o garimpeiro ou a referência."
                        )
                    else:
                        _aplic = [n for n in range(int(_n0), int(_n1) + 1) if n in _bur_f]
                        _pul = (int(_n1) - int(_n0) + 1) - len(_aplic)
                        if not _aplic:
                            st.warning("Nenhum número desta faixa é buraco — nada foi aplicado.")
                        else:
                            with st.spinner("A atualizar…"):
                                for _nn in _aplic:
                                    st.session_state["relatorio"].append(
                                        item_registro_manual_inutilizado(cnpj_limpo, _mf, _sf, _nn)
                                    )
                                reconstruir_dataframes_relatorio_simples()
                            if _pul > 0:
                                st.info(
                                    f"Aplicados **{len(_aplic)}** número(s) que eram buraco na faixa. "
                                    f"**{_pul}** número(s) ignorados (não eram buraco)."
                                )
                            else:
                                st.success(f"Incluídos **{len(_aplic)}** registo(s).")
                            st.rerun()

        # =====================================================================
        # MÓDULO: DESFAZER INUTILIZAÇÃO MANUAL
        # =====================================================================
        inut_manuais = [item for item in st.session_state['relatorio'] if item.get('Arquivo') == "REGISTRO_MANUAL"]
        if inut_manuais:
            with st.expander("🔙 Desfazer inutilização manual", expanded=False):
                _df_m = pd.DataFrame(
                    [
                        {"Chave": i["Chave"], "Tipo": i["Tipo"], "Série": str(i["Série"]), "Nota": i["Número"]}
                        for i in inut_manuais
                    ]
                )
                _dm = sorted(_df_m["Tipo"].astype(str).unique())
                _mdes = st.selectbox("Modelo", _dm, key="desf_mod")
                _sub_d = _df_m[_df_m["Tipo"].astype(str) == _mdes]
                _dsers = sorted(_sub_d["Série"].astype(str).unique())
                _sdes = st.selectbox("Série", _dsers, key="desf_ser")
                _sub2_d = _sub_d[_sub_d["Série"].astype(str) == _sdes].sort_values("Nota")
                _rotulos = {
                    row["Chave"]: f"Nota n.º {int(row['Nota'])}"
                    for _, row in _sub2_d.iterrows()
                }
                _chaves_sel = st.multiselect(
                    "Remover da lista de inutilizadas:",
                    options=list(_rotulos.keys()),
                    format_func=lambda k: _rotulos.get(k, k),
                    key="desf_pick",
                )
                if st.button("Remover seleção e atualizar tabelas", key="desf_btn"):
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
        
        # =====================================================================
        # ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE
        # =====================================================================
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE")
        
        if st.session_state.get('validation_done'):
            if len(st.session_state['df_divergencias']) > 0: 
                st.warning("⚠️ Status atualizados baseados no relatório de autenticidade.")
            else: 
                st.success("✅ O status dos XMLs está alinhado com a SEFAZ.")

        with st.expander("Clique aqui para subir o Excel e atualizar o status real"):
            auth_files = st.file_uploader(
                "Suba um ou mais Excel (.xlsx / .xls) [Col A=Chave, Col F=Status]",
                type=["xlsx", "xls"],
                accept_multiple_files=True,
                key="auth_up",
            )
            st.caption(
                "Pode escolher vários relatórios de uma vez; as chaves são reunidas. "
                "Se a mesma chave aparecer em mais do que um ficheiro, prevalece o status do **último** ficheiro processado."
            )
            if auth_files and st.button("🔄 VALIDAR E ATUALIZAR"):
                auth_dict = {}
                for auth_file in auth_files:
                    df_auth = pd.read_excel(auth_file)
                    for idx, row in df_auth.iterrows():
                        chave_lida = str(row.iloc[0]).strip()
                        status_lido = str(row.iloc[5]).strip().upper()
                        if len(chave_lida) == 44:
                            auth_dict[chave_lida] = status_lido
                        
                lote_full = _lote_recalc_de_relatorio(st.session_state["relatorio"])
                ref_ar, ref_mr, ref_map = buraco_ctx_sessao()
                lote_sem_manual = {
                    k: v
                    for k, v in lote_full.items()
                    if not _item_inutil_manual_sem_xml(v[0])
                }
                H_val = _conjunto_buracos_sem_inutil_manual(lote_sem_manual, ref_ar, ref_mr, ref_map)
                drop_ch_v = set()
                for k, (res, is_p) in lote_full.items():
                    if not _item_inutil_manual_sem_xml(res):
                        continue
                    r0 = res.get("Range", (res["Número"], res["Número"]))
                    ra, rb = int(r0[0]), int(r0[1])
                    ser_s = str(res["Série"]).strip()
                    if not any((res["Tipo"], ser_s, n) in H_val for n in range(ra, rb + 1)):
                        drop_ch_v.add(k)
                if drop_ch_v:
                    st.session_state["relatorio"] = [
                        x for x in st.session_state["relatorio"] if x["Chave"] not in drop_ch_v
                    ]
                    lote_full = _lote_recalc_de_relatorio(st.session_state["relatorio"])

                lote_recalc = lote_full
                audit_map = {}
                canc_list = []
                inut_list = []
                aut_list = []
                geral_list = []
                div_list = []
                
                for k, (res, is_p) in lote_recalc.items():
                    status_final = res["Status"]
                    
                    if res["Chave"] in auth_dict and "CANCEL" in auth_dict[res["Chave"]]:
                        status_final = "CANCELADOS"
                        if res["Status"] == "NORMAIS": 
                            div_list.append({
                                "Chave": res["Chave"], 
                                "Nota": res["Número"], 
                                "Status XML": "AUTORIZADA", 
                                "Status Real": "CANCELADA"
                            })
                    
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
                        "Chave": res["Chave"], 
                        "Status Final": status_final, 
                        "Valor": res["Valor"], 
                        "Ano": res["Ano"], 
                        "Mes": res["Mes"]
                    }
                    
                    if status_final == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        ra, rb = int(r[0]), int(r[1])
                        _man_inut = _inutil_sem_xml_manual(res)
                        for n in range(ra, rb + 1):
                            if _man_inut:
                                if (res["Tipo"], str(res["Série"]).strip(), n) not in H_val:
                                    continue
                            item_inut = registro_detalhado.copy()
                            item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                            geral_list.append(item_inut)
                            if is_p:
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
                                inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                    else:
                        geral_list.append(registro_detalhado)
                        if is_p:
                            sk = (res["Tipo"], res["Série"])
                            if sk not in audit_map:
                                audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                            ult_u = ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])
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
                                if status_final == "CANCELADOS":
                                    canc_list.append(registro_detalhado)
                                elif status_final == "NORMAIS":
                                    aut_list.append(registro_detalhado)
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
                    'df_canceladas': pd.DataFrame(canc_list), 
                    'df_autorizadas': pd.DataFrame(aut_list), 
                    'df_inutilizadas': pd.DataFrame(inut_list), 
                    'df_geral': pd.DataFrame(geral_list), 
                    'df_resumo': pd.DataFrame(res_final), 
                    'df_faltantes': pd.DataFrame(fal_final), 
                    'df_divergencias': pd.DataFrame(div_list), 
                    'st_counts': {
                        "CANCELADOS": len(canc_list), 
                        "INUTILIZADOS": len(inut_list), 
                        "AUTORIZADAS": len(aut_list)
                    }, 
                    'validation_done': True
                })
                aplicar_compactacao_dfs_sessao()
                st.rerun()

        st.divider()

        # =====================================================================
        # ETAPA 3: FILTROS E EXPORTAÇÃO
        # =====================================================================
        st.markdown("### ⚙️ Etapa 3: filtros e exportação")
        with st.expander("Como isto funciona", expanded=False):
            st.markdown(
                """
<div style="background:#fff8fc;border:1px solid #f8bbd9;border-radius:10px;padding:14px 16px;margin-bottom:14px;font-size:0.93rem;line-height:1.55;color:#333;">
<b>1) Filtros (quem entra no Excel e nos XML)</b><br/>
• Em <em>cada</em> lista, <b>deixar vazia</b> = <b>não há filtro</b> nesse critério — contam todas as linhas que os outros critérios ainda deixam passar.<br/>
• <b>Listas dependentes:</b> ao escolher só «Emissão própria», por exemplo, Ano/mês, Modelo, Série, etc. mostram <b>só</b> valores que existem nessa origem no lote (não aparecem séries só de terceiros).<br/><br/>
<b>2) ZIP de XML (até 10 000 ficheiros por parte)</b> — cada parte inclui <b>sempre</b> um Excel dentro da pasta <code>RELATORIO_GARIMPEIRO/</code> só com as linhas do filtro que correspondem àqueles XMLs (não é opcional).<br/>
• <b>Com pastas</b> e/ou <b>tudo na raiz</b> — pode marcar um ou os dois tipos de ZIP.<br/><br/>
<b>3) Excel do filtro completo</b> — <b>opcional</b>: ficheiro à parte com <b>todas</b> as linhas do filtro (não entra nos ZIPs).
</div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("##### Quem entra na exportação (filtros)")
        st.caption("Origem é a única com só 2 opções fixas; as outras listas mudam conforme o que já escolheu acima.")

        for _k_v2 in ("v2_f_orig", "v2_f_mes", "v2_f_mod", "v2_f_ser", "v2_f_stat", "v2_f_op"):
            if _k_v2 not in st.session_state:
                st.session_state[_k_v2] = []

        df_g_base = st.session_state["df_geral"]
        todas_origens = ["EMISSÃO PRÓPRIA", "TERCEIROS"]
        _fo = list(st.session_state.get("v2_f_orig") or [])
        _fm = list(st.session_state.get("v2_f_mes") or [])
        _fmod = list(st.session_state.get("v2_f_mod") or [])
        _fser = list(st.session_state.get("v2_f_ser") or [])
        _fst = list(st.session_state.get("v2_f_stat") or [])
        _fop = list(st.session_state.get("v2_f_op") or [])
        _ap_mes = True

        _opts = v2_opcoes_cascata_etapa3(
            df_g_base,
            _fo,
            _fm,
            _ap_mes,
            _fmod,
            _fser,
            _fst,
            _fop,
        )
        anos_meses = _opts["anos_meses"]
        modelos = _opts["modelos"]
        series = _opts["series"]
        status_opcoes = _opts["status"]
        operacoes_opts = _opts["operacoes"]
        v2_sanear_selecoes_contra_opcoes(anos_meses, modelos, series, status_opcoes, operacoes_opts)
        _ord_mod = ["NF-e", "NFC-e", "CT-e", "MDF-e"]
        modelos = [m for m in _ord_mod if m in modelos] + sorted(
            [m for m in modelos if m not in _ord_mod],
            key=lambda x: (len(str(x)), str(x)),
        )

        _wp = st.session_state.pop("v2_preset_warn", None)
        if _wp:
            st.warning(_wp)

        with st.container():
            f_col1, f_col2, f_col3, f_col4, f_col5, f_col6 = st.columns(6)
            with f_col1:
                filtro_origem = st.multiselect(
                    "Origem (vazio = própria e terceiros)",
                    todas_origens,
                    key="v2_f_orig",
                    help="Vazio = não filtra origem. Com opções = só essas. Própria = emitente = CNPJ da sidebar.",
                )
            with f_col2:
                filtro_meses = st.multiselect(
                    "Ano / mês (vazio = todos)",
                    anos_meses,
                    key="v2_f_mes",
                    help="Vazio = qualquer competência. Lista só meses que ainda existem com os outros filtros.",
                )
            with f_col3:
                filtro_modelos = st.multiselect(
                    "Modelo (vazio = todos)",
                    modelos,
                    key="v2_f_mod",
                    help="Vazio = todos os modelos permitidos pelos outros filtros.",
                )
            with f_col4:
                filtro_series = st.multiselect(
                    "Série (vazio = todas)",
                    series,
                    key="v2_f_ser",
                    help="Vazio = todas as séries permitidas pelos filtros acima (ex.: só emissão própria → só séries da empresa).",
                )
            with f_col5:
                filtro_status = st.multiselect(
                    "Status (vazio = todos)",
                    status_opcoes,
                    key="v2_f_stat",
                    help="Vazio = todos os status. Coluna «Status Final» na exportação.",
                )
            with f_col6:
                if operacoes_opts:
                    filtro_operacao = st.multiselect(
                        "Operação (vazio = entrada e saída)",
                        operacoes_opts,
                        key="v2_f_op",
                        help="Vazio = não filtra por operação. Várias opções = união.",
                    )
                else:
                    st.caption("Operação: sem opções nos dados filtrados.")
                    filtro_operacao = []

        _c_rep, _ = st.columns([1, 5])
        with _c_rep:
            st.button(
                "Repor filtros",
                key="v2_pre_clr",
                on_click=v2_callback_repor_filtros,
            )

        aplicar_mes_so_na_propria = True

        nenhum_filtro = (
            len(filtro_origem) == 0
            and len(filtro_meses) == 0
            and len(filtro_modelos) == 0
            and len(filtro_series) == 0
            and len(filtro_status) == 0
            and len(filtro_operacao) == 0
        )
        if nenhum_filtro:
            st.caption(
                "Todas as listas vazias → exportação inclui **todas** as linhas do relatório geral (confirmação abaixo)."
            )
            confirm_export_total = st.checkbox(
                "Confirmo exportar o relatório geral completo (sem filtrar por colunas).",
                value=True,
                key="v2_confirm_full",
            )
        else:
            confirm_export_total = True

        st.markdown("##### Formato da exportação")
        st.caption(
            "Cada **parte de ZIP** (até 10 000 XML) inclui **automaticamente** um Excel com só as linhas daquele bloco, "
            "em `RELATORIO_GARIMPEIRO/`. À parte pode optar por descarregar **Excel** com **todo** o resultado do filtro."
        )
        z1, z2 = st.columns(2)
        with z1:
            st.markdown("**ZIP de XML — com pastas**")
            v2_zip_org = st.checkbox(
                "Gerar ZIP com XML **dentro de pastas** (organizado, igual à estrutura do garimpo)",
                value=True,
                key="v2_zip_org",
                help="Cada parte tem até 10 000 XML + Excel do bloco na pasta RELATORIO_GARIMPEI/.",
            )
        with z2:
            st.markdown("**ZIP de XML — tudo junto, sem pastas**")
            v2_zip_plano = st.checkbox(
                "Gerar ZIP com **todos os XML na raiz** (sem subpastas; nada separado por pasta)",
                value=True,
                key="v2_zip_plano",
                help="Cada parte tem até 10 000 XML + o mesmo Excel do bloco em RELATORIO_GARIMPEI/.",
            )
        st.markdown("**Relatório completo do filtro (opcional, fora dos ZIPs)**")
        v2_excel_completo = st.checkbox(
            "Gerar **Excel** com todas as linhas do filtro (descarregar à parte)",
            value=False,
            key="v2_excel_completo",
        )

        _quer_alguma_saida = v2_zip_org or v2_zip_plano or v2_excel_completo
        if not _quer_alguma_saida:
            st.info("Marque **pelo menos** um tipo de ZIP ou Excel completo.")

        _btn_dis = (nenhum_filtro and not confirm_export_total) or (df_g_base.empty) or (not _quer_alguma_saida)

        if st.button(
            "Gerar ficheiros (ZIPs com Excel por bloco; completo opcional)",
            type="primary",
            key="v2_btn_export",
            disabled=_btn_dis,
        ):
            df_geral_filtrado = filtrar_df_geral_para_exportacao(
                df_g_base,
                filtro_origem,
                filtro_meses,
                aplicar_mes_so_na_propria,
                filtro_modelos,
                filtro_series,
                filtro_status,
                filtro_operacao,
            )
            if df_geral_filtrado is None or df_geral_filtrado.empty:
                st.warning("Resultado filtrado: 0 linhas. Altere os filtros (multiselects).")
            else:
                with st.spinner("A gerar ZIPs (Excel por bloco) e Excel opcional…"):
                    st.session_state["excel_buffer"] = None
                    gc.collect()

                    ts = datetime.now().strftime("%Y%m%d_%H%M")
                    st.session_state["export_excel_name"] = f"relatorio_completo_{ts}.xlsx"

                    if v2_excel_completo:
                        buffer_excel = io.BytesIO()
                        with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
                            df_geral_filtrado.to_excel(writer, sheet_name="Filtrado", index=False)
                            rs = (
                                df_geral_filtrado.groupby("Status Final", dropna=False)
                                .size()
                                .reset_index(name="Quantidade")
                            )
                            rs.to_excel(writer, sheet_name="Resumo_status", index=False)
                        st.session_state["excel_buffer"] = buffer_excel.getvalue()
                    else:
                        st.session_state["excel_buffer"] = None

                    for f in os.listdir("."):
                        if f.startswith("z_org_final") or f.startswith("z_todos_final"):
                            try:
                                os.remove(f)
                            except OSError:
                                pass

                    filtro_chaves = {
                        k
                        for k in (
                            _chave_para_conjunto_export(x)
                            for x in df_geral_filtrado["Chave"].tolist()
                        )
                        if k
                    }
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
                        excel_fn = (
                            f"RELATORIO_GARIMPEIRO/relatorio_filtrado_pt{Z['seq_bloco']:03d}.xlsx"
                        )
                        xb = excel_bytes_relatorio_bloco(
                            df_geral_filtrado, Z["chaves_bloco"]
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
                            oname = f"z_org_final_pt{Z['curr_org_part']}.zip"
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
                            tname = f"z_todos_final_pt{Z['curr_todos_part']}.zip"
                            Z["z_todos"] = zipfile.ZipFile(tname, "w", zipfile.ZIP_DEFLATED)
                            Z["todos_parts"].append(tname)
                            Z["todos_count"] = 0

                    if v2_zip_org:
                        org_name = f"z_org_final_pt{Z['curr_org_part']}.zip"
                        Z["z_org"] = zipfile.ZipFile(org_name, "w", zipfile.ZIP_DEFLATED)
                        Z["org_parts"].append(org_name)
                    if v2_zip_plano:
                        todos_name = f"z_todos_final_pt{Z['curr_todos_part']}.zip"
                        Z["z_todos"] = zipfile.ZipFile(todos_name, "w", zipfile.ZIP_DEFLATED)
                        Z["todos_parts"].append(todos_name)

                    if os.path.exists(TEMP_UPLOADS_DIR) and (v2_zip_org or v2_zip_plano):
                        for f_name in os.listdir(TEMP_UPLOADS_DIR):
                            f_path = os.path.join(TEMP_UPLOADS_DIR, f_name)
                            with open(f_path, "rb") as f_temp:
                                for name, xml_data in extrair_recursivo(f_temp, f_name):
                                    res, _ = identify_xml_info(xml_data, cnpj_limpo, name)
                                    ck = (
                                        _chave_para_conjunto_export(res["Chave"])
                                        if res
                                        else None
                                    )
                                    if res and ck and ck in filtro_chaves:
                                        Z["xml_matched"] += 1
                                        Z["chaves_bloco"].add(ck)
                                        if v2_zip_org and Z["z_org"] is not None:
                                            Z["z_org"].writestr(
                                                f"{res['Pasta']}/{name}", xml_data
                                            )
                                            Z["org_count"] += 1
                                        if v2_zip_plano and Z["z_todos"] is not None:
                                            Z["z_todos"].writestr(name, xml_data)
                                            Z["todos_count"] += 1
                                        limite = (
                                            v2_zip_org
                                            and Z["org_count"] >= MAX_XML_PER_ZIP
                                        ) or (
                                            v2_zip_plano
                                            and Z["todos_count"] >= MAX_XML_PER_ZIP
                                        )
                                        if limite:
                                            _fechar_bloco_zip()
                                    del xml_data

                    if Z["chaves_bloco"] and (
                        (v2_zip_org and Z["org_count"] > 0)
                        or (v2_zip_plano and Z["todos_count"] > 0)
                    ):
                        excel_fn_last = (
                            f"RELATORIO_GARIMPEIRO/relatorio_filtrado_pt{Z['seq_bloco']:03d}.xlsx"
                        )
                        xb_last = excel_bytes_relatorio_bloco(
                            df_geral_filtrado, Z["chaves_bloco"]
                        )
                        if xb_last:
                            if v2_zip_org and Z["z_org"] is not None and Z["org_count"] > 0:
                                Z["z_org"].writestr(excel_fn_last, xb_last)
                            if (
                                v2_zip_plano
                                and Z["z_todos"] is not None
                                and Z["todos_count"] > 0
                            ):
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

                    if (v2_zip_org or v2_zip_plano) and Z.get("xml_matched", 0) == 0:
                        st.session_state["v2_export_sem_xml"] = (
                            "Nenhum XML em disco correspondeu às chaves do filtro. "
                            "Causas frequentes: pasta do garimpo apagada (use **Reset** só depois de exportar), "
                            "filtros a excluir todas as linhas com chave válida, ou chaves na tabela em formato "
                            "que não bate com o lido dos ficheiros — tente exportação completa sem filtros."
                        )
                    else:
                        st.session_state.pop("v2_export_sem_xml", None)

                    st.session_state.update(
                        {
                            "org_zip_parts": org_parts if v2_zip_org else [],
                            "todos_zip_parts": todos_parts if v2_zip_plano else [],
                            "export_ready": True,
                        }
                    )
                    gc.collect()
                st.rerun()

        if st.session_state.get("export_ready"):
            _sem = st.session_state.pop("v2_export_sem_xml", None)
            if _sem:
                st.warning(_sem)
            st.success("Geração concluída. Use **Baixar XML** para cada ZIP e **Baixar Excel** para o relatório completo, se existir.")
            _parts_o = st.session_state.get("org_zip_parts") or []
            _parts_t = st.session_state.get("todos_zip_parts") or []
            _dl_i = 0
            if _parts_o:
                st.markdown("### ZIP — XML **com pastas** (organizado)")
                for row in chunk_list(_parts_o, 3):
                    cols = st.columns(len(row))
                    for idx, part in enumerate(row):
                        _dl_i += 1
                        if os.path.exists(part):
                            with open(part, "rb") as fp:
                                cols[idx].download_button(
                                    rotulo_download_zip_parte(part),
                                    fp.read(),
                                    file_name=os.path.basename(part),
                                    key=f"v2_dlo_{_dl_i}",
                                    use_container_width=True,
                                )
            if _parts_t:
                st.markdown("### ZIP — **só ficheiros** (tudo na raiz, sem pastas)")
                for row in chunk_list(_parts_t, 3):
                    cols = st.columns(len(row))
                    for idx, part in enumerate(row):
                        _dl_i += 1
                        if os.path.exists(part):
                            with open(part, "rb") as fp:
                                cols[idx].download_button(
                                    rotulo_download_zip_parte(part),
                                    fp.read(),
                                    file_name=os.path.basename(part),
                                    key=f"v2_dlt_{_dl_i}",
                                    use_container_width=True,
                                )

            _xbuf = st.session_state.get("excel_buffer")
            if _xbuf:
                st.download_button(
                    "Baixar Excel",
                    _xbuf,
                    file_name=st.session_state.get(
                        "export_excel_name", "relatorio_completo.xlsx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="v2_dl_xlsx",
                    use_container_width=True,
                )

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


