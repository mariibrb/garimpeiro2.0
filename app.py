import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import random
import gc
import shutil
import pdfplumber
from collections import Counter, defaultdict
from datetime import date

# --- CONFIGURAÇÃO E ESTILO (CLONE ABSOLUTO DO DIAMOND TAX) ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');

        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { 
            background: radial-gradient(circle at top right, #FFDEEF 0%, #F8F9FA 100%) !important; 
        }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #FFDEEF !important;
            min-width: 400px !important;
            max-width: 400px !important;
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
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- VARIÁVEIS DE SISTEMA DE ARQUIVOS (PREVENÇÃO DE QUEDA DE MEMÓRIA) ---
TEMP_EXTRACT_DIR = "temp_garimpo_zips"
TEMP_UPLOADS_DIR = "temp_garimpo_uploads"
MAX_XML_PER_ZIP = 10000  # Máx. XMLs por ficheiro ZIP (Domínio e Etapa 3); reparte em vários lotes
# Se dois números emitidos consecutivos (ordenados) diferem mais que isto, tratamos como outra faixa.
# Assim evitamos milhões de "buracos" falsos (ex.: uma chave/XML errado com nº gigante ou duas séries distantes misturadas).
MAX_SALTO_ENTRE_NOTAS_CONSECUTIVAS = 25000

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
        if modelo is None or (isinstance(modelo, float) and pd.isna(modelo)):
            continue
        modelo = str(modelo).strip()
        if not modelo:
            continue
        serie = row.get("Série")
        if serie is None or (isinstance(serie, float) and pd.isna(serie)):
            serie = ""
        else:
            serie = str(serie).strip()
        if not serie:
            continue
        ult = row.get("Último número")
        if ult is None or (isinstance(ult, float) and pd.isna(ult)):
            continue
        try:
            u = int(float(ult))
        except (TypeError, ValueError):
            continue
        if u <= 0:
            continue
        out[f"{modelo}|{serie}"] = u
    return out


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


def parse_linhas_inutil_manual(text):
    """Linhas: MODELO|SÉRIE|NÚMERO (pipes opcionais com espaços)."""
    triplas = []
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 3:
            continue
        mod, ser, num = parts[0], parts[1], parts[2]
        d = "".join(filter(str.isdigit, num))
        if not d:
            continue
        try:
            n = int(d)
        except ValueError:
            continue
        if n <= 0 or not mod or not ser:
            continue
        triplas.append((mod, ser, n))
    return triplas


def reconstruir_dataframes_relatorio_simples():
    """Recalcula tabelas a partir de st.session_state['relatorio'] (status no próprio item)."""
    lote_recalc = {}
    for item in st.session_state["relatorio"]:
        key = item["Chave"]
        is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
        if key in lote_recalc:
            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
                lote_recalc[key] = (item, is_p)
        else:
            lote_recalc[key] = (item, is_p)

    audit_map = {}
    canc_list = []
    inut_list = []
    aut_list = []
    geral_list = []

    for k, (res, is_p) in lote_recalc.items():
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
            for n in range(r[0], r[1] + 1):
                item_inut = registro_detalhado.copy()
                item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                geral_list.append(item_inut)
        else:
            geral_list.append(registro_detalhado)

        if is_p:
            sk = (res["Tipo"], res["Série"])
            if sk not in audit_map:
                audit_map[sk] = {"nums": set(), "valor": 0.0}

            if res["Status"] == "INUTILIZADOS":
                r = res.get("Range", (res["Número"], res["Número"]))
                for n in range(r[0], r[1] + 1):
                    audit_map[sk]["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
            else:
                if res["Número"] > 0:
                    audit_map[sk]["nums"].add(res["Número"])
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
            fal_final.extend(enumerar_buracos_por_segmento(ns, t, s))

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


def coletar_numeros_por_competencia_emitente(relatorio, ano_ref, mes_ref):
    """Só emissão própria; números no mês de referência e nos meses posteriores."""
    agg = defaultdict(lambda: {"no_mes_ref": set(), "apos_ref": set()})
    for r in relatorio:
        if "EMITIDOS_CLIENTE" not in r.get("Pasta", ""):
            continue
        ano = r.get("Ano", "2000")
        mes = r.get("Mes", "01")
        if str(ano) == "0000":
            continue
        tipo = r.get("Tipo") or ""
        ser = str(r.get("Série", "0"))
        key = (tipo, ser)
        stt = r.get("Status", "")
        if stt == "INUTILIZADOS":
            ra, rb = r.get("Range", (r.get("Número", 0), r.get("Número", 0)))
            try:
                ra, rb = int(ra), int(rb)
            except (TypeError, ValueError):
                continue
            for n in range(ra, rb + 1):
                if _ym_eq(ano, mes, ano_ref, mes_ref):
                    agg[key]["no_mes_ref"].add(n)
                elif _ym_gt(ano, mes, ano_ref, mes_ref):
                    agg[key]["apos_ref"].add(n)
        else:
            try:
                n = int(r.get("Número", 0) or 0)
            except (TypeError, ValueError):
                n = 0
            if n <= 0:
                continue
            if _ym_eq(ano, mes, ano_ref, mes_ref):
                agg[key]["no_mes_ref"].add(n)
            elif _ym_gt(ano, mes, ano_ref, mes_ref):
                agg[key]["apos_ref"].add(n)
    return agg


def montar_df_conferencia_sequencia(relatorio, ano_ref, mes_ref, ref_map):
    """ref_map: chave 'Modelo|Série' -> último nº informado ao fim do mês de referência."""
    linhas = []
    agg = coletar_numeros_por_competencia_emitente(relatorio, ano_ref, mes_ref)
    for kstr, ultimo_u in ref_map.items():
        partes = kstr.split("|", 2)
        if len(partes) < 2:
            continue
        tipo, ser = partes[0].strip(), partes[1].strip()
        key = (tipo, ser)
        bucket = agg.get(key, {"no_mes_ref": set(), "apos_ref": set()})
        nr = bucket["no_mes_ref"]
        ap = bucket["apos_ref"]
        max_no_mes = max(nr) if nr else None
        min_apos = min(ap) if ap else None
        obs = []
        if max_no_mes is None:
            obs.append("Sem notas desta série no mês de referência nos XMLs.")
        elif max_no_mes > ultimo_u:
            obs.append(f"Máx. nos XMLs no mês ref. ({max_no_mes}) é maior que o último informado ({ultimo_u}).")
        elif max_no_mes < ultimo_u:
            obs.append(f"Máx. nos XMLs no mês ref. ({max_no_mes}) é menor que o último informado ({ultimo_u}).")
        else:
            obs.append("Máximo no mês de referência coincide com o último informado.")
        esperado_proximo = ultimo_u + 1
        if min_apos is None:
            obs.append("Nenhuma nota após o mês de referência neste lote.")
        elif min_apos > esperado_proximo:
            obs.append(
                f"Possível falta de sequência: primeiro nº após ref. nos XMLs é {min_apos}; esperado {esperado_proximo}."
            )
        elif min_apos < esperado_proximo:
            obs.append(
                f"Primeiro após ref. nos XMLs: {min_apos} (menor que {esperado_proximo} — verifique competência ou duplicados)."
            )
        else:
            obs.append(f"Sequência coerente: primeiro após ref. = {esperado_proximo}.")
        linhas.append(
            {
                "Modelo": tipo,
                "Série": ser,
                "Último informado (fim mês ref.)": ultimo_u,
                "Máx. XML no mês ref.": max_no_mes if max_no_mes is not None else "—",
                "Primeiro XML após ref.": min_apos if min_apos is not None else "—",
                "Observações": " ".join(obs),
            }
        )
    return pd.DataFrame(linhas)


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
                out.append({"Tipo": tipo_doc, "Serie": serie_str, "Num_Faltante": b})
    return out


# --- FUNÇÃO AUXILIAR PARA O BLOCO DOMÍNIO ---
def extrair_notas_faltantes_dominio(pdf_file):
    notas_faltantes = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                matches = re.findall(r'(\d+)\s+(\d+)\s+(\d+)\s+(?:NFe|NFCe|CTe|NF-e|NFC-e|CT-e)', text, re.IGNORECASE)
                for m in matches:
                    inicio, fim, serie = int(m[0]), int(m[1]), str(m[2])
                    for num in range(inicio, fim + 1):
                        notas_faltantes.append({"Série": serie, "Número": num})
    except: pass
    return notas_faltantes


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


def escrever_zip_dominio_por_chaves(cnpj_limpo, chaves_lista):
    """Gera um ou mais ZIPs (máx. MAX_XML_PER_ZIP XMLs cada). Retorna (lista_caminhos, total_xml)."""
    if not chaves_lista or not os.path.exists(TEMP_UPLOADS_DIR):
        return [], 0
    ch_set = set(chaves_lista)
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

    try:
        for fn in os.listdir(TEMP_UPLOADS_DIR):
            f_path = os.path.join(TEMP_UPLOADS_DIR, fn)
            with open(f_path, "rb") as ft:
                for name, data in extrair_recursivo(ft, fn):
                    res, _ = identify_xml_info(data, cnpj_limpo, name)
                    if res and res["Chave"] in ch_set:
                        if no_lote >= MAX_XML_PER_ZIP:
                            zf.close()
                            part_idx += 1
                            nome = f"faltantes_dominio_final_pt{part_idx}.zip"
                            zf = zipfile.ZipFile(nome, "w", zipfile.ZIP_DEFLATED)
                            parts.append(nome)
                            no_lote = 0
                        zf.writestr(f"{res['Pasta']}/{name}", data)
                        count_xml += 1
                        no_lote += 1
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


# Texto espelhado nos cartões + área “copiar guia” (manter alinhado ao fluxo real da app)
TEXTO_GUIA_GARIMPEIRO = """
O GARIMPEIRO — Guia rápido (para copiar)

PASSO A PASSO
1. Na barra lateral: informe o CNPJ do emitente (cliente) e clique em Liberar operação.
2. Envie ZIP ou XML soltos (volumes grandes são suportados). Depois do primeiro resultado, pode incluir mais ficheiros no topo da página, sem reiniciar o garimpo.
3. Clique em Iniciar grande garimpo e aguarde a leitura.
4. (Opcional) Na lateral: “Último nº por série” — define o mês de referência e a tabela; no ecrã aparece a conferência de sequência em relação aos XMLs.
5. (Opcional) Etapa 2: suba o Excel de autenticidade (coluna A = chave 44 dígitos; coluna F = status) para alinhar cancelamentos com a Sefaz.
6. Inutilizadas sem XML: use as abas Dos buracos (filtro por modelo/série), Faixa de números ou Colar lista.
7. Etapa 3: filtros (mês, modelo, série, status) e exportação em ZIP/Excel; cada ZIP tem até 10 mil XMLs.
8. Exportar lista específica: planilha com chaves na coluna A para gerar ZIP só com esses XMLs do lote.

O QUE O SISTEMA FAZ
• Emissão própria (CNPJ da sidebar = emitente): resumo por série, buracos na numeração, canceladas e inutilizadas; buracos são calculados por “trechos” para evitar listas falsamente enormes.
• Terceiros: totalizador por tipo (NF-e, NFC-e, CT-e, MDF-e).
• Um mesmo documento pode gerar mais do que um XML no disco (ex.: nota e evento) — o mesmo número de chaves pode corresponder a vários ficheiros.

DICAS
• Resetar sistema limpa sessão e temporários; use se trocar de cliente ou quiser recomeçar do zero.
• Modelos na app: NF-e, NFC-e, CT-e, MDF-e (use estes nomes nas tabelas e colagens).
""".strip()


# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 Como usar (passo a passo)</h3>
            <ol>
                <li><b>CNPJ:</b> Na lateral, o CNPJ do <b>emitente</b> (cliente) → Liberar operação.</li>
                <li><b>Lote:</b> Envie ZIP ou XML. Grandes volumes são suportados.</li>
                <li><b>Garimpo:</b> Iniciar grande garimpo e aguardar.</li>
                <li><b>Mais ficheiros:</b> No <b>topo dos resultados</b>, inclua XML/ZIP extra <b>sem reiniciar</b>.</li>
                <li><b>(Opcional)</b> Último nº por série (lateral) → conferência de sequência no ecrã.</li>
                <li><b>(Opcional)</b> Etapa 2 — Excel de autenticidade (chave na col. A, status na col. F).</li>
                <li><b>Inutilizadas sem XML:</b> Abas Dos buracos, Faixa ou Colar lista.</li>
                <li><b>Exportar:</b> Etapa 3 (filtros + ZIP/Excel) ou “Exportar lista específica” (chaves na col. A).</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O que o sistema faz</h3>
            <ul>
                <li><b>Emissão própria:</b> Resumo por série, buracos (por trechos), canceladas e inutilizadas separadas.</li>
                <li><b>Terceiros:</b> Contagem por tipo de documento (NF-e, CT-e, etc.).</li>
                <li><b>Filtros e lotes:</b> Exportação à medida; ZIPs partidos em até 10 mil XMLs cada.</li>
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
            key="garimpeiro_guia_copiar",
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
    'zip_dom_parts'
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
        st.session_state["seq_ref_rows"] = pd.DataFrame(
            [{"Modelo": "NF-e", "Série": "1", "Último número": None}]
        )

with st.sidebar:
    st.markdown("### 🔍 Configuração")
    cnpj_input = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_limpo = "".join(filter(str.isdigit, cnpj_input))
    
    if cnpj_input and len(cnpj_limpo) != 14: 
        st.error("⚠️ CNPJ Inválido.")
        
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): 
            st.session_state['confirmado'] = True

        with st.expander("📌 Último nº por série (mês de referência)"):
            d = date.today()
            def_ano = d.year - 1 if d.month == 1 else d.year
            def_mes = 12 if d.month == 1 else d.month - 1
            a0 = st.session_state["seq_ref_ano"] if st.session_state.get("seq_ref_ano") is not None else def_ano
            m0 = st.session_state["seq_ref_mes"] if st.session_state.get("seq_ref_mes") is not None else def_mes
            sr_ano = st.number_input("Ano de referência", min_value=2000, max_value=2100, value=int(a0), key="seq_in_ano")
            sr_mes = st.number_input("Mês (1–12)", min_value=1, max_value=12, value=int(m0), key="seq_in_mes")
            st.caption(
                "Preencha a tabela: escolha o modelo, a série e o **último número** daquele mês. "
                "Linhas vazias ou sem número são ignoradas ao guardar."
            )
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
                                    "Último número": None,
                                }
                            )
                        st.session_state["seq_ref_rows"] = pd.DataFrame(novas)
                        st.success("Só falta preencher a coluna do último número.")
                        st.rerun()
                    else:
                        st.warning("Resumo por série ainda vazio.")
            _opts = ["NF-e", "NFC-e", "CT-e", "MDF-e"]
            _tbl = st.data_editor(
                st.session_state["seq_ref_rows"],
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Modelo": st.column_config.SelectboxColumn("Modelo", options=_opts, required=True),
                    "Série": st.column_config.TextColumn("Série", required=True, max_chars=10),
                    "Último número": st.column_config.NumberColumn(
                        "Último nº", min_value=0, step=1, format="%d", help="Último emitido no mês de referência"
                    ),
                },
            )
            st.session_state["seq_ref_rows"] = _tbl.copy()
            if st.button("Guardar referência", key="seq_btn_guardar", use_container_width=True):
                parsed = ref_map_from_dataframe(st.session_state["seq_ref_rows"])
                if parsed:
                    st.session_state["seq_ref_ano"] = int(sr_ano)
                    st.session_state["seq_ref_mes"] = int(sr_mes)
                    st.session_state["seq_ref_ultimos"] = parsed
                    st.success(f"{len(parsed)} série(s) guardada(s).")
                else:
                    st.warning("Preencha pelo menos uma linha com modelo, série e último número (> 0).")
            if st.session_state.get("seq_ref_ultimos"):
                st.info(
                    f"Referência ativa: {st.session_state['seq_ref_ano']}/"
                    f"{int(st.session_state['seq_ref_mes']):02d} — "
                    f"{len(st.session_state['seq_ref_ultimos'])} série(s)."
                )
            
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
                        audit_map[sk] = {"nums": set(), "valor": 0.0}
                        
                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        for n in range(r[0], r[1] + 1):
                            audit_map[sk]["nums"].add(n)
                            inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                    else:
                        if res["Número"] > 0:
                            audit_map[sk]["nums"].add(res["Número"])
                            
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
                    
                    fal_final.extend(enumerar_buracos_por_segmento(ns, t, s))

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
                'export_ready': False
            })
            st.rerun()
    else:
        # --- RESULTADOS TELA INICIAL ---
        sc = st.session_state['st_counts']
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS (PRÓPRIAS)", sc.get("AUTORIZADAS", 0))
        c2.metric("❌ CANCELADAS (PRÓPRIAS)", sc.get("CANCELADOS", 0))
        c3.metric("🚫 INUTILIZADAS (PRÓPRIAS)", sc.get("INUTILIZADOS", 0))

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

        st.markdown("### 📊 RESUMO POR SÉRIE")
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)

        st.markdown("### 📥 TERCEIROS — TOTAL POR TIPO")
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

        if (
            st.session_state.get("seq_ref_ultimos")
            and st.session_state.get("seq_ref_ano") is not None
            and st.session_state.get("seq_ref_mes") is not None
        ):
            st.markdown("### 🔗 Conferência de sequência (vs. mês de referência)")
            ar = int(st.session_state["seq_ref_ano"])
            mr = int(st.session_state["seq_ref_mes"])
            st.caption(
                f"Competência de referência: {ar}/{mr:02d}. "
                "Só emissão própria; canceladas, autorizadas e inutilizadas entram na sequência."
            )
            df_seq = montar_df_conferencia_sequencia(
                st.session_state["relatorio"], ar, mr, st.session_state["seq_ref_ultimos"]
            )
            if df_seq.empty:
                st.info("Nenhuma série na referência para exibir.")
            else:
                st.dataframe(df_seq, use_container_width=True, hide_index=True)
        else:
            with st.expander("🔗 Conferência de sequência (opcional)", expanded=False):
                st.markdown(
                    "No menu lateral, em **Último nº por série**, escolha o mês/ano base e preencha a **tabela** "
                    "(modelo, série, último número). Depois do garimpo pode usar **Puxar séries do resumo** para "
                    "trazer modelo e série automaticamente — só falta digitar o último nº."
                )
        
        st.markdown("---")
        col_audit, col_canc, col_inut = st.columns(3)
        
        with col_audit:
            qtd_buracos = len(st.session_state['df_faltantes']) if not st.session_state['df_faltantes'].empty else 0
            st.markdown(f"### ⚠️ BURACOS ({qtd_buracos})")
            st.caption(
                "Somente emissão própria: o CNPJ do emitente do XML é o mesmo da sidebar (vale para NF-e de entrada e de saída emitidas pelo cliente). "
                "Notas em que o cliente só é destinatário (terceiros) não entram nos buracos."
            )
            if not st.session_state['df_faltantes'].empty:
                st.dataframe(st.session_state['df_faltantes'], use_container_width=True, hide_index=True)
            else: 
                st.info("✅ Tudo em ordem.")
                
        with col_canc:
            _q_canc = (
                len(st.session_state["df_canceladas"])
                if not st.session_state["df_canceladas"].empty
                else 0
            )
            st.markdown(f"### ❌ CANCELADAS ({_q_canc})")
            if not st.session_state['df_canceladas'].empty:
                st.dataframe(st.session_state['df_canceladas'], use_container_width=True, hide_index=True)
            else: 
                st.info("ℹ️ Nenhuma nota.")
                
        with col_inut:
            _q_inut = (
                len(st.session_state["df_inutilizadas"])
                if not st.session_state["df_inutilizadas"].empty
                else 0
            )
            st.markdown(f"### 🚫 INUTILIZADAS ({_q_inut})")
            if not st.session_state['df_inutilizadas'].empty:
                st.dataframe(st.session_state['df_inutilizadas'], use_container_width=True, hide_index=True)
            else: 
                st.info("ℹ️ Nenhuma nota.")

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
                "Três formas: escolher entre os **buracos** (com filtro), marcar uma **faixa** de números ou **colar** várias linhas."
            )
            tab_b, tab_f, tab_c = st.tabs(["Dos buracos", "Faixa de números", "Colar lista"])

            with tab_b:
                df_b = st.session_state["df_faltantes"]
                if df_b.empty:
                    st.info("Sem buracos na auditoria — use **Faixa** ou **Colar**, ou faça o garimpo primeiro.")
                else:
                    _mods_b = sorted(df_b["Tipo"].astype(str).unique())
                    _mb = st.selectbox("Modelo", _mods_b, key="inut_b_mod")
                    _sub_b = df_b[df_b["Tipo"].astype(str) == _mb]
                    _sers_b = sorted(_sub_b["Serie"].astype(str).unique())
                    _sb = st.selectbox("Série", _sers_b, key="inut_b_ser")
                    _sub2_b = _sub_b[_sub_b["Serie"].astype(str) == _sb]
                    _nums_b = sorted(_sub2_b["Num_Faltante"].astype(int).unique())
                    st.caption(f"{len(_nums_b)} número(s) em falta neste modelo/série.")
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

            with tab_f:
                _mf = st.selectbox("Modelo", ["NF-e", "NFC-e", "CT-e", "MDF-e"], key="inut_f_mod")
                _sf = st.text_input("Série", value="1", key="inut_f_ser").strip()
                _c1f, _c2f = st.columns(2)
                _n0 = _c1f.number_input("Nota inicial", min_value=1, value=1, step=1, key="inut_f_i")
                _n1 = _c2f.number_input("Nota final", min_value=1, value=1, step=1, key="inut_f_f")
                _MAX_FAIXA_INUT = 5000
                st.caption(f"No máximo {_MAX_FAIXA_INUT} notas por vez (proteção do sistema).")
                if st.button("Marcar faixa inteira", type="primary", key="inut_f_go"):
                    if not _sf:
                        st.warning("Indique a série.")
                    elif _n0 > _n1:
                        st.warning("A nota inicial não pode ser maior que a final.")
                    elif (_n1 - _n0 + 1) > _MAX_FAIXA_INUT:
                        st.warning(f"Reduza a faixa (máximo {_MAX_FAIXA_INUT} notas).")
                    else:
                        with st.spinner("A atualizar…"):
                            for _nn in range(int(_n0), int(_n1) + 1):
                                st.session_state["relatorio"].append(
                                    item_registro_manual_inutilizado(cnpj_limpo, _mf, _sf, _nn)
                                )
                            reconstruir_dataframes_relatorio_simples()
                        st.rerun()

            with tab_c:
                st.caption("Uma nota por linha: `NF-e|1|100` (modelo | série | número).")
                _txt_c = st.text_area(
                    "Linhas",
                    height=120,
                    key="inut_c_txt",
                    placeholder="NF-e|1|100\nNF-e|1|101",
                )
                if st.button("Importar e aplicar", type="primary", key="inut_c_go"):
                    _tri = parse_linhas_inutil_manual(_txt_c)
                    if not _tri:
                        st.warning("Nenhuma linha válida.")
                    else:
                        with st.spinner("A atualizar…"):
                            for _mod, _ser, _nota in _tri:
                                st.session_state["relatorio"].append(
                                    item_registro_manual_inutilizado(cnpj_limpo, _mod, _ser, _nota)
                                )
                            reconstruir_dataframes_relatorio_simples()
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
            auth_file = st.file_uploader("Suba o Excel (.xlsx) [Col A=Chave, Col F=Status]", type=["xlsx", "xls"], key="auth_up")
            if auth_file and st.button("🔄 VALIDAR E ATUALIZAR"):
                df_auth = pd.read_excel(auth_file)
                auth_dict = {}
                
                for idx, row in df_auth.iterrows():
                    chave_lida = str(row.iloc[0]).strip()
                    status_lido = str(row.iloc[5]).strip().upper()
                    if len(chave_lida) == 44:
                        auth_dict[chave_lida] = status_lido
                        
                lote_recalc = {}
                for item in st.session_state['relatorio']:
                    key = item["Chave"]
                    is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
                    if key in lote_recalc:
                        if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: 
                            lote_recalc[key] = (item, is_p)
                    else: 
                        lote_recalc[key] = (item, is_p)

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
                        for n in range(r[0], r[1] + 1):
                            item_inut = registro_detalhado.copy()
                            item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                            geral_list.append(item_inut)
                    else: 
                        geral_list.append(registro_detalhado)

                    if is_p:
                        sk = (res["Tipo"], res["Série"])
                        if sk not in audit_map: 
                            audit_map[sk] = {"nums": set(), "valor": 0.0}
                            
                        if status_final == "INUTILIZADOS":
                            r = res.get("Range", (res["Número"], res["Número"]))
                            for n in range(r[0], r[1] + 1): 
                                audit_map[sk]["nums"].add(n)
                                inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                        else:
                            if res["Número"] > 0:
                                audit_map[sk]["nums"].add(res["Número"])
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
                        fal_final.extend(enumerar_buracos_por_segmento(ns, t, s))
                            
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
                st.rerun()

        st.divider()

        # =====================================================================
        # ETAPA 3: FILTROS AVANÇADOS E EXPORTAÇÃO (NOVO PAINEL DE CONTROLE)
        # =====================================================================
        st.markdown("### ⚙️ ETAPA 3: FILTROS AVANÇADOS E EXPORTAÇÃO")
        
        todas_origens = ["EMISSÃO PRÓPRIA", "TERCEIROS"]
        anos_meses = sorted(list(set([f"{r.get('Ano', '0000')}/{r.get('Mes', '00')}" for r in st.session_state['relatorio'] if r.get('Ano', '0000') != '0000'])))
        modelos = sorted(list(set([r.get('Tipo', '') for r in st.session_state['relatorio']])))
        series = sorted(list(set([str(r.get('Série', '0')) for r in st.session_state['relatorio']])))
        status_opcoes = sorted(list(set([r.get('Status', '') for r in st.session_state['relatorio']]))) 
        
        with st.container():
            f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5)
            with f_col1:
                filtro_origem = st.multiselect("📌 Origem:", todas_origens)
            with f_col2:
                filtro_meses = st.multiselect("📅 Ano/Mês:", anos_meses)
                aplicar_mes_so_na_propria = st.checkbox("Aplicar Mês APENAS na Emissão Própria?", value=True)
            with f_col3:
                filtro_modelos = st.multiselect("📄 Modelo:", modelos)
            with f_col4:
                filtro_series = st.multiselect("🔢 Série:", series)
            with f_col5:
                filtro_status = st.multiselect("✅ Status:", status_opcoes) 

        if st.button("🚀 PROCESSAR E GERAR ARQUIVOS FINAIS"):
            
            with st.spinner("Buscando no HD e montando pacotes..."):
                # Limpa zips antigos
                for f in os.listdir('.'):
                    if f.startswith('z_org_final') or f.startswith('z_todos_final'):
                        try: os.remove(f)
                        except: pass

                # --- 1. APLICA FILTROS NO EXCEL ---
                df_geral_filtrado = st.session_state['df_geral'].copy()
                
                if not df_geral_filtrado.empty:
                    if len(filtro_origem) > 0:
                        df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Origem'].str.contains('|'.join([o.split()[0] for o in filtro_origem]))]
                            
                    if len(filtro_meses) > 0:
                        df_geral_filtrado['Mes_Comp'] = df_geral_filtrado['Ano'] + "/" + df_geral_filtrado['Mes']
                        if aplicar_mes_so_na_propria:
                            df_geral_filtrado = df_geral_filtrado[(df_geral_filtrado['Mes_Comp'].isin(filtro_meses)) | (df_geral_filtrado['Origem'].str.contains('TERCEIROS'))]
                        else:
                            df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Mes_Comp'].isin(filtro_meses)]
                            
                    if len(filtro_modelos) > 0:
                        df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Modelo'].isin(filtro_modelos)]
                        
                    if len(filtro_series) > 0:
                        df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Série'].astype(str).isin(filtro_series)]

                    if len(filtro_status) > 0: 
                        df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Status Final'].isin(filtro_status)]

                # Excel Master
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    df_geral_filtrado.to_excel(writer, sheet_name='Filtrado', index=False)
                st.session_state['excel_buffer'] = buffer_excel.getvalue()

                # --- 2. FILTRAGEM FÍSICA PARA ZIP (Zero RAM) ---
                org_parts, todos_parts, org_count, todos_count, curr_org_part, curr_todos_part = [], [], 0, 0, 1, 1
                org_name, todos_name = f'z_org_final_pt{curr_org_part}.zip', f'z_todos_final_pt{curr_todos_part}.zip'
                
                z_org = zipfile.ZipFile(org_name, "w", zipfile.ZIP_DEFLATED)
                z_todos = zipfile.ZipFile(todos_name, "w", zipfile.ZIP_DEFLATED)
                org_parts.append(org_name); todos_parts.append(todos_name)
                
                filtro_chaves = set(df_geral_filtrado['Chave'].tolist())

                if os.path.exists(TEMP_UPLOADS_DIR):
                    for f_name in os.listdir(TEMP_UPLOADS_DIR):
                        f_path = os.path.join(TEMP_UPLOADS_DIR, f_name)
                        with open(f_path, "rb") as f_temp:
                            for name, xml_data in extrair_recursivo(f_temp, f_name):
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res and res["Chave"] in filtro_chaves:
                                    if org_count >= MAX_XML_PER_ZIP:
                                        z_org.close(); curr_org_part += 1; org_name = f'z_org_final_pt{curr_org_part}.zip'
                                        z_org = zipfile.ZipFile(org_name, "w", zipfile.ZIP_DEFLATED); org_parts.append(org_name); org_count = 0
                                    if todos_count >= MAX_XML_PER_ZIP:
                                        z_todos.close(); curr_todos_part += 1; todos_name = f'z_todos_final_pt{curr_todos_part}.zip'
                                        z_todos = zipfile.ZipFile(todos_name, "w", zipfile.ZIP_DEFLATED); todos_parts.append(todos_name); todos_count = 0

                                    z_org.writestr(f"{res['Pasta']}/{name}", xml_data)
                                    z_todos.writestr(name, xml_data)
                                    org_count += 1; todos_count += 1
                                del xml_data
                
                z_org.close(); z_todos.close()
                st.session_state.update({'org_zip_parts': org_parts, 'todos_zip_parts': todos_parts, 'export_ready': True})
                st.rerun()

        if st.session_state.get('export_ready'):
            st.success("✅ Pacotes prontos!")
            st.markdown("### 📂 DOWNLOAD: ORGANIZADO")
            for row in chunk_list(st.session_state['org_zip_parts'], 3):
                cols = st.columns(len(row))
                for idx, part in enumerate(row):
                    with open(part, 'rb') as f:
                        cols[idx].download_button(f"📥 LOTE {part[-5]}", f.read(), part, use_container_width=True)

            st.markdown("### 📦 DOWNLOAD: SÓ XML")
            for row in chunk_list(st.session_state['todos_zip_parts'], 3):
                cols = st.columns(len(row))
                for idx, part in enumerate(row):
                    with open(part, 'rb') as f:
                        cols[idx].download_button(f"📥 LOTE {part[-5]}", f.read(), part, use_container_width=True)

            st.download_button("📊 RELATÓRIO EXCEL", st.session_state['excel_buffer'], "relatorio.xlsx", use_container_width=True)

        if st.button("⛏️ NOVO GARIMPO / LIMPAR TUDO"):
            limpar_arquivos_temp(); st.session_state.clear(); st.rerun()

        # =====================================================================
        # BLOCO 4: EXPORTAR LISTA ESPECÍFICA (PDF / EXCEL)
        # =====================================================================
        st.divider()
        st.markdown("### 🔎 EXPORTAR LISTA ESPECÍFICA")
        with st.expander("Suba planilha com chaves na coluna A para baixar os XMLs"):
            tab_pdf, tab_xlsx = st.tabs(["📄 PDF (Domínio)", "📊 Excel (chaves)"])

            with tab_pdf:
                pdf_dominio = st.file_uploader("Relatório de notas não lançadas (PDF):", type=["pdf"], key="pdf_dom_final")
                if pdf_dominio and st.button("🔎 BUSCAR XMLS NO LOTE", key="btn_run_dom"):
                    with st.spinner("Analisando e organizando arquivos..."):
                        notas_pdf = extrair_notas_faltantes_dominio(pdf_dominio)
                        if notas_pdf:
                            ch_encontradas = []
                            df_base = st.session_state["df_geral"]
                            for n in notas_pdf:
                                f = df_base[
                                    (df_base["Série"].astype(str) == n["Série"])
                                    & (df_base["Nota"] == n["Número"])
                                    & (df_base["Status Final"] == "NORMAIS")
                                ]
                                if not f.empty:
                                    ch_encontradas.append(f.iloc[0]["Chave"])

                            if ch_encontradas:
                                partes, n_xml = escrever_zip_dominio_por_chaves(cnpj_limpo, ch_encontradas)
                                if partes:
                                    st.session_state["ch_falt_dom"] = ch_encontradas
                                    st.session_state["zip_dom_parts"] = partes
                                    nl = len(partes)
                                    st.success(
                                        f"✅ Sucesso! {len(ch_encontradas)} nota(s) na lista; {n_xml} XML(s) em "
                                        f"{nl} ficheiro(s) ZIP (até {MAX_XML_PER_ZIP} XMLs por lote)."
                                    )
                                else:
                                    st.warning("⚠️ Não foi possível gerar o ZIP.")
                            else:
                                st.warning("⚠️ Nenhum XML correspondente encontrado no lote.")

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
                            partes, n_xml = escrever_zip_dominio_por_chaves(cnpj_limpo, chaves_lidas)
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

            if st.session_state.get("zip_dom_parts"):
                st.caption(f"Cada ficheiro tem no máximo {MAX_XML_PER_ZIP} XMLs.")
                for row in chunk_list(st.session_state["zip_dom_parts"], 3):
                    cols = st.columns(len(row))
                    for idx, part in enumerate(row):
                        if os.path.exists(part):
                            with open(part, "rb") as f_final:
                                cols[idx].download_button(
                                    label=f"📥 {os.path.basename(part)}",
                                    data=f_final.read(),
                                    file_name=os.path.basename(part),
                                    mime="application/zip",
                                    key=f"btn_dl_dom_{part}",
                                    use_container_width=True,
                                )
else:
    st.warning("👈 Insira o CNPJ lateral para começar.")


