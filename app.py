import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import gc

# --- CONFIGURAÇÃO E ESTILO ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');
        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { background: radial-gradient(circle at top right, #FFDEEF 0%, #F8F9FA 100%) !important; }
        [data-testid="stSidebar"] { background-color: #FFFFFF !important; border-right: 1px solid #FFDEEF !important; min-width: 400px !important; }
        div.stButton > button { 
            color: #6C757D !important; background-color: #FFFFFF !important; border: 1px solid #DEE2E6 !important;
            border-radius: 15px !important; font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important;
            height: 60px !important; text-transform: uppercase; width: 100% !important;
            transition: all 0.4s; box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important;
        }
        div.stButton > button:hover { transform: translateY(-5px) !important; border-color: #FF69B4 !important; color: #FF69B4 !important; }
        [data-testid="stFileUploader"] { border: 2px dashed #FF69B4 !important; border-radius: 20px !important; background: #FFFFFF !important; padding: 20px !important; }
        div.stDownloadButton > button { background-color: #FF69B4 !important; color: white !important; border-radius: 15px !important; width: 100% !important; font-weight: 700 !important; }
        h1, h2, h3 { font-family: 'Montserrat', sans-serif; font-weight: 800; color: #FF69B4 !important; text-align: center; }
        .instrucoes-card { background-color: rgba(255, 255, 255, 0.7); border-radius: 15px; padding: 20px; border-left: 5px solid #FF69B4; margin-bottom: 20px; min-height: 280px; }
        [data-testid="stMetric"] { background: white !important; border-radius: 20px !important; border: 1px solid #FFDEEF !important; padding: 15px !important; }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- DEPARTAMENTO 1: LEITURA E IDENTIFICAÇÃO (MOTOR) ---
def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj)))
    nome_puro = os.path.basename(file_name)
    
    if not nome_puro.lower().endswith('.xml') or nome_puro.startswith(('.', '~')):
        return None, False

    resumo = {
        "Arquivo": nome_puro, "Chave": "", "Tipo": "Outros", "Série": "0",
        "Número": 0, "Status": "NORMAIS", "Pasta": "RECEBIDOS_TERCEIROS/OUTROS",
        "Valor": 0.0, "Conteúdo": content_bytes, "Ano": "0000", "Mes": "00"
    }

    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        
        # Identificação de Inutilizadas
        if any(x in tag_l for x in ['<inutnfe', '<retinutnfe', '<procinut']):
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l: resumo["Tipo"] = "CT-e"
            
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            fin = re.search(r'<nnffin>(\d+)</', tag_l).group(1) if re.search(r'<nnffin>(\d+)</', tag_l) else ini
            resumo["Número"], resumo["Range"] = int(ini), (int(ini), int(fin))
            resumo["Ano"] = "20" + re.search(r'<ano>(\d+)</', tag_l).group(1)[-2:] if re.search(r'<ano>(\d+)</', tag_l) else "0000"
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            # Identificação de Notas Normais/Canceladas
            match_ch = re.search(r'<(?:chNFe|chCTe|chMDFe)>(\d{44})</', content_str, re.IGNORECASE)
            if not match_ch:
                match_ch = re.search(r'Id=["\'](?:NFe|CTe|MDFe)?(\d{44})["\']', content_str, re.IGNORECASE)
            
            resumo["Chave"] = match_ch.group(1) if match_ch else ""
            if resumo["Chave"]:
                resumo["Ano"], resumo["Mes"] = "20" + resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])

            # Tipo e Status
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: resumo["Tipo"] = "CT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: resumo["Tipo"] = "MDF-e"
            else: resumo["Tipo"] = "NF-e"

            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: resumo["Status"] = "CANCELADOS"
            elif '110110' in tag_l: resumo["Status"] = "CARTA_CORRECAO"

            if resumo["Status"] == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                resumo["Valor"] = float(v_match.group(1)) if v_match else 0.0

        cnpj_emit = re.search(r'<cnpj>(\d+)</cnpj>', tag_l).group(1) if re.search(r'<cnpj>(\d+)</cnpj>', tag_l) else ""
        if not cnpj_emit and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"):
            cnpj_emit = resumo["Chave"][6:20]
        
        is_p = (cnpj_emit == client_cnpj_clean)
        resumo["Pasta"] = f"EMITIDOS_CLIENTE/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}" if is_p else f"RECEBIDOS_TERCEIROS/{resumo['Tipo']}/{resumo['Ano']}/{resumo['Mes']}"
        
        return resumo, is_p
    except: return None, False

# --- DEPARTAMENTO 2: RECURSIVIDADE ---
def extrair_recursivo(conteudo_bytes, nome_arquivo):
    itens = []
    if nome_arquivo.lower().endswith('.zip'):
        try:
            with zipfile.ZipFile(io.BytesIO(conteudo_bytes)) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith('__MACOSX') or os.path.basename(sub_nome).startswith('.'): continue
                    sub_conteudo = z.read(sub_nome)
                    if sub_nome.lower().endswith('.zip'): itens.extend(extrair_recursivo(sub_conteudo, sub_nome))
                    elif sub_nome.lower().endswith('.xml'): itens.append((os.path.basename(sub_nome), sub_conteudo))
        except: pass
    elif nome_arquivo.lower().endswith('.xml'):
        itens.append((os.path.basename(nome_arquivo), conteudo_bytes))
    return itens

# --- DEPARTAMENTO 3: FUNÇÃO MESTRA DE PROCESSAMENTO (O CORAÇÃO) ---
def processar_lote_completo(lista_relatorio, auth_dict=None):
    """
    Esta função centraliza TODA a lógica de cálculo. 
    Se mudares algo aqui, muda no sistema inteiro.
    """
    lote_unico = {}
    for item in lista_relatorio:
        key = item["Chave"]
        is_p = "EMITIDOS" in item["Pasta"]
        # Prioridade para Canceladas/Inutilizadas no dicionário
        if key not in lote_unico or item["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
            lote_unico[key] = (item, is_p)

    audit_map, canc_list, inut_list, aut_list, geral_list, div_list = {}, [], [], [], [], []

    for k, (res, is_p) in lote_unico.items():
        status_final = res["Status"]
        obs = "Via XML"
        
        # Validação cruzada com Excel de Autenticidade (se fornecido)
        if auth_dict and res["Chave"] in auth_dict:
            status_excel = auth_dict[res["Chave"]]
            if "CANCEL" in status_excel:
                status_final = "CANCELADOS"
                obs = "Via Autenticidade"
                if res["Status"] == "NORMAIS":
                    div_list.append({"Chave": res["Chave"], "Nota": res["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})

        origem_txt = "EMISSÃO PRÓPRIA" if is_p else "TERCEIROS"
        
        # Lógica de Inutilizados (pode ser um range de notas)
        if status_final == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                geral_list.append({"Origem": origem_txt, "Modelo": res["Tipo"], "Série": res["Série"], "Nota": n, "Chave": res["Chave"], "Status Final": "INUTILIZADA", "Valor": 0.0})
                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                    audit_map[sk]["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
        else:
            geral_list.append({"Origem": origem_txt, "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Chave": res["Chave"], "Status Final": status_final, "Valor": res["Valor"]})
            if is_p and res["Número"] > 0:
                sk = (res["Tipo"], res["Série"])
                if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                audit_map[sk]["nums"].add(res["Número"])
                if status_final == "CANCELADOS":
                    canc_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Chave": res["Chave"], "Obs": obs})
                elif status_final == "NORMAIS":
                    aut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Valor": res["Valor"], "Chave": res["Chave"]})
                audit_map[sk]["valor"] += res["Valor"]

    # Cálculo de Buracos
    res_final, fal_final = [], []
    for (t, s), dados in audit_map.items():
        ns = sorted(list(dados["nums"]))
        if ns:
            n_min, n_max = ns[0], ns[-1]
            res_final.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(dados["valor"], 2)})
            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))):
                fal_final.append({"Tipo": t, "Série": s, "Nº Faltante": b})

    # Guardar tudo no Session State de uma vez
    st.session_state.update({
        'df_resumo': pd.DataFrame(res_final), 'df_faltantes': pd.DataFrame(fal_final),
        'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list),
        'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list),
        'df_divergencias': pd.DataFrame(div_list),
        'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}
    })

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO 2.0</h1>", unsafe_allow_html=True)

# Inicialização de variáveis
for k in ['garimpo_ok', 'confirmado', 'relatorio', 'st_counts', 'df_divergencias']:
    if k not in st.session_state:
        if 'df' in k: st.session_state[k] = pd.DataFrame()
        elif k == 'relatorio': st.session_state[k] = []
        elif k == 'st_counts': st.session_state[k] = {}
        else: st.session_state[k] = False

with st.sidebar:
    st.markdown("### 🔍 Configuração")
    cnpj_input = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_limpo = "".join(filter(str.isdigit, cnpj_input))
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    st.divider()
    if st.button("🗑️ RESETAR SISTEMA"): st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded_files = st.file_uploader("Arraste seus arquivos aqui:", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            buf_org, buf_todos = io.BytesIO(), io.BytesIO()
            with zipfile.ZipFile(buf_org, "w") as z_org, zipfile.ZipFile(buf_todos, "w") as z_todos:
                for f in uploaded_files:
                    xmls = extrair_recursivo(f.read(), f.name)
                    for name, data in xmls:
                        res, is_p = identify_xml_info(data, cnpj_limpo, name)
                        if res:
                            st.session_state['relatorio'].append(res)
                            z_org.writestr(f"{res['Pasta']}/{name}", data)
                            z_todos.writestr(name, data)
            
            st.session_state['z_org'] = buf_org.getvalue()
            st.session_state['z_todos'] = buf_todos.getvalue()
            processar_lote_completo(st.session_state['relatorio'])
            st.session_state['garimpo_ok'] = True
            st.rerun()
    else:
        # Área de Resultados e Adição
        st.markdown("### ➕ ADICIONAR MAIS ARQUIVOS")
        extra = st.file_uploader("Adicionar arquivos extras:", accept_multiple_files=True, key="extra")
        if extra and st.button("🔄 ATUALIZAR"):
            for f in extra:
                xmls = extrair_recursivo(f.read(), f.name)
                for name, data in xmls:
                    res, _ = identify_xml_info(data, cnpj_limpo, name)
                    if res: st.session_state['relatorio'].append(res)
            processar_lote_completo(st.session_state['relatorio'])
            st.rerun()

        st.divider()
        sc = st.session_state['st_counts']
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc.get("AUTORIZADAS", 0))
        c2.metric("❌ CANCELADAS", sc.get("CANCELADOS", 0))
        c3.metric("🚫 INUTILIZADAS", sc.get("INUTILIZADOS", 0))

        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        # Etapa de Validação
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM EXCEL")
        auth_f = st.file_uploader("Suba o Excel de Autenticidade:", type=["xlsx"])
        if auth_f and st.button("🔄 VALIDAR"):
            df_auth = pd.read_excel(auth_f)
            # Cria dicionário: Chave -> Status
            auth_dict = {str(r.iloc[0]).strip(): str(r.iloc[5]).strip().upper() for _, r in df_auth.iterrows() if len(str(r.iloc[0]).strip()) == 44}
            processar_lote_completo(st.session_state['relatorio'], auth_dict)
            st.rerun()

        # Botões de Download (Excel)
        buffer_ex = io.BytesIO()
        with pd.ExcelWriter(buffer_ex, engine='xlsxwriter') as writer:
            for sname, df_key in [('Resumo', 'df_resumo'), ('Geral', 'df_geral'), ('Buracos', 'df_faltantes'), ('Divergencias', 'df_divergencias')]:
                if not st.session_state[df_key].empty:
                    st.session_state[df_key].to_excel(writer, sheet_name=sname, index=False)
        
        st.download_button("📊 RELATÓRIO EXCEL", buffer_ex.getvalue(), "auditoria.xlsx")
else:
    st.warning("👈 Insira o CNPJ para começar.")
