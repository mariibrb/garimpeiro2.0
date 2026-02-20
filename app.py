import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import gc

# --- 1. ESTILO (INALTEADO) ---
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

# --- 2. DEPARTAMENTO FISCAL (CORREÇÃO DE PASTAS) ---
def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj)))
    nome_puro = os.path.basename(file_name)
    if not nome_puro.lower().endswith('.xml') or nome_puro.startswith(('.', '~')): return None, False

    resumo = {"Arquivo": nome_puro, "Chave": "", "Tipo": "Outros", "Série": "0", "Número": 0, "Status": "NORMAIS", "Pasta": "RECEBIDOS/OUTROS", "Valor": 0.0, "Conteúdo": content_bytes, "Ano": "0000", "Mes": "00"}
    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        
        # Identificação de Entrada/Saída
        tp_nf_match = re.search(r'<tpnf>([01])</tpnf>', tag_l)
        tp_nf_txt = "SAIDA" if (tp_nf_match and tp_nf_match.group(1) == "1") else "ENTRADA"

        if any(x in tag_l for x in ['<inutnfe', '<retinutnfe', '<procinut']):
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l: resumo["Tipo"] = "CT-e"
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            resumo["Número"] = int(ini)
            resumo["Ano"] = "20" + re.search(r'<ano>(\d+)</', tag_l).group(1)[-2:] if re.search(r'<ano>(\d+)</', tag_l) else "0000"
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            match_ch = re.search(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</', tag_l)
            if not match_ch: match_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
            resumo["Chave"] = match_ch.group(1) if match_ch else ""
            
            if resumo["Chave"]:
                resumo["Ano"], resumo["Mes"] = "20"+resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])
            
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: resumo["Tipo"] = "CT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: resumo["Tipo"] = "MDF-e"
            else: resumo["Tipo"] = "NF-e"

            if any(x in tag_l for x in ['110111', '<cstat>101</cstat>']): resumo["Status"] = "CANCELADOS"
            if resumo["Status"] == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                resumo["Valor"] = float(v_match.group(1)) if v_match else 0.0

        cnpj_emit = re.search(r'<cnpj>(\d+)</cnpj>', tag_l).group(1) if re.search(r'<cnpj>(\d+)</cnpj>', tag_l) else ""
        if not cnpj_emit and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"):
            cnpj_emit = resumo["Chave"][6:20]
        
        # DEFINIÇÃO DO CAMINHO (Removi as barras extras para o Windows aceitar)
        if (cnpj_emit == client_cnpj_clean):
            resumo["Pasta"] = f"EMITIDOS/{resumo['Tipo']}/{tp_nf_txt}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
        else:
            resumo["Pasta"] = f"RECEBIDOS/{resumo['Tipo']}/{resumo['Ano']}/{resumo['Mes']}"
            
        return resumo, (cnpj_emit == client_cnpj_clean)
    except: return None, False

# --- 3. LOGÍSTICA (RECURSIVIDADE) ---
def extrair_recursivo(conteudo_bytes, nome_arquivo):
    itens = []
    if nome_arquivo.lower().endswith('.zip'):
        try:
            with zipfile.ZipFile(io.BytesIO(conteudo_bytes)) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith('__MACOSX') or os.path.basename(sub_nome).startswith('.'): continue
                    sub_c = z.read(sub_nome)
                    if sub_nome.lower().endswith('.zip'): itens.extend(extrair_recursivo(sub_c, sub_nome))
                    elif sub_nome.lower().endswith('.xml'): itens.append((os.path.basename(sub_nome), sub_c))
        except: pass
    elif nome_arquivo.lower().endswith('.xml'): itens.append((os.path.basename(nome_arquivo), conteudo_bytes))
    return itens

# --- 4. FUNÇÃO MESTRA ---
def processar_lote_completo(lista_relatorio, auth_dict=None):
    lote_unico = {}
    for item in lista_relatorio:
        key = item["Chave"]
        if key not in lote_unico or item["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
            lote_unico[key] = (item, "EMITIDOS" in item["Pasta"])

    audit_map, canc_list, inut_list, aut_list, geral_list, div_list = {}, [], [], [], [], []
    for k, (res, is_p) in lote_unico.items():
        status_final, obs = res["Status"], "Via XML"
        if auth_dict and res["Chave"] in auth_dict:
            if "CANCEL" in auth_dict[res["Chave"]]:
                status_final, obs = "CANCELADOS", "Via Autenticidade"
                if res["Status"] == "NORMAIS":
                    div_list.append({"Chave": res["Chave"], "Nota": res["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})

        origem = "EMISSÃO PRÓPRIA" if is_p else "TERCEIROS"
        if status_final == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                geral_list.append({"Origem": origem, "Modelo": res["Tipo"], "Série": res["Série"], "Nota": n, "Chave": res["Chave"], "Status Final": "INUTILIZADA", "Valor": 0.0})
                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                    audit_map[sk]["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
        else:
            geral_list.append({"Origem": origem, "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Chave": res["Chave"], "Status Final": status_final, "Valor": res["Valor"]})
            if is_p and res["Número"] > 0:
                sk = (res["Tipo"], res["Série"])
                if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                audit_map[sk]["nums"].add(res["Número"])
                if status_final == "CANCELADOS":
                    canc_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Chave": res["Chave"], "Obs": obs})
                elif status_final == "NORMAIS":
                    aut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Valor": res["Valor"], "Chave": res["Chave"]})
                audit_map[sk]["valor"] += res["Valor"]

    res_f, fal_f = [], []
    for (t, s), dados in audit_map.items():
        ns = sorted(list(dados["nums"]))
        if ns:
            res_f.append({"Documento": t, "Série": s, "Início": ns[0], "Fim": ns[-1], "Quantidade": len(ns), "Valor Contábil (R$)": round(dados["valor"], 2)})
            for b in sorted(list(set(range(ns[0], ns[-1] + 1)) - set(ns))):
                fal_f.append({"Tipo": t, "Série": s, "Nº Faltante": b})

    st.session_state.update({'df_resumo': pd.DataFrame(res_f), 'df_faltantes': pd.DataFrame(fal_f), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'df_divergencias': pd.DataFrame(div_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}})

# --- 5. INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)
with st.container():
    c1, c2 = st.columns(2)
    with c1: st.markdown('<div class="instrucoes-card"><h3>📖 Instruções</h3><ul><li><b>Etapa 1:</b> Suba os XMLs.</li><li><b>Etapa 2:</b> Valide o status com o Excel.</li></ul></div>', unsafe_allow_html=True)
    with c2: st.markdown('<div class="instrucoes-card"><h3>📊 Resultados</h3><ul><li>ZIPs organizados por Entrada/Saída.</li><li>Relatório de Buracos e Divergências.</li></ul></div>', unsafe_allow_html=True)

for k in ['garimpo_ok', 'confirmado', 'relatorio', 'st_counts', 'df_divergencias', 'z_org', 'z_todos']:
    if k not in st.session_state:
        st.session_state[k] = pd.DataFrame() if 'df' in k else ([] if k == 'relatorio' else ({} if k == 'st_counts' else False))

with st.sidebar:
    st.markdown("### 🔍 Configuração")
    cnpj_input = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_limpo = "".join(filter(str.isdigit, cnpj_input))
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    st.divider()
    if st.button("🗑️ RESETAR TUDO"): st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        up_files = st.file_uploader("Arraste seus XMLs ou ZIPs aqui:", accept_multiple_files=True)
        if up_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            buf_org, buf_todos = io.BytesIO(), io.BytesIO()
            with zipfile.ZipFile(buf_org, "w", zipfile.ZIP_DEFLATED) as z_org, zipfile.ZipFile(buf_todos, "w", zipfile.ZIP_DEFLATED) as z_todos:
                p_bar = st.progress(0)
                for i, f in enumerate(up_files):
                    p_bar.progress((i+1)/len(up_files))
                    xmls = extrair_recursivo(f.read(), f.name)
                    for name, data in xmls:
                        res, is_p = identify_xml_info(data, cnpj_limpo, name)
                        if res:
                            st.session_state['relatorio'].append(res)
                            z_org.writestr(f"{res['Pasta']}/{name}", data)
                            z_todos.writestr(name, data)
                st.session_state['z_org'], st.session_state['z_todos'] = buf_org.getvalue(), buf_todos.getvalue()
            processar_lote_completo(st.session_state['relatorio'])
            st.session_state['garimpo_ok'] = True
            st.rerun()
    else:
        with st.expander("➕ ADICIONAR MAIS NOTAS"):
            extra = st.file_uploader("Arquivos complementares:", accept_multiple_files=True)
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

        st.markdown("### 📊 RESUMO POR SÉRIE")
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.markdown("### 🕵️ ETAPA 2: VALIDAR STATUS")
        with st.expander("Subir Excel de Autenticidade"):
            auth_f = st.file_uploader("Ficheiro .xlsx", type=["xlsx"])
            if auth_f and st.button("🔍 EXECUTAR"):
                df_auth = pd.read_excel(auth_f)
                auth_dict = {str(r.iloc[0]).strip().split('.')[0]: str(r.iloc[5]).strip().upper() for _, r in df_auth.iterrows() if len(str(r.iloc[0]).strip().split('.')[0]) == 44}
                processar_lote_completo(st.session_state['relatorio'], auth_dict)
                st.rerun()

        buf_ex = io.BytesIO()
        with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as wr:
            st.session_state['df_resumo'].to_excel(wr, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(wr, sheet_name='Geral', index=False)
            st.session_state['df_faltantes'].to_excel(wr, sheet_name='Buracos', index=False)

        col1, col2, col3 = st.columns(3)
        col1.download_button("📂 ZIP ORGANIZADO", st.session_state['z_org'], "organizado.zip")
        col2.download_button("📦 TODOS XML", st.session_state['z_todos'], "todos.zip")
        col3.download_button("📊 EXCEL FINAL", buf_ex.getvalue(), "relatorio.xlsx")
