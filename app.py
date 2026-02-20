import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import random
import gc

# --- CONFIGURAÇÃO E ESTILO (CLONE ABSOLUTO DO DIAMOND TAX) ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');
        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { background: radial-gradient(circle at top right, #FFDEEF 0%, #F8F9FA 100%) !important; }
        [data-testid="stSidebar"] { background-color: #FFFFFF !important; border-right: 1px solid #FFDEEF !important; min-width: 400px !important; max-width: 400px !important; }
        div.stButton > button { color: #6C757D !important; background-color: #FFFFFF !important; border: 1px solid #DEE2E6 !important; border-radius: 15px !important; font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important; height: 60px !important; text-transform: uppercase; transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important; width: 100% !important; box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important; }
        div.stButton > button:hover { transform: translateY(-5px) !important; box-shadow: 0 10px 20px rgba(255,105,180,0.2) !important; border-color: #FF69B4 !important; color: #FF69B4 !important; }
        [data-testid="stFileUploader"] { border: 2px dashed #FF69B4 !important; border-radius: 20px !important; background: #FFFFFF !important; padding: 20px !important; }
        div.stDownloadButton > button { background-color: #FF69B4 !important; color: white !important; border: 2px solid #FFFFFF !important; font-weight: 700 !important; border-radius: 15px !important; box-shadow: 0 0 15px rgba(255, 105, 180, 0.3) !important; text-transform: uppercase; width: 100% !important; }
        h1, h2, h3 { font-family: 'Montserrat', sans-serif; font-weight: 800; color: #FF69B4 !important; text-align: center; }
        .instrucoes-card { background-color: rgba(255, 255, 255, 0.7); border-radius: 15px; padding: 20px; border-left: 5px solid #FF69B4; margin-bottom: 20px; min-height: 280px; }
        [data-testid="stMetric"] { background: white !important; border-radius: 20px !important; border: 1px solid #FFDEEF !important; padding: 15px !important; }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- MOTOR DE IDENTIFICAÇÃO (INTEGRAL) ---
def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj))) if client_cnpj else ""
    nome_puro = os.path.basename(file_name)
    if nome_puro.startswith('.') or nome_puro.startswith('~') or not nome_puro.lower().endswith('.xml'):
        return None, False
    
    resumo = {
        "Arquivo": nome_puro, "Chave": "", "Tipo": "Outros", "Série": "0",
        "Número": 0, "Status": "NORMAIS", "Pasta": "",
        "Valor": 0.0, "Conteúdo": content_bytes, "Ano": "0000", "Mes": "00",
        "Operacao": "SAIDA", "Data_Emissao": "",
        "CNPJ_Emit": "", "Nome_Emit": "", "Doc_Dest": "", "Nome_Dest": ""
    }
    
    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        if '<?xml' not in tag_l and '<inf' not in tag_l and '<inut' not in tag_l and '<retinut' not in tag_l: return None, False
        
        tp_nf_match = re.search(r'<tpnf>([01])</tpnf>', tag_l)
        if tp_nf_match:
            resumo["Operacao"] = "ENTRADA" if tp_nf_match.group(1) == "0" else "SAIDA"

        resumo["CNPJ_Emit"] = re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S).group(1) if re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S) else ""
        resumo["Nome_Emit"] = re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""
        resumo["Doc_Dest"] = re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S).group(1) if re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S) else ""
        resumo["Nome_Dest"] = re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""

        data_match = re.search(r'<(?:dhemi|demi|dhregevento)>(\d{4}-\d{2}-\d{2})', tag_l)
        if data_match: resumo["Data_Emissao"] = data_match.group(1)

        if '<inutnfe' in tag_l or '<retinutnfe' in tag_l or '<procinut' in tag_l:
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l: resumo["Tipo"] = "CT-e"
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            fin = re.search(r'<nnffin>(\d+)</', tag_l).group(1) if re.search(r'<nnffin>(\d+)</', tag_l) else ini
            resumo["Número"] = int(ini)
            resumo["Range"] = (int(ini), int(fin))
            resumo["Ano"] = "20" + re.search(r'<ano>(\d+)</', tag_l).group(1)[-2:] if re.search(r'<ano>(\d+)</', tag_l) else "0000"
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            match_ch = re.search(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</', tag_l)
            if not match_ch:
                match_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
                resumo["Chave"] = match_ch.group(1) if match_ch else ""
            else:
                resumo["Chave"] = match_ch.group(1)

            if resumo["Chave"]:
                resumo["Ano"], resumo["Mes"] = "20" + resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])
                if not resumo["Data_Emissao"]: resumo["Data_Emissao"] = f"{resumo['Ano']}-{resumo['Mes']}-01"

            tipo = "NF-e"
            if '<mod>65</mod>' in tag_l: tipo = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: tipo = "CT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: tipo = "MDF-e"
            status = "NORMAIS"
            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: status = "CANCELADOS"
            elif '110110' in tag_l: status = "CARTA_CORRECAO"
            resumo["Tipo"], resumo["Status"] = tipo, status
            if status == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                resumo["Valor"] = float(v_match.group(1)) if v_match else 0.0
            
        if not resumo["CNPJ_Emit"] and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"): 
            resumo["CNPJ_Emit"] = resumo["Chave"][6:20]
        
        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        if is_p:
            resumo["Pasta"] = f"EMITIDOS_CLIENTE/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
        else:
            resumo["Pasta"] = f"RECEBIDOS_TERCEIROS/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Ano']}/{resumo['Mes']}"
        return resumo, is_p
    except: return None, False

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

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 Como usar o sistema (Passo a Passo)</h3>
            <ul>
                <li><b>1. Identificar a Empresa:</b> No menu branco à esquerda, escreva o <b>CNPJ do seu cliente</b> e clique no botão para liberar o sistema.</li>
                <li><b>2. Enviar as Notas:</b> No meio da tela, arraste a sua pasta de notas (pode ser em formato ZIP ou as notas XML soltas).</li>
                <li><b>3. Analisar:</b> Clique no botão <b>"Iniciar Grande Garimpo"</b> e aguarde o fim da barra de progresso.</li>
                <li><b>4. Conferir com o Governo:</b> Na Etapa 2 (final da página), envie a <b>Planilha de Autenticidade</b> da SEFAZ e clique em "Validar e Atualizar".</li>
                <li><b>5. Guardar Resultados:</b> Use os botões coloridos para descarregar o <b>Relatório Master</b> e as notas organizadas.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O que o sistema faz por si</h3>
            <ul>
                <li><b>Acha Notas Perdidas:</b> Identifica automaticamente saltos na numeração (ex: se falta a nota 5 entre a 4 e a 6).</li>
                <li><b>Limpa Cancelamentos:</b> Identifica notas canceladas e retira o valor delas para o faturamento ficar correto.</li>
                <li><b>Arruma a Casa:</b> Organiza tudo em pastas por Ano/Mês e renomeia os arquivos para fácil leitura.</li>
                <li><b>Auditoria Cruzada:</b> Confronta o status do seu arquivo físico com o que consta no site da SEFAZ.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# INICIALIZAÇÃO DE ESTADO
keys_to_init = ['garimpo_ok', 'confirmado', 'z_org', 'z_todos', 'relatorio', 'df_resumo', 'df_faltantes', 'df_canceladas', 'df_inutilizadas', 'df_autorizadas', 'df_geral', 'df_divergencias', 'st_counts', 'dict_arquivos']
for k in keys_to_init:
    if k not in st.session_state:
        if 'df' in k: st.session_state[k] = pd.DataFrame()
        elif 'z_' in k: st.session_state[k] = None
        elif k == 'relatorio': st.session_state[k] = []
        elif k == 'dict_arquivos': st.session_state[k] = {}
        elif k == 'st_counts': st.session_state[k] = {"CANCELADOS": 0, "INUTILIZADOS": 0, "AUTORIZADAS": 0}
        else: st.session_state[k] = False

with st.sidebar:
    cnpj_input = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_limpo = "".join(filter(str.isdigit, cnpj_input))
    if cnpj_input and len(cnpj_limpo) != 14: st.error("⚠️ CNPJ Inválido.")
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    st.divider()
    if st.button("🗑️ RESETAR SISTEMA"):
        st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded_files = st.file_uploader("Arraste seus arquivos aqui:", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            lote_dict = {}
            dict_fisico = {}
            buf_org, buf_todos = io.BytesIO(), io.BytesIO()
            progresso_bar = st.progress(0)
            total_arquivos = len(uploaded_files)
            
            with st.status("⛏️ Minerando...", expanded=True) as status_box:
                with zipfile.ZipFile(buf_org, "w", zipfile.ZIP_STORED) as z_org, \
                     zipfile.ZipFile(buf_todos, "w", zipfile.ZIP_STORED) as z_todos:
                    for i, f in enumerate(uploaded_files):
                        if i % 50 == 0: gc.collect()
                        progresso_bar.progress((i + 1) / total_arquivos)
                        try:
                            f.seek(0)
                            content = f.read()
                            todos_xmls = extrair_recursivo(content, f.name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_dict[key] = (res, is_p)
                                    else:
                                        lote_dict[key] = (res, is_p)
                                        caminho_completo = f"{res['Pasta']}/{name}"
                                        z_org.writestr(caminho_completo, xml_data)
                                        z_todos.writestr(name, xml_data)
                                        dict_fisico[caminho_completo] = xml_data
                        except: continue
                status_box.update(label="✅ Concluído!", state="complete", expanded=False)
            
            # --- AGREGADOR DE ANÁLISES (O CÉREBRO) ---
            rel_list, audit_map, canc_list, inut_list, aut_list, geral_list = [], {}, [], [], [], []
            for k, (res, is_p) in lote_dict.items():
                rel_list.append(res)
                origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
                reg_base = {"Origem": origem_label, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": res["Status"], "Valor": res["Valor"]}
                
                if res["Status"] == "INUTILIZADOS":
                    r = res.get("Range", (res["Número"], res["Número"]))
                    for n in range(r[0], r[1] + 1):
                        item_inut = reg_base.copy(); item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_inut)
                else: geral_list.append(reg_base)

                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        for n in range(r[0], r[1] + 1): audit_map[sk]["nums"].add(n); inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                    else:
                        if res["Número"] > 0:
                            audit_map[sk]["nums"].add(res["Número"])
                            if res["Status"] == "CANCELADOS": canc_list.append(reg_base)
                            elif res["Status"] == "NORMAIS": aut_list.append(reg_base); audit_map[sk]["valor"] += res["Valor"]

            res_final, fal_final = [], []
            for (t, s), dados in audit_map.items():
                ns = sorted(list(dados["nums"]))
                if ns:
                    res_final.append({"Documento": t, "Série": s, "Início": ns[0], "Fim": ns[-1], "Quantidade": len(ns), "Valor Contábil": round(dados["valor"], 2)})
                    for b in sorted(list(set(range(ns[0], ns[-1] + 1)) - set(ns))): fal_final.append({"Tipo": t, "Série": s, "Nº Faltante": b})

            st.session_state.update({'z_org': buf_org.getvalue(), 'z_todos': buf_todos.getvalue(), 'relatorio': rel_list, 'dict_arquivos': dict_fisico, 'df_resumo': pd.DataFrame(res_final), 'df_faltantes': pd.DataFrame(fal_final), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}, 'garimpo_ok': True})
            st.rerun()
    else:
        # --- EXIBIÇÃO DE RESULTADOS ---
        sc = st.session_state['st_counts']
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc["AUTORIZADAS"])
        c2.metric("❌ CANCELADAS", sc["CANCELADOS"])
        c3.metric("🚫 INUTILIZADAS", sc["INUTILIZADOS"])
        st.markdown("### 📊 RESUMO POR SÉRIE")
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.divider()
        col_f, col_c, col_i = st.columns(3)
        with col_f:
            st.markdown("### ⚠️ BURACOS")
            if not st.session_state['df_faltantes'].empty: st.dataframe(st.session_state['df_faltantes'], hide_index=True)
            else: st.info("✅ Tudo em ordem.")
        with col_c:
            st.markdown("### ❌ CANCELADAS")
            if not st.session_state['df_canceladas'].empty: st.dataframe(st.session_state['df_canceladas'], hide_index=True)
            else: st.info("ℹ️ Nenhuma nota.")
        with col_i:
            st.markdown("### 🚫 INUTILIZADAS")
            if not st.session_state['df_inutilizadas'].empty: st.dataframe(st.session_state['df_inutilizadas'], hide_index=True)
            else: st.info("ℹ️ Nenhuma nota.")

        st.divider()
        # --- ETAPA 2: ONDE A MÁGICA DA AUDITORIA RECONSTRÓI TUDO ---
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE")
        with st.expander("Clique aqui para subir o Excel e atualizar o status real"):
            auth_file = st.file_uploader("Suba o Excel (.xlsx) [Col A=Chave, Col F=Status]", type=["xlsx", "xls"], key="auth_up")
            if auth_file and st.button("🔄 VALIDAR E ATUALIZAR"):
                try:
                    df_auth = pd.read_excel(auth_file)
                    auth_dict = {str(row.iloc[0]).strip(): str(row.iloc[5]).strip().upper() for _, row in df_auth.iterrows() if len(str(row.iloc[0]).strip()) == 44}
                    
                    lote_recalc = {}
                    for item in st.session_state['relatorio']:
                        key = item["Chave"]
                        if key in lote_recalc:
                            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_recalc[key] = (item, "EMITIDOS" in item["Pasta"])
                        else: lote_recalc[key] = (item, "EMITIDOS" in item["Pasta"])

                    audit_map, canc_list, inut_list, aut_list, geral_list, div_list = {}, [], [], [], [], []
                    for k, (res, is_p) in lote_recalc.items():
                        st_final = res["Status"]
                        if k in auth_dict and "CANCEL" in auth_dict[k]:
                            st_final = "CANCELADOS"
                            if res["Status"] == "NORMAIS": div_list.append({"Chave": k, "Nota": res["Número"], "Aviso": "Divergência XML vs SEFAZ"})

                        reg = {"Origem": f"{'PRÓPRIA' if is_p else 'TERCEIROS'}", "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data": res["Data_Emissao"], "Chave": k, "Status": st_final, "Valor": 0.0 if st_final == "CANCELADOS" else res["Valor"]}
                        
                        if st_final == "INUTILIZADOS":
                            r = res.get("Range", (res["Número"], res["Número"]))
                            for n in range(r[0], r[1]+1): geral_list.append({**reg, "Nota": n})
                        else: geral_list.append(reg)
                        
                        if is_p:
                            sk = (res["Tipo"], res["Série"]); audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})
                            if st_final == "INUTILIZADOS":
                                r = res.get("Range", (res["Número"], res["Número"]))
                                for n in range(r[0], r[1]+1): audit_map[sk]["nums"].add(n); inut_list.append({"Série": res["Série"], "Nota": n})
                            else:
                                audit_map[sk]["nums"].add(res["Número"])
                                if st_final == "CANCELADOS": canc_list.append(reg)
                                else: aut_list.append(reg); audit_map[sk]["valor"] += res["Valor"]

                    rf_f, ff_f = [], []
                    for (t, s), d in audit_map.items():
                        ns = sorted(list(d["nums"]))
                        if ns:
                            rf_f.append({"Modelo": t, "Série": s, "Início": ns[0], "Fim": ns[-1], "Qtd": len(ns), "Valor Contábil": round(d["valor"], 2)})
                            for b in sorted(list(set(range(ns[0], ns[-1]+1)) - set(ns))): ff_f.append({"Modelo": t, "Série": s, "Nota": b})

                    st.session_state.update({'df_resumo': pd.DataFrame(rf_f), 'df_faltantes': pd.DataFrame(ff_f), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'df_divergencias': pd.DataFrame(div_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}})
                    st.rerun()
                except Exception as e: st.error(f"Erro: {e}")

        # --- EXCEL FINAL ---
        st.divider()
        buf_ex = io.BytesIO()
        with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as writer:
            st.session_state['df_resumo'].to_excel(writer, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(writer, sheet_name='Geral', index=False)
            st.session_state['df_faltantes'].to_excel(writer, sheet_name='Buracos', index=False)
            if not st.session_state['df_divergencias'].empty: st.session_state['df_divergencias'].to_excel(writer, sheet_name='Divergencias', index=False)
        
        c_d1, c_d2, c_d3 = st.columns(3)
        c_d1.download_button("📂 ZIP ORGANIZADO", st.session_state['z_org'], "garimpo.zip")
        c_d2.download_button("📦 SÓ XMLS", st.session_state['z_todos'], "todos.zip")
        c_d3.download_button("📊 EXCEL MASTER", buf_ex.getvalue(), "relatorio.xlsx")

        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral para começar.")
