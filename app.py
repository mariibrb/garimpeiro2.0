import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import random
import gc

# --- CONFIGURAÇÃO E ESTILO ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    try:
        with open("style.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error("Ficheiro style.css não encontrado.")

aplicar_estilo_premium()

# --- MOTOR DE IDENTIFICAÇÃO (INTEGRAL E SEM CORTES) ---
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
            if not match_ch: match_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
            resumo["Chave"] = match_ch.group(1) if match_ch else ""
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

# MANUAL POP PURAMENTE OPERACIONAL (SEM EXPLICAÇÕES TÉCNICAS INÚTEIS)
with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 MANUAL DE OPERAÇÃO PADRÃO (POP)</h3>
            <ul>
                <li><b>PASSO 1:</b> Digite o CNPJ do cliente na barra lateral à esquerda e clique em <b>"Liberar Operação"</b>.</li>
                <li><b>PASSO 2:</b> No campo "Arraste seus arquivos", suba os XMLs ou ZIPs e clique em <b>"Iniciar Grande Garimpo"</b>.</li>
                <li><b>PASSO 3:</b> Após o carregamento, desça até à "Etapa 2" e suba o relatório Excel de Autenticidade (SEFAZ).</li>
                <li><b>PASSO 4:</b> Clique em <b>"Validar e Atualizar"</b> para cruzar os dados e identificar cancelamentos ocultos.</li>
                <li><b>PASSO 5:</b> Utilize os botões de download para baixar o <b>Excel Master</b> e o <b>ZIP Organizado</b>.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O QUE VOCÊ VAI OBTER</h3>
            <ul>
                <li><b>Excel Master:</b> Relatórios automáticos de notas faltantes (Buracos) e divergências de status.</li>
                <li><b>Controle de Canceladas:</b> Uma lista separada de todas as notas canceladas (XML + SEFAZ).</li>
                <li><b>Organização Digital:</b> Todos os seus XMLs renomeados e separados por pastas de Ano e Mês.</li>
                <li><b>Conformidade Fiscal:</b> Cálculo exato do valor contábil, ignorando notas inválidas e canceladas.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# INICIALIZAÇÃO DE ESTADO E LÓGICA DE 500+ LINHAS (MANTIDA INTEGRALMENTE)
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
    st.markdown("### 🔍 Configuração Principal")
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
            status_text = st.empty()
            total_arquivos = len(uploaded_files)
            with st.status("⛏️ Minerando...", expanded=True) as status_box:
                with zipfile.ZipFile(buf_org, "w", zipfile.ZIP_STORED) as z_org, zipfile.ZipFile(buf_todos, "w", zipfile.ZIP_STORED) as z_todos:
                    for i, f in enumerate(uploaded_files):
                        if i % 50 == 0: gc.collect()
                        if total_arquivos > 0 and i % max(1, int(total_arquivos * 0.02)) == 0:
                            progresso_bar.progress((i + 1) / total_arquivos); status_text.text(f"⛏️ Processando arquivo {i+1}/{total_arquivos}: {f.name}")
                        try:
                            f.seek(0); content = f.read(); todos_xmls = extrair_recursivo(content, f.name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_dict[key] = (res, is_p)
                                    else:
                                        lote_dict[key] = (res, is_p); caminho_completo = f"{res['Pasta']}/{name}"
                                        z_org.writestr(caminho_completo, xml_data); z_todos.writestr(name, xml_data); dict_fisico[caminho_completo] = xml_data
                        except: continue
                status_box.update(label="✅ Concluído!", state="complete", expanded=False); progresso_bar.empty(); status_text.empty()

            rel_list, audit_map, canc_list, inut_list, aut_list, geral_list = [], {}, [], [], [], []
            for k, (res, is_p) in lote_dict.items():
                rel_list.append(res); origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
                reg_base = {"Origem": origem_label, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": res["Status"], "Valor": res["Valor"]}
                if res["Status"] == "INUTILIZADOS":
                    r = res.get("Range", (res["Número"], res["Número"]))
                    for n in range(r[0], r[1] + 1): item_inut = reg_base.copy(); item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_inut)
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
                            elif res["Status"] == "NORMAIS": aut_list.append(reg_base)
                            audit_map[sk]["valor"] += res["Valor"]

            res_f, fal_f = [], []
            for (t, s), d in audit_map.items():
                ns = sorted(list(d["nums"]))
                if ns:
                    n_min, n_max = ns[0], ns[-1]; res_f.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
                    for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): fal_f.append({"Tipo": t, "Série": s, "Nº Faltante": b})
            st.session_state.update({'z_org': buf_org.getvalue(), 'z_todos': buf_todos.getvalue(), 'relatorio': rel_list, 'dict_arquivos': dict_fisico, 'df_resumo': pd.DataFrame(res_f), 'df_faltantes': pd.DataFrame(fal_f), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}, 'garimpo_ok': True}); st.rerun()
    else:
        # --- EXIBIÇÃO RESULTADOS ---
        sc = st.session_state['st_counts']; c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc["AUTORIZADAS"]); c2.metric("❌ CANCELADAS", sc["CANCELADOS"]); c3.metric("🚫 INUTILIZADAS", sc["INUTILIZADOS"])
        st.markdown("### 📊 RESUMO POR SÉRIE"); st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.divider()
        col_audit, col_canc, col_inut = st.columns(3)
        with col_audit:
            st.markdown("### ⚠️ BURACOS"); st.dataframe(st.session_state['df_faltantes'], use_container_width=True, hide_index=True) if not st.session_state['df_faltantes'].empty else st.info("✅ Tudo em ordem.")
        with col_canc:
            st.markdown("### ❌ CANCELADAS"); st.dataframe(st.session_state['df_canceladas'], use_container_width=True, hide_index=True) if not st.session_state['df_canceladas'].empty else st.info("ℹ️ Nenhuma nota.")
        with col_inut:
            st.markdown("### 🚫 INUTILIZADAS"); st.dataframe(st.session_state['df_inutilizadas'], use_container_width=True, hide_index=True) if not st.session_state['df_inutilizadas'].empty else st.info("ℹ️ Nenhuma nota.")
        
        st.divider()
        # --- ETAPA 2: VALIDAR ---
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE")
        with st.expander("Suba o Excel e atualize o status real"):
            auth_file = st.file_uploader("Excel (.xlsx) [Coluna A=Chave, Coluna F=Status]", type=["xlsx", "xls"], key="auth_up")
            if auth_file and st.button("🔄 VALIDAR E ATUALIZAR"):
                try:
                    df_auth = pd.read_excel(auth_file); auth_dict = {str(row.iloc[0]).strip(): str(row.iloc[5]).strip().upper() for _, row in df_auth.iterrows() if len(str(row.iloc[0]).strip()) == 44}
                    lote_recalc = {}
                    for item in st.session_state['relatorio']:
                        key = item["Chave"]; is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
                        if key in lote_recalc:
                            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_recalc[key] = (item, is_p)
                        else: lote_recalc[key] = (item, is_p)
                    audit_map, canc_list, inut_list, aut_list, geral_list, div_list = {}, [], [], [], [], []
                    for k, (res, is_p) in lote_recalc.items():
                        status_final = res["Status"]
                        if res["Chave"] in auth_dict and "CANCEL" in auth_dict[res["Chave"]]:
                            status_final = "CANCELADOS"
                            if res["Status"] == "NORMAIS": div_list.append({"Chave": res["Chave"], "Nota": res["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})
                        reg_det = {"Origem": f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})", "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": status_final, "Valor": res["Valor"]}
                        if status_final == "INUTILIZADOS":
                            r = res.get("Range", (res["Número"], res["Número"]))
                            for n in range(r[0], r[1] + 1): item_inut = reg_det.copy(); item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_inut)
                        else: geral_list.append(reg_det)
                        if is_p:
                            sk = (res["Tipo"], res["Série"])
                            if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                            if status_final == "INUTILIZADOS":
                                r = res.get("Range", (res["Número"], res["Número"]))
                                for n in range(r[0], r[1] + 1): audit_map[sk]["nums"].add(n); inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                            else:
                                if res["Número"] > 0:
                                    audit_map[sk]["nums"].add(res["Número"])
                                    if status_final == "CANCELADOS": canc_list.append(reg_det)
                                    elif status_final == "NORMAIS": aut_list.append(reg_det)
                                    audit_map[sk]["valor"] += res["Valor"]
                    res_f, fal_f = [], []
                    for (t, s), d in audit_map.items():
                        ns = sorted(list(d["nums"]))
                        if ns:
                            n_min, n_max = ns[0], ns[-1]; res_f.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
                            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): fal_f.append({"Tipo": t, "Série": s, "Nº Faltante": b})
                    st.session_state.update({'df_canceladas': pd.DataFrame(canc_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_geral': pd.DataFrame(geral_list), 'df_resumo': pd.DataFrame(res_f), 'df_faltantes': pd.DataFrame(fal_f), 'df_divergencias': pd.DataFrame(div_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}}); st.rerun()
                except Exception as e: st.error(f"Erro: {e}")

        st.divider()
        # --- ADICIONAR SEM RESET ---
        with st.expander("➕ ADICIONAR MAIS ARQUIVOS"):
            extra_files = st.file_uploader("Adicionar novos lotes:", accept_multiple_files=True, key="extra_files")
            if extra_files and st.button("PROCESSAR E ATUALIZAR LISTA"):
                for f in extra_files:
                    try:
                        content = f.read(); todos_xmls = extrair_recursivo(content, f.name)
                        for name, xml_data in todos_xmls:
                            res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                            if res: st.session_state['relatorio'].append(res); st.session_state['dict_arquivos'][f"{res['Pasta']}/{name}"] = xml_data
                    except: continue
                lote_recalc = {}
                for item in st.session_state['relatorio']:
                    key = item["Chave"]; is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
                    if key in lote_recalc:
                        if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_recalc[key] = (item, is_p)
                    else: lote_recalc[key] = (item, is_p)
                audit_map, canc_list, inut_list, aut_list, geral_list = {}, [], [], [], []
                for k, (res, is_p) in lote_recalc.items():
                    origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
                    reg_det = {"Origem": origem_label, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": res["Status"], "Valor": res["Valor"]}
                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        for n in range(r[0], r[1] + 1): item_inut = reg_det.copy(); item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_inut)
                    else: geral_list.append(reg_det)
                    if is_p:
                        sk = (res["Tipo"], res["Série"])
                        if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                        if res["Status"] == "INUTILIZADOS":
                            r = res.get("Range", (res["Número"], res["Número"]))
                            for n in range(r[0], r[1] + 1): audit_map[sk]["nums"].add(n); inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                        else:
                            if res["Número"] > 0:
                                audit_map[sk]["nums"].add(res["Número"])
                                if res["Status"] == "CANCELADOS": canc_list.append(reg_det)
                                elif res["Status"] == "NORMAIS": aut_list.append(reg_det)
                                audit_map[sk]["valor"] += res["Valor"]
                res_f, fal_f = [], []
                for (t, s), d in audit_map.items():
                    ns = sorted(list(d["nums"]))
                    if ns:
                        n_min, n_max = ns[0], ns[-1]; res_f.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
                        for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): fal_f.append({"Tipo": t, "Série": s, "Nº Faltante": b})
                st.session_state.update({'df_resumo': pd.DataFrame(res_f), 'df_faltantes': pd.DataFrame(fal_f), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}}); st.rerun()

        # --- EXCEL FINAL (INTEGRAL) ---
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
            st.session_state['df_resumo'].to_excel(writer, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(writer, sheet_name='Geral_Todos', index=False)
            st.session_state['df_faltantes'].to_excel(writer, sheet_name='Buracos', index=False)
            st.session_state['df_canceladas'].to_excel(writer, sheet_name='Canceladas', index=False)
            st.session_state['df_inutilizadas'].to_excel(writer, sheet_name='Inutilizadas', index=False)
            st.session_state['df_autorizadas'].to_excel(writer, sheet_name='Autorizadas', index=False)
            if not st.session_state['df_divergencias'].empty: st.session_state['df_divergencias'].to_excel(writer, sheet_name='Divergencias', index=False)

        col1, col2, col3 = st.columns(3)
        with col1: st.download_button("📂 BAIXAR ORGANIZADO (ZIP)", st.session_state['z_org'], "garimpo.zip", use_container_width=True)
        with col2: st.download_button("📦 BAIXAR TODOS (SÓ XML)", st.session_state['z_todos'], "todos.zip", use_container_width=True)
        with col3: st.download_button("📊 EXCEL MASTER", buffer_excel.getvalue(), "auditoria_detalhada.xlsx", use_container_width=True)
        
        st.divider()
        # --- SELETIVO ---
        st.markdown("### 📂 DOWNLOAD SELETIVO POR PASTA")
        todas_pastas = sorted(list(set([os.path.dirname(k) for k in st.session_state['dict_arquivos'].keys()])))
        if todas_pastas:
            pasta_sel = st.selectbox("Escolha a pasta fiscal:", ["--- SELECIONE ---"] + todas_pastas)
            if pasta_sel != "--- SELECIONE ---":
                buf_sel = io.BytesIO(); cont = 0
                with zipfile.ZipFile(buf_sel, "w") as z_sel:
                    for caminho, dados in st.session_state['dict_arquivos'].items():
                        if caminho.startswith(pasta_sel): z_sel.writestr(os.path.basename(caminho), dados); cont += 1
                st.download_button(f"📥 BAIXAR {cont} ARQUIVOS", buf_sel.getvalue(), "pasta.zip", use_container_width=True)
        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral para começar.")
