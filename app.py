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

# --- MOTOR DE ALTA PERFORMANCE (REGEX PRÉ-COMPILADOS) ---
RE_TPNF = re.compile(r'<tpnf>([01])</tpnf>', re.I)
RE_EMIT = re.compile(r'<emit>.*?<cnpj>(\d+)</cnpj>', re.S | re.I)
RE_EMIT_NOME = re.compile(r'<emit>.*?<xnome>(.*?)</xnome>', re.S | re.I)
RE_DEST = re.compile(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', re.S | re.I)
RE_DEST_NOME = re.compile(r'<dest>.*?<xnome>(.*?)</xnome>', re.S | re.I)
RE_DATA = re.compile(r'<(?:dhemi|demi|dhregevento)>(\d{4}-\d{2}-\d{2})', re.I)
RE_CHAVE = re.compile(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</|id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', re.I)
RE_VALOR = re.compile(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', re.I)
RE_SERIE = re.compile(r'<serie>(\d+)</', re.I)
RE_NNFINI = re.compile(r'<nnfini>(\d+)</', re.I)
RE_NNFFIN = re.compile(r'<nnffin>(\d+)</', re.I)
RE_ANO = re.compile(r'<ano>(\d+)</', re.I)
RE_MOD = re.compile(r'<mod>(\d+)</', re.I)

# --- MOTOR DE IDENTIFICAÇÃO FISCAL INTEGRAL ---
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
        
        tp_m = RE_TPNF.search(content_str)
        if tp_m: resumo["Operacao"] = "ENTRADA" if tp_m.group(1) == "0" else "SAIDA"

        resumo["CNPJ_Emit"] = RE_EMIT.search(content_str).group(1) if RE_EMIT.search(content_str) else ""
        resumo["Nome_Emit"] = RE_EMIT_NOME.search(content_str).group(1).upper() if RE_EMIT_NOME.search(content_str) else ""
        resumo["Doc_Dest"] = RE_DEST.search(content_str).group(1) if RE_DEST.search(content_str) else ""
        resumo["Nome_Dest"] = RE_DEST_NOME.search(content_str).group(1).upper() if RE_DEST_NOME.search(content_str) else ""
        
        data_m = RE_DATA.search(content_str)
        if data_m: resumo["Data_Emissao"] = data_m.group(1)

        if any(tag in content_str.lower() for tag in ['<inutnfe', '<retinutnfe', '<procinut']):
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            mod_m = RE_MOD.search(content_str)
            if mod_m:
                if mod_m.group(1) == '65': resumo["Tipo"] = "NFC-e"
                elif mod_m.group(1) == '57': resumo["Tipo"] = "CT-e"
            resumo["Série"] = RE_SERIE.search(content_str).group(1) if RE_SERIE.search(content_str) else "0"
            ini = RE_NNFINI.search(content_str).group(1) if RE_NNFINI.search(content_str) else "0"
            fin = RE_NNFFIN.search(content_str).group(1) if RE_NNFFIN.search(content_str) else ini
            resumo["Número"], resumo["Range"] = int(ini), (int(ini), int(fin))
            ano_m = RE_ANO.search(content_str)
            resumo["Ano"] = "20" + ano_m.group(1)[-2:] if ano_m else "0000"
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            ch_m = RE_CHAVE.search(content_str)
            resumo["Chave"] = (ch_m.group(1) or ch_m.group(2)) if ch_m else ""
            if resumo["Chave"]:
                resumo["Ano"], resumo["Mes"] = "20" + resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])
                if not resumo["Data_Emissao"]: resumo["Data_Emissao"] = f"{resumo['Ano']}-{resumo['Mes']}-01"

            tipo = "NF-e"
            if '<mod>65</mod>' in content_str: tipo = "NFC-e"
            elif '<mod>57</mod>' in content_str or '<infcte' in content_str: tipo = "CT-e"
            elif '<mod>58</mod>' in content_str or '<infmdfe' in content_str: tipo = "MDF-e"
            
            status = "NORMAIS"
            if '110111' in content_str or '<cstat>101</cstat>' in content_str: status = "CANCELADOS"
            elif '110110' in content_str: status = "CARTA_CORRECAO"
            resumo["Tipo"], resumo["Status"] = tipo, status

            if status == "NORMAIS":
                val_m = RE_VALOR.search(content_str)
                resumo["Valor"] = float(val_m.group(1)) if val_m else 0.0
            
        if not resumo["CNPJ_Emit"] and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"): 
            resumo["CNPJ_Emit"] = resumo["Chave"][6:20]
        
        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        resumo["Pasta"] = f"{'EMITIDOS_CLIENTE' if is_p else 'RECEBIDOS_TERCEIROS'}/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
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
            <h3>📖 MANUAL DE OPERAÇÃO (POP)</h3>
            <ul>
                <li><b>PASSO 1:</b> Insira o CNPJ do cliente na barra lateral e clique em <b>"Liberar Operação"</b>.</li>
                <li><b>PASSO 2:</b> Suba os XMLs ou ZIPs (o sistema abre pastas infinitas) e clique em <b>"Iniciar Grande Garimpo"</b>.</li>
                <li><b>PASSO 3 (CANCELADAS):</b> O sistema identifica notas canceladas via XML e via SEFAZ. O valor delas é zerado no relatório final.</li>
                <li><b>PASSO 4:</b> Na Etapa 2, suba o Excel de Autenticidade para validar o status real perante o Governo.</li>
                <li><b>PASSO 5:</b> Use os botões no final para baixar o <b>Excel Master</b> e o <b>ZIP Organizado</b>.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O QUE VOCÊ VAI OBTER</h3>
            <ul>
                <li><b>Relatório de Buracos:</b> Identificação de notas faltantes na sequência das séries.</li>
                <li><b>Controle de Canceladas:</b> Tabela separada para auditoria de faturamento real.</li>
                <li><b>Divergências de Status:</b> Alerta de notas autorizadas no físico mas canceladas na SEFAZ.</li>
                <li><b>Organização Total:</b> XMLs movidos para pastas lógicas por Ano, Mês e Status.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

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
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    st.divider()
    if st.button("🗑️ RESETAR SISTEMA"):
        st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded_files = st.file_uploader("Arraste seus arquivos aqui:", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            lote_dict, dict_fisico = {}, {}
            buf_org, buf_todos = io.BytesIO(), io.BytesIO()
            total_arquivos = len(uploaded_files)
            with st.status("⛏️ Minerando...", expanded=True) as status_box:
                with zipfile.ZipFile(buf_org, "w", zipfile.ZIP_STORED) as z_org, zipfile.ZipFile(buf_todos, "w", zipfile.ZIP_STORED) as z_todos:
                    for i, f in enumerate(uploaded_files):
                        if i % 50 == 0: gc.collect()
                        try:
                            f.seek(0); content = f.read(); todos_xmls = extrair_recursivo(content, f.name)
                            for name, xml_data in todos_xmls:
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_dict[key] = (res, is_p)
                                    else:
                                        lote_dict[key] = (res, is_p); caminho = f"{res['Pasta']}/{name}"
                                        z_org.writestr(caminho, xml_data); z_todos.writestr(name, xml_data); dict_fisico[caminho] = xml_data
                        except: continue
                status_box.update(label="✅ Concluído!", state="complete", expanded=False)

            rel_list, audit_map, canc_list, inut_list, aut_list, geral_list = [], {}, [], [], [], []
            for k, (res, is_p) in lote_dict.items():
                rel_list.append(res); origem = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
                reg_b = {"Origem": origem, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": res["Status"], "Valor": res["Valor"]}
                if res["Status"] == "INUTILIZADOS":
                    r = res.get("Range", (res["Número"], res["Número"]))
                    for n in range(r[0], r[1] + 1): item_in = reg_b.copy(); item_in.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_in)
                else: geral_list.append(reg_b)
                if is_p:
                    sk = (res["Tipo"], res["Série"]); audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})
                    if res["Status"] == "INUTILIZADOS":
                        r = res.get("Range", (res["Número"], res["Número"]))
                        for n in range(r[0], r[1] + 1): audit_map[sk]["nums"].add(n); inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                    else:
                        if res["Número"] > 0:
                            audit_map[sk]["nums"].add(res["Número"])
                            if res["Status"] == "CANCELADOS": canc_list.append(reg_b)
                            elif res["Status"] == "NORMAIS": aut_list.append(reg_b); audit_map[sk]["valor"] += res["Valor"]
            
            rf, ff = [], []
            for (t, s), d in audit_map.items():
                ns = sorted(list(d["nums"]))
                if ns:
                    n_min, n_max = ns[0], ns[-1]; rf.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
                    for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): ff.append({"Tipo": t, "Série": s, "Nº Faltante": b})
            st.session_state.update({'z_org': buf_org.getvalue(), 'z_todos': buf_todos.getvalue(), 'relatorio': rel_list, 'dict_arquivos': dict_fisico, 'df_resumo': pd.DataFrame(rf), 'df_faltantes': pd.DataFrame(ff), 'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list), 'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}, 'garimpo_ok': True}); st.rerun()
    else:
        # --- BLOCO DE EXIBIÇÃO CORRIGIDO ---
        sc = st.session_state['st_counts']; c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc["AUTORIZADAS"]); c2.metric("❌ CANCELADAS", sc["CANCELADOS"]); c3.metric("🚫 INUTILIZADAS", sc["INUTILIZADOS"])
        st.markdown("### 📊 RESUMO POR SÉRIE"); st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.divider()
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("### ⚠️ BURACOS")
            if not st.session_state['df_faltantes'].empty: st.dataframe(st.session_state['df_faltantes'], use_container_width=True, hide_index=True)
            else: st.info("✅ OK.")
        with col2:
            st.markdown("### ❌ CANCELADAS")
            if not st.session_state['df_canceladas'].empty: st.dataframe(st.session_state['df_canceladas'], use_container_width=True, hide_index=True)
            else: st.info("ℹ️ Nada.")
        with col3:
            st.markdown("### 🚫 INUTILIZADAS")
            if not st.session_state['df_inutilizadas'].empty: st.dataframe(st.session_state['df_inutilizadas'], use_container_width=True, hide_index=True)
            else: st.info("ℹ️ Nada.")
        
        st.divider()
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE")
        with st.expander("Suba o Excel e cruze os dados"):
            auth_file = st.file_uploader("Arquivo (.xlsx)", type=["xlsx"])
            if auth_file and st.button("🔄 VALIDAR E ATUALIZAR"):
                try:
                    df_a = pd.read_excel(auth_file); a_d = {str(r.iloc[0]).strip(): str(r.iloc[5]).strip().upper() for _, r in df_a.iterrows()}
                    l_recalc = {}
                    for item in st.session_state['relatorio']:
                        k, isp = item["Chave"], "EMITIDOS_CLIENTE" in item["Pasta"]
                        if k in l_recalc:
                            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: l_recalc[k] = (item, isp)
                        else: l_recalc[k] = (item, isp)
                    a_map, c_l, i_l, au_l, g_l, d_l = {}, [], [], [], [], []
                    for k, (res, isp) in l_recalc.items():
                        st_f = res["Status"]
                        if res["Chave"] in a_d and "CANCEL" in a_d[res["Chave"]]:
                            st_f = "CANCELADOS"
                            if res["Status"] == "NORMAIS": d_l.append({"Chave": res["Chave"], "Nota": res["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})
                        reg = {"Origem": f"EMISSÃO PRÓPRIA ({res['Operacao']})" if isp else f"TERCEIROS ({res['Operacao']})", "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], "Chave": res["Chave"], "Status Final": st_f, "Valor": res["Valor"]}
                        if st_f == "INUTILIZADOS":
                            r = res.get("Range", (res["Número"], res["Número"]))
                            for n in range(r[0], r[1] + 1): item_in = reg.copy(); item_in.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); g_l.append(item_in)
                        else: g_l.append(reg)
                        if isp:
                            sk = (res["Tipo"], res["Série"]); a_map.setdefault(sk, {"nums": set(), "valor": 0.0})
                            if st_f == "INUTILIZADOS":
                                r = res.get("Range", (res["Número"], res["Número"]))
                                for n in range(r[0], r[1] + 1): a_map[sk]["nums"].add(n); i_l.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                            else:
                                if res["Número"] > 0:
                                    a_map[sk]["nums"].add(res["Número"])
                                    if st_f == "CANCELADOS": c_l.append(reg)
                                    elif st_f == "NORMAIS": au_l.append(reg); a_map[sk]["valor"] += res["Valor"]
                    rf, ff = [], []
                    for (t, s), d in a_map.items():
                        ns = sorted(list(d["nums"]))
                        if ns:
                            n_min, n_max = ns[0], ns[-1]; rf.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
                            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): ff.append({"Tipo": t, "Série": s, "Nº Faltante": b})
                    st.session_state.update({'df_canceladas': pd.DataFrame(c_l), 'df_autorizadas': pd.DataFrame(au_l), 'df_inutilizadas': pd.DataFrame(i_l), 'df_geral': pd.DataFrame(g_l), 'df_resumo': pd.DataFrame(rf), 'df_faltantes': pd.DataFrame(ff), 'df_divergencias': pd.DataFrame(d_l), 'st_counts': {"CANCELADOS": len(c_l), "INUTILIZADOS": len(i_l), "AUTORIZADAS": len(au_l)}})
                    st.success(f"✅ Auditoria Concluída! {len(a_d)} chaves cruzadas."); st.balloons(); st.rerun()
                except Exception as e: st.error(f"Erro: {e}")

        # --- EXCEL FINAL (INTEGRAL COM TODAS AS ABAS) ---
        buffer_ex = io.BytesIO()
        with pd.ExcelWriter(buffer_ex, engine='xlsxwriter') as writer:
            st.session_state['df_resumo'].to_excel(writer, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(writer, sheet_name='Geral', index=False)
            st.session_state['df_faltantes'].to_excel(writer, sheet_name='Buracos', index=False)
            st.session_state['df_canceladas'].to_excel(writer, sheet_name='Canceladas', index=False)
            st.session_state['df_inutilizadas'].to_excel(writer, sheet_name='Inutilizadas', index=False)
            st.session_state['df_autorizadas'].to_excel(writer, sheet_name='Autorizadas', index=False)
            if not st.session_state['df_divergencias'].empty: st.session_state['df_divergencias'].to_excel(writer, sheet_name='Divergencias', index=False)

        st.divider()
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1: st.download_button("📂 ZIP ORGANIZADO", st.session_state['z_org'], "garimpo.zip", use_container_width=True)
        with col_d2: st.download_button("📦 SÓ XMLs", st.session_state['z_todos'], "todos.zip", use_container_width=True)
        with col_d3: st.download_button("📊 EXCEL MASTER", buffer_ex.getvalue(), "relatorio.xlsx", use_container_width=True)
        
        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral para começar.")
