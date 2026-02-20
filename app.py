import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import random
import gc

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_externo():
    try:
        with open("style.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error("Ficheiro style.css não encontrado na pasta raiz.")

aplicar_estilo_externo()

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

# --- FUNÇÃO DE IDENTIFICAÇÃO FISCAL INTEGRAL ---
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
                    if '__MACOSX' in sub_nome: continue
                    if sub_nome.lower().endswith('.zip'): itens.extend(extrair_recursivo(z.read(sub_nome), sub_nome))
                    elif sub_nome.lower().endswith('.xml'): itens.append((os.path.basename(sub_nome), z.read(sub_nome)))
        except: pass
    elif nome_arquivo.lower().endswith('.xml'): itens.append((os.path.basename(nome_arquivo), conteudo_bytes))
    return itens

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 Instruções de Uso</h3>
            <ul>
                <li><b>Digite</b> o CNPJ do cliente na barra lateral e clique em <b>Liberar Operação</b>.</li>
                <li><b>Arraste</b> os arquivos XML ou ZIP para o campo de upload central.</li>
                <li><b>Clique</b> em <b>Iniciar Grande Garimpo</b> para mapear notas e buracos.</li>
                <li><b>Suba</b> o Excel de Autenticidade para atualizar o status real (Canceladas SEFAZ).</li>
                <li><b>Baixe</b> o ZIP Organizado e o Relatório Master no final da página.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O que será obtido?</h3>
            <ul>
                <li><b>Garimpo Profundo:</b> Abertura recursiva de ZIP dentro de ZIP.</li>
                <li><b>Tratamento de Canceladas:</b> Notas canceladas têm valor zerado no resumo.</li>
                <li><b>Auditoria SEFAZ:</b> Cruzamento automático para detetar divergências de status.</li>
                <li><b>Relatório Master:</b> Planilha detalhada com abas de Buracos, Inutilizadas e Geral.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

ks = ['garimpo_ok', 'confirmado', 'z_org', 'z_todos', 'relatorio', 'df_resumo', 'df_faltantes', 'df_canceladas', 'df_inutilizadas', 'df_autorizadas', 'df_geral', 'df_divergencias', 'st_counts', 'dict_arquivos']
for k in ks:
    if k not in st.session_state:
        if 'df' in k: st.session_state[k] = pd.DataFrame()
        elif k == 'relatorio': st.session_state[k] = []
        elif k == 'dict_arquivos': st.session_state[k] = {}
        elif k == 'st_counts': st.session_state[k] = {"CANCELADOS": 0, "INUTILIZADOS": 0, "AUTORIZADAS": 0}
        else: st.session_state[k] = False

with st.sidebar:
    cnpj_in = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_l = "".join(filter(str.isdigit, cnpj_in))
    if len(cnpj_l) == 14 and st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    st.divider()
    if st.button("🗑️ RESETAR SISTEMA"): st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded = st.file_uploader("Suba os arquivos aqui:", accept_multiple_files=True)
        if uploaded and st.button("🚀 INICIAR GRANDE GARIMPO"):
            lote_dict, dict_fisico = {}, {}
            buf_org, buf_todos = io.BytesIO(), io.BytesIO()
            with st.status("⛏️ Minerando...") as status_box:
                with zipfile.ZipFile(buf_org, "w") as z_org, zipfile.ZipFile(buf_todos, "w") as z_todos:
                    for f in uploaded:
                        xmls = extrair_recursivo(f.read(), f.name)
                        for n, d in xmls:
                            res, is_p = identify_xml_info(d, cnpj_l, n)
                            if res:
                                if res["Chave"] in lote_dict:
                                    if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_dict[res["Chave"]] = (res, is_p)
                                else:
                                    lote_dict[res["Chave"]] = (res, is_p); caminho = f"{res['Pasta']}/{n}"
                                    z_org.writestr(caminho, d); z_todos.writestr(n, d); dict_fisico[caminho] = d
            
            rel, audit, c_l, i_l, a_l, g_l = [], {}, [], [], [], []
            for k, (res, is_p) in lote_dict.items():
                rel.append(res); origem = f"{'PRÓPRIA' if is_p else 'TERCEIROS'}"
                base = {"Origem": origem, "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data": res["Data_Emissao"], "Chave": k, "Status": res["Status"], "Valor": res["Valor"]}
                if res["Status"] == "INUTILIZADOS":
                    for num in range(res["Range"][0], res["Range"][1]+1): g_l.append({**base, "Nota": num, "Valor": 0.0})
                else: g_l.append(base)
                if is_p:
                    sk = (res["Tipo"], res["Série"]); audit.setdefault(sk, {"nums": set(), "valor": 0.0})
                    if res["Status"] == "INUTILIZADOS":
                        for num in range(res["Range"][0], res["Range"][1]+1): audit[sk]["nums"].add(num); i_l.append({"Série": res["Série"], "Nota": num})
                    else:
                        audit[sk]["nums"].add(res["Número"]); audit[sk]["valor"] += res["Valor"]
                        if res["Status"] == "CANCELADOS": c_l.append(base)
                        else: a_l.append(base)
            
            rf_f, ff_f = [], []
            for (t, s), d in audit.items():
                ns = sorted(list(d["nums"]))
                if ns:
                    rf_f.append({"Modelo": t, "Série": s, "Início": ns[0], "Fim": ns[-1], "Qtd": len(ns), "Valor": round(d["valor"], 2)})
                    for b in sorted(list(set(range(ns[0], ns[-1]+1)) - set(ns))): ff_f.append({"Modelo": t, "Série": s, "Buraco": b})
            st.session_state.update({'relatorio': rel, 'dict_arquivos': dict_fisico, 'df_resumo': pd.DataFrame(rf_f), 'df_faltantes': pd.DataFrame(ff_f), 'df_canceladas': pd.DataFrame(c_l), 'df_inutilizadas': pd.DataFrame(i_l), 'df_autorizadas': pd.DataFrame(a_l), 'df_geral': pd.DataFrame(g_l), 'st_counts': {"CANCELADOS": len(c_l), "INUTILIZADOS": len(i_l), "AUTORIZADAS": len(a_l)}, 'z_org': buf_org.getvalue(), 'z_todos': buf_todos.getvalue(), 'garimpo_ok': True}); st.rerun()
    else:
        # RESULTADOS
        sc = st.session_state['st_counts']
        st.columns(3)[0].metric("📦 AUTORIZADAS", sc["AUTORIZADAS"])
        st.columns(3)[1].metric("❌ CANCELADAS", sc["CANCELADOS"])
        st.columns(3)[2].metric("🚫 INUTILIZADAS", sc["INUTILIZADOS"])
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.divider()
        col1, col2, col3 = st.columns(3)
        with col1: 
            st.markdown("### ⚠️ BURACOS")
            if not st.session_state['df_faltantes'].empty: st.dataframe(st.session_state['df_faltantes'], hide_index=True)
            else: st.info("✅ OK")
        with col2: 
            st.markdown("### ❌ CANCELADAS")
            if not st.session_state['df_canceladas'].empty: st.dataframe(st.session_state['df_canceladas'], hide_index=True)
            else: st.info("ℹ️ Nada")
        with col3: 
            st.markdown("### 🚫 INUTILIZADAS")
            if not st.session_state['df_inutilizadas'].empty: st.dataframe(st.session_state['df_inutilizadas'], hide_index=True)
            else: st.info("ℹ️ Nada")

        st.divider()
        st.markdown("### 🕵️ ETAPA 2: VALIDAR SEFAZ")
        auth_up = st.file_uploader("Suba o Excel de Autenticidade:", type=["xlsx"])
        if auth_up and st.button("🔄 VALIDAR E ATUALIZAR"):
            try:
                df_a = pd.read_excel(auth_up); a_d = {str(r.iloc[0]).strip(): str(r.iloc[5]).upper() for _, r in df_a.iterrows()}
                l_rec = {}
                for i in st.session_state['relatorio']:
                    k = i["Chave"]; isp = "EMITIDOS" in i["Pasta"]
                    if k in l_rec:
                        if i["Status"] in ["CANCELADOS", "INUTILIZADOS"]: l_rec[k] = (i, isp)
                    else: l_rec[k] = (i, isp)
                
                a_m, c_l, i_l, au_l, g_l, d_l = {}, [], [], [], [], []
                for k, (res, isp) in l_rec.items():
                    st_f = res["Status"]
                    if k in a_d and "CANCEL" in a_d[k]:
                        st_f = "CANCELADOS"
                        if res["Status"] == "NORMAIS": d_l.append({"Chave": k, "Nota": res["Número"], "Aviso": "Divergência"})
                    reg = {"Origem": f"{'PRÓPRIA' if isp else 'TERCEIROS'}", "Modelo": res["Tipo"], "Série": res["Série"], "Nota": res["Número"], "Data": res["Data_Emissao"], "Chave": k, "Status": st_f, "Valor": 0.0 if st_f == "CANCELADOS" else res["Valor"]}
                    if st_f == "INUTILIZADOS":
                        for n in range(res["Range"][0], res["Range"][1]+1): g_l.append({**reg, "Nota": n})
                    else: g_l.append(reg)
                    if isp:
                        sk = (res["Tipo"], res["Série"]); a_m.setdefault(sk, {"nums": set(), "valor": 0.0})
                        if st_f == "INUTILIZADOS":
                            for n in range(res["Range"][0], res["Range"][1]+1): a_m[sk]["nums"].add(n); i_l.append({"Série": res["Série"], "Nota": n})
                        else:
                            a_m[sk]["nums"].add(res["Número"])
                            if st_f == "CANCELADOS": c_l.append(reg)
                            else: au_l.append(reg); a_m[sk]["valor"] += res["Valor"]
                
                rf_f, ff_f = [], []
                for (t, s), d in a_m.items():
                    ns = sorted(list(d["nums"]))
                    if ns:
                        rf_f.append({"Modelo": t, "Série": s, "Início": ns[0], "Fim": ns[-1], "Qtd": len(ns), "Valor": round(d["valor"], 2)})
                        for b in sorted(list(set(range(ns[0], ns[-1]+1)) - set(ns))): ff_f.append({"Modelo": t, "Série": s, "Buraco": b})

                st.session_state.update({'df_resumo': pd.DataFrame(rf_f), 'df_faltantes': pd.DataFrame(ff_f), 'df_canceladas': pd.DataFrame(c_l), 'df_inutilizadas': pd.DataFrame(i_l), 'df_autorizadas': pd.DataFrame(au_l), 'df_geral': pd.DataFrame(g_l), 'df_divergencias': pd.DataFrame(d_l), 'st_counts': {"CANCELADOS": len(c_l), "INUTILIZADOS": len(i_l), "AUTORIZADAS": len(au_l)}})
                st.success("✅ Auditoria Finalizada!"); st.balloons(); st.rerun()
            except Exception as e: st.error(f"Erro: {e}")

        st.divider()
        buf_ex = io.BytesIO()
        with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as wr:
            st.session_state['df_resumo'].to_excel(wr, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(wr, sheet_name='Geral', index=False)
            st.session_state['df_faltantes'].to_excel(wr, sheet_name='Buracos', index=False)
            if not st.session_state['df_divergencias'].empty: st.session_state['df_divergencias'].to_excel(wr, sheet_name='Divergencias', index=False)
        
        c_d1, c_d2, c_d3 = st.columns(3)
        c_d1.download_button("📂 ZIP ORGANIZADO", st.session_state['z_org'], "garimpo.zip")
        c_d2.download_button("📦 SÓ XMLS", st.session_state['z_todos'], "todos.zip")
        c_d3.download_button("📊 EXCEL MASTER", buf_ex.getvalue(), "relatorio.xlsx")
        
        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
