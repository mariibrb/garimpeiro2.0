import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import gc

# --- CONFIGURAÇÃO E ESTILO ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def carregar_estilo():
    """Lê o ficheiro CSS externo para manter o Python limpo."""
    try:
        with open("style.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("⚠️ Ficheiro style.css não encontrado na pasta raiz.")

carregar_estilo()

# --- MOTOR DE IDENTIFICAÇÃO (REGEX OTIMIZADO) ---
RE_CHAVE = re.compile(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</|id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', re.I)
RE_VALOR = re.compile(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', re.I)

def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj))) if client_cnpj else ""
    nome_puro = os.path.basename(file_name)
    if nome_puro.startswith(('.', '~')) or not nome_puro.lower().endswith('.xml'):
        return None, False
    
    resumo = {
        "Arquivo": nome_puro, "Chave": "", "Tipo": "Outros", "Série": "0",
        "Número": 0, "Status": "NORMAIS", "Pasta": "",
        "Valor": 0.0, "Conteúdo": content_bytes, "Ano": "0000", "Mes": "00",
        "Operacao": "SAIDA", "Data_Emissao": "",
        "CNPJ_Emit": "", "Nome_Emit": "", "Doc_Dest": "", "Nome_Dest": ""
    }
    
    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore').lower()
        if not any(x in content_str for x in ['<?xml', '<inf', '<inut']): return None, False
        
        # Identificação de tpNF
        tp_nf_match = re.search(r'<tpnf>([01])</tpnf>', content_str)
        if tp_nf_match: resumo["Operacao"] = "ENTRADA" if tp_nf_match.group(1) == "0" else "SAIDA"

        # Extração de Dados das Partes
        resumo["CNPJ_Emit"] = re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', content_str, re.S).group(1) if re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', content_str, re.S) else ""
        resumo["Nome_Emit"] = re.search(r'<emit>.*?<xnome>(.*?)</xnome>', content_str, re.S).group(1).upper() if re.search(r'<emit>.*?<xnome>(.*?)</xnome>', content_str, re.S) else ""
        resumo["Doc_Dest"] = re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', content_str, re.S).group(1) if re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', content_str, re.S) else ""
        resumo["Nome_Dest"] = re.search(r'<dest>.*?<xnome>(.*?)</xnome>', content_str, re.S).group(1).upper() if re.search(r'<dest>.*?<xnome>(.*?)</xnome>', content_str, re.S) else ""

        # Data de Emissão
        data_match = re.search(r'<(?:dhemi|demi|dhregevento)>(\d{4}-\d{2}-\d{2})', content_str)
        if data_match: resumo["Data_Emissao"] = data_match.group(1)

        # 1. INUTILIZADAS
        if '<inutnfe' in content_str or '<retinutnfe' in content_str:
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            if '<mod>65</mod>' in content_str: resumo["Tipo"] = "NFC-e"
            resumo["Série"] = re.search(r'<serie>(\d+)</', content_str).group(1) if re.search(r'<serie>(\d+)</', content_str) else "0"
            ini = re.search(r'<nnfini>(\d+)</', content_str).group(1) if re.search(r'<nnfini>(\d+)</', content_str) else "0"
            fin = re.search(r'<nnffin>(\d+)</', content_str).group(1) if re.search(r'<nnffin>(\d+)</', content_str) else ini
            resumo["Número"], resumo["Range"] = int(ini), (int(ini), int(fin))
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            # 2. NORMAIS / CANCELADOS
            m_ch = RE_CHAVE.search(content_str)
            if m_ch:
                resumo["Chave"] = m_ch.group(1) or m_ch.group(2)
                resumo["Ano"], resumo["Mes"] = "20" + resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])

            tipo = "NF-e"
            if '<mod>65</mod>' in content_str: tipo = "NFC-e"
            elif '<mod>57</mod>' in content_str or '<infcte' in content_str: tipo = "CT-e"
            
            status = "NORMAIS"
            if any(x in content_str for x in ['110111', '<cstat>101</cstat>']): status = "CANCELADOS"
            resumo["Tipo"], resumo["Status"] = tipo, status

            if status == "NORMAIS":
                v_m = RE_VALOR.search(content_str)
                resumo["Valor"] = float(v_m.group(1)) if v_m else 0.0
        
        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        resumo["Pasta"] = f"{'EMITIDOS_CLIENTE' if is_p else 'RECEBIDOS_TERCEIROS'}/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}"
        return resumo, is_p
    except: return None, False

def extrair_recursivo(conteudo_bytes, nome_arquivo):
    itens = []
    if nome_arquivo.lower().endswith('.zip'):
        try:
            with zipfile.ZipFile(io.BytesIO(conteudo_bytes)) as z:
                for sub_nome in z.namelist():
                    if '__MACOSX' in sub_nome or os.path.basename(sub_nome).startswith('.'): continue
                    sub_c = z.read(sub_nome)
                    if sub_nome.lower().endswith('.zip'): itens.extend(extrair_recursivo(sub_c, sub_nome))
                    elif sub_nome.lower().endswith('.xml'): itens.append((os.path.basename(sub_nome), sub_c))
        except: pass
    elif nome_arquivo.lower().endswith('.xml'):
        itens.append((os.path.basename(nome_arquivo), conteudo_bytes))
    return itens

def processar_lote_dados(lista_resumos, auth_dict=None):
    audit_map, div_list = {}, []
    canc_list, inut_list, aut_list, geral_list = [], [], [], []
    
    lote_unico = {}
    for item in lista_resumos:
        key = item["Chave"]
        if key not in lote_unico or item["Status"] in ["CANCELADOS", "INUTILIZADOS"]:
            lote_unico[key] = item

    for k, res in lote_unico.items():
        is_p = "EMITIDOS_CLIENTE" in res["Pasta"]
        status_final = res["Status"]
        
        if auth_dict and res["Chave"] in auth_dict and "CANCEL" in auth_dict[res["Chave"]]:
            if status_final == "NORMAIS":
                div_list.append({"Chave": res["Chave"], "Nota": res["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})
            status_final = "CANCELADOS"

        reg_base = {
            "Origem": "PRÓPRIA" if is_p else "TERCEIROS", "Modelo": res["Tipo"], "Série": res["Série"], 
            "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "Chave": res["Chave"], 
            "Status Final": status_final, "Valor": res["Valor"]
        }

        if status_final == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                item_in = reg_base.copy(); item_in.update({"Nota": n, "Valor": 0.0}); geral_list.append(item_in)
                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
        else:
            geral_list.append(reg_base)
            if is_p:
                sk = (res["Tipo"], res["Série"])
                audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})
                if res["Número"] > 0:
                    audit_map[sk]["nums"].add(res["Número"])
                    if status_final == "CANCELADOS": canc_list.append(reg_base)
                    elif status_final == "NORMAIS": 
                        aut_list.append(reg_base); audit_map[sk]["valor"] += res["Valor"]

    res_final, fal_final = [], []
    for (t, s), dados in audit_map.items():
        ns = sorted(list(dados["nums"]))
        if ns:
            n_min, n_max = ns[0], ns[-1]
            res_final.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil": round(dados["valor"], 2)})
            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))):
                fal_final.append({"Tipo": t, "Série": s, "Nº Faltante": b})

    return {
        'df_resumo': pd.DataFrame(res_final), 'df_faltantes': pd.DataFrame(fal_final),
        'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list),
        'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list),
        'df_divergencias': pd.DataFrame(div_list),
        'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}
    }

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

# Session State
for k in ['garimpo_ok', 'confirmado', 'relatorio', 'dict_arquivos', 'z_org', 'z_todos']:
    if k not in st.session_state:
        st.session_state[k] = [] if k == 'relatorio' else ({} if k == 'dict_arquivos' else False)

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
            with st.status("⛏️ Minerando...", expanded=True) as status_box:
                with zipfile.ZipFile(buf_org, "w") as z_org, zipfile.ZipFile(buf_todos, "w") as z_todos:
                    for i, f in enumerate(uploaded_files):
                        if i % 50 == 0: gc.collect()
                        xmls = extrair_recursivo(f.read(), f.name)
                        for name, data in xmls:
                            res, is_p = identify_xml_info(data, cnpj_limpo, name)
                            if res:
                                st.session_state['relatorio'].append(res)
                                path = f"{res['Pasta']}/{name}"
                                st.session_state['dict_arquivos'][path] = data
                                z_org.writestr(path, data)
                                z_todos.writestr(name, data)
            
            st.session_state.update(processar_lote_dados(st.session_state['relatorio']))
            st.session_state['z_org'], st.session_state['z_todos'] = buf_org.getvalue(), buf_todos.getvalue()
            st.session_state['garimpo_ok'] = True
            st.rerun()
    else:
        # Exibição de Resultados
        sc = st.session_state.get('st_counts', {})
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc.get("AUTORIZADAS", 0))
        c2.metric("❌ CANCELADAS", sc.get("CANCELADOS", 0))
        c3.metric("🚫 INUTILIZADAS", sc.get("INUTILIZADOS", 0))
        
        st.markdown("### 📊 RESUMO POR SÉRIE")
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        # Etapa 2: Validação
        with st.expander("🕵️ VALIDAR COM RELATÓRIO DE AUTENTICIDADE"):
            auth_file = st.file_uploader("Suba o Excel (.xlsx)", type=["xlsx"])
            if auth_file and st.button("🔄 ATUALIZAR STATUS"):
                df_auth = pd.read_excel(auth_file)
                a_dict = {str(r.iloc[0]).strip(): str(r.iloc[5]).upper() for _, r in df_auth.iterrows()}
                st.session_state.update(processar_lote_dados(st.session_state['relatorio'], a_dict))
                st.rerun()

        # Download Buttons
        col1, col2 = st.columns(2)
        with col1: st.download_button("📂 BAIXAR ZIP ORGANIZADO", st.session_state['z_org'], "garimpo.zip")
        with col2:
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                st.session_state['df_geral'].to_excel(writer, sheet_name='Geral', index=False)
            st.download_button("📊 BAIXAR EXCEL MASTER", buffer_excel.getvalue(), "relatorio.xlsx")

        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral para começar.")
