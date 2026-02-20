import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import gc

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def carregar_estilo():
    try:
        with open("style.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass # Estilo embutido no arquivo CSS externo

carregar_estilo()

# --- FUNÇÕES DE IDENTIFICAÇÃO (COMPLETA) ---
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
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        if not any(x in tag_l for x in ['<?xml', '<inf', '<inut']): return None, False
        
        # tpNF
        tp_nf = re.search(r'<tpnf>([01])</tpnf>', tag_l)
        if tp_nf: resumo["Operacao"] = "ENTRADA" if tp_nf.group(1) == "0" else "SAIDA"

        # Emitente e Destinatário (Regex originais preservadas)
        resumo["CNPJ_Emit"] = re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S).group(1) if re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S) else ""
        resumo["Nome_Emit"] = re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""
        resumo["Doc_Dest"] = re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S).group(1) if re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S) else ""
        resumo["Nome_Dest"] = re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""

        # Datas
        data_m = re.search(r'<(?:dhemi|demi|dhregevento)>(\d{4}-\d{2}-\d{2})', tag_l)
        if data_m: resumo["Data_Emissao"] = data_m.group(1)

        # Inutilizadas
        if '<inutnfe' in tag_l or '<retinutnfe' in tag_l or '<procinut' in tag_l:
            resumo["Status"], resumo["Tipo"] = "INUTILIZADOS", "NF-e"
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            fin = re.search(r'<nnffin>(\d+)</', tag_l).group(1) if re.search(r'<nnffin>(\d+)</', tag_l) else ini
            resumo["Número"], resumo["Range"] = int(ini), (int(ini), int(fin))
            resumo["Ano"] = "20" + re.search(r'<ano>(\d+)</', tag_l).group(1)[-2:] if re.search(r'<ano>(\d+)</', tag_l) else "0000"
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"
        else:
            # Normais / Cancelados
            m_ch = re.search(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</', tag_l)
            if not m_ch: m_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
            resumo["Chave"] = m_ch.group(1) if m_ch else ""

            if resumo["Chave"]:
                resumo["Ano"], resumo["Mes"] = "20" + resumo["Chave"][2:4], resumo["Chave"][4:6]
                resumo["Série"] = str(int(resumo["Chave"][22:25]))
                resumo["Número"] = int(resumo["Chave"][25:34])

            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: resumo["Tipo"] = "CT-e"
            
            if any(x in tag_l for x in ['110111', '<cstat>101</cstat>']): resumo["Status"] = "CANCELADOS"
            elif '110110' in tag_l: resumo["Status"] = "CARTA_CORRECAO"

            if resumo["Status"] == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                resumo["Valor"] = float(v_match.group(1)) if v_match else 0.0

        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        resumo["Pasta"] = f"{'EMITIDOS_CLIENTE' if is_p else 'RECEBIDOS_TERCEIROS'}/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
        return resumo, is_p
    except: return None, False

# --- FUNÇÕES DE APOIO ---
def extrair_recursivo(conteudo_bytes, nome_arquivo):
    itens = []
    if nome_arquivo.lower().endswith('.zip'):
        try:
            with zipfile.ZipFile(io.BytesIO(conteudo_bytes)) as z:
                for sub_nome in z.namelist():
                    if '__MACOSX' in sub_nome or os.path.basename(sub_nome).startswith('.'): continue
                    sc = z.read(sub_nome)
                    if sub_nome.lower().endswith('.zip'): itens.extend(extrair_recursivo(sc, sub_nome))
                    elif sub_nome.lower().endswith('.xml'): itens.append((os.path.basename(sub_nome), sc))
        except: pass
    elif nome_arquivo.lower().endswith('.xml'):
        itens.append((os.path.basename(nome_arquivo), conteudo_bytes))
    return itens

def processar_logica_negocio(lista_resumos, auth_dict=None):
    """Recria exatamente os DataFrames do seu código original."""
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

        origem = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
        reg = {
            "Origem": origem, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], 
            "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], 
            "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"],
            "Chave": res["Chave"], "Status Final": status_final, "Valor": res["Valor"]
        }

        if status_final == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                item_inut = reg.copy(); item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0}); geral_list.append(item_inut)
                if is_p:
                    sk = (res["Tipo"], res["Série"])
                    audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
        else:
            geral_list.append(reg)
            if is_p:
                sk = (res["Tipo"], res["Série"])
                audit_map.setdefault(sk, {"nums": set(), "valor": 0.0})
                if res["Número"] > 0:
                    audit_map[sk]["nums"].add(res["Número"])
                    if status_final == "CANCELADOS": canc_list.append(reg)
                    elif status_final == "NORMAIS": 
                        aut_list.append(reg); audit_map[sk]["valor"] += res["Valor"]

    res_f, fal_f = [], []
    for (t, s), d in audit_map.items():
        ns = sorted(list(d["nums"]))
        if ns:
            n_min, n_max = ns[0], ns[-1]
            res_f.append({"Documento": t, "Série": s, "Início": n_min, "Fim": n_max, "Quantidade": len(ns), "Valor Contábil (R$)": round(d["valor"], 2)})
            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))):
                fal_f.append({"Tipo": t, "Série": s, "Nº Faltante": b})

    return {
        'df_resumo': pd.DataFrame(res_f), 'df_faltantes': pd.DataFrame(fal_f),
        'df_canceladas': pd.DataFrame(canc_list), 'df_inutilizadas': pd.DataFrame(inut_list),
        'df_autorizadas': pd.DataFrame(aut_list), 'df_geral': pd.DataFrame(geral_list),
        'df_divergencias': pd.DataFrame(div_list), 'st_counts': {"CANCELADOS": len(canc_list), "INUTILIZADOS": len(inut_list), "AUTORIZADAS": len(aut_list)}
    }

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

# Session State Keys
ks = ['garimpo_ok', 'confirmado', 'relatorio', 'dict_arquivos', 'z_org', 'z_todos', 'df_resumo', 'df_faltantes', 'df_canceladas', 'df_inutilizadas', 'df_autorizadas', 'df_geral', 'df_divergencias', 'st_counts']
for k in ks:
    if k not in st.session_state:
        if 'df_' in k: st.session_state[k] = pd.DataFrame()
        elif k == 'relatorio': st.session_state[k] = []
        elif k == 'dict_arquivos': st.session_state[k] = {}
        elif k == 'st_counts': st.session_state[k] = {"CANCELADOS": 0, "INUTILIZADOS": 0, "AUTORIZADAS": 0}
        else: st.session_state[k] = False

with st.sidebar:
    st.markdown("### 🔍 Configuração")
    cnpj_in = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_l = "".join(filter(str.isdigit, cnpj_in))
    if len(cnpj_l) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
    if st.button("🗑️ RESETAR SISTEMA"): st.session_state.clear(); st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        files = st.file_uploader("Suba seus arquivos:", accept_multiple_files=True)
        if files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            buf_o, buf_t = io.BytesIO(), io.BytesIO()
            with st.status("⛏️ Minerando...") as stt:
                with zipfile.ZipFile(buf_o, "w") as z_o, zipfile.ZipFile(buf_t, "w") as z_t:
                    for i, f in enumerate(files):
                        if i % 50 == 0: gc.collect()
                        all_xmls = extrair_recursivo(f.read(), f.name)
                        for n, d in all_xmls:
                            res, is_p = identify_xml_info(d, cnpj_l, n)
                            if res:
                                st.session_state['relatorio'].append(res)
                                path = f"{res['Pasta']}/{n}"
                                st.session_state['dict_arquivos'][path] = d
                                z_o.writestr(path, d); z_t.writestr(n, d)
            st.session_state.update(processar_logica_negocio(st.session_state['relatorio']))
            st.session_state['z_org'], st.session_state['z_todos'] = buf_o.getvalue(), buf_t.getvalue()
            st.session_state['garimpo_ok'] = True
            st.rerun()
    else:
        # Exibição (Fiel ao Original)
        sc = st.session_state['st_counts']
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS", sc["AUTORIZADAS"])
        c2.metric("❌ CANCELADAS", sc["CANCELADOS"])
        c3.metric("🚫 INUTILIZADAS", sc["INUTILIZADOS"])
        
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        # Etapa 2 e Downloads (Originais)
        with st.expander("🕵️ ETAPA 2: VALIDAR COM RELATÓRIO"):
            auth = st.file_uploader("Suba o Excel", type=["xlsx"])
            if auth and st.button("🔄 ATUALIZAR"):
                df_a = pd.read_excel(auth)
                a_d = {str(r.iloc[0]).strip(): str(r.iloc[5]).upper() for _, r in df_a.iterrows()}
                st.session_state.update(processar_logica_negocio(st.session_state['relatorio'], a_d))
                st.rerun()

        # Excel Master completo com todas as abas
        buf_ex = io.BytesIO()
        with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as wr:
            st.session_state['df_resumo'].to_excel(wr, sheet_name='Resumo', index=False)
            st.session_state['df_geral'].to_excel(wr, sheet_name='Geral', index=False)
            st.session_state['df_faltantes'].to_excel(wr, sheet_name='Buracos', index=False)
            st.session_state['df_canceladas'].to_excel(wr, sheet_name='Canceladas', index=False)

        col1, col2, col3 = st.columns(3)
        with col1: st.download_button("📂 ZIP ORGANIZADO", st.session_state['z_org'], "organizado.zip")
        with col2: st.download_button("📦 SÓ XMLs", st.session_state['z_todos'], "xmls.zip")
        with col3: st.download_button("📊 EXCEL MASTER", buf_ex.getvalue(), "relatorio.xlsx")

        if st.button("⛏️ NOVO GARIMPO"): st.session_state.clear(); st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral.")
