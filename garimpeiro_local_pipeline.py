"""
Execução local headless do Garimpeiro (uma passagem, sem «Incluir mais»).
Chamar apenas depois de `import app` com `streamlit` substituído por um módulo mínimo (ver garimpeiro_cli.py).
"""
from __future__ import annotations

import gc
import os
import shutil
import time
from pathlib import Path

import pandas as pd
import streamlit as st


def _cli_quiet() -> bool:
    return (os.environ.get("GARIMPEIRO_CLI_QUIET") or "").strip().lower() in (
        "1",
        "true",
        "yes",
    )


def _cli_log(msg: str) -> None:
    """Saída para consola/PowerShell com flush imediato (use `python -u` ou py -3 -u)."""
    if _cli_quiet():
        return
    print(msg, flush=True)


def _cli_xml_progress_cada() -> int:
    """Log de progresso a cada N XML lidos (chaves únicas podem subir devagar por duplicados)."""
    raw = (os.environ.get("GARIMPEIRO_CLI_PROGRESS_XML_CADA") or "8000").strip()
    try:
        n = int(raw)
    except ValueError:
        n = 8000
    return max(1000, min(n, 250000))


def _file_shim(path: Path):
    class _F:
        name = str(path.name)

        def read(self):
            return Path(path).read_bytes()

        def seek(self, _n):
            pass

    return _F()


def _find_sped_txt(entrada: Path, codigo: str) -> Path | None:
    nome = f"SPED_{codigo}.txt"
    for p in entrada.rglob(nome):
        if p.is_file():
            return p
    return None


def _iter_xml_zip(entrada: Path, skip: set[Path]) -> list[Path]:
    out: list[Path] = []
    for p in entrada.rglob("*"):
        if not p.is_file():
            continue
        try:
            rp = p.resolve()
        except OSError:
            continue
        if rp in skip:
            continue
        suf = p.suffix.lower()
        if suf in (".xml", ".zip"):
            out.append(p)
    return sorted(out, key=lambda x: str(x).lower())


def run_garimpeiro_local(
    *,
    entrada: str,
    saida: str,
    cnpj: str,
    modo: str,
    codigo_sped: str | None,
    stem_zip: str | None,
    extracao: str | None = None,
    extracao_pasta: str | None = None,
) -> dict:
    import app as ap

    cnpj_limpo = "".join(c for c in str(cnpj) if c.isdigit())[:14]
    if len(cnpj_limpo) != 14:
        return {"ok": False, "erro": "CNPJ deve ter 14 dígitos."}

    entrada_p = Path(entrada).expanduser().resolve()
    saida_p = Path(saida).expanduser().resolve()
    if not entrada_p.is_dir():
        return {"ok": False, "erro": f"Pasta de entrada inexistente: {entrada_p}"}

    saida_p.mkdir(parents=True, exist_ok=True)
    modo_n = (modo or "pasta").strip().lower()
    if modo_n not in ("pasta", "sped"):
        return {"ok": False, "erro": "modo deve ser 'pasta' ou 'sped'."}
    extr_zip = (extracao or "matriosca").strip().lower()
    extr_pasta = (extracao_pasta or extracao or "matriosca").strip().lower()
    if extr_zip not in ("matriosca", "dominio"):
        return {"ok": False, "erro": "extracao (ZIP) deve ser 'matriosca' ou 'dominio'."}
    if extr_pasta not in ("matriosca", "dominio", "apenas_zip"):
        return {
            "ok": False,
            "erro": "extracao_pasta deve ser 'matriosca', 'dominio' ou 'apenas_zip' (sem espelho XML em pastas).",
        }
    if extr_pasta != "apenas_zip" and extr_pasta != extr_zip:
        extr_pasta = extr_zip

    skip_paths: set[Path] = set()
    texto_sped = ""
    if modo_n == "sped":
        cod = (codigo_sped or "").strip()
        if not cod:
            return {"ok": False, "erro": "modo sped requer --codigo (ex.: 578 → SPED_578.txt)."}
        sped_p = _find_sped_txt(entrada_p, cod)
        if not sped_p:
            return {
                "ok": False,
                "erro": f"Ficheiro SPED_{cod}.txt não encontrado sob {entrada_p} (recursivo).",
            }
        texto_sped = ap._decode_sped_upload_bytes(sped_p.read_bytes()).strip()
        skip_paths.add(sped_p.resolve())
        st.session_state[ap.SPED_SESSION_TEXT_KEY] = texto_sped
        st.session_state[ap.SPED_SESSION_NAME_KEY] = sped_p.name
    else:
        st.session_state.pop(ap.SPED_SESSION_TEXT_KEY, None)
        st.session_state.pop(ap.SPED_SESSION_NAME_KEY, None)

    up_inut = None
    for alt in ("inutilizadas.xlsx", "inutilizadas.xls", "inutilizadas.csv"):
        for p in entrada_p.rglob(alt):
            if p.is_file():
                up_inut = _file_shim(p)
                skip_paths.add(p.resolve())
                break
        if up_inut:
            break

    up_canc = None
    for alt in ("canceladas.xlsx", "canceladas.xls", "canceladas.csv"):
        for p in entrada_p.rglob(alt):
            if p.is_file():
                up_canc = _file_shim(p)
                skip_paths.add(p.resolve())
                break
        if up_canc:
            break

    os.makedirs(ap.TEMP_UPLOADS_DIR, exist_ok=True)
    try:
        for fn in os.listdir(ap.TEMP_UPLOADS_DIR):
            fp = os.path.join(ap.TEMP_UPLOADS_DIR, fn)
            if os.path.isfile(fp):
                try:
                    os.unlink(fp)
                except OSError:
                    pass
    except OSError:
        pass

    st.session_state.pop(ap.SESSION_KEY_FONTES_XML_MEMORIA, None)
    st.session_state.pop(ap.SESSION_KEY_EXTRA_DIGESTS, None)
    st.session_state.pop("garimpo_espelho_indice_td", None)

    arquivos = _iter_xml_zip(entrada_p, skip_paths)
    if not arquivos:
        return {"ok": False, "erro": "Nenhum .xml ou .zip na pasta de entrada (recursivo)."}
    _t_ini = time.perf_counter()
    _cli_log("")
    _cli_log("=== Garimpeiro LOCAL (CLI) ===")
    _cli_log(f"Entrada:  {entrada_p}")
    _cli_log(f"Saída:    {saida_p}")
    _cli_log(
        f"Modo:     {modo_n}"
        + (f"  (código SPED: {codigo_sped})" if modo_n == "sped" else "")
        + f"  |  extração ZIP: {extr_zip}  |  extração pastas: {extr_pasta}"
    )
    _cli_log(f"Ficheiros .xml/.zip encontrados na entrada: {len(arquivos)}  →  a copiar para área temporária…")

    for i, src in enumerate(arquivos, start=1):
        key = ap._garimpo_nome_chave_upload(i, src.name)
        shutil.copy2(src, os.path.join(ap.TEMP_UPLOADS_DIR, key))

    lista_salvos = ap._lista_nomes_fontes_xml_garimpo()
    if not lista_salvos:
        return {"ok": False, "erro": "Lote interno vazio após preparar TEMP."}

    total_ls = len(lista_salvos)
    _cli_log(f"Cópia concluída. Fila de leitura: {total_ls} ficheiro(s).")
    _cli_log("")

    _subdir_esp = ap.garimpe_subdir_espelho_nome(com_sped=(modo_n == "sped"))
    espelho = saida_p / _subdir_esp
    try:
        shutil.rmtree(espelho, ignore_errors=True)
        espelho.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return {"ok": False, "erro": f"Não foi possível preparar a pasta espelho: {e}"}
    st.session_state["garimpo_lote_espelho_root"] = str(espelho)
    st.session_state["garimpo_lote_save_resolved"] = str(saida_p)
    u_stem = ap._v2_sanitize_nome_export((stem_zip or "").strip(), max_len=80)
    st.session_state["mariana_zip_basename"] = u_stem if u_stem else "pacote_apuracao"
    st.session_state["cnpj_widget"] = ap.format_cnpj_visual(cnpj_limpo)
    st.session_state.pop("garimpo_espelho_indice_td", None)
    st.session_state[ap.SESSION_KEY_GARIMPO_EXTRACAO_ZIP] = extr_zip
    st.session_state[ap.SESSION_KEY_GARIMPO_EXTRACAO_PASTA] = extr_pasta
    st.session_state[ap.SESSION_KEY_GARIMPO_EXTRACAO_LOTE] = extr_zip

    lote_dict: dict = {}
    _falhas_leitura = 0
    _prog_xml_iv = _cli_xml_progress_cada()
    if ap._garimpo_escrita_espelho_final_continua_ativa():
        _cli_log(
            "Espelho em disco: gravação **apenas** após ler todo o lote e montar as tabelas (nada durante a leitura)."
        )
    for idx_f, f_name in enumerate(lista_salvos, start=1):
        _cli_log(f"[{idx_f}/{total_ls}] A ler: {f_name}")
        try:
            _xml_lidos_ficheiro = 0
            with ap._abrir_fonte_xml_garimpo_stream(f_name) as file_obj:
                for name, xml_data in ap.extrair_fonte_xml_garimpo(file_obj, f_name):
                    _xml_lidos_ficheiro += 1
                    if _xml_lidos_ficheiro % _prog_xml_iv == 0:
                        _cli_log(
                            f"    … {_xml_lidos_ficheiro} XML processados neste ficheiro · "
                            f"{len(lote_dict)} chaves únicas no acumulado…"
                        )
                    res, is_p = ap.identify_xml_info(xml_data, cnpj_limpo, name)
                    if res:
                        key = res["Chave"]
                        if key in lote_dict:
                            if res["Status"] in [
                                "CANCELADOS",
                                "INUTILIZADOS",
                                "DENEGADOS",
                                "REJEITADOS",
                            ]:
                                lote_dict[key] = (res, is_p)
                        else:
                            lote_dict[key] = (res, is_p)
                    del xml_data
        except Exception as ex:
            _falhas_leitura += 1
            _cli_log(f"    AVISO: falha ao processar este ficheiro ({ex}). Continua…")
            continue
        _cli_log(
            f"    → ficheiro concluído — {len(lote_dict)} documento(s) único(s) no acumulado."
        )

    if _falhas_leitura and not _cli_quiet():
        _cli_log(f"Aviso: {_falhas_leitura} ficheiro(s) com erro (ignorados).")

    if not lote_dict:
        return {
            "ok": False,
            "erro": "0 documentos reconhecidos — confirme o CNPJ emitente e os XML.",
        }

    _cli_log("")
    _cli_log(
        f"Leitura de XML terminada em {time.perf_counter() - _t_ini:.1f}s — a montar resumo, buracos e tabelas…"
    )

    rel_list = []
    ref_ar, ref_mr, ref_map = ap.buraco_ctx_sessao()
    audit_map = {}
    canc_list = []
    inut_list = []
    aut_list = []
    deneg_list = []
    rej_list = []
    geral_list = []

    for _k, (res, is_p) in lote_dict.items():
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
            "UF Destino": res.get("UF_Dest") or "",
            "Chave": res["Chave"],
            "Status Final": res["Status"],
            "Valor": res["Valor"],
            "Ano": res["Ano"],
            "Mes": res["Mes"],
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
            ult_u = ap.ultimo_ref_lookup(ref_map, res["Tipo"], res["Série"])

            if res["Status"] == "INUTILIZADOS":
                r = res.get("Range", (res["Número"], res["Número"]))
                _man_inut = ap._inutil_sem_xml_manual(res)
                for n in range(r[0], r[1] + 1):
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
                    if ap._incluir_em_resumo_por_serie(res, is_p, cnpj_limpo):
                        if sk not in audit_map:
                            audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                        audit_map[sk]["nums"].add(n)
                        if _man_inut or ap.incluir_numero_no_conjunto_buraco(
                            res["Ano"], res["Mes"], n, ref_ar, ref_mr, ult_u
                        ):
                            audit_map[sk]["nums_buraco"].add(n)
            else:
                if res["Status"] == "DENEGADOS":
                    deneg_list.append(registro_base)
                elif res["Status"] == "REJEITADOS":
                    rej_list.append(registro_base)
                elif res["Número"] > 0:
                    if res["Status"] == "CANCELADOS":
                        canc_list.append(registro_base)
                    elif res["Status"] == "NORMAIS":
                        aut_list.append(registro_base)
                    if ap._incluir_em_resumo_por_serie(res, is_p, cnpj_limpo):
                        if sk not in audit_map:
                            audit_map[sk] = {"nums": set(), "nums_buraco": set(), "valor": 0.0}
                        audit_map[sk]["nums"].add(res["Número"])
                        if ap._cancel_sem_xml_manual(res):
                            audit_map[sk]["nums_buraco"].add(res["Número"])
                        elif ap.incluir_numero_no_conjunto_buraco(
                            res["Ano"],
                            res["Mes"],
                            res["Número"],
                            ref_ar,
                            ref_mr,
                            ult_u,
                        ):
                            audit_map[sk]["nums_buraco"].add(res["Número"])
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
        ult_lookup = ap.ultimo_ref_lookup(ref_map, t, s) if ref_ar is not None else None
        fal_final.extend(
            ap.falhas_buraco_por_serie(
                dados["nums_buraco"], t, s, ult_lookup, nums_existentes=dados["nums"]
            )
        )

    st.session_state["relatorio"] = rel_list

    df_faltantes = pd.DataFrame(fal_final)
    _pl_ini = ap._garimpo_aplicar_planilhas_inutil_cancel_no_relatorio(
        rel_list,
        cnpj_limpo,
        df_faltantes,
        up_inut,
        up_canc,
    )
    if _pl_ini.get("msgs"):
        st.session_state["_garimpo_cli_avisos_planilhas"] = _pl_ini["msgs"]
    else:
        st.session_state.pop("_garimpo_cli_avisos_planilhas", None)

    if (_pl_ini.get("inut", 0) + _pl_ini.get("canc", 0)) > 0:
        ap.reconstruir_dataframes_relatorio_simples()
    else:
        st.session_state.update(
            {
                "df_resumo": pd.DataFrame(res_final),
                "df_faltantes": df_faltantes,
                "df_canceladas": pd.DataFrame(canc_list),
                "df_inutilizadas": pd.DataFrame(inut_list),
                "df_autorizadas": pd.DataFrame(aut_list),
                "df_denegadas": pd.DataFrame(deneg_list),
                "df_rejeitadas": pd.DataFrame(rej_list),
                "df_geral": pd.DataFrame(geral_list),
                "st_counts": {
                    "CANCELADOS": len(canc_list),
                    "INUTILIZADOS": len(inut_list),
                    "AUTORIZADAS": len(aut_list),
                    "DENEGADOS": len(deneg_list),
                    "REJEITADOS": len(rej_list),
                },
            }
        )
        ap.aplicar_compactacao_dfs_sessao()
        ap._garimpo_registar_aviso_sped_chaves_sem_xml_no_lote(
            st.session_state.get("df_geral"),
            str(st.session_state.get(ap.SPED_SESSION_TEXT_KEY) or "").strip(),
        )

    df_geral = st.session_state.get("df_geral")
    if df_geral is None or getattr(df_geral, "empty", True):
        return {"ok": False, "erro": "df_geral vazio após montagem."}

    _df_f, filtro = ap._garimpo_df_e_filtro_espelho_so_sped_se_anexado(
        df_geral,
        str(st.session_state.get(ap.SPED_SESSION_TEXT_KEY) or "").strip(),
    )
    if not filtro:
        return {
            "ok": False,
            "erro": "Nenhuma chave para exportar o pacote contabilidade (com SPED: cruza vazio?).",
        }

    _cli_log(
        f"A gravar pacote contabilidade em {espelho.name}/ dentro da saída ("
        + (
            "ZIP + Excel — sem pastas de XML no disco"
            if extr_pasta == "apenas_zip"
            else "pastas + ZIP + Excel"
        )
        + ")…"
    )
    ap._garimpo_gravar_espelho_layout_contabilidade(cnpj_limpo)

    try:
        xb = ap.excel_relatorio_geral_com_dashboard_bytes(df_geral)
        if xb:
            (saida_p / "relatorio_garimpeiro_completo.xlsx").write_bytes(xb)
    except Exception:
        pass

    for nome, key in (("canceladas.xlsx", "df_canceladas"), ("inutilizadas.xlsx", "df_inutilizadas")):
        dfx = st.session_state.get(key)
        if isinstance(dfx, pd.DataFrame) and not dfx.empty:
            b = ap.dataframe_para_excel_bytes(dfx, nome.replace(".xlsx", ""))
            if b:
                try:
                    (saida_p / nome).write_bytes(b)
                except OSError:
                    pass

    gc.collect()
    avisos = list(st.session_state.get("_garimpo_cli_avisos_planilhas") or [])
    sped_xlsx = st.session_state.get(ap.SPED_FALTANTES_XLSX_PATH_KEY)
    _cli_log(f"Concluído em {time.perf_counter() - _t_ini:.1f}s no total.")
    _cli_log("")
    return {
        "ok": True,
        "saida": str(saida_p),
        "espelho": str(espelho),
        "n_documentos": len(lote_dict),
        "avisos_planilhas": avisos,
        "sped_faltantes_xlsx": str(sped_xlsx) if sped_xlsx else None,
    }
