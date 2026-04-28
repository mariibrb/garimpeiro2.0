#!/usr/bin/env python3
"""
Garimpeiro local — linha de comandos (sem servidor Streamlit).
Uso: py -3 garimpeiro_cli.py --entrada "D:\\lote" --saida "D:\\out" --cnpj 12345678000199 --modo pasta
     py -3 garimpeiro_cli.py --entrada "D:\\lote" --saida "D:\\out" --cnpj ... --modo sped --codigo 578
     py -3 garimpeiro_cli.py --entrada "D:\\lote" ... --extracao dominio
     py -3 garimpeiro_cli.py ... --extracao matriosca --extracao-pasta dominio
     py -3 garimpeiro_cli.py ... --extracao-pasta apenas_zip
"""
from __future__ import annotations

import argparse
import os
import sys
import types


def _install_minimal_streamlit():
    """Substitui streamlit antes de importar app.py (UI do módulo fica desativada)."""
    if "streamlit" in sys.modules:
        return

    class _SessionState(dict):
        def get(self, k, default=None):
            try:
                return self[k]
            except KeyError:
                return default

    class _Ctx:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def _noop(*_a, **_k):
        return None

    def _container():
        return _Ctx()

    def _form(*_a, **_k):
        return _Ctx()

    def _columns(*_a, **_k):
        return (_Ctx(), _Ctx())

    def _tabs(_names):
        return tuple(_Ctx() for _ in _names)

    def _cache_data_decorator(_fn=None, **_kwargs):
        def _dec(f):
            return f

        return _dec(_fn) if _fn is not None else _dec

    m = types.ModuleType("streamlit")
    m.session_state = _SessionState()
    m.set_page_config = _noop
    m.markdown = _noop
    m.caption = _noop
    m.title = _noop
    m.header = _noop
    m.subheader = _noop
    m.write = _noop
    m.info = _noop
    m.warning = _noop
    m.error = _noop
    m.success = _noop
    m.exception = _noop
    m.button = lambda *a, **k: False
    m.download_button = lambda *a, **k: False
    m.file_uploader = lambda *a, **k: None
    m.text_input = lambda *a, **k: ""
    m.text_area = lambda *a, **k: ""
    m.selectbox = lambda *a, **k: (a[1][0] if len(a) > 1 and a[1] else None)
    m.multiselect = lambda *a, **k: []
    m.checkbox = lambda *a, **k: False
    m.radio = lambda *a, **k: None
    m.number_input = lambda *a, **k: 0
    m.slider = lambda *a, **k: 0
    m.date_input = lambda *a, **k: None
    m.time_input = lambda *a, **k: None
    m.spinner = lambda *a, **k: _Ctx()
    m.status = lambda *a, **k: _Ctx()
    m.progress = lambda *a, **k: _noop
    m.empty = lambda: _noop
    m.container = _container
    m.sidebar = _Ctx()
    m.expander = lambda *a, **k: _Ctx()
    m.form = _form
    m.form_submit_button = lambda *a, **k: False
    m.columns = _columns
    m.tabs = _tabs
    m.plotly_chart = _noop
    m.dataframe = _noop
    m.table = _noop
    m.metric = _noop
    m.divider = _noop
    m.stop = lambda: (_ for _ in ()).throw(SystemExit(0))
    m.rerun = _noop
    m.cache_data = _cache_data_decorator

    comp = types.ModuleType("streamlit.components")
    comp1 = types.ModuleType("streamlit.components.v1")
    comp1.html = _noop
    comp.v1 = comp1
    m.components = comp

    sys.modules["streamlit"] = m
    sys.modules["streamlit.components"] = comp
    sys.modules["streamlit.components.v1"] = comp1


def main() -> int:
    p = argparse.ArgumentParser(description="Garimpeiro — execução local (uma passagem).")
    p.add_argument("--entrada", required=True, help="Pasta com XML/ZIP (varredura recursiva).")
    p.add_argument("--saida", required=True, help="Pasta onde gravar relatórios e pacote contabilidade.")
    p.add_argument("--cnpj", required=True, help="CNPJ do emitente (14 dígitos).")
    p.add_argument(
        "--modo",
        choices=("pasta", "sped"),
        default="pasta",
        help="pasta = lote completo; sped = filtrar export ao SPED_<codigo>.txt.",
    )
    p.add_argument(
        "--codigo",
        default="",
        help="Código do escritório para localizar SPED_<codigo>.txt (obrigatório se --modo sped).",
    )
    p.add_argument(
        "--stem",
        default="",
        help="Nome base dos ZIPs do pacote (omissão: pacote_apuracao).",
    )
    p.add_argument(
        "--extracao",
        choices=("matriosca", "dominio"),
        default="matriosca",
        help="Layout dos ficheiros .zip do pacote (matriosca=XML/Lote_… dentro do zip; dominio=XML na raiz do zip).",
    )
    p.add_argument(
        "--extracao-pasta",
        choices=("matriosca", "dominio", "apenas_zip"),
        default=None,
        help=(
            "Espelho: matriosca/dominio deve coincidir com --extracao (Recursivo/Domínio alinhados); "
            "apenas_zip = só .zip e Excel sem pastas de XML (omissão = mesmo valor que --extracao)."
        ),
    )
    p.add_argument(
        "--quiet",
        action="store_true",
        help="Sem mensagens de progresso na consola (só erros no stderr e resumo final).",
    )
    args = p.parse_args()

    if args.quiet:
        os.environ["GARIMPEIRO_CLI_QUIET"] = "1"

    os.environ["GARIMPEIRO_HEADLESS"] = "1"
    os.environ.setdefault("GARIMPEIRO_ANALISE_SEM_DISCO_LOCAL", "0")

    _install_minimal_streamlit()

    from garimpeiro_local_pipeline import run_garimpeiro_local

    r = run_garimpeiro_local(
        entrada=args.entrada,
        saida=args.saida,
        cnpj=args.cnpj,
        modo=args.modo,
        codigo_sped=args.codigo.strip() or None,
        stem_zip=args.stem.strip() or None,
        extracao=args.extracao,
        extracao_pasta=args.extracao_pasta,
    )
    if not r.get("ok"):
        print("ERRO:", r.get("erro", "desconhecido"), file=sys.stderr)
        return 1
    print("OK — saída:", r.get("saida"))
    print("Espelho:", r.get("espelho"))
    print("Documentos únicos:", r.get("n_documentos"))
    if r.get("sped_faltantes_xlsx"):
        print("SPED sem XML no lote:", r.get("sped_faltantes_xlsx"))
    for a in r.get("avisos_planilhas") or []:
        print("Aviso:", a)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
