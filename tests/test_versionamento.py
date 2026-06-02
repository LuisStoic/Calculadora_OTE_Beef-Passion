# -*- coding: utf-8 -*-
"""Testes de regressão — versionamento de tabelas (Fase 1).
Roda como script: python tests/test_versionamento.py
Cobre o DoD do PLANO_VERSIONAMENTO.md.
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app

falhas = []
def check(nome, cond, detalhe=""):
    print(("  OK  " if cond else "FALHA ") + nome + (f" — {detalhe}" if detalhe and not cond else ""))
    if not cond:
        falhas.append(nome)

cfg = app.load_cfg()

# 1. Migração preserva chaves e cria versionamento
check("166 chaves no topo (pj/pf)",
      len(cfg["precos_pj"]) == 166 and len(cfg["precos_pf"]) == 166,
      f'pj={len(cfg["precos_pj"])} pf={len(cfg["precos_pf"])}')
check("tabela_atual = 22026",
      cfg.get("tabela_atual", {}).get("id") == "22026")
check("arquivo tem 12026 com 166 chaves",
      cfg["tabelas"] and cfg["tabelas"][0]["id"] == "12026"
      and len(cfg["tabelas"][0]["precos_pj"]) == 166
      and len(cfg["tabelas"][0]["precos_pf"]) == 166)

# 2. precos_da_versao: default = atual; arquivada = 12026
ACEM = "ACÉM PESCOÇO CG|CLASSICO"
pj_atual = app.precos_da_versao(cfg, "PJ", None)
pj_12    = app.precos_da_versao(cfg, "PJ", "12026")
pf_atual = app.precos_da_versao(cfg, "PF", None)
check("default PJ = atual (ACÉM 39.0)", pj_atual[ACEM] == 39.0, f'{pj_atual[ACEM]}')
check("default PF = atual (ACÉM 45.5)", pf_atual[ACEM] == 45.5, f'{pf_atual[ACEM]}')
check("12026 PJ = arquivada (ACÉM 35.0)", pj_12[ACEM] == 35.0, f'{pj_12[ACEM]}')
check("id inexistente cai na atual (fallback)",
      app.precos_da_versao(cfg, "PJ", "99999")[ACEM] == 39.0)

# 3. Diff-zero: helper default == comportamento antigo (topo direto)
ITENS = [
    {"desc": "TOMAHAWK", "cod": "", "peso": 1.0, "preco": 250.0, "total": 250.0},
    {"desc": "PICANHA",  "cod": "", "peso": 2.0, "preco": 300.0, "total": 600.0},
]
ote_row = next(r for r in cfg["ote"]["PJ"] if r["n"] == 3)
produtos = cfg["produtos"]
res_helper = app.calcular_ote(ITENS, ote_row, app.precos_da_versao(cfg, "PJ", None),
                              "PJ", produtos, "PJ:atual")
res_legado = app.calcular_ote(ITENS, ote_row, cfg["precos_pj"], "PJ", produtos, "PJ:legado")
iguais = all(abs(res_helper[k] - res_legado[k]) < 1e-9
             for k in ("fat_comp", "fat_total", "ating", "var_final", "rem_total"))
check("diff-zero: default == caminho antigo", iguais,
      f'helper rem={res_helper["rem_total"]} legado rem={res_legado["rem_total"]}')

# 4. Versão muda o preço de referência e a faixa (TOMAHAWK CG: 250 atual vs 210 em 12026)
item = {"desc": "TOMAHAWK", "cod": "", "peso": 1.0, "preco": 250.0, "total": 250.0}
c_atual = app.classificar_item(item, pj_atual, "PJ", produtos, "PJ:atual")
c_12    = app.classificar_item(item, pj_12,    "PJ", produtos, "PJ:12026")
check("TOMAHAWK ref atual = 250", c_atual["preco_ref"] == 250.0, f'{c_atual["preco_ref"]}')
check("TOMAHAWK ref 12026 = 210", c_12["preco_ref"] == 210.0, f'{c_12["preco_ref"]}')
check("faixa muda entre versões (ideal vs acima)",
      c_atual["faixa"] == "ideal" and c_12["faixa"] == "acima",
      f'atual={c_atual["faixa"]} 12026={c_12["faixa"]}')

# 5. listar_tabelas: atual em primeiro, sem preços
metas = app.listar_tabelas(cfg)
check("listar_tabelas: atual primeiro e marcada",
      metas[0]["id"] == "22026" and metas[0]["atual"] is True)
check("listar_tabelas: inclui 12026",
      any(m["id"] == "12026" for m in metas))
check("listar_tabelas: sem vazar preços",
      all("precos_pj" not in m for m in metas))

print()
if falhas:
    print(f"RESULTADO: {len(falhas)} FALHA(S) -> {falhas}")
    sys.exit(1)
print("RESULTADO: TODOS OS TESTES PASSARAM")
