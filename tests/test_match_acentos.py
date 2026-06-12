# -*- coding: utf-8 -*-
"""Regressão: o matcher de cortes deve ignorar acentos. O PDF do ERP vem sem
acento (ACEM PESCOCO) e a tabela tem acento (ACÉM PESCOÇO); sem normalizar,
cortes válidos caem em 'Sem Referência' e ficam fora do cálculo de comissão.
Roda como script: python tests/test_match_acentos.py
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

# 1. _tokens normaliza acentos
check("_tokens remove acentos", app._tokens("ACÉM PESCOÇO") == app._tokens("ACEM PESCOCO"),
      f'{app._tokens("ACÉM PESCOÇO")} vs {app._tokens("ACEM PESCOCO")}')

cfg = app.load_cfg(); precos = cfg["precos_pj"]; produtos = cfg["produtos"]

# 2. Cortes do PDF (sem acento) casam com a chave acentuada da tabela (score alto)
CASOS = [
    ("ACEM PESCOCO RESERVA", "ACÉM PESCOÇO RESERVA|RESERVA"),
    ("COXAO DURO PECA CG",   "COXÃO DURO PEÇA CG|CLASSICO"),
    ("FILE MIGNON PECA CG",  "FILÉ MIGNON PEÇA CG|CLASSICO"),
    ("ACEM PESCOCO CG",      "ACÉM PESCOÇO CG|CLASSICO"),
]
for desc, esperado in CASOS:
    app._MATCH_INDEX = {}
    p, ch, cat, sc = app._match_preco(desc, precos, "PJ", "", produtos, cache_key="t")
    check(f"match acentuado: {desc}", ch == esperado and sc >= 0.95, f"-> {ch} (score {round(sc,2)})")

# 3. Não-regressão: corte sem acento que já casava continua casando
app._MATCH_INDEX = {}
p, ch, cat, sc = app._match_preco("PICANHA CG", precos, "PJ", "", produtos, cache_key="t")
check("não-regressão: PICANHA CG casa", ch == "PICANHA CG|CLASSICO", f"-> {ch}")

print()
if falhas:
    print(f"RESULTADO: {len(falhas)} FALHA(S) -> {falhas}")
    sys.exit(1)
print("RESULTADO: TODOS OS TESTES PASSARAM")
