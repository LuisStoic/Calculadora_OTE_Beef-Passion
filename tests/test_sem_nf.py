# -*- coding: utf-8 -*-
"""Regra "sem nota fiscal": lançamentos do analítico sem NF não entram no
consolidado do sistema (que é por Cupom/NF-e), então por padrão também não
entram no cálculo do variável. Espelha a regra do intragrupo.
Roda como script: python tests/test_sem_nf.py
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

ote_row = {"n": 3, "meta": 100000, "corte": 70000, "fixo": 3000, "var": 5000}
precos = {"PICANHA|CLASSICO": 100.0}

itens = [
    {"cod": "00111", "desc": "PICANHA", "peso": 10, "preco": 100, "total": 1000,
     "nf": "026725", "cliente": "FFC LANCHES"},                     # com NF -> computa
    {"cod": "00462", "desc": "MOIDA", "peso": 5, "preco": 150, "total": 750,
     "nf": "", "cliente": "BEEF PASSION IND"},                       # sem NF (interno)
    {"cod": "00071", "desc": "PEIXINHO", "peso": 2, "preco": 50, "total": 100,
     "nf": "", "cliente": "CLIENTE REAL SEM NF"},                    # sem NF (nao-interno)
]

# 1. Default (sem_nf_considerar=False): itens sem NF ficam fora do computável
r = app.calcular_ote(itens, ote_row, precos, "PJ", [], "k", sem_nf_considerar=False)
faixas = {it["cod"]: it["faixa"] for it in r["classified"]}
check("com NF -> computa (ideal)", faixas["00111"] == "ideal", faixas.get("00111"))
check("sem NF interno -> faixa sem_nf", faixas["00462"] == "sem_nf", faixas.get("00462"))
check("sem NF externo -> faixa sem_nf", faixas["00071"] == "sem_nf", faixas.get("00071"))
check("fat_comp so conta o item com NF", abs(r["fat_comp"] - 1000.0) < 1e-6, f'fat_comp={r["fat_comp"]}')
check("auditoria conta 2 sem_nf", r["auditoria"]["sem_nf"] == 2, str(r["auditoria"]))
check("r[sem_nf].fat soma os sem NF", abs(r["r"]["sem_nf"]["fat"] - 850.0) < 1e-6, str(r["r"]["sem_nf"]))

# 2. sem_nf_considerar=True: itens sem NF voltam ao fluxo normal (aqui sem ref -> sem_ref)
r2 = app.calcular_ote(itens, ote_row, precos, "PJ", [], "k", sem_nf_considerar=True)
faixas2 = {it["cod"]: it["faixa"] for it in r2["classified"]}
check("considerar=True: sem NF nao vira sem_nf", "sem_nf" not in faixas2.values(), str(faixas2))

# 3. _intragrupo_considerado tem precedencia (nao vira sem_nf mesmo sem NF)
it_intra = {"cod": "00999", "desc": "PICANHA", "peso": 1, "preco": 100, "total": 100,
            "nf": "", "cliente": "BEEF PASSION IND", "_intragrupo_considerado": True}
c = app.classificar_item(it_intra, precos, "PJ", [], "k", sem_nf_considerar=False)
check("intragrupo_considerado nao vira sem_nf", c["faixa"] != "sem_nf", c["faixa"])

print()
if falhas:
    print(f"RESULTADO: {len(falhas)} FALHA(S) -> {falhas}")
    sys.exit(1)
print("RESULTADO: TODOS OS TESTES PASSARAM")
