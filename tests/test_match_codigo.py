# -*- coding: utf-8 -*-
"""Regressão do match por CÓDIGO do produto (cod_bp).
Garante: (1) código com zeros à esquerda casa (00059 == 59); (2) o código só é
aceito se o NOME corrobora (redundância); (3) WAGYU não casa por código (fica
manual); (4) código compartilhado Clássico×Reserva é desempatado pela categoria.
Roda como script: python tests/test_match_codigo.py
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app

falhas = []
def check(nome, cond, detalhe=""):
    print(("  OK  " if cond else "FALHA ") + nome + (f" — {detalhe}" if detalhe and not cond else ""))
    if not cond: falhas.append(nome)

cfg = app.load_cfg()
precos = app.precos_da_versao(cfg, "PJ", None)
produtos = cfg.get("produtos", [])

def m(desc, cod):
    # cache_key único por teste para não reaproveitar índice
    return app._match_preco(desc, precos, "PJ", cod_pdf=cod, produtos=produtos, cache_key="tcod")

# normalização de código
check("_norm_cod remove zeros à esquerda", app._norm_cod("00059") == "59", app._norm_cod("00059"))
check("_norm_cod só dígitos", app._norm_cod(" 059 ") == "59")

# 1. código com zeros à esquerda casa o mesmo produto que sem zeros
ref0, key0, _, sc0 = m("ACEM PESCOCO RESERVA", "00059")
ref1, key1, _, sc1 = m("ACEM PESCOCO RESERVA", "59")
check("código 00059 casa ACÉM PESCOÇO RESERVA", key0 and "ACÉM PESCOÇO RESERVA" in key0, str(key0))
check("00059 e 59 dão o mesmo produto", key0 == key1, f"{key0} vs {key1}")
check("match por código tem score 1.0", sc0 == 1.0, str(sc0))

# 2. redundância de nome: código de ACÉM mas descrição de outro corte -> NÃO usa o código
#    (deve cair no fuzzy e casar pelo NOME, não pelo código errado)
refX, keyX, _, scX = m("BANANINHA", "00059")
check("código ignorado quando o nome não corrobora (BANANINHA, cod de ACÉM)",
      keyX is None or "BANANINHA" in str(keyX), str(keyX))

# 3. WAGYU não casa por código (mantém na revisão manual)
refW, keyW, _, scW = m("ACEM P WAGYU CG", "00538")
check("WAGYU não casa por código (vai à revisão)", keyW is None, str(keyW))

# 4. código compartilhado: LÍNGUA (cod 367) -> Clássico para 'LINGUA BP'
refL, keyL, _, scL = m("LINGUA BP", "00367")
check("código compartilhado desempata por categoria (LINGUA BP -> CLASSICO)",
      keyL and keyL.endswith("|CLASSICO"), str(keyL))

print()
if falhas:
    print(f"RESULTADO: {len(falhas)} FALHA(S) -> {falhas}"); sys.exit(1)
print("RESULTADO: TODOS OS TESTES PASSARAM")
