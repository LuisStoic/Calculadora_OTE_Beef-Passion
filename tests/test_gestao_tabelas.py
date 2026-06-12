# -*- coding: utf-8 -*-
"""Testes de regressão — Fase 2: gestão de tabelas em Configurações.
Roda como script: python tests/test_gestao_tabelas.py
Cobre: senha server-side, criar/editar/promover/excluir versões, audit_log,
backup pré-escrita e não-vazamento da senha. Usa um config.json isolado.
"""
import sys, os, io, json, shutil, tempfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pathlib import Path
import app

falhas = []
def check(nome, cond, detalhe=""):
    print(("  OK  " if cond else "FALHA ") + nome + (f" — {detalhe}" if detalhe and not cond else ""))
    if not cond:
        falhas.append(nome)

# Isola o config para não tocar o de produção
_tmp = Path(tempfile.mkdtemp())
shutil.copy2(app.CFG_PATH, _tmp / "config.json")
app.CFG_PATH = _tmp / "config.json"
app.BACKUP_DIR = _tmp / "backup"
c = app.app.test_client()
SENHA = "beefpassion"
ACEM = "ACÉM PESCOÇO CG|CLASSICO"

# 1. GET não vaza senha e expõe audit_log
g = c.get("/api/config").get_json()
check("GET não vaza senha", "senha" not in g)
check("GET expõe senha_definida", g.get("senha_definida") is True)
check("GET inclui audit_log", isinstance(g.get("audit_log"), list))

# 2. verificar_senha
check("verificar_senha errada", c.post("/api/verificar_senha", json={"senha": "x"}).get_json()["ok"] is False)
check("verificar_senha certa", c.post("/api/verificar_senha", json={"senha": SENHA}).get_json()["ok"] is True)

# 3. Escrita sem senha → 401; De-Para sem senha → 200
check("criar sem senha → 401",
      c.post("/api/tabelas/criar", json={"rotulo": "X", "base": "22026"}).status_code == 401)
check("precos no /api/config sem senha → 401",
      c.post("/api/config", json={"precos_pj": {}}).status_code == 401)
check("depara sem senha → 200",
      c.post("/api/config", json={"depara": {"FOO": {"key": "BAR"}}}).status_code == 200)

# 3b. Datas obrigatórias
check("criar sem datas → 400",
      c.post("/api/tabelas/criar", json={"senha": SENHA, "rotulo": "X"}).status_code == 400)

# 4. Criar versão duplicando 22026, sem tornar atual
r = c.post("/api/tabelas/criar", json={"senha": SENHA, "operador": "LF",
           "rotulo": "Jun/Jul 2026", "inicio": "2026-06-01", "fim": "2026-07-31",
           "base": "22026", "tornar_atual": False})
check("criar duplicando → 200", r.status_code == 200)
nid = r.get_json().get("id")
check("ID segue padrão <seq><ano> (3 de 2026 = 32026)", nid == "32026", str(nid))
cfg = app.load_cfg()
check("atual continua 22026", cfg["tabela_atual"]["id"] == "22026")
nova = next((t for t in cfg["tabelas"] if t["id"] == nid), None)
check("nova duplicou 166 chaves",
      nova and len(nova["precos_pj"]) == 166 and len(nova["precos_pf"]) == 166)

# 5. Editar preços da nova (arquivada) → só ela muda
pj = dict(nova["precos_pj"]); pj[ACEM] = 999.0
r = c.post(f"/api/tabelas/{nid}/precos",
           json={"senha": SENHA, "operador": "LF", "precos_pj": pj, "precos_pf": nova["precos_pf"]})
check("editar preços arquivada → 200", r.status_code == 200)
cfg = app.load_cfg()
check("preço mudou só na arquivada",
      app.precos_da_versao(cfg, "PJ", nid)[ACEM] == 999.0 and
      app.precos_da_versao(cfg, "PJ", None)[ACEM] == 39.0)

# 6. Promover a nova a vigente → swap correto + preço refletido no topo
r = c.post(f"/api/tabelas/{nid}/atual", json={"senha": SENHA, "operador": "LF"})
check("definir atual → 200", r.status_code == 200)
cfg = app.load_cfg()
check("nova virou vigente", cfg["tabela_atual"]["id"] == nid)
check("22026 foi arquivada", any(t["id"] == "22026" for t in cfg["tabelas"]))
check("12026 preservada", any(t["id"] == "12026" for t in cfg["tabelas"]))
check("preço editado refletido no topo", app.precos_da_versao(cfg, "PJ", None)[ACEM] == 999.0)

# 7. Excluir: vigente bloqueada (400); arquivada ok
check("excluir vigente → 400",
      c.post(f"/api/tabelas/{nid}/excluir", json={"senha": SENHA}).status_code == 400)
check("excluir arquivada 22026 → 200",
      c.post("/api/tabelas/22026/excluir", json={"senha": SENHA, "operador": "LF"}).status_code == 200)
cfg = app.load_cfg()
check("22026 removida", not any(t["id"] == "22026" for t in cfg["tabelas"]))

# 8. audit_log registrou cada ação com operador
acoes = [e["acao"] for e in cfg.get("audit_log", [])]
check("log tem criar/editar/definir/excluir",
      all(a in acoes for a in ("criar_tabela", "editar_precos", "definir_atual", "excluir_tabela")),
      str(acoes))
check("log tem operador", all(e.get("operador") for e in cfg.get("audit_log", [])))

# 9. GET /api/tabelas/<id> retorna preços da versão
d = c.get(f"/api/tabelas/{nid}").get_json()
check("GET tabela específica traz preços", d.get("precos_pj", {}).get(ACEM) == 999.0)
check("GET tabela inexistente → 404", c.get("/api/tabelas/99999").status_code == 404)

# 10. ID monotônico: após excluir, o próximo NÃO reusa o id liberado
r = c.post("/api/tabelas/criar", json={"senha": SENHA, "operador": "LF",
           "rotulo": "Ago/Set 2026", "inicio": "2026-08-01", "fim": "2026-09-30",
           "base": "vazia", "tornar_atual": False})
novo2 = r.get_json().get("id")
check("ID monotônico não reusa 22026/32026 após exclusão", novo2 == "42026", str(novo2))

# 11. Alterar código: dedicado, com senha e log próprio
prod_key = app.load_cfg()["produtos"][0]["key"]
check("alterar código sem senha → 401",
      c.post("/api/produtos/codigo", json={"key": prod_key, "codigo": "55555"}).status_code == 401)
check("alterar código inválido → 400",
      c.post("/api/produtos/codigo", json={"senha": SENHA, "key": prod_key, "codigo": "12"}).status_code == 400)
r = c.post("/api/produtos/codigo", json={"senha": SENHA, "operador": "LF", "key": prod_key, "codigo": "55555"})
check("alterar código válido → 200", r.status_code == 200)
cfg = app.load_cfg()
check("código persistido",
      next(p for p in cfg["produtos"] if p["key"] == prod_key)["cod_bp"] == "55555")
check("log dedicado alterar_codigo",
      any(e["acao"] == "alterar_codigo" for e in cfg.get("audit_log", [])))

# 12. Backups pré-escrita criados
n_bk = len(list((app.BACKUP_DIR).glob("config_*.json"))) if app.BACKUP_DIR.exists() else 0
check("snapshots de backup criados", n_bk >= 4, f"n={n_bk}")

shutil.rmtree(_tmp, ignore_errors=True)
print()
if falhas:
    print(f"RESULTADO: {len(falhas)} FALHA(S) -> {falhas}")
    sys.exit(1)
print("RESULTADO: TODOS OS TESTES PASSARAM")
