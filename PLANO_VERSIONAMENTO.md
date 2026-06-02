# Plano — Versionamento de Tabelas de Preço (OTE Beef Passion)

> Status: **aprovado, aguardando go de implementação** (decisão 4b: implementar depois).
> Data do plano: 2026-06-02. Autor: LF + Stoic. Não implementar sem o "go".

## 1. Objetivo

Dar à Calculadora OTE um **histórico versionado de tabelas de preço**, onde:
- a **última versão é usada por default**;
- o operador pode **escolher manualmente** outra versão na tela (ex.: recalcular um mês passado com a tabela vigente naquele mês);
- a atualização nasce na **planilha do Google** (que continua mestra/backup) e é **promovida pela interface do app**, com revisão de diff antes de virar oficial.

## 2. Decisões travadas

| # | Decisão | Escolha |
|---|---|---|
| 1 | Seleção de versão | **(c)** última como default + **seletor manual** na tela |
| 2 | Onde mora o dado durável | **Git** (`config.json` versionado). ≤12 versões/ano. |
| 3 | Backfill da versão anterior | **(a)** a partir do **git** (commit pré-2026-06-02 = `5edad71`) |
| Interface | Como o usuário atualiza | **Jeito B**: planilha mestra → botão "importar/promover" no app → diff → commit |
| 4 | Execução | **(b)** implementar depois (este documento é o registro) |

Competência-aware automático (escolher a tabela pela vigência do mês calculado) **não** foi escolhido; o seletor manual cobre o caso. Fica como melhoria futura opcional (ver §9).

## 3. Modelo de dados (`config.json`)

Hoje: `precos_pj` / `precos_pf` planos (a tabela corrente). Passa a:

```json
"tabelas": [
  {
    "id": "12026",
    "rotulo": "Abr/Mai 2026",
    "inicio": "2026-04-01",
    "fim":    "2026-05-31",
    "fonte":  "git:5edad71",
    "criado_em": "2026-06-02",
    "precos_pj": { "...|CLASSICO": 0.0 },
    "precos_pf": { "...|CLASSICO": 0.0 }
  },
  {
    "id": "22026",
    "rotulo": "Mai/Jun 2026",
    "inicio": "2026-05-01",
    "fim":    "2026-06-30",
    "fonte":  "Sheets 1.6 / commit f15bda6",
    "criado_em": "2026-06-02",
    "precos_pj": { "...": 0.0 },
    "precos_pf": { "...": 0.0 }
  }
]
```

- `produtos`, `vendedores`, `ote`, `senha`, `mult_table`, `depara` permanecem no topo (não versionados nesta fase).
- "Última" = entrada de maior `fim` (ou maior `id`); **calculada**, não persistida → não exige escrita em runtime para o default.

## 4. Migração + backfill (compatibilidade total)

Estender `load_cfg()` (já contém camada de migração):
1. Se existe `precos_pj`/`precos_pf` no topo e **não** existe `tabelas` → embrulhar a tabela corrente na versão **22026** (vigência mai/jun) e remover/depreciar os planos do topo (ou mantê-los como espelho da última para compat de leitura).
2. **Backfill 22026 anterior:** ler o `config.json` do commit `5edad71` (estado que rodou de 30/04 até 02/06) e gravá-lo como versão **12026** (vigência abr/mai). Chaves já no formato correto, zero risco de naming.
3. Idempotente: rodar de novo não duplica versões (checar `id`).

## 5. Seleção de versão em runtime

Helper `tabela_selecionada(cfg, id_escolhido=None)`:
- `id_escolhido` informado (seletor da tela) → usa essa versão;
- senão → **última** (maior `fim`).

Ligar em `/api/calcular` e `/api/exportar_xlsx`: hoje pegam `cfg["precos_pj"|"precos_pf"]` direto (app.py ~1319 e ~1361). Passam a pegar de `tabela_selecionada(...)[canal]`. Frontend ganha um dropdown de versão (default = última) e envia `tabela_id` no payload.

## 6. Fluxo de atualização — Jeito B (promover da planilha)

1. Usuário atualiza preços **na planilha** (aba 1.6, nova `Tabela` id / vigência).
2. No app, botão **"Importar tabela da planilha"**:
   - lê a planilha (CSV publicado ou Sheets API; ver §8),
   - **normaliza** as chaves reaproveitando o parser de 02/06 (colapsar espaço duplo, uppercase, mapear categoria),
   - **merge** sobre as chaves existentes (nunca full-replace; preserva acessórios DIVERSOS e itens fora de vigência),
   - mostra o **diff** (alterados / novos / ausentes) e os **flags de variação grande**;
3. Operador **confirma** (gate);
4. App grava nova entrada em `tabelas[]` e **commita no GitHub via token** → auto-deploy Render.

## 7. Persistência (resolve o Render efêmero)

- Render tem filesystem efêmero: escrever em disco **não persiste**. Por isso o save do app **commita no GitHub** (PAT ou GitHub App como **secret** no Render), em vez de só `save_cfg` local.
- Efeito vale após o redeploy (~1-3 min). Aceitável para tabela mensal.
- Backup: a planilha continua mestra; o git guarda o histórico oficial.

## 8. Leitura da planilha pelo app

Opções (decidir na implementação):
- **CSV publicado** da aba 1.6 ("Publicar na web") — simples, sem credencial, porém expõe a aba publicamente.
- **Sheets API + service account** — privado e estruturado; exige criar service account e compartilhar a planilha com ela. **Preferido** para dado de comissão.

## 9. Riscos e mitigações

| Risco | Mitigação |
|---|---|
| Edição entra no cálculo de comissão com erro | Gate de diff + **validação**: bloquear erro estrutural; sinalizar |Δ| > 20% para confirmação consciente |
| Render efêmero (edição some) | App commita no git via token (§7) |
| Token de escrita vazando | Escopar ao repo único; GitHub App de preferência; rotacionar |
| Edição concorrente / conflito de push | Pull antes de push; trava simples; 1-3 usuários → risco baixo |
| Drift planilha vs app | Jeito B mantém a **planilha como mestra**; app só promove |
| Crescimento do `config.json` | ~50 KB/versão, ~600 KB/ano: trivial. Podar versões antigas para arquivo separado após N anos |
| Naming da planilha (espaço/caixa) | Normalização obrigatória no import (parser de 02/06) |

## 10. Testes de regressão + DoD numérico

**Definition of Done:**
- [ ] `load_cfg()` migra config plano → `tabelas[]` sem perder nenhuma das 166 chaves (pj e pf).
- [ ] Backfill cria a 12026 com exatamente as chaves/valores do commit `5edad71`.
- [ ] Para um mesmo PDF e a **última** versão selecionada, o resultado de `/api/calcular` é **idêntico** ao de hoje (faixas, fat_comp, var_final, rem_total) — diff numérico zero.
- [ ] Selecionar a 12026 muda os preços de referência de forma consistente (teste com ≥1 corte cujo preço diferiu entre 12026 e 22026).
- [ ] Import da planilha sobre estado conhecido reproduz o diff de 02/06 (68 PJ + 14 LOJA), 0 chave criada/removida.
- [ ] Flag de variação dispara para |Δ| > 20% (caso STEAK PASSION RESERVA LOJA 150→99 = -34%).
- [ ] Commit via token aparece no histórico do repo e o Render sobe a versão nova.

## 11. Fases sugeridas

- **Fase 1 (core):** modelo `tabelas[]` + migração + backfill 12026 + helper de seleção + dropdown de versão + testes de regressão. Sem escrita pela tela ainda. Entrega o versionamento e a leitura.
- **Fase 2 (promover da planilha):** leitura da planilha + normalização + diff + gate + commit via token. Entrega o self-service do Jeito B.
- **Fase 3 (opcional):** seleção automática por competência; poda de versões antigas; modo "editor livre" (Jeito A) por cima.

## 12. Premissas e pontos abertos

- Token: definir **PAT** (rápido) vs **GitHub App** (mais seguro) — recomendado App.
- Leitura da planilha: **CSV publicado** vs **service account** — recomendado service account.
- Vigências exatas (datas `inicio`/`fim`) a confirmar com o cliente; assumi mês cheio.
- `produtos`/`cod_bp` seguem globais (não versionados) nesta fase; revisitar se mudarem por vigência.
