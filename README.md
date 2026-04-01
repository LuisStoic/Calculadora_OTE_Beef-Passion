# BeefPassion · Calculadora OTE
**Stoic Capital · v1.0**

---

## Como usar (Windows)

### Primeira vez
1. Extraia esta pasta em qualquer local (ex: `C:\BeefPassion_OTE\`)
2. Dê **duplo clique** no arquivo `INICIAR.bat`
3. Uma janela de terminal vai abrir, instalar as dependências e iniciar o servidor
4. O navegador abrirá automaticamente em `http://localhost:5000`

### Uso diário
- Dê duplo clique em `INICIAR.bat`
- Para encerrar: feche a janela do terminal ou pressione `Ctrl+C`

---

## Arquivos da pasta

| Arquivo | Descrição |
|---|---|
| `INICIAR.bat` | Script de inicialização (duplo clique para abrir) |
| `app.py` | Servidor Python (Flask + pdfplumber) |
| `requirements.txt` | Dependências Python |
| `config.json` | Preços e parâmetros OTE salvos (criado automaticamente) |
| `templates/index.html` | Interface web |

---

## Fluxo de uso

1. **Vendedor & Período** — selecione o vendedor, nível OTE e mês/ano
2. **Relatório de Vendas** — faça upload do PDF de faturamento exportado do ERP
3. **Resultado** — visualize o extrato, exporte em Excel ou imprima
4. **Configurações** — preencha a tabela de preços de referência (fazer uma vez só, fica salvo)

---

## Configurações importantes

### Tabela de Preços (fazer antes do primeiro uso)
- Acesse **Configurações → Tabela de Preços**
- Preencha o **Preço de Tabela (R$/kg)** de cada produto com os preços oficiais vigentes
- Clique em **Salvar Configurações**
- Os preços ficam salvos em `config.json` e persistem entre sessões

### Classificação por faixa
| Faixa | Critério |
|---|---|
| Acima da Tabela | Preço praticado > preço referência |
| Tabela Ideal (≤3%) | Desvio entre 0% e -3% |
| 3 a 10% Desc. | Desvio entre -3% e -10% |
| Abaixo 10% | Desvio < -10% |
| Sem Referência | Produto sem preço cadastrado — excluído do cálculo |

---

## Fontes dos dados

- **Modelo OTE:** `_Beef__Metas_OTE.pdf · Stoic Capital · Luís Fernando · 26/12/2025 · Base G4 Educação`
- **Tabela de Multiplicadores:** `Beef_Metas_OTE.xlsx · aba "1. Tabela de Multiplicador"` — 252 entradas, passo 0,01
- **Extração PDF:** pdfplumber com leitura posicional (coordenadas X/Y) — compatível com FastReport

---

## Solução de problemas

**A janela fecha imediatamente ao clicar em INICIAR.bat**
→ Python não está no PATH. Abra o terminal (`cmd`) e execute:
```
pip install flask pdfplumber openpyxl
cd C:\BeefPassion_OTE
python app.py
```

**"Porta 5000 já está em uso"**
→ Alguma outra aplicação está usando a porta. No `app.py`, linha final, altere `port=5000` para `port=5001` e acesse `http://localhost:5001`.

**PDF não extrai itens**
→ O formato do relatório pode ter variado. Verifique se o PDF foi exportado diretamente do ERP (FastReport), não escaneado ou convertido.
