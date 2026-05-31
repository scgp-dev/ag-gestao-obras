# CLAUDE.md — PMO Construtora A. Gaspar S.A.

## Contexto do Projeto

Sistema de painéis de gestão de obras da **Construtora A. Gaspar S.A.** hospedado no GitHub Pages.
- **Repositório:** `https://github.com/scgp-dev/ag-gestao-obras`
- **URL pública:** `https://scgp-dev.github.io/ag-gestao-obras/`
- **Responsável:** Jailton Santos — Gerente de Engenharia/Planejamento
- **Branch principal:** `main` (deploy automático via GitHub Pages)

---

## Estrutura do Repositório

```
/
├── index.html                        ← Portal raiz A. Gaspar (mapa Leaflet + lista obras)
├── ponte_preview.html                ← Preview da ponte (asset estático)
├── .gitignore                        ← Exclui *.xlsx, dados.json, custos_template.html, PDFs
├── CLAUDE.md                         ← Este arquivo
├── DOCUMENTACAO.md                   ← Documentação técnica para o desenvolvedor
│
└── obras/
    ├── SP-10/                        ← Ponte Graúna-Gaivotas (SIURB/PMSP)
    │   ├── index.html                ← Portal SP-10 (links: Físico, Financeiro, Custos)
    │   ├── fisico.html               ← Painel de Avanço Físico
    │   ├── financeiro.html           ← Painel Financeiro (Medições/SIURB) — requer login Google
    │   ├── kmz/Segmento_1.kmz        ← Traçado da obra (commitado)
    │   └── custos/
    │       ├── index.html            ← Painel Financeiro de Custos ← ARQUIVO PRINCIPAL SP-10
    │       └── dados.json            ← Fonte de dados (não commitado — .gitignore)
    │
    ├── MG-04/                        ← Ponte sobre o Rio São Francisco (DER-MG)
    │   ├── index.html                ← Portal MG-04 (só Custos por enquanto)
    │   ├── kmz/Projeto Geometrico... ← Traçado da obra (commitado)
    │   └── custos/
    │       ├── index.html            ← Painel Financeiro de Custos MG-04
    │       └── dados.json            ← Fonte de dados (não commitado)
    │
    └── MT-02/                        ← OAE 3 — Ponte do Rio Vermelho (RUMO S.A.)
        ├── index.html
        └── MT-002_Relatorio_Gerencial.html
```

---

## Obras Cadastradas

| Código | Obra | Cliente | Status | Contrato |
|--------|------|---------|--------|----------|
| SP-10 | Ponte Graúna-Gaivotas | SIURB/PMSP | Em execução | 029/SIURB/2025 — R$ 347,5 Mi |
| MG-04 | Ponte sobre o Rio São Francisco | DER-MG | Em execução | Em elaboração |
| MT-02 | OAE 3 — Rio Vermelho | RUMO S.A. | Concluída | LRV80/2023 — R$ 146 Mi |

---

## Portal Raiz (index.html)

Usa **Leaflet.js** com mapa interativo (satélite/dark/streets). Cada obra tem:
- Marcador no mapa exatamente nas coordenadas do KMZ
- Painel lateral com info da obra e links para os painéis
- Toggle para visualizar traçado KMZ na obra selecionada

**Coordenadas das obras:**
- SP-10: Lat -23.7227, Lng -46.6729
- MG-04: Lat -14.7476, Lng -43.9339 (centro do KMZ)
- MT-02: Lat -16.4706, Lng -54.6358

Para adicionar nova obra: inserir novo objeto no array `OBRAS` em `index.html`.

---

## Painel de Custos — Arquitetura

Cada painel (`obras/XX/custos/index.html`) é **standalone** — dados JSON embutidos diretamente no HTML.

### Estrutura JavaScript obrigatória (ordem correta no arquivo):
```
1. const DADOS_RAW = [...JSON...];
2. // Faturamento...
   const FATURAMENTO = {...};  ← ou {} se sem faturamento
3. // CATEGORIAS DE APROPRIAÇÃO
   const CATS = [...];
4. const CLS = Object.fromEntries(CATS.map(...));
5. const CORES = {...};
6. function categorizar(r) {...}
7. const DADOS = DADOS_RAW.map(r => ({...r, aprop: categorizar(r)}));
8. const LABEL_MES = {...};
9. let filtros, pgAtual, filtrados, sortCol, sortDir, ...
10. window.addEventListener('DOMContentLoaded', () => { aplicar(); });
11. Funções: aplicar, limpar, renderTudo, renderSB, renderKPIs, renderCurva,
             renderMeses, renderDiagnostico, renderCategorias, renderFornecedores,
             renderDetalhado, sortBy, irPg, pg, setFonte, gerarPDF
```

> ⚠️ **Armadilha crítica ao adaptar painéis:** ao copiar SP-10 para nova obra,
> o bloco entre `DADOS_RAW` e `CLS` (que contém FATURAMENTO + CATS + CORES)
> pode ser perdido se o `IndexOf('};')` encontrar o fechamento errado.
> Sempre verificar que `const CATS = [` e `const CORES = {` existem no arquivo destino.

### Como Reconstruir o SP-10 Após Atualizar os Dados

O `custos_template.html` (template de desenvolvimento, **não commitado**, `.gitignore`)
está em `C:\Users\Jailton Santos\ponte-grauna-painel\obras\SP-10\custos\`.

```powershell
$tmpl  = Get-Content "custos_template.html" -Raw -Encoding UTF8
$dados = Get-Content "dados.json" -Raw -Encoding UTF8
$out   = $tmpl.Replace('%%DADOS_JSON%%', $dados)
[System.IO.File]::WriteAllText("index.html", $out, [System.Text.UTF8Encoding]::new($false))
```

### Adaptar para nova obra (sem template)

Ao derivar `index.html` de SP-10 para uma nova obra, usar PowerShell com substituição
de string direta **sem** usar `IndexOf('};')` para encontrar fim do FATURAMENTO —
esse método é frágil. Em vez disso, usar `IndexOf('const FATURAMENTO = {')` para
localizar o bloco e `IndexOf('const CATS = [')` para verificar integridade.

---

## SP-010 — Faturamento Embutido

Objeto `FATURAMENTO` no JS (~linha 590). 9 NFs emitidas:
```js
const FATURAMENTO = {
  '06/2025': 945315.38,   // NFS 5869 — 1ª Med
  '08/2025': 1070235.50,  // NFS 5979 — 2ª Med
  '09/2025': 1469509.41,  // NFS 5999 — 3ª Med
  '10/2025': 5877844.78,  // NFS 6000 — 4ª Med
  '11/2025': 5528933.79,  // NFS 6001 — 5ª Med
  '01/2026': 16165294.08, // NFS 6040 — 6ª Med
  '02/2026': 10086985.91, // NFS 6089 — 7ª Med
  '03/2026': 10237937.56, // NFS 6126 — 8ª Med
  '04/2026': 11018518.57, // NFN 31   — 9ª Med
};
// Total: R$ 62.400.574,98
```

MG-04 não tem faturamento ainda → `const FATURAMENTO = {};`

---

## Motor de Categorização (`categorizar`)

Ordem de prioridade (qualquer alteração deve ser feita em ambos SP-10 e MG-04):

1. **Aço Estrutural:** Gerdau/ArcelorMittal + VERGALHÃO/ARAME/PREGO
2. **Fundação:** CALTUBOS; ArcelorMittal+BOBINA; Drilling; Poly Fund; MEMPS
3. **Embarcações/Flutuantes:** Gerdau/ArcelorMittal (demais); D Correa; Oceanorte
4. **Equipamentos (Compra):** Convicta IND+C45/Silo; Villa; Luiz Monteiro
5. **Locação de Veículos Leves:** UNIDAS LOCADORA
6. **Locação de Equipamentos:** ~15 fornecedores + LOCACAO DE [equip/caminhao/...]
7. **Mão de Obra:** FOLHA, INSS FOLHA, FGTS, VT, 13°, férias, rescisão, contrib. retrib., TAR TRANSF SALARIO, CONTRIB ADICIONAL SENAI
8. **Encargos Trabalhistas:** Unimed, Sul América, Humana, Pluxee, Jovem Aprendiz
9. **Impostos e Retenções:** Governo, Receita, ISS, IRRF, INSS retido
10. **Consultoria e Projetos:** Multiplano, Enescil, BFA, Pilares, etc.
11. **Combustíveis:** TICKLOG, Raizen, Vibra, ICP (DIESEL → Outros)
12. **Seguros:** AVLA, Sul América Seguros
13. **Aluguel de Imóveis:** Espedito, NEC, RTC + ALUGUEL/ALG REFERENTE/CAUCAO
14. **EPI / Uniformes:** JF Soldas, D Nissa, Vale Safe, Epiflex
15. **Transporte e Logística:** ~10 transportadoras
16. **Prestação de Contas:** Bruno Barbosa + PRESTACAO DE CONTAS no hist
17. **Administrativo:** Bancos, Telecom, TOTVS, MRV/LMA/TOTAL PTA, Passagens
18. **Outros:** fallback

---

## Funcionalidades Implementadas (ambos os painéis)

- KPI Hero Block: Custo Total, Faturamento, Resultado com ícones SVG dinâmicos
- Curva S: barras (custo mensal) + 3 linhas (custo acum., fat. acum., resultado acum.)
- Diagnóstico Mensal: tabela custo × faturado × resultado (bugs corrigidos: meses sem medição e sinal acumulado)
- Curva ABC: composição do custo com % acumulado e classificação A/B/C
- Extrato: tabela paginada com **ordenação clicável** por Data/Fornecedor/Valor/Total
- Seletor de fonte: A / A+ / A++ com persistência localStorage
- Gerador de PDF: `window.print()` com layout dedicado, captura Curva S como imagem
- Filtros globais: mês, categoria, busca livre

---

## Identidade Visual

```css
--g-esc: #1A0D4A   /* azul muito escuro — headers */
--g-pri: #28156E   /* azul A. Gaspar — nav, botões */
--g-med: #3D2490   /* azul médio — destaques */
--g-clr: #EAE7F5   /* lilás claro — cards */
--g-brd: #C5BDE8   /* lilás borda */
```
Fontes: **Barlow + Barlow Condensed** (Google Fonts)

---

## Deploy

```powershell
git add obras/SP-10/custos/index.html   # ou outro arquivo alterado
git commit -m "Mensagem descritiva"
git push origin main
# GitHub Pages deploya em ~1-2 min
```

Remote: `https://github.com/scgp-dev/ag-gestao-obras.git`

---

# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.
