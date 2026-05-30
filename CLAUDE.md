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
├── index.html                        ← Portal raiz A. Gaspar (lista obras)
├── ponte_preview.html                ← Preview da ponte (asset estático)
├── .gitignore                        ← Exclui *.xlsx, dados.json, custos_template.html
├── CLAUDE.md                         ← Este arquivo
├── DOCUMENTACAO.md                   ← Documentação para o desenvolvedor
│
└── obras/
    ├── SP-10/                        ← Ponte Graúna-Gaivotas (SIURB/PMSP)
    │   ├── index.html                ← Portal SP-10 (links: Físico, Financeiro, Custos)
    │   ├── fisico.html               ← Painel de Avanço Físico
    │   ├── financeiro.html           ← Painel Financeiro (Medições/SIURB)
    │   └── custos/
    │       ├── index.html            ← Painel Financeiro de Custos ← ARQUIVO PRINCIPAL
    │       └── dados.json            ← Fonte de dados (não commitado — ver .gitignore)
    │
    └── MT-02/                        ← OAE 3 — Ponte do Rio Vermelho
        ├── index.html
        └── MT-002_Relatorio_Gerencial.html
```

---

## Painel de Custos SP-010 — Arquivo Principal

**Arquivo:** `obras/SP-10/custos/index.html` (~548 KB)

O arquivo é **standalone** — os dados de 2.454 lançamentos (~479 KB de JSON) estão embutidos
diretamente no HTML como `const DADOS_RAW = [...]`. Não depende de servidor nem de API.

### Como Reconstruir Após Atualizar os Dados

1. Exportar planilha Excel → `dados.json` (PowerShell via COM Object — ver DOCUMENTACAO.md)
2. Ter o `custos_template.html` (arquivo de desenvolvimento, não commitado)
3. Executar o PowerShell de build:
   ```powershell
   $tmpl = Get-Content "custos_template.html" -Raw -Encoding UTF8
   $dados = Get-Content "dados.json" -Raw -Encoding UTF8
   $out = $tmpl.Replace('%%DADOS_JSON%%', $dados)
   [System.IO.File]::WriteAllText("index.html", $out, [System.Text.UTF8Encoding]::new($false))
   ```
4. Commitar e fazer push do `index.html`

> **Atenção:** O `custos_template.html` está no `.gitignore` e deve ser mantido localmente.
> Está em `C:\Users\Jailton Santos\ponte-grauna-painel\obras\SP-10\custos\`

---

## Seções do Painel de Custos

| Aba | ID | Descrição |
|-----|----|-----------|
| Visão Geral | `pg-visao` | KPI hero, Curva S, diagnóstico mensal |
| Apropriações | `pg-categorias` | Gráfico de barras + donut por categoria |
| Fornecedores | `pg-fornecedores` | Ranking + gráficos top 15 |
| Extrato | `pg-detalhado` | Tabela paginada, ordenável, com tfoot total |

---

## Faturamento Embutido

Objeto `FATURAMENTO` no JS do index.html (~linha 587). Valores das 9 NFs emitidas:
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
// Total acumulado faturado: R$ 62.400.574,98
```
Para atualizar: editar diretamente no `index.html` e commitar.

---

## Motor de Categorização (`categorizar`)

Função JS (~linha 648) que classifica cada lançamento com base em `nome` (fornecedor)
e `hist` (histórico). **Regras por ordem de prioridade:**

1. **Aço Estrutural:** Gerdau/ArcelorMittal + VERGALHÃO/ARAME/PREGO
2. **Fundação:** CALTUBOS (sempre); ArcelorMittal + BOBINA; Drilling; Poly Fund; MEMPS
3. **Embarcações:** Gerdau/ArcelorMittal (tudo exceto vergalhão/arame/prego); D Correa Viana; Oceanorte
4. **Equipamentos (Compra):** Convicta IND + C45/Silo/Betoneira; Villa Empreend; Luiz Monteiro
5. **Locação de Equipamentos:** ~15 fornecedores específicos + keywords LOCACAO DE [equip]
6. **Locação de Veículos Leves:** UNIDAS LOCADORA
7. **Mão de Obra:** Folha, adiantamento, INSS folha, FGTS, VT, 13°, férias, rescisão, contrib. retributiva, TAR TRANSF SALÁRIOS, CONTRIB ADICIONAL SENAI
8. **Encargos Trabalhistas:** Unimed, Sul América, Humana, Pluxee, Jovem Aprendiz
9. **Impostos e Retenções:** Governo, Receita, INSS retido, ISS, IRRF
10. **Consultoria e Projetos:** ~12 fornecedores (Multiplano, Enescil, BFA, etc.)
11. **Combustíveis:** TICKLOG, Raizen, Vibra, ICP (DIESEL no histórico → Outros)
12. **Seguros:** AVLA, Sul América Seguros
13. **Aluguel de Imóveis:** Espedito, Anasilda, NEC, RTC, Romildo + ALG/ALUGUEL/CAUCAO
14. **EPI / Uniformes:** JF Soldas, D Nissa, Vale Safe, Epiflex
15. **Transporte e Logística:** ~10 transportadoras
16. **Prestação de Contas:** Bruno Barbosa e outros (histórico PRESTACAO DE CONTAS)
17. **Administrativo:** Bancos, Telecom, TOTVS, MRV/LMA/TOTAL PTA, Passagens, etc.
18. **Outros:** Tudo que não se enquadra

---

## Identidade Visual (obrigatório em todos os painéis)

```css
--g-esc: #1A0D4A   /* azul muito escuro — headers, fundos */
--g-pri: #28156E   /* azul A. Gaspar — nav, botões primários */
--g-med: #3D2490   /* azul médio — sub-headers, destaques */
--g-clr: #EAE7F5   /* lilás claro — backgrounds de card */
--g-brd: #C5BDE8   /* lilás borda */
```
Fontes: **Barlow + Barlow Condensed** (Google Fonts)

---

## Deploy

```powershell
git add obras/SP-10/custos/index.html
git commit -m "Mensagem descritiva"
git push origin main
```
GitHub Pages faz deploy automático em ~1-2 min após o push no branch `main`.

---

## Acessibilidade

Seletor de fonte no header (A / A+ / A++). Classes: `fonte-m` (médio) e `fonte-g` (grande)
aplicadas ao `<body>`. Persistência via `localStorage` (chave: `sp010-fonte`).
