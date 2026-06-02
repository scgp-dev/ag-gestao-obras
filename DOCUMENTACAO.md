# Documentação Técnica — PMO Construtora A. Gaspar S.A.

> **Versão:** 1.0 — Mai/2026  
> **Responsável técnico:** Jailton Santos (Gerente de Engenharia/Planejamento)  
> **Contato:** jailtonsantos@construtoraagaspar.com.br

---

## 1. Visão Geral do Sistema

Sistema de painéis de controle gerencial para obras de infraestrutura da **Construtora A. Gaspar S.A.**,
desenvolvido como aplicação web estática hospedada no **GitHub Pages** (zero custo de infraestrutura,
zero backend, zero banco de dados).

Cada painel é um único arquivo HTML standalone que embarca todos os dados e a lógica de visualização.
Os dados são atualizados periodicamente exportando uma planilha Excel e reconstruindo o arquivo.

### 1.1 Acesso

| Recurso | URL |
|---------|-----|
| Portal principal | `https://scgp-dev.github.io/ag-gestao-obras/` |
| Portal SP-010 | `https://scgp-dev.github.io/ag-gestao-obras/obras/SP-10/` |
| Painel Financeiro de Custos SP-010 | `https://scgp-dev.github.io/ag-gestao-obras/obras/SP-10/custos/` |
| Painel Financeiro (Medições) SP-010 | `https://scgp-dev.github.io/ag-gestao-obras/obras/SP-10/financeiro.html` |
| Painel Avanço Físico SP-010 | `https://scgp-dev.github.io/ag-gestao-obras/obras/SP-10/fisico.html` |

### 1.2 Autenticação

Os painéis `financeiro.html` e `fisico.html` usam **Google OAuth 2.0** com whitelist de e-mails
(array `WHITELIST` no início do script de cada painel). O painel de custos não tem autenticação.

Client ID Google: `167891547522-nq3o4hb6qo1g1egq21c2s1sdsr1v124o.apps.googleusercontent.com`

---

## 2. Stack Tecnológica

| Camada | Tecnologia | Justificativa |
|--------|-----------|---------------|
| Linguagem | HTML5 + CSS3 + JavaScript (ES6+) puro | Zero dependências de build |
| Gráficos | Chart.js 4.4.1 (CDN) | Biblioteca madura, sem instalação |
| Fontes | Barlow + Barlow Condensed (Google Fonts) | Identidade visual A. Gaspar |
| Hospedagem | GitHub Pages (branch `main`) | Gratuito, deploy automático via push |
| Dados | JSON embutido no HTML | Sem necessidade de API ou backend |
| Controle de versão | Git + GitHub (`scgp-dev/ag-gestao-obras`) | Histórico completo de alterações |

**Não há:** Node.js, npm, webpack, React, banco de dados, servidor, Docker.

---

## 3. Estrutura de Arquivos

```
ponte-grauna-painel/              ← diretório local de desenvolvimento
│
├── .gitignore                    ← exclui: *.xlsx, dados.json, custos_template.html, PDFs
├── CLAUDE.md                     ← referência técnica rápida (para uso com IA)
├── DOCUMENTACAO.md               ← este arquivo
│
├── index.html                    ← Portal raiz — lista todas as obras
├── ponte_preview.html            ← Asset estático de preview
│
└── obras/
    │
    ├── SP-10/                    ← Obra: Ponte Graúna-Gaivotas (São Paulo)
    │   │                            Contrato 029/SIURB/2025 — R$ 347,5 Mi
    │   │
    │   ├── index.html            ← Portal da obra (Físico | Financeiro | Custos)
    │   ├── fisico.html           ← Painel de Avanço Físico (requer login Google)
    │   ├── financeiro.html       ← Painel Financeiro / Medições (requer login Google)
    │   │
    │   └── custos/               ← Módulo de Controle de Custos
    │       ├── index.html        ← PAINEL PRINCIPAL (548 KB, standalone)
    │       ├── dados.json        ← Exportação do Excel (.gitignored — não sobe)
    │       └── [custos_template.html] ← Template de desenvolvimento (.gitignored)
    │
    └── MT-02/                    ← Obra: OAE 3 — Ponte do Rio Vermelho (Mato Grosso)
        ├── index.html
        └── MT-002_Relatorio_Gerencial.html
```

---

## 4. Painel Financeiro de Custos — Arquitetura Detalhada

### 4.1 Fonte de Dados

**Planilha:** `SP - 010 - ATÉ 30-04-2026.xlsx` (atualizada mensalmente)
- Aba: `Baixas`
- Linhas: ~2.647 (2.454 válidas)
- Colunas utilizadas:

| Col | Campo | Descrição |
|-----|-------|-----------|
| B | TÍTULO | Código do título |
| C | NOME | Razão social do fornecedor |
| D | CC | Centro de custo |
| E | DATA | Data do pagamento (DD/MM/YYYY) |
| F | HISTÓRICO | Descrição do lançamento |
| G | VALOR | Valor bruto |
| H | MULTA | Multa |
| I | DESCONTO | Desconto |
| J | TOTAL PAGO | Valor efetivamente pago |
| K | BANCO | Banco do pagamento |

### 4.2 Processo de Atualização dos Dados

O arquivo `dados.json` é gerado por **PowerShell via COM Object do Excel**:

```powershell
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$wb = $xl.Workbooks.Open("C:\caminho\SP - 010 - ATÉ 30-04-2026.xlsx")
$ws = $wb.Sheets.Item("Baixas")

$rows = $ws.UsedRange.Rows.Count
$records = @()
for ($i = 2; $i -le $rows; $i++) {
    $nome = $ws.Cells.Item($i, 3).Value2
    if ([string]::IsNullOrWhiteSpace($nome)) { continue }
    
    $dataVal = $ws.Cells.Item($i, 5).Value2
    $data = if ($dataVal) { [DateTime]::FromOADate($dataVal).ToString("dd/MM/yyyy") } else { "" }
    $mesAno = if ($dataVal) { [DateTime]::FromOADate($dataVal).ToString("MM/yyyy") } else { "" }
    
    $records += [ordered]@{
        nome   = $nome.Trim()
        data   = $data
        mesAno = $mesAno
        hist   = ($ws.Cells.Item($i, 6).Value2 -replace '"','\"').Trim()
        valor  = [math]::Round($ws.Cells.Item($i, 7).Value2, 2)
        multa  = [math]::Round($ws.Cells.Item($i, 8).Value2, 2)
        desc   = [math]::Round($ws.Cells.Item($i, 9).Value2, 2)
        total  = [math]::Round($ws.Cells.Item($i, 10).Value2, 2)
        banco  = ($ws.Cells.Item($i, 11).Value2)
        titulo = ($ws.Cells.Item($i, 2).Value2)
    }
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$json = $records | ConvertTo-Json -Depth 3 -Compress
[System.IO.File]::WriteAllText("dados.json", $json, [System.Text.UTF8Encoding]::new($false))
Write-Host "Exportados $($records.Count) registros"
```

### 4.3 Build do HTML Final

Com `custos_template.html` e `dados.json` prontos:

```powershell
$tmpl  = Get-Content "custos_template.html" -Raw -Encoding UTF8
$dados = Get-Content "dados.json" -Raw -Encoding UTF8
$out   = $tmpl.Replace('%%DADOS_JSON%%', $dados)
[System.IO.File]::WriteAllText("index.html", $out, [System.Text.UTF8Encoding]::new($false))
```

O template contém o marcador `%%DADOS_JSON%%` onde o JSON é injetado.

### 4.4 Deploy

```bash
git add obras/SP-10/custos/index.html
git commit -m "Atualiza dados até MM/AAAA"
git push origin main
```

GitHub Pages faz o deploy automaticamente em ~1-2 minutos.

---

## 5. Motor de Categorização

### 5.1 Estrutura

A função `categorizar(r)` recebe um registro e retorna uma string com o nome da categoria.
Analisa dois campos: `r.nome` (fornecedor) e `r.hist` (histórico), ambos convertidos para MAIÚSCULO.

```javascript
function categorizar(r) {
  const h = (r.hist||'').toUpperCase();   // histórico
  const n = (r.nome||'').toUpperCase();   // nome do fornecedor
  const hn = h + ' ' + n;                 // concatenação para buscas combinadas
  // ... regras em cascata ...
  return 'Outros';
}
```

### 5.2 Categorias Disponíveis (18 + Outros)

| # | ID | Classe CSS | Cor |
|---|----|-----------|-----|
| 0 | Aço Estrutural | c0 | Azul |
| 1 | Concreto | c1 | Verde |
| 2 | Fundação | c2 | Laranja |
| 3 | Embarcações / Flutuantes | c3 | Azul claro |
| 4 | Equipamentos (Compra) | c4 | Roxo |
| 5 | Locação de Equipamentos | c5 | Verde escuro |
| — | Locação de Veículos Leves | c18 | Azul céu |
| 6 | Mão de Obra | c6 | Marrom |
| 7 | Encargos Trabalhistas | c7 | Vermelho |
| 8 | Impostos e Retenções | c8 | Índigo |
| 9 | Consultoria e Projetos | c9 | Azul índigo |
| 10 | Combustíveis | c10 | Vermelho escuro |
| 11 | Seguros | c11 | Verde escuro |
| 12 | Aluguel de Imóveis | c12 | Marrom |
| 13 | EPI / Uniformes | c13 | Cinza |
| 14 | Transporte e Logística | c14 | Cinza escuro |
| 15 | Prestação de Contas | c15 | Roxo |
| 16 | Administrativo | c16 | Cinza |
| 17 | Outros | c17 | Cinza claro |

### 5.3 Regras de Categorização (por prioridade)

As regras são aplicadas em **cascata** — a primeira que bater retorna. Ordem importa.

**Aço / Fundação (Gerdau e ArcelorMittal):**
```
ArcelorMittal + BOBINA no hist     → Fundação
CALTUBOS (qualquer hist)           → Fundação
Gerdau/ArcelorMittal + VERGALHÃO/ARAME/PREGO → Aço Estrutural
Gerdau/ArcelorMittal + qualquer outra coisa  → Embarcações / Flutuantes
```

**Mão de Obra** (inclui todos os encargos de folha):
```
Keywords no hist: INSS FOLHA, INSS 13, FGTS, VALE TRANSPORT,
                  CONTRIB RETRIB, RESCISAO, 13° SAL, TAR TRANSF SALARIO,
                  FERIAS, ADIANTAMENTO, FOLHA/
CONTRIB ADICIONAL SENAI → Mão de Obra
```

**Aluguel de Imóveis** (inclui variações e erros de grafia):
```
Keywords: ALUGUEL, ALG REFERENTE, CAUCAO, CAUCOES,
          LOCACAO DE IMOVEL, LOCACAO DE ESPACO
Fornecedores conhecidos + hist contendo ALG ou CAUCAO
```

**Combustíveis:**
```
TICKLOG (qualquer campo)  → Combustíveis
DIESEL no hist            → Outros (gerador/maquinário, não frota)
Raizen, Vibra, ICP        → Combustíveis
```

---

## 6. Faturamento — Medições SIURB/PMSP

Os valores de faturamento são embutidos diretamente no JS do painel
(objeto `FATURAMENTO`). São extraídos das **Notas Fiscais emitidas** pela A. Gaspar
para a SIURB/PMSP (PDFs armazenados localmente, não commitados).

| # | NF | Período | Valor |
|---|----|---------|-------|
| 1ª | NFS 5869 | Jun–Jul/2025 | R$ 945.315,38 |
| 2ª | NFS 5979 | Ago/2025 | R$ 1.070.235,50 |
| 3ª | NFS 5999 | Set/2025 | R$ 1.469.509,41 |
| 4ª | NFS 6000 | Out/2025 | R$ 5.877.844,78 |
| 5ª | NFS 6001 | Nov/2025 | R$ 5.528.933,79 |
| 6ª | NFS 6040 | Jan/2026 | R$ 16.165.294,08 |
| 7ª | NFS 6089 | Fev/2026 | R$ 10.086.985,91 |
| 8ª | NFS 6126 | Mar/2026 | R$ 10.237.937,56 |
| 9ª | NFN 31 | Abr/2026 | R$ 11.018.518,57 |
| | **Total** | | **R$ 62.400.574,98** |

Para atualizar quando sair nova medição: editar o objeto `FATURAMENTO`
no `index.html` e adicionar a nova chave `'MM/YYYY': valor`.

---

## 7. Funcionalidades do Painel de Custos

### 7.1 Filtros Globais (afetam todas as abas)
- **Mês:** dropdown com todos os períodos disponíveis
- **Apropriação:** dropdown com as 19 categorias
- **Busca livre:** pesquisa em nome do fornecedor e histórico

### 7.2 Visão Geral
- **KPI Hero Block:** Custo Total, Faturamento, Resultado (Fat.−Custo) com ícones dinâmicos
- **Curva S:** barras (custo mensal) + 3 linhas (custo acumulado, faturado acumulado, resultado acumulado)
- **Diagnóstico Mensal:** tabela custo×faturamento e composição do custo por categoria

### 7.3 Apropriações
- Gráfico de barras horizontal por categoria
- Donut com % de participação
- Tabela de evolução mensal por categoria

### 7.4 Fornecedores
- Top 15 por valor (gráfico de barras)
- Gráfico de concentração (top 5 vs demais)
- Tabela de ranking completa com busca

### 7.5 Extrato
- Tabela paginada (25/50/100 linhas)
- **Ordenação clicável:** Data, Fornecedor, Valor, Total Pago (▲/▼)
- **Total no rodapé:** `tfoot` com soma total dos registros selecionados
- Destaque colorido por categoria (badges)

### 7.6 Acessibilidade
- Seletor de fonte no header: **A** (normal) / **A+** (médio +15%) / **A++** (grande +33%)
- Preferência salva em `localStorage` (chave: `sp010-fonte`)

---

## 8. Identidade Visual

### 8.1 Variáveis CSS

```css
--g-esc: #1A0D4A   /* azul muito escuro — headers, fundos */
--g-pri: #28156E   /* azul A. Gaspar — nav, botões primários */
--g-med: #3D2490   /* azul médio — sub-headers, destaques */
--g-clr: #EAE7F5   /* lilás claro — backgrounds de card */
--g-brd: #C5BDE8   /* lilás borda */
--verde: #166534
--vd-clr: #DCFCE7
--vermelho: #991B1B
--vm-clr: #FEE2E2
--cinza: #64748B
--borda: #E2E8F0
--fundo: #F5F4FB
```

### 8.2 Tipografia
- **Corpo / UI:** `Barlow` (pesos 300, 400, 500, 600)
- **Valores / Títulos de destaque:** `Barlow Condensed` (pesos 300, 400, 600, 700, 800)
- Carregadas via Google Fonts (CDN)

### 8.3 Componentes Reutilizáveis

| Componente | Classe | Uso |
|-----------|--------|-----|
| Card de conteúdo | `.chart-card` | Gráficos |
| Card de tabela | `.table-card` | Tabelas |
| Badge de categoria | `.cat .c0`–`.c18` | Labels coloridos |
| Status bar | `.status-bar` | Barra de info no topo |
| Paginação | `.pag` | Navegação de páginas |
| Separador de seção | `.diag-hd` | Títulos entre seções |

---

## 9. Como Adicionar Uma Nova Obra

1. Criar pasta `obras/[CODIGO]/`
2. Copiar `obras/SP-10/index.html` como base para o portal da obra
3. Adaptar: nome da obra, contratos, links para os painéis
4. Para painel de custos: usar o fluxo Excel → JSON → template → HTML
5. Adicionar card no portal raiz `index.html`
6. Commitar e fazer push

---

## 10. Fluxo de Trabalho Mensal (Atualização de Dados)

```
1. Receber planilha Excel atualizada do financeiro
   └── Verificar aba "Baixas" com novos lançamentos

2. Executar script PowerShell de exportação
   └── Gera dados.json (~479 KB)

3. Verificar novos fornecedores no JSON
   └── Conferir se precisam de regra nova no categorizar()
   └── Verificar se novos lançamentos estão na categoria correta

4. Se nova medição SIURB emitida:
   └── Atualizar objeto FATURAMENTO no index.html
   └── Adicionar linha na tabela de NFs (seção 6 desta documentação)

5. Executar script de build
   └── Gera novo index.html com dados atualizados

6. Abrir index.html localmente para verificação visual

7. Git add + commit + push
   └── GitHub Pages atualiza em ~1-2 minutos
```

---

## 11. Limitações Atuais e Próximos Passos Sugeridos

| Limitação | Impacto | Sugestão |
|-----------|---------|----------|
| Dados embutidos no HTML | Rebuild manual a cada atualização | Integrar com Google Sheets API ou Supabase |
| Sem autenticação no painel de custos | Qualquer pessoa com o link acessa | Adicionar Google OAuth (padrão dos outros painéis) |
| Categorização manual | Novos fornecedores podem ir para "Outros" | Sistema de feedback no painel para reclassificar |
| Sem histórico de versões dos dados | Não dá pra ver dados de meses anteriores | Separar JSON por período |
| custos_template.html perdido se trocar máquina | Não é possível reconstruir sem o template | Criar repositório privado separado para arquivos de desenvolvimento |

---

## 12. Repositório e Credenciais

| Item | Valor |
|------|-------|
| GitHub | `https://github.com/scgp-dev/ag-gestao-obras` |
| GitHub user | `scgp-dev` |
| Git user (commits) | `scgp-dev <jailtonsantos@construtoraagaspar.com.br>` |
| Google OAuth Client ID | `167891547522-nq3o4hb6qo1g1egq21c2s1sdsr1v124o.apps.googleusercontent.com` |
| Whitelist de acesso | jailtonsantos@, marcofrensel@, rodrigo.asfora@, marthaflor@ (todos @construtoraagaspar.com.br) |

---

*Documentação gerada em Mai/2026. Atualizar sempre que houver mudanças estruturais no sistema.*
