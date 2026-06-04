# Portfolio Dashboard — Lucas Taveira

## Visão geral
Dashboard pessoal de acompanhamento de carteira de investimentos.
Site estático hospedado no GitHub Pages — zero backend, zero banco de dados, zero plataforma extra.

## Stack
- **Frontend:** HTML + JavaScript puro (sem frameworks)
- **Estilo:** CSS inline ou arquivo .css único — sem Tailwind, sem dependências de build
- **Gráficos:** Chart.js (via CDN) — sem instalação
- **Leitura do xlsx:** SheetJS (via CDN) — sem instalação
- **Cotações:** Yahoo Finance API (`.SA` para ativos B3) com fallback para Brapi.dev
- **Hospedagem:** GitHub Pages (gratuito, deploy automático via push)

## Estrutura de arquivos
```
portfolio-dashboard/
├── index.html          ← página única do dashboard
├── app.js              ← toda a lógica: leitura xlsx, cotações, cálculos, gráficos
├── style.css           ← estilos do dashboard
├── movimentacao.xlsx   ← base de dados (atualizada mensalmente pelo usuário)
└── CLAUDE.md
```

## Autenticação
- Tela de login simples antes do dashboard
- Senha verificada via hash SHA-256 no JavaScript
- Sessão mantida em sessionStorage (expira ao fechar o navegador)
- Hash da senha configurável no topo do app.js

## Fonte de dados
Arquivo: `movimentacao.xlsx` exportado do portal investidor.b3.com.br
Colunas: Entrada/Saída | Data | Movimentação | Produto | Instituição | Quantidade | Preço unitário | Valor da Operação

Tipos de movimentação relevantes (processar):
- `Transferência - Liquidação` → compras e vendas
- `Dividendo` → proventos
- `Juros Sobre Capital Próprio` → proventos
- `Rendimento` → proventos (ETFs)
- `Bonificação em Ativos` → entrada com custo zero
- `Desdobro` → ajuste de quantidade sem custo
- `Leilão de Fração` → ignorar

Tipos a ignorar:
- `Transferência` → movimentação entre corretoras
- `Atualização` → atualização de custódia

## Cotações em tempo real
Fonte primária: Yahoo Finance
```javascript
// Formato: TICKER.SA para ativos B3
fetch(`https://query1.finance.yahoo.com/v8/finance/chart/${ticker}.SA`)
```

Fallback (se Yahoo bloquear CORS): Brapi.dev
```javascript
fetch(`https://brapi.dev/api/quote/${ticker}`)
```

Ativos em carteira:
| Ticker | Tipo | Qt atual | PM (R$) |
|--------|------|----------|---------|
| IMAB11 | ETF  | 686      | 104,99  |
| IVVB11 | ETF  | 66       | 370,67  |
| DIVO11 | ETF  | 172      | 98,72   |
| ITUB4  | Ação | 49       | 24,79   |
| HASH11 | ETF  | 25       | 51,30   |
| AAPL34 | BDR  | 40       | 66,44   |
| BBAS3  | Ação | 46       | 31,20   |
| BBDC4  | Ação | 36       | 21,95   |
| MOVI3  | Ação | 50       | 15,47   |
| M1TA34 | BDR  | 25       | 30,00   |
| XPBR31 | BDR  | 1        | 166,00  |

## Regras de negócio críticas
- **Custo médio:** ponderado por quantidade; bonificações e desdobros entram com custo zero
- **P&L preço:** (preço atual - custo médio) × quantidade
- **P&L total:** P&L preço + proventos recebidos
- **XIRR:** calculado com método de Newton-Raphson em JavaScript
  - Fluxos: compras (negativo, na data) + proventos (positivo, na data) + valor de mercado hoje (positivo)
- **Formato de datas:** dd/MM/yyyy (padrão dos extratos B3)
- **Moeda:** sempre BRL

## Features do dashboard (prioridade)
1. Tela de login com senha
2. Cards de resumo: Patrimônio total | Custo total | P&L | Proventos recebidos | XIRR a.a.
3. Tabela de posições: ticker | qt | PM | preço atual | var. dia | valor mercado | P&L R$ | P&L % — verde/vermelho
4. Gráfico de proventos por ano (barras)
5. Gráfico de composição: custo vs mercado por ticker (barras)
6. Gráfico de evolução do patrimônio ao longo do tempo (linha)
7. Botão "Atualizar cotações" — rebusca os preços na API
8. Indicação de última atualização das cotações (timestamp)

## Visual
- Tema escuro (fundo #0f0f0f ou similar)
- Cards com bordas sutis e valores em destaque
- Tabela com linhas alternadas e cores verde (#22c55e) para P&L positivo, vermelho (#ef4444) para negativo
- Gráficos com Chart.js, paleta consistente com o tema escuro
- Responsivo — funciona no celular

## Deploy — GitHub Pages
1. Push para branch `main`
2. No GitHub: Settings → Pages → Source: Deploy from branch `main` → pasta `/` (root)
3. URL: `https://lucastaveiraa.github.io/portfolio-dashboard`

## Atualização mensal
1. Baixar extrato atualizado em investidor.b3.com.br
2. Renomear para `movimentacao.xlsx`
3. Substituir o arquivo na pasta local
4. `git add movimentacao.xlsx && git commit -m "update: extrato <mês>" && git push`
5. GitHub Pages atualiza automaticamente em ~1 minuto

## Convenções
- Todo o código em um único app.js (sem módulos ES6 para simplicidade máxima)
- Comentários em português
- Valores monetários sempre em BRL com 2 casas decimais
- Datas no formato dd/MM/yyyy internamente
