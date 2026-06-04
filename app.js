/* ============================================================
   CONFIGURAÇÕES
   ============================================================ */

// Hash SHA-256 de "imeri2026"
const SENHA_HASH = '977547abed280c755c1e36cb7f94b36ace9ab53e71bff476e855312d2e5347a1';

const ARQUIVO_XLSX = 'movimentacao.xlsx';

// Tipos de ativo conhecidos (complementado pelo que vier do xlsx)
const TIPOS_TICKER = {
  IMAB11: 'ETF',  IVVB11: 'ETF',  DIVO11: 'ETF',  HASH11: 'ETF',
  ITUB4:  'Ação', BBAS3:  'Ação', BBDC4:  'Ação', MOVI3:  'Ação',
  AAPL34: 'BDR',  M1TA34: 'BDR',  XPBR31: 'BDR',
};

// Tipos de movimentação a processar
const TIPO_COMPRA_VENDA  = 'transferência - liquidação';
const TIPO_DIVIDENDO     = ['dividendo', 'juros sobre capital próprio', 'rendimento'];
const TIPO_BONIFICACAO   = 'bonificação em ativos';
const TIPO_DESDOBRO      = 'desdobro';
const TIPO_IGNORAR       = ['transferência', 'atualização', 'leilão de fração'];

// Paleta de cores para gráficos
const PALETA = [
  '#3b82f6','#22c55e','#f59e0b','#ef4444','#8b5cf6',
  '#06b6d4','#f97316','#ec4899','#14b8a6','#a855f7','#84cc16',
];

// Estado global da aplicação
const estado = {
  posicoes: {},   // { TICKER: { qt, custoTotal, pm, tipo, proventos, fluxos } }
  historico: [],  // [{ data: Date, capitalInvestido: number }]
  cotacoes: {},
  charts: {},
  carregado: false,
};

/* ============================================================
   AUTENTICAÇÃO
   ============================================================ */

async function sha256(texto) {
  const data = new TextEncoder().encode(texto);
  const buf  = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

async function fazerLogin(e) {
  e.preventDefault();
  const senha = document.getElementById('campo-senha').value;
  const hash  = await sha256(senha);
  if (hash === SENHA_HASH) {
    sessionStorage.setItem('auth', 'ok');
    mostrarDashboard();
  } else {
    document.getElementById('login-erro').textContent = 'Senha incorreta. Tente novamente.';
    document.getElementById('campo-senha').value = '';
    document.getElementById('campo-senha').focus();
  }
}

function fazerLogout() {
  sessionStorage.removeItem('auth');
  document.getElementById('dashboard').classList.add('hidden');
  document.getElementById('tela-login').classList.remove('hidden');
  document.getElementById('campo-senha').value = '';
}

function verificarSessao() {
  return sessionStorage.getItem('auth') === 'ok';
}

function mostrarDashboard() {
  document.getElementById('tela-login').classList.add('hidden');
  document.getElementById('dashboard').classList.remove('hidden');
  if (!estado.carregado) {
    inicializarDashboard();
  }
}

/* ============================================================
   UTILITÁRIOS
   ============================================================ */

function mostrarLoading(visivel) {
  document.getElementById('loading-overlay').classList.toggle('hidden', !visivel);
}

function formatarBRL(valor) {
  if (valor === null || valor === undefined || isNaN(valor)) return '—';
  return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatarPct(valor) {
  if (valor === null || valor === undefined || isNaN(valor)) return '—';
  return (valor >= 0 ? '+' : '') + valor.toFixed(2) + '%';
}

// Converte string "dd/MM/yyyy" ou serial Excel para Date
function parsearData(valor) {
  if (!valor) return null;
  if (typeof valor === 'number') {
    // Serial do Excel
    const d = new Date((valor - 25569) * 86400 * 1000);
    d.setMinutes(d.getMinutes() + d.getTimezoneOffset());
    return d;
  }
  const s = String(valor).trim();
  // dd/MM/yyyy
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  // yyyy-MM-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Extrai ticker do nome do produto (ex: "IMAB11 - ISHARES..." → "IMAB11")
function extrairTicker(produto) {
  if (!produto) return null;
  const s = String(produto).trim().toUpperCase();
  const m = s.match(/^([A-Z]{4}\d{1,2}[A-Z]?\d?)/);
  return m ? m[1] : s.split(/[\s-]/)[0] || null;
}

// Normaliza número que pode vir como string "1.234,56" ou "1234.56"
function parsearNumero(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  const s = String(val).trim().replace(/\./g, '').replace(',', '.');
  return parseFloat(s) || 0;
}

/* ============================================================
   LEITURA DO XLSX
   ============================================================ */

async function lerXLSX() {
  const resp = await fetch(ARQUIVO_XLSX + '?t=' + Date.now());
  if (!resp.ok) throw new Error(`Arquivo ${ARQUIVO_XLSX} não encontrado (${resp.status})`);
  const buf  = await resp.arrayBuffer();
  const wb   = XLSX.read(buf, { type: 'array', cellDates: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows;
}

function semAcento(str) {
  return String(str).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

// Detecta índice de cada coluna pelo cabeçalho
function detectarColunas(cabecalho) {
  const c = {};
  cabecalho.forEach((col, i) => {
    const s = semAcento(col);
    if (s.includes('entrada') || s.includes('saida'))    c.entradaSaida = i;
    else if (s === 'data' || s.startsWith('data '))      c.data         = i;
    else if (s.includes('movimentac'))                    c.tipo         = i;
    else if (s.includes('produto'))                       c.produto      = i;
    else if (s.includes('quantidade'))                    c.quantidade   = i;
    else if (s.includes('preco') || s.includes('unit'))  c.precoUnit    = i;
    else if (s.includes('valor'))                         c.valorOp      = i;
  });
  return c;
}

/* ============================================================
   PROCESSAMENTO DE MOVIMENTAÇÕES
   ============================================================ */

function processarMovimentacoes(rows) {
  // Encontra linha do cabeçalho (primeira que tem pelo menos 5 células preenchidas)
  let idxCab = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const preenchidas = rows[i].filter(v => String(v).trim() !== '').length;
    if (preenchidas >= 5) { idxCab = i; break; }
  }

  const cab = rows[idxCab];
  const c   = detectarColunas(cab);

  // Debug: primeiras 5 linhas de dados
  console.log('[DEBUG] Cabeçalho detectado (linha', idxCab, '):', cab);
  console.log('[DEBUG] Índices de colunas:', c);
  const linhasDebug = rows.slice(idxCab + 1, idxCab + 6).filter(r => r && !r.every(v => String(v).trim() === ''));
  linhasDebug.forEach((row, i) => {
    console.log(`[DEBUG] Linha ${i + 1} raw:`, row);
    console.log(`[DEBUG] Linha ${i + 1} parsed → qt=${parsearNumero(row[c.quantidade])} preco=${parsearNumero(row[c.precoUnit])} tipo="${semAcento(String(row[c.tipo] ?? ''))}" entSai="${String(row[c.entradaSaida] ?? '').trim().toLowerCase()}"`);
  });

  const posicoes = {};
  const fluxosGlobais = [];  // para XIRR global
  const historicoMap  = {};  // "yyyy-MM" → capitalInvestido acumulado

  for (let i = idxCab + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(v => String(v).trim() === '')) continue;

    const entSai = String(row[c.entradaSaida] ?? '').trim().toLowerCase();
    const rawTipo = String(row[c.tipo] ?? '').trim();
    const tipo    = semAcento(rawTipo);
    const produto = String(row[c.produto] ?? '').trim();
    const ticker  = extrairTicker(produto);
    if (!ticker) continue;

    const data      = parsearData(row[c.data]);
    const quantidade = parsearNumero(row[c.quantidade]);
    const precoUnit  = parsearNumero(row[c.precoUnit]);
    const valorOp    = parsearNumero(row[c.valorOp]);

    // Ignorar movimentações sem relevância
    const ignorar = TIPO_IGNORAR.some(t => tipo.includes(semAcento(t)));
    if (ignorar && tipo !== semAcento(TIPO_COMPRA_VENDA)) continue;

    if (!posicoes[ticker]) {
      posicoes[ticker] = {
        tipo: TIPOS_TICKER[ticker] || detectarTipoTicker(ticker),
        qt: 0,
        custoTotal: 0,
        pm: 0,
        proventos: 0,
        fluxos: [],
      };
    }

    const pos = posicoes[ticker];

    // COMPRA / VENDA
    if (tipo.includes('transferencia') && tipo.includes('liquidacao')) {
      const isEntrada = entSai === 'credito' || entSai === 'entrada';
      const isDebito  = entSai === 'debito'  || entSai === 'saida';

      if (isEntrada && quantidade > 0) {
        // Compra: atualiza custo médio ponderado
        const novoCusto = pos.custoTotal + (precoUnit > 0 ? precoUnit * quantidade : valorOp);
        pos.qt          += quantidade;
        pos.custoTotal  = novoCusto;
        pos.pm          = pos.qt > 0 ? pos.custoTotal / pos.qt : 0;
        const saida     = -(precoUnit > 0 ? precoUnit * quantidade : valorOp);
        if (data) pos.fluxos.push({ valor: saida, data });

        // Histórico de capital
        if (data) {
          const chave = `${data.getFullYear()}-${String(data.getMonth()+1).padStart(2,'0')}`;
          historicoMap[chave] = (historicoMap[chave] || 0) + Math.abs(saida);
        }
      } else if (isDebito && quantidade > 0) {
        // Venda: baixa custo proporcionalmente
        const custoVendido = pos.pm * quantidade;
        pos.qt         -= quantidade;
        pos.custoTotal -= custoVendido;
        if (pos.qt <= 0) { pos.qt = 0; pos.custoTotal = 0; }
        pos.pm = pos.qt > 0 ? pos.custoTotal / pos.qt : 0;
        const entrada = precoUnit > 0 ? precoUnit * quantidade : valorOp;
        if (data) pos.fluxos.push({ valor: entrada, data });
      }
    }

    // PROVENTOS
    else if (TIPO_DIVIDENDO.some(t => tipo.includes(t))) {
      const valor = valorOp > 0 ? valorOp : quantidade * precoUnit;
      pos.proventos += valor;
      if (data && valor > 0) pos.fluxos.push({ valor, data });
    }

    // BONIFICAÇÃO — entrada sem custo
    else if (tipo.includes('bonificac')) {
      pos.qt += quantidade;
      // custo médio ajustado: custo total não muda, qt aumenta
      pos.pm = pos.qt > 0 ? pos.custoTotal / pos.qt : 0;
    }

    // DESDOBRO — ajuste de quantidade
    else if (tipo.includes('desdobro')) {
      pos.qt += quantidade;
      pos.pm  = pos.qt > 0 ? pos.custoTotal / pos.qt : 0;
    }
  }

  // Remove posições zeradas
  Object.keys(posicoes).forEach(t => {
    if (posicoes[t].qt <= 0 && posicoes[t].proventos === 0) delete posicoes[t];
  });

  // Constrói série histórica (capital acumulado mês a mês)
  const historico = construirHistorico(historicoMap);

  return { posicoes, historico };
}

function detectarTipoTicker(ticker) {
  if (/\d{2}$/.test(ticker)) return 'ETF';
  if (/34$|31$|32$|33$/.test(ticker)) return 'BDR';
  if (/11$/.test(ticker)) return 'FII';
  return 'Ação';
}

function construirHistorico(mapa) {
  const chaves = Object.keys(mapa).sort();
  if (!chaves.length) return [];
  let acumulado = 0;
  return chaves.map(ch => {
    acumulado += mapa[ch];
    const [ano, mes] = ch.split('-');
    return { data: new Date(Number(ano), Number(mes) - 1, 1), capitalInvestido: acumulado };
  });
}

/* ============================================================
   COTAÇÕES
   ============================================================ */

async function buscarCotacaoYahoo(tickers) {
  const resultados = {};
  // Tenta um ticker de cada vez para evitar rate limit
  for (const tk of tickers) {
    try {
      const url  = `https://query1.finance.yahoo.com/v8/finance/chart/${tk}.SA?interval=1d&range=1d`;
      const resp = await fetch(url, { signal: AbortSignal.timeout(5000) });
      if (!resp.ok) continue;
      const json = await resp.json();
      const meta = json?.chart?.result?.[0]?.meta;
      if (meta?.regularMarketPrice) {
        resultados[tk] = {
          preco:   meta.regularMarketPrice,
          varDia:  meta.regularMarketChangePercent ?? 0,
          fonte: 'Yahoo',
        };
      }
    } catch { /* ignora e usa fallback */ }
  }
  return resultados;
}

async function buscarCotacaoBrapi(tickers) {
  const resultados = {};
  // Brapi aceita lista separada por vírgula
  const chunks = [];
  for (let i = 0; i < tickers.length; i += 10) chunks.push(tickers.slice(i, i + 10));

  for (const chunk of chunks) {
    try {
      const lista = chunk.join(',');
      const url   = `https://brapi.dev/api/quote/${lista}?fundamental=false`;
      const resp  = await fetch(url, { signal: AbortSignal.timeout(8000) });
      if (!resp.ok) continue;
      const json = await resp.json();
      for (const q of (json.results ?? [])) {
        if (q.regularMarketPrice) {
          resultados[q.symbol.replace('.SA', '')] = {
            preco:  q.regularMarketPrice,
            varDia: q.regularMarketChangePercent ?? 0,
            fonte: 'Brapi',
          };
        }
      }
    } catch { /* ignora */ }
  }
  return resultados;
}

async function atualizarCotacoes() {
  const tickers = Object.keys(estado.posicoes).filter(t => estado.posicoes[t].qt > 0);
  if (!tickers.length) return;

  document.getElementById('btn-atualizar').disabled = true;
  mostrarLoading(true);

  try {
    // Brapi como fonte primária (Yahoo bloqueado por CORS no browser)
    let cotacoes = await buscarCotacaoBrapi(tickers);

    // Fallback Yahoo para os que não vieram do Brapi
    const faltantes = tickers.filter(t => !cotacoes[t]);
    if (faltantes.length) {
      const yahoo = await buscarCotacaoYahoo(faltantes);
      Object.assign(cotacoes, yahoo);
    }

    estado.cotacoes = cotacoes;

    atualizarTabelaComCotacoes();
    atualizarCardsComCotacoes();
    renderizarGraficoComposicao();

    const agora = new Date().toLocaleString('pt-BR');
    const fontes = [...new Set(Object.values(cotacoes).map(c => c.fonte))].join('/');
    document.getElementById('ultima-atualizacao').textContent =
      `Última atualização: ${agora} — fonte: ${fontes || '—'}`;
  } catch (err) {
    console.error('Erro ao buscar cotações:', err);
  } finally {
    mostrarLoading(false);
    document.getElementById('btn-atualizar').disabled = false;
  }
}

/* ============================================================
   CÁLCULO DE XIRR (Newton-Raphson)
   ============================================================ */

function calcularXIRR(fluxos) {
  if (!fluxos || fluxos.length < 2) return null;
  const ordered = [...fluxos].sort((a, b) => a.data - b.data);
  const t0 = ordered[0].data.getTime();
  const MS_ANO = 365.25 * 24 * 3600 * 1000;

  function npv(r) {
    return ordered.reduce((s, f) => {
      const t = (f.data.getTime() - t0) / MS_ANO;
      return s + f.valor / Math.pow(1 + r, t);
    }, 0);
  }

  function dnpv(r) {
    return ordered.reduce((s, f) => {
      const t = (f.data.getTime() - t0) / MS_ANO;
      return s - t * f.valor / Math.pow(1 + r, t + 1);
    }, 0);
  }

  let r = 0.1;
  for (let i = 0; i < 200; i++) {
    const d = dnpv(r);
    if (Math.abs(d) < 1e-10) break;
    const nr = r - npv(r) / d;
    if (Math.abs(nr - r) < 1e-8) return nr;
    r = nr;
    if (r < -0.999 || r > 100) break;
  }
  return isFinite(r) ? r : null;
}

function construirFluxosGlobais(posicoes, cotacoes) {
  const todos = [];
  Object.keys(posicoes).forEach(tk => {
    const pos = posicoes[tk];
    todos.push(...pos.fluxos);

    // Valor de mercado atual como fluxo positivo na data de hoje
    const cot = cotacoes[tk];
    if (cot && pos.qt > 0) {
      todos.push({ valor: cot.preco * pos.qt, data: new Date() });
    } else if (pos.qt > 0) {
      // Sem cotação: usa custo como proxy
      todos.push({ valor: pos.custoTotal, data: new Date() });
    }
  });
  return todos;
}

/* ============================================================
   RENDERIZAÇÃO — TABELA
   ============================================================ */

function renderizarTabelaInicial(posicoes) {
  const tbody = document.getElementById('corpo-tabela');
  tbody.innerHTML = '';

  const tickers = Object.keys(posicoes).filter(t => posicoes[t].qt > 0)
    .sort((a, b) => posicoes[b].custoTotal - posicoes[a].custoTotal);

  if (!tickers.length) {
    tbody.innerHTML = '<tr><td colspan="9" class="loading">Nenhuma posição encontrada.</td></tr>';
    return;
  }

  for (const tk of tickers) {
    const pos = posicoes[tk];
    const tr  = document.createElement('tr');
    tr.dataset.ticker = tk;
    tr.innerHTML = `
      <td class="ticker-cell">${tk}</td>
      <td><span class="tipo-badge ${pos.tipo.toLowerCase()}">${pos.tipo}</span></td>
      <td>${pos.qt.toLocaleString('pt-BR')}</td>
      <td>${formatarBRL(pos.pm)}</td>
      <td class="col-preco">—</td>
      <td class="col-var">—</td>
      <td class="col-vmercado">${formatarBRL(pos.custoTotal)}</td>
      <td class="col-pl">—</td>
      <td class="col-plpct">—</td>
    `;
    tbody.appendChild(tr);
  }
}

function atualizarTabelaComCotacoes() {
  const posicoes = estado.posicoes;
  const cotacoes = estado.cotacoes;

  document.querySelectorAll('#corpo-tabela tr[data-ticker]').forEach(tr => {
    const tk  = tr.dataset.ticker;
    const pos = posicoes[tk];
    const cot = cotacoes[tk];
    if (!pos) return;

    if (cot) {
      const vm   = cot.preco * pos.qt;
      const pl   = vm - pos.custoTotal;
      const plPct = pos.custoTotal > 0 ? (pl / pos.custoTotal) * 100 : 0;
      const varDia = cot.varDia;

      tr.querySelector('.col-preco').textContent    = formatarBRL(cot.preco);
      tr.querySelector('.col-var').textContent      = formatarPct(varDia);
      tr.querySelector('.col-var').className        = 'col-var ' + (varDia >= 0 ? 'positivo' : 'negativo');
      tr.querySelector('.col-vmercado').textContent = formatarBRL(vm);
      tr.querySelector('.col-pl').textContent       = formatarBRL(pl);
      tr.querySelector('.col-pl').className         = 'col-pl ' + (pl >= 0 ? 'positivo' : 'negativo');
      tr.querySelector('.col-plpct').textContent    = formatarPct(plPct);
      tr.querySelector('.col-plpct').className      = 'col-plpct ' + (plPct >= 0 ? 'positivo' : 'negativo');
    }
  });
}

/* ============================================================
   RENDERIZAÇÃO — CARDS
   ============================================================ */

function renderizarCardsIniciais(posicoes) {
  const custoTotal   = Object.values(posicoes).reduce((s, p) => s + p.custoTotal, 0);
  const proventos    = Object.values(posicoes).reduce((s, p) => s + p.proventos, 0);

  document.getElementById('card-custo').textContent     = formatarBRL(custoTotal);
  document.getElementById('card-proventos').textContent = formatarBRL(proventos);
  // Patrimônio, P&L e XIRR ficam aguardando cotações
}

function atualizarCardsComCotacoes() {
  const posicoes = estado.posicoes;
  const cotacoes = estado.cotacoes;

  let patrimonio  = 0;
  let custoTotal  = 0;
  let proventos   = 0;

  Object.keys(posicoes).forEach(tk => {
    const pos = posicoes[tk];
    const cot = cotacoes[tk];
    custoTotal += pos.custoTotal;
    proventos  += pos.proventos;
    if (cot && pos.qt > 0) {
      patrimonio += cot.preco * pos.qt;
    } else {
      patrimonio += pos.custoTotal; // fallback
    }
  });

  const pl    = patrimonio - custoTotal + proventos;
  const plPct = custoTotal > 0 ? (pl / custoTotal) * 100 : 0;

  document.getElementById('card-patrimonio').textContent = formatarBRL(patrimonio);
  document.getElementById('card-custo').textContent      = formatarBRL(custoTotal);
  document.getElementById('card-proventos').textContent  = formatarBRL(proventos);

  const elPl = document.getElementById('card-pl');
  elPl.textContent = `${formatarBRL(pl)} (${formatarPct(plPct)})`;
  elPl.className   = 'card-value ' + (pl >= 0 ? 'positivo' : 'negativo');

  // XIRR
  const fluxos = construirFluxosGlobais(posicoes, cotacoes);
  const xirr   = calcularXIRR(fluxos);
  const elXirr = document.getElementById('card-xirr');
  if (xirr !== null) {
    elXirr.textContent = formatarPct(xirr * 100);
    elXirr.className   = 'card-value ' + (xirr >= 0 ? 'positivo' : 'negativo');
  } else {
    elXirr.textContent = '—';
  }
}

/* ============================================================
   GRÁFICOS
   ============================================================ */

function destruirGrafico(id) {
  if (estado.charts[id]) {
    estado.charts[id].destroy();
    delete estado.charts[id];
  }
}

const optBase = {
  responsive: true,
  maintainAspectRatio: true,
  plugins: {
    legend: { labels: { color: '#94a3b8', font: { size: 11 } } },
    tooltip: { callbacks: {
      label: ctx => {
        const v = ctx.parsed.y ?? ctx.parsed;
        return typeof v === 'number'
          ? ' ' + v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
          : String(v);
      }
    }}
  },
  scales: {
    x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: '#2a2a2a' } },
    y: { ticks: { color: '#94a3b8', font: { size: 11 },
          callback: v => v.toLocaleString('pt-BR', { notation: 'compact', style: 'currency', currency: 'BRL' }) },
         grid: { color: '#2a2a2a' } },
  },
};

function renderizarGraficoComposicao() {
  destruirGrafico('composicao');
  const posicoes = estado.posicoes;
  const cotacoes = estado.cotacoes;

  const tickers = Object.keys(posicoes)
    .filter(t => posicoes[t].qt > 0)
    .sort((a, b) => posicoes[b].custoTotal - posicoes[a].custoTotal);

  const custos  = tickers.map(t => posicoes[t].custoTotal);
  const mercado = tickers.map(t => {
    const cot = cotacoes[t];
    return cot ? cot.preco * posicoes[t].qt : posicoes[t].custoTotal;
  });

  estado.charts.composicao = new Chart(
    document.getElementById('grafico-composicao'),
    {
      type: 'bar',
      data: {
        labels: tickers,
        datasets: [
          { label: 'Custo',   data: custos,  backgroundColor: '#3b82f688', borderColor: '#3b82f6', borderWidth: 1 },
          { label: 'Mercado', data: mercado, backgroundColor: '#22c55e88', borderColor: '#22c55e', borderWidth: 1 },
        ],
      },
      options: { ...optBase, plugins: { ...optBase.plugins, legend: { ...optBase.plugins.legend, display: true } } },
    }
  );
}

function renderizarGraficoProventos(posicoes) {
  destruirGrafico('proventos');

  // Agrupa proventos por ano via fluxos positivos (proventos)
  const porAno = {};
  Object.values(posicoes).forEach(pos => {
    pos.fluxos.forEach(f => {
      if (f.valor > 0 && f.data) {
        const ano = f.data.getFullYear();
        porAno[ano] = (porAno[ano] || 0) + f.valor;
      }
    });
  });

  const anos   = Object.keys(porAno).sort();
  const valores = anos.map(a => porAno[a]);

  if (!anos.length) return;

  estado.charts.proventos = new Chart(
    document.getElementById('grafico-proventos'),
    {
      type: 'bar',
      data: {
        labels: anos,
        datasets: [{
          label: 'Proventos',
          data: valores,
          backgroundColor: PALETA[0] + 'aa',
          borderColor: PALETA[0],
          borderWidth: 1,
          borderRadius: 4,
        }],
      },
      options: { ...optBase, plugins: { ...optBase.plugins, legend: { display: false } } },
    }
  );
}

function renderizarGraficoEvolucao(historico) {
  destruirGrafico('evolucao');
  if (!historico.length) return;

  const labels = historico.map(h =>
    h.data.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' })
  );
  const valores = historico.map(h => h.capitalInvestido);

  estado.charts.evolucao = new Chart(
    document.getElementById('grafico-evolucao'),
    {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: 'Capital Investido Acumulado',
          data: valores,
          borderColor: '#3b82f6',
          backgroundColor: '#3b82f611',
          fill: true,
          tension: 0.3,
          pointRadius: 3,
          pointBackgroundColor: '#3b82f6',
        }],
      },
      options: {
        ...optBase,
        plugins: { ...optBase.plugins, legend: { display: false } },
      },
    }
  );
}

/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */

async function inicializarDashboard() {
  mostrarLoading(true);
  try {
    const rows = await lerXLSX();
    const { posicoes, historico } = processarMovimentacoes(rows);
    estado.posicoes  = posicoes;
    estado.historico = historico;
    estado.carregado = true;

    renderizarTabelaInicial(posicoes);
    renderizarCardsIniciais(posicoes);
    renderizarGraficoProventos(posicoes);
    renderizarGraficoEvolucao(historico);
    renderizarGraficoComposicao(); // sem cotações ainda: usa custo == mercado

    // Busca cotações logo após renderizar a estrutura
    await atualizarCotacoes();
  } catch (err) {
    console.error('Erro ao inicializar dashboard:', err);
    document.getElementById('corpo-tabela').innerHTML =
      `<tr><td colspan="9"><div class="aviso">
        ⚠️ Não foi possível carregar o arquivo <strong>movimentacao.xlsx</strong>.<br>
        Certifique-se de que o arquivo está na mesma pasta do index.html.<br>
        <small>${err.message}</small>
      </div></td></tr>`;
  } finally {
    mostrarLoading(false);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  if (verificarSessao()) {
    mostrarDashboard();
  }
});
