// ── URL da API (lida do parâmetro ?api= ou do localStorage) ──
const _paramApi = new URLSearchParams(window.location.search).get('api');
if (_paramApi) {
  localStorage.setItem('_fapi', _paramApi);
  // Remove o parâmetro da barra do browser sem recarregar a página
  window.history.replaceState({}, '', window.location.pathname);
}
const API_URL = localStorage.getItem('_fapi') || '';

const $ = (id) => document.getElementById(id);

const state = {
  despesas: [],
  resumo: null,
};

document.addEventListener('DOMContentLoaded', () => {
  $('refreshBtn').addEventListener('click', refreshAll);

  if (API_URL) {
    refreshAll();
  } else {
    // Primeira configuração: pede a URL uma única vez e guarda
    const url = prompt('Configuração inicial: cole a URL do Apps Script');
    if (url && url.trim()) {
      localStorage.setItem('_fapi', url.trim());
      location.reload();
    } else {
      setStatus('URL não configurada.');
    }
  }
});

async function refreshAll() {
  if (!API_URL) return;
  setStatus('A carregar dados...');
  const results = await Promise.allSettled([loadResumo(), loadDespesas()]);
  const failed = results.filter(r => r.status === 'rejected');
  setStatus(failed.length === 0 ? 'Dados actualizados com sucesso.' : 'Falha: ' + failed.map(r => r.reason?.message).join('; '));
}

function buildUrl(action) {
  const sep = API_URL.includes('?') ? '&' : '?';
  return `${API_URL}${sep}action=${encodeURIComponent(action)}&_t=${Date.now()}`;
}

async function apiGet(action) {
  const res = await fetch(buildUrl(action), { cache: 'no-store' });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || 'Erro na API.');
  return data;
}

async function loadResumo() {
  const data = await apiGet('resumo');
  state.resumo = data;
  renderResumo();
  renderRelatorio();
}

async function loadDespesas() {
  const data = await apiGet('despesas');
  state.despesas = data.items || [];
  renderRelatorio();
}

function renderResumo() {
  const r = state.resumo;
  if (!r) return;
  $('kpiInscritos').textContent  = formatNumber(r.totalInscritos);
  $('kpiInscricoes').textContent = formatMoney(r.receitaInscricoes);
  $('kpiCamisetas').textContent  = formatMoney(r.receitaCamisetas);
  $('kpiTotal').textContent      = formatMoney(r.receitaTotal);
  $('kpiDespesas').textContent   = formatMoney(r.despesasTotais);
  $('kpiSaldo').textContent      = formatMoney(r.saldoLiquido);
}

function renderRelatorio() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) return;

  $('relatorioEntradasBody').innerHTML = `
    <tr><td>Inscrições (${formatNumber(r.totalInscritos)} membros)</td><td class="num">${formatMoney(r.receitaInscricoes)}</td></tr>
    <tr><td>Camisetas</td><td class="num">${formatMoney(r.receitaCamisetas)}</td></tr>`;
  $('relatorioTotalEntradas').textContent = formatMoney(r.receitaTotal);

  $('relatorioSaidasBody').innerHTML = despesas.length
    ? despesas.map(d => `<tr>
        <td>${escapeHtml(d.data||'')}</td>
        <td>${escapeHtml(d.categoria||'')}</td>
        <td>${escapeHtml(d.descricao||'')}</td>
        <td class="num">${formatMoney(d.valorTotal)}</td>
      </tr>`).join('')
    : '<tr><td colspan="4" class="empty">Sem saídas.</td></tr>';

  $('relatorioTotalSaidas').textContent = formatMoney(r.despesasTotais);
  $('relatorioSaldoFinal').textContent  = formatMoney(r.saldoLiquido);
}

/* ── EXPORTAR WORD ── */
function exportarWord() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const fmt = v => Number(v || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2 }) + ' MT';

  const linhasSaidas = despesas.map(d => `
    <tr>
      <td>${escapeHtml(d.data||'')}</td>
      <td>${escapeHtml(d.categoria||'')}</td>
      <td>${escapeHtml(d.descricao||'')}</td>
      <td style="text-align:right">${fmt(d.valorTotal)}</td>
    </tr>`).join('');

  const html = `
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8"/>
  <title>Relatório Consolidado</title>
  <style>
    body  { font-family:Arial,sans-serif; font-size:10pt; color:#1b2430; margin:1.5cm; }
    h1    { font-size:14pt; color:#163B63; text-align:center; text-transform:uppercase; margin:0 0 12pt; }
    h2    { font-size:11pt; margin:10pt 0 4pt; }
    h2.entrada { color:#166534; }
    h2.saida   { color:#991B1B; }
    table { border-collapse:collapse; width:100%; margin-bottom:10pt; }
    th    { background-color:#163B63; color:#fff; padding:4pt 6pt; text-align:left; font-size:9pt; }
    td    { padding:3pt 6pt; border-bottom:1pt solid #e7edf5; font-size:9pt; vertical-align:top; }
    tr.total td  { background-color:#EEF4FB; font-weight:bold; font-size:10pt; }
    tr.saldo td  { background-color:#163B63; color:#fff; font-weight:bold; font-size:12pt; padding:6pt 8pt; }
    .right { text-align:right; }
  </style>
</head>
<body>
  <h1>Relatório Consolidado</h1>

  <h2 class="entrada">Entradas</h2>
  <table>
    <thead><tr><th>Descrição</th><th style="text-align:right">Valor (MT)</th></tr></thead>
    <tbody>
      <tr><td>Inscrições (${r.totalInscritos} membros)</td><td class="right">${fmt(r.receitaInscricoes)}</td></tr>
      <tr><td>Camisetas</td><td class="right">${fmt(r.receitaCamisetas)}</td></tr>
    </tbody>
    <tfoot><tr class="total"><td>TOTAL ENTRADAS</td><td class="right">${fmt(r.receitaTotal)}</td></tr></tfoot>
  </table>

  <h2 class="saida">Saídas</h2>
  <table>
    <thead><tr><th>Data</th><th>Categoria</th><th>Descrição</th><th style="text-align:right">Valor (MT)</th></tr></thead>
    <tbody>${linhasSaidas}</tbody>
    <tfoot><tr class="total"><td colspan="3">TOTAL SAÍDAS</td><td class="right">${fmt(r.despesasTotais)}</td></tr></tfoot>
  </table>

  <table>
    <tbody><tr class="saldo"><td>SALDO LÍQUIDO</td><td class="right">${fmt(r.saldoLiquido)}</td></tr></tbody>
  </table>
</body>
</html>`;

  const blob = new Blob(['\ufeff' + html], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'relatorio_consolidado.doc';
  a.click();
  URL.revokeObjectURL(url);
}

/* ── UTILS ── */
function formatMoney(v) {
  return Number(v || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' MT';
}
function formatNumber(v) { return Number(v || 0).toLocaleString('pt-MZ'); }
function setStatus(t) { $('statusText').textContent = t; }
function escapeHtml(v) {
  return String(v ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;')
    .replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
}
