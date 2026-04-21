const DEFAULT_API_URL = localStorage.getItem('finance_api_url') || '';
const $ = (id) => document.getElementById(id);

const state = {
  apiUrl: DEFAULT_API_URL,
  despesas: [],
  resumo: null,
};

document.addEventListener('DOMContentLoaded', () => {
  $('apiUrl').value = state.apiUrl;
  $('saveApiBtn').addEventListener('click', saveApiUrl);
  $('refreshBtn').addEventListener('click', refreshAll);

  if (state.apiUrl) {
    refreshAll();
  } else {
    setStatus('Cole a URL do Apps Script e clique em Guardar URL.');
  }
});

function saveApiUrl() {
  const url = $('apiUrl').value.trim();
  if (!url) { alert('Cole a URL do Apps Script.'); return; }
  state.apiUrl = url;
  localStorage.setItem('finance_api_url', url);
  setStatus('URL guardada.');
  refreshAll();
}

async function refreshAll() {
  if (!state.apiUrl) { alert('Defina primeiro a URL do Apps Script.'); return; }
  setStatus('A carregar dados...');

  const results = await Promise.allSettled([loadResumo(), loadDespesas()]);
  const failed = results.filter(r => r.status === 'rejected');
  if (failed.length === 0) {
    setStatus('Dados actualizados com sucesso.');
  } else {
    const msgs = failed.map(r => r.reason?.message || 'Erro desconhecido').join('; ');
    setStatus('Alguns dados falharam: ' + msgs);
    console.error('Falhas:', failed);
  }
}

function buildUrl(action) {
  const sep = state.apiUrl.includes('?') ? '&' : '?';
  return `${state.apiUrl}${sep}action=${encodeURIComponent(action)}&_t=${Date.now()}`;
}

async function apiGet(action) {
  const response = await fetch(buildUrl(action), { cache: 'no-store' });
  const data = await response.json();
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
  $('kpiInscritos').textContent = formatNumber(r.totalInscritos);
  $('kpiInscricoes').textContent = formatMoney(r.receitaInscricoes);
  $('kpiCamisetas').textContent = formatMoney(r.receitaCamisetas);
  $('kpiTotal').textContent = formatMoney(r.receitaTotal);
  $('kpiDespesas').textContent = formatMoney(r.despesasTotais);
  $('kpiSaldo').textContent = formatMoney(r.saldoLiquido);
}

function renderRelatorio() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) return;

  $('relatorioEntradasBody').innerHTML = `
    <tr><td>Inscrições (${formatNumber(r.totalInscritos)} membros)</td><td class="num">${formatMoney(r.receitaInscricoes)}</td></tr>
    <tr><td>Camisetas</td><td class="num">${formatMoney(r.receitaCamisetas)}</td></tr>
  `;
  $('relatorioTotalEntradas').textContent = formatMoney(r.receitaTotal);

  if (!despesas.length) {
    $('relatorioSaidasBody').innerHTML = '<tr><td colspan="4" class="empty">Sem saídas.</td></tr>';
    $('relatorioTotalSaidas').textContent = formatMoney(0);
  } else {
    $('relatorioSaidasBody').innerHTML = despesas.map(d => `
      <tr>
        <td>${escapeHtml(d.data || '')}</td>
        <td>${escapeHtml(d.categoria || '')}</td>
        <td>${escapeHtml(d.descricao || '')}</td>
        <td class="num">${formatMoney(d.valorTotal)}</td>
      </tr>
    `).join('');
    $('relatorioTotalSaidas').textContent = formatMoney(r.despesasTotais);
  }

  $('relatorioSaldoFinal').textContent = formatMoney(r.saldoLiquido);
}

/* ── EXPORTAR EXCEL ── */
function exportarExcel() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const wb = XLSX.utils.book_new();

  const entradas = [
    ['ENTRADAS', ''],
    ['Descrição', 'Valor (MT)'],
    ['Inscrições (' + r.totalInscritos + ' membros)', r.receitaInscricoes],
    ['Camisetas', r.receitaCamisetas],
    ['', ''],
    ['TOTAL ENTRADAS', r.receitaTotal],
  ];
  const wsE = XLSX.utils.aoa_to_sheet(entradas);
  wsE['!cols'] = [{ wch: 40 }, { wch: 20 }];

  const saidas = [['SAÍDAS', '', '', ''], ['Data', 'Categoria', 'Descrição', 'Valor (MT)']];
  despesas.forEach(d => saidas.push([d.data, d.categoria, d.descricao, d.valorTotal]));
  saidas.push(['', '', '', '']);
  saidas.push(['', '', 'TOTAL SAÍDAS', r.despesasTotais]);
  saidas.push(['', '', '', '']);
  saidas.push(['', '', 'SALDO LÍQUIDO', r.saldoLiquido]);
  const wsS = XLSX.utils.aoa_to_sheet(saidas);
  wsS['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 40 }, { wch: 20 }];

  XLSX.utils.book_append_sheet(wb, wsE, 'Entradas');
  XLSX.utils.book_append_sheet(wb, wsS, 'Saidas');
  XLSX.writeFile(wb, 'relatorio_consolidado.xlsx');
}

/* ── EXPORTAR WORD ── */
function exportarWord() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const fmt = v => Number(v || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2 }) + ' MT';

  const linhasSaidas = despesas.map(d => `
    <tr>
      <td>${escapeHtml(d.data || '')}</td>
      <td>${escapeHtml(d.categoria || '')}</td>
      <td>${escapeHtml(d.descricao || '')}</td>
      <td style="text-align:right">${fmt(d.valorTotal)}</td>
    </tr>`).join('');

  const html = `
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8"/>
  <title>Relatório Consolidado</title>
  <!--[if gte mso 9]>
  <xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml>
  <![endif]-->
  <style>
    body { font-family: Arial, sans-serif; font-size: 11pt; color: #1b2430; margin: 2cm; }
    h1 { font-size: 16pt; color: #163B63; text-align: center; text-transform: uppercase; margin-bottom: 20pt; }
    h2 { font-size: 12pt; margin: 14pt 0 6pt; }
    h2.entrada { color: #166534; }
    h2.saida   { color: #991B1B; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 14pt; }
    th { background-color: #163B63; color: #ffffff; padding: 6pt 8pt; text-align: left; font-size: 10pt; }
    td { padding: 5pt 8pt; border-bottom: 1pt solid #e7edf5; font-size: 10pt; vertical-align: top; }
    tr.total td { background-color: #EEF4FB; font-weight: bold; font-size: 11pt; }
    tr.saldo td { background-color: #163B63; color: #ffffff; font-weight: bold; font-size: 13pt; }
    .right { text-align: right; }
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
function formatMoney(value) {
  return Number(value || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' MT';
}
function formatNumber(value) {
  return Number(value || 0).toLocaleString('pt-MZ');
}
function setStatus(text) {
  $('statusText').textContent = text;
}
function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}
