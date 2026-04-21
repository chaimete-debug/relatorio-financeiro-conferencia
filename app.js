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
async function exportarWord() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          AlignmentType, WidthType, BorderStyle, ShadingType, HeadingLevel } = docx;

  const bdr = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
  const mg = { top: 80, bottom: 80, left: 120, right: 120 };

  const hCell = (text, bg = '163B63') => new TableCell({
    borders, margins: mg,
    shading: { fill: bg, type: ShadingType.CLEAR },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 22 })] })],
  });
  const dCell = (text, right = false) => new TableCell({
    borders, margins: mg,
    children: [new Paragraph({
      alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text || ''), size: 20 })],
    })],
  });
  const tCell = (text, right = false) => new TableCell({
    borders, margins: mg,
    shading: { fill: 'EEF4FB', type: ShadingType.CLEAR },
    children: [new Paragraph({
      alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text || ''), bold: true, size: 22 })],
    })],
  });

  const fmt = v => Number(v || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2 }) + ' MT';
  const sp = () => new Paragraph({ children: [new TextRun('')] });

  const tblEntradas = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [6000, 3026],
    rows: [
      new TableRow({ children: [hCell('Descrição'), hCell('Valor (MT)')] }),
      new TableRow({ children: [dCell('Inscrições (' + r.totalInscritos + ' membros)'), dCell(fmt(r.receitaInscricoes), true)] }),
      new TableRow({ children: [dCell('Camisetas'), dCell(fmt(r.receitaCamisetas), true)] }),
      new TableRow({ children: [tCell('TOTAL ENTRADAS'), tCell(fmt(r.receitaTotal), true)] }),
    ],
  });

  const tblSaidas = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1500, 2000, 3526, 2000],
    rows: [
      new TableRow({ children: [hCell('Data'), hCell('Categoria'), hCell('Descrição'), hCell('Valor (MT)')] }),
      ...despesas.map(d => new TableRow({ children: [
        dCell(d.data), dCell(d.categoria), dCell(d.descricao), dCell(fmt(d.valorTotal), true),
      ]})),
      new TableRow({ children: [tCell(''), tCell(''), tCell('TOTAL SAÍDAS'), tCell(fmt(r.despesasTotais), true)] }),
    ],
  });

  const tblSaldo = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [6000, 3026],
    rows: [new TableRow({ children: [
      new TableCell({ borders, margins: mg, shading: { fill: '163B63', type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: 'SALDO LÍQUIDO', bold: true, color: 'FFFFFF', size: 26 })] })] }),
      new TableCell({ borders, margins: mg, shading: { fill: '163B63', type: ShadingType.CLEAR },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: fmt(r.saldoLiquido), bold: true, color: 'FFFFFF', size: 26 })] })] }),
    ]})],
  });

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'RELATÓRIO CONSOLIDADO', bold: true, size: 32, color: '163B63' })] }),
        sp(),
        new Paragraph({ children: [new TextRun({ text: 'ENTRADAS', bold: true, size: 24, color: '166534' })] }),
        sp(), tblEntradas, sp(),
        new Paragraph({ children: [new TextRun({ text: 'SAÍDAS', bold: true, size: 24, color: '991B1B' })] }),
        sp(), tblSaidas, sp(),
        tblSaldo,
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'relatorio_consolidado.docx'; a.click();
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
