const DEFAULT_API_URL = localStorage.getItem('finance_api_url') || '';
const $ = (id) => document.getElementById(id);

const state = {
  apiUrl: DEFAULT_API_URL,
  despesas: [],
  ajustes: [],
  consolidado: [],
  resumo: null,
};

document.addEventListener('DOMContentLoaded', () => {
  $('apiUrl').value = state.apiUrl;
  bindEvents();

  if (state.apiUrl) {
    refreshAll();
  } else {
    setStatus('Cole a URL do Apps Script e clique em Guardar URL.');
  }
});

function bindEvents() {
  $('saveApiBtn').addEventListener('click', saveApiUrl);
  $('refreshBtn').addEventListener('click', refreshAll);
  $('reloadDespesasBtn').addEventListener('click', loadDespesas);
  $('reloadAjustesBtn').addEventListener('click', loadAjustes);
  $('reloadConsolidadoBtn').addEventListener('click', loadConsolidado);
  $('despesaForm').addEventListener('submit', submitDespesa);
  $('ajusteForm').addEventListener('submit', submitAjuste);
  $('editAjusteForm').addEventListener('submit', submitEditAjuste);
  $('editDespesaForm').addEventListener('submit', submitEditDespesa);
  $('correcaoReceitaForm').addEventListener('submit', submitCorrecaoReceita);
  $('closeDialogBtn').addEventListener('click', () => $('editDialog').close());
  $('closeEditDespesaBtn').addEventListener('click', () => $('editDespesaDialog').close());
}

function saveApiUrl() {
  const url = $('apiUrl').value.trim();
  if (!url) {
    alert('Cole a URL do Apps Script.');
    return;
  }
  state.apiUrl = url;
  localStorage.setItem('finance_api_url', url);
  setStatus('URL guardada.');
  refreshAll();
}

async function refreshAll() {
  if (!ensureApi()) return;
  setStatus('A carregar dados...');

  // Cada chamada independente — uma falha não bloqueia as outras
  const results = await Promise.allSettled([
    loadResumo(),
    loadDespesas(),
    loadAjustes(),
    loadConsolidado(),
  ]);

  const failed = results.filter(r => r.status === 'rejected');
  if (failed.length === 0) {
    setStatus('Dados actualizados com sucesso.');
  } else {
    const msgs = failed.map(r => r.reason?.message || 'Erro desconhecido').join('; ');
    setStatus('Alguns dados falharam: ' + msgs);
    console.error('Falhas no refreshAll:', failed);
  }
}

function ensureApi() {
  if (!state.apiUrl) {
    alert('Defina primeiro a URL do Apps Script.');
    return false;
  }
  return true;
}

function buildUrl(action) {
  const sep = state.apiUrl.includes('?') ? '&' : '?';
  return `${state.apiUrl}${sep}action=${encodeURIComponent(action)}`;
}

async function apiGet(action) {
  const url = buildUrl(action) + '&_t=' + Date.now();
  const response = await fetch(url, { cache: 'no-store' });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || 'Erro na API.');
  return data;
}

async function apiPost(action, payload) {
  const response = await fetch(buildUrl(action), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
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
  renderDespesas();
  renderRelatorio();
}

async function loadAjustes() {
  const data = await apiGet('ajustes_receita');
  state.ajustes = data.items || [];
  renderAjustes();
}

function renderRelatorio() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) return;

  // ENTRADAS
  const entradas = [
    { desc: 'Inscrições (' + formatNumber(r.totalInscritos) + ' membros)', valor: r.receitaInscricoes },
    { desc: 'Camisetas', valor: r.receitaCamisetas },
  ];
  const totalEntradas = r.receitaTotal;

  $('relatorioEntradasBody').innerHTML = entradas.map(e => `
    <tr>
      <td>${escapeHtml(e.desc)}</td>
      <td class="num">${formatMoney(e.valor)}</td>
    </tr>
  `).join('');
  $('relatorioTotalEntradas').textContent = formatMoney(totalEntradas);

  // SAÍDAS
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

  // SALDO
  $('relatorioSaldoFinal').textContent = formatMoney(r.saldoLiquido);
}

function exportarExcel() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const wb = XLSX.utils.book_new();

  // Folha ENTRADAS
  const entradas = [
    ['ENTRADAS', ''],
    ['Descrição', 'Valor (MT)'],
    ['Inscrições (' + r.totalInscritos + ' membros)', r.receitaInscricoes],
    ['Camisetas', r.receitaCamisetas],
    ['', ''],
    ['TOTAL ENTRADAS', r.receitaTotal],
  ];
  const wsEntradas = XLSX.utils.aoa_to_sheet(entradas);
  wsEntradas['!cols'] = [{ wch: 40 }, { wch: 20 }];

  // Folha SAÍDAS
  const saidas = [['SAÍDAS', '', '', ''], ['Data', 'Categoria', 'Descrição', 'Valor (MT)']];
  despesas.forEach(d => saidas.push([d.data, d.categoria, d.descricao, d.valorTotal]));
  saidas.push(['', '', '', '']);
  saidas.push(['', '', 'TOTAL SAÍDAS', r.despesasTotais]);
  saidas.push(['', '', '', '']);
  saidas.push(['', '', 'SALDO LÍQUIDO', r.saldoLiquido]);
  const wsSaidas = XLSX.utils.aoa_to_sheet(saidas);
  wsSaidas['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 40 }, { wch: 20 }];

  XLSX.utils.book_append_sheet(wb, wsEntradas, 'Entradas');
  XLSX.utils.book_append_sheet(wb, wsSaidas, 'Saidas');

  XLSX.writeFile(wb, 'relatorio_consolidado.xlsx');
}

async function exportarWord() {
  const r = state.resumo;
  const despesas = state.despesas;
  if (!r) { alert('Carregue os dados primeiro.'); return; }

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          AlignmentType, WidthType, BorderStyle, ShadingType, HeadingLevel } = docx;

  const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

  const headerCell = (text, bg = '163B63') => new TableCell({
    borders, margins: cellMargins,
    shading: { fill: bg, type: ShadingType.CLEAR },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 22 })] })],
  });

  const dataCell = (text, right = false) => new TableCell({
    borders, margins: cellMargins,
    children: [new Paragraph({
      alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text || ''), size: 20 })],
    })],
  });

  const totalCell = (text, right = false) => new TableCell({
    borders, margins: cellMargins,
    shading: { fill: 'EEF4FB', type: ShadingType.CLEAR },
    children: [new Paragraph({
      alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text || ''), bold: true, size: 22 })],
    })],
  });

  const fmt = (v) => Number(v || 0).toLocaleString('pt-MZ', { minimumFractionDigits: 2 }) + ' MT';

  // Tabela Entradas
  const tblEntradas = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [6000, 3026],
    rows: [
      new TableRow({ children: [headerCell('Descrição'), headerCell('Valor (MT)')] }),
      new TableRow({ children: [dataCell('Inscrições (' + r.totalInscritos + ' membros)'), dataCell(fmt(r.receitaInscricoes), true)] }),
      new TableRow({ children: [dataCell('Camisetas'), dataCell(fmt(r.receitaCamisetas), true)] }),
      new TableRow({ children: [totalCell('TOTAL ENTRADAS'), totalCell(fmt(r.receitaTotal), true)] }),
    ],
  });

  // Tabela Saídas
  const saidasRows = [
    new TableRow({ children: [headerCell('Data'), headerCell('Categoria'), headerCell('Descrição'), headerCell('Valor (MT)')] }),
    ...despesas.map(d => new TableRow({ children: [
      dataCell(d.data), dataCell(d.categoria), dataCell(d.descricao), dataCell(fmt(d.valorTotal), true),
    ]})),
    new TableRow({ children: [
      totalCell(''), totalCell(''), totalCell('TOTAL SAÍDAS'), totalCell(fmt(r.despesasTotais), true),
    ]}),
  ];

  const tblSaidas = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1500, 2000, 3526, 2000],
    rows: saidasRows,
  });

  // Saldo final
  const tblSaldo = new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [6000, 3026],
    rows: [new TableRow({ children: [
      new TableCell({
        borders, margins: cellMargins,
        shading: { fill: '163B63', type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: 'SALDO LÍQUIDO', bold: true, color: 'FFFFFF', size: 26 })] })],
      }),
      new TableCell({
        borders, margins: cellMargins,
        shading: { fill: '163B63', type: ShadingType.CLEAR },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: fmt(r.saldoLiquido), bold: true, color: 'FFFFFF', size: 26 })],
        })],
      }),
    ]}),
  ]});

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'RELATÓRIO CONSOLIDADO', bold: true, size: 32, color: '163B63' })] }),
        new Paragraph({ children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: 'ENTRADAS', bold: true, size: 24, color: '166534' })] }),
        new Paragraph({ children: [new TextRun('')] }),
        tblEntradas,
        new Paragraph({ children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: 'SAÍDAS', bold: true, size: 24, color: '991B1B' })] }),
        new Paragraph({ children: [new TextRun('')] }),
        tblSaidas,
        new Paragraph({ children: [new TextRun('')] }),
        tblSaldo,
      ],
    }],
  });

  const buffer = await Packer.toBlob(doc);
  const url = URL.createObjectURL(buffer);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'relatorio_consolidado.docx';
  a.click();
  URL.revokeObjectURL(url);
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

function renderDespesas() {
  const body = $('despesasBody');
  if (!state.despesas.length) {
    body.innerHTML = '<tr><td colspan="7" class="empty">Sem despesas lançadas.</td></tr>';
    return;
  }

  body.innerHTML = state.despesas
    .slice()
    .reverse()
    .map((item) => {
      const serialized = encodeURIComponent(JSON.stringify(item));
      return `
      <tr>
        <td>${escapeHtml(item.data || '')}</td>
        <td>${escapeHtml(item.categoria || '')}</td>
        <td>${escapeHtml(item.fornecedor || '')}</td>
        <td>${escapeHtml(item.descricao || '')}</td>
        <td>${escapeHtml(String(item.quantidade || ''))}</td>
        <td class="num">${formatMoney(item.valorTotal)}</td>
        <td><button class="btn" data-item="${serialized}" onclick="openEditDespesaDialog(this)">Editar</button></td>
      </tr>`;
    }).join('');
}

function renderAjustes() {
  const body = $('ajustesBody');
  if (!state.ajustes.length) {
    body.innerHTML = '<tr><td colspan="7" class="empty">Sem ajustes lançados.</td></tr>';
    return;
  }

  body.innerHTML = state.ajustes
    .slice()
    .reverse()
    .map((item) => {
      const serialized = encodeURIComponent(JSON.stringify(item));
      const tagClass = normalize(item.operacao) === 'adicionar' ? 'tag-plus' : 'tag-minus';
      return `
      <tr>
        <td>${escapeHtml(item.data || '')}</td>
        <td>${escapeHtml(item.nome || '')}</td>
        <td>${escapeHtml(item.tipo || '')}</td>
        <td><span class="tag ${tagClass}">${escapeHtml(item.operacao || '')}</span></td>
        <td>${escapeHtml(item.descricao || '')}</td>
        <td class="num">${formatMoney(item.valor)}</td>
        <td><button class="btn" data-item="${serialized}" onclick="openEditDialogFromButton(this)">Editar</button></td>
      </tr>`;
    }).join('');
}

function confirmarDeleteDespesa() {
  const rowNumber = Number($('editDespesaForm').elements.rowNumber.value);
  const descricao = $('editDespesaForm').elements.descricao.value;
  if (!confirm(`Tem a certeza que quer eliminar a despesa "${descricao}"?\nEsta acção não pode ser desfeita.`)) return;
  submitDeleteDespesa(rowNumber);
}

async function submitDeleteDespesa(rowNumber) {
  if (!ensureApi()) return;
  try {
    setStatus('A eliminar despesa...');
    await apiPost('deleteDespesa', { rowNumber });
    $('editDespesaDialog').close();
    await Promise.all([loadResumo(), loadDespesas()]);
    setStatus('Despesa eliminada com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao eliminar despesa.');
    setStatus('Falha ao eliminar despesa.');
  }
}

function openEditDespesaDialog(button) {
  const item = JSON.parse(decodeURIComponent(button.dataset.item));
  const form = $('editDespesaForm');
  form.elements.rowNumber.value = item.rowNumber;
  form.elements.data.value = item.data || '';
  form.elements.categoria.value = item.categoria || '';
  form.elements.fornecedor.value = item.fornecedor || '';
  form.elements.quantidade.value = item.quantidade || '';
  form.elements.descricao.value = item.descricao || '';
  form.elements.valorTotal.value = item.valorTotal || '';
  $('editDespesaDialog').showModal();
}

async function submitEditDespesa(event) {
  event.preventDefault();
  if (!ensureApi()) return;
  const form = event.currentTarget;
  const payload = Object.fromEntries(new FormData(form).entries());
  try {
    setStatus('A corrigir despesa...');
    await apiPost('updateDespesa', payload);
    $('editDespesaDialog').close();
    await Promise.all([loadResumo(), loadDespesas()]);
    setStatus('Despesa corrigida com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao corrigir despesa.');
    setStatus('Falha ao corrigir despesa.');
  }
}

function openEditDialogFromButton(button) {
  const item = JSON.parse(decodeURIComponent(button.dataset.item));
  const form = $('editAjusteForm');

  form.elements.id.value = item.id || '';
  form.elements.data.value = item.data || '';
  form.elements.nome.value = item.nome || '';
  form.elements.tipo.value = item.tipo || 'Inscrição';
  form.elements.operacao.value = item.operacao || 'Adicionar';
  form.elements.referencia.value = item.referencia || '';
  form.elements.descricao.value = item.descricao || '';
  form.elements.valor.value = item.valor || '';
  form.elements.observacoes.value = item.observacoes || '';

  $('editDialog').showModal();
}

async function submitDespesa(event) {
  event.preventDefault();
  if (!ensureApi()) return;

  const form = event.currentTarget;
  const payload = Object.fromEntries(new FormData(form).entries());

  try {
    setStatus('A guardar despesa...');
    await apiPost('addDespesa', payload);
    form.reset();
    await Promise.all([loadResumo(), loadDespesas()]);
    setStatus('Despesa guardada com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao guardar despesa.');
    setStatus('Falha ao guardar despesa.');
  }
}

async function submitAjuste(event) {
  event.preventDefault();
  if (!ensureApi()) return;

  const form = event.currentTarget;
  const payload = Object.fromEntries(new FormData(form).entries());

  try {
    setStatus('A guardar ajuste...');
    await apiPost('addAjusteReceita', payload);
    form.reset();
    await Promise.all([loadResumo(), loadAjustes()]);
    setStatus('Ajuste guardado com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao guardar ajuste.');
    setStatus('Falha ao guardar ajuste.');
  }
}

async function loadConsolidado() {
  const data = await apiGet('consolidado');
  state.consolidado = data.items || [];
  renderConsolidado();
}

function renderConsolidado() {
  const body = $('consolidadoBody');
  const items = state.consolidado.filter(i => String(i.nome || '').trim() !== '');
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="4" class="empty">Sem dados no Consolidado.</td></tr>';
    return;
  }
  body.innerHTML = items.map(item => `
    <tr>
      <td>${escapeHtml(item.nome)}</td>
      <td class="num">${formatMoney(item.valorInscricao)}</td>
      <td class="num">${formatMoney(item.valorCamiseta)}</td>
      <td><button class="btn" onclick="abrirCorrecaoReceita(${item.rowNumber})">Corrigir</button></td>
    </tr>
  `).join('');
}

function abrirCorrecaoReceita(rowNumber) {
  const item = state.consolidado.find(i => i.rowNumber === rowNumber);
  if (!item) return;
  const form = $('correcaoReceitaForm');
  form.elements.rowNumber.value = item.rowNumber;
  form.elements.nome.value = item.nome;
  form.elements.valorInscricaoActual.value = formatMoney(item.valorInscricao);
  form.elements.novoValorInscricao.value = item.valorInscricao || '';
  form.elements.valorCamisetaActual.value = formatMoney(item.valorCamiseta);
  form.elements.novoValorCamiseta.value = item.valorCamiseta || '';
  $('seccaoCorrecaoReceita').style.display = '';
  $('seccaoCorrecaoReceita').scrollIntoView({ behavior: 'smooth' });
}

function fecharCorrecaoReceita() {
  $('seccaoCorrecaoReceita').style.display = 'none';
  $('correcaoReceitaForm').reset();
}

async function submitCorrecaoReceita(event) {
  event.preventDefault();
  if (!ensureApi()) return;
  const form = event.currentTarget;
  const payload = {
    rowNumber: Number(form.elements.rowNumber.value),
    novoValorInscricao: form.elements.novoValorInscricao.value,
    novoValorCamiseta: form.elements.novoValorCamiseta.value,
  };
  try {
    setStatus('A guardar correcção...');
    await apiPost('updateConsolidadoReceita', payload);
    fecharCorrecaoReceita();
    await Promise.all([loadResumo(), loadConsolidado()]);
    setStatus('Correcção guardada com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao guardar correcção.');
    setStatus('Falha ao guardar correcção.');
  }
}

function confirmarDeleteAjuste() {
  const descricao = $('editAjusteForm').elements.descricao.value;
  if (!confirm(`Tem a certeza que quer eliminar o ajuste "${descricao}"?\nEsta acção não pode ser desfeita.`)) return;
  const id = $('editAjusteForm').elements.id.value;
  submitDeleteAjuste(id);
}

async function submitDeleteAjuste(id) {
  if (!ensureApi()) return;
  try {
    setStatus('A eliminar ajuste...');
    await apiPost('deleteAjuste', { id });
    $('editDialog').close();
    await Promise.all([loadResumo(), loadAjustes()]);
    setStatus('Ajuste eliminado com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao eliminar ajuste.');
    setStatus('Falha ao eliminar ajuste.');
  }
}

async function submitEditAjuste(event) {
  event.preventDefault();
  if (!ensureApi()) return;

  const form = event.currentTarget;
  const payload = Object.fromEntries(new FormData(form).entries());

  try {
    setStatus('A actualizar ajuste...');
    await apiPost('updateAjusteReceita', payload);
    $('editDialog').close();
    await Promise.all([loadResumo(), loadAjustes()]);
    setStatus('Ajuste actualizado com sucesso.');
  } catch (error) {
    console.error(error);
    alert(error.message || 'Erro ao actualizar ajuste.');
    setStatus('Falha ao actualizar ajuste.');
  }
}

function formatMoney(value) {
  const n = Number(value || 0);
  return `${n.toLocaleString('pt-MZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} MT`;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('pt-MZ');
}

function setStatus(text) {
  $('statusText').textContent = text;
}

function normalize(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
