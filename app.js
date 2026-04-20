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
  $('correcaoReceitaForm').addEventListener('submit', submitCorrecaoReceita);
  $('closeDialogBtn').addEventListener('click', () => $('editDialog').close());
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
  try {
    await Promise.all([loadResumo(), loadDespesas(), loadAjustes(), loadConsolidado()]);
    setStatus('Dados actualizados com sucesso.');
  } catch (error) {
    console.error(error);
    setStatus('Falha ao actualizar dados.');
    alert(error.message || 'Erro ao carregar dados.');
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
  const response = await fetch(buildUrl(action));
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
}

async function loadDespesas() {
  const data = await apiGet('despesas');
  state.despesas = data.items || [];
  renderDespesas();
}

async function loadAjustes() {
  const data = await apiGet('ajustes_receita');
  state.ajustes = data.items || [];
  renderAjustes();
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
    body.innerHTML = '<tr><td colspan="6" class="empty">Sem despesas lançadas.</td></tr>';
    return;
  }

  body.innerHTML = state.despesas
    .slice()
    .reverse()
    .map((item) => `
      <tr>
        <td>${escapeHtml(item.data || '')}</td>
        <td>${escapeHtml(item.categoria || '')}</td>
        <td>${escapeHtml(item.fornecedor || '')}</td>
        <td>${escapeHtml(item.descricao || '')}</td>
        <td>${escapeHtml(String(item.quantidade || ''))}</td>
        <td class="num">${formatMoney(item.valorTotal)}</td>
      </tr>
    `).join('');
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
