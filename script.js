// ============================================================================
// Integração com SharePoint via REST API
// ============================================================================
class SharePointService {
  constructor(siteUrl) {
    this.siteUrl = siteUrl.replace(/\/$/, '');
  }

  encodeEntity(listName) {
    return `SP.Data.${listName.replace(/ /g, '_x0020_').replace(/_/g, '_x005f_')}ListItem`;
  }

  buildUrl(listName, path = '/items') {
    if (!listName) {
      throw new Error('Lista SharePoint não informada.');
    }
    return `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')${path}`;
  }

  sanitizeFileName(fileName) {
    if (typeof fileName !== 'string') {
      return '';
    }
    return fileName.replace(/'/g, "''");
  }

  async request(url, options = {}) {
    const response = await fetch(url, options);
    if (!response.ok) {
      const text = await response.text();
      const error = new Error(text || response.statusText);
      error.status = response.status;
      throw error;
    }
    if (response.status === 204) {
      return null;
    }
    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }

  async addAttachment(listName, itemId, fileName, fileContent, options = {}) {
    if (!listName) {
      throw new Error('Lista SharePoint não informada.');
    }
    if (!itemId) {
      throw new Error('ID do item inválido para anexar arquivo.');
    }

    const { overwrite = false, contentType = 'application/octet-stream' } = options;
    const normalizedFileName = this.sanitizeFileName(fileName || 'resumo.txt');

    if (overwrite) {
      try {
        await this.deleteAttachment(listName, itemId, normalizedFileName);
      } catch (error) {
        if (error?.status !== 404) {
          throw error;
        }
      }
    }

    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      'Content-Type': contentType
    };
    const url = this.buildUrl(
      listName,
      `/items(${itemId})/AttachmentFiles/add(FileName='${normalizedFileName}')`
    );

    const body = fileContent instanceof Blob
      ? fileContent
      : new Blob([fileContent], { type: contentType });

    await this.request(url, { method: 'POST', headers, body });
    return true;
  }

  async deleteAttachment(listName, itemId, fileName) {
    if (!listName) {
      throw new Error('Lista SharePoint não informada.');
    }
    if (!itemId) {
      throw new Error('ID do item inválido para remover anexo.');
    }

    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    };
    const sanitizedName = this.sanitizeFileName(fileName || 'resumo.txt');
    const url = this.buildUrl(
      listName,
      `/items(${itemId})/AttachmentFiles/getByFileName('${sanitizedName}')`
    );
    await this.request(url, { method: 'POST', headers });
    return true;
  }

  async getFormDigest() {
    try {
      const url = `${this.siteUrl}/_api/contextinfo`;
      const headers = {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      };
      const data = await this.request(url, { method: 'POST', headers });
      return data?.d?.GetContextWebInformation?.FormDigestValue;
    } catch (error) {
      if (typeof _spPageContextInfo !== 'undefined') {
        return _spPageContextInfo.formDigestValue;
      }
      throw error;
    }
  }

  async getItems(listName, params = {}) {
    const url = new URL(this.buildUrl(listName));
    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.append(`$${key}`, value);
      }
    });
    const headers = { Accept: 'application/json;odata=verbose' };
    const data = await this.request(url.toString(), { headers });
    return data?.d?.results ?? [];
  }

  async getItem(listName, id) {
    const url = this.buildUrl(listName, `/items(${id})`);
    const headers = { Accept: 'application/json;odata=verbose' };
    const data = await this.request(url, { headers });
    return data?.d ?? null;
  }

  async createItem(listName, payload) {
    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest
    };
    const body = JSON.stringify({
      __metadata: { type: this.encodeEntity(listName) },
      ...payload
    });
    const data = await this.request(this.buildUrl(listName), { method: 'POST', headers, body });
    return data?.d ?? null;
  }

  async updateItem(listName, id, payload) {
    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    };
    const body = JSON.stringify({
      __metadata: { type: this.encodeEntity(listName) },
      ...payload
    });
    await this.request(this.buildUrl(listName, `/items(${id})`), { method: 'POST', headers, body });
    return true;
  }

  async deleteItem(listName, id) {
    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    };
    await this.request(this.buildUrl(listName, `/items(${id})`), { method: 'POST', headers });
    return true;
  }
}

// ============================================================================
// Estado global e referências da interface
// ============================================================================
const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
const DATE_FMT = new Intl.DateTimeFormat('pt-BR');
const BUDGET_THRESHOLD = 1_000_000;
const EXCHANGE_RATE = 5.6; // 1 USD = 5.6 BRL
const DATE_RANGE_ERROR_MESSAGE = 'A data de término não pode ser anterior à data de início.';

const SITE_URL = window.SHAREPOINT_SITE_URL || 'https://arcelormittal.sharepoint.com/sites/controladorialongos/capex';
const sp = new SharePointService("https://arcelormittal.sharepoint.com/sites/controladorialongos/capex");

const state = {
  projects: [],
  selectedProjectId: null,
  currentDetails: null,
  editingSnapshot: {
    simplePeps: new Set(),
    milestones: new Set(),
    activities: new Set(),
    activityPeps: new Set()
  }
};

const validationState = {
  pepBudget: null,
  pepBudgetDetails: null,
  activityDates: null,
  activityDateDetails: null
};

const companyRules = {
  'Empresa 01': {
    centers: ['Centro 01', 'Centro 02'],
    units: ['Unidade 01', 'Unidade 02'],
    locations: ['Local 01', 'Local 02'],
    depreciation: ['CC-01', 'CC-02']
  },
  'Empresa 02': {
    centers: ['Centro 03'],
    units: ['Unidade 03', 'Unidade 04'],
    locations: ['Local 03'],
    depreciation: ['CC-03']
  }
};

function updateProjectState(projectId, changes = {}) {
  if (!projectId) return;
  const index = state.projects.findIndex((item) => Number(item.Id) === Number(projectId));
  if (index === -1) return;
  state.projects[index] = {
    ...state.projects[index],
    ...changes
  };
}

const newProjectBtn = document.getElementById('newProjectBtn');
const projectSearch = document.getElementById('projectSearch');
const projectList = document.getElementById('projectList');
const projectDetails = document.getElementById('projectDetails');
const overlay = document.getElementById('formOverlay');
const projectForm = document.getElementById('projectForm');
const formTitle = document.getElementById('formTitle');
const closeFormBtn = document.getElementById('closeFormBtn');
const saveProjectBtn = document.getElementById('saveProjectBtn');
const submitForApprovalBtn = document.getElementById('submitForApprovalBtn');
const formActions = projectForm?.querySelector('.form-actions') || null;
const formStatus = document.getElementById('formStatus');
const formErrors = document.getElementById('formErrors');
const formErrorsTitle = document.getElementById('formErrorsTitle');
const errorSummaryList = document.getElementById('errorSummaryList');
const statusField = document.getElementById('statusField');
const budgetHint = document.getElementById('budgetHint');
const dateHint = document.getElementById('dateHint');
const simplePepSection = document.getElementById('simplePepSection');
const simplePepList = document.getElementById('simplePepList');
const addSimplePepBtn = document.getElementById('addSimplePepBtn');
const keyProjectSection = document.getElementById('keyProjectSection');
const milestoneList = document.getElementById('milestoneList');
const addMilestoneBtn = document.getElementById('addMilestoneBtn');

const ganttContainer = document.getElementById('ganttContainer');
const ganttTitleEl = document.getElementById('ganttChartTitle');
const ganttChartEl = document.getElementById('ganttChart');

const summaryOverlay = document.getElementById('summaryOverlay');
const summarySections = document.getElementById('summarySections');
const summaryGanttSection = document.getElementById('summaryGanttSection');
const summaryGanttChart = document.getElementById('summaryGanttChart');
const summaryConfirmBtn = document.getElementById('summaryConfirmBtn');
const summaryEditBtn = document.getElementById('summaryEditBtn');

const formSummaryView = document.getElementById('formSummaryView');
const formSummarySections = document.getElementById('formSummarySections');
const formSummaryGanttSection = document.getElementById('formSummaryGanttSection');
const formSummaryGanttChart = document.getElementById('formSummaryGanttChart');
const formSummaryCloseBtn = document.getElementById('formSummaryCloseBtn');

const companySelect = document.getElementById('company');
const centerSelect = document.getElementById('center');
const unitSelect = document.getElementById('unit');
const locationSelect = document.getElementById('location');
const depreciationSelect = document.getElementById('depreciationCostCenter');

const approvalYearInput = document.getElementById('approvalYear');
const projectBudgetInput = document.getElementById('projectBudget');
const investmentLevelSelect = document.getElementById('investmentLevel');
const projectStartDateInput = document.getElementById('startDate');
const projectEndDateInput = document.getElementById('endDate');

const simplePepTemplate = document.getElementById('simplePepTemplate');
const milestoneTemplate = document.getElementById('milestoneTemplate');
const activityTemplate = document.getElementById('activityTemplate');

const READ_ONLY_STATUSES = new Set(['Aprovado', 'Em Aprovação']);
const APPROVAL_ALLOWED_STATUSES = new Set(['Rascunho', 'Reprovado', 'Reprovado para Revisão']);
const defaultSummaryContext = {
  sections: summarySections,
  ganttSection: summaryGanttSection,
  ganttChart: summaryGanttChart
};

const formSummaryContext = {
  sections: formSummarySections,
  ganttSection: formSummaryGanttSection,
  ganttChart: formSummaryGanttChart
};

let activeSummaryContext = defaultSummaryContext;
let currentFormMode = null;

function normalizeStatusKey(status) {
  return typeof status === 'string' ? status.trim() : '';
}

function isReadOnlyStatus(status) {
  return READ_ONLY_STATUSES.has(normalizeStatusKey(status));
}

function canSubmitForApproval(status) {
  const key = normalizeStatusKey(status);
  return !key || APPROVAL_ALLOWED_STATUSES.has(key);
}

function withSummaryContext(context, callback) {
  if (typeof callback !== 'function') return;
  const previousContext = activeSummaryContext;
  activeSummaryContext = context || previousContext;
  try {
    callback();
  } finally {
    activeSummaryContext = previousContext;
  }
}

function clearSummaryContent(context) {
  if (!context) return;
  if (context.sections) {
    context.sections.innerHTML = '';
  }
  if (context.ganttChart) {
    context.ganttChart.innerHTML = '';
  }
  if (context.ganttSection) {
    context.ganttSection.classList.add('hidden');
  }
}

// ============================================================================
// Gantt Chart
// ============================================================================
let ganttLoaderStarted = false;
let ganttReady = false;
let ganttRefreshScheduled = false;
let summaryTriggerButton = null;

function initGantt() {
  if (ganttLoaderStarted) return;
  if (!window.google || !google.charts) return;
  ganttLoaderStarted = true;
  google.charts.load('current', { packages: ['gantt'] });
  google.charts.setOnLoadCallback(() => {
    ganttReady = true;
    refreshGantt();
  });
}

function queueGanttRefresh() {
  if (ganttRefreshScheduled) return;
  ganttRefreshScheduled = true;
  requestAnimationFrame(() => {
    ganttRefreshScheduled = false;
    refreshGantt();
  });
}

function refreshGantt() {
  if (!ganttContainer) return;
  if (keyProjectSection.classList.contains('hidden')) {
    ganttContainer.classList.add('hidden');
    if (ganttChartEl) {
      ganttChartEl.innerHTML = '';
    }
    return;
  }

  const milestones = collectMilestonesForGantt();
  const draw = () => drawGantt(milestones);

  if (ganttReady && window.google?.visualization?.Gantt) {
    draw();
  } else if (ganttLoaderStarted && window.google?.charts) {
    google.charts.setOnLoadCallback(draw);
  } else {
    initGantt();
  }
}

function collectMilestonesForGantt() {
  const milestones = [];
  if (!milestoneList) return milestones;

  milestoneList.querySelectorAll('.milestone').forEach((milestoneEl, index) => {
    const titleInput = milestoneEl.querySelector('.milestone-title');
    const nome = titleInput?.value.trim() || `Marco ${index + 1}`;
    const atividades = [];

    milestoneEl.querySelectorAll('.activity').forEach((activityEl, actIndex) => {
      const title = activityEl.querySelector('.activity-title')?.value.trim() || `Atividade ${actIndex + 1}`;
      const inicio = activityEl.querySelector('.activity-start')?.value || null;
      const fim = activityEl.querySelector('.activity-end')?.value || null;
      const anual = [];
      const ano = parseNumber(activityEl.querySelector('.activity-pep-year')?.value);
      const amountRaw = parseNumericInputValue(activityEl.querySelector('.activity-pep-amount'));
      const amount = Number.isFinite(amountRaw) ? amountRaw : 0;
      const descricao = activityEl.querySelector('.activity-pep-title')?.value.trim() || '';

      if (descricao || ano || amount > 0) {
        anual.push({
          ano,
          capex_brl: amount,
          descricao
        });
      }

      atividades.push({
        titulo: title,
        inicio,
        fim,
        anual
      });
    });

    milestones.push({
      nome,
      atividades
    });
  });

  return milestones;
}

function drawGantt(milestones, options = {}) {
  const {
    container = ganttContainer,
    chartElement = ganttChartEl,
    titleElement = ganttTitleEl,
    emptyMessage = 'Nenhuma atividade para exibir'
  } = options;

  if (!container || !chartElement) return null;
  if (!window.google?.visualization?.DataTable || !window.google?.visualization?.Gantt) {
    return null;
  }

  const parseDate = (value, { isStart } = {}) => {
    if (!value) return null;
    const suffix = isStart ? 'T00:00:00' : 'T23:59:59';
    const date = new Date(`${value}${suffix}`);
    return Number.isNaN(date.getTime()) ? null : date;
  };

  const rows = [];
  let idCounter = 0;
  let minDate = null;
  let maxDate = null;

  const safeMilestones = Array.isArray(milestones) ? milestones : [];

  safeMilestones.forEach((milestone) => {
    if (!milestone) return;

    const activities = Array.isArray(milestone.atividades) ? milestone.atividades : [];
    const validActivities = [];
    let milestoneStart = null;
    let milestoneEnd = null;

    activities.forEach((activity, index) => {
      if (!activity) return;

      const startDate = parseDate(activity.inicio, { isStart: true });
      const endDate = parseDate(activity.fim, { isStart: false });

      if (!startDate || !endDate || endDate < startDate) {
        return;
      }

      if (!milestoneStart || startDate < milestoneStart) milestoneStart = startDate;
      if (!milestoneEnd || endDate > milestoneEnd) milestoneEnd = endDate;

      if (!minDate || startDate < minDate) minDate = startDate;
      if (!maxDate || endDate > maxDate) maxDate = endDate;

      const anualList = Array.isArray(activity.anual) ? activity.anual : [];
      const totalCapex = anualList.reduce((total, year) => total + (Number(year?.capex_brl) || 0), 0);
      const tooltipLines = anualList.map((year) => {
        const yearLabel = year?.ano ?? 'Ano não informado';
        const description = year?.descricao ? ` - ${year.descricao}` : '';
        const value = Number(year?.capex_brl) || 0;
        return `${yearLabel}: ${BRL.format(value)}${description}`;
      });

      const tooltipParts = [`CAPEX total: ${BRL.format(totalCapex)}`];
      if (tooltipLines.length) {
        tooltipParts.push(...tooltipLines);
      }

      validActivities.push({
        index,
        title: activity.titulo || `Atividade ${index + 1}`,
        startDate,
        endDate,
        duration: endDate.getTime() - startDate.getTime(),
        tooltip: tooltipParts.join('<br/>')
      });
    });

    if (!validActivities.length) {
      return;
    }

    idCounter += 1;
    const milestoneId = `ms-${idCounter}`;
    const milestoneName = milestone.nome || `Marco ${idCounter}`;

    if (milestoneStart && milestoneEnd) {
      const milestoneDuration = milestoneEnd.getTime() - milestoneStart.getTime();
      rows.push([
        milestoneId,
        milestoneName,
        'Marco',
        milestoneStart,
        milestoneEnd,
        milestoneDuration,
        0,
        null,
        milestoneName
      ]);

      if (!minDate || milestoneStart < minDate) minDate = milestoneStart;
      if (!maxDate || milestoneEnd > maxDate) maxDate = milestoneEnd;
    }

    validActivities.forEach((activity) => {
      rows.push([
        `${milestoneId}-${activity.index}`,
        activity.title,
        'Atividade',
        activity.startDate,
        activity.endDate,
        activity.duration,
        0,
        milestoneId,
        activity.tooltip
      ]);
    });
  });

  if (!rows.length || !minDate || !maxDate) {
    container.classList.remove('hidden');
    if (titleElement) {
      titleElement.classList.remove('hidden');
    }
    chartElement.innerHTML = `<p class="gantt-empty">${emptyMessage}</p>`;
    return { minDate: null, maxDate: null, rowCount: 0, chart: null };
  }

  chartElement.innerHTML = '';

  const data = new google.visualization.DataTable();
  data.addColumn('string', 'Task ID');
  data.addColumn('string', 'Task Name');
  data.addColumn('string', 'Resource');
  data.addColumn('date', 'Start Date');
  data.addColumn('date', 'End Date');
  data.addColumn('number', 'Duration');
  data.addColumn('number', 'Percent Complete');
  data.addColumn('string', 'Dependencies');
  data.addColumn({ type: 'string', role: 'tooltip', p: { html: true } });
  data.addRows(rows);

  const chart = new google.visualization.Gantt(chartElement);
  const chartOptions = {
    height: Math.max(200, rows.length * 40 + 40),
    tooltip: { isHtml: true },
    gantt: {
      criticalPathEnabled: false,
      arrow: {
        angle: 0,
        width: 0,
        color: '#ffffff',
        radius: 0
      },
      trackHeight: 30,
      palette: [
        { color: '#460a78', dark: '#be2846', light: '#e63c41' },
        { color: '#f58746', dark: '#e63c41', light: '#ffbe6e' }
      ],
      minDate,
      maxDate
    }
  };

  chart.draw(data, chartOptions);

  container.classList.remove('hidden');
  if (titleElement) {
    titleElement.classList.remove('hidden');
  }

  return { minDate, maxDate, rowCount: rows.length, chart };
}

// ============================================================================
// Inicialização
// ============================================================================
function setApprovalYearToCurrent() {
  if (!approvalYearInput) {
    return;
  }
  const currentYear = new Date().getFullYear();
  approvalYearInput.value = currentYear;
  approvalYearInput.max = currentYear;
}

function init() {
  if (investmentLevelSelect) {
    investmentLevelSelect.disabled = true;
    investmentLevelSelect.setAttribute('aria-readonly', 'true');
  }

  bindEvents();
  updateCompanyDependentFields(companySelect?.value || '');
  setApprovalYearToCurrent();
  updateInvestmentLevelField();
  loadProjects();
  initGantt();
  window.addEventListener('load', initGantt, { once: true });
}

function bindEvents() {
  newProjectBtn.addEventListener('click', () => openProjectForm('create'));
  closeFormBtn.addEventListener('click', handleCloseFormRequest);
  // Habilita o fechamento do formulário pela tecla ESC.
  document.addEventListener('keydown', handleOverlayEscape);
  document.addEventListener('input', handleGlobalDateInput);
  if (projectSearch) {
    projectSearch.addEventListener('input', () => renderProjectList());
  }

  projectForm.addEventListener('submit', handleFormSubmit);
  projectForm.addEventListener('focusin', handleFormFocusCapture);
  if (saveProjectBtn) {
    saveProjectBtn.addEventListener('click', (event) => {
      event.preventDefault();
      projectForm.dataset.submitIntent = 'save';
      openSummaryOverlay(saveProjectBtn);
    });
  }

  if (submitForApprovalBtn) {
    submitForApprovalBtn.addEventListener('click', (event) => {
      event.preventDefault();
      projectForm.dataset.submitIntent = 'approval';
      openSummaryOverlay(submitForApprovalBtn);
    });
  }

  if (summaryConfirmBtn) {
    summaryConfirmBtn.addEventListener('click', handleSummaryConfirm);
  }

  if (summaryEditBtn) {
    summaryEditBtn.addEventListener('click', () => closeSummaryOverlay());
  }

  if (formSummaryCloseBtn) {
    formSummaryCloseBtn.addEventListener('click', () => closeForm());
  }

  if (companySelect) {
    companySelect.addEventListener('change', (event) => {
      updateCompanyDependentFields(event.target.value);
    });
  }

  projectBudgetInput.addEventListener('input', () => {
    updateInvestmentLevelField();
    updateBudgetSections();
    validatePepBudget();
  });

  if (projectStartDateInput) {
    const handleProjectStartChange = (event) => {
      validateActivityDates({ changedInput: event.target });
    };
    projectStartDateInput.addEventListener('input', handleProjectStartChange);
    projectStartDateInput.addEventListener('change', handleProjectStartChange);
  }

  if (projectEndDateInput) {
    const handleProjectEndChange = (event) => {
      validateActivityDates({ changedInput: event.target });
    };
    projectEndDateInput.addEventListener('input', handleProjectEndChange);
    projectEndDateInput.addEventListener('change', handleProjectEndChange);
  }

  approvalYearInput.addEventListener('input', () => updateSimplePepYears());

  addSimplePepBtn.addEventListener('click', () => {
    ensureSimplePepRow();
  });

  addMilestoneBtn.addEventListener('click', () => {
    ensureMilestoneBlock();
  });

  simplePepList.addEventListener('click', (event) => {
    if (event.target.classList.contains('remove-row')) {
      const row = event.target.closest('.pep-row');
      row?.remove();
      validatePepBudget();
    }
  });

  simplePepList.addEventListener('input', (event) => {
    if (event.target.classList?.contains('pep-amount')) {
      validatePepBudget({ changedInput: event.target });
    }
  });

  milestoneList.addEventListener('click', (event) => {
    const button = event.target.closest('button');
    if (!button) return;

    if (button.classList.contains('remove-milestone')) {
      button.closest('.milestone')?.remove();
      queueGanttRefresh();
      validatePepBudget();
      validateActivityDates();
      validateAllDateRanges();
      return;
    }
    if (button.classList.contains('add-activity')) {
      const milestone = button.closest('.milestone');
      addActivityBlock(milestone);
      queueGanttRefresh();
      return;
    }
    if (button.classList.contains('remove-activity')) {
      const activity = button.closest('.activity');
      activity?.remove();
      queueGanttRefresh();
      validatePepBudget();
      validateActivityDates();
      validateAllDateRanges();
    }
  });

  const handleMilestoneFormChange = (event) => {
    if (!event.target) return;
    if (event.target.classList?.contains('activity-start')) {
      const activity = event.target.closest('.activity');
      updateActivityPepYear(activity, { force: true });
      validateActivityDates({ changedInput: event.target });
    }
    if (event.target.classList?.contains('activity-end')) {
      validateActivityDates({ changedInput: event.target });
    }
    if (event.target.classList?.contains('activity-pep-amount')) {
      validatePepBudget({ changedInput: event.target });
    }
    queueGanttRefresh();
  };

  milestoneList.addEventListener('input', handleMilestoneFormChange);
  milestoneList.addEventListener('change', handleMilestoneFormChange);
}

// ============================================================================
// Carregamento e renderização da lista de projetos
// ============================================================================
async function loadProjects() {
  try {
    const currentUserId = _spPageContextInfo.userId; // pega o ID do usuário logado
    const results = await sp.getItems('Projects', { 
      orderby: 'Created desc',
      filter: `AuthorId eq ${currentUserId}`
    });
    state.projects = results;
    renderProjectList();
  } catch (error) {
    console.error('Erro ao carregar projetos', error);
  }
}

function renderProjectList() {
  const filter = (projectSearch?.value || '').toLowerCase();
  projectList.innerHTML = '';

  const filtered = state.projects.filter((item) =>
    item.Title?.toLowerCase().includes(filter)
  );

  if (filtered.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'hint';
    empty.textContent = 'Nenhum projeto encontrado.';
    projectList.append(empty);
    return;
  }

  filtered.forEach((item) => {
    const card = document.createElement('article');
    card.className = 'project-card';
    if (state.selectedProjectId === item.Id) {
      card.classList.add('selected');
    }
    card.dataset.id = item.Id;

    const accent = document.createElement('span');
    accent.className = 'project-card-accent';
    accent.style.background = statusColor(item.status);

    const content = document.createElement('div');
    content.className = 'project-card-content';

    const status = document.createElement('span');
    status.className = 'project-card-status';
    status.textContent = item.status || 'Sem status';
    status.style.color = statusColor(item.status);

    const title = document.createElement('span');
    title.className = 'project-card-title';
    title.textContent = item.Title || 'Projeto sem título';
    content.append(status, title);
    if (item.budgetBrl) {
      const budgetRow = document.createElement('div');
      budgetRow.className = 'project-card-bottom';
      const budget = document.createElement('span');
      budget.className = 'project-card-meta';
      budget.textContent = BRL.format(item.budgetBrl);
      budgetRow.append(budget);
      content.append(budgetRow);
    }
    card.append(accent, content);
    card.addEventListener('click', () => selectProject(item.Id));
    projectList.append(card);
  });
}

async function selectProject(projectId) {
  state.selectedProjectId = projectId;
  renderProjectList();
  await loadProjectDetails(projectId);
}

async function loadProjectDetails(projectId) {
  projectDetails.innerHTML = '';
  const loader = document.createElement('p');
  loader.textContent = 'Carregando…';
  loader.className = 'hint';
  projectDetails.append(loader);

  try {
    const project = await sp.getItem('Projects', projectId);
    const [milestones, activities, peps] = await Promise.all([
      sp.getItems('Milestones', { filter: `projectsIdId eq ${projectId}` }),
      sp.getItems('Activities', { filter: `projectsIdId eq ${projectId}` }),
      sp.getItems('Peps', { filter: `projectsIdId eq ${projectId}` })
    ]);

    const detail = {
      project,
      milestones,
      activities,
      peps,
      simplePeps: peps.filter((pep) => !pep.activitiesIdId),
      activityPeps: peps.filter((pep) => pep.activitiesIdId)
    };

    state.currentDetails = detail;
    renderProjectDetails(detail);
  } catch (error) {
    console.error('Erro ao carregar detalhes do projeto', error);
    projectDetails.innerHTML = '';
    const errorBox = document.createElement('p');
    errorBox.className = 'hint';
    errorBox.textContent = 'Não foi possível carregar os dados do projeto.';
    projectDetails.append(errorBox);
  }
}

function renderProjectDetails(detail) {
  projectDetails.innerHTML = '';
  if (!detail?.project) {
    projectDetails.append(createEmptyState());
    return;
  }

  const { project } = detail;

  const wrapper = document.createElement('div');
  wrapper.className = 'project-overview';

  const header = document.createElement('div');
  header.className = 'project-overview__header';
  const title = document.createElement('h2');
  title.className = 'project-overview__title';
  title.textContent = project.Title || 'Projeto sem título';
  const status = document.createElement('span');
  status.className = 'status-pill';
  status.style.background = statusColor(project.status);
  status.textContent = project.status || 'Sem status';
  header.append(title, status);

  if (project.status === 'Aprovado') {
    const info = document.createElement('p');
    info.className = 'project-overview__hint';
    info.textContent = 'Projeto aprovado - somente leitura.';
    header.append(info);
  }

  wrapper.append(header);

  const highlightGrid = document.createElement('div');
  highlightGrid.className = 'project-overview__grid';
  highlightGrid.append(
    createHighlightBox('Orçamento', project.budgetBrl ? BRL.format(project.budgetBrl) : '—', { variant: 'budget' }),
    createHighlightBox('Responsável', project.projectLeader || project.projectUser || 'Não informado')
  );
  wrapper.append(highlightGrid);

  const timelineGrid = document.createElement('div');
  timelineGrid.className = 'project-overview__grid';
  timelineGrid.append(
    createHighlightBox('Data de Início', formatDateValue(project.startDate)),
    createHighlightBox('Data de Conclusão', formatDateValue(project.endDate))
  );
  wrapper.append(timelineGrid);

  const descriptionSection = document.createElement('section');
  descriptionSection.className = 'project-description';
  const descTitle = document.createElement('h3');
  descTitle.textContent = 'Descrição do Projeto';
  const descText = document.createElement('p');
  descText.className = 'project-description__text';
  descText.textContent = project.proposedSolution || project.businessNeed || 'Sem descrição informada.';
  descriptionSection.append(descTitle, descText);
  wrapper.append(descriptionSection);

  const actions = document.createElement('div');
  actions.className = 'project-overview__actions';

  const statusKey = normalizeStatusKey(project.status);
  const editableStatuses = new Set(['Rascunho', 'Reprovado', 'Reprovado para Revisão']);
  const viewOnlyStatuses = new Set(['Aprovado', 'Em Aprovação']);

  if (viewOnlyStatuses.has(statusKey)) {
    const viewBtn = document.createElement('button');
    viewBtn.type = 'button';
    viewBtn.className = 'btn ghost';
    viewBtn.textContent = 'Visualizar Projeto';
    viewBtn.addEventListener('click', () => openProjectForm('edit', detail));
    actions.append(viewBtn);
  } else if (editableStatuses.has(statusKey)) {
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn primary';
    editBtn.textContent = 'Editar Projeto';
    editBtn.addEventListener('click', () => openProjectForm('edit', detail));
    actions.append(editBtn);

    if (canSubmitForApproval(statusKey)) {
      const approveBtn = document.createElement('button');
      approveBtn.type = 'button';
      approveBtn.className = 'btn accent';
      approveBtn.textContent = 'Enviar para Aprovação';
      approveBtn.addEventListener('click', () => {
        openProjectForm('edit', detail);
        requestAnimationFrame(() => {
          projectForm.dataset.submitIntent = 'approval';
          openSummaryOverlay(submitForApprovalBtn || approveBtn);
        });
      });
      actions.append(approveBtn);
    }
  } else {
    const fallbackBtn = document.createElement('button');
    fallbackBtn.type = 'button';
    fallbackBtn.className = 'btn ghost';
    fallbackBtn.textContent = 'Visualizar Projeto';
    fallbackBtn.addEventListener('click', () => openProjectForm('edit', detail));
    actions.append(fallbackBtn);
  }

  if (actions.childElementCount) {
    wrapper.append(actions);
  }

  projectDetails.append(wrapper);
}

function createEmptyState() {
  const empty = document.createElement('div');
  empty.className = 'empty-state';
  const title = document.createElement('h2');
  title.textContent = 'Selecione um projeto';
  const text = document.createElement('p');
  text.textContent = 'Escolha um item na lista ao lado para visualizar os detalhes.';
  empty.append(title, text);
  return empty;
}

function createHighlightBox(label, value, options = {}) {
  const { variant } = options;
  const box = document.createElement('div');
  box.className = 'project-highlight';
  if (variant) {
    box.classList.add(`project-highlight--${variant}`);
  }

  const labelEl = document.createElement('span');
  labelEl.className = 'project-highlight__label';
  labelEl.textContent = label;

  const valueEl = document.createElement('span');
  valueEl.className = 'project-highlight__value';
  valueEl.textContent = value || '—';
  if (variant === 'budget') {
    valueEl.classList.add('project-highlight__value--budget');
  }

  box.append(labelEl, valueEl);
  return box;
}

function formatDateValue(value) {
  if (!value) {
    return '—';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '—';
  }

  return DATE_FMT.format(date);
}

function populateSelectOptions(selectElement, options = [], selectedValue = '') {
  if (!selectElement) return;

  selectElement.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = 'Selecione...';
  if (!selectedValue) {
    placeholderOption.selected = true;
  }
  selectElement.appendChild(placeholderOption);

  let hasSelectedOption = false;

  options.forEach((text) => {
    const option = document.createElement('option');
    option.value = text;
    option.textContent = text;
    if (selectedValue && text === selectedValue) {
      option.selected = true;
      hasSelectedOption = true;
    }
    selectElement.appendChild(option);
  });

  if (selectedValue && !hasSelectedOption) {
    const fallbackOption = document.createElement('option');
    fallbackOption.value = selectedValue;
    fallbackOption.textContent = selectedValue;
    fallbackOption.selected = true;
    selectElement.appendChild(fallbackOption);
  }
}

function updateCompanyDependentFields(companyValue, selectedValues = {}) {
  const rules = companyRules[companyValue] || {
    centers: [],
    units: [],
    locations: [],
    depreciation: []
  };

  populateSelectOptions(centerSelect, rules.centers, selectedValues.center);
  populateSelectOptions(unitSelect, rules.units, selectedValues.unit);
  populateSelectOptions(locationSelect, rules.locations, selectedValues.location);
  populateSelectOptions(depreciationSelect, rules.depreciation, selectedValues.depreciation);
}

// ============================================================================
// Formulário: abertura, preenchimento e coleta dos dados
function setFormFieldsDisabled(disabled) {
  if (!projectForm) return;
  const elements = projectForm.querySelectorAll('input, textarea, select, button');

  elements.forEach((element) => {
    if (!element) return;
    if (element.id === 'closeFormBtn') return;
    if (element.closest('.form-summary')) return;

    if (disabled) {
      if (!Object.prototype.hasOwnProperty.call(element.dataset, 'readonlyOriginalDisabled')) {
        element.dataset.readonlyOriginalDisabled = element.disabled ? 'true' : 'false';
      }
      element.disabled = true;
      element.setAttribute('aria-disabled', 'true');
    } else {
      if (!Object.prototype.hasOwnProperty.call(element.dataset, 'readonlyOriginalDisabled')) {
        if (element.disabled) {
          element.setAttribute('aria-disabled', 'true');
        } else {
          element.removeAttribute('aria-disabled');
        }
        return;
      }
      const original = element.dataset.readonlyOriginalDisabled;
      if (original === 'true') {
        element.disabled = true;
        element.setAttribute('aria-disabled', 'true');
      } else {
        element.disabled = false;
        element.removeAttribute('aria-disabled');
      }
      if (original !== undefined) {
        delete element.dataset.readonlyOriginalDisabled;
      }
      if (element.disabled) {
        element.setAttribute('aria-disabled', 'true');
      }
    }
  });
}

function setFormMode(mode, options = {}) {
  if (!projectForm) return;
  const { showApprovalButton = false, refreshSummary = false } = options;
  const targetMode = mode === 'readonly' ? 'readonly' : 'edit';
  const modeChanged = currentFormMode !== targetMode;
  currentFormMode = targetMode;

  if (targetMode === 'readonly') {
    if (modeChanged) {
      projectForm.classList.add('project-form--readonly');
      setFormFieldsDisabled(true);
      if (formSummaryView) {
        formSummaryView.classList.remove('hidden');
      }
      if (formActions) {
        formActions.classList.add('hidden');
      }
      if (saveProjectBtn) {
        saveProjectBtn.classList.add('hidden');
      }
      if (submitForApprovalBtn) {
        submitForApprovalBtn.classList.add('hidden');
      }
    }

    if (formSummaryView && (modeChanged || refreshSummary)) {
      populateSummaryContent({ context: formSummaryContext, refreshGantt: true });
    }
  } else {
    if (modeChanged) {
      projectForm.classList.remove('project-form--readonly');
      setFormFieldsDisabled(false);
      if (formSummaryView) {
        formSummaryView.classList.add('hidden');
        clearSummaryContent(formSummaryContext);
      }
    }

    if (formActions) {
      formActions.classList.remove('hidden');
    }
    if (saveProjectBtn) {
      saveProjectBtn.classList.remove('hidden');
    }
    if (submitForApprovalBtn) {
      submitForApprovalBtn.classList.toggle('hidden', !showApprovalButton);
    }
  }
}

function applyStatusBehavior(status) {
  const statusKey = normalizeStatusKey(status);
  if (isReadOnlyStatus(statusKey)) {
    if (formTitle) {
      formTitle.textContent = 'Resumo do Projeto';
    }
    setFormMode('readonly', { refreshSummary: true });
    if (projectForm) {
      projectForm.dataset.submitIntent = 'save';
    }
    return;
  }

  const showApprovalButton = canSubmitForApproval(statusKey);
  setFormMode('edit', { showApprovalButton });
  if (projectForm) {
    projectForm.dataset.submitIntent = 'save';
  }
}

// ============================================================================
function openProjectForm(mode, detail = null) {
  projectForm.reset();
  resetFormStatus();
  resetValidationState();
  projectForm.dataset.mode = mode;
  projectForm.dataset.projectId = detail?.project?.Id || '';
  projectForm.dataset.submitIntent = 'save';

  currentFormMode = null;
  setFormMode('edit', { showApprovalButton: true });

  updateInvestmentLevelField();

  if (companySelect) {
    companySelect.value = '';
  }
  updateCompanyDependentFields('');

  state.editingSnapshot = {
    simplePeps: new Set(),
    milestones: new Set(),
    activities: new Set(),
    activityPeps: new Set()
  };

  simplePepList.innerHTML = '';
  milestoneList.innerHTML = '';

  if (mode === 'create') {
    formTitle.textContent = 'Novo Projeto';
    statusField.value = 'Rascunho';
    setApprovalYearToCurrent();
    updateBudgetSections({ clear: true });
  } else if (detail) {
    fillFormWithProject(detail);
  }

  const statusKey = detail?.project?.status || statusField.value || 'Rascunho';
  applyStatusBehavior(statusKey);

  updateSimplePepYears();
  overlay.classList.remove('hidden');
  queueGanttRefresh();
  validateAllDateRanges();
}

function fillFormWithProject(detail) {
  const { project, simplePeps, milestones, activities, activityPeps } = detail;
  formTitle.textContent = `Editar Projeto #${project.Id}`;
  statusField.value = project.status || 'Rascunho';

  document.getElementById('projectName').value = project.Title || '';
  document.getElementById('category').value = project.category || '';
  document.getElementById('investmentType').value = project.investmentType || '';
  document.getElementById('assetType').value = project.assetType || '';
  document.getElementById('projectFunction').value = project.projectFunction || '';

  document.getElementById('approvalYear').value = project.approvalYear || '';
  document.getElementById('startDate').value = project.startDate ? project.startDate.substring(0, 10) : '';
  document.getElementById('endDate').value = project.endDate ? project.endDate.substring(0, 10) : '';

  document.getElementById('projectBudget').value = sanitizeNumericInputValue(project.budgetBrl);
  updateInvestmentLevelField();
  document.getElementById('fundingSource').value = project.fundingSource || '';
  const selectedCompany = project.company || '';
  if (companySelect) {
    companySelect.value = selectedCompany;
  }
  updateCompanyDependentFields(selectedCompany, {
    center: project.center || '',
    unit: project.unit || '',
    location: project.location || '',
    depreciation: project.depreciationCostCenter || ''
  });

  document.getElementById('projectUser').value = project.projectUser || '';
  document.getElementById('projectLeader').value = project.projectLeader || '';

  document.getElementById('businessNeed').value = project.businessNeed || '';
  document.getElementById('proposedSolution').value = project.proposedSolution || '';

  document.getElementById('kpiType').value = project.kpiType || '';
  document.getElementById('kpiName').value = project.kpiName || '';
  document.getElementById('kpiDescription').value = project.kpiDescription || '';
  document.getElementById('kpiCurrent').value = project.kpiCurrent ?? '';
  document.getElementById('kpiExpected').value = project.kpiExpected ?? '';

  updateBudgetSections({ preserve: true });

  if (project.budgetBrl < BUDGET_THRESHOLD) {
    simplePeps.forEach((pep) => {
      const row = createSimplePepRow({
        id: pep.Id,
        title: pep.Title,
        amount: pep.amountBrl,
        year: pep.year
      });
      simplePepList.append(row);
      state.editingSnapshot.simplePeps.add(Number(pep.Id));
    });
    if (!simplePeps.length) {
      ensureSimplePepRow();
    }
  } else {
    milestones.forEach((milestone) => {
      const block = createMilestoneBlock({
        id: milestone.Id,
        title: milestone.Title
      });
      const relatedActivities = activities.filter((act) => act.milestonesIdId === milestone.Id);
      relatedActivities.forEach((activity) => {
        const relatedPeps = activityPeps.filter((pep) => pep.activitiesId === activity.Id);
        const primaryPep = relatedPeps[0] || null;
        const activityBlock = addActivityBlock(block, {
          id: activity.Id,
          title: activity.Title,
          start: activity.startDate,
          end: activity.endDate,
          supplier: activity.supplier,
          description: activity.activityDescription,
          pepId: primaryPep?.Id,
          pepTitle: primaryPep?.Title,
          pepAmount: primaryPep?.amountBrl,
          pepYear: primaryPep?.year
        });
        relatedPeps.forEach((pep) => {
          state.editingSnapshot.activityPeps.add(Number(pep.Id));
        });
        state.editingSnapshot.activities.add(Number(activity.Id));
      });
      if (!relatedActivities.length) {
        addActivityBlock(block);
      }
      milestoneList.append(block);
      state.editingSnapshot.milestones.add(Number(milestone.Id));
    });
    if (!milestones.length) {
      ensureMilestoneBlock();
    }
  }

  queueGanttRefresh();
  validatePepBudget();
  validateActivityDates();
  validateAllDateRanges();
}

function closeForm() {
  overlay.classList.add('hidden');
  closeSummaryOverlay({ restoreFocus: false });
}

function openSummaryOverlay(triggerButton = null) {
  const validation = runFormValidations({ scrollOnError: true, focusFirstError: true });
  if (!validation.valid) {
    return;
  }

  if (!summaryOverlay) {
    return;
  }

  summaryTriggerButton = triggerButton || null;
  if (summaryConfirmBtn) {
    const intent = projectForm.dataset.submitIntent || 'save';
    summaryConfirmBtn.textContent = intent === 'approval' ? 'Enviar para Aprovação' : 'Confirmar';
  }
  populateSummaryOverlay();
  summaryOverlay.classList.remove('hidden');
  summaryOverlay.scrollTop = 0;
  if (summaryConfirmBtn) {
    summaryConfirmBtn.focus();
  }
}

function closeSummaryOverlay(options = {}) {
  const { restoreFocus = true } = options;
  if (!summaryOverlay || summaryOverlay.classList.contains('hidden')) {
    summaryTriggerButton = null;
    return;
  }

  summaryOverlay.classList.add('hidden');
  if (summarySections) {
    summarySections.innerHTML = '';
  }
  if (summaryGanttChart) {
    summaryGanttChart.innerHTML = '';
  }

  if (restoreFocus && summaryTriggerButton) {
    summaryTriggerButton.focus();
  }
  summaryTriggerButton = null;
}

function handleSummaryConfirm() {
  closeSummaryOverlay({ restoreFocus: false });
  projectForm.requestSubmit();
}

function populateSummaryOverlay() {
  if (!summarySections) return;
  populateSummaryContent({ context: defaultSummaryContext, refreshGantt: true });
}

function getSummarySectionsData() {
  return [
    {
      title: 'Sobre o Projeto',
      entries: [
        { label: 'Nome do Projeto', value: getFieldDisplayValue('projectName') },
        { label: 'Orçamento do Projeto', value: formatCurrencyField('projectBudget') },
        { label: 'Nível de Investimento', value: getFieldDisplayValue('investmentLevel') },
        { label: 'Ano de Aprovação', value: getFieldDisplayValue('approvalYear') },
        { label: 'Data de Início', value: formatDateValue(document.getElementById('startDate')?.value) },
        { label: 'Data de Término', value: formatDateValue(document.getElementById('endDate')?.value) }
      ]
    },
    {
      title: 'Origem e Função',
      entries: [
        { label: 'Origem da Verba', value: getFieldDisplayValue('fundingSource') },
        { label: 'Função do Projeto', value: getFieldDisplayValue('projectFunction') },
        { label: 'Tipo de Investimento', value: getFieldDisplayValue('investmentType') },
        { label: 'Tipo de Ativo', value: getFieldDisplayValue('assetType') }
      ]
    },
    {
      title: 'Informações Operacionais',
      entries: [
        { label: 'Empresa', value: getFieldDisplayValue('company') },
        { label: 'Centro', value: getFieldDisplayValue('center') },
        { label: 'Unidade', value: getFieldDisplayValue('unit') },
        { label: 'Local de Implantação', value: getFieldDisplayValue('location') },
        { label: 'C. Custo Depreciação', value: getFieldDisplayValue('depreciationCostCenter') },
        { label: 'Categoria', value: getFieldDisplayValue('category') },
        { label: 'Usuário do Projeto', value: getFieldDisplayValue('projectUser') },
        { label: 'Líder do Projeto', value: getFieldDisplayValue('projectLeader') }
      ]
    },
    {
      title: 'Detalhamento Complementar',
      entries: [
        { label: 'Necessidade do Negócio', value: getFieldDisplayValue('businessNeed'), fullWidth: true },
        { label: 'Solução da Proposta', value: getFieldDisplayValue('proposedSolution'), fullWidth: true }
      ]
    },
    {
      title: 'Indicadores de Desempenho',
      entries: [
        { label: 'Tipo de KPI', value: getFieldDisplayValue('kpiType') },
        { label: 'Nome do KPI', value: getFieldDisplayValue('kpiName') },
        { label: 'KPI Atual', value: formatNumberField('kpiCurrent') },
        { label: 'KPI Esperado', value: formatNumberField('kpiExpected') },
        { label: 'Descrição do KPI', value: getFieldDisplayValue('kpiDescription'), fullWidth: true }
      ]
    }
  ];
}

function populateSummaryContent(options = {}) {
  const { context = defaultSummaryContext, refreshGantt = false } = options;
  const sections = context?.sections;
  if (!sections) return;

  withSummaryContext(context, () => {
    clearSummaryContent(context);
    getSummarySectionsData().forEach((section) => {
      createSummarySection(section.title, section.entries);
    });
    renderPepSummary();
    renderMilestoneSummary();
    populateSummaryGantt({ refreshFirst: refreshGantt });
  });
}

function createSummarySection(title, entries = []) {
  const sections = activeSummaryContext?.sections;
  if (!sections || !entries.length) return;

  const section = document.createElement('section');
  section.className = 'summary-section';

  const heading = document.createElement('h3');
  heading.textContent = title;
  section.appendChild(heading);

  const list = document.createElement('div');
  list.className = 'summary-list';

  entries.forEach((entry) => {
    if (!entry?.label) return;
    const item = document.createElement('div');
    item.className = 'summary-item';
    if (entry.fullWidth) {
      item.classList.add('summary-item--full');
    }

    const label = document.createElement('span');
    label.className = 'summary-label';
    label.textContent = entry.label;

    const value = document.createElement('p');
    value.className = 'summary-value';
    value.textContent = resolveSummaryValue(entry.value);

    item.append(label, value);
    list.appendChild(item);
  });

  if (list.children.length) {
    section.appendChild(list);
    sections.appendChild(section);
  }
}

function renderPepSummary() {
  const sections = activeSummaryContext?.sections;
  if (!sections) return;
  if (!simplePepList || !milestoneList) return;

  const rows = [];

  if (!simplePepSection.classList.contains('hidden')) {
    simplePepList.querySelectorAll('.pep-row').forEach((row) => {
      const element = getSelectOptionText(row.querySelector('.pep-title'));
      const amount = formatCurrencyValueFromElement(row.querySelector('.pep-amount'));
      const year = row.querySelector('.pep-year')?.value ?? '';

      if (
        resolveSummaryValue(element) === '—' &&
        resolveSummaryValue(amount) === '—' &&
        resolveSummaryValue(year) === '—'
      ) {
        return;
      }

      rows.push({
        element,
        amount,
        year
      });
    });
  } else if (!keyProjectSection.classList.contains('hidden')) {
    milestoneList.querySelectorAll('.activity').forEach((activity) => {
      const element = getSelectOptionText(activity.querySelector('.activity-pep-title'));
      const amount = formatCurrencyValueFromElement(activity.querySelector('.activity-pep-amount'));
      const year = activity.querySelector('.activity-pep-year')?.value ?? '';
      const activityTitle = activity.querySelector('.activity-title')?.value ?? '';

      if (
        resolveSummaryValue(element) === '—' &&
        resolveSummaryValue(amount) === '—' &&
        resolveSummaryValue(year) === '—'
      ) {
        return;
      }

      rows.push({
        element,
        amount,
        year,
        activity: activityTitle
      });
    });
  }

  if (!rows.length) {
    return;
  }

  const hasActivityColumn = rows.some((row) => resolveSummaryValue(row.activity) !== '—');

  const section = document.createElement('section');
  section.className = 'summary-section';

  const heading = document.createElement('h3');
  heading.textContent = 'Elemento PEP';
  section.appendChild(heading);

  const table = document.createElement('table');
  table.className = 'summary-table';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const headers = ['Elemento PEP', 'Valor (R$)', 'Ano'];
  if (hasActivityColumn) {
    headers.push('Atividade');
  }
  headers.forEach((label) => {
    const th = document.createElement('th');
    th.textContent = label;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    const cells = [
      resolveSummaryValue(row.element),
      resolveSummaryValue(row.amount),
      resolveSummaryValue(row.year)
    ];

    if (hasActivityColumn) {
      cells.push(resolveSummaryValue(row.activity));
    }

    cells.forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  section.appendChild(table);
  sections.appendChild(section);
}

function renderMilestoneSummary() {
  const sections = activeSummaryContext?.sections;
  if (!sections || !milestoneList) return;
  if (keyProjectSection.classList.contains('hidden')) return;

  const milestones = milestoneList.querySelectorAll('.milestone');
  if (!milestones.length) return;

  const section = document.createElement('section');
  section.className = 'summary-section';

  const heading = document.createElement('h3');
  heading.textContent = 'Key Projects';
  section.appendChild(heading);

  const wrapper = document.createElement('div');
  wrapper.className = 'summary-milestones';

  milestones.forEach((milestone, index) => {
    const card = document.createElement('article');
    card.className = 'summary-milestone';

    const titleInput = milestone.querySelector('.milestone-title');
    const resolvedTitle = resolveSummaryValue(titleInput?.value);
    const titleText = resolvedTitle === '—' ? `Marco ${index + 1}` : resolvedTitle;

    const title = document.createElement('h4');
    title.textContent = titleText;
    card.appendChild(title);

    const activities = milestone.querySelectorAll('.activity');
    if (activities.length) {
      const activityContainer = document.createElement('div');
      activityContainer.className = 'summary-activities';

      activities.forEach((activity, actIndex) => {
        const activityCard = document.createElement('article');
        activityCard.className = 'summary-activity';

        const activityTitleInput = activity.querySelector('.activity-title');
        const resolvedActivityTitle = resolveSummaryValue(activityTitleInput?.value);
        const activityTitle = resolvedActivityTitle === '—' ? `Atividade ${actIndex + 1}` : resolvedActivityTitle;

        const headingEl = document.createElement('h5');
        headingEl.textContent = activityTitle;
        activityCard.appendChild(headingEl);

        const detailList = document.createElement('div');
        detailList.className = 'summary-list summary-list--activity';

        const detailItems = [
          { label: 'Período', value: buildActivityPeriod(activity) },
          { label: 'Valor da Atividade', value: formatCurrencyValueFromElement(activity.querySelector('.activity-pep-amount')) },
          { label: 'Elemento PEP', value: getSelectOptionText(activity.querySelector('.activity-pep-title')) },
          { label: 'Ano do PEP', value: activity.querySelector('.activity-pep-year')?.value ?? '' },
          { label: 'Fornecedor', value: activity.querySelector('.activity-supplier')?.value ?? '' },
          { label: 'Descrição', value: activity.querySelector('.activity-description')?.value ?? '', fullWidth: true }
        ];

        detailItems.forEach((item) => {
          if (!item?.label) return;
          const detailItem = document.createElement('div');
          detailItem.className = 'summary-item';
          if (item.fullWidth) {
            detailItem.classList.add('summary-item--full');
          }

          const detailLabel = document.createElement('span');
          detailLabel.className = 'summary-label';
          detailLabel.textContent = item.label;

          const detailValue = document.createElement('p');
          detailValue.className = 'summary-value';
          detailValue.textContent = resolveSummaryValue(item.value);

          detailItem.append(detailLabel, detailValue);
          detailList.appendChild(detailItem);
        });

        activityCard.appendChild(detailList);
        activityContainer.appendChild(activityCard);
      });

      card.appendChild(activityContainer);
    }

    wrapper.appendChild(card);
  });

  section.appendChild(wrapper);
  sections.appendChild(section);
}

function populateSummaryGantt(options = {}) {
  const { refreshFirst = false } = options;
  const context = activeSummaryContext;
  const ganttSection = context?.ganttSection;
  const ganttChart = context?.ganttChart;
  if (!ganttSection || !ganttChart) return;

  if (context) {
    context.lastGanttResult = undefined;
  }

  if (refreshFirst) {
    refreshGantt();
    requestAnimationFrame(() => populateSummaryGantt());
    return;
  }

  if (keyProjectSection.classList.contains('hidden')) {
    ganttSection.classList.add('hidden');
    ganttChart.innerHTML = '';
    if (context) {
      context.lastGanttResult = null;
    }
    return;
  }

  const assignResult = (result) => {
    if (context) {
      context.lastGanttResult = result || null;
    }
    return result;
  };

  const drawSummary = () =>
    assignResult(
      drawGantt(collectMilestonesForGantt(), {
        container: ganttSection,
        chartElement: ganttChart,
        titleElement: ganttSection.querySelector('h3'),
        emptyMessage: 'Nenhuma atividade para exibir'
      })
    );

  if (ganttReady && window.google?.visualization?.Gantt) {
    drawSummary();
  } else if (ganttLoaderStarted && window.google?.charts) {
    google.charts.setOnLoadCallback(drawSummary);
  } else if (window.google?.charts) {
    initGantt();
    google.charts.setOnLoadCallback(drawSummary);
  } else {
    ganttSection.classList.remove('hidden');
    ganttChart.innerHTML = '<p class="gantt-empty">Nenhuma atividade para exibir</p>';
    if (context) {
      context.lastGanttResult = null;
    }
  }
}

async function waitForSummaryRender(context, options = {}) {
  if (!context) return;
  const { timeout = 2000 } = options;
  const start = Date.now();

  while (context.lastGanttResult === undefined && Date.now() - start < timeout) {
    await new Promise((resolve) => setTimeout(resolve, 50));
  }

  const result = context.lastGanttResult;
  if (result?.chart && window.google?.visualization?.events) {
    await new Promise((resolve) => {
      const listener = google.visualization.events.addListener(result.chart, 'ready', () => {
        google.visualization.events.removeListener(listener);
        resolve();
      });
    });
  }

  await new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

function buildActivityPeriod(activity) {
  if (!activity) return '';
  const startValue = activity.querySelector('.activity-start')?.value;
  const endValue = activity.querySelector('.activity-end')?.value;
  const start = formatDateValue(startValue);
  const end = formatDateValue(endValue);

  const hasStart = start !== '—';
  const hasEnd = end !== '—';

  if (hasStart && hasEnd) {
    return `${start} a ${end}`;
  }
  if (hasStart) {
    return `A partir de ${start}`;
  }
  if (hasEnd) {
    return `Até ${end}`;
  }
  return '';
}

function getFieldDisplayValue(fieldId) {
  const field = document.getElementById(fieldId);
  if (!field) return '';
  if (field.tagName === 'SELECT') {
    return getSelectOptionText(field);
  }
  return field.value ?? '';
}

function getSelectOptionText(selectElement) {
  if (!selectElement) return '';
  const option = selectElement.options?.[selectElement.selectedIndex];
  if (option) {
    return option.textContent?.trim() ?? '';
  }
  return selectElement.value ?? '';
}

function formatCurrencyField(fieldId) {
  const field = document.getElementById(fieldId);
  if (!field) return '';
  const raw = field.value;
  if (raw === undefined || raw === null || raw === '') {
    return '';
  }
  const value = parseNumericInputValue(raw);
  if (!Number.isFinite(value)) {
    return '';
  }
  return BRL.format(value);
}

function formatNumberField(fieldId) {
  const field = document.getElementById(fieldId);
  if (!field) return '';
  const raw = field.value;
  if (raw === undefined || raw === null || raw === '') {
    return '';
  }
  const normalized = String(raw).replace(',', '.');
  const value = Number.parseFloat(normalized);
  if (!Number.isFinite(value)) {
    return raw;
  }
  return value.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
}

function formatCurrencyValueFromElement(element) {
  if (!element) return '';
  const raw = element.value;
  if (raw === undefined || raw === null || raw === '') {
    return '';
  }
  const value = parseNumericInputValue(element);
  if (!Number.isFinite(value)) {
    return '';
  }
  return BRL.format(value);
}

function resolveSummaryValue(value) {
  if (value === null || value === undefined) {
    return '—';
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      return '—';
    }
    return value.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
  }
  const text = String(value).trim();
  return text ? text : '—';
}

/**
 * Solicita confirmação antes de fechar o formulário.
 * Usada tanto pelo botão Fechar quanto pela tecla ESC.
 * Também fecha o resumo, caso esteja aberto, antes de perguntar ao usuário.
 */
function handleCloseFormRequest() {
  if (summaryOverlay && !summaryOverlay.classList.contains('hidden')) {
    closeSummaryOverlay();
    return;
  }

  if (!overlay || overlay.classList.contains('hidden')) {
    return;
  }

  const shouldClose = window.confirm(
    'Tem certeza que deseja fechar o formulário? Suas alterações não salvas serão perdidas.'
  );

  if (shouldClose) {
    closeForm();
  }
}

function handleOverlayEscape(event) {
  if (event.key !== 'Escape') return;
  handleCloseFormRequest();
}

function updateBudgetSections(options = {}) {
  const { preserve = false, clear = false } = options;
  const value = getProjectBudgetValue();
  const isNumber = Number.isFinite(value);
  setSectionInteractive(simplePepSection, false);
  setSectionInteractive(keyProjectSection, false);

  if (!isNumber) {
    simplePepSection.classList.add('hidden');
    keyProjectSection.classList.add('hidden');
    if (clear) {
      simplePepList.innerHTML = '';
      milestoneList.innerHTML = '';
    }
    updateBudgetHintMessage({ budget: NaN });
    return;
  }

  if (value >= BUDGET_THRESHOLD) {
    simplePepSection.classList.add('hidden');
    keyProjectSection.classList.remove('hidden');
    setSectionInteractive(keyProjectSection, true);
    if (!preserve) {
      simplePepList.innerHTML = '';
    }
    if (!milestoneList.children.length) {
      ensureMilestoneBlock();
    }
  } else {
    keyProjectSection.classList.add('hidden');
    simplePepSection.classList.remove('hidden');
    setSectionInteractive(simplePepSection, true);
    if (!preserve) {
      milestoneList.innerHTML = '';
    }
    if (!simplePepList.children.length && !preserve) {
      ensureSimplePepRow();
    }
  }

  updateBudgetHintMessage({ budget: value });
  queueGanttRefresh();
}

function setSectionInteractive(section, enabled) {
  if (!section) return;
  section.querySelectorAll('input, textarea, button').forEach((element) => {
    if (element.type === 'hidden') return;
    element.disabled = !enabled;
  });
}

// ============================================================================
// Validações de formulário
// ============================================================================
function resetFormStatus() {
  if (!formStatus) return;
  formStatus.textContent = '';
  formStatus.classList.remove('show', 'feedback--info', 'feedback--success', 'feedback--warning', 'feedback--error');
}

function clearErrorSummary() {
  if (!formErrors) return;
  formErrors.classList.remove('show', 'feedback--info', 'feedback--success', 'feedback--warning', 'feedback--error');
  if (formErrorsTitle) {
    formErrorsTitle.textContent = '';
  }
  if (errorSummaryList) {
    errorSummaryList.innerHTML = '';
  }
}

function renderErrorSummary(issues = [], options = {}) {
  if (!formErrors || !errorSummaryList) {
    return;
  }

  const entries = Array.isArray(issues) ? issues.filter(Boolean) : [];
  if (!entries.length) {
    clearErrorSummary();
    return;
  }

  const tone = options.tone || 'error';
  const titleMessage = options.title || (tone === 'warning'
    ? 'Alguns campos estão incompletos.'
    : 'Não foi possível salvar. Verifique os erros abaixo.');

  formErrors.classList.add('show');
  formErrors.classList.remove('feedback--info', 'feedback--success', 'feedback--warning', 'feedback--error');
  formErrors.classList.add(`feedback--${tone}`);
  if (formErrorsTitle) {
    formErrorsTitle.textContent = titleMessage;
  }

  errorSummaryList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  entries.forEach((issue) => {
    const messages = Array.isArray(issue.items)
      ? [...new Set(issue.items.filter((item) => Boolean(item && String(item).trim())))]
      : [];
    const listItem = document.createElement('li');
    listItem.className = 'error-summary__item';

    if (issue.title) {
      const strong = document.createElement('strong');
      strong.textContent = issue.title;
      listItem.appendChild(strong);
    }

    if (messages.length === 1) {
      const messageText = issue.title ? ` — ${messages[0]}` : messages[0];
      listItem.appendChild(document.createTextNode(messageText));
    } else if (messages.length > 1) {
      if (issue.title) {
        listItem.appendChild(document.createTextNode(':'));
      }
      const nestedList = document.createElement('ul');
      nestedList.className = 'error-summary__sublist';
      messages.forEach((message) => {
        const nestedItem = document.createElement('li');
        nestedItem.textContent = message;
        nestedList.appendChild(nestedItem);
      });
      listItem.appendChild(nestedList);
    } else if (issue.message) {
      const messageText = issue.title ? ` — ${issue.message}` : issue.message;
      listItem.appendChild(document.createTextNode(messageText));
    } else if (!issue.title) {
      return;
    }

    fragment.appendChild(listItem);
  });

  errorSummaryList.appendChild(fragment);
}

function ensureFieldControlWrapper(element) {
  if (!element) return null;
  const parent = element.parentElement;
  if (!parent) return null;
  if (parent.classList.contains('field-control')) {
    return parent;
  }
  if (parent.classList.contains('field-group')) {
    const wrapper = document.createElement('div');
    wrapper.className = 'field-control';
    parent.insertBefore(wrapper, element);
    wrapper.appendChild(element);
    return wrapper;
  }
  return parent;
}

function clearFieldErrors() {
  if (!projectForm) return;
  projectForm.querySelectorAll('.field-error').forEach((node) => node.remove());
  projectForm.querySelectorAll('.field-group.has-error').forEach((group) => group.classList.remove('has-error'));
  projectForm.querySelectorAll('.field-invalid').forEach((input) => input.classList.remove('field-invalid'));
}

function applyFieldError(element, message) {
  if (!element) return;
  const group = element.closest('.field-group');
  if (!group) return;
  group.classList.add('has-error');
  element.classList.add('field-invalid');
  const wrapper = ensureFieldControlWrapper(element);
  if (!wrapper) return;
  let feedback = wrapper.querySelector('.field-error');
  if (!feedback) {
    feedback = document.createElement('span');
    feedback.className = 'field-error';
    wrapper.appendChild(feedback);
  }
  feedback.textContent = message;
}

function getFieldLabel(element) {
  if (!element) return 'Campo';
  const group = element.closest('.field-group');
  const label = group?.querySelector('label');
  if (label?.textContent) {
    return label.textContent.trim();
  }
  return element.getAttribute('aria-label') || element.name || element.id || 'Campo';
}

function collectInvalidFields() {
  if (!projectForm) return [];
  const invalid = [];
  const elements = projectForm.querySelectorAll('input, select, textarea');

  elements.forEach((element) => {
    if (!element?.willValidate) return;
    if (element.disabled) return;
    if (element.type === 'hidden') return;
    if (element.closest('.hidden')) return;

    const validity = element.validity;
    if (!validity || validity.valid) {
      return;
    }

    const label = getFieldLabel(element);
    let type = 'general';
    let message = element.validationMessage || `Verifique “${label}”.`;

    if (validity.valueMissing) {
      type = 'required';
      message = `Preencha “${label}”.`;
    } else if (element.dataset?.dateRangeInvalid === 'true' || message === DATE_RANGE_ERROR_MESSAGE) {
      type = 'date';
      message = `${label}: ${DATE_RANGE_ERROR_MESSAGE}`;
    }

    invalid.push({ element, label, message, type });
  });

  return invalid;
}

function resetValidationState() {
  validationState.pepBudget = null;
  validationState.pepBudgetDetails = null;
  validationState.activityDates = null;
  validationState.activityDateDetails = null;
  clearDateRangeValidity();
  clearFieldErrors();
  clearErrorSummary();
}

function setValidationError(key, message, details = null) {
  validationState[key] = message || null;
  const detailsKey = `${key}Details`;
  if (detailsKey in validationState) {
    validationState[detailsKey] = details || null;
  }
}

function rememberFieldPreviousValue(element) {
  if (!element || typeof element !== 'object') return;
  if (!('dataset' in element)) return;
  element.dataset.previousValue = element.value ?? '';
}

function normalizeNumericString(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value.toString() : '';
  }

  let text = String(value).trim();
  if (!text) {
    return '';
  }

  let sanitized = text.replace(/\s+/g, '');
  let sign = '';
  if (sanitized.startsWith('-')) {
    sign = '-';
    sanitized = sanitized.slice(1);
  }

  sanitized = sanitized.replace(/[^0-9.,]/g, '');
  if (!sanitized) {
    return '';
  }

  const separators = sanitized.match(/[.,]/g) || [];
  const uniqueSeparators = new Set(separators);
  const lastComma = sanitized.lastIndexOf(',');
  const lastDot = sanitized.lastIndexOf('.');
  let decimalIndex = Math.max(lastComma, lastDot);
  let decimalSeparator = decimalIndex >= 0 ? sanitized[decimalIndex] : null;
  let integerPart = sanitized;
  let fractionalPart = '';

  if (decimalSeparator) {
    integerPart = sanitized.slice(0, decimalIndex);
    fractionalPart = sanitized.slice(decimalIndex + 1);

    if (
      fractionalPart.length > 2 &&
      separators.length > 1 &&
      uniqueSeparators.size === 1
    ) {
      decimalSeparator = null;
      integerPart = sanitized;
      fractionalPart = '';
    }
  }

  integerPart = integerPart.replace(/[.,]/g, '');
  if (!integerPart) {
    integerPart = '0';
  }

  if (!decimalSeparator) {
    return sign + integerPart;
  }

  fractionalPart = fractionalPart.replace(/[.,]/g, '');
  if (!fractionalPart) {
    return sign + integerPart;
  }

  return `${sign}${integerPart}.${fractionalPart}`;
}

function coerceNumericValue(value) {
  const normalized = normalizeNumericString(value);
  if (!normalized) {
    return NaN;
  }
  const number = Number.parseFloat(normalized);
  return Number.isFinite(number) ? number : NaN;
}

function sanitizeNumericInputValue(value) {
  const number = coerceNumericValue(value);
  return Number.isFinite(number) ? number.toString() : '';
}

function parseNumericInputValue(source) {
  if (!source) return 0;

  if (typeof source === 'object' && source !== null) {
    if (typeof source.valueAsNumber === 'number' && !Number.isNaN(source.valueAsNumber)) {
      return source.valueAsNumber;
    }
    const numericValue = coerceNumericValue(source.value);
    return Number.isFinite(numericValue) ? numericValue : 0;
  }

  const numericValue = coerceNumericValue(source);
  return Number.isFinite(numericValue) ? numericValue : 0;
}

function determineInvestmentLevel(budgetBrl) {
  if (!Number.isFinite(budgetBrl) || budgetBrl < 0) {
    return '';
  }

  const budgetUsd = budgetBrl / EXCHANGE_RATE;

  if (budgetUsd >= 150_000_000) {
    return 'N1';
  }
  if (budgetUsd >= 10_000_000) {
    return 'N2';
  }
  if (budgetUsd >= 2_000_000) {
    return 'N3';
  }
  return 'N4';
}

function getProjectBudgetValue() {
  if (!projectBudgetInput) return NaN;
  const rawValue = projectBudgetInput.value;
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return NaN;
  }
  return parseNumericInputValue(projectBudgetInput);
}

function updateInvestmentLevelField(budgetBrl = getProjectBudgetValue()) {
  if (!investmentLevelSelect) return;
  const level = determineInvestmentLevel(budgetBrl);
  investmentLevelSelect.value = level;
}

function getPepAmountInputs() {
  const simplePepInputs = Array.from(simplePepList.querySelectorAll('.pep-amount'));
  const activityPepInputs = Array.from(milestoneList.querySelectorAll('.activity-pep-amount'));
  return [...simplePepInputs, ...activityPepInputs];
}

function calculatePepTotal() {
  return getPepAmountInputs().reduce((sum, input) => sum + parseNumericInputValue(input), 0);
}

function updateBudgetHintMessage({ budget = getProjectBudgetValue(), total = calculatePepTotal() } = {}) {
  if (!budgetHint) return;

  if (!Number.isFinite(budget)) {
    if (total > 0) {
      budgetHint.textContent = 'Informe o orçamento do projeto para validar os valores de PEP.';
      budgetHint.style.color = '#c62828';
    } else {
      budgetHint.textContent = '';
      budgetHint.style.color = '';
    }
    return;
  }

  const remaining = budget - total;
  const tolerance = 0.009;
  if (remaining >= -tolerance) {
    const safeRemaining = Math.max(0, Math.round(remaining * 100) / 100);
    budgetHint.textContent = `💰 Orçamento restante: ${BRL.format(safeRemaining)} (Total PEPs: ${BRL.format(total)} de ${BRL.format(budget)})`;
    budgetHint.style.color = '#2e7d32';
    return;
  }

  const exceeded = Math.abs(Math.round(remaining * 100) / 100);
  budgetHint.textContent = `⚠️ Orçamento excedido em ${BRL.format(exceeded)} (Total PEPs: ${BRL.format(total)} de ${BRL.format(budget)})`;
  budgetHint.style.color = '#c62828';
}

function validatePepBudget(options = {}) {
  const { changedInput = null } = options;
  const budget = getProjectBudgetValue();
  const total = calculatePepTotal();

  updateBudgetHintMessage({ budget, total });

  if (!Number.isFinite(budget)) {
    setValidationError('pepBudget', null, null);
    return true;
  }

  const remaining = budget - total;
  const roundedRemaining = Math.round(remaining * 100) / 100;

  if (remaining < -0.009) {
    const exceeded = Math.abs(roundedRemaining);
    const message = `A soma dos PEPs (${BRL.format(total)}) ultrapassa o orçamento do projeto (${BRL.format(budget)}) em ${BRL.format(exceeded)}.`;
    setValidationError('pepBudget', message, { budget, total, remaining: roundedRemaining });

    if (changedInput) {
      const previousValue = changedInput.dataset?.previousValue ?? '';
      if (changedInput.value !== previousValue) {
        changedInput.value = previousValue;
      }
    }
    return false;
  }

  setValidationError('pepBudget', null, { budget, total, remaining: roundedRemaining });
  if (changedInput) {
    rememberFieldPreviousValue(changedInput);
  }
  return true;
}

function parseDateInputValue(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function getDateRangePairs() {
  const pairs = [];

  if (projectStartDateInput || projectEndDateInput) {
    pairs.push({ start: projectStartDateInput, end: projectEndDateInput });
  }

  if (milestoneList) {
    milestoneList.querySelectorAll('.milestone').forEach((milestone) => {
      const milestoneStart = milestone.querySelector('.milestone-start');
      const milestoneEnd = milestone.querySelector('.milestone-end');
      if (milestoneStart || milestoneEnd) {
        pairs.push({ start: milestoneStart, end: milestoneEnd });
      }
    });

    milestoneList.querySelectorAll('.activity').forEach((activity) => {
      const startInput = activity.querySelector('.activity-start');
      const endInput = activity.querySelector('.activity-end');
      if (startInput || endInput) {
        pairs.push({ start: startInput, end: endInput });
      }
    });
  }

  return pairs;
}

function validateDateRange(startInput, endInput, options = {}) {
  const { report = false } = options;
  const relevantInputs = [startInput, endInput].filter(
    (input) => input && typeof input.setCustomValidity === 'function'
  );

  if (!relevantInputs.length) {
    return true;
  }

  relevantInputs.forEach((input) => {
    input.setCustomValidity('');
    if (input.dataset) {
      delete input.dataset.dateRangeInvalid;
    }
  });

  const startDate = startInput ? parseDateInputValue(startInput.value) : null;
  const endDate = endInput ? parseDateInputValue(endInput.value) : null;

  if (startDate && endDate && endDate < startDate) {
    const invalidTarget = endInput && typeof endInput.setCustomValidity === 'function' ? endInput : startInput;
    if (invalidTarget) {
      invalidTarget.setCustomValidity(DATE_RANGE_ERROR_MESSAGE);
      if (invalidTarget.dataset) {
        invalidTarget.dataset.dateRangeInvalid = 'true';
      }
      if (report && typeof invalidTarget.reportValidity === 'function') {
        invalidTarget.reportValidity();
      }
    }
    return false;
  }

  return true;
}

function validateAllDateRanges(options = {}) {
  const { report = false } = options;
  let firstInvalidField = null;

  getDateRangePairs().forEach(({ start, end }) => {
    const isValid = validateDateRange(start, end);
    if (!isValid && !firstInvalidField) {
      firstInvalidField = (end && typeof end.setCustomValidity === 'function') ? end : start;
    }
  });

  if (report && firstInvalidField && typeof firstInvalidField.reportValidity === 'function') {
    firstInvalidField.reportValidity();
  }

  return !firstInvalidField;
}

function clearDateRangeValidity() {
  const inputs = new Set();
  getDateRangePairs().forEach(({ start, end }) => {
    if (start) inputs.add(start);
    if (end) inputs.add(end);
  });

  inputs.forEach((input) => {
    if (input && typeof input.setCustomValidity === 'function') {
      input.setCustomValidity('');
      if (input.dataset) {
        delete input.dataset.dateRangeInvalid;
      }
    }
  });
}

function updateDateHintMessage({
  hasProjectStart = false,
  hasProjectEnd = false,
  hasStartIssue = false,
  hasEndIssue = false,
  activityCount = 0
} = {}) {
  if (!dateHint) return;

  if (!hasProjectStart || !hasProjectEnd) {
    dateHint.textContent = '';
    dateHint.style.color = '';
    return;
  }

  if (activityCount === 0) {
    dateHint.textContent = '✅ Todas as atividades estão dentro do intervalo do projeto.';
    dateHint.style.color = '#2e7d32';
    return;
  }

  if (!hasStartIssue && !hasEndIssue) {
    dateHint.textContent = '✅ Todas as atividades estão dentro do intervalo do projeto.';
    dateHint.style.color = '#2e7d32';
    return;
  }

  const messages = [];
  if (hasStartIssue) {
    messages.push('⚠️ Atividade começa antes da data inicial do projeto.');
  }
  if (hasEndIssue) {
    messages.push('⚠️ Atividade termina depois da data final do projeto.');
  }
  dateHint.textContent = messages.join(' ');
  dateHint.style.color = '#c62828';
}

function validateActivityDates(options = {}) {
  const { changedInput = null } = options;
  const projectStart = parseDateInputValue(projectStartDateInput?.value);
  const projectEnd = parseDateInputValue(projectEndDateInput?.value);

  const activities = milestoneList.querySelectorAll('.activity');
  const activityCount = activities.length;

  if (!projectStart || !projectEnd) {
    updateDateHintMessage({
      hasProjectStart: Boolean(projectStart),
      hasProjectEnd: Boolean(projectEnd),
      activityCount
    });
    setValidationError('activityDates', null, null);
    return true;
  }

  let invalidMessage = null;
  let invalidField = null;
  let invalidTitle = null;
  let hasStartIssue = false;
  let hasEndIssue = false;
  for (const activity of activities) {
    const startInput = activity.querySelector('.activity-start');
    const endInput = activity.querySelector('.activity-end');
    const title = activity.querySelector('.activity-title')?.value?.trim() || 'Atividade';
    const startDate = parseDateInputValue(startInput?.value);
    const endDate = parseDateInputValue(endInput?.value);

    if (startDate && startDate < projectStart) {
      hasStartIssue = true;
      if (!invalidMessage) {
        invalidMessage = `A data de início da atividade "${title}" não pode ser anterior à data de início do projeto.`;
        invalidField = changedInput === projectStartDateInput ? projectStartDateInput : startInput;
        invalidTitle = title;
      }
    }

    if (endDate && endDate > projectEnd) {
      hasEndIssue = true;
      if (!invalidMessage) {
        invalidMessage = `A data de término da atividade "${title}" não pode ser posterior à data de término do projeto.`;
        invalidField = changedInput === projectEndDateInput ? projectEndDateInput : endInput;
        invalidTitle = title;
      }
    }
  }

  updateDateHintMessage({
    hasProjectStart: true,
    hasProjectEnd: true,
    hasStartIssue,
    hasEndIssue,
    activityCount
  });

  if (invalidMessage) {
    setValidationError('activityDates', invalidMessage, {
      field: invalidField || null,
      activityTitle: invalidTitle,
      hasStartIssue,
      hasEndIssue
    });
    if (invalidField) {
      const previousValue = invalidField.dataset?.previousValue ?? '';
      if (invalidField.value !== previousValue) {
        invalidField.value = previousValue;
      }
    }
    return false;
  }

  setValidationError('activityDates', null, {
    field: null,
    activityTitle: null,
    hasStartIssue: false,
    hasEndIssue: false
  });
  if (changedInput) {
    rememberFieldPreviousValue(changedInput);
  }
  return true;
}

function runFormValidations(options = {}) {
  const { scrollOnError = false, focusFirstError = false } = options;

  clearFieldErrors();

  const pepValid = validatePepBudget();
  const activityValid = validateActivityDates();
  const dateRangesValid = validateAllDateRanges();

  const invalidFields = collectInvalidFields();
  invalidFields.forEach((issue) => {
    applyFieldError(issue.element, issue.message);
  });

  const issues = [];

  const requiredIssues = invalidFields.filter((issue) => issue.type === 'required');
  if (requiredIssues.length) {
    issues.push({
      title: 'Campos obrigatórios',
      items: requiredIssues.map((issue) => issue.message),
      type: 'required',
      focusElement: requiredIssues[0]?.element || null
    });
  }

  const dateFieldIssues = invalidFields.filter((issue) => issue.type === 'date');
  const dateMessages = dateFieldIssues.map((issue) => issue.message);
  if (validationState.activityDates) {
    dateMessages.push(validationState.activityDates);
  }
  const uniqueDateMessages = [...new Set(dateMessages.filter(Boolean))];
  if (uniqueDateMessages.length) {
    const activityField = validationState.activityDateDetails?.field || null;
    const isFieldAlreadyHighlighted = activityField
      ? invalidFields.some((issue) => issue.element === activityField)
      : false;
    if (activityField && !isFieldAlreadyHighlighted) {
      applyFieldError(activityField, validationState.activityDates || uniqueDateMessages[0]);
    }
    issues.push({
      title: 'Datas',
      items: uniqueDateMessages,
      type: 'date',
      focusElement: activityField || dateFieldIssues[0]?.element || null
    });
  }

  if (!pepValid && validationState.pepBudgetDetails) {
    const { budget = 0, total = 0, remaining = 0 } = validationState.pepBudgetDetails;
    const difference = Math.abs(Math.round(remaining * 100) / 100);
    const detailMessage = remaining < 0
      ? `Orçamento: ${BRL.format(budget)} · Soma dos PEPs: ${BRL.format(total)} · Excedente: ${BRL.format(difference)}.`
      : `Orçamento: ${BRL.format(budget)} · Soma dos PEPs: ${BRL.format(total)} · Restante: ${BRL.format(difference)}.`;
    const pepInputs = getPepAmountInputs().filter((input) => !input.closest('.hidden'));
    const pepFocus = pepInputs[0] || projectBudgetInput;
    if (projectBudgetInput && !invalidFields.some((issue) => issue.element === projectBudgetInput)) {
      applyFieldError(projectBudgetInput, 'Ajuste o orçamento do projeto ou redistribua os valores de PEP.');
    }
    if (pepFocus && pepFocus !== projectBudgetInput && !invalidFields.some((issue) => issue.element === pepFocus)) {
      applyFieldError(pepFocus, 'Revise os valores dos PEPs para manter o projeto dentro do orçamento.');
    }
    issues.push({
      title: 'Orçamento x PEPs',
      items: [detailMessage],
      type: 'pep',
      focusElement: pepFocus || null
    });
  }

  const otherIssues = invalidFields.filter((issue) => issue.type === 'general');
  if (otherIssues.length) {
    issues.push({
      title: 'Validações',
      items: otherIssues.map((issue) => issue.message),
      type: 'general',
      focusElement: otherIssues[0]?.element || null
    });
  }

  const isValid = issues.length === 0 && pepValid && activityValid && dateRangesValid;

  if (isValid) {
    clearErrorSummary();
    if (formStatus && (formStatus.classList.contains('feedback--warning') || formStatus.classList.contains('feedback--error'))) {
      resetFormStatus();
    }
    return { valid: true, issues: [] };
  }

  if (scrollOnError) {
    scrollFormToTop();
  }

  const hasNonRequiredIssue = issues.some((issue) => issue.type !== 'required');
  const severity = hasNonRequiredIssue ? 'error' : 'warning';
  const statusMessage = severity === 'warning'
    ? 'Alguns campos estão incompletos.'
    : 'Não foi possível salvar. Verifique os erros abaixo.';

  showStatus(statusMessage, { type: severity });
  renderErrorSummary(issues, { tone: severity, title: statusMessage });

  if (focusFirstError) {
    const focusCandidate = issues.find((issue) => issue.focusElement)?.focusElement
      || invalidFields[0]?.element
      || validationState.activityDateDetails?.field
      || null;
    if (focusCandidate && typeof focusCandidate.focus === 'function') {
      requestAnimationFrame(() => {
        try {
          focusCandidate.focus({ preventScroll: true });
        } catch (error) {
          focusCandidate.focus();
        }
      });
    }
  }

  return { valid: false, issues };
}

function handleGlobalDateInput(event) {
  const target = event.target;
  if (!target || typeof target.matches !== 'function') {
    return;
  }

  if (target === projectStartDateInput || target === projectEndDateInput) {
    validateDateRange(projectStartDateInput, projectEndDateInput, { report: true });
    return;
  }

  if (target.classList?.contains('activity-start') || target.classList?.contains('activity-end')) {
    const activity = target.closest('.activity');
    const startInput = activity?.querySelector('.activity-start');
    const endInput = activity?.querySelector('.activity-end');
    validateDateRange(startInput, endInput, { report: true });
    return;
  }

  if (target.classList?.contains('milestone-start') || target.classList?.contains('milestone-end')) {
    const milestone = target.closest('.milestone');
    const startInput = milestone?.querySelector('.milestone-start');
    const endInput = milestone?.querySelector('.milestone-end');
    validateDateRange(startInput, endInput, { report: true });
  }
}

function handleFormFocusCapture(event) {
  const target = event.target;
  if (!target) return;

  if (target === projectStartDateInput || target === projectEndDateInput) {
    rememberFieldPreviousValue(target);
    return;
  }

  if (target.classList?.contains('pep-amount')) {
    rememberFieldPreviousValue(target);
    return;
  }

  if (
    target.classList?.contains('activity-pep-amount') ||
    target.classList?.contains('activity-start') ||
    target.classList?.contains('activity-end')
  ) {
    rememberFieldPreviousValue(target);
  }
}

function updateActivityPepYear(activityElement, options = {}) {
  if (!activityElement) return;
  const pepYearInput = activityElement.querySelector('.activity-pep-year');
  if (!pepYearInput) return;

  const startValue = activityElement.querySelector('.activity-start')?.value || '';
  const startYear = startValue ? parseInt(startValue.substring(0, 4), 10) : NaN;
  const fallbackYear = options.fallbackYear ?? parseNumber(approvalYearInput.value);
  const resolvedYear = Number.isFinite(startYear) ? startYear : fallbackYear;

  if (options.force || !pepYearInput.value) {
    pepYearInput.value = resolvedYear ?? '';
  }
}

function updateSimplePepYears() {
  const year = parseInt(approvalYearInput.value, 10) || '';
  simplePepList.querySelectorAll('.pep-year').forEach((input) => {
    input.value = year;
  });
  if (milestoneList) {
    milestoneList.querySelectorAll('.activity').forEach((activity) => {
      const startValue = activity.querySelector('.activity-start')?.value;
      const forceUpdate = !startValue;
      updateActivityPepYear(activity, { fallbackYear: year || null, force: forceUpdate });
    });
  }
}

function ensureSimplePepRow() {
  const row = createSimplePepRow({ year: parseInt(approvalYearInput.value, 10) || '' });
  simplePepList.append(row);
}

function ensureMilestoneBlock() {
  const block = createMilestoneBlock();
  milestoneList.append(block);
  addActivityBlock(block);
  queueGanttRefresh();
}

function createSimplePepRow({ id = '', title = '', amount = '', year = '' } = {}) {
  const fragment = simplePepTemplate.content.cloneNode(true);
  const row = fragment.querySelector('.pep-row');
  row.dataset.pepId = id;
  row.querySelector('.pep-title').value = title || '';
  row.querySelector('.pep-amount').value = sanitizeNumericInputValue(amount);
  row.querySelector('.pep-year').value = year ?? '';
  return row;
}

function createMilestoneBlock({ id = '', title = '' } = {}) {
  const fragment = milestoneTemplate.content.cloneNode(true);
  const block = fragment.querySelector('.milestone');
  block.dataset.milestoneId = id;
  block.querySelector('.milestone-title').value = title || '';
  return block;
}

function addActivityBlock(milestoneElement, data = {}) {
  if (!milestoneElement) return null;
  const fragment = activityTemplate.content.cloneNode(true);
  const activity = fragment.querySelector('.activity');
  const startInput = activity.querySelector('.activity-start');
  const endInput = activity.querySelector('.activity-end');
  const amountInput = activity.querySelector('.activity-pep-amount');
  const pepTitleInput = activity.querySelector('.activity-pep-title');
  const pepYearInput = activity.querySelector('.activity-pep-year');

  activity.dataset.activityId = data.id || '';
  activity.dataset.pepId = data.pepId || '';

  activity.querySelector('.activity-title').value = data.title || '';
  if (amountInput) {
    amountInput.value = sanitizeNumericInputValue(data.pepAmount);
  }
  if (startInput) {
    startInput.value = data.start ? data.start.substring(0, 10) : '';
  }
  if (endInput) {
    endInput.value = data.end ? data.end.substring(0, 10) : '';
  }
  activity.querySelector('.activity-supplier').value = data.supplier || '';
  activity.querySelector('.activity-description').value = data.description || '';
  if (pepTitleInput) {
    pepTitleInput.value = data.pepTitle || '';
  }
  if (pepYearInput) {
    const startYear = startInput?.value ? parseInt(startInput.value.substring(0, 4), 10) : null;
    const fallbackYear = parseNumber(approvalYearInput.value) || null;
    const resolvedYear = data.pepYear ?? (Number.isFinite(startYear) ? startYear : null) ?? fallbackYear;
    pepYearInput.value = resolvedYear ?? '';
  }

  milestoneElement.querySelector('.activity-list').append(activity);
  if (!data.pepYear) {
    updateActivityPepYear(activity);
  }
  validateDateRange(startInput, endInput);
  queueGanttRefresh();
  return activity;
}

// ============================================================================
// Envio do formulário e persistência (CRUD)
// ============================================================================
async function handleFormSubmit(event) {
  event.preventDefault();
  const mode = projectForm.dataset.mode;
  const projectId = projectForm.dataset.projectId;
  const submitIntent = projectForm.dataset.submitIntent || 'save';
  const isApproval = submitIntent === 'approval';

  const validation = runFormValidations({ scrollOnError: true, focusFirstError: true });
  if (!validation.valid) {
    return;
  }

  const normalizedStatus = isApproval ? 'Em Aprovação' : 'Rascunho';
  statusField.value = normalizedStatus;

  const payload = collectProjectData();
  payload.status = normalizedStatus;

  scrollFormToTop();
  showStatus(isApproval ? 'Enviando para aprovação…' : 'Salvando…', { type: 'info' });

  let resolvedId = Number(projectId) || null;

  try {
    let savedProjectId = projectId;
    if (mode === 'create') {
      const result = await sp.createItem('Projects', payload);
      savedProjectId = result?.Id;
    } else {
      await sp.updateItem('Projects', Number(projectId), payload);
    }

    resolvedId = Number(savedProjectId || projectId);
    if (!Number.isFinite(resolvedId)) {
      throw new Error('ID do projeto inválido após salvar.');
    }

    await persistRelatedRecords(resolvedId, payload);

    if (isApproval) {
      const approvalSummary = buildApprovalSummary(resolvedId, payload);
      const jsonContent = JSON.stringify(approvalSummary, null, 2);
      const jsonBlob = new Blob([jsonContent], { type: 'application/json' });

      await sp.addAttachment('Projects', resolvedId, 'resumo.txt', jsonBlob, {
        overwrite: true,
        contentType: 'application/json'
      });

      await sp.updateItem('Projects', resolvedId, { status: 'Em Aprovação' });
    }

    if (resolvedId) {
      updateProjectState(resolvedId, {
        Title: payload.Title,
        status: payload.status,
        budgetBrl: payload.budgetBrl,
        investmentLevel: payload.investmentLevel
      });
      renderProjectList();
      if (state.currentDetails?.project?.Id === resolvedId) {
        state.currentDetails = {
          ...state.currentDetails,
          project: {
            ...state.currentDetails.project,
            Title: payload.Title,
            status: payload.status,
            budgetBrl: payload.budgetBrl,
            investmentLevel: payload.investmentLevel
          }
        };
        renderProjectDetails(state.currentDetails);
      }
    }

    const successMessage = isApproval
      ? 'Projeto enviado para aprovação!'
      : 'Projeto salvo com sucesso!';
    showStatus(successMessage, { type: 'success' });
    await loadProjects();
    if (resolvedId) {
      await selectProject(resolvedId);
    }
    closeForm();
  } catch (error) {
    console.error('Erro ao salvar projeto', error);

    if (isApproval && Number.isFinite(resolvedId)) {
      try {
        await sp.updateItem('Projects', resolvedId, { status: 'Rascunho' });
        updateProjectState(resolvedId, { status: 'Rascunho' });
        renderProjectList();
        if (state.currentDetails?.project?.Id === resolvedId) {
          state.currentDetails = {
            ...state.currentDetails,
            project: {
              ...state.currentDetails.project,
              status: 'Rascunho'
            }
          };
          renderProjectDetails(state.currentDetails);
        }
      } catch (rollbackError) {
        console.error('Erro ao reverter status após falha no envio', rollbackError);
      }
      statusField.value = 'Rascunho';
    }

    scrollFormToTop();
    const statusMessage = isApproval
      ? 'Não foi possível enviar para aprovação. Verifique os erros abaixo.'
      : 'Não foi possível salvar. Verifique os erros abaixo.';
    showStatus(statusMessage, { type: 'error' });
    renderErrorSummary(
      [
        {
          title: isApproval ? 'Erro ao enviar para aprovação' : 'Erro ao salvar',
          items: [
            isApproval
              ? 'Não foi possível concluir o envio para aprovação. Verifique os dados, tente novamente ou contate o suporte.'
              : 'Não foi possível salvar o projeto. Verifique os dados e tente novamente.'
          ],
          type: 'general'
        }
      ],
      { tone: 'error', title: statusMessage }
    );
  } finally {
    projectForm.dataset.submitIntent = 'save';
  }
}

function collectProjectData() {
  const budgetValue = getProjectBudgetValue();
  const budgetBrl = Number.isFinite(budgetValue) ? budgetValue : 0;
  const investmentLevelValue = determineInvestmentLevel(budgetValue);

  const data = {
    Title: document.getElementById('projectName').value.trim(),
    category: document.getElementById('category').value.trim(),
    investmentType: document.getElementById('investmentType').value.trim(),
    assetType: document.getElementById('assetType').value.trim(),
    projectFunction: document.getElementById('projectFunction').value.trim(),
    approvalYear: parseNumber(document.getElementById('approvalYear').value),
    startDate: document.getElementById('startDate').value || null,
    endDate: document.getElementById('endDate').value || null,
    budgetBrl,
    investmentLevel: investmentLevelValue,
    fundingSource: document.getElementById('fundingSource').value.trim(),
    depreciationCostCenter: document.getElementById('depreciationCostCenter').value.trim(),
    company: document.getElementById('company').value.trim(),
    center: document.getElementById('center').value.trim(),
    unit: document.getElementById('unit').value.trim(),
    location: document.getElementById('location').value.trim(),
    projectUser: document.getElementById('projectUser').value.trim(),
    projectLeader: document.getElementById('projectLeader').value.trim(),
    businessNeed: document.getElementById('businessNeed').value.trim(),
    proposedSolution: document.getElementById('proposedSolution').value.trim(),
    kpiType: document.getElementById('kpiType').value.trim(),
    kpiName: document.getElementById('kpiName').value.trim(),
    kpiDescription: document.getElementById('kpiDescription').value.trim(),
    kpiCurrent: document.getElementById('kpiCurrent').value.trim(),
    kpiExpected: document.getElementById('kpiExpected').value.trim()
  };
  return data;
}

async function persistRelatedRecords(projectId, projectData) {
  if (!projectId) return;
  const approvalYear = projectData.approvalYear;
  const budget = projectData.budgetBrl;

  if (budget >= BUDGET_THRESHOLD) {
    await persistKeyProjects(projectId);
    await cleanupSimplePeps();
  } else {
    await persistSimplePeps(projectId, approvalYear);
    await cleanupKeyProjects();
  }
}

function collectSimplePepDataForSummary() {
  if (!simplePepList) return [];

  const peps = [];
  simplePepList.querySelectorAll('.pep-row').forEach((row) => {
    const idValue = parseNumber(row.dataset.pepId);
    const titleSelect = row.querySelector('.pep-title');
    const titleValue = titleSelect?.value?.trim() || '';
    const titleDisplay = getSelectOptionText(titleSelect);
    const amountInput = row.querySelector('.pep-amount');
    const amountText = amountInput?.value ?? '';
    const yearInput = row.querySelector('.pep-year');
    const yearText = yearInput?.value ?? '';

    const hasData = Boolean(titleValue) || amountText.trim() !== '' || yearText.trim() !== '';
    if (!hasData) {
      return;
    }

    const parsedId = Number.isFinite(idValue) ? idValue : null;
    const parsedAmount = parseNumericInputValue(amountInput);
    const parsedYear = parseNumber(yearText);

    peps.push({
      id: parsedId,
      title: titleValue,
      titleDisplay,
      amountBrl: Number.isFinite(parsedAmount) ? parsedAmount : 0,
      amountText,
      year: Number.isFinite(parsedYear) ? parsedYear : null,
      yearText,
      type: 'simple'
    });
  });

  return peps;
}

function collectMilestonesForSummary() {
  if (!milestoneList) return [];

  const milestones = [];
  milestoneList.querySelectorAll('.milestone').forEach((milestoneEl) => {
    const milestoneIdValue = parseNumber(milestoneEl.dataset.milestoneId);
    const resolvedMilestoneId = Number.isFinite(milestoneIdValue) ? milestoneIdValue : null;
    const titleValue = milestoneEl.querySelector('.milestone-title')?.value.trim() || '';
    const activities = [];

    milestoneEl.querySelectorAll('.activity').forEach((activityEl) => {
      const activityIdValue = parseNumber(activityEl.dataset.activityId);
      const resolvedActivityId = Number.isFinite(activityIdValue) ? activityIdValue : null;
      const title = activityEl.querySelector('.activity-title')?.value.trim() || '';
      const startDateValue = activityEl.querySelector('.activity-start')?.value || '';
      const endDateValue = activityEl.querySelector('.activity-end')?.value || '';
      const supplier = activityEl.querySelector('.activity-supplier')?.value.trim() || '';
      const description = activityEl.querySelector('.activity-description')?.value.trim() || '';

      const pepTitleSelect = activityEl.querySelector('.activity-pep-title');
      const pepTitleValue = pepTitleSelect?.value?.trim() || '';
      const pepTitleDisplay = getSelectOptionText(pepTitleSelect);
      const pepAmountInput = activityEl.querySelector('.activity-pep-amount');
      const pepAmountText = pepAmountInput?.value ?? '';
      const pepYearInput = activityEl.querySelector('.activity-pep-year');
      const pepYearText = pepYearInput?.value ?? '';
      const pepIdValue = parseNumber(activityEl.dataset.pepId);

      const hasActivityData = [
        title,
        startDateValue,
        endDateValue,
        supplier,
        description,
        pepTitleValue,
        pepAmountText,
        pepYearText
      ].some((value) => (typeof value === 'string' ? value.trim() !== '' : value !== null));

      if (!hasActivityData) {
        return;
      }

      const parsedPepAmount = parseNumericInputValue(pepAmountInput);
      const parsedPepYear = parseNumber(pepYearText);

      const pep =
        pepTitleValue || pepAmountText.trim() !== '' || pepYearText.trim() !== ''
          ? {
              id: Number.isFinite(pepIdValue) ? pepIdValue : null,
              title: pepTitleValue,
              titleDisplay: pepTitleDisplay,
              amountBrl: Number.isFinite(parsedPepAmount) ? parsedPepAmount : 0,
              amountText: pepAmountText,
              year: Number.isFinite(parsedPepYear) ? parsedPepYear : null,
              yearText: pepYearText
            }
          : null;

      activities.push({
        id: resolvedActivityId,
        milestoneId: resolvedMilestoneId,
        title,
        startDate: startDateValue || null,
        endDate: endDateValue || null,
        supplier,
        description,
        pep
      });
    });

    if (!activities.length && !titleValue) {
      return;
    }

    milestones.push({
      id: resolvedMilestoneId,
      title: titleValue,
      activities
    });
  });

  return milestones;
}

function collectActivitiesForSummary(milestones) {
  const activities = [];
  const safeMilestones = Array.isArray(milestones) ? milestones : [];

  safeMilestones.forEach((milestone) => {
    const milestoneId = milestone?.id ?? null;
    const milestoneTitle = milestone?.title ?? '';
    const milestoneActivities = Array.isArray(milestone?.activities) ? milestone.activities : [];

    milestoneActivities.forEach((activity) => {
      if (!activity) return;
      activities.push({
        ...activity,
        milestoneId,
        milestoneTitle
      });
    });
  });

  return activities;
}

function collectPepEntriesForSummary(simplePeps, activities) {
  const peps = Array.isArray(simplePeps)
    ? simplePeps.map((pep) => ({ ...pep }))
    : [];

  const activityList = Array.isArray(activities) ? activities : [];
  activityList.forEach((activity) => {
    if (!activity?.pep) return;
    peps.push({
      ...activity.pep,
      type: 'activity',
      activityId: activity.id ?? null,
      activityTitle: activity.title ?? '',
      milestoneId: activity.milestoneId ?? null,
      milestoneTitle: activity.milestoneTitle ?? ''
    });
  });

  return peps;
}

function collectProjectDisplayValues() {
  return {
    investmentLevel: getSelectOptionText(investmentLevelSelect),
    company: getSelectOptionText(companySelect),
    center: getSelectOptionText(centerSelect),
    unit: getSelectOptionText(unitSelect),
    location: getSelectOptionText(locationSelect),
    depreciationCostCenter: getSelectOptionText(depreciationSelect),
    category: getSelectOptionText(document.getElementById('category')),
    investmentType: getSelectOptionText(document.getElementById('investmentType')),
    assetType: getSelectOptionText(document.getElementById('assetType')),
    kpiType: getSelectOptionText(document.getElementById('kpiType'))
  };
}

function buildApprovalSummary(projectId, projectData = {}) {
  const simplePeps = collectSimplePepDataForSummary();
  const milestones = collectMilestonesForSummary();
  const activities = collectActivitiesForSummary(milestones);
  const peps = collectPepEntriesForSummary(simplePeps, activities);
  const displayValues = collectProjectDisplayValues();

  const numericId = Number(projectId);
  const resolvedId = Number.isFinite(numericId) ? numericId : projectId;

  return {
    project: {
      id: resolvedId,
      ...projectData,
      displayValues
    },
    milestones,
    activities,
    peps
  };
}

async function persistSimplePeps(projectId, approvalYear) {
  const currentIds = new Set();

  for (const row of simplePepList.querySelectorAll('.pep-row')) {
    const id = row.dataset.pepId;
    const title = row.querySelector('.pep-title').value.trim();
    const amount = parseNumericInputValue(row.querySelector('.pep-amount')) || 0;
    const year = parseNumber(row.querySelector('.pep-year').value) || approvalYear;
    const payload = {
      Title: title,
      amountBrl: amount,
      year,
      projectsIdId: projectId
    };
    if (id) {
      await sp.updateItem('Peps', Number(id), payload);
      currentIds.add(Number(id));
    } else {
      const created = await sp.createItem('Peps', payload);
      currentIds.add(Number(created.Id));
    }
  }

  const toDelete = [...state.editingSnapshot.simplePeps].filter((id) => !currentIds.has(id));
  for (const id of toDelete) {
    await sp.deleteItem('Peps', Number(id));
  }
}

async function cleanupSimplePeps() {
  for (const id of state.editingSnapshot.simplePeps) {
    await sp.deleteItem('Peps', Number(id));
  }
  state.editingSnapshot.simplePeps.clear();
}

async function persistKeyProjects(projectId) {
  const milestoneIds = new Set();
  const activityIds = new Set();
  const activityPepIds = new Set();

  for (const milestone of milestoneList.querySelectorAll('.milestone')) {
    const id = milestone.dataset.milestoneId;
    const title = milestone.querySelector('.milestone-title').value.trim();
    const payload = {
      Title: title,
      projectsIdId: projectId
    };
    let milestoneId = Number(id);
    if (id) {
      await sp.updateItem('Milestones', milestoneId, payload);
    } else {
      const created = await sp.createItem('Milestones', payload);
      milestoneId = created.Id;
      milestone.dataset.milestoneId = milestoneId;
    }
    milestoneIds.add(Number(milestoneId));

    for (const activity of milestone.querySelectorAll('.activity')) {
      const activityIdRaw = activity.dataset.activityId;
      const activityPayload = {
        Title: activity.querySelector('.activity-title').value.trim(),
        startDate: activity.querySelector('.activity-start').value || null,
        endDate: activity.querySelector('.activity-end').value || null,
        activityDescription: activity.querySelector('.activity-description').value.trim(),
        supplier: activity.querySelector('.activity-supplier').value.trim(),
        projectsIdId: projectId,
        milestonesIdId: milestoneId
      };
      let activityId = Number(activityIdRaw);
      if (activityIdRaw) {
        await sp.updateItem('Activities', activityId, activityPayload);
      } else {
        const createdActivity = await sp.createItem('Activities', activityPayload);
        activityId = createdActivity.Id;
        activity.dataset.activityId = activityId;
      }
      activityIds.add(Number(activityId));

      const pepTitle = activity.querySelector('.activity-pep-title')?.value.trim() || '';
      const pepAmount = parseNumericInputValue(activity.querySelector('.activity-pep-amount')) || 0;
      const pepYearInput = activity.querySelector('.activity-pep-year');
      let pepYear = parseNumber(pepYearInput?.value);
      if (!pepYear) {
        const startValue = activity.querySelector('.activity-start')?.value;
        const startYear = startValue ? parseInt(startValue.substring(0, 4), 10) : NaN;
        if (Number.isFinite(startYear)) {
          pepYear = startYear;
        } else {
          const fallback = parseNumber(approvalYearInput.value);
          if (fallback) {
            pepYear = fallback;
          }
        }
        if (pepYearInput) {
          pepYearInput.value = pepYear ?? '';
        }
      }

      const pepIdRaw = activity.dataset.pepId;
      const hasPepData = Boolean(pepTitle) || pepAmount > 0;

      if (hasPepData) {
        const pepPayload = {
          Title: pepTitle,
          amountBrl: pepAmount,
          year: pepYear,
          projectsIdId: projectId,
          activitiesIdId: activityId
        };
        let pepId = Number(pepIdRaw);
        if (pepIdRaw) {
          await sp.updateItem('Peps', pepId, pepPayload);
        } else {
          const createdPep = await sp.createItem('Peps', pepPayload);
          pepId = createdPep.Id;
          activity.dataset.pepId = pepId;
        }
        activityPepIds.add(Number(pepId));
      } else if (pepIdRaw) {
        await sp.deleteItem('Peps', Number(pepIdRaw));
        activity.dataset.pepId = '';
      }
    }
  }

  await deleteMissing('Peps', state.editingSnapshot.activityPeps, activityPepIds);
  await deleteMissing('Activities', state.editingSnapshot.activities, activityIds);
  await deleteMissing('Milestones', state.editingSnapshot.milestones, milestoneIds);
}

async function cleanupKeyProjects() {
  await deleteMissing('Peps', state.editingSnapshot.activityPeps, new Set());
  await deleteMissing('Activities', state.editingSnapshot.activities, new Set());
  await deleteMissing('Milestones', state.editingSnapshot.milestones, new Set());
}

async function deleteMissing(listName, previousSet, currentSet) {
  for (const id of previousSet) {
    if (!currentSet.has(id)) {
      await sp.deleteItem(listName, Number(id));
    }
  }
  previousSet.clear();
}

// ============================================================================
// Utilitários
// ============================================================================
function scrollFormToTop() {
  const scrollElement = (element) => {
    if (!element) return;
    if (typeof element.scrollTo === 'function') {
      element.scrollTo(0, 0);
    } else {
      element.scrollTop = 0;
    }
  };

  scrollElement(overlay);
  scrollElement(projectForm);
}

function showStatus(message, options = {}) {
  if (!formStatus) return;
  if (message === null || message === undefined || message === '') {
    resetFormStatus();
    return;
  }

  let tone = 'info';

  if (typeof options === 'boolean') {
    tone = options ? 'success' : 'info';
  } else if (typeof options === 'string') {
    tone = options;
  } else if (typeof options === 'object' && options !== null) {
    if (options.type) {
      tone = options.type;
    } else if (typeof options.success === 'boolean') {
      tone = options.success ? 'success' : 'info';
    }
  }

  if (!['info', 'success', 'warning', 'error'].includes(tone)) {
    tone = 'info';
  }

  formStatus.textContent = message;
  formStatus.classList.add('show');
  formStatus.classList.remove('feedback--info', 'feedback--success', 'feedback--warning', 'feedback--error');
  formStatus.classList.add(`feedback--${tone}`);
}

function statusColor(status) {
  switch (status) {
    case 'Rascunho':
      return '#414141';
    case 'Em Aprovação':
      return '#970886';
    case 'Reprovado para Revisão':
      return '#fe8f46';
    case 'Aprovado':
      return '#3d9308';
    case 'Reprovado':
      return '#f83241';
    default:
      return '#414141';
  }
}

function parseNumber(value) {
  const number = parseInt(value, 10);
  return Number.isFinite(number) ? number : null;
}

// ============================================================================
// Execução
// ============================================================================
init();
