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

  async request(url, options = {}) {
    const response = await fetch(url, options);
    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || response.statusText);
    }
    if (response.status === 204) {
      return null;
    }
    const text = await response.text();
    return text ? JSON.parse(text) : null;
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
  activityDates: null
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
const sendApprovalBtn = document.getElementById('sendApprovalBtn');
const saveProjectBtn = document.getElementById('saveProjectBtn');
const formStatus = document.getElementById('formStatus');
const formErrors = document.getElementById('formErrors');
const statusField = document.getElementById('statusField');
const budgetHint = document.getElementById('budgetHint');
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

const approvalYearInput = document.getElementById('approvalYear');
const projectBudgetInput = document.getElementById('projectBudget');
const projectStartDateInput = document.getElementById('startDate');
const projectEndDateInput = document.getElementById('endDate');

const simplePepTemplate = document.getElementById('simplePepTemplate');
const milestoneTemplate = document.getElementById('milestoneTemplate');
const activityTemplate = document.getElementById('activityTemplate');

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
      const amountRaw = parseFloat(activityEl.querySelector('.activity-pep-amount')?.value);
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

function drawGantt(milestones) {
  if (!ganttContainer || !ganttChartEl || !window.google?.visualization) return;

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

  const rows = [];
  let idCounter = 0;

  milestones.forEach((milestone) => {
    if (!milestone) return;
    idCounter += 1;
    let msStart = null;
    let msEnd = null;
    const activityRows = [];
    const activities = Array.isArray(milestone.atividades) ? milestone.atividades : [];

    activities.forEach((activity, index) => {
      const startDate = activity.inicio ? new Date(activity.inicio) : new Date();
      const endDate = activity.fim ? new Date(activity.fim) : new Date(startDate.getTime() + 1000 * 60 * 60 * 24);
      if (!msStart || startDate < msStart) msStart = startDate;
      if (!msEnd || endDate > msEnd) msEnd = endDate;

      const totalCapex = activity.anual.reduce((total, year) => total + (year.capex_brl || 0), 0);
      const tooltipLines = activity.anual.map((year) => {
        const yearLabel = year.ano ?? 'Ano não informado';
        const description = year.descricao ? ` - ${year.descricao}` : '';
        return `${yearLabel}: ${BRL.format(year.capex_brl || 0)}${description}`;
      });

      activityRows.push([
        `ms-${idCounter}-${index}`,
        activity.titulo || `Atividade ${index + 1}`,
        'Atividade',
        startDate,
        endDate,
        null,
        0,
        `ms-${idCounter}`,
        `CAPEX total: ${BRL.format(totalCapex)}${tooltipLines.length ? `<br/>${tooltipLines.join('<br/>')}` : ''}`
      ]);
    });

    if (msStart && msEnd) {
      rows.push([
        `ms-${idCounter}`,
        milestone.nome,
        'milestone',
        msStart,
        msEnd,
        null,
        0,
        null,
        milestone.nome
      ]);
    }

    rows.push(...activityRows);
  });

  if (!rows.length) {
    ganttChartEl.innerHTML = '';
    ganttContainer.classList.add('hidden');
    return;
  }

  ganttContainer.classList.remove('hidden');
  if (ganttTitleEl) {
    ganttTitleEl.classList.remove('hidden');
  }

  data.addRows(rows);
  const chart = new google.visualization.Gantt(ganttChartEl);
  chart.draw(data, {
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
      ]
    }
  });
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
  bindEvents();
  setApprovalYearToCurrent();
  loadProjects();
  initGantt();
  window.addEventListener('load', initGantt, { once: true });
}

function bindEvents() {
  newProjectBtn.addEventListener('click', () => openProjectForm('create'));
  closeFormBtn.addEventListener('click', closeForm);
  // Habilita o fechamento do formulário pela tecla ESC.
  document.addEventListener('keydown', handleOverlayEscape);
  if (projectSearch) {
    projectSearch.addEventListener('input', () => renderProjectList());
  }

  projectForm.addEventListener('submit', handleFormSubmit);
  projectForm.addEventListener('focusin', handleFormFocusCapture);
  if (saveProjectBtn) {
    saveProjectBtn.addEventListener('click', (event) => {
      event.preventDefault();
      openSummaryOverlay('save', saveProjectBtn);
    });
  }

  if (sendApprovalBtn) {
    sendApprovalBtn.addEventListener('click', (event) => {
      event.preventDefault();
      openSummaryOverlay('approval', sendApprovalBtn);
    });
  }

  if (summaryConfirmBtn) {
    summaryConfirmBtn.addEventListener('click', handleSummaryConfirm);
  }

  if (summaryEditBtn) {
    summaryEditBtn.addEventListener('click', () => closeSummaryOverlay());
  }

  projectBudgetInput.addEventListener('input', () => {
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

  const statusKey = project.status || '';
  const canEditAndSend = ['Rascunho', 'Reprovado para Revisão'].includes(statusKey);
  const viewOnlyStatuses = ['Aprovado', 'Reprovado', 'Em Aprovação'];

  if (viewOnlyStatuses.includes(statusKey)) {
    const viewBtn = document.createElement('button');
    viewBtn.type = 'button';
    viewBtn.className = 'btn ghost';
    viewBtn.textContent = 'Visualizar Projeto';
    viewBtn.addEventListener('click', () => openProjectForm('edit', detail));
    actions.append(viewBtn);
  } else if (canEditAndSend) {
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn primary';
    editBtn.textContent = 'Editar Projeto';
    editBtn.addEventListener('click', () => openProjectForm('edit', detail));
    actions.append(editBtn);

    const approveBtn = document.createElement('button');
    approveBtn.type = 'button';
    approveBtn.className = 'btn accent';
    approveBtn.textContent = 'Enviar para Aprovação';
    approveBtn.addEventListener('click', () => {
      openProjectForm('edit', detail);
      requestAnimationFrame(() => {
        sendApprovalBtn?.focus();
      });
    });
    actions.append(approveBtn);
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

// ============================================================================
// Formulário: abertura, preenchimento e coleta dos dados
// ============================================================================
function openProjectForm(mode, detail = null) {
  projectForm.reset();
  formStatus.classList.remove('show');
  resetValidationState();
  projectForm.dataset.mode = mode;
  projectForm.dataset.action = 'save';
  projectForm.dataset.projectId = detail?.project?.Id || '';

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
    sendApprovalBtn.classList.remove('hidden');
    setApprovalYearToCurrent();
    updateBudgetSections({ clear: true });
  } else if (detail) {
    fillFormWithProject(detail);
  }

  updateSimplePepYears();
  overlay.classList.remove('hidden');
  queueGanttRefresh();
}

function fillFormWithProject(detail) {
  const { project, simplePeps, milestones, activities, activityPeps } = detail;
  formTitle.textContent = `Editar Projeto #${project.Id}`;
  statusField.value = project.status || 'Rascunho';
  sendApprovalBtn.classList.toggle('hidden', !['Rascunho', 'Reprovado para Revisão'].includes(project.status));

  document.getElementById('projectName').value = project.Title || '';
  document.getElementById('category').value = project.category || '';
  document.getElementById('investmentType').value = project.investmentType || '';
  document.getElementById('assetType').value = project.assetType || '';
  document.getElementById('projectFunction').value = project.projectFunction || '';

  document.getElementById('approvalYear').value = project.approvalYear || '';
  document.getElementById('startDate').value = project.startDate ? project.startDate.substring(0, 10) : '';
  document.getElementById('endDate').value = project.endDate ? project.endDate.substring(0, 10) : '';

  document.getElementById('projectBudget').value = project.budgetBrl ?? '';
  document.getElementById('investmentLevel').value = project.investmentLevel || '';
  document.getElementById('fundingSource').value = project.fundingSource || '';
  document.getElementById('depreciationCostCenter').value = project.depreciationCostCenter || '';

  document.getElementById('company').value = project.company || '';
  document.getElementById('center').value = project.center || '';
  document.getElementById('unit').value = project.unit || '';
  document.getElementById('location').value = project.location || '';

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
}

function closeForm() {
  overlay.classList.add('hidden');
  closeSummaryOverlay({ restoreFocus: false });
}

function openSummaryOverlay(action, triggerButton = null) {
  const normalizedAction = action === 'approval' ? 'approval' : 'save';
  projectForm.dataset.action = normalizedAction;

  const formValid = projectForm.reportValidity();
  const pepValid = validatePepBudget();
  const activityValid = validateActivityDates();

  if (!formValid || !pepValid || !activityValid) {
    return;
  }

  if (!summaryOverlay) {
    projectForm.requestSubmit();
    return;
  }

  summaryTriggerButton = triggerButton || null;
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

  summarySections.innerHTML = '';

  const sections = [
    {
      title: 'Dados Gerais',
      entries: [
        { label: 'Nome do Projeto', value: getFieldDisplayValue('projectName') },
        { label: 'Categoria', value: getFieldDisplayValue('category') },
        { label: 'Tipo de Investimento', value: getFieldDisplayValue('investmentType') },
        { label: 'Tipo de Ativo', value: getFieldDisplayValue('assetType') },
        { label: 'Função do Projeto', value: getFieldDisplayValue('projectFunction') }
      ]
    },
    {
      title: 'Planejamento Temporal',
      entries: [
        { label: 'Ano de Aprovação', value: getFieldDisplayValue('approvalYear') },
        { label: 'Data de Início', value: formatDateValue(document.getElementById('startDate')?.value) },
        { label: 'Data de Término', value: formatDateValue(document.getElementById('endDate')?.value) }
      ]
    },
    {
      title: 'Orçamento e Recursos',
      entries: [
        { label: 'Orçamento do Projeto', value: formatCurrencyField('projectBudget') },
        { label: 'Nível de Investimento', value: getFieldDisplayValue('investmentLevel') },
        { label: 'Origem da Verba', value: getFieldDisplayValue('fundingSource') },
        { label: 'C Custo Depreciação', value: getFieldDisplayValue('depreciationCostCenter') }
      ]
    },
    {
      title: 'Localização Operacional',
      entries: [
        { label: 'Empresa', value: getFieldDisplayValue('company') },
        { label: 'Centro', value: getFieldDisplayValue('center') },
        { label: 'Unidade', value: getFieldDisplayValue('unit') },
        { label: 'Localização', value: getFieldDisplayValue('location') }
      ]
    },
    {
      title: 'Pessoas Envolvidas',
      entries: [
        { label: 'Solicitante', value: getFieldDisplayValue('projectUser') },
        { label: 'Gestor do Projeto', value: getFieldDisplayValue('projectLeader') }
      ]
    },
    {
      title: 'Justificativa e Solução',
      entries: [
        { label: 'Necessidade do Negócio', value: getFieldDisplayValue('businessNeed') },
        { label: 'Solução Proposta', value: getFieldDisplayValue('proposedSolution') }
      ]
    },
    {
      title: 'Indicadores (KPIs)',
      entries: [
        { label: 'Tipo de KPI', value: getFieldDisplayValue('kpiType') },
        { label: 'Nome do KPI', value: getFieldDisplayValue('kpiName') },
        { label: 'Descrição do KPI', value: getFieldDisplayValue('kpiDescription') },
        { label: 'Valor Atual', value: formatNumberField('kpiCurrent') },
        { label: 'Valor Esperado', value: formatNumberField('kpiExpected') }
      ]
    }
  ];

  sections.forEach((section) => createSummarySection(section.title, section.entries));

  renderPepSummary();
  renderMilestoneSummary();
  populateSummaryGantt({ refreshFirst: true });
}

function createSummarySection(title, entries = []) {
  if (!summarySections || !entries.length) return;

  const section = document.createElement('section');
  section.className = 'summary-section';

  const heading = document.createElement('h3');
  heading.textContent = title;
  section.appendChild(heading);

  const list = document.createElement('dl');
  list.className = 'summary-list';

  entries.forEach((entry) => {
    if (!entry?.label) return;
    const dt = document.createElement('dt');
    dt.textContent = entry.label;
    const dd = document.createElement('dd');
    dd.textContent = resolveSummaryValue(entry.value);
    list.append(dt, dd);
  });

  section.appendChild(list);
  summarySections.appendChild(section);
}

function renderPepSummary() {
  if (!summarySections) return;
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
  heading.textContent = 'PEPs';
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
  summarySections.appendChild(section);
}

function renderMilestoneSummary() {
  if (!summarySections || !milestoneList) return;
  if (keyProjectSection.classList.contains('hidden')) return;

  const milestones = milestoneList.querySelectorAll('.milestone');
  if (!milestones.length) return;

  const section = document.createElement('section');
  section.className = 'summary-section';

  const heading = document.createElement('h3');
  heading.textContent = 'Marcos e Atividades';
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

        const detailList = document.createElement('dl');
        detailList.className = 'summary-list summary-list--activity';

        const detailItems = [
          { label: 'Período', value: buildActivityPeriod(activity) },
          { label: 'Valor da Atividade', value: formatCurrencyValueFromElement(activity.querySelector('.activity-pep-amount')) },
          { label: 'Elemento PEP', value: getSelectOptionText(activity.querySelector('.activity-pep-title')) },
          { label: 'Ano do PEP', value: activity.querySelector('.activity-pep-year')?.value ?? '' },
          { label: 'Fornecedor', value: activity.querySelector('.activity-supplier')?.value ?? '' },
          { label: 'Descrição', value: activity.querySelector('.activity-description')?.value ?? '' }
        ];

        detailItems.forEach((item) => {
          const dt = document.createElement('dt');
          dt.textContent = item.label;
          const dd = document.createElement('dd');
          dd.textContent = resolveSummaryValue(item.value);
          detailList.append(dt, dd);
        });

        activityCard.appendChild(detailList);
        activityContainer.appendChild(activityCard);
      });

      card.appendChild(activityContainer);
    }

    wrapper.appendChild(card);
  });

  section.appendChild(wrapper);
  summarySections.appendChild(section);
}

function populateSummaryGantt(options = {}) {
  const { refreshFirst = false } = options;
  if (!summaryGanttSection || !summaryGanttChart) return;

  if (refreshFirst) {
    refreshGantt();
    requestAnimationFrame(() => populateSummaryGantt());
    return;
  }

  if (keyProjectSection.classList.contains('hidden')) {
    summaryGanttSection.classList.add('hidden');
    summaryGanttChart.innerHTML = '';
    return;
  }

  const chartHtml = ganttChartEl?.innerHTML?.trim();
  if (!chartHtml) {
    summaryGanttSection.classList.add('hidden');
    summaryGanttChart.innerHTML = '';
    return;
  }

  summaryGanttSection.classList.remove('hidden');
  summaryGanttChart.innerHTML = chartHtml;
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
 * Fecha o formulário quando o usuário pressiona ESC.
 * Ao detectar a tecla, confirma com o usuário antes de fechar o formulário.
 * A verificação garante que a overlay exista e esteja visível,
 * evitando erros quando o formulário já está fechado.
 */
function handleOverlayEscape(event) {
  if (event.key !== 'Escape') return;
  if (summaryOverlay && !summaryOverlay.classList.contains('hidden')) {
    closeSummaryOverlay();
    return;
  }
  if (!overlay || overlay.classList.contains('hidden')) return;

  const shouldClose = window.confirm(
    'Tem certeza que deseja fechar o formulário? Suas alterações não salvas serão perdidas.'
  );

  if (shouldClose) {
    closeForm();
  }
}

function updateBudgetSections(options = {}) {
  const { preserve = false, clear = false } = options;
  const value = parseFloat(projectBudgetInput.value.replace(',', '.'));
  const isNumber = !Number.isNaN(value) && Number.isFinite(value);
  setSectionInteractive(simplePepSection, false);
  setSectionInteractive(keyProjectSection, false);

  if (!isNumber) {
    simplePepSection.classList.add('hidden');
    keyProjectSection.classList.add('hidden');
    budgetHint.textContent = '';
    if (clear) {
      simplePepList.innerHTML = '';
      milestoneList.innerHTML = '';
    }
    return;
  }

  if (value >= BUDGET_THRESHOLD) {
    simplePepSection.classList.add('hidden');
    keyProjectSection.classList.remove('hidden');
    setSectionInteractive(keyProjectSection, true);
    budgetHint.textContent = 'Projeto classificado como KEY Project (>= R$ 1.000.000,00).';
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
    budgetHint.textContent = 'Projeto com orçamento inferior a R$ 1.000.000,00.';
    if (!preserve) {
      milestoneList.innerHTML = '';
    }
    if (!simplePepList.children.length) {
      ensureSimplePepRow();
    }
  }

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
function clearFormErrorMessage() {
  formErrors.textContent = '';
  formErrors.classList.remove('show');
  delete formErrors.dataset.validation;
}

function resetValidationState() {
  validationState.pepBudget = null;
  validationState.activityDates = null;
  clearFormErrorMessage();
}

function setValidationError(key, message) {
  validationState[key] = message || null;
  const firstMessage = validationState.pepBudget || validationState.activityDates;
  if (firstMessage) {
    showError(firstMessage);
    formErrors.dataset.validation = 'true';
  } else if (formErrors.dataset.validation === 'true') {
    clearFormErrorMessage();
  }
}

function rememberFieldPreviousValue(element) {
  if (!element || typeof element !== 'object') return;
  if (!('dataset' in element)) return;
  element.dataset.previousValue = element.value ?? '';
}

function parseNumericInputValue(source) {
  if (!source) return 0;
  const rawValue = typeof source === 'string' ? source : source.value;
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return 0;
  }
  const normalized = String(rawValue).replace(',', '.');
  const number = Number.parseFloat(normalized);
  return Number.isFinite(number) ? number : 0;
}

function getProjectBudgetValue() {
  if (!projectBudgetInput) return NaN;
  const rawValue = projectBudgetInput.value;
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return NaN;
  }
  return parseNumericInputValue(projectBudgetInput);
}

function getPepAmountInputs() {
  const simplePepInputs = Array.from(simplePepList.querySelectorAll('.pep-amount'));
  const activityPepInputs = Array.from(milestoneList.querySelectorAll('.activity-pep-amount'));
  return [...simplePepInputs, ...activityPepInputs];
}

function calculatePepTotal() {
  return getPepAmountInputs().reduce((sum, input) => sum + parseNumericInputValue(input), 0);
}

function validatePepBudget(options = {}) {
  const { changedInput = null } = options;
  const budget = getProjectBudgetValue();

  if (!Number.isFinite(budget)) {
    setValidationError('pepBudget', null);
    return true;
  }

  const total = calculatePepTotal();
  if (total - budget > 0.009) {
    const message = `A soma dos PEPs (${BRL.format(total)}) ultrapassa o orçamento do projeto (${BRL.format(budget)}).`;
    setValidationError('pepBudget', message);

    if (changedInput) {
      const previousValue = changedInput.dataset?.previousValue ?? '';
      if (changedInput.value !== previousValue) {
        changedInput.value = previousValue;
      }
    }
    return false;
  }

  setValidationError('pepBudget', null);
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

function validateActivityDates(options = {}) {
  const { changedInput = null } = options;
  const projectStart = parseDateInputValue(projectStartDateInput?.value);
  const projectEnd = parseDateInputValue(projectEndDateInput?.value);

  if (!projectStart && !projectEnd) {
    setValidationError('activityDates', null);
    return true;
  }

  let invalidMessage = null;
  let invalidField = null;

  const activities = milestoneList.querySelectorAll('.activity');
  for (const activity of activities) {
    const startInput = activity.querySelector('.activity-start');
    const endInput = activity.querySelector('.activity-end');
    const title = activity.querySelector('.activity-title')?.value?.trim() || 'Atividade';
    const startDate = parseDateInputValue(startInput?.value);
    const endDate = parseDateInputValue(endInput?.value);

    if (projectStart && startDate && startDate < projectStart) {
      invalidMessage = `A data de início da atividade "${title}" não pode ser anterior à data de início do projeto.`;
      invalidField = changedInput === projectStartDateInput ? projectStartDateInput : startInput;
      break;
    }

    if (projectEnd && endDate && endDate > projectEnd) {
      invalidMessage = `A data de término da atividade "${title}" não pode ser posterior à data de término do projeto.`;
      invalidField = changedInput === projectEndDateInput ? projectEndDateInput : endInput;
      break;
    }
  }

  if (invalidMessage) {
    setValidationError('activityDates', invalidMessage);
    if (invalidField) {
      const previousValue = invalidField.dataset?.previousValue ?? '';
      if (invalidField.value !== previousValue) {
        invalidField.value = previousValue;
      }
    }
    return false;
  }

  setValidationError('activityDates', null);
  if (changedInput) {
    rememberFieldPreviousValue(changedInput);
  }
  return true;
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
  row.querySelector('.pep-amount').value = amount ?? '';
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
    amountInput.value = data.pepAmount ?? '';
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
  queueGanttRefresh();
  return activity;
}

// ============================================================================
// Envio do formulário e persistência (CRUD)
// ============================================================================
async function handleFormSubmit(event) {
  event.preventDefault();
  const mode = projectForm.dataset.mode;
  const action = projectForm.dataset.action || 'save';
  const projectId = projectForm.dataset.projectId;

  if (!projectForm.reportValidity()) {
    return;
  }

  const pepValid = validatePepBudget();
  const activityDatesValid = validateActivityDates();
  if (!pepValid || !activityDatesValid) {
    return;
  }

  const baseStatus = statusField.value || 'Rascunho';
  const nextStatus = action === 'approval' ? 'Em Aprovação' : baseStatus;
  statusField.value = mode === 'create' ? (action === 'approval' ? 'Em Aprovação' : 'Rascunho') : nextStatus;
  projectForm.dataset.action = 'save';

  const payload = collectProjectData();
  payload.status = statusField.value;

  showStatus('Salvando projeto…');

  try {
    let savedProjectId = projectId;
    if (mode === 'create') {
      const result = await sp.createItem('Projects', payload);
      savedProjectId = result?.Id;
    } else {
      await sp.updateItem('Projects', Number(projectId), payload);
    }

    await persistRelatedRecords(Number(savedProjectId || projectId), payload);

    const resolvedId = Number(savedProjectId || projectId);
    if (resolvedId) {
      updateProjectState(resolvedId, {
        Title: payload.Title,
        status: payload.status,
        budgetBrl: payload.budgetBrl
      });
      renderProjectList();
      if (state.currentDetails?.project?.Id === resolvedId) {
        state.currentDetails = {
          ...state.currentDetails,
          project: {
            ...state.currentDetails.project,
            Title: payload.Title,
            status: payload.status,
            budgetBrl: payload.budgetBrl
          }
        };
        renderProjectDetails(state.currentDetails);
      }
    }

    showStatus('Projeto salvo com sucesso.', true);
    await loadProjects();
    if (savedProjectId) {
      await selectProject(Number(savedProjectId));
    }
    closeForm();
  } catch (error) {
    console.error('Erro ao salvar projeto', error);
    showError('Não foi possível salvar o projeto. Verifique os dados e tente novamente.');
  }
}

function collectProjectData() {
  const data = {
    Title: document.getElementById('projectName').value.trim(),
    category: document.getElementById('category').value.trim(),
    investmentType: document.getElementById('investmentType').value.trim(),
    assetType: document.getElementById('assetType').value.trim(),
    projectFunction: document.getElementById('projectFunction').value.trim(),
    approvalYear: parseNumber(document.getElementById('approvalYear').value),
    startDate: document.getElementById('startDate').value || null,
    endDate: document.getElementById('endDate').value || null,
    budgetBrl: parseFloat(document.getElementById('projectBudget').value) || 0,
    investmentLevel: document.getElementById('investmentLevel').value.trim(),
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
    kpiCurrent: parseFloatOrNull(document.getElementById('kpiCurrent').value),
    kpiExpected: parseFloatOrNull(document.getElementById('kpiExpected').value)
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

async function persistSimplePeps(projectId, approvalYear) {
  const currentIds = new Set();

  for (const row of simplePepList.querySelectorAll('.pep-row')) {
    const id = row.dataset.pepId;
    const title = row.querySelector('.pep-title').value.trim();
    const amount = parseFloat(row.querySelector('.pep-amount').value) || 0;
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
      const pepAmount = parseFloat(activity.querySelector('.activity-pep-amount')?.value) || 0;
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
          projectsId: projectId,
          activitiesId: activityId
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
function showStatus(message, success = false) {
  formStatus.textContent = message;
  formStatus.classList.add('show');
  formStatus.classList.toggle('error', !success && message.toLowerCase().includes('erro'));
}

function showError(message) {
  formErrors.textContent = message;
  formErrors.classList.add('show');
  delete formErrors.dataset.validation;
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

function parseFloatOrNull(value) {
  const number = parseFloat(value);
  return Number.isFinite(number) ? number : null;
}

// ============================================================================
// Execução
// ============================================================================
init();
