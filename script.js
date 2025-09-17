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

const SITE_URL = window.SHAREPOINT_SITE_URL || 'https://<seu-tenant>.sharepoint.com/sites/<seu-site>';
const sp = new SharePointService(SITE_URL);

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

const newProjectBtn = document.getElementById('newProjectBtn');
const projectSearch = document.getElementById('projectSearch');
const projectList = document.getElementById('projectList');
const projectDetails = document.getElementById('projectDetails');
const overlay = document.getElementById('formOverlay');
const projectForm = document.getElementById('projectForm');
const formTitle = document.getElementById('formTitle');
const closeFormBtn = document.getElementById('closeFormBtn');
const sendApprovalBtn = document.getElementById('sendApprovalBtn');
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

const approvalYearInput = document.getElementById('approvalYear');
const projectBudgetInput = document.getElementById('projectBudget');

const simplePepTemplate = document.getElementById('simplePepTemplate');
const milestoneTemplate = document.getElementById('milestoneTemplate');
const activityTemplate = document.getElementById('activityTemplate');
const activityPepTemplate = document.getElementById('activityPepTemplate');

// ============================================================================
// Inicialização
// ============================================================================
function init() {
  bindEvents();
  const currentYear = new Date().getFullYear();
  approvalYearInput.value = currentYear;
  approvalYearInput.max = currentYear;
  loadProjects();
}

function bindEvents() {
  newProjectBtn.addEventListener('click', () => openProjectForm('create'));
  closeFormBtn.addEventListener('click', closeForm);
  projectSearch.addEventListener('input', () => renderProjectList());

  projectForm.addEventListener('submit', handleFormSubmit);
  sendApprovalBtn.addEventListener('click', () => {
    projectForm.dataset.action = 'approval';
    projectForm.requestSubmit();
  });

  projectBudgetInput.addEventListener('input', () => updateBudgetSections());
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
    }
  });

  milestoneList.addEventListener('click', (event) => {
    if (event.target.classList.contains('remove-milestone')) {
      event.target.closest('.milestone')?.remove();
      return;
    }
    if (event.target.classList.contains('add-activity')) {
      const milestone = event.target.closest('.milestone');
      addActivityBlock(milestone);
      return;
    }
    if (event.target.classList.contains('remove-activity')) {
      event.target.closest('.activity')?.remove();
      return;
    }
    if (event.target.classList.contains('add-activity-pep')) {
      const activity = event.target.closest('.activity');
      addActivityPepRow(activity);
      return;
    }
    if (event.target.classList.contains('remove-activity-pep')) {
      event.target.closest('.activity-pep')?.remove();
    }
  });
}

// ============================================================================
// Carregamento e renderização da lista de projetos
// ============================================================================
async function loadProjects() {
  try {
    const results = await sp.getItems('Projects', { orderby: 'Created desc' });
    state.projects = results;
    renderProjectList();
  } catch (error) {
    console.error('Erro ao carregar projetos', error);
  }
}

function renderProjectList() {
  const filter = projectSearch.value.toLowerCase();
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

    const title = document.createElement('h3');
    title.textContent = item.Title || 'Projeto sem título';

    const status = document.createElement('span');
    status.className = 'status';
    status.style.background = statusColor(item.status);
    status.textContent = item.status || 'Sem status';

    const info = document.createElement('p');
    const budget = item.budgetBrl ? ` • ${BRL.format(item.budgetBrl)}` : '';
    info.textContent = `${item.approvalYear || ''}${budget}`.trim();

    card.append(title, status, info);
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
      sp.getItems('Milestones', { filter: `projectsId eq ${projectId}` }),
      sp.getItems('Activities', { filter: `projectsId eq ${projectId}` }),
      sp.getItems('Peps', { filter: `projectsId eq ${projectId}` })
    ]);

    const detail = {
      project,
      milestones,
      activities,
      peps,
      simplePeps: peps.filter((pep) => !pep.activitiesId),
      activityPeps: peps.filter((pep) => pep.activitiesId)
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

  const { project, milestones, activities, simplePeps, activityPeps } = detail;
  if (project.status === 'Reprovado') {
    return;
  }

  const wrapper = document.createElement('div');

  const header = document.createElement('div');
  header.className = 'details-header';
  const title = document.createElement('h2');
  title.textContent = project.Title || 'Projeto sem título';
  const status = document.createElement('span');
  status.className = 'status';
  status.style.background = statusColor(project.status);
  status.textContent = project.status || 'Sem status';
  header.append(title, status);

  if (['Rascunho', 'Reprovado para Revisão'].includes(project.status)) {
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn primary';
    editBtn.textContent = 'Editar Projeto';
    editBtn.addEventListener('click', () => openProjectForm('edit', detail));
    header.append(editBtn);
  } else if (project.status === 'Aprovado') {
    const info = document.createElement('p');
    info.className = 'hint';
    info.textContent = 'Projeto aprovado - somente leitura.';
    header.append(info);
  }

  wrapper.append(header);

  const overview = document.createElement('div');
  overview.className = 'details-grid';

  overview.append(
    createDetailBox('Ano de Aprovação', project.approvalYear),
    createDetailBox('Orçamento (R$)', project.budgetBrl ? BRL.format(project.budgetBrl) : ''),
    createDetailBox('Nível de Investimento', project.investmentLevel),
    createDetailBox('Origem da Verba', project.fundingSource)
  );

  overview.append(
    createDetailBox('Empresa', project.company),
    createDetailBox('Centro', project.center),
    createDetailBox('Unidade', project.unit),
    createDetailBox('Local de Implantação', project.location)
  );

  overview.append(
    createDetailBox('Project User', project.projectUser),
    createDetailBox('Coordenador do Projeto', project.projectLeader),
    createDetailBox('Período', formatPeriod(project.startDate, project.endDate))
  );

  wrapper.append(overview);

  wrapper.append(createTextSection('Sumário do Projeto', project.businessNeed));
  wrapper.append(createTextSection('Comentário', project.proposedSolution));

  const kpiSection = document.createElement('div');
  kpiSection.className = 'detail-box';
  const kpiTitle = document.createElement('h4');
  kpiTitle.textContent = 'Indicadores de Desempenho';
  const kpiContent = document.createElement('p');
  const pieces = [
    project.kpiType ? `Tipo: ${project.kpiType}` : '',
    project.kpiName ? `Nome: ${project.kpiName}` : '',
    project.kpiDescription ? `Descrição: ${project.kpiDescription}` : '',
    project.kpiCurrent !== null && project.kpiCurrent !== undefined ? `Atual: ${project.kpiCurrent}` : '',
    project.kpiExpected !== null && project.kpiExpected !== undefined ? `Esperado: ${project.kpiExpected}` : ''
  ].filter(Boolean);
  kpiContent.innerHTML = pieces.join('<br>') || 'Sem indicadores informados.';
  kpiSection.append(kpiTitle, kpiContent);
  wrapper.append(kpiSection);

  if (project.budgetBrl < BUDGET_THRESHOLD && simplePeps.length) {
    wrapper.append(createPepSection(simplePeps));
  }

  if (project.budgetBrl >= BUDGET_THRESHOLD && milestones.length) {
    wrapper.append(createKeyProjectsSection(milestones, activities, activityPeps));
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

function createDetailBox(label, value) {
  const box = document.createElement('div');
  box.className = 'detail-box';
  const title = document.createElement('h4');
  title.textContent = label;
  const text = document.createElement('p');
  text.textContent = value || '—';
  box.append(title, text);
  return box;
}

function createTextSection(label, content) {
  const container = document.createElement('section');
  const title = document.createElement('h3');
  title.className = 'section-title';
  title.textContent = label;
  const box = document.createElement('div');
  box.className = 'detail-box';
  const text = document.createElement('p');
  text.textContent = content || 'Sem informações.';
  box.append(text);
  container.append(title, box);
  return container;
}

function createPepSection(peps) {
  const container = document.createElement('section');
  const title = document.createElement('h3');
  title.className = 'section-title';
  title.textContent = 'PEPs do Projeto';
  const list = document.createElement('div');
  list.className = 'inline-list';
  peps.forEach((pep) => {
    const article = document.createElement('article');
    const heading = document.createElement('h4');
    heading.textContent = pep.Title || 'Elemento PEP';
    const text = document.createElement('p');
    const amount = pep.amountBrl ? BRL.format(pep.amountBrl) : '—';
    text.textContent = `Valor: ${amount} • Ano: ${pep.year || '—'}`;
    article.append(heading, text);
    list.append(article);
  });
  container.append(title, list);
  return container;
}

function createKeyProjectsSection(milestones, activities, activityPeps) {
  const container = document.createElement('section');
  const title = document.createElement('h3');
  title.className = 'section-title';
  title.textContent = 'Key Projects';
  container.append(title);

  const list = document.createElement('div');
  list.className = 'inline-list';

  milestones.forEach((milestone) => {
    const article = document.createElement('article');
    const heading = document.createElement('h4');
    heading.textContent = milestone.Title || 'Marco';
    article.append(heading);

    const milestoneActivities = activities.filter((act) => act.milestonesId === milestone.Id);
    milestoneActivities.forEach((activity) => {
      const activityBox = document.createElement('div');
      activityBox.className = 'detail-box';
      const activityTitle = document.createElement('strong');
      activityTitle.textContent = activity.Title || 'Atividade';
      const description = document.createElement('div');
      description.innerHTML = [
        formatPeriod(activity.startDate, activity.endDate),
        activity.supplier ? `Fornecedor: ${activity.supplier}` : '',
        activity.activityDescription || ''
      ].filter(Boolean).join('<br>');

      const pepsForActivity = activityPeps.filter((pep) => pep.activitiesId === activity.Id);
      if (pepsForActivity.length) {
        const pepList = document.createElement('ul');
        pepsForActivity.forEach((pep) => {
          const li = document.createElement('li');
          const amount = pep.amountBrl ? BRL.format(pep.amountBrl) : '—';
          li.textContent = `${pep.Title || 'PEP'} • ${amount} • Ano ${pep.year || '—'}`;
          pepList.append(li);
        });
        description.appendChild(pepList);
      }

      activityBox.append(activityTitle, description);
      article.append(activityBox);
    });

    list.append(article);
  });

  container.append(list);
  return container;
}

// ============================================================================
// Formulário: abertura, preenchimento e coleta dos dados
// ============================================================================
function openProjectForm(mode, detail = null) {
  projectForm.reset();
  formStatus.classList.remove('show');
  formErrors.classList.remove('show');
  formErrors.textContent = '';
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
    updateBudgetSections({ clear: true });
  } else if (detail) {
    fillFormWithProject(detail);
  }

  updateSimplePepYears();
  overlay.classList.remove('hidden');
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
      const relatedActivities = activities.filter((act) => act.milestonesId === milestone.Id);
      relatedActivities.forEach((activity) => {
        const activityBlock = addActivityBlock(block, {
          id: activity.Id,
          title: activity.Title,
          start: activity.startDate,
          end: activity.endDate,
          supplier: activity.supplier,
          description: activity.activityDescription
        }, false);
        const relatedPeps = activityPeps.filter((pep) => pep.activitiesId === activity.Id);
        relatedPeps.forEach((pep) => {
          const pepRow = createActivityPepRow({
            id: pep.Id,
            title: pep.Title,
            amount: pep.amountBrl,
            year: pep.year
          });
          activityBlock.querySelector('.activity-pep-list').append(pepRow);
          state.editingSnapshot.activityPeps.add(Number(pep.Id));
        });
        if (!relatedPeps.length) {
          addActivityPepRow(activityBlock);
        }
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
}

function closeForm() {
  overlay.classList.add('hidden');
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
}

function setSectionInteractive(section, enabled) {
  if (!section) return;
  section.querySelectorAll('input, textarea, button').forEach((element) => {
    if (element.type === 'hidden') return;
    element.disabled = !enabled;
  });
}

function updateSimplePepYears() {
  const year = parseInt(approvalYearInput.value, 10) || '';
  simplePepList.querySelectorAll('.pep-year').forEach((input) => {
    input.value = year;
  });
}

function ensureSimplePepRow() {
  const row = createSimplePepRow({ year: parseInt(approvalYearInput.value, 10) || '' });
  simplePepList.append(row);
}

function ensureMilestoneBlock() {
  const block = createMilestoneBlock();
  milestoneList.append(block);
  addActivityBlock(block, {}, true);
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

function addActivityBlock(milestoneElement, data = {}, addDefaultPep = true) {
  if (!milestoneElement) return null;
  const fragment = activityTemplate.content.cloneNode(true);
  const activity = fragment.querySelector('.activity');
  activity.dataset.activityId = data.id || '';
  activity.querySelector('.activity-title').value = data.title || '';
  activity.querySelector('.activity-start').value = data.start ? data.start.substring(0, 10) : '';
  activity.querySelector('.activity-end').value = data.end ? data.end.substring(0, 10) : '';
  activity.querySelector('.activity-supplier').value = data.supplier || '';
  activity.querySelector('.activity-description').value = data.description || '';
  milestoneElement.querySelector('.activity-list').append(activity);
  if (addDefaultPep) {
    addActivityPepRow(activity);
  }
  return activity;
}

function createActivityPepRow({ id = '', title = '', amount = '', year = '' } = {}) {
  const fragment = activityPepTemplate.content.cloneNode(true);
  const row = fragment.querySelector('.activity-pep');
  row.dataset.pepId = id;
  row.querySelector('.activity-pep-title').value = title || '';
  row.querySelector('.activity-pep-amount').value = amount ?? '';
  row.querySelector('.activity-pep-year').value = year ?? '';
  return row;
}

function addActivityPepRow(activityElement, data = {}) {
  if (!activityElement) return null;
  const list = activityElement.querySelector('.activity-pep-list');
  const row = createActivityPepRow(data);
  list.append(row);
  return row;
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
      projectsId: projectId,
      activitiesId: null
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
      projectsId: projectId
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
        projectsId: projectId,
        milestonesId: milestoneId
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

      for (const pepRow of activity.querySelectorAll('.activity-pep')) {
        const pepIdRaw = pepRow.dataset.pepId;
        const pepPayload = {
          Title: pepRow.querySelector('.activity-pep-title').value.trim(),
          amountBrl: parseFloat(pepRow.querySelector('.activity-pep-amount').value) || 0,
          year: parseNumber(pepRow.querySelector('.activity-pep-year').value),
          projectsId: projectId,
          activitiesId: activityId
        };
        let pepId = Number(pepIdRaw);
        if (pepIdRaw) {
          await sp.updateItem('Peps', pepId, pepPayload);
        } else {
          const createdPep = await sp.createItem('Peps', pepPayload);
          pepId = createdPep.Id;
          pepRow.dataset.pepId = pepId;
        }
        activityPepIds.add(Number(pepId));
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
}

function formatPeriod(start, end) {
  if (!start && !end) return 'Sem datas definidas';
  const startLabel = start ? DATE_FMT.format(new Date(start)) : '—';
  const endLabel = end ? DATE_FMT.format(new Date(end)) : '—';
  return `${startLabel} até ${endLabel}`;
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
