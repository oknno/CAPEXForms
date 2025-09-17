//
//  Script principal do protótipo CAPEX Forms.
//  Aqui concentro tanto a camada de integração com SharePoint quanto os
//  comportamentos da interface montada em HTML estático.
//

// ============================================================================
// Classe de integração com SharePoint (SPRestApi)
// ============================================================================
class SPRestApi {
    /**
     * Cria uma instância da API REST do SharePoint.
     * @param {string} site - URL do site SharePoint. https://<seu contoso>.sharepoint.com/sites/<seu site>
     * @param {string|null} lista - Nome da lista padrão (opcional).
     */
    constructor(site = 'https://<seu contoso>.sharepoint.com/sites/<seu site>', lista = null) {
        this.site = site;
        this.listaAtual = lista;
    }

    /**
     * Define a lista atual para operações subsequentes.
     * @param {string} listaName - Nome da lista.
     */
    setLista(listaName) {
        this.listaAtual = listaName;
    }

    /**
     * Cria uma nova instância da API com uma lista específica.
     * @param {string} listaName - Nome da lista.
     * @returns {SPRestApi} Nova instância com lista definida.
     */
    getLista(listaName) {
        return new SPRestApi(this.site, listaName);
    }

    /**
     * Codifica o nome da lista para o formato esperado pelo SharePoint.
     * @param {string} lista - Nome da lista.
     * @returns {string} Tipo de entidade codificado.
     */
    encodeEntityType(lista) {
        return "SP.Data." + lista.replace(/ /g, '_x0020_').replace(/_/g, '_x005f_') + "ListItem";
    }

    /**
     * Constrói a URL da API para uma lista.
     * @param {string} lista - Nome da lista.
     * @param {string} endpoint - Caminho adicional da API.
     * @returns {string} URL completa da API.
     */
    buildListUrl(lista, endpoint = '') {
        if (!lista) throw new Error("Lista não definida.");
        return `${this.site}${ this.site.charAt(this.site.length-1) === '/'? '': '' }_api/web/lists/getbytitle('${lista}')${endpoint}`;
    }

    /**
     * Executa uma requisição HTTP genérica.
     * @param {string} url - URL da requisição.
     * @param {string} method - Método HTTP.
     * @param {Object} headers - Cabeçalhos da requisição.
     * @param {any} body - Corpo da requisição.
     * @returns {Promise<Object>} Resposta da API.
     */
    async request(url, method = 'GET', headers = {}, body = null) {
        const response = await fetch(url, { method, headers, body });
        const json = await response.json();
        return json;
    }

    /**
     * Obtém o valor do Form Digest necessário para requisições POST.
     * @returns {Promise<string>} Valor do Form Digest.
     */
    async getFormDigestValue() {
        try {
            const url = `${this.site}/_api/contextinfo`;
            const headers = {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            };
            const data = await this.request(url, "POST", headers);
            return data.d.GetContextWebInformation.FormDigestValue;
        } catch (error) {
            console.error("Erro ao obter o Form Digest:", error);
            return _spPageContextInfo.formDigestValue;
        }
    }

    /**
     * Adiciona um item à lista.
     * @param {Object} data - Dados do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do item criado ou false.
     */
    async addItem(data = {}, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, "/items");
        const headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigest
        };
        const payload = JSON.stringify({
            "__metadata": { "type": this.encodeEntityType(lista) },
            ...data
        });
        const response = await this.request(url, "POST", headers, payload);
        return response.error ? false : response;
    }

    /**
     * Adiciona um anexo a um item da lista.
     * @param {number} itemId - ID do item.
     * @param {string} fileName - Nome do arquivo.
     * @param {Blob|ArrayBuffer} fileContent - Conteúdo do arquivo.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do anexo ou false.
     */
    async addAttachment(itemId, fileName, fileContent, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "Content-Type": "application/octet-stream"
        };
        const response = await fetch(url, { method: "POST", headers, body: fileContent });
        if (!response.ok) {
            const error = await response.json();
            console.error("Erro ao adicionar anexo:", error);
            return false;
        }
        const result = await response.json();
        return result.d;
    }

    /**
     * Atualiza um item existente na lista.
     * @param {number} id - ID do item.
     * @param {Object} data - Dados atualizados.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Confirmação ou false.
     */
    async updateItem(id, data = {}, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        };
        const payload = JSON.stringify({
            "__metadata": { "type": this.encodeEntityType(lista) },
            ...data
        });
        const response = await fetch(url, { method: "POST", headers, body: payload });
        const text = await response.text();
        if (!response.ok) {
            try {
                const errorData = text ? JSON.parse(text) : {};
                console.error('Erro detalhado do SharePoint:', errorData.error?.message?.value || errorData);
            } catch (parseErr) {
                console.error('Erro ao atualizar item no SharePoint:', text || parseErr);
            }
            return false;
        }
        if (!text) {
            return {};
        }
        try {
            return JSON.parse(text);
        } catch (parseErr) {
            console.warn('Resposta vazia ou inválida ao atualizar item:', parseErr);
            return {};
        }
    }

    /**
     * Exclui um item da lista.
     * @param {number} id - ID do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<boolean>} True se sucesso.
     */
    async deleteItem(id, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        };
        const response = await fetch(url, { method: "POST", headers });
        if (!response.ok) throw new Error(await response.text());
        return true;
    }

    /**
     * Executa qualquer requisição arbitrária à API do SharePoint.
     * @param {string} api - Caminho da API.
     * @param {string} [method="GET"] - Método HTTP.
     * @param {any} [body=null] - Corpo da requisição.
     * @param {Object} [headers={}] - Cabeçalhos adicionais.
     * @returns {Promise<Object>} Resposta da API.
     */
    async anyRequest(api, method = "GET", body = null, headers = {}) {
        const url = `${this.site}/_api/${api}`;
        const defaultHeaders = {
            "accept": "application/json;odata=verbose",
            "accept-language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "charset": "UTF-8"
        };
        return await this.request(url, method, { ...defaultHeaders, ...headers }, body);
    }

    /**
     * Obtém itens da lista com parâmetros opcionais de consulta.
     * @param {Object} [params={}] - Parâmetros de consulta.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object>} Lista de itens.
     */
    async getItems(params = {}, lista = this.listaAtual) {
        const url = new URL(this.buildListUrl(lista, "/items"));
        for (const parameter in params) {
            url.searchParams.append(`$${parameter}`, params[parameter]);
        }
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url.toString(), { method: 'GET', headers });
        return await response.json();
    }

    /**
     * Recupera um item específico da lista pelo ID.
     * @param {number} id - ID do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object>} Item recuperado.
     */
    async getItemById(id, lista = this.listaAtual) {
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém metadados da lista, como campos e tipos.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Metadados da lista.
     */
    async getListMetadata(lista = this.listaAtual) {
        const url = this.buildListUrl(lista, "/fields");
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d.results;
    }

    /**
     * Obtém informações do usuário atual logado.
     * @returns {Promise<Object>} Dados do usuário.
     */
    async getUserInfo() {
        const url = `${this.site}/_api/web/currentuser`;
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém informações gerais do site atual.
     * @returns {Promise<Object>} Dados do site.
     */
    async getSiteInfo() {
        const url = `${this.site}/_api/web`;
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Pesquisa itens na lista com base em um filtro OData.
     * @param {string} filtro - Filtro OData.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Itens filtrados.
     */
    async searchItems(filtro, lista = this.listaAtual) {
        const url = this.buildListUrl(lista, `/items?$filter=${encodeURIComponent(filtro)}`);
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d.results;
    }
}

// ============================================================================
// Escopo principal: lógica da interface e orquestração dos fluxos
// ============================================================================
(function () {
  // ========================================================================
  // Utilitários (formatadores, parsers, helpers)
  // ========================================================================
  const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });

  function parseNumberBRL(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    const normalized = String(val).replace(/\./g, '').replace(',', '.').replace(/[^\d.]/g, '');
    return Number(normalized || 0);
  }

  function formatDate(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      return isNaN(d) ? '' : d.toLocaleDateString('pt-BR');
    } catch (e) {
      return '';
    }
  }

  function getStatusColor(status) {
    switch (status) {
      case 'Rascunho': return '#414141';
      case 'Em Aprovação': return '#970886';
      case 'Reprovado': return '#f83241';
      case 'Reprovado para Revisão': return '#fe8f46';
      case 'Aprovado': return '#3d9308';
      default: return '#414141';
    }
  }

  function getValueFromSelector(queryOrId, defaultValue = "", parent = document) {
    let e = null;
    try {
      e = typeof parent.getElementById === "function" ? parent.getElementById(queryOrId) : parent.querySelector('#' + queryOrId);
    } catch (error) {
    }
    if (e === null) e = parent.querySelector(queryOrId);
    if (e === null) e = { value: defaultValue };
    return e.value;
  }

  // ========================================================================
  // Controle de estado global (referências e variáveis compartilhadas)
  // ========================================================================
  const REQ_THRESHOLD = 1000000; // 1 milhão
  const SharePoint = new SPRestApi('https://arcelormittal.sharepoint.com/sites/controladorialongos/capex/');

  const form = document.getElementById('capexForm');
  const statusBox = document.getElementById('status');
  const submitBtn = form.querySelector('button[type="submit"]');
  const errorsBox = document.getElementById('errors');
  const milestonesWrap = document.getElementById('milestones');
  const addMilestoneBtn = document.getElementById('addMilestoneBtn');
  const projectBudgetInput = document.getElementById('projectBudget');
  const approvalYearInput = document.getElementById('approvalYear');
  const capexFlag = document.getElementById('capexFlag');
  const milestoneSection = document.getElementById('milestoneSection');
  const pepSection = document.getElementById('pepSection');
  const pepDropdown = document.getElementById('pepDropdown');
  const pepValueInput = document.getElementById('pepValue');
  const projectList = document.getElementById('projectList');
  const projectDetails = document.getElementById('projectDetails');
  const appContainer = document.getElementById('app');
  const newProjectBtn = document.getElementById('newProjectBtn');
  const saveDraftBtn = document.getElementById('saveDraftBtn');
  const backBtn = document.getElementById('backBtn');
  const milestoneTpl = document.getElementById('milestoneTemplate');
  const activityTpl = document.getElementById('activityTemplate');

  google.charts.load('current', { packages: ['gantt'] });

  const currentYear = new Date().getFullYear();
  approvalYearInput.max = currentYear;
  approvalYearInput.placeholder = currentYear;

  let milestoneCount = 0;
  let currentProjectsId = null;
  let resetFormWithoutAlert = true;
  let availablePeps = [];

  // ========================================================================
  // Manipulação de UI (mostrar/ocultar formulário, detalhes, feedback)
  // ========================================================================
  function populatePepSelect(selectEl, preferredValue = '') {
    if (!selectEl) return;
    const desiredValue = preferredValue || selectEl.dataset.selectedPep || selectEl.value || '';
    const currentScroll = selectEl.scrollTop;
    selectEl.innerHTML = '';

    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = 'Selecione…';
    selectEl.appendChild(placeholder);

    availablePeps.forEach((pep) => {
      const option = document.createElement('option');
      option.value = pep.code;
      option.textContent = pep.code;
      option.dataset.amount = pep.amount;
      selectEl.appendChild(option);
    });

    selectEl.value = desiredValue;
    if (selectEl.value !== desiredValue) {
      selectEl.value = '';
    }
    selectEl.dataset.selectedPep = desiredValue || selectEl.value || '';
    if (selectEl.value) {
      selectEl.classList.remove('is-invalid');
    }
    selectEl.scrollTop = currentScroll;
  }

  function updatePepValueDisplay() {
    if (!pepDropdown || !pepValueInput) return;
    const selectedCode = pepDropdown.value || '';
    pepDropdown.dataset.selectedPep = selectedCode;
    const info = availablePeps.find(pep => pep.code === selectedCode);
    if (info) {
      pepValueInput.value = info.amount ?? 0;
      if (selectedCode) {
        pepDropdown.classList.remove('is-invalid');
      }
    } else {
      pepValueInput.value = '';
    }
  }

  async function loadPepOptions() {
    if (pepDropdown && pepDropdown.options.length === 0) {
      populatePepSelect(pepDropdown, pepDropdown.dataset.selectedPep || pepDropdown.value);
    }
    try {
      const PepsList = SharePoint.getLista('Peps');
      const res = await PepsList.getItems({ select: 'Title,amountBrl' });
      const items = res.d?.results || [];
      const unique = new Map();
      items.forEach(item => {
        const code = (item.Title || '').trim();
        if (!code) return;
        if (!unique.has(code)) {
          const amount = Number(item.amountBrl ?? 0);
          unique.set(code, { code, amount: Number.isFinite(amount) ? amount : 0 });
        }
      });
      availablePeps = Array.from(unique.values()).sort((a, b) => a.code.localeCompare(b.code, 'pt-BR'));
    } catch (error) {
      availablePeps = [];
    }

    if (pepDropdown) {
      populatePepSelect(pepDropdown, pepDropdown.dataset.selectedPep || pepDropdown.value);
      updatePepValueDisplay();
    }
    [...document.querySelectorAll('.act-pep')].forEach(select => {
      populatePepSelect(select, select.dataset.selectedPep || select.value);
    });
  }

  function updateStatus(message = '', type = 'info') {
    if (!statusBox) return;
    statusBox.textContent = message;
    statusBox.className = `status ${type}`;
  }

  function overThreshold() {
    return parseNumberBRL(projectBudgetInput.value) >= REQ_THRESHOLD;
  }

  function updateCapexFlag() {
    const n = parseNumberBRL(projectBudgetInput.value);
    if (!n) {
      capexFlag.textContent = '';
      return;
    }
    capexFlag.innerHTML = n >= REQ_THRESHOLD
      ? `<span class="ok">Orçamento do Projeto ${BRL.format(n)} &ge; ${BRL.format(REQ_THRESHOLD)} — marcos obrigatórios.</span>`
      : `Orçamento do Projeto ${BRL.format(n)} &lt; ${BRL.format(REQ_THRESHOLD)} — marcos não necessários.`;
  }

  function updateMilestoneVisibility() {
    const budget = parseNumberBRL(projectBudgetInput.value);
    const showMilestones = budget >= REQ_THRESHOLD;
    const showPep = budget > 0 && budget < REQ_THRESHOLD;

    milestoneSection.style.display = showMilestones ? 'block' : 'none';
    if (!showMilestones) {
      milestonesWrap.innerHTML = '';
      milestoneCount = 0;
      refreshGantt();
    }

    if (pepSection) {
      pepSection.style.display = showPep ? 'block' : 'none';
      if (!showPep) {
        if (pepDropdown) {
          pepDropdown.value = '';
          pepDropdown.dataset.selectedPep = '';
          pepDropdown.classList.remove('is-invalid');
        }
        if (pepValueInput) {
          pepValueInput.value = '';
        }
      } else if (pepDropdown) {
        updatePepValueDisplay();
      }
    }
  }

  function resetForm() {
    form.reset();
    currentProjectsId = null;
    [...form.elements].forEach(el => el.disabled = false);
    if (saveDraftBtn) saveDraftBtn.style.display = 'inline-flex';
    submitBtn.style.display = 'inline-flex';
    if (pepDropdown) {
      pepDropdown.value = '';
      pepDropdown.dataset.selectedPep = '';
      pepDropdown.classList.remove('is-invalid');
    }
    if (pepValueInput) {
      pepValueInput.value = '';
    }
    if (pepSection) {
      pepSection.style.display = 'none';
    }
    milestonesWrap.innerHTML = '';
    milestoneCount = 0;
    refreshGantt();
  }

  function showForm() {
    if (appContainer) appContainer.style.display = 'none';
    form.style.display = 'block';
    if (backBtn) backBtn.style.display = 'inline-flex';
    if (newProjectBtn) newProjectBtn.style.display = 'none';
    document.body.style.overflow = 'auto';
  }

  function showProjectList() {
    form.style.display = 'none';
    if (appContainer) appContainer.style.display = 'flex';
    if (backBtn) backBtn.style.display = 'none';
    if (newProjectBtn) newProjectBtn.style.display = 'inline-block';
    resetForm();
    document.body.style.overflow = 'hidden';
  }

  function showProjectDetails(item) {
    if (!projectDetails) return;

    projectDetails.innerHTML = '';

    const wrapper = document.createElement('div');
    wrapper.className = 'project-details';

    if (!item) {
      const empty = document.createElement('div');
      empty.className = 'empty';
      const emptyTitle = document.createElement('p');
      emptyTitle.className = 'empty-title';
      emptyTitle.textContent = 'Selecione um projeto';
      const emptyCopy = document.createElement('p');
      emptyCopy.textContent = 'Clique em um projeto na lista ao lado para ver os detalhes';
      empty.append(emptyTitle, emptyCopy);
      wrapper.appendChild(empty);
      projectDetails.appendChild(wrapper);
      return;
    }

    const createDetailCard = (label, value, valueClass = '') => {
      const card = document.createElement('div');
      card.className = 'detail-card';
      const heading = document.createElement('h3');
      heading.textContent = label;
      const text = document.createElement('p');
      if (valueClass) text.className = valueClass;
      text.textContent = value;
      card.append(heading, text);
      return card;
    };

    const header = document.createElement('div');
    header.className = 'details-header';
    const titleEl = document.createElement('h1');
    titleEl.textContent = item.Title || '';
    const statusBadge = document.createElement('span');
    statusBadge.className = 'status-badge';
    const statusValue = item.status || '';
    statusBadge.textContent = statusValue;
    statusBadge.style.background = getStatusColor(statusValue);
    header.append(titleEl, statusBadge);

    const grid = document.createElement('div');
    grid.className = 'details-grid';

    const budgetCard = createDetailCard('Orçamento', BRL.format(item.budgetBrl ?? 0), 'budget-value');
    const responsible = createDetailCard('Responsável', item.projectLeader || item.projectUser || '');
    const startDate = createDetailCard('Data de Início', formatDate(item.startDate));
    const endDate = createDetailCard('Data de Conclusão', formatDate(item.endDate));

    const descriptionCard = createDetailCard('Descrição do Projeto', item.businessNeed || item.proposedSolution || '');
    descriptionCard.classList.add('detail-desc');

    grid.append(budgetCard, responsible, startDate, endDate, descriptionCard);

    const actions = document.createElement('div');
    actions.className = 'detail-actions';

    const addActionButton = (label, className, handler) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = className;
      button.textContent = label;
      if (typeof handler === 'function') {
        button.addEventListener('click', handler);
      }
      actions.appendChild(button);
      return button;
    };

    const status = item.status || '';

    switch (status) {
      case 'Rascunho':
        addActionButton('Editar Projeto', 'btn secondary action-btn', () => editProject(item.Id));
        addActionButton('Enviar para Aprovação', 'btn primary action-btn approve');
        break;
      case 'Reprovado para Revisão':
        addActionButton('Editar Projeto', 'btn secondary action-btn', () => editProject(item.Id));
        break;
      case 'Aprovado':
      case 'Em Aprovação':
        addActionButton('Visualizar Projeto', 'btn secondary action-btn', () => editProject(item.Id));
        break;
      case 'Reprovado':
        break;
      default:
        addActionButton('Editar Projeto', 'btn secondary action-btn', () => editProject(item.Id));
        break;
    }

    wrapper.append(header, grid);
    if (actions.childElementCount > 0) {
      wrapper.appendChild(actions);
    }

    projectDetails.appendChild(wrapper);
  }

  // ========================================================================
  // CRUD de Projetos (create, update, delete, list)
  // ========================================================================
  function getProjectData() {
    return {
      nome: getValueFromSelector('projectName').trim(),
      ano_aprovacao: parseInt(getValueFromSelector('approvalYear', 0), 10),
      capex_budget_brl: parseNumberBRL(getValueFromSelector('projectBudget')),
      nivel_investimento: getValueFromSelector('investmentLevel').trim(),
      origem_verba: getValueFromSelector('fundingSource').trim(),
      project_user: getValueFromSelector('projectUser').trim(),
      project_leader: getValueFromSelector('projectLeader').trim(),
      empresa: getValueFromSelector('company').trim(),
      centro: getValueFromSelector('center').trim(),
      unidade: getValueFromSelector('unit').trim(),
      local_implantacao: getValueFromSelector('projectLocation').trim(),
      ccusto_depreciacao: getValueFromSelector('depreciationCostCenter').trim(),
      categoria: getValueFromSelector('category').trim(),
      tipo_investimento: getValueFromSelector('investmentType').trim(),
      tipo_ativo: getValueFromSelector('assetType').trim(),
      funcao_projeto: getValueFromSelector('projectFunction').trim(),
      data_inicio: getValueFromSelector('startDate', '').trim(),
      data_fim: getValueFromSelector('endDate', '').trim(),
      sumario: getValueFromSelector('projectSummary', '').trim(),
      comentario: getValueFromSelector('projectComment', '').trim(),
      kpi_tipo: getValueFromSelector('kpiType', '').trim(),
      kpi_nome: getValueFromSelector('kpiName', '').trim(),
      kpi_descricao: getValueFromSelector('kpiDescription', '').trim(),
      kpi_atual: parseNumberBRL(getValueFromSelector('kpiCurrent', 0).trim()),
      kpi_esperado: parseNumberBRL(getValueFromSelector('kpiExpected', 0).trim())
    };
  }

  async function loadUserProjects() {
    if (!projectList) return;
    projectList.innerHTML = '';
    const projetos = SharePoint.getLista('Projects');
    const res = await projetos.getItems({
      select: 'Id,Title,budgetBrl,status',
      filter: `AuthorId eq '${_spPageContextInfo.userId}'`
    });
    const items = res.d?.results || [];

    items.forEach(item => {
      const card = document.createElement('div');
      card.className = 'project-card';

      const statusBadge = document.createElement('span');
      statusBadge.className = 'status-badge';
      const statusValue = item.status || '';
      statusBadge.style.background = getStatusColor(statusValue);
      statusBadge.textContent = statusValue || '';

      const title = document.createElement('h3');
      title.textContent = item.Title || '';

      const budget = document.createElement('p');
      const budgetValue = item.budgetBrl;
      budget.textContent = BRL.format(budgetValue ?? 0);

      card.append(statusBadge, title, budget);
      card.addEventListener('click', async () => {
        const fullItem = await projetos.getItemById(item.Id);
        showProjectDetails(fullItem);
        [...projectList.children].forEach(c => c.classList.remove('selected'));
        card.classList.add('selected');
      });
      projectList.appendChild(card);
    });
    showProjectList();
    showProjectDetails(null);
  }

  function fillForm(item) {
    document.getElementById('projectName').value = item.Title || '';
    document.getElementById('approvalYear').value = item.approvalYear ?? '';
    document.getElementById('projectBudget').value = item.budgetBrl ?? '';
    document.getElementById('investmentLevel').value = item.investmentLevel ?? '';
    document.getElementById('fundingSource').value = item.fundingSource ?? '';
    document.getElementById('unit').value = item.unit ?? '';
    document.getElementById('center').value = item.center ?? '';
    document.getElementById('projectLocation').value = item.location ?? '';
    document.getElementById('projectUser').value = item.projectUser ?? '';
    document.getElementById('projectLeader').value = item.projectLeader ?? '';
    document.getElementById('company').value = item.company ?? '';
    document.getElementById('depreciationCostCenter').value = item.depreciationCostCenter ?? '';
    document.getElementById('category').value = item.category ?? '';
    document.getElementById('investmentType').value = item.investmentType ?? '';
    document.getElementById('assetType').value = item.assetType ?? '';
    document.getElementById('projectFunction').value = item.projectFunction ?? '';
    const startVal = item.startDate ?? '';
    const endVal = item.endDate ?? '';
    document.getElementById('startDate').value = startVal ? startVal.substring(0, 10) : '';
    document.getElementById('endDate').value = endVal ? endVal.substring(0, 10) : '';
    document.getElementById('projectSummary').value = item.businessNeed ?? '';
    document.getElementById('projectComment').value = item.proposedSolution ?? '';
    document.getElementById('kpiType').value = item.kpiType ?? '';
    document.getElementById('kpiName').value = item.kpiName ?? '';
    document.getElementById('kpiDescription').value = item.kpiDescription ?? '';
    document.getElementById('kpiCurrent').value = item.kpiCurrent ?? '';
    document.getElementById('kpiExpected').value = item.kpiExpected ?? '';
    updateCapexFlag();
    updateMilestoneVisibility();
  }

  async function editProject(id) {
    const item = await SharePoint.getLista('Projects').getItemById(id);
    currentProjectsId = id;
    fillForm(item);
    const msData = await fetchProjectStructure(id);
    setMilestonesData(msData);
    const statusValue = item.status || '';
    const editable = ['Rascunho', 'Reprovado para Revisão'].includes(statusValue);
    [...form.elements].forEach(el => el.disabled = !editable);
    if (saveDraftBtn) saveDraftBtn.style.display = editable ? 'inline-flex' : 'none';
    submitBtn.style.display = editable ? 'inline-flex' : 'none';
    showForm();
    updateStatus(`Status atual: ${statusValue}`, 'info');
  }

  async function saveDraft() {
    const data = getProjectData();
    const milestones = getMilestonesData();
    const payload = {
      Title: data.nome,
      approvalYear: data.ano_aprovacao,
      budgetBrl: data.capex_budget_brl,
      investmentLevel: data.nivel_investimento,
      fundingSource: data.origem_verba,
      projectUser: data.project_user,
      projectLeader: data.project_leader,
      company: data.empresa,
      center: data.centro,
      unit: data.unidade,
      location: data.local_implantacao,
      depreciationCostCenter: data.ccusto_depreciacao,
      category: data.categoria,
      investmentType: data.tipo_investimento,
      assetType: data.tipo_ativo,
      projectFunction: data.funcao_projeto,
      startDate: data.data_inicio || null,
      endDate: data.data_fim || null,
      businessNeed: data.sumario,
      proposedSolution: data.comentario,
      kpiType: data.kpi_tipo,
      kpiName: data.kpi_nome,
      kpiDescription: data.kpi_descricao,
      kpiCurrent: data.kpi_atual,
      kpiExpected: data.kpi_esperado,
      status: 'Rascunho'
    };
    updateStatus('Salvando rascunho...', 'info');
    try {
      let info;
      const projetos = SharePoint.getLista('Projects');
      if (currentProjectsId) {
        info = await projetos.updateItem(currentProjectsId, payload);
      } else {
        info = await projetos.addItem(payload);
        currentProjectsId = info.d?.Id || info.d?.ID;
      }
      await clearProjectStructure(currentProjectsId);
      await saveProjectStructure(currentProjectsId, milestones, data.ano_aprovacao);
      updateStatus('Rascunho salvo.', 'success');
      await loadUserProjects();
    } catch (e) {
      updateStatus('Erro ao salvar rascunho.', 'error');
    }
  }

  async function updateProjectStatus(id, status) {
    try {
      await SharePoint.getLista('Projects').updateItem(id, { status });
      await loadUserProjects();
      if (currentProjectsId === id) {
        await editProject(id);
      }
    } catch (e) {
      console.error(e);
    }
  }

  // ========================================================================
  // CRUD de Milestones
  // ========================================================================
  function addMilestone(nameDefault) {
    milestoneCount++;
    const node = milestoneTpl.content.cloneNode(true);
    const ms = node.querySelector('[data-milestone]');
    const nameInput = node.querySelector('.milestone-name');
    const nameSummaryHeader = node.querySelector('summary');
    const actsWrap = node.querySelector('[data-activities]');
    const btnAddAct = node.querySelector('[data-add-activity]');
    const btnRemove = node.querySelector('[data-remove-milestone]');

    nameInput.value = nameDefault || `Milestone ${milestoneCount}`;
    if (nameSummaryHeader) {
      nameSummaryHeader.textContent = nameInput.value;
      nameInput.addEventListener('input', e => (nameSummaryHeader.textContent = e.target.value));
    }

    btnAddAct.addEventListener('click', () => {
      addActivity(actsWrap);
      refreshGantt();
    });
    btnRemove.addEventListener('click', () => {
      ms.remove();
      renumberMilestones();
      refreshGantt();
    });

    milestonesWrap.appendChild(node);
    const justAdded = milestonesWrap.lastElementChild.querySelector('[data-activities]');
    addActivity(justAdded);
    renumberMilestones();
    refreshGantt();
  }

  function renumberMilestones() {
    const all = [...milestonesWrap.querySelectorAll('.milestone-name')];
    let idx = 1;
    for (const el of all) {
      if (!el.value || /^Milestone\s+\d+$/i.test(el.value.trim())) {
        el.value = `Milestone ${idx}`;
      }
      idx++;
    }
  }

  function setMilestonesData(milestones) {
    milestonesWrap.innerHTML = '';
    milestoneCount = 0;
    milestones.forEach(ms => {
      addMilestone(ms.nome);
      const msNode = milestonesWrap.lastElementChild;
      const actsWrap = msNode.querySelector('[data-activities]');
      actsWrap.innerHTML = '';
      (ms.atividades || []).forEach(act => {
        addActivity(actsWrap);
        const actNode = actsWrap.lastElementChild;
        actNode.querySelector('.act-title').value = act.titulo || '';
        const start = actNode.querySelector('.act-start');
        const end = actNode.querySelector('.act-end');
        if (act.inicio) start.value = act.inicio.substring(0,10);
        if (act.fim) end.value = act.fim.substring(0,10);
        start.dispatchEvent(new Event('change'));
        end.dispatchEvent(new Event('change'));
        const pepSelect = actNode.querySelector('.act-pep');
        if (pepSelect) {
          pepSelect.dataset.selectedPep = act.pep || '';
          populatePepSelect(pepSelect, act.pep || '');
          if (act.pep) {
            pepSelect.classList.remove('is-invalid');
          }
        }
        const overview = actNode.querySelector('.act-overview');
        if (overview) overview.value = act.descricao || '';
        (act.anual || []).forEach(a => {
          const row = actNode.querySelector(`.row[data-year="${a.ano}"]`);
          if (row) {
            row.querySelector('.act-capex').value = a.capex_brl;
          }
        });
      });
    });
    refreshGantt();
  }

  // ========================================================================
  // CRUD de Activities
  // ========================================================================
  function addActivity(container) {
    const node = activityTpl.content.cloneNode(true);
    const act = node.querySelector('[data-activity]');
    const btnRemove = node.querySelector('[data-remove-activity]');
    const start = node.querySelector('.act-start');
    const end = node.querySelector('.act-end');
    const yearWrap = node.querySelector('[data-year-fields]');
    const pepSelect = node.querySelector('.act-pep');

    if (pepSelect) {
      populatePepSelect(pepSelect, pepSelect.dataset.selectedPep || '');
      pepSelect.addEventListener('change', () => {
        pepSelect.dataset.selectedPep = pepSelect.value || '';
        if (pepSelect.value) {
          pepSelect.classList.remove('is-invalid');
        }
      });
    }

    function validateDates() {
      if (start.value && end.value && start.value > end.value) {
        end.setCustomValidity('A data de fim não pode ser anterior à data de início.');
      } else {
        end.setCustomValidity('');
      }
    }
    function updateYearFields() {
      if (!start.value || !end.value) {
        refreshGantt();
        return;
      }
      const startYear = new Date(start.value).getFullYear();
      const endYear = new Date(end.value).getFullYear();
      const years = [];
      for (let y = startYear; y <= endYear; y++) {
        const previousRow = yearWrap.querySelector(`.row[data-year="${y}"]`);
        years.push(y+'');
        if(previousRow !== null) continue;
        const row = document.createElement('div');
        row.className = 'row act-year';
        row.dataset.year = y;
        row.innerHTML = `
          <div class="act-year-field act-year-value">
            <label>Valor CAPEX da atividade (BRL) - ${y}</label>
            <input type="number" class="act-capex" data-year="${y}" min="0" step="0.01" inputmode="decimal" required placeholder="Ex.: 250000,00" />
          </div>
        `;
        yearWrap.appendChild(row);
      }

      [...yearWrap.querySelectorAll('.row[data-year]')].forEach(ye=>{
        if(!years.includes(ye.dataset.year)) ye.remove();
      });

      refreshGantt();
    }
    start.addEventListener('change', () => { validateDates(); updateYearFields(); });
    end.addEventListener('change', () => { validateDates(); updateYearFields(); });

    btnRemove.addEventListener('click', () => {
      act.remove();
      refreshGantt();
    });

    const today = new Date().toISOString().substring(0,10);
    const tomorrow = new Date( new Date().getTime() + 1000 * 60 * 60 * 24 ).toISOString().substring(0,10);
    start.value = today;
    end.value = tomorrow;

    container.appendChild(node);
    updateYearFields();
    refreshGantt();
  }

  function getMilestonesData() {
    const milestones = [];
    const msNodes = [...milestonesWrap.querySelectorAll('[data-milestone]')];
    const today = new Date().toISOString().substring(0,10);
    msNodes.forEach(ms => {
      const nome = getValueFromSelector('.milestone-name', "", ms).trim();
      const acts = [...ms.querySelectorAll('[data-activity]')].map(a => {
        const anual = [...a.querySelectorAll('.row[data-year]')].map(row => ({
          ano: parseInt(row.dataset.year, 10),
          capex_brl: parseNumberBRL(getValueFromSelector('.act-capex', 0, row)),
        }));
        const pepSelect = a.querySelector('.act-pep');
        const pepCode = pepSelect ? (pepSelect.value || '').trim() : '';
        return {
          titulo: getValueFromSelector('.act-title', "", a).trim(),
          inicio: getValueFromSelector('.act-start', today, a),
          fim: getValueFromSelector('.act-end', today, a),
          descricao: getValueFromSelector('.act-overview', "", a).trim(),
          pep: pepCode,
          anual,
        };
      });
      milestones.push({ nome, atividades: acts });
    });
    return milestones;
  }

  // ========================================================================
  // CRUD de Peps
  // ========================================================================
  async function clearProjectStructure(projectsId) {
  const Milestones = SharePoint.getLista('milestones');
  const Activities = SharePoint.getLista('activities');
  const Peps = SharePoint.getLista('peps');

  const msRes = await Milestones.getItems({ select: 'Id', filter: `projectsId eq ${projectsId}` });
  const marcos = msRes.d?.results || [];
  for (const ms of marcos) {
    const actRes = await Activities.getItems({ select: 'Id', filter: `milestonesId eq ${ms.Id}` });
    const acts = actRes.d?.results || [];
    for (const act of acts) {
      const alRes = await Peps.getItems({ select: 'Id', filter: `activitiesId eq ${act.Id}` });
      const als = alRes.d?.results || [];
      for (const al of als) {
        await Peps.deleteItem(al.Id);
      }
      await Activities.deleteItem(act.Id);
    }
    await Milestones.deleteItem(ms.Id);
  }
}

async function saveProjectStructure(projectsId, milestones, projectApprovalYear) {
  const Milestones = SharePoint.getLista('milestones');
  const Activities = SharePoint.getLista('activities');
  const Peps = SharePoint.getLista('peps');
  const projectLookupId = Number(projectsId);

  if (!Number.isFinite(projectLookupId)) {
    throw new Error('Project ID inválido para salvar a estrutura.');
  }

  const approvalYearNumber = Number(projectApprovalYear);
  const projectYear = Number.isFinite(approvalYearNumber) ? approvalYearNumber : null;

  const milestonesList = Array.isArray(milestones) ? milestones : [];
  for (const milestone of milestonesList) {
    const milestonePayload = {
      Title: (milestone?.nome || '').trim(),
      projectsId: projectLookupId
    };

    const infoMarco = await Milestones.addItem(milestonePayload);
    const marcoIdRaw = infoMarco?.d?.Id ?? infoMarco?.d?.ID;
    const marcoId = Number(marcoIdRaw);
    if (!Number.isFinite(marcoId)) continue;

    const atividades = Array.isArray(milestone?.atividades) ? milestone.atividades : [];
    for (const atividade of atividades) {
      const activityPayload = {
        Title: (atividade?.titulo || '').trim(),
        startDate: atividade?.inicio || null,
        endDate: atividade?.fim || null,
        activityDescription: atividade?.descricao || '',
        milestonesId: marcoId,
        projectsId: projectLookupId
      };

      const infoAtv = await Activities.addItem(activityPayload);
      const atvIdRaw = infoAtv?.d?.Id ?? infoAtv?.d?.ID;
      const atvId = Number(atvIdRaw);

      const anualEntries = Array.isArray(atividade?.anual) ? atividade.anual : [];
      for (const anual of anualEntries) {
        const amountNumber = Number(anual?.capex_brl ?? 0);
        const annualYearNumber = Number(anual?.ano);

        const pepPayload = {
          Title: String((atividade?.pep || '') || atividade?.titulo || '').trim(),
          amountBrl: Number.isFinite(amountNumber) ? amountNumber : 0,
          year: Number.isFinite(projectYear) ? projectYear : (Number.isFinite(annualYearNumber) ? annualYearNumber : null),
          projectsId: projectLookupId,
          activitiesId: atvId
        };

        await Peps.addItem(pepPayload);
      }
    }
  }
}

async function fetchProjectStructure(projectsId) {
  const Milestones = SharePoint.getLista('milestones');
  const Activities = SharePoint.getLista('activities');
  const Peps = SharePoint.getLista('peps');

  const msRes = await Milestones.getItems({ select: 'Id,Title', filter: `projectsId eq ${projectsId}` });
  const result = [];

  for (const ms of msRes.d?.results || []) {
    const actRes = await Activities.getItems({ select: 'Id,Title,startDate,endDate,activityDescription', filter: `milestonesId eq ${ms.Id}` });
    const acts = [];

    for (const act of actRes.d?.results || []) {
      const alRes = await Peps.getItems({ select: 'Id,Title,year,amountBrl', filter: `activitiesId eq ${act.Id}` });
      const anual = (alRes.d?.results || []).map(a => ({
        ano: a.year,
        capex_brl: a.amountBrl,
        pepTitle: a.Title
      }));

      acts.push({
        titulo: act.Title,
        inicio: act.startDate,
        fim: act.endDate,
        descricao: act.activityDescription || '',
        anual
      });
    }

    result.push({ nome: ms.Title, atividades: acts });
  }

  return result;
}

  // ========================================================================
  // Validações do formulário
  // ========================================================================
  function validateForm() {
    const errs = [];
    const errsEl = [];
    errorsBox.style.display = 'none';
    errorsBox.innerHTML = '';

    const reqFields = [
      { id: 'projectName', label: 'Nome do Projeto' },
      { id: 'approvalYear', label: 'Ano de Aprovação' },
      { id: 'projectBudget', label: 'Orçamento do Projeto em R$' },
      { id: 'investmentLevel', label: 'Nível de Investimento' },
      { id: 'fundingSource', label: 'Origem da Verba' },
      { id: 'projectUser', label: 'Project User' },
      { id: 'projectLeader', label: 'Coordenador do Projeto' },
      { id: 'company', label: 'Empresa' },
      { id: 'center', label: 'Centro' },
      { id: 'unit', label: 'Unidade' },
      { id: 'projectLocation', label: 'Local de Implantação' },
      { id: 'depreciationCostCenter', label: 'C Custo Depreciação' },
      { id: 'category', label: 'Categoria' },
      { id: 'investmentType', label: 'Tipo de Investimento' },
      { id: 'assetType', label: 'Tipo de Ativo' },
      { id: 'projectFunction', label: 'Função do Projeto' },
      { id: 'startDate', label: 'Data de Início' },
      { id: 'endDate', label: 'Data de Término' },
      { id: 'projectSummary', label: 'Sumário do Projeto' },
      { id: 'projectComment', label: 'Comentário' },
      { id: 'kpiType', label: 'Tipo de KPI' },
      { id: 'kpiName', label: 'Nome do KPI' },
      { id: 'kpiDescription', label: 'Descrição do KPI' },
      { id: 'kpiCurrent', label: 'KPI Atual' },
      { id: 'kpiExpected', label: 'KPI Esperado' },
    ];
    for (const f of reqFields) {
      const el = document.getElementById(f.id);
      if (!el.value || (el.type === 'number' && parseNumberBRL(el.value) < 0)) {
        errs.push(`Preencha o campo: <strong>${f.label}</strong>.`);
        errsEl.push(el);
      } else {
        el.classList.remove('is-invalid');
      }
    }

    const yearVal = parseInt(approvalYearInput.value, 10);
    if (isNaN(yearVal) || yearVal > currentYear) {
      errsEl.push(approvalYearInput);
      errs.push(`O <strong>ano de aprovação</strong> deve ser menor ou igual a ${currentYear}.`);
    } else {
      approvalYearInput.classList.remove('is-invalid');
    }

    const mustHaveMilestone = overThreshold();
    const milestones = [...milestonesWrap.querySelectorAll('[data-milestone]')];

    if (mustHaveMilestone && milestones.length === 0) {
      errs.push('O valor CAPEX é superior ou igual a R$ 1.000.000,00. Adicione pelo menos <strong>1 marco</strong>.');
    }

    if (!mustHaveMilestone && pepSection && pepSection.style.display !== 'none') {
      if (pepDropdown) {
        if (!pepDropdown.value) {
          errs.push('Selecione um <strong>elemento PEP</strong> na seção Elementos PEP.');
          errsEl.push(pepDropdown);
        } else {
          pepDropdown.classList.remove('is-invalid');
        }
      }
    }

    milestones.forEach((ms, i) => {
      const idx = i + 1;
      const name = ms.querySelector('.milestone-name');
      if (!name.value.trim()) {
        errs.push(`Informe o <strong>nome do marco ${idx}</strong>.`);
      }
      const acts = [...ms.querySelectorAll('[data-activity]')];
      if (acts.length === 0) {
        errs.push(`O <strong>marco ${idx}</strong> deve possuir pelo menos 1 atividade.`);
      }
      acts.forEach((a, j) => {
        const jdx = j + 1;
        const title = a.querySelector('.act-title');
        const start = a.querySelector('.act-start');
        const end = a.querySelector('.act-end');
        const overviewEl = a.querySelector('.act-overview');
        const pepSelect = a.querySelector('.act-pep');
        const yearRows = [...a.querySelectorAll('.row[data-year]')];

        if (!title.value.trim()) errs.push(`Atividade ${jdx} do marco ${idx}: informe o <strong>título</strong>.`);
        if (!start.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de início</strong>.`);
        if (!end.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de fim</strong>.`);
        if (start.value && end.value && start.value > end.value) {
          errs.push(`Atividade ${jdx} do marco ${idx}: a <strong>data de início</strong> não pode ser posterior à <strong>data de fim</strong>.`);
        }
        if (overviewEl) {
          if (!overviewEl.value.trim()) {
            errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>descrição da atividade</strong>.`);
            errsEl.push(overviewEl);
          }
        }
        if (pepSelect) {
          if (!pepSelect.value) {
            errs.push(`Atividade ${jdx} do marco ${idx}: selecione o <strong>elemento PEP</strong>.`);
            errsEl.push(pepSelect);
          } else {
            pepSelect.classList.remove('is-invalid');
          }
        }
        if (yearRows.length === 0) {
          errs.push(`Atividade ${jdx} do marco ${idx}: defina <strong>datas de início e fim</strong> válidas para gerar campos anuais.`);
        }
        yearRows.forEach(row => {
          const ano = row.dataset.year;
          const cap = parseNumberBRL(getValueFromSelector('.act-capex', 0, row));
          if (isNaN(cap) || cap < 0) {
            errs.push(`Atividade ${jdx} do marco ${idx}, ano ${ano}: informe o <strong>valor CAPEX</strong> (BRL) válido (≥ 0).`);
          }
        });
      });
    });

    if (errs.length) {
      const ul = document.createElement('ul');
      errs.forEach(e => {
        const li = document.createElement('li');
        li.innerHTML = e;
        ul.appendChild(li);
      });
      errorsBox.appendChild(document.createTextNode('Por favor, corrija os itens abaixo:'));
      errorsBox.appendChild(ul);
      errorsBox.style.display = 'block';
      errsEl.forEach(inel=>inel.classList.add('is-invalid'));
      try {
        errsEl[0].scrollIntoView({ behavior: "smooth" });
        errsEl[0].focus();
      } catch (error) {

      }
      return false;
    }
    return true;
  }

  // ========================================================================
  // Renderização de Gantt
  // ========================================================================
  function drawGantt(milestones) {
    const chartEl = document.getElementById('ganttChart');
    const titleEl = document.getElementById('ganttCharTitle');

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
    milestones.forEach((ms) => {
      idCounter++;
      let msStart = null;
      let msEnd = null;
      const actRows = [];
      (ms.atividades || []).forEach((act, index) => {
        const startDate = act.inicio ? new Date(act.inicio) : new Date();
        const endDate = act.fim ? new Date(act.fim) : new Date(startDate.getTime() + 1000 * 60 * 60 * 24);
        if (startDate && (!msStart || startDate < msStart)) msStart = startDate;
        if (endDate && (!msEnd || endDate > msEnd)) msEnd = endDate;
        const totalCapex = (act.anual || []).reduce((sum, y) => sum + (y.capex_brl || 0), 0);
        const descTooltip = (act.anual || []).map((y) => `${y.ano}: ${BRL.format(y.capex_brl)}`).join('<br/>');
        const pepTooltip = act.pep ? `<br/>PEP: ${act.pep}` : '';
        actRows.push([
          `ms-${idCounter}-${index}`,
          act.titulo || `Atividade ${index + 1}`,
          'Atividade',
          startDate,
          endDate,
          null,
          0,
          `ms-${idCounter}`,
          `CAPEX total: ${BRL.format(totalCapex)}${descTooltip ? '<br/>' + descTooltip : ''}${pepTooltip}`
        ]);
      });

      if (msStart && msEnd) {
        rows.push([
          `ms-${idCounter}`,
          ms.nome,
          'milestone',
          msStart,
          msEnd,
          null,
          0,
          null,
          ms.nome
        ]);
      }
      rows.push(...actRows);
    });

    if (!rows.length) {
      chartEl.innerHTML = '';
      titleEl.style.display = 'none';
      return;
    }

    titleEl.style.display = 'block';

    data.addRows(rows);
    const chart = new google.visualization.Gantt(chartEl);
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

  function refreshGantt() {
    const milestones = getMilestonesData();
    const draw = () => drawGantt(milestones);
    if (google.visualization && google.visualization.Gantt) {
      draw();
    } else {
      google.charts.setOnLoadCallback(draw);
    }
  }

  // ========================================================================
  // Listeners de eventos (submit, reset, botões, inputs)
  // ========================================================================
  if (newProjectBtn) {
    newProjectBtn.addEventListener('click', () => {
      resetFormWithoutAlert = true;
      resetForm();
      showForm();
      updateStatus('', 'info');
      resetFormWithoutAlert = false;
    });
  }

  addMilestoneBtn.addEventListener('click', () => addMilestone());
  milestoneSection.addEventListener('input', refreshGantt);
  milestoneSection.addEventListener('change', refreshGantt);
  if (saveDraftBtn) saveDraftBtn.addEventListener('click', saveDraft);
  if (backBtn) backBtn.addEventListener('click', showProjectList);

  if (projectBudgetInput) {
    projectBudgetInput.addEventListener('input', () => {
      updateCapexFlag();
      updateMilestoneVisibility();
      refreshGantt();
    });
    projectBudgetInput.addEventListener('change', () => {
      updateCapexFlag();
      updateMilestoneVisibility();
      refreshGantt();
    });
  }

  if (pepDropdown) {
    pepDropdown.addEventListener('change', () => {
      updatePepValueDisplay();
    });
  }

  form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    updateCapexFlag();
    updateMilestoneVisibility();
    if (!validateForm()) return;

    updateStatus('Enviando dados...', 'info');
    submitBtn.disabled = true;

    const payload = {
      projeto: getProjectData(),
      milestones: getMilestonesData()
    };

    const Projetos = SharePoint.getLista('Projects');

    try {
      const infoProjeto = await Projetos.addItem({
        Title: payload.projeto.nome,
        approvalYear: payload.projeto.ano_aprovacao,
        budgetBrl: payload.projeto.capex_budget_brl,
        investmentLevel: payload.projeto.nivel_investimento,
        fundingSource: payload.projeto.origem_verba,
        projectUser: payload.projeto.project_user,
        projectLeader: payload.projeto.project_leader,
        company: payload.projeto.empresa,
        center: payload.projeto.centro,
        unit: payload.projeto.unidade,
        location: payload.projeto.local_implantacao,
        depreciationCostCenter: payload.projeto.ccusto_depreciacao,
        category: payload.projeto.categoria,
        investmentType: payload.projeto.tipo_investimento,
        assetType: payload.projeto.tipo_ativo,
        projectFunction: payload.projeto.funcao_projeto,
        startDate: payload.projeto.data_inicio || null,
        endDate: payload.projeto.data_fim || null,
        businessNeed: payload.projeto.sumario,
        proposedSolution: payload.projeto.comentario,
        kpiType: payload.projeto.kpi_tipo,
        kpiName: payload.projeto.kpi_nome,
        kpiDescription: payload.projeto.kpi_descricao,
        kpiCurrent: payload.projeto.kpi_atual,
        kpiExpected: payload.projeto.kpi_esperado
      });

      const projectsId = Number(infoProjeto?.d?.Id ?? infoProjeto?.d?.ID);
      if (!Number.isFinite(projectsId)) {
        throw new Error('Project ID inválido retornado pelo SharePoint.');
      }

      await saveProjectStructure(projectsId, payload.milestones, payload.projeto.ano_aprovacao);
      updateStatus('Formulário enviado com sucesso!', 'success');
      refreshGantt();
      await loadUserProjects();
    } catch (error) {
      updateStatus('Ops, algo deu errado.', 'error');
    } finally {
      submitBtn.disabled = false;
    }
  });

  form.addEventListener('reset', ev => {
    if (resetFormWithoutAlert === false && !confirm('Tem certeza que deseja limpar o formulário?')) {
      ev.preventDefault();
    } else {
      updateStatus('Formulário limpo.', 'info');
    }
  });

  updateCapexFlag();
  updateMilestoneVisibility();
  loadPepOptions();
  refreshGantt();
  loadUserProjects();
})();
