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
            "Title": "",
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
        if (response.ok) return { d: { Id: parseInt(id) } };
        const errorData = await response.json();
        console.error('Erro detalhado do SharePoint:', errorData.error.message.value);
        return false;
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

(function () {
  const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  const REQ_THRESHOLD = 500000; // 1.5 milhões

  const SharePoint = new SPRestApi('https://arcelormittal.sharepoint.com/sites/controladorialongos/capex/');

  const form = document.getElementById('capexForm');
  const statusBox = document.getElementById('status');
  const submitBtn = form.querySelector('button[type="submit"]');
  const errorsBox = document.getElementById('errors');
  const milestonesWrap = document.getElementById('milestones');
  const addMilestoneBtn = document.getElementById('addMilestoneBtn');
  const projectValueInput = document.getElementById('projectValue');
  const approvalYearInput = document.getElementById('approvalYear');
  const capexFlag = document.getElementById('capexFlag');
  const milestoneSection = document.getElementById('milestoneSection');
  const projectList = document.getElementById('projectList');
  const projectDetails = document.getElementById('projectDetails');
  const appContainer = document.getElementById('app');
  const newProjectBtn = document.getElementById('newProjectBtn');
  const saveDraftBtn = document.getElementById('saveDraftBtn');
  const backBtn = document.getElementById('backBtn');
  google.charts.load('current', { packages: ['gantt'] });

  const milestoneTpl = document.getElementById('milestoneTemplate');
  const activityTpl = document.getElementById('activityTemplate');

  const currentYear = new Date().getFullYear();
  approvalYearInput.max = currentYear;
  approvalYearInput.placeholder = currentYear;

  let milestoneCount = 0;
  let currentProjectId = null;
  let resetFormWithoutAlert = true;

  // Helpers
  function updateStatus(message = '', type = 'info') {
    if (!statusBox) return;
    statusBox.textContent = message;
    statusBox.className = `status ${type}`;
  }

  function parseNumberBRL(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    // aceita ponto ou vírgula como separador decimal
    const normalized = String(val).replace(/\./g, '').replace(',', '.').replace(/[^\d.]/g, '');
    return Number(normalized || 0);
  }

  function overThreshold() {
    return parseNumberBRL(projectValueInput.value) > REQ_THRESHOLD;
  }

  function updateCapexFlag() {
    const n = parseNumberBRL(projectValueInput.value);
    if (!n) { capexFlag.textContent = ''; return; }
    capexFlag.innerHTML = n > REQ_THRESHOLD
      ? `<span class="ok">CAPEX BUDGET ${BRL.format(n)} &gt; ${BRL.format(REQ_THRESHOLD)} — marcos obrigatórios.</span>`
      : `CAPEX BUDGET ${BRL.format(n)} ≤ ${BRL.format(REQ_THRESHOLD)} — marcos não necessários.`;
  }

  function updateMilestoneVisibility() {
    const show = overThreshold();
    milestoneSection.style.display = show ? 'block' : 'none';
    if (!show) {
      milestonesWrap.innerHTML = '';
      milestoneCount = 0;
      refreshGantt();
    }
  }

  function resetForm() {
    form.reset();
    currentProjectId = null;
    [...form.elements].forEach(el => el.disabled = false);
    if (saveDraftBtn) saveDraftBtn.style.display = 'inline-flex';
    submitBtn.style.display = 'inline-flex';
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

  function getStatusColor(status) {
    switch (status) {
      case 'Rascunho': return '#414141';
      case 'Em Aprovação': return '#970886';
      case 'Recusado': return '#f83241';
      case 'Aprovado': return '#fe8f46';
      default: return '#414141';
    }
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

  function showProjectDetails(item) {
    if (!projectDetails) return;
    if (!item) {
      projectDetails.innerHTML = '<div class="project-details"><div class="empty"><p class="empty-title">Selecione um projeto</p><p>Clique em um projeto na lista ao lado</p><p>para ver os detalhes</p></div></div>';
    }
    projectDetails.innerHTML = `
      <div class="project-details">
        <div class="details-header">
          <h1>${item.Title}</h1>
          <span class="status-badge" style="background:${getStatusColor(item.Status)}">${item.Status}</span>
        </div>
        <div class="details-grid">
          <div class="detail-card">
            <h3>Orçamento</h3>
            <p class="budget-value">${BRL.format(item.CapexBudgetBRL || 0)}</p>
          </div>
          <div class="detail-card">
            <h3>Responsável</h3>
            <p>${item.Responsavel || item.ProjectLeader || ''}</p>
          </div>
          <div class="detail-card">
            <h3>Data de Início</h3>
            <p>${formatDate(item.DataInicio)}</p>
          </div>
          <div class="detail-card">
            <h3>Data de Conclusão</h3>
            <p>${formatDate(item.DataFim || item.DataConclusao)}</p>
          </div>
        <div class="detail-desc">
          <h3>Descrição do Projeto</h3>
          <p>${item.Descricao || ''}</p>
        </div>
      </div>
        <div class="detail-actions">
          <button type="button" class="action-btn edit" id="editProjectDetails">Editar Projeto</button>
          ${item.Status === 'Rascunho' ? '<button type="button" class="action-btn approve">Enviar para Aprovação</button>' : ''}
        </div>
      </div>`;
    const editBtn = document.getElementById('editProjectDetails');
    if (editBtn) {
      const editable = item.Status === 'Rascunho' || item.Status === 'Reprovado para Revisão';
      if (editable) {
        editBtn.addEventListener('click', () => editProject(item.Id));
      } else {
        editBtn.disabled = true;
      }
    }
  }

  if (newProjectBtn) {
    newProjectBtn.addEventListener('click', () => {
      resetFormWithoutAlert = true;
      resetForm();
      showForm();
      updateStatus('', 'info');
      resetFormWithoutAlert = false;
    });
  }

  async function loadUserProjects() {
    if (!projectList) return;
    projectList.innerHTML = '';
    const projetos = SharePoint.getLista('Projetos');
    const res = await projetos.getItems({
      select: 'Id,Title,CapexBudgetBRL,Status',
      filter: `AuthorId eq '${_spPageContextInfo.userId}'`
    });
    const items = res.d?.results || [];

    items.forEach(item => {
      const card = document.createElement('div');
      card.className = 'project-card';
      card.innerHTML = `
        <span class="status-badge" style="background:${getStatusColor(item.Status)}">${item.Status}</span>
        <h3>${item.Title}</h3>
        <p>${BRL.format(item.CapexBudgetBRL || 0)}</p>`;
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
    document.getElementById('approvalYear').value = item.AnoAprovacao || '';
    document.getElementById('unit').value = item.Unidade || '';
    document.getElementById('center').value = item.Centro || '';
    document.getElementById('projectLocation').value = item.LocalImplantacao || '';
    document.getElementById('projectUser').value = item.ProjectUser || '';
    document.getElementById('projectLeader').value = item.ProjectLeader || '';
    document.getElementById('investmentType').value = item.TipoInvestimento || '';
    document.getElementById('assetType').value = item.TipoAtivo || '';
    document.getElementById('usefulLife').value = item.VidaUtilAnos || '';
    document.getElementById('projectValue').value = item.CapexBudgetBRL || '';
    document.getElementById('necessidade').value = item.NecessidadeNegocio || '';
    document.getElementById('solucao').value = item.SolucaoProposta || '';
    document.getElementById('kpi').value = item.KpiImpactado || '';
    document.getElementById('kpiDesc').value = item.KpiDescricao || '';
    document.getElementById('kpi_actual').value = item.KpiValorAtual || '';
    document.getElementById('kpi_expected').value = item.KpiValorEsperado || '';
    updateCapexFlag();
    updateMilestoneVisibility();
  }

  async function editProject(id) {
    const item = await SharePoint.getLista('Projetos').getItemById(id);
    currentProjectId = id;
    fillForm(item);
    const msData = await fetchProjectStructure(id);
    setMilestonesData(msData);
    const editable = ['Rascunho', 'Reprovado para Revisão'].includes(item.Status);
    [...form.elements].forEach(el => el.disabled = !editable);
    if (saveDraftBtn) saveDraftBtn.style.display = editable ? 'inline-flex' : 'none';
    submitBtn.style.display = editable ? 'inline-flex' : 'none';
    showForm();
    updateStatus(`Status atual: ${item.Status}`, 'info');
  }

  async function saveDraft() {
    const data = getProjectData();
    const milestones = getMilestonesData();
    const payload = {
      Title: data.nome,
      AnoAprovacao: data.ano_aprovacao,
      CapexBudgetBRL: data.capex_budget_brl,
      Centro: data.centro,
      KpiDescricao: data.kpiDesc,
      KpiImpactado: data.kpi,
      KpiValorAtual: data.kpi_actual,
      KpiValorEsperado: data.kpi_expected,
      LocalImplantacao: data.local_implantacao,
      NecessidadeNegocio: data.necessidade,
      ProjectLeader: data.project_leader,
      ProjectUser: data.project_user,
      SolucaoProposta: data.solucao,
      TipoAtivo: data.tipo_ativo,
      TipoInvestimento: data.tipo_investimento,
      Unidade: data.unidade,
      VidaUtilAnos: data.vida_util,
      Status: 'Rascunho'
    };
    updateStatus('Salvando rascunho...', 'info');
    try {
      let info;
      const projetos = SharePoint.getLista('Projetos');
      if (currentProjectId) {
        info = await projetos.updateItem(currentProjectId, payload);
      } else {
        info = await projetos.addItem(payload);
        currentProjectId = info.d?.Id || info.d?.ID;
      }
      await clearProjectStructure(currentProjectId);
      await saveProjectStructure(currentProjectId, milestones);
      updateStatus('Rascunho salvo.', 'success');
      await loadUserProjects();
    } catch (e) {
      updateStatus('Erro ao salvar rascunho.', 'error');
    }
  }

  async function updateProjectStatus(id, status) {
    try {
      await SharePoint.getLista('Projetos').updateItem(id, { Status: status });
      await loadUserProjects();
      if (currentProjectId === id) {
        await editProject(id);
      }
    } catch (e) {
      console.error(e);
    }
  }

  function addMilestone(nameDefault) {
    milestoneCount++;
    const node = milestoneTpl.content.cloneNode(true);
    const ms = node.querySelector('[data-milestone]');
    const nameInput = node.querySelector('.milestone-name');
    const nameSummaryHeader = node.querySelector('summary');
    const actsWrap = node.querySelector('[data-activities]');
    const btnAddAct = node.querySelector('[data-add-activity]');
    const btnRemove = node.querySelector('[data-remove-milestone]');
    
    nameInput.addEventListener('input', e=>(nameSummaryHeader.textContent = e.target.value));

    nameInput.value = nameDefault || `Milestone ${milestoneCount}`;
    nameSummaryHeader.textContent = nameInput.value;

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
    // Adiciona a primeira atividade por padrão para incentivar o preenchimento
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

  function addActivity(container) {
    const node = activityTpl.content.cloneNode(true);
    const act = node.querySelector('[data-activity]');
    const btnRemove = node.querySelector('[data-remove-activity]');
    const start = node.querySelector('.act-start');
    const end = node.querySelector('.act-end');
    const yearWrap = node.querySelector('[data-year-fields]');

    // regra: início <= fim
    function validateDates() {
      if (start.value && end.value && start.value > end.value) {
        end.setCustomValidity('A data de fim não pode ser anterior à data de início.');
      } else {
        end.setCustomValidity('');
      }
    }
    function updateYearFields() {
      //yearWrap.innerHTML = '';
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
          <div class="c-4">
            <label>Valor CAPEX da atividade (BRL) - ${y}</label>
            <input type="number" class="act-capex" data-year="${y}" min="0" step="0.01" inputmode="decimal" required placeholder="Ex.: 250000,00" />
          </div>
          <div class="c-8">
            <label>Descrição - ${y}</label>
            <textarea class="act-desc" data-year="${y}" required placeholder="Detalhe a atividade, entregáveis e premissas."></textarea>
          </div>
        `;
        yearWrap.appendChild(row);
      }

      [...yearWrap.querySelectorAll('.row[data-year]')].forEach(ye=>{
        if(!years.includes(ye.dataset.year)) ye.remove();
      })

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

  function getProjectData(){
    return {
        nome: getValueFromSelector('projectName').trim(),
        numero_solicitacao: getValueFromSelector('requestNumber').trim(),
        ano_aprovacao: parseInt(getValueFromSelector('approvalYear', 0), 10),
        codigo: getValueFromSelector('projectCode').trim(),
        unidade: getValueFromSelector('unit').trim(),
        centro: getValueFromSelector('center').trim(),
        local_implantacao: getValueFromSelector('projectLocation').trim(),
        project_user: getValueFromSelector('projectUser').trim(),
        project_leader: getValueFromSelector('projectLeader').trim(),
        tipo_investimento: getValueFromSelector('investmentType').trim(),
        tipo_ativo: getValueFromSelector('assetType').trim(),
        vida_util: parseInt(getValueFromSelector('usefulLife', 0), 10),
        capex_budget_brl: parseNumberBRL(getValueFromSelector('projectValue')),
        necessidade: getValueFromSelector('necessidade', "").trim(),
        solucao: getValueFromSelector('solucao', "").trim(),
        kpi: getValueFromSelector('kpi', "").trim(),
        kpiDesc: getValueFromSelector('kpiDesc', "").trim(),
        kpi_actual: parseNumberBRL(getValueFromSelector('kpi_actual', 0).trim()),
        kpi_expected: parseNumberBRL(getValueFromSelector('kpi_expected', 0).trim())
      }
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
          capex_brl: parseNumberBRL(getValueFromSelector('.act-capex'), 0, row),
          descricao: getValueFromSelector('.act-desc', "", row).trim(),
        }));
        return {
          titulo: getValueFromSelector('.act-title', "", a).trim(),
          inicio: getValueFromSelector('.act-start', today, a),
          fim: getValueFromSelector('.act-end', today, a),
          elementoPep: getValueFromSelector('[name="kpi"]', "", a),
          anual,
        };
      });
      milestones.push({ nome, atividades: acts });
    });
    return milestones;
  }

  async function clearProjectStructure(projectId) {
    const Marcos = SharePoint.getLista('Marcos');
    const Atividades = SharePoint.getLista('Atividades1');
    const Alocacoes = SharePoint.getLista('AlocacoesAnuais');

    const msRes = await Marcos.getItems({ select: 'Id', filter: `ProjetoId eq ${projectId}` });
    const marcos = msRes.d?.results || [];
    for (const ms of marcos) {
      const actRes = await Atividades.getItems({ select: 'Id', filter: `MarcoId eq ${ms.Id}` });
      const acts = actRes.d?.results || [];
      for (const act of acts) {
        const alRes = await Alocacoes.getItems({ select: 'Id', filter: `AtividadeId eq ${act.Id}` });
        const als = alRes.d?.results || [];
        for (const al of als) {
          await Alocacoes.deleteItem(al.Id);
        }
        await Atividades.deleteItem(act.Id);
      }
      await Marcos.deleteItem(ms.Id);
    }
  }

  async function saveProjectStructure(projectId, milestones) {
    const Marcos = SharePoint.getLista('Marcos');
    const Atividades = SharePoint.getLista('Atividades1');
    const Alocacoes = SharePoint.getLista('AlocacoesAnuais');
    for (const milestone of milestones) {
      const infoMarco = await Marcos.addItem({ Title: milestone.nome, ProjetoId: projectId });
      const marcoId = infoMarco.d?.Id || infoMarco.d?.ID;
      for (const atividade of milestone.atividades || []) {
        const infoAtv = await Atividades.addItem({
          Title: atividade.titulo,
          DataInicio: atividade.inicio,
          DataFim: atividade.fim,
          ElementoPEP: atividade.elementoPep,
          MarcoId: marcoId
        });
        const atvId = infoAtv.d?.Id || infoAtv.d?.ID;
        for (const anual of atividade.anual || []) {
          await Alocacoes.addItem({
            Title: '',
            Ano: anual.ano,
            CapexBRL: anual.capex_brl,
            Descricao: anual.descricao,
            AtividadeId: atvId
          });
        }
      }
    }
  }

  async function fetchProjectStructure(projectId) {
    const Marcos = SharePoint.getLista('Marcos');
    const Atividades = SharePoint.getLista('Atividades1');
    const Alocacoes = SharePoint.getLista('AlocacoesAnuais');
    const msRes = await Marcos.getItems({ select: 'Id,Title', filter: `ProjetoId eq ${projectId}` });
    const result = [];
    for (const ms of msRes.d?.results || []) {
      const actRes = await Atividades.getItems({ select: 'Id,Title,DataInicio,DataFim,ElementoPEP', filter: `MarcoId eq ${ms.Id}` });
      const acts = [];
      for (const act of actRes.d?.results || []) {
        const alRes = await Alocacoes.getItems({ select: 'Ano,CapexBRL,Descricao', filter: `AtividadeId eq ${act.Id}` });
        const anual = (alRes.d?.results || []).map(a => ({ ano: a.Ano, capex_brl: a.CapexBRL, descricao: a.Descricao }));
        acts.push({ titulo: act.Title, inicio: act.DataInicio, fim: act.DataFim, elementoPep: act.ElementoPEP, anual });
      }
      result.push({ nome: ms.Title, atividades: acts });
    }
    return result;
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
        actNode.querySelector('[name="kpi"]').value = act.elementoPep || '';
        (act.anual || []).forEach(a => {
          const row = actNode.querySelector(`.row[data-year="${a.ano}"]`);
          if (row) {
            row.querySelector('.act-capex').value = a.capex_brl;
            row.querySelector('.act-desc').value = a.descricao;
          }
        });
      });
    });
    refreshGantt();
  }

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
    milestones.forEach((ms, i) => {
      idCounter++;
      let msStart = null;
      let msEnd = null;
      const actRows = [];
      // calcula o intervalo do marco
      ms.atividades.forEach((act, j) => {
        const startDate = act.inicio ? new Date(act.inicio) : new Date();
        const endDate = act.fim ? new Date(act.fim) : new Date( startDate.getTime() + 1000 * 60 * 60 * 24 );
        if (startDate && (!msStart || startDate < msStart)) msStart = startDate;
        if (endDate && (!msEnd || endDate > msEnd)) msEnd = endDate;
        const totalCapex = (act.anual || []).reduce((sum, y) => sum + (y.capex_brl || 0), 0);
        const descTooltip = (act.anual || []).map(y => `${y.ano}: ${BRL.format(y.capex_brl)} - ${y.descricao}`).join('<br/>');
        actRows.push([
            `ms-${idCounter}-${j}`,//Task ID
            act.titulo || `Atividade ${j+1}`, //Task Name
            "Atividade", //Resource
            startDate, //Start Date
            endDate, //End Date
            null, //Duration
            0, //Percent Complete
            `ms-${idCounter}`, //Dependencies
            `CAPEX total: ${BRL.format(totalCapex)}${descTooltip ? '<br/>' + descTooltip : ''}`    //tooltip
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
          { color: '#460a78', dark: '#be2846', light: '#e63c41' }, // milestones nas cores da empresa
          { color: '#f58746', dark: '#e63c41', light: '#ffbe6e' } // atividades nas cores da empresa
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

  // UI events
  addMilestoneBtn.addEventListener('click', () => addMilestone());
  milestoneSection.addEventListener('input', refreshGantt);
  milestoneSection.addEventListener('change', refreshGantt);
  if (saveDraftBtn) saveDraftBtn.addEventListener('click', saveDraft);
  if (backBtn) backBtn.addEventListener('click', showProjectList);

  projectValueInput.addEventListener('input', () => {
    updateCapexFlag();
    updateMilestoneVisibility();
    refreshGantt();
  });
  projectValueInput.addEventListener('change', () => {
    updateCapexFlag();
    updateMilestoneVisibility();
    refreshGantt();
  });

  // Validation
  function validateForm() {
    const errs = [];
    const errsEl = [];
    errorsBox.style.display = 'none';
    errorsBox.innerHTML = '';

    // Valida campos básicos do projeto
    const reqFields = [
      { id: 'projectName', label: 'Nome do Projeto' },
      // { id: 'requestNumber', label: 'Número da solicitação' },
      { id: 'approvalYear', label: 'Ano de aprovação do Projeto' },
      // { id: 'projectCode', label: 'Código do projeto' },
      { id: 'unit', label: 'Unidade' },
      { id: 'center', label: 'Centro' },
      { id: 'projectLocation', label: 'Local da implantação do projeto' },
      { id: 'projectUser', label: 'Project User do Projeto' },
      { id: 'projectLeader', label: 'Project Leader' },
      { id: 'investmentType', label: 'Tipo de investimento' },
      { id: 'assetType', label: 'Tipo de ativo' },
      { id: 'usefulLife', label: 'Vida útil do projeto' },
      { id: 'projectValue', label: 'CAPEX BUDGET do Projeto' },
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

    // Requisito de marcos se CAPEX > 1,5 mi
    const mustHaveMilestone = overThreshold();
    const milestones = [...milestonesWrap.querySelectorAll('[data-milestone]')];

    if (mustHaveMilestone && milestones.length === 0) {
      errs.push('O valor CAPEX é superior a R$ 1,5 milhão. Adicione pelo menos <strong>1 marco</strong>.');
    }

    // Para cada marco: nome e pelo menos 1 atividade com campos válidos
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
        const yearRows = [...a.querySelectorAll('.row[data-year]')];

        if (!title.value.trim()) errs.push(`Atividade ${jdx} do marco ${idx}: informe o <strong>título</strong>.`);
        if (!start.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de início</strong>.`);
        if (!end.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de fim</strong>.`);
        if (start.value && end.value && start.value > end.value) {
          errs.push(`Atividade ${jdx} do marco ${idx}: a <strong>data de início</strong> não pode ser posterior à <strong>data de fim</strong>.`);
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

          console.log(getValueFromSelector('.act-desc', "", row), row, ano);
          if (getValueFromSelector('.act-desc', "", row).trim().length === 0) {
            errs.push(`Atividade ${jdx} do marco ${idx}, ano ${ano}: informe a <strong>descrição</strong>.`);
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

  function getValueFromSelector(queryOrId, defaultValue = "", parent = document){
    let e = null;
    try {
      e = typeof parent.getElementById === "function"? parent.getElementById(queryOrId): parent.querySelector('#'+queryOrId);
    } catch (error) {
    } 
    if( e === null ) e = parent.querySelector(queryOrId);
    if( e === null ) e = { value: defaultValue };
    return e.value;
  }

form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    updateCapexFlag();
    updateMilestoneVisibility();
    if (!validateForm()) return;

    updateStatus('Enviando dados...', 'info');
    submitBtn.disabled = true;

    // Se chegou aqui, tudo válido. Monte um payload de exemplo e mostre no console.
    const payload = {
      projeto: getProjectData(),
      milestones: getMilestonesData()
    };

    const Projetos = SharePoint.getLista('Projetos');

    try {
      const infoProjeto = await Projetos.addItem({
        Title: payload.projeto.nome,
        AnoAprovacao: payload.projeto.ano_aprovacao,
        CapexBudgetBRL: payload.projeto.capex_budget_brl,
        Centro: payload.projeto.centro,
        KpiDescricao: payload.projeto.kpiDesc,
        KpiImpactado: payload.projeto.kpi,
        KpiValorAtual: payload.projeto.kpi_actual,
        KpiValorEsperado: payload.projeto.kpi_expected,
        LocalImplantacao: payload.projeto.local_implantacao,
        NecessidadeNegocio: payload.projeto.necessidade,
        ProjectLeader: payload.projeto.project_leader,
        ProjectUser: payload.projeto.project_user,
        SolucaoProposta: payload.projeto.solucao,
        TipoAtivo: payload.projeto.tipo_ativo,
        TipoInvestimento: payload.projeto.tipo_investimento,
        Unidade: payload.projeto.unidade,
        VidaUtilAnos: payload.projeto.vida_util
      });

      await saveProjectStructure(infoProjeto.d.ID, payload.milestones);
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

  // Estado inicial
  updateCapexFlag();
  updateMilestoneVisibility();
  refreshGantt();
  loadUserProjects();
})();