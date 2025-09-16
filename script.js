//
//  Script principal do protótipo CAPEX Forms.
//  Aqui concentro tanto a camada de integração com SharePoint quanto os
//  comportamentos da interface montada em HTML estático.
//
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

//
//  A partir daqui começa a lógica de interface, encapsulada em uma IIFE
//  para não poluir o escopo global quando o script é carregado no SharePoint.
//
(function () {
  // Formatação monetária utilizada em toda a interface
  const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  // Limiar que define quando o fluxo de marcos precisa ser exibido
  const REQ_THRESHOLD = 500000; // 1.5 milhões

  // Cliente já configurado apontando para o site utilizado nas demonstrações
  const SharePoint = new SPRestApi('https://arcelormittal.sharepoint.com/sites/controladorialongos/capex/');

  // Cache de elementos chave presentes no formulário e na área de leitura
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
  const projectList = document.getElementById('projectList');
  const projectDetails = document.getElementById('projectDetails');
  const appContainer = document.getElementById('app');
  const newProjectBtn = document.getElementById('newProjectBtn');
  const saveDraftBtn = document.getElementById('saveDraftBtn');
  const backBtn = document.getElementById('backBtn');
  google.charts.load('current', { packages: ['gantt'] });

  // Templates HTML usados para gerar marcos e atividades dinamicamente
  const milestoneTpl = document.getElementById('milestoneTemplate');
  const activityTpl = document.getElementById('activityTemplate');

  const currentYear = new Date().getFullYear();
  approvalYearInput.max = currentYear;
  approvalYearInput.placeholder = currentYear;

  // Estados auxiliares controlando marcos, projeto atual e reset silencioso
  let milestoneCount = 0;
  let currentProjectId = null;
  let resetFormWithoutAlert = true;

  // Helpers
  // Feedback visual ao usuário sobre o progresso das ações
  function updateStatus(message = '', type = 'info') {
    if (!statusBox) return;
    statusBox.textContent = message;
    statusBox.className = `status ${type}`;
  }

  // Converte valores com formatação brasileira em números nativos
  function parseNumberBRL(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    // aceita ponto ou vírgula como separador decimal
    const normalized = String(val).replace(/\./g, '').replace(',', '.').replace(/[^\d.]/g, '');
    return Number(normalized || 0);
  }

  // Aponta se o orçamento atual excede o limite para marcos obrigatórios
  function overThreshold() {
    return parseNumberBRL(projectBudgetInput.value) > REQ_THRESHOLD;
  }

  // Atualiza a legenda que orienta quando marcos devem ser adicionados
  function updateCapexFlag() {
    const n = parseNumberBRL(projectBudgetInput.value);
    if (!n) { capexFlag.textContent = ''; return; }
    capexFlag.innerHTML = n > REQ_THRESHOLD
      ? `<span class="ok">Orçamento do Projeto ${BRL.format(n)} &gt; ${BRL.format(REQ_THRESHOLD)} — marcos obrigatórios.</span>`
      : `Orçamento do Projeto ${BRL.format(n)} ≤ ${BRL.format(REQ_THRESHOLD)} — marcos não necessários.`;
  }

  // Esconde ou revela a seção de marcos de acordo com o orçamento
  function updateMilestoneVisibility() {
    const show = overThreshold();
    milestoneSection.style.display = show ? 'block' : 'none';
    if (!show) {
      milestonesWrap.innerHTML = '';
      milestoneCount = 0;
      refreshGantt();
    }
  }

  // Limpa o formulário e volta ao estado padrão
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

  // Alterna a UI para o modo de edição/cadastro
  function showForm() {
    if (appContainer) appContainer.style.display = 'none';
    form.style.display = 'block';
    if (backBtn) backBtn.style.display = 'inline-flex';
    if (newProjectBtn) newProjectBtn.style.display = 'none';
    document.body.style.overflow = 'auto';
  }

  // Retorna para a visão em lista e oculta o formulário
  function showProjectList() {
    form.style.display = 'none';
    if (appContainer) appContainer.style.display = 'flex';
    if (backBtn) backBtn.style.display = 'none';
    if (newProjectBtn) newProjectBtn.style.display = 'inline-block';
    resetForm();
    document.body.style.overflow = 'hidden';
  }

  // Mantém consistência das cores exibidas no selo de status
  function getStatusColor(status) {
    switch (status) {
      case 'Rascunho': return '#414141';
      case 'Em Aprovação': return '#970886';
      case 'Recusado': return '#f83241';
      case 'Aprovado': return '#fe8f46';
      default: return '#414141';
    }
  }

  // Formata datas vindas do SharePoint para o padrão brasileiro
  function formatDate(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      return isNaN(d) ? '' : d.toLocaleDateString('pt-BR');
    } catch (e) {
      return '';
    }
  }

  // Renderiza os detalhes resumidos do projeto no painel principal
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
    statusBadge.textContent = item.Status || '';
    statusBadge.style.background = getStatusColor(item.Status);
    header.append(titleEl, statusBadge);

    const grid = document.createElement('div');
    grid.className = 'details-grid';

    const budgetCard = createDetailCard('Orçamento', BRL.format(item.CapexBudgetBRL || 0), 'budget-value');
    const responsible = createDetailCard('Responsável', item.Responsavel || item.ProjectLeader || '');
    const startDate = createDetailCard('Data de Início', formatDate(item.DataInicio));
    const endDate = createDetailCard('Data de Conclusão', formatDate(item.DataFim || item.DataConclusao));

    const descriptionCard = createDetailCard('Descrição do Projeto', item.SumarioProjeto || item.NecessidadeNegocio || item.ComentarioProjeto || item.Descricao || '');
    descriptionCard.classList.add('detail-desc');

    grid.append(budgetCard, responsible, startDate, endDate, descriptionCard);

    const actions = document.createElement('div');
    actions.className = 'detail-actions';

    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn secondary action-btn';
    editBtn.id = 'editProjectDetails';
    editBtn.textContent = 'Editar Projeto';
    actions.appendChild(editBtn);

    if (item.Status === 'Rascunho' || item.Status === 'Reprovado para Revisão') {
      const approveBtn = document.createElement('button');
      approveBtn.type = 'button';
      approveBtn.className = 'btn primary action-btn approve';
      approveBtn.textContent = 'Enviar para Aprovação';
      actions.appendChild(approveBtn);
    }

    wrapper.append(header, grid, actions);
    projectDetails.appendChild(wrapper);

    const isEditable = item.Status !== 'Aprovado';
    if (isEditable) {
      editBtn.addEventListener('click', () => editProject(item.Id));
    } else {
      editBtn.disabled = true;
    }
  }

  // Botão superior que leva o usuário direto para o formulário de criação
  if (newProjectBtn) {
    newProjectBtn.addEventListener('click', () => {
      resetFormWithoutAlert = true;
      resetForm();
      showForm();
      updateStatus('', 'info');
      resetFormWithoutAlert = false;
    });
  }

  // Busca e renderiza cartões com os projetos do usuário atual
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

      const statusBadge = document.createElement('span');
      statusBadge.className = 'status-badge';
      statusBadge.style.background = getStatusColor(item.Status);
      statusBadge.textContent = item.Status || '';

      const title = document.createElement('h3');
      title.textContent = item.Title || '';

      const budget = document.createElement('p');
      budget.textContent = BRL.format(item.CapexBudgetBRL || 0);

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

  // Preenche o formulário com os dados recuperados do SharePoint
  function fillForm(item) {
    document.getElementById('projectName').value = item.Title || '';
    document.getElementById('approvalYear').value = item.AnoAprovacao || '';
    document.getElementById('projectBudget').value = item.CapexBudgetBRL || '';
    document.getElementById('investmentLevel').value = item.NivelInvestimento || '';
    document.getElementById('fundingSource').value = item.OrigemVerba || '';
    document.getElementById('unit').value = item.Unidade || '';
    document.getElementById('center').value = item.Centro || '';
    document.getElementById('projectLocation').value = item.LocalImplantacao || '';
    document.getElementById('projectUser').value = item.ProjectUser || '';
    document.getElementById('projectLeader').value = item.ProjectLeader || '';
    document.getElementById('company').value = item.Empresa || '';
    document.getElementById('depreciationCostCenter').value = item.CCustoDepreciacao || '';
    document.getElementById('category').value = item.Categoria || '';
    document.getElementById('investmentType').value = item.TipoInvestimento || '';
    document.getElementById('assetType').value = item.TipoAtivo || '';
    document.getElementById('projectFunction').value = item.FuncaoProjeto || '';
    document.getElementById('startDate').value = item.DataInicio ? item.DataInicio.substring(0, 10) : (item.DataInicioProjeto ? item.DataInicioProjeto.substring(0,10) : '');
    document.getElementById('endDate').value = item.DataFim ? item.DataFim.substring(0, 10) : (item.DataFimProjeto ? item.DataFimProjeto.substring(0,10) : '');
    document.getElementById('projectSummary').value = item.SumarioProjeto || item.NecessidadeNegocio || '';
    document.getElementById('projectComment').value = item.ComentarioProjeto || item.SolucaoProposta || '';
    document.getElementById('kpiType').value = item.TipoKPI || item.KpiImpactado || '';
    document.getElementById('kpiName').value = item.NomeKPI || '';
    document.getElementById('kpiDescription').value = item.KpiDescricao || '';
    document.getElementById('kpiCurrent').value = item.KpiValorAtual || '';
    document.getElementById('kpiExpected').value = item.KpiValorEsperado || '';
    updateCapexFlag();
    updateMilestoneVisibility();
  }

  // Abre um projeto específico em modo de edição quando permitido
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

  // Persiste o formulário como rascunho e salva a estrutura hierárquica
  async function saveDraft() {
    const data = getProjectData();
    const milestones = getMilestonesData();
    const payload = {
      Title: data.nome,
      AnoAprovacao: data.ano_aprovacao,
      CapexBudgetBRL: data.capex_budget_brl,
      NivelInvestimento: data.nivel_investimento,
      OrigemVerba: data.origem_verba,
      ProjectUser: data.project_user,
      ProjectLeader: data.project_leader,
      Empresa: data.empresa,
      Centro: data.centro,
      Unidade: data.unidade,
      LocalImplantacao: data.local_implantacao,
      CCustoDepreciacao: data.ccusto_depreciacao,
      Categoria: data.categoria,
      TipoInvestimento: data.tipo_investimento,
      TipoAtivo: data.tipo_ativo,
      FuncaoProjeto: data.funcao_projeto,
      DataInicio: data.data_inicio || null,
      DataFim: data.data_fim || null,
      DataInicioProjeto: data.data_inicio || null,
      DataFimProjeto: data.data_fim || null,
      SumarioProjeto: data.sumario,
      ComentarioProjeto: data.comentario,
      NecessidadeNegocio: data.sumario,
      SolucaoProposta: data.comentario,
      TipoKPI: data.kpi_tipo,
      KpiImpactado: data.kpi_tipo,
      NomeKPI: data.kpi_nome,
      KpiDescricao: data.kpi_descricao,
      KpiValorAtual: data.kpi_atual,
      KpiValorEsperado: data.kpi_esperado,
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

  // Atualiza rapidamente o status do item e re-renderiza a lista
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

  // Cria um novo marco e garante que ele venha com uma atividade inicial
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
    // Adiciona a primeira atividade por padrão para incentivar o preenchimento
    const justAdded = milestonesWrap.lastElementChild.querySelector('[data-activities]');
    addActivity(justAdded);
    renumberMilestones();
    refreshGantt();
  }

  // Atualiza a numeração padrão dos marcos conforme adições/remoções
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

  // Insere uma atividade dentro de um marco, controlando datas e alocações
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
            <textarea class="act-desc" data-year="${y}" required maxlength="600" placeholder="Detalhe a atividade, entregáveis e premissas."></textarea>
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

  // Coleta os campos principais do formulário para montar o payload do projeto
  function getProjectData(){
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
    }
  }

  // Extrai toda a hierarquia de marcos, atividades e alocações anuais
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
          descricao: getValueFromSelector('.act-desc', "", row).trim(),
        }));
        return {
          titulo: getValueFromSelector('.act-title', "", a).trim(),
          inicio: getValueFromSelector('.act-start', today, a),
          fim: getValueFromSelector('.act-end', today, a),
          elementoPep: getValueFromSelector('[name="kpi"]', "", a),
          descricao: getValueFromSelector('.act-overview', "", a).trim(),
          fornecedor: getValueFromSelector('.act-supplier', "", a).trim(),
          anual,
        };
      });
      milestones.push({ nome, atividades: acts });
    });
    return milestones;
  }

  // Remove registros relacionados antes de salvar uma nova versão da estrutura
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

  // Persiste marcos, atividades e alocações nas listas secundárias do SharePoint
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
          DescricaoAtividade: atividade.descricao,
          FornecedorAtividade: atividade.fornecedor,
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

  // Recarrega marcos, atividades e alocações para edição posterior
  async function fetchProjectStructure(projectId) {
    const Marcos = SharePoint.getLista('Marcos');
    const Atividades = SharePoint.getLista('Atividades1');
    const Alocacoes = SharePoint.getLista('AlocacoesAnuais');
    const msRes = await Marcos.getItems({ select: 'Id,Title', filter: `ProjetoId eq ${projectId}` });
    const result = [];
    for (const ms of msRes.d?.results || []) {
      const actRes = await Atividades.getItems({ select: 'Id,Title,DataInicio,DataFim,ElementoPEP,DescricaoAtividade,FornecedorAtividade', filter: `MarcoId eq ${ms.Id}` });
      const acts = [];
      for (const act of actRes.d?.results || []) {
        const alRes = await Alocacoes.getItems({ select: 'Ano,CapexBRL,Descricao', filter: `AtividadeId eq ${act.Id}` });
        const anual = (alRes.d?.results || []).map(a => ({ ano: a.Ano, capex_brl: a.CapexBRL, descricao: a.Descricao }));
        acts.push({
          titulo: act.Title,
          inicio: act.DataInicio,
          fim: act.DataFim,
          elementoPep: act.ElementoPEP,
          descricao: act.DescricaoAtividade,
          fornecedor: act.FornecedorAtividade,
          anual
        });
      }
      result.push({ nome: ms.Title, atividades: acts });
    }
    return result;
  }

  // Recria dinamicamente a interface com base nos dados carregados
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
        const overview = actNode.querySelector('.act-overview');
        const supplier = actNode.querySelector('.act-supplier');
        if (overview) overview.value = act.descricao || '';
        if (supplier) supplier.value = act.fornecedor || '';
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

  // Monta o gráfico de Gantt respeitando o estilo e cores definidos
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

  // Atualiza o gráfico sempre que algum dado de marcos/atividades muda
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
  // Principais listeners responsáveis por manter a UI sincronizada com as ações
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

  // Validation
  // Bloco extenso de validações que cobre as regras discutidas com o usuário
  function validateForm() {
    const errs = [];
    const errsEl = [];
    errorsBox.style.display = 'none';
    errorsBox.innerHTML = '';

    // Valida campos básicos do projeto
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
        const overviewEl = a.querySelector('.act-overview');
        const supplierEl = a.querySelector('.act-supplier');
        const yearRows = [...a.querySelectorAll('.row[data-year]')];

        if (!title.value.trim()) errs.push(`Atividade ${jdx} do marco ${idx}: informe o <strong>título</strong>.`);
        if (!start.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de início</strong>.`);
        if (!end.value) errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>data de fim</strong>.`);
        if (start.value && end.value && start.value > end.value) {
          errs.push(`Atividade ${jdx} do marco ${idx}: a <strong>data de início</strong> não pode ser posterior à <strong>data de fim</strong>.`);
        }
        if (!overviewEl || !overviewEl.value.trim()) {
          errs.push(`Atividade ${jdx} do marco ${idx}: informe a <strong>descrição da atividade</strong>.`);
          if (overviewEl) errsEl.push(overviewEl);
        }
        if (!supplierEl || !supplierEl.value.trim()) {
          errs.push(`Atividade ${jdx} do marco ${idx}: informe o <strong>fornecedor da atividade</strong>.`);
          if (supplierEl) errsEl.push(supplierEl);
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

  // Função utilitária para recuperar valores sem duplicar lógica de busca
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

  // Fluxo completo de submissão que simula a integração final com SharePoint
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
        NivelInvestimento: payload.projeto.nivel_investimento,
        OrigemVerba: payload.projeto.origem_verba,
        ProjectUser: payload.projeto.project_user,
        ProjectLeader: payload.projeto.project_leader,
        Empresa: payload.projeto.empresa,
        Centro: payload.projeto.centro,
        Unidade: payload.projeto.unidade,
        LocalImplantacao: payload.projeto.local_implantacao,
        CCustoDepreciacao: payload.projeto.ccusto_depreciacao,
        Categoria: payload.projeto.categoria,
        TipoInvestimento: payload.projeto.tipo_investimento,
        TipoAtivo: payload.projeto.tipo_ativo,
        FuncaoProjeto: payload.projeto.funcao_projeto,
        DataInicio: payload.projeto.data_inicio || null,
        DataFim: payload.projeto.data_fim || null,
        DataInicioProjeto: payload.projeto.data_inicio || null,
        DataFimProjeto: payload.projeto.data_fim || null,
        SumarioProjeto: payload.projeto.sumario,
        ComentarioProjeto: payload.projeto.comentario,
        NecessidadeNegocio: payload.projeto.sumario,
        SolucaoProposta: payload.projeto.comentario,
        TipoKPI: payload.projeto.kpi_tipo,
        KpiImpactado: payload.projeto.kpi_tipo,
        NomeKPI: payload.projeto.kpi_nome,
        KpiDescricao: payload.projeto.kpi_descricao,
        KpiValorAtual: payload.projeto.kpi_atual,
        KpiValorEsperado: payload.projeto.kpi_esperado
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

  // Reset personalizado que confirma com o usuário antes de apagar campos
  form.addEventListener('reset', ev => {
    if (resetFormWithoutAlert === false && !confirm('Tem certeza que deseja limpar o formulário?')) {
      ev.preventDefault();
    } else {
      updateStatus('Formulário limpo.', 'info');
    }
  });

  // Estado inicial
  // Reaplica cálculos e carrega os projetos assim que o script é executado
  updateCapexFlag();
  updateMilestoneVisibility();
  refreshGantt();
  loadUserProjects();
})();