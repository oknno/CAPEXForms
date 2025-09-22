/*!
 * script.js
 * Descrição: Gerencia integração com SharePoint, estado do formulário CAPEX e visualizações auxiliares.
 * Autor: Matheus Okano
 * Versão: v1.0
 * Última atualização: 2025-09-19
 * Nota: Documentação e comentários adicionados sem alterar a lógica do sistema.
 */

// ============================================================================
// Tipagens de domínio via JSDoc para melhor leitura e tooling
// ============================================================================

/**
 * @typedef {('Rascunho'|'Em Aprovação'|'Aprovado'|'Reprovado'|'Reprovado para Revisão')} ProjectStatus
 * Representa os status possíveis de um projeto no fluxo CAPEX.
 */

/**
 * @typedef {Object} Project
 * @property {number} Id - Identificador único do item SharePoint.
 * @property {string} Title - Nome do projeto exibido em cards e resumos.
 * @property {ProjectStatus} [status] - Situação atual do projeto.
 * @property {number} [budgetBrl] - Orçamento em reais, utilizado nas validações de PEP.
 * @property {string} [investmentLevel] - Nível de investimento calculado em função do orçamento.
 * @property {string} [company]
 * @property {string} [center]
 * @property {string} [unit]
 * @property {string} [location]
 * @property {string} [depreciationCostCenter]
 * @property {string} [projectUser]
 * @property {string} [projectLeader]
 * @property {string} [businessNeed]
 * @property {string} [proposedSolution]
 * @property {string} [kpiType]
 * @property {string} [kpiName]
 * @property {string} [kpiDescription]
 * @property {string|number} [kpiCurrent]
 * @property {string|number} [kpiExpected]
 * @property {string} [category]
 * @property {string} [investmentType]
 * @property {string} [assetType]
 * @property {number} [approvalYear]
 * @property {string} [startDate]
 * @property {string} [endDate]
 * @property {string} [fundingSource]
 */

/**
 * @typedef {Object} Pep
 * @property {number|null} id - ID SharePoint do PEP, se existir.
 * @property {string} title - Nome do PEP selecionado.
 * @property {string} [titleDisplay] - Texto pronto para exibição.
 * @property {number} amountBrl - Valor em reais associado ao PEP.
 * @property {string} [amountText]
 * @property {number|null} year - Ano de execução do PEP.
 * @property {string} [yearText]
 * @property {string} [type] - Identifica se veio de atividade.
 * @property {number|null} [activityId]
 * @property {string} [activityTitle]
 * @property {number|null} [milestoneId]
 * @property {string} [milestoneTitle]
 */

/**
 * @typedef {Object} Activity
 * @property {number|null} id - ID SharePoint da atividade, quando disponível.
 * @property {number|null} milestoneId - Referência ao marco pai.
 * @property {string} title - Nome da atividade.
 * @property {string|null} startDate - Data inicial no formato ISO.
 * @property {string|null} endDate - Data final no formato ISO.
 * @property {string} description - Descrição detalhada.
 * @property {string} supplier - Fornecedor vinculado.
 * @property {Pep|null} [pep] - PEP associado à atividade.
 */

/**
 * @typedef {Object} Milestone
 * @property {number|null} id - ID do marco em SharePoint.
 * @property {string} title - Nome do marco.
 * @property {Activity[]} activities - Atividades relacionadas ao marco.
 */

/**
 * @typedef {Object} SummaryPayload
 * @property {{id:number|string, displayValues:Object} & Project} project - Dados consolidados do projeto.
 * @property {Milestone[]} milestones - Marcos coletados do formulário.
 * @property {Activity[]} activities - Lista linearizada de atividades.
 * @property {Pep[]} peps - PEPs simples e vinculados às atividades.
 */

// ============================================================================
// Integração com SharePoint via REST API
// ============================================================================
/**
 * Serviço dedicado à comunicação com listas do SharePoint usando REST.
 * Responsável por CRUD de itens e anexos JSON dos resumos.
 */
class SharePointService {
  /**
   * @param {string} siteUrl - URL base do site SharePoint onde as listas residem.
   */
  constructor(siteUrl) {
    this.siteUrl = siteUrl.replace(/\/$/, '');
  }

  /**
   * Constrói o nome da entidade REST conforme convenção SharePoint.
   * @param {string} listName - Nome amigável da lista.
   * @returns {string} Nome da entidade REST no formato SP.Data.<List>ListItem.
   */
  encodeEntity(listName) {
    return `SP.Data.${listName.replace(/ /g, '_x0020_').replace(/_/g, '_x005f_')}ListItem`;
  }

  /**
   * Monta a URL de acesso à lista e rota desejada.
   * @param {string} listName - Título da lista no SharePoint.
   * @param {string} [path='/items'] - Segmento adicional para recursos (ex.: /items(1)).
   * @returns {string} URL completa pronta para uso em fetch.
   * @throws {Error} Quando o nome da lista não é informado.
   */
  buildUrl(listName, path = '/items') {
    if (!listName) {
      throw new Error('Lista SharePoint não informada.');
    }
    const safeListName = String(listName).replace(/'/g, "''");
    return `${this.siteUrl}/_api/web/lists/getbytitle('${safeListName}')${path}`;
  }

  /**
   * Normaliza nome de arquivo de anexo para evitar caracteres ilegais.
   * @param {string} fileName - Nome original informado.
   * @returns {string} Nome sanitizado pronto para upload.
   */
  sanitizeFileName(fileName) {
    if (typeof fileName !== 'string') {
      return '';
    }
    const normalized = fileName.normalize ? fileName.normalize('NFKC') : fileName;
    const trimmed = normalized.trim();
    if (!trimmed) {
      return '';
    }
    const withoutControl = trimmed.replace(/[\u0000-\u001f\u007f]/g, '');
    const replacedInvalid = withoutControl.replace(/[\\/:*?"<>|]/g, '_');
    const collapsedSpaces = replacedInvalid.replace(/\s+/g, ' ').trim();
    if (!collapsedSpaces) {
      return '';
    }
    return collapsedSpaces.replace(/'/g, "''");
  }

  /**
   * Executa requisição REST com logging e tratamento de erros padrão.
   * @param {string} url - URL alvo no SharePoint.
   * @param {RequestInit} [options={}] - Configuração fetch (método, headers, body).
   * @returns {Promise<null|Object>} Corpo JSON parseado ou null para 204/resposta vazia.
   * @throws {Error} Para falhas de rede ou respostas não OK.
   */
  async request(url, options = {}) {
    let response;
    try {
      // Passo 1: dispara fetch e captura falhas de rede antes da avaliação HTTP
      response = await fetch(url, options);
    } catch (networkError) {
      console.error('Falha na requisição SharePoint', {
        url,
        error: networkError
      });
      throw new Error('Não foi possível conectar ao SharePoint. Tente novamente mais tarde.');
    }

    if (!response.ok) {
      // Passo 2: log detalhado para respostas HTTP não bem-sucedidas
      const responseText = await response.text();
      console.error('Erro retornado pela API do SharePoint', {
        url,
        status: response.status,
        statusText: response.statusText,
        responseText
      });
      const message = responseText || response.statusText || 'Erro desconhecido na API do SharePoint.';
      const error = new Error(message);
      error.status = response.status;
      error.url = url;
      throw error;
    }

    if (response.status === 204) {
      return null;
    }

    const text = await response.text();
    if (!text) {
      return null;
    }

    try {
      // Passo 3: parseia JSON e propaga erro contextual caso a resposta seja inválida
      return JSON.parse(text);
    } catch (parseError) {
      console.error('Não foi possível interpretar a resposta do SharePoint como JSON', {
        url,
        status: response.status,
        rawBody: text
      });
      const error = new Error('Resposta inválida recebida do SharePoint.');
      error.status = response.status;
      error.url = url;
      throw error;
    }
  }

  /**
   * Envia anexo JSON ao SharePoint garantindo validação de conteúdo e tamanho.
   * @param {string} listName - Lista onde o item reside.
   * @param {number|string} itemId - Identificador do item pai.
   * @param {string} fileName - Nome sugerido do arquivo (forçado para .json).
   * @param {Blob|Object|string} fileContent - Conteúdo que será persistido.
   * @param {{overwrite?:boolean}} [options={}] - Controla substituição prévia do arquivo.
   * @returns {Promise<boolean>} Indica sucesso do upload.
   * @throws {Error} Quando validações de conteúdo ou tamanho falham.
   */
async addAttachment(listName, itemId, fileName, fileContent, options = {}) {
  if (!listName) throw new Error('Lista SharePoint não informada.');
  if (!itemId) throw new Error('ID do item inválido para anexar arquivo.');

  // 👉 força sempre extensão .json e content-type correto
  const { overwrite = false } = options;
  const rawFileName = fileName?.endsWith('.json') ? fileName : `resumo_${itemId}.json`;
  const sanitizedFileName = this.sanitizeFileName(rawFileName);
  // encodeURIComponent evita caracteres especiais na rota AttachmentFiles
  // encodeURIComponent protege a rota getByFileName contra caracteres especiais
  const encodedFileName = encodeURIComponent(sanitizedFileName);

  if (!sanitizedFileName) {
    throw new Error('Nome do arquivo inválido.');
  }

  if (overwrite) {
    try {
      await this.deleteAttachment(listName, itemId, sanitizedFileName);
    } catch (error) {
      if (error?.status !== 404) {
        throw error;
      }
    }
  }

  const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024; // 10MB

  // 🔑 Sempre serializa objeto para JSON formatado e valida conteúdo
  let bodyContent = fileContent;
  if (fileContent instanceof Blob) {
    bodyContent = fileContent;
  } else if (typeof fileContent === 'object' && fileContent !== null) {
    try {
      bodyContent = JSON.stringify(fileContent, null, 2);
    } catch (serializationError) {
      console.error('Falha ao serializar o conteúdo do anexo em JSON.', serializationError);
      throw new Error('Não foi possível preparar o anexo JSON.');
    }
  } else if (typeof fileContent === 'string') {
    const trimmed = fileContent.trim();
    if (!trimmed) {
      throw new Error('O conteúdo do anexo JSON está vazio.');
    }
    try {
      JSON.parse(trimmed);
    } catch (jsonError) {
      console.error('Conteúdo do anexo não é um JSON válido.', jsonError);
      throw new Error('O conteúdo do anexo não é um JSON válido.');
    }
    bodyContent = trimmed;
  } else if (fileContent !== undefined && fileContent !== null) {
    const textContent = String(fileContent).trim();
    if (!textContent) {
      throw new Error('O conteúdo do anexo JSON está vazio.');
    }
    try {
      JSON.parse(textContent);
    } catch (jsonError) {
      console.error('Conteúdo do anexo não é um JSON válido.', jsonError);
      throw new Error('O conteúdo do anexo não é um JSON válido.');
    }
    bodyContent = textContent;
  } else {
    throw new Error('O conteúdo do anexo JSON está vazio.');
  }

  if (!(bodyContent instanceof Blob)) {
    const normalizedContent =
      typeof bodyContent === 'string' ? bodyContent.trim() : String(bodyContent ?? '').trim();
    if (!normalizedContent) {
      throw new Error('O conteúdo do anexo JSON está vazio.');
    }
    bodyContent = normalizedContent;
  }

  const digest = await this.getFormDigest();
  const headers = {
    Accept: 'application/json;odata=verbose',
    'X-RequestDigest': digest,
    'Content-Type': 'application/json'
  };

  const url = this.buildUrl(
    listName,
    `/items(${itemId})/AttachmentFiles/add(FileName='${encodedFileName}')`
  );

  const body = bodyContent instanceof Blob
    ? bodyContent
    : new Blob([bodyContent], { type: 'application/json' });

  if (body.size > MAX_ATTACHMENT_SIZE) {
    throw new Error('O anexo excede o tamanho máximo permitido de 10MB.');
  }

  console.log("🔎 Salvando anexo em:", url, "Arquivo:", sanitizedFileName);

  await this.request(url, { method: 'POST', headers, body });
  return true;
}

  /**
   * Remove anexo JSON existente, validando presença antes de efetuar POST de exclusão.
   * @param {string} listName - Lista alvo.
   * @param {number|string} itemId - ID do item pai.
   * @param {string} fileName - Nome do arquivo a remover.
   * @returns {Promise<boolean>} True quando exclusão ocorre ou arquivo não existe.
   * @throws {Error} Propaga falhas de rede exceto 404 ignorado quando overwrite.
   */
async deleteAttachment(listName, itemId, fileName) {
  if (!listName) throw new Error('Lista SharePoint não informada.');
  if (!itemId) throw new Error('ID do item inválido para remover anexo.');

  const rawFileName = fileName || `resumo_${itemId}.json`;
  const sanitizedFileName = this.sanitizeFileName(rawFileName);
  if (!sanitizedFileName) {
    console.warn("⚠️ Nome inválido em deleteAttachment, abortando:", fileName);
    return false;
  }

  const encodedFileName = encodeURIComponent(sanitizedFileName);

  // 🔎 Verifica se o anexo existe antes de tentar remover
  const attachmentsUrl = this.buildUrl(listName, `/items(${itemId})/AttachmentFiles`);
  const headersCheck = { Accept: 'application/json;odata=verbose' };
  const attachments = await this.request(attachmentsUrl, { method: 'GET', headers: headersCheck });

  const found = attachments?.d?.results?.some(
    (att) => att.FileName && att.FileName.toLowerCase() === sanitizedFileName.toLowerCase()
  );

  if (!found) {
    console.log(`⚠️ Anexo "${sanitizedFileName}" não existe no item ${itemId}, ignorando exclusão.`);
    return false;
  }

  const digest = await this.getFormDigest();
  const headers = {
    Accept: 'application/json;odata=verbose',
    'X-RequestDigest': digest,
    'IF-MATCH': '*',
    'X-HTTP-Method': 'DELETE'
  };

  const url = this.buildUrl(
    listName,
    `/items(${itemId})/AttachmentFiles/getByFileName('${encodedFileName}')`
  );

  console.log("🔎 Removendo anexo existente:", url, "Arquivo:", sanitizedFileName);

  await this.request(url, { method: 'POST', headers });
  return true;
}

  /**
   * Obtém token X-RequestDigest necessário para operações de escrita.
   * @returns {Promise<string>} Valor do form digest atual.
   * @throws {Error} Quando SharePoint não retorna digest e _spPageContextInfo não está disponível.
   */
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
        // Fallback SharePoint: reutiliza digest exposto globalmente quando disponível
        return _spPageContextInfo.formDigestValue;
      }
      throw error;
    }
  }

  /**
   * Lista itens de uma lista SharePoint com parâmetros OData opcionais.
   * @param {string} listName - Nome da lista.
   * @param {Object} [params={}] - Parâmetros como select, filter, orderby.
   * @returns {Promise<Object[]>} Coleção de itens no formato JSON padrão.
   */
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

  /**
   * Recupera um item específico da lista.
   * @param {string} listName - Lista alvo.
   * @param {number|string} id - Identificador do item.
   * @returns {Promise<Object|null>} Item retornado pela API ou null.
   */
  async getItem(listName, id) {
    const url = this.buildUrl(listName, `/items(${id})`);
    const headers = { Accept: 'application/json;odata=verbose' };
    const data = await this.request(url, { headers });
    return data?.d ?? null;
  }

  /**
   * Cria novo item na lista informada.
   * @param {string} listName - Lista alvo.
   * @param {Object} payload - Dados do item (campos customizados inclusos).
   * @returns {Promise<Object|null>} Item criado retornado pelo SharePoint.
   */
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

  /**
   * Atualiza item existente utilizando verbo MERGE e cabeçalho IF-MATCH *.
   * @param {string} listName - Lista alvo.
   * @param {number|string} id - ID do item a alterar.
   * @param {Object} payload - Campos a atualizar.
   * @returns {Promise<boolean>} True ao concluir sem erros.
   */
  async updateItem(listName, id, payload) {
    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      // IF-MATCH:* + X-HTTP-Method:MERGE evita conflitos de versão mantendo semântica REST
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

  /**
   * Exclui item via DELETE lógico no SharePoint (X-HTTP-Method: DELETE).
   * @param {string} listName - Lista alvo.
   * @param {number|string} id - ID do item a remover.
   * @returns {Promise<boolean>} Indica que a operação foi concluída.
   */
  async deleteItem(listName, id) {
    const digest = await this.getFormDigest();
    const headers = {
      Accept: 'application/json;odata=verbose',
      'X-RequestDigest': digest,
      // Cabeçalhos padrão SharePoint para exclusão (garantem remoção independente da versão)
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
/*
 * Mantém formatadores, constantes de negócio e caches utilizados em diferentes fluxos.
 * Os seletores DOM abaixo são consumidos por handlers espalhados pelo arquivo; evitar
 * reatribuição desses nós para preservar performance e consistência de estado.
 */
const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
const DATE_FMT = new Intl.DateTimeFormat('pt-BR');
const BUDGET_THRESHOLD = 1_000_000;
const EXCHANGE_RATE = 5.6; // 1 USD = 5.6 BRL
const DATE_RANGE_ERROR_MESSAGE = 'A data de término não pode ser anterior à data de início.';

const SITE_URL = window.SHAREPOINT_SITE_URL || 'https://arcelormittal.sharepoint.com/sites/controladorialongos/capex';
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

const validationState = {
  pepBudget: null,
  pepBudgetDetails: null,
  activityDates: null,
  activityDateDetails: null
};

const EMPTY_ARRAY = Object.freeze([]);

function createCenterDictionary(centerCodes, units, locations) {
  if (!Array.isArray(centerCodes) || !centerCodes.length) return {};
  const normalizedUnits = Array.isArray(units) ? [...units] : [];
  const normalizedLocations = Array.isArray(locations) ? [...locations] : [];
  return centerCodes.reduce((acc, code) => {
    acc[code] = {
      units: normalizedUnits,
      locations: normalizedLocations
    };
    return acc;
  }, {});
}

function normalizeCompanyRules(rules) {
  if (!rules || typeof rules !== 'object') return {};
  return Object.entries(rules).reduce((acc, [companyCode, entry]) => {
    if (!entry || typeof entry !== 'object') return acc;
    const depreciation = Array.isArray(entry.depreciation) ? [...entry.depreciation] : [];
    const centers = {};
    const centersSource = entry.centers;
    if (centersSource && typeof centersSource === 'object' && !Array.isArray(centersSource)) {
      Object.entries(centersSource).forEach(([centerCode, centerEntry]) => {
        if (!centerEntry || typeof centerEntry !== 'object') return;
        centers[centerCode] = {
          units: Array.isArray(centerEntry.units) ? [...centerEntry.units] : [],
          locations: Array.isArray(centerEntry.locations) ? [...centerEntry.locations] : []
        };
      });
    } else if (Array.isArray(centersSource)) {
      const sharedUnits = Array.isArray(entry.units) ? [...entry.units] : [];
      const sharedLocations = Array.isArray(entry.locations) ? [...entry.locations] : [];
      centersSource.forEach((centerCode) => {
        centers[centerCode] = {
          units: sharedUnits,
          locations: sharedLocations
        };
      });
    }
    acc[companyCode] = { depreciation, centers };
    return acc;
  }, {});
}

function markCompanyRulesNormalized(rules) {
  if (rules && typeof rules === 'object' && !rules.__normalized) {
    Object.defineProperty(rules, '__normalized', { value: true, enumerable: false, configurable: true });
  }
  return rules || {};
}

const companyHierarchy = {
  "ACBR": {
    "AC01": {
      "units": [
        "NL - NOVVA LOGÍSTICA LTDA"
      ],
      "locations": [
        "GR",
        "TI"
      ]
    },
    "AC02": {
      "units": [
        "NL - NOVVA LOGÍSTICA LTDA"
      ],
      "locations": [
        "GR",
        "TI"
      ]
    }
  },
  "AMTL": {
    "TL01": {
      "units": [
        "SU - SUPRIMENTOS SITREL",
        "TL - SITREL"
      ],
      "locations": [
        "ET",
        "GR",
        "GT",
        "LA",
        "LE",
        "MA",
        "MS",
        "MU",
        "RH",
        "SU",
        "TI",
        "TL"
      ]
    }
  },
  "BF00": {
    "E001": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "EC"
      ]
    },
    "E201": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "CO"
      ]
    },
    "E202": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "GA"
      ]
    },
    "E204": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "CO"
      ]
    },
    "E205": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "BU",
        "CO"
      ]
    },
    "E210": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "FA"
      ]
    },
    "E213": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": []
    },
    "E501": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "RD"
      ]
    },
    "E507": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "RQ"
      ]
    },
    "E601": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "CB"
      ]
    },
    "E602": {
      "units": [
        "BF - BIOFLORESTAS"
      ],
      "locations": [
        "FQ"
      ]
    }
  },
  "BMJF": {
    "4000": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "SU - JUIZ DE FORA - SUPRIMENTOS",
        "UJ - JUIZ DE FORA - SIDERURGIA",
        "UJ - JUIZ DE FORA - TREFILARIA"
      ],
      "locations": [
        "AC",
        "AF",
        "EU",
        "GR",
        "GT",
        "JF",
        "LA",
        "LE",
        "MA",
        "MS",
        "RH",
        "TI",
        "TR"
      ]
    },
    "4100": {
      "units": [
        "BM - BARRA MANSA - SIDERURGIA",
        "BM - BARRA MANSA - TREFILARIA",
        "CD - COMITÊ DIGITAL",
        "SU - BARRA MANSA - ALMOXARIFADO"
      ],
      "locations": [
        "AC",
        "BM",
        "ET",
        "GR",
        "GT",
        "LA",
        "LE",
        "MA",
        "MS",
        "MU",
        "RH",
        "TI",
        "TR"
      ]
    },
    "4200": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "RS - RESENDE - SIDERURGIA",
        "RS - RESENDE - TREFILARIA",
        "SU - RESENDE - ALMOXARIFADO"
      ],
      "locations": [
        "AC",
        "ET",
        "GR",
        "GT",
        "LA",
        "LE",
        "MA",
        "MS",
        "MU",
        "RH",
        "RS",
        "TI",
        "TR"
      ]
    },
    "5000": {
      "units": [
        "ME - METÁLICOS CONTAGEM"
      ],
      "locations": [
        "BH"
      ]
    },
    "6003": {
      "units": [
        "BF - FLORESTAS SF"
      ],
      "locations": [
        "BS"
      ]
    },
    "6004": {
      "units": [
        "BF - FLORESTAS SF"
      ],
      "locations": [
        "RI"
      ]
    },
    "6006": {
      "units": [
        "BF - FLORESTAS SF"
      ],
      "locations": [
        "SC"
      ]
    },
    "6007": {
      "units": [
        "BF - FLORESTAS SF"
      ],
      "locations": [
        "SR"
      ]
    },
    "7115": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CL"
      ]
    },
    "7500": {
      "units": [
        "BC - VP COMERCIAL",
        "CD - COMITÊ DIGITAL",
        "EC - ESCRITÓRIO CENTRAL - ECA",
        "LP - DIR. LOGÍSTICA E PLAN.",
        "SH - TI - SHARED SERVICE",
        "SU - SUPRIMENTOS CORPORATIVO"
      ],
      "locations": [
        "CO",
        "IA",
        "RD",
        "TI"
      ]
    },
    "7502": {
      "units": [
        "BC - VP COMERCIAL",
        "CD - COMITÊ DIGITAL",
        "LP - DIR. LOGÍSTICA E PLAN.",
        "SH - TI - SHARED SERVICE"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "7510": {
      "units": [
        "BC - VP COMERCIAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9206": {
      "units": [
        "ME - PÁTIO SUC.IRACEMÁPOLIS",
        "UP - PÁTIO SUC.IRACEMÁPOLIS"
      ],
      "locations": [
        "IR"
      ]
    },
    "9216": {
      "units": [
        "ME - ENTREPOSTO MARACANAÚ"
      ],
      "locations": [
        "MA"
      ]
    },
    "9217": {
      "units": [
        "ME - ENTREPOSTO JABOATÃO"
      ],
      "locations": [
        "JA"
      ]
    },
    "9227": {
      "units": [
        "ME - RCS ALVORADA"
      ],
      "locations": [
        "AV"
      ]
    },
    "9307": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CL"
      ]
    },
    "9308": {
      "units": [
        "BC - LOJA SALVADOR"
      ],
      "locations": [
        "RD"
      ]
    },
    "9309": {
      "units": [
        "BC - DBA GUARAPUAVA"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9333": {
      "units": [
        "BC - CDB RIO DE JANEIRO"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9335": {
      "units": [
        "BC - DBA BOA VISTA",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9344": {
      "units": [
        "ME - RCS MANAUS"
      ],
      "locations": [
        "MN"
      ]
    },
    "9345": {
      "units": [
        "BC - DBA PATO BRANCO"
      ],
      "locations": [
        "RD",
        "TI"
      ]
    },
    "9351": {
      "units": [
        "ME - ENTREPOSTO PINHAIS"
      ],
      "locations": [
        "CT"
      ]
    },
    "9353": {
      "units": [
        "BC - LOJA BONSUCESSO"
      ],
      "locations": [
        "RD"
      ]
    },
    "9355": {
      "units": [
        "BC - DBA BOA VISTA"
      ],
      "locations": []
    },
    "9360": {
      "units": [
        "BC - CL CONFINS",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9361": {
      "units": [
        "BC - CDB SÃO PAULO"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9366": {
      "units": [
        "BC - CDB SALVADOR"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9367": {
      "units": [
        "BC - HUB SALVADOR",
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "HB",
        "RD",
        "TI"
      ]
    },
    "9368": {
      "units": [
        "BC - CDB FORTALEZA"
      ],
      "locations": []
    },
    "9372": {
      "units": [
        "BC - LOJA ARICANDUVA"
      ],
      "locations": [
        "RD"
      ]
    },
    "9374": {
      "units": [
        "BC - LOJA MOGI DAS CRUZES"
      ],
      "locations": [
        "RD"
      ]
    },
    "9375": {
      "units": [
        "BC - LOJA SJ PINHAIS"
      ],
      "locations": [
        "RD"
      ]
    },
    "9377": {
      "units": [
        "BC - CL EXTREMA",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "CL",
        "RD",
        "TI"
      ]
    },
    "9378": {
      "units": [
        "BC - LOJA ITAIM PAULISTA"
      ],
      "locations": [
        "RD"
      ]
    },
    "9379": {
      "units": [
        "BC - LOJA OSASCO"
      ],
      "locations": [
        "RD"
      ]
    },
    "9381": {
      "units": [
        "BC - LOJA ITAQUAQUECETUBA"
      ],
      "locations": [
        "RD"
      ]
    },
    "9382": {
      "units": [
        "BC - LOJA BELO HORIZONTE"
      ],
      "locations": [
        "RD"
      ]
    },
    "9397": {
      "units": [
        "BC - LOJA SANTA BÁRBARA"
      ],
      "locations": []
    },
    "9404": {
      "units": [
        "ME - ENTREPOSTO GUARULHOS"
      ],
      "locations": [
        "GU"
      ]
    },
    "9416": {
      "units": [
        "ME - ENTREPOSTO CANDEIAS"
      ],
      "locations": [
        "CA"
      ]
    },
    "9450": {
      "units": [
        "BC - BELGO PRONTO CURITIBA"
      ],
      "locations": [
        "BP"
      ]
    },
    "9460": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "FT - FÁBRICA DE TELAS SP",
        "SU - ALMOXARIFADO FÁBR. TELAS SP"
      ],
      "locations": [
        "EU",
        "GR",
        "GT",
        "LE",
        "MA",
        "MS",
        "RH",
        "SP",
        "TI",
        "TT"
      ]
    },
    "9515": {
      "units": [
        "ME - ENTREPOSTO FIAT GOIANIA"
      ],
      "locations": [
        "GO"
      ]
    },
    "9600": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CL"
      ]
    },
    "9607": {
      "units": [
        "BC - CDB BELO HORIZONTE",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9618": {
      "units": [
        "BC - CDB CURITIBA"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9633": {
      "units": [
        "BC - CDB MARABÁ"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9650": {
      "units": [
        "ME - ENTREPOSTO BAURU"
      ],
      "locations": [
        "BU"
      ]
    },
    "9651": {
      "units": [
        "BC - DBA CASCAVEL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9655": {
      "units": [
        "BC - DBA JUIZ DE FORA"
      ],
      "locations": []
    },
    "9660": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CL"
      ]
    },
    "9661": {
      "units": [
        "BC - DBA RIO DAS PEDRAS",
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CS",
        "TI"
      ]
    },
    "9670": {
      "units": [
        "BC - DBA JUIZ DE FORA",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9682": {
      "units": [
        "BC - HUB RIO DAS PEDRAS",
        "CD - COMITÊ DIGITAL",
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "HB",
        "RD",
        "TI"
      ]
    },
    "9683": {
      "units": [
        "BC - LOJA RECIFE"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9684": {
      "units": [
        "BC - CDB BELÉM"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9689": {
      "units": [
        "BC - CDB RECIFE"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9696": {
      "units": [
        "BC - LOJA GUARULHOS"
      ],
      "locations": [
        "RD"
      ]
    },
    "9703": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "CL"
      ]
    },
    "9718": {
      "units": [
        "BC - CDB SÃO PAULO II"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9719": {
      "units": [
        "BC - VP COMERCIAL"
      ],
      "locations": [
        "BP"
      ]
    },
    "9728": {
      "units": [
        "BC - BELGO PRONTO RIO DAS PEDRAS"
      ],
      "locations": [
        "RD",
        "TI"
      ]
    },
    "9729": {
      "units": [
        "BC - BELGO PRONTO RIO DAS PEDRAS",
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "BP"
      ]
    },
    "9733": {
      "units": [
        "BC - CDB RIO DAS PEDRAS"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9840": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "GA - GUILMAN AMORIM",
        "SU - SUPRIMENTOS - MONLEVADE",
        "UM - MONLEVADE"
      ],
      "locations": [
        "AC",
        "AF",
        "EU",
        "GA",
        "GR",
        "GT",
        "JM",
        "LA",
        "LE",
        "MA",
        "MS",
        "RE",
        "RH",
        "SI",
        "TI"
      ]
    },
    "9860": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "SU - ALMOXARIFADO PIRACICABA",
        "UP - PIRACICABA"
      ],
      "locations": [
        "AC",
        "EU",
        "GR",
        "IR",
        "LA",
        "LE",
        "MA",
        "MS",
        "PI",
        "RH",
        "TI"
      ]
    },
    "9870": {
      "units": [
        "CD - COMITÊ DIGITAL",
        "SU - ALMOXARIFADO SABARÁ",
        "US - SABARÁ"
      ],
      "locations": [
        "BA",
        "EU",
        "GR",
        "LE",
        "MA",
        "MS",
        "RH",
        "SA",
        "TI"
      ]
    },
    "9880": {
      "units": [
        "BC - BELGO PRONTO MARACANAÚ",
        "CD - COMITÊ DIGITAL",
        "LP - DIR. LOGÍSTICA E PLAN."
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "9881": {
      "units": [
        "LP - DIR. LOGÍSTICA E PLAN.",
        "RP - TREFILARIA RIO DAS PEDRAS",
        "SU - ALMOXARIFADO RIO DAS PEDRAS"
      ],
      "locations": [
        "CS",
        "GR",
        "RP"
      ]
    },
    "9910": {
      "units": [
        "BC - CDB FORTALEZA",
        "CD - COMITÊ DIGITAL"
      ],
      "locations": [
        "BP",
        "RD",
        "TI"
      ]
    },
    "EN01": {
      "units": [
        "CE - COMERCIALIZADORA DE ENERGIA"
      ],
      "locations": []
    },
    "MA10": {
      "units": [
        "MA - MINA DO ANDRADE"
      ],
      "locations": [
        "BR",
        "FR",
        "GR",
        "MA",
        "MI",
        "MS",
        "OV",
        "SF",
        "TI",
        "US"
      ]
    },
    "SA20": {
      "units": [
        "MS - MINA SERRA AZUL"
      ],
      "locations": [
        "BR",
        "FR",
        "GR",
        "MA",
        "MI",
        "MS",
        "OV",
        "SF",
        "TI",
        "US"
      ]
    }
  }
};

const companyRules = {
  BMJF: {
    depreciation: ['CC-01', 'CC-02'],
    centers: companyHierarchy.BMJF
  },
  BF00: {
    depreciation: ['CC-03'],
    centers: companyHierarchy.BF00
  },
  ACBR: {
    depreciation: ['CC-03'],
    centers: companyHierarchy.ACBR
  },
  AMTL: {
    depreciation: ['CC-03'],
    centers: companyHierarchy.AMTL
  }
};

const defaultCompanyRules = markCompanyRulesNormalized(normalizeCompanyRules(companyRules));

function resolveCompanyRules() {
  if (typeof window !== 'undefined') {
    const globalRules = window.companyRules;
    if (!globalRules) {
      window.companyRules = defaultCompanyRules;
      return window.companyRules;
    }
    if (!globalRules.__normalized) {
      window.companyRules = markCompanyRulesNormalized(normalizeCompanyRules(globalRules));
    }
    return window.companyRules;
  }
  return defaultCompanyRules;
}

function safeArray(value) {
  return Array.isArray(value) ? value : EMPTY_ARRAY;
}

/**
 * Atualiza o estado local de um projeto específico preservando imutabilidade do array.
 * @param {number|string} projectId - Identificador do projeto a atualizar.
 * @param {Partial<Project>} [changes={}] - Campos alterados retornados após operações CRUD.
 */
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
const floatingCloseBtn = document.querySelector('.form-close-btn');
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

const summaryTitle = document.getElementById('summaryTitle');

const formSummaryView = document.getElementById('formSummaryView');
const formSummarySections = document.getElementById('formSummarySections');
const formSummaryGanttSection = document.getElementById('formSummaryGanttSection');
const formSummaryGanttChart = document.getElementById('formSummaryGanttChart');
const formSummaryCloseBtn = document.getElementById('formSummaryCloseBtn');

if (summaryTitle && !summaryTitle.hasAttribute('tabindex')) {
  summaryTitle.setAttribute('tabindex', '-1');
}

const companySelect = document.getElementById('company');
const centerSelect = document.getElementById('center');
const unitSelect = document.getElementById('unit');
const locationSelect = document.getElementById('location');
const approvalYearInput = document.getElementById('approvalYear');
const projectBudgetInput = document.getElementById('projectBudget');
const investmentLevelSelect = document.getElementById('investmentLevel');
const projectStartDateInput = document.getElementById('startDate');
const projectEndDateInput = document.getElementById('endDate');

const simplePepTemplate = document.getElementById('simplePepTemplate');
const milestoneTemplate = document.getElementById('milestoneTemplate');
const activityTemplate = document.getElementById('activityTemplate');

const PROJECT_STATUSES = Object.freeze({
  APPROVED: 'Aprovado',
  DRAFT: 'Rascunho',
  REJECTED: 'Reprovado',
  REJECTED_FOR_REVIEW: 'Reprovado para Revisão',
  IN_APPROVAL: 'Em Aprovação'
});

/**
 * Retorna a cor utilizada para representar o status do projeto em chips e cards.
 * @param {ProjectStatus|string} status - Status atual do projeto.
 * @returns {string} Cor em hexadecimal.
 */
function statusColor(status) {
  switch (status) {
    case PROJECT_STATUSES.DRAFT:
      return '#414141';
    case PROJECT_STATUSES.IN_APPROVAL:
      return '#970886';
    case PROJECT_STATUSES.REJECTED_FOR_REVIEW:
      return '#fe8f46';
    case PROJECT_STATUSES.APPROVED:
      return '#3d9308';
    case PROJECT_STATUSES.REJECTED:
      return '#f83241';
    default:
      return '#414141';
  }
}

const STATUS_GROUPS = Object.freeze({
  READ_ONLY: new Set([PROJECT_STATUSES.APPROVED, PROJECT_STATUSES.IN_APPROVAL]),
  APPROVAL_ALLOWED: new Set([
    PROJECT_STATUSES.DRAFT,
    PROJECT_STATUSES.REJECTED,
    PROJECT_STATUSES.REJECTED_FOR_REVIEW
  ])
});
const defaultSummaryContext = {
  sections: summarySections,
  ganttSection: summaryGanttSection,
  ganttChart: summaryGanttChart
};

if (typeof window !== 'undefined') {
  resolveCompanyRules();
}

const formSummaryContext = {
  sections: formSummarySections,
  ganttSection: formSummaryGanttSection,
  ganttChart: formSummaryGanttChart
};

let activeSummaryContext = defaultSummaryContext;
let currentFormMode = null;
let renderProjectListFrame = null;

/**
 * Normaliza status removendo espaços extras para comparações simples.
 * @param {string} status - Status retornado do SharePoint.
 * @returns {string} Representação simplificada.
 */
function normalizeStatusKey(status) {
  return typeof status === 'string' ? status.trim() : '';
}

/**
 * Verifica se o status atual deve bloquear edição do formulário.
 * @param {ProjectStatus|string} status - Status do projeto.
 * @returns {boolean} True quando o modo leitura deve ser aplicado.
 */
function isReadOnlyStatus(status) {
  return STATUS_GROUPS.READ_ONLY.has(normalizeStatusKey(status));
}

/**
 * Define se o botão "Enviar para Aprovação" deve estar disponível.
 * @param {ProjectStatus|string} status - Status atual.
 * @returns {boolean} True quando envio é permitido.
 */
function canSubmitForApproval(status) {
  const key = normalizeStatusKey(status);
  return !key || STATUS_GROUPS.APPROVAL_ALLOWED.has(key);
}

/**
 * Utilitário de debounce para evitar múltiplas execuções durante digitação.
 * @param {Function} fn - Função original a ser adiada.
 * @param {number} [delay=200] - Janela em milissegundos.
 * @returns {Function} Função decorada que respeita o intervalo informado.
 */
function debounce(fn, delay = 200) {
  let timerId = null;
  return function debouncedFunction(...args) {
    if (timerId) {
      clearTimeout(timerId);
    }
    timerId = setTimeout(() => {
      timerId = null;
      fn.apply(this, args);
    }, delay);
  };
}

/**
 * Executa callback com contexto de resumo temporariamente ajustado.
 * @param {Object} context - Referências de seções/gráfico a utilizar.
 * @param {Function} callback - Rotina a executar com o contexto ativo.
 */
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

/**
 * Limpa áreas de resumo (seções e gráfico) antes de nova renderização.
 * @param {Object} context - Referências de DOM do resumo.
 */
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

/**
 * Inicializa carregamento do pacote Google Charts para o gráfico de Gantt.
 * Evita requisições duplicadas utilizando flags locais.
 */
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

/**
 * Agenda atualização do Gantt usando requestAnimationFrame para otimizar repaints.
 */
function queueGanttRefresh() {
  if (ganttRefreshScheduled) return;
  ganttRefreshScheduled = true;
  requestAnimationFrame(() => {
    ganttRefreshScheduled = false;
    refreshGantt();
  });
}

/**
 * Atualiza visualização do gráfico de Gantt com base nas atividades cadastradas.
 * Oculta o container quando não há dados válidos.
 */
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

/**
 * Varre o DOM para construir estrutura mínima de marcos e atividades usada pelo Gantt.
 * @returns {Array<{nome:string, atividades:Array}>} Lista de marcos com atividades.
 */
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

/**
 * Renderiza gráfico de Gantt no container informado.
 * @param {Array} milestones - Estrutura gerada por collectMilestonesForGantt.
 * @param {Object} [options={}] - Permite customizar elementos e mensagens.
 * @returns {Object|null} Informações sobre datas e linhas renderizadas.
 */
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
/**
 * Define o ano de aprovação padrão igual ao ano corrente e protege campo contra valores futuros.
 */
function setApprovalYearToCurrent() {
  if (!approvalYearInput) {
    return;
  }
  const currentYear = new Date().getFullYear();
  approvalYearInput.value = currentYear;
  approvalYearInput.max = currentYear;
}

/**
 * Fluxo principal de inicialização: prepara selects, registra eventos e carrega dados iniciais.
 */
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

/**
 * Centraliza o registro de listeners de interface, garantindo foco e acessibilidade.
 */
function bindEvents() {
  newProjectBtn.addEventListener('click', () => openProjectForm('create'));
  closeFormBtn.addEventListener('click', handleCloseFormRequest);
  if (floatingCloseBtn) {
    floatingCloseBtn.addEventListener('click', handleCloseFormRequest);
  }
  // Habilita o fechamento do formulário pela tecla ESC.
  document.addEventListener('keydown', handleOverlayEscape);
  document.addEventListener('input', handleGlobalDateInput);
  if (projectSearch) {
    projectSearch.addEventListener('input', () => renderProjectList({ defer: true }));
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

  const debouncedBudgetRecalculation = debounce(() => {
    updateInvestmentLevelField();
    updateBudgetSections();
    validatePepBudget();
  }, 180);
  projectBudgetInput.addEventListener('input', debouncedBudgetRecalculation);

  const scheduleActivityDateValidation = debounce((input) => {
    if (input) {
      validateActivityDates({ changedInput: input });
    } else {
      validateActivityDates();
    }
  }, 180);

  const schedulePepBudgetValidation = debounce((input) => {
    if (input) {
      validatePepBudget({ changedInput: input });
    } else {
      validatePepBudget();
    }
  }, 180);

  if (projectStartDateInput) {
    const handleProjectStartChange = (event) => {
      scheduleActivityDateValidation(event.target);
    };
    projectStartDateInput.addEventListener('input', handleProjectStartChange);
    projectStartDateInput.addEventListener('change', handleProjectStartChange);
  }

  if (projectEndDateInput) {
    const handleProjectEndChange = (event) => {
      scheduleActivityDateValidation(event.target);
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
      schedulePepBudgetValidation(event.target);
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
      scheduleActivityDateValidation(event.target);
    }
    if (event.target.classList?.contains('activity-end')) {
      scheduleActivityDateValidation(event.target);
    }
    if (event.target.classList?.contains('activity-pep-amount')) {
      schedulePepBudgetValidation(event.target);
    }
    queueGanttRefresh();
  };

  milestoneList.addEventListener('input', handleMilestoneFormChange);
  milestoneList.addEventListener('change', handleMilestoneFormChange);
}

// ============================================================================
// Carregamento e renderização da lista de projetos
// ============================================================================
/**
 * Recupera projetos do SharePoint filtrando pelo autor logado e atualiza o estado local.
 * @returns {Promise<void>} Promessa resolvida após renderizar lista.
 */
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

/**
 * Renderiza cards de projetos no painel lateral, com filtro opcional e defer.
 * @param {{defer?:boolean}} [options={}] - Quando defer é true, usa requestAnimationFrame.
 */
function renderProjectList(options = {}) {
  const filter = (projectSearch?.value || '').toLowerCase();

  const drawList = () => {
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
  };

  if (options?.defer) {
    if (renderProjectListFrame) {
      cancelAnimationFrame(renderProjectListFrame);
    }
    renderProjectListFrame = requestAnimationFrame(() => {
      renderProjectListFrame = null;
      drawList();
    });
    return;
  }

  if (renderProjectListFrame) {
    cancelAnimationFrame(renderProjectListFrame);
    renderProjectListFrame = null;
  }
  drawList();
}

/**
 * Seleciona projeto na lista, busca detalhes completos e atualiza painel principal.
 * @param {number|string} projectId - Identificador do projeto selecionado.
 * @returns {Promise<void>} Promessa concluída após renderizar detalhes.
 */
async function selectProject(projectId) {
  if (renderProjectListFrame) {
    cancelAnimationFrame(renderProjectListFrame);
    renderProjectListFrame = null;
  }
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

/**
 * Atualiza painel principal com dados ricos do projeto (cards, ações e descrição).
 * @param {{project?: Project}} detail - Objeto completo retornado pela API.
 */
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

  if (project.status === PROJECT_STATUSES.APPROVED) {
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
  const editableStatuses = STATUS_GROUPS.APPROVAL_ALLOWED;
  const viewOnlyStatuses = STATUS_GROUPS.READ_ONLY;

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

/**
 * Produz conteúdo padrão quando nenhum projeto está selecionado.
 * @returns {HTMLDivElement} Elemento com mensagem de orientação.
 */
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

/**
 * Monta card compacto exibindo meta-informações do projeto (orçamento, datas).
 * @param {string} label - Rótulo apresentado no topo do card.
 * @param {string} value - Valor formatado a exibir.
 * @param {{variant?:string}} [options={}] - Permite aplicar estilos específicos.
 * @returns {HTMLDivElement} Elemento configurado.
 */
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

/**
 * Converte string de data em formato local pt-BR ou retorna traço quando inválido.
 * @param {string|null|undefined} value - Data em formato ISO.
 * @returns {string} Data formatada ou '—'.
 */
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

/**
 * Repopula options de um select conforme regras de negócio, preservando valor selecionado.
 * @param {HTMLSelectElement|null} selectElement - Select alvo.
 * @param {string[]} [options=[]] - Valores disponíveis.
 * @param {string} [selectedValue=''] - Valor que deve permanecer selecionado.
 */
function populateSelectOptions(selectElement, options = [], selectedValue = '') {
  if (!selectElement || selectElement.tagName !== 'SELECT') {
    return;
  }

  const initialPlaceholderOption = selectElement.options?.[0];
  const placeholderText = initialPlaceholderOption && initialPlaceholderOption.value === ''
    ? initialPlaceholderOption.textContent || 'Selecione...'
    : 'Selecione...';

  const normalizedOptions = Array.isArray(options)
    ? options.map((option) => String(option))
    : [];
  const normalizedSelectedValue = selectedValue != null ? String(selectedValue) : '';

  selectElement.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = placeholderText;
  selectElement.appendChild(placeholderOption);

  const fragment = document.createDocumentFragment();
  normalizedOptions.forEach((text) => {
    const option = document.createElement('option');
    option.value = text;
    option.textContent = text;
    fragment.appendChild(option);
  });

  selectElement.appendChild(fragment);

  if (normalizedSelectedValue && normalizedOptions.includes(normalizedSelectedValue)) {
    selectElement.value = normalizedSelectedValue;
  } else {
    selectElement.value = '';
  }
}

function fillOptions(selectEl, items = [], { placeholder = 'Selecione...' } = {}) {
  if (!selectEl) return;
  const prev = selectEl.value;
  const normalizedItems = Array.isArray(items) ? items.map((item) => String(item)) : [];
  selectEl.innerHTML = '';
  if (placeholder) {
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = placeholder;
    selectEl.appendChild(opt0);
  }
  const frag = document.createDocumentFragment();
  for (const v of normalizedItems) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    frag.appendChild(opt);
  }
  selectEl.appendChild(frag);
  if (normalizedItems.includes(prev)) selectEl.value = prev; // preserva seleção válida
}

function getUnits(companyCode, centerCode) {
  if (!companyCode || !centerCode) return EMPTY_ARRAY;
  const rules = resolveCompanyRules();
  const companyRule = rules?.[companyCode];
  const centerRule = companyRule?.centers?.[centerCode];
  return safeArray(centerRule?.units);
}

function getLocations(companyCode, centerCode) {
  if (!companyCode || !centerCode) return EMPTY_ARRAY;
  const rules = resolveCompanyRules();
  const companyRule = rules?.[companyCode];
  const centerRule = companyRule?.centers?.[centerCode];
  return safeArray(centerRule?.locations);
}

function applyCompany(companyCode, { centerEl, locationEl, unitEl }) {
  const rules = resolveCompanyRules();
  const centersMap = rules?.[companyCode]?.centers || {};
  const centers = Object.keys(centersMap);
  fillOptions(centerEl, centers, { placeholder: 'Selecione o centro' });
  fillOptions(locationEl, EMPTY_ARRAY, { placeholder: 'Selecione a localização' });
  fillOptions(unitEl, EMPTY_ARRAY, { placeholder: 'Selecione a unidade' });
}

function applyCenter(companyCode, centerCode, { locationEl, unitEl }) {
  const locs  = getLocations(companyCode, centerCode);
  const units = getUnits(companyCode, centerCode);
  fillOptions(locationEl, locs,  { placeholder: 'Selecione a localização' });
  fillOptions(unitEl,      units,{ placeholder: 'Selecione a unidade' });
}

(function wireCompanyCenterLocationUnit() {
  function ready() {
    if (document.readyState === 'loading') return false;
    const rules = resolveCompanyRules();
    return Object.keys(rules || {}).length > 0;
  }

  function run() {
    const companyEl  = document.getElementById('company');
    const centerEl   = document.getElementById('center');
    const locationEl = document.getElementById('location');
    const unitEl     = document.getElementById('unit');
    if (!companyEl || !centerEl || !locationEl || !unitEl) {
      console.warn('[cascata] Ajuste os IDs: company, center, location, unit.');
      return;
    }

    const els = { companyEl, centerEl, locationEl, unitEl };

    const handleCompanyChange = (e) => {
      const companyValue = e.target.value;
      if (typeof updateCompanyDependentFields === 'function') {
        const depValue = (document.getElementById('depreciationCostCenter')?.value || '').trim();
        updateCompanyDependentFields(companyValue, {
          center: '',
          location: '',
          unit: '',
          depreciationCostCenter: depValue
        });
        return;
      }
      applyCompany(companyValue, els);
    };

    const handleCenterChange = (e) => {
      const companyValue = companyEl.value || '';
      const centerValue = e.target.value;
      if (typeof updateCompanyDependentFields === 'function') {
        const depValue = (document.getElementById('depreciationCostCenter')?.value || '').trim();
        updateCompanyDependentFields(companyValue, {
          center: centerValue,
          location: '',
          unit: '',
          depreciationCostCenter: depValue
        });
        return;
      }
      applyCenter(companyValue, centerValue, els);
    };

    companyEl.addEventListener('change', handleCompanyChange);
    centerEl.addEventListener('change', handleCenterChange);

    const initCompany = companyEl.value || '';
    if (initCompany) {
      if (typeof updateCompanyDependentFields === 'function') {
        const depValue = (document.getElementById('depreciationCostCenter')?.value || '').trim();
        updateCompanyDependentFields(initCompany, {
          center: centerEl.value || '',
          location: locationEl.value || '',
          unit: unitEl.value || '',
          depreciationCostCenter: depValue
        });
      } else {
        applyCompany(initCompany, els);
        const initCenter = centerEl.value || '';
        if (initCenter) applyCenter(initCompany, initCenter, els);
      }
    }
  }

  if (!ready()) {
    const t0 = Date.now();
    const iv = setInterval(() => {
      if (ready()) { clearInterval(iv); run(); }
      else if (Date.now() - t0 > 2000) { clearInterval(iv); console.warn('[cascata] Mapas/DOM não prontos.'); }
    }, 50);
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => { /* no-op */ }, { once: true });
    }
  } else {
    run();
  }
})();

/**
 * Atualiza selects dependentes (centro, unidade etc.) conforme empresa escolhida.
 * @param {string} companyValue - Empresa selecionada.
 * @param {{center?:string, unit?:string, location?:string, depreciation?:string, depreciationCostCenter?:string}} [selectedValues={}] - Valores previamente gravados.
 */
function updateCompanyDependentFields(companyValue, selectedValues) {
  const companySelect = document.getElementById('company');
  const centerSelect = document.getElementById('center');
  const locationSelect = document.getElementById('location');
  const unitSelect = document.getElementById('unit');
  const depreciationFieldElement = document.getElementById('depreciationCostCenter');

  if (!centerSelect && !locationSelect && !unitSelect && !depreciationFieldElement) {
    return;
  }

  const hasSelectedValues = arguments.length > 1 && selectedValues && typeof selectedValues === 'object';
  const preserve = hasSelectedValues
    ? {
        center: selectedValues.center != null ? String(selectedValues.center) : '',
        location: selectedValues.location != null ? String(selectedValues.location) : '',
        unit: selectedValues.unit != null ? String(selectedValues.unit) : '',
        depreciation: (
          selectedValues.depreciationCostCenter ?? selectedValues.depreciation ?? ''
        )?.toString().trim()
      }
    : {
        center: centerSelect?.value || '',
        location: locationSelect?.value || '',
        unit: unitSelect?.value || '',
        depreciation: (depreciationFieldElement?.value ?? '').trim()
      };

  const allRules = typeof resolveCompanyRules === 'function' ? resolveCompanyRules() : {};
  const companyKey = companyValue != null ? String(companyValue) : (companySelect?.value || '');
  const rules = allRules?.[companyKey] || {};

  const centersSource = rules.centers;
  const centers = Array.isArray(centersSource)
    ? centersSource.map((center) => String(center))
    : centersSource && typeof centersSource === 'object'
      ? Object.keys(centersSource)
      : [];

  populateSelectOptions(centerSelect, centers, preserve.center);

  const activeCenter = centerSelect?.value || '';
  let locations = [];
  let units = [];

  if (Array.isArray(rules.locations)) {
    locations = rules.locations.map((location) => String(location));
  } else if (centersSource && typeof centersSource === 'object') {
    const centerRule = centersSource[activeCenter];
    locations = Array.isArray(centerRule?.locations)
      ? centerRule.locations.map((location) => String(location))
      : [];
    units = Array.isArray(centerRule?.units)
      ? centerRule.units.map((unit) => String(unit))
      : [];
  }

  if (Array.isArray(rules.units) && units.length === 0) {
    units = rules.units.map((unit) => String(unit));
  }

  populateSelectOptions(locationSelect, locations, preserve.location);
  populateSelectOptions(unitSelect, units, preserve.unit);

  // --- Depreciação: suporta SELECT ou INPUT ---
  const depreciationField = document.getElementById('depreciationCostCenter');
  const preservedDep = (
    (selectedValues && (selectedValues.depreciationCostCenter ?? selectedValues.depreciation)) ??
    depreciationField?.value ??
    ''
  ).trim();

  // Se houver regras de depreciação por empresa e o campo for SELECT, ainda permita popular.
  const depreciationList = Array.isArray(rules?.depreciation)
    ? rules.depreciation.map(String)
    : [];

  // Caso SELECT: popula como antes + opção legada se preciso
  if (depreciationField && depreciationField.tagName === 'SELECT') {
    populateSelectOptions(depreciationField, depreciationList, preservedDep);

    if (preservedDep && !depreciationList.includes(preservedDep)) {
      const legacyOption = document.createElement('option');
      legacyOption.value = preservedDep;
      legacyOption.textContent = preservedDep + ' (existente)';
      legacyOption.dataset.legacy = 'true';
      depreciationField.appendChild(legacyOption);
      depreciationField.value = preservedDep;
    }
  } else if (depreciationField) {
    // Caso INPUT: apenas preservar/escrever
    if (preservedDep) depreciationField.value = preservedDep;
    if (!depreciationField.placeholder) depreciationField.placeholder = 'Ex.: CC-01';
  }
}

function isValidSelection({ company, center, location, unit }) {
  const rules = resolveCompanyRules();
  const companyRule = rules?.[company];
  if (!companyRule) return false;
  const centerRule = companyRule.centers?.[center];
  if (!centerRule) return false;

  const locations = safeArray(centerRule.locations);
  const units = safeArray(centerRule.units);

  const locationOk = locations.length === 0 ? !location : locations.includes(location);
  if (!locationOk) return false;

  const unitOk = units.length === 0 ? !unit : units.includes(unit);
  return unitOk;
}

const form = document.querySelector('form#projectForm');
if (form) {
  form.addEventListener('submit', (ev) => {
    const sels = {
      company:  document.getElementById('company')?.value || '',
      center:   document.getElementById('center')?.value || '',
      location: document.getElementById('location')?.value || '',
      unit:     document.getElementById('unit')?.value || '',
    };
    if (!isValidSelection(sels)) {
      ev.preventDefault();
      alert('Seleção inválida: verifique Empresa, Centro, Localização e Unidade.');
    }
  });
}

// ============================================================================
// Formulário: abertura, preenchimento e coleta dos dados
/**
 * Alterna habilitação dos controles do formulário respeitando estados originais.
 * @param {boolean} disabled - Indica se campos devem ser bloqueados.
 */
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

/**
 * Aplica modo de formulário (edição ou leitura) ajustando visibilidade de seções.
 * @param {'edit'|'readonly'} mode - Modo desejado.
 * @param {{showApprovalButton?:boolean, refreshSummary?:boolean}} [options={}] - Ajustes complementares.
 */
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

/**
 * Ajusta fluxo de exibição com base no status do projeto (ex.: somente leitura quando aprovado).
 * @param {ProjectStatus|string} status - Status em avaliação.
 */
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
/**
 * Prepara overlay para criação/edição populando campos e aplicando regras iniciais.
 * @param {'create'|'edit'} mode - Contexto do formulário.
 * @param {{project?: Project}} [detail=null] - Dados carregados quando em edição.
 */
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
    statusField.value = PROJECT_STATUSES.DRAFT;
    setApprovalYearToCurrent();
    updateBudgetSections({ clear: true });
  } else if (detail) {
    fillFormWithProject(detail);
  }

  const statusKey = detail?.project?.status || statusField.value || PROJECT_STATUSES.DRAFT;
  applyStatusBehavior(statusKey);

  updateSimplePepYears();
  overlay.classList.remove('hidden');
  queueGanttRefresh();
  validateAllDateRanges();
}

/**
 * Preenche o formulário com dados existentes de um projeto selecionado para edição.
 * @param {{project: Project, simplePeps:Array, milestones:Array, activities:Array, activityPeps:Array}} detail - Pacote de dados relacionado ao projeto.
 */
function fillFormWithProject(detail) {
  const { project, simplePeps, milestones, activities, activityPeps } = detail;
  formTitle.textContent = `Editar Projeto #${project.Id}`;
  statusField.value = project.status || PROJECT_STATUSES.DRAFT;

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
    depreciationCostCenter: project.depreciationCostCenter || ''
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
        const actId = Number(activity.Id);
        const relatedPeps = activityPeps.filter(
          (pep) => Number(pep.activitiesIdId ?? pep.activitiesId) === actId
        );
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
    const isEditingMode = projectForm?.dataset.mode === 'edit';
    if (!milestones.length && !isEditingMode) {
      ensureMilestoneBlock();
    }
  }

  queueGanttRefresh();
  validatePepBudget();
  validateActivityDates();
  validateAllDateRanges();
}

/**
 * Fecha overlay de formulário e garante fechamento do resumo embutido.
 */
function closeForm() {
  overlay.classList.add('hidden');
  closeSummaryOverlay({ restoreFocus: false });
}

/**
 * Exibe overlay de resumo para revisão final, após validar campos obrigatórios.
 * @param {HTMLButtonElement|null} [triggerButton=null] - Botão que acionou o resumo para restaurar foco.
 */
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
  if (summaryTitle) {
    summaryTitle.focus();
  } else if (summaryConfirmBtn) {
    summaryConfirmBtn.focus();
  }
}

/**
 * Oculta overlay de resumo e limpa conteúdo, restaurando foco se necessário.
 * @param {{restoreFocus?:boolean}} [options={}] - Controla retorno do foco ao botão acionador.
 */
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

/**
 * Confirmador do overlay: fecha modal e dispara submit programático.
 */
function handleSummaryConfirm() {
  closeSummaryOverlay({ restoreFocus: false });
  projectForm.requestSubmit();
}

/**
 * Atualiza overlay principal de resumo preenchendo seções e Gantt.
 */
function populateSummaryOverlay() {
  if (!summarySections) return;
  populateSummaryContent({ context: defaultSummaryContext, refreshGantt: true });
}

/**
 * Agrupa dados do formulário em estrutura declarativa para renderização no resumo.
 * @returns {Array<{title:string, entries:Array}>} Conjunto de seções e campos.
 */
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

/**
 * Renderiza seções e gráficos do resumo conforme contexto ativo (overlay ou formulário).
 * @param {{context?:Object, refreshGantt?:boolean}} [options={}] - Define destino e necessidade de atualizar o Gantt.
 */
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

/**
 * Cria uma seção de resumo com base nos pares label/valor fornecidos.
 * @param {string} title - Título da seção.
 * @param {Array<{label:string, value:*, fullWidth?:boolean}>} entries - Campos exibidos na seção.
 */
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

/**
 * Consolida dados de PEPs simples ou vinculados a atividades para exibição no resumo.
 */
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

/**
 * Exibe marcos e atividades no resumo final quando Key Projects está habilitado.
 */
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

/**
 * Alimenta gráfico Gantt no resumo, respeitando disponibilidade do Google Charts.
 * @param {{refreshFirst?:boolean}} [options={}] - Quando true, força refresh antes da captura.
 */
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

/**
 * Aguarda renderização completa do Gantt para garantir captura ou navegação suave.
 * @param {Object} context - Contexto ativo do resumo.
 * @param {{timeout?:number}} [options={}] - Tempo máximo de espera em milissegundos.
 * @returns {Promise<void>} Resolve após gráfico disparar evento 'ready' ou timeout.
 */
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

/**
 * Monta string amigável representando período de uma atividade.
 * @param {Element} activity - Elemento DOM da atividade.
 * @returns {string} Intervalo formatado ou vazio quando não informado.
 */
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

/**
 * Recupera valor textual exibível de campos do formulário, respeitando selects.
 * @param {string} fieldId - ID do elemento.
 * @returns {string} Texto puro ou vazio.
 */
function getFieldDisplayValue(fieldId) {
  const field = document.getElementById(fieldId);
  if (!field) return '';
  if (field.tagName === 'SELECT') {
    return getSelectOptionText(field);
  }
  return field.value ?? '';
}

/**
 * Extrai texto da option selecionada ou retorna valor direto do select.
 * @param {HTMLSelectElement|null} selectElement - Elemento alvo.
 * @returns {string} Texto amigável.
 */
function getSelectOptionText(selectElement) {
  if (!selectElement) return '';
  const option = selectElement.options?.[selectElement.selectedIndex];
  if (option) {
    return option.textContent?.trim() ?? '';
  }
  return selectElement.value ?? '';
}

/**
 * Converte valor numérico de um campo para moeda brasileira.
 * @param {string} fieldId - ID do input.
 * @returns {string} Valor formatado ou vazio.
 */
/**
 * Converte valor numérico de input para moeda brasileira.
 * @param {string} fieldId - ID do campo.
 * @returns {string} Valor formatado ou string vazia.
 */
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

/**
 * Converte campo numérico para string localizada, mantendo fallback em caso de parsing inválido.
 * @param {string} fieldId - ID do input.
 * @returns {string} Valor formatado ou original.
 */
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

/**
 * Lê valor numérico de um input e devolve formato monetário pt-BR.
 * @param {HTMLInputElement|null} element - Campo de origem.
 * @returns {string} Valor formatado ou string vazia.
 */
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

/**
 * Normaliza valores exibidos no resumo, retornando '—' quando não informado.
 * @param {*} value - Valor bruto.
 * @returns {string} Texto pronto para exibição.
 */
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
 * Solicita confirmação antes de fechar o formulário e garante fechamento do resumo ativo.
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

/**
 * Intercepta tecla ESC para iniciar fluxo de fechamento controlado do formulário.
 * @param {KeyboardEvent} event - Evento de teclado.
 */
function handleOverlayEscape(event) {
  if (event.key !== 'Escape') return;
  handleCloseFormRequest();
}

/**
 * Alterna visibilidade entre seção de PEP simples e Key Projects com base no orçamento.
 * @param {{preserve?:boolean, clear?:boolean}} [options={}] - Controla limpeza de listas dinâmicas.
 */
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
    if (!milestoneList.children.length && !preserve) {
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

/**
 * Habilita ou desabilita interações dentro de uma seção específica.
 * @param {HTMLElement|null} section - Fieldset alvo.
 * @param {boolean} enabled - Indica se inputs permanecem ativos.
 */
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

/**
 * Executa validações nativas HTML5 e retorna lista estruturada de problemas.
 * @returns {Array<{element:HTMLElement, label:string, message:string, type:string}>} Erros detectados.
 */
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

/**
 * Reestabelece o estado das validações customizadas e limpa mensagens auxiliares.
 */
function resetValidationState() {
  validationState.pepBudget = null;
  validationState.pepBudgetDetails = null;
  validationState.activityDates = null;
  validationState.activityDateDetails = null;
  clearDateRangeValidity();
  clearFieldErrors();
  clearErrorSummary();
}

/**
 * Atualiza caches de mensagens e detalhes usados pelos validadores personalizados.
 * @param {string} key - Nome da propriedade em validationState.
 * @param {string|null} message - Mensagem registrada.
 * @param {*} [details=null] - Metadados adicionais.
 */
function setValidationError(key, message, details = null) {
  validationState[key] = message || null;
  const detailsKey = `${key}Details`;
  if (detailsKey in validationState) {
    validationState[detailsKey] = details || null;
  }
}

/**
 * Memoriza valor anterior de um campo para restauração em validações corretivas.
 * @param {HTMLElement|null} element - Campo monitorado.
 */
function rememberFieldPreviousValue(element) {
  if (!element || typeof element !== 'object') return;
  if (!('dataset' in element)) return;
  element.dataset.previousValue = element.value ?? '';
}

/**
 * Normaliza strings numéricas aceitando formatos com vírgula/ponto.
 * @param {*} value - Valor bruto informado pelo usuário.
 * @returns {string} Representação normalizada.
 */
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

/**
 * Converte valores variados em número de ponto flutuante, respeitando normalização.
 * @param {*} value - Valor a converter.
 * @returns {number} Número coerente ou NaN.
 */
function coerceNumericValue(value) {
  const normalized = normalizeNumericString(value);
  if (!normalized) {
    return NaN;
  }
  const number = Number.parseFloat(normalized);
  return Number.isFinite(number) ? number : NaN;
}

/**
 * Gera string numérica pronta para setar em inputs tipo number/text.
 * @param {*} value - Valor de origem.
 * @returns {string} Valor sanitizado ou vazio.
 */
function sanitizeNumericInputValue(value) {
  const number = coerceNumericValue(value);
  return Number.isFinite(number) ? number.toString() : '';
}

/**
 * Interpreta valor numérico a partir de input ou string.
 * @param {HTMLInputElement|number|string} source - Fonte do valor.
 * @returns {number} Número válido ou 0.
 */
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

/**
 * Calcula nível de investimento com base no orçamento convertido para USD.
 * @param {number} budgetBrl - Orçamento em reais.
 * @returns {string} Código N1..N4 ou vazio.
 */
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

/**
 * Recupera valor numérico do orçamento informado no formulário.
 * @returns {number} Valor em reais ou NaN se inválido.
 */
function getProjectBudgetValue() {
  if (!projectBudgetInput) return NaN;
  const rawValue = projectBudgetInput.value;
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return NaN;
  }
  return parseNumericInputValue(projectBudgetInput);
}

/**
 * Atualiza select de nível de investimento conforme valor do orçamento.
 * @param {number} [budgetBrl=getProjectBudgetValue()] - Valor base para cálculo.
 */
function updateInvestmentLevelField(budgetBrl = getProjectBudgetValue()) {
  if (!investmentLevelSelect) return;
  const level = determineInvestmentLevel(budgetBrl);
  investmentLevelSelect.value = level;
}

/**
 * Retorna lista de inputs que compõem o somatório de PEP.
 * @returns {HTMLInputElement[]} Inputs de valores PEP.
 */
function getPepAmountInputs() {
  const simplePepInputs = Array.from(simplePepList.querySelectorAll('.pep-amount'));
  const activityPepInputs = Array.from(milestoneList.querySelectorAll('.activity-pep-amount'));
  return [...simplePepInputs, ...activityPepInputs];
}

/**
 * Soma todos os valores de PEP considerando listas simples e atividades.
 * @returns {number} Total em reais.
 */
function calculatePepTotal() {
  return getPepAmountInputs().reduce((sum, input) => sum + parseNumericInputValue(input), 0);
}

/**
 * Atualiza mensagem informativa abaixo do formulário com saldo/ excedente de orçamento.
 * @param {{budget?:number,total?:number}} [param0={}] - Valores utilizados no cálculo.
 */
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

/**
 * Valida se somatório de PEP está dentro do orçamento informado.
 * @param {{changedInput?:HTMLInputElement}} [options={}] - Campo modificado recentemente.
 * @returns {boolean} True quando orçamento e PEP estão consistentes.
 */
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

/**
 * Executa validação em todos os pares de datas, reportando o primeiro inválido.
 * @param {{report?:boolean}} [options={}] - Quando true, chama reportValidity no primeiro campo inválido.
 * @returns {boolean} True se nenhum intervalo estiver inconsistente.
 */
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

/**
 * Limpa customValidity aplicado em todos os campos de data do formulário.
 */
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

/**
 * Atualiza mensagem informativa relacionada às datas do projeto e atividades.
 * @param {{hasProjectStart?:boolean,hasProjectEnd?:boolean,hasStartIssue?:boolean,hasEndIssue?:boolean,activityCount?:number}} [param0={}] - Indicadores para montagem do texto.
 */
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

/**
 * Coordena validações customizadas e nativas retornando resumo das inconsistências.
 * @param {{scrollOnError?:boolean, focusFirstError?:boolean}} [options={}] - Comportamento pós-validação.
 * @returns {{valid:boolean, issues:Array}} Resultado consolidado.
 */
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

/**
 * Escuta eventos de input em campos de data para aplicar validações contextuais.
 * @param {Event} event - Evento disparado no formulário.
 */
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

/**
 * Armazena valor atual de campos sensíveis antes da edição para suportar rollback.
 * @param {FocusEvent} event - Evento focusin delegado do formulário.
 */
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

/**
 * Ajusta ano do PEP de uma atividade considerando data de início ou ano de aprovação.
 * @param {HTMLElement|null} activityElement - Container da atividade.
 * @param {{fallbackYear?:number|null, force?:boolean}} [options={}] - Estratégia de atualização.
 */
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

/**
 * Propaga ano padrão para PEPs simples e atividades quando o campo principal muda.
 */
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

/**
 * Garante ao menos uma linha de PEP simples pronta para preenchimento.
 */
function ensureSimplePepRow() {
  const row = createSimplePepRow({ year: parseInt(approvalYearInput.value, 10) || '' });
  simplePepList.append(row);
}

/**
 * Adiciona bloco de marco padrão e agenda atualização do Gantt.
 */
function ensureMilestoneBlock() {
  const block = createMilestoneBlock();
  milestoneList.append(block);
  addActivityBlock(block);
  queueGanttRefresh();
}

/**
 * Normaliza valores de ID para uso consistente em data attributes.
 * @param {*} value - Valor original informado.
 * @returns {string} Representação textual segura ou string vazia.
 */
function resolveDatasetId(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const stringValue = String(value);
  return stringValue ? stringValue : '';
}

/**
 * Cria linha PEP a partir do template aplicando dados já salvos, quando houver.
 * @param {{id?:string|number,title?:string,amount?:number|string,year?:number|string}} [param0={}] - Dados iniciais.
 * @returns {HTMLElement} Linha gerada.
 */
function createSimplePepRow({ id = '', title = '', amount = '', year = '' } = {}) {
  const fragment = simplePepTemplate.content.cloneNode(true);
  const row = fragment.querySelector('.pep-row');
  row.dataset.pepId = resolveDatasetId(id);
  row.querySelector('.pep-title').value = title || '';
  row.querySelector('.pep-amount').value = sanitizeNumericInputValue(amount);
  row.querySelector('.pep-year').value = year ?? '';
  return row;
}

/**
 * Cria bloco de marco a partir do template e aplica identificadores existentes.
 * @param {{id?:string|number,title?:string}} [param0={}] - Dados iniciais.
 * @returns {HTMLElement} Bloco de marco.
 */
function createMilestoneBlock({ id = '', title = '' } = {}) {
  const fragment = milestoneTemplate.content.cloneNode(true);
  const block = fragment.querySelector('.milestone');
  block.dataset.milestoneId = resolveDatasetId(id);
  block.querySelector('.milestone-title').value = title || '';
  return block;
}

/**
 * Cria atividade dentro de um marco, preenchendo campos quando dados são fornecidos.
 * @param {HTMLElement|null} milestoneElement - Container do marco.
 * @param {Object} [data={}] - Valores opcionais (id, datas, PEP etc.).
 * @returns {HTMLElement|null} Atividade recém criada.
 */
function addActivityBlock(milestoneElement, data = {}) {
  if (!milestoneElement) return null;
  const fragment = activityTemplate.content.cloneNode(true);
  const activity = fragment.querySelector('.activity');
  const startInput = activity.querySelector('.activity-start');
  const endInput = activity.querySelector('.activity-end');
  const amountInput = activity.querySelector('.activity-pep-amount');
  const pepTitleInput = activity.querySelector('.activity-pep-title');
  const pepYearInput = activity.querySelector('.activity-pep-year');

  const resolvedMilestoneId = resolveDatasetId(
    data.milestoneId ?? milestoneElement.dataset?.milestoneId ?? ''
  );
  activity.dataset.activityId = resolveDatasetId(data.id);
  activity.dataset.pepId = resolveDatasetId(data.pepId);
  activity.dataset.milestoneId = resolvedMilestoneId;

  activity.querySelector('.activity-title').value = data.title || '';
  if (amountInput) {
    amountInput.type = 'text';
    amountInput.setAttribute('inputmode', 'decimal');
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
/**
 * Trata submissão do formulário executando persistência no SharePoint e anexos JSON.
 * @param {Event} event - Evento submit interceptado.
 */
async function handleFormSubmit(event) {
  event.preventDefault();
  const mode = projectForm.dataset.mode;
  const projectId = projectForm.dataset.projectId;
  const submitIntent = projectForm.dataset.submitIntent || 'save';
  const isApproval = submitIntent === 'approval';

  document.querySelectorAll('.activity-pep-amount').forEach((inp) => {
    inp.value = sanitizeNumericInputValue(inp.value);
  });

  const validation = runFormValidations({ scrollOnError: true, focusFirstError: true });
  if (!validation.valid) {
    return;
  }

  const normalizedStatus = isApproval
    ? PROJECT_STATUSES.IN_APPROVAL
    : PROJECT_STATUSES.DRAFT;
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

      const attachmentOptions = {
        contentType: 'application/json',
        ...(mode !== 'create' ? { overwrite: true } : {})
      };

      await sp.addAttachment('Projects', resolvedId, 'resumo.txt', jsonBlob, attachmentOptions);

      await sp.updateItem('Projects', resolvedId, { status: PROJECT_STATUSES.IN_APPROVAL });
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
        await sp.updateItem('Projects', resolvedId, { status: PROJECT_STATUSES.DRAFT });
        updateProjectState(resolvedId, { status: PROJECT_STATUSES.DRAFT });
        renderProjectList();
        if (state.currentDetails?.project?.Id === resolvedId) {
          state.currentDetails = {
            ...state.currentDetails,
            project: {
              ...state.currentDetails.project,
              status: PROJECT_STATUSES.DRAFT
            }
          };
          renderProjectDetails(state.currentDetails);
        }
      } catch (rollbackError) {
        console.error('Erro ao reverter status após falha no envio', rollbackError);
      }
      statusField.value = PROJECT_STATUSES.DRAFT;
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

/**
 * Extrai dados do formulário para montar payload a ser enviado à lista Projects.
 * @returns {Project} Objeto com campos normalizados.
 */
function collectProjectData() {
  const budgetValue = getProjectBudgetValue();
  const budgetBrl = Number.isFinite(budgetValue) ? budgetValue : 0;
  const investmentLevelValue = determineInvestmentLevel(budgetValue);
  const depField = document.getElementById('depreciationCostCenter');
  const depreciationValue = depField?.value || '';

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
    depreciationCostCenter: depreciationValue,
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

/**
 * Decide entre persistir PEPs simples ou estrutura Key Projects de acordo com o orçamento.
 * @param {number} projectId - ID do projeto salvo.
 * @param {Project} projectData - Dados coletados do formulário.
 */
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

/**
 * Extrai PEPs simples do formulário para uso em resumos e persistência.
 * @returns {Pep[]} Lista de PEPs simples.
 */
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

/**
 * Consolida dados de marcos e atividades presentes no formulário para resumos e anexos.
 * @returns {Milestone[]} Estrutura normalizada.
 */
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

/**
 * Lineariza atividades associadas aos marcos para facilitar exportação.
 * @param {Milestone[]} milestones - Marcos coletados previamente.
 * @returns {Activity[]} Lista de atividades enriquecida com referência ao marco.
 */
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

/**
 * Agrupa PEPs simples e vinculados às atividades para compor anexos e resumos.
 * @param {Pep[]} simplePeps - PEPs independentes coletados.
 * @param {Activity[]} activities - Atividades com eventual PEP.
 * @returns {Pep[]} Conjunto consolidado.
 */
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

/**
 * Reúne rótulos amigáveis de selects para compor anexos e resumos.
 * @returns {Object} Mapa de valores exibidos ao usuário.
 */
function collectProjectDisplayValues() {
  const depField = document.getElementById('depreciationCostCenter');
  const depValue = depField?.value || '';
  return {
    investmentLevel: getSelectOptionText(investmentLevelSelect),
    company: getSelectOptionText(companySelect),
    center: getSelectOptionText(centerSelect),
    unit: getSelectOptionText(unitSelect),
    location: getSelectOptionText(locationSelect),
    depreciationCostCenter: depValue,
    category: getSelectOptionText(document.getElementById('category')),
    investmentType: getSelectOptionText(document.getElementById('investmentType')),
    assetType: getSelectOptionText(document.getElementById('assetType')),
    kpiType: getSelectOptionText(document.getElementById('kpiType'))
  };
}

/**
 * Estrutura payload resumido do projeto utilizado em anexos JSON.
 * @param {number|string} projectId - Identificador do projeto.
 * @param {Project} [projectData={}] - Dados coletados do formulário.
 * @returns {SummaryPayload} Resumo consolidado.
 */
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

/**
 * Sincroniza PEPs simples com a lista SharePoint, criando, atualizando e removendo conforme necessário.
 * @param {number} projectId - ID do projeto pai.
 * @param {number} approvalYear - Ano de aprovação usado como fallback.
 */
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
      row.dataset.pepId = resolveDatasetId(id);
      currentIds.add(Number(id));
    } else {
      const created = await sp.createItem('Peps', payload);
      const createdId = created?.Id;
      row.dataset.pepId = resolveDatasetId(createdId);
      if (Number.isFinite(Number(createdId))) {
        currentIds.add(Number(createdId));
      }
    }
  }

  const toDelete = [...state.editingSnapshot.simplePeps].filter((id) => !currentIds.has(id));
  for (const id of toDelete) {
    await sp.deleteItem('Peps', Number(id));
  }
}

/**
 * Remove PEPs previamente associados quando projeto migra para modo Key Projects.
 */
async function cleanupSimplePeps() {
  for (const id of state.editingSnapshot.simplePeps) {
    await sp.deleteItem('Peps', Number(id));
  }
  state.editingSnapshot.simplePeps.clear();
}

/**
 * Sincroniza marcos, atividades e PEPs vinculados com suas respectivas listas SharePoint.
 * @param {number} projectId - ID do projeto pai.
 */
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
    }
    milestone.dataset.milestoneId = resolveDatasetId(milestoneId);
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
      }
      activity.dataset.activityId = resolveDatasetId(activityId);
      activity.dataset.milestoneId = resolveDatasetId(milestoneId);
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
        }
        activity.dataset.pepId = resolveDatasetId(pepId);
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

/**
 * Remove registros de Key Projects associados ao projeto quando orçamento cai abaixo do limiar.
 */
async function cleanupKeyProjects() {
  await deleteMissing('Peps', state.editingSnapshot.activityPeps, new Set());
  await deleteMissing('Activities', state.editingSnapshot.activities, new Set());
  await deleteMissing('Milestones', state.editingSnapshot.milestones, new Set());
}

/**
 * Remove itens SharePoint que estavam relacionados anteriormente mas não existem mais localmente.
 * @param {string} listName - Nome da lista alvo.
 * @param {Set<number>} previousSet - IDs conhecidos antes da edição.
 * @param {Set<number>} currentSet - IDs que permanecem após edição.
 */
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
/**
 * Faz scroll para o topo do overlay e do formulário, garantindo visibilidade das mensagens.
 */
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

/**
 * Atualiza componente de status com mensagem contextual e tom visual.
 * @param {string} message - Texto apresentado.
 * @param {Object|string|boolean} [options={}] - Configuração de tom.
 */
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

/**
 * Converte string em inteiro seguro ou retorna null.
 * @param {*} value - Valor de origem.
 * @returns {number|null} Inteiro válido.
 */
function parseNumber(value) {
  const number = parseInt(value, 10);
  return Number.isFinite(number) ? number : null;
}

// ============================================================================
// PEP title dropdown utilities and wiring
// ============================================================================

const PEP_OPTIONS_DEFAULT = [
  'DESP.ENGENHARIA / DETALHAMENTO PROJETO',
  'AQUISIÇÃO DE EQUIPAMENTOS NACIONAIS',
  'AQUISIÇÃO DE EQUIPAMENTOS IMPORTADOS',
  'AQUISIÇÃO DE VEÍCULOS',
  'DESPESAS COM OBRAS CIVIS',
  'DESP.MONTAGEM EQUIPTOS/ESTRUTURAS/OUTRAS',
  'AQ.DE COMPONENTES/MAT.INSTAL./FERRAMENTA',
  'DESPESAS COM MEIO AMBIENTE',
  'DESPESAS COM SEGURANÇA',
  'DESPESAS COM SEGUROS',
  'DESP.CONSULTORIA INTERNA (AMS)-TEC.INFOR',
  'DESP.CONSULTORIA EXTERNA - TEC.INFOR',
  'AQUISIÇÃO DE HARDWARE (NOTEBOOKS, ETC)',
  'AQUISIÇÃO DE SOFTWARE',
  'AQUISIÇÃO DE IMÓVEIS',
  'DESP.GERENCIAMENTO E COORDENAÇÃO',
  'CONTINGÊNCIAS'
];

const PEP_OPTIONS_BF00 = [
  'CONSTRUÇÃO/REFORMA DE FORNOS',
  'CONSTRUÇÃO/REFORMA DE QUEIMADORES',
  'INSTALAÇÕES INDUSTRIAIS',
  'INSTALAÇÕES PREDIAIS',
  'COMPUTADORES E PERIFÉRICOS',
  'SOFTWARES',
  'CERTIFICAÇÕES E LICENÇAS',
  "INFRAESTRUTURA UPC'S",
  'CONTRUÇÃO DE FORNOS',
  'CONSTRUÇÃO QUEIMADORES',
  'MÓDULOS MOVIMENTAÇÃO',
  'MÓDULOS MECANIZAÇÃO',
  'CONSTRUÇÃO DE VIVEIROS',
  'MELHORIAS INDUSTRIAIS',
  'MELHORIAS AMBIENTAIS',
  'TECNOLOGIA DA INFORMAÇÃO',
  'MÁQUINAS E EQUIPAMENTOS',
  'FERRAMENTAS',
  'IMPLEMENTOS AGRÍCOLAS',
  'MÓVEIS E UTENSÍLIOS',
  'VEÍCULOS LEVES',
  'VEÍCULOS PESADOS'
];

/**
 * Remove duplicados preservando ordem e ignorando valores vazios.
 * @param {unknown[]} arr - Coleção de valores para normalizar.
 * @returns {string[]} Lista sanitizada de strings únicas.
 */
function uniquePreserveOrder(arr) {
  if (!Array.isArray(arr)) {
    return [];
  }

  const seen = new Set();
  const result = [];

  for (const item of arr) {
    if (item === undefined || item === null) {
      continue;
    }

    const normalized = String(item).trim();

    if (!normalized || seen.has(normalized)) {
      continue;
    }

    seen.add(normalized);
    result.push(normalized);
  }

  return result;
}

/**
 * Obtém o código da empresa considerando input existente.
 * @returns {string} Código em caixa alta ou string vazia.
 */
function getCurrentCompanyCode() {
  return (document.getElementById('company')?.value || '').trim().toUpperCase();
}

/**
 * Retorna lista de opções de PEP conforme empresa.
 * @param {string} companyCode - Código atual da empresa.
 * @returns {string[]} Lista de opções disponíveis.
 */
function getPepOptionsForCompany(companyCode) {
  return uniquePreserveOrder(companyCode === 'BF00' ? PEP_OPTIONS_BF00 : PEP_OPTIONS_DEFAULT);
}

/**
 * Popula SELECT com opções de PEP preservando valor atual.
 * @param {HTMLSelectElement} selectEl - Elemento SELECT alvo.
 * @param {string[]} options - Opções disponíveis.
 * @param {string} selected - Valor previamente selecionado.
 */
function populatePepTitleSelect(selectEl, options, selected) {
  if (!selectEl) {
    return;
  }

  const normalizedSelected = selected == null ? '' : String(selected).trim();
  const normalizedOptions = Array.isArray(options) ? options : [];

  while (selectEl.firstChild) {
    selectEl.removeChild(selectEl.firstChild);
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Selecione…';
  if (!normalizedSelected) {
    placeholder.selected = true;
  }
  selectEl.appendChild(placeholder);

  let hasSelected = false;

  for (const optionValue of normalizedOptions) {
    const option = document.createElement('option');
    option.value = optionValue;
    option.textContent = optionValue;
    if (optionValue === normalizedSelected) {
      option.selected = true;
      hasSelected = true;
    }
    selectEl.appendChild(option);
  }

  if (normalizedSelected && !hasSelected) {
    const existingOption = document.createElement('option');
    existingOption.value = normalizedSelected;
    existingOption.textContent = `${normalizedSelected} (existente)`;
    existingOption.selected = true;
    selectEl.appendChild(existingOption);
  }
}

/**
 * Garante que o campo de título do PEP seja SELECT, preservando atributos.
 * @param {HTMLElement} el - Elemento original.
 * @returns {{select: HTMLSelectElement|null, currentValue: string}} Elemento SELECT e valor atual.
 */
function ensurePepTitleSelect(el) {
  if (!el) {
    return { select: null, currentValue: '' };
  }

  if (el.tagName && el.tagName.toUpperCase() === 'SELECT') {
    return { select: /** @type {HTMLSelectElement} */ (el), currentValue: el.value != null ? String(el.value) : '' };
  }

  const currentValue = el.value != null ? String(el.value) : '';
  const select = document.createElement('select');

  for (const attr of Array.from(el.attributes)) {
    if (attr.name === 'type' || attr.name === 'value') {
      continue;
    }

    if (attr.name === 'class') {
      select.className = el.className || '';
    } else {
      select.setAttribute(attr.name, attr.value);
    }
  }

  if (!select.className && el.className) {
    select.className = el.className;
  }

  if (el.dataset) {
    for (const [key, value] of Object.entries(el.dataset)) {
      select.dataset[key] = value;
    }
  }

  if (typeof el.required === 'boolean' && el.required) {
    select.required = true;
  }

  if (typeof el.disabled === 'boolean') {
    select.disabled = el.disabled;
  }

  el.replaceWith(select);

  return { select, currentValue };
}

/**
 * Inicializa uma linha de PEP convertendo e populando o título.
 * @param {Element} rowEl - Linha de PEP alvo.
 */
function initPepRow(rowEl) {
  if (!rowEl) {
    return;
  }

  const control = rowEl.querySelector('.pep-title, [name="pepTitle"]');
  if (!control) {
    return;
  }

  const { select, currentValue } = ensurePepTitleSelect(/** @type {HTMLElement} */ (control));
  if (!select) {
    return;
  }

  const options = getPepOptionsForCompany(getCurrentCompanyCode());
  populatePepTitleSelect(select, options, currentValue);
}

/**
 * Atualiza todas as linhas de PEP existentes.
 */
function refreshAllPepTitleSelects() {
  const rows = document.querySelectorAll('.pep-row');
  rows.forEach(initPepRow);
}

let boundCompanyElement = null;
let pepRowsObserver = null;

function bindCompanyChangeListener() {
  const companyField = document.getElementById('company');
  if (!companyField || companyField === boundCompanyElement) {
    return;
  }

  if (boundCompanyElement) {
    boundCompanyElement.removeEventListener('change', refreshAllPepTitleSelects);
  }

  companyField.addEventListener('change', refreshAllPepTitleSelects);
  boundCompanyElement = companyField;
}

function observePepRowMutations() {
  if (pepRowsObserver || typeof MutationObserver !== 'function' || !document.body) {
    return;
  }

  pepRowsObserver = new MutationObserver((mutations) => {
    let shouldRebindCompany = false;

    for (const mutation of mutations) {
      for (const node of mutation.addedNodes) {
        if (!(node instanceof Element)) {
          continue;
        }

        if (!shouldRebindCompany && (node.id === 'company' || node.querySelector('#company'))) {
          shouldRebindCompany = true;
        }

        if (node.classList.contains('pep-row')) {
          initPepRow(node);
        }

        node.querySelectorAll('.pep-row').forEach(initPepRow);
      }
    }

    if (shouldRebindCompany) {
      bindCompanyChangeListener();
    }
  });

  pepRowsObserver.observe(document.body, { childList: true, subtree: true });
}

function bootstrapPepTitleDropdowns() {
  bindCompanyChangeListener();
  refreshAllPepTitleSelects();
  observePepRowMutations();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', bootstrapPepTitleDropdowns, { once: true });
} else {
  bootstrapPepTitleDropdowns();
}

// ============================================================================
// Execução
// ============================================================================
init();
