// Única fronteira HTTP/Apps Script. Sem estado de pedidos guardado globalmente.
function createAppsScriptApplication_() {
  const domain = createCondominiumDomain_(GARAGE_RESIDENTIAL_STRUCTURE, OPEN_STATUSES, PENDING_STATUSES);
  const ports = createAppsScriptPorts_(CONFIG, REQUIRED_HEADERS, domain);
  return createSecureApplication_(createCondominiumApplication_(ports, CONFIG, domain), ports);
}

function parseRequestBody_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      const payload = JSON.parse(e.postData.contents);
      if (payload && typeof payload === 'object') return payload;
    } catch (err) { /* Compatibilidade com clientes que enviam parâmetros de formulário. */ }
  }
  const payload = {};
  if (e && e.parameter) Object.keys(e.parameter).forEach(key => payload[key] = e.parameter[key]);
  return payload;
}

function routeRequest_(method, event) {
  const payload = method === 'GET' ? (event && event.parameter || {}) : parseRequestBody_(event);
  const rawAction = payload.action;
  const action = method === 'GET'
    ? (rawAction === null || rawAction === undefined ? '' : String(rawAction).trim()) || 'status'
    : String(rawAction || '').trim();
  const startedAt=Date.now();
  const result = createAppsScriptApplication_().dispatch(method, action, payload);
  if(action === 'general.status' || (action === 'status' && payload.details === 'public') || result.token && ['garage.loginPin','garage.loginAdmin','garageLoginPin','garageLoginAdmin'].includes(action)){
    result.serviceVersion='3.7.0-rc1';
    result.serverDurationMs=Math.max(0,Date.now()-startedAt);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) { return routeRequest_('GET', e); }
function doPost(e) { return routeRequest_('POST', e); }

// Instalação explícita: consultas e reportes não criam nem alteram o esquema das folhas.
function setupBackend_() {
  const result = createAppsScriptApplication_().initialize();
  Logger.log('Backend preparado.');
  return result;
}
function setupApp() { return setupBackend_(); }
function resetPinManualmente_() {
  createAppsScriptApplication_().clearLegacyPin();
  Logger.log('PIN removido.');
}

// Executar uma vez para acrescentar apenas a folha de ocorrências gerais.
function setupGeneralOccurrences() {
  const domain=createCondominiumDomain_(GARAGE_RESIDENTIAL_STRUCTURE,OPEN_STATUSES,PENDING_STATUSES);
  createAppsScriptPorts_(CONFIG,REQUIRED_HEADERS,domain).initializeGeneral();
  Logger.log('Folha OCORRENCIAS_GERAIS preparada. Dados existentes preservados.');
}
