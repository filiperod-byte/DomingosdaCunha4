// Adaptadores Google. Cada pedido recebe uma instância nova (sem cache entre pedidos).
function createAppsScriptPorts_(CONFIG, REQUIRED_HEADERS, domain) {
  const { safeText_, toInt_, normalizePoint_ } = domain;
  let spreadsheet;
  let settingsCache;
  let residentsCache;
  function residentRows() {
    if (!residentsCache) residentsCache = getSheetObjects_('CONDOMINOS');
    return residentsCache;
  }
function setupIfNeeded_() {
  const ss = getSpreadsheet_();
  ensureSheetStructure_(ss, CONFIG.SHEETS.REGISTOS, REQUIRED_HEADERS.REGISTOS);
  ensureSheetStructure_(ss, CONFIG.SHEETS.ESTADO, REQUIRED_HEADERS.ESTADO);
  ensureSheetStructure_(ss, CONFIG.SHEETS.FECHOS, REQUIRED_HEADERS.FECHOS);
}

function ensureSheetStructure_(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastColumn() === 0 || sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,expectedHeaders.length).setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const current = getHeaders_(sheet);
  const missing = expectedHeaders.filter(h => current.indexOf(h) === -1);
  if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getSpreadsheet_() { if (!spreadsheet) spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); return spreadsheet; }

function getSheet_(name) { const sh = getSpreadsheet_().getSheetByName(name); if (!sh) throw new Error('Folha não encontrada: ' + name); return sh; }

function getHeaders_(sheet) { const lastCol = sheet.getLastColumn(); if (lastCol < 1) return []; return sheet.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v || '').trim()); }

function getSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values=sheet.getDataRange().getValues();
  if(values.length<2)return [];
  const headers=values[0].map(v=>String(v||'').trim());
  return values.slice(1).map((row, idx) => {
    const obj = { _rowIndex: idx + 2 };
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}

function appendObjectRow_(sheetName, obj) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '')); }

function updateObjectRow_(sheetName, rowIndex, patch) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); const row = sh.getRange(rowIndex,1,1,headers.length).getValues()[0]; const merged = {}; headers.forEach((h,i) => merged[h] = row[i]); Object.keys(patch).forEach(k => merged[k] = patch[k]); sh.getRange(rowIndex,1,1,headers.length).setValues([headers.map(h => Object.prototype.hasOwnProperty.call(merged,h) ? merged[h] : '')]); }

function upsertStateRow_(floor, point, data) { const existing = findStateRowByPoint_(floor, point); if (existing) updateObjectRow_(CONFIG.SHEETS.ESTADO, existing._rowIndex, data); else appendObjectRow_(CONFIG.SHEETS.ESTADO, data); }

function findStateRowByPoint_(floor, point) { const f = toInt_(floor); const p = normalizePoint_(point); return getSheetObjects_(CONFIG.SHEETS.ESTADO).find(r => toInt_(r.FLOOR) === f && normalizePoint_(r.POINT) === p) || null; }

function getScriptProperties_() { return PropertiesService.getScriptProperties(); }

function hashPin_(pin, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '::' + String(pin), Utilities.Charset.UTF_8).map(b => { const v = b < 0 ? b + 256 : b; return (v < 16 ? '0' : '') + v.toString(16); }).join(''); }

function getRootFolder_() { return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID); }

function getOrCreateSubfolder_(name) { const root = getRootFolder_(); const it = root.getFoldersByName(name); return it.hasNext() ? it.next() : root.createFolder(name); }

function saveIncomingPhoto_(opts) {
  const base64 = safeText_(opts.base64);
  const dataUrl = safeText_(opts.dataUrl);
  let mimeType = safeText_(opts.mimeType);
  let fileName = safeText_(opts.fileName);
  if (!base64 && !dataUrl) return { fileId:'', fileUrl:'', fileName:'' };
  let finalBase64 = base64 || (dataUrl.indexOf(',') > -1 ? dataUrl.split(',')[1] : '');
  if (!mimeType && dataUrl.indexOf(';base64,') > -1) mimeType = dataUrl.substring(5, dataUrl.indexOf(';base64,'));
  if (!finalBase64) throw new Error('Foto inválida.');
  if (!mimeType) mimeType = 'image/jpeg';
  if (!fileName) fileName = buildDefaultPhotoName_(opts.prefix || 'file', opts.occurrenceId || Utilities.getUuid(), mimeType);
  const file = getOrCreateSubfolder_(safeText_(opts.folderName)).createFile(Utilities.newBlob(Utilities.base64Decode(finalBase64), mimeType, fileName));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { Logger.log(err); }
  return { fileId:file.getId(), fileUrl:file.getUrl(), fileName:file.getName() };
}

function buildDefaultPhotoName_(prefix, occurrenceId, mimeType) { return [prefix || 'file', occurrenceId || Utilities.getUuid(), Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss')].join('_') + '.' + extensionFromMimeType_(mimeType); }

function extensionFromMimeType_(mime) { return ({ 'image/jpeg':'jpg','image/jpg':'jpg','image/png':'png','image/webp':'webp','image/heic':'heic' })[mime] || 'jpg'; }

function safeSendEmail_(to, subject, body, opts) { try { if (!to) return; MailApp.sendEmail(Object.assign({ to, subject, body }, opts || {})); } catch (err) { Logger.log('Falha email: ' + err); } }

function withScriptLock_(cb) { const lock = LockService.getScriptLock(); lock.waitLock(30000); try { return cb(); } finally { lock.releaseLock(); } }

function generateOccurrenceId_() { return 'OCC-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMddHHmmss') + '-' + Utilities.getUuid().substring(0,8).toUpperCase(); }

function isoNow_() { return new Date().toISOString(); }

function getSheet(name) { const ss = getSpreadsheet_(); let sh = ss.getSheetByName(name); if (!sh) { sh = ss.insertSheet(name); initSheet(name, sh); } return sh; }

function initSheet(name, sheet) { const h = { CONDOMINOS:['ID','DataRegisto','Nome','Piso','Fracao','Email','Telemovel','Estado','PIN','PINAtivo','DataAprovacao','AprovadoPor','UltimoAcesso','Observacoes','TentativasFalhadas','BloqueadoAte'], CONSULTAS:['ID','DataHora','PINUsado','Nome','Piso','Fracao','Email','Motivo','MotivoOutro','CodigoMostrado','Resultado','UserAgent'], CONFIG:['Chave','Valor','Descricao'], LOG:['DataHora','Tipo','Acao','Email','Mensagem'] }; if (h[name]) { sheet.appendRow(h[name]); sheet.getRange(1,1,1,h[name].length).setFontWeight('bold'); } }

function setupApp() { const c = getSheet('CONFIG'); getSheet('CONDOMINOS'); getSheet('CONSULTAS'); getSheet('LOG'); const defaults = [['NOME_CONDOMINIO','Domingos da Cunha 4','Nome do condomínio'],['CODIGO_CADEADO','0000','Código atual do cadeado'],['ADMIN_EMAIL',CONFIG.ADMIN_EMAIL || 'admin@email.com','Email do administrador'],['ADMIN_PIN','123456','PIN do administrador'],['TEMPO_VISIVEL_SEGUNDOS','15','Segundos que o código fica visível'],['PIN_DIGITOS','6','Número de dígitos do PIN'],['MAX_TENTATIVAS_PIN','5','Tentativas antes de bloquear'],['BLOQUEIO_MINUTOS','15','Minutos de bloqueio'],['TEXTO_AVISO','Após utilização, confirme que a caixa fica corretamente fechada.','Aviso mostrado com o código']]; const keys = c.getDataRange().getValues().map(r => r[0]); defaults.forEach(r => { if (keys.indexOf(r[0]) === -1) c.appendRow(r); }); return 'Setup concluído'; }

  const residentFields = {
    id: 'ID', registeredAt: 'DataRegisto', name: 'Nome', floor: 'Piso', fraction: 'Fracao',
    email: 'Email', phone: 'Telemovel', status: 'Estado', pin: 'PIN', pinActive: 'PINAtivo',
    approvedAt: 'DataAprovacao', approvedBy: 'AprovadoPor', lastAccess: 'UltimoAcesso',
    notes: 'Observacoes', failedAttempts: 'TentativasFalhadas', blockedUntil: 'BloqueadoAte'
  };
  const consultationFields = {
    id: 'ID', at: 'DataHora', pin: 'PINUsado', name: 'Nome', floor: 'Piso', fraction: 'Fracao',
    email: 'Email', reason: 'Motivo', reasonOther: 'MotivoOutro', code: 'CodigoMostrado',
    result: 'Resultado', userAgent: 'UserAgent'
  };
  function decode(row, fields) {
    const record = { legacyRow: row._rowIndex };
    Object.keys(fields).forEach(key => record[key] = row[fields[key]]);
    return record;
  }
  function encode(record, fields) {
    const row = {};
    Object.keys(fields).forEach(key => {
      if (Object.prototype.hasOwnProperty.call(record, key)) row[fields[key]] = record[key];
    });
    return row;
  }
  function setting(key) {
    if (!settingsCache) {
      settingsCache = Object.create(null);
      getSheetObjects_('CONFIG').forEach(row => {
        if (!Object.prototype.hasOwnProperty.call(settingsCache, row.Chave)) settingsCache[row.Chave] = String(row.Valor);
      });
    }
    return Object.prototype.hasOwnProperty.call(settingsCache, key) ? settingsCache[key] : null;
  }
  function saveSetting(key, value) {
    const row = getSheetObjects_('CONFIG').find(row => row.Chave === key);
    if (row) updateObjectRow_('CONFIG', row._rowIndex, { Valor: value });
    else appendObjectRow_('CONFIG', { Chave: key, Valor: value, Descricao: '' });
    settingsCache = undefined;
  }
  function entityId(prefix) {
    return (prefix || 'ID') + '_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
  }
  return {
    initialize: function () {
      return withScriptLock_(function () {
        setupIfNeeded_();
        const result = setupApp();
        settingsCache = undefined;
        return result;
      });
    },
    clock: { now: () => new Date(), format: (date, pattern) => Utilities.formatDate(new Date(date), CONFIG.TZ, pattern) },
    ids: { occurrence: generateOccurrenceId_, entity: entityId, pinCandidate: () => String(Math.floor(100000 + Math.random() * 900000)) },
    crypto: { uuid: () => Utilities.getUuid(), hashPin: hashPin_ },
    properties: {
      getProperty: key => getScriptProperties_().getProperty(key),
      setProperties: (values, clear) => getScriptProperties_().setProperties(values, clear),
      deleteProperty: key => getScriptProperties_().deleteProperty(key)
    },
    lock: { run: withScriptLock_ },
    mail: { send: safeSendEmail_ },
    files: { save: saveIncomingPhoto_ },
    environment: { effectiveEmail: () => Session.getEffectiveUser().getEmail(), activeEmail: () => Session.getActiveUser().getEmail() },
    settings: { get: setting, set: saveSetting },
    occurrences: {
      list: () => getSheetObjects_(CONFIG.SHEETS.ESTADO),
      update: (row, patch) => updateObjectRow_(CONFIG.SHEETS.ESTADO, row, patch),
      upsert: upsertStateRow_
    },
    events: { append: record => appendObjectRow_(CONFIG.SHEETS.REGISTOS, record) },
    closures: { append: record => appendObjectRow_(CONFIG.SHEETS.FECHOS, record) },
    residents: {
      list: () => residentRows().map(row => decode(row, residentFields)),
      get: row => {
        const record = residentRows().find(record => record._rowIndex === row);
        if (!record) throw new Error('Morador não encontrado.');
        return decode(record, residentFields);
      },
      add: record => { try { return appendObjectRow_('CONDOMINOS', encode(record, residentFields)); } finally { residentsCache = undefined; } },
      update: (row, patch) => { try { return updateObjectRow_('CONDOMINOS', row, encode(patch, residentFields)); } finally { residentsCache = undefined; } }
    },
    consultations: {
      list: () => getSheetObjects_('CONSULTAS').map(row => decode(row, consultationFields)),
      add: record => appendObjectRow_('CONSULTAS', encode(record, consultationFields))
    },
    log: { add: (type, action, email, message) => appendObjectRow_('LOG', {
      DataHora: new Date(), Tipo: type, Acao: action, Email: email || '', Mensagem: message || ''
    }) }
  };
}
