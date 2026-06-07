/*************************************************
 * DOMINGOS DA CUNHA 4 - BACKEND UNIFICADO V3
 * Google Apps Script Web App
 *
 * Módulos:
 * - Extintores: reporte pendente, validação admin, ocorrências abertas, fecho, QR/backoffice
 * - Garagem/Cadeado: registo, aprovação, PIN de utilizador, código do quadro elétrico
 *************************************************/

const CONFIG = {
  APP_NAME: 'Domingos da Cunha 4',
  TZ: 'Europe/Lisbon',
  SPREADSHEET_ID: '16vz1ZIKkCI7NfAf2ZSVCNLEafBn4Svtu7UaRe0WdpBU',
  ROOT_FOLDER_ID: '1Ppu0Tk5zLrOWBQGwJtH_GE2kpSanWyb9',
  ADMIN_EMAIL: 'filiperod@gmail.com',
  PIN_MIN_LENGTH: 4,
  SHEETS: {
    REGISTOS: 'REGISTOS',
    ESTADO: 'ESTADO_EXTINTORES',
    FECHOS: 'HISTORICO_FECHOS'
  },
  SUBFOLDERS: {
    REPORTS: 'REPORTES_EXTINTORES',
    CLOSES: 'FECHOS_EXTINTORES'
  },
  PROPS: {
    PIN_HASH: 'BACKOFFICE_PIN_HASH',
    PIN_SALT: 'BACKOFFICE_PIN_SALT',
    PIN_UPDATED_AT: 'BACKOFFICE_PIN_UPDATED_AT'
  }
};

const REQUIRED_HEADERS = {
  REGISTOS: [
    'TIMESTAMP','EVENTO','OCCURRENCE_ID','FLOOR','FLOOR_LABEL','POINT','LOCATION','REPORTED_BY','REASON','NOTES','PHOTO_FILE_ID','PHOTO_FILE_URL','PHOTO_FILE_NAME','STATUS_BEFORE','STATUS_AFTER','SOURCE','CLIENT_TS','EXTRA_JSON'
  ],
  ESTADO: [
    'LAST_UPDATED','STATUS','OCCURRENCE_ID','FLOOR','FLOOR_LABEL','POINT','LOCATION','REPORTED_AT','REPORTED_BY','REASON','NOTES','PHOTO_FILE_ID','PHOTO_FILE_URL','PHOTO_FILE_NAME','CLOSED_AT','CLOSE_NOTES','CLOSE_PHOTO_FILE_ID','CLOSE_PHOTO_FILE_URL','CLOSE_PHOTO_FILE_NAME','LAST_EVENT'
  ],
  FECHOS: [
    'TIMESTAMP','OCCURRENCE_ID','FLOOR','FLOOR_LABEL','POINT','LOCATION','REPORTED_AT','REPORTED_BY','REASON','OPEN_NOTES','CLOSE_NOTES','CLOSE_PHOTO_FILE_ID','CLOSE_PHOTO_FILE_URL','CLOSE_PHOTO_FILE_NAME','SOURCE','CLIENT_TS'
  ]
};

const OPEN_STATUSES = ['ABERTA', 'OPEN', 'ALERTA', 'REPORTADO'];
const PENDING_STATUSES = ['PENDENTE_VALIDACAO', 'PENDENTE', 'PENDING_VALIDATION'];

const GARAGE_RESIDENTIAL_STRUCTURE = {
  '10': ['A', 'B'],
  '9': ['A', 'B', 'C'],
  '8': ['A', 'B', 'C'],
  '7': ['A', 'B', 'C'],
  '6': ['A', 'B', 'C'],
  '5': ['A', 'B', 'C'],
  '4': ['A', 'B', 'C'],
  '3': ['A', 'B', 'C'],
  '2': ['A', 'B', 'C'],
  '-1': ['GARAGEM'],
  '-2': ['GARAGEM'],
  '-3': ['LUGAR_GARAGEM']
};

function doGet(e) { return routeRequest_('GET', e); }
function doPost(e) { return routeRequest_('POST', e); }

function routeRequest_(method, e) {
  try {
    setupIfNeeded_();
    if (method === 'GET') {
      const action = getActionFromGet_(e) || 'status';
      switch (action) {
        case 'status': return jsonResponse_(handleGetStatus_());
        case 'pinStatus': return jsonResponse_(handleGetPinStatus_());
        case 'openOccurrences': return jsonResponse_(handleGetOpenOccurrences_());
        case 'pendingOccurrences': return jsonResponse_(handleGetPendingOccurrences_());
        case 'health': return jsonResponse_(handleGetHealth_());
        case 'garage.publicConfig':
        case 'garagePublicConfig':
        case 'garage.structure':
        case 'garageStructure':
        case 'garage.dashboard':
        case 'garageDashboard':
        case 'garage.pending':
        case 'garagePending':
        case 'garage.approved':
        case 'garageApproved':
        case 'garage.history':
        case 'garageHistory':
        case 'garage.adminConfig':
        case 'garageAdminConfig':
          return jsonResponse_(handleGarageGet_(action));
        default:
          return jsonResponse_({ success:false, message:'Ação GET inválida: ' + action });
      }
    }

    const payload = parseRequestBody_(e);
    const action = String(payload.action || '').trim();
    switch (action) {
      case 'report': return jsonResponse_(handlePostReport_(payload));
      case 'approveOccurrence': return jsonResponse_(handlePostApproveOccurrence_(payload));
      case 'rejectOccurrence': return jsonResponse_(handlePostRejectOccurrence_(payload));
      case 'setPin': return jsonResponse_(handlePostSetPin_(payload));
      case 'validatePin': return jsonResponse_(handlePostValidatePin_(payload));
      case 'resetPin': return jsonResponse_(handlePostResetPin_(payload));
      case 'closeOccurrence': return jsonResponse_(handlePostCloseOccurrence_(payload));
      case 'garage.register':
      case 'garageRegister':
      case 'garage.loginPin':
      case 'garageLoginPin':
      case 'garage.failedAttempt':
      case 'garageFailedAttempt':
      case 'garage.resendPin':
      case 'garageResendPin':
      case 'garage.recoverCode':
      case 'garageRecoverCode':
      case 'garage.getCode':
      case 'garageGetCode':
      case 'garage.loginAdmin':
      case 'garageLoginAdmin':
      case 'garage.approve':
      case 'garageApprove':
      case 'garage.reject':
      case 'garageReject':
      case 'garage.block':
      case 'garageBlock':
      case 'garage.unblock':
      case 'garageUnblock':
      case 'garage.regeneratePin':
      case 'garageRegeneratePin':
      case 'garage.changeCode':
      case 'garageChangeCode':
      case 'garage.saveConfig':
      case 'garageSaveConfig':
        return jsonResponse_(handleGaragePost_(action, payload));
      default:
        return jsonResponse_({ success:false, message:'Ação POST inválida: ' + action });
    }
  } catch (err) {
    return jsonResponse_({ success:false, message: err && err.message ? err.message : 'Erro interno no backend' });
  }
}

function handleGetStatus_() {
  const openRows = getOpenStateRows_();
  return {
    success:true,
    reported: openRows.map(row => ({ floor: toInt_(row.FLOOR), point: String(row.POINT || '').trim() })),
    totalOpen: openRows.length,
    updatedAt: isoNow_()
  };
}

function handleGetPinStatus_() {
  return { success:true, pinConfigured:isPinConfigured_(), updatedAt:getScriptProperties_().getProperty(CONFIG.PROPS.PIN_UPDATED_AT) || '' };
}

function publicOccurrenceFromRow_(row) {
  return {
    id: String(row.OCCURRENCE_ID || ''),
    floor: toInt_(row.FLOOR),
    floorLabel: String(row.FLOOR_LABEL || formatFloorLabel_(row.FLOOR)),
    point: String(row.POINT || ''),
    location: String(row.LOCATION || ''),
    reportedBy: String(row.REPORTED_BY || ''),
    reason: String(row.REASON || ''),
    notes: String(row.NOTES || ''),
    createdAt: String(row.REPORTED_AT || ''),
    photoUrl: String(row.PHOTO_FILE_URL || ''),
    status: String(row.STATUS || '')
  };
}

function handleGetOpenOccurrences_() {
  const rows = getOpenStateRows_().sort((a,b) => String(b.REPORTED_AT || '').localeCompare(String(a.REPORTED_AT || ''))).map(publicOccurrenceFromRow_);
  return { success:true, occurrences: rows };
}

function handleGetPendingOccurrences_() {
  const rows = getPendingStateRows_().sort((a,b) => String(b.REPORTED_AT || '').localeCompare(String(a.REPORTED_AT || ''))).map(publicOccurrenceFromRow_);
  return { success:true, occurrences: rows };
}

function handleGetHealth_() {
  return {
    success:true,
    appName: CONFIG.APP_NAME,
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    rootFolderId: CONFIG.ROOT_FOLDER_ID,
    adminEmail: getAdminEmail_({}),
    pinConfigured: isPinConfigured_(),
    garageConfigReady: !!getConfig('NOME_CONDOMINIO'),
    now: isoNow_()
  };
}

function handlePostReport_(payload) {
  return withScriptLock_(function () {
    const floor = toInt_(payload.floor);
    const point = normalizePoint_(payload.point);
    const location = safeText_(payload.location);
    const reportedBy = safeText_(payload.name || payload.reportedBy);
    const reason = safeText_(payload.reason);
    const notes = safeText_(payload.notes);
    const source = safeText_(payload.source);
    const clientTs = safeText_(payload.clientTs);

    if (isNaN(floor)) throw new Error('Piso inválido.');
    if (!point) throw new Error('Ponto obrigatório.');
    if (!reportedBy) throw new Error('Nome obrigatório.');
    if (!reason) throw new Error('Motivo obrigatório.');

    const floorLabel = formatFloorLabel_(floor);
    const nowIso = isoNow_();
    const existingState = findStateRowByPoint_(floor, point);
    const alreadyOpen = !!(existingState && isOpenStatus_(existingState.STATUS));
    const alreadyPending = !!(existingState && isPendingStatus_(existingState.STATUS));
    const occurrenceId = (alreadyOpen || alreadyPending) ? String(existingState.OCCURRENCE_ID || '') : generateOccurrenceId_();
    const statusAfter = alreadyOpen ? 'ABERTA' : 'PENDENTE_VALIDACAO';
    const eventName = alreadyOpen ? 'REPORT_ADICIONAL' : (alreadyPending ? 'REPORT_ATUALIZADO_PENDENTE' : 'REPORT_PENDENTE_VALIDACAO');

    const photo = saveIncomingPhoto_({
      base64: payload.photoBase64,
      dataUrl: payload.photoDataUrl,
      mimeType: payload.photoType,
      fileName: payload.photoName,
      folderName: CONFIG.SUBFOLDERS.REPORTS,
      occurrenceId: occurrenceId,
      prefix: 'report'
    });

    appendObjectRow_(CONFIG.SHEETS.REGISTOS, {
      TIMESTAMP: nowIso,
      EVENTO: eventName,
      OCCURRENCE_ID: occurrenceId,
      FLOOR: floor,
      FLOOR_LABEL: floorLabel,
      POINT: point,
      LOCATION: location,
      REPORTED_BY: reportedBy,
      REASON: reason,
      NOTES: notes,
      PHOTO_FILE_ID: photo.fileId,
      PHOTO_FILE_URL: photo.fileUrl,
      PHOTO_FILE_NAME: photo.fileName,
      STATUS_BEFORE: existingState ? String(existingState.STATUS || '') : '',
      STATUS_AFTER: statusAfter,
      SOURCE: source,
      CLIENT_TS: clientTs,
      EXTRA_JSON: JSON.stringify({ alreadyOpen: alreadyOpen, pendingValidation: !alreadyOpen })
    });

    upsertStateRow_(floor, point, {
      LAST_UPDATED: nowIso,
      STATUS: statusAfter,
      OCCURRENCE_ID: occurrenceId,
      FLOOR: floor,
      FLOOR_LABEL: floorLabel,
      POINT: point,
      LOCATION: location || (existingState ? String(existingState.LOCATION || '') : ''),
      REPORTED_AT: alreadyOpen && existingState ? String(existingState.REPORTED_AT || nowIso) : nowIso,
      REPORTED_BY: reportedBy,
      REASON: reason,
      NOTES: notes,
      PHOTO_FILE_ID: photo.fileId || (existingState ? String(existingState.PHOTO_FILE_ID || '') : ''),
      PHOTO_FILE_URL: photo.fileUrl || (existingState ? String(existingState.PHOTO_FILE_URL || '') : ''),
      PHOTO_FILE_NAME: photo.fileName || (existingState ? String(existingState.PHOTO_FILE_NAME || '') : ''),
      CLOSED_AT: '',
      CLOSE_NOTES: '',
      CLOSE_PHOTO_FILE_ID: '',
      CLOSE_PHOTO_FILE_URL: '',
      CLOSE_PHOTO_FILE_NAME: '',
      LAST_EVENT: eventName
    });

    sendReportEmail_({ adminEmail:getAdminEmail_(payload), occurrenceId, floor, floorLabel, point, location, reportedBy, reason, notes, photoUrl:photo.fileUrl, alreadyOpen, pendingValidation: !alreadyOpen });

    return { success:true, occurrenceId, alreadyOpen, pendingValidation: !alreadyOpen, floor, point };
  });
}

function handlePostApproveOccurrence_(payload) {
  return withScriptLock_(function () {
    const row = findPendingStateRow_(safeText_(payload.occurrenceId), toInt_(payload.floor), normalizePoint_(payload.point));
    if (!row) throw new Error('Reporte pendente não encontrado.');
    const nowIso = isoNow_();
    updateObjectRow_(CONFIG.SHEETS.ESTADO, row._rowIndex, { LAST_UPDATED: nowIso, STATUS: 'ABERTA', LAST_EVENT: 'VALIDACAO_ADMIN' });
    appendObjectRow_(CONFIG.SHEETS.REGISTOS, {
      TIMESTAMP: nowIso,
      EVENTO: 'VALIDACAO_ADMIN',
      OCCURRENCE_ID: String(row.OCCURRENCE_ID || ''),
      FLOOR: toInt_(row.FLOOR),
      FLOOR_LABEL: String(row.FLOOR_LABEL || formatFloorLabel_(row.FLOOR)),
      POINT: String(row.POINT || ''),
      LOCATION: String(row.LOCATION || ''),
      REPORTED_BY: safeText_(payload.adminEmail || 'Admin'),
      REASON: 'VALIDADO',
      NOTES: safeText_(payload.notes || ''),
      STATUS_BEFORE: String(row.STATUS || ''),
      STATUS_AFTER: 'ABERTA',
      SOURCE: safeText_(payload.source || 'backoffice')
    });
    return { success:true, occurrenceId:String(row.OCCURRENCE_ID || '') };
  });
}

function handlePostRejectOccurrence_(payload) {
  return withScriptLock_(function () {
    const row = findPendingStateRow_(safeText_(payload.occurrenceId), toInt_(payload.floor), normalizePoint_(payload.point));
    if (!row) throw new Error('Reporte pendente não encontrado.');
    const nowIso = isoNow_();
    updateObjectRow_(CONFIG.SHEETS.ESTADO, row._rowIndex, { LAST_UPDATED: nowIso, STATUS: 'REJEITADA', CLOSED_AT: nowIso, CLOSE_NOTES: safeText_(payload.notes || 'Rejeitado pela administração'), LAST_EVENT: 'REJEICAO_ADMIN' });
    appendObjectRow_(CONFIG.SHEETS.REGISTOS, {
      TIMESTAMP: nowIso,
      EVENTO: 'REJEICAO_ADMIN',
      OCCURRENCE_ID: String(row.OCCURRENCE_ID || ''),
      FLOOR: toInt_(row.FLOOR),
      FLOOR_LABEL: String(row.FLOOR_LABEL || formatFloorLabel_(row.FLOOR)),
      POINT: String(row.POINT || ''),
      LOCATION: String(row.LOCATION || ''),
      REPORTED_BY: safeText_(payload.adminEmail || 'Admin'),
      REASON: 'REJEITADO',
      NOTES: safeText_(payload.notes || ''),
      STATUS_BEFORE: String(row.STATUS || ''),
      STATUS_AFTER: 'REJEITADA',
      SOURCE: safeText_(payload.source || 'backoffice')
    });
    return { success:true, occurrenceId:String(row.OCCURRENCE_ID || '') };
  });
}

function handlePostSetPin_(payload) {
  return withScriptLock_(function () {
    const pin = normalizePin_(payload.pin);
    if (!pin) throw new Error('PIN obrigatório.');
    if (pin.length < CONFIG.PIN_MIN_LENGTH) throw new Error('O PIN deve ter pelo menos ' + CONFIG.PIN_MIN_LENGTH + ' dígitos.');
    if (!/^\d+$/.test(pin)) throw new Error('O PIN deve conter apenas dígitos.');
    if (isPinConfigured_()) throw new Error('O PIN já está definido.');
    const salt = Utilities.getUuid();
    getScriptProperties_().setProperties({ [CONFIG.PROPS.PIN_HASH]: hashPin_(pin, salt), [CONFIG.PROPS.PIN_SALT]: salt, [CONFIG.PROPS.PIN_UPDATED_AT]: isoNow_() }, true);
    return { success:true, message:'PIN definido com sucesso.' };
  });
}

function handlePostValidatePin_(payload) {
  const pin = normalizePin_(payload.pin);
  if (!isPinConfigured_()) return { success:true, valid:false, message:'Ainda não existe PIN definido.' };
  if (!pin) return { success:true, valid:false, message:'PIN vazio.' };
  const props = getScriptProperties_();
  return { success:true, valid: hashPin_(pin, props.getProperty(CONFIG.PROPS.PIN_SALT) || '') === (props.getProperty(CONFIG.PROPS.PIN_HASH) || '') };
}

function handlePostResetPin_(payload) {
  return withScriptLock_(function () {
    clearPin_();
    sendPinResetEmail_({ adminEmail:getAdminEmail_(payload) });
    return { success:true, message:'PIN resetado.' };
  });
}

function handlePostCloseOccurrence_(payload) {
  return withScriptLock_(function () {
    const occurrenceId = safeText_(payload.occurrenceId);
    const floor = toInt_(payload.floor);
    const point = normalizePoint_(payload.point);
    const closeNotes = safeText_(payload.closeNotes || payload.notes);
    const source = safeText_(payload.source);
    const clientTs = safeText_(payload.clientTs);
    let stateRow = occurrenceId ? findOpenStateRowByOccurrenceId_(occurrenceId) : null;
    if (!stateRow && !isNaN(floor) && point) stateRow = findOpenStateRowByPoint_(floor, point);
    if (!stateRow) throw new Error('Ocorrência aberta não encontrada.');

    const closePhoto = saveIncomingPhoto_({ base64:payload.closePhotoBase64, dataUrl:payload.closePhotoDataUrl, mimeType:payload.closePhotoType, fileName:payload.closePhotoName, folderName:CONFIG.SUBFOLDERS.CLOSES, occurrenceId:String(stateRow.OCCURRENCE_ID || ''), prefix:'close' });
    if (!closePhoto.fileId) throw new Error('A fotografia de fecho é obrigatória.');
    const nowIso = isoNow_();

    appendObjectRow_(CONFIG.SHEETS.FECHOS, {
      TIMESTAMP: nowIso,
      OCCURRENCE_ID: String(stateRow.OCCURRENCE_ID || ''),
      FLOOR: toInt_(stateRow.FLOOR),
      FLOOR_LABEL: String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)),
      POINT: String(stateRow.POINT || ''),
      LOCATION: String(stateRow.LOCATION || ''),
      REPORTED_AT: String(stateRow.REPORTED_AT || ''),
      REPORTED_BY: String(stateRow.REPORTED_BY || ''),
      REASON: String(stateRow.REASON || ''),
      OPEN_NOTES: String(stateRow.NOTES || ''),
      CLOSE_NOTES: closeNotes,
      CLOSE_PHOTO_FILE_ID: closePhoto.fileId,
      CLOSE_PHOTO_FILE_URL: closePhoto.fileUrl,
      CLOSE_PHOTO_FILE_NAME: closePhoto.fileName,
      SOURCE: source,
      CLIENT_TS: clientTs
    });

    appendObjectRow_(CONFIG.SHEETS.REGISTOS, {
      TIMESTAMP: nowIso, EVENTO:'FECHO', OCCURRENCE_ID:String(stateRow.OCCURRENCE_ID || ''), FLOOR:toInt_(stateRow.FLOOR), FLOOR_LABEL:String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)), POINT:String(stateRow.POINT || ''), LOCATION:String(stateRow.LOCATION || ''), REASON:'FECHO', NOTES:closeNotes, PHOTO_FILE_ID:closePhoto.fileId, PHOTO_FILE_URL:closePhoto.fileUrl, PHOTO_FILE_NAME:closePhoto.fileName, STATUS_BEFORE:String(stateRow.STATUS || ''), STATUS_AFTER:'OK', SOURCE:source, CLIENT_TS:clientTs
    });

    updateObjectRow_(CONFIG.SHEETS.ESTADO, stateRow._rowIndex, { LAST_UPDATED:nowIso, STATUS:'OK', CLOSED_AT:nowIso, CLOSE_NOTES:closeNotes, CLOSE_PHOTO_FILE_ID:closePhoto.fileId, CLOSE_PHOTO_FILE_URL:closePhoto.fileUrl, CLOSE_PHOTO_FILE_NAME:closePhoto.fileName, LAST_EVENT:'FECHO' });
    sendCloseEmail_({ adminEmail:getAdminEmail_(payload), occurrenceId:String(stateRow.OCCURRENCE_ID || ''), floor:toInt_(stateRow.FLOOR), floorLabel:String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)), point:String(stateRow.POINT || ''), location:String(stateRow.LOCATION || ''), reason:String(stateRow.REASON || ''), reportedBy:String(stateRow.REPORTED_BY || ''), closeNotes, closePhotoUrl:closePhoto.fileUrl });
    return { success:true, occurrenceId:String(stateRow.OCCURRENCE_ID || ''), floor:toInt_(stateRow.FLOOR), point:String(stateRow.POINT || '') };
  });
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

function getSpreadsheet_() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getSheet_(name) { const sh = getSpreadsheet_().getSheetByName(name); if (!sh) throw new Error('Folha não encontrada: ' + name); return sh; }
function getHeaders_(sheet) { const lastCol = sheet.getLastColumn(); if (lastCol < 1) return []; return sheet.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v || '').trim()); }
function getSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !headers.length) return [];
  return sheet.getRange(2,1,lastRow-1,headers.length).getValues().map((row, idx) => {
    const obj = { _rowIndex: idx + 2 };
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}
function appendObjectRow_(sheetName, obj) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '')); }
function updateObjectRow_(sheetName, rowIndex, patch) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); const row = sh.getRange(rowIndex,1,1,headers.length).getValues()[0]; const merged = {}; headers.forEach((h,i) => merged[h] = row[i]); Object.keys(patch).forEach(k => merged[k] = patch[k]); sh.getRange(rowIndex,1,1,headers.length).setValues([headers.map(h => Object.prototype.hasOwnProperty.call(merged,h) ? merged[h] : '')]); }
function upsertStateRow_(floor, point, data) { const existing = findStateRowByPoint_(floor, point); if (existing) updateObjectRow_(CONFIG.SHEETS.ESTADO, existing._rowIndex, data); else appendObjectRow_(CONFIG.SHEETS.ESTADO, data); }
function findStateRowByPoint_(floor, point) { const f = toInt_(floor); const p = normalizePoint_(point); return getSheetObjects_(CONFIG.SHEETS.ESTADO).find(r => toInt_(r.FLOOR) === f && normalizePoint_(r.POINT) === p) || null; }
function findOpenStateRowByPoint_(floor, point) { const r = findStateRowByPoint_(floor, point); return r && isOpenStatus_(r.STATUS) ? r : null; }
function findOpenStateRowByOccurrenceId_(id) { const target = safeText_(id); if (!target) return null; return getSheetObjects_(CONFIG.SHEETS.ESTADO).find(r => safeText_(r.OCCURRENCE_ID) === target && isOpenStatus_(r.STATUS)) || null; }
function findPendingStateRow_(id, floor, point) { const rows = getSheetObjects_(CONFIG.SHEETS.ESTADO); const target = safeText_(id); if (target) return rows.find(r => safeText_(r.OCCURRENCE_ID) === target && isPendingStatus_(r.STATUS)) || null; if (!isNaN(floor) && point) return rows.find(r => toInt_(r.FLOOR) === floor && normalizePoint_(r.POINT) === point && isPendingStatus_(r.STATUS)) || null; return null; }
function getOpenStateRows_() { return getSheetObjects_(CONFIG.SHEETS.ESTADO).filter(r => isOpenStatus_(r.STATUS)); }
function getPendingStateRows_() { return getSheetObjects_(CONFIG.SHEETS.ESTADO).filter(r => isPendingStatus_(r.STATUS)); }

function getScriptProperties_() { return PropertiesService.getScriptProperties(); }
function isPinConfigured_() { const p = getScriptProperties_(); return !!(p.getProperty(CONFIG.PROPS.PIN_HASH) && p.getProperty(CONFIG.PROPS.PIN_SALT)); }
function clearPin_() { const p = getScriptProperties_(); p.deleteProperty(CONFIG.PROPS.PIN_HASH); p.deleteProperty(CONFIG.PROPS.PIN_SALT); p.deleteProperty(CONFIG.PROPS.PIN_UPDATED_AT); }
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

function getAdminEmail_(payload) { const fromPayload = safeText_(payload && payload.adminEmail); if (isRealEmail_(fromPayload)) return fromPayload; if (isRealEmail_(CONFIG.ADMIN_EMAIL)) return CONFIG.ADMIN_EMAIL; const eff = safeText_(Session.getEffectiveUser().getEmail()); if (isRealEmail_(eff)) return eff; const act = safeText_(Session.getActiveUser().getEmail()); if (isRealEmail_(act)) return act; return ''; }
function isRealEmail_(email) { return !!email && String(email).indexOf('@') > -1 && String(email).indexOf('COLOCAR_') === -1; }
function safeSendEmail_(to, subject, body, opts) { try { if (!to) return; MailApp.sendEmail(Object.assign({ to, subject, body }, opts || {})); } catch (err) { Logger.log('Falha email: ' + err); } }
function sendReportEmail_(d) { if (!d.adminEmail) return; const status = d.pendingValidation ? 'PENDENTE DE VALIDAÇÃO' : 'ABERTA'; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reporte extintor - ' + status, ['Reporte de extintor recebido.', '', 'Estado: ' + status, 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Localização: ' + (d.location || '—'), 'Reportado por: ' + d.reportedBy, 'Motivo: ' + d.reason, 'Observação: ' + (d.notes || '—'), 'Foto: ' + (d.photoUrl || '—'), '', 'Data/Hora: ' + isoNow_()].join('\n')); }
function sendCloseEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Ocorrência fechada', ['Ocorrência fechada.', '', 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Foto fecho: ' + (d.closePhotoUrl || '—'), 'Observação: ' + (d.closeNotes || '—')].join('\n')); }
function sendPinResetEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reset do PIN Backoffice', 'O PIN do backoffice foi resetado.'); }

function parseRequestBody_(e) { if (e && e.postData && e.postData.contents) { try { const p = JSON.parse(e.postData.contents); if (p && typeof p === 'object') return p; } catch (err) {} } const out = {}; if (e && e.parameter) Object.keys(e.parameter).forEach(k => out[k] = e.parameter[k]); return out; }
function getActionFromGet_(e) { return e && e.parameter ? safeText_(e.parameter.action) : ''; }
function jsonResponse_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function withScriptLock_(cb) { const lock = LockService.getScriptLock(); lock.waitLock(30000); try { return cb(); } finally { lock.releaseLock(); } }
function generateOccurrenceId_() { return 'OCC-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMddHHmmss') + '-' + Utilities.getUuid().substring(0,8).toUpperCase(); }
function isoNow_() { return new Date().toISOString(); }
function safeText_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function toInt_(v) { if (v === null || v === undefined || v === '') return NaN; return parseInt(String(v).trim(), 10); }
function normalizePoint_(v) { return safeText_(v).toUpperCase(); }
function normalizePin_(v) { return safeText_(v).replace(/\s+/g, ''); }
function formatFloorLabel_(floor) { const n = toInt_(floor); if (isNaN(n)) return String(floor || ''); if (n < 0) return 'p' + n; return String(n); }
function isOpenStatus_(s) { return OPEN_STATUSES.indexOf(safeText_(s).toUpperCase()) > -1; }
function isPendingStatus_(s) { return PENDING_STATUSES.indexOf(safeText_(s).toUpperCase()) > -1; }
function setupBackend_() { setupIfNeeded_(); Logger.log('Backend preparado.'); }
function resetPinManualmente_() { clearPin_(); Logger.log('PIN removido.'); }

/*************************************************
 * GARAGEM / CADEADO
 *************************************************/
function handleGarageGet_(action) {
  setupApp();
  switch (action) {
    case 'garage.publicConfig': case 'garagePublicConfig': return getPublicConfig();
    case 'garage.structure': case 'garageStructure': return getGarageStructure_();
    case 'garage.dashboard': case 'garageDashboard': return getDashboardData();
    case 'garage.pending': case 'garagePending': return getPedidosPendentes();
    case 'garage.approved': case 'garageApproved': return getConominiosAprovados();
    case 'garage.history': case 'garageHistory': return getHistoricoConsultas(50);
    case 'garage.adminConfig': case 'garageAdminConfig': return getAdminConfig();
    default: return { ok:false, success:false, msg:'Ação GET Garagem inválida: ' + action };
  }
}
function handleGaragePost_(action, payload) {
  setupApp();
  switch (action) {
    case 'garage.register': case 'garageRegister': return registarCondomino({ nome:payload.nome || payload.name, piso:payload.piso || payload.floor, fracao:payload.fracao || payload.fraction, email:payload.email, telemovel:payload.telemovel || payload.telefone || payload.phone });
    case 'garage.loginPin': case 'garageLoginPin': return loginComPIN(payload.pin, payload.userAgent || payload.ua || '');
    case 'garage.failedAttempt': case 'garageFailedAttempt': registarTentativaFalhada(payload.pin); return { ok:true, success:true };
    case 'garage.resendPin': case 'garageResendPin': case 'garage.recoverCode': case 'garageRecoverCode': return reenviarPIN(payload.email);
    case 'garage.getCode': case 'garageGetCode': return obterCodigo(payload.pin, payload.motivo || payload.reason, payload.motivoOutro || payload.reasonOther || '', payload.userAgent || payload.ua || '');
    case 'garage.loginAdmin': case 'garageLoginAdmin': return loginAdmin(payload.email, payload.pin);
    case 'garage.approve': case 'garageApprove': return aprovarCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');
    case 'garage.reject': case 'garageReject': return rejeitarCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');
    case 'garage.block': case 'garageBlock': return bloquearCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');
    case 'garage.unblock': case 'garageUnblock': return desbloquearCondomino(Number(payload.row));
    case 'garage.regeneratePin': case 'garageRegeneratePin': return regenerarPIN(Number(payload.row));
    case 'garage.changeCode': case 'garageChangeCode': return alterarCodigo(payload.codigo || payload.code || payload.novoCodigo, payload.adminEmail || payload.email || 'Admin');
    case 'garage.saveConfig': case 'garageSaveConfig': return guardarConfigs(payload.configs && typeof payload.configs === 'object' ? payload.configs : payload);
    default: return { ok:false, success:false, msg:'Ação POST Garagem inválida: ' + action };
  }
}
function getSheet(name) { const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); let sh = ss.getSheetByName(name); if (!sh) { sh = ss.insertSheet(name); initSheet(name, sh); } return sh; }
function initSheet(name, sheet) { const h = { CONDOMINOS:['ID','DataRegisto','Nome','Piso','Fracao','Email','Telemovel','Estado','PIN','PINAtivo','DataAprovacao','AprovadoPor','UltimoAcesso','Observacoes','TentativasFalhadas','BloqueadoAte'], CONSULTAS:['ID','DataHora','PINUsado','Nome','Piso','Fracao','Email','Motivo','MotivoOutro','CodigoMostrado','Resultado','UserAgent'], CONFIG:['Chave','Valor','Descricao'], LOG:['DataHora','Tipo','Acao','Email','Mensagem'] }; if (h[name]) { sheet.appendRow(h[name]); sheet.getRange(1,1,1,h[name].length).setFontWeight('bold'); } }
function setupApp() { const c = getSheet('CONFIG'); getSheet('CONDOMINOS'); getSheet('CONSULTAS'); getSheet('LOG'); const defaults = [['NOME_CONDOMINIO','Domingos da Cunha 4','Nome do condomínio'],['CODIGO_CADEADO','0000','Código atual do cadeado'],['ADMIN_EMAIL',CONFIG.ADMIN_EMAIL || 'admin@email.com','Email do administrador'],['ADMIN_PIN','123456','PIN do administrador'],['TEMPO_VISIVEL_SEGUNDOS','15','Segundos que o código fica visível'],['PIN_DIGITOS','6','Número de dígitos do PIN'],['MAX_TENTATIVAS_PIN','5','Tentativas antes de bloquear'],['BLOQUEIO_MINUTOS','15','Minutos de bloqueio'],['TEXTO_AVISO','Após utilização, confirme que a caixa fica corretamente fechada.','Aviso mostrado com o código']]; const keys = c.getDataRange().getValues().map(r => r[0]); defaults.forEach(r => { if (keys.indexOf(r[0]) === -1) c.appendRow(r); }); return 'Setup concluído'; }
function getConfig(key) { const data = getSheet('CONFIG').getDataRange().getValues(); for (let i=1;i<data.length;i++) if (data[i][0] === key) return String(data[i][1]); return null; }
function setConfig(key, value) { const sh = getSheet('CONFIG'); const data = sh.getDataRange().getValues(); for (let i=1;i<data.length;i++) if (data[i][0] === key) { sh.getRange(i+1,2).setValue(value); return; } sh.appendRow([key,value,'']); }
function addLog(tipo, acao, email, msg) { getSheet('LOG').appendRow([new Date(), tipo, acao, email || '', msg || '']); }
function generateId(prefix) { return (prefix || 'ID') + '_' + Date.now() + '_' + Math.floor(Math.random()*1000); }
function generateUniquePIN() { const data = getSheet('CONDOMINOS').getDataRange().getValues(); const pins = data.slice(1).map(r => String(r[8])); let pin, a = 0; do { pin = String(Math.floor(100000 + Math.random()*900000)); a++; if (a > 100) throw new Error('Não foi possível gerar PIN único.'); } while (pins.indexOf(pin) !== -1 || pin === getConfig('ADMIN_PIN')); return pin; }
function getCondominoByPIN(pin) { const sh = getSheet('CONDOMINOS'); const data = sh.getDataRange().getValues(); const headers = data[0]; for (let i=1;i<data.length;i++) if (String(data[i][8]) === String(pin)) { const o = {}; headers.forEach((h,idx) => o[h] = data[i][idx]); o._row = i+1; return o; } return null; }
function updateCondominoRow(row, updates) { const sh = getSheet('CONDOMINOS'); const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; Object.keys(updates).forEach(k => { const c = headers.indexOf(k); if (c !== -1) sh.getRange(row,c+1).setValue(updates[k]); }); }
function getPublicConfig() { return { nomeCondominio:getConfig('NOME_CONDOMINIO') || 'Condomínio', tempoVisivel:parseInt(getConfig('TEMPO_VISIVEL_SEGUNDOS')) || 15, adminEmail:getConfig('ADMIN_EMAIL') || '', structure:getGarageStructure_().floors }; }
function getGarageStructure_() { const labels = {'10':'10.º andar','9':'9.º andar','8':'8.º andar','7':'7.º andar','6':'6.º andar','5':'5.º andar','4':'4.º andar','3':'3.º andar','2':'2.º andar','-1':'Garagem p-1','-2':'Garagem p-2','-3':'Garagem p-3'}; return { ok:true, success:true, floors:Object.keys(GARAGE_RESIDENTIAL_STRUCTURE).map(f => ({ piso:f, label:labels[f] || f, fraccoes:GARAGE_RESIDENTIAL_STRUCTURE[f] })).sort((a,b) => Number(b.piso) - Number(a.piso)) }; }
function normalizeGarageFloor_(v) { return String(v || '').trim().toUpperCase().replace(/º|°/g,'').replace(/ANDAR|PISO/g,'').replace(/\s+/g,''); }
function normalizeGarageFraction_(v) { return String(v || '').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/FRACCAO|FRACAO|FRAÇÃO/g,'').replace(/\s+/g,''); }
function validateGarageResident_(piso, fracao) { const f = normalizeGarageFloor_(piso); const frac = normalizeGarageFraction_(fracao); const allowed = GARAGE_RESIDENTIAL_STRUCTURE[f]; if (!allowed) throw new Error('Piso/zona inválido para registo.'); const norm = allowed.map(normalizeGarageFraction_); if (norm.indexOf(frac) === -1) throw new Error('Fração/garagem inválida para ' + f + '.'); return { piso:f, fracao:frac }; }
function getConfiguredAdminEmail_() { const e = getConfig('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL || ''; return isRealEmail_(e) ? e : ''; }
function registarCondomino(d) { try { if (!d || !d.nome) throw new Error('Nome obrigatório.'); if (!d.email) throw new Error('Email obrigatório.'); const v = validateGarageResident_(d.piso, d.fracao); d.piso = v.piso; d.fracao = v.fracao; const sh = getSheet('CONDOMINOS'); const data = sh.getDataRange().getValues(); const email = String(d.email).toLowerCase().trim(); for (let i=1;i<data.length;i++) if (String(data[i][5]).toLowerCase().trim() === email) { const estado = String(data[i][7] || '').toUpperCase(); if (estado === 'APROVADO') return { ok:false, success:false, tipo:'email_existente', recoveryAllowed:true, msg:'Este email já tem acesso aprovado. Pode recuperar o PIN.' }; if (estado === 'PENDENTE') return { ok:false, success:false, tipo:'pendente', msg:'Já existe um pedido pendente com este email.' }; if (estado === 'BLOQUEADO') return { ok:false, success:false, tipo:'bloqueado', msg:'Este email está bloqueado. Contacte a administração.' }; if (estado === 'REJEITADO') return { ok:false, success:false, tipo:'rejeitado', msg:'Este email já teve um pedido não aprovado. Contacte a administração.' }; }
    sh.appendRow([generateId('COND'), new Date(), d.nome, d.piso, d.fracao, d.email, d.telemovel || '', 'PENDENTE', '', false, '', '', '', '', 0, '']); addLog('INFO','REGISTO',d.email,'Novo pedido: ' + d.nome); emailRegisto(d.email, d.nome); emailAdminNovoPedido(d); return { ok:true, success:true }; } catch (err) { return { ok:false, success:false, msg:'Erro ao registar: ' + err.message, message:'Erro ao registar: ' + err.message }; } }
function loginComPIN(pin) { try { const c = getCondominoByPIN(pin); if (!c) return { ok:false, tipo:'invalido', msg:'PIN inválido.' }; if (c.Estado === 'PENDENTE') return { ok:false, tipo:'pendente', msg:'O seu pedido ainda está pendente.' }; if (c.Estado === 'REJEITADO') return { ok:false, tipo:'rejeitado', msg:'Pedido rejeitado.' }; if (c.Estado === 'BLOQUEADO') return { ok:false, tipo:'bloqueado', msg:'Acesso bloqueado.' }; if (c.Estado === 'APROVADO') { updateCondominoRow(c._row, { UltimoAcesso:new Date(), TentativasFalhadas:0, BloqueadoAte:'' }); return { ok:true, success:true, nome:c.Nome, piso:c.Piso, fracao:c.Fracao }; } return { ok:false, tipo:'invalido', msg:'Estado desconhecido.' }; } catch (err) { return { ok:false, tipo:'erro', msg:'Erro interno.' }; } }
function registarTentativaFalhada(pin) {}
function reenviarPIN(email) { try { const data = getSheet('CONDOMINOS').getDataRange().getValues(); const target = String(email || '').toLowerCase().trim(); for (let i=1;i<data.length;i++) if (String(data[i][5]).toLowerCase().trim() === target) { const estado = String(data[i][7] || '').toUpperCase(); if (estado !== 'APROVADO' || !data[i][8]) return { ok:false, success:false, msg:'Não existe conta aprovada com este email.' }; emailPINRecuperacao(email, data[i][2], data[i][8]); return { ok:true, success:true, msg:'PIN enviado para o email registado.' }; } return { ok:false, success:false, msg:'Email não encontrado.' }; } catch (err) { return { ok:false, success:false, msg:'Erro ao reenviar PIN: ' + err.message }; } }
function obterCodigo(pin, motivo, motivoOutro, ua) { const c = getCondominoByPIN(pin); if (!c || c.Estado !== 'APROVADO') return { ok:false, msg:'Acesso inválido.' }; const codigo = getConfig('CODIGO_CADEADO'); getSheet('CONSULTAS').appendRow([generateId('CONS'), new Date(), pin, c.Nome, c.Piso, c.Fracao, c.Email, motivo, motivoOutro || '', codigo, 'SUCESSO', ua || '']); return { ok:true, success:true, codigo, aviso:getConfig('TEXTO_AVISO'), tempo:parseInt(getConfig('TEMPO_VISIVEL_SEGUNDOS')) || 15 }; }
function loginAdmin(email, pin) { return String(email || '').toLowerCase() === String(getConfig('ADMIN_EMAIL') || '').toLowerCase() && String(pin) === String(getConfig('ADMIN_PIN')) ? { ok:true, success:true } : { ok:false, msg:'Credenciais inválidas.' }; }
function getDashboardData() { const cond = getSheet('CONDOMINOS').getDataRange().getValues().slice(1); const cons = getSheet('CONSULTAS').getDataRange().getValues().slice(1); const hoje = new Date(); return { ok:true, pendentes:cond.filter(r => r[7] === 'PENDENTE').length, ativos:cond.filter(r => r[7] === 'APROVADO').length, consultasHoje:cons.filter(r => { const d = new Date(r[1]); return d.getDate() === hoje.getDate() && d.getMonth() === hoje.getMonth() && d.getFullYear() === hoje.getFullYear(); }).length, codigoAtual:getConfig('CODIGO_CADEADO') }; }
function getPedidosPendentes() { const data = getSheet('CONDOMINOS').getDataRange().getValues(); const out=[]; for(let i=1;i<data.length;i++) if(data[i][7] === 'PENDENTE') out.push({row:i+1,id:data[i][0],nome:data[i][2],piso:data[i][3],fracao:data[i][4],email:data[i][5],telemovel:data[i][6],dataRegisto:data[i][1] ? Utilities.formatDate(new Date(data[i][1]), CONFIG.TZ, 'dd/MM/yyyy HH:mm') : ''}); return out; }
function getConominiosAprovados() { const data = getSheet('CONDOMINOS').getDataRange().getValues(); const out=[]; for(let i=1;i<data.length;i++) if(data[i][7] === 'APROVADO') out.push({row:i+1,id:data[i][0],nome:data[i][2],piso:data[i][3],fracao:data[i][4],email:data[i][5],pin:data[i][8],ultimoAcesso:data[i][12] ? Utilities.formatDate(new Date(data[i][12]), CONFIG.TZ, 'dd/MM/yyyy HH:mm') : 'Nunca'}); return out; }
function aprovarCondomino(row, adminEmail) { const sh = getSheet('CONDOMINOS'); const r = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0]; const pin = generateUniquePIN(); sh.getRange(row,8).setValue('APROVADO'); sh.getRange(row,9).setValue(pin); sh.getRange(row,10).setValue(true); sh.getRange(row,11).setValue(new Date()); sh.getRange(row,12).setValue(adminEmail || 'Admin'); emailAprovacao(r[5], r[2], pin); return { ok:true, success:true }; }
function rejeitarCondomino(row) { const sh = getSheet('CONDOMINOS'); const r = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0]; sh.getRange(row,8).setValue('REJEITADO'); emailRejeicao(r[5], r[2]); return { ok:true, success:true }; }
function bloquearCondomino(row) { getSheet('CONDOMINOS').getRange(row,8).setValue('BLOQUEADO'); return { ok:true, success:true }; }
function desbloquearCondomino(row) { getSheet('CONDOMINOS').getRange(row,8).setValue('APROVADO'); return { ok:true, success:true }; }
function regenerarPIN(row) { const sh=getSheet('CONDOMINOS'); const r=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0]; const pin=generateUniquePIN(); sh.getRange(row,9).setValue(pin); emailPINRecuperacao(r[5], r[2], pin); return { ok:true, success:true }; }
function alterarCodigo(codigo) { if (!/^\d+$/.test(String(codigo)) || String(codigo).length < 4) return { ok:false, msg:'Código inválido.' }; setConfig('CODIGO_CADEADO', codigo); return { ok:true, success:true }; }
function getHistoricoConsultas(limite) { return getSheet('CONSULTAS').getDataRange().getValues().slice(1).reverse().slice(0, limite || 50).map(r => ({ dataHora:r[1] ? Utilities.formatDate(new Date(r[1]), CONFIG.TZ, 'dd/MM/yyyy HH:mm') : '', pin:r[2], nome:r[3], piso:r[4], fracao:r[5], email:r[6], motivo:r[7], motivoOutro:r[8], codigo:r[9], resultado:r[10] })); }
function getAdminConfig() { return { nomeCondominio:getConfig('NOME_CONDOMINIO'), adminEmail:getConfig('ADMIN_EMAIL'), codigoAtual:getConfig('CODIGO_CADEADO'), tempoVisivel:getConfig('TEMPO_VISIVEL_SEGUNDOS'), maxTentativas:getConfig('MAX_TENTATIVAS_PIN'), bloqueioMinutos:getConfig('BLOQUEIO_MINUTOS'), textoAviso:getConfig('TEXTO_AVISO') }; }
function guardarConfigs(configs) { Object.keys(configs).forEach(k => { if (configs[k] !== '') setConfig(k, configs[k]); }); return { ok:true, success:true }; }
function emailRegisto(email,nome) { safeSendEmail_(email, 'Pedido de acesso recebido — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nRecebemos o seu pedido de acesso. Após aprovação receberá o PIN por email.', { replyTo:getConfiguredAdminEmail_() }); }
function emailAdminNovoPedido(d) { safeSendEmail_(getConfig('ADMIN_EMAIL'), 'Novo pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Novo pedido:\n\nNome: ' + d.nome + '\nPiso: ' + d.piso + '\nFração/Garagem: ' + d.fracao + '\nEmail: ' + d.email + '\nTelemóvel: ' + (d.telemovel || '—'), { replyTo:getConfiguredAdminEmail_() }); }
function emailAprovacao(email,nome,pin) { safeSendEmail_(email, 'Acesso aprovado — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin + '\n\nGuarde-o num local seguro.', { replyTo:getConfiguredAdminEmail_() }); }
function emailRejeicao(email,nome) { safeSendEmail_(email, 'Pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu pedido não foi aprovado. Contacte a administração.', { replyTo:getConfiguredAdminEmail_() }); }
function emailPINRecuperacao(email,nome,pin) { safeSendEmail_(email, 'Recuperação de PIN — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin, { replyTo:getConfiguredAdminEmail_() }); }
