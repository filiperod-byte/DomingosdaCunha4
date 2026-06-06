/*************************************************
 * DOMINGOS DA CUNHA 4 - BACKEND UNIFICADO
 * Google Apps Script Web App
 *
 * Funcionalidades:
 * Módulos:
 * - Extintores: reportes, estado, backoffice, fecho com foto e QR codes
 * - Garagem/Cadeado: registo de condóminos, aprovação, PIN, consulta do código e histórico
 *
 * IMPORTANTE:
 * - Existe apenas um doGet(e) e um doPost(e).
 * - As chamadas de GitHub Pages usam fetch/FormData para este Web App.
 * - Mantém compatibilidade com as ações antigas dos extintores: status, report, etc.
 *************************************************/

/*************************************************
 * CONFIGURAÇÃO
 *************************************************/
const CONFIG = {
  APP_NAME: 'Extintores Report',
  TZ: 'Europe/Lisbon',

  SPREADSHEET_ID: '16vz1ZIKkCI7NfAf2ZSVCNLEafBn4Svtu7UaRe0WdpBU',
  ROOT_FOLDER_ID: '1Ppu0Tk5zLrOWBQGwJtH_GE2kpSanWyb9',

  // Opcional: coloca aqui um email fixo para receber notificações.
  // Se deixares vazio, o script tenta usar o email efetivo do owner.
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
    'TIMESTAMP',
    'EVENTO',
    'OCCURRENCE_ID',
    'FLOOR',
    'FLOOR_LABEL',
    'POINT',
    'LOCATION',
    'REPORTED_BY',
    'REASON',
    'NOTES',
    'PHOTO_FILE_ID',
    'PHOTO_FILE_URL',
    'PHOTO_FILE_NAME',
    'STATUS_BEFORE',
    'STATUS_AFTER',
    'SOURCE',
    'CLIENT_TS',
    'EXTRA_JSON'
  ],

  ESTADO: [
    'LAST_UPDATED',
    'STATUS',
    'OCCURRENCE_ID',
    'FLOOR',
    'FLOOR_LABEL',
    'POINT',
    'LOCATION',
    'REPORTED_AT',
    'REPORTED_BY',
    'REASON',
    'NOTES',
    'PHOTO_FILE_ID',
    'PHOTO_FILE_URL',
    'PHOTO_FILE_NAME',
    'CLOSED_AT',
    'CLOSE_NOTES',
    'CLOSE_PHOTO_FILE_ID',
    'CLOSE_PHOTO_FILE_URL',
    'CLOSE_PHOTO_FILE_NAME',
    'LAST_EVENT'
  ],

  FECHOS: [
    'TIMESTAMP',
    'OCCURRENCE_ID',
    'FLOOR',
    'FLOOR_LABEL',
    'POINT',
    'LOCATION',
    'REPORTED_AT',
    'REPORTED_BY',
    'REASON',
    'OPEN_NOTES',
    'CLOSE_NOTES',
    'CLOSE_PHOTO_FILE_ID',
    'CLOSE_PHOTO_FILE_URL',
    'CLOSE_PHOTO_FILE_NAME',
    'SOURCE',
    'CLIENT_TS'
  ]
};

const OPEN_STATUSES = ['ABERTA', 'OPEN', 'ALERTA', 'REPORTADO'];

/*************************************************
 * ENTRYPOINTS WEB APP
 *************************************************/
function doGet(e) {
  return routeRequest_('GET', e);
}

function doPost(e) {
  return routeRequest_('POST', e);
}

/*************************************************
 * ROUTER
 *************************************************/
function routeRequest_(method, e) {
  try {
    setupIfNeeded_();

    if (method === 'GET') {
      const action = getActionFromGet_(e) || 'status';

      switch (action) {
        case 'status':
          return jsonResponse_(handleGetStatus_());

        case 'pinStatus':
          return jsonResponse_(handleGetPinStatus_());

        case 'openOccurrences':
          return jsonResponse_(handleGetOpenOccurrences_());

        case 'health':
          return jsonResponse_(handleGetHealth_());

        case 'garage.publicConfig':
        case 'garagePublicConfig':
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
          return jsonResponse_({
            success: false,
            message: 'Ação GET inválida: ' + action
          });
      }
    }

    const payload = parseRequestBody_(e);
    const action = String(payload.action || '').trim();

    switch (action) {
      case 'report':
        return jsonResponse_(handlePostReport_(payload));

      case 'setPin':
        return jsonResponse_(handlePostSetPin_(payload));

      case 'validatePin':
        return jsonResponse_(handlePostValidatePin_(payload));

      case 'resetPin':
        return jsonResponse_(handlePostResetPin_(payload));

      case 'closeOccurrence':
        return jsonResponse_(handlePostCloseOccurrence_(payload));

      case 'garage.register':
      case 'garageRegister':
      case 'garage.loginPin':
      case 'garageLoginPin':
      case 'garage.failedAttempt':
      case 'garageFailedAttempt':
      case 'garage.resendPin':
      case 'garageResendPin':
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
        return jsonResponse_({
          success: false,
          message: 'Ação POST inválida: ' + action
        });
    }
  } catch (err) {
    return jsonResponse_({
      success: false,
      message: err && err.message ? err.message : 'Erro interno no backend'
    });
  }
}

/*************************************************
 * GET ACTIONS
 *************************************************/
function handleGetStatus_() {
  const openRows = getOpenStateRows_();

  const reported = openRows.map(function (row) {
    return {
      floor: toInt_(row.FLOOR),
      point: String(row.POINT || '').trim()
    };
  });

  return {
    success: true,
    reported: reported,
    totalOpen: openRows.length,
    updatedAt: isoNow_()
  };
}

function handleGetPinStatus_() {
  return {
    success: true,
    pinConfigured: isPinConfigured_(),
    updatedAt: getScriptProperties_().getProperty(CONFIG.PROPS.PIN_UPDATED_AT) || ''
  };
}

function handleGetOpenOccurrences_() {
  const openRows = getOpenStateRows_()
    .sort(function (a, b) {
      return String(b.REPORTED_AT || '').localeCompare(String(a.REPORTED_AT || ''));
    })
    .map(function (row) {
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
        photoUrl: String(row.PHOTO_FILE_URL || '')
      };
    });

  return {
    success: true,
    occurrences: openRows
  };
}

function handleGetHealth_() {
  return {
    success: true,
    appName: CONFIG.APP_NAME,
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    rootFolderId: CONFIG.ROOT_FOLDER_ID,
    adminEmail: getAdminEmail_({}),
    pinConfigured: isPinConfigured_(),
    garageConfigReady: !!getConfig('NOME_CONDOMINIO'),
    now: isoNow_()
  };
}


/*************************************************
 * GARAGEM - ROTAS API
 *************************************************/
function handleGarageGet_(action) {
  garageSetupIfNeeded_();

  switch (action) {
    case 'garage.publicConfig':
    case 'garagePublicConfig':
      return getPublicConfig();

    case 'garage.dashboard':
    case 'garageDashboard':
      return getDashboardData();

    case 'garage.pending':
    case 'garagePending':
      return getPedidosPendentes();

    case 'garage.approved':
    case 'garageApproved':
      return getConominiosAprovados();

    case 'garage.history':
    case 'garageHistory':
      return getHistoricoConsultas(50);

    case 'garage.adminConfig':
    case 'garageAdminConfig':
      return getAdminConfig();

    default:
      return {
        ok: false,
        success: false,
        msg: 'Ação GET Garagem inválida: ' + action,
        message: 'Ação GET Garagem inválida: ' + action
      };
  }
}

function handleGaragePost_(action, payload) {
  garageSetupIfNeeded_();

  switch (action) {
    case 'garage.register':
    case 'garageRegister':
      return registarCondomino({
        nome: payload.nome || payload.name,
        piso: payload.piso || payload.floor,
        fracao: payload.fracao || payload.fraction,
        email: payload.email,
        telemovel: payload.telemovel || payload.telefone || payload.phone
      });

    case 'garage.loginPin':
    case 'garageLoginPin':
      return loginComPIN(payload.pin, payload.userAgent || payload.ua || '');

    case 'garage.failedAttempt':
    case 'garageFailedAttempt':
      registarTentativaFalhada(payload.pin);
      return { ok: true, success: true };

    case 'garage.resendPin':
    case 'garageResendPin':
      return reenviarPIN(payload.email);

    case 'garage.getCode':
    case 'garageGetCode':
      return obterCodigo(payload.pin, payload.motivo || payload.reason, payload.motivoOutro || payload.reasonOther || '', payload.userAgent || payload.ua || '');

    case 'garage.loginAdmin':
    case 'garageLoginAdmin':
      return loginAdmin(payload.email, payload.pin);

    case 'garage.approve':
    case 'garageApprove':
      return aprovarCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');

    case 'garage.reject':
    case 'garageReject':
      return rejeitarCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');

    case 'garage.block':
    case 'garageBlock':
      return bloquearCondomino(Number(payload.row), payload.adminEmail || payload.email || 'Admin');

    case 'garage.unblock':
    case 'garageUnblock':
      return desbloquearCondomino(Number(payload.row));

    case 'garage.regeneratePin':
    case 'garageRegeneratePin':
      return regenerarPIN(Number(payload.row));

    case 'garage.changeCode':
    case 'garageChangeCode':
      return alterarCodigo(payload.codigo || payload.code || payload.novoCodigo, payload.adminEmail || payload.email || 'Admin');

    case 'garage.saveConfig':
    case 'garageSaveConfig':
      return guardarConfigs(payload.configs && typeof payload.configs === 'object' ? payload.configs : payload);

    default:
      return {
        ok: false,
        success: false,
        msg: 'Ação POST Garagem inválida: ' + action,
        message: 'Ação POST Garagem inválida: ' + action
      };
  }
}

function garageSetupIfNeeded_() {
  setupApp();
}

/*************************************************
 * POST ACTIONS
 *************************************************/
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

    if (isNaN(floor)) {
      throw new Error('FLOOR inválido.');
    }
    if (!point) {
      throw new Error('POINT é obrigatório.');
    }
    if (!reportedBy) {
      throw new Error('NAME é obrigatório.');
    }
    if (!reason) {
      throw new Error('REASON é obrigatório.');
    }

    const floorLabel = formatFloorLabel_(floor);
    const nowIso = isoNow_();

    const existingState = findStateRowByPoint_(floor, point);
    const alreadyOpen = !!(existingState && isOpenStatus_(existingState.STATUS));
    const occurrenceId = alreadyOpen
      ? String(existingState.OCCURRENCE_ID || '')
      : generateOccurrenceId_();

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
      EVENTO: alreadyOpen ? 'REPORT_ADICIONAL' : 'REPORT',
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
      STATUS_AFTER: 'ABERTA',
      SOURCE: source,
      CLIENT_TS: clientTs,
      EXTRA_JSON: JSON.stringify({
        alreadyOpen: alreadyOpen
      })
    });

    upsertStateRow_(floor, point, {
      LAST_UPDATED: nowIso,
      STATUS: 'ABERTA',
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
      LAST_EVENT: alreadyOpen ? 'REPORT_ADICIONAL' : 'REPORT'
    });

    sendReportEmail_({
      adminEmail: getAdminEmail_(payload),
      occurrenceId: occurrenceId,
      floor: floor,
      floorLabel: floorLabel,
      point: point,
      location: location,
      reportedBy: reportedBy,
      reason: reason,
      notes: notes,
      photoUrl: photo.fileUrl,
      alreadyOpen: alreadyOpen
    });

    return {
      success: true,
      occurrenceId: occurrenceId,
      alreadyOpen: alreadyOpen,
      floor: floor,
      point: point
    };
  });
}

function handlePostSetPin_(payload) {
  return withScriptLock_(function () {
    const pin = normalizePin_(payload.pin);

    if (!pin) {
      throw new Error('PIN obrigatório.');
    }
    if (pin.length < CONFIG.PIN_MIN_LENGTH) {
      throw new Error('O PIN deve ter pelo menos ' + CONFIG.PIN_MIN_LENGTH + ' dígitos.');
    }
    if (!/^\d+$/.test(pin)) {
      throw new Error('O PIN deve conter apenas dígitos.');
    }
    if (isPinConfigured_()) {
      throw new Error('O PIN já está definido. Usa a recuperação/reset se precisares de o redefinir.');
    }

    const salt = Utilities.getUuid();
    const hash = hashPin_(pin, salt);

    getScriptProperties_().setProperties({
      [CONFIG.PROPS.PIN_HASH]: hash,
      [CONFIG.PROPS.PIN_SALT]: salt,
      [CONFIG.PROPS.PIN_UPDATED_AT]: isoNow_()
    }, true);

    return {
      success: true,
      message: 'PIN definido com sucesso.'
    };
  });
}

function handlePostValidatePin_(payload) {
  const pin = normalizePin_(payload.pin);

  if (!isPinConfigured_()) {
    return {
      success: true,
      valid: false,
      message: 'Ainda não existe PIN definido.'
    };
  }

  if (!pin) {
    return {
      success: true,
      valid: false,
      message: 'PIN vazio.'
    };
  }

  const props = getScriptProperties_();
  const salt = props.getProperty(CONFIG.PROPS.PIN_SALT) || '';
  const storedHash = props.getProperty(CONFIG.PROPS.PIN_HASH) || '';
  const incomingHash = hashPin_(pin, salt);

  return {
    success: true,
    valid: incomingHash === storedHash
  };
}

function handlePostResetPin_(payload) {
  return withScriptLock_(function () {
    const adminEmail = getAdminEmail_(payload);

    clearPin_();

    sendPinResetEmail_({
      adminEmail: adminEmail
    });

    return {
      success: true,
      message: 'PIN resetado. Na próxima utilização poderá ser definido novo PIN.'
    };
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

    let stateRow = null;

    if (occurrenceId) {
      stateRow = findOpenStateRowByOccurrenceId_(occurrenceId);
    }

    if (!stateRow && !isNaN(floor) && point) {
      stateRow = findOpenStateRowByPoint_(floor, point);
    }

    if (!stateRow) {
      throw new Error('Ocorrência aberta não encontrada.');
    }

    const closePhoto = saveIncomingPhoto_({
      base64: payload.closePhotoBase64,
      dataUrl: payload.closePhotoDataUrl,
      mimeType: payload.closePhotoType,
      fileName: payload.closePhotoName,
      folderName: CONFIG.SUBFOLDERS.CLOSES,
      occurrenceId: String(stateRow.OCCURRENCE_ID || ''),
      prefix: 'close'
    });

    if (!closePhoto.fileId) {
      throw new Error('A fotografia de fecho é obrigatória.');
    }

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
      TIMESTAMP: nowIso,
      EVENTO: 'FECHO',
      OCCURRENCE_ID: String(stateRow.OCCURRENCE_ID || ''),
      FLOOR: toInt_(stateRow.FLOOR),
      FLOOR_LABEL: String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)),
      POINT: String(stateRow.POINT || ''),
      LOCATION: String(stateRow.LOCATION || ''),
      REPORTED_BY: '',
      REASON: 'FECHO',
      NOTES: closeNotes,
      PHOTO_FILE_ID: closePhoto.fileId,
      PHOTO_FILE_URL: closePhoto.fileUrl,
      PHOTO_FILE_NAME: closePhoto.fileName,
      STATUS_BEFORE: String(stateRow.STATUS || ''),
      STATUS_AFTER: 'OK',
      SOURCE: source,
      CLIENT_TS: clientTs,
      EXTRA_JSON: JSON.stringify({
        closedOccurrenceId: String(stateRow.OCCURRENCE_ID || '')
      })
    });

    updateObjectRow_(CONFIG.SHEETS.ESTADO, stateRow._rowIndex, {
      LAST_UPDATED: nowIso,
      STATUS: 'OK',
      CLOSED_AT: nowIso,
      CLOSE_NOTES: closeNotes,
      CLOSE_PHOTO_FILE_ID: closePhoto.fileId,
      CLOSE_PHOTO_FILE_URL: closePhoto.fileUrl,
      CLOSE_PHOTO_FILE_NAME: closePhoto.fileName,
      LAST_EVENT: 'FECHO'
    });

    sendCloseEmail_({
      adminEmail: getAdminEmail_(payload),
      occurrenceId: String(stateRow.OCCURRENCE_ID || ''),
      floor: toInt_(stateRow.FLOOR),
      floorLabel: String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)),
      point: String(stateRow.POINT || ''),
      location: String(stateRow.LOCATION || ''),
      reason: String(stateRow.REASON || ''),
      reportedBy: String(stateRow.REPORTED_BY || ''),
      closeNotes: closeNotes,
      closePhotoUrl: closePhoto.fileUrl
    });

    return {
      success: true,
      occurrenceId: String(stateRow.OCCURRENCE_ID || ''),
      floor: toInt_(stateRow.FLOOR),
      point: String(stateRow.POINT || '')
    };
  });
}

/*************************************************
 * SETUP / ESTRUTURA
 *************************************************/
function setupIfNeeded_() {
  const ss = getSpreadsheet_();

  ensureSheetStructure_(ss, CONFIG.SHEETS.REGISTOS, REQUIRED_HEADERS.REGISTOS);
  ensureSheetStructure_(ss, CONFIG.SHEETS.ESTADO, REQUIRED_HEADERS.ESTADO);
  ensureSheetStructure_(ss, CONFIG.SHEETS.FECHOS, REQUIRED_HEADERS.FECHOS);
}

function ensureSheetStructure_(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastCol === 0 || lastRow === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentHeaders = getHeaders_(sheet);
  const missing = expectedHeaders.filter(function (h) {
    return currentHeaders.indexOf(h) === -1;
  });

  if (missing.length > 0) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

/*************************************************
 * SHEETS - LEITURA / ESCRITA
 *************************************************/
function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Folha não encontrada: ' + sheetName);
  }
  return sheet;
}

function getHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v || '').trim();
  });
}

function getSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;

  if (lastRow < 2 || lastCol < 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return values.map(function (row, idx) {
    const obj = { _rowIndex: idx + 2 };
    headers.forEach(function (h, i) {
      obj[h] = row[i];
    });
    return obj;
  });
}

function appendObjectRow_(sheetName, obj) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);

  const row = headers.map(function (h) {
    return obj.hasOwnProperty(h) ? obj[h] : '';
  });

  sheet.appendRow(row);
}

function updateObjectRow_(sheetName, rowIndex, patchObj) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const merged = {};

  headers.forEach(function (h, i) {
    merged[h] = current[i];
  });

  Object.keys(patchObj).forEach(function (key) {
    merged[key] = patchObj[key];
  });

  const output = headers.map(function (h) {
    return merged.hasOwnProperty(h) ? merged[h] : '';
  });

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([output]);
}

function upsertStateRow_(floor, point, dataObj) {
  const existing = findStateRowByPoint_(floor, point);

  if (existing) {
    updateObjectRow_(CONFIG.SHEETS.ESTADO, existing._rowIndex, dataObj);
  } else {
    appendObjectRow_(CONFIG.SHEETS.ESTADO, dataObj);
  }
}

function findStateRowByPoint_(floor, point) {
  const rows = getSheetObjects_(CONFIG.SHEETS.ESTADO);
  const floorNum = toInt_(floor);
  const pointNorm = normalizePoint_(point);

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (toInt_(row.FLOOR) === floorNum && normalizePoint_(row.POINT) === pointNorm) {
      return row;
    }
  }

  return null;
}

function findOpenStateRowByPoint_(floor, point) {
  const row = findStateRowByPoint_(floor, point);
  if (row && isOpenStatus_(row.STATUS)) {
    return row;
  }
  return null;
}

function findOpenStateRowByOccurrenceId_(occurrenceId) {
  const target = String(occurrenceId || '').trim();
  if (!target) return null;

  const rows = getSheetObjects_(CONFIG.SHEETS.ESTADO);

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (String(row.OCCURRENCE_ID || '').trim() === target && isOpenStatus_(row.STATUS)) {
      return row;
    }
  }

  return null;
}

function getOpenStateRows_() {
  return getSheetObjects_(CONFIG.SHEETS.ESTADO).filter(function (row) {
    return isOpenStatus_(row.STATUS);
  });
}

/*************************************************
 * PIN
 *************************************************/
function getScriptProperties_() {
  return PropertiesService.getScriptProperties();
}

function isPinConfigured_() {
  const props = getScriptProperties_();
  const hash = props.getProperty(CONFIG.PROPS.PIN_HASH);
  const salt = props.getProperty(CONFIG.PROPS.PIN_SALT);
  return !!(hash && salt);
}

function clearPin_() {
  const props = getScriptProperties_();
  props.deleteProperty(CONFIG.PROPS.PIN_HASH);
  props.deleteProperty(CONFIG.PROPS.PIN_SALT);
  props.deleteProperty(CONFIG.PROPS.PIN_UPDATED_AT);
}

function hashPin_(pin, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '::' + String(pin),
    Utilities.Charset.UTF_8
  );

  return bytesToHex_(digest);
}

function bytesToHex_(bytes) {
  return bytes.map(function (b) {
    const value = b < 0 ? b + 256 : b;
    const hex = value.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/*************************************************
 * DRIVE / FOTOS
 *************************************************/
function getRootFolder_() {
  return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
}

function getOrCreateSubfolder_(subfolderName) {
  const root = getRootFolder_();
  const existing = root.getFoldersByName(subfolderName);

  if (existing.hasNext()) {
    return existing.next();
  }

  return root.createFolder(subfolderName);
}

function saveIncomingPhoto_(opts) {
  const base64 = safeText_(opts.base64);
  const dataUrl = safeText_(opts.dataUrl);
  let mimeType = safeText_(opts.mimeType);
  let fileName = safeText_(opts.fileName);
  const folderName = safeText_(opts.folderName);
  const occurrenceId = safeText_(opts.occurrenceId);
  const prefix = safeText_(opts.prefix) || 'file';

  if (!base64 && !dataUrl) {
    return {
      fileId: '',
      fileUrl: '',
      fileName: ''
    };
  }

  let finalBase64 = base64;

  if (!finalBase64 && dataUrl.indexOf(',') > -1) {
    finalBase64 = dataUrl.split(',')[1];
  }

  if (!mimeType && dataUrl.indexOf(';base64,') > -1) {
    mimeType = dataUrl.substring(5, dataUrl.indexOf(';base64,'));
  }

  if (!finalBase64) {
    throw new Error('Foto inválida.');
  }

  if (!mimeType) {
    mimeType = 'image/jpeg';
  }

  if (!fileName) {
    fileName = buildDefaultPhotoName_(prefix, occurrenceId, mimeType);
  }

  const folder = getOrCreateSubfolder_(folderName);
  const bytes = Utilities.base64Decode(finalBase64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Não foi possível aplicar sharing à foto: ' + err);
  }

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: file.getName()
  };
}

function buildDefaultPhotoName_(prefix, occurrenceId, mimeType) {
  const ext = extensionFromMimeType_(mimeType);
  return [
    prefix || 'file',
    occurrenceId || Utilities.getUuid(),
    Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss')
  ].join('_') + '.' + ext;
}

function extensionFromMimeType_(mimeType) {
  const map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/heic': 'heic'
  };

  return map[mimeType] || 'jpg';
}

/*************************************************
 * EMAILS
 *************************************************/
function getAdminEmail_(payload) {
  const fromPayload = safeText_(payload && payload.adminEmail);
  if (isRealEmail_(fromPayload)) return fromPayload;

  const fixed = safeText_(CONFIG.ADMIN_EMAIL);
  if (isRealEmail_(fixed)) return fixed;

  const effective = safeText_(Session.getEffectiveUser().getEmail());
  if (isRealEmail_(effective)) return effective;

  const active = safeText_(Session.getActiveUser().getEmail());
  if (isRealEmail_(active)) return active;

  return '';
}

function isRealEmail_(email) {
  return !!email && email.indexOf('@') > -1 && email.indexOf('COLOCAR_') === -1;
}

function sendReportEmail_(data) {
  if (!data.adminEmail) return;

  const subject =
    '[' + CONFIG.APP_NAME + '] ' +
    (data.alreadyOpen ? 'Informação adicional' : 'Novo reporte') +
    ' - ' + data.floorLabel + ' / ' + data.point;

  const body = [
    'Foi registado um reporte de extintor.',
    '',
    'Ocorrência ID: ' + data.occurrenceId,
    'Piso: ' + data.floorLabel,
    'Ponto: ' + data.point,
    'Localização: ' + (data.location || '—'),
    'Reportado por: ' + data.reportedBy,
    'Motivo: ' + data.reason,
    'Observação: ' + (data.notes || '—'),
    'Foto: ' + (data.photoUrl || '—'),
    'Tipo: ' + (data.alreadyOpen ? 'Informação adicional' : 'Novo reporte'),
    '',
    'Data/Hora: ' + isoNow_()
  ].join('\n');

  safeSendEmail_(data.adminEmail, subject, body);
}

function sendCloseEmail_(data) {
  if (!data.adminEmail) return;

  const subject =
    '[' + CONFIG.APP_NAME + '] Ocorrência fechada - ' +
    data.floorLabel + ' / ' + data.point;

  const body = [
    'Uma ocorrência foi fechada.',
    '',
    'Ocorrência ID: ' + data.occurrenceId,
    'Piso: ' + data.floorLabel,
    'Ponto: ' + data.point,
    'Localização: ' + (data.location || '—'),
    'Reportado por: ' + (data.reportedBy || '—'),
    'Motivo original: ' + (data.reason || '—'),
    'Observação de fecho: ' + (data.closeNotes || '—'),
    'Foto de fecho: ' + (data.closePhotoUrl || '—'),
    '',
    'Data/Hora: ' + isoNow_()
  ].join('\n');

  safeSendEmail_(data.adminEmail, subject, body);
}

function sendPinResetEmail_(data) {
  if (!data.adminEmail) return;

  const subject = '[' + CONFIG.APP_NAME + '] Recuperação / reset do PIN';
  const body = [
    'Foi solicitado reset ao PIN do back office.',
    '',
    'O PIN anterior foi removido.',
    'Na próxima utilização será possível definir um novo PIN.',
    '',
    'Data/Hora: ' + isoNow_()
  ].join('\n');

  safeSendEmail_(data.adminEmail, subject, body);
}

function safeSendEmail_(recipient, subject, body) {
  try {
    if (!recipient) return;
    GmailApp.sendEmail(recipient, subject, body);
  } catch (err) {
    Logger.log('Falha no envio de email: ' + err);
  }
}

/*************************************************
 * HELPERS GERAIS
 *************************************************/
function parseRequestBody_(e) {
  const payload = {};

  if (e && e.postData && e.postData.contents) {
    try {
      const parsed = JSON.parse(e.postData.contents);
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (err) {
      // ignora e tenta parâmetros normais
    }
  }

  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (key) {
      payload[key] = e.parameter[key];
    });
  }

  return payload;
}

function getActionFromGet_(e) {
  if (!e || !e.parameter) return '';
  return String(e.parameter.action || '').trim();
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function generateOccurrenceId_() {
  return 'OCC-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMddHHmmss') + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function isoNow_() {
  return new Date().toISOString();
}

function safeText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function toInt_(value) {
  if (value === null || value === undefined || value === '') return NaN;
  return parseInt(String(value).trim(), 10);
}

function normalizePoint_(value) {
  return safeText_(value).toUpperCase();
}

function normalizePin_(value) {
  return safeText_(value).replace(/\s+/g, '');
}

function formatFloorLabel_(floor) {
  const n = toInt_(floor);

  if (isNaN(n)) return String(floor || '');
  if (n < 0) return 'p' + n;
  return String(n);
}

function isOpenStatus_(status) {
  const norm = safeText_(status).toUpperCase();
  return OPEN_STATUSES.indexOf(norm) > -1;
}

/*************************************************
 * FUNÇÕES MANUAIS ÚTEIS
 *************************************************/
function setupBackend_() {
  setupIfNeeded_();
  Logger.log('Backend preparado com sucesso.');
}

function resetPinManualmente_() {
  clearPin_();
  Logger.log('PIN removido com sucesso.');
}

function verOcorrenciasAbertas_() {
  Logger.log(JSON.stringify(handleGetOpenOccurrences_(), null, 2));
}

function verEstadoAtual_() {
  Logger.log(JSON.stringify(handleGetStatus_(), null, 2));
}

// ============================================================
// GARAGEM APP — Codigo.gs
// Google Apps Script + Sheets + HtmlService
// ============================================================

var SHEET_ID = CONFIG.SPREADSHEET_ID; // usado pelo módulo Garagem; não usar getActiveSpreadsheet em Web App

// ── Folhas ───────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(name, sheet);
  }
  return sheet;
}

function initSheet(name, sheet) {
  var headers = {
    CONDOMINOS: ['ID','DataRegisto','Nome','Piso','Fracao','Email','Telemovel','Estado','PIN','PINAtivo','DataAprovacao','AprovadoPor','UltimoAcesso','Observacoes','TentativasFalhadas','BloqueadoAte'],
    CONSULTAS:  ['ID','DataHora','PINUsado','Nome','Piso','Fracao','Email','Motivo','MotivoOutro','CodigoMostrado','Resultado','UserAgent'],
    CONFIG:     ['Chave','Valor','Descricao'],
    LOG:        ['DataHora','Tipo','Acao','Email','Mensagem']
  };
  if (headers[name]) {
    sheet.appendRow(headers[name]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
  }
}

// ── Setup inicial ────────────────────────────────────────────
function setupApp() {
  var configSheet = getSheet('CONFIG');
  getSheet('CONDOMINOS');
  getSheet('CONSULTAS');
  getSheet('LOG');

  var defaults = [
    ['NOME_CONDOMINIO',       'Condomínio Exemplo',           'Nome do condomínio'],
    ['CODIGO_CADEADO',        '000000',                       'Código atual do cadeado'],
    ['ADMIN_EMAIL',           'admin@email.com',              'Email do administrador'],
    ['ADMIN_PIN',             '123456',                       'PIN do administrador'],
    ['TEMPO_VISIVEL_SEGUNDOS','60',                           'Segundos que o código fica visível'],
    ['PIN_DIGITOS',           '6',                            'Número de dígitos do PIN'],
    ['MAX_TENTATIVAS_PIN',    '5',                            'Tentativas antes de bloquear'],
    ['BLOQUEIO_MINUTOS',      '15',                           'Minutos de bloqueio após tentativas falhadas'],
    ['TEXTO_AVISO',           'Após utilização, confirme que a caixa fica corretamente fechada.', 'Aviso mostrado com o código']
  ];

  var existingData = configSheet.getDataRange().getValues();
  var existingKeys = existingData.map(function(r){ return r[0]; });

  defaults.forEach(function(row) {
    if (existingKeys.indexOf(row[0]) === -1) {
      configSheet.appendRow(row);
    }
  });

  return 'Setup concluído com sucesso!';
}

// ── Config helpers ───────────────────────────────────────────
function getConfig(key) {
  var data = getSheet('CONFIG').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return String(data[i][1]);
  }
  return null;
}

function setConfig(key, value) {
  var sheet = getSheet('CONFIG');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

// ── Log ──────────────────────────────────────────────────────
function addLog(tipo, acao, email, mensagem) {
  getSheet('LOG').appendRow([
    new Date(),
    tipo,
    acao,
    email || '',
    mensagem || ''
  ]);
}

// ── ID único ─────────────────────────────────────────────────
function generateId(prefix) {
  return (prefix || 'ID') + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// ── PIN único ────────────────────────────────────────────────
function generateUniquePIN() {
  var sheet = getSheet('CONDOMINOS');
  var data = sheet.getDataRange().getValues();
  var existingPINs = data.slice(1).map(function(r){ return String(r[8]); });

  var pin, attempts = 0;
  do {
    pin = String(Math.floor(100000 + Math.random() * 900000));
    attempts++;
    if (attempts > 100) throw new Error('Não foi possível gerar PIN único.');
  } while (existingPINs.indexOf(pin) !== -1 ||
           pin === getConfig('ADMIN_PIN'));

  return pin;
}

// ── Condómino por PIN ────────────────────────────────────────
function getCondominoByPIN(pin) {
  var sheet = getSheet('CONDOMINOS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][8]) === String(pin)) {
      var obj = {};
      headers.forEach(function(h, idx){ obj[h] = data[i][idx]; });
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

function updateCondominoRow(row, updates) {
  var sheet = getSheet('CONDOMINOS');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Object.keys(updates).forEach(function(key) {
    var col = headers.indexOf(key);
    if (col !== -1) sheet.getRange(row, col + 1).setValue(updates[key]);
  });
}

// ============================================================
// ROTAS PÚBLICAS (chamadas pelo cliente via google.script.run)
// ============================================================

// ── Rotas Garagem ─────────────────────────────────────────────
// Em GitHub Pages não existe google.script.run.
// O acesso passa pelo router único em routeRequest_(), com ações garage.*.
// ── Obter configs públicos ───────────────────────────────────
function getPublicConfig() {
  return {
    nomeCondominio: getConfig('NOME_CONDOMINIO') || 'Condomínio',
    tempoVisivel:   parseInt(getConfig('TEMPO_VISIVEL_SEGUNDOS')) || 60,
    adminEmail:     getConfig('ADMIN_EMAIL') || ''
  };
}

// ── Registo de condómino ─────────────────────────────────────
function registarCondomino(dados) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var allData = sheet.getDataRange().getValues();

    // Verificar email duplicado
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][5]).toLowerCase() === String(dados.email).toLowerCase()) {
        var estado = allData[i][7];
        if (estado === 'PENDENTE') return { ok: false, msg: 'Já existe um pedido pendente com este email.' };
        if (estado === 'APROVADO') return { ok: false, msg: 'Este email já tem acesso aprovado.' };
        if (estado === 'BLOQUEADO') return { ok: false, msg: 'Este email encontra-se bloqueado. Contacte a administração.' };
      }
    }

    var id = generateId('COND');
    sheet.appendRow([
      id,
      new Date(),
      dados.nome,
      dados.piso,
      dados.fracao,
      dados.email,
      dados.telemovel || '',
      'PENDENTE',
      '',    // PIN (vazio até aprovação)
      false, // PINAtivo
      '',    // DataAprovacao
      '',    // AprovadoPor
      '',    // UltimoAcesso
      '',    // Observacoes
      0,     // TentativasFalhadas
      ''     // BloqueadoAte
    ]);

    addLog('INFO', 'REGISTO', dados.email, 'Novo pedido de registo: ' + dados.nome);

    // Email para o condómino
    emailRegisto(dados.email, dados.nome);
    // Email para o admin
    emailAdminNovoPedido(dados);

    return { ok: true };
  } catch (err) {
    addLog('ERRO', 'REGISTO', dados.email, err.message);
    return { ok: false, msg: 'Erro ao registar: ' + err.message };
  }
}

// ── Login com PIN ────────────────────────────────────────────
function loginComPIN(pin, userAgent) {
  try {
    var cond = getCondominoByPIN(pin);

    if (!cond) {
      addLog('ALERTA', 'LOGIN', '', 'Tentativa com PIN inválido: ' + pin);
      return { ok: false, tipo: 'invalido', msg: 'PIN inválido. Confirme o código e tente novamente.' };
    }

    // Verificar bloqueio temporário
    if (cond.BloqueadoAte) {
      var bloqueadoAte = new Date(cond.BloqueadoAte);
      if (bloqueadoAte > new Date()) {
        var minutos = Math.ceil((bloqueadoAte - new Date()) / 60000);
        return { ok: false, tipo: 'bloqueio_temp', msg: 'Demasiadas tentativas falhadas. Tente novamente em ' + minutos + ' minuto(s).' };
      } else {
        // Bloqueio expirou, resetar
        updateCondominoRow(cond._row, { TentativasFalhadas: 0, BloqueadoAte: '' });
      }
    }

    var estado = cond.Estado;

    if (estado === 'PENDENTE')  return { ok: false, tipo: 'pendente',  msg: 'O seu pedido ainda está pendente de aprovação.' };
    if (estado === 'REJEITADO') return { ok: false, tipo: 'rejeitado', msg: 'O seu pedido foi rejeitado pela administração.' };
    if (estado === 'BLOQUEADO') return { ok: false, tipo: 'bloqueado', msg: 'O seu acesso encontra-se bloqueado. Contacte a administração.' };

    if (estado === 'APROVADO') {
      updateCondominoRow(cond._row, {
        UltimoAcesso: new Date(),
        TentativasFalhadas: 0,
        BloqueadoAte: ''
      });
      addLog('INFO', 'LOGIN', cond.Email, 'Login bem-sucedido: ' + cond.Nome);
      return {
        ok: true,
        nome: cond.Nome,
        piso: cond.Piso,
        fracao: cond.Fracao
      };
    }

    return { ok: false, tipo: 'invalido', msg: 'Estado desconhecido. Contacte a administração.' };

  } catch (err) {
    addLog('ERRO', 'LOGIN', '', err.message);
    return { ok: false, tipo: 'erro', msg: 'Erro interno. Tente novamente.' };
  }
}

// ── Registar tentativa falhada ───────────────────────────────
function registarTentativaFalhada(pin) {
  var cond = getCondominoByPIN(pin);
  if (!cond) return;

  var max = parseInt(getConfig('MAX_TENTATIVAS_PIN')) || 5;
  var minutos = parseInt(getConfig('BLOQUEIO_MINUTOS')) || 15;
  var tentativas = parseInt(cond.TentativasFalhadas || 0) + 1;

  if (tentativas >= max) {
    var bloqueadoAte = new Date(new Date().getTime() + minutos * 60000);
    updateCondominoRow(cond._row, { TentativasFalhadas: tentativas, BloqueadoAte: bloqueadoAte });
    addLog('ALERTA', 'LOGIN', cond.Email, 'Bloqueio temporário após ' + tentativas + ' tentativas falhadas');
  } else {
    updateCondominoRow(cond._row, { TentativasFalhadas: tentativas });
  }
}

// ── Reenviar PIN ─────────────────────────────────────────────
function reenviarPIN(email) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][5]).toLowerCase() === String(email).toLowerCase()) {
        var estado = data[i][7];
        var pin    = data[i][8];
        var nome   = data[i][2];

        if (estado !== 'APROVADO') {
          return { ok: false, msg: 'Não existe conta aprovada com este email.' };
        }

        emailPINRecuperacao(email, nome, pin);
        addLog('INFO', 'LOGIN', email, 'PIN reenviado por pedido do condómino');
        return { ok: true };
      }
    }
    return { ok: false, msg: 'Email não encontrado.' };
  } catch (err) {
    return { ok: false, msg: 'Erro ao reenviar PIN.' };
  }
}

// ── Obter código do cadeado ──────────────────────────────────
function obterCodigo(pin, motivo, motivoOutro, userAgent) {
  try {
    var cond = getCondominoByPIN(pin);
    if (!cond || cond.Estado !== 'APROVADO') {
      return { ok: false, msg: 'Acesso inválido.' };
    }

    var codigo = getConfig('CODIGO_CADEADO');
    var aviso  = getConfig('TEXTO_AVISO');
    var tempo  = parseInt(getConfig('TEMPO_VISIVEL_SEGUNDOS')) || 60;

    // Registar consulta
    var id = generateId('CONS');
    getSheet('CONSULTAS').appendRow([
      id,
      new Date(),
      pin,
      cond.Nome,
      cond.Piso,
      cond.Fracao,
      cond.Email,
      motivo,
      motivoOutro || '',
      codigo,
      'SUCESSO',
      userAgent || ''
    ]);

    addLog('INFO', 'CONSULTA', cond.Email, 'Código consultado por ' + cond.Nome + ' — Motivo: ' + motivo);

    return {
      ok: true,
      codigo: codigo,
      aviso: aviso,
      tempo: tempo
    };
  } catch (err) {
    addLog('ERRO', 'CONSULTA', '', err.message);
    return { ok: false, msg: 'Erro ao obter código.' };
  }
}

// ============================================================
// ÁREA DE ADMINISTRAÇÃO
// ============================================================

function loginAdmin(email, pin) {
  var adminEmail = getConfig('ADMIN_EMAIL');
  var adminPIN   = getConfig('ADMIN_PIN');

  if (email.toLowerCase() === adminEmail.toLowerCase() && pin === adminPIN) {
    addLog('INFO', 'LOGIN', email, 'Login de administrador');
    return { ok: true };
  }
  addLog('ALERTA', 'LOGIN', email, 'Tentativa de login de admin falhada');
  return { ok: false, msg: 'Credenciais inválidas.' };
}

function getDashboardData() {
  try {
    var condSheet = getSheet('CONDOMINOS');
    var consSheet = getSheet('CONSULTAS');

    var condData = condSheet.getDataRange().getValues().slice(1);
    var consData = consSheet.getDataRange().getValues().slice(1);

    var pendentes = condData.filter(function(r){ return r[7] === 'PENDENTE'; }).length;
    var ativos    = condData.filter(function(r){ return r[7] === 'APROVADO'; }).length;

    var hoje = new Date();
    var consultasHoje = consData.filter(function(r){
      var d = new Date(r[1]);
      return d.getDate() === hoje.getDate() &&
             d.getMonth() === hoje.getMonth() &&
             d.getFullYear() === hoje.getFullYear();
    }).length;

    return {
      ok: true,
      pendentes: pendentes,
      ativos: ativos,
      consultasHoje: consultasHoje,
      codigoAtual: getConfig('CODIGO_CADEADO')
    };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function getPedidosPendentes() {
  var data = getSheet('CONDOMINOS').getDataRange().getValues();
  var headers = data[0];
  var pendentes = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'PENDENTE') {
      pendentes.push({
        row: i + 1,
        id:      data[i][0],
        nome:    data[i][2],
        piso:    data[i][3],
        fracao:  data[i][4],
        email:   data[i][5],
        telemovel: data[i][6],
        dataRegisto: data[i][1] ? Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : ''
      });
    }
  }
  return pendentes;
}

function getConominiosAprovados() {
  var data = getSheet('CONDOMINOS').getDataRange().getValues();
  var aprovados = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'APROVADO') {
      aprovados.push({
        row: i + 1,
        id:      data[i][0],
        nome:    data[i][2],
        piso:    data[i][3],
        fracao:  data[i][4],
        email:   data[i][5],
        pin:     data[i][8],
        dataAprovacao: data[i][10] ? Utilities.formatDate(new Date(data[i][10]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        ultimoAcesso:  data[i][12] ? Utilities.formatDate(new Date(data[i][12]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : 'Nunca'
      });
    }
  }
  return aprovados;
}

function aprovarCondomino(rowNum, adminEmail) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nome  = row[2];
    var email = row[5];
    var pin   = generateUniquePIN();

    sheet.getRange(rowNum, 8).setValue('APROVADO');
    sheet.getRange(rowNum, 9).setValue(pin);
    sheet.getRange(rowNum, 10).setValue(true);
    sheet.getRange(rowNum, 11).setValue(new Date());
    sheet.getRange(rowNum, 12).setValue(adminEmail || 'Admin');

    emailAprovacao(email, nome, pin);
    addLog('INFO', 'APROVACAO', email, 'Condómino aprovado: ' + nome + ' | PIN: ' + pin);

    return { ok: true };
  } catch (err) {
    addLog('ERRO', 'APROVACAO', '', err.message);
    return { ok: false, msg: err.message };
  }
}

function rejeitarCondomino(rowNum, adminEmail) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nome  = row[2];
    var email = row[5];

    sheet.getRange(rowNum, 8).setValue('REJEITADO');
    emailRejeicao(email, nome);
    addLog('INFO', 'REJEICAO', email, 'Condómino rejeitado: ' + nome);

    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function bloquearCondomino(rowNum, adminEmail) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nome  = row[2];
    var email = row[5];

    sheet.getRange(rowNum, 8).setValue('BLOQUEADO');
    addLog('INFO', 'BLOQUEIO', email, 'Condómino bloqueado: ' + nome);

    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function desbloquearCondomino(rowNum) {
  try {
    var sheet = getSheet('CONDOMINOS');
    sheet.getRange(rowNum, 8).setValue('APROVADO');
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function regenerarPIN(rowNum) {
  try {
    var sheet = getSheet('CONDOMINOS');
    var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nome  = row[2];
    var email = row[5];
    var pin   = generateUniquePIN();

    sheet.getRange(rowNum, 9).setValue(pin);
    emailPINRecuperacao(email, nome, pin);
    addLog('INFO', 'REGENERAR_PIN', email, 'PIN regenerado para: ' + nome);

    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function alterarCodigo(novoCodigo, adminEmail) {
  try {
    if (!/^\d+$/.test(novoCodigo) || novoCodigo.length < 4) {
      return { ok: false, msg: 'Código inválido. Deve ter pelo menos 4 dígitos.' };
    }
    setConfig('CODIGO_CADEADO', novoCodigo);
    addLog('INFO', 'ALTERACAO_CODIGO', adminEmail || 'Admin', 'Código do cadeado alterado');
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function getHistoricoConsultas(limite) {
  limite = limite || 50;
  var data = getSheet('CONSULTAS').getDataRange().getValues();
  var rows = data.slice(1).reverse().slice(0, limite);
  return rows.map(function(r) {
    return {
      dataHora:  r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
      pin:       r[2],
      nome:      r[3],
      piso:      r[4],
      fracao:    r[5],
      email:     r[6],
      motivo:    r[7],
      motivoOutro: r[8],
      codigo:    r[9],
      resultado: r[10]
    };
  });
}

function getAdminConfig() {
  return {
    nomeCondominio: getConfig('NOME_CONDOMINIO'),
    adminEmail:     getConfig('ADMIN_EMAIL'),
    codigoAtual:    getConfig('CODIGO_CADEADO'),
    tempoVisivel:   getConfig('TEMPO_VISIVEL_SEGUNDOS'),
    maxTentativas:  getConfig('MAX_TENTATIVAS_PIN'),
    bloqueioMinutos: getConfig('BLOQUEIO_MINUTOS'),
    textoAviso:     getConfig('TEXTO_AVISO')
  };
}

function guardarConfigs(configs) {
  try {
    Object.keys(configs).forEach(function(key) {
      setConfig(key, configs[key]);
    });
    addLog('INFO', 'ALTERACAO_CODIGO', 'Admin', 'Configurações atualizadas');
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// ============================================================
// EMAILS
// ============================================================

function emailRegisto(email, nome) {
  try {
    var nomeCondominio = getConfig('NOME_CONDOMINIO');
    MailApp.sendEmail({
      to: email,
      subject: 'Pedido de acesso recebido — ' + nomeCondominio,
      htmlBody: templateEmail(
        'Pedido recebido ✓',
        'Olá ' + nome + ',',
        'Recebemos o seu pedido de acesso à app do código da caixa do automático da luz da garagem.<br><br>' +
        'O pedido será validado pela administração do condomínio.<br><br>' +
        'Após aprovação receberá o seu PIN pessoal por email.',
        '#2563EB'
      )
    });
  } catch(e) { addLog('ERRO', 'REGISTO', email, 'Falha ao enviar email: ' + e.message); }
}

function emailAdminNovoPedido(dados) {
  try {
    var adminEmail = getConfig('ADMIN_EMAIL');
    var nomeCondominio = getConfig('NOME_CONDOMINIO');
    MailApp.sendEmail({
      to: adminEmail,
      subject: 'Novo pedido de acesso à app da garagem — ' + nomeCondominio,
      htmlBody: templateEmail(
        'Novo pedido de acesso',
        'Foi submetido um novo pedido de acesso:',
        '<b>Nome:</b> ' + dados.nome + '<br>' +
        '<b>Piso:</b> ' + dados.piso + '<br>' +
        '<b>Fração:</b> ' + dados.fracao + '<br>' +
        '<b>Email:</b> ' + dados.email + '<br>' +
        '<b>Telemóvel:</b> ' + (dados.telemovel || 'Não fornecido') + '<br><br>' +
        'Aceda à área de administração para aprovar ou rejeitar o pedido.',
        '#F59E0B'
      )
    });
  } catch(e) { addLog('ERRO', 'REGISTO', '', 'Falha ao notificar admin: ' + e.message); }
}

function emailAprovacao(email, nome, pin) {
  try {
    var nomeCondominio = getConfig('NOME_CONDOMINIO');
    MailApp.sendEmail({
      to: email,
      subject: 'Acesso aprovado — ' + nomeCondominio,
      htmlBody: templateEmail(
        'Acesso aprovado ✓',
        'Olá ' + nome + ',',
        'O seu acesso foi aprovado com sucesso!<br><br>' +
        'O seu PIN pessoal é:<br>' +
        '<div style="font-size:36px;font-weight:bold;letter-spacing:8px;color:#1E3A5F;text-align:center;padding:20px;background:#DBEAFE;border-radius:12px;margin:16px 0;">' + pin + '</div>' +
        'Use este PIN para entrar na app e consultar o código quando necessário.<br><br>' +
        '<b>Guarde este PIN num local seguro e não o partilhe.</b>',
        '#16A34A'
      )
    });
  } catch(e) { addLog('ERRO', 'APROVACAO', email, 'Falha ao enviar email de aprovação: ' + e.message); }
}

function emailRejeicao(email, nome) {
  try {
    var nomeCondominio = getConfig('NOME_CONDOMINIO');
    var adminEmail = getConfig('ADMIN_EMAIL');
    MailApp.sendEmail({
      to: email,
      subject: 'Pedido de acesso — ' + nomeCondominio,
      htmlBody: templateEmail(
        'Pedido analisado',
        'Olá ' + nome + ',',
        'Após análise do seu pedido, a administração do condomínio não pôde aprovar o acesso neste momento.<br><br>' +
        'Se considera que existe algum engano ou pretende obter mais informações, contacte diretamente a administração:<br>' +
        '<b>' + adminEmail + '</b>',
        '#DC2626'
      )
    });
  } catch(e) {}
}

function emailPINRecuperacao(email, nome, pin) {
  try {
    var nomeCondominio = getConfig('NOME_CONDOMINIO');
    MailApp.sendEmail({
      to: email,
      subject: 'Recuperação de PIN — ' + nomeCondominio,
      htmlBody: templateEmail(
        'O seu PIN pessoal',
        'Olá ' + nome + ',',
        'Conforme solicitado, aqui está o seu PIN de acesso:<br>' +
        '<div style="font-size:36px;font-weight:bold;letter-spacing:8px;color:#1E3A5F;text-align:center;padding:20px;background:#DBEAFE;border-radius:12px;margin:16px 0;">' + pin + '</div>' +
        '<b>Se não pediu este email, ignore-o.</b>',
        '#2563EB'
      )
    });
  } catch(e) {}
}

function templateEmail(titulo, saudacao, corpo, cor) {
  cor = cor || '#2563EB';
  return '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#F4F6F8;padding:24px;">' +
    '<div style="background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">' +
      '<div style="background:' + cor + ';padding:24px;text-align:center;">' +
        '<h2 style="color:#fff;margin:0;font-size:20px;">' + titulo + '</h2>' +
      '</div>' +
      '<div style="padding:24px;">' +
        '<p style="color:#111827;font-size:16px;margin:0 0 12px;">' + saudacao + '</p>' +
        '<p style="color:#374151;font-size:15px;line-height:1.6;margin:0;">' + corpo + '</p>' +
      '</div>' +
      '<div style="background:#F4F6F8;padding:16px;text-align:center;">' +
        '<p style="color:#9CA3AF;font-size:12px;margin:0;">' + (getConfig('NOME_CONDOMINIO') || 'Condomínio') + ' — App de Acesso à Garagem</p>' +
      '</div>' +
    '</div>' +
  '</div>';
}