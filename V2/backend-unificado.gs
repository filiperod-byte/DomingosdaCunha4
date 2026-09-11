// GERADO por node backend/build.cjs. Editar backend/src, não este ficheiro.
// Refatoração de compatibilidade: consultar backend/README.md antes de publicar.

// --- config.js ---
// Configuração e esquema atuais. Não colocar credenciais neste ficheiro.
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


// --- domain.js ---
// Regras puras: não dependem de Sheets, HTTP ou serviços Google.
function createCondominiumDomain_(residentialStructure, openStatuses, pendingStatuses) {
  const GARAGE_RESIDENTIAL_STRUCTURE = residentialStructure;
  const OPEN_STATUSES = openStatuses;
  const PENDING_STATUSES = pendingStatuses;
function safeText_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function toInt_(v) { if (v === null || v === undefined || v === '') return NaN; return parseInt(String(v).trim(), 10); }

function normalizePoint_(v) { return safeText_(v).toUpperCase(); }

function normalizePin_(v) { return safeText_(v).replace(/\s+/g, ''); }

function formatFloorLabel_(floor) { const n = toInt_(floor); if (isNaN(n)) return String(floor || ''); if (n < 0) return 'p' + n; return String(n); }

function isOpenStatus_(s) { return OPEN_STATUSES.indexOf(safeText_(s).toUpperCase()) > -1; }

function isPendingStatus_(s) { return PENDING_STATUSES.indexOf(safeText_(s).toUpperCase()) > -1; }

function isRealEmail_(email) { return !!email && String(email).indexOf('@') > -1 && String(email).indexOf('COLOCAR_') === -1; }

function getGarageStructure_() { const labels = {'10':'10.º andar','9':'9.º andar','8':'8.º andar','7':'7.º andar','6':'6.º andar','5':'5.º andar','4':'4.º andar','3':'3.º andar','2':'2.º andar','-1':'Garagem p-1','-2':'Garagem p-2','-3':'Garagem p-3'}; return { ok:true, success:true, floors:Object.keys(GARAGE_RESIDENTIAL_STRUCTURE).map(f => ({ piso:f, label:labels[f] || f, fraccoes:GARAGE_RESIDENTIAL_STRUCTURE[f] })).sort((a,b) => Number(b.piso) - Number(a.piso)) }; }

function normalizeGarageFloor_(v) { return String(v || '').trim().toUpperCase().replace(/º|°/g,'').replace(/ANDAR|PISO/g,'').replace(/\s+/g,''); }

function normalizeGarageFraction_(v) { return String(v || '').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/FRACCAO|FRACAO|FRAÇÃO/g,'').replace(/\s+/g,''); }

function validateGarageResident_(piso, fracao) { const f = normalizeGarageFloor_(piso); const frac = normalizeGarageFraction_(fracao); const allowed = GARAGE_RESIDENTIAL_STRUCTURE[f]; if (!allowed) throw new Error('Piso/zona inválido para registo.'); const norm = allowed.map(normalizeGarageFraction_); if (norm.indexOf(frac) === -1) throw new Error('Fração/garagem inválida para ' + f + '.'); return { piso:f, fracao:frac }; }
return { safeText_, toInt_, normalizePoint_, normalizePin_, formatFloorLabel_, isOpenStatus_, isPendingStatus_, isRealEmail_, getGarageStructure_, normalizeGarageFloor_, normalizeGarageFraction_, validateGarageResident_ };
}


// --- notifications.js ---
// Conteúdo das notificações separado do transporte de email.
function createNotificationService_(ports, CONFIG, domain) {
  const { isRealEmail_ } = domain;
  const getConfig = key => ports.settings.get(key);
  const isoNow_ = () => ports.clock.now().toISOString();
  const safeSendEmail_ = (to, subject, body, options) => ports.mail.send(to, subject, body, options);
function sendReportEmail_(d) { if (!d.adminEmail) return; const status = d.pendingValidation ? 'PENDENTE DE VALIDAÇÃO' : 'ABERTA'; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reporte extintor - ' + status, ['Reporte de extintor recebido.', '', 'Estado: ' + status, 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Localização: ' + (d.location || '—'), 'Reportado por: ' + d.reportedBy, 'Motivo: ' + d.reason, 'Observação: ' + (d.notes || '—'), 'Foto: ' + (d.photoUrl || '—'), '', 'Data/Hora: ' + isoNow_()].join('\n')); }

function sendCloseEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Ocorrência fechada', ['Ocorrência fechada.', '', 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Foto fecho: ' + (d.closePhotoUrl || '—'), 'Observação: ' + (d.closeNotes || '—')].join('\n')); }

function sendPinResetEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reset do PIN Backoffice', 'O PIN do backoffice foi resetado.'); }

function emailRegisto(email,nome) { safeSendEmail_(email, 'Pedido de acesso recebido — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nRecebemos o seu pedido de acesso. Após aprovação receberá o PIN por email.', { replyTo:getConfiguredAdminEmail_() }); }

function emailAdminNovoPedido(d) { safeSendEmail_(getConfig('ADMIN_EMAIL'), 'Novo pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Novo pedido:\n\nNome: ' + d.nome + '\nPiso: ' + d.piso + '\nFração/Garagem: ' + d.fracao + '\nEmail: ' + d.email + '\nTelemóvel: ' + (d.telemovel || '—'), { replyTo:getConfiguredAdminEmail_() }); }

function emailAprovacao(email,nome,pin) { safeSendEmail_(email, 'Acesso aprovado — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin + '\n\nGuarde-o num local seguro.', { replyTo:getConfiguredAdminEmail_() }); }

function emailRejeicao(email,nome) { safeSendEmail_(email, 'Pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu pedido não foi aprovado. Contacte a administração.', { replyTo:getConfiguredAdminEmail_() }); }

function emailPINRecuperacao(email,nome,pin) { safeSendEmail_(email, 'Recuperação de PIN — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin, { replyTo:getConfiguredAdminEmail_() }); }

function getConfiguredAdminEmail_() { const e = getConfig('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL || ''; return isRealEmail_(e) ? e : ''; }
return { sendReportEmail_, sendCloseEmail_, sendPinResetEmail_, emailRegisto, emailAdminNovoPedido, emailAprovacao, emailRejeicao, emailPINRecuperacao };
}


// --- occurrences.js ---
// Casos de uso dos extintores. Os registos legados mantêm o contrato atual.
function createOccurrenceService_(ports, CONFIG, domain, notifications, getAdminEmail_) {
  const { safeText_, toInt_, normalizePoint_, formatFloorLabel_, isOpenStatus_, isPendingStatus_ } = domain;
  const withScriptLock_ = cb => ports.lock.run(cb);
  const isoNow_ = () => ports.clock.now().toISOString();
  const generateOccurrenceId_ = () => ports.ids.occurrence();
  const saveIncomingPhoto_ = options => ports.files.save(options);
  const { sendReportEmail_, sendCloseEmail_ } = notifications;
function handleGetStatus_() {
  const openRows = getOpenStateRows_();
  return {
    success:true,
    reported: openRows.map(row => ({ floor: toInt_(row.FLOOR), point: String(row.POINT || '').trim() })),
    totalOpen: openRows.length,
    updatedAt: isoNow_()
  };
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

    ports.events.append( {
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

    ports.occurrences.upsert(floor, point, {
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
    ports.occurrences.update( row._rowIndex, { LAST_UPDATED: nowIso, STATUS: 'ABERTA', LAST_EVENT: 'VALIDACAO_ADMIN' });
    ports.events.append( {
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
    ports.occurrences.update( row._rowIndex, { LAST_UPDATED: nowIso, STATUS: 'REJEITADA', CLOSED_AT: nowIso, CLOSE_NOTES: safeText_(payload.notes || 'Rejeitado pela administração'), LAST_EVENT: 'REJEICAO_ADMIN' });
    ports.events.append( {
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

    ports.closures.append( {
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

    ports.events.append( {
      TIMESTAMP: nowIso, EVENTO:'FECHO', OCCURRENCE_ID:String(stateRow.OCCURRENCE_ID || ''), FLOOR:toInt_(stateRow.FLOOR), FLOOR_LABEL:String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)), POINT:String(stateRow.POINT || ''), LOCATION:String(stateRow.LOCATION || ''), REASON:'FECHO', NOTES:closeNotes, PHOTO_FILE_ID:closePhoto.fileId, PHOTO_FILE_URL:closePhoto.fileUrl, PHOTO_FILE_NAME:closePhoto.fileName, STATUS_BEFORE:String(stateRow.STATUS || ''), STATUS_AFTER:'OK', SOURCE:source, CLIENT_TS:clientTs
    });

    ports.occurrences.update( stateRow._rowIndex, { LAST_UPDATED:nowIso, STATUS:'OK', CLOSED_AT:nowIso, CLOSE_NOTES:closeNotes, CLOSE_PHOTO_FILE_ID:closePhoto.fileId, CLOSE_PHOTO_FILE_URL:closePhoto.fileUrl, CLOSE_PHOTO_FILE_NAME:closePhoto.fileName, LAST_EVENT:'FECHO' });
    sendCloseEmail_({ adminEmail:getAdminEmail_(payload), occurrenceId:String(stateRow.OCCURRENCE_ID || ''), floor:toInt_(stateRow.FLOOR), floorLabel:String(stateRow.FLOOR_LABEL || formatFloorLabel_(stateRow.FLOOR)), point:String(stateRow.POINT || ''), location:String(stateRow.LOCATION || ''), reason:String(stateRow.REASON || ''), reportedBy:String(stateRow.REPORTED_BY || ''), closeNotes, closePhotoUrl:closePhoto.fileUrl });
    return { success:true, occurrenceId:String(stateRow.OCCURRENCE_ID || ''), floor:toInt_(stateRow.FLOOR), point:String(stateRow.POINT || '') };
  });
}

function findStateRowByPoint_(floor, point) { const f = toInt_(floor); const p = normalizePoint_(point); return ports.occurrences.list().find(r => toInt_(r.FLOOR) === f && normalizePoint_(r.POINT) === p) || null; }

function findOpenStateRowByPoint_(floor, point) { const r = findStateRowByPoint_(floor, point); return r && isOpenStatus_(r.STATUS) ? r : null; }

function findOpenStateRowByOccurrenceId_(id) { const target = safeText_(id); if (!target) return null; return ports.occurrences.list().find(r => safeText_(r.OCCURRENCE_ID) === target && isOpenStatus_(r.STATUS)) || null; }

function findPendingStateRow_(id, floor, point) { const rows = ports.occurrences.list(); const target = safeText_(id); if (target) return rows.find(r => safeText_(r.OCCURRENCE_ID) === target && isPendingStatus_(r.STATUS)) || null; if (!isNaN(floor) && point) return rows.find(r => toInt_(r.FLOOR) === floor && normalizePoint_(r.POINT) === point && isPendingStatus_(r.STATUS)) || null; return null; }

function getOpenStateRows_() { return ports.occurrences.list().filter(r => isOpenStatus_(r.STATUS)); }

function getPendingStateRows_() { return ports.occurrences.list().filter(r => isPendingStatus_(r.STATUS)); }
  return {
    status: handleGetStatus_, open: handleGetOpenOccurrences_, pending: handleGetPendingOccurrences_,
    report: handlePostReport_, approve: handlePostApproveOccurrence_,
    reject: handlePostRejectOccurrence_, close: handlePostCloseOccurrence_
  };
}


// --- legacy-pin.js ---
// Compatibilidade com o PIN do backoffice antigo. Não constitui uma sessão autenticada.
function createLegacyPinService_(ports, CONFIG, domain, notifications, getAdminEmail_) {
  const { normalizePin_ } = domain;
  const withScriptLock_ = cb => ports.lock.run(cb);
  const isoNow_ = () => ports.clock.now().toISOString();
  const hashPin_ = (pin, salt) => ports.crypto.hashPin(pin, salt);
  const { sendPinResetEmail_ } = notifications;
function handleGetPinStatus_() {
  return { success:true, pinConfigured:isPinConfigured_(), updatedAt:ports.properties.getProperty(CONFIG.PROPS.PIN_UPDATED_AT) || '' };
}

function handlePostSetPin_(payload) {
  return withScriptLock_(function () {
    const pin = normalizePin_(payload.pin);
    if (!pin) throw new Error('PIN obrigatório.');
    if (pin.length < CONFIG.PIN_MIN_LENGTH) throw new Error('O PIN deve ter pelo menos ' + CONFIG.PIN_MIN_LENGTH + ' dígitos.');
    if (!/^\d+$/.test(pin)) throw new Error('O PIN deve conter apenas dígitos.');
    if (isPinConfigured_()) throw new Error('O PIN já está definido.');
    const salt = ports.crypto.uuid();
    ports.properties.setProperties({ [CONFIG.PROPS.PIN_HASH]: hashPin_(pin, salt), [CONFIG.PROPS.PIN_SALT]: salt, [CONFIG.PROPS.PIN_UPDATED_AT]: isoNow_() }, false);
    return { success:true, message:'PIN definido com sucesso.' };
  });
}

function handlePostValidatePin_(payload) {
  const pin = normalizePin_(payload.pin);
  if (!isPinConfigured_()) return { success:true, valid:false, message:'Ainda não existe PIN definido.' };
  if (!pin) return { success:true, valid:false, message:'PIN vazio.' };
  const props = ports.properties;
  return { success:true, valid: hashPin_(pin, props.getProperty(CONFIG.PROPS.PIN_SALT) || '') === (props.getProperty(CONFIG.PROPS.PIN_HASH) || '') };
}

function handlePostResetPin_(payload) {
  return withScriptLock_(function () {
    clearPin_();
    sendPinResetEmail_({ adminEmail:getAdminEmail_(payload) });
    return { success:true, message:'PIN resetado.' };
  });
}

function isPinConfigured_() { const p = ports.properties; return !!(p.getProperty(CONFIG.PROPS.PIN_HASH) && p.getProperty(CONFIG.PROPS.PIN_SALT)); }

function clearPin_() { const p = ports.properties; p.deleteProperty(CONFIG.PROPS.PIN_HASH); p.deleteProperty(CONFIG.PROPS.PIN_SALT); p.deleteProperty(CONFIG.PROPS.PIN_UPDATED_AT); }
  return { status: handleGetPinStatus_, set: handlePostSetPin_, validate: handlePostValidatePin_,
    reset: handlePostResetPin_, configured: isPinConfigured_, clear: clearPin_ };
}


// --- residents.js ---
// Identidade e ciclo de vida dos moradores: não pertencem ao módulo de garagem.
// Mantém as respostas e regras legadas. Ver as limitações de segurança em README.md.
function createResidentService_(ports, domain, notifications) {
  const { validateGarageResident_ } = domain;
  const repo = ports.residents;
  const setting = key => ports.settings.get(key);
  const byPin = pin => repo.list().find(c => String(c.pin) === String(pin)) || null;
  const dateText = date => ports.clock.format(date, 'dd/MM/yyyy HH:mm');

  function uniquePin() {
    const pins = repo.list().map(c => String(c.pin));
    for (let attempt = 0; attempt < 100; attempt++) {
      const pin = ports.ids.pinCandidate();
      if (pins.indexOf(pin) === -1 && pin !== setting('ADMIN_PIN')) return pin;
    }
    throw new Error('Não foi possível gerar PIN único.');
  }

  function register(d) {
    try {
      if (!d || !d.nome) throw new Error('Nome obrigatório.');
      if (!d.email) throw new Error('Email obrigatório.');
      const location = validateGarageResident_(d.piso, d.fracao);
      d = Object.assign({}, d, { piso: location.piso, fracao: location.fracao });
      const email = String(d.email).toLowerCase().trim();
      for (const c of repo.list()) {
        if (String(c.email).toLowerCase().trim() !== email) continue;
        const status = String(c.status || '').toUpperCase();
        if (status === 'APROVADO') return { ok: false, success: false, tipo: 'email_existente', recoveryAllowed: true, msg: 'Este email já tem acesso aprovado. Pode recuperar o PIN.' };
        if (status === 'PENDENTE') return { ok: false, success: false, tipo: 'pendente', msg: 'Já existe um pedido pendente com este email.' };
        if (status === 'BLOQUEADO') return { ok: false, success: false, tipo: 'bloqueado', msg: 'Este email está bloqueado. Contacte a administração.' };
        if (status === 'REJEITADO') return { ok: false, success: false, tipo: 'rejeitado', msg: 'Este email já teve um pedido não aprovado. Contacte a administração.' };
      }
      repo.add({
        id: ports.ids.entity('COND'), registeredAt: ports.clock.now(), name: d.nome,
        floor: d.piso, fraction: d.fracao, email: d.email, phone: d.telemovel || '',
        status: 'PENDENTE', pin: '', pinActive: false, approvedAt: '', approvedBy: '',
        lastAccess: '', notes: '', failedAttempts: 0, blockedUntil: ''
      });
      ports.log.add('INFO', 'REGISTO', d.email, 'Novo pedido: ' + d.nome);
      notifications.emailRegisto(d.email, d.nome);
      notifications.emailAdminNovoPedido(d);
      return { ok: true, success: true };
    } catch (err) {
      return { ok: false, success: false, msg: 'Erro ao registar: ' + err.message, message: 'Erro ao registar: ' + err.message };
    }
  }

  function login(pin) {
    try {
      const c = byPin(pin);
      if (!c) return { ok: false, tipo: 'invalido', msg: 'PIN inválido.' };
      if (c.status === 'PENDENTE') return { ok: false, tipo: 'pendente', msg: 'O seu pedido ainda está pendente.' };
      if (c.status === 'REJEITADO') return { ok: false, tipo: 'rejeitado', msg: 'Pedido rejeitado.' };
      if (c.status === 'BLOQUEADO') return { ok: false, tipo: 'bloqueado', msg: 'Acesso bloqueado.' };
      if (c.status === 'APROVADO') {
        repo.update(c.legacyRow, { lastAccess: ports.clock.now(), failedAttempts: 0, blockedUntil: '' });
        return { ok: true, success: true, nome: c.name, piso: c.floor, fracao: c.fraction };
      }
      return { ok: false, tipo: 'invalido', msg: 'Estado desconhecido.' };
    } catch (err) {
      return { ok: false, tipo: 'erro', msg: 'Erro interno.' };
    }
  }

  function resendPin(email) {
    try {
      const target = String(email || '').toLowerCase().trim();
      const c = repo.list().find(c => String(c.email).toLowerCase().trim() === target);
      if (!c) return { ok: false, success: false, msg: 'Email não encontrado.' };
      if (String(c.status || '').toUpperCase() !== 'APROVADO' || !c.pin) return { ok: false, success: false, msg: 'Não existe conta aprovada com este email.' };
      notifications.emailPINRecuperacao(email, c.name, c.pin);
      return { ok: true, success: true, msg: 'PIN enviado para o email registado.' };
    } catch (err) {
      return { ok: false, success: false, msg: 'Erro ao reenviar PIN: ' + err.message };
    }
  }

  function loginAdmin(email, pin) {
    return String(email || '').toLowerCase() === String(setting('ADMIN_EMAIL') || '').toLowerCase()
      && String(pin) === String(setting('ADMIN_PIN'))
      ? { ok: true, success: true } : { ok: false, msg: 'Credenciais inválidas.' };
  }

  function pending() {
    return repo.list().filter(c => c.status === 'PENDENTE').map(c => ({
      row: c.legacyRow, id: c.id, nome: c.name, piso: c.floor, fracao: c.fraction,
      email: c.email, telemovel: c.phone, dataRegisto: c.registeredAt ? dateText(c.registeredAt) : ''
    }));
  }

  function approved() {
    return repo.list().filter(c => c.status === 'APROVADO').map(c => ({
      row: c.legacyRow, id: c.id, nome: c.name, piso: c.floor, fracao: c.fraction,
      email: c.email, pin: c.pin, ultimoAcesso: c.lastAccess ? dateText(c.lastAccess) : 'Nunca'
    }));
  }

  function approve(row, adminEmail) {
    const c = repo.get(row);
    const pin = uniquePin();
    repo.update(row, { status: 'APROVADO', pin: pin, pinActive: true, approvedAt: ports.clock.now(), approvedBy: adminEmail || 'Admin' });
    notifications.emailAprovacao(c.email, c.name, pin);
    return { ok: true, success: true };
  }

  function reject(row) {
    const c = repo.get(row);
    repo.update(row, { status: 'REJEITADO' });
    notifications.emailRejeicao(c.email, c.name);
    return { ok: true, success: true };
  }

  function changeStatus(row, status) {
    repo.get(row);
    repo.update(row, { status: status });
    return { ok: true, success: true };
  }

  function regeneratePin(row) {
    const c = repo.get(row);
    const pin = uniquePin();
    repo.update(row, { pin: pin });
    notifications.emailPINRecuperacao(c.email, c.name, pin);
    return { ok: true, success: true };
  }

  return {
    register: register, login: login, loginAdmin: loginAdmin, resendPin: resendPin,
    pending: pending, approved: approved, approve: approve, reject: reject,
    block: row => changeStatus(row, 'BLOQUEADO'), unblock: row => changeStatus(row, 'APROVADO'),
    regeneratePin: regeneratePin,
    // Legado: esta ação nunca contou tentativas. A implementação segura exige alteração coordenada do login.
    failedAttempt: () => ({ ok: true, success: true })
  };
}


// --- accesses.js ---
// Consulta e gestão de códigos de recursos comuns. Independente da tecnologia de armazenamento.
function createAccessService_(ports) {
  const setting = key => ports.settings.get(key);

  function getCode(pin, reason, reasonOther, userAgent) {
    const c = ports.residents.list().find(c => String(c.pin) === String(pin));
    if (!c || c.status !== 'APROVADO') return { ok: false, msg: 'Acesso inválido.' };
    const code = setting('CODIGO_CADEADO');
    ports.consultations.add({
      id: ports.ids.entity('CONS'), at: ports.clock.now(), pin: pin,
      name: c.name, floor: c.floor, fraction: c.fraction, email: c.email,
      reason: reason, reasonOther: reasonOther || '', code: code, result: 'SUCESSO', userAgent: userAgent || ''
    });
    return { ok: true, success: true, codigo: code, aviso: setting('TEXTO_AVISO'), tempo: parseInt(setting('TEMPO_VISIVEL_SEGUNDOS')) || 15 };
  }

  function dashboard() {
    const residents = ports.residents.list();
    const today = ports.clock.now();
    const todayCount = ports.consultations.list().filter(record => {
      const date = new Date(record.at);
      return date.getDate() === today.getDate() && date.getMonth() === today.getMonth() && date.getFullYear() === today.getFullYear();
    }).length;
    return {
      ok: true, pendentes: residents.filter(c => c.status === 'PENDENTE').length,
      ativos: residents.filter(c => c.status === 'APROVADO').length,
      consultasHoje: todayCount, codigoAtual: setting('CODIGO_CADEADO')
    };
  }

  function history(limit) {
    return ports.consultations.list().slice().reverse().slice(0, limit || 50).map(record => ({
      dataHora: record.at ? ports.clock.format(record.at, 'dd/MM/yyyy HH:mm') : '',
      pin: record.pin, nome: record.name, piso: record.floor, fracao: record.fraction,
      email: record.email, motivo: record.reason, motivoOutro: record.reasonOther,
      codigo: record.code, resultado: record.result
    }));
  }

  function changeCode(code) {
    if (!/^\d+$/.test(String(code)) || String(code).length < 4) return { ok: false, msg: 'Código inválido.' };
    ports.settings.set('CODIGO_CADEADO', code);
    return { ok: true, success: true };
  }

  return { getCode: getCode, dashboard: dashboard, history: history, changeCode: changeCode };
}


// --- application.js ---
// Composição dos módulos e contrato de pedidos independente do transporte HTTP.
// As ações garage.* são aliases de compatibilidade, não a organização interna da app.
function createCondominiumApplication_(ports, config, domain) {
  const notifications = createNotificationService_(ports, config, domain);
  function adminEmail(payload) {
    const candidate = domain.safeText_(payload && payload.adminEmail);
    if (domain.isRealEmail_(candidate)) return candidate;
    if (domain.isRealEmail_(config.ADMIN_EMAIL)) return config.ADMIN_EMAIL;
    const effective = domain.safeText_(ports.environment.effectiveEmail());
    if (domain.isRealEmail_(effective)) return effective;
    const active = domain.safeText_(ports.environment.activeEmail());
    return domain.isRealEmail_(active) ? active : '';
  }
  const occurrences = createOccurrenceService_(ports, config, domain, notifications, adminEmail);
  const pins = createLegacyPinService_(ports, config, domain, notifications, adminEmail);
  const residents = createResidentService_(ports, domain, notifications);
  const accesses = createAccessService_(ports);
  const setting = key => ports.settings.get(key);
  const routes = { GET: Object.create(null), POST: Object.create(null) };

  function route(method, names, handler) {
    names.split(' ').forEach(name => routes[method][name] = handler);
  }
  route('GET', 'status', occurrences.status);
  route('GET', 'pinStatus', pins.status);
  route('GET', 'openOccurrences', occurrences.open);
  route('GET', 'pendingOccurrences', occurrences.pending);
  route('GET', 'health', () => ({
    success: true, appName: config.APP_NAME, spreadsheetId: config.SPREADSHEET_ID,
    rootFolderId: config.ROOT_FOLDER_ID, adminEmail: adminEmail({}), pinConfigured: pins.configured(),
    garageConfigReady: !!setting('NOME_CONDOMINIO'), now: ports.clock.now().toISOString()
  }));
  route('GET', 'garage.publicConfig garagePublicConfig', () => ({
    nomeCondominio: setting('NOME_CONDOMINIO') || 'Condomínio',
    tempoVisivel: parseInt(setting('TEMPO_VISIVEL_SEGUNDOS')) || 15,
    adminEmail: setting('ADMIN_EMAIL') || '', structure: domain.getGarageStructure_().floors
  }));
  route('GET', 'garage.structure garageStructure', domain.getGarageStructure_);
  route('GET', 'garage.dashboard garageDashboard', accesses.dashboard);
  route('GET', 'garage.pending garagePending', residents.pending);
  route('GET', 'garage.approved garageApproved', residents.approved);
  route('GET', 'garage.history garageHistory', () => accesses.history(50));
  route('GET', 'garage.adminConfig garageAdminConfig', () => ({
    nomeCondominio: setting('NOME_CONDOMINIO'), adminEmail: setting('ADMIN_EMAIL'),
    codigoAtual: setting('CODIGO_CADEADO'), tempoVisivel: setting('TEMPO_VISIVEL_SEGUNDOS'),
    maxTentativas: setting('MAX_TENTATIVAS_PIN'), bloqueioMinutos: setting('BLOQUEIO_MINUTOS'), textoAviso: setting('TEXTO_AVISO')
  }));
  route('POST', 'report', occurrences.report);
  route('POST', 'approveOccurrence', occurrences.approve);
  route('POST', 'rejectOccurrence', occurrences.reject);
  route('POST', 'closeOccurrence', occurrences.close);
  route('POST', 'setPin', pins.set);
  route('POST', 'validatePin', pins.validate);
  route('POST', 'resetPin', pins.reset);
  route('POST', 'garage.register garageRegister', p => residents.register({
    nome: p.nome || p.name, piso: p.piso || p.floor, fracao: p.fracao || p.fraction,
    email: p.email, telemovel: p.telemovel || p.telefone || p.phone
  }));
  route('POST', 'garage.loginPin garageLoginPin', p => residents.login(p.pin));
  route('POST', 'garage.failedAttempt garageFailedAttempt', residents.failedAttempt);
  route('POST', 'garage.resendPin garageResendPin garage.recoverCode garageRecoverCode', p => residents.resendPin(p.email));
  route('POST', 'garage.getCode garageGetCode', p => accesses.getCode(p.pin, p.motivo || p.reason, p.motivoOutro || p.reasonOther || '', p.userAgent || p.ua || ''));
  route('POST', 'garage.loginAdmin garageLoginAdmin', p => residents.loginAdmin(p.email, p.pin));
  route('POST', 'garage.approve garageApprove', p => residents.approve(Number(p.row), p.adminEmail || p.email || 'Admin'));
  route('POST', 'garage.reject garageReject', p => residents.reject(Number(p.row)));
  route('POST', 'garage.block garageBlock', p => residents.block(Number(p.row)));
  route('POST', 'garage.unblock garageUnblock', p => residents.unblock(Number(p.row)));
  route('POST', 'garage.regeneratePin garageRegeneratePin', p => residents.regeneratePin(Number(p.row)));
  route('POST', 'garage.changeCode garageChangeCode', p => accesses.changeCode(p.codigo || p.code || p.novoCodigo));
  route('POST', 'garage.saveConfig garageSaveConfig', p => {
    const configs = p.configs && typeof p.configs === 'object' ? p.configs : p;
    Object.keys(configs).forEach(key => { if (configs[key] !== '') ports.settings.set(key, configs[key]); });
    return { ok: true, success: true };
  });

  function dispatch(method, action, payload) {
    try {
      const handler = routes[method] && routes[method][action];
      if (!handler) return { success: false, message: 'Ação ' + method + ' inválida: ' + action };
      return handler(payload || {});
    } catch (err) {
      return { success: false, message: err && err.message ? err.message : 'Erro interno no backend' };
    }
  }
  return { dispatch: dispatch, initialize: () => ports.initialize(), clearLegacyPin: pins.clear };
}


// --- security.js ---
// Política de acesso aplicada a TODAS as entradas HTTP. Sem dependências Google.
function createSecureApplication_(app, ports) {
  const hash = value => ports.crypto.hashPin(String(value), 'dc4-session-v1');
  const now = () => ports.clock.now().getTime();
  const read = key => JSON.parse(ports.properties.getProperty(key) || 'null');
  const write = (key, value) => ports.properties.setProperties({ [key]: JSON.stringify(value) }, false);
  const canonical = action => action.replace(/^garage([A-Z])/, (_, c) => 'garage.' + c.toLowerCase());
  const deny = (message, code) => ({ ok: false, success: false, valid: false, msg: message, message, code: code || 'AUTH_REQUIRED' });
  const active = r => r && r.status === 'APROVADO' && String(r.pinActive).toLowerCase() === 'true' && !(new Date(r.blockedUntil).getTime() > now());
  function fingerprint(role, resident) {
    return hash(role === 'admin' ? ports.settings.get('ADMIN_EMAIL') + ':' + ports.settings.get('ADMIN_PIN') : resident.pin + ':' + resident.status + ':' + resident.pinActive);
  }
  function issue(role, resident) {
    const id = role === 'admin' ? 'admin' : String(resident.id);
    const subject = hash(role + ':' + id);
    const token = subject + '.' + ports.crypto.uuid() + ports.crypto.uuid();
    const expiresAt = now() + (role === 'admin' ? 30 * 60 : 8 * 60 * 60) * 1000;
    write('DC4_SESSION_' + subject, { digest: hash(token), role, id, expiresAt, fingerprint: fingerprint(role, resident) });
    return { token, expiresAt, role };
  }
  function verify(token) {
    if (typeof token !== 'string' || !/^[a-f0-9]{64}\.[a-f0-9-]{72}$/.test(token)) return null;
    const key = 'DC4_SESSION_' + token.split('.')[0];
    const session = read(key);
    if (!session || session.expiresAt <= now() || session.digest !== hash(token)) return null;
    const resident = session.role === 'resident' ? ports.residents.list().find(r => String(r.id) === session.id) : null;
    if (session.role === 'resident' && !active(resident)) return null;
    if (session.fingerprint !== fingerprint(session.role, resident)) return null;
    return { ...session, resident, key };
  }
  function throttle(bucket, max, duration) {
    return ports.lock.run(() => {
      const key = 'DC4_RATE_' + bucket;
      let state = read(key);
      if (!state || state.until <= now()) state = { count: 0, until: now() + duration };
      if (state.count >= max) return false;
      state.count++; write(key, state); return true;
    });
  }
  const publicGet = new Set(['status', 'garage.publicConfig', 'garage.structure']);
  const adminGet = new Set(['health', 'openOccurrences', 'pendingOccurrences', 'garage.dashboard', 'garage.pending', 'garage.approved', 'garage.history', 'garage.adminConfig']);
  const adminPost = new Set(['approveOccurrence', 'rejectOccurrence', 'closeOccurrence', 'garage.approve', 'garage.reject', 'garage.block', 'garage.unblock', 'garage.regeneratePin', 'garage.changeCode', 'garage.saveConfig']);
  function dispatch(method, action, input) {
    try {
      action = canonical(action);
      const p = Object.assign({}, input || {});
      // O transporte POST permite consultas autenticadas sem colocar tokens no URL.
      if (method === 'POST' && p._method === 'GET') method = 'GET';
      if (method === 'GET' && publicGet.has(action)) return app.dispatch(method, action, p);
      if (method === 'POST' && ['garage.loginAdmin', 'garage.loginPin'].includes(action)) {
        const admin = action === 'garage.loginAdmin';
        if (!throttle(admin ? 'admin-login' : 'resident-login', admin ? 10 : 30, 15 * 60 * 1000)) return deny('Demasiadas tentativas. Aguarde 15 minutos.', 'RATE_LIMITED');
        const pin = String(p.pin || '');
        if (!/^\d{6,12}$/.test(pin) || (admin && ['123456', '000000'].includes(pin))) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        const resident = admin ? null : ports.residents.list().find(r => String(r.pin) === pin);
        if (admin ? !ports.settings.get('ADMIN_PIN') || !ports.settings.get('ADMIN_EMAIL') : !active(resident)) return deny('Credenciais inválidas ou acesso indisponível.', 'INVALID_CREDENTIALS');
        const result = app.dispatch(method, action, p);
        if (!(result.ok || result.success)) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        return Object.assign({}, result, issue(admin ? 'admin' : 'resident', resident));
      }
      if (method === 'POST' && ['garage.register', 'garage.resendPin', 'garage.recoverCode'].includes(action)) {
        if (!throttle('public-mail', 10, 60 * 60 * 1000)) return deny('Limite de pedidos atingido. Tente mais tarde.', 'RATE_LIMITED');
        if (action !== 'garage.register') {
          const resident = ports.residents.list().find(r => String(r.email).trim().toLowerCase() === String(p.email || '').trim().toLowerCase());
          if (active(resident)) app.dispatch('POST', 'garage.resendPin', { email: resident.email });
          return { ok: true, success: true, msg: 'Se existir uma conta aprovada, receberá o PIN no email registado.' };
        }
        return app.dispatch(method, action, p);
      }
      // A criação/reset do PIN antigo deixa de estar exposta por HTTP.
      if (['pinStatus', 'setPin', 'validatePin', 'resetPin'].includes(action)) return deny('Entre pela área de administração da app.', 'LEGACY_LOGIN_DISABLED');
      const session = verify(p.token);
      if (!session) return deny('Sessão terminada ou inválida. Volte a entrar.');
      if (method === 'POST' && action === 'auth.logout') {
        ports.properties.deleteProperty(session.key); return { ok: true, success: true };
      }
      if (method === 'POST' && action === 'auth.session') return { ok: true, success: true, role: session.role, expiresAt: session.expiresAt };
      if (adminGet.has(action) && method === 'GET' || adminPost.has(action) && method === 'POST') {
        if (session.role !== 'admin') return deny('Acesso reservado à administração.', 'FORBIDDEN');
        p.adminEmail = ports.settings.get('ADMIN_EMAIL');
        if (action === 'garage.saveConfig') {
          const allowed = new Set(['NOME_CONDOMINIO', 'ADMIN_EMAIL', 'ADMIN_PIN', 'CODIGO_CADEADO', 'TEMPO_VISIVEL_SEGUNDOS', 'MAX_TENTATIVAS_PIN', 'BLOQUEIO_MINUTOS', 'TEXTO_AVISO']);
          const configs = p.configs || {};
          if (Object.keys(configs).some(k => !allowed.has(k))) return deny('Configuração não permitida.', 'INVALID_CONFIG');
          if (configs.ADMIN_PIN && (!/^\d{6,12}$/.test(String(configs.ADMIN_PIN)) || ['123456', '000000'].includes(String(configs.ADMIN_PIN)))) return deny('O PIN administrativo deve ter pelo menos 6 dígitos.', 'INVALID_CONFIG');
          p.configs = configs;
        }
        return app.dispatch(method, action, p);
      }
      if (method === 'POST' && action === 'garage.getCode' && session.role === 'resident') {
        p.pin = session.resident.pin;
        return app.dispatch(method, action, p);
      }
      if (method === 'POST' && action === 'report') {
        p.name = p.reportedBy = session.role === 'resident' ? session.resident.name : 'Administração';
        p.adminEmail = ports.settings.get('ADMIN_EMAIL');
        return app.dispatch(method, action, p);
      }
      return deny('Ação não autorizada.', 'FORBIDDEN');
    } catch (err) { return deny('Não foi possível processar o pedido.', 'SERVER_ERROR'); }
  }
  return { dispatch, initialize: app.initialize, clearLegacyPin: app.clearLegacyPin };
}


// --- apps-script-ports.js ---
// Adaptadores Google. Cada pedido recebe uma instância nova (sem cache entre pedidos).
function createAppsScriptPorts_(CONFIG, REQUIRED_HEADERS, domain) {
  const { safeText_, toInt_, normalizePoint_ } = domain;
  let spreadsheet;
  let settingsCache;
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
      list: () => getSheetObjects_('CONDOMINOS').map(row => decode(row, residentFields)),
      get: row => {
        const record = getSheetObjects_('CONDOMINOS').find(record => record._rowIndex === row);
        if (!record) throw new Error('Morador não encontrado.');
        return decode(record, residentFields);
      },
      add: record => appendObjectRow_('CONDOMINOS', encode(record, residentFields)),
      update: (row, patch) => updateObjectRow_('CONDOMINOS', row, encode(patch, residentFields))
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


// --- entrypoints.js ---
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
  const result = createAppsScriptApplication_().dispatch(method, action, payload);
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
