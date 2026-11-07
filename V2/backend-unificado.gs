// GERADO por node backend/build.cjs. Editar backend/src, não este ficheiro.
// Refatoração de compatibilidade: consultar backend/README.md antes de publicar.

// --- ../../occurrence-catalog.js ---
/* Catálogo partilhado entre interface e backend. */
const DC4_OCCURRENCE_CATALOG = {
  floors: [10,9,8,7,6,5,4,3,2,0,-1,-2,-3],
  statuses: {RECEBIDO:'Recebido',EM_ANALISE:'Em análise',AGENDADO:'Intervenção agendada',RESOLVIDO:'Resolvido',REJEITADO:'Não validado'},
  categories: [
    {id:'iluminacao',label:'Iluminação',scope:'floor',reasons:['Luz apagada','Luz intermitente','Sensor / temporizador','Outro']},
    {id:'portas',label:'Portas e portas corta-fogo',scope:'floor',reasons:['Não fecha corretamente','Fechadura / puxador','Porta danificada','Acesso obstruído','Outro']},
    {id:'limpeza',label:'Limpeza',scope:'both',reasons:['Sujidade','Resíduos abandonados','Derrame','Outro']},
    {id:'conservacao',label:'Pavimento, paredes e teto',scope:'floor',reasons:['Dano / peça solta','Infiltração / humidade','Outro']},
    {id:'elevadores',label:'Elevadores',scope:'general',reasons:['Fora de serviço','Portas não abrem / fecham','Ruído / funcionamento irregular','Outro']},
    {id:'portao',label:'Portão da garagem',scope:'general',reasons:['Não abre / fecha','Sensor / comando','Ruído / dano','Outro']},
    {id:'entrada',label:'Porta de entrada / intercomunicador',scope:'general',reasons:['Não abre / fecha','Código / chave / fechadura','Intercomunicador','Outro']},
    {id:'agua',label:'Abastecimento de água / bombas',scope:'general',reasons:['Falta de água','Pressão irregular','Fuga de água','Ruído nas bombas','Outro']},
    {id:'exterior',label:'Exterior / fachada',scope:'general',reasons:['Dano / peça solta','Infiltração','Outro']},
    {id:'outro',label:'Outro problema',scope:'both',reasons:['Outro']}
  ]
};
if (typeof window !== 'undefined') window.DC4_OCCURRENCE_CATALOG = DC4_OCCURRENCE_CATALOG;


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
    'LAST_UPDATED','STATUS','OCCURRENCE_ID','FLOOR','FLOOR_LABEL','POINT','LOCATION','REPORTED_AT','REPORTED_BY','REASON','NOTES','PHOTO_FILE_ID','PHOTO_FILE_URL','PHOTO_FILE_NAME','CLOSED_AT','CLOSE_NOTES','CLOSE_PHOTO_FILE_ID','CLOSE_PHOTO_FILE_URL','CLOSE_PHOTO_FILE_NAME','LAST_EVENT','PUBLIC_NOTES'
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
function approvedDescription_(row) {
  // Before 3.7.2 only an immediately preceding admin validation proves provenance.
  return safeText_(row.LAST_EVENT === 'VALIDACAO_ADMIN' ? row.NOTES : row.PUBLIC_NOTES);
}
function handleGetStatus_(payload) {
  const openRows = getOpenStateRows_();
  const withDetails = payload && payload.details === 'public';
  // Descrição da ocorrência aberta autorizada para consulta pública; excluir autor e fotografias.
  const categories = ['Extintor em falta', 'Extintor danificado', 'Suporte vazio',
    'Selo / verificação em falta', 'Acesso obstruído', 'Outro'];
  const result = {
    success: true,
    reported: openRows.map(row => {
      const item = { floor: toInt_(row.FLOOR), point: String(row.POINT || '').trim() };
      if (withDetails) {
        const reason = String(row.REASON || '').trim();
        item.reason = categories.includes(reason) ? reason : 'Outra anomalia';
        item.description = approvedDescription_(row);
        const date = new Date(row.REPORTED_AT);
        item.reportedAt = row.REPORTED_AT && !isNaN(date.getTime()) ? date.toISOString() : '';
      }
      return item;
    }),
    totalOpen: openRows.length,
    updatedAt: isoNow_()
  };
  if (withDetails) result.publicDetailsVersion = 2;
  return result;
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
      PUBLIC_NOTES: alreadyOpen ? approvedDescription_(existingState) : '',
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
    if (payload.publicNotes !== undefined && safeText_(payload.publicNotes) !== safeText_(row.NOTES)) {
      throw new Error('As notas mudaram. Atualize os pedidos e reveja o texto antes de validar.');
    }
    const nowIso = isoNow_();
    ports.occurrences.update( row._rowIndex, { LAST_UPDATED: nowIso, STATUS: 'ABERTA', LAST_EVENT: 'VALIDACAO_ADMIN_372', PUBLIC_NOTES: safeText_(payload.publicNotes) });
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
      email: c.email, ultimoAcesso: c.lastAccess ? dateText(c.lastAccess) : 'Nunca'
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
      id: ports.ids.entity('CONS'), at: ports.clock.now(),
      name: c.name, floor: c.floor, fraction: c.fraction, email: c.email,
      reason: reason, reasonOther: reasonOther || '', result: 'SUCESSO', userAgent: userAgent || ''
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
      nome: record.name, piso: record.floor, fracao: record.fraction,
      email: record.email, motivo: record.reason, motivoOutro: record.reasonOther,
      resultado: record.result
    }));
  }

  function changeCode(code) {
    if (!/^\d+$/.test(String(code)) || String(code).length < 4) return { ok: false, msg: 'Código inválido.' };
    ports.settings.set('CODIGO_CADEADO', code);
    return { ok: true, success: true };
  }

  return { getCode: getCode, dashboard: dashboard, history: history, changeCode: changeCode };
}


// --- general-occurrences.js ---
// Eventos duráveis: uma linha contém o estado e o recibo do pedido, sob lock.
function createGeneralOccurrenceService_(ports) {
  function invalid(message) { const e=new Error(message); e.code='VALIDATION_ERROR'; throw e; }
  const catalog = DC4_OCCURRENCE_CATALOG;
  const publicStates = ['EM_ANALISE','AGENDADO','RESOLVIDO'];
  const text = (v,max) => String(v == null ? '' : v).trim().slice(0,max);
  const hash = v => ports.crypto.hashPin(JSON.stringify(v),'dc4-general-v1');
  const all = () => ports.general.list().map(r => ({...r, data:JSON.parse(r.DATA_JSON)}));
  function latest(events) {
    const map = new Map(); events.forEach(e => map.set(e.data.id,e.data)); return [...map.values()];
  }
  function publicItem(o) {
    return {id:o.id,scope:o.scope,floor:o.floor,category:o.category,reason:o.reason,
      location:o.publicLocation,description:o.publicDescription,status:o.status,
      reportedAt:o.reportedAt,updatedAt:o.updatedAt,scheduledFor:o.scheduledFor||'',
      resolution:o.publicResolution||'',confirmations:(o.confirmations||[]).length};
  }
  function list() { return {success:true,occurrences:latest(all()).filter(o=>publicStates.includes(o.status)).map(publicItem)}; }
  function adminList() { return {success:true,occurrences:latest(all())}; }
  function mutate(action,p,change) {
    try { return ports.lock.run(()=>{
      const requestId=text(p.requestId,100),actor=p._actor;
      if(!actor || !/^[a-zA-Z0-9_-]{16,80}$/.test(requestId))invalid('Identificador de envio inválido. Atualize a página.');
      const clean={...p};['token','action','_method','requestId','_actor','_actorName','adminEmail'].forEach(k=>delete clean[k]);
      const digest=hash(clean),events=all();
      const old=events.find(e=>e.REQUEST_ID===requestId&&e.ACTOR_ID===actor&&e.ACTION===action);
      if(old){if(old.PAYLOAD_HASH!==digest)invalid('O pedido já foi recebido com outro conteúdo. Consulte as ocorrências antes de criar um novo.');return {success:true,occurrenceId:old.data.id,replayed:true};}
      const state=change(latest(events));
      const at=ports.clock.now().toISOString();state.updatedAt=at;
      state.revision=(state.revision||0)+1;
      ports.general.append({EVENT_ID:ports.crypto.uuid(),REQUEST_ID:requestId,ACTOR_ID:actor,ACTION:action,PAYLOAD_HASH:digest,AT:at,DATA_JSON:JSON.stringify(state)});
      return {success:true,occurrenceId:state.id};
    }); } catch (e) { return {success:false,code:e.code||'RETRY_SAME_REQUEST',message:e.code?e.message:'Não foi possível confirmar a gravação. Repita o mesmo envio.'}; }
  }
  function report(p) {return mutate('report',p,()=>{
    const scope=p.scope,floor=scope==='floor'?Number(p.floor):null;
    if(!['floor','general'].includes(scope)||scope==='floor'&&(p.floor==null||String(p.floor).trim()===''||!catalog.floors.includes(floor)))invalid('Localização inválida.');
    const cat=catalog.categories.find(c=>c.id===p.category&&(c.scope==='both'||c.scope===scope));
    if(!cat||!cat.reasons.includes(p.reason))invalid('Tipo ou motivo inválido.');
    const description=text(p.description,2000),location=text(p.location,160);
    if(!description)invalid('Descreva o problema.');
    if(scope==='general'&&!location)invalid('Indique o equipamento ou local onde observou o problema.');
    const photoData=String(p.photoBase64||'');
    if(photoData.length>1600000||photoData&&!['image/jpeg','image/png','image/webp'].includes(p.photoType))invalid('Fotografia inválida ou demasiado grande.');
    const id='GER-'+ports.crypto.uuid(),at=ports.clock.now().toISOString();
    const photo=ports.files.save({base64:photoData,mimeType:p.photoType,fileName:'ocorrencia-'+id+'.jpg',folderName:'OCORRENCIAS_GERAIS_PRIVADAS',occurrenceId:id,prefix:'geral',private:true});
    return {id,scope,floor,category:cat.id,reason:p.reason,location,description,
      reportedBy:p._actorName,reporterId:p._actor,photoUrl:photo.fileUrl||'',
      status:'RECEBIDO',reportedAt:at,publicLocation:'',publicDescription:'',internalNotes:'',confirmations:[],history:[{at,status:'RECEBIDO'}]};
  });}
  function update(p){return mutate('update',p,items=>{
    const old=items.find(o=>o.id===p.occurrenceId);if(!old)invalid('Ocorrência não encontrada.');
    if(Number(p.revision)!==old.revision)invalid('Esta ocorrência foi alterada. Atualize a lista antes de guardar.');
    if(!Object.prototype.hasOwnProperty.call(catalog.statuses,p.status))invalid('Estado inválido.');
    const description=text(p.publicDescription,2000),resolution=text(p.publicResolution,2000);
    if(publicStates.includes(p.status)&&!description)invalid('Preencha a descrição para os moradores.');
    if(p.status==='RESOLVIDO'&&!resolution)invalid('Indique o que foi feito para resolver.');
    const scheduledFor=text(p.scheduledFor,10);
    if(scheduledFor&&(!/^\d{4}-\d{2}-\d{2}$/.test(scheduledFor)||isNaN(Date.parse(scheduledFor))))invalid('Data inválida.');
    if(p.status==='AGENDADO'&&!scheduledFor)invalid('Indique a data da intervenção.');
    return {...old,status:p.status,publicDescription:description,publicLocation:text(p.publicLocation,160),publicResolution:resolution,scheduledFor,
      internalNotes:text(p.internalNotes,4000),history:[...(old.history||[]),{at:ports.clock.now().toISOString(),status:p.status,by:p._actorName}].slice(-100)};
  });}
  function confirm(p){return mutate('confirm',p,items=>{
    const old=items.find(o=>o.id===p.occurrenceId);if(!old||!['EM_ANALISE','AGENDADO'].includes(old.status))invalid('Esta ocorrência já não está aberta para confirmação.');
    return {...old,confirmations:[...new Set([...(old.confirmations||[]),p._actor])]};
  });}
  return {list,adminList,report,update,confirm};
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
  const general = createGeneralOccurrenceService_(ports);
  const pins = createLegacyPinService_(ports, config, domain, notifications, adminEmail);
  const residents = createResidentService_(ports, domain, notifications);
  const accesses = createAccessService_(ports);
  const setting = key => ports.settings.get(key);
  const routes = { GET: Object.create(null), POST: Object.create(null) };

  function route(method, names, handler) {
    names.split(' ').forEach(name => routes[method][name] = handler);
  }
  route('GET','general.status',general.list);
  route('GET','general.admin',general.adminList);
  route('POST','general.report',general.report);
  route('POST','general.update',general.update);
  route('POST','general.confirm',general.confirm);
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
  const publicGet = new Set(['general.status', 'status', 'garage.publicConfig', 'garage.structure']);
  const adminGet = new Set(['general.admin', 'health', 'openOccurrences', 'pendingOccurrences', 'garage.dashboard', 'garage.pending', 'garage.approved', 'garage.history', 'garage.adminConfig']);
  const adminPost = new Set(['general.update', 'approveOccurrence', 'rejectOccurrence', 'closeOccurrence', 'garage.approve', 'garage.reject', 'garage.block', 'garage.unblock', 'garage.regeneratePin', 'garage.changeCode', 'garage.saveConfig']);
  function dispatch(method, action, input) {
    try {
      action = canonical(action);
      const p = Object.assign({}, input || {});
      // O transporte POST permite consultas autenticadas sem colocar tokens no URL.
      if (method === 'POST' && p._method === 'GET') method = 'GET';
      if (method === 'GET' && publicGet.has(action)) return app.dispatch(method, action, p);
      if (method === 'POST' && ['garage.loginAdmin', 'garage.loginPin'].includes(action)) {
        const timings = {};
        let checkpoint = now();
        const mark = key => { const t = now(); timings[key] = Math.max(0, t - checkpoint); checkpoint = t; };
        const admin = action === 'garage.loginAdmin';
        if (!throttle(admin ? 'admin-login' : 'resident-login', admin ? 10 : 30, 15 * 60 * 1000)) return deny('Demasiadas tentativas. Aguarde 15 minutos.', 'RATE_LIMITED');
        mark('rateLimitMs');
        const pin = String(p.pin || '');
        if (!/^\d{6,12}$/.test(pin) || (admin && ['123456', '000000'].includes(pin))) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        const resident = admin ? null : ports.residents.list().find(r => String(r.pin) === pin);
        if (admin ? !ports.settings.get('ADMIN_PIN') || !ports.settings.get('ADMIN_EMAIL') : !active(resident)) return deny('Credenciais inválidas ou acesso indisponível.', 'INVALID_CREDENTIALS');
        mark('validationMs');
        const result = app.dispatch(method, action, p);
        mark('accessUpdateMs');
        if (!(result.ok || result.success)) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        const session = issue(admin ? 'admin' : 'resident', resident);
        mark('sessionMs');
        return Object.assign({}, result, session, { loginTimingsMs: timings });
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
      p._actor = session.role + ':' + session.id;
      p._actorName = session.role === 'resident' ? session.resident.name : 'Administração';
      // Financeiro: apenas POST autenticado; identidade e fração derivadas da sessão verificada.
      // Não encaminhar dados financeiros pelo router legado, e nunca aceitar roles fornecidas no body.
      if (method === 'POST' && /^finance\.(admin|resident)\./.test(action)) {
        const requiredRole=action.startsWith('finance.admin.')?'admin':'resident';
        if(session.role!==requiredRole)return deny('Sem autorização para esta área financeira.','FORBIDDEN');
        const resident = session.resident;
        const actor = session.role === 'admin'
          ? { role:'admin', email:String(ports.settings.get('ADMIN_EMAIL') || '').trim().toLowerCase() }
          : { role:'resident',
              email:String(resident.email || '').trim().toLowerCase(),
              fractionId:String(resident.floor || '')+String(resident.fraction || '') };
        if (!actor.email) return deny('Utilizador sem email autorizado.', 'FORBIDDEN');
        try { return financeDispatch_(p, actor); }
        catch (err) { return deny('Não foi possível executar o pedido financeiro.', 'FINANCE_ERROR'); }
      }
      if (method === 'POST' && ['general.report','general.confirm'].includes(action)) return app.dispatch(method,action,p);
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

// Sheets interprets strings starting with "=" as formulas even in appendRow.
function literalCell_(value) {
  return typeof value === 'string' && /^\s*=/.test(value) ? "'" + value : value;
}

function appendObjectRow_(sheetName, obj) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? literalCell_(obj[h]) : '')); }

function updateObjectRow_(sheetName, rowIndex, patch) { const sh = getSheet_(sheetName); const headers = getHeaders_(sh); const row = sh.getRange(rowIndex,1,1,headers.length).getValues()[0]; const merged = {}; headers.forEach((h,i) => merged[h] = row[i]); Object.keys(patch).forEach(k => merged[k] = patch[k]); sh.getRange(rowIndex,1,1,headers.length).setValues([headers.map(h => Object.prototype.hasOwnProperty.call(merged,h) ? literalCell_(merged[h]) : '')]); }

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
  try { if (!opts.private) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { Logger.log(err); }
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
    general: {list: () => getSheetObjects_('OCORRENCIAS_GERAIS'), append: row => appendObjectRow_('OCORRENCIAS_GERAIS',row)},
    migrateSecurity372: () => withScriptLock_(() => {
      const history = getSheet_('CONSULTAS');
      const headers = getHeaders_(history);
      const columns = ['PINUsado', 'CodigoMostrado'].map(name => headers.indexOf(name) + 1);
      if (columns.includes(0)) throw new Error('Verifique os cabeçalhos de CONSULTAS antes da migração.');
      ensureSheetStructure_(getSpreadsheet_(), CONFIG.SHEETS.ESTADO, REQUIRED_HEADERS.ESTADO);
      const rows = history.getLastRow() - 1;
      if (rows > 0) columns.forEach(column => history.getRange(2, column, rows, 1).clearContent());
    }),
    initializeGeneral: () => withScriptLock_(() => ensureSheetStructure_(getSpreadsheet_(),'OCORRENCIAS_GERAIS',['EVENT_ID','REQUEST_ID','ACTOR_ID','ACTION','PAYLOAD_HASH','AT','DATA_JSON'])),
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


// --- finance.js ---
/**
 * DC4 Finance — módulo para Google Apps Script (pré-produção)
 * NÃO PUBLICAR sem integrar no doPost e validação de sessão existentes.
 * Propriedades do script: FINANCE_SPREADSHEET_ID; FINANCE_PHOTOS_FOLDER_ID.
 * O router do projeto deve VALIDAR O TOKEN ANTES de chamar financeDispatch_(req, actor).
 * actor = { role: 'admin'|'resident', email: email validado, fractionId: fração validada }.
 * Nunca confiar em role, email ou fractionId enviados pelo browser.
 */
const FNC = {
  sheet: {
    invoices: 'FATURAS', settings:'WALLBOX_CONFIG', fractions:'WALLBOX_FRACOES',
    readings:'LEITURAS_WALLBOX', payments:'WALLBOX_PAGAMENTOS',
    discounts:'WALLBOX_DESCONTOS', tariff:'TARIFAS_EDP', history:'HISTORICO',
    cycles:'WALLBOX_CICLOS', rates:'WALLBOX_TARIFAS_CICLO', notifications:'WALLBOX_NOTIFICACOES'
  }
};

function financeDispatch_(req, actor) {
  if (!actor || !actor.role || !actor.email) throw new Error('AUTH_REQUIRED');
  const action = String(req.action || '');
  if (!/^finance\.(admin|resident)\./.test(action)) throw new Error('Ação não autorizada');
  if (action.indexOf('finance.admin.')===0 && actor.role!=='admin') throw new Error('FORBIDDEN');
  if (action.indexOf('finance.resident.')===0 && actor.role!=='resident') throw new Error('FORBIDDEN');
  const map = {
    'finance.admin.bootstrap': () => fncBootstrap_(),
    'finance.admin.invoices': () => fncInvoiceDashboard_(),
    'finance.admin.wallboxes': () => fncWallboxDashboard_(),
    'finance.admin.correctInitial': () => fncCorrectInitial_(req, actor),
    'finance.admin.correctReading': () => fncCorrectReading_(req, actor),
    'finance.admin.preflight': () => fncPreflight_(),
    'finance.admin.scanNow': () => financeScanFolder_(),
    'finance.admin.extractInvoice': () => fncExtractOneInvoice_(req.id),
    'finance.admin.scannerStatus': () => financeScannerStatus_(),
    'finance.admin.enableScanner': () => financeInstallScannerTrigger_(),
    'finance.admin.disableScanner': () => financeDisableScannerTrigger_(),
    'finance.admin.saveDiscount': () => fncSaveDiscount_(req, actor),
    'finance.admin.notifyCycle': () => fncNotifyCycle_(req,actor),
    'finance.admin.setNotifications': () => fncSetNotifications_(req,actor),
    'finance.admin.upsertFraction': () => fncUpsertFraction_(req, actor),
    'finance.admin.deactivateFraction': () => fncDeactivateFraction_(req, actor),
    'finance.admin.updateInvoice': () => fncUpdateInvoice_(req,actor),
    'finance.admin.validateInvoice': () => {
      const result=fncValidateInvoice_(req,actor);
      if(result.cycleId && String(fncSettings_().NOTIFICACOES_ATIVAS).toUpperCase()==='SIM')
        return {...result,notifications:fncNotifyCycle_({cycleId:result.cycleId},actor)};
      return result;
    },
    'finance.admin.markInvoicePaid': () => fncMarkInvoicePaid_(req, actor),
    'finance.admin.validateReading': () => fncValidateReading_(req, actor),
    'finance.admin.getReadingPhoto': () => fncGetReadingPhoto_(req,actor),
    'finance.admin.markWallboxPaid': () => fncMarkWallboxPaid_(req, actor),
    'finance.resident.eligibility': () => fncResidentEligibility_(actor),
    'finance.resident.status': () => fncResidentStatus_(actor),
    'finance.resident.amendReading': () => fncAmendReading_(req, actor),
    'finance.resident.submitReading': () => fncSubmitReading_(req, actor)
  };
  if (!map[action]) throw new Error('Ação financeira desconhecida');
  return map[action]();
}
function fncPrivateFolder_(name) {
  const folders=DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID).getFoldersByName(name);
  if(!folders.hasNext())throw new Error('Pasta privada não encontrada: '+name);
  const folder=folders.next();
  if(folders.hasNext())throw new Error('Existem várias pastas privadas com o mesmo nome: '+name);
  return folder;
}
function fncFinanceFileId_(){
  const override=PropertiesService.getScriptProperties().getProperty('FINANCE_SPREADSHEET_ID');
  if(override)return override;
  // Localização automática dentro da Drive que já é usada pelo backend do condomínio.
  const files=fncPrivateFolder_('FATURAS').getFilesByName('DC4 — Gestão Financeira (pré-produção)');
  if(!files.hasNext())throw new Error('Base de dados financeira não encontrada no Drive');
  const file=files.next();
  if(files.hasNext())throw new Error('Existem várias bases financeiras com o mesmo nome');
  return file.getId();
}
// Um único acesso à Spreadsheet financeira por execução (sem cache de leituras de células).
// A Apps Script reinicializa o contexto em novas execuções; dentro da mesma execução
// reutilizamos apenas a referência ao ficheiro, nunca os valores das tabelas.
let fncFinanceSpreadsheetRef_;
function fncDb_(){
  if(!fncFinanceSpreadsheetRef_)
    fncFinanceSpreadsheetRef_=SpreadsheetApp.openById(fncFinanceFileId_());
  return fncFinanceSpreadsheetRef_;
}
function fncTab_(name) {
  const sh=fncDb_().getSheetByName(name);
  if (!sh) throw new Error('Separador não encontrado: '+name);
  return sh;
}
function fncData_(name) {
  const values=fncTab_(name).getDataRange().getValues();
  if (!values.length) return [];
  const keys=values[0].map(String);
  return values.slice(1).map((row,i)=>({row:i+2,...Object.fromEntries(keys.map((key,j)=>[key,row[j]]))}))
    .filter(r=>Object.values(r).some(v=>v!=='' && v!=null && v!==r.row));
}
function fncSafeCell_(v){
  // Proteger Google Sheets contra injeção de fórmulas, incluindo nomes de ficheiros no Drive.
  return typeof v==='string' && /^[\s\u0000-\u001f]*[=]/.test(v) ? "'"+v : v;
}
function fncUpdate_(name,row,patch) {
  const sh=fncTab_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.entries(patch).forEach(([col,v])=>{
    const i=headers.indexOf(col);
    if (i<0) throw new Error('Coluna inexistente '+col);
    sh.getRange(row,i+1).setValue(fncSafeCell_(v));
  });
}
function fncAppend_(name,object) {
  const sh=fncTab_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => fncSafeCell_(object[h]===undefined?'':object[h])));
}
function fncLock_(fn) {
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function fncNow_(){return new Date();}
function fncNumber_(v){
  const s=String(v).trim().replace(/\s/g,'').replace(',','.');
  const x=Number(s);
  if(s==='' || !Number.isFinite(x)) throw new Error('Número inválido');
  return x;
}
function fncRequired_(val,label){if (!String(val??'').trim())throw new Error(label+' obrigatório');return String(val).trim();}
function fncAudit_(id,action,before,after,actor,note) {
  fncAppend_(FNC.sheet.history,{'Data':fncNow_(),'ID documento':id,
    'Ação':action,'Antes':JSON.stringify(before),'Depois':JSON.stringify(after),
    'Utilizador':actor.email,'Notas':note||'','Origem':'App DC4'});
}
function fncSettings_(){
  const rows=fncData_(FNC.sheet.settings);
  return Object.fromEntries(rows.map(r=>[String(r['Chave']),r['Valor']]));
}
/** Percentagem em vigor na data, sem aplicar alterações agendadas antes da vigência. */
function fncDiscountAt_(date){
  const day=fncIsoDate_(date);
  const history=fncData_(FNC.sheet.discounts)
    .filter(d=>d['Início vigência'] &&
      fncIsoDate_(d['Início vigência'])<=day)
    .sort((a,b)=>fncIsoDate_(a['Início vigência']).localeCompare(fncIsoDate_(b['Início vigência'])));
  if(history.length)return fncNumber_(history[history.length-1]['Desconto %']);
  const config=fncData_(FNC.sheet.settings).find(r=>r.Chave==='DESCONTO_FIXOS');
  if(!config || config.Valor==='' || config.Valor===null || config.Valor===undefined)return '';
  const when=config['Início vigência'];
  if(when && fncIsoDate_(when)>day)return '';
  return fncNumber_(config.Valor);
}


/** Vistas independentes: abrir faturas não lê wallboxes e vice-versa. */
function fncInvoiceDashboard_(){
  return {ok:true,invoices:fncData_(FNC.sheet.invoices).map(r=>({
    id:r.ID,supplier:r['Fornecedor'],description:r['Descrição'],
    number:r['Documento'],issued:r['Emissão'],due:r['Vencimento'],
    nif:r['NIF'],vat:r['IVA (€)'],subtotal:r['Base tributável (€)'],
    value:r['Total (€)'],validation:r['Validação'],payment:r['Pagamento'],
    documentType:r['Tipo'],method:r['Método'],
    extraction:r['Estado extração']||'Não processado',
    iban:r['IBAN fornecedor'],
    entity:r['Entidade MB'],reference:r['Referência MB'],url:r['PDF original']
  }))};
}
function fncWallboxDashboard_(){
  const latest=new Map();
  fncData_(FNC.sheet.fractions).forEach(r=>latest.set(String(r['ID fração']),r));
  const discount=fncDiscountAt_(fncNow_());
  const config=fncSettings_();
  return {ok:true,settings:{
    discount:discount===''?null:Number(discount),
    notifications:String(config.NOTIFICACOES_ATIVAS).toUpperCase()==='SIM'
  },fractions:[...latest.values()].map(r=>({
    id:r['ID fração'],name:r['Fração'],status:r['Estado'],email:r['Email autorizado'],
    startDate:fncDateDisplay_(r['Data adesão']),
    initialReading:r['Leitura inicial (kWh)'],initialDate:fncDateDisplay_(r['Data leitura inicial']),
    endDate:fncDateDisplay_(r['Data desativação'])
  })),
  readings:fncData_(FNC.sheet.readings).map(r=>({
    id:r['ID leitura'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],
    date:fncDateDisplay_(r['Data leitura']),kwh:r['Leitura acumulada kWh'],
    state:r['Estado']
  })),
  cycles:fncData_(FNC.sheet.cycles).map(r=>({
    id:r['ID ciclo'],invoiceId:r['Fatura ID'],state:r['Estado'],
    notifiedAt:fncDateDisplay_(r['Notificado em'])
  })),
  payments:fncData_(FNC.sheet.payments).map(r=>({
    id:r['ID apuramento'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],
    total:r['Total a pagar €'],state:r['Estado']
  }))};
}
function fncDateDisplay_(value){
  if(value===null||value===undefined||value==='')return '';
  if(value instanceof Date)return Utilities.formatDate(value,CONFIG.TZ,'yyyy-MM-dd');
  return String(value).slice(0,10);
}
/** Apenas para mostrar o atalho aos moradores: não carrega o histórico financeiro. */
function fncResidentEligibility_(actor){
  const id=String(actor.fractionId||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const email=String(actor.email||'').trim().toLowerCase();
  const latest=new Map();
  fncData_(FNC.sheet.fractions).forEach(r=>latest.set(String(r['ID fração']),r));
  const membership=latest.get(id);
  return {ok:true,eligible:Boolean(membership &&
    membership['Estado']==='Ativa' &&
    String(membership['Email autorizado']).trim().toLowerCase()===email)};
}
/** Corrigir a base de contagem sem reescrever movimentos financeiros. */
function fncCorrectInitial_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração');
    const membership=fncData_(FNC.sheet.fractions).filter(r=>String(r['ID fração'])===id).slice(-1)[0];
    if(!membership)throw new Error('Fração sem registo de wallbox');
    const date=fncRequired_(req.initialDate,'Data inicial');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Data inválida');
    const kwh=fncNumber_(req.initialReading);
    if(kwh<0)throw new Error('Consumo acumulado inválido');
    const reason=fncRequired_(req.reason,'Motivo da correção').slice(0,250);
    const updatedEmail=String(req.email??membership['Email autorizado']).trim().toLowerCase();
    if(updatedEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(updatedEmail))
      throw new Error('Email inválido');
    if(membership['Estado']==='Ativa' &&
       updatedEmail!==String(membership['Email autorizado']||'').trim().toLowerCase() &&
       !fncApprovedResident_(id,updatedEmail))
      throw new Error('O novo email não pertence a um morador aprovado nesta fração');
    const membershipStart=fncDateDisplay_(membership['Data adesão']);
    const paid=fncData_(FNC.sheet.payments).some(p=>p['Fração']===id &&
      (!membershipStart||fncDateDisplay_(p['Data inicial'])>=membershipStart));
    const approved=fncData_(FNC.sheet.readings).some(r=>r['Fração']===id &&
      r['Estado']==='Validada' && (!membershipStart||fncDateDisplay_(r['Data leitura'])>=membershipStart));
    if(paid||approved)throw new Error('Já existem leituras validadas ou apuramentos. Corrigir por regularização auditada, não alterando a base.');
    const pending=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===id&&r['Estado']==='Submetida');
    if(pending.some(r=>fncDateDisplay_(r['Data leitura'])<=date ||
       Number(r['Leitura acumulada kWh'])<kwh))
      throw new Error('A nova base não pode ultrapassar uma leitura pendente nem a respetiva data.');
    const previous={kwh:membership['Leitura inicial (kWh)'],
      date:fncDateDisplay_(membership['Data leitura inicial']),
      photoId:membership['Foto inicial Drive ID'],email:membership['Email autorizado']};
    const patch={'Leitura inicial (kWh)':kwh,'Data leitura inicial':date,'Atualizado em':fncNow_(),
      'Email autorizado':updatedEmail,'Observações':'Correção administrativa: '+reason};
    if(req.photoBase64)patch['Foto inicial Drive ID']=fncStorePhoto_(req.photoBase64,actor,'revisao-inicial-'+id);
    fncUpdate_(FNC.sheet.fractions,membership.row,patch);
    fncAudit_(id,'Corrigir leitura inicial',previous,
      {kwh,date,email:updatedEmail,photoId:patch['Foto inicial Drive ID']||previous.photoId},actor,reason);
    return {ok:true};
  });
}
/** Correção de leitura submetida, enquanto aguarda conferência. */
function fncEditPendingReading_(req,actor,ownedFraction){
  const id=fncRequired_(req.id,'Leitura');
  const reading=fncData_(FNC.sheet.readings).find(r=>String(r['ID leitura'])===id);
  if(!reading||reading['Estado']!=='Submetida')throw new Error('Só é possível corrigir leituras ainda por validar.');
  if(ownedFraction && reading['Fração']!==ownedFraction)
    throw new Error('Esta leitura não pertence à tua fração');
  const date=fncRequired_(req.readingDate,'Data da leitura');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Data inválida');
  const kwh=fncNumber_(req.kwh);
  if(kwh<0)throw new Error('kWh inválidos');
  const membership=fncData_(FNC.sheet.fractions).filter(r=>r['ID fração']===reading['Fração']).slice(-1)[0];
  if(!membership)throw new Error('Fração não encontrada');
  const joined=fncDateDisplay_(membership['Data adesão']);
  if(joined && date<joined)throw new Error('Data anterior à adesão à wallbox');
  const earlier=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===reading['Fração']&&
    r['Estado']==='Validada'&&fncDateDisplay_(r['Data leitura'])>=joined);
  const reference=earlier.sort((a,b)=>fncDateDisplay_(a['Data leitura']).localeCompare(fncDateDisplay_(b['Data leitura']))).slice(-1)[0];
  const baseKwh=reference?Number(reference['Leitura acumulada kWh']):Number(membership['Leitura inicial (kWh)']);
  const baseDate=fncDateDisplay_(reference?reference['Data leitura']:membership['Data leitura inicial']);
  if(date<=baseDate || kwh<baseKwh)throw new Error('Leitura inferior à anterior ou data inválida');
  if(fncData_(FNC.sheet.payments).some(p=>p['Fração']===reading['Fração']&&p['Ciclo EDP']===reading['Ciclo EDP']))
    throw new Error('Já existe apuramento, não é possível corrigir esta leitura.');
  const prev={kwh:reading['Leitura acumulada kWh'],date:fncDateDisplay_(reading['Data leitura']),
    photoId:reading['Foto Drive ID']};
  const patch={'Leitura acumulada kWh':kwh,'Data leitura':date};
  if(req.photoBase64)patch['Foto Drive ID']=fncStorePhoto_(req.photoBase64,actor,'revisao-'+id);
  fncUpdate_(FNC.sheet.readings,reading.row,patch);
  fncAudit_(id,'Corrigir leitura submetida',prev,
    {kwh,date,photoId:patch['Foto Drive ID']||prev.photoId},actor,
    String(req.reason||'Correção antes da validação'));
  return {ok:true};
}
function fncCorrectReading_(req,actor){return fncLock_(()=>fncEditPendingReading_(req,actor,null));}
function fncAmendReading_(req,actor){
  return fncLock_(()=>{
    const membership=fncResidentFraction_(actor);
    return fncEditPendingReading_(req,actor,String(membership['ID fração']));
  });
}

function fncBootstrap_(){
  const s=fncSettings_(),discount=fncDiscountAt_(fncNow_());
  // O histórico das adesões é mantido na folha; a interface só mostra
  // o estado mais recente de cada fração, sem duplicações visuais.
  const allFractions=fncData_(FNC.sheet.fractions),latest=new Map();
  allFractions.forEach(r=>latest.set(String(r['ID fração']),r));
  return {ok:true,settings:{discount:discount===''?null:Number(discount),
    notifications:String(s.NOTIFICACOES_ATIVAS).toUpperCase()==='SIM'},
    fractions:[...latest.values()].map(r=>({
      id:r['ID fração'],name:r['Fração'],status:r['Estado'],email:r['Email autorizado'],
      startDate:r['Data adesão'],initialReading:r['Leitura inicial (kWh)'],
      initialDate:r['Data leitura inicial'],endDate:r['Data desativação'],row:r.row
    })),
    invoices:fncData_(FNC.sheet.invoices).map(r=>({
      id:r.ID, supplier:r['Fornecedor'],description:r['Descrição'],due:r['Vencimento'],
      value:r['Total (€)'],validation:r['Validação'],payment:r['Pagamento'],
      documentType:r['Tipo'],method:r['Método'],iban:r['IBAN fornecedor'],entity:r['Entidade MB'],
      reference:r['Referência MB'],url:r['PDF original']
    })),
    readings:fncData_(FNC.sheet.readings).map(r=>({
      id:r['ID leitura'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],date:r['Data leitura'],
      kwh:r['Leitura acumulada kWh'],state:r['Estado'],row:r.row
    })),
    cycles:fncData_(FNC.sheet.cycles).map(r=>({
      id:r['ID ciclo'],invoiceId:r['Fatura ID'],state:r['Estado'],notifiedAt:r['Notificado em']
    })),
    payments:fncData_(FNC.sheet.payments).map(r=>({
      id:r['ID apuramento'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],
      total:r['Total a pagar €'],state:r['Estado'],row:r.row
    }))};
}

/**
 * Diagnóstico apenas de leitura: não configura nem ativa serviços.
 * Acessível apenas a administradores autenticados pelo router principal.
 */
function fncPreflight_(){
  const checks=[];
  const check=(id,ok,message)=>checks.push({id,ok:Boolean(ok),message});
  try{
    const sh=fncDb_();
    const names=[FNC.sheet.invoices,FNC.sheet.settings,FNC.sheet.fractions,
      FNC.sheet.readings,FNC.sheet.payments,FNC.sheet.discounts,FNC.sheet.tariff,
      FNC.sheet.history,FNC.sheet.cycles,FNC.sheet.rates,FNC.sheet.notifications];
    const missing=names.filter(name=>!sh.getSheetByName(name));
    check('finance_sheet',!missing.length,
      missing.length?'Separadores em falta: '+missing.join(', '):'Estrutura financeira localizada');
  }catch(err){
    check('finance_sheet',false,'Sem acesso à base de dados financeira');
    return {ok:true,ready:false,checks};
  }
  const settings=fncSettings_(),fractions=fncData_(FNC.sheet.fractions);
  const discount=settings.DESCONTO_FIXOS;
  check('discount',discount!==''&&discount!==null&&discount!==undefined&&
    Number(discount)>=0&&Number(discount)<=100,'Desconto global definido pela administração');
  const active=fractions.filter(r=>r['Estado']==='Ativa');
  check('fractions',active.length>0,'Pelo menos uma fração ativa');
  const validResidents=active.every(r=>
    fncApprovedResident_(r['ID fração'],r['Email autorizado']));
  check('residents',validResidents,'Moradores das frações ativas com conta aprovada');
  const invoices=fncData_(FNC.sheet.invoices);
  const validatedEdp=invoices.some(r=>r['Validação']==='Validada' &&
    /(^|\s)edp(\s|$)/i.test(String(r['Fornecedor']||'')));
  check('edp_invoice',validatedEdp,'Fatura de eletricidade validada');
  const cycles=fncData_(FNC.sheet.cycles).filter(r=>r['Estado']==='Aberto');
  check('cycle',cycles.length>0,'Existe ciclo de leitura aberto');
  try{
    const id=PropertiesService.getScriptProperties().getProperty('FINANCE_PHOTOS_FOLDER_ID');
    const photos=id?DriveApp.getFolderById(id):fncPrivateFolder_('WALLBOX_LEITURAS_PRIVADAS');
    check('private_photos',Boolean(photos.getId()),'Pasta privada das fotografias acessível');
  }catch(err){check('private_photos',false,'Não foi possível aceder à pasta privada de fotografias');}
  check('notifications',String(settings.NOTIFICACOES_ATIVAS).toUpperCase()==='SIM',
    'Notificações estão ativadas pela administração');
  // Não apresenta emails, IBAN, nomes nem fotografias.
  return {ok:true,ready:checks.every(x=>x.ok),checks};
}
function fncApprovedResident_(fractionId,email){
  const expected=String(email||'').trim().toLowerCase();
  if(!expected)return false;
  const match=/^(10|[2-9])([A-C])$/.exec(String(fractionId||''));
  if(!match)return false;
  const sheet=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('CONDOMINOS');
  if(!sheet)return false;
  const all=sheet.getDataRange().getValues();
  if(all.length<2)return false;
  const h=all[0].map(x=>String(x).trim());
  const pos=n=>h.indexOf(n);
  const needed=['Piso','Fracao','Email','Estado','PINAtivo'];
  if(needed.some(x=>pos(x)<0))return false;
  return all.slice(1).some(r=>
    String(r[pos('Piso')]).replace(/\D/g,'')===match[1] &&
    String(r[pos('Fracao')]).toUpperCase().trim()===match[2] &&
    String(r[pos('Email')]).trim().toLowerCase()===expected &&
    String(r[pos('Estado')]).trim().toUpperCase()==='APROVADO' &&
    ['true','sim','1'].includes(String(r[pos('PINAtivo')]).trim().toLowerCase()));
}

function fncSaveDiscount_(req,actor){
  return fncLock_(()=>{
    const pct=fncNumber_(req.percent);
    if(pct<0||pct>100)throw new Error('Desconto entre 0 e 100%');
    const rows=fncData_(FNC.sheet.settings), record=rows.find(r=>r.Chave==='DESCONTO_FIXOS');
    if(!record)throw new Error('Configuração do desconto inexistente');
    const old=record.Valor;
    const effective=fncRequired_(req.effectiveFrom,'Data de vigência');
    const oldDiscounts=fncData_(FNC.sheet.discounts);
    // Numa configuração preexistente, preservar o valor atual se o novo desconto
    // só entrar em vigor no futuro, mesmo sem histórico anterior.
    if(!oldDiscounts.length && old!=='' && old!==null && old!==undefined &&
       effective>fncIsoDate_(fncNow_()).slice(0,10)){
      fncAppend_(FNC.sheet.discounts,{'ID':'DESC-BASE-'+Utilities.getUuid(),
        'Início vigência':fncIsoDate_(fncNow_()).slice(0,10),
        'Fim vigência':effective,'Desconto %':fncNumber_(old),
        'Motivo':'Valor anterior à nova vigência',
        'Definido por':actor.email,'Registado em':fncNow_(),
        'Notas':'Valor inicial preservado ao agendar alteração futura'});
    }
    // Vigência global: nunca alterar/apagar descontos que sustentam apuramentos fechados.
    if(oldDiscounts.length){
      const last=oldDiscounts[oldDiscounts.length-1];
      if(last['Fim vigência']==='')fncUpdate_(FNC.sheet.discounts,last.row,{'Fim vigência':effective});
    }
    fncAppend_(FNC.sheet.discounts,{'ID':'DESC-'+Utilities.getUuid(),
      'Início vigência':effective,'Desconto %':pct,'Motivo':String(req.reason||'Alteração administrativa'),
      'Definido por':actor.email,'Registado em':fncNow_(),'Notas':'Aplicar apenas a ciclos futuros ainda não fechados'});
    fncUpdate_(FNC.sheet.settings,record.row,{'Valor':pct,'Início vigência':effective,
      'Atualizado em':fncNow_(),'Atualizado por':actor.email});
    fncAudit_('CONFIG_WALLBOX','Desconto',old,pct,actor);
    return {ok:true,discount:pct};
  });
}
function fncParseFraction_(id){
  const result=/^(10|[2-9])([A-C])$/.exec(String(id));
  if(!result)throw new Error('Fração inválida');
  return {floor:result[1],frac:result[2]};
}
function fncUpsertFraction_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração').toUpperCase().replace(/[^0-9A-Z]/g,'');
    const label=fncRequired_(req.label,'Descrição');
    const email=fncRequired_(req.email,'Email autorizado').toLowerCase();
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Email inválido');
    const initial=fncNumber_(req.initialReading);
    if(initial<0)throw new Error('Leitura inicial inválida');
    const start=fncRequired_(req.initialDate,'Data da leitura inicial');
    const {floor,frac}=fncParseFraction_(id);
    if(!GARAGE_RESIDENTIAL_STRUCTURE[String(floor)] ||
       !GARAGE_RESIDENTIAL_STRUCTURE[String(floor)].includes(frac))
      throw new Error('Fração desconhecida na estrutura do condomínio');
    if(!fncApprovedResident_(id,email))
      throw new Error('O morador deve estar registado e aprovado nesta fração antes da ativação');
    const fraction=fncData_(FNC.sheet.fractions).filter(r=>String(r['ID fração']).toUpperCase()===id).slice(-1)[0];
    if(fraction && String(fraction['Estado'])==='Ativa')throw new Error('Fração já ativa — desative primeiro');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(start))throw new Error('Data inicial inválida');
    const data={'ID fração':id,'Fração':label,'Estado':'Ativa',
      'Email autorizado':email,'Data adesão':start,'Leitura inicial (kWh)':initial,
      'Data leitura inicial':start,'Foto inicial Drive ID':fncStorePhoto_(req.initialPhotoBase64,actor,'inicio-'+id),
      'Atualizado em':fncNow_(),'Observações':String(req.note||'')};
    // Cada nova adesão cria uma linha própria: conserva a base/histórico da adesão anterior.
    fncAppend_(FNC.sheet.fractions,data);
    fncAudit_(id,'Ativar wallbox',fraction?.Estado||'',data,actor);
    return {ok:true};
  });
}
function fncDeactivateFraction_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração');
    const fraction=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id && r['Estado']==='Ativa');
    if(!fraction)throw new Error('Fração não está ativa');
    const date=fncRequired_(req.endDate,'Data fim de utilização');
    const final=fncNumber_(req.finalReading);
    const last=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===id&&r['Estado']==='Validada')
      .filter(r=>String(r['Data leitura'])>=String(fraction['Data adesão'])).slice(-1)[0];
    const minReading=last?Number(last['Leitura acumulada kWh']):Number(fraction['Leitura inicial (kWh)']);
    if(final<minReading)throw new Error('Leitura final inferior à última leitura válida');
    const finalPhotoId=fncStorePhoto_(req.finalPhotoBase64,actor,'fim-'+id);
    fncUpdate_(FNC.sheet.fractions,fraction.row,{'Estado':'Inativa','Data desativação':date,
      'Leitura final (kWh)':final,'Foto final Drive ID':finalPhotoId, 'Atualizado em':fncNow_()});
    fncAudit_(id,'Desativar wallbox','Ativa','Inativa',actor,'Histórico preservado; regularizar valores pendentes.');
    return {ok:true};
  });
}
function fncValidIban_(s) {
  if(!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(s))return false;
  const digits=s.slice(4)+s.slice(0,4);
  let remainder=0;
  for(const ch of digits){
    const numerical=ch>='A'&&ch<='Z'?String(ch.charCodeAt(0)-55):ch;
    for(const d of numerical)remainder=(remainder*10+Number(d))%97;
  }
  return remainder===1;
}
function fncUpdateInvoice_(req,actor) {
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Documento');
    const row=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===id);
    if(!row)throw new Error('Documento não encontrado');
    if(row['Validação']==='Validada'||row['Pagamento']==='Pago')
      throw new Error('Documento validado ou pago não pode ser editado sem anular a validação');
    const type=fncRequired_(req.documentType,'Tipo de documento');
    if(!['Fatura','Fatura + contribuição audiovisual','Nota despesas/honorários','Orçamento / honorários','Documento por classificar'].includes(type))
      throw new Error('Tipo de documento inválido');
    const supplier=fncRequired_(req.supplier,'Fornecedor').slice(0,160);
    const description=fncRequired_(req.description,'Serviço').slice(0,300);
    const value=fncNumber_(req.value);
    if(value<0||value>1000000)throw new Error('Valor fora do limite permitido');
    const due=String(req.due||'').trim();
    if(due&&!/^\d{4}-\d{2}-\d{2}$/.test(due))throw new Error('Data de vencimento inválida');
    const iban=String(req.iban||'').toUpperCase().replace(/\s/g,'');
    if(iban&&!fncValidIban_(iban))throw new Error('IBAN inválido');
    const entity=String(req.entity||'').replace(/\s/g,''),reference=String(req.reference||'').replace(/\s/g,'');
    if(entity&&!/^\d{5}$/.test(entity))throw new Error('Entidade Multibanco inválida');
    if(reference&&!/^\d{9}$/.test(reference))throw new Error('Referência Multibanco inválida');
    if(Boolean(entity)!==Boolean(reference))throw new Error('Entidade e referência são obrigatórias em conjunto');
    const method=fncRequired_(req.method,'Método de pagamento');
    if(!['Transferência','Multibanco','Débito direto','Por confirmar','Transferência apenas após fatura'].includes(method))
      throw new Error('Método de pagamento inválido');
    if(method==='Transferência'&&!iban)throw new Error('Indique o IBAN antes de validar a transferência');
    if(method==='Multibanco'&&!entity)throw new Error('Indique entidade e referência Multibanco');
    const patch={'Tipo':type,'Fornecedor':supplier,'Descrição':description,'Total (€)':value,
      'Vencimento':due,'Método':method,'IBAN fornecedor':iban,'Entidade MB':entity,'Referência MB':reference,
      'Validação':'Por validar','Validado por':''};
    fncUpdate_(FNC.sheet.invoices,row.row,patch);
    fncAudit_(id,'Corrigir documento',
      {tipo:row['Tipo'],valor:row['Total (€)'],metodo:row['Método']},
      {tipo:type,valor:value,metodo:method},actor,'Correções pendentes de validação');
    return {ok:true,id};
  });
}

function fncValidateInvoice_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Fatura');
    const invoice=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===id);
    if(!invoice)throw new Error('Fatura não encontrada');
    const kind=String(invoice['Tipo']||'').toLowerCase();
    if(kind.includes('orçamento')||kind.includes('por classificar'))
      throw new Error('Classificar e validar os dados antes de autorizar o pagamento');
    if(fncNumber_(invoice['Total (€)'])<=0)throw new Error('Total a pagar inválido');
    if(String(invoice['Validação'])==='Validada')return {ok:true,alreadyValidated:true};
    const isEdp=String(invoice['Fornecedor']).toLowerCase().includes('edp');
    const cycle=isEdp?fncOpenCycle_(invoice,actor):null;
    fncUpdate_(FNC.sheet.invoices,invoice.row,{'Validação':'Validada','Validado por':actor.email});
    fncAudit_(id,'Validar fatura',invoice['Validação'],'Validada',actor);
    if(isEdp)return {ok:true,edpCycleReady:true,cycleId:cycle,
      message:'Ciclo aberto. Notificação dos moradores depende de configuração e confirmação administrativa.'};
    return {ok:true};
  });
}
function fncMarkInvoicePaid_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Fatura'), date=fncRequired_(req.paidAt,'Data pagamento');
    const row=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===id);
    if(!row)throw new Error('Fatura não encontrada');
    if(String(row['Validação'])!=='Validada')throw new Error('Validar fatura primeiro');
    if(String(row['Pagamento'])==='Pago')return {ok:true,alreadyPaid:true};
    if(/orçamento|por classificar/i.test(String(row['Tipo']||'')))throw new Error('Documento não pagável');
    if(fncNumber_(row['Total (€)'])<=0)throw new Error('Valor inválido para pagamento');
    fncUpdate_(FNC.sheet.invoices,row.row,{'Pagamento':'Pago','Data pag.':date});
    fncAudit_(id,'Registar pagamento',row['Pagamento'],'Pago',actor,'Confirmado pelo administrador. Não executa transferência.');
    return {ok:true};
  });
}
function fncResidentFraction_(actor){
  const id=String(actor.fractionId||'').toUpperCase().replace(/[^0-9A-Z]/g,'');
  if(!id)throw new Error('Fração não associada à sessão');
  const row=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id&&r['Estado']==='Ativa');
  if(!row||row['Estado']!=='Ativa')throw new Error('Serviço de wallbox não ativo para a fração');
  if(String(row['Email autorizado']).toLowerCase()!==String(actor.email).toLowerCase())
    throw new Error('Morador não autorizado nesta fração');
  if(!fncApprovedResident_(id,actor.email))
    throw new Error('Conta de morador não aprovada ou desativada');
  return row;
}
function fncResidentStatus_(actor){
  const fraction=fncResidentFraction_(actor);
  // Um novo morador não recebe o histórico financeiro de ocupantes anteriores.
  const since=fncIsoDate_(fraction['Data adesão']);
  const readings=fncData_(FNC.sheet.readings)
    .filter(r=>r['Fração']===fraction['ID fração']&&fncIsoDate_(r['Data leitura'])>=since);
  const payments=fncData_(FNC.sheet.payments)
    .filter(r=>r['Fração']===fraction['ID fração']&&fncIsoDate_(r['Data inicial'])>=since);
  return {ok:true,fraction:fraction['Fração'],
    initialReading:fraction['Leitura inicial (kWh)'],initialDate:fraction['Data leitura inicial'],
    openCycle:(fncData_(FNC.sheet.cycles).filter(c=>c['Estado']==='Aberto').slice(-1)[0]||{})['ID ciclo']||null,
    readings:readings.map(r=>({id:r['ID leitura'],cycle:r['Ciclo EDP'],date:r['Data leitura'],kwh:r['Leitura acumulada kWh'],state:r['Estado']})),
    payments:payments.map(r=>({cycle:r['Ciclo EDP'],total:r['Total a pagar €'],state:r['Estado']}))};
}
function fncSubmitReading_(req,actor){
  return fncLock_(()=>{
    const fraction=fncResidentFraction_(actor);
    const kwh=fncNumber_(req.kwh), date=fncRequired_(req.readingDate,'Data leitura');
    if(kwh<0)throw new Error('Leitura inválida');
    const cycle=fncRequired_(req.cycleId,'Ciclo');
    const activeCycle=fncData_(FNC.sheet.cycles).find(c=>c['ID ciclo']===cycle && c['Estado']==='Aberto');
    if(!activeCycle)throw new Error('Ciclo não autorizado ou já encerrado');
    const all=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===fraction['ID fração']);
    if(all.some(r=>r['Ciclo EDP']===cycle && r['Estado']!=='Rejeitada'))throw new Error('Já existe leitura neste ciclo');
    const last=all.filter(r=>r['Estado']==='Validada')
      .filter(r=>fncIsoDate_(r['Data leitura'])>=fncIsoDate_(fraction['Data adesão']))
      .sort((a,b)=>fncIsoDate_(a['Data leitura']).localeCompare(fncIsoDate_(b['Data leitura']))).slice(-1)[0];
    const base=last?Number(last['Leitura acumulada kWh']):Number(fraction['Leitura inicial (kWh)']);
    if(kwh<base)throw new Error('Leitura inferior à última validada');
    const baseDate=last?fncIsoDate_(last['Data leitura']):fncIsoDate_(fraction['Data leitura inicial']);
    if(fncIsoDate_(date)<=baseDate)throw new Error('A nova leitura deve ser posterior à última validada');
    // O endpoint existente deve aceitar imagens apenas após auth. Nunca guardar base64 no Sheet.
    const photoId=fncStorePhoto_(req.photoBase64,actor,cycle);
    const rid='LR-'+Utilities.getUuid();
    fncAppend_(FNC.sheet.readings,{'ID leitura':rid,'Fração':fraction['ID fração'],
      'Ciclo EDP':cycle,'Data leitura':date,'Leitura acumulada kWh':kwh,
      'Foto Drive ID':photoId,'Estado':'Submetida','Submetido por':actor.email,'Submetido em':fncNow_()});
    return {ok:true,readingId:rid,message:'Leitura submetida para validação'};
  });
}
function fncStorePhoto_(b64,actor,cycle){
  const folderId=PropertiesService.getScriptProperties().getProperty('FINANCE_PHOTOS_FOLDER_ID');
  const folder=folderId?DriveApp.getFolderById(folderId):fncPrivateFolder_('WALLBOX_LEITURAS_PRIVADAS');
  const data=fncRequired_(b64,'Fotografia').replace(/^data:image\/(?:jpeg|png);base64,/,'');
  if(data.length>4300000 || !/^[A-Za-z0-9+/=]+$/.test(data))throw new Error('Imagem inválida ou demasiado grande');
  const isPng=/^data:image\/png;base64,/.test(String(b64));
  const bytes=Utilities.base64Decode(data), mime=isPng?'image/png':'image/jpeg';
  const ext=isPng?'.png':'.jpg';
  return folder.createFile(Utilities.newBlob(bytes,mime,'wallbox-'+cycle+'-'+Utilities.getUuid()+ext)).getId();
}
function fncGetReadingPhoto_(req,actor){
  // Apenas administradores com sessão válida podem ler fotografias privadas.
  // O ficheiro é resolvido pelo ID de leitura registado no Sheets, nunca por ID arbitrário do cliente.
  const id=fncRequired_(req.id,'ID de leitura');
  const reading=fncData_(FNC.sheet.readings).find(r=>r['ID leitura']===id);
  if(!reading || !reading['Foto Drive ID'])throw new Error('Fotografia não encontrada');
  const file=DriveApp.getFileById(String(reading['Foto Drive ID']));
  const blob=file.getBlob(),bytes=blob.getBytes();
  if(bytes.length>3.1*1024*1024)throw new Error('Imagem demasiado grande para visualizar na App');
  const mime=String(blob.getContentType()||'image/jpeg');
  if(!['image/jpeg','image/png'].includes(mime))throw new Error('Formato da fotografia não autorizado');
  return {ok:true,dataUrl:'data:'+mime+';base64,'+Utilities.base64Encode(bytes)};
}

function fncValidateReading_(req,actor){
  return fncLock_(()=>{
    const rid=fncRequired_(req.id,'Leitura');
    const r=fncData_(FNC.sheet.readings).find(x=>x['ID leitura']===rid);
    if(!r||r['Estado']!=='Submetida')throw new Error('Leitura não está pendente');
    const cycle=fncData_(FNC.sheet.cycles).find(x=>x['ID ciclo']===r['Ciclo EDP'] && x['Estado']==='Aberto');
    const tariffs=fncData_(FNC.sheet.rates).find(x=>x['ID ciclo']===r['Ciclo EDP']);
    if(!cycle||!tariffs)throw new Error('Ciclo e tarifas ainda não configurados');
    // O desconto é congelado no ciclo para que uma alteração administrativa futura
    // não altere a valorização da eletricidade já faturada a outras frações.
    const pct=cycle['Desconto aplicado %']==='' ?
      fncDiscountAt_(fncNow_()) : cycle['Desconto aplicado %'];
    if(pct===''||pct===null||pct===undefined)throw new Error('Definir desconto antes de validar a leitura');
    const totalActive=fncData_(FNC.sheet.fractions).filter(x=>x['Estado']==='Ativa').length;
    if(totalActive>1)throw new Error('Apuramento com várias wallboxes exige rateio conjunto; nenhuma cobrança é autorizada automaticamente');
    const fraction=fncData_(FNC.sheet.fractions).find(x=>x['ID fração']===r['Fração']&&x['Estado']==='Ativa');
    if(!fraction)throw new Error('Fração não encontrada');
    const prev=fncData_(FNC.sheet.readings).filter(x=>x['Fração']===r['Fração']&&x['Estado']==='Validada')
      .filter(x=>fncIsoDate_(x['Data leitura'])>=fncIsoDate_(fraction['Data adesão']))
      .sort((a,b)=>fncIsoDate_(a['Data leitura']).localeCompare(fncIsoDate_(b['Data leitura']))).slice(-1)[0];
    const fromDate=prev?prev['Data leitura']:fraction['Data leitura inicial'];
    const fromKwh=prev?prev['Leitura acumulada kWh']:fraction['Leitura inicial (kWh)'];
    if(fromDate===''||fromKwh==='')throw new Error('Leitura inicial não configurada');
    if(fncIsoDate_(r['Data leitura'])<=fncIsoDate_(fromDate))
      throw new Error('Leitura fora da sequência temporal');
    const inputs={
      invoiceStart:fncIsoDate_(tariffs['Início faturação']),
      invoiceEnd:fncIsoDate_(tariffs['Fim faturação']),
      buildingKwh:tariffs['Consumo prédio kWh'],
      energyExVat:tariffs['Energia sem IVA'],energyVat:tariffs['IVA energia'],
      iecExVat:tariffs['IEC sem IVA'],iecVat:tariffs['IVA IEC'],
      powerExVat:tariffs['Potência sem IVA'],powerVat:tariffs['IVA potência'],
      dgegExVat:tariffs['DGEG sem IVA'],dgegVat:tariffs['IVA DGEG'],
      avExVat:tariffs['Audiovisual sem IVA'],avVat:tariffs['IVA audiovisual'],
      readingStart:fncIsoDate_(fromDate),readingEnd:fncIsoDate_(r['Data leitura']),
      readingKwhStart:fromKwh,readingKwhEnd:r['Leitura acumulada kWh'],discountPct:pct
    };
    const calc=financeCalculateWallbox_(inputs);
    const paid=fncData_(FNC.sheet.payments).find(x=>x['Fração']===r['Fração']&&x['Ciclo EDP']===r['Ciclo EDP']);
    if(paid)throw new Error('Já existe apuramento para este ciclo e fração');
    const pid='WP-'+Utilities.getUuid();
    if(cycle['Desconto aplicado %']==='')fncUpdate_(FNC.sheet.cycles,cycle.row,{'Desconto aplicado %':pct});
    fncAppend_(FNC.sheet.payments,{'ID apuramento':pid,'Fração':r['Fração'],
      'Ciclo EDP':r['Ciclo EDP'],'Data inicial':inputs.readingStart,'Data final':inputs.readingEnd,
      'Leitura inicial kWh':fromKwh,'Leitura final kWh':r['Leitura acumulada kWh'],
      'Consumo kWh':calc.kwh,'Energia + IEC €':calc.energy,'Custos fixos antes desc. €':calc.fixedBeforeDiscount,
      'Desconto %':pct,'Parcela custos fixos €':calc.fixedAfterDiscount,'Total a pagar €':calc.total,
      'Estado':'A pagar','Emissão':fncNow_(),'Validado por':actor.email,
      'Foto leitura Drive ID':r['Foto Drive ID']});
    fncUpdate_(FNC.sheet.readings,r.row,{'Estado':'Validada','Validado por':actor.email,'Validado em':fncNow_()});
    fncAudit_(rid,'Validar leitura e apurar valor','Submetida','Validada',actor,'Apuramento '+pid);
    return {ok:true,paymentId:pid,amount:calc.total};
  });
}
function fncMarkWallboxPaid_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Apuramento');
    const row=fncData_(FNC.sheet.payments).find(r=>r['ID apuramento']===id);
    if(!row || !['A pagar','Pago'].includes(row.Estado))throw new Error('Apuramento não está a pagamento');
    if(row.Estado==='Pago')return {ok:true,alreadyPaid:true};
    fncUpdate_(FNC.sheet.payments,row.row,{'Estado':'Pago','Pago em':fncRequired_(req.paidAt,'Data recebimento'),'Recebido por':actor.email});
    fncAudit_(id,'Recebimento wallbox','A pagar','Pago',actor);
    return {ok:true};
  });
}
/**
 * Cálculo isolado e testável do simulador original.
 * O total só pode ser lançado depois de ler tarifas da fatura e desconto aprovado.
 * inputs: {energyExVat,energyVat,iecExVat,iecVat,powerExVat,powerVat,
 * dgegExVat,dgegVat,avExVat,avVat,invoiceStart,invoiceEnd,buildingKwh,
 * readingStart,readingEnd,readingKwhStart,readingKwhEnd,discountPct}
 */
function financeCalculateWallbox_(i){
  function date_(str){const d=new Date(String(str).slice(0,10)+'T12:00:00Z');if(isNaN(d.getTime()))throw new Error('Data inválida');return d;}
  const p=x=>fncNumber_(i[x]), first=date_(i.invoiceStart), last=date_(i.invoiceEnd);
  const r0=date_(i.readingStart),r1=date_(i.readingEnd);
  const invoiceDays=Math.round((last-first)/86400000)+1,days=Math.round((r1-r0)/86400000);
  const kwh=p('readingKwhEnd')-p('readingKwhStart');
  if(invoiceDays<=0||days<=0||kwh<0||p('buildingKwh')<=0)throw new Error('Períodos ou consumos inválidos');
  const dim=d=>new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+1,0)).getUTCDate();
  const months=12*(last.getUTCFullYear()-first.getUTCFullYear())+
    (last.getUTCMonth()-first.getUTCMonth())+
    last.getUTCDate()/dim(last)-(first.getUTCDate()-1)/dim(first);
  const withVat=(a,b)=>p(a)*(1+p(b));
  const fixedDaily=withVat('powerExVat','powerVat')+
    (withVat('dgegExVat','dgegVat')+withVat('avExVat','avVat'))*months/invoiceDays;
  const share=Math.min(1,(kwh/days)/(p('buildingKwh')/invoiceDays));
  const rawFixed=fixedDaily*days*share;
  const discount=p('discountPct');
  if(discount<0||discount>100)throw new Error('Desconto inválido');
  const fixed=rawFixed*(1-discount/100);
  const energy=kwh*(withVat('energyExVat','energyVat')+withVat('iecExVat','iecVat'));
  const total=Math.round((energy+fixed+Number.EPSILON)*100)/100;
  return {invoiceDays,readingDays:days,kwh,share,
    energy:Math.round(energy*100)/100,fixedBeforeDiscount:Math.round(rawFixed*100)/100,
    discountPct:discount,fixedAfterDiscount:Math.round(fixed*100)/100,total};
}

/** Fatura EDP aprovada abre um ciclo; tarifas são copiadas e imutáveis nesse ciclo. */
function fncIsoDate_(value){
  if(value instanceof Date)return Utilities.formatDate(value,'Europe/Lisbon','yyyy-MM-dd');
  const s=String(value||'').trim();if(!/^\d{4}-\d{2}-\d{2}/.test(s))throw Error('Data inválida: '+s);
  return s.slice(0,10);
}
function fncTariffMap_(){
  return Object.fromEntries(fncData_(FNC.sheet.tariff).map(x=>[String(x['Parâmetro']),x['Valor']]));
}
function fncOpenCycle_(invoice,actor){
  const id='CYCLE-'+invoice.ID, existing=fncData_(FNC.sheet.cycles).find(x=>x['ID ciclo']===id);
  if(existing)return id;
  const t=fncTariffMap_();
  const invoiceStart=fncIsoDate_(invoice['Período início']),invoiceEnd=fncIsoDate_(invoice['Período fim']);
  if(fncIsoDate_(t['Período início'])!==invoiceStart||fncIsoDate_(t['Período fim'])!==invoiceEnd)
    throw Error('Período da tabela EDP não corresponde ao documento. Rever tarifas primeiro.');
  const keys=['Consumo prédio','Energia (sem IVA)','IVA energia','IEC (sem IVA)','IVA IEC',
    'Potência (sem IVA)','IVA potência','DGEG (sem IVA)','IVA DGEG',
    'Contribuição audiovisual (sem IVA)','IVA audiovisual'];
  for(const k of keys)fncNumber_(t[k]);
  const fields={'ID ciclo':id,'Fatura ID':invoice.ID,'Início faturação':invoiceStart,
    'Fim faturação':invoiceEnd,'Consumo prédio kWh':t['Consumo prédio'],
    'Energia sem IVA':t['Energia (sem IVA)'],'IVA energia':t['IVA energia'],
    'IEC sem IVA':t['IEC (sem IVA)'],'IVA IEC':t['IVA IEC'],
    'Potência sem IVA':t['Potência (sem IVA)'],'IVA potência':t['IVA potência'],
    'DGEG sem IVA':t['DGEG (sem IVA)'],'IVA DGEG':t['IVA DGEG'],
    'Audiovisual sem IVA':t['Contribuição audiovisual (sem IVA)'],
    'IVA audiovisual':t['IVA audiovisual'],
    'Desconto %':fncDiscountAt_(fncNow_()),'Criado em':fncNow_,
    'Validado por':actor.email,'Origem EDP':invoice['PDF original']};
  fields['Criado em']=fncNow_();
  fncAppend_(FNC.sheet.rates,fields);
  fncAppend_(FNC.sheet.cycles,{'ID ciclo':id,'Fatura ID':invoice.ID,
    'Início faturação':invoiceStart,'Fim faturação':invoiceEnd,'Estado':'Aberto',
    'Validado por':actor.email,'Criado em':fncNow_(),
    'Desconto aplicado %':fncDiscountAt_(fncNow_()),
    'Observações':'Aguardam-se leituras das frações ativas.'});
  return id;
}
function fncNotifyCycle_(req,actor){
  const settings=fncSettings_();
  if(String(settings.NOTIFICACOES_ATIVAS).toUpperCase()!=='SIM')
    throw Error('Notificações ainda desativadas na configuração.');
  const id=fncRequired_(req.cycleId,'Ciclo');
  return fncLock_(()=>{
    const c=fncData_(FNC.sheet.cycles).find(r=>r['ID ciclo']===id);
    if(!c||c['Estado']!=='Aberto')throw Error('Ciclo não está aberto');
    const active=fncData_(FNC.sheet.fractions).filter(f=>f['Estado']==='Ativa'&&f['Email autorizado']&&fncApprovedResident_(f['ID fração'],f['Email autorizado']));
    const sent=fncData_(FNC.sheet.notifications).filter(n=>n['Ciclo']===id&&n['Estado']==='Enviado');
    const sentFractions=new Set(sent.map(n=>String(n['Fração'])));
    const link='https://filiperod-byte.github.io/DomingosdaCunha4/V2/wallbox.html';
    let count=0;
    for(const f of active){
      const email=String(f['Email autorizado']).trim(),fractionId=String(f['ID fração']);
      if(sentFractions.has(fractionId))continue;
      MailApp.sendEmail({to:email,subject:'Domingos da Cunha 4 — Nova leitura da wallbox',
        body:'A fatura de eletricidade do condomínio já foi validada.\n'+
          'Pode agora submeter a leitura acumulada da sua wallbox, com fotografia, na App:\n'+link+
          '\n\nO valor a transferir será calculado após validação da leitura pela administração.'});
      fncAppend_(FNC.sheet.notifications,{'ID':'WN-'+Utilities.getUuid(),'Ciclo':id,
        'Fração':fractionId,'Email':email,'Estado':'Enviado','Enviado em':fncNow_()});
      count++;
    }
    // Sem destinatários não se considera o ciclo notificado.
    if(active.length && fncData_(FNC.sheet.notifications)
      .filter(n=>n['Ciclo']===id&&n['Estado']==='Enviado').length>=active.length){
      fncUpdate_(FNC.sheet.cycles,c.row,{'Notificado em':fncNow_()});
    }
    if(count)fncAudit_(id,'Enviar pedidos de leitura','',count+' email(s)',actor);
    return {ok:true,notified:count,eligible:active.length};
  });
}

/** Ativação/desativação das notificações no backend, auditada. */
function fncSetNotifications_(req,actor){
  return fncLock_(()=>{
    const enabled=req.enabled===true || req.enabled==='true';
    const record=fncData_(FNC.sheet.settings).find(r=>r['Chave']==='NOTIFICACOES_ATIVAS');
    if(!record)throw Error('Configuração de notificações não encontrada');
    const old=record['Valor'],v=enabled?'SIM':'NAO';
    fncUpdate_(FNC.sheet.settings,record.row,{'Valor':v,'Atualizado em':fncNow_(),'Atualizado por':actor.email});
    fncAudit_('CONFIG_WALLBOX','Notificações',old,v,actor);
    return {ok:true,enabled};
  });
}

/**
 * Descoberta automática de novos documentos na Drive /FATURAS/AAAA/MM.
 * Não tenta classificar imagens/PDF digitalizados como faturas fiáveis.
 * Apenas regista novos documentos para validação humana/extração posterior.
 * Propriedade FINANCE_FATURAS_ROOT_ID = ID da pasta FATURAS privada.
 */
function financeScanFolder_(){
  const scanned=fncLock_(()=>{
    const root=PropertiesService.getScriptProperties().getProperty('FINANCE_FATURAS_ROOT_ID');
    const folder=root?DriveApp.getFolderById(root):fncPrivateFolder_('FATURAS');
    const seen=new Set(fncData_(FNC.sheet.invoices).map(r=>String(r['ID Drive']||'')));
    let count=0;
    const years=folder.getFolders();
    while(years.hasNext()){
      const year=years.next();
      if(!/^\d{4}$/.test(year.getName()))continue;
      const months=year.getFolders();
      while(months.hasNext()){
        const month=months.next();
        if(!/^(0[1-9]|1[0-2])$/.test(month.getName()))continue;
        const files=month.getFiles();
        while(files.hasNext()){
          const file=files.next(),id=file.getId();
          if(seen.has(id))continue;
          const mimetype=file.getMimeType();
          if(!['application/pdf','image/jpeg','image/png'].includes(mimetype))continue;
          const docid='DC4-DOC-'+Utilities.getUuid();
          fncAppend_(FNC.sheet.invoices,{
            'ID':docid,'Tipo':'Documento por classificar',
            'Fornecedor':'Por identificar','Descrição':file.getName(),
            'Documento':file.getName(),'Validação':'Por validar','Pagamento':'Por confirmar',
            'PDF original':'https://drive.google.com/file/d/'+id+'/view',
            'Observações':'Importado automaticamente; verificar antes de registar pagamento.',
            'ID Drive':id
          });
          seen.add(id);count++;
        }
      }
    }
    return {ok:true,newDocuments:count};
  });
  // O registo do ficheiro é independente da extração: mesmo quando o OCR falha,
  // o documento permanece consultável e não entra automaticamente no circuito de pagamentos.
  const extraction=fncExtractPendingInvoices_(2);
  return {...scanned,extraction};
}

/**
 * A consulta das faturas não depende do âmbito OAuth de gestão de triggers.
 * ScriptApp.getProjectTriggers requer script.scriptapp e pode falhar numa Web App
 * mesmo quando Drive e Sheets estão autorizados. Responder sem lançar erro:
 * é importante manter a sincronização manual e explicar a autorização em falta.
 */
function financeScannerStatus_(){
  const prop=PropertiesService.getScriptProperties();
  try{
    const handlers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='financeScanFolder_');
    const enabled=handlers.length>0;
    prop.setProperties({'FINANCE_SCANNER_ENABLED':enabled?'SIM':'NAO'},false);
    return {ok:true,enabled,count:handlers.length,verified:true,authorizationRequired:false};
  }catch(err){
    return {ok:true,enabled:prop.getProperty('FINANCE_SCANNER_ENABLED')==='SIM',
      verified:false,authorizationRequired:true,
      message:'Não é possível verificar os agendamentos sem autorização do Google Apps Script. A pesquisa manual de faturas continua disponível.'};
  }
}
function fncScannerPermissionResult_(err){
  const reason=String(err&&err.message||err||'');
  // Mensagem original apenas nos registos do editor, nunca na resposta pública da App.
  console.error('Erro ao configurar sincronização de faturas: '+reason);
  const permission=/permiss|authoriz|autoriz|scope|script\\.scriptapp|not allowed|access denied/i.test(reason);
  return {ok:false,requiresAuthorization:permission,
    message:permission?
      'Autoriza a gestão de agendamentos no editor Google Apps Script executando ativarSincronizacaoFaturas.':
      'Não foi possível gerir a monitorização. Consulta os registos de execução do Google Apps Script.'};
}
function financeDisableScannerTrigger_(){
  try{
    const handlers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='financeScanFolder_');
    handlers.forEach(t=>ScriptApp.deleteTrigger(t));
    PropertiesService.getScriptProperties().setProperties({'FINANCE_SCANNER_ENABLED':'NAO'},false);
    return {ok:true,enabled:false,removed:handlers.length};
  }catch(err){return fncScannerPermissionResult_(err);}
}

/** Criar o trigger uma única vez, após autorizar o projeto Apps Script. */
function financeInstallScannerTrigger_(){
  try{
    const matches=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='financeScanFolder_');
    if(!matches.length)ScriptApp.newTrigger('financeScanFolder_').timeBased().everyMinutes(15).create();
    PropertiesService.getScriptProperties().setProperties({'FINANCE_SCANNER_ENABLED':'SIM'},false);
    return matches.length?{ok:true,enabled:true,alreadyInstalled:true}:{ok:true,enabled:true,installed:true};
  }catch(err){return fncScannerPermissionResult_(err);}
}


// --- invoice-extraction.js ---
/**
 * Extração conservadora de dados de faturas via OCR do Google Drive.
 * Usa exclusivamente as permissões da conta que executa o Apps Script.
 * O PDF original não é modificado. Documento OCR temporário é enviado para o lixo.
 * Nunca aprova faturas nem cria pagamentos.
 */

/** Formatos portugueses (1.234,56 / 1234.56 / 1234,56). */
function fncParseEuro_(value){
  if(value===null||value===undefined)return null;
  let s=String(value).replace(/[\s\u00a0€]/g,'').replace(/[^\d.,-]/g,'');
  if(!s)return null;
  const c=s.lastIndexOf(','),p=s.lastIndexOf('.');
  if(c>=0&&p>=0)s=c>p?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');
  else if(c>=0)s=s.replace(',','.');
  else if(p>=0&&s.length-p-1===3)s=s.replace(/\./g,'');
  const n=Number(s);
  return Number.isFinite(n)&&n>=0&&n<=1000000?Math.round(n*100)/100:null;
}
function fncISOInvoiceDate_(value){
  const s=String(value||'').trim();
  let y,m,d;
  if(/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)){[y,m,d]=s.split(/[-/]/).map(Number);}
  else if(/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)){[d,m,y]=s.split(/[-/]/).map(Number);}
  else return '';
  const dt=new Date(Date.UTC(y,m-1,d));
  if(dt.getUTCFullYear()!==y||dt.getUTCMonth()!==m-1||dt.getUTCDate()!==d)return '';
  return [y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
}
function fncFirstMatch_(text,expressions){
  for(const exp of expressions){const m=text.match(exp);if(m&&m[1])return String(m[1]).trim();}
  return '';
}
/** Pura: só extrai factos com etiquetas inequívocas, nunca usa nome do PDF como valor. */
function fncParseInvoiceText_(source){
  const text=String(source||'').replace(/\r/g,'').replace(/[\u00a0]/g,' ').slice(0,80000);
  const m=text.match(/\b(?:factura|fatura)\s*(?:n[º°.o]?\s*)?(?:\d+\s+)?(\d{4}\/\d{1,10})\b/i);
  let number=m?m[1]:fncFirstMatch_(text,[
    /\b(?:Factura|Fatura)\s*(?:n[º°.]?\s*)?(FT\s*[\w/. -]*\d)\b/i,
    /\b(FT\s*\d{4}[\w/. -]*\d)\b/i
  ]);
  if(number)number=number.replace(/\s+/g,' ').slice(0,90);
  const issuance=fncISOInvoiceDate_(fncFirstMatch_(text,[
    /Data\s+(?:de\s+)?Emiss[aã]o\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i,
    /\bData\s+(?:da\s+)?Fatura\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i
  ]));
  let due=fncISOInvoiceDate_(fncFirstMatch_(text,[
    /Data\s+(?:de\s+)?Vencimento\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i,
    /Data\s+Limite\s+(?:de\s+)?Pagamento\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i
  ]));
  // Alguns PDFs extraídos em colunas listam duas datas sob os cabeçalhos
  // EMISSÃO / VENCIMENTO. Só mapear se houver ambos os valores adjacentes.
  let issued=issuance;
  if(!issued){
    const columns=text.match(/DATA\s+DE\s+EMISS[AÃ]O\s*\n\s*DATA\s+VENCIMENTO\s+(\d{4}-\d{2}-\d{2})\s*\n\s*(\d{4}-\d{2}-\d{2})/i);
    if(columns){
      due=fncISOInvoiceDate_(columns[1]);
      issued=fncISOInvoiceDate_(columns[2]);
    }
  }
  if(!issued)issued=fncISOInvoiceDate_(fncFirstMatch_(text,[
    /(?:^|\n)\s*Data\s*:\s*(\d{2}\/\d{2}\/\d{4}|\d{4}-\d{2}-\d{2})/im
  ]));
  // O OCR pode mover € para antes do valor, ou omiti-lo. Não aceitar
  // números sem uma etiqueta financeira explícita.
  const amount='(\\d{1,3}(?:[. ]\\d{3})*,\\d{2}|\\d{1,8}[,.]\\d{2})';
  function amounts(labels){
    const vals=[];
    for(const label of labels){
      const regex=new RegExp('(?:^|\\n)\\s*'+label+'[ \\t:\\n]{0,15}€?[ \\t]*'+amount,'gim');
      for(const match of text.matchAll(regex)){
        const v=fncParseEuro_(match[1]);
        if(v!==null)vals.push(v);
      }
    }
    return [...new Set(vals)];
  }
  const totals=amounts([
    'Total\\s+(?:da\\s+)?(?:Factura|Fatura)',
    'Total\\s+a\\s+Pagar','Valor\\s+Total',
    'Total\\s+(?:do\\s+)?Documento','TOTAL'
  ]);
  const ambiguousTotal=totals.length>1;
  let gross=totals.length===1?totals[0]:null;
  const net=amounts([
    'Subtotal\\s+(?:da\\s+)?(?:Factura|Fatura)',
    'Valor\\s+sem\\s+IVA','Valor\\s+il[ií]quido','Sem\\s*/\\s*IVA'
  ]);
  let subtotal=net.length===1?net[0]:null;
  const tax=text.match(/\bIVA\s*[\d.,]+\s*%\s*\(\s*Incid[eê]ncia\s*:\s*([\d.,]+)\s*\)\s*([\d.,]+)\s*€?/i);
  let iva=tax?fncParseEuro_(tax[2]):fncParseEuro_(fncFirstMatch_(text,[
    /(?:^|\n)\s*IVA\s*[:\s]*€\s*([\d.,]+)\s*(?:€|\n|$)/im
  ]));
  if(tax&&subtotal===null)subtotal=fncParseEuro_(tax[1]);
  // Orçamentos podem ter o valor "C/IVA 23%" sem outro campo de total.
  if(gross===null && !ambiguousTotal &&
    /(?:^|\n)\s*(?:Orc\.|Honorários|Documento valido por)/im.test(text)){
    gross=fncParseEuro_(fncFirstMatch_(text,[
      /C\s*\/\s*IVA\s*\d+(?:[.,]\d+)?%\s*([\d.,]+)\s*€/i
    ]));
  }
  // Só calcular a soma quando base e IVA estão ambos identificados no PDF.
  if(gross===null&&!ambiguousTotal&&subtotal!==null&&iva!==null)
    gross=Math.round((subtotal+iva)*100)/100;
  if(gross!==null&&subtotal!==null&&iva!==null&&Math.abs(gross-subtotal-iva)>0.03){
    gross=null;iva=null;subtotal=null;
  }
  // Primeiro a tipologia: uma nota ou orçamento não é automaticamente fatura.
  const fold=text.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const documentType=/nota\s+de\s+despesas\s+e\s+honorarios/.test(fold)?
    'Nota despesas/honorários':
    /(?:\borc\.?\s*:|documento\s+valido\s+por\s+\d+\s+dias|\borcamento\b)/.test(fold)?
      'Orçamento / honorários':
      /(?:\bfactura\b|\bfatura\b|\bFT\s*\d{4}[/_-]\d+)/i.test(text)?
        'Fatura':'Documento por classificar';
  if(!number)number=fncFirstMatch_(text,[/\b(FT\s*\d{4}\s*[/_-]\s*\d+)\b/i])
    .replace(/\s+/g,' ').replace(/\s*\/\s*/g,'/');

  const issuer=text.split(/\n\s*Para\s*\n/i)[0];
  const lines=issuer.split('\n').map(x=>x.trim()).filter(Boolean);
  const company=lines.find(x=>/\b(?:Lda\.?|Limitada|S\.?\s*A\.?)\b/i.test(x)&&x.length<170)||'';
  // Nome empresarial até Lda/S.A.; morada do fornecedor pode estar na mesma linha OCR.
  let vendor=company?(company.match(/^(.{3,130}?\b(?:Lda|Limitada|S\.?\s*A\.?)\.?)(?:\s|$)/i)||[])[1]||company:'';
  if(!vendor&&documentType==='Nota despesas/honorários'){
    const pos=lines.findIndex(x=>/^Advogados$/i.test(x));
    if(pos>=2)vendor=lines[pos-2]+' / '+lines[pos-1];
  }
  vendor=vendor.trim().slice(0,150);
  let nif=fncFirstMatch_(issuer,[
    /N[ºo°.]?\s*Contribuinte\s*:\s*(\d{9})\b/i,
    /\b(?:NIF|NIPC)\s*:\s*(?:PT\s*)?(\d{9})\b/i,
    /Cons\.\s*Reg\.\s*Comercial\s*:\s*[^\n]*?(\d{9})\b/i
  ]);
  if(!nif&&documentType==='Orçamento / honorários')
    nif=fncFirstMatch_(text,[
      /(?:NIF|NIPC)\s*:\s*(?:PT\s*)?(\d[\d ]{8,14})\b/i
    ]).replace(/\s/g,'').slice(0,9);

  const article=text.match(/Lista\s+de\s+Artigos\s*[\s\S]*?\n[^\n]*(?:DESCRI[CÇ][AÃ]O|ARTIGO)[^\n]*\n\s*([^\n]{8,220})/i);
  let description=article?article[1].trim():'';
  if(description)description=description
    .replace(/\s+\d+(?:\.\d+)?\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s*$/,'')
    .slice(0,280);
  if(!description&&documentType==='Fatura'){
    const lines=text.match(/C[ÓO]DIGO\s+DESCRI[CÇ][AÃ]O[^\n]*\n\s*([^\n]{5,130})\n([^\n]{0,130})/i);
    if(lines)description=(lines[1]+' '+lines[2]).replace(/\s+\d+(?:[.,]\d+)?\s+un\b.*$/i,'').trim().slice(0,280);
  }
  if(!description&&documentType==='Nota despesas/honorários')
    description=fncFirstMatch_(text,[/Assunto\s*:\s*([^\n]{3,180})/i]);
  if(!description&&documentType==='Orçamento / honorários')
    description=fncFirstMatch_(text,[
      /(?:^|\n)\s*\d+(?:\.\d+)?\s+([^\n]{8,160}?)\s+un\s+\d+\s+[\d.,]+\s*€/im
    ]);
  if(documentType==='Orçamento / honorários'&&gross===null){
    gross=fncParseEuro_(fncFirstMatch_(text,[/C\s*\/\s*IVA\s*\d+(?:[.,]\d+)?%\s*([\d.,]+)\s*€/i]));
  }
  const ibanRaw=fncFirstMatch_(text,[
    /\b(PT50(?:[ ]*\d){21})\b/i
  ]).replace(/\s/g,'');
  const iban=/^PT50\d{21}$/.test(ibanRaw)?ibanRaw:'';
  const method=/D[eé]bito\s+Diret[oa]/i.test(text)?'Débito direto':'Por confirmar';
  const strong=Boolean(vendor&&gross!==null&&number&&issuance&&documentType==='Fatura');
  return {supplier:vendor,nif,number,issued,due,value:gross,iva,subtotal,description,method,
    iban,documentType,state:strong?'Extraído — por validar':'Extração parcial — por validar',
    recognized:[vendor,nif,number,issuance,due,gross!==null?String(gross):'',description].filter(Boolean).length};
}
/**
 * Converte PDF/imagem para Docs com OCR da Drive API v3 e exporta texto simples.
 * Não requer serviço avançado "Drive", mas usa UrlFetchApp + token da conta Apps Script.
 */
function fncDriveOcrText_(fileId){
  const file=DriveApp.getFileById(fileId);
  const mime=file.getMimeType();
  if(!['application/pdf','image/jpeg','image/png'].includes(mime))
    throw new Error('Formato não suportado para extração');
  const blob=file.getBlob(),media=blob.getBytes();
  if(media.length>7*1024*1024)throw new Error('PDF acima de 7 MB — revisão manual');
  const token=ScriptApp.getOAuthToken(),boundary='dc4ocr'+Utilities.getUuid().replace(/-/g,'');
  const metadata=JSON.stringify({name:'DC4-OCR-TEMP-'+Utilities.getUuid(),
    mimeType:'application/vnd.google-apps.document'});
  const head='--'+boundary+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+metadata+
    '\r\n--'+boundary+'\r\nContent-Type: '+mime+'\r\n\r\n';
  const tail='\r\n--'+boundary+'--\r\n';
  const bytes=Utilities.newBlob(head).getBytes().concat(media,Utilities.newBlob(tail).getBytes());
  const options={method:'post',contentType:'multipart/related; boundary='+boundary,
    headers:{Authorization:'Bearer '+token},payload:bytes,muteHttpExceptions:true};
  const upload=UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=pt&fields=id,mimeType',
    options);
  if(upload.getResponseCode()<200||upload.getResponseCode()>=300)
    throw new Error('Conversão do PDF indisponível (Google Drive HTTP '+upload.getResponseCode()+
      '). Verificar OAuth Drive e UrlFetchApp no editor.');
  let tempId='';
  try{tempId=JSON.parse(upload.getContentText()).id;}catch(_){}
  if(!tempId)throw new Error('Google Drive não devolveu documento OCR');
  try{
    const txt=UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(tempId)+'/export?mimeType=text%2Fplain',
      {headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
    if(txt.getResponseCode()!==200)throw new Error('A exportação OCR não foi concluída');
    const content=txt.getContentText('UTF-8');
    if(!String(content).trim())throw new Error('Documento sem texto reconhecido');
    return content;
  }finally{
    try{DriveApp.getFileById(tempId).setTrashed(true);}catch(e){console.warn('OCR temporário não foi eliminado: '+e.message);}
  }
}
/** Aplica somente dados reconhecidos a documento por validar. */

/** Acrescenta só os campos técnicos ausentes, sem alterar os já existentes. */
function fncEnsureExtractionColumns_(){
  const sheet=fncTab_(FNC.sheet.invoices);
  const cols=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
  const extras=['IVA (€)','Base tributável (€)','Estado extração','Data extração'].filter(c=>!cols.includes(c));
  if(extras.length)sheet.getRange(1,sheet.getLastColumn()+1,1,extras.length).setValues([extras]);
}
/** Processa um pequeno lote de faturas não validadas; não precisa de trigger adicional. */
function fncExtractPendingInvoices_(maxInvoices){
  const count=Math.min(2,Math.max(0,Number(maxInvoices)||0));
  if(!count)return {processed:0,extracted:0,failed:0};
  const pending=fncData_(FNC.sheet.invoices)
    .filter(r=>{
      if(r['Validação']!=='Por validar'||r['Pagamento']==='Pago'||!r['ID Drive'])return false;
      const state=String(r['Estado extração']||'');
      if(/^Falhou/i.test(state))return false;
      if(!state)return true;
      // Reprocessar uma única vez os PDFs que a versão anterior interpretou mal.
      // Depois marcar "Leitor v2" para não repetir OCR a cada 15 minutos.
      return !String(r['Observações']||'').includes('Leitor v2') &&
        (!r['Total (€)']||r['Fornecedor']==='Por identificar'||
          /\b(?:Rua|Impasse|Avenida)\b/i.test(String(r['Fornecedor']||'')));
    })
    .slice(0,count);
  if(!pending.length)return {processed:0,extracted:0,failed:0};
  fncEnsureExtractionColumns_();
  let success=0,failed=0;
  for(const invoice of pending){
    try{
      const r=fncExtractOneInvoice_(invoice.ID);
      if(r.ok)success++;
    }catch(err){
      failed++;
      console.error('Falha na extração de '+invoice.ID+': '+String(err.message||err));
      // Falhas não bloqueiam novas faturas nem repetem OCR em cada abertura.
      fncLock_(()=>{
        const row=fncData_(FNC.sheet.invoices).find(x=>x.ID===invoice.ID);
        if(row && row['Validação']==='Por validar')
          fncUpdate_(FNC.sheet.invoices,row.row,{
            'Estado extração':'Falhou — revisão manual',
            'Data extração':fncNow_()
          });
      });
    }
  }
  return {processed:pending.length,extracted:success,failed};
}
function fncExtractOneInvoice_(invoiceId){
  const original=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===String(invoiceId));
  if(!original)throw new Error('Documento não encontrado');
  if(original['Validação']==='Validada'||original['Pagamento']==='Pago')throw new Error('Documento fechado');
  const fileId=original['ID Drive'];
  if(!fileId)throw new Error('Documento sem ID de Drive');
  // O OCR não mantém bloqueios globais enquanto aguarda o Google Drive.
  fncEnsureExtractionColumns_();
  const parsed=fncParseInvoiceText_(fncDriveOcrText_(String(fileId)));
  return fncLock_(()=>{
    // Conferir novamente o registo ANTES da escrita; pode ter sido validado enquanto o OCR corria.
    const item=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===String(invoiceId));
    if(!item||item['Validação']==='Validada'||item['Pagamento']==='Pago')
      throw new Error('Documento foi entretanto validado; extração não aplicada');
    if(String(item['ID Drive'])!==String(fileId))throw new Error('Documento original alterado');
    const patch={'Estado extração':parsed.state,'Data extração':fncNow_()};
  const prior=String(item['Estado extração']||'');
  const previousOCR=/^(Extra[ií]d|Extração parcial|Falhou)/i.test(prior);
  const set=(field,value,placeholder)=>{
    if(value===null||value===undefined||value==='')return;
    const current=item[field];
    const existing=String(current==null?'':current);
    const empty=current===''||current==null||(field==='Total (€)'&&current===0);
    // Só corrigir leituras anteriores da própria App com morada colada ao
    // fornecedor, nunca substituir nomes que o administrador corrigiu.
    const contaminated=field==='Fornecedor'&&previousOCR&&
      existing.startsWith(String(value)+' ')&&
      /\b(?:Rua|Impasse|Largo|Av\.|Avenida|Pra[cç]a|Travessa)\b/i.test(existing.slice(String(value).length));
    const filePlaceholder=field==='Descrição'&&/\.(?:pdf|png|jpe?g)$/i.test(existing);
    if(empty||contaminated||filePlaceholder||
       (placeholder&&existing===String(placeholder)))patch[field]=value;
  };
  if(parsed.recognized>=2){
    set('Tipo',parsed.documentType,'Documento por classificar');
    set('Fornecedor',parsed.supplier,'Por identificar');
    set('NIF',parsed.nif);
    // Número/documento já corrigido pelo administrador é preservado.
    set('Documento',parsed.number,item['Descrição']);
    set('Emissão',parsed.issued);
    set('Vencimento',parsed.due);
    set('Total (€)',parsed.value);
    if(parsed.method==='Débito direto')set('Método',parsed.method,'Por confirmar');
    set('Descrição',parsed.description,item['Documento']);
    set('IVA (€)',parsed.iva);
    set('Base tributável (€)',parsed.subtotal);
    if(parsed.iban && fncValidIban_(parsed.iban))set('IBAN fornecedor',parsed.iban);
  }
  const details=['Leitor v2: dados extraídos do PDF — CONFIRMAR antes de validar ou pagar'];
  if(parsed.value!==null)details.push('total '+parsed.value.toFixed(2)+' €');
  if(parsed.iva!==null)details.push('IVA '+parsed.iva.toFixed(2)+' €');
  if(parsed.iban)details.push('IBAN observado no documento — confirmar titularidade no banco');
  if(parsed.documentType==='Orçamento / honorários')
    details.push('Orçamento: não tratar como fatura pagável');
  if(parsed.value===null)details.push('Total por apurar — revisão manual necessária');
  const oldNotes=String(item['Observações']||'');
  patch['Observações']=oldNotes && !/^Importado automaticamente|^Extração automática|^Dados extraídos/i.test(oldNotes)?
    oldNotes+'; '+details.join('; '):details.join('; ');
  const keys=fncTab_(FNC.sheet.invoices).getRange(1,1,1,fncTab_(FNC.sheet.invoices).getLastColumn())
    .getValues()[0].map(String);
  for(const k of Object.keys(patch))if(!keys.includes(k))delete patch[k];
  fncUpdate_(FNC.sheet.invoices,item.row,patch);
  return {ok:true,id:item.ID,extraction:parsed.state,recognized:parsed.recognized,amount:parsed.value};
  });
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
  const startedAt=Date.now();
  const result = createAppsScriptApplication_().dispatch(method, action, payload);
  if(action === 'general.status' || (action === 'status' && payload.details === 'public') || result.token && ['garage.loginPin','garage.loginAdmin','garageLoginPin','garageLoginAdmin'].includes(action)){
    result.serviceVersion='3.7.2';
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

// Executar no editor após publicar o serviço 3.7.2. Pode ser repetido.
function migrateSecurity372() {
  const domain = createCondominiumDomain_(GARAGE_RESIDENTIAL_STRUCTURE, OPEN_STATUSES, PENDING_STATUSES);
  createAppsScriptPorts_(CONFIG, REQUIRED_HEADERS, domain).migrateSecurity372();
}

// Executar uma vez para acrescentar apenas a folha de ocorrências gerais.
function setupGeneralOccurrences() {
  const domain=createCondominiumDomain_(GARAGE_RESIDENTIAL_STRUCTURE,OPEN_STATUSES,PENDING_STATUSES);
  createAppsScriptPorts_(CONFIG,REQUIRED_HEADERS,domain).initializeGeneral();
  Logger.log('Folha OCORRENCIAS_GERAIS preparada. Dados existentes preservados.');
}


/**
 * Configuração pontual pelo PROPRIETÁRIO no editor do Google Apps Script.
 * O sufixo "_" da função interna impede a sua apresentação no seletor
 * "Executar". Este wrapper público não é chamado pela interface HTTP.
 * Repetir a execução não cria triggers duplicados.
 */
function ativarSincronizacaoFaturas() {
  // Pedir autorização no EDITOR antes de entrar no instalador. Se faltar consentimento,
  // o Apps Script interrompe esta execução e mostra o ecrã de permissões.
  // Não executar isto em pedidos do navegador: é um comando manual de configuração.
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/script.scriptapp'
  ]);
  // Compatibilidade: alguns projetos Apps Script tinham o helper sem "_".
  // É feita a deteção sem invocar funções inexistentes (evita ReferenceError).
  const installer=typeof financeInstallScannerTrigger_==='function'
    ? financeInstallScannerTrigger_
    : (typeof financeInstallScannerTrigger==='function'
        ? financeInstallScannerTrigger : null);
  if(!installer)throw new Error(
    'Instalador de faturas não encontrado. Atualiza o backend Apps Script.');
  const result=installer();
  if(!result || !result.ok) {
    const message=result && result.message || 'Não foi possível ativar a sincronização.';
    Logger.log(message);
    throw new Error(message);
  }
  Logger.log(result.alreadyInstalled ?
    'Sincronização de faturas já estava ativa (15 em 15 minutos).' :
    'Sincronização de faturas ativada (15 em 15 minutos).');
  return result;
}
