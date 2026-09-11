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
