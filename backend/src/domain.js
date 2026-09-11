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
