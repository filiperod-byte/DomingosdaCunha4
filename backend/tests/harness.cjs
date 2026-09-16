// Emulador mínimo apenas para testes. Nunca comunica com Google ou com o endpoint publicado.
const vm = require('node:vm');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const FIXED = '2026-09-11T10:00:00.000Z';

function harness(source, options = {}) {
  const counts = { reads: 0, opens: 0, writes: 0, mails: 0, locks: 0, releases: 0, files: 0 };
  const sheets = new Map();
  const emails = [];
  const uploads = [];
  class FixedDate extends Date {
    constructor(...args) { super(...(args.length ? args : [FIXED])); }
    static now() { return new Date(FIXED).getTime(); }
  }
  class Sheet {
    constructor(rows = []) { this.rows = rows.map(row => row.slice()); }
    getLastColumn() { return Math.max(0, ...this.rows.map(row => row.length)); }
    getLastRow() { return this.rows.length; }
    setFrozenRows() { return this; }
    appendRow(row) { counts.writes++; this.rows.push(row.slice()); return this; }
    getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
    getRange(row, column, height = 1, width = 1) {
      if (![row, column, height, width].every(value => Number.isInteger(value) && value > 0)) throw new Error('Intervalo inválido.');
      const self = this;
      return {
        getValues() { counts.reads++; return Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => self.rows[row - 1 + i]?.[column - 1 + j] ?? '')); },
        setValues(values) {
          counts.writes++;
          values.forEach((valuesRow, i) => valuesRow.forEach((value, j) => {
            self.rows[row - 1 + i] ||= [];
            self.rows[row - 1 + i][column - 1 + j] = value;
          }));
          return this;
        },
        setValue(value) { return this.setValues([[value]]); },
        setFontWeight() { return this; }
      };
    }
  }
  let uuid = 0;
  let fileId = 0;
  const props = {};
  const folders = new Map();
  const root = {
    getFoldersByName(name) { return { hasNext: () => folders.has(name), next: () => folders.get(name) }; },
    createFolder(name) {
      const folder = { createFile(blob) {
        counts.files++;
        const id = 'TEST-FILE-' + (++fileId);
        uploads.push({ folder: name, id, type: blob.type, name: blob.name, bytes: Array.from(blob.bytes) });
        return { setSharing() {}, getId: () => id, getUrl: () => 'https://example.invalid/file/' + id, getName: () => blob.name };
      } };
      folders.set(name, folder);
      return folder;
    }
  };
  const context = vm.createContext({
    Date: FixedDate, Math: Object.assign(Object.create(Math), { random: () => 0.42 }),
    Logger: { log() {} },
    SpreadsheetApp: { openById() {
      counts.opens++;
      return { getSheetByName: name => sheets.get(name) || null, insertSheet: name => {
        counts.writes++;
        const sheet = new Sheet(); sheets.set(name, sheet); return sheet;
      } };
    } },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: key => props[key] ?? null,
      setProperties(values, clear) {
        if (clear) Object.keys(props).forEach(key => delete props[key]);
        Object.assign(props, values);
      },
      deleteProperty: key => delete props[key]
    }) },
    LockService: { getScriptLock: () => ({ waitLock() { counts.locks++; }, releaseLock() { counts.releases++; } }) },
    MailApp: { sendEmail(email) {
      if (options.mailFails) throw new Error('Synthetic mail failure');
      counts.mails++; emails.push(JSON.parse(JSON.stringify(email)));
    } },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'owner@example.invalid' }), getActiveUser: () => ({ getEmail: () => '' }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: value => ({ setMimeType: () => JSON.parse(value) }) },
    DriveApp: { getFolderById: () => root, Access: { ANYONE_WITH_LINK: 'link' }, Permission: { VIEW: 'view' } },
    Utilities: {
      getUuid: () => (++uuid).toString().padStart(8, '0') + '-0000-4000-8000-000000000000',
      formatDate(date, timezone, pattern) {
        const d = new Date(date);
        const stamp = d.toISOString();
        if (pattern === 'yyyyMMddHHmmss') return stamp.replace(/[-:T]/g, '').slice(0, 14);
        if (pattern === 'yyyyMMdd_HHmmss') return stamp.slice(0, 10).replace(/-/g, '') + '_' + stamp.slice(11, 19).replace(/:/g, '');
        return stamp;
      },
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (algorithm, value) => [...crypto.createHash(algorithm).update(value).digest()].map(value => value > 127 ? value - 256 : value),
      base64Decode: value => [...Buffer.from(value, 'base64')],
      newBlob: (bytes, type, name) => ({ bytes, type, name })
    }
  });
  // Configuração sintética mesmo quando se executa o ficheiro original.
  source = source.replace(/SPREADSHEET_ID:\s*'[^']*'/, "SPREADSHEET_ID: 'TEST-SHEET'")
    .replace(/ROOT_FOLDER_ID:\s*'[^']*'/, "ROOT_FOLDER_ID: 'TEST-DRIVE'")
    .replace(/ADMIN_EMAIL:\s*'[^']*'/, "ADMIN_EMAIL: 'admin@example.invalid'");
  if (options.legacyCore) source = source.replace('createSecureApplication_(createCondominiumApplication_(ports, CONFIG, domain), ports)', 'createCondominiumApplication_(ports, CONFIG, domain)');
  vm.runInContext(source, context);
  const post = (action, data = {}) => context.doPost({ postData: { contents: JSON.stringify({ action, ...data }) } });
  const get = action => context.doGet({ parameter: { action } });
  function initialize() {
    // A implementação antiga tinha duas rotinas de preparação separadas.
    if (context.setupIfNeeded_) context.setupIfNeeded_();
    context.setupApp();
  }
  function snapshot() {
    return JSON.parse(JSON.stringify({
      sheets: Object.fromEntries([...sheets].sort(([a], [b]) => a.localeCompare(b)).map(([name, sheet]) => [name, sheet.rows])),
      emails, uploads, props
    }));
  }
  return { context, counts, sheets, props, get, post, initialize, snapshot };
}

function bundled() {
  return fs.readFileSync(path.join(__dirname, '../../V2/backend-unificado.gs'), 'utf8');
}
module.exports = { harness, bundled, FIXED };
