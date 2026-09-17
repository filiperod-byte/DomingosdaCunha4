const test=require('node:test');
const assert=require('node:assert/strict');
const {harness,bundled}=require('./harness.cjs');
function setup(){
 const b=harness(bundled());b.initialize();
 const sheet=[...b.sheets.values()].find(s=>s.rows[0].includes('LAST_EVENT'));
 function add(patch){const record={STATUS:'ABERTA',FLOOR:9,POINT:'E1',REASON:'Extintor em falta',REPORTED_AT:'2026-09-14T10:00:00.000Z',REPORTED_BY:'Pessoa privada',NOTES:'Manómetro sem pressão',PHOTO_FILE_URL:'https://example.invalid/private',...patch};sheet.appendRow(sheet.rows[0].map(k=>record[k]??''));}
 return {b,add,details:()=>b.context.doGet({parameter:{action:'status',details:'public'}})};
}
test('consulta pública fornece motivo e data, descrição, sem autor ou fotografias',()=>{
 const {b,add,details}=setup();add({});
 const result=details();assert.equal(result.publicDetailsVersion,2);
 assert.deepEqual(result.reported,[{floor:9,point:'E1',reason:'Extintor em falta',description:'Manómetro sem pressão',reportedAt:'2026-09-14T10:00:00.000Z'}]);
 assert.doesNotMatch(JSON.stringify(result),/Pessoa privada|private/);
 assert.deepEqual(b.get('status').reported,[{floor:9,point:'E1'}]);
});
test('texto livre não é publicado como motivo e pendentes não aparecem',()=>{
 const {add,details}=setup();add({REASON:'Nome e telefone privados'});add({POINT:'E2',STATUS:'PENDENTE_VALIDACAO'});
 const result=details();assert.equal(result.reported.length,1);assert.equal(result.reported[0].reason,'Outra anomalia');assert.doesNotMatch(JSON.stringify(result),/telefone/);
});
test('diagnóstico público identifica serviço e tempo sem credenciais',()=>{
 const {details}=setup();const r=details();
 assert.equal(r.serviceVersion,'3.6.1-rc1');assert.ok(r.serverDurationMs>=0);
 assert.doesNotMatch(JSON.stringify(r),/ADMIN_PIN|SPREADSHEET_ID|token/);
});
