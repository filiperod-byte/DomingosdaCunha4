const test=require('node:test'),assert=require('node:assert/strict');
const {harness,bundled}=require('./harness.cjs');
test('login lê moradores uma vez, mede etapas e vê bloqueios no pedido seguinte',()=>{
 const b=harness(bundled());b.initialize();
 const s=b.sheets.get('CONDOMINOS');
 const row={ID:'r1',Nome:'Teste',Piso:'9',Fracao:'B',Estado:'APROVADO',PIN:'654321',PINAtivo:true};
 s.rows.push(s.rows[0].map(h=>row[h]??''));
 const before=b.counts.reads;
 const r=b.post('garage.loginPin',{pin:'654321'});
 assert.ok(r.token);assert.equal(b.counts.reads-before,3); // lista + cabeçalhos/linha da atualização
 assert.equal(r.serviceVersion,'3.6.1-rc1');assert.ok(r.serverDurationMs>=0);
 assert.deepEqual(Object.keys(r.loginTimingsMs).sort(),['accessUpdateMs','rateLimitMs','sessionMs','validationMs']);
 assert.ok(Object.values(r.loginTimingsMs).every(v=>typeof v==='number'&&v>=0));
 s.rows[1][s.rows[0].indexOf('Estado')]='BLOQUEADO';
 assert.equal(b.post('garage.loginPin',{pin:'654321'}).code,'INVALID_CREDENTIALS');
 assert.equal(b.post('auth.session',{token:r.token}).code,'AUTH_REQUIRED');
});
