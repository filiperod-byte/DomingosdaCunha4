// Only run with production credentials on the trusted main branch.
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project = '1yWP-0b4jnZ6AYgSfm-iVAbS6EyNzIgiIhnDNA4KXb1vOSWhQ4r_29m-f';
const endpoint = new URL(JSON.parse(readFileSync('V2/config.json')).backendUrl);
assert.equal(endpoint.origin, 'https://script.google.com');
const deploymentId = endpoint.pathname.match(/^\/macros\/s\/([\w-]+)\/exec$/)?.[1];
assert.ok(deploymentId, 'URL de implementação inválido.');
const target = 'V2/backend-unificado.gs';
const normalize = source => source.replace(/\r\n/g, '\n').trim();
const next = normalize(readFileSync(target, 'utf8'));

// Accept the previous committed bundle or the new bundle (retry after interrupted upload).
const commit = execFileSync('git', ['log', '-1', '--format=%H', '--', target], {encoding:'utf8'}).trim();
const previous = normalize(execFileSync('git', ['show', commit + '^:' + target], {encoding:'utf8'}));
const credentials = JSON.parse(process.env.CLASPRC_JSON || '{}').tokens?.default;
assert.ok(credentials?.refresh_token && credentials.client_id && credentials.client_secret,
  'Credencial clasp incompleta. Verifique CLASPRC_JSON.');
const auth = await fetch('https://oauth2.googleapis.com/token', {method:'POST',
  body:new URLSearchParams({grant_type:'refresh_token', refresh_token:credentials.refresh_token,
    client_id:credentials.client_id, client_secret:credentials.client_secret})});
if (!auth.ok) throw Error('Falha ao renovar a autorização Google: HTTP ' + auth.status);
const {access_token: token} = await auth.json();
async function api(path, method='GET', body) {
  const response = await fetch('https://script.googleapis.com/v1/projects/' + project + path, {
    method, headers:{Authorization:'Bearer ' + token, 'Content-Type':'application/json'},
    body:body && JSON.stringify(body)});
  if (!response.ok) throw Error(method + ' ' + path.split('?')[0] + ' HTTP ' + response.status);
  return response.json();
}
const deploymentPath = '/deployments/' + deploymentId;
const deployment = await api(deploymentPath);
assert.ok(deployment.entryPoints?.some(e=>e.webApp?.url === endpoint.href), 'Implementação não corresponde à app.');
const config = deployment.deploymentConfig;
assert.ok(config.versionNumber, 'Implementação sem versão.');
const live = await api('/content?versionNumber=' + config.versionNumber);
const draft = await api('/content');
const candidates = live.files.filter(f=>f.type === 'SERVER_JS' && [previous,next].includes(normalize(f.source)));
assert.equal(candidates.length,1,'Código publicado diferente da base revista. Publicação suspensa.');
const name = candidates[0].name;
function canonical(files, replace) {
  return files.map(({name:n,type,source})=>[n,type,replace && n===name && type==='SERVER_JS' ? next : normalize(source)])
    .sort((a,b)=>JSON.stringify(a.slice(0,2)).localeCompare(JSON.stringify(b.slice(0,2))));
}
const equal = (a,b)=>JSON.stringify(a)===JSON.stringify(b);
const desired = canonical(live.files,true);
if (!equal(canonical(draft.files),desired)) {
  assert.ok(equal(canonical(draft.files),canonical(live.files)),'Existem alterações por publicar no editor.');
  await api('/content','PUT',{files:draft.files.map(({name:n,type,source})=>({name:n,type,source:n===name && type==='SERVER_JS'?next:source}))});
}
if (!equal(canonical(live.files),desired)) {
  const {versionNumber} = await api('/versions','POST',{description:'GitHub ' + commit.slice(0,12)});
  const saved = await api('/content?versionNumber=' + versionNumber);
  assert.ok(equal(canonical(saved.files),desired),'A versão criada diverge da versão pretendida.');
  assert.equal((await api(deploymentPath)).deploymentConfig.versionNumber,config.versionNumber,'A implementação foi alterada durante a publicação.');
  await api(deploymentPath,'PUT',{deploymentConfig:{versionNumber, manifestFileName:config.manifestFileName,description:config.description}});
  assert.equal((await api(deploymentPath)).deploymentConfig.versionNumber,versionNumber,'Versão publicada por confirmar.');
}
console.log('Implementação verificada no URL existente.');
