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

// Only accept a published source found in the trusted first-parent main history.
// After merging a PR, "git log -1 -- <file>" can pick an intermediate feature-branch
// commit that was never live. That falsely rejects legitimate deployments.
// Preserve all drift checks: unknown live code and unrelated drafts still abort.
const commit = execFileSync('git', ['rev-parse', 'HEAD'], {encoding:'utf8'}).trim();
const mainRevisions = execFileSync('git', ['rev-list', '--first-parent', '--max-count=12', 'HEAD'], {encoding:'utf8'})
  .trim().split(String.fromCharCode(10)).filter(Boolean);
const approved = new Set([next]);
for (const revision of mainRevisions) {
  try {
    const source = execFileSync('git', ['show', revision + ':' + target], {encoding:'utf8',stdio:['ignore','pipe','ignore']});
    approved.add(normalize(source));
  } catch (_) { /* Revision predates the backend file. */ }
}
console.log('A validar implementação face a ' + approved.size + ' versões da main.');
const credentials = JSON.parse(process.env.CLASPRC_JSON || '{}').tokens?.default;
assert.ok(credentials?.refresh_token && credentials.client_id && credentials.client_secret,
  'Credencial clasp incompleta. Verifique CLASPRC_JSON.');
const auth = await fetch('https://oauth2.googleapis.com/token', {method:'POST',
  body:new URLSearchParams({grant_type:'refresh_token', refresh_token:credentials.refresh_token,
    client_id:credentials.client_id, client_secret:credentials.client_secret})});
if (!auth.ok) throw Error('Falha ao renovar a autorização Google: HTTP ' + auth.status);
const {access_token: token} = await auth.json();
async function api(path, method='GET', body) {
  const url = new URL('https://script.googleapis.com/v1/projects/' + project + path);
  if (method === 'GET') url.searchParams.set('fields', '*');
  const response = await fetch(url, {
    cache:'no-store',
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
const candidates = live.files.filter(f=>f.type === 'SERVER_JS' && approved.has(normalize(f.source)));
assert.equal(candidates.length,1,'Código publicado diferente da base revista. Publicação suspensa.');
const name = candidates[0].name;
function canonical(files, replace) {
  return files.map(({name:n,type,source})=>[n,type,replace && n===name && type==='SERVER_JS' ? next : normalize(source)])
    .sort((a,b)=>JSON.stringify(a.slice(0,2)).localeCompare(JSON.stringify(b.slice(0,2))));
}
const equal = (a,b)=>JSON.stringify(a)===JSON.stringify(b);
const desired = canonical(live.files,true);
function draftChanges(liveFiles,editorFiles) {
  const key = f => JSON.stringify([f.name,f.type]);
  const published = new Map(liveFiles.map(f=>[key(f),f]));
  const editor = new Map(editorFiles.map(f=>[key(f),f]));
  return [...new Set([...published.keys(),...editor.keys()])].sort().flatMap(k=>{
    const before=published.get(k),after=editor.get(k);
    if(!before)return [{name:after.name,type:after.type,change:'ADICIONADO'}];
    if(!after)return [{name:before.name,type:before.type,change:'REMOVIDO'}];
    return normalize(before.source)===normalize(after.source)?[]:
      [{name:before.name,type:before.type,change:'MODIFICADO'}];
  });
}
if (!equal(canonical(draft.files),desired)) {
  const differences=draftChanges(live.files,draft.files);
  if(differences.length) {
    // Diagnóstico seguro: apenas nomes e estado; nunca imprimir código,
    // tokens, valores de configuração ou conteúdo de ficheiros.
    console.error('DIVERGÊNCIA NO EDITOR — '+differences.length+' ficheiro(s) por conciliar:');
    for(const item of differences)
      console.error(' - '+item.change+': '+JSON.stringify(item.name)+' ('+item.type+')');
  }
  assert.ok(differences.length===0,'Existem alterações por publicar no editor.');
  await api('/content','PUT',{files:draft.files.map(({name:n,type,source})=>({name:n,type,source:n===name && type==='SERVER_JS'?next:source}))});
}
if (!equal(canonical(live.files),desired)) {
  const {versionNumber} = await api('/versions','POST',{description:'GitHub ' + commit.slice(0,12)});
  const saved = await api('/content?versionNumber=' + versionNumber);
  assert.ok(equal(canonical(saved.files),desired),'A versão criada diverge da versão pretendida.');
  const current = (await api(deploymentPath)).deploymentConfig.versionNumber;
  if (current !== config.versionNumber) {
    const active = await api('/content?versionNumber=' + current);
    assert.ok(equal(canonical(active.files),desired),'A implementação mudou para código diferente durante a publicação.');
    console.log('A versão pretendida já está ativa.');
    process.exit(0);
  }
  await api(deploymentPath,'PUT',{deploymentConfig:{versionNumber, manifestFileName:config.manifestFileName,description:config.description}});
  for (let attempt = 0; attempt < 12; attempt++) {
    if ((await api(deploymentPath)).deploymentConfig.versionNumber === versionNumber) break;
    if (attempt === 11) throw Error('Versão publicada por confirmar.');
    await new Promise(resolve => setTimeout(resolve, 5000));
  }
}
console.log('Implementação verificada no URL existente.');
