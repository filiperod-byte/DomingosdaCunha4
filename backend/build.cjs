// Node.js, sem dependências externas. Gera o único ficheiro a copiar para Apps Script.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const files = [
  'config.js', 'domain.js', 'notifications.js', 'occurrences.js', 'legacy-pin.js',
  'residents.js', 'accesses.js', 'application.js', 'security.js', 'apps-script-ports.js', 'entrypoints.js'
];
const banner = '// GERADO por node backend/build.cjs. Editar backend/src, não este ficheiro.\n'
  + '// Refatoração de compatibilidade: consultar backend/README.md antes de publicar.\n';
const output = banner + files.map(name => '\n// --- ' + name + ' ---\n' + fs.readFileSync(path.join(__dirname, 'src', name), 'utf8')).join('\n');
new vm.Script(output, { filename: 'backend-unificado.gs' });
const destination = path.join(__dirname, '..', 'V2', 'backend-unificado.gs');
if (process.argv.includes('--check')) {
  if (!fs.existsSync(destination) || fs.readFileSync(destination, 'utf8') !== output) {
    console.error('Bundle desatualizado. Executar node backend/build.cjs.');
    process.exitCode = 1;
  } else console.log('Bundle atualizado e sintaxe válida.');
} else {
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  fs.writeFileSync(destination, output);
  console.log('Gerado V2/backend-unificado.gs');
}
