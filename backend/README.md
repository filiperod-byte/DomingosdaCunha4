# Backend do condomínio — primeira etapa da reestruturação

O backend continua a executar no Apps Script e a guardar os dados no Sheets. A lógica passa a organizar-se por funções do condomínio, com os serviços Google concentrados num adaptador. A interface existente mantém as ações e respostas anteriores, incluindo os aliases `garage.*`.

A segunda etapa acrescenta `security.js` à entrada HTTP e `session-client.js` à interface. Ambos devem ser instalados em conjunto. A integração real com Google e os testes visuais no navegador continuam pendentes; a proposta permanece em rascunho até essa validação.

## Sessões e permissões

As ações administrativas exigem uma sessão de administrador; consultas do código exigem sessão de morador e os reportes recebem a identidade validada no servidor. O acesso antigo por `validatePin`, `setPin` e `resetPin` deixa de estar exposto. A entrada administrativa é `V2/admin.html`.

Tokens são guardados em `sessionStorage`, enviados no corpo POST, guardados como hash no servidor e expiram após 30 minutos (administração) ou 8 horas (morador). Existe uma sessão por conta: novo login termina a anterior. Bloqueio, desativação e alteração de PIN invalidam sessões nos pedidos seguintes. Logout revoga a sessão no servidor quando há ligação; sem rede permanece a expiração. Não se usa o valor de `dc4_admin_unlocked` como autorização no servidor.

Há limites globais por janela: 10 logins administrativos e 30 logins de moradores por 15 minutos, e 10 pedidos públicos de registo/recuperação por hora. Contam também pedidos bem-sucedidos. Isto limita tentativas e spam, mas um atacante pode esgotar a janela e impedir temporariamente acessos legítimos. O Apps Script não fornece aqui um mecanismo robusto de limitação por IP. A proteção contra abuso e o modelo de credenciais devem evoluir na migração.

O PIN administrativo tem 6 a 12 dígitos. Os valores de fábrica `123456` e `000000` são recusados. Alterar diretamente `ADMIN_PIN` na folha CONFIG antes de validar a instalação, usando um valor próprio. Não partilhar o PIN no GitHub nem nesta conversa.

Persistem limitações: PINs de moradores e o PIN administrativo no Sheets não estão cifrados; as fotografias conservam partilha por ligação; não foi realizada uma auditoria completa de XSS ou de ficheiros enviados. A sessão no navegador não protege contra XSS. Restringir editores da base de dados e planear autenticação dedicada na migração.

## Organização

| Ficheiro em `src/` | Responsabilidade |
| --- | --- |
| `config.js` | Configuração e esquema legado |
| `domain.js` | Normalização e validações sem serviços externos |
| `residents.js` | Registo, aprovação e ciclo de vida dos moradores |
| `occurrences.js` | Reporte, aprovação e fecho de ocorrências |
| `accesses.js` | Consulta do código do cadeado e histórico |
| `legacy-pin.js` | Compatibilidade com o PIN do backoffice antigo |
| `notifications.js` | Conteúdo e destinatários das notificações |
| `application.js` | Composição dos serviços e despacho das ações |
| `apps-script-ports.js` | Sheets, Drive, email, propriedades, relógio, identificadores e locks |
| `entrypoints.js` | Entrada HTTP do Apps Script e operações manuais |

Os serviços recebem um objeto `ports`. Os repositórios expõem operações como `residents.list/get/add/update`, `occurrences.list/update/upsert` e `consultations.list/add`. O adaptador traduz os nomes dos campos dos moradores e consultas para as colunas do Sheets. Os serviços não conhecem números de colunas nem chamam serviços Google diretamente.

As ocorrências ainda usam campos do esquema legado, como `OCCURRENCE_ID`, e a API dos moradores ainda identifica operações por `row`. Estas dependências estão explícitas: uma migração deve substituir gradualmente o número da linha por um identificador estável e mapear o esquema antigo.

## Construção e verificação

Requer Node.js 20 ou posterior, sem dependências npm:

```sh
node backend/build.cjs
node --test backend/tests/*.test.cjs
node backend/build.cjs --check
```

Editar `backend/src/`, gerar e incluir também `V2/backend-unificado.gs` no commit. O ficheiro `.gs` é o artefacto de instalação; não editar diretamente. O gerador concatena os módulos numa ordem explícita e verifica a sintaxe. Não utiliza imports Node nem módulos ES no código executado pelo Apps Script.

Os testes executam apenas em memória, sem rede, emails reais ou acesso a folhas. `tests/compatibility.json` contém resultados obtidos do backend original no commit `c89622f3f70645290e02e0865c44111f4a07e3c3`, usando dados, identificadores, relógio e destinatários fictícios. Oito cenários comparam respostas, persistência, notificações e fotografias. Outros testes verificam consultas sem escritas, inicialização, propriedades, linhas inválidas, colunas reordenadas e execução sem serviços Google.

**Compatibilidade não prova segurança:** os testes de referência executam explicitamente apenas o núcleo legado, sem a política HTTP. Os testes de `security.test.cjs` e `client.test.cjs` verificam separadamente a fronteira autenticada e o cliente. Os resultados de referência incluem comportamentos vulneráveis que já não são expostos pela entrada HTTP. A integração real com Google, permissões Drive, quotas, entrega de email e formatação no fuso horário não foram validadas por estes testes.

## Alterações de comportamento intencionais

- A criação e preparação de folhas acontece apenas em `setupBackend_()` ou no alias `setupApp()`. Pedidos normais já não inicializam o esquema. Uma instalação nova precisa dessa operação manual antes de receber pedidos.
- Cada pedido reutiliza a spreadsheet aberta e a configuração lida. Um pedido seguinte recebe um novo adaptador, evitando configuração desatualizada entre pedidos.
- Definir o PIN antigo preserva outras propriedades do script; a versão anterior apagava propriedades alheias ao PIN.
- Operações de administração sobre linhas inexistentes ou sobre o cabeçalho são rejeitadas.

## Etapas seguintes

1. **Validar integração e reforçar credenciais.** Testar Apps Script, Drive e interface em conjunto; substituir gradualmente os PINs partilhados e o armazenamento de credenciais no Sheets por autenticação dedicada.
2. **Modelo do condomínio.** Unificar identidade e permissões, introduzir identificadores estáveis e tratar extintores como um tipo de equipamento/ocorrência. O PIN legado deve ser retirado de forma coordenada. Rever também fotografias públicas por ligação, autoria dos reportes e destinatários fornecidos pelo cliente.
3. **Migração de infraestrutura.** Implementar adaptadores para a plataforma escolhida e uma nova entrada HTTP, migrar e verificar os dados, trocar o endpoint da interface e retirar o anterior após validação. As portas são atualmente síncronas; fornecedores com SDK assíncrono exigirão adaptar chamadas para `async/await`. Não é uma migração sem alterações ao código.

O adaptador Sheets ainda lê tabelas inteiras em várias operações e conserva as limitações de concorrência e de consistência do backend original. Um futuro adaptador deve permitir consultas filtradas e transações adequadas. A separação atual reduz o trabalho de migração, mas não resolve por si só escala, segurança ou modelação de dados.

## Publicação

Ver [instruções de instalação](../V2/INSTALAR_BACKEND.md). Alterar o GitHub não atualiza o Apps Script publicado. Esta branch não altera o endpoint, as folhas ou o Drive em produção.


## Consulta pública do motivo (interface v3.4.0)

`status` com `details=public` acrescenta `reason` e `reportedAt` aos pontos com ocorrência aberta e `publicDetailsVersion: 1` à resposta. O pedido sem esse parâmetro conserva o contrato anterior. Só são publicadas as categorias previstas no formulário; texto livre é apresentado como «Outra anomalia». Não são incluídos nomes, observações, fotografias ou reportes pendentes. Ao alterar as categorias do formulário, rever também esta lista no serviço de ocorrências.

A interface continua a funcionar com o backend anterior, mas só consegue mostrar o motivo após atualizar o ficheiro Apps Script e a versão da implantação. Esta alteração não exige executar `setupApp()` nem mudar os dados da spreadsheet. Preservar a configuração dos recursos Google ao substituir o código.
