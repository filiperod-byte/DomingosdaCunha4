# Instalar o backend modular com sessões

As alterações incluem a interface GitHub e o Apps Script. Guardar o código no editor Google não atualiza uma implantação versionada `/exec`. Não substituir apenas uma das partes em produção: os clientes antigos não enviam a sessão agora obrigatória.

## Preparar o código

O ficheiro completo para copiar é [`backend-unificado.gs`](backend-unificado.gs). Não copiar os testes nem os ficheiros JavaScript da interface para o Apps Script. O fonte está em `backend/src/` e pode ser gerado com:

```sh
node backend/build.cjs
node --test backend/tests/*.test.cjs
node backend/build.cjs --check
```

## Primeiro validar numa cópia

1. Guardar uma cópia do código Google atual e registar a versão da implantação atual.
2. Criar cópias de teste da spreadsheet e da pasta Drive; no início do `.gs`, adaptar `SPREADSHEET_ID`, `ROOT_FOLDER_ID` e `ADMIN_EMAIL` aos recursos de teste. Não usar dados pessoais reais nos testes.
3. Copiar o ficheiro `.gs` completo para um projeto Apps Script V8. Não manter outro ficheiro com `doGet`, `doPost` ou as mesmas constantes globais. Preservar separadamente quaisquer funções próprias que não pertençam a este backend.
4. Executar `setupApp()` uma vez. Autorizar os serviços Google para os recursos de teste. A inicialização preserva os dados existentes e deixa de ocorrer em cada pedido.
5. Na folha `CONFIG`, confirmar `ADMIN_EMAIL` e definir `ADMIN_PIN` com 6 a 12 dígitos próprios. `123456` e `000000` são recusados. Os PINs dos moradores continuam a ter 6 dígitos; confirmar `PINAtivo=TRUE` nos moradores aprovados que devem entrar.
6. Criar uma implantação de teste e configurar uma cópia da interface desta branch para esse URL (ambos `config.json` e `V2/config.json`). Validar administrador, morador, QR, reportes/fotografias, bloqueio, logout e emails. Depois de alterar o PIN administrativo, voltar a entrar.

## Publicação coordenada

Preparar o novo código no editor do projeto Google de produção, com os identificadores corretos, sem publicar ainda uma nova versão. Confirmar o PIN administrativo próprio na base de dados. Manter uma cópia anterior para recuperação.

Depois de validar a cópia, publicar os ficheiros da interface desta branch e atualizar a implantação Google existente para uma nova versão, na mesma janela de manutenção. Manter a implantação preserva o URL `/exec`; se criar outra, atualizar ambos os `config.json`. O service worker tem nova versão para renovar os ficheiros da app; fechar/reabrir a app e confirmar a atualização nos dispositivos.

É necessário voltar a entrar na app. A administração de extintores passa a usar o login da administração em `V2/admin.html`. O reset anónimo do PIN antigo foi desativado; a recuperação administrativa faz-se pelo editor autorizado da folha CONFIG.

Os testes locais não validam permissões Google, quotas, entrega real de emails nem funcionamento visual no telemóvel. A branch continua em rascunho até à validação dessa integração. Nenhum endpoint Google foi alterado automaticamente.
