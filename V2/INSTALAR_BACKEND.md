# Preparar o backend modular no Apps Script

Esta refatorização mantém o contrato da interface atual, mas conserva falhas de autenticação e autorização do backend anterior. **Usar primeiro um ambiente de teste; não publicar esta etapa isoladamente no endpoint público.** Consultar [arquitetura, limitações e próximos passos](../backend/README.md).

## Gerar e testar

Na raiz do repositório, com Node.js 20 ou posterior:

```sh
node backend/build.cjs
node --test backend/tests/backend.test.cjs
node backend/build.cjs --check
```

O código fonte está em `backend/src/`. O ficheiro gerado para o editor Google é `V2/backend-unificado.gs`.

## Validar numa cópia

1. Criar uma spreadsheet e pasta Drive de teste, com dados fictícios. Configurar os identificadores e o email de teste em `backend/src/config.js` e gerar novamente o ficheiro `.gs`. A configuração original aponta para recursos existentes: substituí-la antes de executar a cópia.
2. Criar um projeto Apps Script de teste com runtime V8. Copiar o conteúdo gerado para um ficheiro do projeto, sem manter outro backend com as mesmas funções globais.
3. Executar `setupApp()` no editor. É o alias público da rotina `setupBackend_()` e prepara as folhas necessárias. Autorizar apenas os recursos de teste pretendidos. A inicialização é manual: as consultas HTTP já não criam folhas.
4. Validar os fluxos numa implantação de teste com acesso restrito. Se o acesso restrito impedir os pedidos da interface estática, validar as funções no editor e concluir a autenticação antes de expor o endpoint.
5. Verificar registo, aprovação, login, consulta do código, reporte, aprovação e fecho com fotografia, bem como emails, fuso horário e permissões dos ficheiros. Os testes locais não substituem esta integração.

## Quando a versão estiver pronta para produção

Após corrigir autenticação e permissões e validar a integração, gerar com a configuração correta, atualizar o código no projeto Apps Script de produção e publicar uma nova versão da implantação existente. Se for mantida a implantação, o URL `/exec` mantém-se; uma implantação nova exige atualizar o endpoint usado pela interface.

A atualização dos ficheiros no GitHub e a atualização da implantação Google são operações distintas. Nenhuma implantação Google foi alterada por esta proposta. Guardar a versão anterior para permitir voltar atrás.
