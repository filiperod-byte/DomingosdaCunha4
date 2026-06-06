# Instalar backend unificado no Google Apps Script

Este ficheiro explica como aplicar o backend unificado da V2.

## Objetivo

A V2 usa GitHub Pages como frontend e Google Apps Script como backend.

O backend deve ter apenas:

- um `doGet(e)`
- um `doPost(e)`
- um router único para os módulos:
  - Extintores
  - Garagem / Cadeado

## Ficheiro a usar

Copiar o conteúdo completo de:

```text
V2/backend-unificado.gs
```

para o Google Apps Script.

## Atenção importante

No Apps Script, todos os ficheiros `.gs` são carregados ao mesmo tempo.

Por isso, não basta renomear ficheiros antigos.

Se existirem vários ficheiros `.gs` com:

```js
const CONFIG = ...
function doGet(e) { ... }
function doPost(e) { ... }
```

vai dar conflito.

## Instalação recomendada

1. Abrir o projeto no Google Apps Script.
2. Criar uma cópia de segurança do código antigo fora do Apps Script, por exemplo no GitHub ou num ficheiro `.txt`.
3. No ficheiro principal `.gs`, colar todo o conteúdo de `V2/backend-unificado.gs`.
4. Nos outros ficheiros `.gs` antigos, apagar o conteúdo ou comentar tudo com:

```js
/*
  código antigo aqui dentro
*/
```

5. Os ficheiros `.html` antigos (`Index.html`, `Admin.html`, `Style.html`) podem ficar. Já não são usados pela V2, mas não criam conflito se não forem chamados.
6. Guardar o projeto.
7. Executar manualmente:

```js
setupBackend_()
```

8. Executar manualmente:

```js
setupApp()
```

9. Autorizar permissões quando o Google pedir.
10. Fazer nova implementação do Web App:

```text
Deploy / Implementar → Manage deployments / Gerir implementações → Editar → Nova versão
```

11. Confirmar configuração:

```text
Execute as: Me / Eu
Who has access: Anyone / Qualquer pessoa
```

12. Publicar a nova versão.

## Testes

Abrir no browser:

```text
https://script.google.com/macros/s/AKfycbxgAehbnaj2kGv-3K6PDRzPXHbiFF4nCimJ6Adje4-d917-MFhcHE9xjLOb-8cwqKdzQw/exec?action=health
```

Deve devolver JSON com:

```json
{
  "success": true
}
```

Depois testar:

```text
https://filiperod-byte.github.io/DomingosdaCunha4/V2/garagem.html
```

E:

```text
https://filiperod-byte.github.io/DomingosdaCunha4/V2/garagem-admin.html
```

## Nota

A app antiga da garagem não precisa de ser apagada enquanto ideia/projeto.

O que não pode continuar é ter código antigo ativo com `doGet`, `doPost` ou `CONFIG` duplicados dentro do mesmo projeto Apps Script.
