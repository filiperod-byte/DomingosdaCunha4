# Domingos da Cunha 4 — v2 unificado

Esta pasta junta os dois módulos do condomínio:

- **Extintores**: mantém a app atual em `../index.html` e `../backoffice.html`.
- **Garagem / Cadeado**: nova versão estática em GitHub Pages, usando `fetch` para comunicar com o backend Apps Script.

## URLs de teste

- Menu v2: `https://filiperod-byte.github.io/DomingosdaCunha4/v2/`
- Garagem: `https://filiperod-byte.github.io/DomingosdaCunha4/v2/garagem.html`
- Admin Garagem: `https://filiperod-byte.github.io/DomingosdaCunha4/v2/garagem-admin.html`

## Backend

Copiar o ficheiro `backend-unificado.gs` para o projeto Apps Script e publicar uma nova versão do Web App.

Definições recomendadas do deployment:

- Executar como: **Eu**
- Quem tem acesso: **Qualquer pessoa**



## Backend modular (proposta de reestruturação)

O código fonte do backend passa a estar em [`backend/src/`](../backend/src/). O ficheiro `backend-unificado.gs` é gerado por `node backend/build.cjs`. Ver [arquitetura e testes](../backend/README.md) e [instalação](INSTALAR_BACKEND.md). A interface e o backend incluem sessões validadas no servidor e devem ser instalados em conjunto. Falta validar a integração real com Google antes de publicar.
