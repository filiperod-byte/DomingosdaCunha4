# Login — backend 3.6.1-rc1

Candidata compatível com a interface 3.6.0. Não requer alterações ao Sheets.

## Diagnóstico e alteração
O login seguro lia CONDOMINOS para validar a conta e o serviço de moradores voltava a ler a mesma lista. A lista é agora reutilizada apenas durante o pedido; uma escrita invalida-a e cada pedido cria um adaptador novo. Não há cache de autorizações entre pedidos. Bloqueios, PIN ativo, expiração, limites de tentativas e revogação mantêm-se.

O teste conta três leituras de valores no login, em vez de quatro: lista, cabeçalhos e linha para atualizar o último acesso. Não foi alterada a escrita do último acesso nem a emissão de sessão.

Logins bem-sucedidos devolvem serviceVersion, serverDurationMs e loginTimingsMs com rateLimitMs (inclui espera pelo lock), validationMs, accessUpdateMs e sessionMs. São durações, sem credenciais ou dados pessoais. serverDurationMs mede apenas a execução dentro de routeRequest_, não o arranque Google nem a rede. O cliente atual já regista o total e serverMs em dc4Diagnostics(); as etapas estão na resposta JSON do login.

## Validação
56 testes Node passaram; bundle gerado e verificado. Chromium com viewport móvel e API simulada: menu, mapa, G4, regressos, login no envio, expiração do servidor com texto/fotografia preservados, reautenticação e envio, falha de rede com formulário preservado, repetição após recuperação e logout local. Testes de backend validam também revogação e bloqueio entre pedidos. Não foram feitos logins reais, reportes reais ou envios de email.

Limitações: API simulada e service worker desativado no Chromium. Ganho real de latência só pode ser medido após implementação e login pelo proprietário. A falha simulada de rede ocorre antes de escrever; não prova proteção contra duplicação se o servidor gravar e a resposta se perder.

## Implementação
Guardar cópia atual; substituir pelo V2/backend-unificado.gs preservando os IDs de configuração; guardar e atualizar a implementação existente para Nova versão. Não executar setupApp. A interface permanece 3.6.0 e o serviço identifica-se como 3.6.1-rc1.

Depois, efetuar um login normal e comparar a duração com os anteriores 6–9 segundos. Não partilhar PINs ou tokens. A publicação em produção do backend depende desta implementação manual.
