# Estabilização — 3.6.0-rc1

Estado: interface 3.6.0 validada para publicação; backend 3.6.0-rc1 confirmado no serviço.

## Diagnóstico confirmado

- Recuperação do utilizador ocorria depois de uma consulta remota.
- Dois formulários de reporte tinham comportamentos distintos na expiração da sessão.
- A submissão automática do PIN e o botão podiam iniciar dois logins.
- A configuração do endpoint guardava uma promessa rejeitada, impedindo nova tentativa nessa página.
- Pedidos públicos simultâneos iguais não eram agrupados.
- A leitura de registos fazia duas leituras de valores (cabeçalhos e dados).
- O serviço publicado não identificava a versão nem o tempo de execução.

Medições externas deste ambiente: status 21 762 ms; garage.publicConfig 21 764 ms.
Amostra limitada, pedidos em paralelo; inclui ligação, redirecionamentos e execução.
Não permite atribuir a demora ao Sheets ou garantir melhoria de latência.
Consultar backend/baseline-timings.json. Não contém códigos nem dados pessoais.

## Alterações preparadas

- Um único formulário de reporte, o mesmo do QR, com login apenas quando necessário.
- Regresso ao mapa quando o formulário foi aberto pelo mapa; QR direto regressa ao menu.
- Menu recuperado sem esperar pela configuração remota.
- Bloqueio da submissão repetida do login.
- Expiração no cadeado regressa ao login, conserva o motivo e volta à tarefa.
- PIN removido da memória da app após login e não reenviado para consultar o código.
- Login que termine depois de logout não repõe a sessão local.
- Pedidos públicos simultâneos partilham a chamada, sem cache persistente de dados.
- Uma leitura de valores para obter cabeçalhos e registos no adaptador Sheets.
- status com details=public devolve serviceVersion e serverDurationMs.
- window.dc4Diagnostics() permite consultar tempos sem corpos de pedidos, PINs ou tokens.

## Validação

55 testes Node passam, incluindo contratos legados, segurança, navegação e novos cenários.
O bundle Apps Script corresponde às fontes e tem sintaxe válida.
Não foram feitos reportes reais, logins reais nem envios de email nos testes.
Não foram executados testes em navegador: instalação de Chromium falhou por timeout.
Não considerar esta candidata validada no telemóvel.

## Passo que requer intervenção do proprietário

1. Guardar cópia do Apps Script atual.
2. Colar V2/backend-unificado.gs, preservando os IDs próprios de configuração.
3. Guardar e atualizar a implementação existente para uma nova versão, mantendo o URL.
4. Não executar setupApp: não há alterações no esquema do Sheets.
5. Consultar o endpoint existente com ?action=status&details=public e verificar
   serviceVersion e serverDurationMs. O diagnóstico é público; não enviar PINs.

O backend mantém o contrato da app 3.5.2. A interface candidata fica na branch de
estabilização até à validação. A branch não é um endereço de aplicação publicada.

## Verificação final ainda necessária

- Navegador e telemóvel: login, mapa, reporte, regresso ao mapa e menu, código.
- Sessão expirada durante texto/fotografia: reautenticar e continuar.
- Rede lenta/offline e atualização da app instalada.
- Comparar tempos totais com o tempo do serviço após a implementação.
- Validar em navegador antes de integrar a candidata em main.

A sessão continua por separador, em sessionStorage, com validade controlada pelo
servidor. Não foi criada sessão persistente entre fechos da aplicação.

## Validação após implementação

Serviço confirmado: 3.6.0-rc1. Consulta status com detalhes:
serverDurationMs=367; tempo total externo=24696 ms. A maior parte da espera
não pertence ao trecho medido; não atribuir ao telemóvel ou Sheets sem mais dados.

Chromium 153 executado com Playwright e viewport 390x844. Passaram os percursos:
menu com sessão, mapa com quatro extintores no piso -2, G4 para formulário,
descrição, regresso ao mapa e menu, QR sem sessão, autenticação no envio e rede
indisponível. API simulada e service worker desativado nestes testes; não houve
escritas reais. A instalação pelo download original falhou, mas uma distribuição
alternativa do Chromium permitiu executar os testes.

Teste em aparelho físico e atualização de uma PWA já instalada continuam pendentes.
O backend pode manter o identificador rc1: não é necessário voltar a implementá-lo
apenas para alterar esse identificador.
