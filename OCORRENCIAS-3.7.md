# Ocorrências — candidata 3.7.0

## Funcionalidades
- Menu Ocorrências com seleção visual do edifício (pisos existentes, incluindo 0 e garagens) e Prédio / Zonas gerais.
- Por piso: extintores identificados, iluminação, portas, limpeza, conservação e outro. Geral: elevadores, portão, entrada/intercomunicador, água/bombas, exterior, limpeza e outro.
- QR dos extintores e histórico existentes preservados. Novo QR por piso abre general-report.html?floor=N; seleção de tipo e motivo no formulário.
- Administração gera a seleção imprimível das 13 etiquetas de piso; SVG locais, sem chamadas a geradores QR externos. Logótipo colocado fora do QR para preservar a leitura.
- Consulta pública só de ocorrências gerais validadas, com descrição revista, data, estado, intervenção e resolução. Recebido e Não validado ficam reservados à administração.
- Identidade verificada ao enviar; reutiliza sessão válida, registo sujeito a aprovação e autenticação/fotografia do formulário dos extintores.
- Também verifiquei contabiliza cada conta uma vez; não cria outra ocorrência.
- Administração: recebido, em análise, intervenção agendada, resolvido e não validado; texto público separado das notas internas, original e histórico consultáveis.
- Fotografias novas não ativam partilha pública; continuam sujeitas às permissões da pasta Drive. O endpoint público não devolve links nem autores.

## Implementação pelo proprietário
1. Guardar cópia do Apps Script atual.
2. Substituir pelo conteúdo completo de V2/backend-unificado.gs, preservando os IDs de configuração.
3. Guardar e executar **setupGeneralOccurrences** uma vez no editor. Cria apenas OCORRENCIAS_GERAIS e respetivos cabeçalhos; repetir não apaga dados. **Não executar setupApp.**
4. Atualizar a implementação existente: Gerir implementações → lápis → Nova versão → Implementar. Manter o URL atual.
5. Confirmar a implementação para permitir verificar general.status e publicar a interface 3.7.0.

O backend identifica-se como 3.7.0-rc1. É compatível com a interface atual 3.6.0. A interface candidata só deve ser integrada após o backend estar preparado. O estado público general.status deverá devolver success:true e occurrences:[], se ainda não houver ocorrências validadas.

## Persistência e repetição de envios
A folha nova é um registo de eventos. Cada alteração acrescenta uma linha com o estado completo e o recibo (conta, ação, identificador e hash do conteúdo), sob ScriptLock. Uma repetição devolve o recibo existente. A interface mantém o mesmo pedido e congela a edição enquanto o resultado é incerto. Uma rejeição de validação permite corrigir o formulário. A fotografia pode ficar órfã se houver falha antes de gravar a linha; não é criada uma segunda ocorrência por repetição após gravação.

A proteção é específica das operações gerais novas. Os reportes dedicados aos extintores mantêm o comportamento anterior. O formulário e identificador em curso mantêm-se enquanto a página estiver aberta; fechar/recarregar pode perder esse rascunho. Não há envio automático ao recuperar rede.

As edições administrativas exigem a revisão atual para não sobrescrever alterações de outro administrador. Notas e fotos não saem na consulta pública. O catálogo partilhado define localizações/categorias tanto na interface como no serviço.

## Verificação
- 62 testes Node: contratos antigos, permissões, bloqueio/sessões, validação, publicação, privacidade, confirmação única, conflito de revisão e repetição após gravação com resposta perdida.
- Chromium em ecrã móvel com backend executado num emulador dos serviços Google: piso -2, G4 dedicado, piso 0, prédio geral, login ao enviar, fotografia, expiração/reautenticação, publicação administrativa, confirmação, indisponibilidade e repetição sem duplicados.
- Percursos anteriores dos extintores também passaram no Chromium.
- 13 SVG descodificados com verificação exata dos URLs. Impressão A4 inspecionada e QR impresso do piso -2 descodificado.
- Não houve logins, reportes, emails ou alterações de dados reais nestes testes.

Limites: API emulada e service worker desativado nos testes em navegador. Falta validar implementação real e atualização da app instalada. A primeira versão das ocorrências gerais não envia emails automáticos: a administração consulta a lista na app. Permissões das fotografias dependem também da pasta Drive. O custo da leitura do histórico cresce com o número de eventos; medir antes de otimizar.
