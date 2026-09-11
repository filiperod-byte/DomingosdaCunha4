// Conteúdo das notificações separado do transporte de email.
function createNotificationService_(ports, CONFIG, domain) {
  const { isRealEmail_ } = domain;
  const getConfig = key => ports.settings.get(key);
  const isoNow_ = () => ports.clock.now().toISOString();
  const safeSendEmail_ = (to, subject, body, options) => ports.mail.send(to, subject, body, options);
function sendReportEmail_(d) { if (!d.adminEmail) return; const status = d.pendingValidation ? 'PENDENTE DE VALIDAÇÃO' : 'ABERTA'; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reporte extintor - ' + status, ['Reporte de extintor recebido.', '', 'Estado: ' + status, 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Localização: ' + (d.location || '—'), 'Reportado por: ' + d.reportedBy, 'Motivo: ' + d.reason, 'Observação: ' + (d.notes || '—'), 'Foto: ' + (d.photoUrl || '—'), '', 'Data/Hora: ' + isoNow_()].join('\n')); }

function sendCloseEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Ocorrência fechada', ['Ocorrência fechada.', '', 'Ocorrência: ' + d.occurrenceId, 'Piso: ' + d.floorLabel, 'Ponto: ' + d.point, 'Foto fecho: ' + (d.closePhotoUrl || '—'), 'Observação: ' + (d.closeNotes || '—')].join('\n')); }

function sendPinResetEmail_(d) { if (!d.adminEmail) return; safeSendEmail_(d.adminEmail, '[' + CONFIG.APP_NAME + '] Reset do PIN Backoffice', 'O PIN do backoffice foi resetado.'); }

function emailRegisto(email,nome) { safeSendEmail_(email, 'Pedido de acesso recebido — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nRecebemos o seu pedido de acesso. Após aprovação receberá o PIN por email.', { replyTo:getConfiguredAdminEmail_() }); }

function emailAdminNovoPedido(d) { safeSendEmail_(getConfig('ADMIN_EMAIL'), 'Novo pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Novo pedido:\n\nNome: ' + d.nome + '\nPiso: ' + d.piso + '\nFração/Garagem: ' + d.fracao + '\nEmail: ' + d.email + '\nTelemóvel: ' + (d.telemovel || '—'), { replyTo:getConfiguredAdminEmail_() }); }

function emailAprovacao(email,nome,pin) { safeSendEmail_(email, 'Acesso aprovado — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin + '\n\nGuarde-o num local seguro.', { replyTo:getConfiguredAdminEmail_() }); }

function emailRejeicao(email,nome) { safeSendEmail_(email, 'Pedido de acesso — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu pedido não foi aprovado. Contacte a administração.', { replyTo:getConfiguredAdminEmail_() }); }

function emailPINRecuperacao(email,nome,pin) { safeSendEmail_(email, 'Recuperação de PIN — ' + getConfig('NOME_CONDOMINIO'), 'Olá ' + nome + ',\n\nO seu PIN pessoal é: ' + pin, { replyTo:getConfiguredAdminEmail_() }); }

function getConfiguredAdminEmail_() { const e = getConfig('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL || ''; return isRealEmail_(e) ? e : ''; }
return { sendReportEmail_, sendCloseEmail_, sendPinResetEmail_, emailRegisto, emailAdminNovoPedido, emailAprovacao, emailRejeicao, emailPINRecuperacao };
}
