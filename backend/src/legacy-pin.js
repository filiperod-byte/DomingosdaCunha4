// Compatibilidade com o PIN do backoffice antigo. Não constitui uma sessão autenticada.
function createLegacyPinService_(ports, CONFIG, domain, notifications, getAdminEmail_) {
  const { normalizePin_ } = domain;
  const withScriptLock_ = cb => ports.lock.run(cb);
  const isoNow_ = () => ports.clock.now().toISOString();
  const hashPin_ = (pin, salt) => ports.crypto.hashPin(pin, salt);
  const { sendPinResetEmail_ } = notifications;
function handleGetPinStatus_() {
  return { success:true, pinConfigured:isPinConfigured_(), updatedAt:ports.properties.getProperty(CONFIG.PROPS.PIN_UPDATED_AT) || '' };
}

function handlePostSetPin_(payload) {
  return withScriptLock_(function () {
    const pin = normalizePin_(payload.pin);
    if (!pin) throw new Error('PIN obrigatório.');
    if (pin.length < CONFIG.PIN_MIN_LENGTH) throw new Error('O PIN deve ter pelo menos ' + CONFIG.PIN_MIN_LENGTH + ' dígitos.');
    if (!/^\d+$/.test(pin)) throw new Error('O PIN deve conter apenas dígitos.');
    if (isPinConfigured_()) throw new Error('O PIN já está definido.');
    const salt = ports.crypto.uuid();
    ports.properties.setProperties({ [CONFIG.PROPS.PIN_HASH]: hashPin_(pin, salt), [CONFIG.PROPS.PIN_SALT]: salt, [CONFIG.PROPS.PIN_UPDATED_AT]: isoNow_() }, false);
    return { success:true, message:'PIN definido com sucesso.' };
  });
}

function handlePostValidatePin_(payload) {
  const pin = normalizePin_(payload.pin);
  if (!isPinConfigured_()) return { success:true, valid:false, message:'Ainda não existe PIN definido.' };
  if (!pin) return { success:true, valid:false, message:'PIN vazio.' };
  const props = ports.properties;
  return { success:true, valid: hashPin_(pin, props.getProperty(CONFIG.PROPS.PIN_SALT) || '') === (props.getProperty(CONFIG.PROPS.PIN_HASH) || '') };
}

function handlePostResetPin_(payload) {
  return withScriptLock_(function () {
    clearPin_();
    sendPinResetEmail_({ adminEmail:getAdminEmail_(payload) });
    return { success:true, message:'PIN resetado.' };
  });
}

function isPinConfigured_() { const p = ports.properties; return !!(p.getProperty(CONFIG.PROPS.PIN_HASH) && p.getProperty(CONFIG.PROPS.PIN_SALT)); }

function clearPin_() { const p = ports.properties; p.deleteProperty(CONFIG.PROPS.PIN_HASH); p.deleteProperty(CONFIG.PROPS.PIN_SALT); p.deleteProperty(CONFIG.PROPS.PIN_UPDATED_AT); }
  return { status: handleGetPinStatus_, set: handlePostSetPin_, validate: handlePostValidatePin_,
    reset: handlePostResetPin_, configured: isPinConfigured_, clear: clearPin_ };
}
