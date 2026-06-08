// Ajusta o texto do modal de reporte antes da submissão.
// Enquanto ainda está a ser preenchido, deve dizer "A registar por" e não "Reportado por".
(function () {
  function fixReporterText() {
    const subtitle = document.getElementById('modalSubtitle');
    if (!subtitle) return;
    subtitle.textContent = subtitle.textContent
      .replace('Reportado por:', 'A registar por:')
      .replace('Registado por:', 'A registar por:');
  }

  function patchOpenModal() {
    if (typeof window.openModal !== 'function' || window.openModal.__reportTextFixed) return;
    const original = window.openModal;
    window.openModal = function patchedOpenModal() {
      const result = original.apply(this, arguments);
      window.setTimeout(fixReporterText, 0);
      return result;
    };
    window.openModal.__reportTextFixed = true;
  }

  document.addEventListener('DOMContentLoaded', function () {
    patchOpenModal();
    window.setTimeout(patchOpenModal, 250);
  });
})();
