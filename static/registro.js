(() => {
  function initRegistroValidation() {
    const cursoEl = document.getElementById('curso');
    const codigoEl = document.getElementById('codigo');
    const submitBtn = document.getElementById('submit-btn');

    if (!cursoEl || !codigoEl || !submitBtn) return;

    // Caja de alerta (debajo del código)
    let codeAlert = document.getElementById('code-exists-alert');
    if (!codeAlert) {
      codeAlert = document.createElement('div');
      codeAlert.id = 'code-exists-alert';
      codeAlert.className = 'alert alert-warning d-none mt-2';

      // Insertar después del input de código
      const wrapper = codigoEl.parentElement;
      if (wrapper) wrapper.appendChild(codeAlert);
    }

    let lastQueryKey = '';
    let debounceTimer = null;

    function setBlocked(blocked, message) {
      if (blocked) {
        codeAlert.textContent = message || 'Este código ya firmó.';
        codeAlert.classList.remove('d-none');
        submitBtn.disabled = true;
      } else {
        codeAlert.textContent = '';
        codeAlert.classList.add('d-none');
        submitBtn.disabled = false;
      }
    }

    async function check() {
      const curso = (cursoEl.value || '').trim();
      const codigo = (codigoEl.value || '').trim();

      if (!curso || !codigo || codigo.length < 3) {
        setBlocked(false);
        return;
      }

      const key = `${curso}::${codigo}`;
      if (key === lastQueryKey) return;
      lastQueryKey = key;

      try {
        const url = `/api/firmas/existe?curso=${encodeURIComponent(curso)}&codigo=${encodeURIComponent(codigo)}`;
        const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
        const data = await res.json();

        if (data && data.ok && data.exists === true) {
          setBlocked(true, 'Este código ya firmó en este curso.');
        } else {
          setBlocked(false);
        }
      } catch {
        // Si falla la validación, no bloquear el registro
        setBlocked(false);
      }
    }

    function scheduleCheck() {
      if (debounceTimer) window.clearTimeout(debounceTimer);
      debounceTimer = window.setTimeout(check, 350);
    }

    codigoEl.addEventListener('input', scheduleCheck);
    codigoEl.addEventListener('blur', check);
    cursoEl.addEventListener('change', () => {
      lastQueryKey = '';
      scheduleCheck();
    });

    // Primer check si hay valores pre-cargados
    scheduleCheck();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initRegistroValidation);
  } else {
    initRegistroValidation();
  }
})();
