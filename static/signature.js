(() => {
  function initSignaturePad() {
    const canvas = document.getElementById('signature-pad');
    const hiddenInput = document.getElementById('firma');
    const clearBtn = document.getElementById('clear-signature');
    const errorBox = document.getElementById('signature-error');
    const form = document.getElementById('signature-form');
    const submitBtn = document.getElementById('submit-btn');

    if (!canvas || !hiddenInput || !form) return;

    const ctx = canvas.getContext('2d');
    let drawing = false;
    let hasDrawn = false;

    function showError(message) {
      if (!errorBox) return;
      errorBox.textContent = message;
      errorBox.classList.remove('d-none');
    }

    function hideError() {
      if (!errorBox) return;
      errorBox.textContent = '';
      errorBox.classList.add('d-none');
    }

    function resizeCanvas() {
      const ratio = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();

      // Reset transform so scaling doesn't compound
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      canvas.width = Math.max(1, Math.round(rect.width * ratio));
      canvas.height = Math.max(1, Math.round(rect.height * ratio));
      ctx.setTransform(ratio, 0, 0, ratio, 0, 0);

      // Styling
      ctx.lineWidth = 2;
      ctx.lineCap = 'round';
      ctx.strokeStyle = '#000000';

      // Wipe on resize (simple + predictable)
      clear();
    }

    function getPointFromEvent(evt) {
      const rect = canvas.getBoundingClientRect();
      return {
        x: evt.clientX - rect.left,
        y: evt.clientY - rect.top,
      };
    }

    function start(evt) {
      drawing = true;
      hideError();
      const p = getPointFromEvent(evt);
      ctx.beginPath();
      ctx.moveTo(p.x, p.y);
    }

    function move(evt) {
      if (!drawing) return;
      const p = getPointFromEvent(evt);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      hasDrawn = true;
    }

    function end() {
      drawing = false;
    }

    function clear() {
      drawing = false;
      hasDrawn = false;
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      const ratio = window.devicePixelRatio || 1;
      ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
      ctx.beginPath();
      hiddenInput.value = '';
      hideError();
    }

    // Pointer events (mouse + touch + pen)
    canvas.addEventListener('pointerdown', (evt) => {
      evt.preventDefault();
      canvas.setPointerCapture(evt.pointerId);
      start(evt);
    });
    canvas.addEventListener('pointermove', (evt) => {
      evt.preventDefault();
      move(evt);
    });
    canvas.addEventListener('pointerup', (evt) => {
      evt.preventDefault();
      end();
    });
    canvas.addEventListener('pointercancel', (evt) => {
      evt.preventDefault();
      end();
    });
    canvas.addEventListener('pointerleave', () => {
      end();
    });

    if (clearBtn) {
      clearBtn.addEventListener('click', (evt) => {
        evt.preventDefault();
        clear();
      });
    }

    window.addEventListener('resize', () => resizeCanvas());

    form.addEventListener('submit', (evt) => {
      hideError();

      if (!hasDrawn) {
        evt.preventDefault();
        showError('Por favor, firma antes de enviar.');
        return;
      }

      hiddenInput.value = canvas.toDataURL('image/png');

      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.dataset.originalText = submitBtn.textContent || 'Enviar';
        submitBtn.textContent = 'Enviando…';
      }
    });

    resizeCanvas();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSignaturePad);
  } else {
    initSignaturePad();
  }
})();
