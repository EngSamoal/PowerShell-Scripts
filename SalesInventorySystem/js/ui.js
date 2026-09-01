/**
 * ui.js — small reusable Arabic/RTL UI helpers: toasts, confirm modal,
 * and screen switching. Keeping these in one place means every screen
 * gets the same look for errors/confirmations instead of the browser's
 * default alert()/confirm() (which don't respect RTL well).
 */

const UI = {
  /** Shows a brief toast message. type: 'success' | 'error' | 'info'. */
  toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    requestAnimationFrame(() => toast.classList.add('toast-visible'));

    setTimeout(() => {
      toast.classList.remove('toast-visible');
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  },

  /** Shows a friendly error toast. Use for any caught technical exception. */
  error(message) {
    this.toast(message || 'تعذر إتمام العملية. يرجى المحاولة مرة أخرى.', 'error');
  },

  success(message) {
    this.toast(message, 'success');
  },

  /**
   * Shows an Arabic confirmation modal and resolves true/false with the
   * user's choice. Use before any risky action (delete, void, restore...).
   */
  confirm({ title = 'تأكيد', message, confirmLabel = 'تأكيد', cancelLabel = 'إلغاء', danger = false }) {
    return new Promise((resolve) => {
      this.showModal(`
        <div class="modal">
          <h3 class="modal-title">${title}</h3>
          <p class="modal-message">${message}</p>
          <div class="modal-actions">
            <button type="button" class="btn btn-secondary" data-action="cancel">${cancelLabel}</button>
            <button type="button" class="btn ${danger ? 'btn-danger' : 'btn-primary'}" data-action="confirm">${confirmLabel}</button>
          </div>
        </div>
      `);

      const overlay = document.getElementById('modal-overlay');
      const close = (result) => {
        this.closeModal();
        resolve(result);
      };

      overlay.querySelector('[data-action="cancel"]').addEventListener('click', () => close(false));
      overlay.querySelector('[data-action="confirm"]').addEventListener('click', () => close(true));
    });
  },

  /** Shows arbitrary modal HTML (a <div class="modal">...</div>, optionally .modal-wide). Caller wires up its own events. */
  showModal(html) {
    const overlay = document.getElementById('modal-overlay');
    overlay.innerHTML = html;
    overlay.classList.add('modal-overlay-visible');
  },

  /** Closes and clears whatever modal is currently shown. */
  closeModal() {
    const overlay = document.getElementById('modal-overlay');
    overlay.classList.remove('modal-overlay-visible');
    overlay.innerHTML = '';
  },

  /** Shows one screen (by id) and hides the others; highlights the matching nav item. */
  showScreen(screenId) {
    document.querySelectorAll('.screen').forEach((el) => el.classList.remove('screen-active'));
    const target = document.getElementById(screenId);
    if (target) target.classList.add('screen-active');

    document.querySelectorAll('.nav-item').forEach((el) => {
      el.classList.toggle('nav-item-active', el.dataset.screen === screenId);
    });
  },
};
