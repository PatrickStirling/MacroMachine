export function createTextPrompt() {
  let activePrompt = null;
  let activeCleanup = null;

  const destroyActive = (value, resolve) => {
    if (activeCleanup) {
      activeCleanup();
      activeCleanup = null;
    }
    if (activePrompt) {
      try { activePrompt.remove(); } catch (_) {}
      activePrompt = null;
    }
    resolve(value);
  };

  const open = ({
    title,
    label,
    initialValue,
    confirmText,
    cancelText,
    placeholder,
    inputType,
    multiline,
    rows,
    selectOnOpen,
  } = {}) => new Promise((resolve) => {
    if (activeCleanup) {
      activeCleanup();
      activeCleanup = null;
    }
    if (activePrompt) {
      try { activePrompt.remove(); } catch (_) {}
      activePrompt = null;
    }

    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const form = document.createElement('form');
    form.className = 'add-control-form';

    const header = document.createElement('header');
    const headerWrap = document.createElement('div');
    const titleEl = document.createElement('h3');
    titleEl.textContent = title || 'Enter value';
    headerWrap.appendChild(titleEl);

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = 'x';
    closeBtn.setAttribute('aria-label', 'Close');
    header.appendChild(headerWrap);
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'form-body';
    const field = document.createElement('label');
    field.textContent = label || 'Value';
    const control = multiline ? document.createElement('textarea') : document.createElement('input');
    if (!multiline) {
      control.type = inputType || 'text';
    } else {
      const nextRows = Number.isFinite(rows) ? rows : 8;
      control.rows = nextRows > 0 ? nextRows : 8;
    }
    control.value = initialValue || '';
    if (placeholder) control.placeholder = placeholder;
    field.appendChild(control);
    body.appendChild(field);

    const actions = document.createElement('div');
    actions.className = 'modal-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = cancelText || 'Cancel';
    const okBtn = document.createElement('button');
    okBtn.type = 'submit';
    okBtn.textContent = confirmText || 'OK';
    actions.appendChild(cancelBtn);
    actions.appendChild(okBtn);

    form.appendChild(header);
    form.appendChild(body);
    form.appendChild(actions);
    overlay.appendChild(form);
    document.body.appendChild(overlay);
    activePrompt = overlay;

    const cleanup = (value) => destroyActive(value, resolve);
    const onCancel = () => cleanup(null);
    const onSubmit = (ev) => {
      ev.preventDefault();
      cleanup(control.value);
    };
    const onOverlayClick = (ev) => {
      if (ev.target === overlay) onCancel();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') {
        ev.preventDefault();
        onCancel();
      }
    };

    cancelBtn.addEventListener('click', onCancel);
    closeBtn.addEventListener('click', onCancel);
    form.addEventListener('submit', onSubmit);
    overlay.addEventListener('click', onOverlayClick);
    document.addEventListener('keydown', onKeyDown);

    activeCleanup = () => {
      cancelBtn.removeEventListener('click', onCancel);
      closeBtn.removeEventListener('click', onCancel);
      form.removeEventListener('submit', onSubmit);
      overlay.removeEventListener('click', onOverlayClick);
      document.removeEventListener('keydown', onKeyDown);
    };

    control.focus();
    if (selectOnOpen !== false && typeof control.select === 'function') {
      control.select();
    }
  });

  return { open };
}
