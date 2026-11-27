/* Componentes de interface: aviso (toast), menu suspenso, paleta de cores e janela modal. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const ui = (XL.ui = {});

  // ---------- Toast ----------
  let toastBox = null;
  ui.toast = (msg, opts = {}) => {
    if (!toastBox) { toastBox = h('div.ui-toasts'); document.body.appendChild(toastBox); }
    const t = h('div.ui-toast' + (opts.type ? '.' + opts.type : ''), { html: msg });
    toastBox.appendChild(t);
    requestAnimationFrame(() => t.classList.add('show'));
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, opts.duration || 3800);
  };

  // ---------- Popup genérico ----------
  let openPopup = null;
  function closePopup() {
    if (openPopup) { openPopup.el.remove(); openPopup.cleanup(); openPopup = null; }
  }
  ui.closePopup = closePopup;
  ui.popup = (anchor, el) => showPopup(anchor, el);
  function showPopup(anchor, el) {
    closePopup();
    document.body.appendChild(el);
    // âncora pode ser um elemento ou um ponto {x, y} (menu do botão direito)
    const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : { left: anchor.x, right: anchor.x, top: anchor.y, bottom: anchor.y };
    if (!anchor.getBoundingClientRect) anchor = { contains: () => false };
    const w = el.offsetWidth, hh = el.offsetHeight;
    let x = r.left, y = r.bottom + 2;
    if (x + w > window.innerWidth - 8) x = Math.max(8, window.innerWidth - w - 8);
    if (y + hh > window.innerHeight - 8) y = Math.max(8, r.top - hh - 2);
    el.style.left = x + 'px';
    el.style.top = y + 'px';
    const onDown = (e) => { if (!el.contains(e.target) && !anchor.contains(e.target)) closePopup(); };
    const onKey = (e) => { if (e.key === 'Escape') closePopup(); };
    setTimeout(() => { document.addEventListener('mousedown', onDown, true); document.addEventListener('keydown', onKey, true); });
    openPopup = {
      el,
      cleanup: () => { document.removeEventListener('mousedown', onDown, true); document.removeEventListener('keydown', onKey, true); },
    };
  }

  // items: [{label, icon, onClick, checked, sep, hint, data}]
  ui.menu = (anchor, items, opts = {}) => {
    const el = h('div.ui-menu', { role: 'menu' });
    if (opts.title) el.appendChild(h('div.ui-menu-title', opts.title));
    for (const it of items) {
      if (it.sep) { el.appendChild(h('div.ui-menu-sep')); continue; }
      if (it.header) { el.appendChild(h('div.ui-menu-header', it.header)); continue; }
      const b = h('button.ui-menu-item', {
        type: 'button',
        role: 'menuitem',
        disabled: !!it.disabled,
        onmousedown: (e) => e.preventDefault(),
        onclick: () => { closePopup(); it.onClick && it.onClick(); },
      },
        h('span.ui-menu-ico', { html: it.icon || (it.checked ? '✓' : '') }),
        h('span.ui-menu-lbl', it.label),
        it.hint ? h('span.ui-menu-hint', it.hint) : null);
      if (it.data) for (const k in it.data) b.dataset[k] = it.data[k];
      if (it.swatch) b.querySelector('.ui-menu-ico').appendChild(h('span.ui-swatch', { style: it.swatch }));
      el.appendChild(b);
    }
    showPopup(anchor, el);
    return closePopup;
  };

  // ---------- Paleta de cores (cores do tema do Office) ----------
  const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
  const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
  function mix(hex, t) {
    // t > 0 clareia, t < 0 escurece
    const n = parseInt(hex.slice(1), 16);
    let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
    const f = (c) => Math.round(t > 0 ? c + (255 - c) * t : c * (1 + t));
    r = f(r); g = f(g); b = f(b);
    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  }
  ui.palette = (anchor, onPick, opts = {}) => {
    const el = h('div.ui-palette');
    const sw = (c, title) => h('button.ui-sw', {
      type: 'button', title: title || c, style: { background: c }, dataset: { color: c },
      onmousedown: (e) => e.preventDefault(),
      onclick: () => { closePopup(); onPick(c); },
    });
    if (opts.noneLabel) {
      el.appendChild(h('button.ui-pal-none', {
        type: 'button', onmousedown: (e) => e.preventDefault(), onclick: () => { closePopup(); onPick(null); },
      }, opts.noneLabel));
    }
    el.appendChild(h('div.ui-pal-title', 'Cores do Tema'));
    const grid = h('div.ui-pal-grid');
    THEME.forEach((c) => grid.appendChild(sw(c)));
    el.appendChild(grid);
    const tints = h('div.ui-pal-grid.tints');
    [0.8, 0.6, 0.4, -0.25, -0.5].forEach((t) => THEME.forEach((c) => {
      const base = c === '#FFFFFF' ? mix('#FFFFFF', -[0.05, 0.15, 0.25, 0.35, 0.5][[0.8, 0.6, 0.4, -0.25, -0.5].indexOf(t)]) : c === '#000000' ? mix('#000000', [0.5, 0.35, 0.25, 0.15, 0.05][[0.8, 0.6, 0.4, -0.25, -0.5].indexOf(t)]) : mix(c, t);
      tints.appendChild(sw(base));
    }));
    el.appendChild(tints);
    el.appendChild(h('div.ui-pal-title', 'Cores Padrão'));
    const std = h('div.ui-pal-grid');
    STANDARD.forEach((c) => std.appendChild(sw(c)));
    el.appendChild(std);
    showPopup(anchor, el);
  };

  // ---------- Modal ----------
  const openModals = new Set();
  ui.closeAllModals = () => openModals.forEach((m) => m.close(true));
  ui.modal = ({ title, body, buttons = [], width = 480, onClose, className }) => {
    const back = h('div.ui-modal-back');
    const box = h('div.ui-modal' + (className ? '.' + className : ''), { style: { width: width + 'px' }, role: 'dialog', 'aria-modal': 'true' });
    const handle = {};
    const close = (silent) => {
      if (!openModals.has(handle)) return;
      openModals.delete(handle);
      back.remove();
      document.removeEventListener('keydown', onKey, true);
      if (!silent && onClose) onClose();
    };
    const onKey = (e) => {
      if (!back.isConnected) { close(true); return; }
      if (e.key === 'Escape') { e.stopPropagation(); close(); }
    };
    handle.close = close;
    openModals.add(handle);
    document.addEventListener('keydown', onKey, true);
    if (title) box.appendChild(h('div.ui-modal-head', h('div.ui-modal-title', title), h('button.ui-modal-x', { type: 'button', title: 'Fechar', onclick: () => close(), html: '&times;' })));
    const content = h('div.ui-modal-body');
    if (typeof body === 'string') content.innerHTML = body;
    else if (body) content.appendChild(body);
    box.appendChild(content);
    if (buttons.length) {
      const foot = h('div.ui-modal-foot');
      buttons.forEach((b) => foot.appendChild(h('button.btn' + (b.primary ? '.btn-primary' : ''), {
        type: 'button', onclick: () => { if (b.onClick && b.onClick(() => close()) === false) return; if (!b.keepOpen) close(); },
      }, b.label)));
      box.appendChild(foot);
    }
    back.appendChild(box);
    back.addEventListener('mousedown', (e) => { if (e.target === back && !className) close(); });
    document.body.appendChild(back);
    return { close: () => close(), box, content };
  };
})();
