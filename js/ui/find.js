/* Localizar e Substituir (janela flutuante que não bloqueia a planilha). */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const K = XL.SHEET_KEY;

  const F = (XL.Find = { panel: null, app: null });

  F.close = () => {
    if (F.panel) { F.panel.remove(); F.panel = null; }
    if (F.app) { const a = F.app; F.app = null; a.grid.focus(); }
  };

  F.open = (app, mode) => {
    if (F.panel && F.app === app) { F.setMode(mode); F.findInput.focus(); F.findInput.select(); return; }
    F.close();
    F.app = app;
    const findInput = h('input.fd-search', { type: 'text', dataset: { role: 'find' }, placeholder: 'O que você procura?' });
    const replInput = h('input.fd-search', { type: 'text', dataset: { role: 'replace' }, placeholder: 'Trocar por...' });
    const caseCb = h('input', { type: 'checkbox', dataset: { role: 'case' } });
    const wholeCb = h('input', { type: 'checkbox', dataset: { role: 'whole' } });
    const lookIn = h('select.fd-cat', { dataset: { role: 'lookin' } }, h('option', { value: 'formulas' }, 'Fórmulas'), h('option', { value: 'values' }, 'Valores'));
    const status = h('div.fr-status');
    const results = h('div.fr-results');
    const tabFind = h('button.fr-tab', { type: 'button', dataset: { tab: 'find' } }, 'Localizar');
    const tabRepl = h('button.fr-tab', { type: 'button', dataset: { tab: 'replace' } }, 'Substituir');
    const replRow = h('div.fr-row', h('label.fd-lbl', 'Substituir por:'), replInput);
    const btn = (label, role, fn, primary) => h('button.btn.btn-sm' + (primary ? '.btn-primary' : '.btn-ghost'), { type: 'button', dataset: { role }, onclick: fn }, label);
    const bReplAll = btn('Substituir tudo', 'replaceall', () => F.replaceAll());
    const bRepl = btn('Substituir', 'replaceone', () => F.replaceOne());
    const panel = h('div.xl-find', { role: 'dialog', 'aria-label': 'Localizar e substituir' },
      h('div.fr-head', h('div.fr-tabs', tabFind, tabRepl), h('button.ui-modal-x', { type: 'button', title: 'Fechar (Esc)', onclick: F.close, html: '&times;' })),
      h('div.fr-body',
        h('div.fr-row', h('label.fd-lbl', 'Localizar:'), findInput),
        replRow,
        h('div.fr-opts',
          h('label.sd-check', caseCb, ' Diferenciar maiúsculas e minúsculas'),
          h('label.sd-check', wholeCb, ' Coincidir conteúdo da célula inteira'),
          h('label.sd-check', 'Examinar: ', lookIn)),
        h('div.fr-btns', bReplAll, bRepl, btn('Localizar tudo', 'findall', () => F.findAll()), btn('Localizar próxima', 'findnext', () => F.findNext(), true)),
        status, results));
    Object.assign(F, { panel, findInput, replInput, caseCb, wholeCb, lookIn, status, results, replRow, tabFind, tabRepl, bReplAll, bRepl });
    tabFind.addEventListener('click', () => F.setMode('find'));
    tabRepl.addEventListener('click', () => F.setMode('replace'));
    panel.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); F.close(); }
      if (e.key === 'Enter' && (e.target === findInput || e.target === replInput)) { e.preventDefault(); F.findNext(); }
    });
    document.body.appendChild(panel);
    const r = app.grid.el.getBoundingClientRect();
    panel.style.top = Math.max(8, r.top + 8) + 'px';
    panel.style.left = Math.max(8, r.right - panel.offsetWidth - 24) + 'px';
    F.setMode(mode);
    findInput.focus();
    app.emit('findopen', { mode });
  };

  F.setMode = (mode) => {
    F.mode = mode;
    F.tabFind.classList.toggle('on', mode === 'find');
    F.tabRepl.classList.toggle('on', mode === 'replace');
    F.replRow.style.display = mode === 'replace' ? '' : 'none';
    F.bReplAll.style.display = mode === 'replace' ? '' : 'none';
    F.bRepl.style.display = mode === 'replace' ? '' : 'none';
  };

  const opts = () => ({ q: F.findInput.value, matchCase: F.caseCb.checked, whole: F.wholeCb.checked, look: F.lookIn.value });
  const textFor = (S, r, c, look) => {
    const cell = S.cell(r, c);
    if (!cell || cell.kind === 'empty') return null;
    return look === 'values' ? S.textOf(r, c) : XL.editTextFor(cell);
  };
  const matches = (text, o) => {
    if (text === null) return false;
    const a = o.matchCase ? text : text.toLocaleLowerCase('pt-BR');
    const b = o.matchCase ? o.q : o.q.toLocaleLowerCase('pt-BR');
    return o.whole ? a === b : a.includes(b);
  };
  // Células que combinam, na ordem linha por linha. Com várias células selecionadas, só procura nelas.
  function allMatches(o) {
    const app = F.app;
    const S = app.sheet;
    const g = app.grid.range();
    const scoped = g.r1 !== g.r2 || g.c1 !== g.c2;
    const out = [];
    for (const k of S.cells.keys()) {
      const r = Math.floor(k / K), c = k % K;
      if (scoped && (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2)) continue;
      if (matches(textFor(S, r, c, o.look), o)) out.push({ r, c });
    }
    return out.sort((a, b) => a.r - b.r || a.c - b.c);
  }
  const say = (msg, kind) => { F.status.textContent = msg; F.status.className = 'fr-status' + (kind ? ' ' + kind : ''); };

  F.findNext = () => {
    const o = opts();
    if (!o.q) { say('Digite o que deseja localizar.', 'warn'); return; }
    const list = allMatches(o);
    F.results.innerHTML = '';
    if (!list.length) { say('Não encontramos o que você está procurando.', 'warn'); F.app.emit('find', { found: false, text: o.q }); return; }
    const { ar, ac } = F.app.grid.sel;
    const next = list.find((m) => m.r > ar || (m.r === ar && m.c > ac)) || list[0];
    const g = F.app.grid.range();
    const scoped = g.r1 !== g.r2 || g.c1 !== g.c2;
    if (scoped) { F.app.grid.sel.ar = next.r; F.app.grid.sel.ac = next.c; F.app.grid.renderSel(); F.app.onSelectionChange(true); }
    else F.app.grid.setSel(next.r, next.c);
    say(`Encontrado em ${XL.toRef(next.r, next.c)} (${list.indexOf(next) + 1} de ${list.length}).`, 'ok');
    F.app.emit('find', { found: true, ref: XL.toRef(next.r, next.c), text: o.q });
  };

  F.findAll = () => {
    const o = opts();
    if (!o.q) { say('Digite o que deseja localizar.', 'warn'); return; }
    const list = allMatches(o);
    const S = F.app.sheet;
    F.results.innerHTML = '';
    if (!list.length) { say('Não encontramos o que você está procurando.', 'warn'); return; }
    say(`${list.length} célula(s) encontrada(s). Clique em uma para ir até ela.`, 'ok');
    F.results.appendChild(h('div.fr-res.head', h('span', 'Célula'), h('span', 'Valor')));
    list.forEach((m) => {
      F.results.appendChild(h('button.fr-res', { type: 'button', onclick: () => F.app.grid.setSel(m.r, m.c) },
        h('span', XL.toRef(m.r, m.c)), h('span', S.textOf(m.r, m.c))));
    });
    F.app.emit('findall', { count: list.length, text: o.q });
  };

  function replaceIn(text, o, repl) {
    if (o.whole) return repl;
    const re = new RegExp(o.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), o.matchCase ? 'g' : 'gi');
    return text.replace(re, () => repl);
  }
  function applyReplace(r, c, o, repl) {
    const S = F.app.sheet;
    const cell = S.cell(r, c);
    const src = XL.editTextFor(cell);
    const nt = replaceIn(o.look === 'values' && cell.kind !== 'formula' ? S.textOf(r, c) : src, o, repl);
    try { S.setInput(r, c, nt); return true; } catch (e) { return false; }
  }

  F.replaceOne = () => {
    const o = opts();
    if (!o.q) { say('Digite o que deseja localizar.', 'warn'); return; }
    const app = F.app;
    const { ar, ac } = app.grid.sel;
    if (matches(textFor(app.sheet, ar, ac, o.look), o)) {
      app.pushUndo();
      applyReplace(ar, ac, o, F.replInput.value);
      app.render();
      app.emit('replace', { count: 1, find: o.q, repl: F.replInput.value });
      app.emit('change', {});
    }
    F.findNext();
  };

  F.replaceAll = () => {
    const o = opts();
    if (!o.q) { say('Digite o que deseja localizar.', 'warn'); return; }
    const app = F.app;
    const list = allMatches(o);
    if (!list.length) { say('Não encontramos nada para substituir.', 'warn'); return; }
    app.pushUndo();
    let n = 0;
    for (const m of list) if (applyReplace(m.r, m.c, o, F.replInput.value)) n++;
    if (app.sheet.filter) app.sheet.applyFilter();
    app.render();
    F.results.innerHTML = '';
    say(`Concluído. Fizemos ${n} substituição(ões).`, 'ok');
    app.emit('replace', { count: n, find: o.q, repl: F.replInput.value, all: true });
    app.emit('change', {});
  };
})();
