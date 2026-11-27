/* Diálogos de dados: Classificar, menu do AutoFiltro, Personalizar AutoFiltro, Remover Duplicatas e Imprimir Títulos. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const D = (XL.Dialogs = {});

  const colLabel = (S, g, c, header) => (header && S.textOf(g.r1, c)) || 'Coluna ' + XL.numToCol(c + 1);
  const isNumericCol = (S, g, c, header) => {
    let nums = 0, texts = 0;
    for (let r = g.r1 + (header ? 1 : 0); r <= g.r2; r++) {
      const v = S.value(r, c);
      if (typeof v === 'number') nums++; else if (v !== null && v !== '') texts++;
    }
    return nums > 0 && nums >= texts;
  };

  // ================= Classificar (vários níveis) =================
  D.sort = (app) => {
    const d = app.dataRegion();
    if (!d) { XL.ui.toast('Clique em uma célula dentro dos dados que você quer classificar.'); return; }
    const S = app.sheet;
    const g = d.g;
    const levels = [{ col: Math.max(g.c1, Math.min(g.c2, app.grid.sel.ac)), desc: false }];
    app.grid.setRange(g);
    let header = d.header;
    const list = h('div.sd-levels');
    const headerCb = h('input', { type: 'checkbox' });
    headerCb.checked = header;
    headerCb.addEventListener('change', () => { header = headerCb.checked; draw(); });

    function draw() {
      list.innerHTML = '';
      list.appendChild(h('div.sd-row.sd-head', h('span', 'Coluna'), h('span', 'Ordem'), h('span')));
      levels.forEach((lv, i) => {
        const colSel = h('select.fd-cat', { dataset: { level: i, role: 'col' } });
        for (let c = g.c1; c <= g.c2; c++) colSel.appendChild(h('option', { value: String(c) }, colLabel(S, g, c, header)));
        colSel.value = String(lv.col);
        const numeric = isNumericCol(S, g, lv.col, header);
        const ordSel = h('select.fd-cat', { dataset: { level: i, role: 'order' } },
          h('option', { value: 'asc' }, numeric ? 'Do Menor para o Maior' : 'De A a Z'),
          h('option', { value: 'desc' }, numeric ? 'Do Maior para o Menor' : 'De Z a A'));
        ordSel.value = lv.desc ? 'desc' : 'asc';
        colSel.addEventListener('change', () => { lv.col = Number(colSel.value); draw(); });
        ordSel.addEventListener('change', () => { lv.desc = ordSel.value === 'desc'; });
        const del = h('button.sd-del', { type: 'button', title: 'Excluir este nível', html: '&times;' });
        del.disabled = levels.length === 1;
        del.addEventListener('click', () => { levels.splice(i, 1); draw(); });
        list.appendChild(h('div.sd-row', h('span.sd-lbl', i === 0 ? 'Classificar por' : 'E depois por'), colSel, ordSel, del));
      });
    }
    const addBtn = h('button.btn.btn-soft.btn-sm.sd-add', { type: 'button' }, '+ Adicionar Nível');
    addBtn.addEventListener('click', () => {
      const used = new Set(levels.map((l) => l.col));
      let c = g.c1;
      while (used.has(c) && c < g.c2) c++;
      levels.push({ col: c, desc: false });
      draw();
    });
    draw();
    XL.ui.modal({
      title: 'Classificar', width: 560,
      body: h('div.sd',
        h('div.sd-top', addBtn, h('label.sd-check', headerCb, ' Meus dados contêm cabeçalhos')),
        list,
        h('p.cd-note', `Intervalo: ${XL.rangeToStr(g)}. Cada linha se move inteira, então os dados continuam corretos.`)),
      buttons: [{ label: 'OK', primary: true, onClick: () => app.sortRange(g, levels.map((l) => ({ col: l.col, desc: l.desc })), header) }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
  };

  // ================= Menu do AutoFiltro =================
  D.filterMenu = (app, col, anchor) => {
    const S = app.sheet;
    const f = S.filter;
    if (!f) return;
    const g = f.g;
    const header = S.textOf(g.r1, col) || 'Coluna ' + XL.numToCol(col + 1);
    const crit = f.crit[col];
    const vals = new Map();
    let blanks = false, numeric = true;
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      if (!S.rowPasses(r, col)) continue;
      const t = S.textOf(r, col);
      if (t === '') { blanks = true; continue; }
      const v = S.value(r, col);
      if (typeof v !== 'number') numeric = false;
      if (!vals.has(t)) vals.set(t, v);
    }
    const items = [...vals.entries()].sort((a, b) => XL.Formula.compareValues(a[1], b[1]));
    const checkedInit = (t) => !crit || (crit.type === 'values' ? crit.values.includes(t) : S.passes(findRow(t), crit, col));
    function findRow(t) { for (let r = g.r1 + 1; r <= g.r2; r++) if (S.textOf(r, col) === t) return r; return g.r1 + 1; }

    const el = h('div.ui-menu.xl-fmenu');
    const close = () => XL.ui.closePopup();
    const act = (label, fn, opts = {}) => h('button.ui-menu-item', {
      type: 'button', disabled: !!opts.disabled, dataset: opts.data || {},
      onmousedown: (e) => e.preventDefault(),
      onclick: () => { close(); fn(); },
    }, h('span.ui-menu-ico', { html: opts.icon || '' }), h('span.ui-menu-lbl', label));

    const app2 = app;
    const sortCol = (desc) => app2.sortRange(Object.assign({}, g), [{ col, desc }], true);
    el.append(
      act(numeric ? 'Classificar do Menor para o Maior' : 'Classificar de A a Z', () => sortCol(false), { icon: XL.icons.sortAZ, data: { cmd: 'fsort', arg: 'asc' } }),
      act(numeric ? 'Classificar do Maior para o Menor' : 'Classificar de Z a A', () => sortCol(true), { icon: XL.icons.sortZA, data: { cmd: 'fsort', arg: 'desc' } }),
      h('div.ui-menu-sep'),
      act(`Limpar Filtro de “${header}”`, () => app.setFilterCriteria(col, null), { disabled: !crit, data: { cmd: 'fclear' } }));

    // Filtros de Número / Texto
    const ops = numeric
      ? [['=', 'É Igual a...'], ['<>', 'É Diferente de...'], ['>', 'É Maior do que...'], ['>=', 'É Maior ou Igual a...'], ['<', 'É Menor do que...'], ['<=', 'É Menor ou Igual a...']]
      : [['=', 'É Igual a...'], ['<>', 'É Diferente de...'], ['comeca', 'Começa com...'], ['termina', 'Termina com...'], ['contem', 'Contém...'], ['naocontem', 'Não Contém...']];
    const sub = h('div.xl-fsub', { style: { display: 'none' } },
      ops.map(([op, label]) => act(label, () => D.customFilter(app, col, op, header, numeric), { data: { cmd: 'fcond', arg: op } })));
    const subBtn = h('button.ui-menu-item.xl-fsub-btn', {
      type: 'button', dataset: { cmd: 'fcondmenu' }, onmousedown: (e) => e.preventDefault(),
      onclick: () => { sub.style.display = sub.style.display === 'none' ? '' : 'none'; },
    }, h('span.ui-menu-ico', { html: crit && crit.type === 'cond' ? '✓' : '' }), h('span.ui-menu-lbl', numeric ? 'Filtros de Número' : 'Filtros de Texto'), h('span.ui-menu-hint', '›'));
    el.append(subBtn, sub);

    // lista de valores
    const search = h('input.xl-fsearch', { type: 'search', placeholder: 'Pesquisar' });
    const listEl = h('div.xl-flist');
    const boxes = [];
    const mk = (label, value, checked, cls) => {
      const cb = h('input', { type: 'checkbox' });
      cb.checked = checked;
      const lab = h('label.xl-fitem' + (cls ? '.' + cls : ''), { dataset: { value } }, cb, h('span', label));
      return { cb, lab, value };
    };
    const all = mk('(Selecionar Tudo)', '__all__', true, 'all');
    listEl.appendChild(all.lab);
    for (const [t] of items) { const b = mk(t, t, checkedInit(t)); boxes.push(b); listEl.appendChild(b.lab); }
    let blankBox = null;
    if (blanks) { blankBox = mk('(Vazias)', '__blank__', !crit || !!crit.blanks); boxes.push(blankBox); listEl.appendChild(blankBox.lab); }
    const syncAll = () => {
      const vis = boxes.filter((b) => b.lab.style.display !== 'none');
      all.cb.checked = vis.every((b) => b.cb.checked);
      all.cb.indeterminate = !all.cb.checked && vis.some((b) => b.cb.checked);
    };
    boxes.forEach((b) => b.cb.addEventListener('change', syncAll));
    all.cb.addEventListener('change', () => {
      boxes.forEach((b) => { if (b.lab.style.display !== 'none') b.cb.checked = all.cb.checked; });
      syncAll();
    });
    search.addEventListener('input', () => {
      const q = search.value.trim().toLocaleLowerCase('pt-BR');
      boxes.forEach((b) => {
        const show = !q || (b.value !== '__blank__' && b.value.toLocaleLowerCase('pt-BR').includes(q));
        b.lab.style.display = show ? '' : 'none';
        if (q) b.cb.checked = show;
      });
      syncAll();
    });
    syncAll();
    const ok = h('button.btn.btn-primary.btn-sm.xl-fmenu-ok', { type: 'button' }, 'OK');
    ok.addEventListener('click', () => {
      const q = search.value.trim();
      const chosen = boxes.filter((b) => b.cb.checked && (!q || b.lab.style.display !== 'none'));
      const values = chosen.filter((b) => b !== blankBox).map((b) => b.value);
      const wantBlank = !!blankBox && chosen.includes(blankBox);
      close();
      if (values.length === items.length && (!blankBox || wantBlank)) app.setFilterCriteria(col, null);
      else app.setFilterCriteria(col, { type: 'values', values, blanks: wantBlank });
      app.grid.focus();
    });
    const cancel = h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => { close(); app.grid.focus(); } }, 'Cancelar');
    el.append(h('div.ui-menu-sep'), search, listEl, h('div.xl-fmenu-foot', ok, cancel));
    XL.ui.popup(anchor, el);
    app.emit('filtermenu', { col: XL.numToCol(col + 1) });
  };

  D.customFilter = (app, col, op, header, numeric) => {
    const ops = numeric
      ? [['=', 'é igual a'], ['<>', 'é diferente de'], ['>', 'é maior do que'], ['>=', 'é maior ou igual a'], ['<', 'é menor do que'], ['<=', 'é menor ou igual a']]
      : [['=', 'é igual a'], ['<>', 'é diferente de'], ['comeca', 'começa com'], ['termina', 'termina com'], ['contem', 'contém'], ['naocontem', 'não contém']];
    const sel = h('select.fd-cat', { dataset: { role: 'op' } }, ops.map(([v, l]) => h('option', { value: v }, l)));
    sel.value = op;
    const inp = h('input.fd-search', { type: 'text', dataset: { role: 'val' }, placeholder: numeric ? 'Ex.: 5000' : 'Ex.: Sul' });
    XL.ui.modal({
      title: 'Personalizar AutoFiltro', width: 440,
      body: h('div.cf', h('p', 'Mostrar linhas onde: ', h('b', header)), h('div.cf-row', sel, inp)),
      buttons: [{
        label: 'OK', primary: true,
        onClick: () => {
          if (inp.value.trim() === '') { XL.ui.toast('Digite um valor para comparar.'); return false; }
          app.setFilterCriteria(col, { type: 'cond', op: sel.value, value: inp.value.trim() });
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
    setTimeout(() => inp.focus(), 0);
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); inp.closest('.ui-modal').querySelector('.btn-primary').click(); } });
  };

  // ================= Remover Duplicatas =================
  D.removeDuplicates = (app) => {
    const d = app.dataRegion();
    if (!d) { XL.ui.toast('Clique em uma célula dentro da tabela onde quer remover duplicatas.'); return; }
    const S = app.sheet;
    const g = d.g;
    app.grid.setRange(g);
    let header = d.header;
    const cols = [];
    for (let c = g.c1; c <= g.c2; c++) cols.push({ c, on: true });
    const listEl = h('div.dd-list');
    const draw = () => {
      listEl.innerHTML = '';
      cols.forEach((x) => {
        const cb = h('input', { type: 'checkbox', dataset: { col: XL.numToCol(x.c + 1) } });
        cb.checked = x.on;
        cb.addEventListener('change', () => { x.on = cb.checked; });
        listEl.appendChild(h('label.xl-fitem', cb, h('span', colLabel(S, g, x.c, header))));
      });
    };
    const headerCb = h('input', { type: 'checkbox' });
    headerCb.checked = header;
    headerCb.addEventListener('change', () => { header = headerCb.checked; draw(); });
    const setAll = (v) => { cols.forEach((x) => (x.on = v)); draw(); };
    draw();
    XL.ui.modal({
      title: 'Remover Duplicatas', width: 440,
      body: h('div.dd',
        h('p', 'Para excluir valores duplicados, marque as colunas que devem ser comparadas. Linhas com os mesmos valores nessas colunas são consideradas repetidas — a primeira é mantida.'),
        h('div.dd-top',
          h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => setAll(true) }, 'Selecionar Tudo'),
          h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => setAll(false) }, 'Desmarcar Tudo'),
          h('label.sd-check', headerCb, ' Meus dados contêm cabeçalhos')),
        listEl),
      buttons: [{
        label: 'OK', primary: true,
        onClick: () => {
          const chosen = cols.filter((x) => x.on).map((x) => x.c);
          if (!chosen.length) { XL.ui.toast('Marque pelo menos uma coluna.'); return false; }
          setTimeout(() => app.removeDuplicates(g, chosen, header), 0);
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
  };

  // ================= Imprimir Títulos =================
  D.printTitles = (app) => {
    const S = app.sheet;
    const t = S.page.titleRows;
    const inp = h('input.fd-search', { type: 'text', dataset: { role: 'titles' }, placeholder: 'Ex.: $1:$1', value: t ? `$${t.r1 + 1}:$${t.r2 + 1}` : '' });
    XL.ui.modal({
      title: 'Imprimir Títulos', width: 420,
      body: h('div',
        h('p', 'As linhas escolhidas aqui aparecem no topo de ', h('b', 'todas'), ' as páginas impressas — perfeito para repetir o cabeçalho da tabela.'),
        h('label.fd-lbl', 'Linhas a repetir na parte superior:'), inp,
        h('p.cd-note', 'Use $1:$1 para repetir só a linha 1, ou $1:$2 para as linhas 1 e 2. Deixe vazio para não repetir.')),
      buttons: [{
        label: 'OK', primary: true,
        onClick: () => {
          const v = inp.value.replace(/\s|\$/g, '');
          if (!v) { app.setPage({ titleRows: null }); return; }
          const m = /^(\d+)(?::(\d+))?$/.exec(v);
          if (!m) { XL.ui.toast('Formato inválido. Use algo como <b>$1:$1</b>.'); return false; }
          const a = Number(m[1]) - 1, b = Number(m[2] || m[1]) - 1;
          if (a < 0 || b < 0) { XL.ui.toast('Formato inválido. Use algo como <b>$1:$1</b>.'); return false; }
          app.setPage({ titleRows: { r1: Math.min(a, b), r2: Math.max(a, b) } });
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
    setTimeout(() => inp.focus(), 0);
  };
})();
