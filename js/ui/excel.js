/* ExcelApp: monta o simulador completo (faixa de opções, barra de fórmulas, grade, barra de status)
   e centraliza os comandos. Emite eventos que o guia das lições observa. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const CELL_STYLES = {
    bom: { fill: '#C6EFCE', color: '#006100' },
    ruim: { fill: '#FFC7CE', color: '#9C0006' },
    neutro: { fill: '#FFEB9C', color: '#9C5700' },
    titulo: { b: true, size: 18, color: '#44546A' },
    total: { b: true, border: { t: 1, b: 1 } },
    cabecalho: { fill: '#107C41', color: '#FFFFFF', b: true },
    normal: { fill: null, color: null, b: null, i: null, u: null, size: null, border: null, font: null },
  };
  const FREQUENT = ['SOMA', 'MÉDIA', 'SE', 'PROCV', 'PROCX', 'CONT.SE', 'SOMASE', 'MÁXIMO', 'MÍNIMO', 'CONT.NÚM', 'CONT.VALORES', 'HOJE'];
  const clone = (o) => JSON.parse(JSON.stringify(o || {}));
  // comandos que abrem uma janela (não devolvem o foco para a grade)
  const NOFOCUS = ['fx', 'insert-fn', 'colwidth', 'shortcuts', 'find', 'replace', 'sort-dialog', 'dedupe', 'print-preview',
    'print-titles', 'insert-dialog', 'delete-dialog', 'sheet-rename', 'goto', 'sheet-delete', 'format-cells-dialog', 'texttocols',
    'table-create', 'cf', 'dv-dialog', 'chart-title', 'table-rename'];

  class ExcelApp extends XL.Emitter {
    constructor(container, opts = {}) {
      super();
      this.wb = new XL.Workbook(opts.rows || 100, opts.cols || 26);
      this.undoStack = [];
      this.redoStack = [];
      this.clip = null;
      this.clipTSV = '';
      this.painterFmt = null;
      this.formulaBarVisible = true;
      this.build(container);
      this.grid.setSel(0, 0);
      this.render();
    }

    // planilha ativa da pasta de trabalho
    get sheet() { return this.wb.active; }

    // ---------- Montagem ----------
    build(container) {
      this.root = h('div.xl-app');
      this.ribbon = new XL.Ribbon(this);

      this.namebox = h('input.xl-namebox', { type: 'text', 'aria-label': 'Caixa de Nome', spellcheck: 'false', title: 'Caixa de Nome: mostra a célula ativa. Digite um endereço (ex.: H10) e Enter para ir até ele.' });
      this.namebox.addEventListener('focus', () => this.namebox.select());
      this.namebox.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); this.gotoRef(this.namebox.value); }
        if (e.key === 'Escape') { e.preventDefault(); this.updateNameBox(); this.grid.focus(); }
      });
      this.namebox.addEventListener('blur', () => this.updateNameBox());

      this.fbar = h('input.xl-finput', { type: 'text', 'aria-label': 'Barra de Fórmulas', spellcheck: 'false', autocomplete: 'off' });
      this.fbar.addEventListener('focus', () => { if (!this.grid.editing) this.grid.beginEdit('edit', undefined, 'bar'); });
      this.fbar.addEventListener('input', () => this.grid.onBarInput());
      this.fbar.addEventListener('keydown', (e) => { if (this.grid.editing) this.grid.onEditKey(e); });
      this.fbar.addEventListener('click', () => this.grid.onEditInput(true));
      this.fbar.addEventListener('keyup', (e) => { if (/Arrow|Home|End/.test(e.key)) this.grid.onEditInput(true); });

      const fbtn = (label, title, fn, cls) => h('button.xl-fbtn' + (cls ? '.' + cls : ''), {
        type: 'button', title, html: label, onmousedown: (e) => e.preventDefault(), onclick: fn,
      });
      this.fxBtn = fbtn('<i>fx</i>', 'Inserir Função', () => this.exec('fx'), 'fx');
      this.fbarRow = h('div.xl-fbar',
        this.namebox,
        h('span.xl-fdots', '⋮'),
        h('div.xl-fbtns',
          fbtn('✕', 'Cancelar (Esc)', () => this.grid.cancel(), 'cancel'),
          fbtn('✓', 'Inserir (Enter)', () => this.grid.commit(), 'ok'),
          this.fxBtn),
        this.fbar);

      this.grid = new XL.Grid(this);

      this.tabsBox = h('div.xl-sheettabs');
      this.sheetbar = h('div.xl-sheetbar',
        h('button.xl-sheetnav', { type: 'button', title: 'Planilha anterior (Ctrl+PageUp)', onmousedown: (e) => e.preventDefault(), onclick: () => this.showSheetOffset(-1) }, '◀'),
        h('button.xl-sheetnav', { type: 'button', title: 'Próxima planilha (Ctrl+PageDown)', onmousedown: (e) => e.preventDefault(), onclick: () => this.showSheetOffset(1) }, '▶'),
        this.tabsBox,
        h('button.xl-sheetadd', { type: 'button', title: 'Nova planilha', onmousedown: (e) => e.preventDefault(), onclick: () => this.exec('sheet-add') }, '+'));

      this.modeEl = h('span.xl-status-mode', 'Pronto');
      this.filterEl = h('span.xl-status-filter');
      this.statsEl = h('span.xl-status-stats');
      this.warnEl = h('span.xl-status-warn');
      this.statusbar = h('div.xl-status', this.modeEl, this.filterEl, this.warnEl, this.statsEl);

      this.root.append(this.ribbon.tabsEl, this.ribbon.bodyEl, this.fbarRow, this.grid.el, this.sheetbar, this.statusbar);
      container.appendChild(this.root);
      this.ribbon.fit();
      this.renderTabs();
      this.onResize = () => { this.grid.positionOverlays(); this.ribbon.fit(); };
      window.addEventListener('resize', this.onResize);
    }

    destroy() {
      if (XL.Find && XL.Find.app === this) { XL.Find.app = null; XL.Find.close(); }
      if (XL.Print && XL.Print.current && XL.Print.current.app === this) XL.Print.current.close(true);
      window.removeEventListener('resize', this.onResize);
      this.grid.destroy();
      XL.ui.closePopup();
      this.root.remove();
    }

    // ---------- Estado visível ----------
    render() {
      this.grid.render();
      this.updateFormulaBar();
      this.updateStatus();
      this.ribbon.sync();
    }
    setMode(m) { this.modeEl.textContent = m; }
    updateNameBox(text) {
      if (document.activeElement === this.namebox && text === undefined) return;
      this.namebox.value = text || XL.toRef(this.grid.sel.ar, this.grid.sel.ac);
    }
    updateFormulaBar() {
      if (this.grid.editing) return;
      const { ar, ac } = this.grid.sel;
      this.fbar.value = XL.editTextFor(this.sheet.cell(ar, ac));
    }
    updateStatus() {
      const g = this.grid.range();
      let count = 0, n = 0, sum = 0, fmt = null;
      if (g.r1 !== g.r2 || g.c1 !== g.c2) {
        for (const [k, cell] of this.sheet.cells) {
          if (cell.kind === 'empty') continue;
          const r = Math.floor(k / XL.SHEET_KEY), c = k % XL.SHEET_KEY;
          if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) continue;
          const v = this.sheet.value(r, c);
          if (v === null || v === '') continue;
          count++;
          if (typeof v === 'number') { n++; sum += v; if (!fmt) fmt = cell.fmt; }
        }
      }
      const f = (x) => XL.formatNumber(x, fmt && fmt.numFmt !== 'geral' ? fmt : { numFmt: 'numero', dec: Number.isInteger(x) ? 0 : 2 });
      this.statsEl.innerHTML = count > 1
        ? (n ? `<span>Média: <b>${f(sum / n)}</b></span>` : '') + `<span>Contagem: <b>${count}</b></span>` + (n ? `<span>Soma: <b>${f(sum)}</b></span>` : '')
        : '';
      this.warnEl.textContent = this.sheet.circular ? '⚠ Referência circular' : '';
      const fs = this.sheet.filterStats();
      this.filterEl.textContent = fs && fs.shown < fs.total ? `${fs.shown} de ${fs.total} registros encontrados` : fs ? 'Filtro ativo' : '';
    }

    onSelectionChange(final) {
      const g = this.grid.range();
      if (!this.grid.isAway()) this.sheet.sel = Object.assign({}, this.grid.sel);
      const multi = g.r1 !== g.r2 || g.c1 !== g.c2;
      if (!final && multi) this.updateNameBox(`${g.r2 - g.r1 + 1}L x ${g.c2 - g.c1 + 1}C`);
      else this.updateNameBox();
      this.updateFormulaBar();
      this.updateStatus();
      this.ribbon.sync();
      if (final) this.ribbon.updateContext();
      const data = { ref: XL.toRef(this.grid.sel.ar, this.grid.sel.ac), range: XL.rangeToStr(g) };
      this.emit(final ? 'select' : 'selecting', data);
    }

    // ---------- Desfazer ----------
    pushUndo() {
      this.undoStack.push(this.wb.snapshot());
      if (this.undoStack.length > 100) this.undoStack.shift();
      this.redoStack = [];
    }
    undo() {
      if (this.grid.editing) { this.grid.cancel(); return; }
      if (!this.undoStack.length) { XL.ui.toast('Nada para desfazer.'); return; }
      this.redoStack.push(this.wb.snapshot());
      this.wb.restore(this.undoStack.pop());
      this.afterWorkbookChange();
      this.emit('undo', {});
      this.emit('change', {});
    }
    redo() {
      if (!this.redoStack.length) return;
      this.undoStack.push(this.wb.snapshot());
      this.wb.restore(this.redoStack.pop());
      this.afterWorkbookChange();
      this.emit('redo', {});
      this.emit('change', {});
    }
    afterWorkbookChange() {
      this.grid.reloadSheet();
      this.renderTabs();
      this.updateFormulaBar();
      this.updateNameBox();
      this.updateStatus();
      this.ribbon.sync();
    }

    // ---------- Entrada de dados ----------
    commitInput(r, c, text, sheet) {
      sheet = sheet || this.sheet;
      let t = text;
      if (t.length > 1 && t[0] === '=') {
        try {
          const norm = XL.Formula.normalize(t, { sheetName: (n) => { const s = this.wb.byName(n); return s && s.name; } });
          const ast = XL.Formula.parse(norm.slice(1));
          const issue = XL.Formula.validate(ast);
          if (issue) return { ok: false, msg: issue };
          const missing = XL.Formula.tokenize(norm.slice(1), true).find((tk) => tk.sheet && !this.wb.byName(tk.sheet));
          if (missing) return { ok: false, msg: `Não existe uma planilha chamada "${missing.sheet}" nesta pasta de trabalho. Confira o nome na aba lá embaixo.` };
          t = norm;
        } catch (e) {
          return { ok: false, msg: e.message };
        }
      }
      if (t[0] !== '=' && sheet.validationAt) {
        const rule = sheet.validationAt(r, c);
        if (rule && !sheet.checkValidation(rule, t)) return { ok: false, dv: rule };
      }
      const before = sheet.cell(r, c);
      const prevInput = before ? XL.editTextFor(before) : '';
      if (t === prevInput && (!before || before.kind !== 'formula')) {
        this.emit('commit', { r, c, ref: XL.toRef(r, c), input: t, sheet: sheet.name, unchanged: true });
        return { ok: true };
      }
      this.pushUndo();
      sheet.setInput(r, c, t);
      if (sheet.tables.length && this.afterTableCommit) this.afterTableCommit(sheet, r, c, t);
      if (sheet.filter) sheet.extendFilter();
      this.render();
      this.emit('commit', { r, c, ref: XL.toRef(r, c), input: t, sheet: sheet.name });
      this.emit('change', {});
      return { ok: true };
    }

    // ---------- Comandos ----------
    exec(cmd, arg) {
      const g = this.grid;
      if (g.editing && !['insert-fn', 'fx'].includes(cmd)) {
        if (!g.commit()) return;
      }
      XL.ui.closePopup();
      const S = this.sheet;
      const { ar, ac } = g.sel;
      const cur = S.fmt(ar, ac);
      switch (cmd) {
        case 'bold': this.applyFmt({ b: cur.b ? null : true }); break;
        case 'italic': this.applyFmt({ i: cur.i ? null : true }); break;
        case 'underline': this.applyFmt({ u: cur.u ? null : true }); break;
        case 'font': this.applyFmt({ font: arg === 'Aptos Narrow' ? null : arg }); break;
        case 'size': this.applyFmt({ size: Number(arg) === 11 ? null : Number(arg) }); break;
        case 'grow': case 'shrink': {
          const s = cur.size || 11;
          const i = SIZES.findIndex((x) => x >= s);
          const ni = cmd === 'grow' ? Math.min(SIZES.length - 1, (SIZES[i] === s ? i + 1 : i)) : Math.max(0, i - 1);
          this.applyFmt({ size: SIZES[ni] === 11 ? null : SIZES[ni] });
          break;
        }
        case 'fill': this.applyFmt({ fill: arg || null }); break;
        case 'color': this.applyFmt({ color: arg || null }); break;
        case 'align': this.applyFmt({ align: cur.align === arg ? null : arg }); break;
        case 'valign': this.applyFmt({ valign: cur.valign === arg || arg === 'bottom' ? null : arg }); break;
        case 'wrap': this.applyFmt({ wrap: cur.wrap ? null : true }); break;
        case 'border': this.applyBorder(arg || 'bottom'); break;
        case 'numfmt':
          if (arg === 'custom') { XL.Dialogs.formatCells(this); break; }
          this.applyFmt({ numFmt: arg === 'geral' ? null : arg, dec: null, grouping: null, code: null }); break;
        case 'currency': this.applyFmt({ numFmt: 'contabil', dec: null, grouping: null, code: null }); break;
        case 'percent': this.applyFmt({ numFmt: 'porcentagem', dec: null, grouping: null, code: null }); break;
        case 'thousands': this.applyFmt({ numFmt: 'numero', dec: null, grouping: null, code: null }); break;
        case 'dec-inc': this.changeDecimals(1); break;
        case 'dec-dec': this.changeDecimals(-1); break;
        case 'cellstyle': this.applyFmt(CELL_STYLES[arg] || {}); break;
        case 'clear': this.clearRange(arg || 'all'); break;
        case 'copy': this.copy(false); break;
        case 'cut': this.copy(true); break;
        case 'paste': this.paste(arg); break;
        case 'painter':
          this.painterFmt = clone(cur);
          this.grid.el.classList.add('painter');
          XL.ui.toast('🖌️ Pincel ativo: clique ou arraste sobre as células que vão receber o formato.');
          break;
        case 'autosum': this.autosum(arg || 'SOMA'); break;
        case 'fx': this.openFunctionDialog(); break;
        case 'insert-fn': this.insertFunction(arg); break;
        case 'show-formulas': g.setShowFormulas(!g.showFormulas); this.ribbon.sync(); break;
        case 'gridlines': g.setGridlines(!!arg); break;
        case 'headings': g.setHeadings(!!arg); break;
        case 'formulabar': this.formulaBarVisible = !!arg; this.fbarRow.style.display = arg ? '' : 'none'; g.positionOverlays(); break;
        case 'recalc': S.invalidate(); this.render(); XL.ui.toast('Planilha recalculada.'); break;
        case 'autofit': for (let c = g.range().c1; c <= g.range().c2; c++) g.autofit(c); break;
        case 'colwidth': this.askColWidth(); break;
        case 'undo': this.undo(); break;
        case 'redo': this.redo(); break;
        case 'shortcuts': this.showShortcuts(); break;
        case 'soon': XL.ui.toast('🔒 ' + (XL.SOON[arg] || XL.SOON.generic), { type: 'soon', duration: 4500 }); break;
        default: this.execExtra(cmd, arg); // comandos de dados, planilhas e página (excel-data.js)
      }
      this.emit('command', { cmd, arg });
      if (!NOFOCUS.includes(cmd)) g.focus();
    }

    applyFmt(patch) {
      const g = this.grid.range();
      this.pushUndo();
      XL.eachInRange(g, (r, c) => this.sheet.setFmt(r, c, patch));
      this.render();
      this.emit('format', { range: XL.rangeToStr(g), patch });
      this.emit('change', {});
    }

    applyBorder(kind) {
      const g = this.grid.range();
      const S = this.sheet;
      this.pushUndo();
      const setB = (r, c, patch) => {
        if (r < 0 || c < 0 || r >= S.rows || c >= S.cols) return;
        const b = Object.assign({}, S.fmt(r, c).border || {}, patch);
        for (const k in b) if (!b[k]) delete b[k];
        S.setFmt(r, c, { border: Object.keys(b).length ? b : null });
      };
      XL.eachInRange(g, (r, c) => {
        if (kind === 'all') setB(r, c, { t: 1, r: 1, b: 1, l: 1 });
        else if (kind === 'none') S.setFmt(r, c, { border: null });
        else if (kind === 'outer') {
          const p = {};
          if (r === g.r1) p.t = 1;
          if (r === g.r2) p.b = 1;
          if (c === g.c1) p.l = 1;
          if (c === g.c2) p.r = 1;
          setB(r, c, p);
        }
        else if (kind === 'bottom' && r === g.r2) setB(r, c, { b: 1 });
        else if (kind === 'top' && r === g.r1) setB(r, c, { t: 1 });
        else if (kind === 'left' && c === g.c1) setB(r, c, { l: 1 });
        else if (kind === 'right' && c === g.c2) setB(r, c, { r: 1 });
      });
      if (kind === 'none') {
        for (let c = g.c1; c <= g.c2; c++) { setB(g.r1 - 1, c, { b: 0 }); setB(g.r2 + 1, c, { t: 0 }); }
        for (let r = g.r1; r <= g.r2; r++) { setB(r, g.c1 - 1, { r: 0 }); setB(r, g.c2 + 1, { l: 0 }); }
      }
      this.render();
      this.emit('format', { range: XL.rangeToStr(g), border: kind });
      this.emit('change', {});
    }

    changeDecimals(delta) {
      const { ar, ac } = this.grid.sel;
      const f = this.sheet.fmt(ar, ac);
      let cur;
      if (f.dec !== undefined && f.dec !== null) cur = f.dec;
      else if (['numero', 'moeda', 'contabil'].includes(f.numFmt)) cur = 2;
      else if (f.numFmt === 'porcentagem') cur = 0;
      else {
        const v = this.sheet.value(ar, ac);
        const s = typeof v === 'number' ? XL.fmtGeneral(v) : '';
        cur = s.includes(',') && !s.includes('E') ? s.split(',')[1].length : 0;
      }
      const nd = Math.max(0, Math.min(15, cur + delta));
      const patch = { dec: nd };
      if (!f.numFmt || f.numFmt === 'geral') { patch.numFmt = 'numero'; patch.grouping = false; }
      this.applyFmt(patch);
    }

    clearRange(kind) {
      const g = this.grid.range();
      this.pushUndo();
      XL.eachInRange(g, (r, c) => {
        if (kind !== 'formats') this.sheet.clearContent(r, c);
        if (kind !== 'contents') this.sheet.replaceFmt(r, c, {});
      });
      this.sheet.invalidate();
      this.render();
      this.emit('clear', { range: XL.rangeToStr(g), kind });
      this.emit('change', {});
    }

    // ---------- Pincel de formatação ----------
    applyPainter() {
      const f = this.painterFmt;
      this.cancelPainter();
      const g = this.grid.range();
      this.pushUndo();
      XL.eachInRange(g, (r, c) => this.sheet.replaceFmt(r, c, f));
      this.render();
      this.emit('command', { cmd: 'painter-apply' });
      this.emit('change', {});
    }
    cancelPainter() { this.painterFmt = null; this.grid.el.classList.remove('painter'); }

    // ---------- AutoSoma ----------
    autosum(fn) {
      const g = this.grid.range();
      const S = this.sheet;
      // como no Excel: com linhas filtradas, a AutoSoma usa SUBTOTAL (que ignora as escondidas)
      const sub = { SOMA: 9, MÉDIA: 1, 'CONT.NÚM': 2, MÁXIMO: 4, MÍNIMO: 5 }[fn];
      if (S.filter && S.hiddenRows.size && sub) return this.autosumBuild(g, (rng) => `=SUBTOTAL(${sub};${rng})`, 'SUBTOTAL');
      return this.autosumBuild(g, (rng) => `=${fn}(${rng})`, fn);
    }
    autosumBuild(g, make, fn) {
      const S = this.sheet;
      const isNum = (r, c) => r >= 0 && c >= 0 && typeof S.value(r, c) === 'number' && S.has(r, c);
      if (g.r1 === g.r2 && g.c1 === g.c2) {
        const r = g.r1, c = g.c1;
        let rangeStr = '';
        if (isNum(r - 1, c)) {
          let i = r - 1;
          while (isNum(i - 1, c)) i--;
          rangeStr = XL.rangeToStr({ r1: i, c1: c, r2: r - 1, c2: c });
        } else if (isNum(r, c - 1)) {
          let j = c - 1;
          while (isNum(r, j - 1)) j--;
          rangeStr = XL.rangeToStr({ r1: r, c1: j, r2: r, c2: c - 1 });
        }
        const text = make(rangeStr);
        this.grid.beginEdit('enter', text);
        const end = text.length - 1;
        const start = end - rangeStr.length;
        this.grid.editor.setSelectionRange(rangeStr ? start : end, end);
        if (rangeStr) this.grid.lastPoint = { start, end, text, cell: null };
        this.grid.drawRefs();
        this.emit('autosum', { fn, range: rangeStr });
        return;
      }
      // várias células: coloca o resultado logo abaixo de cada coluna selecionada
      this.pushUndo();
      const lastEmpty = [];
      for (let c = g.c1; c <= g.c2; c++) lastEmpty.push(!S.has(g.r2, c));
      const useLast = g.r2 > g.r1 && lastEmpty.every(Boolean);
      const targetRow = useLast ? g.r2 : g.r2 + 1;
      const endRow = useLast ? g.r2 - 1 : g.r2;
      if (targetRow >= S.rows) return;
      for (let c = g.c1; c <= g.c2; c++) {
        S.setInput(targetRow, c, make(XL.rangeToStr({ r1: g.r1, c1: c, r2: endRow, c2: c })));
      }
      this.render();
      this.emit('autosum', { fn, multi: true });
      this.emit('change', {});
    }

    // ---------- Preenchimento ----------
    fill(src, dst, dir) {
      this.pushUndo();
      this.sheet.fill(src, dst, dir);
      this.render();
      this.grid.setRange(dst);
      this.emit('fill', { src: XL.rangeToStr(src), dst: XL.rangeToStr(dst), dir });
      this.emit('change', {});
    }

    // ---------- Copiar / colar ----------
    selectionTSV() {
      const g = this.grid.range();
      const lines = [];
      for (let r = g.r1; r <= g.r2; r++) {
        const row = [];
        for (let c = g.c1; c <= g.c2; c++) row.push(this.sheet.display(r, c).text);
        lines.push(row.join('\t'));
      }
      return lines.join('\n');
    }
    copy(cut) {
      const g = this.grid.range();
      const cells = [];
      for (let r = g.r1; r <= g.r2; r++) {
        const row = [];
        for (let c = g.c1; c <= g.c2; c++) {
          const x = this.sheet.cell(r, c);
          row.push(x ? { kind: x.kind, input: x.input, value: this.sheet.value(r, c), fmt: clone(x.fmt) } : null);
        }
        cells.push(row);
      }
      this.clip = { g, cells, cut };
      this.clipTSV = this.selectionTSV();
      this.grid.showCopy(g);
      try { navigator.clipboard && navigator.clipboard.writeText(this.clipTSV).catch(() => {}); } catch (e) { /* sem permissão */ }
      this.emit(cut ? 'cut' : 'copy', { range: XL.rangeToStr(g) });
    }
    clearCopy() { this.clip = null; this.grid.showCopy(null); }
    paste(mode) {
      if (!this.clip) { XL.ui.toast('Copie algo primeiro: selecione células e pressione Ctrl+C.'); return; }
      const { g, cells, cut } = this.clip;
      const S = this.sheet;
      const { ar, ac } = this.grid.sel;
      this.pushUndo();
      if (cut) XL.eachInRange(g, (r, c) => { S.clearContent(r, c); S.replaceFmt(r, c, {}); });
      cells.forEach((row, i) => row.forEach((src, j) => {
        const r = ar + i, c = ac + j;
        if (r >= S.rows || c >= S.cols) return;
        if (mode === 'formats') { S.replaceFmt(r, c, src ? src.fmt : {}); return; }
        if (mode === 'values') { S.setValue(r, c, src ? src.value : null); return; }
        if (!src || src.kind !== 'formula') S.putCell(r, c, src);
        else {
          const input = cut ? src.input : XL.Formula.shiftFormula(src.input, r - (g.r1 + i), c - (g.c1 + j));
          try { S.setInput(r, c, input); } catch (e) { S.setValue(r, c, input); }
        }
        S.replaceFmt(r, c, src ? src.fmt : {});
      }));
      S.invalidate();
      const dst = { r1: ar, c1: ac, r2: Math.min(S.rows - 1, ar + cells.length - 1), c2: Math.min(S.cols - 1, ac + cells[0].length - 1) };
      if (cut) this.clearCopy();
      this.render();
      this.grid.setRange(dst);
      this.emit('paste', { mode: mode || 'all', range: XL.rangeToStr(dst) });
      this.emit('change', {});
    }
    pasteText(text) {
      const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
      if (!rows.length) return;
      const { ar, ac } = this.grid.sel;
      this.pushUndo();
      rows.forEach((row, i) => row.forEach((v, j) => {
        const r = ar + i, c = ac + j;
        if (r >= this.sheet.rows || c >= this.sheet.cols) return;
        try { this.sheet.setInput(r, c, v); } catch (e) { this.sheet.setValue(r, c, v); }
      }));
      this.clearCopy();
      this.render();
      this.grid.setRange({ r1: ar, c1: ac, r2: Math.min(this.sheet.rows - 1, ar + rows.length - 1), c2: Math.min(this.sheet.cols - 1, ac + rows[0].length - 1) });
      this.emit('paste', { mode: 'text' });
      this.emit('change', {});
    }

    // ---------- Funções ----------
    insertFunction(name) {
      const g = this.grid;
      if (g.editing && g.activeInput().value[0] === '=') {
        g.insertText(name + '(');
        g.focusActiveInput();
      } else {
        if (g.editing) g.cancel();
        g.beginEdit('enter', '=' + name + '(');
      }
      this.emit('insertfn', { name });
    }

    openFunctionDialog() {
      const F = XL.Functions;
      let selected = null;
      const search = h('input.fd-search', { type: 'text', placeholder: 'Ex.: somar, média, procurar, data...' });
      const cat = h('select.fd-cat', h('option', { value: '_freq' }, 'Mais usadas'), h('option', { value: '_all' }, 'Todas'), F.CATEGORIES.map((c) => h('option', { value: c }, c)));
      const list = h('div.fd-list', { tabindex: '0' });
      const info = h('div.fd-info');
      const pick = (d) => {
        selected = d;
        list.querySelectorAll('.fd-item').forEach((x) => x.classList.toggle('on', x.dataset.name === d.name));
        info.innerHTML = `<div class="fd-syntax">${XL.escapeHtml(d.syntax)}</div><div>${XL.escapeHtml(d.desc)}</div>`;
      };
      const insert = (close) => { if (!selected) return false; close(); this.insertFunction(selected.name); };
      let modal;
      const fill = () => {
        const q = XL.normName(search.value.trim());
        let items;
        if (q) items = F.list.filter((d) => XL.normName(d.name).includes(q) || XL.normName(d.desc).includes(q) || (d.en && d.en.includes(q)));
        else if (cat.value === '_freq') items = FREQUENT.map((n) => F.get(n)).filter(Boolean);
        else if (cat.value === '_all') items = F.list;
        else items = F.byCategory(cat.value);
        list.innerHTML = '';
        items.forEach((d) => list.appendChild(h('div.fd-item', {
          dataset: { name: d.name },
          onclick: () => pick(d),
          ondblclick: () => insert(modal.close),
        }, d.name)));
        if (items.length) pick(items[0]); else { selected = null; info.textContent = 'Nenhuma função encontrada.'; }
      };
      search.addEventListener('input', fill);
      cat.addEventListener('change', () => { search.value = ''; fill(); });
      const body = h('div.fd',
        h('label.fd-lbl', 'Procure por uma função:'), search,
        h('label.fd-lbl', 'Ou selecione uma categoria:'), cat,
        h('label.fd-lbl', 'Selecione uma função:'), list, info);
      modal = XL.ui.modal({
        title: 'Inserir Função', body, width: 460,
        buttons: [{ label: 'OK', primary: true, onClick: (close) => insert(close) }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      fill();
      search.focus();
    }

    askColWidth() {
      const g = this.grid.range();
      const inp = h('input.fd-search', { type: 'number', min: '1', max: '500', value: String(Math.round(this.grid.colWidth(g.c1) / 7)) });
      XL.ui.modal({
        title: 'Largura da Coluna', width: 300,
        body: h('div', h('label.fd-lbl', 'Largura da coluna (em caracteres):'), inp),
        buttons: [{ label: 'OK', primary: true, onClick: () => {
          const v = Number(inp.value);
          if (!(v > 0)) return false;
          this.pushUndo();
          for (let c = g.c1; c <= g.c2; c++) this.grid.setColWidth(c, v * 7 + 5);
          this.render();
          this.emit('colwidth', {});
        } }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => inp.select(), 30);
    }

    showShortcuts() {
      const rows = [
        ['Enter / Shift+Enter', 'Confirma e desce / sobe'], ['Tab / Shift+Tab', 'Confirma e vai para a direita / esquerda'],
        ['Esc', 'Cancela a edição'], ['F2', 'Edita a célula sem apagar o conteúdo'], ['F4', 'Alterna referência absoluta ($A$1)'],
        ['Delete', 'Apaga o conteúdo da seleção'], ['Ctrl+Z / Ctrl+Y', 'Desfazer / Refazer'], ['Ctrl+C / Ctrl+X / Ctrl+V', 'Copiar / Recortar / Colar'],
        ['Ctrl+N / Ctrl+I / Ctrl+S', 'Negrito / Itálico / Sublinhado'], ['Alt+=', 'AutoSoma'], ['Ctrl+Setas', 'Pula até a borda dos dados'],
        ['Shift+Setas', 'Estende a seleção'], ['Ctrl+A', 'Seleciona tudo'], ['Ctrl+`', 'Mostrar Fórmulas'], ['Shift+F3', 'Inserir Função'],
        ['Alt+Enter', 'Quebra de linha dentro da célula'],
        ['Ctrl+L / Ctrl+U', 'Localizar / Substituir'], ['Ctrl+Shift+L', 'Liga/desliga o Filtro'],
        ['Ctrl+ + / Ctrl+ -', 'Inserir / excluir linhas ou colunas'], ['Ctrl+PageDown / PageUp', 'Próxima / anterior planilha'],
        ['Ctrl+P', 'Visualizar Impressão'], ['Ctrl+T', 'Criar Tabela'], ['Alt+↓', 'Abrir a lista suspensa da célula'], ['Ctrl+1', 'Formatar Células'], ['Ctrl+E', 'Preenchimento Relâmpago'], ['Ctrl+D', 'Preencher para baixo (copiar)'],
      ];
      XL.ui.modal({
        title: 'Atalhos de teclado', width: 520,
        body: h('table.kb-table', rows.map(([k, d]) => h('tr', h('td', h('kbd', k)), h('td', d)))),
        buttons: [{ label: 'Fechar', primary: true }],
        onClose: () => this.grid.focus(),
      });
    }

    // ---------- API usada pelas lições ----------
    reset() {
      if (this.grid.editing) this.grid.cancel();
      this.cancelPainter();
      this.selectedChart = null;
      XL.ui.closeAllModals && XL.ui.closeAllModals();
      if (XL.Find) XL.Find.close();
      if (XL.Print && XL.Print.current) XL.Print.current.close(true);
      this.wb.reset();
      this.undoStack = [];
      this.redoStack = [];
      this.clearCopy();
      this.grid.showFormulas = false;
      this.grid.setGuides([]);
      this.grid.reloadSheet();
      this.renderTabs();
      this.render();
      this.grid.setSel(0, 0, 0, 0, false);
      this.grid.scroller.scrollTop = 0;
      this.grid.scroller.scrollLeft = 0;
      this.ribbon.setTab('home', true);
    }
    // Aceita { A1: 'texto', B2: 15, C3: '=SOMA(B1:B2)', D4: { v: 0.1, fmt: { numFmt: 'porcentagem' } } }
    // sheetName opcional: escreve em outra planilha (criando-a se não existir).
    setCells(map, sheetName) {
      let S = this.sheet;
      if (sheetName) S = this.wb.byName(sheetName) || this.wb.addSheet(sheetName);
      for (const ref in map) {
        const p = XL.parseRef(ref);
        let v = map[ref];
        let fmt = null;
        if (v && typeof v === 'object') { fmt = v.fmt; v = v.v; }
        if (v !== undefined && v !== null) {
          if (typeof v === 'number') S.setValue(p.r, p.c, v);
          else S.setInput(p.r, p.c, String(v));
        }
        if (fmt) S.setFmt(p.r, p.c, fmt);
      }
      if (sheetName) this.renderTabs();
      this.render();
    }
    setCell(ref, input) {
      const p = XL.parseRef(ref);
      this.pushUndo();
      this.sheet.setInput(p.r, p.c, String(input));
      this.render();
      this.emit('commit', { r: p.r, c: p.c, ref: XL.toRef(p.r, p.c), input: String(input) });
      this.emit('change', {});
    }
    setFormat(rangeStr, patch) {
      XL.eachInRange(XL.parseRange(rangeStr), (r, c) => this.sheet.setFmt(r, c, patch));
      this.render();
    }
    setColWidths(map, sheetName) {
      const S = (sheetName && this.wb.byName(sheetName)) || this.sheet;
      for (const col in map) S.colW[XL.colToNum(col) - 1] = map[col];
      this.grid.applyColWidths();
      this.render();
    }
    select(rangeStr) {
      const g = XL.parseRange(rangeStr);
      if (this.grid.editing) this.grid.cancel();
      this.grid.setRange(g);
      this.grid.focus();
    }
    gotoRef(text) {
      const g = XL.parseRange(String(text).replace(/\s/g, ''));
      if (!g || g.r2 >= this.sheet.rows || g.c2 >= this.sheet.cols) {
        XL.ui.toast('Referência inválida. Digite algo como <b>H10</b> ou <b>B2:D6</b>.');
        this.updateNameBox();
        return false;
      }
      if (this.grid.editing) this.grid.cancel();
      this.grid.setRange(g);
      this.emit('namebox', { ref: XL.rangeToStr(g) });
      this.grid.focus();
      return true;
    }
  }

  XL.ExcelApp = ExcelApp;
})();
