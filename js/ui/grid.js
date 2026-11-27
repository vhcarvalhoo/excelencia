/* Grade de células: desenho, seleção, edição na célula, modo "Aponte", alça de preenchimento. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const K = XL.SHEET_KEY;
  const REF_COLORS = ['#2f6fdb', '#d6403a', '#8a4ad0', '#18935a', '#d9821a', '#139aa6', '#c23a8c', '#6d7d16'];
  const ROWH_W = 44;
  const DEFAULT_W = 80;

  class Grid {
    constructor(app) {
      this.app = app;
      this.R = this.sheet.rows;
      this.C = this.sheet.cols;
      this.sel = { ar: 0, ac: 0, fr: 0, fc: 0 };
      this.editing = null;
      this.showFormulas = false;
      this.gridlines = true;
      this.headings = true;
      this.rendered = new Set();
      this.hlCells = [];
      this.hlHeads = [];
      this.drag = null;
      this.lastPoint = null;
      this.copyRange = null;
      this.guides = [];
      this.ac = null;
      this.build();
    }

    // a grade sempre mostra a planilha ativa da pasta de trabalho
    get sheet() { return this.app.sheet; }

    // ---------- Construção ----------
    build() {
      this.el = h('div.xl-grid-host');
      this.scroller = h('div.xl-grid-scroll');
      this.wrap = h('div.xl-grid-wrap');
      this.table = h('table.xl-grid');
      const cg = h('colgroup');
      cg.appendChild(h('col', { style: { width: ROWH_W + 'px' } }));
      this.colEls = [];
      for (let c = 0; c < this.C; c++) {
        const col = h('col', { style: { width: this.colWidth(c) + 'px' } });
        this.colEls.push(col);
        cg.appendChild(col);
      }
      this.table.appendChild(cg);
      const thead = h('thead');
      const hr = h('tr');
      this.corner = h('th.xl-corner', { title: 'Selecionar tudo' }, h('span.xl-corner-tri'));
      hr.appendChild(this.corner);
      this.colHeads = [];
      for (let c = 0; c < this.C; c++) {
        const th = h('th.xl-colh', { dataset: { c } }, XL.numToCol(c + 1), h('span.xl-resizer', { dataset: { c }, title: 'Arraste para mudar a largura · duplo clique para autoajustar' }));
        this.colHeads.push(th);
        hr.appendChild(th);
      }
      thead.appendChild(hr);
      this.table.appendChild(thead);
      const tbody = h('tbody');
      this.tds = [];
      this.rowHeads = [];
      this.trs = [];
      for (let r = 0; r < this.R; r++) {
        const tr = h('tr');
        this.trs.push(tr);
        const th = h('th.xl-rowh', { dataset: { r } }, String(r + 1));
        this.rowHeads.push(th);
        tr.appendChild(th);
        const row = [];
        for (let c = 0; c < this.C; c++) {
          const td = h('td.xl-c', { dataset: { r, c } }, h('div.xl-cv'));
          row.push(td);
          tr.appendChild(td);
        }
        this.tds.push(row);
        tbody.appendChild(tr);
      }
      this.table.appendChild(tbody);
      this.updateTableWidth();

      this.refLayer = h('div.xl-layer');
      this.selBox = h('div.xl-selbox');
      this.handle = h('div.xl-handle', { title: 'Alça de preenchimento: arraste para copiar/continuar · duplo clique para preencher até o fim' });
      this.copyBox = h('div.xl-copybox', { style: { display: 'none' } });
      this.fillBox = h('div.xl-fillbox', { style: { display: 'none' } });
      this.guideLayer = h('div.xl-layer.guides');
      this.breakLayer = h('div.xl-layer.breaks');
      this.filterLayer = h('div.xl-filterlayer');
      this.chartLayer = h('div.xl-chartlayer');
      this.dvLayer = h('div.xl-filterlayer.dv');
      this.editor = h('textarea.xl-editor.idle', { spellcheck: 'false', autocomplete: 'off', rows: '1', wrap: 'off', 'aria-label': 'Editor de célula' });
      this.wrap.append(this.table, this.breakLayer, this.refLayer, this.selBox, this.copyBox, this.fillBox, this.handle, this.filterLayer, this.dvLayer, this.chartLayer, this.guideLayer, this.editor);
      this.scroller.appendChild(this.wrap);
      this.el.appendChild(this.scroller);

      this.acEl = h('div.xl-ac', { style: { display: 'none' } });
      this.tipEl = h('div.xl-tip', { style: { display: 'none' } });
      document.body.append(this.acEl, this.tipEl);

      // eventos
      this.table.addEventListener('mousedown', (e) => this.onMouseDown(e));
      this.table.addEventListener('mouseover', (e) => this.onMouseOver(e));
      this.table.addEventListener('dblclick', (e) => this.onDblClick(e));
      this.handle.addEventListener('mousedown', (e) => this.onHandleDown(e));
      this.handle.addEventListener('dblclick', (e) => this.onHandleDbl(e));
      this.editor.addEventListener('mousedown', (e) => { if (!this.editing) e.preventDefault(); });
      this.editor.addEventListener('keydown', (e) => this.onKey(e));
      this.editor.addEventListener('input', () => this.onEditorInput());
      this.editor.addEventListener('copy', (e) => this.onCopyEvent(e, false));
      this.editor.addEventListener('cut', (e) => this.onCopyEvent(e, true));
      this.editor.addEventListener('paste', (e) => this.onPasteEvent(e));
      this.editor.addEventListener('click', () => this.onEditInput(true));
      this.editor.addEventListener('keyup', (e) => { if (/Arrow|Home|End/.test(e.key) && this.editing) this.onEditInput(true); });
      this.table.addEventListener('contextmenu', (e) => this.onContextMenu(e));
      this.scroller.addEventListener('scroll', () => {
        this.hidePopups(true);
        if (this.sheet.freeze) this.positionOverlays();
        this.app.emit('scroll', { top: this.scroller.scrollTop, left: this.scroller.scrollLeft });
      });
      this._mu = (e) => this.onMouseUp(e);
      this._mm = (e) => this.onMouseMove(e);
      document.addEventListener('mouseup', this._mu);
      document.addEventListener('mousemove', this._mm);
    }

    destroy() {
      document.removeEventListener('mouseup', this._mu);
      document.removeEventListener('mousemove', this._mm);
      this.acEl.remove();
      this.tipEl.remove();
      if (this.errEl) this.errEl.remove();
    }

    colWidth(c) { return this.sheet.colW[c] || DEFAULT_W; }
    setColWidth(c, w) {
      w = Math.max(6, Math.round(w));
      this.sheet.colW[c] = w;
      this.colEls[c].style.width = w + 'px';
      this.updateTableWidth();
      this.positionOverlays();
    }
    applyColWidths() {
      for (let c = 0; c < this.C; c++) this.colEls[c].style.width = this.colWidth(c) + 'px';
      this.updateTableWidth();
    }
    updateTableWidth() {
      let w = this.headings ? ROWH_W : 0;
      for (let c = 0; c < this.C; c++) w += this.colWidth(c);
      this.table.style.width = w + 'px';
    }

    focus() {
      if (this.editing && this.editing.source === 'bar') return;
      this.editor.focus({ preventScroll: true });
    }

    // ---------- Seleção ----------
    clampR(r) { return Math.max(0, Math.min(this.R - 1, r)); }
    clampC(c) { return Math.max(0, Math.min(this.C - 1, c)); }
    range() {
      const s = this.sel;
      return { r1: Math.min(s.ar, s.fr), c1: Math.min(s.ac, s.fc), r2: Math.max(s.ar, s.fr), c2: Math.max(s.ac, s.fc) };
    }
    setSel(r, c, fr, fc, final = true) {
      r = this.clampR(r); c = this.clampC(c);
      fr = fr === undefined ? r : this.clampR(fr);
      fc = fc === undefined ? c : this.clampC(fc);
      this.sel = { ar: r, ac: c, fr, fc };
      this.renderSel();
      this.scrollIntoView(fr, fc);
      this.app.onSelectionChange(final);
    }
    setRange(g, final = true) { this.setSel(g.r1, g.c1, g.r2, g.c2, final); this.scrollIntoView(g.r1, g.c1); }
    setFocusCell(r, c, final = true) {
      this.sel.fr = this.clampR(r);
      this.sel.fc = this.clampC(c);
      this.renderSel();
      this.scrollIntoView(this.sel.fr, this.sel.fc);
      this.app.onSelectionChange(final);
    }

    renderSel() {
      const g = this.range();
      for (const td of this.hlCells) td.classList.remove('in-range');
      for (const th of this.hlHeads) th.classList.remove('hl', 'full');
      this.hlCells = [];
      this.hlHeads = [];
      const { ar, ac } = this.sel;
      if (g.r1 !== g.r2 || g.c1 !== g.c2) {
        XL.eachInRange(g, (r, c) => {
          if (r === ar && c === ac) return;
          const td = this.tds[r][c];
          td.classList.add('in-range');
          this.hlCells.push(td);
        });
      }
      const fullCol = g.r1 === 0 && g.r2 === this.R - 1;
      const fullRow = g.c1 === 0 && g.c2 === this.C - 1;
      for (let c = g.c1; c <= g.c2; c++) { const th = this.colHeads[c]; th.classList.add('hl'); if (fullCol) th.classList.add('full'); this.hlHeads.push(th); }
      for (let r = g.r1; r <= g.r2; r++) { const th = this.rowHeads[r]; th.classList.add('hl'); if (fullRow) th.classList.add('full'); this.hlHeads.push(th); }
      this.positionOverlays();
    }

    isHidden(r) { return this.sheet.hiddenRows.has(r); }
    // Próxima linha visível andando dr linhas (pula as escondidas pelo filtro).
    stepRow(r, dr) {
      if (!dr) return r;
      const dir = Math.sign(dr);
      let n = this.clampR(r + dr);
      while (n >= 0 && n < this.R && this.isHidden(n)) n += dir;
      return n < 0 || n >= this.R ? r : n;
    }
    moveBy(dr, dc, extend) {
      if (extend) this.setFocusCell(this.stepRow(this.sel.fr, dr), this.sel.fc + dc);
      else this.setSel(this.stepRow(this.sel.ar, dr), this.sel.ac + dc);
    }

    rect(g) {
      let r1 = this.clampR(g.r1), r2 = this.clampR(g.r2);
      const c1 = this.clampC(g.c1), c2 = this.clampC(g.c2);
      while (r1 < r2 && this.isHidden(r1)) r1++;
      while (r2 > r1 && this.isHidden(r2)) r2--;
      if (this.isHidden(r1)) {
        let n = r1;
        while (n < this.R - 1 && this.isHidden(n)) n++;
        const a = this.tds[n][c1], b = this.tds[n][c2];
        return { x: a.offsetLeft, y: a.offsetTop, w: b.offsetLeft + b.offsetWidth - a.offsetLeft, h: 0 };
      }
      const a = this.tds[r1][c1], b = this.tds[r2][c2];
      return { x: a.offsetLeft, y: a.offsetTop, w: b.offsetLeft + b.offsetWidth - a.offsetLeft, h: b.offsetTop + b.offsetHeight - a.offsetTop };
    }
    place(el, g, pad = 1) {
      const r = this.rect(g);
      Object.assign(el.style, { left: r.x - pad + 'px', top: r.y - pad + 'px', width: r.w + pad + 'px', height: r.h + pad + 'px' });
    }

    positionOverlays() {
      const g = this.range();
      this.place(this.selBox, g);
      const r = this.rect(g);
      Object.assign(this.handle.style, { left: r.x + r.w - 4 + 'px', top: r.y + r.h - 4 + 'px' });
      this.handle.style.display = this.editing ? 'none' : '';
      this.positionEditor();
      if (this.copyRange) this.place(this.copyBox, this.copyRange);
      this.drawRefs();
      this.drawGuides();
      this.drawFilterButtons();
      this.drawBreaks();
      this.drawDv();
    }

    // Seta da lista suspensa e mensagem de entrada da validação de dados
    drawDv() {
      this.dvLayer.innerHTML = '';
      if (this.editing) return;
      const g = this.range();
      if (g.r1 !== g.r2 || g.c1 !== g.c2) return;
      const rule = this.sheet.validationAt(g.r1, g.c1);
      if (!rule) return;
      const td = this.tds[g.r1][g.c1];
      if (rule.type === 'list') {
        const b = h('button.xl-dvbtn', { type: 'button', title: 'Escolher da lista (Alt+↓)', html: '<svg viewBox="0 0 12 12" width="10" height="10"><path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>' });
        b.style.left = td.offsetLeft + td.offsetWidth + 1 + 'px';
        b.style.top = td.offsetTop + td.offsetHeight - 20 + 'px';
        b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
        b.addEventListener('click', (e) => { e.stopPropagation(); this.app.openDvList(g.r1, g.c1, b); });
        this.dvLayer.appendChild(b);
      }
      if (rule.inputMsg || rule.inputTitle) {
        const tip = h('div.xl-dvtip', rule.inputTitle ? h('b', rule.inputTitle) : null, rule.inputMsg ? h('div', rule.inputMsg) : null);
        tip.style.left = td.offsetLeft + 8 + 'px';
        tip.style.top = td.offsetTop + td.offsetHeight + 4 + 'px';
        this.dvLayer.appendChild(tip);
      }
    }

    // Editando uma fórmula enquanto olha outra planilha (para apontar células dela)?
    isAway() { return !!this.editing && this.editing.sheet !== this.sheet; }

    positionEditor() {
      const td = this.tds[this.editing ? this.editing.r : this.sel.ar][this.editing ? this.editing.c : this.sel.ac];
      const ed = this.editor;
      ed.classList.toggle('away', this.isAway());
      ed.style.left = td.offsetLeft + 'px';
      ed.style.top = td.offsetTop + 'px';
      ed.style.minWidth = td.offsetWidth - 1 + 'px';
      ed.style.height = td.offsetHeight - 1 + 'px';
      if (this.editing) {
        const cell = this.sheet.cell(this.editing.r, this.editing.c);
        const f = (cell && cell.fmt) || {};
        ed.style.fontWeight = f.b ? 'bold' : '';
        ed.style.fontStyle = f.i ? 'italic' : '';
        ed.style.fontFamily = f.font ? `"${f.font}", sans-serif` : '';
        ed.style.fontSize = f.size ? f.size + 'pt' : '';
        ed.style.width = '0px';
        ed.style.width = Math.max(td.offsetWidth - 1, ed.scrollWidth + 8) + 'px';
      }
    }

    scrollIntoView(r, c) {
      if (this.isHidden(r)) return;
      const td = this.tds[r][c];
      const sc = this.scroller;
      const f = this.sheet.freeze || { rows: 0, cols: 0 };
      let headH = this.headings ? this.colHeads[0].offsetHeight : 0;
      let rowW = this.headings ? ROWH_W : 0;
      for (let i = 0; i < f.rows; i++) if (!this.isHidden(i)) headH += this.trs[i].offsetHeight;
      for (let i = 0; i < f.cols; i++) rowW += this.colWidth(i);
      if (r >= f.rows) {
        const top = this.staticTop(r);
        if (top - headH < sc.scrollTop) sc.scrollTop = top - headH;
        else if (top + td.offsetHeight > sc.scrollTop + sc.clientHeight) sc.scrollTop = top + td.offsetHeight - sc.clientHeight;
      }
      if (c >= f.cols) {
        const left = this.staticLeft(c);
        if (left - rowW < sc.scrollLeft) sc.scrollLeft = left - rowW;
        else if (left + td.offsetWidth > sc.scrollLeft + sc.clientWidth) sc.scrollLeft = left + td.offsetWidth - sc.clientWidth;
      }
    }
    // posição "real" (sem o efeito do congelamento) de uma linha/coluna
    staticTop(r) { return this.rowHeads[r].offsetTop; }
    staticLeft(c) { let x = this.headings ? ROWH_W : 0; for (let i = 0; i < c; i++) x += this.colWidth(i); return x; }

    // ---------- Linhas escondidas, congelamento, botões de filtro, quebras de página ----------
    applyHidden() {
      const hid = this.sheet.hiddenRows;
      const f = this.sheet.filter;
      const filtering = !!f && hid.size > 0;
      for (let r = 0; r < this.R; r++) {
        const hide = hid.has(r);
        const tr = this.trs[r];
        if (tr._hidden !== hide) { tr.style.display = hide ? 'none' : ''; tr._hidden = hide; }
        this.rowHeads[r].classList.toggle('filtered', filtering && r > f.g.r1 && r <= f.g.r2);
      }
    }

    applyFreeze() {
      for (const el of this.frozenEls || []) {
        el.style.position = ''; el.style.top = ''; el.style.left = ''; el.style.zIndex = '';
        el.classList.remove('frz', 'frz-b', 'frz-r');
      }
      this.frozenEls = [];
      const f = this.sheet.freeze;
      if (!f || (!f.rows && !f.cols)) return;
      const headH = this.headings ? this.colHeads[0].offsetHeight : 0;
      const rowW = this.headings ? ROWH_W : 0;
      const tops = [], lefts = [];
      let top = headH, left = rowW;
      for (let r = 0; r < f.rows; r++) { tops[r] = top; if (!this.isHidden(r)) top += this.trs[r].offsetHeight; }
      for (let c = 0; c < f.cols; c++) { lefts[c] = left; left += this.colWidth(c); }
      const mark = (el, t, l, z) => {
        el.style.position = 'sticky';
        if (t !== null) el.style.top = t + 'px';
        if (l !== null) el.style.left = l + 'px';
        el.style.zIndex = z;
        this.frozenEls.push(el);
      };
      for (let r = 0; r < this.R; r++) {
        const inRow = r < f.rows;
        if (inRow) mark(this.rowHeads[r], tops[r], null, 6);
        for (let c = 0; c < (inRow ? this.C : f.cols); c++) {
          const td = this.tds[r][c];
          const inCol = c < f.cols;
          mark(td, inRow ? tops[r] : null, inCol ? lefts[c] : null, inRow && inCol ? 3 : 2);
          td.classList.add('frz');
          if (inRow && r === f.rows - 1) td.classList.add('frz-b');
          if (inCol && c === f.cols - 1) td.classList.add('frz-r');
        }
      }
      for (let c = 0; c < f.cols; c++) mark(this.colHeads[c], null, lefts[c], 6);
    }

    drawFilterButtons() {
      this.filterLayer.innerHTML = '';
      const f = this.sheet.filter;
      if (!f) return;
      for (let c = f.g.c1; c <= f.g.c2; c++) {
        const td = this.tds[f.g.r1][c];
        const on = !!f.crit[c];
        const b = h('button.xl-filterbtn' + (on ? '.on' : ''), {
          type: 'button', title: on ? 'Filtro aplicado nesta coluna' : 'Filtrar / classificar esta coluna', dataset: { c },
          html: on
            ? '<svg viewBox="0 0 12 12" width="11" height="11"><path d="M1.5 2h9L7 6.2V10l-2 1V6.2z" fill="currentColor"/></svg>'
            : '<svg viewBox="0 0 12 12" width="10" height="10"><path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>',
        });
        b.style.left = td.offsetLeft + td.offsetWidth - 19 + 'px';
        b.style.top = td.offsetTop + td.offsetHeight - 19 + 'px';
        b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
        b.addEventListener('click', (e) => { e.stopPropagation(); if (this.editing && !this.commit()) return; this.app.openFilterMenu(c, b); });
        this.filterLayer.appendChild(b);
      }
    }

    drawBreaks() {
      this.breakLayer.innerHTML = '';
      if (!this.sheet.showBreaks || !XL.Print) return;
      const p = XL.Print.paginate(this.sheet);
      if (!p.area) return;
      const box = this.rect(p.area);
      for (const c of p.colBreaks) {
        const x = this.tds[p.area.r1][c].offsetLeft;
        this.breakLayer.appendChild(h('div.xl-break.v', { style: { left: x - 1 + 'px', top: box.y + 'px', height: box.h + 'px' } }));
      }
      for (const r of p.rowBreaks) {
        const y = this.rowHeads[r].offsetTop;
        this.breakLayer.appendChild(h('div.xl-break.hz', { style: { top: y - 1 + 'px', left: box.x + 'px', width: box.w + 'px' } }));
      }
    }

    // Mostra a planilha ativa (após trocar de planilha ou desfazer).
    reloadSheet() {
      this.applyColWidths();
      this.showCopy(null);
      const s = this.sheet.sel;
      this.sel = s ? Object.assign({}, s) : { ar: 0, ac: 0, fr: 0, fc: 0 };
      this.render();
      this.renderSel();
      this.scrollIntoView(this.sel.ar, this.sel.ac);
    }

    // ---------- Menu do botão direito ----------
    onContextMenu(e) {
      e.preventDefault();
      if (this.editing && !this.commit()) return;
      const th = e.target.closest('th');
      const cell = this.cellAt(e.target);
      const g = this.range();
      let kind = 'cell';
      if (th && th.classList.contains('xl-rowh')) {
        const r = +th.dataset.r;
        if (!(g.c1 === 0 && g.c2 === this.C - 1 && r >= g.r1 && r <= g.r2)) this.setSel(r, 0, r, this.C - 1);
        kind = 'row';
      } else if (th && th.classList.contains('xl-colh')) {
        const c = +th.dataset.c;
        if (!(g.r1 === 0 && g.r2 === this.R - 1 && c >= g.c1 && c <= g.c2)) this.setSel(0, c, this.R - 1, c);
        kind = 'col';
      } else if (cell) {
        if (!(cell.r >= g.r1 && cell.r <= g.r2 && cell.c >= g.c1 && cell.c <= g.c2)) this.setSel(cell.r, cell.c);
      } else return;
      this.app.showContextMenu(kind, { x: e.clientX, y: e.clientY });
    }

    // ---------- Desenho das células ----------
    render() {
      const sh = this.sheet;
      this.applyHidden();
      const keys = new Set();
      // estilos de tabela e formatação condicional entram como "camadas" sobre o formato da célula
      this.tblMap = sh.tableStyles();
      this.cfMap = sh.conditionalStyles();
      for (const k of this.tblMap.keys()) keys.add(k);
      for (const k of this.cfMap.keys()) keys.add(k);
      const bottom = new Set(), right = new Set(), topS = new Set(), leftS = new Set();
      for (const [k, cell] of sh.cells) {
        keys.add(k);
        const b = cell.fmt.border;
        if (!b) continue;
        const r = Math.floor(k / K), c = k % K;
        if (b.b) bottom.add(k);
        if (b.r) right.add(k);
        if (b.t) { if (r > 0) { bottom.add(k - K); keys.add(k - K); } else topS.add(k); }
        if (b.l) { if (c > 0) { right.add(k - 1); keys.add(k - 1); } else leftS.add(k); }
      }
      // células vizinhas (à esquerda) podem mudar o corte do texto
      for (const k of [...keys]) if (k % K > 0) keys.add(k - 1);
      for (const k of this.rendered) if (!keys.has(k)) this.paintTd(k, false, false, false, false);
      const numeric = [];
      for (const k of keys) if (this.paintTd(k, bottom.has(k), right.has(k), topS.has(k), leftS.has(k))) numeric.push(k);
      this.rendered = keys;
      // números que não cabem na coluna viram ####, como no Excel
      const tooWide = numeric.filter((k) => {
        const td = this.tds[Math.floor(k / K)][k % K];
        return td.firstChild.scrollWidth > td.clientWidth;
      });
      for (const k of tooWide) {
        const td = this.tds[Math.floor(k / K)][k % K];
        td.firstChild.textContent = '#'.repeat(Math.max(1, Math.floor((td.clientWidth - 6) / 8)));
      }
      this.applyFreeze();
      this.positionOverlays();
      if (XL.Charts) XL.Charts.draw(this);
    }

    // Retorna true se a célula mostra um número (para checar ####).
    paintTd(k, bB, bR, bT, bL) {
      const r = Math.floor(k / K), c = k % K;
      if (r >= this.R || c >= this.C) return false;
      const td = this.tds[r][c];
      const cv = td.firstChild;
      const cell = this.sheet.cell(r, c);
      const f = (cell && cell.fmt) || {};
      const d = cell ? this.sheet.display(r, c, this.showFormulas) : { text: '', align: null };
      cv.textContent = d.text;
      let align = f.align || d.align || 'left';
      if (this.showFormulas && cell && cell.kind === 'formula') align = 'left';
      cv.style.textAlign = align;
      const ts = (this.tblMap && this.tblMap.get(k)) || {};
      const cs = (this.cfMap && this.cfMap.get(k)) || {};
      td.style.backgroundColor = cs.fill || f.fill || ts.fill || '';
      td.style.backgroundImage = cs.bar ? `linear-gradient(90deg, ${cs.bar.color} 0%, ${cs.bar.color}99 ${cs.bar.pct}%, transparent ${cs.bar.pct}%)` : '';
      cv.style.fontWeight = cs.b || f.b || (ts.b && f.b === undefined) ? 'bold' : '';
      cv.style.fontStyle = f.i ? 'italic' : '';
      cv.style.textDecoration = f.u ? 'underline' : '';
      cv.style.color = cs.color || f.color || ts.color || '';
      cv.style.fontFamily = f.font ? `"${f.font}", sans-serif` : '';
      cv.style.fontSize = f.size ? f.size + 'pt' : '';
      td.style.verticalAlign = { top: 'top', middle: 'middle' }[f.valign] || '';
      cv.classList.toggle('wrap', !!f.wrap);
      const rightEmpty = c + 1 >= this.C || !this.sheet.has(r, c + 1);
      const spill = align === 'left' && rightEmpty && !d.num && !f.wrap;
      cv.classList.toggle('clip', !spill);
      td.style.borderBottomColor = bB ? '#000' : '';
      td.style.borderRightColor = bR ? '#000' : '';
      td.style.boxShadow = [bT ? 'inset 0 1px 0 #000' : '', bL ? 'inset 1px 0 0 #000' : ''].filter(Boolean).join(',');
      return !!d.num && !this.showFormulas;
    }

    setShowFormulas(v) { this.showFormulas = v; this.render(); }
    setGridlines(v) { this.gridlines = v; this.el.classList.toggle('no-gridlines', !v); }
    setHeadings(v) { this.headings = v; this.el.classList.toggle('no-headings', !v); this.updateTableWidth(); this.positionOverlays(); }

    // ---------- Realces: referências (fórmula em edição), guias (lições), cópia ----------
    drawRefs() {
      this.refLayer.innerHTML = '';
      if (!this.editing) return;
      const text = this.activeInput().value;
      if (text[0] !== '=') return;
      const here = XL.normName(this.sheet.name);
      const away = this.isAway();
      XL.Formula.refsIn(text).forEach((ref, i) => {
        const g = ref.g;
        if (g.r1 >= this.R || g.c1 >= this.C) return;
        if (ref.sheet ? XL.normName(ref.sheet) !== here : away) return; // referência de outra planilha
        const color = REF_COLORS[i % REF_COLORS.length];
        const box = h('div.xl-refbox', { style: { borderColor: color, background: color + '14' } });
        this.place(box, { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, this.R - 1), c2: Math.min(g.c2, this.C - 1) }, 0);
        this.refLayer.appendChild(box);
      });
    }

    setGuides(list) { this.guides = list || []; this.drawGuides(); }
    drawGuides() {
      this.guideLayer.innerHTML = '';
      for (const g of this.guides) {
        const box = h('div.xl-guidebox');
        this.place(box, g, 2);
        this.guideLayer.appendChild(box);
      }
    }
    scrollToRange(g) { this.scrollIntoView(this.clampR(g.r2), this.clampC(g.c2)); this.scrollIntoView(g.r1, g.c1); }

    showCopy(g) {
      this.copyRange = g;
      this.copyBox.style.display = g ? '' : 'none';
      if (g) this.place(this.copyBox, g);
    }

    // ---------- Mouse ----------
    cellAt(target) {
      const td = target.closest && target.closest('td.xl-c');
      return td ? { r: +td.dataset.r, c: +td.dataset.c } : null;
    }

    onMouseDown(e) {
      if (e.button !== 0) return;
      if (this.app.selectedChart) this.app.selectChart(null);
      if (e.target.classList.contains('xl-resizer')) { this.startResize(e, +e.target.dataset.c); return; }
      const th = e.target.closest('th');
      if (th) {
        e.preventDefault();
        if (this.editing && !this.commit()) return;
        if (th.classList.contains('xl-colh')) {
          const c = +th.dataset.c;
          if (e.shiftKey) this.setSel(0, this.sel.ac, this.R - 1, c, false);
          else this.setSel(0, c, this.R - 1, c, false);
          this.drag = { type: 'col', anchor: e.shiftKey ? this.sel.ac : c };
        } else if (th.classList.contains('xl-rowh')) {
          const r = +th.dataset.r;
          if (e.shiftKey) this.setSel(this.sel.ar, 0, r, this.C - 1, false);
          else this.setSel(r, 0, r, this.C - 1, false);
          this.drag = { type: 'row', anchor: e.shiftKey ? this.sel.ar : r };
        } else {
          this.setSel(0, 0, this.R - 1, this.C - 1, true);
        }
        this.focus();
        return;
      }
      const cell = this.cellAt(e.target);
      if (!cell) return;
      e.preventDefault();
      if (this.editing) {
        if (this.isPointable()) { this.startPoint(cell); return; }
        if (!this.commit()) return;
      }
      this.tabCol = null;
      if (e.shiftKey) this.setFocusCell(cell.r, cell.c, false);
      else this.setSel(cell.r, cell.c, cell.r, cell.c, false);
      this.drag = { type: 'sel' };
      this.focus();
    }

    onMouseOver(e) {
      if (!this.drag) return;
      const cell = this.cellAt(e.target);
      const th = e.target.closest && e.target.closest('th');
      const d = this.drag;
      if (d.type === 'col') {
        const c = cell ? cell.c : th && th.dataset.c !== undefined ? +th.dataset.c : null;
        if (c !== null) { this.sel = { ar: 0, ac: d.anchor, fr: this.R - 1, fc: c }; this.renderSel(); this.app.onSelectionChange(false); }
        return;
      }
      if (d.type === 'row') {
        const r = cell ? cell.r : th && th.dataset.r !== undefined ? +th.dataset.r : null;
        if (r !== null) { this.sel = { ar: d.anchor, ac: 0, fr: r, fc: this.C - 1 }; this.renderSel(); this.app.onSelectionChange(false); }
        return;
      }
      if (!cell) return;
      if (d.type === 'sel') this.setFocusCell(cell.r, cell.c, false);
      else if (d.type === 'point') this.updatePoint(cell);
      else if (d.type === 'fill') this.updateFill(cell);
    }

    onMouseMove(e) {
      if (this.drag && this.drag.type === 'resize') {
        const d = this.drag;
        this.setColWidth(d.c, d.w0 + (e.clientX - d.x0));
      }
    }

    onMouseUp() {
      const d = this.drag;
      if (!d) return;
      this.drag = null;
      if (d.type === 'fill') this.finishFill(d);
      else if (d.type === 'resize') { this.app.render(); this.app.emit('colwidth', { c: d.c, w: this.colWidth(d.c) }); }
      else if (d.type === 'point') this.focusActiveInput();
      else {
        this.app.onSelectionChange(true);
        if (this.app.painterFmt) this.app.applyPainter();
      }
    }

    onDblClick(e) {
      if (e.target.classList.contains('xl-resizer')) { this.autofit(+e.target.dataset.c); return; }
      const cell = this.cellAt(e.target);
      if (!cell || this.editing) return;
      this.setSel(cell.r, cell.c);
      this.beginEdit('edit');
    }

    startResize(e, c) {
      e.preventDefault();
      e.stopPropagation();
      if (this.editing && !this.commit()) return;
      this.app.pushUndo();
      this.drag = { type: 'resize', c, x0: e.clientX, w0: this.colWidth(c) };
    }

    autofit(c) {
      const meas = h('span.xl-measure');
      this.wrap.appendChild(meas);
      let max = 0;
      for (let r = 0; r < this.R; r++) {
        if (!this.sheet.has(r, c)) continue;
        const cv = this.tds[r][c].firstChild;
        const d = this.sheet.display(r, c, this.showFormulas);
        meas.style.font = getComputedStyle(cv).font;
        meas.textContent = d.text;
        max = Math.max(max, meas.offsetWidth);
      }
      meas.remove();
      if (max === 0) return;
      this.app.pushUndo();
      this.setColWidth(c, max + 12);
      this.app.render();
      this.app.emit('colwidth', { c, w: this.colWidth(c) });
    }

    // ---------- Alça de preenchimento ----------
    onHandleDown(e) {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      if (this.editing && !this.commit()) return;
      this.drag = { type: 'fill', src: this.range(), dst: null, dir: null };
      this.app.setMode('Pronto');
    }
    updateFill(cell) {
      const d = this.drag, s = d.src;
      const dDown = cell.r - s.r2, dUp = s.r1 - cell.r, dRight = cell.c - s.c2, dLeft = s.c1 - cell.c;
      const vert = Math.max(dDown, dUp), hor = Math.max(dRight, dLeft);
      d.dst = null; d.dir = null;
      if (vert > 0 || hor > 0) {
        if (vert >= hor) {
          if (dDown > 0) { d.dir = 'down'; d.dst = { r1: s.r1, c1: s.c1, r2: cell.r, c2: s.c2 }; }
          else { d.dir = 'up'; d.dst = { r1: cell.r, c1: s.c1, r2: s.r2, c2: s.c2 }; }
        } else if (dRight > 0) { d.dir = 'right'; d.dst = { r1: s.r1, c1: s.c1, r2: s.r2, c2: cell.c }; }
        else { d.dir = 'left'; d.dst = { r1: s.r1, c1: cell.c, r2: s.r2, c2: s.c2 }; }
      }
      if (d.dst) { this.fillBox.style.display = ''; this.place(this.fillBox, d.dst); }
      else this.fillBox.style.display = 'none';
    }
    finishFill(d) {
      this.fillBox.style.display = 'none';
      if (d.dst) this.app.fill(d.src, d.dst, d.dir);
    }
    onHandleDbl(e) {
      e.preventDefault();
      const g = this.range();
      const adj = (r) => (g.c1 > 0 && this.sheet.has(r, g.c1 - 1)) || (g.c2 + 1 < this.C && this.sheet.has(r, g.c2 + 1));
      let end = g.r2;
      while (end + 1 < this.R && adj(end + 1)) end++;
      if (end > g.r2) this.app.fill(g, { r1: g.r1, c1: g.c1, r2: end, c2: g.c2 }, 'down');
    }

    // ---------- Edição ----------
    activeInput() { return this.editing && this.editing.source === 'bar' ? this.app.fbar : this.editor; }
    focusActiveInput() { this.activeInput().focus({ preventScroll: true }); }

    beginEdit(mode, initial, source = 'cell') {
      if (this.editing) return;
      const { ar: r, ac: c } = this.sel;
      const cell = this.sheet.cell(r, c);
      const text = initial !== undefined ? initial : XL.editTextFor(cell);
      this.editing = { r, c, mode, source, sheet: this.sheet };
      this.editor.classList.remove('idle');
      if (this.editor.value !== text) this.editor.value = text;
      this.app.fbar.value = text;
      this.lastPoint = null;
      this.handle.style.display = 'none';
      this.positionEditor();
      if (source === 'cell') {
        this.editor.focus({ preventScroll: true });
        const n = this.editor.value.length;
        this.editor.setSelectionRange(n, n);
      }
      this.app.setMode(mode === 'enter' ? 'Digite' : 'Editar');
      this.app.emit('editstart', { r, c });
      this.onEditInput();
    }

    // Chamado quando o texto de qualquer editor muda.
    onEditorInput() {
      if (!this.editing) {
        // digitar com a célula selecionada começa a edição (modo "Digite")
        const v = this.editor.value;
        if (v === '') return;
        this.beginEdit('enter', v);
        return;
      }
      if (this.editing.source === 'cell') this.app.fbar.value = this.editor.value;
      this.lastPoint = null;
      this.onEditInput();
    }
    onBarInput() {
      if (!this.editing) return;
      this.editor.value = this.app.fbar.value;
      this.lastPoint = null;
      this.onEditInput();
    }

    onEditInput(caretOnly) {
      if (!this.editing) return;
      this.hideError();
      if (!caretOnly) { this.positionEditor(); this.drawRefs(); }
      const pointing = this.drag && this.drag.type === 'point';
      if (!pointing) this.app.setMode(this.editing.mode === 'enter' ? 'Digite' : 'Editar');
      this.updateAC();
      this.updateTip();
    }

    commit(dr = 0, dc = 0) {
      if (!this.editing) return true;
      const { r, c, sheet } = this.editing;
      const text = this.activeInput().value;
      const res = this.app.commitInput(r, c, text, sheet);
      if (!res.ok) {
        if (res.dv) { this.showDvError(res.dv, r, c); return false; }
        this.showError(res.msg); this.focusActiveInput(); return false;
      }
      const away = this.isAway();
      this.endEdit();
      if (away) this.app.showSheet(sheet); // volta para a planilha da fórmula
      const nc = this.tabTarget(c, dr, dc);
      if (dr || dc) this.setSel(this.stepRow(r, dr), nc);
      else this.renderSel();
      return true;
    }

    // Como no Excel: depois de preencher uma linha com Tab, o Enter volta para a coluna onde você começou.
    tabTarget(c, dr, dc) {
      if (dc > 0 && !dr) { if (this.tabCol === undefined || this.tabCol === null) this.tabCol = c; return c + dc; }
      if (dr > 0 && !dc && this.tabCol !== undefined && this.tabCol !== null) { const t = this.tabCol; this.tabCol = null; return t; }
      this.tabCol = null;
      return c + dc;
    }

    cancel() {
      if (!this.editing) return;
      const away = this.isAway();
      const sheet = this.editing.sheet;
      this.endEdit();
      if (away) this.app.showSheet(sheet);
      this.renderSel();
    }

    endEdit() {
      const src = this.editing && this.editing.source;
      this.editing = null;
      this.lastPoint = null;
      this.editor.classList.add('idle');
      this.editor.value = '';
      ['fontWeight', 'fontStyle', 'fontFamily', 'fontSize', 'width'].forEach((p) => (this.editor.style[p] = ''));
      this.hidePopups();
      this.hideError();
      this.refLayer.innerHTML = '';
      this.app.setMode('Pronto');
      this.app.updateFormulaBar();
      if (src === 'bar') this.app.fbar.blur();
      this.editor.focus({ preventScroll: true });
      this.app.emit('editend');
    }

    // Está num ponto da fórmula onde clicar numa célula insere a referência?
    isPointable() {
      if (!this.editing) return false;
      const inp = this.activeInput();
      const t = inp.value;
      if (t[0] !== '=') return false;
      const pos = inp.selectionStart;
      if (this.lastPoint && this.lastPoint.text === t && this.lastPoint.end === inp.selectionEnd) return true;
      const before = t.slice(0, pos).replace(/\s+$/, '');
      const ch = before[before.length - 1];
      return !!ch && '=(;,+-*/^&<>:'.includes(ch);
    }
    startPoint(cell) {
      const inp = this.activeInput();
      let start, end;
      if (this.lastPoint && this.lastPoint.text === inp.value && this.lastPoint.end === inp.selectionEnd) {
        start = this.lastPoint.start; end = this.lastPoint.end;
      } else { start = inp.selectionStart; end = inp.selectionEnd; }
      this.drag = { type: 'point', anchor: cell };
      this.replacePoint(start, end, this.pointRef({ r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c }), cell);
      this.app.setMode('Aponte');
    }
    updatePoint(cell) {
      const a = this.drag.anchor;
      const g = { r1: Math.min(a.r, cell.r), c1: Math.min(a.c, cell.c), r2: Math.max(a.r, cell.r), c2: Math.max(a.c, cell.c) };
      this.replacePoint(this.lastPoint.start, this.lastPoint.end, this.pointRef(g), cell);
    }
    // Texto da referência ao clicar/arrastar durante uma fórmula. Dentro de tabelas, usa [@Coluna] ou Tabela1[Coluna].
    pointRef(g) {
      const prefix = this.pointPrefix();
      if (!prefix && this.editing) {
        const S = this.sheet;
        const t = S.tableAt(g.r1, g.c1);
        if (t && g.c1 === g.c2 && t === S.tableAt(g.r2, g.c2)) {
          const dr = S.tableDataRows(t);
          const header = S.tableHeader(t, g.c1);
          const ed = this.editing;
          if (header) {
            if (g.r1 === g.r2 && g.r1 === ed.r && g.r1 >= dr.r1 && g.r1 <= dr.r2 && XL.inRange(t.g, ed.r, ed.c)) return XL.srefText(header);
            if (g.r1 === dr.r1 && g.r2 === dr.r2) return XL.srefText(header, t.name);
          }
        }
      }
      return prefix + XL.rangeToStr(g);
    }

    showDvError(rule, r, c) {
      const msg = rule.errorMsg || 'Este valor não corresponde às restrições de validação de dados definidas para esta célula.';
      this.hidePopups();
      XL.ui.modal({
        title: rule.errorTitle || 'Valor inválido', width: 430, className: 'dv-modal',
        body: `<div class="dv-err"><span class="dv-x">✖</span><p>${XL.escapeHtml(msg)}</p></div>`,
        buttons: [
          { label: 'Repetir', primary: true, role: 'dv-retry', onClick: () => { setTimeout(() => { const inp = this.activeInput(); inp.focus(); inp.select(); }, 0); } },
          { label: 'Cancelar', role: 'dv-cancel', onClick: () => { this.cancel(); } },
        ],
      });
      this.app.emit('dverror', { ref: XL.toRef(r, c) });
    }

    // Ao apontar células de outra planilha, a referência leva o nome dela: Janeiro!B6
    pointPrefix() { return this.isAway() ? XL.quoteSheet(this.sheet.name) + '!' : ''; }
    replacePoint(start, end, ref, cell) {
      const inp = this.activeInput();
      const t = inp.value.slice(0, start) + ref + inp.value.slice(end);
      inp.value = t;
      const pos = start + ref.length;
      inp.setSelectionRange(pos, pos);
      if (inp === this.editor) this.app.fbar.value = t; else this.editor.value = t;
      this.lastPoint = { start, end: pos, text: t, cell };
      this.positionEditor();
      this.drawRefs();
      this.hidePopups();
      this.updateTip();
    }

    // ---------- Teclado ----------
    onKey(e) {
      if (this.editing) { this.onEditKey(e); return; }
      const k = e.key;
      const ctrl = e.ctrlKey || e.metaKey;
      const app = this.app;
      if (ctrl) {
        const lk = k.toLowerCase();
        if (lk === 'z') { e.preventDefault(); app.undo(); return; }
        if (lk === 'y') { e.preventDefault(); app.redo(); return; }
        if (lk === 'l' && e.shiftKey) { e.preventDefault(); app.exec('filter-toggle'); return; }
        if (lk === 'n' || lk === 'b') { e.preventDefault(); app.exec('bold'); return; }
        if (lk === 'i') { e.preventDefault(); app.exec('italic'); return; }
        if (lk === 's') { e.preventDefault(); app.exec('underline'); return; }
        if (lk === 'l' || lk === 'f') { e.preventDefault(); app.exec('find'); return; }
        if (lk === 'u' || lk === 'h') { e.preventDefault(); app.exec('replace'); return; }
        if (lk === 'p') { e.preventDefault(); app.exec('print-preview'); return; }
        if (k === '1') { e.preventDefault(); app.exec('format-cells-dialog'); return; }
        if (lk === 'e') { e.preventDefault(); app.exec('flashfill'); return; }
        if (lk === 't') { e.preventDefault(); app.exec('table-create'); return; }
        if (lk === 'd') { e.preventDefault(); app.exec('fill-down'); return; }
        if (k === '+' || k === '=') { e.preventDefault(); app.exec('insert-smart'); return; }
        if (k === '-') { e.preventDefault(); app.exec('delete-smart'); return; }
        if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); app.showSheetOffset(k === 'PageDown' ? 1 : -1); return; }
        if (lk === 'a') { e.preventDefault(); this.setSel(0, 0, this.R - 1, this.C - 1); return; }
        if (lk === 'c') { app.copy(false); return; } // o evento "copy" também é disparado
        if (lk === 'x') { app.copy(true); return; }
        if (k === '`' || k === "'" || k === '"') { e.preventDefault(); app.exec('show-formulas'); return; }
        if (k === 'Home') { e.preventDefault(); this.setSel(0, 0); return; }
        if (k.startsWith('Arrow')) { e.preventDefault(); this.ctrlArrow(k, e.shiftKey); return; }
        return;
      }
      if (e.altKey && (k === '=' || k === '+')) { e.preventDefault(); app.exec('autosum', 'SOMA'); return; }
      if (e.altKey && k === 'ArrowDown') {
        const rule = this.sheet.validationAt(this.sel.ar, this.sel.ac);
        if (rule && rule.type === 'list') { e.preventDefault(); app.openDvList(this.sel.ar, this.sel.ac, this.dvLayer.querySelector('.xl-dvbtn') || this.tds[this.sel.ar][this.sel.ac]); return; }
      }
      const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], PageDown: [20, 0], PageUp: [-20, 0] };
      if (moves[k]) {
        e.preventDefault();
        const [dr, dc] = moves[k];
        this.tabCol = null;
        this.moveBy(dr, dc, e.shiftKey);
        return;
      }
      switch (k) {
        case 'Enter': {
          e.preventDefault();
          const dr = e.shiftKey ? -1 : 1;
          this.setSel(this.stepRow(this.sel.ar, dr), this.tabTarget(this.sel.ac, dr, 0));
          return;
        }
        case 'Tab': e.preventDefault(); this.setSel(this.sel.ar, this.tabTarget(this.sel.ac, 0, e.shiftKey ? -1 : 1)); return;
        case 'F2': e.preventDefault(); this.beginEdit('edit'); return;
        case 'Delete': e.preventDefault(); if (app.selectedChart) app.exec('chart-delete'); else app.exec('clear', 'contents'); return;
        case 'Backspace': e.preventDefault(); this.beginEdit('enter', ''); return;
        case 'Escape': e.preventDefault(); app.clearCopy(); if (app.painterFmt) app.cancelPainter(); return;
        case 'Home': e.preventDefault(); this.setSel(this.sel.ar, 0); return;
        case 'F4': e.preventDefault(); return;
      }
      if (e.shiftKey && k === 'F3') { e.preventDefault(); app.exec('fx'); }
    }

    ctrlArrow(k, extend) {
      const [dr, dc] = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
      let r = extend ? this.sel.fr : this.sel.ar, c = extend ? this.sel.fc : this.sel.ac;
      const has = (rr, cc) => rr >= 0 && cc >= 0 && rr < this.R && cc < this.C && this.sheet.has(rr, cc);
      if (has(r, c) && has(r + dr, c + dc)) {
        while (has(r + dr, c + dc)) { r += dr; c += dc; }
      } else {
        r += dr; c += dc;
        while (r >= 0 && c >= 0 && r < this.R && c < this.C && !has(r, c)) { r += dr; c += dc; }
        if (!(r >= 0 && c >= 0 && r < this.R && c < this.C)) { r -= dr; c -= dc; }
      }
      if (extend) this.setFocusCell(r, c); else this.setSel(r, c);
    }

    onEditKey(e) {
      const k = e.key;
      if (this.ac) {
        if (k === 'ArrowDown' || k === 'ArrowUp') {
          e.preventDefault();
          this.ac.idx = (this.ac.idx + (k === 'ArrowDown' ? 1 : -1) + this.ac.items.length) % this.ac.items.length;
          this.renderAC();
          return;
        }
        if (k === 'Tab') { e.preventDefault(); this.acceptAC(this.ac.idx); return; }
        if (k === 'Escape') { e.preventDefault(); this.hidePopups(); return; }
      }
      if (k === 'Enter') {
        e.preventDefault();
        if (e.altKey) { this.insertText('\n'); return; }
        this.commit(e.shiftKey ? -1 : 1, 0);
        return;
      }
      if (k === 'Tab') { e.preventDefault(); this.commit(0, e.shiftKey ? -1 : 1); return; }
      if (k === 'Escape') { e.preventDefault(); this.cancel(); return; }
      if (k === 'F4') {
        e.preventDefault();
        const inp = this.activeInput();
        const res = XL.Formula.toggleAbsolute(inp.value, inp.selectionStart);
        if (res) {
          inp.value = res.text;
          inp.setSelectionRange(res.caret, res.caret);
          if (inp === this.editor) this.app.fbar.value = res.text; else this.editor.value = res.text;
          this.lastPoint = null;
          this.onEditInput();
          this.app.emit('f4', {});
        }
        return;
      }
      if (k === 'F2') {
        e.preventDefault();
        this.editing.mode = this.editing.mode === 'enter' ? 'edit' : 'enter';
        this.app.setMode(this.editing.mode === 'enter' ? 'Digite' : 'Editar');
        return;
      }
      const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k] && this.editing.mode === 'enter') {
        const [dr, dc] = arrows[k];
        const inp = this.activeInput();
        if (inp.value[0] === '=') {
          if (this.isPointable()) {
            e.preventDefault();
            const base = (this.lastPoint && this.lastPoint.cell) || { r: this.editing.r, c: this.editing.c };
            const cell = { r: this.clampR(base.r + dr), c: this.clampC(base.c + dc) };
            const lp = this.lastPoint && this.lastPoint.text === inp.value && this.lastPoint.end === inp.selectionStart ? this.lastPoint : null;
            const start = lp ? lp.start : inp.selectionStart;
            const end = lp ? lp.end : inp.selectionStart;
            this.replacePoint(start, end, this.pointRef({ r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c }), cell);
            this.app.setMode('Aponte');
          }
          return;
        }
        e.preventDefault();
        this.commit(dr, dc);
      }
    }

    insertText(s) {
      const inp = this.activeInput();
      const a = inp.selectionStart, b = inp.selectionEnd;
      inp.value = inp.value.slice(0, a) + s + inp.value.slice(b);
      inp.setSelectionRange(a + s.length, a + s.length);
      if (inp === this.editor) this.app.fbar.value = inp.value; else this.editor.value = inp.value;
      this.onEditInput();
    }

    // ---------- Área de transferência do sistema ----------
    onCopyEvent(e, cut) {
      if (this.editing) return;
      const tsv = this.app.selectionTSV();
      e.clipboardData.setData('text/plain', tsv);
      e.preventDefault();
      if (!this.app.clip || this.app.clipTSV !== tsv) this.app.copy(cut);
    }
    onPasteEvent(e) {
      if (this.editing) return;
      e.preventDefault();
      const text = e.clipboardData.getData('text/plain');
      if (this.app.clip && text.replace(/\r/g, '') === (this.app.clipTSV || '').replace(/\r/g, '')) this.app.paste();
      else this.app.pasteText(text);
    }

    // ---------- Autocompletar e dica de sintaxe ----------
    hidePopups(keepState) {
      this.acEl.style.display = 'none';
      this.tipEl.style.display = 'none';
      if (!keepState) this.ac = null;
    }
    popupAnchor() {
      const inp = this.activeInput();
      const r = inp.getBoundingClientRect();
      return { x: r.left, y: r.bottom + 2 };
    }
    updateAC() {
      const inp = this.activeInput();
      const t = inp.value;
      this.ac = null;
      this.acEl.style.display = 'none';
      if (t[0] !== '=') return;
      const pos = inp.selectionStart;
      if (pos !== inp.selectionEnd) return;
      const m = /(?:^|[=(;,+\-*/^&<>:\s])([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9_.]*)$/.exec(t.slice(0, pos));
      if (!m) return;
      const prefix = m[1];
      if (/^[A-Za-z]{1,3}\d+$/.test(prefix)) return;
      const items = XL.Functions.search(prefix);
      if (!items.length) return;
      this.ac = { items, idx: 0, start: pos - prefix.length, end: pos };
      this.renderAC();
    }
    renderAC() {
      const el = this.acEl;
      el.innerHTML = '';
      this.ac.items.forEach((d, i) => {
        el.appendChild(h('div.xl-ac-item' + (i === this.ac.idx ? '.on' : ''), {
          onmousedown: (e) => { e.preventDefault(); this.acceptAC(i); },
        }, h('span.xl-ac-ico', 'fx'), h('span.xl-ac-name', d.name)));
      });
      const cur = this.ac.items[this.ac.idx];
      el.appendChild(h('div.xl-ac-desc', cur.desc));
      el.appendChild(h('div.xl-ac-hint', 'Tab para inserir · ↑↓ para escolher'));
      const p = this.popupAnchor();
      Object.assign(el.style, { display: '', left: p.x + 'px', top: p.y + 'px' });
      this.tipEl.style.display = 'none';
    }
    acceptAC(i) {
      const d = this.ac.items[i];
      const inp = this.activeInput();
      const t = inp.value;
      const nt = t.slice(0, this.ac.start) + d.name + '(' + t.slice(this.ac.end);
      const pos = this.ac.start + d.name.length + 1;
      inp.value = nt;
      inp.setSelectionRange(pos, pos);
      if (inp === this.editor) this.app.fbar.value = nt; else this.editor.value = nt;
      this.ac = null;
      this.onEditInput();
      this.focusActiveInput();
    }
    updateTip() {
      if (this.ac) return;
      const inp = this.activeInput();
      const ctx = XL.Formula.funcContext(inp.value, inp.selectionStart);
      const def = ctx && XL.Functions.get(ctx.name);
      if (!def) { this.tipEl.style.display = 'none'; return; }
      const n = def.args.length;
      const hi = def.args[n - 1] === '...' ? Math.min(ctx.arg, n - 2) : ctx.arg;
      const argHtml = def.args.map((a, i) => (i === hi ? `<b>${XL.escapeHtml(a)}</b>` : XL.escapeHtml(a))).join('; ');
      this.tipEl.innerHTML = `<span class="fn">${def.name}</span>(${argHtml})`;
      const p = this.popupAnchor();
      Object.assign(this.tipEl.style, { display: '', left: p.x + 'px', top: p.y + 'px' });
    }

    // ---------- Erro de fórmula ----------
    showError(msg) {
      this.hidePopups();
      if (!this.errEl) { this.errEl = h('div.xl-ferr'); document.body.appendChild(this.errEl); }
      this.errEl.innerHTML = `<b>Há um problema com esta fórmula.</b><br>${XL.escapeHtml(msg)}<div class="xl-ferr-foot">Corrija e pressione Enter, ou Esc para cancelar.</div>`;
      const p = this.popupAnchor();
      Object.assign(this.errEl.style, { display: '', left: p.x + 'px', top: p.y + 4 + 'px' });
      this.app.emit('formulaerror', { msg });
    }
    hideError() { if (this.errEl) this.errEl.style.display = 'none'; }
  }

  XL.Grid = Grid;
})();
