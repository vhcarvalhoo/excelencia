/* Extensão do ExcelApp: abas de planilha, inserir/excluir linhas e colunas, classificar, filtrar,
   remover duplicatas, congelar painéis, localizar e configurar página. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const I = XL.icons;

  Object.assign(XL.ExcelApp.prototype, {
    execExtra(cmd, arg) {
      switch (cmd) {
        // planilhas
        case 'sheet-add': return this.cmdSheetAdd();
        case 'sheet-delete': return this.cmdSheetDelete(this.sheet);
        case 'sheet-rename': return this.startRename(this.sheet);
        // linhas e colunas
        case 'insert-rows': { const g = this.grid.range(); return this.structural('insert', 'row', g.r1, g.r2 - g.r1 + 1); }
        case 'delete-rows': { const g = this.grid.range(); return this.structural('delete', 'row', g.r1, g.r2 - g.r1 + 1); }
        case 'insert-cols': { const g = this.grid.range(); return this.structural('insert', 'col', g.c1, g.c2 - g.c1 + 1); }
        case 'delete-cols': { const g = this.grid.range(); return this.structural('delete', 'col', g.c1, g.c2 - g.c1 + 1); }
        case 'insert-dialog': case 'insert-smart': return this.cellsDialog('insert');
        case 'delete-dialog': case 'delete-smart': return this.cellsDialog('delete');
        // classificar e filtrar
        case 'sort': return this.sortQuick(arg === 'desc');
        case 'sort-dialog': return XL.Dialogs.sort(this);
        case 'filter-toggle': return this.filterToggle();
        case 'filter-clear': return this.filterClear();
        case 'filter-reapply': return this.filterReapply();
        case 'filter-cell': return this.filterByCell();
        case 'dedupe': return XL.Dialogs.removeDuplicates(this);
        case 'texttocols': return XL.Dialogs.textToColumns(this);
        case 'flashfill': return XL.Dialogs.flashFill(this);
        case 'fill-down': return this.fillDownCopy();
        case 'format-cells-dialog': return XL.Dialogs.formatCells(this);
        // exibir
        case 'freeze': return this.freezePanes(arg);
        case 'find': return XL.Find.open(this, 'find');
        case 'replace': return XL.Find.open(this, 'replace');
        case 'goto':
          this.namebox.focus();
          XL.ui.toast('Digite o endereço (ex.: <b>H10</b>) na Caixa de Nome e pressione Enter.');
          return;
        // página
        case 'page-orient': return this.setPage({ orientation: arg });
        case 'page-size': return this.setPage({ size: arg });
        case 'page-margins': return this.setPage({ margins: arg });
        case 'page-fit': return this.setPage({ fit: arg });
        case 'page-fitw': case 'page-fith': {
          const f = this.sheet.page.fit;
          let w = f === 'sheet' || f === 'cols', hh = f === 'sheet' || f === 'rows';
          if (cmd === 'page-fitw') w = arg === '1'; else hh = arg === '1';
          return this.setPage({ fit: w && hh ? 'sheet' : w ? 'cols' : hh ? 'rows' : 'none' });
        }
        case 'print-area': {
          if (arg === 'clear') return this.setPage({ printArea: null });
          const g = this.grid.range();
          if (g.r1 === g.r2 && g.c1 === g.c2) { XL.ui.toast('Selecione primeiro o intervalo que deseja imprimir (ex.: A1:F20).'); return; }
          return this.setPage({ printArea: g });
        }
        case 'print-titles': return XL.Dialogs.printTitles(this);
        case 'print-gridlines': return this.setPage({ printGridlines: !!arg });
        case 'print-preview': return XL.Print.preview(this);
      }
    },

    // ================= Abas de planilha =================
    renderTabs() {
      const box = this.tabsBox;
      if (!box) return;
      box.innerHTML = '';
      this.wb.sheets.forEach((s, i) => {
        const b = h('button.xl-sheettab' + (i === this.wb.activeIndex ? '.active' : ''), {
          type: 'button', dataset: { sheet: s.name, i },
          title: 'Clique para abrir · duplo clique para renomear · botão direito para mais opções',
        }, s.name);
        b.addEventListener('mousedown', (e) => e.preventDefault());
        b.addEventListener('click', () => this.onTabClick(i));
        b.addEventListener('dblclick', () => this.startRename(s));
        b.addEventListener('contextmenu', (e) => { e.preventDefault(); this.tabMenu(s, { x: e.clientX, y: e.clientY }); });
        box.appendChild(b);
      });
    },

    onTabClick(i) {
      const g = this.grid;
      const target = this.wb.sheets[i];
      if (!target || (target === this.sheet && !g.isAway())) return;
      if (g.editing) {
        // editando uma fórmula: dá para ir até outra planilha e clicar nas células dela
        const isFormula = g.activeInput().value[0] === '=';
        if (isFormula && (g.isPointable() || g.isAway())) { this.showSheet(i, true); return; }
        if (!g.commit()) return;
      }
      this.showSheet(i);
    },

    showSheet(sheetOrIndex, keepEdit) {
      const i = typeof sheetOrIndex === 'number' ? sheetOrIndex : this.wb.sheets.indexOf(sheetOrIndex);
      if (i < 0 || i >= this.wb.sheets.length) return;
      const g = this.grid;
      if (!g.isAway()) this.sheet.sel = Object.assign({}, g.sel);
      const changed = i !== this.wb.activeIndex;
      this.wb.activeIndex = i;
      g.reloadSheet();
      if (keepEdit && g.editing) {
        if (g.editing.source !== 'bar') {
          g.editing.source = 'bar';
          const n = this.fbar.value.length;
          this.fbar.focus();
          this.fbar.setSelectionRange(n, n);
        }
        g.positionEditor();
        g.drawRefs();
        this.setMode(g.isAway() ? 'Aponte' : 'Digite');
      }
      this.renderTabs();
      this.updateNameBox();
      this.updateFormulaBar();
      this.updateStatus();
      this.ribbon.sync();
      if (changed) this.emit('sheetswitch', { name: this.sheet.name });
    },

    showSheetOffset(d) {
      const i = Math.max(0, Math.min(this.wb.sheets.length - 1, this.wb.activeIndex + d));
      if (i === this.wb.activeIndex) return;
      if (this.grid.editing && !this.grid.commit()) return;
      this.showSheet(i);
      this.grid.focus();
    },

    cmdSheetAdd(name) {
      this.pushUndo();
      const s = this.wb.addSheet(name, this.wb.activeIndex + 1);
      this.showSheet(s);
      this.emit('sheetadd', { name: s.name });
      this.emit('change', {});
      return s;
    },

    cmdSheetDelete(sheet) {
      if (this.wb.sheets.length === 1) { XL.ui.toast('Uma pasta de trabalho precisa ter pelo menos uma planilha.'); return; }
      const doIt = () => {
        this.pushUndo();
        const name = sheet.name;
        this.wb.removeSheet(sheet);
        this.grid.reloadSheet();
        this.renderTabs();
        this.render();
        this.emit('sheetdelete', { name });
        this.emit('change', {});
      };
      if (sheet.cells.size) {
        XL.ui.modal({
          title: 'Excluir planilha',
          body: `<p>A planilha <b>${XL.escapeHtml(sheet.name)}</b> tem dados. Deseja excluí-la?</p><p style="color:#6b7480">Fórmulas de outras planilhas que usam dados dela vão mostrar <b>#REF!</b>. No simulador você pode desfazer com Ctrl+Z.</p>`,
          buttons: [{ label: 'Excluir', primary: true, onClick: doIt }, { label: 'Cancelar' }],
          onClose: () => this.grid.focus(),
        });
      } else doIt();
    },

    startRename(sheet) {
      const g = this.grid;
      if (g.editing && !g.commit()) return;
      this.renderTabs();
      const btn = this.tabsBox.querySelector(`[data-i="${this.wb.sheets.indexOf(sheet)}"]`);
      if (!btn) return;
      const inp = h('input.xl-tabinput', { type: 'text', value: sheet.name, maxlength: '31', 'aria-label': 'Nome da planilha', spellcheck: 'false' });
      btn.replaceWith(inp);
      inp.focus();
      inp.select();
      let done = false;
      const finish = (ok, fromBlur) => {
        if (done) return;
        if (ok) {
          const errMsg = this.renameSheet(sheet, inp.value);
          if (errMsg) {
            XL.ui.toast('⚠ ' + errMsg);
            if (!fromBlur) { inp.focus(); inp.select(); return; }
          }
        }
        done = true;
        this.renderTabs();
        this.grid.focus();
      };
      inp.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
      inp.addEventListener('blur', () => finish(true, true));
      this.emit('renamestart', { name: sheet.name });
    },

    renameSheet(sheet, name) {
      name = String(name).trim();
      if (name === sheet.name) return null;
      const e = this.wb.validateName(name, sheet);
      if (e) return e;
      this.pushUndo();
      const old = sheet.name;
      this.wb.renameSheet(sheet, name);
      this.render();
      this.renderTabs();
      this.emit('sheetrename', { from: old, to: sheet.name });
      this.emit('change', {});
      return null;
    },

    tabMenu(sheet, pt) {
      XL.ui.menu(pt, [
        { label: 'Inserir nova planilha', data: { cmd: 'sheet-add' }, onClick: () => this.exec('sheet-add') },
        { label: 'Excluir', data: { cmd: 'sheet-delete' }, onClick: () => this.cmdSheetDelete(sheet) },
        { label: 'Renomear', data: { cmd: 'sheet-rename' }, onClick: () => this.startRename(sheet) },
      ]);
    },

    // ================= Menu do botão direito =================
    showContextMenu(kind, pt) {
      const m = (label, cmd, arg, icon) => ({ label, icon: icon ? I[icon] : '', data: { cmd, arg: arg === undefined ? '' : String(arg) }, onClick: () => this.exec(cmd, arg) });
      const items = [
        m('Recortar', 'cut', undefined, 'cut'),
        m('Copiar', 'copy', undefined, 'copy'),
        m('Colar', 'paste'),
        m('Colar Valores', 'paste', 'values'),
        { sep: true },
      ];
      if (kind === 'row') items.push(m('Inserir', 'insert-rows', undefined, 'insertCells'), m('Excluir', 'delete-rows', undefined, 'deleteCells'));
      else if (kind === 'col') items.push(m('Inserir', 'insert-cols', undefined, 'insertCells'), m('Excluir', 'delete-cols', undefined, 'deleteCells'));
      else items.push(m('Inserir...', 'insert-dialog', undefined, 'insertCells'), m('Excluir...', 'delete-dialog', undefined, 'deleteCells'));
      items.push(m('Limpar Conteúdo', 'clear', 'contents'));
      items.push({ sep: true }, m('Formatar Células...', 'format-cells-dialog', undefined, 'formatCells'));
      if (kind === 'cell') {
        items.push({ sep: true },
          m('Filtrar pelo valor da célula', 'filter-cell'),
          m('Classificar de A a Z', 'sort', 'asc', 'sortAZ'),
          m('Classificar de Z a A', 'sort', 'desc', 'sortZA'));
      }
      XL.ui.menu(pt, items);
    },

    // ================= Inserir / excluir =================
    structural(kind, axis, at, n) {
      this.pushUndo();
      this.clearCopy();
      this.wb.structural(this.sheet, kind, axis, at, n);
      this.grid.applyColWidths();
      this.render();
      this.grid.renderSel();
      const label = axis === 'row' ? `${at + 1}:${at + n}` : `${XL.numToCol(at + 1)}:${XL.numToCol(at + n)}`;
      this.emit(kind, { axis, at, n, ref: label });
      this.emit('change', {});
    },

    cellsDialog(kind) {
      const g = this.grid.range();
      const fullRow = g.c1 === 0 && g.c2 === this.sheet.cols - 1;
      const fullCol = g.r1 === 0 && g.r2 === this.sheet.rows - 1;
      if (fullRow) return this.exec(kind + '-rows');
      if (fullCol) return this.exec(kind + '-cols');
      const verb = kind === 'insert' ? 'Inserir' : 'Excluir';
      let choice = 'row';
      const opt = (value, label) => {
        const r = h('input', { type: 'radio', name: 'cd', value });
        if (value === 'row') r.checked = true;
        r.addEventListener('change', () => { choice = value; });
        return h('label.cd-opt', r, ' ' + label);
      };
      XL.ui.modal({
        title: verb, width: 340,
        body: h('div.cd',
          opt('row', 'Linha inteira'), opt('col', 'Coluna inteira'),
          h('p.cd-note', 'Dica: clique com o botão direito no número da linha (ou letra da coluna) para inserir/excluir direto.')),
        buttons: [{ label: 'OK', primary: true, onClick: () => this.exec(kind + (choice === 'row' ? '-rows' : '-cols')) }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },

    // Ctrl+D: copia a primeira linha da seleção para as de baixo (sem continuar séries).
    fillDownCopy() {
      const g = this.grid.range();
      if (g.r1 === g.r2) { XL.ui.toast('Selecione a célula de cima e as de baixo que devem receber a cópia (ex.: B2:B10).'); return; }
      const S = this.sheet;
      this.pushUndo();
      for (let c = g.c1; c <= g.c2; c++) {
        const src = S.cell(g.r1, c);
        for (let r = g.r1 + 1; r <= g.r2; r++) {
          if (src && src.kind === 'formula') S.setInput(r, c, XL.Formula.shiftFormula(src.input, r - g.r1, 0));
          else S.putCell(r, c, src);
          S.replaceFmt(r, c, src ? src.fmt : {});
        }
      }
      this.render();
      this.emit('fill', { src: XL.rangeToStr({ r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c2 }), dst: XL.rangeToStr(g), dir: 'down' });
      this.emit('change', {});
    },

    // ================= Classificar =================
    // Intervalo de dados para classificar/filtrar: a seleção, ou o bloco contínuo ao redor da célula.
    dataRegion() {
      const S = this.sheet;
      const g = this.grid.range();
      const { ar, ac } = this.grid.sel;
      const single = g.r1 === g.r2 && g.c1 === g.c2;
      if (S.filter) {
        const f = S.filter.g;
        if (ar >= f.r1 && ar <= f.r2 && ac >= f.c1 && ac <= f.c2) return { g: Object.assign({}, f), header: true, filter: true };
      }
      if (single || g.c1 === g.c2) {
        const reg = S.currentRegion(ar, ac);
        if (reg.r1 === reg.r2 && reg.c1 === reg.c2 && !S.has(ar, ac)) return null;
        if (!single && (reg.c1 < g.c1 || reg.c2 > g.c2)) XL.ui.toast('A seleção foi <b>expandida</b> para incluir as colunas vizinhas — assim cada linha continua inteira.');
        return { g: reg, header: S.detectHeader(reg) };
      }
      return { g, header: S.detectHeader(g) };
    },

    sortQuick(desc) {
      const d = this.dataRegion();
      if (!d) { XL.ui.toast('Clique em uma célula dentro dos dados que você quer classificar.'); return; }
      const col = Math.max(d.g.c1, Math.min(d.g.c2, this.grid.sel.ac));
      this.sortRange(d.g, [{ col, desc }], d.header);
    },

    sortRange(g, keys, header) {
      const S = this.sheet;
      this.pushUndo();
      S.sortRange(g, keys, header);
      if (S.filter) S.applyFilter();
      this.render();
      this.emit('sort', { range: XL.rangeToStr(g), header, keys: keys.map((k) => ({ col: XL.numToCol(k.col + 1), desc: !!k.desc })) });
      this.emit('change', {});
    },

    removeDuplicates(g, cols, header) {
      const S = this.sheet;
      this.pushUndo();
      const res = S.removeDuplicates(g, cols, header);
      if (S.filter) S.applyFilter();
      this.render();
      this.emit('dedupe', res);
      this.emit('change', {});
      XL.ui.modal({
        title: 'Remover Duplicatas', width: 420,
        body: res.removed
          ? `<p><b>${res.removed}</b> ${res.removed === 1 ? 'valor duplicado encontrado e removido' : 'valores duplicados encontrados e removidos'}; <b>${res.remaining}</b> ${res.remaining === 1 ? 'valor exclusivo permanece' : 'valores exclusivos permanecem'}.</p>`
          : '<p>Nenhum valor duplicado foi encontrado.</p>',
        buttons: [{ label: 'OK', primary: true }],
        onClose: () => this.grid.focus(),
      });
      return res;
    },

    // ================= Filtro =================
    filterToggle() {
      const S = this.sheet;
      if (S.filter) {
        this.pushUndo();
        S.filter = null;
        S.applyFilter();
        this.render();
        this.emit('filter', { on: false });
        this.emit('change', {});
        return;
      }
      const d = this.dataRegion();
      if (!d || d.g.r1 === d.g.r2) { XL.ui.toast('Clique em uma célula dentro da tabela (com cabeçalho) antes de ativar o filtro.'); return; }
      this.pushUndo();
      S.filter = { g: Object.assign({}, d.g), crit: {} };
      S.applyFilter();
      this.render();
      this.emit('filter', { on: true, range: XL.rangeToStr(d.g) });
      this.emit('change', {});
    },

    setFilterCriteria(col, crit) {
      const S = this.sheet;
      if (!S.filter) return;
      this.pushUndo();
      if (crit) S.filter.crit[col] = crit;
      else delete S.filter.crit[col];
      S.applyFilter();
      this.render();
      const st = S.filterStats();
      this.emit('filter', { on: true, col: XL.numToCol(col + 1), crit, shown: st.shown, total: st.total });
      this.emit('change', {});
    },

    filterClear() {
      const S = this.sheet;
      if (!S.filter) return;
      this.pushUndo();
      S.filter.crit = {};
      S.applyFilter();
      this.render();
      this.emit('filter', { on: true, cleared: true });
      this.emit('change', {});
    },

    filterReapply() {
      const S = this.sheet;
      if (!S.filter) return;
      S.applyFilter();
      this.render();
      this.emit('filter', { on: true, reapplied: true });
    },

    filterByCell() {
      const S = this.sheet;
      const { ar, ac } = this.grid.sel;
      if (!S.filter) this.filterToggle();
      if (!S.filter) return;
      if (ar <= S.filter.g.r1) return;
      this.setFilterCriteria(ac, { type: 'values', values: [S.textOf(ar, ac)], blanks: S.textOf(ar, ac) === '' });
    },

    openFilterMenu(col, anchor) { XL.Dialogs.filterMenu(this, col, anchor); },

    // ================= Congelar painéis =================
    freezePanes(kind) {
      const S = this.sheet;
      const { ar, ac } = this.grid.sel;
      if (kind === 'panes' && ar === 0 && ac === 0) {
        XL.ui.toast('Selecione a célula logo <b>abaixo</b> e à <b>direita</b> do que deseja congelar (ex.: <b>B2</b> congela a linha 1 e a coluna A).');
        return;
      }
      this.pushUndo();
      if (kind === 'none') S.freeze = null;
      else if (kind === 'top') S.freeze = { rows: 1, cols: 0 };
      else if (kind === 'first') S.freeze = { rows: 0, cols: 1 };
      else S.freeze = { rows: ar, cols: ac };
      this.grid.scroller.scrollTop = 0;
      this.grid.scroller.scrollLeft = 0;
      this.render();
      this.emit('freeze', { rows: S.freeze ? S.freeze.rows : 0, cols: S.freeze ? S.freeze.cols : 0 });
      this.emit('change', {});
    },

    // ================= Página =================
    setPage(patch) {
      const S = this.sheet;
      this.pushUndo();
      Object.assign(S.page, patch);
      S.showBreaks = true;
      this.render();
      this.ribbon.sync();
      this.emit('pagesetup', Object.assign({}, S.page));
      this.emit('change', {});
      if (XL.Print && XL.Print.current) XL.Print.current.refresh();
    },
  });
})();
