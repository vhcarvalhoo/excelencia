/* Extensão do ExcelApp (Módulo 7): tabelas, gráficos, formatação condicional e validação de dados. */
(function () {
  'use strict';
  const XL = window.XL;
  const K = XL.SHEET_KEY;
  const proto = XL.ExcelApp.prototype;
  const baseExtra = proto.execExtra;
  const overlap = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

  Object.assign(proto, {
    execExtra(cmd, arg) {
      switch (cmd) {
        // tabelas
        case 'table-create': return XL.Dialogs.createTable(this, arg || 'azul');
        case 'table-style': return this.tableUpdate((t) => { t.style = arg; }, 'tablestyle');
        case 'table-banded': return this.tableUpdate((t) => { t.banded = !!arg; }, 'tablebanded');
        case 'table-totals': return this.tableTotals(!!arg);
        case 'table-filterbtn': return this.tableFilterBtn(!!arg);
        case 'table-rename': return this.tableRename(arg);
        case 'table-convert': return this.tableConvert();
        // gráficos
        case 'chart-insert': {
          if (arg !== 'rec') return this.insertChart(arg || 'col');
          const t = this.recommendChart();
          if (t) XL.ui.toast(t === 'line' ? '💡 Recomendado: <b>Linhas</b> — seus rótulos parecem uma sequência no tempo.' : '💡 Recomendado: <b>Colunas Agrupadas</b> — o melhor para comparar categorias.');
          return this.insertChart(t || 'col');
        }
        case 'chart-type': return this.chartUpdate((c) => { c.type = arg; });
        case 'chart-title': return XL.Dialogs.chartTitle(this);
        case 'chart-legend': return this.chartUpdate((c) => { c.legend = !!arg; });
        case 'chart-labels': return this.chartUpdate((c) => { c.labels = !!arg; });
        case 'chart-delete': return this.deleteChart();
        // formatação condicional
        case 'cf': {
          const a = String(arg || '');
          if (a.startsWith('bar-')) return this.addCF({ kind: 'bar', color: a.slice(4) });
          if (a.startsWith('scale-')) return this.addCF({ kind: 'scale', scale: a.slice(6) });
          return XL.Dialogs.conditional(this, a);
        }
        case 'cf-clear': return this.clearCF(arg);
        // validação
        case 'dv-dialog': return XL.Dialogs.validation(this);
        case 'dv-clear': return this.clearDV();
      }
      return baseExtra.call(this, cmd, arg);
    },

    // ================= Tabelas =================
    activeTable() { const { ar, ac } = this.grid.sel; return this.sheet.tableAt(ar, ac); },

    createTable(style, rangeStr) {
      const S = this.sheet;
      let g = rangeStr ? XL.parseRange(String(rangeStr).replace(/[$=\s]/g, '')) : null;
      if (!g) { const d = this.dataRegion(); g = d && d.g; }
      if (!g || g.r1 === g.r2) { XL.ui.toast('Selecione os dados (com a linha de cabeçalho) antes de criar a tabela.'); return null; }
      if (S.tables.some((t) => overlap(t.g, g))) { XL.ui.toast('Uma tabela não pode sobrepor outra tabela.'); return null; }
      this.pushUndo();
      const t = { name: S.nextTableName(), g: Object.assign({}, g), style: style || 'azul', banded: true, totalRow: false };
      for (let c = g.c1; c <= g.c2; c++) if (!S.has(g.r1, c)) S.setInput(g.r1, c, 'Coluna' + (c - g.c1 + 1));
      S.tables.push(t);
      if (!S.filter) { S.filter = { g: Object.assign({}, g), crit: {}, table: t.name }; S.applyFilter(); }
      this.render();
      this.ribbon.updateContext();
      this.ribbon.setTab('tabledesign', true);
      this.emit('table', { name: t.name, range: XL.rangeToStr(g) });
      this.emit('change', {});
      return t;
    },

    // mantém o filtro da tabela do mesmo tamanho dela (sem a linha de totais)
    syncTableFilter(t) {
      const S = this.sheet;
      if (S.filter && S.filter.table === t.name) {
        S.filter.g = Object.assign({}, t.g, { r2: S.tableDataRows(t).r2 });
        S.applyFilter();
      }
    },

    // A tabela cresce quando você digita logo abaixo ou à direita; fórmulas viram colunas calculadas.
    afterTableCommit(S, r, c, input) {
      for (const t of S.tables) {
        const g = t.g;
        if (!t.totalRow && r === g.r2 + 1 && c >= g.c1 && c <= g.c2 && input !== '') {
          g.r2++;
          for (let cc = g.c1; cc <= g.c2; cc++) {
            if (cc === c) continue;
            const above = S.cell(r - 1, cc);
            if (above && above.kind === 'formula' && !S.has(r, cc)) S.setInput(r, cc, XL.Formula.shiftFormula(above.input, 1, 0));
            if (above && !Object.keys(S.fmt(r, cc)).length) S.replaceFmt(r, cc, above.fmt);
          }
          this.syncTableFilter(t);
          this.emit('tableexpand', { name: t.name, axis: 'row', range: XL.rangeToStr(g) });
        } else if (c === g.c2 + 1 && r >= g.r1 && r <= g.r2 && input !== '') {
          g.c2++;
          if (r !== g.r1 && !S.has(g.r1, c)) S.setInput(g.r1, c, 'Coluna' + (c - g.c1 + 1));
          this.syncTableFilter(t);
          this.emit('tableexpand', { name: t.name, axis: 'col', range: XL.rangeToStr(g) });
        }
        // coluna calculada: a primeira fórmula numa coluna vazia da tabela é copiada para todas as linhas
        if (input.startsWith('=') && XL.inRange(g, r, c)) {
          const dr = S.tableDataRows(t);
          if (r >= dr.r1 && r <= dr.r2) {
            let empty = true;
            for (let rr = dr.r1; rr <= dr.r2; rr++) if (rr !== r && S.has(rr, c)) { empty = false; break; }
            if (empty && dr.r2 > dr.r1) {
              for (let rr = dr.r1; rr <= dr.r2; rr++) if (rr !== r) S.setInput(rr, c, XL.Formula.shiftFormula(input, rr - r, 0));
              this.emit('calccolumn', { name: t.name, col: XL.numToCol(c + 1) });
              XL.ui.toast('⚡ <b>Coluna calculada</b>: a fórmula foi preenchida em toda a coluna da tabela.');
            }
          }
        }
      }
    },

    tableUpdate(fn, evt) {
      const t = this.activeTable();
      if (!t) { XL.ui.toast('Clique em uma célula dentro da tabela primeiro.'); return; }
      this.pushUndo();
      fn(t);
      this.render();
      this.ribbon.setTab(this.ribbon.active, true);
      this.emit(evt || 'tablechange', { name: t.name });
      this.emit('change', {});
    },

    tableTotals(on) {
      const t = this.activeTable();
      if (!t || !!t.totalRow === on) return;
      const S = this.sheet;
      this.pushUndo();
      if (on) {
        const r = t.g.r2 + 1;
        if (r >= S.rows) return;
        for (let c = t.g.c1; c <= t.g.c2; c++) if (S.has(r, c)) { XL.ui.toast('A linha logo abaixo da tabela precisa estar vazia para a Linha de Totais.'); this.undoStack.pop(); return; }
        t.g.r2 = r;
        t.totalRow = true;
        S.setInput(r, t.g.c1, 'Total');
        const dr = S.tableDataRows(t);
        for (let c = t.g.c2; c > t.g.c1; c--) {
          let num = false;
          for (let rr = dr.r1; rr <= dr.r2; rr++) if (typeof S.value(rr, c) === 'number') { num = true; break; }
          if (num) {
            S.setInput(r, c, `=SUBTOTAL(109;${t.name}[${S.tableHeader(t, c)}])`);
            S.replaceFmt(r, c, S.fmt(dr.r2, c));
            break;
          }
        }
      } else {
        for (let c = t.g.c1; c <= t.g.c2; c++) { S.clearContent(t.g.r2, c); S.replaceFmt(t.g.r2, c, {}); }
        t.g.r2--;
        t.totalRow = false;
      }
      this.syncTableFilter(t);
      this.render();
      this.ribbon.setTab(this.ribbon.active, true);
      this.emit('tabletotals', { name: t.name, on });
      this.emit('change', {});
    },

    tableFilterBtn(on) {
      const t = this.activeTable();
      if (!t) return;
      const S = this.sheet;
      this.pushUndo();
      if (on) S.filter = { g: Object.assign({}, t.g, { r2: S.tableDataRows(t).r2 }), crit: {}, table: t.name };
      else if (S.filter && S.filter.table === t.name) S.filter = null;
      S.applyFilter();
      this.render();
      this.emit('filter', { on });
    },

    tableRename(name) {
      const t = this.activeTable();
      if (!t) return;
      name = String(name || '').trim();
      if (name === t.name) return;
      if (!/^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*$/.test(name)) { XL.ui.toast('Nome de tabela inválido: comece com uma letra e não use espaços.'); this.ribbon.setTab(this.ribbon.active, true); return; }
      if (this.wb.sheets.some((s) => s.tableByName(name))) { XL.ui.toast('Já existe uma tabela com esse nome.'); this.ribbon.setTab(this.ribbon.active, true); return; }
      this.pushUndo();
      const old = t.name;
      t.name = name;
      if (this.sheet.filter && this.sheet.filter.table === old) this.sheet.filter.table = name;
      // atualiza as fórmulas que usam Tabela[Coluna]
      for (const s of this.wb.sheets) for (const cell of s.cells.values()) {
        if (cell.kind !== 'formula') continue;
        const toks = XL.Formula.tokenize(cell.input.slice(1), true);
        let changed = false;
        const body = toks.map((tk) => {
          if (tk.type === 'sref' && tk.table && XL.normName(tk.table) === XL.normName(old)) { changed = true; return name + tk.text.slice(tk.table.length); }
          return tk.text;
        }).join('');
        if (changed) s.setFormulaText(cell, '=' + body);
      }
      this.wb.invalidateAll();
      this.render();
      this.ribbon.setTab(this.ribbon.active, true);
      this.emit('tablerename', { from: old, to: name });
      this.emit('change', {});
    },

    // Volta a ser um intervalo comum: as referências estruturadas viram referências normais
    tableConvert() {
      const t = this.activeTable();
      if (!t) return;
      const S = this.sheet;
      this.pushUndo();
      const styles = S.tableStyles();
      for (const [k, st] of styles) {
        const r = Math.floor(k / K), c = k % K;
        if (!XL.inRange(t.g, r, c)) continue;
        const f = S.fmt(r, c);
        const patch = {};
        if (st.fill && !f.fill) patch.fill = st.fill;
        if (st.color && !f.color) patch.color = st.color;
        if (st.b && f.b === undefined) patch.b = true;
        S.setFmt(r, c, patch);
      }
      const dr = S.tableDataRows(t);
      for (const [k, cell] of S.cells) {
        if (cell.kind !== 'formula') continue;
        const r = Math.floor(k / K);
        const toks = XL.Formula.tokenize(cell.input.slice(1), true);
        let changed = false;
        const body = toks.map((tk) => {
          if (tk.type !== 'sref') return tk.text;
          const mine = tk.table ? XL.normName(tk.table) === XL.normName(t.name) : XL.inRange(t.g, r, k % K);
          if (!mine) return tk.text;
          const c = S.tableColIndex(t, tk.col);
          if (c < 0) return tk.text;
          changed = true;
          const L = XL.numToCol(c + 1);
          return tk.thisRow ? L + (r + 1) : `$${L}$${dr.r1 + 1}:$${L}$${dr.r2 + 1}`;
        }).join('');
        if (changed) S.setFormulaText(cell, '=' + body);
      }
      S.tables = S.tables.filter((x) => x !== t);
      if (S.filter && S.filter.table === t.name) { S.filter = null; S.applyFilter(); }
      S.invalidate();
      this.render();
      this.ribbon.updateContext();
      this.emit('tableconvert', { name: t.name });
      this.emit('change', {});
    },

    // ================= Gráficos =================
    // Rótulos em sequência de tempo (meses, datas, anos) pedem linhas; categorias pedem colunas.
    recommendChart() {
      let g = this.grid.range();
      if (g.r1 === g.r2 && g.c1 === g.c2) { const d = this.dataRegion(); g = d && d.g; }
      if (!g) return null;
      const cats = this.sheet.chartData({ g, type: 'col' }).cats.map((c) => String(c).toLowerCase());
      const months = /^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/;
      const timeLike = cats.length >= 3 && cats.every((c) => months.test(c) || /^\d{4}$/.test(c) || /^\d{1,2}\/\d{1,2}/.test(c) || /^(t|tri)\s?\d/.test(c));
      return timeLike ? 'line' : 'col';
    },

    selectChart(id) {
      if (this.selectedChart === id) return;
      this.selectedChart = id;
      XL.Charts.draw(this.grid);
      this.ribbon.updateContext();
      if (id) this.emit('chartselect', { id });
    },
    activeChart() { return this.sheet.charts.find((c) => c.id === this.selectedChart) || null; },

    insertChart(type) {
      const S = this.sheet;
      let g = this.grid.range();
      if (g.r1 === g.r2 && g.c1 === g.c2) { const d = this.dataRegion(); g = d && d.g; }
      if (!g) { XL.ui.toast('Selecione os dados do gráfico (com os rótulos) antes de inserir.'); return; }
      const probe = { g, type };
      if (!S.chartData(probe).series.length) { XL.ui.toast('Não encontramos números na seleção para fazer o gráfico.'); return; }
      this.pushUndo();
      const col = Math.min(S.cols - 1, g.c2 + 2);
      const ch = {
        id: 'g' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36),
        type, g: Object.assign({}, g), x: this.grid.staticLeft(col), y: this.grid.rowHeads[g.r1].offsetTop,
        w: 440, h: 270, title: null, legend: true, labels: false,
      };
      // não deixa um gráfico novo cobrir outro: desce até achar espaço livre
      for (let moved = true, guard = 0; moved && guard < 20; guard++) {
        moved = false;
        for (const o of S.charts) {
          if (ch.x < o.x + o.w && o.x < ch.x + ch.w && ch.y < o.y + o.h && o.y < ch.y + ch.h) { ch.y = o.y + o.h + 12; moved = true; }
        }
      }
      S.charts.push(ch);
      this.render();
      this.selectChart(ch.id);
      this.ribbon.setTab('chartdesign', true);
      this.emit('chart', { type, range: XL.rangeToStr(g), id: ch.id });
      this.emit('change', {});
    },

    chartUpdate(fn) {
      const ch = this.activeChart();
      if (!ch) { XL.ui.toast('Clique no gráfico primeiro para selecioná-lo.'); return; }
      this.pushUndo();
      fn(ch);
      XL.Charts.draw(this.grid);
      this.ribbon.setTab(this.ribbon.active, true);
      this.emit('chartchange', { id: ch.id, type: ch.type, title: ch.title, legend: ch.legend, labels: ch.labels });
      this.emit('change', {});
    },

    deleteChart() {
      const ch = this.activeChart();
      if (!ch) return;
      this.pushUndo();
      this.sheet.charts = this.sheet.charts.filter((c) => c !== ch);
      this.selectedChart = null;
      XL.Charts.draw(this.grid);
      this.ribbon.updateContext();
      this.emit('chartdelete', { id: ch.id });
      this.emit('change', {});
    },

    // ================= Formatação condicional =================
    addCF(rule) {
      const g = this.grid.range();
      this.pushUndo();
      this.sheet.cf.unshift(Object.assign({ g: Object.assign({}, g) }, rule));
      this.render();
      this.emit('cf', Object.assign({ range: XL.rangeToStr(g) }, rule));
      this.emit('change', {});
    },
    clearCF(scope) {
      const g = this.grid.range();
      this.pushUndo();
      this.sheet.cf = scope === 'all' ? [] : this.sheet.cf.filter((r) => !overlap(r.g, g));
      this.render();
      this.emit('cfclear', { scope });
      this.emit('change', {});
    },

    // ================= Validação de dados =================
    addDV(rule) {
      const g = this.grid.range();
      this.pushUndo();
      this.sheet.dv = this.sheet.dv.filter((r) => !(r.g.r1 >= g.r1 && r.g.r2 <= g.r2 && r.g.c1 >= g.c1 && r.g.c2 <= g.c2));
      this.sheet.dv.push(Object.assign({ g: Object.assign({}, g) }, rule));
      this.render();
      this.emit('dv', Object.assign({ range: XL.rangeToStr(g) }, rule));
      this.emit('change', {});
    },
    clearDV() {
      const g = this.grid.range();
      this.pushUndo();
      this.sheet.dv = this.sheet.dv.filter((r) => !overlap(r.g, g));
      this.render();
      this.emit('dvclear', {});
      this.emit('change', {});
    },
    openDvList(r, c, anchor) {
      const rule = this.sheet.validationAt(r, c);
      if (!rule || rule.type !== 'list') return;
      const vals = this.sheet.listValues(rule);
      // como no Excel, a lista abre logo abaixo da célula
      const td = this.grid.tds[r] && this.grid.tds[r][c];
      XL.ui.menu(td || anchor, vals.map((v) => ({
        label: v, data: { dv: v },
        onClick: () => {
          if (this.grid.sel.ar !== r || this.grid.sel.ac !== c) this.grid.setSel(r, c);
          const res = this.commitInput(r, c, v);
          if (res.ok) this.emit('dvpick', { ref: XL.toRef(r, c), value: v });
          this.grid.focus();
        },
      })));
      this.emit('dvopen', { ref: XL.toRef(r, c) });
    },
  });
})();
