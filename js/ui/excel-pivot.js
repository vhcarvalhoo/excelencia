/* Extensão do ExcelApp (Módulo 8): Tabelas Dinâmicas — criar, mover campos, configurações de valor,
   agrupar, campo calculado, atualizar, filtros/classificação, Segmentação de Dados, Linha do Tempo e Gráfico Dinâmico. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const proto = XL.ExcelApp.prototype;
  const baseExtra = proto.execExtra;
  const baseCommit = proto.commitInput;
  const baseInsertChart = proto.insertChart;
  const norm = (s) => XL.normName(String(s));
  const uid = (p) => p + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
  const AREAS = { filters: 'Filtros', cols: 'Colunas', rows: 'Linhas', vals: 'Valores' };
  XL.PIVOT_AREAS = AREAS;
  const enterSubmits = (inp) => inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const b = inp.closest('.ui-modal').querySelector('.btn-primary'); if (b) b.click(); } });

  Object.assign(proto, {
    execExtra(cmd, arg) {
      switch (cmd) {
        case 'pivot-create': return this.pivotCreateDialog();
        case 'pivot-refresh': return this.pivotRefresh(false);
        case 'pivot-refresh-all': case 'refresh': return this.pivotRefresh(true);
        case 'pivot-settings': return this.pivotSettingsDialog(arg);
        case 'pivot-group': return this.pivotGroupDialog();
        case 'pivot-ungroup': return this.pivotUngroup();
        case 'pivot-calc': return this.pivotCalcDialog();
        case 'pivot-slicer': return this.pivotSlicerDialog();
        case 'pivot-timeline': return this.pivotTimelineDialog();
        case 'pivot-chart': return this.pivotChart(arg || 'col');
        case 'pivot-fieldlist': this.pivotPaneOff = !this.pivotPaneOff; this.updatePivotPane(true); this.ribbon.sync(); return;
        case 'pivot-rename': return this.pivotRename(arg);
        case 'pivot-subtotals': return this.pivotUpdate((p) => { p.subtotals = arg !== 'off'; }, 'pivotlayout');
        case 'pivot-grand': return this.pivotUpdate((p) => { p.grandRows = arg === 'on' || arg === 'rows'; p.grandCols = arg === 'on' || arg === 'cols'; }, 'pivotlayout');
        case 'pivot-style': return this.pivotUpdate((p) => { p.style = arg; }, 'pivotstyle');
      }
      return baseExtra.call(this, cmd, arg);
    },

    // ---------- Localizar a Tabela Dinâmica ----------
    activePivot() {
      if (!this.grid) return null;
      const { ar, ac } = this.grid.sel;
      return this.sheet.pivotAt ? this.sheet.pivotAt(ar, ac) : null;
    },
    // a Tabela Dinâmica com que estamos trabalhando: a da célula ativa ou a última usada nesta planilha
    workPivot() {
      const p = this.activePivot();
      if (p) return p;
      const S = this.sheet;
      return (S.pivots || []).find((x) => x.id === this.lastPivot) || null;
    },

    // Células da Tabela Dinâmica não podem ser digitadas
    commitInput(r, c, text, sheet) {
      const S = sheet || this.sheet;
      if (S.pivotAt && S.pivotAt(r, c) && String(text) !== '') {
        return { ok: false, msg: 'Não é possível alterar esta parte de uma Tabela Dinâmica. Para mudar o resultado, use a <b>Lista de Campos</b> — ou altere os dados de origem e clique em <b>Atualizar</b>.' };
      }
      return baseCommit.call(this, r, c, text, sheet);
    },

    // ================= Criar =================
    pivotCreateDialog() {
      const S = this.sheet;
      const { ar, ac } = this.grid.sel;
      if (S.pivotAt(ar, ac)) { XL.ui.toast('Esta célula já faz parte de uma Tabela Dinâmica.'); return; }
      const t = S.tableAt(ar, ac);
      let src = '';
      if (t) src = t.name;
      else {
        const d = this.dataRegion();
        if (!d) { XL.ui.toast('Clique em uma célula dos dados (com cabeçalhos) antes de criar a Tabela Dinâmica.'); return; }
        const g = d.g;
        src = XL.quoteSheet(S.name) + '!$' + XL.numToCol(g.c1 + 1) + '$' + (g.r1 + 1) + ':$' + XL.numToCol(g.c2 + 1) + '$' + (g.r2 + 1);
        this.grid.setRange(g);
      }
      const inp = h('input.fd-search', { type: 'text', value: src, dataset: { role: 'pv-src' } });
      const rNew = h('input', { type: 'radio', name: 'pvdest', value: 'new', dataset: { role: 'pv-new' } });
      const rOld = h('input', { type: 'radio', name: 'pvdest', value: 'old', dataset: { role: 'pv-existing' } });
      rNew.checked = true;
      const loc = h('input.fd-search', { type: 'text', value: '', disabled: true, dataset: { role: 'pv-loc' }, placeholder: 'Ex.: H3' });
      const sync = () => { loc.disabled = !rOld.checked; if (rOld.checked) loc.focus(); };
      rNew.addEventListener('change', sync); rOld.addEventListener('change', sync);
      enterSubmits(inp); enterSubmits(loc);
      XL.ui.modal({
        title: 'Criar Tabela Dinâmica', width: 440,
        body: h('div.pvd',
          h('div.dv-sec', 'Escolha os dados que deseja analisar'),
          h('label.fd-lbl', 'Tabela/Intervalo:'), inp,
          h('div.dv-sec', 'Escolha onde deseja que o relatório da Tabela Dinâmica seja colocado'),
          h('label.sd-check', rNew, ' Nova Planilha'),
          h('label.sd-check', rOld, ' Planilha Existente'),
          h('label.fd-lbl', 'Local:'), loc),
        buttons: [{ label: 'OK', primary: true, onClick: () => { if (!this.createPivot(inp.value, rOld.checked ? loc.value : null)) return false; } }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => inp.select(), 0);
    },

    parsePivotSource(text) {
      const t = String(text || '').trim().replace(/^=/, '');
      for (const s of this.wb.sheets) { const tb = s.tableByName(t); if (tb) return { table: tb.name }; }
      const i = t.lastIndexOf('!');
      const sheetName = i > 0 ? t.slice(0, i).replace(/^'|'$/g, '') : this.sheet.name;
      const S = this.wb.byName(sheetName);
      const g = XL.parseRange(t.slice(i + 1).replace(/\$/g, ''));
      if (!S || !g) return null;
      return { sheet: S.name, range: XL.rangeToStr(g) };
    },

    createPivot(srcText, where) {
      const src = this.parsePivotSource(srcText);
      if (!src) { XL.ui.toast('Referência da fonte de dados inválida. Use o nome de uma tabela (ex.: <b>Vendas</b>) ou um intervalo (ex.: <b>Dados!$A$1:$F$30</b>).'); return null; }
      if (src.range && XL.parseRange(src.range).r1 === XL.parseRange(src.range).r2) { XL.ui.toast('A fonte precisa ter a linha de cabeçalho e pelo menos uma linha de dados.'); return null; }
      let at = { r: 2, c: 0 };
      let target;
      if (where) {
        const t = String(where).trim().replace(/\$/g, '');
        const i = t.lastIndexOf('!');
        target = i > 0 ? this.wb.byName(t.slice(0, i).replace(/^'|'$/g, '')) : this.sheet;
        const pr = XL.parseRef(t.slice(i + 1));
        if (!target || !pr) { XL.ui.toast('Local inválido. Digite uma célula, como <b>H3</b>.'); return null; }
        at = { r: pr.r, c: pr.c };
      }
      this.pushUndo();
      if (!target) {
        const idx = this.wb.activeIndex;
        target = this.wb.addSheet(this.wb.nextName(), idx);
        this.wb.activeIndex = idx;
      } else this.wb.activeIndex = this.wb.sheets.indexOf(target);
      const p = {
        id: uid('pv'), name: target.nextPivotName(), src, at, rows: [], cols: [], vals: [], filters: [],
        groups: {}, calc: [], hidden: {}, sort: {}, style: 'azul', subtotals: true, grandRows: true, grandCols: true,
      };
      target.pivots.push(p);
      target.renderPivot(p);
      this.lastPivot = p.id;
      this.pivotPaneOff = false;
      this.afterWorkbookChange();
      this.grid.setSel(at.r, at.c);
      this.render();
      this.ribbon.updateContext();
      this.ribbon.setTab('pivotanalyze', true);
      this.updatePivotPane(true);
      this.emit('sheetadd', { name: target.name });
      this.emit('pivot', { name: p.name, sheet: target.name });
      this.emit('change', {});
      return p;
    },

    // ================= Alterações =================
    pivotUpdate(fn, evt, data, pv) {
      const p = pv || this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return null; }
      this.pushUndo();
      if (fn(p) === false) { this.undoStack.pop(); return null; }
      this.lastPivot = p.id;
      this.sheet.renderPivot(p);
      this.grid.applyColWidths();
      this.render();
      if (['pivotanalyze', 'pivotdesign'].includes(this.ribbon.active)) this.ribbon.setTab(this.ribbon.active, true);
      this.updatePivotPane(true);
      this.emit(evt || 'pivotchange', Object.assign({ name: p.name }, data || {}));
      this.emit('change', {});
      return p;
    },

    // move um campo para uma área (filters | cols | rows | vals) ou remove (area null)
    pivotMoveField(field, area, opts = {}) {
      return this.pivotUpdate((p) => {
        const fd = XL.pivotFields(this.wb, p).find((f) => norm(f.name) === norm(field));
        if (!fd) return false;
        const from = opts.from; // { area, i } quando o campo vem de uma área
        if (from && from.area === 'vals') p.vals.splice(from.i, 1);
        else if (from) p[from.area] = p[from.area].filter((x) => (from.area === 'filters' ? norm(x.field) : norm(x)) !== norm(field));
        if (!area) return;
        if (area === 'vals') {
          const fn = fd.type === 'num' ? 'sum' : 'count';
          const base = XL.PIVOT_FN[fn] + ' de ' + fd.name;
          let name = base, n = 2;
          while (p.vals.some((v) => norm(XL.pivotValName(v)) === norm(name))) name = base + n++;
          const v = { field: fd.name, fn, name };
          if (opts.index !== undefined) p.vals.splice(opts.index, 0, v); else p.vals.push(v);
          return;
        }
        // um campo fica em só uma das áreas Filtros/Colunas/Linhas
        p.rows = p.rows.filter((x) => norm(x) !== norm(fd.name));
        p.cols = p.cols.filter((x) => norm(x) !== norm(fd.name));
        p.filters = p.filters.filter((x) => norm(x.field) !== norm(fd.name));
        if (area === 'cols' && p.cols.length) {
          XL.ui.toast('Neste simulador, a área <b>Colunas</b> aceita um campo por vez — o campo anterior voltou para a lista.');
          p.cols = [];
        }
        const item = area === 'filters' ? { field: fd.name, sel: null } : fd.name;
        const list = p[area];
        if (opts.index !== undefined && opts.index <= list.length) list.splice(opts.index, 0, item); else list.push(item);
      }, 'pivotfield', { field, area: area || 'none' });
    },
    // marcar a caixa na lista: texto/data vão para Linhas, números para Valores (como no Excel)
    pivotToggleField(field, on) {
      const p = this.workPivot();
      if (!p) return;
      const fd = XL.pivotFields(this.wb, p).find((f) => norm(f.name) === norm(field));
      if (!fd) return;
      if (on) return this.pivotMoveField(fd.name, fd.type === 'num' || fd.calc ? 'vals' : 'rows');
      return this.pivotUpdate((q) => {
        q.rows = q.rows.filter((x) => norm(x) !== norm(field));
        q.cols = q.cols.filter((x) => norm(x) !== norm(field));
        q.filters = q.filters.filter((x) => norm(x.field) !== norm(field));
        q.vals = q.vals.filter((x) => norm(x.field) !== norm(field));
      }, 'pivotfield', { field, area: 'none' });
    },

    pivotRefresh(all) {
      const p = this.workPivot();
      if (!all && !p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica para atualizá-la.'); return; }
      const any = this.wb.sheets.some((s) => (s.pivots || []).length);
      if (!any) { XL.ui.toast('Não há Tabelas Dinâmicas para atualizar nesta pasta de trabalho.'); return; }
      this.pushUndo();
      if (all) XL.refreshPivots(this.wb); else this.sheet.renderPivot(p);
      this.grid.applyColWidths();
      this.render();
      this.updatePivotPane(true);
      XL.ui.toast(all ? '🔄 Todas as Tabelas Dinâmicas foram <b>atualizadas</b>.' : '🔄 Tabela Dinâmica <b>atualizada</b> com os dados atuais.');
      this.emit('pivotrefresh', { all: !!all });
      this.emit('change', {});
    },

    pivotRename(name) {
      const p = this.workPivot();
      if (!p) return;
      name = String(name || '').trim();
      if (!name || name === p.name) { this.ribbon.setTab(this.ribbon.active, true); return; }
      if (this.wb.sheets.some((s) => (s.pivots || []).some((x) => x !== p && norm(x.name) === norm(name)))) { XL.ui.toast('Já existe uma Tabela Dinâmica com esse nome.'); this.ribbon.setTab(this.ribbon.active, true); return; }
      this.pivotUpdate((q) => { q.name = name; }, 'pivotrename', { to: name });
    },

    // ================= Configurações do Campo de Valor =================
    pivotSettingsDialog(index) {
      const p = this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return; }
      let i = index !== undefined && index !== '' ? Number(index) : p.vcols ? p.vcols[this.grid.sel.ac] : undefined;
      if (i === undefined || !p.vals[i]) i = p.vals.length === 1 ? 0 : undefined;
      if (i === undefined) { XL.ui.toast('Clique numa célula de <b>valores</b> (ou use a setinha do campo na área Valores).'); return; }
      const v = p.vals[i];
      const auto = (fn) => XL.PIVOT_FN[fn] + ' de ' + v.field;
      const name = h('input.fd-search', { type: 'text', value: XL.pivotValName(v), dataset: { role: 'pv-vname' } });
      let touched = norm(name.value) !== norm(auto(v.fn || 'sum'));
      name.addEventListener('input', () => { touched = true; });
      const fn = h('select.fd-cat.pv-list', { size: 5, dataset: { role: 'pv-fn' } }, Object.keys(XL.PIVOT_FN).map((k) => h('option', { value: k }, XL.PIVOT_FN[k])));
      fn.value = v.fn || 'sum';
      fn.addEventListener('change', () => { if (!touched) name.value = auto(fn.value); });
      const show = h('select.fd-cat', { dataset: { role: 'pv-show' } },
        [['normal', 'Sem Cálculo'], ['pctTotal', '% do Total Geral'], ['pctCol', '% do Total da Coluna'], ['pctRow', '% do Total da Linha']].map(([k, l]) => h('option', { value: k }, l)));
      show.value = v.show || 'normal';
      const nf = h('select.fd-cat', { dataset: { role: 'pv-numfmt' } },
        [['', 'Geral'], ['numero', 'Número'], ['moeda', 'Moeda (R$)'], ['contabil', 'Contábil'], ['porcentagem', 'Porcentagem']].map(([k, l]) => h('option', { value: k }, l)));
      nf.value = v.numFmt || '';
      enterSubmits(name);
      XL.ui.modal({
        title: 'Configurações do Campo de Valor', width: 440,
        body: h('div.pvd',
          h('p.cf-q', 'Nome da Origem: ', h('b', v.field)),
          h('label.fd-lbl', 'Nome Personalizado:'), name,
          h('div.dv-sec', 'Resumir Valores por'),
          h('p.cd-note', 'Escolha o tipo de cálculo que deseja usar para resumir os dados do campo selecionado:'), fn,
          h('div.dv-sec', 'Mostrar Valores como'), show,
          h('div.dv-sec', 'Formato do Número'), nf),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const nm = name.value.trim() || auto(fn.value);
            if (p.vals.some((x, j) => j !== i && norm(XL.pivotValName(x)) === norm(nm))) { XL.ui.toast('Já existe um campo de valor com esse nome.'); return false; }
            this.pivotUpdate((q) => {
              const x = q.vals[i];
              x.fn = fn.value; x.name = nm; x.show = show.value; x.numFmt = nf.value || undefined; x.dec = nf.value === 'porcentagem' ? 1 : undefined;
            }, 'pivotsettings', { field: v.field, fn: fn.value, show: show.value, numFmt: nf.value }, p);
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },

    // ================= Agrupar =================
    pivotGroupField() {
      const p = this.workPivot();
      if (!p) return null;
      const f = p.rmap && p.rmap[this.grid.sel.ar];
      if (f && this.grid.sel.ac === p.at.c) return f;
      return null;
    },
    pivotGroupDialog() {
      const p = this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return; }
      const field = this.pivotGroupField();
      if (!field) { XL.ui.toast('Clique em um <b>rótulo de linha</b> (uma data ou um número) para agrupar.'); return; }
      const fd = XL.pivotFields(this.wb, p).find((f) => norm(f.name) === norm(field));
      if (!fd || fd.virtual || (fd.type !== 'date' && fd.type !== 'num') || (p.groups[fd.name] && p.groups[fd.name].by)) {
        XL.ui.toast('Não é possível agrupar essa seleção. O Agrupar funciona com campos de <b>data</b> ou de <b>números</b>.');
        return;
      }
      if (fd.type === 'date') {
        const cur = (p.groups[fd.name] && p.groups[fd.name].by) || ['months'];
        const units = [['months', 'Meses'], ['quarters', 'Trimestres'], ['years', 'Anos']];
        const boxes = units.map(([u, l]) => { const b = h('input', { type: 'checkbox', dataset: { unit: u } }); b.checked = cur.includes(u); return h('label.sd-check.pv-unit', b, ' ' + l); });
        XL.ui.modal({
          title: 'Agrupamento', width: 360,
          body: h('div.pvd', h('p.cf-q', 'Campo: ', h('b', fd.name)), h('div.dv-sec', 'Por'), h('div.pv-units', boxes),
            h('p.cd-note', 'Escolha um ou mais períodos. Com Meses e Anos, o Excel cria o campo "Anos" e separa os meses de cada ano.')),
          buttons: [{
            label: 'OK', primary: true,
            onClick: () => {
              const by = [...document.querySelectorAll('.pv-units input:checked')].map((b) => b.dataset.unit);
              if (!by.length) { XL.ui.toast('Escolha pelo menos um período.'); return false; }
              this.pivotUpdate((q) => {
                q.groups[fd.name] = { by };
                const i = q.rows.findIndex((x) => norm(x) === norm(fd.name));
                const extra = ['years', 'quarters'].filter((u) => by.includes(u) && u !== (['months', 'quarters', 'years'].find((x) => by.includes(x))));
                const names = extra.map((u) => (u === 'years' ? 'Anos' : 'Trimestres'));
                if (i >= 0) q.rows.splice(i, 0, ...names.filter((n) => !q.rows.includes(n)));
                q.hidden[fd.name] = [];
              }, 'pivotgroup', { field: fd.name, by }, p);
            },
          }, { label: 'Cancelar' }],
          onClose: () => this.grid.focus(),
        });
        return;
      }
      const items = XL.pivotItems(this.wb, p, fd.name).map((x) => x.key).filter((x) => typeof x === 'number');
      const lo = Math.floor(Math.min(...items)), hi = Math.ceil(Math.max(...items));
      const a = h('input.fd-search', { type: 'text', value: String(lo), dataset: { role: 'pv-gstart' } });
      const b = h('input.fd-search', { type: 'text', value: String(hi), dataset: { role: 'pv-gend' } });
      const s = h('input.fd-search', { type: 'text', value: '10', dataset: { role: 'pv-gby' } });
      [a, b, s].forEach(enterSubmits);
      XL.ui.modal({
        title: 'Agrupamento', width: 360,
        body: h('div.pvd', h('p.cf-q', 'Campo: ', h('b', fd.name)), h('label.fd-lbl', 'Começando em:'), a, h('label.fd-lbl', 'Terminando em:'), b, h('label.fd-lbl', 'Por:'), s),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const n = (x) => XL.parseNumber ? XL.parseNumber(x.value) : Number(String(x.value).replace(',', '.'));
            const st = n(a), en = n(b), by = n(s);
            if (![st, en, by].every((x) => typeof x === 'number' && isFinite(x)) || by <= 0 || en < st) { XL.ui.toast('Confira os números do agrupamento.'); return false; }
            this.pivotUpdate((q) => { q.groups[fd.name] = { start: st, end: en, step: by }; q.hidden[fd.name] = []; }, 'pivotgroup', { field: fd.name, step: by }, p);
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },
    pivotUngroup() {
      const p = this.workPivot();
      if (!p) return;
      let field = this.pivotGroupField();
      if (field && !p.groups[field]) {
        const v = XL.pivotFields(this.wb, p).find((f) => norm(f.name) === norm(field));
        field = v && v.virtual ? v.virtual.base : null;
      }
      if (!field || !p.groups[field]) { XL.ui.toast('Clique em um rótulo de um campo agrupado para desagrupar.'); return; }
      this.pivotUpdate((q) => {
        delete q.groups[field];
        q.rows = q.rows.filter((x) => x !== 'Anos' && x !== 'Trimestres');
        q.cols = q.cols.filter((x) => x !== 'Anos' && x !== 'Trimestres');
        q.filters = q.filters.filter((x) => x.field !== 'Anos' && x.field !== 'Trimestres');
      }, 'pivotgroup', { field, by: null });
    },

    // ================= Campo Calculado =================
    pivotCalcDialog() {
      const p = this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return; }
      const fields = XL.pivotFields(this.wb, p).filter((f) => !f.virtual && !f.calc);
      let n = 1;
      while (fields.some((f) => norm(f.name) === norm('Campo' + n)) || (p.calc || []).some((c) => norm(c.name) === norm('Campo' + n))) n++;
      const name = h('input.fd-search', { type: 'text', value: 'Campo' + n, dataset: { role: 'pv-cname' } });
      const formula = h('input.fd-search', { type: 'text', value: '= 0', dataset: { role: 'pv-cformula' } });
      const list = h('select.fd-cat.pv-list', { size: Math.min(7, fields.length), dataset: { role: 'pv-cfields' } }, fields.map((f) => h('option', { value: f.name }, f.name)));
      const insert = () => {
        if (!list.value) return;
        const nm = /[^A-Za-zÀ-ÿ0-9_]/.test(list.value) ? "'" + list.value + "'" : list.value;
        let v = formula.value;
        if (v.trim() === '= 0' || v.trim() === '=0') v = '=';
        formula.value = v + nm;
        formula.focus();
      };
      list.addEventListener('dblclick', insert);
      enterSubmits(name); enterSubmits(formula);
      XL.ui.modal({
        title: 'Inserir Campo Calculado', width: 460,
        body: h('div.pvd',
          h('label.fd-lbl', 'Nome:'), name,
          h('label.fd-lbl', 'Fórmula:'), formula,
          h('div.dv-sec', 'Campos'), list,
          h('button.btn.btn-sm.pv-ins', { type: 'button', dataset: { role: 'pv-cinsert' }, onclick: insert }, 'Inserir Campo'),
          h('p.cd-note', 'A conta é feita sobre as SOMAS de cada grupo. Ex.: =Valor-Custo dá o lucro de cada linha da Tabela Dinâmica.')),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const nm = name.value.trim();
            let f = formula.value.trim();
            if (!nm || /[\\/?*[\]:']/.test(nm)) { XL.ui.toast('Digite um nome válido para o campo.'); return false; }
            if (fields.some((x) => norm(x.name) === norm(nm))) { XL.ui.toast('Esse nome já é de um campo da fonte de dados.'); return false; }
            if (f[0] !== '=') f = '=' + f;
            try { XL.pivotCalcCheck(this.wb, p, f); } catch (e) { XL.ui.toast('Há um problema com esta fórmula: ' + XL.escapeHtml(e.message)); return false; }
            this.pivotUpdate((q) => {
              q.calc = (q.calc || []).filter((c) => norm(c.name) !== norm(nm));
              q.calc.push({ name: nm, formula: f });
              if (!q.vals.some((v) => norm(v.field) === norm(nm))) q.vals.push({ field: nm, fn: 'sum', name: 'Soma de ' + nm });
            }, 'pivotcalc', { name: nm, formula: f }, p);
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => name.select(), 0);
    },

    // ================= Menus das setinhas da Tabela Dinâmica =================
    drawPivotButtons() {
      const g = this.grid;
      for (const p of this.sheet.pivots || []) {
        for (const bt of p.btns || []) {
          const td = g.tds[bt.r] && g.tds[bt.r][bt.c];
          if (!td) continue;
          const b = h('button.xl-filterbtn' + (bt.on ? '.on' : ''), {
            type: 'button', title: bt.kind === 'filter' ? 'Filtrar a Tabela Dinâmica por ' + bt.field : 'Filtrar / classificar: ' + bt.field,
            dataset: { pvbtn: bt.kind, field: bt.field },
            html: bt.on
              ? '<svg viewBox="0 0 12 12" width="11" height="11"><path d="M1.5 2h9L7 6.2V10l-2 1V6.2z" fill="currentColor"/></svg>'
              : '<svg viewBox="0 0 12 12" width="10" height="10"><path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>',
          });
          b.style.left = td.offsetLeft + td.offsetWidth - 19 + 'px';
          b.style.top = td.offsetTop + td.offsetHeight - 19 + 'px';
          b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
          b.addEventListener('click', (e) => { e.stopPropagation(); if (g.editing && !g.commit()) return; this.openPivotMenu(p, bt, b); });
          g.filterLayer.appendChild(b);
        }
      }
    },

    openPivotMenu(p, bt, anchor) {
      this.lastPivot = p.id;
      const items = XL.pivotItems(this.wb, p, bt.field);
      const field = bt.field;
      const isFilter = bt.kind === 'filter';
      const f = isFilter ? p.filters.find((x) => norm(x.field) === norm(field)) : null;
      const hidden = new Set(isFilter ? [] : (p.hidden[field] || []));
      const selected = (lbl) => (isFilter ? !f.sel || f.sel.includes(lbl) : !hidden.has(lbl));
      const el = h('div.ui-menu.pv-menu', { dataset: { field } });
      if (!isFilter) {
        const valName = p.vals[0] ? XL.pivotValName(p.vals[0]) : null;
        const sortBtn = (label, key) => h('button.ui-menu-item', {
          type: 'button', dataset: { pvsort: key }, onmousedown: (e) => e.preventDefault(),
          onclick: () => { XL.ui.closePopup(); this.pivotSort(p, field, key); },
        }, h('span.ui-menu-ico'), h('span.ui-menu-lbl', label));
        el.append(h('div.ui-menu-header', 'Campo: ' + field),
          sortBtn('Classificar de A a Z', 'asc'), sortBtn('Classificar de Z a A', 'desc'));
        if (valName) el.append(sortBtn('Classificar do Maior para o Menor (' + valName + ')', 'vdesc'), sortBtn('Classificar do Menor para o Maior (' + valName + ')', 'vasc'));
        el.append(h('div.ui-menu-sep'));
      } else el.append(h('div.ui-menu-header', 'Filtrar por ' + field));
      const all = h('input', { type: 'checkbox', dataset: { item: '*' } });
      const boxes = items.map((it) => { const b = h('input', { type: 'checkbox', dataset: { item: it.label } }); b.checked = selected(it.label); return b; });
      const syncAll = () => { all.checked = boxes.every((b) => b.checked); all.indeterminate = !all.checked && boxes.some((b) => b.checked); };
      all.addEventListener('change', () => { boxes.forEach((b) => { b.checked = all.checked; }); });
      boxes.forEach((b) => b.addEventListener('change', syncAll));
      syncAll();
      el.append(h('div.pv-items',
        h('label.pv-item', all, isFilter ? ' (Tudo)' : ' (Selecionar Tudo)'),
        items.map((it, i) => h('label.pv-item', boxes[i], ' ' + it.label))));
      el.append(h('div.pv-menu-foot',
        h('button.btn.btn-primary.btn-sm', {
          type: 'button', dataset: { role: 'pv-ok' }, onclick: () => {
            const sel = items.filter((it, i) => boxes[i].checked).map((it) => it.label);
            if (!sel.length) { XL.ui.toast('Escolha pelo menos um item.'); return; }
            XL.ui.closePopup();
            this.pivotFilter(p, field, isFilter, sel.length === items.length ? null : sel, items);
          },
        }, 'OK'),
        h('button.btn.btn-sm', { type: 'button', onclick: () => XL.ui.closePopup() }, 'Cancelar')));
      XL.ui.popup(anchor, el);
      this.emit('pivotmenu', { field });
    },
    pivotSort(p, field, key) {
      this.pivotUpdate((q) => {
        q.sort = q.sort || {};
        q.sort[field] = key === 'asc' || key === 'desc' ? { by: 'label', dir: key } : { by: 'value', dir: key === 'vdesc' ? 'desc' : 'asc', val: 0 };
      }, 'pivotsort', { field, key }, p);
    },
    pivotFilter(p, field, isFilter, sel, items) {
      this.pivotUpdate((q) => {
        if (isFilter) { const f = q.filters.find((x) => norm(x.field) === norm(field)); if (f) f.sel = sel; }
        else q.hidden[field] = sel ? items.filter((it) => !sel.includes(it.label)).map((it) => it.label) : [];
      }, 'pivotfilter', { field, sel }, p);
    },

    // ================= Segmentação de Dados e Linha do Tempo =================
    pivotSlicerDialog(timeline) {
      const p = this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return; }
      let fields = XL.pivotFields(this.wb, p).filter((f) => !f.calc && !f.virtual);
      if (timeline) fields = fields.filter((f) => f.type === 'date');
      if (!fields.length) { XL.ui.toast('A fonte de dados não tem nenhuma coluna de <b>datas</b> para a Linha do Tempo.'); return; }
      const boxes = fields.map((f) => h('input', { type: 'checkbox', dataset: { field: f.name } }));
      XL.ui.modal({
        title: timeline ? 'Inserir Linhas do Tempo' : 'Inserir Segmentação de Dados', width: 340,
        body: h('div.pvd.pv-pick', boxes.map((b, i) => h('label.sd-check', b, ' ' + fields[i].name))),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const pick = fields.filter((f, i) => boxes[i].checked).map((f) => f.name);
            if (!pick.length) { XL.ui.toast('Marque pelo menos um campo.'); return false; }
            this.insertSlicers(p, pick, timeline ? 'timeline' : 'slicer');
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },
    pivotTimelineDialog() { return this.pivotSlicerDialog(true); },

    insertSlicers(p, fields, kind) {
      const S = this.sheet;
      this.pushUndo();
      const g = this.grid;
      const right = p.out ? p.out.c2 + 2 : p.at.c + 4;
      const hit = (x, y, o) => S.charts.concat(S.slicers).find((c) => x < c.x + c.w && c.x < x + o.w && y < c.y + c.h && c.y < y + o.h);
      // lugares possíveis: à direita da Tabela Dinâmica ou abaixo de todas elas (num painel, à esquerda dos gráficos)
      const first = kind === 'timeline' ? { w: 420, h: 150 } : { w: 190, h: 230 };
      const cands = [{ x: g.staticLeft(Math.min(S.cols - 1, right)), y: g.rowHeads[Math.max(0, p.at.r)].offsetTop }];
      const lowest = Math.max(0, ...S.pivots.map((q) => (q.out ? q.out.r2 : 0)));
      if (lowest + 2 < S.rows) cands.push({ x: g.staticLeft(p.at.c), y: g.rowHeads[lowest + 2].offsetTop });
      const spot = cands.find((cd) => !hit(cd.x, cd.y, first)) || cands[0];
      const x0 = spot.x, y0 = spot.y;
      let x = x0;
      for (const field of fields) {
        const o = kind === 'timeline' ? { w: 420, h: 150 } : { w: 190, h: 230 };
        let y = y0;
        // não cobre gráficos nem outras segmentações: desce até achar espaço
        for (let guard = 0, c = hit(x, y, o); c && guard < 30; guard++, c = hit(x, y, o)) y = c.y + c.h + 12;
        S.slicers.push(Object.assign({ id: uid('sl'), kind, field, pivots: [p.id], sel: null, range: null, x, y, multi: false }, o));
        x += o.w + 14;
      }
      this.render();
      this.emit(kind, { fields });
      this.emit('change', {});
    },

    slicerPick(sl, label, add) {
      const p = this.sheet.pivotById(sl.pivots[0]);
      if (!p) return;
      const all = XL.pivotItems(this.wb, p, sl.field).map((x) => x.label);
      let sel = sl.sel ? sl.sel.slice() : null;
      if (add || sl.multi) {
        if (!sel) sel = all.filter((x) => x !== label);
        else if (sel.includes(label)) sel = sel.filter((x) => x !== label);
        else sel.push(label);
        if (!sel.length) sel = null;
      } else sel = [label];
      if (sel && sel.length === all.length) sel = null;
      this.slicerSet(sl, (s) => { s.sel = sel; }, 'slicerpick', { field: sl.field, sel });
    },
    slicerSet(sl, fn, evt, data) {
      this.pushUndo();
      fn(sl);
      for (const id of sl.pivots) { const p = this.sheet.pivotById(id); if (p) this.sheet.renderPivot(p); }
      this.grid.applyColWidths();
      this.render();
      this.updatePivotPane(true);
      this.emit(evt, data);
      this.emit('change', {});
    },
    // Conexões de Relatório: uma segmentação pode filtrar várias Tabelas Dinâmicas da mesma fonte
    slicerConnectionsDialog(sl) {
      const S = this.sheet;
      const base = S.pivotById(sl.pivots[0]);
      const same = S.pivots.filter((p) => !base || JSON.stringify(p.src) === JSON.stringify(base.src));
      const boxes = same.map((p) => { const b = h('input', { type: 'checkbox', dataset: { pivot: p.name } }); b.checked = sl.pivots.includes(p.id); return b; });
      XL.ui.modal({
        title: `Conexões de Relatório (${sl.field})`, width: 380,
        body: h('div.pvd.pv-pick', h('p.cf-q', 'Marque as Tabelas Dinâmicas que esta segmentação deve filtrar:'),
          same.map((p, i) => h('label.sd-check', boxes[i], ' ' + p.name + '  ', h('span.pv-tag', '(' + S.name + ')')))),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const ids = same.filter((p, i) => boxes[i].checked).map((p) => p.id);
            if (!ids.length) { XL.ui.toast('Marque pelo menos uma Tabela Dinâmica.'); return false; }
            this.pushUndo();
            sl.pivots = ids;
            for (const p of S.pivots) S.renderPivot(p);
            this.grid.applyColWidths();
            this.render();
            this.emit('slicerconn', { field: sl.field, pivots: same.filter((p) => ids.includes(p.id)).map((p) => p.name) });
            this.emit('change', {});
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },

    slicerDelete(sl) {
      this.pushUndo();
      this.sheet.slicers = this.sheet.slicers.filter((x) => x !== sl);
      for (const id of sl.pivots) { const p = this.sheet.pivotById(id); if (p) this.sheet.renderPivot(p); }
      this.render();
      this.emit('slicerdelete', { field: sl.field });
      this.emit('change', {});
    },

    // ================= Gráfico Dinâmico =================
    pivotChart(type) {
      const p = this.workPivot();
      if (!p) { XL.ui.toast('Clique em uma célula da Tabela Dinâmica primeiro.'); return; }
      if (!p.vals.length) { XL.ui.toast('Coloque pelo menos um campo em <b>Valores</b> antes de criar o Gráfico Dinâmico.'); return; }
      const S = this.sheet;
      this.pushUndo();
      const col = Math.min(S.cols - 1, (p.out ? p.out.c2 : p.at.c) + 2);
      const ch = {
        id: uid('g'), type, pivot: p.id, g: Object.assign({}, p.out), x: this.grid.staticLeft(col), y: this.grid.rowHeads[Math.max(0, p.at.r)].offsetTop,
        w: 440, h: 270, title: null, legend: true, labels: false,
      };
      for (let moved = true, guard = 0; moved && guard < 20; guard++) {
        moved = false;
        for (const o of S.charts.concat(S.slicers)) if (ch.x < o.x + o.w && o.x < ch.x + ch.w && ch.y < o.y + o.h && o.y < ch.y + ch.h) { ch.y = o.y + o.h + 12; moved = true; }
      }
      S.charts.push(ch);
      this.render();
      this.selectChart(ch.id);
      this.ribbon.setTab('chartdesign', true);
      this.emit('pivotchart', { type, id: ch.id, name: p.name });
      this.emit('chart', { type, id: ch.id, pivot: p.name });
      this.emit('change', {});
    },
    // Inserir › Gráfico com a célula dentro da Tabela Dinâmica cria um Gráfico Dinâmico (como no Excel)
    insertChart(type) {
      if (this.activePivot() && type !== 'rec') return this.pivotChart(type);
      return baseInsertChart.call(this, type);
    },
  });

  // ---------- Menu do botão direito dentro da Tabela Dinâmica ----------
  const baseCtx = proto.showContextMenu;
  proto.showContextMenu = function (kind, pt) {
    const p = kind === 'cell' ? this.activePivot() : null;
    if (!p) return baseCtx.call(this, kind, pt);
    this.lastPivot = p.id;
    const m = (label, cmd, arg, icon) => ({ label, icon: icon ? XL.icons[icon] : '', data: { cmd, arg: arg === undefined ? '' : String(arg) }, onClick: () => this.exec(cmd, arg) });
    const items = [m('Copiar', 'copy', undefined, 'copy'), { sep: true }, m('Atualizar', 'pivot-refresh', undefined, 'refreshSm')];
    if (p.vcols && p.vcols[this.grid.sel.ac] !== undefined && this.grid.sel.ar >= (p.dataTop || 0)) items.push(m('Configurações do Campo de Valor...', 'pivot-settings'));
    if (this.pivotGroupField()) items.push({ sep: true }, m('Agrupar...', 'pivot-group'), m('Desagrupar', 'pivot-ungroup'));
    items.push({ sep: true }, m(this.pivotPaneOff ? 'Mostrar Lista de Campos' : 'Ocultar Lista de Campos', 'pivot-fieldlist'));
    XL.ui.menu(pt, items);
  };
})();
