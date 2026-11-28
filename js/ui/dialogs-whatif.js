/* Módulo 10: Atingir Meta, Tabela de Dados, Gerenciador de Cenários e Solver. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const K = XL.SHEET_KEY;
  const proto = XL.ExcelApp.prototype;
  const baseExtra = proto.execExtra;
  const baseCommit = proto.commitInput;
  const enterSubmits = (inp) => inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const b = inp.closest('.ui-modal').querySelector('.btn-primary'); if (b) b.click(); } });
  const abs = (r, c) => '$' + XL.numToCol(c + 1) + '$' + (r + 1);
  const absRange = (g) => abs(g.r1, g.c1) + (g.r1 !== g.r2 || g.c1 !== g.c2 ? ':' + abs(g.r2, g.c2) : '');
  const ref1 = (t) => XL.parseRef(String(t || '').trim().replace(/^=/, '').replace(/\$/g, ''));
  const fmtNum = (v, f) => XL.formatValue(v, f || {}).text;
  const field = (label, inp) => h('div.wi-field', h('label.fd-lbl', label), inp);

  Object.assign(proto, {
    execExtra(cmd, arg) {
      switch (cmd) {
        case 'goalseek': return this.goalSeekDialog();
        case 'datatable': return this.dataTableDialog();
        case 'scenario-manager': return this.scenarioManager();
        case 'solver': return this.solverDialog();
      }
      return baseExtra.call(this, cmd, arg);
    },

    commitInput(r, c, text, sheet) {
      const S = sheet || this.sheet;
      if (S.dataTableAt && S.dataTableAt(r, c)) return { ok: false, msg: 'Não é possível alterar parte de uma Tabela de Dados. Ela é calculada pelo Excel a partir da fórmula e da célula de entrada.' };
      return baseCommit.call(this, r, c, text, sheet);
    },

    // ================= Atingir Meta =================
    goalSeekDialog() {
      const { ar, ac } = this.grid.sel;
      const set = h('input.fd-search', { type: 'text', value: abs(ar, ac), dataset: { role: 'gs-set' } });
      const to = h('input.fd-search', { type: 'text', value: '', dataset: { role: 'gs-to' } });
      const chg = h('input.fd-search', { type: 'text', value: '', dataset: { role: 'gs-change' } });
      [set, to, chg].forEach(enterSubmits);
      XL.ui.modal({
        title: 'Atingir Meta', width: 360,
        body: h('div.pvd', field('Definir célula:', set), field('Para valor:', to), field('Alternando célula:', chg),
          h('p.cd-note', '"Definir célula" tem a fórmula do resultado; "Alternando célula" é o número que o Excel vai testar até chegar ao valor.')),
        buttons: [{ label: 'OK', primary: true, onClick: () => this.runGoalSeek(set.value, to.value, chg.value) }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => to.focus(), 0);
    },
    runGoalSeek(setText, toText, chgText) {
      const S = this.sheet;
      const sp = ref1(setText), cp = ref1(chgText);
      const tv = XL.parseInput(String(toText).trim());
      if (!sp) { XL.ui.toast('Informe a célula a definir (ex.: B4).'); return false; }
      const sc = S.cell(sp.r, sp.c);
      if (!sc || sc.kind !== 'formula') { XL.ui.toast('A célula a definir deve conter uma <b>fórmula</b>.'); return false; }
      if (tv.kind !== 'number') { XL.ui.toast('Digite um número em "Para valor".'); return false; }
      if (!cp) { XL.ui.toast('Informe a célula que o Excel pode alterar (ex.: B1).'); return false; }
      const cc = S.cell(cp.r, cp.c);
      if (cc && cc.kind === 'formula') { XL.ui.toast('A célula a alternar deve conter um <b>valor</b>, não uma fórmula.'); return false; }
      const res = S.goalSeek(XL.toRef(sp.r, sp.c), tv.value, XL.toRef(cp.r, cp.c));
      this.pushUndo();
      if (res.ok) { S.setValue(cp.r, cp.c, res.x); this.render(); this.updateFormulaBar(); }
      const now = S.value(sp.r, sp.c);
      const setRef = XL.toRef(sp.r, sp.c);
      setTimeout(() => XL.ui.modal({
        title: 'Status do Atingir Meta', width: 380, className: 'gs-status',
        body: h('div.pvd',
          h('p', res.ok ? `Atingir Meta com a Célula ${setRef} encontrou uma solução.` : `Atingir Meta com a Célula ${setRef} pode não ter encontrado uma solução.`),
          h('table.gs-tbl', h('tr', h('td', 'Valor de destino:'), h('td', fmtNum(tv.value))), h('tr', h('td', 'Valor atual:'), h('td', fmtNum(now, sc.fmt))))),
        buttons: [
          { label: 'OK', primary: true, role: 'gs-ok', onClick: () => { if (!res.ok) this.undoStack.pop(); this.emit('goalseek', { set: setRef, change: XL.toRef(cp.r, cp.c), target: tv.value, ok: res.ok, x: res.x }); this.emit('change', {}); } },
          { label: 'Cancelar', role: 'gs-cancel', onClick: () => { this.exec('undo'); } },
        ],
        onClose: () => this.grid.focus(),
      }), 0);
    },

    // ================= Tabela de Dados =================
    dataTableDialog() {
      const g = this.grid.range();
      if (g.r2 - g.r1 < 1 || g.c2 - g.c1 < 1) { XL.ui.toast('Selecione a tabela inteira primeiro: a fórmula no topo (ou à esquerda) e os valores a testar.'); return; }
      const row = h('input.fd-search', { type: 'text', value: '', dataset: { role: 'dt-row' } });
      const col = h('input.fd-search', { type: 'text', value: '', dataset: { role: 'dt-col' } });
      [row, col].forEach(enterSubmits);
      XL.ui.modal({
        title: 'Tabela de Dados', width: 380,
        body: h('div.pvd', field('Célula de entrada da linha:', row), field('Célula de entrada da coluna:', col),
          h('p.cd-note', 'Valores a testar numa COLUNA (um embaixo do outro)? Use "Célula de entrada da coluna": a célula do modelo que receberá cada valor.')),
        buttons: [{ label: 'OK', primary: true, onClick: () => { if (!this.createDataTable(g, row.value, col.value)) return false; } }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => col.focus(), 0);
    },
    createDataTable(g, rowText, colText) {
      const S = this.sheet;
      const rp = String(rowText).trim() ? ref1(rowText) : null, cp = String(colText).trim() ? ref1(colText) : null;
      if ((String(rowText).trim() && !rp) || (String(colText).trim() && !cp) || (!rp && !cp)) { XL.ui.toast('Célula de entrada inválida. Digite uma célula, como <b>B1</b>.'); return false; }
      if (S.dataTables.some((t) => !(t.g.r2 < g.r1 || g.r2 < t.g.r1 || t.g.c2 < g.c1 || g.c2 < t.g.c1))) { XL.ui.toast('Já existe uma Tabela de Dados nessa área.'); return false; }
      this.pushUndo();
      const t = { g: Object.assign({}, g), rowIn: rp ? abs(rp.r, rp.c) : null, colIn: cp ? abs(cp.r, cp.c) : null };
      S.dataTables.push(t);
      const f = '=TABELA(' + (t.rowIn || '') + ';' + (t.colIn || '') + ')';
      for (let r = g.r1 + 1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) S.setInput(r, c, f);
      this.render();
      this.emit('datatable', { range: XL.rangeToStr(g), rowIn: t.rowIn, colIn: t.colIn });
      this.emit('change', {});
      return true;
    },

    // ================= Gerenciador de Cenários =================
    scenarioManager(pickName) {
      const S = this.sheet;
      const list = h('select.fd-cat.pv-list', { size: 6, dataset: { role: 'sc-list' } }, S.scenarios.map((s) => h('option', { value: s.name }, s.name)));
      const info = h('div.sc-info');
      const cur = () => S.scenarios.find((s) => s.name === list.value) || null;
      const showInfo = () => { const s = cur(); info.innerHTML = ''; if (s) info.append(h('div', h('b', 'Células variáveis: '), s.cells), h('div', h('b', 'Comentário: '), s.comment || '—')); };
      list.addEventListener('change', showInfo);
      if (S.scenarios.length) list.value = pickName && S.scenarios.some((s) => s.name === pickName) ? pickName : S.scenarios[0].name;
      showInfo();
      const need = () => { if (!cur()) { XL.ui.toast('Escolha um cenário na lista.'); return false; } return true; };
      XL.ui.modal({
        title: 'Gerenciador de Cenários', width: 460,
        body: h('div.pvd', h('label.fd-lbl', 'Cenários:'),
          S.scenarios.length ? list : h('p.cd-note', 'Nenhum cenário definido. Para adicionar, clique em Adicionar.'), info),
        buttons: [
          { label: 'Adicionar...', role: 'sc-add', onClick: () => { setTimeout(() => this.scenarioEdit(null), 0); } },
          { label: 'Editar...', role: 'sc-edit', onClick: () => { if (!need()) return false; const s = cur(); setTimeout(() => this.scenarioEdit(s), 0); } },
          { label: 'Excluir', role: 'sc-del', onClick: () => { if (!need()) return false; this.pushUndo(); S.scenarios = S.scenarios.filter((x) => x !== cur()); setTimeout(() => this.scenarioManager(), 0); } },
          { label: 'Mostrar', role: 'sc-show', keepOpen: true, onClick: () => { if (!need()) return false; this.scenarioShow(cur()); } },
          { label: 'Resumo...', role: 'sc-summary', onClick: () => { setTimeout(() => this.scenarioSummaryDialog(), 0); } },
          { label: 'Fechar', primary: true, role: 'sc-close' },
        ],
        onClose: () => this.grid.focus(),
      });
    },
    scenarioEdit(sc) {
      const g = this.grid.range();
      const name = h('input.fd-search', { type: 'text', value: sc ? sc.name : '', dataset: { role: 'sc-name' } });
      const cells = h('input.fd-search', { type: 'text', value: sc ? sc.cells : absRange(g), dataset: { role: 'sc-cells' } });
      const comment = h('input.fd-search', { type: 'text', value: sc ? sc.comment || '' : '', dataset: { role: 'sc-comment' } });
      [name, cells, comment].forEach(enterSubmits);
      XL.ui.modal({
        title: sc ? 'Editar Cenário' : 'Adicionar Cenário', width: 420,
        body: h('div.pvd', field('Nome do cenário:', name), field('Células variáveis:', cells), h('p.cd-note', 'Use ; para separar células que não estão lado a lado (até 32 células).'), field('Comentário:', comment)),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const nm = name.value.trim();
            if (!nm) { XL.ui.toast('Dê um nome ao cenário.'); return false; }
            if (this.sheet.scenarios.some((s) => s !== sc && XL.normName(s.name) === XL.normName(nm))) { XL.ui.toast('Já existe um cenário com esse nome.'); return false; }
            const list = XL.whatIfCells(cells.value);
            if (!list || !list.length) { XL.ui.toast('Células variáveis inválidas. Ex.: <b>$B$1:$B$2</b>.'); return false; }
            if (list.length > 32) { XL.ui.toast('Um cenário pode ter no máximo 32 células variáveis.'); return false; }
            setTimeout(() => this.scenarioValues(sc, nm, cells.value.trim(), comment.value.trim(), list), 0);
          },
        }, { label: 'Cancelar', onClick: () => { setTimeout(() => this.scenarioManager(), 0); } }],
      });
      setTimeout(() => name.focus(), 0);
    },
    scenarioValues(sc, nm, cellsText, comment, list) {
      const S = this.sheet;
      const inputs = list.map((p, i) => {
        const v = sc && sc.values[i] !== undefined ? sc.values[i] : XL.editTextFor(S.cell(p.r, p.c));
        const inp = h('input.fd-search', { type: 'text', value: v, dataset: { role: 'sc-val', cell: XL.toRef(p.r, p.c) } });
        enterSubmits(inp);
        const lbl = S.textOf(p.r, p.c - 1);
        return h('div.sc-row', h('span.sc-lbl', `${i + 1}:  ${abs(p.r, p.c)}${lbl && p.c > 0 && typeof S.value(p.r, p.c - 1) === 'string' ? ' (' + lbl + ')' : ''}`), inp);
      });
      XL.ui.modal({
        title: 'Valores de Cenário', width: 420,
        body: h('div.pvd', h('p.cf-q', 'Insira valores para cada célula variável.'), inputs),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const values = inputs.map((row) => row.querySelector('input').value.trim());
            this.pushUndo();
            if (sc) Object.assign(sc, { name: nm, cells: cellsText, comment, values });
            else S.scenarios.push({ name: nm, cells: cellsText, comment, values });
            this.emit('scenario', { name: nm, cells: cellsText, values });
            this.emit('change', {});
            setTimeout(() => this.scenarioManager(nm), 0);
          },
        }, { label: 'Cancelar', onClick: () => { setTimeout(() => this.scenarioManager(), 0); } }],
      });
      setTimeout(() => { const f = inputs[0] && inputs[0].querySelector('input'); if (f) f.select(); }, 0);
    },
    scenarioShow(sc) {
      const S = this.sheet;
      const list = XL.whatIfCells(sc.cells) || [];
      this.pushUndo();
      list.forEach((p, i) => { if (sc.values[i] !== undefined) S.setInput(p.r, p.c, String(sc.values[i])); });
      this.render();
      this.updateFormulaBar();
      this.emit('scenarioshow', { name: sc.name });
      this.emit('change', {});
    },
    scenarioSummaryDialog() {
      const S = this.sheet;
      if (!S.scenarios.length) { XL.ui.toast('Não há cenários nesta planilha.'); return; }
      const res = h('input.fd-search', { type: 'text', value: abs(this.grid.sel.ar, this.grid.sel.ac), dataset: { role: 'sc-result' } });
      enterSubmits(res);
      XL.ui.modal({
        title: 'Resumo do Cenário', width: 380,
        body: h('div.pvd', h('label.sd-check', h('input', { type: 'radio', checked: true }), ' Resumo do cenário'), field('Células de resultado:', res),
          h('p.cd-note', 'As células cujo valor você quer comparar entre os cenários (ex.: o lucro).')),
        buttons: [{ label: 'OK', primary: true, onClick: () => { if (!this.scenarioSummary(res.value)) return false; } }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
    },
    scenarioSummary(resultText) {
      const S = this.sheet;
      const results = XL.whatIfCells(resultText);
      if (!results || !results.length) { XL.ui.toast('Células de resultado inválidas.'); return false; }
      const vars = [];
      const seen = new Set();
      for (const sc of S.scenarios) for (const p of XL.whatIfCells(sc.cells) || []) { const k = p.r * K + p.c; if (!seen.has(k)) { seen.add(k); vars.push(p); } }
      this.pushUndo();
      let name = 'Resumo do Cenário', n = 2;
      while (this.wb.byName(name)) name = 'Resumo do Cenário ' + n++;
      const idx = this.wb.activeIndex;
      const T = this.wb.addSheet(name, idx);
      const head = { b: true, fill: '#D9D9D9' };
      const label = (p) => { const l = p.c > 0 && typeof S.value(p.r, p.c - 1) === 'string' ? S.textOf(p.r, p.c - 1) : ''; return abs(p.r, p.c) + (l ? ' (' + l + ')' : ''); };
      const put = (r, c, v, fmt) => { if (v !== null && v !== undefined) T.setValue(r, c, v); if (fmt) T.setFmt(r, c, fmt); };
      put(1, 1, 'Resumo do Cenário', { b: true, size: 13 });
      put(2, 2, 'Valores atuais:', head);
      S.scenarios.forEach((sc, j) => put(2, 3 + j, sc.name, head));
      put(3, 1, 'Células variáveis:', { b: true });
      let r = 4;
      const valOf = (sc, p) => {
        const cells = XL.whatIfCells(sc.cells) || [];
        const i = cells.findIndex((q) => q.r === p.r && q.c === p.c);
        if (i < 0) return S.value(p.r, p.c);
        const pi = XL.parseInput(String(sc.values[i]));
        return pi.kind === 'number' ? pi.value : sc.values[i];
      };
      for (const p of vars) {
        const f = Object.assign({}, S.fmt(p.r, p.c));
        delete f.fill; delete f.b;
        put(r, 1, label(p));
        put(r, 2, S.value(p.r, p.c), f);
        S.scenarios.forEach((sc, j) => put(r, 3 + j, valOf(sc, p), f));
        r++;
      }
      put(r, 1, 'Células de resultado:', { b: true });
      r++;
      for (const p of results) {
        const f = Object.assign({}, S.fmt(p.r, p.c));
        delete f.fill;
        put(r, 1, label(p));
        put(r, 2, S.value(p.r, p.c), f);
        S.scenarios.forEach((sc, j) => {
          const ov = new Map();
          for (const q of vars) { const v = valOf(sc, q); if (typeof v === 'number') ov.set(q.r * K + q.c, v); }
          put(r, 3 + j, S.evalWith(p.r, p.c, ov), f);
        });
        r++;
      }
      put(r + 1, 1, 'Observações: a coluna Valores atuais mostra as células variáveis no momento em que o resumo foi criado.', { i: true, color: '#595959' });
      T.colW[0] = 24; T.colW[1] = 190; T.colW[2] = 110;
      S.scenarios.forEach((_, j) => { T.colW[3 + j] = 110; });
      this.wb.activeIndex = idx;
      this.afterWorkbookChange();
      this.render();
      this.emit('sheetadd', { name });
      this.emit('scenariosummary', { sheet: name, results: resultText });
      this.emit('change', {});
      return true;
    },

    // ================= Solver =================
    solverDialog() {
      const S = this.sheet;
      const m = JSON.parse(JSON.stringify(S.solver || { obj: abs(this.grid.sel.ar, this.grid.sel.ac), goal: 'max', target: '0', vars: '', cons: [], nonneg: true, method: 'GRG Não Linear' }));
      const obj = h('input.fd-search', { type: 'text', value: m.obj, dataset: { role: 'sv-obj' } });
      const radio = (val, label, role) => { const r = h('input', { type: 'radio', name: 'svgoal', value: val, dataset: { role } }); r.checked = m.goal === val; return h('label.sv-radio', r, ' ' + label); };
      const target = h('input.fd-search.sv-target', { type: 'text', value: m.target || '0', dataset: { role: 'sv-target' } });
      const vars = h('input.fd-search', { type: 'text', value: m.vars, dataset: { role: 'sv-vars' } });
      const consBox = h('div.sv-cons', { dataset: { role: 'sv-cons' } });
      const OPS = { '<=': '<=', '=': '=', '>=': '>=', int: 'núm' };
      const drawCons = () => {
        consBox.innerHTML = '';
        if (!m.cons.length) consBox.appendChild(h('div.sv-empty', 'Nenhuma restrição.'));
        m.cons.forEach((k, i) => consBox.appendChild(h('div.sv-con', { dataset: { i } }, h('span', `${k.ref} ${OPS[k.op]} ${k.op === 'int' ? 'inteiro' : k.rhs}`),
          h('button.sv-del', { type: 'button', title: 'Excluir restrição', dataset: { role: 'sv-cdel' }, onclick: () => { m.cons.splice(i, 1); drawCons(); } }, '✕'))));
      };
      drawCons();
      const cref = h('input.fd-search', { type: 'text', placeholder: 'Referência', dataset: { role: 'sv-cref' } });
      const cop = h('select.fd-cat', { dataset: { role: 'sv-cop' } }, Object.keys(OPS).map((k) => h('option', { value: k }, OPS[k])));
      const crhs = h('input.fd-search', { type: 'text', placeholder: 'Restrição', dataset: { role: 'sv-crhs' } });
      cop.addEventListener('change', () => { crhs.disabled = cop.value === 'int'; });
      const add = h('button.btn.btn-sm', {
        type: 'button', dataset: { role: 'sv-cadd' },
        onclick: () => {
          if (!XL.whatIfCells(cref.value)) { XL.ui.toast('Referência de célula inválida na restrição.'); return; }
          if (cop.value !== 'int' && !String(crhs.value).trim()) { XL.ui.toast('Digite o valor ou a célula da restrição.'); return; }
          m.cons.push({ ref: cref.value.trim(), op: cop.value, rhs: cop.value === 'int' ? '' : crhs.value.trim() });
          cref.value = ''; crhs.value = '';
          drawCons();
        },
      }, 'Adicionar');
      const nonneg = h('input', { type: 'checkbox', dataset: { role: 'sv-nonneg' } });
      nonneg.checked = m.nonneg !== false;
      const method = h('select.fd-cat', { dataset: { role: 'sv-method' } }, ['GRG Não Linear', 'LP Simplex', 'Evolutionary'].map((x) => h('option', { value: x }, x)));
      method.value = m.method || 'GRG Não Linear';
      const read = () => {
        m.obj = obj.value.trim();
        m.goal = (document.querySelector('input[name="svgoal"]:checked') || {}).value || 'max';
        m.target = target.value.trim();
        m.vars = vars.value.trim();
        m.nonneg = nonneg.checked;
        m.method = method.value;
        return m;
      };
      const save = () => { this.pushUndo(); S.solver = JSON.parse(JSON.stringify(read())); this.emit('solvermodel', Object.assign({}, S.solver)); };
      XL.ui.modal({
        title: 'Parâmetros do Solver', width: 560,
        body: h('div.pvd.sv',
          field('Definir Objetivo:', obj),
          h('div.sv-goal', h('span.fd-lbl', 'Para:'), radio('max', 'Máx.', 'sv-max'), radio('min', 'Mín.', 'sv-min'), radio('value', 'Valor de:', 'sv-val'), target),
          field('Alterando Células Variáveis:', vars),
          h('label.fd-lbl', 'Sujeito às Restrições:'), consBox,
          h('div.sv-addrow', cref, cop, crhs, add),
          h('label.sd-check', nonneg, ' Tornar Variáveis Irrestritas Não Negativas'),
          field('Selecionar um Método de Solução:', method),
          h('p.cd-note', 'No Excel, o Solver é um suplemento: ative uma vez em Arquivo › Opções › Suplementos › Suplemento Solver. Ele aparece em Dados › Análise.')),
        buttons: [
          { label: 'Resolver', primary: true, role: 'sv-solve', onClick: () => { save(); return this.runSolver(); } },
          { label: 'Fechar', role: 'sv-close', onClick: () => { save(); } },
        ],
        onClose: () => this.grid.focus(),
      });
    },
    runSolver() {
      const S = this.sheet;
      const res = S.solve(S.solver);
      if (res.status === 'error') { XL.ui.toast(res.msg); return false; }
      const before = res.vars.map((p) => S.value(p.r, p.c));
      if (res.status === 'ok') {
        this.pushUndo();
        res.vars.forEach((p, i) => S.setValue(p.r, p.c, res.x[i]));
        this.render();
        this.updateFormulaBar();
      }
      const msg = res.status === 'ok' ? 'O Solver encontrou uma solução. Todas as Restrições e condições de adequação foram satisfeitas.'
        : res.status === 'unbounded' ? 'Os valores da Célula Objetivo não convergem: o resultado pode crescer sem limite. Faltam restrições?'
          : 'O Solver não pôde encontrar uma solução viável: nenhuma combinação respeita todas as restrições.';
      const keep = h('input', { type: 'radio', name: 'svkeep', value: 'keep', dataset: { role: 'sv-keep' } });
      const restore = h('input', { type: 'radio', name: 'svkeep', value: 'restore', dataset: { role: 'sv-restore' } });
      keep.checked = true;
      setTimeout(() => XL.ui.modal({
        title: 'Resultados do Solver', width: 460,
        body: h('div.pvd', h('p.sv-msg' + (res.status === 'ok' ? '.ok' : '.bad'), msg),
          res.status === 'ok' ? h('div', h('label.sd-check', keep, ' Manter Solução do Solver'), h('label.sd-check', restore, ' Restaurar Valores Originais'),
            h('p.cd-note', `Método usado: ${res.method}. Valor do objetivo: ${XL.formatValue(res.value, {}).text}.`)) : null),
        buttons: [{
          label: 'OK', primary: true, role: 'sv-ok',
          onClick: () => {
            const kept = res.status === 'ok' && keep.checked;
            if (res.status === 'ok' && !kept) this.exec('undo');
            this.emit('solver', { status: res.status, kept, x: res.x, before });
            this.emit('change', {});
          },
        }],
        onClose: () => this.grid.focus(),
      }), 0);
    },
  });
})();
