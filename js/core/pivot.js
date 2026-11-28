/* Tabela Dinâmica: lê a fonte (tabela ou intervalo), filtra, agrupa, resume e escreve o resultado nas células.
   Também guarda as Segmentações de Dados e Linhas do Tempo ligadas a ela. Sem interface aqui. */
(function () {
  'use strict';
  const XL = window.XL;
  const K = XL.SHEET_KEY;
  const P = XL.Sheet.prototype;

  const MONTHS = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const FN = { sum: 'Soma', count: 'Contagem', avg: 'Média', max: 'Máx.', min: 'Mín.' };
  const UNIT_NAMES = { years: 'Anos', quarters: 'Trimestres', months: 'Meses' };
  const ORDER = ['months', 'quarters', 'years'];
  const GRAND = '\u0000grand';
  const norm = (s) => XL.normName(String(s));
  const isNum = (v) => typeof v === 'number';
  const isBlank = (v) => v === null || v === undefined || v === '';
  const isDateFmt = (f) => !!f && (f.numFmt === 'data' || f.numFmt === 'datalonga' || (f.numFmt === 'custom' && XL.isDateCode(f.code || '')));
  const cmpKey = (a, b) => {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return String(a).localeCompare(String(b), 'pt-BR', { sensitivity: 'base', numeric: true });
  };
  const finest = (by) => ORDER.find((u) => by.includes(u)) || 'months';

  // ---------------- Fonte de dados ----------------
  function source(wb, p) {
    let S = null, g = null;
    if (p.src.table) {
      for (const s of wb.sheets) {
        const t = s.tableByName(p.src.table);
        if (t) { S = s; g = { r1: t.g.r1, c1: t.g.c1, r2: s.tableDataRows(t).r2, c2: t.g.c2 }; break; }
      }
    } else {
      S = wb.byName(p.src.sheet);
      g = XL.parseRange(p.src.range);
    }
    if (!S || !g || g.r2 <= g.r1) return null;
    const headers = [];
    const seen = new Map();
    for (let c = g.c1; c <= g.c2; c++) {
      const v = S.value(g.r1, c);
      let name = isBlank(v) ? 'Coluna' + (c - g.c1 + 1) : XL.formatValue(v, S.fmt(g.r1, c)).text;
      const k = norm(name);
      if (seen.has(k)) { seen.set(k, seen.get(k) + 1); name += seen.get(k); } else seen.set(k, 1);
      headers.push(name);
    }
    const recs = [];
    const fmts = headers.map(() => null);
    const kinds = headers.map(() => ({ num: 0, text: 0, date: 0 }));
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const row = [];
      let any = false;
      for (let c = g.c1; c <= g.c2; c++) {
        let v = S.value(r, c);
        if (v && v.isXLError) v = v.code;
        if (!isBlank(v)) any = true;
        row.push(v);
        const j = c - g.c1;
        if (isBlank(v)) continue;
        const f = S.fmt(r, c);
        if (isNum(v)) { if (isDateFmt(f)) kinds[j].date++; else kinds[j].num++; if (!fmts[j]) fmts[j] = f; } else kinds[j].text++;
      }
      if (any) recs.push(row);
    }
    const type = kinds.map((k) => (k.date && k.date >= k.num && k.date >= k.text ? 'date' : k.num && k.num >= k.text ? 'num' : 'text'));
    return { S, g, headers, recs, type, fmts };
  }

  // Campos disponíveis: colunas da fonte + campos de grupo (Anos, Trimestres) + campos calculados
  function fieldList(wb, p, src) {
    src = src || source(wb, p);
    if (!src) return [];
    const out = src.headers.map((name, j) => ({ name, j, type: src.type[j] }));
    for (const base in p.groups || {}) {
      const gr = p.groups[base];
      if (!gr.by) continue;
      const fine = finest(gr.by);
      for (const u of ['years', 'quarters']) if (gr.by.includes(u) && u !== fine) out.push({ name: UNIT_NAMES[u], virtual: { base, unit: u }, type: 'text' });
    }
    for (const cf of p.calc || []) out.push({ name: cf.name, calc: cf, type: 'num' });
    return out;
  }

  function dateItem(serial, unit) {
    const d = XL.serialToParts(serial);
    if (unit === 'years') return { key: d.y, label: String(d.y), value: String(d.y) };
    if (unit === 'quarters') { const q = Math.floor((d.m - 1) / 3) + 1; return { key: q, label: 'Trim' + q, value: 'Trim' + q }; }
    return { key: d.m, label: MONTHS[d.m - 1], value: MONTHS[d.m - 1] };
  }

  // ---------------- Cálculo ----------------
  function makeCtx(wb, p, src) {
    const fields = fieldList(wb, p, src);
    const fmap = new Map(fields.map((f) => [norm(f.name), f]));
    const idx = new Map(src.headers.map((h, j) => [norm(h), j]));
    const item = (fname, rec) => {
      const fd = fmap.get(norm(fname));
      if (!fd || fd.calc) return null;
      if (fd.virtual) {
        const v = rec[idx.get(norm(fd.virtual.base))];
        return isNum(v) ? dateItem(v, fd.virtual.unit) : { key: '￿', label: '(vazio)', value: '(vazio)' };
      }
      const v = rec[fd.j];
      if (isBlank(v)) return { key: '￿', label: '(vazio)', value: '(vazio)' };
      const gr = (p.groups || {})[fd.name];
      if (gr && gr.by && isNum(v)) return dateItem(v, finest(gr.by));
      if (gr && gr.step && isNum(v)) {
        const lo = gr.start + Math.floor((v - gr.start) / gr.step) * gr.step;
        if (v < gr.start) return { key: -Infinity, label: '<' + XL.fmtGeneral(gr.start), value: '<' + XL.fmtGeneral(gr.start) };
        if (v > gr.end) return { key: Infinity, label: '>' + XL.fmtGeneral(gr.end), value: '>' + XL.fmtGeneral(gr.end) };
        const lbl = XL.fmtGeneral(lo) + '-' + XL.fmtGeneral(lo + gr.step - 1);
        return { key: lo, label: lbl, value: lbl };
      }
      if (fd.type === 'date' && isNum(v)) return { key: v, label: XL.formatValue(v, { numFmt: 'data' }).text, value: v, fmt: { numFmt: 'data' } };
      if (isNum(v)) { const f = src.fmts[fd.j] || {}; return { key: v, label: XL.formatValue(v, f).text, value: v, fmt: f.numFmt ? { numFmt: f.numFmt, dec: f.dec, code: f.code } : null }; }
      return { key: String(v).toLocaleLowerCase('pt-BR'), label: String(v), value: String(v) };
    };
    const numOf = (fname, rec) => { const fd = fmap.get(norm(fname)); if (!fd || fd.virtual || fd.calc) return null; return rec[fd.j]; };
    return { fields, fmap, item, numOf };
  }

  // Avalia a fórmula do campo calculado trocando os nomes dos campos pelas somas
  function evalCalc(wb, formula, names, sumOf) {
    let f = String(formula || '').trim();
    if (f[0] === '=') f = f.slice(1);
    const sorted = names.slice().sort((a, b) => b.length - a.length);
    let out = '';
    let i = 0;
    const isWord = (ch) => !!ch && /[A-Za-zÀ-ÿ0-9_.]/.test(ch);
    while (i < f.length) {
      if (f[i] === '"') { const j = f.indexOf('"', i + 1); out += f.slice(i, j < 0 ? f.length : j + 1); i = j < 0 ? f.length : j + 1; continue; }
      if (f[i] === "'") {
        const j = f.indexOf("'", i + 1);
        const nm = f.slice(i + 1, j < 0 ? f.length : j);
        const hit = names.find((n) => norm(n) === norm(nm));
        if (hit) { out += '(' + numText(sumOf(hit)) + ')'; i = j + 1; continue; }
      }
      if (!isWord(f[i - 1])) {
        const low = f.slice(i).toLocaleLowerCase('pt-BR');
        const hit = sorted.find((n) => low.startsWith(n.toLocaleLowerCase('pt-BR')) && !isWord(f[i + n.length]));
        if (hit) { out += '(' + numText(sumOf(hit)) + ')'; i += hit.length; continue; }
      }
      out += f[i++];
    }
    const ast = XL.Formula.parse(XL.Formula.normalize('=' + out).slice(1));
    const ctx = wb.sheets[0].ctx;
    return XL.Formula.scalar(XL.Formula.evaluate(ast, ctx));
  }
  const numText = (n) => (isNum(n) ? String(n).replace('.', ',') : '0');

  function aggregate(cx, wb, p, v, recs) {
    if (!recs.length) return null;
    const fd = cx.fmap.get(norm(v.field));
    if (fd && fd.calc) {
      const names = cx.fields.filter((f) => !f.calc && !f.virtual).map((f) => f.name);
      const sumOf = (name) => recs.reduce((s, r) => { const x = cx.numOf(name, r); return isNum(x) ? s + x : s; }, 0);
      try { return evalCalc(wb, fd.calc.formula, names, sumOf); } catch (e) { return XL.Formula.err ? XL.Formula.err(XL.Formula.E.NAME) : null; }
    }
    const vals = recs.map((r) => cx.numOf(v.field, r));
    if (v.fn === 'count') { const n = vals.filter((x) => !isBlank(x)).length; return n; }
    const nums = vals.filter(isNum);
    if (v.fn === 'avg') return nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : null;
    if (v.fn === 'max') return nums.length ? Math.max(...nums) : null;
    if (v.fn === 'min') return nums.length ? Math.min(...nums) : null;
    return nums.reduce((s, x) => s + x, 0);
  }

  // Filtros de relatório, rótulos ocultos, segmentações e linhas do tempo
  function passes(cx, p, rec, extra) {
    for (const f of p.filters || []) {
      if (!f.sel) continue;
      const it = cx.item(f.field, rec);
      if (!it || !f.sel.includes(it.label)) return false;
    }
    for (const fname in p.hidden || {}) {
      const hid = p.hidden[fname];
      if (!hid || !hid.length) continue;
      const it = cx.item(fname, rec);
      if (it && hid.includes(it.label)) return false;
    }
    for (const sl of extra) {
      if (sl.kind === 'timeline') {
        if (!sl.range) continue;
        const v = cx.numOf(sl.field, rec);
        if (!isNum(v) || v < sl.range.from || v > sl.range.to) return false;
      } else if (sl.sel && sl.sel.length) {
        const it = cx.item(sl.field, rec);
        if (!it || !sl.sel.includes(it.label)) return false;
      }
    }
    return true;
  }

  function compute(wb, p, host) {
    const src = source(wb, p);
    if (!src) return { error: 'A fonte de dados da Tabela Dinâmica não foi encontrada.' };
    const cx = makeCtx(wb, p, src);
    const extra = host ? host.slicers.filter((s) => s.pivots.includes(p.id)) : [];
    const recs = src.recs.filter((r) => passes(cx, p, r, extra));
    const vals = (p.vals || []).filter((v) => cx.fmap.has(norm(v.field)));
    const rows = (p.rows || []).filter((f) => cx.fmap.has(norm(f)));
    const cols = (p.cols || []).filter((f) => cx.fmap.has(norm(f)));

    // árvore das linhas
    const root = { recs, children: null, depth: -1 };
    const grow = (node, level) => {
      if (level >= rows.length) return;
      const map = new Map();
      for (const r of node.recs) {
        const it = cx.item(rows[level], r);
        const k = typeof it.key + ':' + it.key;
        if (!map.has(k)) map.set(k, { item: it, recs: [], depth: level, field: rows[level] });
        map.get(k).recs.push(r);
      }
      node.children = [...map.values()];
      node.children.forEach((ch) => grow(ch, level + 1));
    };
    grow(root, 0);

    // itens das colunas
    let colItems = [null];
    if (cols.length) {
      const map = new Map();
      for (const r of recs) { const it = cx.item(cols[0], r); const k = typeof it.key + ':' + it.key; if (!map.has(k)) map.set(k, it); }
      colItems = [...map.values()];
      const s = (p.sort || {})[cols[0]];
      colItems.sort((a, b) => cmpKey(a.key, b.key) * (s && s.dir === 'desc' ? -1 : 1));
    }
    const inCol = (ci, rec) => ci === null || ci === GRAND || (() => { const it = cx.item(cols[0], rec); return it.key === ci.key && typeof it.key === typeof ci.key; })();
    const agg = (v, list, ci) => aggregate(cx, wb, p, v, ci === null || ci === GRAND ? list : list.filter((r) => inCol(ci, r)));

    // ordenação das linhas (por rótulo ou por valor)
    const sortTree = (node) => {
      if (!node.children) return;
      const s = (p.sort || {})[node.children[0] && node.children[0].field];
      const dir = s && s.dir === 'desc' ? -1 : 1;
      if (s && s.by === 'value' && vals[s.val || 0]) {
        const v = vals[s.val || 0];
        const val = (n) => { const x = agg(v, n.recs, null); return isNum(x) ? x : -Infinity; };
        node.children.sort((a, b) => (val(a) - val(b)) * dir);
      } else node.children.sort((a, b) => cmpKey(a.item.key, b.item.key) * dir);
      node.children.forEach(sortTree);
    };
    sortTree(root);

    const show = (v, x, rowRecs, ci) => {
      if (!v.show || v.show === 'normal' || !isNum(x)) return x;
      let base;
      if (v.show === 'pctTotal') base = agg(v, recs, null);
      else if (v.show === 'pctCol') base = agg(v, recs, ci);
      else if (v.show === 'pctRow') base = agg(v, rowRecs, null);
      return isNum(base) && base !== 0 ? x / base : null;
    };
    return { src, cx, recs, root, rows, cols, vals, colItems, agg, show };
  }

  // ---------------- Escrita nas células ----------------
  P.pivotAt = function (r, c) { return (this.pivots || []).find((p) => p.out && XL.inRange(p.out, r, c)) || null; };
  P.pivotById = function (id) { return (this.pivots || []).find((p) => p.id === id) || null; };
  P.nextPivotName = function () {
    let i = 1;
    const all = this.wb ? this.wb.sheets.flatMap((s) => s.pivots || []) : this.pivots;
    while (all.some((p) => p.name === 'TabelaDinâmica' + i)) i++;
    return 'TabelaDinâmica' + i;
  };

  const valName = (v) => v.name || (FN[v.fn || 'sum'] + ' de ' + v.field);

  P.renderPivot = function (p) {
    const S = this;
    const wb = S.wb;
    // apaga o resultado anterior
    if (p.out) {
      for (let r = p.out.r1; r <= p.out.r2; r++) for (let c = p.out.c1; c <= p.out.c2; c++) S.cells.delete(r * K + c);
    }
    const put = [];
    const W = (r, c, v, fmt) => put.push({ r, c, v, fmt: fmt || {} });
    const res = compute(wb, p, S);
    const st = XL.TABLE_STYLES[p.style] || XL.TABLE_STYLES.azul;
    const head = { b: true, fill: st.band };
    const C0 = p.at.c;
    const nf = (p.filters || []).length;
    const top = Math.max(p.at.r, nf ? nf + 1 : 0);
    p.btns = [];
    p.vcols = {}; // coluna → índice do campo de valor
    p.rmap = {}; // linha → campo de linha (para Agrupar)
    p.error = res.error || null;
    if (res.error) {
      W(top, C0, res.error, { color: '#C00000' });
    } else {
      // filtros de relatório
      (p.filters || []).forEach((f, i) => {
        const r = top - 1 - nf + i;
        const lbl = !f.sel ? '(Tudo)' : f.sel.length === 1 ? f.sel[0] : '(Vários Itens)';
        W(r, C0, f.field, { b: true });
        W(r, C0 + 1, lbl, {});
        p.btns.push({ r, c: C0 + 1, kind: 'filter', field: f.field, on: !!f.sel });
      });
      const { rows, cols, vals, colItems, root, agg, show } = res;
      if (!rows.length && !cols.length && !vals.length) {
        W(top, C0, p.name, { b: true, color: '#1F4E79' });
        W(top + 1, C0, 'Para criar um relatório, escolha os campos', { i: true, color: '#595959' });
        W(top + 2, C0, 'na Lista de Campos da Tabela Dinâmica.', { i: true, color: '#595959' });
      } else {
        const hasCols = cols.length > 0;
        const nv = vals.length;
        const LC = rows.length || hasCols ? 1 : 0; // coluna de rótulos
        const leaves = [];
        for (const ci of colItems) { if (nv) vals.forEach((v) => leaves.push({ ci, v })); else leaves.push({ ci, v: null }); }
        if (hasCols && p.grandCols !== false && nv) vals.forEach((v) => leaves.push({ ci: GRAND, v }));
        const hdr = hasCols ? (nv > 1 ? 3 : 2) : 1;
        const pctFmt = { numFmt: 'porcentagem', dec: 2 };
        const vfmt = (v) => (v.show && v.show !== 'normal' ? pctFmt : v.numFmt ? { numFmt: v.numFmt, dec: v.dec } : {});
        // cabeçalhos
        if (hasCols) {
          W(top, C0, nv === 1 ? valName(vals[0]) : '', head);
          W(top, C0 + LC, 'Rótulos de Coluna', head);
          p.btns.push({ r: top, c: C0 + LC, kind: 'cols', field: cols[0], on: !!((p.hidden || {})[cols[0]] || []).length });
          for (let i = 1; i < leaves.length; i++) W(top, C0 + LC + i, '', head);
          if (nv > 1) {
            W(top + 1, C0, '', head);
            leaves.forEach((lf, i) => {
              const first = i === 0 || leaves[i - 1].ci !== lf.ci;
              W(top + 1, C0 + LC + i, first ? (lf.ci === GRAND ? '' : lf.ci.value) : '', Object.assign({}, head, lf.ci && lf.ci.fmt ? lf.ci.fmt : {}));
            });
          }
          const hr = top + hdr - 1;
          W(hr, C0, rows.length ? 'Rótulos de Linha' : '', head);
          if (rows.length) p.btns.push({ r: hr, c: C0, kind: 'rows', field: rows[0], on: !!((p.hidden || {})[rows[0]] || []).length });
          leaves.forEach((lf, i) => {
            let t;
            if (nv > 1) t = lf.ci === GRAND ? 'Total ' + valName(lf.v) : valName(lf.v);
            else t = lf.ci === GRAND ? 'Total Geral' : lf.ci.value;
            W(hr, C0 + LC + i, t, Object.assign({}, head, nv <= 1 && lf.ci && lf.ci !== GRAND && lf.ci.fmt ? lf.ci.fmt : {}));
          });
        } else {
          if (LC) { W(top, C0, 'Rótulos de Linha', head); p.btns.push({ r: top, c: C0, kind: 'rows', field: rows[0], on: !!((p.hidden || {})[rows[0]] || []).length }); }
          leaves.forEach((lf, i) => W(top, C0 + LC + i, lf.v ? valName(lf.v) : '', head));
        }
        leaves.forEach((lf, i) => { if (lf.v) p.vcols[C0 + LC + i] = p.vals.indexOf(lf.v); });
        p.dataTop = top + hdr;
        // linhas de dados
        let r = top + hdr;
        const writeVals = (rr, recsOf, bold) => leaves.forEach((lf, i) => {
          if (!lf.v) return;
          const x = show(lf.v, agg(lf.v, recsOf, lf.ci), recsOf, lf.ci);
          W(rr, C0 + LC + i, x, Object.assign({}, vfmt(lf.v), bold ? { b: true } : {}));
        });
        const walk = (node) => {
          for (const ch of node.children || []) {
            const parent = !!ch.children;
            p.rmap[r] = ch.field;
            W(r, C0, ch.item.value, Object.assign({}, ch.item.fmt || {}, parent ? { b: true } : {}, ch.depth ? { indent: ch.depth, align: 'left' } : { align: 'left' }));
            if (!parent || p.subtotals !== false) writeVals(r, ch.recs, parent);
            r++;
            walk(ch);
          }
        };
        if (rows.length) {
          walk(root);
          if (p.grandRows !== false) {
            W(r, C0, 'Total Geral', head);
            leaves.forEach((lf, i) => {
              if (!lf.v) { W(r, C0 + LC + i, '', head); return; }
              const x = show(lf.v, agg(lf.v, root.recs, lf.ci), root.recs, lf.ci);
              W(r, C0 + LC + i, x, Object.assign({}, vfmt(lf.v), head));
            });
            r++;
          }
        } else {
          if (LC) W(r, C0, '', {});
          writeVals(r, root.recs, false);
        }
      }
    }
    // grava e calcula a área ocupada
    let g = null;
    for (const x of put) {
      if (x.r < 0 || x.c < 0 || x.r >= S.rows || x.c >= S.cols) continue;
      S.setValue(x.r, x.c, x.v === '' ? null : x.v);
      S.replaceFmt(x.r, x.c, x.fmt);
      if (!g) g = { r1: x.r, c1: x.c, r2: x.r, c2: x.c };
      else { g.r1 = Math.min(g.r1, x.r); g.c1 = Math.min(g.c1, x.c); g.r2 = Math.max(g.r2, x.r); g.c2 = Math.max(g.c2, x.c); }
    }
    p.out = g;
    // ajusta a largura das colunas ao atualizar (como o Excel faz)
    if (g) {
      for (let c = g.c1; c <= g.c2; c++) {
        let w = 0;
        for (let rr = g.r1; rr <= g.r2; rr++) {
          const cell = S.cell(rr, c);
          if (!cell || cell.kind === 'empty') continue;
          const f = cell.fmt || {};
          const t = XL.formatValue(S.value(rr, c), f).text;
          w = Math.max(w, t.length * 7.2 + 18 + (f.indent || 0) * 14 + (p.btns.some((b) => b.r === rr && b.c === c) ? 18 : 0));
        }
        if (w) S.colW[c] = Math.round(Math.min(260, Math.max(64, w, S.colW[c] && p.keepWidths ? S.colW[c] : 0)));
      }
    }
    S.invalidate();
    return res;
  };

  // Atualiza todas as Tabelas Dinâmicas (Atualizar Tudo) ou só as de uma fonte
  XL.refreshPivots = (wb, filter) => {
    for (const s of wb.sheets) for (const p of s.pivots || []) if (!filter || filter(p)) s.renderPivot(p);
  };

  // Itens de um campo (para menus de filtro, segmentação de dados e linha do tempo), sem filtros aplicados
  XL.pivotItems = (wb, p, field) => {
    const src = source(wb, p);
    if (!src) return [];
    const cx = makeCtx(wb, p, src);
    const map = new Map();
    for (const r of src.recs) { const it = cx.item(field, r); if (!it) continue; const k = typeof it.key + ':' + it.key; if (!map.has(k)) map.set(k, it); }
    return [...map.values()].sort((a, b) => cmpKey(a.key, b.key));
  };
  // Itens que ainda têm dados depois dos outros filtros (a segmentação acinzenta os demais)
  XL.pivotLiveItems = (wb, p, host, field, except) => {
    const src = source(wb, p);
    if (!src) return new Set();
    const cx = makeCtx(wb, p, src);
    const extra = host.slicers.filter((s) => s.pivots.includes(p.id) && s !== except);
    const set = new Set();
    for (const r of src.recs) if (passes(cx, p, r, extra)) { const it = cx.item(field, r); if (it) set.add(it.label); }
    return set;
  };
  XL.pivotFields = (wb, p) => fieldList(wb, p);
  XL.pivotSource = source;
  XL.pivotValName = valName;
  XL.PIVOT_FN = FN;
  XL.pivotDateRange = (wb, p, field) => {
    const src = source(wb, p);
    if (!src) return null;
    const j = src.headers.findIndex((h) => norm(h) === norm(field));
    const nums = src.recs.map((r) => r[j]).filter(isNum);
    return nums.length ? { from: Math.min(...nums), to: Math.max(...nums) } : null;
  };
  XL.pivotCalcCheck = (wb, p, formula) => {
    const names = fieldList(wb, p).filter((f) => !f.calc && !f.virtual).map((f) => f.name);
    const v = evalCalc(wb, formula, names, () => 1);
    if (v && v.isXLError && v.code === '#NOME?') throw new Error('A fórmula usa um nome que não é um campo da Tabela Dinâmica.');
    return v;
  };

  // Dados para o Gráfico Dinâmico: categorias = itens das linhas (último nível), séries = valores ou itens das colunas
  XL.pivotChartData = (wb, sheet, p) => {
    const res = compute(wb, p, sheet);
    if (res.error || !res.vals.length) return { cats: [], series: [] };
    const leavesOf = (node, path) => (node.children ? node.children.flatMap((ch) => leavesOf(ch, path.concat(ch.item.label))) : [{ node, label: path.join(' - ') }]);
    const cats = res.rows.length ? leavesOf(res.root, []) : [{ node: res.root, label: 'Total' }];
    const series = [];
    const v0 = res.vals[0];
    if (res.cols.length) {
      res.colItems.forEach((ci) => series.push({ name: ci.label, vals: cats.map((c) => { const x = res.show(v0, res.agg(v0, c.node.recs, ci), c.node.recs, ci); return isNum(x) ? x : null; }) }));
    } else {
      res.vals.forEach((v) => series.push({ name: valName(v), vals: cats.map((c) => { const x = res.show(v, res.agg(v, c.node.recs, null), c.node.recs, null); return isNum(x) ? x : null; }) }));
    }
    return { cats: cats.map((c) => c.label), series };
  };

  // o gráfico ligado a uma Tabela Dinâmica usa os dados dela
  const baseChartData = P.chartData;
  P.chartData = function (chart) {
    if (chart && chart.pivot) {
      const p = this.pivotById(chart.pivot);
      return p ? XL.pivotChartData(this.wb, this, p) : { cats: [], series: [] };
    }
    return baseChartData.call(this, chart);
  };
})();
