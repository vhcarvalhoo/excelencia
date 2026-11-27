/* Modelo da planilha: células, formatos, cálculo, desfazer, preenchimento, classificação, filtro e
   configurações de página. Uma planilha pertence a uma pasta de trabalho (XL.Workbook). */
(function () {
  'use strict';
  const XL = window.XL;
  const { Matrix, err, isErr, E } = XL.Formula;
  const K = 1024; // chave numérica: r*K + c

  const clone = (o) => JSON.parse(JSON.stringify(o === undefined ? null : o));
  const defaultPage = () => ({ orientation: 'retrato', size: 'A4', margins: 'normal', fit: 'none', printArea: null, titleRows: null, printGridlines: false });

  class Sheet {
    constructor(rows = 100, cols = 26, name = 'Planilha1') {
      this.rows = rows;
      this.cols = cols;
      this.name = name;
      this.wb = null;
      this.cells = new Map();
      this.colW = {};
      this.cache = new Map();
      this.computing = new Set();
      this.circular = false;
      this.filter = null; // { g, crit: { [col]: {...} } }
      this.hiddenRows = new Set();
      this.freeze = null; // { rows, cols }
      this.page = defaultPage();
      this.showBreaks = false;
      this.sel = null;
      const self = this;
      const other = (s) => {
        if (!s) return self;
        if (self.wb) return self.wb.byName(s);
        return XL.normName(s) === XL.normName(self.name) ? self : null;
      };
      this.ctx = {
        cell: (r, c, s) => { const sh = other(s); return sh ? sh.value(r, c) : err(E.REF); },
        range: (r1, c1, r2, c2, s) => { const sh = other(s); return sh ? sh.rangeValues(r1, c1, r2, c2) : err(E.REF); },
        lastRow: (s) => { const sh = other(s); return sh ? sh.lastRow() : 0; },
      };
    }

    key(r, c) { return r * K + c; }
    cell(r, c) { return this.cells.get(r * K + c); }
    has(r, c) { const x = this.cell(r, c); return !!x && x.kind !== 'empty'; }

    ensure(r, c) {
      const k = r * K + c;
      let x = this.cells.get(k);
      if (!x) { x = { kind: 'empty', input: '', value: null, ast: null, fmt: {} }; this.cells.set(k, x); }
      return x;
    }
    cleanup(r, c) {
      const k = r * K + c;
      const x = this.cells.get(k);
      if (x && x.kind === 'empty' && Object.keys(x.fmt).length === 0) this.cells.delete(k);
    }
    clearCache() { this.cache.clear(); this.circular = false; }
    invalidate() { if (this.wb) this.wb.invalidateAll(); else this.clearCache(); }

    // Define o conteúdo como se o usuário tivesse digitado. Fórmulas inválidas lançam erro.
    setInput(r, c, text) {
      const cell = this.ensure(r, c);
      let p = XL.parseInput(text);
      if (cell.fmt.numFmt === 'texto' && p.kind !== 'empty') p = { kind: 'text', value: String(text) };
      if (p.kind === 'formula') {
        const norm = XL.Formula.normalize(String(text));
        let ast;
        try { ast = XL.Formula.parse(norm.slice(1)); } catch (e) {
          if (norm.includes('#REF!')) ast = { t: 'err', v: E.REF };
          else throw e;
        }
        Object.assign(cell, { kind: 'formula', input: norm, value: null, ast });
      } else if (p.kind === 'empty') {
        Object.assign(cell, { kind: 'empty', input: '', value: null, ast: null });
      } else {
        // como no Excel: digitar 15 numa célula com formato de porcentagem vira 15%
        if (p.kind === 'number' && !p.fmt && cell.fmt.numFmt === 'porcentagem') p.value /= 100;
        Object.assign(cell, { kind: p.kind, input: String(text), value: p.value, ast: null });
        if (p.fmt && (!cell.fmt.numFmt || cell.fmt.numFmt === 'geral')) Object.assign(cell.fmt, p.fmt);
      }
      this.cleanup(r, c);
      this.invalidate();
    }

    // Troca o texto de uma fórmula já existente (usado em ajustes automáticos).
    setFormulaText(cell, input) {
      cell.input = input;
      try { cell.ast = XL.Formula.parse(input.slice(1)); } catch (e) { cell.ast = { t: 'err', v: E.REF }; }
    }

    // Grava um valor já calculado (usado em "Colar Valores").
    setValue(r, c, v) {
      const cell = this.ensure(r, c);
      if (v === null || v === undefined) Object.assign(cell, { kind: 'empty', input: '', value: null, ast: null });
      else if (typeof v === 'number') Object.assign(cell, { kind: 'number', input: XL.fmtGeneral(v), value: v, ast: null });
      else if (typeof v === 'boolean') Object.assign(cell, { kind: 'bool', input: v ? 'VERDADEIRO' : 'FALSO', value: v, ast: null });
      else if (isErr(v)) Object.assign(cell, { kind: 'text', input: v.code, value: v.code, ast: null });
      else Object.assign(cell, { kind: 'text', input: String(v), value: String(v), ast: null });
      this.cleanup(r, c);
      this.invalidate();
    }

    // Copia o conteúdo constante de outra célula sem reinterpretar o texto.
    putCell(r, c, src) {
      if (!src || src.kind === 'empty') { this.clearContent(r, c); return; }
      if (src.kind === 'formula') { this.setInput(r, c, src.input); return; }
      const cell = this.ensure(r, c);
      Object.assign(cell, { kind: src.kind, input: src.input, value: src.value, ast: null });
      this.invalidate();
    }

    clearContent(r, c) {
      const cell = this.cell(r, c);
      if (!cell) return;
      Object.assign(cell, { kind: 'empty', input: '', value: null, ast: null });
      this.cleanup(r, c);
      this.invalidate();
    }

    setFmt(r, c, patch) {
      const cell = this.ensure(r, c);
      for (const k in patch) {
        if (patch[k] === null || patch[k] === undefined) delete cell.fmt[k];
        else cell.fmt[k] = typeof patch[k] === 'object' ? clone(patch[k]) : patch[k];
      }
      this.cleanup(r, c);
    }
    replaceFmt(r, c, fmt) {
      const cell = this.ensure(r, c);
      cell.fmt = clone(fmt || {});
      this.cleanup(r, c);
    }
    fmt(r, c) { const x = this.cell(r, c); return x ? x.fmt : {}; }

    lastRow() {
      let m = 0;
      for (const k of this.cells.keys()) m = Math.max(m, Math.floor(k / K));
      return Math.min(m, this.rows - 1);
    }
    // Retângulo que contém todas as células com conteúdo ou formatação visível.
    usedRange() {
      let g = null;
      for (const [k, x] of this.cells) {
        if (x.kind === 'empty' && !x.fmt.fill && !x.fmt.border) continue;
        const r = Math.floor(k / K), c = k % K;
        if (!g) g = { r1: r, c1: c, r2: r, c2: c };
        else { g.r1 = Math.min(g.r1, r); g.c1 = Math.min(g.c1, c); g.r2 = Math.max(g.r2, r); g.c2 = Math.max(g.c2, c); }
      }
      return g;
    }

    rangeValues(r1, c1, r2, c2) {
      if (r1 < 0 || c1 < 0 || r2 >= this.rows || c2 >= this.cols) return err(E.REF);
      const rows = [];
      let hidden = null;
      for (let r = r1; r <= r2; r++) {
        const row = [];
        for (let c = c1; c <= c2; c++) row.push(this.value(r, c));
        rows.push(row);
        if (this.hiddenRows.has(r)) { hidden = hidden || new Array(r2 - r1 + 1).fill(false); hidden[r - r1] = true; }
      }
      const m = new Matrix(rows);
      if (hidden) m.hiddenRows = hidden;
      return m;
    }

    value(r, c) {
      if (r < 0 || c < 0 || r >= this.rows || c >= this.cols) return err(E.REF);
      const k = r * K + c;
      if (this.cache.has(k)) return this.cache.get(k);
      const cell = this.cells.get(k);
      if (!cell || cell.kind === 'empty') return null;
      if (cell.kind !== 'formula') return cell.value;
      if (this.computing.has(k)) { this.circular = true; return 0; }
      this.computing.add(k);
      let v;
      try { v = XL.Formula.evaluate(cell.ast, this.ctx); } catch (e) { console.error(e); v = err(E.VALUE); }
      this.computing.delete(k);
      if (v instanceof Matrix) v = v.h && v.w ? v.rows[0][0] : err(E.CALC);
      if (v === null || v === undefined) v = 0;
      if (typeof v === 'number' && !isFinite(v)) v = err(E.NUM);
      this.cache.set(k, v);
      return v;
    }

    display(r, c, showFormulas) {
      const cell = this.cell(r, c);
      if (!cell || cell.kind === 'empty') return { text: '', align: null };
      if (showFormulas && cell.kind === 'formula') return { text: cell.input, align: 'left' };
      return XL.formatValue(this.value(r, c), cell.fmt);
    }

    // ---------- Desfazer ----------
    snapshot() {
      const cells = [];
      for (const [k, x] of this.cells) cells.push([k, x.kind, x.input, x.kind === 'formula' ? null : x.value, clone(x.fmt)]);
      return {
        cells, colW: Object.assign({}, this.colW), filter: clone(this.filter), hidden: [...this.hiddenRows],
        freeze: clone(this.freeze), page: clone(this.page), showBreaks: this.showBreaks,
      };
    }
    restore(s) {
      this.cells.clear();
      for (const [k, kind, input, value, fmt] of s.cells) {
        const x = { kind, input, value, ast: null, fmt: clone(fmt) };
        if (kind === 'formula') {
          try { x.ast = XL.Formula.parse(input.slice(1)); } catch (e) { x.ast = { t: 'err', v: E.NAME }; }
        }
        this.cells.set(k, x);
      }
      this.colW = Object.assign({}, s.colW);
      this.filter = clone(s.filter || null);
      this.hiddenRows = new Set(s.hidden || []);
      this.freeze = clone(s.freeze || null);
      this.page = Object.assign(defaultPage(), clone(s.page || {}));
      this.showBreaks = !!s.showBreaks;
      this.invalidate();
    }
    clear() {
      this.cells.clear(); this.colW = {}; this.filter = null; this.hiddenRows.clear(); this.freeze = null;
      this.page = defaultPage(); this.showBreaks = false; this.sel = null; this.invalidate();
    }

    // ---------- Região de dados ----------
    // Bloco contínuo de dados ao redor de uma célula (como Ctrl+* no Excel).
    currentRegion(r, c) {
      const g = { r1: r, c1: c, r2: r, c2: c };
      const any = (r1, c1, r2, c2) => {
        for (let i = Math.max(0, r1); i <= Math.min(this.rows - 1, r2); i++)
          for (let j = Math.max(0, c1); j <= Math.min(this.cols - 1, c2); j++) if (this.has(i, j)) return true;
        return false;
      };
      for (let guard = 0; guard < 500; guard++) {
        let grew = false;
        if (g.r1 > 0 && any(g.r1 - 1, g.c1 - 1, g.r1 - 1, g.c2 + 1)) { g.r1--; grew = true; }
        if (g.r2 < this.rows - 1 && any(g.r2 + 1, g.c1 - 1, g.r2 + 1, g.c2 + 1)) { g.r2++; grew = true; }
        if (g.c1 > 0 && any(g.r1 - 1, g.c1 - 1, g.r2 + 1, g.c1 - 1)) { g.c1--; grew = true; }
        if (g.c2 < this.cols - 1 && any(g.r1 - 1, g.c2 + 1, g.r2 + 1, g.c2 + 1)) { g.c2++; grew = true; }
        if (!grew) break;
      }
      return g;
    }
    // A primeira linha parece um cabeçalho? (textos, em negrito ou de tipo diferente dos dados)
    detectHeader(g) {
      if (g.r2 <= g.r1) return false;
      let allText = true, anyBold = false, differs = false;
      for (let c = g.c1; c <= g.c2; c++) {
        const x = this.cell(g.r1, c);
        if (x && x.kind !== 'empty' && x.kind !== 'text') allText = false;
        if (x && x.fmt.b) anyBold = true;
        const below = this.cell(g.r1 + 1, c);
        if (below && below.kind !== 'empty' && below.kind !== 'text') differs = true;
      }
      return allText && (anyBold || differs);
    }

    // Lê linhas inteiras (só as colunas c1..c2) para mover depois.
    takeRow(r, c1, c2) {
      const out = [];
      for (let c = c1; c <= c2; c++) {
        const x = this.cell(r, c);
        out.push(x ? { kind: x.kind, input: x.input, value: x.value, fmt: clone(x.fmt) } : null);
      }
      return out;
    }
    putRow(r, c1, row, fromRow) {
      row.forEach((src, i) => {
        const c = c1 + i;
        if (!src) { this.clearContent(r, c); this.replaceFmt(r, c, {}); return; }
        if (src.kind === 'formula') this.setInput(r, c, XL.Formula.shiftFormula(src.input, r - fromRow, 0));
        else this.putCell(r, c, src);
        this.replaceFmt(r, c, src.fmt);
      });
    }

    // ---------- Classificar ----------
    // keys: [{ col (índice absoluto), desc }]
    sortRange(g, keys, hasHeader) {
      const start = g.r1 + (hasHeader ? 1 : 0);
      if (start > g.r2) return;
      const rows = [];
      for (let r = start; r <= g.r2; r++) {
        rows.push({ r, data: this.takeRow(r, g.c1, g.c2), vals: keys.map((k) => this.value(r, k.col)) });
      }
      rows.sort((a, b) => {
        for (let i = 0; i < keys.length; i++) {
          const x = a.vals[i], y = b.vals[i];
          const ex = x === null || x === '', ey = y === null || y === '';
          if (ex || ey) { if (ex && ey) continue; return ex ? 1 : -1; } // vazios sempre no fim
          let cmp = XL.Formula.compareValues(isErr(x) ? 'zzzz' : x, isErr(y) ? 'zzzz' : y);
          if (keys[i].desc) cmp = -cmp;
          if (cmp) return cmp;
        }
        return a.r - b.r;
      });
      rows.forEach((row, i) => this.putRow(start + i, g.c1, row.data, row.r));
      this.invalidate();
    }

    // ---------- Remover duplicatas ----------
    removeDuplicates(g, cols, hasHeader) {
      const start = g.r1 + (hasHeader ? 1 : 0);
      const seen = new Set();
      const keep = [];
      let removed = 0;
      for (let r = start; r <= g.r2; r++) {
        const key = cols.map((c) => {
          const v = this.value(r, c);
          return typeof v === 'string' ? 's:' + v.trim().toLocaleLowerCase('pt-BR') : typeof v + ':' + String(v);
        }).join('\u0001');
        if (seen.has(key)) { removed++; continue; }
        seen.add(key);
        keep.push({ r, data: this.takeRow(r, g.c1, g.c2) });
      }
      if (removed) {
        keep.forEach((row, i) => this.putRow(start + i, g.c1, row.data, row.r));
        for (let r = start + keep.length; r <= g.r2; r++) {
          for (let c = g.c1; c <= g.c2; c++) { this.clearContent(r, c); this.replaceFmt(r, c, {}); }
        }
        this.invalidate();
      }
      return { removed, remaining: keep.length };
    }

    // ---------- Filtro (AutoFiltro) ----------
    textOf(r, c) { return this.display(r, c).text; }
    passes(r, crit, c) {
      if (!crit) return true;
      if (crit.type === 'values') {
        const t = this.textOf(r, c);
        if (t === '') return !!crit.blanks;
        return crit.values.includes(t);
      }
      // condição personalizada
      const v = this.value(r, c);
      const t = this.textOf(r, c).toLocaleLowerCase('pt-BR');
      const q = String(crit.value === undefined ? '' : crit.value);
      const ql = q.toLocaleLowerCase('pt-BR');
      const qn = XL.parseInput(q);
      const num = qn.kind === 'number' ? qn.value : null;
      const cmp = num !== null && typeof v === 'number' ? Math.sign(v - num) : t.localeCompare(ql, 'pt-BR', { sensitivity: 'base' });
      switch (crit.op) {
        case '=': return num !== null && typeof v === 'number' ? v === num : t === ql;
        case '<>': return num !== null && typeof v === 'number' ? v !== num : t !== ql;
        case '>': return (num === null || typeof v === 'number') && cmp > 0;
        case '>=': return (num === null || typeof v === 'number') && cmp >= 0;
        case '<': return (num === null || typeof v === 'number') && cmp < 0;
        case '<=': return (num === null || typeof v === 'number') && cmp <= 0;
        case 'contem': return t.includes(ql);
        case 'naocontem': return !t.includes(ql);
        case 'comeca': return t.startsWith(ql);
        case 'termina': return t.endsWith(ql);
      }
      return true;
    }
    rowPasses(r, exceptCol) {
      const f = this.filter;
      if (!f) return true;
      for (const col in f.crit) {
        if (Number(col) === exceptCol) continue;
        if (!this.passes(r, f.crit[col], Number(col))) return false;
      }
      return true;
    }
    // Estende o filtro para baixo se novas linhas foram digitadas logo abaixo.
    extendFilter() {
      const g = this.filter.g;
      const rowHas = (r) => { for (let c = g.c1; c <= g.c2; c++) if (this.has(r, c)) return true; return false; };
      while (g.r2 + 1 < this.rows && rowHas(g.r2 + 1)) g.r2++;
    }
    applyFilter() {
      this.hiddenRows.clear();
      if (this.filter) {
        this.extendFilter();
        const g = this.filter.g;
        for (let r = g.r1 + 1; r <= g.r2; r++) if (!this.rowPasses(r)) this.hiddenRows.add(r);
      }
      this.invalidate();
    }
    filterStats() {
      if (!this.filter) return null;
      const g = this.filter.g;
      const total = g.r2 - g.r1;
      let shown = 0;
      for (let r = g.r1 + 1; r <= g.r2; r++) if (!this.hiddenRows.has(r)) shown++;
      return { total, shown };
    }

    // Ajusta estados guardados por linha/coluna após inserir/excluir.
    adjustState(op) {
      const fix = (g) => {
        if (!g) return null;
        const a = op.axis === 'row' ? [g.r1 + 1, g.r2 + 1] : [g.c1 + 1, g.c2 + 1];
        const s = XL.Formula.adjustSpan(a[0], a[1], op);
        if (!s) return null;
        return op.axis === 'row' ? Object.assign({}, g, { r1: s[0] - 1, r2: s[1] - 1 }) : Object.assign({}, g, { c1: s[0] - 1, c2: s[1] - 1 });
      };
      if (this.filter) {
        const h = this.filter.g.r1 + 1; // linha do cabeçalho (1-based)
        const headerDeleted = op.kind === 'delete' && op.axis === 'row' && h >= op.at && h <= op.at + op.n - 1;
        const ng = fix(this.filter.g);
        if (!ng || headerDeleted) this.filter = null;
        else {
          if (op.axis === 'col') {
            const crit = {};
            for (const col in this.filter.crit) {
              const v = XL.Formula.adjustSpan(Number(col) + 1, Number(col) + 1, op);
              if (v) crit[v[0] - 1] = this.filter.crit[col];
            }
            this.filter.crit = crit;
          }
          this.filter.g = ng;
        }
      }
      this.page.printArea = fix(this.page.printArea);
      if (op.axis === 'row' && this.page.titleRows) {
        const t = fix({ r1: this.page.titleRows.r1, r2: this.page.titleRows.r2, c1: 0, c2: 0 });
        this.page.titleRows = t ? { r1: t.r1, r2: t.r2 } : null;
      }
      if (op.axis === 'col') {
        const w = {};
        for (const c in this.colW) {
          const v = XL.Formula.adjustSpan(Number(c) + 1, Number(c) + 1, op);
          if (v) w[v[0] - 1] = this.colW[c];
        }
        this.colW = w;
      }
      this.applyFilter();
    }

    // ---------- Alça de preenchimento ----------
    // src e dst são intervalos; dst contém src e cresce na direção dir.
    fill(src, dst, dir) {
      const vertical = dir === 'down' || dir === 'up';
      const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
      for (let line = lines[0]; line <= lines[1]; line++) {
        let srcPos = [];
        if (vertical) for (let r = src.r1; r <= src.r2; r++) srcPos.push([r, line]);
        else for (let c = src.c1; c <= src.c2; c++) srcPos.push([line, c]);
        const tgtPos = [];
        if (dir === 'down') for (let r = src.r2 + 1; r <= dst.r2; r++) tgtPos.push([r, line]);
        if (dir === 'up') for (let r = src.r1 - 1; r >= dst.r1; r--) tgtPos.push([r, line]);
        if (dir === 'right') for (let c = src.c2 + 1; c <= dst.c2; c++) tgtPos.push([line, c]);
        if (dir === 'left') for (let c = src.c1 - 1; c >= dst.c1; c--) tgtPos.push([line, c]);
        if (dir === 'up' || dir === 'left') srcPos = srcPos.reverse();
        this.fillLine(srcPos, tgtPos);
      }
      this.invalidate();
    }

    fillLine(srcPos, tgtPos) {
      const n = srcPos.length;
      const srcCells = srcPos.map(([r, c]) => this.cell(r, c));
      const kinds = srcCells.map((x) => (x ? x.kind : 'empty'));
      const allNum = kinds.every((k) => k === 'number');
      const first = srcCells[0];
      let gen = null;

      if (allNum && n >= 2) {
        const vals = srcCells.map((x) => x.value);
        const step = (vals[n - 1] - vals[0]) / (n - 1);
        gen = (i) => ({ value: vals[n - 1] + step * (i + 1) });
      } else if (allNum && n === 1 && first.fmt.numFmt === 'data') {
        gen = (i) => ({ value: first.value + i + 1 });
      } else if (kinds.every((k) => k === 'text')) {
        const series = textSeries(srcCells.map((x) => x.input));
        if (series) gen = (i) => ({ text: series(i) });
      }

      tgtPos.forEach(([r, c], i) => {
        if (r < 0 || c < 0 || r >= this.rows || c >= this.cols) return;
        const si = i % n;
        const s = srcCells[si];
        const [sr, sc] = srcPos[si];
        const fmt = s ? clone(s.fmt) : {};
        if (gen) {
          const g = gen(i);
          if (g.text !== undefined) this.setInput(r, c, g.text);
          else this.setValue(r, c, Number(g.value.toPrecision(15)));
        } else if (!s || s.kind === 'empty') {
          this.clearContent(r, c);
        } else if (s.kind === 'formula') {
          this.setInput(r, c, XL.Formula.shiftFormula(s.input, r - sr, c - sc));
        } else if (s.kind === 'text' && n === 1 && /\d+$/.test(s.input)) {
          const m = /^(.*?)(\d+)$/.exec(s.input);
          this.setInput(r, c, m[1] + String(parseInt(m[2], 10) + i + 1).padStart(m[2].length, '0'));
        } else {
          this.putCell(r, c, s);
        }
        this.replaceFmt(r, c, fmt);
      });
    }
  }

  // Séries de texto conhecidas: meses e dias da semana (completos ou abreviados).
  const LISTS = [
    XL.MONTHS,
    ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'],
    XL.WEEKDAYS,
    ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'],
  ];
  function textSeries(texts) {
    const norm = (s) => XL.stripAccents(String(s).trim().toLowerCase());
    for (const list of LISTS) {
      const nl = list.map(norm);
      const idx = texts.map((t) => nl.indexOf(norm(t)));
      if (idx.some((i) => i < 0)) continue;
      const last = idx[idx.length - 1];
      const step = idx.length >= 2 ? (idx[1] - idx[0] + nl.length) % nl.length || 1 : 1;
      const sample = texts[texts.length - 1].trim();
      const caseFn = sample === sample.toUpperCase()
        ? (s) => s.toUpperCase()
        : sample[0] === sample[0].toUpperCase()
          ? (s) => s[0].toUpperCase() + s.slice(1)
          : (s) => s;
      return (i) => caseFn(list[(last + step * (i + 1)) % list.length]);
    }
    return null;
  }

  XL.Sheet = Sheet;
  XL.SHEET_KEY = K;
})();
