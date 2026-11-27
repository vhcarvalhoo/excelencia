/* Recursos do Módulo 7 no modelo da planilha: tabelas (com referências estruturadas),
   formatação condicional, validação de dados e dados dos gráficos. */
(function () {
  'use strict';
  const XL = window.XL;
  const { Matrix, err, isErr, E } = XL.Formula;
  const K = XL.SHEET_KEY;
  const P = XL.Sheet.prototype;
  const inG = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
  XL.inRange = inG;

  // ======================= Tabelas =======================
  XL.TABLE_STYLES = {
    azul: { label: 'Azul', head: '#4472C4', band: '#D9E1F2', total: '#B4C6E7' },
    verde: { label: 'Verde', head: '#3E8E41', band: '#E2EFDA', total: '#C6E0B4' },
    laranja: { label: 'Laranja', head: '#C55A11', band: '#FCE4D6', total: '#F8CBAD' },
    cinza: { label: 'Cinza', head: '#595959', band: '#EDEDED', total: '#D9D9D9' },
  };
  P.tableAt = function (r, c) { return this.tables.find((t) => inG(t.g, r, c)) || null; };
  P.tableByName = function (name) {
    const k = XL.normName(name);
    return this.tables.find((t) => XL.normName(t.name) === k) || null;
  };
  P.tableDataRows = function (t) { return { r1: t.g.r1 + 1, r2: t.g.r2 - (t.totalRow ? 1 : 0) }; };
  P.tableHeader = function (t, c) { return String(this.textOf(t.g.r1, c)).trim(); };
  P.tableColIndex = function (t, name) {
    const k = XL.normName(String(name).trim());
    for (let c = t.g.c1; c <= t.g.c2; c++) if (XL.normName(this.tableHeader(t, c)) === k) return c;
    return -1;
  };
  // Valor de [@Coluna] (linha atual) ou de Tabela1[Coluna] (coluna inteira de dados)
  P.srefValue = function (node, cur, asRange) {
    let sh = this, t = null;
    if (node.table) {
      t = this.tableByName(node.table);
      if (!t && this.wb) for (const s of this.wb.sheets) { const x = s.tableByName(node.table); if (x) { t = x; sh = s; break; } }
    } else if (cur) t = this.tableAt(cur.r, cur.c);
    if (!t) return err(E.REF);
    const c = sh.tableColIndex(t, node.col);
    if (c < 0) return err(E.REF);
    const dr = sh.tableDataRows(t);
    if (node.thisRow) {
      if (!cur || sh !== this || cur.r < dr.r1 || cur.r > dr.r2) return err(E.VALUE);
      return asRange ? sh.rangeValues(cur.r, c, cur.r, c) : sh.value(cur.r, c);
    }
    if (dr.r2 < dr.r1) return new Matrix([[null]]);
    return sh.rangeValues(dr.r1, c, dr.r2, c);
  };
  // Texto da referência estruturada como o Excel escreve
  XL.srefText = (header, tableName) => {
    const simple = /^[A-Za-zÀ-ÿ0-9_]+$/.test(header);
    if (tableName) return `${tableName}[${header}]`;
    return simple ? `[@${header}]` : `[@[${header}]]`;
  };
  // Estilo visual de cada célula das tabelas (aplicado quando a célula não tem cor própria)
  P.tableStyles = function () {
    const out = new Map();
    for (const t of this.tables) {
      const st = XL.TABLE_STYLES[t.style] || XL.TABLE_STYLES.azul;
      const dr = this.tableDataRows(t);
      for (let c = t.g.c1; c <= t.g.c2; c++) {
        out.set(t.g.r1 * K + c, { fill: st.head, color: '#FFFFFF', b: true });
        for (let r = dr.r1; r <= dr.r2; r++) {
          if (t.banded !== false && (r - dr.r1) % 2 === 0) out.set(r * K + c, { fill: st.band });
          else out.set(r * K + c, {});
        }
        if (t.totalRow) out.set(t.g.r2 * K + c, { fill: st.total, b: true });
      }
    }
    return out;
  };
  P.nextTableName = function () {
    const all = this.wb ? [].concat(...this.wb.sheets.map((s) => s.tables)) : this.tables;
    let i = 1;
    while (all.some((t) => XL.normName(t.name) === 'TABELA' + i)) i++;
    return 'Tabela' + i;
  };

  // ======================= Formatação condicional =======================
  XL.CF_STYLES = {
    vermelho: { label: 'Preenchimento Vermelho Claro e Texto Vermelho Escuro', fill: '#FFC7CE', color: '#9C0006' },
    amarelo: { label: 'Preenchimento Amarelo e Texto Amarelo Escuro', fill: '#FFEB9C', color: '#9C5700' },
    verde: { label: 'Preenchimento Verde e Texto Verde Escuro', fill: '#C6EFCE', color: '#006100' },
    vermelhoClaro: { label: 'Preenchimento Vermelho Claro', fill: '#FFC7CE' },
    textoVermelho: { label: 'Texto Vermelho', color: '#C00000' },
    negrito: { label: 'Negrito', b: true },
    azul: { label: 'Preenchimento Azul Claro', fill: '#DDEBF7', color: '#1F4E79' },
  };
  const hex = (h) => { const n = parseInt(h.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  const mix = (a, b, t) => { const x = hex(a), y = hex(b); return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join(''); };
  XL.CF_SCALES = {
    gw: { label: 'Branco – Verde', colors: ['#FFFFFF', '#63BE7B'] },
    wb: { label: 'Branco – Azul', colors: ['#FFFFFF', '#5A8AC6'] },
    gyr: { label: 'Verde – Amarelo – Vermelho', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
    ryg: { label: 'Vermelho – Amarelo – Verde', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  };
  XL.CF_BARS = { azul: '#5A8AC6', verde: '#63BE7B', laranja: '#F4A460' };

  // valor de comparação: número, texto ou fórmula ("=$H$1")
  P.cfValue = function (v) {
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    if (s.startsWith('=')) {
      try { const r = XL.Formula.evaluate(XL.Formula.parse(XL.Formula.normalize(s).slice(1)), this.ctx); return XL.Formula.scalar(r); } catch (e) { return null; }
    }
    const p = XL.parseInput(s);
    return p.kind === 'number' ? p.value : s;
  };
  // Resultado: Map(chave → { fill, color, b, bar: { pct, color } })
  P.conditionalStyles = function () {
    const out = new Map();
    const put = (r, c, st) => {
      const k = r * K + c;
      const cur = out.get(k) || {};
      for (const p in st) if (cur[p] === undefined) cur[p] = st[p];
      out.set(k, cur);
    };
    for (const rule of this.cf) {
      const g = rule.g;
      const cells = [];
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        if (this.hiddenRows.has(r)) continue;
        cells.push({ r, c, v: this.value(r, c) });
      }
      const nums = cells.filter((x) => typeof x.v === 'number').map((x) => x.v);
      const style = XL.CF_STYLES[rule.style] || XL.CF_STYLES.vermelho;
      const cmp = (a, b) => XL.Formula.compareValues(a, b);
      let test = null;
      switch (rule.kind) {
        case 'gt': { const v = this.cfValue(rule.v1); test = (x) => x !== null && x !== '' && cmp(x, v) > 0; break; }
        case 'lt': { const v = this.cfValue(rule.v1); test = (x) => x !== null && x !== '' && cmp(x, v) < 0; break; }
        case 'eq': { const v = this.cfValue(rule.v1); test = (x) => x !== null && cmp(x, v) === 0; break; }
        case 'between': {
          const a = this.cfValue(rule.v1), b = this.cfValue(rule.v2);
          const lo = cmp(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
          test = (x) => x !== null && x !== '' && cmp(x, lo) >= 0 && cmp(x, hi) <= 0;
          break;
        }
        case 'text': { const s = String(rule.v1 || '').toLocaleLowerCase('pt-BR'); test = (x) => x !== null && String(isErr(x) ? x.code : x).toLocaleLowerCase('pt-BR').includes(s); break; }
        case 'dup': case 'uniq': {
          const key = (x) => (typeof x === 'string' ? 's:' + x.toLocaleLowerCase('pt-BR') : typeof x + ':' + x);
          const count = new Map();
          cells.forEach((x) => { if (x.v === null || x.v === '') return; const k2 = key(x.v); count.set(k2, (count.get(k2) || 0) + 1); });
          test = (x) => x !== null && x !== '' && (rule.kind === 'dup' ? count.get(key(x)) > 1 : count.get(key(x)) === 1);
          break;
        }
        case 'top': case 'bottom': {
          const n = Math.max(1, Math.trunc(Number(rule.v1) || 10));
          const sorted = nums.slice().sort((a, b) => (rule.kind === 'top' ? b - a : a - b));
          const lim = sorted[Math.min(n, sorted.length) - 1];
          test = (x) => typeof x === 'number' && (rule.kind === 'top' ? x >= lim : x <= lim);
          break;
        }
        case 'above': case 'below': {
          const avg = nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0;
          test = (x) => typeof x === 'number' && (rule.kind === 'above' ? x > avg : x < avg);
          break;
        }
        case 'formula': {
          for (const cell of cells) {
            let ok = false;
            try {
              const f = XL.Formula.shiftFormula(XL.Formula.normalize(rule.formula), cell.r - g.r1, cell.c - g.c1);
              const prev = this.curCell;
              this.curCell = { r: cell.r, c: cell.c };
              const v = XL.Formula.evaluate(XL.Formula.parse(f.slice(1)), this.ctx);
              this.curCell = prev;
              const b = XL.Formula.toBool(XL.Formula.scalar(v));
              ok = b === true;
            } catch (e) { ok = false; }
            if (ok) put(cell.r, cell.c, { fill: style.fill, color: style.color, b: style.b });
          }
          continue;
        }
        case 'bar': {
          if (!nums.length) continue;
          const lo = Math.min(0, ...nums), hi = Math.max(...nums);
          const color = XL.CF_BARS[rule.color] || XL.CF_BARS.azul;
          for (const cell of cells) {
            if (typeof cell.v !== 'number') continue;
            const pct = hi === lo ? 100 : Math.max(0, Math.min(100, ((cell.v - lo) / (hi - lo)) * 100));
            put(cell.r, cell.c, { bar: { pct, color } });
          }
          continue;
        }
        case 'scale': {
          if (!nums.length) continue;
          const sc = XL.CF_SCALES[rule.scale] || XL.CF_SCALES.gw;
          const lo = Math.min(...nums), hi = Math.max(...nums);
          for (const cell of cells) {
            if (typeof cell.v !== 'number') continue;
            const t = hi === lo ? 0.5 : (cell.v - lo) / (hi - lo);
            const cs = sc.colors;
            const fill = cs.length === 2 ? mix(cs[0], cs[1], t) : t < 0.5 ? mix(cs[0], cs[1], t * 2) : mix(cs[1], cs[2], (t - 0.5) * 2);
            put(cell.r, cell.c, { fill });
          }
          continue;
        }
      }
      if (!test) continue;
      for (const cell of cells) if (test(cell.v)) put(cell.r, cell.c, { fill: style.fill, color: style.color, b: style.b });
    }
    return out;
  };

  // ======================= Validação de dados =======================
  P.validationAt = function (r, c) {
    for (let i = this.dv.length - 1; i >= 0; i--) if (inG(this.dv[i].g, r, c)) return this.dv[i];
    return null;
  };
  P.listValues = function (rule) {
    const src = String(rule.source || '').trim();
    if (src.startsWith('=')) {
      try {
        const v = XL.Formula.evaluate(XL.Formula.parse(XL.Formula.normalize(src).slice(1)), this.ctx);
        const vals = v instanceof Matrix ? v.flat() : [v];
        return vals.filter((x) => x !== null && x !== '' && !isErr(x)).map((x) => XL.formatValue(x, {}).text);
      } catch (e) { return []; }
    }
    return src.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
  };
  // Confere o que foi digitado. Retorna true se for válido.
  P.checkValidation = function (rule, text) {
    if (String(text).trim() === '') return true;
    const p = XL.parseInput(text);
    if (rule.type === 'list') {
      const t = String(text).trim().toLocaleLowerCase('pt-BR');
      return this.listValues(rule).some((v) => v.toLocaleLowerCase('pt-BR') === t);
    }
    let x;
    if (rule.type === 'textlen') x = String(text).length;
    else {
      if (p.kind !== 'number') return false;
      x = p.value;
      if (rule.type === 'whole' && !Number.isInteger(x)) return false;
    }
    const a = this.cfValue(rule.v1), b = this.cfValue(rule.v2);
    switch (rule.op || 'between') {
      case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b);
      case 'notbetween': return x < Math.min(a, b) || x > Math.max(a, b);
      case 'eq': return x === a;
      case 'ne': return x !== a;
      case 'gt': return x > a;
      case 'lt': return x < a;
      case 'ge': return x >= a;
      case 'le': return x <= a;
    }
    return true;
  };

  // ======================= Dados dos gráficos =======================
  // Categorias = primeira coluna de texto; séries = colunas numéricas (nome no cabeçalho).
  P.chartData = function (chart) {
    const g = chart.g;
    const isNumCol = (c, r1) => { for (let r = r1; r <= g.r2; r++) if (typeof this.value(r, c) === 'number') return true; return false; };
    const headerRow = (() => {
      for (let c = g.c1; c <= g.c2; c++) { const v = this.value(g.r1, c); if (v !== null && typeof v !== 'string') return false; }
      return g.r2 > g.r1;
    })();
    const r1 = headerRow ? g.r1 + 1 : g.r1;
    const catCol = !isNumCol(g.c1, r1) && g.c2 > g.c1 ? g.c1 : null;
    const cats = [];
    for (let r = r1; r <= g.r2; r++) {
      if (this.hiddenRows.has(r)) continue;
      cats.push({ r, label: catCol !== null ? this.textOf(r, catCol) : String(r - r1 + 1) });
    }
    const series = [];
    for (let c = catCol !== null ? catCol + 1 : g.c1; c <= g.c2; c++) {
      if (!isNumCol(c, r1)) continue;
      const name = headerRow ? this.textOf(g.r1, c) : 'Série' + (series.length + 1);
      series.push({ name, col: c, vals: cats.map((x) => { const v = this.value(x.r, c); return typeof v === 'number' ? v : null; }) });
    }
    return { cats: cats.map((x) => x.label), series };
  };
})();
