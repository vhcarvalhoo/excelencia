/* Motor de fórmulas: tokenização, análise sintática, avaliação e utilitários de edição.
   Segue a convenção do Excel em português: separador de argumentos ";" e vírgula decimal. */
(function () {
  'use strict';
  const XL = window.XL;

  // ---------- Tipos de valor ----------
  class XLError {
    constructor(code, msg) { this.code = code; this.msg = msg || ''; }
    get isXLError() { return true; }
    toString() { return this.code; }
  }
  const E = { DIV0: '#DIV/0!', VALUE: '#VALOR!', REF: '#REF!', NAME: '#NOME?', NA: '#N/D', NUM: '#NÚM!', NULL: '#NULO!', CALC: '#CALC!' };
  const err = (code, msg) => new XLError(code, msg);
  const isErr = (v) => v instanceof XLError;

  class Matrix {
    constructor(rows) { this.rows = rows; }
    get h() { return this.rows.length; }
    get w() { return this.rows[0] ? this.rows[0].length : 0; }
    flat() { return [].concat(...this.rows); }
    map(f) { return new Matrix(this.rows.map((row) => row.map(f))); }
  }

  class FormulaSyntaxError extends Error {
    constructor(msg) { super(msg); this.name = 'FormulaSyntaxError'; }
  }

  // ---------- Tokenização ----------
  const REF_PART = '\\$?[A-Za-z]{1,3}\\$?\\d+';
  const RE_RANGE = new RegExp('^(' + REF_PART + ')(?::(' + REF_PART + '))?(?![A-Za-zÀ-ÿ0-9_.(])');
  const RE_COLRANGE = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![A-Za-zÀ-ÿ0-9_.(])/;
  const RE_NUM = /^\d+(?:[.,]\d+)?(?:[eE][+-]?\d+)?/;
  const RE_IDENT = /^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*/;
  const RE_ERR = /^#(?:DIV\/0!|VALOR!|REF!|NOME\?|N\/D|NÚM!|NULO!|CALC!)/i;
  const RE_SHEET = /^(?:'((?:[^']|'')+)'|([A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*))!/;
  const RE_SREF = /^([A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*)?\[(?:(@)(?:\[([^[\]]+)\]|([^[\]]+))|([^[\]@#][^[\]]*))\]/;
  const OPS2 = ['<>', '<=', '>='];
  const OPS1 = '+-*/^&=<>%():';

  // Nome de planilha como aparece numa fórmula (com aspas simples quando precisa).
  XL.quoteSheet = (name) =>
    /^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
      ? name
      : "'" + String(name).replace(/'/g, "''") + "'";

  function tokenize(src, lenient) {
    const toks = [];
    let i = 0;
    const n = src.length;
    while (i < n) {
      const ch = src[i];
      const rest = src.slice(i);
      const pos = i;
      let m;
      if (/\s/.test(ch)) {
        m = /^\s+/.exec(rest);
        toks.push({ type: 'ws', text: m[0], pos });
        i += m[0].length;
        continue;
      }
      if (ch === '"') {
        let j = i + 1, val = '', closed = false;
        while (j < n) {
          if (src[j] === '"') {
            if (src[j + 1] === '"') { val += '"'; j += 2; continue; }
            closed = true;
            break;
          }
          val += src[j];
          j++;
        }
        if (!closed) {
          if (lenient) { toks.push({ type: 'str', text: src.slice(i), value: val, pos, open: true }); break; }
          throw new FormulaSyntaxError('Um texto foi aberto com aspas (") mas não foi fechado.');
        }
        toks.push({ type: 'str', text: src.slice(i, j + 1), value: val, pos });
        i = j + 1;
        continue;
      }
      if (ch === '#' && (m = RE_ERR.exec(rest))) {
        toks.push({ type: 'err', text: m[0], value: m[0].toUpperCase(), pos });
        i += m[0].length;
        continue;
      }
      // referência estruturada de tabela: [@Qtd]  [@[Preço unit.]]  Tabela1[Total]
      if ((ch === '[' || /[A-Za-zÀ-ÿ_]/.test(ch)) && (m = RE_SREF.exec(rest))) {
        toks.push({ type: 'sref', text: m[0], table: m[1] || null, thisRow: !!m[2], col: (m[3] || m[4] || m[5]).trim(), pos });
        i += m[0].length;
        continue;
      }
      // referência a outra planilha: Planilha2!A1  ou  'Minha planilha'!A1:B5
      if ((ch === "'" || /[A-Za-zÀ-ÿ_]/.test(ch)) && (m = RE_SHEET.exec(rest))) {
        const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
        const after = rest.slice(m[0].length);
        let m2;
        if ((m2 = RE_RANGE.exec(after))) {
          toks.push({ type: m2[2] ? 'range' : 'ref', text: m[0] + m2[0], refText: m2[0], prefix: m[0], sheet, pos });
          i += m[0].length + m2[0].length;
          continue;
        }
        if ((m2 = RE_COLRANGE.exec(after))) {
          toks.push({ type: 'colrange', text: m[0] + m2[0], refText: m2[0], prefix: m[0], sheet, pos });
          i += m[0].length + m2[0].length;
          continue;
        }
        if (lenient) { toks.push({ type: 'sheetprefix', text: m[0], sheet, pos }); i += m[0].length; continue; }
        throw new FormulaSyntaxError(`Depois de "${m[0]}" deve vir uma referência de célula, como ${m[0]}A1.`);
      }
      if ((m = RE_RANGE.exec(rest))) {
        toks.push({ type: m[2] ? 'range' : 'ref', text: m[0], refText: m[0], prefix: '', pos });
        i += m[0].length;
        continue;
      }
      if ((m = RE_COLRANGE.exec(rest))) {
        toks.push({ type: 'colrange', text: m[0], refText: m[0], prefix: '', pos });
        i += m[0].length;
        continue;
      }
      if ((m = RE_NUM.exec(rest))) {
        toks.push({ type: 'num', text: m[0], value: parseFloat(m[0].replace(',', '.')), pos });
        i += m[0].length;
        continue;
      }
      if ((m = RE_IDENT.exec(rest))) {
        const word = m[0];
        let j = i + word.length;
        while (j < n && /\s/.test(src[j])) j++;
        const up = XL.normName(word);
        if (src[j] === '(') toks.push({ type: 'func', text: word, name: up, pos });
        else if (up === 'VERDADEIRO' || up === 'TRUE') toks.push({ type: 'bool', text: word, value: true, pos });
        else if (up === 'FALSO' || up === 'FALSE') toks.push({ type: 'bool', text: word, value: false, pos });
        else toks.push({ type: 'name', text: word, pos });
        i += word.length;
        continue;
      }
      const two = src.substr(i, 2);
      if (OPS2.includes(two)) { toks.push({ type: 'op', text: two, pos }); i += 2; continue; }
      if (ch === ';' || ch === ',') { toks.push({ type: 'sep', text: ch, pos }); i++; continue; }
      if (OPS1.includes(ch)) { toks.push({ type: 'op', text: ch, pos }); i++; continue; }
      if (lenient) { toks.push({ type: 'unknown', text: ch, pos }); i++; continue; }
      throw new FormulaSyntaxError(`O caractere "${ch}" não pode ser usado nesta posição da fórmula.`);
    }
    return toks;
  }

  // ---------- Análise sintática (descendente recursiva) ----------
  function parse(src) {
    const toks = tokenize(src).filter((t) => t.type !== 'ws');
    let p = 0;
    const isOp = (...ops) => { const t = toks[p]; return !!t && t.type === 'op' && ops.includes(t.text); };
    const isSep = () => { const t = toks[p]; return !!t && t.type === 'sep'; };

    function comparison() {
      let l = concat();
      while (isOp('=', '<>', '<', '>', '<=', '>=')) { const op = toks[p++].text; l = { t: 'bin', op, l, r: concat() }; }
      return l;
    }
    function concat() {
      let l = additive();
      while (isOp('&')) { p++; l = { t: 'bin', op: '&', l, r: additive() }; }
      return l;
    }
    function additive() {
      let l = mult();
      while (isOp('+', '-')) { const op = toks[p++].text; l = { t: 'bin', op, l, r: mult() }; }
      return l;
    }
    function mult() {
      let l = power();
      while (isOp('*', '/')) { const op = toks[p++].text; l = { t: 'bin', op, l, r: power() }; }
      return l;
    }
    function power() {
      let l = unary();
      while (isOp('^')) { p++; l = { t: 'bin', op: '^', l, r: unary() }; }
      return l;
    }
    function unary() {
      if (isOp('-')) { p++; return { t: 'neg', e: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      return postfix();
    }
    function postfix() {
      let e = primary();
      while (isOp('%')) { p++; e = { t: 'pct', e }; }
      return e;
    }
    function primary() {
      const t = toks[p];
      if (!t) throw new FormulaSyntaxError('A fórmula está incompleta: falta um valor no final.');
      p++;
      switch (t.type) {
        case 'num': return { t: 'num', v: t.value };
        case 'str': return { t: 'str', v: t.value };
        case 'bool': return { t: 'bool', v: t.value };
        case 'err': return { t: 'err', v: t.value };
        case 'ref': {
          const a = XL.parseRef(t.refText);
          return a ? { t: 'ref', r: a.r, c: a.c, sheet: t.sheet } : { t: 'err', v: E.REF };
        }
        case 'range': {
          const g = XL.parseRange(t.refText);
          return g ? Object.assign({ t: 'range', sheet: t.sheet }, g) : { t: 'err', v: E.REF };
        }
        case 'colrange': {
          const [a, b] = t.refText.replace(/\$/g, '').split(':');
          const c1 = XL.colToNum(a) - 1, c2 = XL.colToNum(b) - 1;
          return { t: 'colrange', c1: Math.min(c1, c2), c2: Math.max(c1, c2), sheet: t.sheet };
        }
        case 'name': return { t: 'name', name: t.text };
        case 'sref': return { t: 'sref', table: t.table, col: t.col, thisRow: t.thisRow };
        case 'func': {
          const fname = t.text.toUpperCase();
          if (!isOp('(')) throw new FormulaSyntaxError(`Faltou abrir o parêntese da função ${fname}.`);
          p++;
          const args = [];
          if (isOp(')')) { p++; }
          else {
            for (;;) {
              if (isSep() || isOp(')')) args.push({ t: 'missing' });
              else args.push(comparison());
              if (isSep()) { p++; continue; }
              if (isOp(')')) { p++; break; }
              if (!toks[p]) throw new FormulaSyntaxError(`Faltou fechar o parêntese da função ${fname}.`);
              throw new FormulaSyntaxError(`Separe os argumentos da função ${fname} com ponto e vírgula (;).`);
            }
          }
          return { t: 'func', name: t.name, args };
        }
        case 'op':
          if (t.text === '(') {
            const e = comparison();
            if (!isOp(')')) throw new FormulaSyntaxError('Faltou fechar um parêntese.');
            p++;
            return { t: 'paren', e };
          }
          throw new FormulaSyntaxError(`O operador "${t.text}" está em uma posição inválida.`);
        case 'sep':
          throw new FormulaSyntaxError('O ponto e vírgula (;) só separa argumentos dentro de uma função.');
        default:
          throw new FormulaSyntaxError(`"${t.text}" não é válido aqui.`);
      }
    }

    if (!toks.length) throw new FormulaSyntaxError('A fórmula está vazia. Digite algo depois do sinal de igual (=).');
    const ast = comparison();
    if (p < toks.length) {
      const t = toks[p];
      if (t.type === 'op' && t.text === ')') throw new FormulaSyntaxError('Há um parêntese ")" fechando sem um "(" correspondente.');
      throw new FormulaSyntaxError(`Há algo sobrando perto de "${t.text}". Talvez falte um operador (+, -, *, /) ou um ponto e vírgula.`);
    }
    return ast;
  }

  // Verifica número de argumentos (o Excel não deixa confirmar se estiver errado).
  function validate(ast) {
    let msg = null;
    (function walk(n) {
      if (!n || msg) return;
      if (n.t === 'func') {
        const def = XL.Functions.get(n.name);
        if (def) {
          if (n.args.length < def.min) msg = `Você inseriu poucos argumentos para a função ${def.name}. Sintaxe: ${def.syntax}`;
          else if (n.args.length > def.max) msg = `Você inseriu argumentos demais para a função ${def.name}. Sintaxe: ${def.syntax}`;
        }
        n.args.forEach(walk);
      } else if (n.t === 'bin') { walk(n.l); walk(n.r); }
      else if (n.e) walk(n.e);
    })(ast);
    return msg;
  }

  // ---------- Coerção ----------
  function scalar(v) {
    if (v instanceof Matrix) return v.h === 1 && v.w === 1 ? v.rows[0][0] : err(E.VALUE);
    return v;
  }
  function toNum(v) {
    v = scalar(v);
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isErr(v)) return v;
    const s = String(v);
    if (s.trim() === '') return err(E.VALUE);
    const p = XL.parseInput(s);
    if (p.kind === 'number') return p.value;
    return err(E.VALUE);
  }
  function toStr(v) {
    v = scalar(v);
    if (v === null || v === undefined) return '';
    if (isErr(v)) return v;
    if (typeof v === 'number') return XL.fmtGeneral(v);
    if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
    return String(v);
  }
  function toBool(v) {
    v = scalar(v);
    if (v === null || v === undefined) return false;
    if (isErr(v)) return v;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    const up = XL.normName(String(v).trim());
    if (up === 'VERDADEIRO' || up === 'TRUE') return true;
    if (up === 'FALSO' || up === 'FALSE') return false;
    return err(E.VALUE);
  }
  const rank = (v) => (typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0);
  function compareValues(x, y) {
    if (x === null || x === undefined) x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0;
    if (y === null || y === undefined) y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0;
    const rx = rank(x), ry = rank(y);
    if (rx !== ry) return rx < ry ? -1 : 1;
    if (rx === 2) return Math.sign(x.localeCompare(y, 'pt-BR', { sensitivity: 'accent' }));
    if (rx === 3) return x === y ? 0 : x ? 1 : -1;
    return x < y ? -1 : x > y ? 1 : 0;
  }

  // ---------- Avaliação ----------
  function lift1(v, f) { return v instanceof Matrix ? v.map(f) : f(v); }
  function lift2(a, b, f) {
    const am = a instanceof Matrix, bm = b instanceof Matrix;
    if (!am && !bm) return f(a, b);
    const A = am ? a : new Matrix([[a]]);
    const B = bm ? b : new Matrix([[b]]);
    const h = Math.max(A.h, B.h), w = Math.max(A.w, B.w);
    const pick = (M, i, j) => {
      const r = M.h === 1 ? 0 : i, c = M.w === 1 ? 0 : j;
      return r < M.h && c < M.w ? M.rows[r][c] : err(E.NA);
    };
    const rows = [];
    for (let i = 0; i < h; i++) {
      const row = [];
      for (let j = 0; j < w; j++) row.push(f(pick(A, i, j), pick(B, i, j)));
      rows.push(row);
    }
    return h === 1 && w === 1 ? rows[0][0] : new Matrix(rows);
  }
  function binScalar(op, x, y) {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    switch (op) {
      case '+': case '-': case '*': case '/': case '^': {
        const m = toNum(x); if (isErr(m)) return m;
        const n = toNum(y); if (isErr(n)) return n;
        if (op === '+') return m + n;
        if (op === '-') return m - n;
        if (op === '*') return m * n;
        if (op === '/') return n === 0 ? err(E.DIV0) : m / n;
        const r = Math.pow(m, n);
        return isFinite(r) ? r : err(E.NUM);
      }
      case '&': return toStr(x) + toStr(y);
      default: {
        const c = compareValues(x, y);
        switch (op) {
          case '=': return c === 0;
          case '<>': return c !== 0;
          case '<': return c < 0;
          case '>': return c > 0;
          case '<=': return c <= 0;
          case '>=': return c >= 0;
        }
      }
    }
    return err(E.VALUE);
  }

  function evaluate(node, ctx) {
    switch (node.t) {
      case 'num': case 'str': case 'bool': return node.v;
      case 'err': return err(node.v);
      case 'missing': return null;
      case 'paren': return evaluate(node.e, ctx);
      case 'ref': return ctx.cell(node.r, node.c, node.sheet);
      case 'range': return ctx.range(node.r1, node.c1, node.r2, node.c2, node.sheet);
      case 'colrange': return ctx.range(0, node.c1, ctx.lastRow(node.sheet), node.c2, node.sheet);
      case 'name': return err(E.NAME);
      case 'sref': return ctx.sref ? ctx.sref(node, false) : err(E.REF);
      case 'neg': return lift1(evaluate(node.e, ctx), (v) => { const n = toNum(v); return isErr(n) ? n : -n; });
      case 'pct': return lift1(evaluate(node.e, ctx), (v) => { const n = toNum(v); return isErr(n) ? n : n / 100; });
      case 'bin': return lift2(evaluate(node.l, ctx), evaluate(node.r, ctx), (x, y) => binScalar(node.op, x, y));
      case 'func': {
        const def = XL.Functions.get(node.name);
        if (!def) return err(E.NAME);
        if (node.args.length < def.min || node.args.length > def.max) return err(E.VALUE);
        const args = node.args.map((a) => (a.t === 'ref' ? ctx.range(a.r, a.c, a.r, a.c, a.sheet)
          : a.t === 'sref' && a.thisRow && ctx.sref ? ctx.sref(a, true) : evaluate(a, ctx)));
        try {
          const v = def.fn(args, ctx);
          return v === undefined ? err(E.VALUE) : v;
        } catch (e) {
          if (e instanceof XLError) return e;
          console.error(e);
          return err(E.VALUE);
        }
      }
    }
    return err(E.VALUE);
  }

  // ---------- Utilitários de edição ----------
  // Padroniza a fórmula como o Excel faz ao confirmar: nomes em maiúsculas, ";" como separador,
  // e fecha parênteses que ficaram abertos.
  // opts.sheetName(nome) devolve o nome oficial da planilha (para corrigir maiúsculas/aspas).
  function normalize(text, opts) {
    let toks;
    try { toks = tokenize(text.slice(1)); } catch (e) { return text; }
    let depth = 0;
    const body = toks.map((t) => {
      switch (t.type) {
        case 'func': { const def = XL.Functions.get(t.name); return def ? def.name : t.text.toUpperCase(); }
        case 'ref': case 'range': case 'colrange': {
          if (!t.sheet) return t.refText.toUpperCase();
          const canon = (opts && opts.sheetName && opts.sheetName(t.sheet)) || t.sheet;
          return XL.quoteSheet(canon) + '!' + t.refText.toUpperCase();
        }
        case 'bool': return t.value ? 'VERDADEIRO' : 'FALSO';
        case 'err': return t.value;
        case 'sep': return ';';
        case 'op':
          if (t.text === '(') depth++;
          else if (t.text === ')') depth--;
          return t.text;
        default: return t.text;
      }
    }).join('');
    return '=' + body + (depth > 0 ? ')'.repeat(depth) : '');
  }

  function shiftRef(ref, dr, dc) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref);
    if (!m) return ref;
    let c = XL.colToNum(m[2]);
    let r = parseInt(m[4], 10);
    if (!m[1]) c += dc;
    if (!m[3]) r += dr;
    if (c < 1 || r < 1) return '#REF!';
    return m[1] + XL.numToCol(c) + m[3] + r;
  }
  // Ajusta referências relativas ao copiar/arrastar uma fórmula.
  function shiftFormula(formula, dr, dc) {
    if (!formula || formula[0] !== '=') return formula;
    return mapRefTokens(formula, (t) => {
      let res;
      if (t.type === 'ref') res = shiftRef(t.refText, dr, dc);
      else if (t.type === 'range') {
        const parts = t.refText.split(':').map((x) => shiftRef(x, dr, dc));
        res = parts.includes('#REF!') ? '#REF!' : parts.join(':');
      } else {
        const parts = t.refText.split(':').map((x) => {
          if (x[0] === '$') return x;
          const c = XL.colToNum(x) + dc;
          return c < 1 ? '#REF!' : XL.numToCol(c);
        });
        res = parts.includes('#REF!') ? '#REF!' : parts.join(':');
      }
      return res === '#REF!' ? res : t.prefix + res;
    });
  }

  // Reescreve os tokens de referência de uma fórmula. fn(token) devolve o novo texto (ou undefined).
  function mapRefTokens(formula, fn) {
    if (!formula || formula[0] !== '=') return formula;
    const toks = tokenize(formula.slice(1), true);
    let changed = false;
    const body = toks.map((t) => {
      if (t.type !== 'ref' && t.type !== 'range' && t.type !== 'colrange') return t.text;
      const r = fn(t);
      if (r === undefined || r === t.text) return t.text;
      changed = true;
      return r;
    }).join('');
    return changed ? '=' + body : formula;
  }

  // Ajuste ao inserir/excluir linhas ou colunas.
  // op = { kind: 'insert'|'delete', axis: 'row'|'col', at (1-based), n, max }
  function adjustPart(x, op) {
    const end = op.at + op.n - 1;
    if (op.kind === 'insert') { const y = x >= op.at ? x + op.n : x; return y > op.max ? null : y; }
    if (x >= op.at && x <= end) return null;
    return x > end ? x - op.n : x;
  }
  function adjustSpan(a, b, op) {
    if (op.kind === 'insert') {
      const na = a >= op.at ? a + op.n : a, nb = b >= op.at ? b + op.n : b;
      return na > op.max ? null : [na, Math.min(nb, op.max)];
    }
    const end = op.at + op.n - 1;
    if (a >= op.at && b <= end) return null;
    const na = a < op.at ? a : a > end ? a - op.n : op.at;
    const nb = b < op.at ? b : b > end ? b - op.n : op.at - 1;
    return nb < na ? null : [na, nb];
  }
  const RE_PART = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
  function adjustStructure(formula, op, pointsToTarget) {
    return mapRefTokens(formula, (t) => {
      if (!pointsToTarget(t.sheet)) return undefined;
      let res;
      if (t.type === 'colrange') {
        if (op.axis === 'row') return undefined;
        const [a, b] = t.refText.split(':');
        const ma = /^(\$?)(\w+)$/.exec(a), mb = /^(\$?)(\w+)$/.exec(b);
        const span = adjustSpan(XL.colToNum(ma[2]), XL.colToNum(mb[2]), op);
        if (!span) return '#REF!';
        res = ma[1] + XL.numToCol(span[0]) + ':' + mb[1] + XL.numToCol(span[1]);
      } else {
        const parts = t.refText.split(':').map((x) => RE_PART.exec(x));
        const val = (m) => (op.axis === 'row' ? parseInt(m[4], 10) : XL.colToNum(m[2]));
        const build = (m, v) => (op.axis === 'row' ? m[1] + m[2].toUpperCase() + m[3] + v : m[1] + XL.numToCol(v) + m[3] + m[4]);
        if (parts.length === 1) {
          const v = adjustPart(val(parts[0]), op);
          if (v === null) return '#REF!';
          res = build(parts[0], v);
        } else {
          const span = adjustSpan(val(parts[0]), val(parts[1]), op);
          if (!span) return '#REF!';
          res = build(parts[0], span[0]) + ':' + build(parts[1], span[1]);
        }
      }
      return t.prefix + res;
    });
  }
  // Ao renomear/excluir planilhas.
  function renameSheetRefs(formula, oldName, newName) {
    const k = XL.normName(oldName);
    return mapRefTokens(formula, (t) => (t.sheet && XL.normName(t.sheet) === k ? XL.quoteSheet(newName) + '!' + t.refText : undefined));
  }
  function dropSheetRefs(formula, name) {
    const k = XL.normName(name);
    return mapRefTokens(formula, (t) => (t.sheet && XL.normName(t.sheet) === k ? '#REF!' : undefined));
  }

  // Lista as referências usadas numa fórmula (para colorir enquanto edita).
  function refsIn(text) {
    if (!text || text[0] !== '=') return [];
    const out = [];
    const seen = new Set();
    for (const t of tokenize(text.slice(1), true)) {
      let g = null;
      if (t.type === 'ref' || t.type === 'range') g = XL.parseRange(t.refText.replace(/\$/g, ''));
      else if (t.type === 'colrange') {
        const [a, b] = t.refText.replace(/\$/g, '').split(':');
        const c1 = XL.colToNum(a) - 1, c2 = XL.colToNum(b) - 1;
        g = { r1: 0, r2: 1e6, c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
      }
      if (!g) continue;
      const key = t.text.replace(/\$/g, '').toUpperCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ g, text: key, sheet: t.sheet || null });
    }
    return out;
  }

  function usesFunction(input, fnName) {
    if (!input || input[0] !== '=') return false;
    const target = XL.Functions.get(fnName);
    return tokenize(input.slice(1), true).some((t) => t.type === 'func' && XL.Functions.get(t.name) === target);
  }

  // F4: alterna $ na referência sob o cursor (A1 → $A$1 → A$1 → $A1 → A1).
  function toggleAbsolute(text, caret) {
    if (!text || text[0] !== '=') return null;
    const toks = tokenize(text.slice(1), true);
    for (const t of toks) {
      if (t.type !== 'ref' && t.type !== 'range') continue;
      const tokStart = t.pos + 1;
      const start = tokStart + t.prefix.length, end = start + t.refText.length;
      if (caret < tokStart || caret > end) continue;
      const parts = t.refText.split(':');
      const m0 = /^(\$?)[A-Za-z]+(\$?)\d+$/.exec(parts[0]);
      const state = (m0[1] ? 2 : 0) + (m0[2] ? 1 : 0); // 0=A1 3=$A$1 1=A$1 2=$A1
      const next = { 0: [true, true], 3: [false, true], 1: [true, false], 2: [false, false] }[state];
      const repl = parts.map((x) => {
        const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(x);
        return (next[0] ? '$' : '') + m[1].toUpperCase() + (next[1] ? '$' : '') + m[2];
      }).join(':');
      const nt = text.slice(0, start) + repl + text.slice(end);
      return { text: nt, caret: start + repl.length };
    }
    return null;
  }

  // Descobre em qual função (e argumento) o cursor está, para a dica de sintaxe.
  function funcContext(text, caret) {
    if (!text || text[0] !== '=') return null;
    const toks = tokenize(text.slice(1, caret), true);
    const stack = [];
    let pending = null;
    for (const t of toks) {
      if (t.type === 'func') { pending = t.name; continue; }
      if (t.type === 'op' && t.text === '(') { stack.push({ name: pending, arg: 0 }); pending = null; continue; }
      if (t.type === 'op' && t.text === ')') { stack.pop(); continue; }
      if (t.type === 'sep') { if (stack.length) stack[stack.length - 1].arg++; continue; }
      if (t.type !== 'ws') pending = null;
    }
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i];
    return null;
  }

  XL.XLError = XLError;
  XL.Formula = {
    tokenize, parse, validate, evaluate, normalize, shiftFormula, refsIn, usesFunction, toggleAbsolute, funcContext,
    mapRefTokens, adjustStructure, adjustSpan, renameSheetRefs, dropSheetRefs,
    XLError, Matrix, E, err, isErr, toNum, toStr, toBool, scalar, compareValues, FormulaSyntaxError,
  };
})();
