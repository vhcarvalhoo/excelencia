/* Utilidades gerais: referências de célula, DOM, eventos, números pt-BR, datas. */
(function () {
  'use strict';
  const XL = (window.XL = window.XL || {});

  // ---------- Referências de célula ----------
  XL.colToNum = (s) => {
    let n = 0;
    for (const ch of String(s).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  };
  XL.numToCol = (n) => {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };
  XL.toRef = (r, c) => XL.numToCol(c + 1) + (r + 1);
  XL.parseRef = (ref) => {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
    if (!m) return null;
    const r = parseInt(m[2], 10) - 1;
    if (r < 0) return null;
    return { r, c: XL.colToNum(m[1]) - 1 };
  };
  XL.parseRange = (str) => {
    const parts = String(str).trim().split(':');
    if (parts.length > 2) return null;
    const a = XL.parseRef(parts[0]);
    if (!a) return null;
    const b = parts[1] !== undefined ? XL.parseRef(parts[1]) : a;
    if (!b) return null;
    return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
  };
  XL.rangeToStr = (g) =>
    g.r1 === g.r2 && g.c1 === g.c2 ? XL.toRef(g.r1, g.c1) : XL.toRef(g.r1, g.c1) + ':' + XL.toRef(g.r2, g.c2);
  XL.eachInRange = (g, fn) => {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c);
  };

  // ---------- Texto ----------
  XL.stripAccents = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  XL.normName = (s) => XL.stripAccents(s).toUpperCase();
  XL.escapeHtml = (s) =>
    String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

  // ---------- Eventos ----------
  class Emitter {
    constructor() { this._l = {}; }
    on(evt, fn) { (this._l[evt] = this._l[evt] || []).push(fn); return () => this.off(evt, fn); }
    off(evt, fn) { const a = this._l[evt]; if (a) { const i = a.indexOf(fn); if (i >= 0) a.splice(i, 1); } }
    emit(evt, data) {
      (this._l[evt] || []).slice().forEach((f) => f(data));
      if (evt !== '*') (this._l['*'] || []).slice().forEach((f) => f(evt, data));
    }
  }
  XL.Emitter = Emitter;

  // ---------- DOM ----------
  // XL.h('div.classe#id', {atributos}, ...filhos)
  XL.h = function (tag, attrs, ...children) {
    const m = /^([a-z0-9]+)?((?:[.#][\w-]+)*)$/i.exec(tag);
    const el = document.createElement((m && m[1]) || 'div');
    if (m && m[2]) {
      m[2].replace(/([.#])([\w-]+)/g, (_, t, v) => {
        if (t === '.') el.classList.add(v);
        else el.id = v;
      });
    }
    if (attrs && (typeof attrs !== 'object' || attrs instanceof Node || Array.isArray(attrs))) {
      children.unshift(attrs);
      attrs = null;
    }
    if (attrs) {
      for (const k in attrs) {
        const v = attrs[k];
        if (v === null || v === undefined || v === false) continue;
        if (k === 'class') String(v).split(/\s+/).filter(Boolean).forEach((c) => el.classList.add(c));
        else if (k === 'style') {
          if (typeof v === 'string') el.style.cssText = v;
          else for (const p in v) {
            if (p.startsWith('--')) el.style.setProperty(p, v[p]);
            else el.style[p] = v[p];
          }
        } else if (k === 'html') el.innerHTML = v;
        else if (k === 'text') el.textContent = v;
        else if (k === 'dataset') Object.assign(el.dataset, v);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
        else if (k === 'value') el.value = v;
        else el.setAttribute(k, v === true ? '' : v);
      }
    }
    const add = (c) => {
      if (c === null || c === undefined || c === false) return;
      if (Array.isArray(c)) c.forEach(add);
      else el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
    };
    children.forEach(add);
    return el;
  };

  // ---------- Armazenamento local (progresso) ----------
  XL.store = {
    get(k, d) {
      try {
        const v = localStorage.getItem('excelencia:' + k);
        return v == null ? d : JSON.parse(v);
      } catch (e) {
        return d;
      }
    },
    set(k, v) {
      try { localStorage.setItem('excelencia:' + k, JSON.stringify(v)); } catch (e) { /* sem armazenamento */ }
    },
  };

  // ---------- Números (pt-BR) ----------
  const nfCache = new Map();
  XL.nf = (min, max, grouping) => {
    const k = min + '|' + max + '|' + grouping;
    let f = nfCache.get(k);
    if (!f) {
      f = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: grouping });
      nfCache.set(k, f);
    }
    return f;
  };

  // Formato "Geral": até ~11 dígitos significativos, vírgula decimal.
  XL.fmtGeneral = (n) => {
    if (n === 0) return '0';
    if (!isFinite(n)) return '#NÚM!';
    const a = Math.abs(n);
    if (a >= 1e11 || a < 1e-9) {
      return n.toExponential(5).replace(/\.?0+e/, 'e').replace('.', ',').replace('e+', 'E+').replace('e-', 'E-');
    }
    const p = Number(n.toPrecision(11));
    return XL.nf(0, 10, false).format(p);
  };

  // Converte texto em número no padrão brasileiro. Retorna null se não for número.
  XL.parseNumberStr = (str) => {
    const s = String(str).trim();
    if (!s) return null;
    if (/^[-+]?\d+(,\d+)?([eE][-+]?\d+)?$/.test(s)) return parseFloat(s.replace(',', '.'));
    if (/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
    if (/^[-+]?,\d+$/.test(s)) return parseFloat(s.replace(',', '.'));
    return null;
  };

  // ---------- Datas (número de série do Excel) ----------
  const EPOCH = Date.UTC(1899, 11, 30);
  XL.dateToSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
  XL.serialToParts = (s) => {
    const dt = new Date(EPOCH + Math.floor(s) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() };
  };
  XL.todaySerial = () => {
    const n = new Date();
    return XL.dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
  };
  const daysInMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  XL.MONTHS = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  XL.WEEKDAYS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];

  // ---------- Interpretação do que o usuário digita ----------
  XL.parseInput = (text) => {
    const s = String(text === null || text === undefined ? '' : text);
    if (s.trim() === '') return { kind: 'empty' };
    if (s[0] === "'") return { kind: 'text', value: s.slice(1) };
    if (s[0] === '=' && s.length > 1) return { kind: 'formula' };
    const t = s.trim();
    const up = XL.normName(t);
    if (up === 'VERDADEIRO' || up === 'TRUE') return { kind: 'bool', value: true };
    if (up === 'FALSO' || up === 'FALSE') return { kind: 'bool', value: false };
    let n = XL.parseNumberStr(t);
    if (n !== null) return { kind: 'number', value: n };
    let m = /^([-+]?[\d.,]+)\s*%$/.exec(t);
    if (m) {
      n = XL.parseNumberStr(m[1]);
      if (n !== null) {
        const dec = (m[1].split(',')[1] || '').length;
        return { kind: 'number', value: n / 100, fmt: { numFmt: 'porcentagem', dec } };
      }
    }
    m = /^(-?)\s*R\$\s*([\d.,]+)$/i.exec(t);
    if (m) {
      n = XL.parseNumberStr(m[2]);
      if (n !== null) return { kind: 'number', value: m[1] ? -n : n, fmt: { numFmt: 'moeda' } };
    }
    m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2}|\d{4}))?$/.exec(t);
    if (m) {
      const d = +m[1], mo = +m[2];
      let y = m[3] ? +m[3] : new Date().getFullYear();
      if (m[3] && m[3].length === 2) y += y < 30 ? 2000 : 1900;
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= daysInMonth(y, mo)) {
        return { kind: 'number', value: XL.dateToSerial(y, mo, d), fmt: { numFmt: 'data' } };
      }
    }
    return { kind: 'text', value: s };
  };

  // ---------- Exibição de valores ----------
  const pad2 = (n) => String(n).padStart(2, '0');
  XL.formatNumber = (n, fmt) => {
    const f = fmt || {};
    const nf = f.numFmt || 'geral';
    const dec = f.dec;
    switch (nf) {
      case 'numero': {
        const d = dec !== undefined && dec !== null ? dec : 2;
        return XL.nf(d, d, f.grouping !== false).format(n);
      }
      case 'moeda':
      case 'contabil': {
        const d = dec !== undefined && dec !== null ? dec : 2;
        if (nf === 'contabil' && n === 0) return 'R$ -';
        const s = XL.nf(d, d, true).format(Math.abs(n));
        return (n < 0 ? '-' : '') + 'R$ ' + s;
      }
      case 'porcentagem': {
        const d = dec !== undefined && dec !== null ? dec : 0;
        return XL.nf(d, d, false).format(n * 100) + '%';
      }
      case 'data': {
        if (n < 0 || n > 2958465) return '########';
        const p = XL.serialToParts(n);
        return pad2(p.d) + '/' + pad2(p.m) + '/' + p.y;
      }
      case 'datalonga': {
        if (n < 0 || n > 2958465) return '########';
        const p = XL.serialToParts(n);
        return `${XL.WEEKDAYS[p.wd]}, ${p.d} de ${XL.MONTHS[p.m - 1]} de ${p.y}`;
      }
      case 'hora': {
        let secs = Math.round((n - Math.floor(n)) * 86400);
        const hh = Math.floor(secs / 3600); secs -= hh * 3600;
        const mm = Math.floor(secs / 60);
        return pad2(hh) + ':' + pad2(mm) + ':' + pad2(secs - mm * 60);
      }
      case 'custom':
        return XL.formatCode(n, f.code || 'Geral');
      default:
        if (dec !== undefined && dec !== null) return XL.nf(dec, dec, false).format(n);
        return XL.fmtGeneral(n);
    }
  };

  // ---------- Códigos de formato personalizados (como em Formatar Células › Personalizado e na função TEXTO) ----------
  // Números: 0 # , (decimal) . (milhar) %  · Datas: d dd ddd dddd m mm mmm mmmm aa aaaa (ou yy yyyy) h hh mm ss
  // Texto fixo: "entre aspas" ou \x · Seções: positivo;negativo;zero
  const MON3 = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const WD3 = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
  function splitSections(code) {
    const out = [];
    let cur = '', q = false;
    for (let i = 0; i < code.length; i++) {
      const ch = code[i];
      if (ch === '"') q = !q;
      if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
      if (ch === '\\' && !q && i + 1 < code.length) { cur += ch + code[++i]; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  function tokenizeCode(sec) {
    const toks = [];
    for (let i = 0; i < sec.length; i++) {
      const ch = sec[i];
      if (ch === '"') { let j = i + 1, s = ''; while (j < sec.length && sec[j] !== '"') s += sec[j++]; toks.push({ t: 'lit', s }); i = j; continue; }
      if (ch === '\\') { toks.push({ t: 'lit', s: sec[i + 1] || '' }); i++; continue; }
      if (ch === '[') { while (i < sec.length && sec[i] !== ']') i++; continue; }
      if (ch === '_') { toks.push({ t: 'lit', s: ' ' }); i++; continue; }
      if (ch === '*') { i++; continue; }
      if (/[dmayhs]/i.test(ch)) {
        let j = i;
        while (j < sec.length && sec[j].toLowerCase() === ch.toLowerCase()) j++;
        toks.push({ t: 'date', s: sec.slice(i, j).toLowerCase() });
        i = j - 1;
        continue;
      }
      if ('0#?.,'.includes(ch)) { toks.push({ t: 'num', s: ch }); continue; }
      if (ch === '%') { toks.push({ t: 'pct' }); continue; }
      if (ch === '@') { toks.push({ t: 'at' }); continue; }
      toks.push({ t: 'lit', s: ch });
    }
    return toks;
  }
  XL.isDateCode = (code) => tokenizeCode(splitSections(String(code))[0]).some((t) => t.t === 'date');
  function formatDateCode(v, toks) {
    if (v < 0 || v > 2958465) return '########';
    const p = XL.serialToParts(v);
    let secs = Math.round((v - Math.floor(v)) * 86400);
    const hh = Math.floor(secs / 3600); secs -= hh * 3600;
    const mi = Math.floor(secs / 60); secs -= mi * 60;
    const dates = toks.map((t, i) => (t.t === 'date' ? i : -1)).filter((i) => i >= 0);
    return toks.map((t, i) => {
      if (t.t !== 'date') return t.t === 'lit' ? t.s : t.t === 'pct' ? '%' : '';
      const s = t.s, c = s[0], n = s.length;
      if (c === 'd') return n === 1 ? String(p.d) : n === 2 ? pad2(p.d) : n === 3 ? WD3[p.wd] : XL.WEEKDAYS[p.wd];
      if (c === 'a' || c === 'y') return n <= 2 ? pad2(p.y % 100) : String(p.y);
      if (c === 'h') return n === 1 ? String(hh) : pad2(hh);
      if (c === 's') return n === 1 ? String(secs) : pad2(secs);
      // "m" é minuto quando vem depois de hora ou antes de segundo
      const k = dates.indexOf(i);
      const prev = k > 0 ? toks[dates[k - 1]].s[0] : '';
      const next = k < dates.length - 1 ? toks[dates[k + 1]].s[0] : '';
      if (n <= 2 && (prev === 'h' || next === 's')) return n === 1 ? String(mi) : pad2(mi);
      return n === 1 ? String(p.m) : n === 2 ? pad2(p.m) : n === 3 ? MON3[p.m - 1] : n === 4 ? XL.MONTHS[p.m - 1] : XL.MONTHS[p.m - 1][0];
    }).join('');
  }
  function formatNumCode(v, toks, neg) {
    const idx = toks.map((t, i) => (t.t === 'num' ? i : -1)).filter((i) => i >= 0);
    const pct = toks.some((t) => t.t === 'pct');
    const lit = (t) => (t.t === 'lit' ? t.s : t.t === 'pct' ? '%' : '');
    if (!idx.length) return toks.map(lit).join('');
    const first = idx[0], last = idx[idx.length - 1];
    const spec = toks.slice(first, last + 1).filter((t) => t.t === 'num').map((t) => t.s).join('');
    const ci = spec.indexOf(',');
    const intSpec = ci >= 0 ? spec.slice(0, ci) : spec;
    const decSpec = ci >= 0 ? spec.slice(ci + 1).replace(/[.,]/g, '') : '';
    const grouping = intSpec.includes('.');
    const minInt = (intSpec.match(/0/g) || []).length;
    const maxDec = decSpec.length;
    const minDec = (decSpec.match(/0/g) || []).length;
    const x = Math.abs(v) * (pct ? 100 : 1);
    let [ip, dp = ''] = x.toFixed(maxDec).split('.');
    while (dp.length > minDec && dp.endsWith('0')) dp = dp.slice(0, -1);
    if (ip === '0' && minInt === 0) ip = '';
    ip = ip.padStart(minInt, '0');
    if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    const num = (neg && (ip || dp) && Number(ip + dp) !== 0 ? '-' : '') + ip + (dp ? ',' + dp : '');
    return toks.slice(0, first).map(lit).join('') + num + toks.slice(last + 1).map(lit).join('');
  }
  XL.formatCode = (n, code) => {
    code = String(code);
    if (/^(geral|general)$/i.test(code.trim())) return XL.fmtGeneral(n);
    const secs = splitSections(code);
    let sec = secs[0];
    let v = n, neg = n < 0;
    if (n < 0 && secs[1] !== undefined) { sec = secs[1]; v = -n; neg = false; }
    else if (n === 0 && secs[2] !== undefined) sec = secs[2];
    const toks = tokenizeCode(sec);
    if (toks.some((t) => t.t === 'date')) return formatDateCode(v, toks);
    return formatNumCode(v, toks, neg);
  };

  XL.formatValue = (v, fmt) => {
    if (v === null || v === undefined) return { text: '', align: null };
    if (v && v.isXLError) return { text: v.code, align: 'center', err: true };
    if (typeof v === 'boolean') return { text: v ? 'VERDADEIRO' : 'FALSO', align: 'center' };
    if (typeof v === 'number') return { text: XL.formatNumber(v, fmt), align: 'right', num: true };
    return { text: String(v), align: 'left' };
  };

  // Texto exibido na Barra de Fórmulas para uma célula.
  XL.editTextFor = (cell) => {
    if (!cell || cell.kind === 'empty') return '';
    if (cell.kind === 'formula' || cell.kind === 'text') return cell.input;
    if (cell.kind === 'bool') return cell.value ? 'VERDADEIRO' : 'FALSO';
    const v = cell.value;
    const nf = cell.fmt && cell.fmt.numFmt;
    if (nf === 'porcentagem') return XL.nf(0, 10, false).format(Number((v * 100).toPrecision(12))) + '%';
    if (nf === 'data' || nf === 'datalonga' || (nf === 'custom' && XL.isDateCode(cell.fmt.code || ''))) return XL.formatNumber(v, { numFmt: 'data' });
    return XL.nf(0, 10, false).format(v);
  };
})();
