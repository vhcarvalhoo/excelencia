/* Janelas "Formatar Células" (Ctrl+1) e "Texto para Colunas". */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const D = XL.Dialogs;

  // ================= Formatar Células =================
  const CATS = [
    ['geral', 'Geral'], ['numero', 'Número'], ['moeda', 'Moeda'], ['contabil', 'Contábil'], ['data', 'Data'],
    ['hora', 'Hora'], ['porcentagem', 'Porcentagem'], ['texto', 'Texto'], ['custom', 'Personalizado'],
  ];
  const DATE_CODES = ['dd/mm/aaaa', 'd/m/aa', 'd "de" mmmm "de" aaaa', 'dddd, d "de" mmmm', 'mmm/aa', 'mmmm', 'dddd'];
  const CUSTOM_CODES = ['0', '0,00', '#.##0', '#.##0,00', '0000', '0%', '0,0%', 'R$ #.##0,00', '#.##0,0 "kg"', '"Nota: "0,0',
    'dd/mm/aaaa', 'd "de" mmmm "de" aaaa', 'dddd', 'mmm/aa', 'hh:mm'];

  D.formatCells = (app) => {
    const S = app.sheet;
    const g = app.grid.range();
    const { ar, ac } = app.grid.sel;
    const cur = S.fmt(ar, ac);
    const sample = (() => { const v = S.value(ar, ac); return typeof v === 'number' ? v : 1234.5; })();
    let cat = cur.numFmt || 'geral';
    if (cat === 'datalonga') cat = 'data';
    const st = {
      dec: cur.dec !== undefined && cur.dec !== null ? cur.dec : 2,
      grouping: cur.grouping !== false,
      dateCode: cur.numFmt === 'custom' && XL.isDateCode(cur.code || '') ? cur.code : 'dd/mm/aaaa',
      code: cur.numFmt === 'custom' ? cur.code : '',
    };
    if (cur.numFmt === 'custom' && XL.isDateCode(cur.code || '')) cat = 'data';

    const list = h('div.fc-cats');
    const panel = h('div.fc-panel');
    const preview = h('div.fc-preview');

    // formato que será aplicado, conforme a categoria e as opções
    const patch = () => {
      switch (cat) {
        case 'geral': return { numFmt: null, dec: null, grouping: null, code: null };
        case 'numero': return { numFmt: 'numero', dec: st.dec, grouping: st.grouping ? null : false, code: null };
        case 'moeda': case 'contabil': case 'porcentagem': return { numFmt: cat, dec: st.dec, grouping: null, code: null };
        case 'data': return st.dateCode === 'dd/mm/aaaa' ? { numFmt: 'data', dec: null, grouping: null, code: null } : { numFmt: 'custom', code: st.dateCode, dec: null, grouping: null };
        case 'hora': return { numFmt: 'hora', dec: null, grouping: null, code: null };
        case 'texto': return { numFmt: 'texto', dec: null, grouping: null, code: null };
        case 'custom': return { numFmt: 'custom', code: st.code || 'Geral', dec: null, grouping: null };
      }
      return {};
    };
    const updatePreview = () => {
      const p = patch();
      const f = {};
      for (const k in p) if (p[k] !== null && p[k] !== undefined) f[k] = p[k];
      let txt;
      try { txt = cat === 'texto' ? String(S.textOf(ar, ac) || sample) : XL.formatNumber(sample, f); } catch (e) { txt = '(formato inválido)'; }
      preview.innerHTML = `<span class="fc-plbl">Exemplo</span><span class="fc-pval">${XL.escapeHtml(txt)}</span>`;
    };
    const decInput = () => {
      const i = h('input.fd-search.fc-dec', { type: 'number', min: '0', max: '15', value: String(st.dec), dataset: { role: 'dec' } });
      i.addEventListener('input', () => { st.dec = Math.max(0, Math.min(15, Number(i.value) || 0)); updatePreview(); });
      return h('label.fd-lbl', 'Casas decimais: ', i);
    };
    const codeList = (codes, current, onPick, role) => {
      const box = h('div.fd-list.fc-codes', { dataset: { role } });
      codes.forEach((c) => {
        const it = h('div.fd-item' + (c === current ? '.on' : ''), { dataset: { code: c } }, c);
        it.addEventListener('click', () => { box.querySelectorAll('.fd-item').forEach((x) => x.classList.toggle('on', x === it)); onPick(c); });
        box.appendChild(it);
      });
      return box;
    };
    const drawPanel = () => {
      panel.innerHTML = '';
      panel.appendChild(preview);
      if (cat === 'geral') panel.appendChild(h('p.cd-note', 'Sem formato específico: o Excel mostra o número como ele é.'));
      if (['numero', 'moeda', 'contabil', 'porcentagem'].includes(cat)) panel.appendChild(decInput());
      if (cat === 'numero') {
        const cb = h('input', { type: 'checkbox' });
        cb.checked = st.grouping;
        cb.addEventListener('change', () => { st.grouping = cb.checked; updatePreview(); });
        panel.appendChild(h('label.sd-check', cb, ' Usar separador de milhar (.)'));
      }
      if (cat === 'data') { panel.appendChild(h('label.fd-lbl', 'Tipo:')); panel.appendChild(codeList(DATE_CODES, st.dateCode, (c) => { st.dateCode = c; updatePreview(); }, 'datecodes')); }
      if (cat === 'texto') panel.appendChild(h('p.cd-note', 'O conteúdo é tratado como texto, mesmo que seja um número (bom para CEP, telefone e códigos).'));
      if (cat === 'custom') {
        const inp = h('input.fd-search', { type: 'text', value: st.code, dataset: { role: 'code' }, placeholder: 'Ex.: 0000 ou #.##0,00 "kg"', spellcheck: 'false' });
        inp.addEventListener('input', () => { st.code = inp.value; updatePreview(); });
        panel.appendChild(h('label.fd-lbl', 'Tipo:'));
        panel.appendChild(inp);
        panel.appendChild(codeList(CUSTOM_CODES, st.code, (c) => { st.code = c; inp.value = c; updatePreview(); }, 'customcodes'));
        panel.appendChild(h('p.cd-note', h('b', '0'), ' = dígito obrigatório · ', h('b', '#'), ' = dígito opcional · ', h('b', ','), ' = decimal · ', h('b', '.'), ' = milhar · texto entre ', h('b', '"aspas"'), ' · datas: d, m, aaaa'));
      }
      updatePreview();
    };
    CATS.forEach(([id, label]) => {
      const it = h('div.fd-item.fc-cat' + (id === cat ? '.on' : ''), { dataset: { cat: id } }, label);
      it.addEventListener('click', () => {
        cat = id;
        list.querySelectorAll('.fc-cat').forEach((x) => x.classList.toggle('on', x === it));
        drawPanel();
      });
      list.appendChild(it);
    });
    drawPanel();
    XL.ui.modal({
      title: 'Formatar Células', width: 560,
      body: h('div.fc', h('div.fc-left', h('div.fd-lbl', 'Categoria:'), list), panel),
      buttons: [{
        label: 'OK', primary: true,
        onClick: () => {
          if (cat === 'custom' && !st.code.trim()) { XL.ui.toast('Digite um código de formato no campo Tipo.'); return false; }
          app.grid.setRange(g);
          app.applyFmt(patch());
          app.emit('formatcells', { cat, code: cat === 'custom' ? st.code : cat === 'data' ? st.dateCode : null });
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
  };

  // ================= Texto para Colunas =================
  D.textToColumns = (app) => {
    const S = app.sheet;
    let g = app.grid.range();
    if (g.c1 !== g.c2) { XL.ui.toast('Selecione <b>uma coluna só</b> com o texto que deve ser dividido.'); return; }
    if (g.r1 === g.r2) {
      // só uma célula: usa o bloco contínuo dessa coluna
      let r1 = g.r1, r2 = g.r1;
      while (r1 > 0 && S.has(r1 - 1, g.c1)) r1--;
      while (r2 < S.rows - 1 && S.has(r2 + 1, g.c1)) r2++;
      g = { r1, r2, c1: g.c1, c2: g.c1 };
    }
    const rows = [];
    for (let r = g.r1; r <= g.r2; r++) { const c = S.cell(r, g.c1); rows.push(c && c.kind !== 'empty' ? XL.editTextFor(c) : ''); }
    if (!rows.some(Boolean)) { XL.ui.toast('A seleção está vazia.'); return; }
    app.grid.setRange(g);
    const delims = { tab: false, semicolon: false, comma: false, space: false, other: false };
    // sugere o delimitador mais comum
    const sample = rows.join('\n');
    if (sample.includes(';')) delims.semicolon = true;
    else if (sample.includes('\t')) delims.tab = true;
    else if (sample.includes(',')) delims.comma = true;
    let otherChar = '';
    let merge = false;
    const split = (s) => {
      const chars = [];
      if (delims.tab) chars.push('\t');
      if (delims.semicolon) chars.push(';');
      if (delims.comma) chars.push(',');
      if (delims.space) chars.push(' ');
      if (delims.other && otherChar) chars.push(otherChar);
      if (!chars.length) return [s];
      const re = new RegExp('[' + chars.map((c) => c.replace(/[\]\\^-]/g, '\\$&')).join('') + ']' + (merge ? '+' : ''));
      return s.split(re);
    };
    const prev = h('div.tc-preview');
    const drawPrev = () => {
      prev.innerHTML = '';
      const tbl = h('table.tc-table');
      rows.slice(0, 6).forEach((s) => tbl.appendChild(h('tr', split(s).map((p) => h('td', p)))));
      prev.appendChild(tbl);
    };
    const cb = (key, label) => {
      const i = h('input', { type: 'checkbox', dataset: { delim: key } });
      i.checked = delims[key];
      i.addEventListener('change', () => { delims[key] = i.checked; drawPrev(); });
      return h('label.sd-check', i, ' ' + label);
    };
    const other = h('input.fd-search.tc-other', { type: 'text', maxlength: '1', dataset: { role: 'other' } });
    other.addEventListener('input', () => { otherChar = other.value; delims.other = !!otherChar; const o = other.parentNode && other.parentNode.querySelector('input[data-delim="other"]'); if (o) o.checked = delims.other; drawPrev(); });
    const mergeCb = h('input', { type: 'checkbox' });
    mergeCb.addEventListener('change', () => { merge = mergeCb.checked; drawPrev(); });
    const dest = h('input.fd-search', { type: 'text', value: XL.toRef(g.r1, g.c1), dataset: { role: 'dest' } });
    const otherLbl = cb('other', 'Outros:');
    otherLbl.appendChild(other);
    drawPrev();
    XL.ui.modal({
      title: 'Texto para Colunas', width: 560,
      body: h('div.tc',
        h('p', 'Divide o texto de cada célula em várias colunas, usando um caractere separador (delimitador).'),
        h('div.fd-lbl', 'Delimitadores:'),
        h('div.tc-delims', cb('tab', 'Tabulação'), cb('semicolon', 'Ponto e vírgula'), cb('comma', 'Vírgula'), cb('space', 'Espaço'), otherLbl),
        h('label.sd-check', mergeCb, ' Considerar delimitadores consecutivos como um só'),
        h('div.fd-lbl', 'Visualização dos dados:'), prev,
        h('label.fd-lbl', 'Destino (primeira célula):'), dest),
      buttons: [{
        label: 'Concluir', primary: true,
        onClick: () => {
          const d = XL.parseRef(dest.value.replace(/\$/g, ''));
          if (!d) { XL.ui.toast('Destino inválido. Use algo como <b>A1</b>.'); return false; }
          app.pushUndo();
          let maxCols = 0;
          rows.forEach((s, i) => {
            const parts = split(s);
            maxCols = Math.max(maxCols, parts.length);
            parts.forEach((p, j) => {
              const r = d.r + i, c = d.c + j;
              if (r >= S.rows || c >= S.cols) return;
              try { S.setInput(r, c, p.startsWith('=') ? "'" + p : p); } catch (e) { S.setValue(r, c, p); }
            });
          });
          app.render();
          app.grid.setRange({ r1: d.r, c1: d.c, r2: Math.min(S.rows - 1, d.r + rows.length - 1), c2: Math.min(S.cols - 1, d.c + maxCols - 1) });
          app.emit('texttocols', { cols: maxCols, rows: rows.length });
          app.emit('change', {});
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
  };

  // ================= Preenchimento Relâmpago (aprende o padrão pelos exemplos) =================
  const DELIMS = [' ', ',', ';', '@', '.', '-', '/', '_'];
  const proper = (s) => s.toLocaleLowerCase('pt-BR').replace(/(^|[^A-Za-zÀ-ÿ])([A-Za-zÀ-ÿ])/g, (_, p, c) => p + c.toLocaleUpperCase('pt-BR'));
  const CASES = { n: (s) => s, p: proper, u: (s) => s.toLocaleUpperCase('pt-BR'), l: (s) => s.toLocaleLowerCase('pt-BR') };
  function atoms(src, preferEnd) {
    const out = [];
    src.forEach((text, j) => {
      if (!text) return;
      out.push({ k: ['w', j], v: text });
      for (const d of DELIMS) {
        const parts = text.split(d).filter((p) => p !== '');
        if (parts.length < 2) continue;
        const idx = parts.map((_, i) => i);
        const order = preferEnd ? idx.slice().reverse() : idx;
        for (const i of order) {
          const fromEnd = preferEnd;
          out.push({ k: ['t', j, d, fromEnd ? i - parts.length : i], v: parts[i] });
        }
        for (const i of order) out.push({ k: ['i', j, d, preferEnd ? i - parts.length : i], v: parts[i][0] });
      }
    });
    return out;
  }
  function evalAtom(k, src) {
    const text = src[k[1]];
    if (!text) return null;
    if (k[0] === 'w') return text;
    const parts = text.split(k[2]).filter((p) => p !== '');
    const i = k[3] < 0 ? parts.length + k[3] : k[3];
    if (parts.length < 2 || i < 0 || i >= parts.length) return null;
    return k[0] === 'i' ? parts[i][0] : parts[i];
  }
  function learn(example, preferEnd) {
    const out = example.out;
    const list = atoms(example.src, preferEnd);
    const prog = [];
    let pos = 0;
    while (pos < out.length) {
      let best = null;
      for (const a of list) {
        for (const c in CASES) {
          const s = CASES[c](a.v);
          if (!s || s.length <= (best ? best.len : 0)) continue;
          if (s.length === 1 && !/[A-ZÀ-Ý]/.test(out[pos]) && a.k[0] !== 'w') continue; // iniciais só em maiúsculas
          if (out.startsWith(s, pos)) best = { k: a.k, c, len: s.length };
        }
      }
      if (best) { prog.push({ atom: best.k, c: best.c }); pos += best.len; }
      else {
        const last = prog[prog.length - 1];
        if (last && last.lit !== undefined) last.lit += out[pos]; else prog.push({ lit: out[pos] });
        pos++;
      }
    }
    return prog;
  }
  function run(prog, src) {
    let s = '';
    for (const p of prog) {
      if (p.lit !== undefined) { s += p.lit; continue; }
      const v = evalAtom(p.atom, src);
      if (v === null) return null;
      s += CASES[p.c](v);
    }
    return s;
  }
  XL.FlashFill = { learn, run };

  D.flashFill = (app) => {
    const S = app.sheet;
    const { ar, ac } = app.grid.sel;
    const reg = S.currentRegion(ar, ac);
    const c = ac;
    if (reg.c1 === reg.c2) { XL.ui.toast('O Preenchimento Relâmpago precisa de dados nas colunas ao lado. Digite um exemplo ao lado dos dados.'); return; }
    const header = S.detectHeader(reg);
    const srcCols = [];
    for (let j = reg.c1; j <= reg.c2; j++) if (j !== c) srcCols.push(j);
    const srcOf = (r) => srcCols.map((j) => { const x = S.cell(r, j); return x && x.kind !== 'empty' ? S.textOf(r, j) : ''; });
    const examples = [], targets = [];
    for (let r = reg.r1 + (header ? 1 : 0); r <= reg.r2; r++) {
      const src = srcOf(r);
      if (!src.some(Boolean)) continue;
      const cell = S.cell(r, c);
      if (cell && cell.kind !== 'empty' && cell.kind !== 'formula') examples.push({ r, src, out: S.textOf(r, c) });
      else if (!cell || cell.kind === 'empty') targets.push({ r, src });
    }
    if (!examples.length) { XL.ui.toast('Digite pelo menos um <b>exemplo</b> na coluna antes de usar o Preenchimento Relâmpago (Ctrl+E).'); return; }
    if (!targets.length) { XL.ui.toast('Não há células vazias para preencher nesta coluna.'); return; }
    let prog = null;
    for (const ex of examples) {
      for (const pe of [false, true]) {
        const p = learn(ex, pe);
        if (examples.every((e) => run(p, e.src) === e.out)) { prog = p; break; }
      }
      if (prog) break;
    }
    if (!prog || prog.every((p) => p.lit !== undefined)) {
      XL.ui.toast('O Preenchimento Relâmpago não reconheceu um padrão. Digite mais um ou dois exemplos e tente de novo.');
      return;
    }
    app.pushUndo();
    let n = 0;
    for (const t of targets) {
      const v = run(prog, t.src);
      if (v === null) continue;
      S.setValue(t.r, c, v);
      n++;
    }
    app.render();
    const rr = targets.map((t) => t.r);
    app.grid.setRange({ r1: Math.min(...rr), c1: c, r2: Math.max(...rr), c2: c });
    app.emit('flashfill', { col: XL.numToCol(c + 1), count: n });
    app.emit('change', {});
    XL.ui.toast(`⚡ Preenchimento Relâmpago: <b>${n}</b> ${n === 1 ? 'célula preenchida' : 'células preenchidas'}.`);
  };
})();
