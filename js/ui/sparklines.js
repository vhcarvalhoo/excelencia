/* Módulo 11: Minigráficos (linha, coluna, ganho/perda) desenhados dentro das células. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const P = XL.Sheet.prototype;
  const proto = XL.ExcelApp.prototype;
  const baseExtra = proto.execExtra;
  const COLOR = '#4472C4', NEG = '#C00000', HIGH = '#3E9B4F', LOW = '#E06C1F';
  const uid = () => 'sp' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
  const enterSubmits = (inp) => inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const b = inp.closest('.ui-modal').querySelector('.btn-primary'); if (b) b.click(); } });
  const clean = (t) => String(t || '').trim().replace(/^=/, '').replace(/\$/g, '');

  // ---------------- Modelo ----------------
  // Cada local (célula) recebe uma linha (ou coluna) dos dados.
  XL.sparkLayout = (dataText, locText) => {
    const D = XL.parseRange(clean(dataText)), L = XL.parseRange(clean(locText));
    if (!D) return { msg: 'O intervalo de dados não é válido. Ex.: B3:G6' };
    if (!L) return { msg: 'O intervalo de local não é válido. Ex.: H3:H6' };
    if (L.r1 !== L.r2 && L.c1 !== L.c2) return { msg: 'O intervalo de local precisa ser uma única linha ou uma única coluna.' };
    const nLoc = (L.r2 - L.r1 + 1) * (L.c2 - L.c1 + 1);
    const pairs = [];
    const rows = D.r2 - D.r1 + 1, cols = D.c2 - D.c1 + 1;
    for (let i = 0; i < nLoc; i++) {
      const r = L.r1 + (L.r1 === L.r2 ? 0 : i), c = L.c1 + (L.r1 === L.r2 ? i : 0);
      let src;
      if (nLoc === 1) src = D;
      else if (rows === nLoc) src = { r1: D.r1 + i, c1: D.c1, r2: D.r1 + i, c2: D.c2 };
      else if (cols === nLoc) src = { r1: D.r1, c1: D.c1 + i, r2: D.r2, c2: D.c1 + i };
      else return { msg: `O intervalo de dados tem ${rows} linha(s) e ${cols} coluna(s), mas o local tem ${nLoc} célula(s). Eles precisam combinar.` };
      pairs.push({ r, c, src });
    }
    return { pairs, D, L };
  };
  P.sparkAt = function (r, c) {
    return (this.sparks || []).find((g) => { const L = XL.parseRange(clean(g.loc)); return L && XL.inRange(L, r, c); }) || null;
  };

  // ---------------- Desenho ----------------
  function svgFor(type, vals, w, hgt, g) {
    const nums = vals.filter((v) => typeof v === 'number');
    if (!nums.length || w < 6 || hgt < 6) return '';
    const n = vals.length;
    const max = Math.max(...nums), min = Math.min(...nums);
    const hiI = vals.indexOf(max), loI = vals.indexOf(min);
    const out = [];
    if (type === 'line') {
      const lo = min, hi = max === min ? min + 1 : max;
      const X = (i) => (n === 1 ? w / 2 : 2 + (i * (w - 4)) / (n - 1));
      const Y = (v) => hgt - 2 - ((v - lo) / (hi - lo)) * (hgt - 4);
      const pts = vals.map((v, i) => (typeof v === 'number' ? [X(i), Y(v)] : null));
      out.push(`<path d="${pts.filter(Boolean).map((p, k) => (k ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1)).join('')}" fill="none" stroke="${COLOR}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>`);
      pts.forEach((p, i) => {
        if (!p) return;
        let col = null;
        if (g.high && i === hiI) col = HIGH; else if (g.low && i === loI) col = LOW; else if (g.markers) col = COLOR;
        if (col) out.push(`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="2.6" fill="${col}"/>`);
      });
    } else {
      const bw = Math.max(1.5, (w - 2) / n - 1.5);
      const step = (w - 2) / n;
      if (type === 'winloss') {
        const mid = hgt / 2;
        vals.forEach((v, i) => {
          if (typeof v !== 'number' || v === 0) return;
          const x = 1 + i * step + (step - bw) / 2;
          out.push(`<rect x="${x.toFixed(1)}" y="${v > 0 ? 2 : mid + 1}" width="${bw.toFixed(1)}" height="${(mid - 3).toFixed(1)}" fill="${v > 0 ? COLOR : NEG}"/>`);
        });
      } else {
        // eixo automático por minigráfico (como no Excel): parte perto do menor valor, salvo se houver negativos
        const lo = min < 0 ? min : max === min ? 0 : Math.max(0, min - (max - min) * 0.25);
        const hi = max > 0 ? max : 0;
        if (hi === lo) return '';
        const Y = (v) => hgt - 1 - ((v - lo) / (hi - lo)) * (hgt - 3);
        vals.forEach((v, i) => {
          if (typeof v !== 'number') return;
          const x = 1 + i * step + (step - bw) / 2;
          const y0 = Y(0), y1 = Y(v);
          let col = v < 0 ? NEG : COLOR;
          if (g.high && i === hiI) col = HIGH; else if (g.low && i === loI) col = LOW;
          out.push(`<rect x="${x.toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(1, Math.abs(y1 - y0)).toFixed(1)}" fill="${col}"/>`);
        });
      }
    }
    return `<svg viewBox="0 0 ${w} ${hgt}" width="${w}" height="${hgt}" xmlns="http://www.w3.org/2000/svg">${out.join('')}</svg>`;
  }

  Object.assign(proto, {
    drawSparks() {
      const g = this.grid;
      if (!g.sparkLayer) { g.sparkLayer = h('div.xl-sparklayer'); g.chartLayer.parentNode.insertBefore(g.sparkLayer, g.chartLayer); }
      const layer = g.sparkLayer;
      layer.innerHTML = '';
      const S = this.sheet;
      for (const grp of S.sparks || []) {
        const lay = XL.sparkLayout(grp.data, grp.loc);
        if (!lay.pairs) continue;
        for (const p of lay.pairs) {
          const td = g.tds[p.r] && g.tds[p.r][p.c];
          if (!td || !td.offsetHeight || S.hiddenRows.has(p.r)) continue;
          const vals = [];
          XL.eachInRange(p.src, (r, c) => { const v = S.value(r, c); vals.push(typeof v === 'number' ? v : null); });
          const w = td.offsetWidth - 6, hh = td.offsetHeight - 5;
          const box = h('div.xl-spark', { dataset: { cell: XL.toRef(p.r, p.c), type: grp.type }, html: svgFor(grp.type, vals, w, hh, grp) });
          Object.assign(box.style, { left: td.offsetLeft + 3 + 'px', top: td.offsetTop + 2 + 'px', width: w + 'px', height: hh + 'px' });
          layer.appendChild(box);
        }
      }
    },

    activeSpark() {
      if (!this.grid) return null;
      const { ar, ac } = this.grid.sel;
      return this.sheet.sparkAt ? this.sheet.sparkAt(ar, ac) : null;
    },

    execExtra(cmd, arg) {
      switch (cmd) {
        case 'spark-insert': return this.sparkDialog(arg || 'line');
        case 'spark-type': return this.sparkUpdate((g) => { g.type = arg; });
        case 'spark-high': return this.sparkUpdate((g) => { g.high = !!arg; });
        case 'spark-low': return this.sparkUpdate((g) => { g.low = !!arg; });
        case 'spark-markers': return this.sparkUpdate((g) => { g.markers = !!arg; });
        case 'spark-clear': {
          const g = this.activeSpark();
          if (!g) return;
          this.pushUndo();
          this.sheet.sparks = this.sheet.sparks.filter((x) => x !== g);
          this.render();
          this.ribbon.updateContext();
          this.emit('sparkclear', {});
          this.emit('change', {});
          return;
        }
      }
      return baseExtra.call(this, cmd, arg);
    },

    sparkDialog(type) {
      const g = this.grid.range();
      const NAMES = { line: 'Linha', col: 'Coluna', winloss: 'Ganho/Perda' };
      const data = h('input.fd-search', { type: 'text', value: '', dataset: { role: 'sp-data' }, placeholder: 'Ex.: B3:G6' });
      const loc = h('input.fd-search', { type: 'text', value: XL.rangeToStr(g).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), dataset: { role: 'sp-loc' } });
      [data, loc].forEach(enterSubmits);
      XL.ui.modal({
        title: 'Criar Minigráficos', width: 400,
        body: h('div.pvd',
          h('p.cf-q', 'Tipo: ', h('b', NAMES[type])),
          h('div.dv-sec', 'Escolha os dados desejados'), h('label.fd-lbl', 'Intervalo de dados:'), data,
          h('div.dv-sec', 'Escolha onde deseja posicionar os minigráficos'), h('label.fd-lbl', 'Intervalo de local:'), loc,
          h('p.cd-note', 'Cada célula do local recebe uma linha dos dados (ou uma coluna, se o local estiver na horizontal).')),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const lay = XL.sparkLayout(data.value, loc.value);
            if (!lay.pairs) { XL.ui.toast(lay.msg); return false; }
            this.pushUndo();
            const S = this.sheet;
            // um local só pode ter um minigráfico
            S.sparks = S.sparks.filter((x) => { const L = XL.parseRange(clean(x.loc)); return !lay.pairs.some((p) => XL.inRange(L, p.r, p.c)); });
            const grp = { id: uid(), type, data: XL.rangeToStr(lay.D), loc: XL.rangeToStr(lay.L), high: false, low: false, markers: false };
            S.sparks.push(grp);
            this.grid.setSel(lay.L.r1, lay.L.c1, lay.L.r2, lay.L.c2);
            this.render();
            this.ribbon.updateContext();
            this.ribbon.setTab('sparkdesign', true);
            this.emit('spark', { type, data: grp.data, loc: grp.loc });
            this.emit('change', {});
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => data.focus(), 0);
    },

    sparkUpdate(fn) {
      const g = this.activeSpark();
      if (!g) { XL.ui.toast('Clique numa célula com minigráfico primeiro.'); return; }
      this.pushUndo();
      fn(g);
      this.render();
      this.ribbon.setTab(this.ribbon.active, true);
      this.emit('sparkchange', { type: g.type, high: g.high, low: g.low, markers: g.markers });
      this.emit('change', {});
    },
  });
})();
