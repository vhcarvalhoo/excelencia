/* Gráficos do simulador: desenho em SVG (colunas, barras, linhas, pizza), mover/redimensionar,
   seleção e dica ao passar o mouse. Paleta validada (6 cores, ordem fixa, daltonismo ok). */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const esc = XL.escapeHtml;
  // azul e laranja no estilo do Office; demais ajustadas para passar no validador (contraste e daltonismo)
  const PALETTE = ['#4472C4', '#E06C1F', '#8E5BC9', '#3E9B4F', '#C1508F', '#A67C00'];
  const INK = '#262626', INK2 = '#595959', GRID = '#E7E7E7', AXIS = '#A6A6A6';
  const TYPES = { col: 'Colunas', bar: 'Barras', line: 'Linhas', pie: 'Pizza' };

  const fmtVal = (v) => XL.nf(0, Number.isInteger(v) ? 0 : 2, true).format(v);
  const fmtAxis = (v) => (Math.abs(v) >= 10000 ? XL.nf(0, 1, true).format(v / 1000) + ' mil' : XL.nf(0, 2, true).format(v));
  const textW = (s, size = 11) => String(s).length * size * 0.55;
  function niceScale(min, max) {
    if (max === min) max = min + 1;
    const raw = (max - min) / 5;
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    const f = raw / p;
    const step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
    const ticks = [];
    for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Number(v.toPrecision(12)));
    return { lo, hi, ticks };
  }
  const clip = (s, w, size = 11) => { s = String(s); const max = Math.max(3, Math.floor(w / (size * 0.55))); return s.length > max ? s.slice(0, max - 1) + '…' : s; };
  // barra com a ponta de dados arredondada (4px) e base reta
  function colPath(x, w, y0, y1) {
    const r = Math.min(4, w / 2, Math.abs(y0 - y1));
    if (y1 <= y0) return `M${x},${y0}V${y1 + r}Q${x},${y1} ${x + r},${y1}H${x + w - r}Q${x + w},${y1} ${x + w},${y1 + r}V${y0}Z`;
    return `M${x},${y0}V${y1 - r}Q${x},${y1} ${x + r},${y1}H${x + w - r}Q${x + w},${y1} ${x + w},${y1 - r}V${y0}Z`;
  }
  function barPath(y, hgt, x0, x1) {
    const r = Math.min(4, hgt / 2, Math.abs(x1 - x0));
    if (x1 >= x0) return `M${x0},${y}H${x1 - r}Q${x1},${y} ${x1},${y + r}V${y + hgt - r}Q${x1},${y + hgt} ${x1 - r},${y + hgt}H${x0}Z`;
    return `M${x0},${y}H${x1 + r}Q${x1},${y} ${x1},${y + r}V${y + hgt - r}Q${x1},${y + hgt} ${x1 + r},${y + hgt}H${x0}Z`;
  }
  const tipAttr = (series, cat, v) => `data-tip="${esc(`Série "${series}" · Ponto "${cat}" · Valor: ${fmtVal(v)}`)}"`;

  function render(chart, S) {
    const W = chart.w, H = chart.h;
    const data = S.chartData(chart);
    const out = [];
    const title = chart.title !== null && chart.title !== undefined ? chart.title : data.series.length === 1 ? data.series[0].name : 'Título do Gráfico';
    let top = 14;
    if (title) { out.push(`<text x="${W / 2}" y="27" text-anchor="middle" class="ct-title">${esc(clip(title, W - 30, 15))}</text>`); top = 44; }
    if (!data.series.length || !data.cats.length) {
      out.push(`<text x="${W / 2}" y="${H / 2}" text-anchor="middle" class="ct-ax">Não há números no intervalo do gráfico.</text>`);
      return wrap(out, W, H);
    }
    // pizza: só a primeira série; mais de 6 fatias viram "Outros" (cores nunca se repetem)
    if (chart.type === 'pie') {
      let items = data.cats.map((c, i) => ({ label: c, v: data.series[0].vals[i] })).filter((x) => x.v !== null && x.v > 0);
      if (items.length > 6) {
        const rest = items.slice(5).reduce((s, x) => s + x.v, 0);
        items = items.slice(0, 5).concat([{ label: 'Outros', v: rest }]);
      }
      const legendH = chart.legend !== false ? 30 : 0;
      const total = items.reduce((s, x) => s + x.v, 0) || 1;
      const cx = W / 2, cy = top + (H - top - legendH) / 2;
      const r = Math.max(20, Math.min(W / 2 - 20, (H - top - legendH) / 2 - 8));
      let a0 = -Math.PI / 2;
      items.forEach((it, i) => {
        const a1 = a0 + (it.v / total) * Math.PI * 2;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = (a) => `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;
        const d = items.length === 1 ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0Z` : `M${cx},${cy}L${p(a0)}A${r},${r} 0 ${large},1 ${p(a1)}Z`;
        out.push(`<path d="${d}" fill="${PALETTE[i]}" stroke="#fff" stroke-width="2" class="ct-mark" ${tipAttr(data.series[0].name, it.label, it.v)}/>`);
        if (chart.labels) {
          const am = (a0 + a1) / 2, pct = Math.round((it.v / total) * 100);
          if (a1 - a0 > 0.35) out.push(`<text x="${cx + r * 0.62 * Math.cos(am)}" y="${cy + r * 0.62 * Math.sin(am) + 4}" text-anchor="middle" class="ct-pie-lbl">${pct}%</text>`);
        }
        a0 = a1;
      });
      if (legendH) out.push(legend(items.map((x) => x.label), W, H));
      return wrap(out, W, H);
    }
    const showLegend = chart.legend !== false && data.series.length > 1;
    const legendH = showLegend ? 28 : 0;
    const vals = [].concat(...data.series.map((s) => s.vals)).filter((v) => v !== null);
    const sc = niceScale(Math.min(0, ...vals), Math.max(0, ...vals));
    const ns = data.series.length, n = data.cats.length;
    if (chart.type === 'bar') {
      const labW = Math.min(120, Math.max(...data.cats.map((c) => textW(c))) + 12);
      const L = 10 + labW, R = W - 20, T = top, B = H - legendH - 24;
      const x = (v) => L + ((v - sc.lo) / (sc.hi - sc.lo)) * (R - L);
      sc.ticks.forEach((t) => {
        out.push(`<line x1="${x(t)}" x2="${x(t)}" y1="${T}" y2="${B}" stroke="${t === 0 ? AXIS : GRID}" stroke-width="1"/>`);
        out.push(`<text x="${x(t)}" y="${B + 16}" text-anchor="middle" class="ct-ax">${fmtAxis(t)}</text>`);
      });
      const band = (B - T) / n;
      const barH = Math.max(3, Math.min(24, (band * 0.72 - 2 * (ns - 1)) / ns));
      const groupH = ns * barH + 2 * (ns - 1);
      data.cats.forEach((cat, i) => {
        const cy = T + band * (i + 0.5);
        out.push(`<text x="${L - 8}" y="${cy + 4}" text-anchor="end" class="ct-ax">${esc(clip(cat, labW - 8))}</text>`);
        data.series.forEach((s, j) => {
          const v = s.vals[i];
          if (v === null) return;
          const y = cy - groupH / 2 + j * (barH + 2);
          out.push(`<path d="${barPath(y, barH, x(0), x(v))}" fill="${PALETTE[j % PALETTE.length]}" class="ct-mark" ${tipAttr(s.name, cat, v)}/>`);
          if (chart.labels) out.push(`<text x="${x(v) + (v >= 0 ? 5 : -5)}" y="${y + barH / 2 + 4}" text-anchor="${v >= 0 ? 'start' : 'end'}" class="ct-lbl">${fmtVal(v)}</text>`);
        });
      });
    } else {
      const labW = Math.max(...sc.ticks.map((t) => textW(fmtAxis(t)))) + 10;
      const L = 12 + labW, R = W - 16, T = top, B = H - legendH - 26;
      const y = (v) => B - ((v - sc.lo) / (sc.hi - sc.lo)) * (B - T);
      sc.ticks.forEach((t) => {
        out.push(`<line x1="${L}" x2="${R}" y1="${y(t)}" y2="${y(t)}" stroke="${t === 0 ? AXIS : GRID}" stroke-width="1"/>`);
        out.push(`<text x="${L - 7}" y="${y(t) + 4}" text-anchor="end" class="ct-ax">${fmtAxis(t)}</text>`);
      });
      const band = (R - L) / n;
      data.cats.forEach((cat, i) => out.push(`<text x="${L + band * (i + 0.5)}" y="${B + 17}" text-anchor="middle" class="ct-ax">${esc(clip(cat, band - 4))}</text>`));
      if (chart.type === 'line') {
        data.series.forEach((s, j) => {
          const color = PALETTE[j % PALETTE.length];
          const pts = s.vals.map((v, i) => (v === null ? null : [L + band * (i + 0.5), y(v)]));
          const d = pts.filter(Boolean).map((p, k) => (k ? 'L' : 'M') + p[0] + ',' + p[1]).join('');
          out.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`);
          pts.forEach((p, i) => {
            if (!p) return;
            out.push(`<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="${color}" stroke="#fff" stroke-width="2" class="ct-mark" ${tipAttr(s.name, data.cats[i], s.vals[i])}/>`);
            if (chart.labels) out.push(`<text x="${p[0]}" y="${p[1] - 9}" text-anchor="middle" class="ct-lbl">${fmtVal(s.vals[i])}</text>`);
          });
        });
      } else {
        const barW = Math.max(3, Math.min(24, (band * 0.72 - 2 * (ns - 1)) / ns));
        const groupW = ns * barW + 2 * (ns - 1);
        data.cats.forEach((cat, i) => {
          const cx = L + band * (i + 0.5);
          data.series.forEach((s, j) => {
            const v = s.vals[i];
            if (v === null) return;
            const x = cx - groupW / 2 + j * (barW + 2);
            out.push(`<path d="${colPath(x, barW, y(0), y(v))}" fill="${PALETTE[j % PALETTE.length]}" class="ct-mark" ${tipAttr(s.name, cat, v)}/>`);
            if (chart.labels) out.push(`<text x="${x + barW / 2}" y="${v >= 0 ? y(v) - 5 : y(v) + 13}" text-anchor="middle" class="ct-lbl">${fmtVal(v)}</text>`);
          });
        });
      }
    }
    if (showLegend) out.push(legend(data.series.map((s) => s.name), W, H));
    return wrap(out, W, H);
  }

  function legend(names, W, H) {
    const items = names.map((nm, i) => ({ nm: clip(nm, 110), w: 16 + textW(clip(nm, 110), 11.5) + 14, color: PALETTE[i % PALETTE.length] }));
    const total = items.reduce((s, x) => s + x.w, 0);
    let x = Math.max(8, (W - total) / 2);
    const y = H - 14;
    return items.map((it) => {
      const g = `<rect x="${x}" y="${y - 9}" width="10" height="10" rx="2" fill="${it.color}"/><text x="${x + 15}" y="${y}" class="ct-leg">${esc(it.nm)}</text>`;
      x += it.w;
      return g;
    }).join('');
  }
  const wrap = (out, W, H) => `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg" role="img">${out.join('')}</svg>`;

  // ---------- Objetos de gráfico na grade ----------
  function draw(grid) {
    const app = grid.app;
    const S = grid.sheet;
    const layer = grid.chartLayer;
    const ids = new Set(S.charts.map((c) => c.id));
    [...layer.children].forEach((n) => { if (n.classList.contains('xl-chart') && !ids.has(n.dataset.id)) n.remove(); });
    for (const ch of S.charts) {
      let node = layer.querySelector(`.xl-chart[data-id="${ch.id}"]`);
      if (!node) {
        node = h('div.xl-chart', { dataset: { id: ch.id, type: ch.type } }, h('div.xl-chart-svg'), h('div.xl-chart-handle', { title: 'Arraste para redimensionar' }));
        layer.appendChild(node);
        bind(node, app, ch.id);
      }
      node.dataset.type = ch.type;
      node.dataset.range = XL.rangeToStr(ch.g);
      Object.assign(node.style, { left: ch.x + 'px', top: ch.y + 'px', width: ch.w + 'px', height: ch.h + 'px' });
      node.classList.toggle('sel', app.selectedChart === ch.id);
      node.firstChild.innerHTML = render(ch, S);
    }
  }

  function bind(node, app, id) {
    const find = () => app.sheet.charts.find((c) => c.id === id);
    node.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      const g = app.grid;
      if (g.editing && !g.commit()) return;
      app.selectChart(id);
      const ch = find();
      if (!ch) return;
      const resize = e.target.classList.contains('xl-chart-handle');
      const x0 = e.clientX, y0 = e.clientY, start = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
      let moved = false;
      const mm = (ev) => {
        const dx = ev.clientX - x0, dy = ev.clientY - y0;
        if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
        if (!moved) { moved = true; app.pushUndo(); }
        if (resize) { ch.w = Math.max(240, start.w + dx); ch.h = Math.max(160, start.h + dy); node.style.width = ch.w + 'px'; node.style.height = ch.h + 'px'; node.firstChild.innerHTML = render(ch, app.sheet); }
        else { ch.x = Math.max(0, start.x + dx); ch.y = Math.max(0, start.y + dy); node.style.left = ch.x + 'px'; node.style.top = ch.y + 'px'; }
      };
      const mu = () => {
        document.removeEventListener('mousemove', mm);
        document.removeEventListener('mouseup', mu);
        if (moved) app.emit(resize ? 'chartresize' : 'chartmove', { id });
        app.grid.focus();
      };
      document.addEventListener('mousemove', mm);
      document.addEventListener('mouseup', mu);
    });
    node.addEventListener('click', (e) => { e.stopPropagation(); if (app.selectedChart !== id) app.selectChart(id); });
    // dica com o valor ao passar o mouse
    node.addEventListener('mousemove', (e) => {
      const t = e.target.closest && e.target.closest('[data-tip]');
      const tip = XL.Charts.tipEl();
      if (!t) { tip.style.display = 'none'; return; }
      tip.textContent = t.getAttribute('data-tip');
      Object.assign(tip.style, { display: '', left: e.clientX + 14 + 'px', top: e.clientY + 12 + 'px' });
    });
    node.addEventListener('mouseleave', () => { XL.Charts.tipEl().style.display = 'none'; });
  }

  let tip = null;
  XL.Charts = {
    PALETTE, TYPES, render, draw,
    tipEl() { if (!tip) { tip = h('div.xl-charttip', { style: { display: 'none' } }); document.body.appendChild(tip); } return tip; },
  };
})();
