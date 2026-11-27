/* Impressão simulada: calcula as páginas (tamanho, orientação, margens, dimensionamento, área e títulos)
   e mostra a Visualização de Impressão como no Excel. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const CM = 37.795; // pixels por centímetro (96 dpi)
  const PAPER = { A4: [794, 1123], Carta: [816, 1056], 'Ofício': [816, 1344] };
  const MARGINS = {
    normal: { t: 1.9, b: 1.9, l: 1.8, r: 1.8 },
    estreita: { t: 1.9, b: 1.9, l: 0.6, r: 0.6 },
    larga: { t: 2.54, b: 2.54, l: 2.54, r: 2.54 },
  };
  const DEFAULT_W = 80;

  const P = (XL.Print = { current: null, PAPER, MARGINS });

  const colW = (S, c) => S.colW[c] || DEFAULT_W;
  const rowH = (S, r, c1, c2) => {
    let m = 11;
    for (let c = c1; c <= c2; c++) { const f = S.fmt(r, c); if (f.size) m = Math.max(m, f.size); }
    return Math.round(m * 1.82);
  };

  // Divide a área de impressão em páginas.
  P.paginate = (S) => {
    const p = S.page;
    const area = p.printArea || S.usedRange();
    let [W, H] = PAPER[p.size] || PAPER.A4;
    if (p.orientation === 'paisagem') [W, H] = [H, W];
    const m = MARGINS[p.margins] || MARGINS.normal;
    const cw = W - (m.l + m.r) * CM, ch = H - (m.t + m.b) * CM;
    const out = { area, pages: [], colBreaks: [], rowBreaks: [], scale: 1, W, H, cw, ch, m };
    if (!area) return out;
    const cols = [], rows = [];
    for (let c = area.c1; c <= area.c2; c++) cols.push(c);
    for (let r = area.r1; r <= area.r2; r++) if (!S.hiddenRows.has(r)) rows.push(r);
    const totalW = cols.reduce((s, c) => s + colW(S, c), 0);
    const totalH = rows.reduce((s, r) => s + rowH(S, r, area.c1, area.c2), 0);
    const titles = [];
    if (p.titleRows) for (let r = p.titleRows.r1; r <= p.titleRows.r2; r++) titles.push(r);
    const titleH = titles.reduce((s, r) => s + rowH(S, r, area.c1, area.c2), 0);
    let scale = 1;
    if (p.fit === 'sheet') scale = Math.min(1, cw / totalW, ch / totalH);
    else if (p.fit === 'cols') scale = Math.min(1, cw / totalW);
    else if (p.fit === 'rows') scale = Math.min(1, ch / totalH);
    scale = Math.max(0.1, scale);
    out.scale = scale;
    // colunas
    const colChunks = [];
    let cur = [], w = 0;
    for (const c of cols) {
      const cwid = colW(S, c) * scale;
      if (cur.length && w + cwid > cw + 0.5) { colChunks.push(cur); cur = []; w = 0; }
      cur.push(c); w += cwid;
    }
    if (cur.length) colChunks.push(cur);
    // linhas (a partir da 2ª página, os títulos ocupam espaço no topo)
    const rowChunks = [];
    cur = []; let hgt = 0;
    for (const r of rows) {
      const avail = rowChunks.length && titles.length && !titles.includes(r) ? ch - titleH * scale : ch;
      const rh = rowH(S, r, area.c1, area.c2) * scale;
      if (cur.length && hgt + rh > avail + 0.5) { rowChunks.push(cur); cur = []; hgt = 0; }
      cur.push(r); hgt += rh;
    }
    if (cur.length) rowChunks.push(cur);
    // ordem do Excel: primeiro para baixo, depois para a direita
    for (const cc of colChunks) {
      rowChunks.forEach((rc) => {
        const rep = titles.length && rc[0] > titles[titles.length - 1] ? titles : [];
        out.pages.push({ cols: cc, rows: rep.concat(rc), titleCount: rep.length });
      });
    }
    out.colBreaks = colChunks.slice(1).map((x) => x[0]);
    out.rowBreaks = rowChunks.slice(1).map((x) => x[0]);
    return out;
  };

  // Desenha uma página como tabela.
  function renderPage(S, pg, info, k) {
    const page = h('div.pp-page', { style: { width: info.W * k + 'px', height: info.H * k + 'px' } });
    const content = h('div.pp-content', {
      style: {
        left: info.m.l * CM * k + 'px', top: info.m.t * CM * k + 'px',
        width: info.cw / (info.scale) + 'px', transform: `scale(${info.scale * k})`,
      },
    });
    const tbl = h('table.pp-table' + (S.page.printGridlines ? '.grid' : ''));
    const cg = h('colgroup');
    let tw = 0;
    pg.cols.forEach((c) => { cg.appendChild(h('col', { style: { width: colW(S, c) + 'px' } })); tw += colW(S, c); });
    tbl.appendChild(cg);
    tbl.style.width = tw + 'px';
    pg.rows.forEach((r, i) => {
      const tr = h('tr', { style: { height: rowH(S, r, pg.cols[0], pg.cols[pg.cols.length - 1]) + 'px' } });
      if (i < pg.titleCount) tr.classList.add('pp-title');
      pg.cols.forEach((c) => {
        const cell = S.cell(r, c);
        const f = (cell && cell.fmt) || {};
        const d = S.display(r, c);
        const td = h('td', d.text);
        Object.assign(td.style, {
          textAlign: f.align || d.align || 'left', fontWeight: f.b ? 'bold' : '', fontStyle: f.i ? 'italic' : '',
          textDecoration: f.u ? 'underline' : '', color: f.color || '', background: f.fill || '',
          fontSize: f.size ? f.size + 'pt' : '',
        });
        const b = f.border || {};
        ['t', 'r', 'b', 'l'].forEach((s) => { if (b[s]) td.style['border' + { t: 'Top', r: 'Right', b: 'Bottom', l: 'Left' }[s]] = '1px solid #000'; });
        tr.appendChild(td);
      });
      tbl.appendChild(tr);
    });
    content.appendChild(tbl);
    page.appendChild(content);
    return page;
  }

  // Visualização de Impressão (como Arquivo > Imprimir).
  P.preview = (app) => {
    const S = app.sheet;
    let idx = 0;
    const view = h('div.pp-view');
    const counter = h('span.pp-count');
    const pagesBadge = h('div.pp-badge');
    const mkSel = (label, key, opts, role) => {
      const sel = h('select.fd-cat', { dataset: { role } }, opts.map(([v, l]) => h('option', { value: v }, l)));
      sel.value = S.page[key];
      sel.addEventListener('change', () => app.setPage({ [key]: sel.value }));
      return h('label.pp-field', h('span', label), sel);
    };
    const gridCb = h('input', { type: 'checkbox', dataset: { role: 'gridlines' } });
    gridCb.checked = S.page.printGridlines;
    gridCb.addEventListener('change', () => app.setPage({ printGridlines: gridCb.checked }));
    const side = h('div.pp-side',
      h('button.btn.btn-primary.btn-block.pp-print', {
        type: 'button',
        onclick: () => {
          const n = P.paginate(app.sheet).pages.length;
          XL.ui.toast(`🖨️ Simulação: no Excel de verdade, ${n} ${n === 1 ? 'página seria enviada' : 'páginas seriam enviadas'} para a impressora agora.`);
          app.emit('print', { pages: n });
        },
      }, '🖨️ Imprimir'),
      pagesBadge,
      h('div.pp-sec', 'Configurações'),
      mkSel('Orientação', 'orientation', [['retrato', 'Orientação Retrato'], ['paisagem', 'Orientação Paisagem']], 'orientation'),
      mkSel('Tamanho do papel', 'size', [['A4', 'A4 (21 x 29,7 cm)'], ['Carta', 'Carta (21,6 x 27,9 cm)'], ['Ofício', 'Ofício (21,6 x 35,6 cm)']], 'size'),
      mkSel('Margens', 'margins', [['normal', 'Margens Normais'], ['estreita', 'Margens Estreitas'], ['larga', 'Margens Largas']], 'margins'),
      mkSel('Dimensionamento', 'fit', [['none', 'Sem Dimensionamento'], ['sheet', 'Ajustar Planilha em Uma Página'], ['cols', 'Ajustar Todas as Colunas em Uma Página'], ['rows', 'Ajustar Todas as Linhas em Uma Página']], 'fit'),
      h('label.sd-check.pp-check', gridCb, ' Imprimir linhas de grade'),
      h('div.pp-info'));
    const nav = h('div.pp-nav',
      h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => { idx = Math.max(0, idx - 1); draw(); } }, '◀'),
      counter,
      h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => { idx++; draw(); } }, '▶'));
    const body = h('div.pp', side, h('div.pp-main', view, nav));

    function draw() {
      const S2 = app.sheet;
      const info = P.paginate(S2);
      const n = info.pages.length;
      idx = Math.max(0, Math.min(idx, n - 1));
      view.innerHTML = '';
      if (!n) {
        view.appendChild(h('div.pp-empty', 'Não encontramos nada para imprimir. Digite algum dado na planilha primeiro.'));
        counter.textContent = '0 de 0';
        pagesBadge.textContent = '';
        return;
      }
      const availH = Math.max(300, (view.clientHeight || 520) - 20);
      const availW = Math.max(280, (view.clientWidth || 600) - 40);
      const k = Math.min(availH / info.H, availW / info.W);
      view.appendChild(renderPage(S2, info.pages[idx], info, k));
      counter.textContent = `${idx + 1} de ${n}`;
      pagesBadge.innerHTML = `📄 <b>${n}</b> ${n === 1 ? 'página' : 'páginas'}`;
      const p = S2.page;
      side.querySelector('.pp-info').innerHTML = [
        `Área: <b>${p.printArea ? XL.rangeToStr(p.printArea) : 'planilha inteira (' + XL.rangeToStr(info.area) + ')'}</b>`,
        `Títulos repetidos: <b>${p.titleRows ? `linhas ${p.titleRows.r1 + 1}${p.titleRows.r2 > p.titleRows.r1 ? '–' + (p.titleRows.r2 + 1) : ''}` : 'nenhum'}</b>`,
        `Escala: <b>${Math.round(info.scale * 100)}%</b>`,
      ].join('<br>');
    }

    // Como o "backstage" do Excel: cobre só a janela do Excel (o painel do guia continua livre).
    if (P.current) P.current.close(true);
    const onKey = (e) => { if (e.key === 'Escape' && !document.querySelector('.ui-modal-back')) { e.preventDefault(); e.stopPropagation(); close(); } };
    const close = (silent) => {
      document.removeEventListener('keydown', onKey, true);
      overlay.remove();
      P.current = null;
      if (silent) return;
      app.sheet.showBreaks = true;
      app.render();
      app.grid.focus();
      app.emit('printclose', {});
    };
    const overlay = h('div.pp-overlay',
      h('div.pp-top',
        h('button.btn.btn-ghost.btn-sm.pp-back', { type: 'button', onclick: () => close() }, '← Voltar à planilha'),
        h('span.pp-title', 'Imprimir'),
        h('span.pp-hint', 'Esc também fecha')),
      body);
    app.root.appendChild(overlay);
    document.addEventListener('keydown', onKey, true);
    P.current = {
      app,
      close,
      refresh: () => {
        side.querySelectorAll('select').forEach((s) => {
          const key = { orientation: 'orientation', size: 'size', margins: 'margins', fit: 'fit' }[s.dataset.role];
          if (key) s.value = app.sheet.page[key];
        });
        gridCb.checked = app.sheet.page.printGridlines;
        draw();
      },
    };
    requestAnimationFrame(draw);
    draw();
    app.emit('printpreview', { pages: P.paginate(S).pages.length });
  };
})();
