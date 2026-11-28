/* Lista de Campos da Tabela Dinâmica (arrastar entre Filtros, Colunas, Linhas e Valores),
   Segmentação de Dados e Linha do Tempo desenhadas sobre a grade. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const proto = XL.ExcelApp.prototype;
  const norm = (s) => XL.normName(String(s));
  const MONTHS = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const AREA_ICON = { filters: '⏷', cols: '⫼', rows: '☰', vals: 'Σ' };

  // ======================= Lista de Campos =======================
  class PivotPane {
    constructor(app) {
      this.app = app;
      this.pid = null;
      this.query = '';
      this.el = h('aside.pv-pane', { dataset: { role: 'pv-pane' } });
      this.el.addEventListener('xl-dnd', (e) => {
        const info = this.infoOf(e.target);
        if (!info) return;
        const to = e.detail && e.detail.to;
        const area = to && to.closest ? to.closest('.pv-area') : null;
        this.drop(info, area, e.detail && e.detail.y);
      });
    }

    infoOf(node) {
      const chip = node.closest && node.closest('.pv-chip');
      if (chip) return { field: chip.dataset.field, from: { area: chip.dataset.area, i: +chip.dataset.i } };
      const f = node.closest && node.closest('.pv-field');
      if (f) return { field: f.dataset.field };
      return null;
    }

    render(p) {
      const app = this.app;
      this.pid = p.id;
      const fields = XL.pivotFields(app.wb, p);
      const used = (name) => p.rows.some((x) => norm(x) === norm(name)) || p.cols.some((x) => norm(x) === norm(name)) ||
        p.filters.some((x) => norm(x.field) === norm(name)) || p.vals.some((x) => norm(x.field) === norm(name));
      const search = h('input.pv-search', { type: 'text', placeholder: 'Pesquisar', value: this.query });
      const list = h('div.pv-fields');
      const fill = () => {
        list.innerHTML = '';
        for (const f of fields) {
          if (this.query && !norm(f.name).includes(norm(this.query))) continue;
          const cb = h('input', { type: 'checkbox', dataset: { field: f.name } });
          cb.checked = used(f.name);
          cb.addEventListener('change', () => app.pivotToggleField(f.name, cb.checked));
          const row = h('div.pv-field', { dataset: { field: f.name }, title: 'Marque, ou arraste para uma das áreas abaixo' },
            cb, h('span.pv-fname', f.name), f.calc ? h('span.pv-tag', 'calc.') : null);
          row.addEventListener('mousedown', (e) => { if (e.target !== cb) this.startDrag(e, { field: f.name }); });
          list.appendChild(row);
        }
      };
      fill();
      search.addEventListener('input', () => { this.query = search.value; fill(); });
      const area = (key) => {
        const items = key === 'vals' ? p.vals.map((v) => ({ field: v.field, label: XL.pivotValName(v) }))
          : key === 'filters' ? p.filters.map((x) => ({ field: x.field, label: x.field }))
            : p[key].map((x) => ({ field: x, label: x }));
        return h('div.pv-area', { dataset: { area: key } },
          h('div.pv-area-title', h('span.pv-area-ico', AREA_ICON[key]), XL.PIVOT_AREAS[key]),
          h('div.pv-chips', items.map((it, i) => {
            const caret = h('button.pv-caret', { type: 'button', title: 'Opções do campo', html: XL.icons.caret });
            const chip = h('div.pv-chip', { dataset: { field: it.field, area: key, i, name: it.label } }, h('span.pv-chip-lbl', it.label), caret);
            caret.addEventListener('mousedown', (e) => e.stopPropagation());
            caret.addEventListener('click', (e) => { e.stopPropagation(); this.chipMenu(p, key, i, it, caret); });
            chip.addEventListener('mousedown', (e) => this.startDrag(e, { field: it.field, from: { area: key, i } }));
            return chip;
          })));
      };
      this.el.innerHTML = '';
      this.el.append(
        h('div.pv-head', h('span', 'Campos da Tabela Dinâmica'),
          h('button.pv-close', { type: 'button', title: 'Fechar a lista de campos', dataset: { role: 'pv-close' }, html: '&times;', onclick: () => app.exec('pivot-fieldlist') })),
        h('div.pv-sub', 'Escolha os campos para adicionar ao relatório:'),
        search, list,
        h('div.pv-sub', 'Arraste os campos entre as áreas abaixo:'),
        h('div.pv-areas', area('filters'), area('cols'), area('rows'), area('vals')));
    }

    chipMenu(p, key, i, it, anchor) {
      const app = this.app;
      const mv = (label, to, extra) => ({ label, data: Object.assign({ pv: 'move-' + to }, extra || {}), onClick: () => app.pivotMoveField(it.field, to, { from: { area: key, i } }) });
      const list = key === 'vals' ? p.vals : key === 'filters' ? p.filters : p[key];
      const items = [];
      if (i > 0) items.push({ label: 'Mover para Cima', data: { pv: 'up' }, onClick: () => app.pivotMoveField(it.field, key, { from: { area: key, i }, index: i - 1 }) });
      if (i < list.length - 1) items.push({ label: 'Mover para Baixo', data: { pv: 'down' }, onClick: () => app.pivotMoveField(it.field, key, { from: { area: key, i }, index: i + 1 }) });
      if (items.length) items.push({ sep: true });
      if (key !== 'filters') items.push(mv('Mover para Filtros de Relatório', 'filters'));
      if (key !== 'rows') items.push(mv('Mover para Rótulos de Linha', 'rows'));
      if (key !== 'cols') items.push(mv('Mover para Rótulos de Coluna', 'cols'));
      if (key !== 'vals') items.push(mv('Mover para Valores', 'vals'));
      items.push({ sep: true }, { label: 'Remover Campo', data: { pv: 'remove' }, onClick: () => app.pivotMoveField(it.field, null, { from: { area: key, i } }) });
      if (key === 'vals') items.push({ sep: true }, { label: 'Configurações do Campo de Valor...', data: { pv: 'settings' }, onClick: () => app.pivotSettingsDialog(i) });
      XL.ui.menu(anchor, items);
    }

    // arrastar com o mouse (com "fantasma" seguindo o ponteiro)
    startDrag(e, info) {
      if (e.button !== 0) return;
      e.preventDefault();
      const x0 = e.clientX, y0 = e.clientY;
      let ghost = null, over = null;
      const mm = (ev) => {
        if (!ghost && Math.abs(ev.clientX - x0) + Math.abs(ev.clientY - y0) < 5) return;
        if (!ghost) { ghost = h('div.pv-ghost', info.field); document.body.appendChild(ghost); }
        Object.assign(ghost.style, { left: ev.clientX + 10 + 'px', top: ev.clientY + 8 + 'px' });
        const el = document.elementFromPoint(ev.clientX, ev.clientY);
        const a = el && el.closest ? el.closest('.pv-area') : null;
        if (a !== over) { if (over) over.classList.remove('over'); over = a; if (over) over.classList.add('over'); }
      };
      const mu = (ev) => {
        document.removeEventListener('mousemove', mm);
        document.removeEventListener('mouseup', mu);
        if (over) over.classList.remove('over');
        if (!ghost) return;
        ghost.remove();
        const el = document.elementFromPoint(ev.clientX, ev.clientY);
        const a = el && el.closest ? el.closest('.pv-area') : null;
        if (a || (info.from && !(el && el.closest && el.closest('.pv-pane')))) this.drop(info, a, ev.clientY);
        else if (info.from && el && el.closest && el.closest('.pv-fields')) this.drop(info, null, ev.clientY);
      };
      document.addEventListener('mousemove', mm);
      document.addEventListener('mouseup', mu);
    }

    drop(info, areaEl, y) {
      const app = this.app;
      if (!areaEl) { if (info.from) app.pivotMoveField(info.field, null, { from: info.from }); return; }
      const area = areaEl.dataset.area;
      const chips = [...areaEl.querySelectorAll('.pv-chip')];
      let index = chips.length;
      if (y !== undefined && y !== null) { index = chips.filter((c) => { const r = c.getBoundingClientRect(); return r.top + r.height / 2 < y; }).length; }
      if (info.from && info.from.area === area && info.from.i < index) index--;
      if (info.from && info.from.area === area && info.from.i === index) return;
      app.pivotMoveField(info.field, area, { from: info.from, index });
    }
  }

  Object.assign(proto, {
    updatePivotPane(force) {
      if (!this.grid) return;
      const p = this.activePivot();
      const show = !!p && !this.pivotPaneOff;
      if (show && !this.pivotPane) {
        this.pivotPane = new PivotPane(this);
        this.midEl = h('div.xl-mid');
        this.root.insertBefore(this.midEl, this.grid.el);
        this.midEl.append(this.grid.el, this.pivotPane.el);
      }
      if (!this.pivotPane) return;
      const was = this.pivotPane.el.style.display !== 'none';
      this.pivotPane.el.style.display = show ? '' : 'none';
      if (show && (force || this.pivotPane.pid !== p.id || !was)) this.pivotPane.render(p);
      if (show !== was) { this.grid.positionOverlays(); this.ribbon.fit(); }
    },

    // ======================= Segmentação de Dados / Linha do Tempo =======================
    drawSlicers() {
      const g = this.grid;
      const layer = g.chartLayer;
      const S = this.sheet;
      const list = S.slicers || [];
      const ids = new Set(list.map((s) => s.id));
      layer.querySelectorAll('.xl-slicer').forEach((n) => { if (!ids.has(n.dataset.id)) n.remove(); });
      for (const sl of list) {
        let node = layer.querySelector(`.xl-slicer[data-id="${sl.id}"]`);
        if (!node) {
          node = h('div.xl-slicer' + (sl.kind === 'timeline' ? '.timeline' : ''), { dataset: { id: sl.id, field: sl.field, kind: sl.kind } });
          layer.appendChild(node);
          this.bindSlicer(node, sl.id);
        }
        Object.assign(node.style, { left: sl.x + 'px', top: sl.y + 'px', width: sl.w + 'px', height: sl.h + 'px' });
        node.innerHTML = '';
        const p = S.pivotById(sl.pivots[0]);
        if (!p) { node.append(h('div.sl-head', h('span.sl-title', sl.field)), h('div.cd-note', 'Tabela Dinâmica removida.')); continue; }
        if (sl.kind === 'timeline') this.fillTimeline(node, sl, p);
        else this.fillSlicer(node, sl, p);
      }
    },

    fillSlicer(node, sl, p) {
      const items = XL.pivotItems(this.wb, p, sl.field);
      const live = XL.pivotLiveItems(this.wb, p, this.sheet, sl.field, sl);
      const on = (lbl) => !sl.sel || sl.sel.includes(lbl);
      node.append(
        h('div.sl-head',
          h('span.sl-title', sl.field),
          h('button.sl-tool' + (sl.multi ? '.on' : ''), { type: 'button', title: 'Seleção Múltipla (ou segure Ctrl ao clicar)', dataset: { role: 'sl-multi' }, html: '<svg viewBox="0 0 16 16" width="14" height="14"><path d="M2 4h7M2 8h7M2 12h7" stroke="currentColor" stroke-width="1.4"/><path d="m10.5 7 1.5 1.5L15 5.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>' }),
          h('button.sl-tool' + (sl.sel ? '.active' : ''), { type: 'button', title: 'Limpar Filtro', dataset: { role: 'sl-clear' }, html: '<svg viewBox="0 0 16 16" width="14" height="14"><path d="M1.5 2.5h10L8 7v4.5l-2 1.2V7z" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="m10.5 9.5 4 4m0-4-4 4" stroke="' + (sl.sel ? '#C42B1C' : 'currentColor') + '" stroke-width="1.4"/></svg>' })),
        h('div.sl-items', items.map((it) => h('button.sl-item' + (on(it.label) ? '.on' : '') + (live.has(it.label) ? '' : '.nodata'), { type: 'button', dataset: { item: it.label }, title: it.label }, it.label))));
    },

    timelineMonths(p, field) {
      const rg = XL.pivotDateRange(this.wb, p, field);
      if (!rg) return [];
      const a = XL.serialToParts(rg.from), b = XL.serialToParts(rg.to);
      const out = [];
      for (let y = a.y, m = a.m; y < b.y || (y === b.y && m <= b.m);) {
        const from = XL.dateToSerial(y, m, 1);
        const to = XL.dateToSerial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1;
        out.push({ y, m, from, to, key: y + '-' + m });
        m++; if (m > 12) { m = 1; y++; }
      }
      return out;
    },

    fillTimeline(node, sl, p) {
      const months = this.timelineMonths(p, sl.field);
      const inR = (mo) => sl.range && mo.from >= sl.range.from && mo.to <= sl.range.to;
      let cap = 'Todos os Períodos';
      if (sl.range) {
        const a = XL.serialToParts(sl.range.from), b = XL.serialToParts(sl.range.to);
        cap = a.y === b.y && a.m === b.m ? `${MONTHS[a.m - 1]} de ${a.y}` : `${MONTHS[a.m - 1]} de ${a.y} – ${MONTHS[b.m - 1]} de ${b.y}`;
      }
      const years = h('div.tl-years');
      const strip = h('div.tl-strip');
      months.forEach((mo, i) => {
        years.appendChild(h('span.tl-year', { style: { width: '44px' } }, i === 0 || mo.m === 1 ? String(mo.y) : ''));
        strip.appendChild(h('div.tl-month' + (inR(mo) ? '.on' : ''), { dataset: { ym: mo.key }, title: `${MONTHS[mo.m - 1]} de ${mo.y}` }, MONTHS[mo.m - 1]));
      });
      node.append(
        h('div.sl-head', h('span.sl-title', sl.field),
          h('button.sl-tool' + (sl.range ? '.active' : ''), { type: 'button', title: 'Limpar Filtro', dataset: { role: 'sl-clear' }, html: '<svg viewBox="0 0 16 16" width="14" height="14"><path d="M1.5 2.5h10L8 7v4.5l-2 1.2V7z" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="m10.5 9.5 4 4m0-4-4 4" stroke="' + (sl.range ? '#C42B1C' : 'currentColor') + '" stroke-width="1.4"/></svg>' })),
        h('div.tl-cap', h('span', cap), h('span.tl-level', 'MESES')),
        h('div.tl-scroll', years, strip));
    },

    bindSlicer(node, id) {
      const find = () => this.sheet.slicers.find((s) => s.id === id);
      const monthOf = (el) => { const m = el && el.closest && el.closest('.tl-month'); return m ? m.dataset.ym : null; };
      const applyRange = (sl, k1, k2) => {
        const p = this.sheet.pivotById(sl.pivots[0]);
        const ms = this.timelineMonths(p, sl.field);
        const i1 = ms.findIndex((m) => m.key === k1), i2 = ms.findIndex((m) => m.key === k2);
        if (i1 < 0 || i2 < 0) return;
        const a = ms[Math.min(i1, i2)], b = ms[Math.max(i1, i2)];
        this.slicerSet(sl, (s) => { s.range = { from: a.from, to: b.to }; }, 'timelinepick', { field: sl.field, from: a.key, to: b.key });
      };
      let tlStart = null;
      node.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        e.stopPropagation();
        const sl = find();
        if (!sl) return;
        if (this.grid.editing && !this.grid.commit()) return;
        if (this.selectedChart) this.selectChart(null);
        const ym = monthOf(e.target);
        if (ym) {
          e.preventDefault();
          tlStart = ym;
          let last = ym;
          const mm = (ev) => {
            const k = monthOf(document.elementFromPoint(ev.clientX, ev.clientY));
            if (!k || k === last) return;
            last = k;
            const keys = [...node.querySelectorAll('.tl-month')].map((x) => x.dataset.ym);
            const i1 = keys.indexOf(tlStart), i2 = keys.indexOf(k);
            node.querySelectorAll('.tl-month').forEach((x, i) => x.classList.toggle('on', i >= Math.min(i1, i2) && i <= Math.max(i1, i2)));
          };
          const mu = () => { document.removeEventListener('mousemove', mm); document.removeEventListener('mouseup', mu); applyRange(sl, tlStart, last); tlStart = '__done'; };
          document.addEventListener('mousemove', mm);
          document.addEventListener('mouseup', mu);
          return;
        }
        if (!e.target.closest('.sl-head') || e.target.closest('.sl-tool')) return;
        e.preventDefault();
        const x0 = e.clientX, y0 = e.clientY, start = { x: sl.x, y: sl.y };
        let moved = false;
        const mm = (ev) => {
          const dx = ev.clientX - x0, dy = ev.clientY - y0;
          if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
          if (!moved) { moved = true; this.pushUndo(); }
          sl.x = Math.max(0, start.x + dx); sl.y = Math.max(0, start.y + dy);
          node.style.left = sl.x + 'px'; node.style.top = sl.y + 'px';
        };
        const mu = () => { document.removeEventListener('mousemove', mm); document.removeEventListener('mouseup', mu); if (moved) this.emit('slicermove', { field: sl.field }); this.grid.focus(); };
        document.addEventListener('mousemove', mm);
        document.addEventListener('mouseup', mu);
      });
      node.addEventListener('click', (e) => {
        e.stopPropagation();
        const sl = find();
        if (!sl) return;
        const tool = e.target.closest('.sl-tool');
        if (tool) {
          if (tool.dataset.role === 'sl-multi') this.slicerSet(sl, (s) => { s.multi = !s.multi; }, 'slicermulti', { field: sl.field, multi: !sl.multi });
          else if (sl.kind === 'timeline') this.slicerSet(sl, (s) => { s.range = null; }, 'timelinepick', { field: sl.field, from: null });
          else this.slicerSet(sl, (s) => { s.sel = null; }, 'slicerpick', { field: sl.field, sel: null });
          return;
        }
        const it = e.target.closest('.sl-item');
        if (it) { this.slicerPick(sl, it.dataset.item, e.ctrlKey || e.metaKey); return; }
        const ym = monthOf(e.target);
        if (ym) { if (tlStart === '__done') { tlStart = null; return; } applyRange(sl, ym, ym); }
      });
      // demonstração: arrastar de um mês até outro
      node.addEventListener('xl-dnd', (e) => {
        const sl = find();
        const a = monthOf(e.target), b = monthOf(e.detail && e.detail.to);
        if (sl && a && b) applyRange(sl, a, b);
      });
      node.addEventListener('contextmenu', (e) => {
        e.preventDefault(); e.stopPropagation();
        const sl = find();
        if (!sl) return;
        XL.ui.menu({ x: e.clientX, y: e.clientY }, [
          { label: 'Conexões de Relatório...', data: { pv: 'sl-conn' }, onClick: () => this.slicerConnectionsDialog(sl) },
          { label: `Limpar Filtro de "${sl.field}"`, data: { pv: 'sl-clear' }, onClick: () => this.slicerSet(sl, (s) => { s.sel = null; s.range = null; }, sl.kind === 'timeline' ? 'timelinepick' : 'slicerpick', { field: sl.field, sel: null }) },
          { sep: true },
          { label: `Remover "${sl.field}"`, data: { pv: 'sl-remove' }, onClick: () => this.slicerDelete(sl) },
        ]);
      });
    },
  });

  // a lista de campos acompanha a planilha ativa (trocar de aba, desfazer...)
  const baseAfter = proto.afterWorkbookChange;
  proto.afterWorkbookChange = function () {
    baseAfter.call(this);
    if (this.pivotPane) this.updatePivotPane(true);
  };
})();
