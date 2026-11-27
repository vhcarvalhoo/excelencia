/* Faixa de Opções: guias, grupos e botões. Cada botão tem data-cmd para o guia de lições destacar. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const I = XL.icons;

  const FONTS = ['Aptos Narrow', 'Aptos', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Georgia', 'Segoe UI', 'Times New Roman', 'Verdana'];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const NUMFMTS = [
    ['geral', 'Geral'], ['numero', 'Número'], ['moeda', 'Moeda'], ['contabil', 'Contábil'],
    ['data', 'Data Abreviada'], ['datalonga', 'Data Completa'], ['hora', 'Hora'], ['porcentagem', 'Porcentagem'], ['texto', 'Texto'], ['custom', 'Personalizado...'],
  ];

  // Recursos que ainda serão ensinados — mostram onde aparecem na trilha.
  const SOON = {
    merge: 'Mesclar células: <b>Nível Intermediário</b> (dica: prefira "Centralizar na Seleção").',
    orient: 'Orientação do texto: <b>Nível Intermediário</b>.',
    indent: 'Recuo: <b>Nível Intermediário</b>.',
    pivot: 'Tabelas Dinâmicas: <b>Nível Avançado</b>.',
    scatter: 'Gráfico de Dispersão: <b>Nível Avançado</b> (Análise de dados).',
    sparkline: 'Minigráficos: <b>Nível Avançado</b> (Dashboards).',
    slicer: 'Segmentação de Dados e Linha do Tempo: <b>Nível Avançado</b>.',
    pictures: 'Imagens e ilustrações não fazem parte das lições atuais.',
    powerquery: 'Power Query (Obter e Transformar Dados): <b>Nível Expert</b> — a ferramenta mais poderosa para importar e tratar dados!',
    refresh: 'Atualizar consultas: <b>Nível Expert</b> (Power Query).',
    whatif: 'Teste de Hipóteses (Atingir Meta, Cenários): <b>Nível Avançado</b>.',
    forecast: 'Planilha de Previsão: <b>Nível Avançado</b>.',
    names: 'Nomes definidos: <b>Nível Intermediário</b>.',
    audit: 'Auditoria de fórmulas: <b>Nível Intermediário</b>.',
    automate: 'Macros, VBA e Office Scripts: <b>Nível Expert</b> (Automação).',
    generic: 'Este recurso ainda não está disponível no simulador.',
  };
  XL.SOON = SOON;

  class Ribbon {
    constructor(app) {
      this.app = app;
      this.active = 'home';
      this.lastBorder = 'bottom';
      this.lastFill = '#FFFF00';
      this.lastColor = '#C00000';
      this.tabsEl = h('div.xl-tabs', { role: 'tablist' });
      this.bodyEl = h('div.xl-ribbon');
      this.tabs = this.defineTabs();
      this.renderTabBar();
      this.setTab('home', true);
    }

    // Guias contextuais (Design da Tabela, Design do Gráfico) só aparecem quando fazem sentido.
    visible(t) { return !t.contextual || !!t.contextual(); }
    renderTabBar() {
      this.tabsEl.innerHTML = '';
      for (const t of this.tabs) {
        if (!this.visible(t)) continue;
        const b = h('button.xl-tab' + (t.contextual ? '.ctx' : ''), {
          type: 'button', role: 'tab', dataset: { tab: t.id },
          onmousedown: (e) => e.preventDefault(),
          onclick: () => this.setTab(t.id),
        }, t.label);
        if (t.id === 'file') b.classList.add('file');
        b.classList.toggle('active', t.id === this.active);
        this.tabsEl.appendChild(b);
      }
    }
    reveal(b) {
      const T = this.tabsEl;
      if (!b || !T.isConnected) return;
      const tr = T.getBoundingClientRect(), br = b.getBoundingClientRect();
      if (br.right > tr.right) T.scrollLeft += br.right - tr.right + 12;
      else if (br.left < tr.left) T.scrollLeft -= tr.left - br.left + 12;
    }
    updateContext() {
      const app = this.app;
      if (!app.grid) return;
      const vis = this.tabs.filter((t) => t.contextual && this.visible(t)).map((t) => t.id).join(',');
      if (vis !== this.ctxVis) {
        this.ctxVis = vis;
        this.renderTabBar();
        // em janelas estreitas a guia contextual nasce fora da tela: rola a barra até ela
        const ctx = this.tabsEl.querySelectorAll('.xl-tab.ctx');
        if (ctx.length) this.reveal(ctx[ctx.length - 1]);
      }
      const cur = this.tabs.find((t) => t.id === this.active);
      const t = app.activeTable ? app.activeTable() : null;
      const key = (t ? t.name : '') + '|' + (app.selectedChart || '');
      if (cur && cur.contextual && !this.visible(cur)) this.setTab('home', true);
      else if (cur && cur.contextual && key !== this.ctxKey) this.setTab(this.active, true);
      this.ctxKey = key;
    }

    setTab(id, silent) {
      if (id === 'file') {
        // Arquivo > Imprimir: o que mais se usa da guia Arquivo
        this.app.exec('print-preview');
        return;
      }
      const want = this.tabs.find((t) => t.id === id);
      if (!want || !this.visible(want)) id = 'home';
      this.active = id;
      if (want && want.contextual && !this.tabsEl.querySelector(`[data-tab="${id}"]`)) this.renderTabBar();
      this.tabsEl.querySelectorAll('.xl-tab').forEach((b) => b.classList.toggle('active', b.dataset.tab === id));
      this.reveal(this.tabsEl.querySelector('.xl-tab.active'));
      const tab = this.tabs.find((t) => t.id === id);
      this.bodyEl.innerHTML = '';
      if (tab.groups) tab.groups.forEach((g) => this.bodyEl.appendChild(this.renderGroup(g)));
      else this.bodyEl.appendChild(h('div.rb-empty', tab.empty || 'Esta guia não é usada nas lições atuais.'));
      this.fit();
      this.sync();
      if (!silent) this.app.emit('tab', { tab: id });
      this.app.grid && this.app.grid.focus();
    }

    // ---------- Renderização ----------
    renderGroup(g) {
      const body = h('div.rb-group-body');
      g.items.forEach((it) => body.appendChild(this.renderItem(it)));
      return h('div.rb-group', body, h('div.rb-group-label', g.label));
    }

    renderItem(it) {
      switch (it.t) {
        case 'col': return h('div.rb-col', it.items.map((x) => this.renderItem(x)));
        case 'row': return h('div.rb-row', it.items.map((x) => this.renderItem(x)));
        case 'sep': return h('span.rb-sep');
        case 'select': return this.renderSelect(it);
        case 'check': return this.renderCheck(it);
        case 'input': return this.renderInput(it);
        default: return this.renderButton(it);
      }
    }

    iconHtml(it) {
      if (it.glyph) return `<span class="rb-glyph">${it.glyph}</span>`;
      return it.icon ? I[it.icon] || '' : '';
    }

    renderButton(it) {
      const big = it.t === 'big';
      const app = this.app;
      const main = h('button.rb-btn' + (big ? '.rb-big' : ''), {
        type: 'button', title: it.title || (it.label || '').replace('\n', ' '),
        dataset: { cmd: it.cmd || '', arg: it.arg !== undefined ? String(it.arg) : '' },
        onmousedown: (e) => e.preventDefault(),
      });
      const iconWrap = h('span.rb-ico', { html: this.iconHtml(it) });
      if (it.colorBar) {
        const bar = h('span.rb-colorbar', { style: { background: it.cmd === 'fill' ? this.lastFill : this.lastColor } });
        iconWrap.appendChild(bar);
        main.dataset.colorbar = it.cmd;
      }
      main.appendChild(iconWrap);
      if (it.label) main.appendChild(h('span.rb-lbl', { html: XL.escapeHtml(it.label).replace('\n', '<br>') + (it.menuOnly && big ? ' ' + I.caret : '') }));
      if (it.menuOnly && !big) main.appendChild(h('span.rb-caret', { html: I.caret }));

      const openMenu = (anchor) => {
        if (it.cmd === 'fill' || it.cmd === 'color') {
          XL.ui.palette(anchor, (c) => {
            if (it.cmd === 'fill') this.lastFill = c || this.lastFill; else this.lastColor = c || this.lastColor;
            this.updateColorBars();
            app.exec(it.cmd, c);
          }, { noneLabel: it.cmd === 'fill' ? 'Sem Preenchimento' : 'Automático' });
          return;
        }
        const items = (typeof it.menu === 'function' ? it.menu() : it.menu).map((m) => m.sep || m.header ? m : ({
          label: m.label, icon: m.icon ? I[m.icon] : m.glyph ? `<span class="rb-glyph sm">${m.glyph}</span>` : '', hint: m.hint, swatch: m.swatch,
          data: { cmd: m.cmd, arg: m.arg !== undefined ? String(m.arg) : '' },
          onClick: () => {
            if (it.cmd === 'border' && m.cmd === 'border') this.lastBorder = m.arg;
            app.exec(m.cmd, m.arg);
          },
        }));
        XL.ui.menu(anchor, items);
      };

      if (it.menuOnly) {
        main.dataset.menu = it.menuId || it.cmd || '';
        main.addEventListener('click', () => openMenu(main));
        return main;
      }
      main.addEventListener('click', () => {
        let arg = it.arg;
        if (it.cmd === 'border') arg = this.lastBorder;
        if (it.cmd === 'fill') arg = this.lastFill;
        if (it.cmd === 'color') arg = this.lastColor;
        app.exec(it.cmd, arg);
      });
      if (!it.menu && !it.palette) return main;
      const arrow = h('button.rb-btn.rb-arrow', {
        type: 'button', title: (it.title || it.label || '') + ' (mais opções)',
        dataset: { menu: it.menuId || it.cmd },
        onmousedown: (e) => e.preventDefault(),
        html: I.caret,
      });
      arrow.addEventListener('click', () => openMenu(arrow));
      return h('span.rb-split' + (big ? '.big' : ''), main, arrow);
    }

    renderSelect(it) {
      const sel = h('select.rb-select', { title: it.title, dataset: { cmd: it.cmd }, style: { width: it.width + 'px' } },
        it.options.map((o) => Array.isArray(o) ? h('option', { value: o[0] }, o[1]) : h('option', { value: String(o) }, String(o))));
      sel.addEventListener('change', () => {
        const v = sel.value;
        this.app.exec(it.cmd, it.cmd === 'size' ? Number(v) : v);
      });
      this['sel_' + it.cmd] = sel;
      return it.label ? h('label.rb-sellbl', h('span', it.label), sel) : sel;
    }

    renderInput(it) {
      const inp = h('input.rb-input', { type: 'text', title: it.title || '', dataset: { cmd: it.cmd }, style: { width: it.width + 'px' } });
      inp.value = it.get() || '';
      const commit = () => { if (inp.isConnected && inp.value !== (it.get() || '')) this.app.exec(it.cmd, inp.value); };
      inp.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Enter') { e.preventDefault(); commit(); this.app.grid.focus(); }
        else if (e.key === 'Escape') { inp.value = it.get() || ''; this.app.grid.focus(); }
      });
      inp.addEventListener('change', commit);
      return it.label ? h('label.rb-sellbl.rb-inlbl', h('span', it.label), inp) : inp;
    }

    renderCheck(it) {
      const cb = h('input', { type: 'checkbox', dataset: { cmd: it.cmd } });
      cb.checked = !!it.get();
      cb.addEventListener('change', () => this.app.exec(it.cmd, cb.checked));
      return h('label.rb-check', cb, it.label);
    }

    // Em qual guia está um botão? (usado pelas demonstrações do guia)
    findTabFor(pred) {
      for (const t of this.tabs) {
        if (!t.groups) continue;
        let found = false;
        const walk = (items) => items.forEach((it) => {
          if (found || !it) return;
          if (pred(it)) found = true;
          else if (it.items) walk(it.items);
        });
        t.groups.forEach((g) => walk(g.items));
        if (found) return t.id;
      }
      return null;
    }

    // Como no Excel: se a janela for estreita, os botões pequenos perdem o rótulo.
    fit() {
      const b = this.bodyEl;
      b.classList.remove('compact', 'compact2');
      if (!b.isConnected) return;
      if (b.scrollWidth > b.clientWidth + 2) b.classList.add('compact');
      if (b.scrollWidth > b.clientWidth + 2) b.classList.add('compact2');
    }

    updateColorBars() {
      this.bodyEl.querySelectorAll('[data-colorbar]').forEach((b) => {
        const bar = b.querySelector('.rb-colorbar');
        if (bar) bar.style.background = b.dataset.colorbar === 'fill' ? this.lastFill : this.lastColor;
      });
    }

    // Reflete o formato da célula ativa (negrito ligado, fonte, formato de número...).
    sync() {
      const app = this.app;
      if (!app.grid) return;
      const { ar, ac } = app.grid.sel;
      const f = app.sheet.fmt(ar, ac);
      const on = (cmd, arg, v) => {
        const sel = `.rb-btn[data-cmd="${cmd}"]` + (arg !== undefined ? `[data-arg="${arg}"]` : '');
        this.bodyEl.querySelectorAll(sel).forEach((b) => b.classList.toggle('on', !!v));
      };
      on('bold', undefined, f.b);
      on('italic', undefined, f.i);
      on('underline', undefined, f.u);
      on('wrap', undefined, f.wrap);
      ['left', 'center', 'right'].forEach((a) => on('align', a, f.align === a));
      ['top', 'middle', 'bottom'].forEach((a) => on('valign', a, f.valign === a));
      on('show-formulas', undefined, app.grid.showFormulas);
      if (this.sel_font && this.bodyEl.contains(this.sel_font)) this.sel_font.value = f.font || 'Aptos Narrow';
      if (this.sel_size && this.bodyEl.contains(this.sel_size)) this.sel_size.value = String(f.size || 11);
      if (this.sel_numfmt && this.bodyEl.contains(this.sel_numfmt)) this.sel_numfmt.value = f.numFmt === 'custom' ? 'custom' : f.numFmt || 'geral';
      on('filter-toggle', undefined, !!app.sheet.filter);
      const p = app.sheet.page;
      const fw = this['sel_page-fitw'], fh = this['sel_page-fith'];
      if (fw && this.bodyEl.contains(fw)) fw.value = p.fit === 'sheet' || p.fit === 'cols' ? '1' : 'auto';
      if (fh && this.bodyEl.contains(fh)) fh.value = p.fit === 'sheet' || p.fit === 'rows' ? '1' : 'auto';
      const pg = this.bodyEl.querySelector('input[data-cmd="print-gridlines"]');
      if (pg) pg.checked = !!p.printGridlines;
    }

    // ---------- Definição das guias ----------
    defineTabs() {
      const app = this.app;
      const soon = (key, icon, label, t = 'btn', extra = {}) => Object.assign({ t, cmd: 'soon', arg: key, icon, label }, extra);
      const fnMenu = (cat) => () => XL.Functions.byCategory(cat).map((d) => ({ label: d.name, cmd: 'insert-fn', arg: d.name, hint: '' }));
      const autosumMenu = [
        { label: 'Soma', cmd: 'autosum', arg: 'SOMA', glyph: 'Σ' },
        { label: 'Média', cmd: 'autosum', arg: 'MÉDIA' },
        { label: 'Contar Números', cmd: 'autosum', arg: 'CONT.NÚM' },
        { label: 'Máx.', cmd: 'autosum', arg: 'MÁXIMO' },
        { label: 'Mín.', cmd: 'autosum', arg: 'MÍNIMO' },
        { sep: true },
        { label: 'Mais Funções...', cmd: 'fx' },
      ];
      const borderMenu = [
        { label: 'Borda Inferior', cmd: 'border', arg: 'bottom' },
        { label: 'Borda Superior', cmd: 'border', arg: 'top' },
        { label: 'Borda Esquerda', cmd: 'border', arg: 'left' },
        { label: 'Borda Direita', cmd: 'border', arg: 'right' },
        { sep: true },
        { label: 'Sem Borda', cmd: 'border', arg: 'none' },
        { label: 'Todas as Bordas', cmd: 'border', arg: 'all', icon: 'borders' },
        { label: 'Bordas Externas', cmd: 'border', arg: 'outer' },
      ];
      const cellStyles = [
        { header: 'Bom, Ruim e Neutro' },
        { label: 'Bom', cmd: 'cellstyle', arg: 'bom', swatch: 'background:#c6efce' },
        { label: 'Ruim', cmd: 'cellstyle', arg: 'ruim', swatch: 'background:#ffc7ce' },
        { label: 'Neutro', cmd: 'cellstyle', arg: 'neutro', swatch: 'background:#ffeb9c' },
        { header: 'Títulos' },
        { label: 'Título', cmd: 'cellstyle', arg: 'titulo' },
        { label: 'Total', cmd: 'cellstyle', arg: 'total' },
        { label: 'Cabeçalho verde', cmd: 'cellstyle', arg: 'cabecalho', swatch: 'background:#107c41' },
        { sep: true },
        { label: 'Normal (remover estilo)', cmd: 'cellstyle', arg: 'normal' },
      ];

      const TS = XL.TABLE_STYLES;
      const tblGlyph = (k) => `<span class="tbl-sw"><i style="background:${TS[k].head}"></i><i style="background:${TS[k].band}"></i><i></i><i style="background:${TS[k].band}"></i></span>`;
      const tableMenu = [{ header: 'Estilos de Tabela' }].concat(Object.keys(TS).map((k) => ({
        label: 'Estilo ' + TS[k].label, cmd: 'table-create', arg: k, swatch: `background:linear-gradient(${TS[k].head} 0 34%, ${TS[k].band} 34% 67%, #fff 67%)`,
      })));
      const grad = (cols) => `background:linear-gradient(90deg, ${cols.join(', ')})`;
      const cfMenu = [
        { header: 'Realçar Regras das Células' },
        { label: 'É Maior do que...', cmd: 'cf', arg: 'gt' },
        { label: 'É Menor do que...', cmd: 'cf', arg: 'lt' },
        { label: 'Está Entre...', cmd: 'cf', arg: 'between' },
        { label: 'É Igual a...', cmd: 'cf', arg: 'eq' },
        { label: 'Texto que Contém...', cmd: 'cf', arg: 'text' },
        { label: 'Valores Duplicados...', cmd: 'cf', arg: 'dup' },
        { header: 'Regras de Primeiros/Últimos' },
        { label: 'Primeiros Itens...', cmd: 'cf', arg: 'top' },
        { label: 'Últimos Itens...', cmd: 'cf', arg: 'bottom' },
        { label: 'Acima da Média...', cmd: 'cf', arg: 'above' },
        { label: 'Abaixo da Média...', cmd: 'cf', arg: 'below' },
        { header: 'Barras de Dados' },
        ...Object.keys(XL.CF_BARS).map((k) => ({ label: 'Barra ' + k[0].toUpperCase() + k.slice(1), cmd: 'cf', arg: 'bar-' + k, swatch: grad([XL.CF_BARS[k], XL.CF_BARS[k] + '33']) })),
        { header: 'Escalas de Cor' },
        ...Object.keys(XL.CF_SCALES).map((k) => ({ label: XL.CF_SCALES[k].label, cmd: 'cf', arg: 'scale-' + k, swatch: grad(XL.CF_SCALES[k].colors) })),
        { sep: true },
        { label: 'Nova Regra (usar uma fórmula)...', cmd: 'cf', arg: 'formula' },
        { sep: true },
        { label: 'Limpar Regras das Células Selecionadas', cmd: 'cf-clear', arg: 'sel' },
        { label: 'Limpar Regras da Planilha Inteira', cmd: 'cf-clear', arg: 'all' },
      ];
      const chartTypes = [
        { label: 'Colunas Agrupadas', arg: 'col', icon: 'chartCol' },
        { label: 'Barras Agrupadas', arg: 'bar', icon: 'chartBar' },
        { label: 'Linhas', arg: 'line', icon: 'chartLine' },
        { label: 'Pizza', arg: 'pie', icon: 'chartPie' },
      ];
      const tbl = () => (app.grid && app.activeTable ? app.activeTable() : null);
      const chart = () => (app.grid && app.activeChart ? app.activeChart() : null);

      return [
        { id: 'file', label: 'Arquivo' },
        {
          id: 'home', label: 'Página Inicial', groups: [
            { label: 'Área de Transferência', items: [
              { t: 'big', cmd: 'paste', icon: 'paste', label: 'Colar', title: 'Colar (Ctrl+V)', menuId: 'paste', menu: [
                { label: 'Colar', cmd: 'paste', icon: 'paste' },
                { label: 'Colar Valores', cmd: 'paste', arg: 'values', glyph: '123' },
                { label: 'Colar Formatação', cmd: 'paste', arg: 'formats', icon: 'painter' },
              ] },
              { t: 'col', items: [
                { cmd: 'cut', icon: 'cut', title: 'Recortar (Ctrl+X)' },
                { cmd: 'copy', icon: 'copy', title: 'Copiar (Ctrl+C)' },
                { cmd: 'painter', icon: 'painter', title: 'Pincel de Formatação: copia o formato de uma célula para outra' },
              ] },
            ] },
            { label: 'Fonte', items: [{ t: 'col', items: [
              { t: 'row', items: [
                { t: 'select', cmd: 'font', options: FONTS, width: 128, title: 'Fonte' },
                { t: 'select', cmd: 'size', options: SIZES, width: 54, title: 'Tamanho da Fonte' },
                { cmd: 'grow', glyph: 'A<sup>+</sup>', title: 'Aumentar Tamanho da Fonte' },
                { cmd: 'shrink', glyph: 'A<sup>−</sup>', title: 'Diminuir Tamanho da Fonte' },
              ] },
              { t: 'row', items: [
                { cmd: 'bold', glyph: '<b>N</b>', title: 'Negrito (Ctrl+N)' },
                { cmd: 'italic', glyph: '<i style="font-family:Georgia,serif">I</i>', title: 'Itálico (Ctrl+I)' },
                { cmd: 'underline', glyph: '<u>S</u>', title: 'Sublinhado (Ctrl+S)' },
                { t: 'sep' },
                { cmd: 'border', icon: 'borders', title: 'Bordas', menuId: 'border', menu: borderMenu },
                { cmd: 'fill', icon: 'fill', title: 'Cor de Preenchimento', colorBar: true, palette: true, menuId: 'fill' },
                { cmd: 'color', glyph: '<span style="font-weight:600">A</span>', title: 'Cor da Fonte', colorBar: true, palette: true, menuId: 'color' },
              ] },
            ] }] },
            { label: 'Alinhamento', items: [{ t: 'col', items: [
              { t: 'row', items: [
                { cmd: 'valign', arg: 'top', icon: 'vTop', title: 'Alinhar em Cima' },
                { cmd: 'valign', arg: 'middle', icon: 'vMid', title: 'Alinhar no Meio' },
                { cmd: 'valign', arg: 'bottom', icon: 'vBottom', title: 'Alinhar Embaixo' },
                { t: 'sep' },
                soon('orient', 'orient', null, 'btn', { title: 'Orientação' }),
                { cmd: 'wrap', icon: 'wrap', title: 'Quebrar Texto Automaticamente' },
              ] },
              { t: 'row', items: [
                { cmd: 'align', arg: 'left', icon: 'alignLeft', title: 'Alinhar à Esquerda' },
                { cmd: 'align', arg: 'center', icon: 'alignCenter', title: 'Centralizar' },
                { cmd: 'align', arg: 'right', icon: 'alignRight', title: 'Alinhar à Direita' },
                { t: 'sep' },
                soon('indent', 'indentDec', null, 'btn', { title: 'Diminuir Recuo' }),
                soon('indent', 'indentInc', null, 'btn', { title: 'Aumentar Recuo' }),
                soon('merge', 'merge', null, 'btn', { title: 'Mesclar e Centralizar' }),
              ] },
            ] }] },
            { label: 'Número', items: [{ t: 'col', items: [
              { t: 'row', items: [{ t: 'select', cmd: 'numfmt', options: NUMFMTS, width: 150, title: 'Formato de Número' }] },
              { t: 'row', items: [
                { cmd: 'currency', icon: 'currency', title: 'Formato de Número de Contabilização (R$)' },
                { cmd: 'percent', glyph: '%', title: 'Estilo de Porcentagem' },
                { cmd: 'thousands', glyph: '<span style="font-size:11px">000</span>', title: 'Separador de Milhares' },
                { t: 'sep' },
                { cmd: 'dec-inc', glyph: '<span class="dec">←0<br>,00</span>', title: 'Aumentar Casas Decimais' },
                { cmd: 'dec-dec', glyph: '<span class="dec">,00<br>→0</span>', title: 'Diminuir Casas Decimais' },
                { cmd: 'format-cells-dialog', glyph: '<span style="font-size:12px">⋯</span>', title: 'Mais formatos: Formatar Células (Ctrl+1)' },
              ] },
            ] }] },
            { label: 'Estilos', items: [{ t: 'col', items: [
              { cmd: 'cf', icon: 'condFormat', label: 'Formatação Condicional', menuOnly: true, menuId: 'cf', menu: cfMenu },
              { cmd: 'table-create', icon: 'table', label: 'Formatar como Tabela', menuOnly: true, menuId: 'fmt-table', menu: tableMenu },
              { cmd: 'cellstyle', icon: 'cellStyles', label: 'Estilos de Célula', menuOnly: true, menuId: 'cellstyle', menu: cellStyles },
            ] }] },
            { label: 'Células', items: [{ t: 'col', items: [
              { cmd: 'insert-cells', icon: 'insertCells', label: 'Inserir', menuOnly: true, menuId: 'insert-cells', menu: [
                { label: 'Inserir Células...', cmd: 'insert-dialog', icon: 'insertCells' },
                { label: 'Inserir Linhas na Planilha', cmd: 'insert-rows' },
                { label: 'Inserir Colunas na Planilha', cmd: 'insert-cols' },
                { sep: true },
                { label: 'Inserir Planilha', cmd: 'sheet-add' },
              ] },
              { cmd: 'delete-cells', icon: 'deleteCells', label: 'Excluir', menuOnly: true, menuId: 'delete-cells', menu: [
                { label: 'Excluir Células...', cmd: 'delete-dialog', icon: 'deleteCells' },
                { label: 'Excluir Linhas da Planilha', cmd: 'delete-rows' },
                { label: 'Excluir Colunas da Planilha', cmd: 'delete-cols' },
                { sep: true },
                { label: 'Excluir Planilha', cmd: 'sheet-delete' },
              ] },
              { cmd: 'format-cells', icon: 'formatCells', label: 'Formatar', menuOnly: true, menuId: 'format-cells', menu: [
                { label: 'AutoAjuste da Largura da Coluna', cmd: 'autofit' },
                { label: 'Largura da Coluna...', cmd: 'colwidth' },
                { sep: true },
                { label: 'Renomear Planilha', cmd: 'sheet-rename' },
                { sep: true },
                { label: 'Formatar Células...', cmd: 'format-cells-dialog', hint: 'Ctrl+1' },
              ] },
            ] }] },
            { label: 'Edição', items: [
              { t: 'col', items: [
                { cmd: 'autosum', arg: 'SOMA', glyph: '<span style="font-size:17px">Σ</span>', label: 'AutoSoma', title: 'AutoSoma (Alt+=)', menuId: 'autosum', menu: autosumMenu },
                { cmd: 'fill-menu', icon: 'fillDown', label: 'Preencher', menuOnly: true, menuId: 'fill-menu', menu: [
                  { label: 'Para Baixo', cmd: 'fill-down', hint: 'Ctrl+D' },
                  { label: 'Preenchimento Relâmpago', cmd: 'flashfill', icon: 'flashFill', hint: 'Ctrl+E' },
                ] },
                { cmd: 'clear', icon: 'clear', label: 'Limpar', menuOnly: true, menuId: 'clear', menu: [
                  { label: 'Limpar Tudo', cmd: 'clear', arg: 'all' },
                  { label: 'Limpar Formatos', cmd: 'clear', arg: 'formats' },
                  { label: 'Limpar Conteúdo', cmd: 'clear', arg: 'contents' },
                ] },
              ] },
              { t: 'big', icon: 'sortFilter', label: 'Classificar\ne Filtrar', menuOnly: true, menuId: 'sortfilter', menu: [
                { label: 'Classificar de A a Z', cmd: 'sort', arg: 'asc', icon: 'sortAZ' },
                { label: 'Classificar de Z a A', cmd: 'sort', arg: 'desc', icon: 'sortZA' },
                { label: 'Personalizar Classificação...', cmd: 'sort-dialog' },
                { sep: true },
                { label: 'Filtro', cmd: 'filter-toggle', hint: 'Ctrl+Shift+L' },
                { label: 'Limpar', cmd: 'filter-clear' },
                { label: 'Reaplicar', cmd: 'filter-reapply' },
              ] },
              { t: 'big', icon: 'find', label: 'Localizar e\nSelecionar', menuOnly: true, menuId: 'findselect', menu: [
                { label: 'Localizar...', cmd: 'find', hint: 'Ctrl+L' },
                { label: 'Substituir...', cmd: 'replace', hint: 'Ctrl+U' },
                { label: 'Ir para...', cmd: 'goto' },
              ] },
            ] },
          ],
        },
        {
          id: 'insert', label: 'Inserir', groups: [
            { label: 'Tabelas', items: [soon('pivot', 'pivot', 'Tabela\nDinâmica', 'big'), { t: 'big', cmd: 'table-create', icon: 'tableBig', label: 'Tabela', title: 'Tabela (Ctrl+T)' }] },
            { label: 'Ilustrações', items: [soon('pictures', 'pictures', 'Imagens', 'big')] },
            { label: 'Gráficos', items: [
              { t: 'big', cmd: 'chart-insert', arg: 'rec', icon: 'chartRec', label: 'Gráficos\nRecomendados', title: 'Gráficos Recomendados: o simulador escolhe o tipo pelos seus dados' },
              { t: 'col', items: [
                { t: 'row', items: [{ cmd: 'chart-insert', arg: 'col', icon: 'chartCol', title: 'Inserir Gráfico de Colunas' }, { cmd: 'chart-insert', arg: 'bar', icon: 'chartBar', title: 'Inserir Gráfico de Barras' }] },
                { t: 'row', items: [{ cmd: 'chart-insert', arg: 'line', icon: 'chartLine', title: 'Inserir Gráfico de Linhas' }, { cmd: 'chart-insert', arg: 'pie', icon: 'chartPie', title: 'Inserir Gráfico de Pizza' }] },
                { t: 'row', items: [soon('scatter', 'chartScatter', null, 'btn', { title: 'Dispersão' })] },
              ] },
            ] },
            { label: 'Minigráficos', items: [soon('sparkline', 'sparkline', 'Minigráficos', 'big')] },
            { label: 'Filtros', items: [soon('slicer', 'slicer', 'Segmentação\nde Dados', 'big')] },
          ],
        },
        { id: 'draw', label: 'Desenhar', empty: 'A guia Desenhar (caneta e marca-texto) não é usada nas lições.' },
        {
          id: 'layout', label: 'Layout da Página', groups: [
            { label: 'Configurar Página', items: [
              { t: 'big', icon: 'margins', label: 'Margens', menuOnly: true, menuId: 'page-margins', menu: [
                { label: 'Normal', cmd: 'page-margins', arg: 'normal', hint: 'Sup. 1,9 · Esq. 1,8 cm' },
                { label: 'Estreita', cmd: 'page-margins', arg: 'estreita', hint: 'Sup. 1,9 · Esq. 0,6 cm' },
                { label: 'Larga', cmd: 'page-margins', arg: 'larga', hint: '2,54 cm em todos os lados' },
              ] },
              { t: 'big', icon: 'orientation', label: 'Orientação', menuOnly: true, menuId: 'page-orient', menu: [
                { label: 'Retrato', cmd: 'page-orient', arg: 'retrato' },
                { label: 'Paisagem', cmd: 'page-orient', arg: 'paisagem' },
              ] },
              { t: 'big', icon: 'paperSize', label: 'Tamanho', menuOnly: true, menuId: 'page-size', menu: [
                { label: 'A4', cmd: 'page-size', arg: 'A4', hint: '21 x 29,7 cm' },
                { label: 'Carta', cmd: 'page-size', arg: 'Carta', hint: '21,6 x 27,9 cm' },
                { label: 'Ofício', cmd: 'page-size', arg: 'Ofício', hint: '21,6 x 35,6 cm' },
              ] },
              { t: 'big', icon: 'printArea', label: 'Área de\nImpressão', menuOnly: true, menuId: 'print-area', menu: [
                { label: 'Definir Área de Impressão', cmd: 'print-area', arg: 'set' },
                { label: 'Limpar Área de Impressão', cmd: 'print-area', arg: 'clear' },
              ] },
              { t: 'big', cmd: 'print-titles', icon: 'printTitles', label: 'Imprimir\nTítulos' },
            ] },
            { label: 'Dimensionar para Ajustar', items: [{ t: 'col', items: [
              { t: 'select', cmd: 'page-fitw', label: 'Largura:', options: [['auto', 'Automática'], ['1', '1 página']], width: 104, title: 'Largura da impressão' },
              { t: 'select', cmd: 'page-fith', label: 'Altura:', options: [['auto', 'Automática'], ['1', '1 página']], width: 104, title: 'Altura da impressão' },
            ] }] },
            { label: 'Opções de Planilha', items: [{ t: 'col', items: [
              { t: 'check', cmd: 'print-gridlines', label: 'Imprimir linhas de grade', get: () => app.sheet.page.printGridlines },
            ] }] },
            { label: 'Visualizar', items: [{ t: 'big', cmd: 'print-preview', icon: 'preview', label: 'Visualizar\nImpressão', title: 'Visualizar Impressão (Ctrl+P)' }] },
          ],
        },
        {
          id: 'formulas', label: 'Fórmulas', groups: [
            { label: 'Biblioteca de Funções', items: [
              { t: 'big', cmd: 'fx', icon: 'fx', label: 'Inserir\nFunção', title: 'Inserir Função (Shift+F3)' },
              { t: 'big', cmd: 'autosum', arg: 'SOMA', icon: 'sigmaBig', label: 'AutoSoma', menuOnly: true, menuId: 'autosum-big', menu: autosumMenu },
              { t: 'big', icon: 'finance', label: 'Financeira', menuOnly: true, menuId: 'fn-fin', menu: fnMenu('Financeira') },
              { t: 'big', icon: 'logic', label: 'Lógica', menuOnly: true, menuId: 'fn-log', menu: fnMenu('Lógica') },
              { t: 'big', icon: 'text', label: 'Texto', menuOnly: true, menuId: 'fn-txt', menu: fnMenu('Texto') },
              { t: 'big', icon: 'dateTime', label: 'Data e\nHora', menuOnly: true, menuId: 'fn-dt', menu: fnMenu('Data e Hora') },
              { t: 'big', icon: 'lookup', label: 'Pesquisa e\nReferência', menuOnly: true, menuId: 'fn-pr', menu: fnMenu('Pesquisa e Referência') },
              { t: 'big', icon: 'math', label: 'Matemática e\nTrigonometria', menuOnly: true, menuId: 'fn-mat', menu: fnMenu('Matemática e Trigonometria') },
              { t: 'big', icon: 'more', label: 'Mais\nFunções', menuOnly: true, menuId: 'fn-est', menu: fnMenu('Estatística') },
            ] },
            { label: 'Nomes Definidos', items: [soon('names', 'names', 'Gerenciador\nde Nomes', 'big')] },
            { label: 'Auditoria de Fórmulas', items: [
              { t: 'col', items: [
                soon('audit', 'tracePrec', 'Rastrear Precedentes'),
                soon('audit', 'traceDep', 'Rastrear Dependentes'),
              ] },
              { t: 'col', items: [
                { cmd: 'show-formulas', icon: 'showFormulas', label: 'Mostrar Fórmulas', title: 'Mostrar Fórmulas (Ctrl+`)' },
                soon('audit', 'errorCheck', 'Verificação de Erros'),
                soon('audit', 'evalFormula', 'Avaliar Fórmula'),
              ] },
            ] },
            { label: 'Cálculo', items: [{ t: 'big', cmd: 'recalc', icon: 'calc', label: 'Calcular\nAgora' }] },
          ],
        },
        {
          id: 'data', label: 'Dados', groups: [
            { label: 'Obter e Transformar Dados', items: [
              soon('powerquery', 'getData', 'Obter\nDados', 'big'),
              { t: 'col', items: [soon('powerquery', 'csv', 'De Texto/CSV'), soon('powerquery', 'web', 'Da Web'), soon('powerquery', 'fromTable', 'De Tabela/Intervalo')] },
            ] },
            { label: 'Consultas e Conexões', items: [soon('refresh', 'refresh', 'Atualizar\nTudo', 'big')] },
            { label: 'Classificar e Filtrar', items: [
              { t: 'col', items: [
                { cmd: 'sort', arg: 'asc', icon: 'sortAZ', title: 'Classificar de A a Z (menor para o maior)' },
                { cmd: 'sort', arg: 'desc', icon: 'sortZA', title: 'Classificar de Z a A (maior para o menor)' },
              ] },
              { t: 'big', cmd: 'sort-dialog', icon: 'sortBig', label: 'Classificar' },
              { t: 'big', cmd: 'filter-toggle', icon: 'filter', label: 'Filtro', title: 'Filtro (Ctrl+Shift+L)' },
              { t: 'col', items: [
                { cmd: 'filter-clear', icon: 'clear', label: 'Limpar' },
                { cmd: 'filter-reapply', icon: 'refreshSm', label: 'Reaplicar' },
              ] },
            ] },
            { label: 'Ferramentas de Dados', items: [
              { t: 'big', cmd: 'texttocols', icon: 'textToCols', label: 'Texto para\nColunas' },
              { t: 'big', cmd: 'flashfill', icon: 'flashFillBig', label: 'Preenchimento\nRelâmpago', title: 'Preenchimento Relâmpago (Ctrl+E)' },
              { t: 'big', cmd: 'dedupe', icon: 'removeDup', label: 'Remover\nDuplicatas' },
              { t: 'big', cmd: 'dv-dialog', icon: 'validation', label: 'Validação\nde Dados', title: 'Validação de Dados: listas suspensas e regras de entrada', menuId: 'dv', menu: [
                { label: 'Validação de Dados...', cmd: 'dv-dialog', icon: 'validation' },
                { label: 'Limpar Validação das Células', cmd: 'dv-clear' },
              ] },
            ] },
            { label: 'Previsão', items: [soon('whatif', 'whatIf', 'Teste de\nHipóteses', 'big'), soon('forecast', 'forecast', 'Planilha de\nPrevisão', 'big')] },
          ],
        },
        { id: 'review', label: 'Revisão', empty: 'A guia Revisão (comentários, proteção) não é usada nas lições atuais.' },
        {
          id: 'view', label: 'Exibir', groups: [
            { label: 'Mostrar', items: [{ t: 'col', items: [
              { t: 'check', cmd: 'gridlines', label: 'Linhas de Grade', get: () => !app.grid || app.grid.gridlines },
              { t: 'check', cmd: 'headings', label: 'Títulos', get: () => !app.grid || app.grid.headings },
              { t: 'check', cmd: 'formulabar', label: 'Barra de Fórmulas', get: () => app.formulaBarVisible !== false },
            ] }] },
            { label: 'Janela', items: [{
              t: 'big', icon: 'freeze', label: 'Congelar\nPainéis', menuOnly: true, menuId: 'freeze',
              menu: () => [
                app.sheet.freeze
                  ? { label: 'Descongelar Painéis', cmd: 'freeze', arg: 'none', hint: 'libera linhas e colunas' }
                  : { label: 'Congelar Painéis', cmd: 'freeze', arg: 'panes', hint: 'acima e à esquerda da célula ativa' },
                { label: 'Congelar Linha Superior', cmd: 'freeze', arg: 'top' },
                { label: 'Congelar Primeira Coluna', cmd: 'freeze', arg: 'first' },
              ],
            }] },
          ],
        },
        {
          id: 'automate', label: 'Automatizar', groups: [
            { label: 'Automação', items: [soon('automate', 'script', 'Novo\nScript', 'big'), soon('automate', 'macro', 'Macros', 'big')] },
          ],
        },
        {
          id: 'help', label: 'Ajuda', groups: [
            { label: 'Ajuda', items: [{ t: 'big', cmd: 'shortcuts', icon: 'keyboard', label: 'Atalhos de\nTeclado' }] },
          ],
        },
        {
          id: 'tabledesign', label: 'Design da Tabela', contextual: () => !!tbl(), groups: [
            { label: 'Propriedades', items: [{ t: 'col', items: [
              { t: 'input', cmd: 'table-rename', label: 'Nome da Tabela:', width: 104, title: 'Nome da Tabela: usado nas fórmulas, ex.: =SOMA(Vendas[Total])', get: () => (tbl() || {}).name },
            ] }] },
            { label: 'Ferramentas', items: [{ t: 'col', items: [
              { cmd: 'dedupe', icon: 'removeDup', label: 'Remover Duplicatas' },
              { cmd: 'table-convert', icon: 'table', label: 'Converter em Intervalo', title: 'Converter em Intervalo: volta a ser células comuns (mantém a formatação)' },
            ] }] },
            { label: 'Opções de Estilo de Tabela', items: [{ t: 'col', items: [
              { t: 'check', cmd: 'table-totals', label: 'Linha de Totais', get: () => (tbl() || {}).totalRow },
              { t: 'check', cmd: 'table-banded', label: 'Linhas em Tiras', get: () => tbl() && tbl().banded !== false },
              { t: 'check', cmd: 'table-filterbtn', label: 'Botão de Filtro', get: () => { const t = tbl(); return !!(t && app.sheet.filter && app.sheet.filter.table === t.name); } },
            ] }] },
            { label: 'Estilos de Tabela', items: [{ t: 'row', items: Object.keys(TS).map((k) => ({ cmd: 'table-style', arg: k, glyph: tblGlyph(k), title: 'Estilo ' + TS[k].label })) }] },
          ],
        },
        {
          id: 'chartdesign', label: 'Design do Gráfico', contextual: () => !!chart(), groups: [
            { label: 'Tipo', items: [{ t: 'big', icon: 'chartCol', label: 'Alterar Tipo\nde Gráfico', menuOnly: true, menuId: 'chart-type', menu: chartTypes.map((x) => Object.assign({ cmd: 'chart-type' }, x)) }] },
            { label: 'Layouts de Gráfico', items: [{ t: 'col', items: [
              { cmd: 'chart-title', glyph: '<span style="font-size:11px;font-weight:700">T</span>', label: 'Título do Gráfico' },
              { t: 'check', cmd: 'chart-legend', label: 'Legenda', get: () => (chart() || {}).legend },
              { t: 'check', cmd: 'chart-labels', label: 'Rótulos de Dados', get: () => (chart() || {}).labels },
            ] }] },
            { label: 'Excluir', items: [{ t: 'big', cmd: 'chart-delete', icon: 'clear', label: 'Excluir\nGráfico', title: 'Excluir Gráfico (Delete)' }] },
          ],
        },
      ];
    }
  }

  XL.Ribbon = Ribbon;
})();
