/* Janelas do Módulo 7: Criar Tabela, Título do Gráfico, Formatação Condicional e Validação de Dados. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const D = XL.Dialogs;
  const enterSubmits = (inp) => inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const b = inp.closest('.ui-modal').querySelector('.btn-primary'); if (b) b.click(); } });

  // ================= Criar Tabela =================
  D.createTable = (app, style) => {
    const S = app.sheet;
    const d = app.dataRegion();
    const g = d ? d.g : app.grid.range();
    if (S.tableAt(app.grid.sel.ar, app.grid.sel.ac)) { XL.ui.toast('Esta célula já faz parte de uma tabela.'); return; }
    app.grid.setRange(g);
    const inp = h('input.fd-search', { type: 'text', value: '=$' + XL.numToCol(g.c1 + 1) + '$' + (g.r1 + 1) + ':$' + XL.numToCol(g.c2 + 1) + '$' + (g.r2 + 1), dataset: { role: 'tbl-range' } });
    const cb = h('input', { type: 'checkbox', disabled: true });
    cb.checked = true;
    enterSubmits(inp);
    XL.ui.modal({
      title: 'Criar Tabela', width: 380,
      body: h('div', h('label.fd-lbl', 'Onde estão os dados da tabela?'), inp,
        h('label.sd-check', { style: 'margin-top:10px' }, cb, ' Minha tabela tem cabeçalhos'),
        h('p.cd-note', 'A primeira linha do intervalo vira o cabeçalho da tabela.')),
      buttons: [{ label: 'OK', primary: true, onClick: () => { if (!app.createTable(style, inp.value)) return false; } }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
    setTimeout(() => inp.select(), 0);
  };

  // ================= Título do gráfico =================
  D.chartTitle = (app) => {
    const ch = app.activeChart();
    if (!ch) { XL.ui.toast('Clique no gráfico primeiro para selecioná-lo.'); return; }
    const data = app.sheet.chartData(ch);
    const cur = ch.title !== null && ch.title !== undefined ? ch.title : data.series.length === 1 ? data.series[0].name : 'Título do Gráfico';
    const inp = h('input.fd-search', { type: 'text', value: cur, dataset: { role: 'chart-title' } });
    enterSubmits(inp);
    XL.ui.modal({
      title: 'Título do Gráfico', width: 380,
      body: h('div', h('label.fd-lbl', 'Texto do título (deixe vazio para não mostrar título):'), inp),
      buttons: [{ label: 'OK', primary: true, onClick: () => app.chartUpdate((c) => { c.title = inp.value.trim(); }) }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
    setTimeout(() => inp.select(), 0);
  };

  // ================= Formatação condicional =================
  const CF_TEXT = {
    gt: ['É Maior do que', 'Formatar células que são MAIORES DO QUE:'],
    lt: ['É Menor do que', 'Formatar células que são MENORES DO QUE:'],
    between: ['Está Entre', 'Formatar células que estão ENTRE:'],
    eq: ['É Igual a', 'Formatar células que são IGUAIS A:'],
    text: ['Texto que Contém', 'Formatar células que contêm o texto:'],
    dup: ['Valores Duplicados', 'Formatar células que contêm valores:'],
    top: ['Primeiros Itens', 'Formatar células que estão entre os PRIMEIROS:'],
    bottom: ['Últimos Itens', 'Formatar células que estão entre os ÚLTIMOS:'],
    above: ['Acima da Média', 'Formatar células que estão ACIMA DA MÉDIA'],
    below: ['Abaixo da Média', 'Formatar células que estão ABAIXO DA MÉDIA'],
    formula: ['Nova Regra (fórmula)', 'Formatar valores em que esta fórmula é verdadeira:'],
  };
  D.conditional = (app, kind) => {
    const info = CF_TEXT[kind];
    if (!info) return;
    const S = app.sheet;
    const g = app.grid.range();
    const nums = [];
    XL.eachInRange(g, (r, c) => { const v = S.value(r, c); if (typeof v === 'number') nums.push(v); });
    const avg = nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0;
    const sug = Math.round(avg);
    const v1 = h('input.fd-search', { type: 'text', dataset: { role: 'cf-v1' } });
    const v2 = h('input.fd-search', { type: 'text', dataset: { role: 'cf-v2' } });
    if (['gt', 'lt', 'eq', 'between'].includes(kind)) v1.value = String(sug);
    if (kind === 'between') v2.value = String(nums.length ? Math.round(Math.max(...nums)) : 0);
    if (kind === 'top' || kind === 'bottom') v1.value = '3';
    if (kind === 'formula') v1.value = '=';
    const dupSel = h('select.fd-cat', { dataset: { role: 'cf-dup' } }, h('option', { value: 'dup' }, 'Duplicados'), h('option', { value: 'uniq' }, 'Exclusivos'));
    const styleSel = h('select.fd-cat', { dataset: { role: 'cf-style' } }, Object.keys(XL.CF_STYLES).map((k) => h('option', { value: k }, XL.CF_STYLES[k].label)));
    const swatch = h('span.cf-sample', 'AaBbCc 123');
    const paint = () => { const st = XL.CF_STYLES[styleSel.value]; Object.assign(swatch.style, { background: st.fill || '#fff', color: st.color || '#000', fontWeight: st.b ? '700' : '400' }); };
    styleSel.addEventListener('change', paint);
    paint();
    [v1, v2].forEach(enterSubmits);
    const rows = [h('p.cf-q', info[1])];
    if (['gt', 'lt', 'eq', 'text', 'top', 'bottom'].includes(kind)) rows.push(v1);
    if (kind === 'between') rows.push(h('div.cfd-row', v1, h('span', ' e '), v2));
    if (kind === 'dup') rows.push(dupSel);
    if (kind === 'formula') {
      rows.push(v1);
      rows.push(h('p.cd-note', 'Escreva a fórmula para a primeira célula da seleção (', XL.rangeToStr({ r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c1 }), '). Use $ para travar a coluna: ', h('code', '=$D2<1'), ' pinta a linha inteira.'));
    }
    rows.push(h('label.fd-lbl', 'com:'), h('div.cfd-row', styleSel, swatch));
    XL.ui.modal({
      title: info[0], width: 540,
      body: h('div.cfd', rows),
      buttons: [{
        label: 'OK', primary: true,
        onClick: () => {
          const rule = { kind: kind === 'dup' ? dupSel.value : kind, style: styleSel.value };
          if (['gt', 'lt', 'eq', 'between', 'text', 'top', 'bottom'].includes(kind)) {
            if (v1.value.trim() === '') { XL.ui.toast('Digite um valor.'); return false; }
            rule.v1 = v1.value.trim();
          }
          if (kind === 'between') rule.v2 = v2.value.trim();
          if (kind === 'formula') {
            const f = v1.value.trim();
            if (f.length < 2 || f[0] !== '=') { XL.ui.toast('A regra precisa ser uma fórmula começando com =.'); return false; }
            try { XL.Formula.parse(XL.Formula.normalize(f).slice(1)); } catch (e) { XL.ui.toast('Há um problema com esta fórmula: ' + XL.escapeHtml(e.message)); return false; }
            rule.formula = XL.Formula.normalize(f);
          }
          app.grid.setRange(g);
          app.addCF(rule);
        },
      }, { label: 'Cancelar' }],
      onClose: () => app.grid.focus(),
    });
    setTimeout(() => { if (rows.includes(v1) || kind === 'between') { v1.focus(); v1.select(); } }, 0);
  };

  // ================= Validação de dados =================
  D.validation = (app) => {
    const S = app.sheet;
    const g = app.grid.range();
    const cur = S.validationAt(g.r1, g.c1) || {};
    const type = h('select.fd-cat', { dataset: { role: 'dv-type' } },
      h('option', { value: 'any' }, 'Qualquer valor'), h('option', { value: 'whole' }, 'Número inteiro'), h('option', { value: 'decimal' }, 'Decimal'),
      h('option', { value: 'list' }, 'Lista'), h('option', { value: 'date' }, 'Data'), h('option', { value: 'textlen' }, 'Comprimento do texto'));
    type.value = cur.type || 'any';
    const op = h('select.fd-cat', { dataset: { role: 'dv-op' } },
      [['between', 'está entre'], ['notbetween', 'não está entre'], ['eq', 'é igual a'], ['ne', 'é diferente de'], ['gt', 'é maior do que'], ['lt', 'é menor do que'], ['ge', 'é maior ou igual a'], ['le', 'é menor ou igual a']]
        .map(([v, l]) => h('option', { value: v }, l)));
    op.value = cur.op || 'between';
    const vmin = h('input.fd-search', { type: 'text', value: cur.v1 !== undefined ? cur.v1 : '', dataset: { role: 'dv-min' } });
    const vmax = h('input.fd-search', { type: 'text', value: cur.v2 !== undefined ? cur.v2 : '', dataset: { role: 'dv-max' } });
    const source = h('input.fd-search', { type: 'text', value: cur.source || '', dataset: { role: 'dv-source' }, placeholder: 'Ex.: Pendente;Pago;Atrasado  ou  =$H$2:$H$5' });
    const inTitle = h('input.fd-search', { type: 'text', value: cur.inputTitle || '', dataset: { role: 'dv-intitle' } });
    const inMsg = h('input.fd-search', { type: 'text', value: cur.inputMsg || '', dataset: { role: 'dv-inmsg' } });
    const errTitle = h('input.fd-search', { type: 'text', value: cur.errorTitle || '', dataset: { role: 'dv-errtitle' } });
    const errMsg = h('input.fd-search', { type: 'text', value: cur.errorMsg || '', dataset: { role: 'dv-errmsg' } });
    const crit = h('div.dv-crit');
    const draw = () => {
      crit.innerHTML = '';
      const t = type.value;
      if (t === 'list') crit.append(h('label.fd-lbl', 'Fonte:'), source, h('p.cd-note', 'Itens separados por ponto e vírgula, ou um intervalo da planilha começando com = (ex.: =$H$2:$H$5).'));
      else if (t !== 'any') {
        crit.append(h('label.fd-lbl', 'Dados:'), op);
        const two = op.value === 'between' || op.value === 'notbetween';
        crit.append(h('label.fd-lbl', two ? 'Mínimo:' : 'Valor:'), vmin);
        if (two) crit.append(h('label.fd-lbl', 'Máximo:'), vmax);
        if (t === 'date') crit.append(h('p.cd-note', 'Use datas como 01/01/2026.'));
      }
    };
    type.addEventListener('change', draw);
    op.addEventListener('change', draw);
    draw();
    XL.ui.modal({
      title: 'Validação de Dados', width: 500,
      body: h('div.dv',
        h('div.dv-sec', 'Configurações'),
        h('label.fd-lbl', 'Permitir:'), type, crit,
        h('div.dv-sec', 'Mensagem de entrada (opcional)'),
        h('div.dv-two', h('div', h('label.fd-lbl', 'Título:'), inTitle), h('div', h('label.fd-lbl', 'Mensagem:'), inMsg)),
        h('div.dv-sec', 'Alerta de erro (opcional)'),
        h('div.dv-two', h('div', h('label.fd-lbl', 'Título:'), errTitle), h('div', h('label.fd-lbl', 'Mensagem:'), errMsg))),
      buttons: [
        {
          label: 'OK', primary: true,
          onClick: () => {
            const t = type.value;
            app.grid.setRange(g);
            if (t === 'any') { app.clearDV(); return; }
            const rule = { type: t, op: op.value, inputTitle: inTitle.value.trim(), inputMsg: inMsg.value.trim(), errorTitle: errTitle.value.trim(), errorMsg: errMsg.value.trim() };
            if (t === 'list') {
              if (!source.value.trim()) { XL.ui.toast('Informe a Fonte da lista.'); return false; }
              rule.source = source.value.trim();
            } else {
              if (vmin.value.trim() === '') { XL.ui.toast('Informe o valor.'); return false; }
              rule.v1 = vmin.value.trim();
              if (op.value === 'between' || op.value === 'notbetween') {
                if (vmax.value.trim() === '') { XL.ui.toast('Informe o Máximo.'); return false; }
                rule.v2 = vmax.value.trim();
              }
            }
            app.addDV(rule);
          },
        },
        { label: 'Limpar Tudo', onClick: () => { app.grid.setRange(g); app.clearDV(); } },
        { label: 'Cancelar' },
      ],
      onClose: () => app.grid.focus(),
    });
  };
})();
