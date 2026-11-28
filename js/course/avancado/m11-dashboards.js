/* Nível Avançado · Módulo 11 — Dashboards (3 lições + demonstrações). */
(function () {
  'use strict';
  const XL = window.XL;
  const head = { b: true, fill: '#DDEBF7' };
  const moeda = { numFmt: 'moeda' };
  const { cell, drag, type, key, tab, btn, menu, item, el, rclick, typeIn, fill, look } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const ok = (say) => el('.ui-modal .btn-primary', say || 'Clique em <b>OK</b>');
  const esc = XL.escapeHtml;
  const write = (ref, formula, say) => [cell(ref), type(formula, say || `Digite <code>${esc(formula)}</code>`), enter()];
  const commit = (x, ref, text) => { const p = XL.parseRef(ref); x.app.select(ref); return x.app.commitInput(p.r, p.c, String(text)); };
  const fillDown = (x, from, to) => { x.app.select(from); x.app.fill(XL.parseRange(from), XL.parseRange(from + ':' + to), 'down'); };
  const near = (a, b, eps = 1e-6) => typeof a === 'number' && Math.abs(a - b) < eps;
  const norm = (s) => XL.normName(String(s));

  // ======================================================================
  // 11.1 Princípios de um bom dashboard
  // ======================================================================
  const PROD = [['Mouse', 5400], ['Notebook', 48000], ['Headset', 3900], ['Monitor', 21000], ['Teclado', 7200], ['Cadeira gamer', 9500]];
  const ORD = PROD.slice().sort((a, b) => b[1] - a[1]);
  const chart11 = (x) => x.sheet.charts.find((c) => XL.rangeToStr(c.g) === 'A1:B7') || null;
  XL.registerLesson({
    id: 'a11-1',
    title: 'Princípios de um bom dashboard',
    summary: [
      'Um dashboard responde a <b>uma pergunta</b> de quem vai ler — não mostra tudo o que você tem.',
      '<b>Ordene</b> os dados (do maior para o menor) e use <b>barras horizontais</b> para comparar categorias com nomes longos.',
      'O <b>título conta a conclusão</b> ("Notebook vende mais que todos os outros juntos"), não só o assunto ("Vendas").',
      '<b>Menos é mais</b>: rótulos nos dados em vez de eixos cheios, sem linhas de grade, poucas cores e com significado.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Produto', fmt: head }, B1: { v: 'Vendas', fmt: head } };
      PROD.forEach(([p, v], i) => { m['A' + (i + 2)] = p; m['B' + (i + 2)] = { v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 120, B: 110 });
    },
    steps: [
      {
        title: 'Cinco princípios',
        text: `<ol style="padding-left:18px;margin:6px 0">
          <li><b>Uma pergunta</b>: "qual produto puxa as vendas?"</li>
          <li><b>Ordem</b>: do mais importante para o menos importante.</li>
          <li><b>O gráfico certo</b>: barras para comparar, linhas para tempo.</li>
          <li><b>Título com a mensagem</b>, não só o assunto.</li>
          <li><b>Menos tinta</b>: tire tudo o que não ajuda a ler.</li></ol>
          <p>Vamos aplicar cada um nestes dados de vendas.</p>`,
        target: 'cell:A1:B7',
      },
      {
        title: 'Ordem: do maior para o menor',
        task: 'Clique numa célula de <b>Vendas</b> (ex.: <b>B3</b>) e use <b>Dados › Classificar de Z a A</b>.',
        hint: 'É o botão "ZA" com a seta para baixo, no grupo Classificar e Filtrar.',
        target: ['cell:B2:B7', '.rb-btn[data-cmd="sort"][data-arg="desc"]'],
        check: (x) => ORD.every(([p], i) => x.val('A' + (i + 2)) === p),
        solve: (x) => { x.app.select('B3'); x.app.exec('sort', 'desc'); },
        success: 'Agora quem lê vê primeiro o que importa. Em lista desordenada, o olho fica pulando para achar o maior.',
      },
      {
        title: 'O gráfico certo',
        text: '<p>Para comparar categorias, <b>barras horizontais</b> são as mais fáceis de ler — e os nomes dos produtos cabem inteiros do lado.</p>',
        task: 'Selecione <b>A1:B7</b> e clique em <b>Inserir › Gráfico de Barras</b>.',
        target: ['cell:A1:B7', '.rb-btn[data-cmd="chart-insert"][data-arg="bar"]'],
        check: (x) => { const c = chart11(x); return !!c && c.type === 'bar'; },
        solve: (x) => { x.app.select('A1:B7'); x.app.insertChart('bar'); },
        success: 'O maior fica no topo, como na tabela. Uma pizza com 6 fatias pareceria parecida e seria bem mais difícil de comparar.',
      },
      {
        title: 'Um título que conta a história',
        text: '<p>"Vendas" diz o assunto. Um bom título diz a <b>conclusão</b>: Notebook sozinho (R$ 48 mil) vende mais que todos os outros juntos (R$ 47 mil).</p>',
        task: 'Com o gráfico selecionado, use <b>Design do Gráfico › Título do Gráfico</b> e escreva <code>Notebook vende mais que todos os outros juntos</code>.',
        target: ['.xl-chart', '.rb-btn[data-cmd="chart-title"]'],
        check: (x) => { const c = chart11(x); return !!c && /notebook/i.test(String(c.title || '')) && String(c.title).length > 20; },
        solve: (x) => { const c = chart11(x); x.app.selectChart(c.id); x.app.chartUpdate((q) => { q.title = 'Notebook vende mais que todos os outros juntos'; }); },
        success: 'Quem bate o olho já entende a mensagem, mesmo sem ler os números.',
      },
      {
        title: 'Valores nas barras',
        task: 'Marque <b>Design do Gráfico › Rótulos de Dados</b>.',
        target: ['.xl-chart', 'input[data-cmd="chart-labels"]'],
        check: (x) => { const c = chart11(x); return !!c && !!c.labels; },
        solve: (x) => { const c = chart11(x); x.app.selectChart(c.id); x.app.chartUpdate((q) => { q.labels = true; }); },
        success: 'Com o valor escrito em cada barra, ninguém precisa seguir as linhas até o eixo.',
      },
      {
        title: 'Menos tinta',
        text: '<p>Num painel, as <b>linhas de grade</b> da planilha só poluem. O Excel deixa escondê-las — só na tela, os dados continuam lá.</p>',
        task: 'Na guia <b>Exibir</b>, desmarque <b>Linhas de Grade</b>.',
        target: 'input[data-cmd="gridlines"]',
        check: (x) => x.app.grid.gridlines === false,
        solve: (x) => x.app.exec('gridlines', false),
        success: 'Visual limpo: dados ordenados, gráfico certo, título com mensagem e sem ruído. Esses princípios valem para qualquer relatório.',
      },
    ],
  });

  R('a11-1', {
    'Ordem: do maior para o menor': [
      cell('B3', 'Clique numa célula de <b>Vendas</b>'),
      btn('sort', 'desc', 'Na guia <b>Dados</b>, clique em <b>Classificar de Z a A</b>'),
    ],
    'O gráfico certo': [
      drag('A1', 'B7', 'Selecione de <b>A1</b> até <b>B7</b>'),
      btn('chart-insert', 'bar', 'Na guia <b>Inserir</b>, clique em <b>Gráfico de Barras</b>'),
    ],
    'Um título que conta a história': [
      el('.xl-chart', 'Clique no gráfico'),
      btn('chart-title', undefined, 'Em <b>Design do Gráfico</b>, clique em <b>Título do Gráfico</b>'),
      typeIn('input[data-role="chart-title"]', 'Notebook vende mais que todos os outros juntos', 'Escreva a conclusão como título'),
      ok(),
    ],
    'Valores nas barras': [
      el('.xl-chart', 'Clique no gráfico'),
      tab('chartdesign', 'Clique na guia <b>Design do Gráfico</b>'),
      el('input[data-cmd="chart-labels"]', 'Marque <b>Rótulos de Dados</b>'),
    ],
    'Menos tinta': [
      tab('view', 'Clique na guia <b>Exibir</b>'),
      el('input[data-cmd="gridlines"]', 'Desmarque <b>Linhas de Grade</b>'),
    ],
  });

  // ======================================================================
  // 11.2 KPIs e minigráficos
  // ======================================================================
  const REG = [['Norte', [42, 45, 44, 48, 51, 53]], ['Nordeste', [38, 36, 37, 35, 34, 33]], ['Sul', [55, 54, 58, 60, 59, 63]], ['Sudeste', [80, 84, 82, 88, 91, 95]]]
    .map(([n, v]) => [n, v.map((x) => x * 1000)]);
  const MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun'];
  const sparkOf = (x) => x.sheet.sparks.find((s) => s.loc === 'H3:H6') || null;
  XL.registerLesson({
    id: 'a11-2',
    title: 'KPIs e minigráficos',
    summary: [
      '<b>KPI</b> (indicador-chave) é um número que responde "estamos bem?" — sempre comparado a algo: meta, mês anterior, ano passado.',
      'Variação percentual: <code>=(novo-antigo)/antigo</code>, com formato %. Formatação condicional destaca o que piorou.',
      '<b>Minigráficos</b> (Inserir › Minigráficos) são gráficos pequenos dentro de uma célula: a tendência de cada linha, lado a lado.',
      'Na guia <b>Minigráfico</b>: tipo (Linha, Coluna, Ganho/Perda), Ponto Alto/Baixo e Marcadores.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Vendas por região — 1º semestre', fmt: { b: true, size: 13 } }, A2: { v: 'Região', fmt: head }, H2: { v: 'Tendência', fmt: head }, I2: { v: 'Total', fmt: head }, J2: { v: 'Var. jun/jan', fmt: head } };
      MES.forEach((mm, j) => { m[XL.numToCol(j + 2) + '2'] = { v: mm, fmt: head }; });
      REG.forEach(([n, vals], i) => {
        const r = i + 3;
        m['A' + r] = n;
        vals.forEach((v, j) => { m[XL.numToCol(j + 2) + r] = { v, fmt: { numFmt: 'numero', dec: 0 } }; });
        m['I' + r] = { v: null, fmt: { numFmt: 'moeda', b: true } };
        m['J' + r] = { v: null, fmt: { numFmt: 'porcentagem', dec: 1, b: true } };
      });
      x.set(m);
      x.widths({ A: 90, B: 62, C: 62, D: 62, E: 62, F: 62, G: 62, H: 110, I: 115, J: 95 });
    },
    steps: [
      {
        title: 'Números que respondem "estamos bem?"',
        text: `<p>A tabela tem as vendas mensais de 4 regiões. Num dashboard, ninguém quer ler 24 números: quer <b>KPIs</b> (o total, a variação) e a <b>tendência</b> de cada região.</p>
          <p>Vamos montar as colunas Total e Variação e, na coluna <b>Tendência</b>, colocar <b>minigráficos</b> — gráficos do tamanho de uma célula.</p>`,
        target: 'cell:A2:J6',
      },
      {
        title: 'KPI: o total',
        task: 'Em <b>I3</b>, digite <code>=SOMA(B3:G3)</code> e arraste até <b>I6</b>.',
        target: 'cell:I3:I6',
        check: (x) => REG.every(([, v], i) => near(x.val('I' + (i + 3)), v.reduce((s, y) => s + y, 0))),
        solve: (x) => { commit(x, 'I3', '=SOMA(B3:G3)'); fillDown(x, 'I3', 'I6'); },
      },
      {
        title: 'KPI: a variação',
        text: '<p>Um número sozinho não diz se é bom. Compare junho com janeiro: <code>=(G3-B3)/B3</code> dá a variação em %.</p>',
        task: 'Em <b>J3</b>, digite <code>=(G3-B3)/B3</code> e arraste até <b>J6</b>.',
        target: 'cell:J3:J6',
        check: (x) => REG.every(([, v], i) => near(x.val('J' + (i + 3)), (v[5] - v[0]) / v[0])),
        solve: (x) => { commit(x, 'J3', '=(G3-B3)/B3'); fillDown(x, 'J3', 'J6'); },
        success: 'O Sudeste cresceu 18,8%; o Nordeste <b>caiu 13,2%</b>. Esse é o tipo de coisa que o dashboard precisa gritar.',
      },
      {
        title: 'Destaque o que piorou',
        task: 'Selecione <b>J3:J6</b> e use <b>Formatação Condicional › É Menor do que...</b> com o valor <code>0</code> (vermelho).',
        target: ['cell:J3:J6', '[data-menu="cf"]'],
        check: (x) => x.sheet.cf.some((r) => r.kind === 'lt' && XL.rangeToStr(r.g) === 'J3:J6' && x.sheet.cfValue(r.v1) === 0),
        solve: (x) => { x.app.select('J3:J6'); x.app.addCF({ kind: 'lt', v1: '0', style: 'vermelho' }); },
        success: 'Vermelho só onde há problema: a cor ganha significado em vez de enfeitar.',
      },
      {
        title: 'Minigráficos de linha',
        task: 'Selecione <b>H3:H6</b>, clique em <b>Inserir › Minigráficos › Linha</b>, em <b>Intervalo de dados</b> digite <code>B3:G6</code> e clique em OK.',
        target: ['cell:H3:H6', '.rb-btn[data-cmd="spark-insert"][data-arg="line"]'],
        check: (x) => { const s = sparkOf(x); return !!s && s.data === 'B3:G6' && s.type === 'line'; },
        solve: (x) => { x.app.pushUndo(); x.sheet.sparks.push({ id: 'sp-licao', type: 'line', data: 'B3:G6', loc: 'H3:H6', high: false, low: false, markers: false }); x.app.select('H3'); x.app.render(); x.app.ribbon.updateContext(); x.app.emit('spark', {}); },
        success: 'Cada região ganhou um mini-gráfico com a sua tendência: dá para ver de relance que o Nordeste é o único que desce.',
      },
      {
        title: 'Ponto alto e ponto baixo',
        task: 'Com uma célula de H3:H6 selecionada, marque <b>Ponto Alto</b> e <b>Ponto Baixo</b> na guia <b>Minigráfico</b>.',
        target: ['cell:H3:H6', 'input[data-cmd="spark-high"]'],
        check: (x) => { const s = sparkOf(x); return !!s && s.high && s.low; },
        solve: (x) => { x.app.select('H3'); x.app.sparkUpdate((g) => { g.high = true; g.low = true; }); },
        success: 'O melhor mês aparece em verde e o pior em laranja, em cada região.',
      },
      {
        title: 'Troque para colunas',
        task: 'Na guia <b>Minigráfico</b>, clique em <b>Coluna</b>.',
        target: '.rb-btn[data-cmd="spark-type"][data-arg="col"]',
        check: (x) => { const s = sparkOf(x); return !!s && s.type === 'col'; },
        solve: (x) => { x.app.select('H3'); x.app.sparkUpdate((g) => { g.type = 'col'; }); },
        success: 'Colunas mostram melhor o tamanho de cada mês; linhas, a tendência. Já o <b>Ganho/Perda</b> serve para dados com positivos e negativos (como lucro ou variação).',
      },
    ],
  });

  R('a11-2', {
    'KPI: o total': write('I3', '=SOMA(B3:G3)').concat([fill('I3', 'I6', 'Arraste a alça de preenchimento até <b>I6</b>')]),
    'KPI: a variação': write('J3', '=(G3-B3)/B3', 'Digite <code>=(G3-B3)/B3</code> — (novo − antigo) ÷ antigo').concat([fill('J3', 'J6', 'Arraste a alça até <b>J6</b>')]),
    'Destaque o que piorou': [
      drag('J3', 'J6', 'Selecione as variações: <b>J3:J6</b>'),
      menu('cf', 'Na guia <b>Página Inicial</b>, clique em <b>Formatação Condicional</b>'),
      item('cf', 'lt', 'Escolha <b>É Menor do que...</b>'),
      typeIn('input[data-role="cf-v1"]', '0', 'Digite <code>0</code>'),
      ok(),
    ],
    'Minigráficos de linha': [
      drag('H3', 'H6', 'Selecione onde os minigráficos vão ficar: <b>H3:H6</b>'),
      btn('spark-insert', 'line', 'Na guia <b>Inserir</b>, grupo Minigráficos, clique em <b>Linha</b>'),
      typeIn('input[data-role="sp-data"]', 'B3:G6', 'Em <b>Intervalo de dados</b>, digite <code>B3:G6</code> (os 6 meses das 4 regiões)'),
      look('input[data-role="sp-loc"]', 'O local já veio da seleção: <b>$H$3:$H$6</b>'),
      ok(),
    ],
    'Ponto alto e ponto baixo': [
      cell('H3', 'Clique num minigráfico'),
      tab('sparkdesign', 'Clique na guia <b>Minigráfico</b>'),
      el('input[data-cmd="spark-high"]', 'Marque <b>Ponto Alto</b>'),
      el('input[data-cmd="spark-low"]', 'Marque <b>Ponto Baixo</b>'),
    ],
    'Troque para colunas': [
      cell('H3', 'Clique num minigráfico'),
      btn('spark-type', 'col', 'Na guia <b>Minigráfico</b>, clique em <b>Coluna</b>'),
    ],
  });

  // ======================================================================
  // 11.3 Dashboard interativo completo
  // ======================================================================
  const PRECO = { Notebook: 3500, Monitor: 1200, Teclado: 150, Mouse: 80 };
  const LINHAS = [
    ['05/01/2026', 'Sul', 'Notebook', 2], ['09/01/2026', 'Norte', 'Mouse', 10], ['14/01/2026', 'Sudeste', 'Monitor', 3], ['22/01/2026', 'Sul', 'Teclado', 6],
    ['03/02/2026', 'Sudeste', 'Notebook', 1], ['11/02/2026', 'Sul', 'Monitor', 2], ['17/02/2026', 'Norte', 'Teclado', 8], ['25/02/2026', 'Sudeste', 'Mouse', 12],
    ['04/03/2026', 'Norte', 'Monitor', 4], ['12/03/2026', 'Sudeste', 'Notebook', 3], ['19/03/2026', 'Sul', 'Mouse', 15], ['27/03/2026', 'Norte', 'Notebook', 2],
    ['02/04/2026', 'Sul', 'Teclado', 5], ['10/04/2026', 'Norte', 'Notebook', 1], ['16/04/2026', 'Sudeste', 'Notebook', 4], ['28/04/2026', 'Sul', 'Monitor', 3],
    ['06/05/2026', 'Sudeste', 'Mouse', 20], ['13/05/2026', 'Sul', 'Teclado', 7], ['21/05/2026', 'Norte', 'Monitor', 2], ['29/05/2026', 'Sudeste', 'Notebook', 3],
    ['03/06/2026', 'Norte', 'Notebook', 2], ['11/06/2026', 'Sudeste', 'Monitor', 5], ['18/06/2026', 'Sul', 'Notebook', 3], ['26/06/2026', 'Norte', 'Mouse', 9],
  ].map(([d, r, p, q]) => ({ d, r, p, q, v: q * PRECO[p] }));
  const somaNotebook = LINHAS.filter((l) => l.p === 'Notebook').reduce((s, l) => s + l.v, 0);
  const pivotBase = (id, name, at, def) => Object.assign({
    id, name, src: { table: 'Vendas' }, at, rows: [], cols: [], vals: [{ field: 'Valor', fn: 'sum', name: 'Soma de Valor', numFmt: 'moeda' }], filters: [],
    groups: {}, calc: [], hidden: {}, sort: {}, style: 'azul', subtotals: true, grandRows: true, grandCols: true,
  }, def);
  const painel = (x) => x.app.wb.byName('Painel');
  const pv = (x, i) => { const S = painel(x); return S ? S.pivots[i] : null; };
  const pchart = (x, i) => { const S = painel(x), p = pv(x, i); return S && p ? S.charts.find((c) => c.pivot === p.id) : null; };
  const slicer = (x) => { const S = painel(x); return S ? S.slicers.find((s) => s.kind === 'slicer' && norm(s.field) === norm('Produto')) : null; };
  const total = (x, i) => { const S = painel(x), p = pv(x, i); if (!S || !p || !p.out) return null; const c = Object.keys(p.vcols || {}).map(Number)[0]; return S.value(p.out.r2, c); };
  const goPivot = (x, i) => { const S = painel(x); if (x.app.sheet !== S) { x.app.wb.activeIndex = x.app.wb.sheets.indexOf(S); x.app.afterWorkbookChange(); } const p = pv(x, i); x.app.grid.setSel(p.dataTop, p.at.c); x.app.ribbon.updateContext(); return p; };
  XL.registerLesson({
    id: 'a11-3',
    title: 'Dashboard interativo completo',
    summary: [
      'Dashboard interativo = <b>Tabelas Dinâmicas</b> (os cálculos) + <b>Gráficos Dinâmicos</b> (a visão) + <b>Segmentações</b> (os filtros).',
      'Uma segmentação pode filtrar <b>várias</b> Tabelas Dinâmicas da mesma fonte: botão direito › <b>Conexões de Relatório</b>.',
      'Organize: filtros à esquerda ou no topo, gráficos lado a lado, título claro, sem linhas de grade.',
      'Com os dados numa <b>Tabela</b>, basta colar as vendas novas e clicar em <b>Dados › Atualizar Tudo</b>.',
    ],
    setup: (x) => {
      const app = x.app;
      const V = app.wb.active;
      V.name = 'Vendas';
      const m = { A1: { v: 'Data', fmt: { b: true } }, B1: { v: 'Região', fmt: { b: true } }, C1: { v: 'Produto', fmt: { b: true } }, D1: { v: 'Qtd', fmt: { b: true } }, E1: { v: 'Valor', fmt: { b: true } } };
      LINHAS.forEach((l, i) => { const r = i + 2; m['A' + r] = { v: l.d, fmt: { numFmt: 'data' } }; m['B' + r] = l.r; m['C' + r] = l.p; m['D' + r] = l.q; m['E' + r] = { v: l.v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 90, B: 80, C: 85, D: 50, E: 100 });
      const g = { r1: 0, c1: 0, r2: LINHAS.length, c2: 4 };
      V.tables.push({ name: 'Vendas', g, style: 'azul', banded: true, totalRow: false });
      const S = app.wb.addSheet('Painel', 0);
      app.wb.activeIndex = 0;
      S.setValue(0, 0, 'Painel de Vendas — 1º semestre de 2026');
      S.setFmt(0, 0, { b: true, size: 14, color: '#1F4E79' });
      const p1 = pivotBase('pv-reg', 'TabelaDinâmica1', { r: 2, c: 0 }, { rows: ['Região'] });
      const p2 = pivotBase('pv-mes', 'TabelaDinâmica2', { r: 9, c: 0 }, { rows: ['Data'], groups: { Data: { by: ['months'] } } });
      S.pivots.push(p1, p2);
      S.renderPivot(p1); S.renderPivot(p2);
      app.afterWorkbookChange();
      app.grid.setSel(4, 0);
      app.render();
      app.ribbon.updateContext();
    },
    steps: [
      {
        title: 'As peças de um painel',
        text: `<p>A planilha <b>Painel</b> já tem duas Tabelas Dinâmicas da tabela <b>Vendas</b>: vendas por <b>região</b> (A3) e por <b>mês</b> (A10).</p>
          <p>Vamos transformá-las num painel interativo: um gráfico para cada uma e uma <b>segmentação de Produto</b> que filtra <b>tudo de uma vez</b>.</p>`,
        target: ['cell:A3:B7', 'cell:A10:B17'],
      },
      {
        title: 'Gráfico das regiões',
        task: 'Clique em <b>A5</b> (Tabela Dinâmica das regiões) e use <b>Análise da Tabela Dinâmica › Gráfico Dinâmico › Colunas Agrupadas</b>.',
        target: ['cell:A3:B7', '[data-menu="pivot-chart"]'],
        check: (x) => { const c = pchart(x, 0); return !!c && c.type === 'col'; },
        solve: (x) => { goPivot(x, 0); x.app.pivotChart('col'); },
      },
      {
        title: 'Gráfico dos meses',
        task: 'Clique em <b>A12</b> (Tabela Dinâmica dos meses) e crie um <b>Gráfico Dinâmico › Linhas</b>.',
        target: ['cell:A10:B17', '[data-menu="pivot-chart"]'],
        check: (x) => { const c = pchart(x, 1); return !!c && c.type === 'line'; },
        solve: (x) => { goPivot(x, 1); x.app.pivotChart('line'); },
        success: 'Dois gráficos, cada um ligado à sua Tabela Dinâmica: colunas para comparar regiões, linha para ver o tempo.',
      },
      {
        title: 'Uma segmentação de Produto',
        task: 'Clique em <b>A5</b> e use <b>Análise da Tabela Dinâmica › Inserir Segmentação de Dados</b>, marcando <b>Produto</b>.',
        target: '.rb-btn[data-cmd="pivot-slicer"]',
        check: (x) => !!slicer(x),
        solve: (x) => { const p = goPivot(x, 0); x.app.insertSlicers(p, ['Produto'], 'slicer'); },
        success: 'A segmentação ficou embaixo das tabelas, à esquerda dos gráficos. Por enquanto ela só filtra a Tabela Dinâmica das regiões.',
      },
      {
        title: 'Conecte as duas Tabelas Dinâmicas',
        text: '<p>Para um único filtro controlar o painel inteiro, a segmentação precisa estar ligada às <b>duas</b> Tabelas Dinâmicas.</p>',
        task: 'Clique com o <b>botão direito</b> no título da segmentação › <b>Conexões de Relatório...</b>, marque também <b>TabelaDinâmica2</b> e clique em OK.',
        target: '.xl-slicer[data-field="Produto"]',
        check: (x) => { const s = slicer(x); return !!s && s.pivots.length === 2; },
        solve: (x) => { const s = slicer(x); const S = painel(x); x.app.pushUndo(); s.pivots = S.pivots.map((p) => p.id); S.pivots.forEach((p) => S.renderPivot(p)); x.app.render(); x.app.emit('slicerconn', {}); },
      },
      {
        title: 'Filtre o painel inteiro',
        task: 'Na segmentação, clique em <b>Notebook</b>.',
        target: '.xl-slicer[data-field="Produto"]',
        check: (x) => { const s = slicer(x); return !!s && s.sel && s.sel.join() === 'Notebook' && near(total(x, 0), somaNotebook) && near(total(x, 1), somaNotebook); },
        solve: (x) => x.app.slicerPick(slicer(x), 'Notebook', false),
        success: `As duas tabelas e os dois gráficos mostram só notebooks (${XL.formatValue(somaNotebook, moeda).text}) — com <b>um</b> clique. Experimente outros produtos.`,
      },
      {
        title: 'Acabamento',
        task: 'Na guia <b>Exibir</b>, desmarque <b>Linhas de Grade</b> para dar cara de painel.',
        target: 'input[data-cmd="gridlines"]',
        check: (x) => x.app.grid.gridlines === false,
        solve: (x) => x.app.exec('gridlines', false),
        success: 'Seu primeiro dashboard interativo está pronto! Para atualizar com vendas novas: cole-as na tabela Vendas e clique em <b>Dados › Atualizar Tudo</b>. Parabéns por concluir o Nível Avançado!',
      },
    ],
  });

  R('a11-3', {
    'Gráfico das regiões': [
      cell('A5', 'Clique na Tabela Dinâmica das regiões'),
      menu('pivot-chart', 'Em <b>Análise da Tabela Dinâmica</b>, clique em <b>Gráfico Dinâmico</b>'),
      item('pivot-chart', 'col', 'Escolha <b>Colunas Agrupadas</b>'),
    ],
    'Gráfico dos meses': [
      cell('A12', 'Clique na Tabela Dinâmica dos meses'),
      menu('pivot-chart', 'Clique em <b>Gráfico Dinâmico</b>'),
      item('pivot-chart', 'line', 'Escolha <b>Linhas</b>'),
    ],
    'Uma segmentação de Produto': [
      cell('A5', 'Clique na Tabela Dinâmica das regiões'),
      btn('pivot-slicer', undefined, 'Clique em <b>Inserir Segmentação de Dados</b>'),
      el('.pv-pick input[data-field="Produto"]', 'Marque <b>Produto</b>'),
      ok(),
    ],
    'Conecte as duas Tabelas Dinâmicas': [
      rclick('.xl-slicer[data-field="Produto"] .sl-title', 'Clique com o <b>botão direito</b> no título da segmentação'),
      el('.ui-menu-item[data-pv="sl-conn"]', 'Escolha <b>Conexões de Relatório...</b>'),
      el('input[data-pivot="TabelaDinâmica2"]', 'Marque <b>TabelaDinâmica2</b> (a dos meses)'),
      ok(),
    ],
    'Filtre o painel inteiro': [
      el('.xl-slicer[data-field="Produto"] .sl-item[data-item="Notebook"]', 'Clique em <b>Notebook</b>'),
      look('.xl-chart', 'Os gráficos acompanham o filtro'),
    ],
    'Acabamento': [
      tab('view', 'Clique na guia <b>Exibir</b>'),
      el('input[data-cmd="gridlines"]', 'Desmarque <b>Linhas de Grade</b>'),
    ],
  });
})();
