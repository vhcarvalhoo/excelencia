/* Nível Avançado · Módulo 8 — Tabelas Dinâmicas (4 lições + demonstrações). */
(function () {
  'use strict';
  const XL = window.XL;
  const moeda = { numFmt: 'moeda' };
  const head = { b: true };
  const { cell, drag, type, key, tab, btn, menu, item, el, set, typeIn, look, sheet, dnd } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const ok = (say) => el('.ui-modal .btn-primary', say || 'Clique em <b>OK</b>');
  const norm = (s) => XL.normName(String(s));

  // ---------------- Dados de vendas (planilha "Vendas", tabela "Vendas") ----------------
  const PRECO = { Notebook: [3500, 2600], Monitor: [1200, 850], Teclado: [150, 90], Mouse: [80, 45] };
  const LINHAS = [
    ['05/01/2026', 'Ana', 'Sul', 'Notebook', 2], ['09/01/2026', 'Bruno', 'Norte', 'Mouse', 10], ['14/01/2026', 'Carla', 'Sudeste', 'Monitor', 3],
    ['22/01/2026', 'Diego', 'Sul', 'Teclado', 6], ['03/02/2026', 'Ana', 'Sudeste', 'Notebook', 1], ['11/02/2026', 'Bruno', 'Sul', 'Monitor', 2],
    ['17/02/2026', 'Carla', 'Norte', 'Teclado', 8], ['25/02/2026', 'Diego', 'Sudeste', 'Mouse', 12], ['04/03/2026', 'Ana', 'Norte', 'Monitor', 4],
    ['12/03/2026', 'Bruno', 'Sudeste', 'Notebook', 3], ['19/03/2026', 'Carla', 'Sul', 'Mouse', 15], ['27/03/2026', 'Diego', 'Norte', 'Notebook', 2],
    ['02/04/2026', 'Ana', 'Sul', 'Teclado', 5], ['10/04/2026', 'Bruno', 'Norte', 'Notebook', 1], ['16/04/2026', 'Carla', 'Sudeste', 'Notebook', 4],
    ['28/04/2026', 'Diego', 'Sul', 'Monitor', 3], ['06/05/2026', 'Ana', 'Sudeste', 'Mouse', 20], ['13/05/2026', 'Bruno', 'Sul', 'Teclado', 7],
    ['21/05/2026', 'Carla', 'Norte', 'Monitor', 2], ['29/05/2026', 'Diego', 'Sudeste', 'Notebook', 3], ['03/06/2026', 'Ana', 'Norte', 'Notebook', 2],
    ['11/06/2026', 'Bruno', 'Sudeste', 'Monitor', 5], ['18/06/2026', 'Carla', 'Sul', 'Notebook', 3], ['26/06/2026', 'Diego', 'Norte', 'Mouse', 9],
  ];
  const REG = LINHAS.map(([d, v, r, p, q]) => ({ d, v, r, p, q, valor: q * PRECO[p][0], custo: q * PRECO[p][1], mes: +d.slice(3, 5) }));
  const soma = (pred, k = 'valor') => REG.filter(pred).reduce((s, x) => s + x[k], 0);
  const TOTAL = soma(() => true);
  const N = REG.length;

  // escreve a planilha Vendas como Tabela "Vendas"
  const buildVendas = (x) => {
    const app = x.app;
    const S = app.wb.active;
    S.name = 'Vendas';
    const m = { A1: { v: 'Data', fmt: head }, B1: { v: 'Vendedor', fmt: head }, C1: { v: 'Região', fmt: head }, D1: { v: 'Produto', fmt: head }, E1: { v: 'Qtd', fmt: head }, F1: { v: 'Valor', fmt: head }, G1: { v: 'Custo', fmt: head } };
    REG.forEach((x2, i) => {
      const r = i + 2;
      m['A' + r] = { v: x2.d, fmt: { numFmt: 'data' } }; m['B' + r] = x2.v; m['C' + r] = x2.r; m['D' + r] = x2.p; m['E' + r] = x2.q;
      m['F' + r] = { v: x2.valor, fmt: moeda }; m['G' + r] = { v: x2.custo, fmt: moeda };
    });
    x.set(m);
    x.widths({ A: 90, B: 80, C: 80, D: 85, E: 50, F: 100, G: 100 });
    const g = { r1: 0, c1: 0, r2: N, c2: 6 };
    S.tables.push({ name: 'Vendas', g, style: 'azul', banded: true, totalRow: false });
    S.filter = { g: Object.assign({}, g), crit: {}, table: 'Vendas' };
    S.applyFilter();
    app.renderTabs();
    app.render();
  };
  // cria a planilha "Resumo" com uma Tabela Dinâmica pronta
  const buildPivot = (x, def) => {
    const app = x.app;
    const wb = app.wb;
    const S = wb.addSheet('Resumo', 0);
    wb.activeIndex = 0;
    const p = Object.assign({
      id: 'pv-licao', name: 'TabelaDinâmica1', src: { table: 'Vendas' }, at: { r: 2, c: 0 }, rows: [], cols: [], vals: [], filters: [],
      groups: {}, calc: [], hidden: {}, sort: {}, style: 'azul', subtotals: true, grandRows: true, grandCols: true,
    }, def);
    S.pivots.push(p);
    S.renderPivot(p);
    app.afterWorkbookChange();
    app.grid.setSel(2, 0);
    app.render();
    app.ribbon.updateContext();
    return p;
  };

  // ---------------- Auxiliares das verificações ----------------
  const allPivots = (x) => x.app.wb.sheets.flatMap((s) => (s.pivots || []).map((p) => ({ s, p })));
  const pv = (x) => { const a = allPivots(x); return a.length ? a[0].p : null; };
  const pvSheet = (x) => { const a = allPivots(x); return a.length ? a[0].s : null; };
  const has = (list, name) => (list || []).some((f) => norm(f) === norm(name));
  // valor de uma linha da Tabela Dinâmica (pelo rótulo) na coluna do campo de valor i
  const pvVal = (x, label, i = 0) => {
    const p = pv(x), S = pvSheet(x);
    if (!p || !p.out) return null;
    const c = Object.keys(p.vcols || {}).map(Number).find((cc) => p.vcols[cc] === i);
    if (c === undefined) return null;
    for (let r = p.dataTop || p.out.r1; r <= p.out.r2; r++) if (norm(S.textOf(r, p.at.c)) === norm(label)) return S.value(r, c);
    return null;
  };
  const near = (a, b) => typeof a === 'number' && Math.abs(a - b) < 1e-6;
  const goPivot = (x) => { const s = pvSheet(x); if (s && x.app.sheet !== s) { x.app.wb.activeIndex = x.app.wb.sheets.indexOf(s); x.app.afterWorkbookChange(); } const p = pv(x); x.app.grid.setSel(p.dataTop || p.at.r, p.at.c); x.app.ribbon.updateContext(); return p; };

  // ======================================================================
  // 8.1 Criando sua primeira Tabela Dinâmica
  // ======================================================================
  const REGIOES = ['Norte', 'Sudeste', 'Sul'];
  const maiorRegiao = REGIOES.slice().sort((a, b) => soma((y) => y.r === b) - soma((y) => y.r === a))[0];
  XL.registerLesson({
    id: 'a8-1',
    title: 'Criando sua primeira Tabela Dinâmica',
    summary: [
      '<b>Inserir › Tabela Dinâmica</b> resume milhares de linhas sem nenhuma fórmula. Use uma <b>Tabela</b> como fonte: ela cresce e a Tabela Dinâmica acompanha.',
      'Na <b>Lista de Campos</b>: <b>Linhas</b> = itens um embaixo do outro; <b>Colunas</b> = itens lado a lado; <b>Valores</b> = o que somar/contar; <b>Filtros</b> = filtro do relatório inteiro.',
      'Texto marcado vai para Linhas; número vai para Valores (Soma). Arraste para trocar de área.',
      '<b>Configurações do Campo de Valor</b>: Soma, Contagem, Média... e o formato do número.',
      'A setinha de <b>Rótulos de Linha</b> filtra e classifica — inclusive pelo valor (do maior para o menor).',
    ],
    setup: (x) => { buildVendas(x); x.app.grid.setSel(0, 0); },
    steps: [
      {
        title: 'Uma pergunta, 24 linhas',
        text: `<p>O gerente quer saber: <b>quanto cada região vendeu de cada produto?</b></p>
          <p>Dá para responder com SOMASES, mas seriam várias fórmulas. A <b>Tabela Dinâmica</b> monta esse resumo em segundos, só arrastando campos — e muda de formato quando você quiser.</p>
          <p>Repare que os dados já estão numa <b>Tabela</b> chamada <code>Vendas</code>: a melhor fonte para uma Tabela Dinâmica.</p>`,
        target: 'cell:A1:G25',
      },
      {
        title: 'Crie a Tabela Dinâmica',
        task: 'Clique em qualquer célula dos dados (ex.: <b>B3</b>), vá em <b>Inserir › Tabela Dinâmica</b> e confirme com <b>OK</b> (Nova Planilha).',
        hint: 'O botão "Tabela Dinâmica" é o primeiro da guia Inserir.',
        target: ['cell:A1:G25', '.rb-btn[data-cmd="pivot-create"]'],
        check: (x) => { const p = pv(x); return !!p && (p.src.table === 'Vendas' || !!p.src.range); },
        solve: (x) => { x.app.select('B3'); x.app.createPivot('Vendas', null); },
        success: 'Surgiu uma planilha nova com a Tabela Dinâmica (ainda vazia) e, à direita, a <b>Lista de Campos</b>. Lá no alto apareceram as guias <b>Análise da Tabela Dinâmica</b> e <b>Design</b>.',
      },
      {
        title: 'Regiões nas linhas',
        text: '<p>Na <b>Lista de Campos</b>, cada coluna da fonte virou um <b>campo</b>. Ao marcar um campo de <b>texto</b>, o Excel o coloca em <b>Linhas</b>.</p>',
        task: 'Na Lista de Campos, marque <b>Região</b>.',
        hint: 'Se a lista sumiu, clique numa célula da Tabela Dinâmica (ex.: A3).',
        target: '.pv-field[data-field="Região"]',
        check: (x) => { const p = pv(x); return !!p && has(p.rows, 'Região'); },
        solve: (x) => { goPivot(x); x.app.pivotMoveField('Região', 'rows'); },
      },
      {
        title: 'Some os valores',
        text: '<p>Campos de <b>número</b> vão para <b>Valores</b>, e o Excel já começa somando.</p>',
        task: 'Marque o campo <b>Valor</b>.',
        target: '.pv-field[data-field="Valor"]',
        check: (x) => { const p = pv(x); return !!p && p.vals.some((v) => norm(v.field) === norm('Valor') && (v.fn || 'sum') === 'sum') && near(pvVal(x, 'Total Geral'), TOTAL); },
        solve: (x) => { goPivot(x); x.app.pivotMoveField('Valor', 'vals'); },
        success: `Pronto: <b>Soma de Valor</b> por região, com o <b>Total Geral</b> (${XL.formatValue(TOTAL, moeda).text}) no fim. Nenhuma fórmula digitada!`,
      },
      {
        title: 'Produtos nas colunas',
        text: '<p>Agora vamos cruzar: regiões nas linhas e <b>produtos nas colunas</b>. Para escolher a área, <b>arraste</b> o campo até ela.</p>',
        task: 'Arraste o campo <b>Produto</b> da lista até a área <b>Colunas</b>.',
        target: ['.pv-field[data-field="Produto"]', '.pv-area[data-area="cols"]'],
        check: (x) => { const p = pv(x); return !!p && has(p.cols, 'Produto') && has(p.rows, 'Região'); },
        solve: (x) => { goPivot(x); x.app.pivotMoveField('Produto', 'cols'); },
        success: 'Uma tabela cruzada completa: cada célula é a soma de uma região com um produto. Experimente depois arrastar Região para Colunas e Produto para Linhas — o resumo gira (por isso o nome "dinâmica").',
      },
      {
        title: 'Formato de moeda',
        text: '<p>Os valores apareceram sem R$. O formato certo fica nas <b>Configurações do Campo de Valor</b> — assim ele continua valendo quando a Tabela Dinâmica mudar.</p>',
        task: 'Clique numa célula de valor (ex.: <b>B5</b>), vá em <b>Análise da Tabela Dinâmica › Configurações do Campo</b>, escolha o formato <b>Moeda</b> e clique em <b>OK</b>.',
        hint: 'Também dá para abrir pela setinha de "Soma de Valor" na área Valores › Configurações do Campo de Valor.',
        target: ['cell:B5', '.rb-btn[data-cmd="pivot-settings"]'],
        check: (x) => { const p = pv(x); return !!p && p.vals[0] && p.vals[0].numFmt === 'moeda'; },
        solve: (x) => { const p = goPivot(x); x.app.pivotUpdate((q) => { q.vals[0].numFmt = 'moeda'; }, 'pivotsettings', {}, p); },
        success: 'Na mesma janela fica o <b>Resumir Valores por</b>: troque Soma por <b>Contagem</b> (quantos pedidos) ou <b>Média</b> (ticket médio) sem mudar mais nada.',
      },
      {
        title: 'Classifique pelo valor',
        text: '<p>Qual região vendeu mais? Em vez de procurar com os olhos, <b>classifique pelo valor</b>.</p>',
        task: 'Clique na setinha de <b>Rótulos de Linha</b> e escolha <b>Classificar do Maior para o Menor</b>.',
        target: '[data-pvbtn="rows"]',
        check: (x) => { const p = pv(x); const s = p && (p.sort || {})['Região']; return !!s && s.by === 'value' && s.dir === 'desc'; },
        solve: (x) => { const p = goPivot(x); x.app.pivotSort(p, 'Região', 'vdesc'); },
        success: `<b>${maiorRegiao}</b> subiu para o topo. A mesma setinha também <b>filtra</b> itens (desmarque o que não quer ver). Parabéns: você montou um relatório de vendas completo sem nenhuma fórmula!`,
      },
    ],
  });

  R('a8-1', {
    'Crie a Tabela Dinâmica': [
      cell('B3', 'Clique em uma célula qualquer dos dados'),
      btn('pivot-create', undefined, 'Na guia <b>Inserir</b>, clique em <b>Tabela Dinâmica</b>'),
      look('input[data-role="pv-src"]', 'A fonte já veio preenchida com a tabela <b>Vendas</b>'),
      look('input[data-role="pv-new"]', '"Nova Planilha": o relatório fica numa aba só dele'),
      ok(),
    ],
    'Regiões nas linhas': [el('.pv-field[data-field="Região"] input', 'Na Lista de Campos, marque <b>Região</b>')],
    'Some os valores': [el('.pv-field[data-field="Valor"] input', 'Marque <b>Valor</b> — ele vai para a área Valores como <b>Soma de Valor</b>')],
    'Produtos nas colunas': [dnd('.pv-field[data-field="Produto"]', '.pv-area[data-area="cols"]', 'Arraste <b>Produto</b> e solte na área <b>Colunas</b>')],
    'Formato de moeda': [
      cell('B5', 'Clique numa célula de valor'),
      btn('pivot-settings', undefined, 'Na guia <b>Análise da Tabela Dinâmica</b>, clique em <b>Configurações do Campo</b>'),
      look('select[data-role="pv-fn"]', 'Aqui você escolheria Soma, Contagem, Média...'),
      set('select[data-role="pv-numfmt"]', 'moeda', 'Em <b>Formato do Número</b>, escolha <b>Moeda</b>'),
      ok(),
    ],
    'Classifique pelo valor': [
      el('[data-pvbtn="rows"]', 'Clique na setinha de <b>Rótulos de Linha</b>'),
      el('[data-pvsort="vdesc"]', 'Escolha <b>Classificar do Maior para o Menor</b>'),
    ],
  });

  // ======================================================================
  // 8.2 Agrupamentos e campos calculados
  // ======================================================================
  const LUCRO = TOTAL - soma(() => true, 'custo');
  const NOVO_F2 = 10500;
  XL.registerLesson({
    id: 'a8-2',
    title: 'Agrupamentos e campos calculados',
    summary: [
      'Clique numa data da Tabela Dinâmica › <b>Agrupar Seleção</b> › Meses (e Anos, se houver mais de um ano). Números também agrupam em faixas (0-99, 100-199...).',
      '<b>Campos, Itens e Conjuntos › Campo Calculado</b> cria um campo novo a partir dos outros: <code>=Valor-Custo</code>.',
      'O mesmo campo pode entrar duas vezes em Valores — por exemplo, a soma e o <b>% do Total Geral</b>.',
      'Mudou a fonte? A Tabela Dinâmica <b>não</b> se atualiza sozinha: clique em <b>Atualizar</b> (ou Dados › Atualizar Tudo).',
    ],
    setup: (x) => { buildVendas(x); buildPivot(x, { rows: ['Data'], vals: [{ field: 'Valor', fn: 'sum', name: 'Soma de Valor', numFmt: 'moeda' }] }); },
    steps: [
      {
        title: 'Datas demais',
        text: `<p>A planilha <b>Resumo</b> tem uma Tabela Dinâmica com a <b>Soma de Valor por Data</b>: são ${N} datas diferentes, uma por linha. Assim é impossível ver a tendência do semestre.</p>
          <p>Vamos <b>agrupar</b> as datas em meses e criar novas medidas.</p>`,
        target: 'cell:A3:B28',
      },
      {
        title: 'Agrupe por mês',
        task: 'Clique em uma data da Tabela Dinâmica (ex.: <b>A5</b>), vá em <b>Análise da Tabela Dinâmica › Agrupar Seleção</b>, deixe marcado <b>Meses</b> e clique em <b>OK</b>.',
        hint: 'Também dá pelo botão direito sobre a data › Agrupar...',
        target: ['cell:A5', '.rb-btn[data-cmd="pivot-group"]'],
        check: (x) => { const p = pv(x); const g = p && p.groups.Data; return !!g && !!g.by && g.by.includes('months') && near(pvVal(x, 'jan'), soma((y) => y.mes === 1)); },
        solve: (x) => { const p = goPivot(x); x.app.pivotUpdate((q) => { q.groups.Data = { by: ['months'] }; }, 'pivotgroup', {}, p); },
        success: 'De 24 linhas para 6 meses: agora dá para ver o semestre. Com dados de vários anos, marque também <b>Anos</b> — o Excel cria o campo "Anos" e separa jan/2025 de jan/2026.',
      },
      {
        title: 'Campo calculado: lucro',
        text: '<p>A fonte tem <b>Valor</b> e <b>Custo</b>, mas não o lucro. Em vez de criar uma coluna nova nos dados, crie um <b>Campo Calculado</b> na própria Tabela Dinâmica.</p>',
        task: 'Em <b>Análise da Tabela Dinâmica › Campos, Itens e Conjuntos › Campo Calculado</b>, dê o nome <code>Lucro</code>, a fórmula <code>=Valor-Custo</code> e clique em <b>OK</b>.',
        target: ['[data-menu="pivot-calcs"]'],
        check: (x) => { const p = pv(x); if (!p) return false; const i = p.vals.findIndex((v) => norm(v.field) === norm('Lucro')); return i >= 0 && near(pvVal(x, 'Total Geral', i), LUCRO); },
        solve: (x) => { const p = goPivot(x); x.app.pivotUpdate((q) => { q.calc.push({ name: 'Lucro', formula: '=Valor-Custo' }); q.vals.push({ field: 'Lucro', fn: 'sum', name: 'Soma de Lucro' }); }, 'pivotcalc', {}, p); },
        success: `Apareceu a coluna <b>Soma de Lucro</b> (${XL.formatValue(LUCRO, moeda).text} no total). O campo calculado também entrou na Lista de Campos e funciona como qualquer outro.`,
      },
      {
        title: 'Valor duas vezes',
        text: '<p>Quanto cada mês representa do semestre? Para mostrar a soma <b>e</b> a porcentagem lado a lado, coloque o campo <b>Valor</b> de novo em Valores.</p>',
        task: 'Arraste o campo <b>Valor</b> da lista até a área <b>Valores</b> (ele já está lá — tudo bem, entra de novo).',
        target: ['.pv-field[data-field="Valor"]', '.pv-area[data-area="vals"]'],
        check: (x) => { const p = pv(x); return !!p && p.vals.filter((v) => norm(v.field) === norm('Valor')).length >= 2; },
        solve: (x) => { goPivot(x); x.app.pivotMoveField('Valor', 'vals'); },
        success: 'Entrou como <b>Soma de Valor2</b>. Agora vamos transformá-lo em porcentagem.',
      },
      {
        title: 'Mostrar como % do total',
        task: 'Na área Valores, clique na setinha de <b>Soma de Valor2</b> › <b>Configurações do Campo de Valor</b>, e em <b>Mostrar Valores como</b> escolha <b>% do Total Geral</b>. Clique em OK.',
        target: '.pv-area[data-area="vals"]',
        check: (x) => { const p = pv(x); return !!p && p.vals.some((v) => norm(v.field) === norm('Valor') && v.show === 'pctTotal'); },
        solve: (x) => { const p = goPivot(x); const i = p.vals.map((v) => norm(v.field)).lastIndexOf(norm('Valor')); x.app.pivotUpdate((q) => { q.vals[i].show = 'pctTotal'; }, 'pivotsettings', {}, p); },
        success: 'Cada mês agora mostra sua fatia do semestre, e o Total Geral dá 100%. Dica: renomeie o campo para algo como "% do semestre" no Nome Personalizado.',
      },
      {
        title: 'A fonte mudou',
        text: '<p>Uma venda foi corrigida: a de <b>05/01</b> (Ana, Notebook) foi de <b>R$ 10.500</b>, não de R$ 7.000.</p>',
        task: 'Vá à planilha <b>Vendas</b> e, em <b>F2</b>, digite <code>10500</code>.',
        target: ['.xl-sheettab[data-sheet="Vendas"]', 'cell:Vendas!F2'],
        check: (x) => x.val('Vendas!F2') === NOVO_F2,
        solve: (x) => { const S = x.app.wb.byName('Vendas'); x.app.commitInput(1, 5, String(NOVO_F2), S); },
        success: 'Agora volte ao Resumo: a Tabela Dinâmica <b>ainda mostra o valor antigo</b>. Ela guarda uma cópia dos dados e só refaz o resumo quando você pede.',
      },
      {
        title: 'Atualize a Tabela Dinâmica',
        task: 'Volte à planilha <b>Resumo</b>, clique numa célula da Tabela Dinâmica e use <b>Análise da Tabela Dinâmica › Atualizar</b>.',
        hint: 'Também funciona: botão direito › Atualizar, ou Dados › Atualizar Tudo.',
        target: ['.xl-sheettab[data-sheet="Resumo"]', '.rb-btn[data-cmd="pivot-refresh"]'],
        check: (x) => x.happened('pivotrefresh') && near(pvVal(x, 'Total Geral'), TOTAL - REG[0].valor + NOVO_F2),
        solve: (x) => { goPivot(x); x.app.pivotRefresh(false); },
        success: 'Janeiro e o Total Geral foram atualizados. Regra de ouro: <b>mudou a fonte, clique em Atualizar</b>. Se a fonte for uma Tabela, até as linhas novas entram sozinhas no resumo.',
      },
    ],
  });

  R('a8-2', {
    'Agrupe por mês': [
      cell('A5', 'Clique em uma das datas da Tabela Dinâmica'),
      btn('pivot-group', undefined, 'Na guia <b>Análise da Tabela Dinâmica</b>, clique em <b>Agrupar Seleção</b>'),
      look('.pv-units', '<b>Meses</b> já vem marcado'),
      ok(),
    ],
    'Campo calculado: lucro': [
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      menu('pivot-calcs', 'Clique em <b>Campos, Itens e Conjuntos</b>'),
      item('pivot-calc', undefined, 'Escolha <b>Campo Calculado...</b>'),
      typeIn('input[data-role="pv-cname"]', 'Lucro', 'No <b>Nome</b>, digite <code>Lucro</code>'),
      typeIn('input[data-role="pv-cformula"]', '=Valor-Custo', 'Na <b>Fórmula</b>, digite <code>=Valor-Custo</code> (os nomes dos campos, como numa conta)'),
      ok(),
    ],
    'Valor duas vezes': [
      cell('A5', 'Clique dentro da Tabela Dinâmica para ver a Lista de Campos'),
      dnd('.pv-field[data-field="Valor"]', '.pv-area[data-area="vals"]', 'Arraste <b>Valor</b> até a área <b>Valores</b>'),
    ],
    'Mostrar como % do total': [
      cell('A5'),
      el('.pv-chip[data-name="Soma de Valor2"] .pv-caret', 'Clique na setinha de <b>Soma de Valor2</b>'),
      el('.ui-menu-item[data-pv="settings"]', 'Escolha <b>Configurações do Campo de Valor...</b>'),
      set('select[data-role="pv-show"]', 'pctTotal', 'Em <b>Mostrar Valores como</b>, escolha <b>% do Total Geral</b>'),
      ok(),
    ],
    'A fonte mudou': [
      sheet('Vendas', 'Clique na aba <b>Vendas</b>'),
      cell('F2'),
      type(String(NOVO_F2), 'Digite <code>10500</code>'),
      enter(),
    ],
    'Atualize a Tabela Dinâmica': [
      sheet('Resumo', 'Volte para a aba <b>Resumo</b> — repare que o valor de janeiro não mudou'),
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      btn('pivot-refresh', undefined, 'Em <b>Análise da Tabela Dinâmica</b>, clique em <b>Atualizar</b>'),
    ],
  });

  // ======================================================================
  // 8.3 Segmentação de Dados e Linha do Tempo
  // ======================================================================
  const slicerOf = (x, field, kind = 'slicer') => { const S = pvSheet(x); return S ? S.slicers.find((s) => s.kind === kind && norm(s.field) === norm(field)) : null; };
  const MAR = XL.dateToSerial(2026, 3, 1), MAI = XL.dateToSerial(2026, 5, 31);
  XL.registerLesson({
    id: 'a8-3',
    title: 'Segmentação de Dados e Linha do Tempo',
    summary: [
      '<b>Inserir Segmentação de Dados</b> cria botões para filtrar a Tabela Dinâmica com um clique — e mostra na tela o que está filtrado.',
      'Para escolher vários itens: botão <b>Seleção Múltipla</b> ou <kbd>Ctrl</kbd>+clique. O funil com ✕ <b>limpa o filtro</b>.',
      'A <b>Linha do Tempo</b> filtra campos de data por período: clique num mês ou arraste de um mês a outro.',
      'Itens sem dados (por causa de outros filtros) aparecem apagados na segmentação.',
    ],
    setup: (x) => { buildVendas(x); buildPivot(x, { rows: ['Vendedor'], vals: [{ field: 'Valor', fn: 'sum', name: 'Soma de Valor', numFmt: 'moeda' }] }); },
    steps: [
      {
        title: 'Filtros que todo mundo entende',
        text: `<p>A Tabela Dinâmica na planilha <b>Resumo</b> mostra as vendas por vendedor. Quem abre o relatório quer perguntar: "e só no Sul?", "e só no 2º trimestre?".</p>
          <p>A <b>Segmentação de Dados</b> e a <b>Linha do Tempo</b> são filtros em forma de botões: rápidos de usar e deixam claro o que está filtrado.</p>`,
        target: 'cell:A3:B8',
      },
      {
        title: 'Insira uma Segmentação de Dados',
        task: 'Com uma célula da Tabela Dinâmica selecionada, clique em <b>Análise da Tabela Dinâmica › Inserir Segmentação de Dados</b>, marque <b>Região</b> e clique em OK.',
        target: '.rb-btn[data-cmd="pivot-slicer"]',
        check: (x) => !!slicerOf(x, 'Região'),
        solve: (x) => { const p = goPivot(x); x.app.insertSlicers(p, ['Região'], 'slicer'); },
        success: 'Apareceu uma caixa com um botão para cada região. Você pode arrastá-la pelo título.',
      },
      {
        title: 'Filtre com um clique',
        task: 'Na segmentação, clique em <b>Sul</b>.',
        target: '.xl-slicer[data-field="Região"]',
        check: (x) => { const s = slicerOf(x, 'Região'); return !!s && !!s.sel && s.sel.length === 1 && s.sel[0] === 'Sul' && near(pvVal(x, 'Total Geral'), soma((y) => y.r === 'Sul')); },
        solve: (x) => { const s = slicerOf(x, 'Região'); goPivot(x); x.app.slicerPick(s, 'Sul', false); },
        success: 'A Tabela Dinâmica agora mostra só as vendas do Sul — e o botão destacado deixa isso claro para quem lê.',
      },
      {
        title: 'Vários itens',
        task: 'Ligue a <b>Seleção Múltipla</b> (ícone de lista com ✓ no título da segmentação) e clique também em <b>Norte</b>.',
        hint: 'Outra forma: segure <kbd>Ctrl</kbd> enquanto clica em Norte.',
        target: '.xl-slicer[data-field="Região"]',
        check: (x) => { const s = slicerOf(x, 'Região'); return !!s && !!s.sel && s.sel.length === 2 && s.sel.includes('Sul') && s.sel.includes('Norte'); },
        solve: (x) => { const s = slicerOf(x, 'Região'); goPivot(x); x.app.slicerSet(s, (q) => { q.sel = ['Sul', 'Norte']; }, 'slicerpick', {}); },
      },
      {
        title: 'Limpe o filtro',
        task: 'Clique no ícone de <b>funil com ✕</b> (Limpar Filtro) no título da segmentação.',
        target: '.xl-slicer[data-field="Região"] [data-role="sl-clear"]',
        check: (x) => { const s = slicerOf(x, 'Região'); return !!s && !s.sel && x.happened('slicerpick'); },
        solve: (x) => { const s = slicerOf(x, 'Região'); x.app.slicerSet(s, (q) => { q.sel = null; }, 'slicerpick', {}); },
        success: 'Todas as regiões de volta. Para remover a segmentação inteira, clique com o botão direito nela › Remover.',
      },
      {
        title: 'Insira uma Linha do Tempo',
        task: 'Clique numa célula da Tabela Dinâmica e use <b>Análise da Tabela Dinâmica › Inserir Linha do Tempo</b>. Marque <b>Data</b> e clique em OK.',
        target: '.rb-btn[data-cmd="pivot-timeline"]',
        check: (x) => !!slicerOf(x, 'Data', 'timeline'),
        solve: (x) => { const p = goPivot(x); x.app.insertSlicers(p, ['Data'], 'timeline'); },
        success: 'A Linha do Tempo mostra os meses de janeiro a junho. Ela só aceita campos de <b>data</b>.',
      },
      {
        title: 'Escolha um período',
        task: 'Na Linha do Tempo, <b>arraste</b> do mês <b>mar</b> até <b>mai</b>.',
        hint: 'Clique em "mar", segure o botão do mouse e solte em "mai".',
        target: '.xl-slicer.timeline',
        check: (x) => { const s = slicerOf(x, 'Data', 'timeline'); return !!s && !!s.range && s.range.from === MAR && s.range.to === MAI; },
        solve: (x) => { const s = slicerOf(x, 'Data', 'timeline'); x.app.slicerSet(s, (q) => { q.range = { from: MAR, to: MAI }; }, 'timelinepick', {}); },
        success: 'A Tabela Dinâmica mostra só março a maio, e a legenda da Linha do Tempo diz o período. Segmentação e Linha do Tempo se somam: dá para ver "Sul, de março a maio" com dois cliques.',
      },
    ],
  });

  R('a8-3', {
    'Insira uma Segmentação de Dados': [
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      btn('pivot-slicer', undefined, 'Em <b>Análise da Tabela Dinâmica</b>, clique em <b>Inserir Segmentação de Dados</b>'),
      el('.pv-pick input[data-field="Região"]', 'Marque <b>Região</b>'),
      ok(),
    ],
    'Filtre com um clique': [el('.xl-slicer[data-field="Região"] .sl-item[data-item="Sul"]', 'Clique em <b>Sul</b>')],
    'Vários itens': [
      el('.xl-slicer[data-field="Região"] [data-role="sl-multi"]', 'Ligue a <b>Seleção Múltipla</b>'),
      el('.xl-slicer[data-field="Região"] .sl-item[data-item="Norte"]', 'Clique em <b>Norte</b> — agora Sul e Norte ficam selecionados'),
    ],
    'Limpe o filtro': [el('.xl-slicer[data-field="Região"] [data-role="sl-clear"]', 'Clique em <b>Limpar Filtro</b>')],
    'Insira uma Linha do Tempo': [
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      btn('pivot-timeline', undefined, 'Clique em <b>Inserir Linha do Tempo</b>'),
      el('.pv-pick input[data-field="Data"]', 'Marque <b>Data</b>'),
      ok(),
    ],
    'Escolha um período': [
      dnd('.xl-slicer.timeline .tl-month[data-ym="2026-3"]', '.xl-slicer.timeline .tl-month[data-ym="2026-5"]', 'Clique em <b>mar</b> e arraste até <b>mai</b>'),
    ],
  });

  // ======================================================================
  // 8.4 Gráficos dinâmicos
  // ======================================================================
  const pchart = (x) => { const S = pvSheet(x); return S ? S.charts.find((c) => c.pivot) : null; };
  XL.registerLesson({
    id: 'a8-4',
    title: 'Gráficos dinâmicos',
    summary: [
      '<b>Análise da Tabela Dinâmica › Gráfico Dinâmico</b> cria um gráfico ligado à Tabela Dinâmica (ou use Inserir › Gráfico com a célula dentro dela).',
      'Tudo o que você muda na Tabela Dinâmica — campos, filtros, segmentações — muda o gráfico junto.',
      'Linhas → categorias do gráfico; Colunas → séries (uma cor por item).',
      'Linhas para tendência no tempo; colunas para comparar. Troque em Design do Gráfico › Alterar Tipo.',
    ],
    setup: (x) => { buildVendas(x); buildPivot(x, { rows: ['Data'], groups: { Data: { by: ['months'] } }, vals: [{ field: 'Valor', fn: 'sum', name: 'Soma de Valor', numFmt: 'moeda' }] }); },
    steps: [
      {
        title: 'Do resumo ao gráfico',
        text: `<p>A Tabela Dinâmica do <b>Resumo</b> já tem o faturamento por mês. Um <b>Gráfico Dinâmico</b> transforma esse resumo em imagem — e continua ligado a ele.</p>
          <p>A diferença para um gráfico comum: quando você mexe nos campos ou nos filtros, o gráfico <b>se redesenha sozinho</b>.</p>`,
        target: 'cell:A3:B10',
      },
      {
        title: 'Crie o Gráfico Dinâmico',
        task: 'Clique dentro da Tabela Dinâmica e use <b>Análise da Tabela Dinâmica › Gráfico Dinâmico › Linhas</b>.',
        target: '[data-menu="pivot-chart"]',
        check: (x) => { const c = pchart(x); return !!c && c.type === 'line'; },
        solve: (x) => { goPivot(x); x.app.pivotChart('line'); },
        success: 'Uma linha por mês, de janeiro a junho: dá para ver a tendência do semestre de relance.',
      },
      {
        title: 'Uma linha por região',
        task: 'Arraste o campo <b>Região</b> até a área <b>Colunas</b> da Lista de Campos.',
        hint: 'Se a Lista de Campos sumiu, clique numa célula da Tabela Dinâmica (ex.: A5).',
        target: ['.pv-field[data-field="Região"]', '.pv-area[data-area="cols"]'],
        check: (x) => { const p = pv(x); const c = pchart(x); return !!p && !!c && has(p.cols, 'Região') && pvSheet(x).chartData(c).series.length === 3; },
        solve: (x) => { goPivot(x); x.app.pivotMoveField('Região', 'cols'); },
        success: 'A Tabela Dinâmica ganhou uma coluna por região — e o gráfico ganhou <b>uma linha (série) por região</b>, sem você tocar nele.',
      },
      {
        title: 'Troque para colunas',
        text: '<p>Para <b>comparar</b> as regiões mês a mês, colunas lado a lado são mais fáceis de ler do que linhas cruzadas.</p>',
        task: 'Clique no gráfico e use <b>Design do Gráfico › Alterar Tipo de Gráfico › Colunas Agrupadas</b>.',
        target: ['.xl-chart', '[data-menu="chart-type"]'],
        check: (x) => { const c = pchart(x); return !!c && c.type === 'col'; },
        solve: (x) => { const c = pchart(x); x.app.selectChart(c.id); x.app.chartUpdate((q) => { q.type = 'col'; }); },
      },
      {
        title: 'Segmentação para o gráfico',
        task: 'Clique numa célula da Tabela Dinâmica, insira uma <b>Segmentação de Dados</b> de <b>Produto</b> e clique em <b>Notebook</b>.',
        hint: 'Análise da Tabela Dinâmica › Inserir Segmentação de Dados › marque Produto › OK. Depois clique em Notebook.',
        target: '.rb-btn[data-cmd="pivot-slicer"]',
        check: (x) => { const s = slicerOf(x, 'Produto'); const c = pchart(x); if (!s || !c || !s.sel || s.sel.join() !== 'Notebook') return false; const d = pvSheet(x).chartData(c); const tot = d.series.reduce((a, se) => a + se.vals.reduce((b, v) => b + (v || 0), 0), 0); return near(tot, soma((y) => y.p === 'Notebook')); },
        solve: (x) => { const p = goPivot(x); if (!slicerOf(x, 'Produto')) x.app.insertSlicers(p, ['Produto'], 'slicer'); x.app.slicerPick(slicerOf(x, 'Produto'), 'Notebook', false); },
        success: 'O gráfico agora mostra só os notebooks: um <b>painel interativo</b> — clique em outro produto e veja tudo mudar. Esse é o começo de um dashboard!',
      },
    ],
  });

  R('a8-4', {
    'Crie o Gráfico Dinâmico': [
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      menu('pivot-chart', 'Em <b>Análise da Tabela Dinâmica</b>, clique em <b>Gráfico Dinâmico</b>'),
      item('pivot-chart', 'line', 'Escolha <b>Linhas</b>'),
    ],
    'Uma linha por região': [
      cell('A5', 'Clique dentro da Tabela Dinâmica para ver a Lista de Campos'),
      dnd('.pv-field[data-field="Região"]', '.pv-area[data-area="cols"]', 'Arraste <b>Região</b> até a área <b>Colunas</b>'),
      look('.xl-chart', 'O gráfico ganhou uma linha por região'),
    ],
    'Troque para colunas': [
      el('.xl-chart', 'Clique no gráfico para selecioná-lo'),
      menu('chart-type', 'Em <b>Design do Gráfico</b>, clique em <b>Alterar Tipo de Gráfico</b>'),
      item('chart-type', 'col', 'Escolha <b>Colunas Agrupadas</b>'),
    ],
    'Segmentação para o gráfico': [
      cell('A5', 'Clique dentro da Tabela Dinâmica'),
      btn('pivot-slicer', undefined, 'Clique em <b>Inserir Segmentação de Dados</b>'),
      el('.pv-pick input[data-field="Produto"]', 'Marque <b>Produto</b>'),
      ok(),
      el('.xl-slicer[data-field="Produto"] .sl-item[data-item="Notebook"]', 'Clique em <b>Notebook</b> — o gráfico acompanha'),
    ],
  });
})();
