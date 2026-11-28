/* Nível Avançado · Módulo 10 — Análise e cenários: Atingir Meta, Cenários e Tabela de Dados, Solver e previsões. */
(function () {
  'use strict';
  const XL = window.XL;
  const head = { b: true, fill: '#DDEBF7' };
  const moeda = { numFmt: 'moeda' };
  const destaque = { b: true, fill: '#FFF2CC', numFmt: 'moeda' };
  const { cell, drag, type, key, tab, btn, menu, item, el, set, typeIn, look } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const ok = (say) => el('.ui-modal .btn-primary', say || 'Clique em <b>OK</b>');
  const esc = XL.escapeHtml;
  const write = (ref, formula, say) => [cell(ref), type(formula, say || `Digite <code>${esc(formula)}</code>`), enter()];
  const commit = (x, ref, text) => { const p = XL.parseRef(ref); x.app.select(ref); return x.app.commitInput(p.r, p.c, String(text)); };
  const near = (a, b, eps = 1e-6) => typeof a === 'number' && Math.abs(a - b) < eps;

  // ======================================================================
  // 10.1 Atingir Meta
  // ======================================================================
  const pgto = (pv, r, n) => (pv * r) / (1 - Math.pow(1 + r, -n));
  const seekGoal = (x, setRef, target, changeRef) => {
    const S = x.sheet;
    const res = S.goalSeek(setRef, target, changeRef);
    const p = XL.parseRef(changeRef);
    x.app.pushUndo();
    S.setValue(p.r, p.c, res.x);
    x.app.render();
    x.app.emit('goalseek', { set: setRef, change: changeRef, target, ok: res.ok, x: res.x });
  };
  XL.registerLesson({
    id: 'a10-1',
    title: 'Atingir Meta',
    summary: [
      'Fórmulas vão das entradas para o resultado. <b>Atingir Meta</b> faz o caminho inverso: você diz o resultado e o Excel descobre a entrada.',
      'Dados › Teste de Hipóteses › Atingir Meta: <b>Definir célula</b> (a fórmula), <b>Para valor</b> (a meta) e <b>Alternando célula</b> (o número que pode mudar).',
      'A célula a alternar precisa ser um <b>valor</b>, e a célula definida precisa ser uma <b>fórmula</b> que dependa dela.',
      'Clássicos: valor máximo do empréstimo, prazo necessário e <b>ponto de equilíbrio</b> (lucro = 0).',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Valor do empréstimo', fmt: { b: true } }, B1: { v: 30000, fmt: moeda },
        A2: { v: 'Juros ao mês', fmt: { b: true } }, B2: { v: 0.015, fmt: { numFmt: 'porcentagem', dec: 1 } },
        A3: { v: 'Prazo (meses)', fmt: { b: true } }, B3: 36,
        A4: { v: 'Parcela', fmt: { b: true } }, B4: { v: '=PGTO(B2;B3;-B1)', fmt: destaque },
        D1: { v: 'Preço de venda', fmt: { b: true } }, E1: { v: 80, fmt: moeda },
        D2: { v: 'Custo por unidade', fmt: { b: true } }, E2: { v: 35, fmt: moeda },
        D3: { v: 'Custos fixos do mês', fmt: { b: true } }, E3: { v: 9000, fmt: moeda },
        D4: { v: 'Unidades vendidas', fmt: { b: true } }, E4: 100,
        D5: { v: 'Lucro do mês', fmt: { b: true } }, E5: { v: '=E4*(E1-E2)-E3', fmt: destaque },
      });
      x.widths({ A: 150, B: 110, C: 30, D: 150, E: 110 });
    },
    steps: [
      {
        title: 'O caminho inverso',
        text: `<p>Um empréstimo de R$ 30.000, a 1,5% ao mês, em 36 meses, dá uma parcela de <b>${XL.formatValue(pgto(30000, 0.015, 36), moeda).text}</b> (a fórmula PGTO em B4).</p>
          <p>Mas a pergunta real costuma ser ao contrário: <b>"eu posso pagar R$ 900 por mês — quanto posso pegar emprestado?"</b>. Em vez de tentar valores no chute, use o <b>Atingir Meta</b>.</p>`,
        target: 'cell:A1:B4',
      },
      {
        title: 'Quanto posso pegar?',
        task: 'Clique em <b>B4</b> e use <b>Dados › Teste de Hipóteses › Atingir Meta</b>: Definir célula <b>B4</b>, Para valor <code>900</code>, Alternando célula <b>B1</b>. Clique em OK duas vezes.',
        hint: 'O botão "Teste de Hipóteses" fica no grupo Previsão, à direita da guia Dados.',
        target: ['cell:B4', '[data-menu="whatif"]'],
        check: (x) => near(x.val('B4'), 900, 0.01) && x.happened('goalseek', (d) => d.change === 'B1' && d.ok),
        solve: (x) => seekGoal(x, 'B4', 900, 'B1'),
        success: `Com parcelas de R$ 900, dá para pegar cerca de <b>${XL.formatValue((900 * (1 - Math.pow(1.015, -36))) / 0.015, moeda).text}</b>. O Excel testou valores em B1 até a parcela chegar exatamente em 900.`,
      },
      {
        title: 'Desfaça',
        text: '<p>E se você precisar mesmo dos R$ 30.000? Aí o que muda é o <b>prazo</b>. Primeiro, volte B1 ao valor original.</p>',
        task: 'Pressione <kbd>Ctrl</kbd>+<kbd>Z</kbd> para desfazer o Atingir Meta.',
        target: 'cell:B1',
        check: (x) => x.val('B1') === 30000,
        solve: (x) => { x.app.select('B1'); x.app.exec('undo'); if (x.val('B1') !== 30000) commit(x, 'B1', 30000); },
      },
      {
        title: 'Em quantos meses?',
        task: 'Use o <b>Atingir Meta</b> de novo: Definir célula <b>B4</b>, Para valor <code>900</code>, mas agora Alternando célula <b>B3</b> (o prazo).',
        target: ['cell:B3:B4', '[data-menu="whatif"]'],
        check: (x) => x.val('B1') === 30000 && near(x.val('B4'), 900, 0.01) && x.happened('goalseek', (d) => d.change === 'B3' && d.ok),
        solve: (x) => seekGoal(x, 'B4', 900, 'B3'),
        success: 'Cerca de <b>46,6 meses</b>. Como ninguém paga 0,6 parcela, o prazo real seria 47 meses — dá para arredondar com <code>ARREDONDAR.PARA.CIMA</code>.',
      },
      {
        title: 'Ponto de equilíbrio',
        text: '<p>Ao lado, uma loja vende a R$ 80 um produto que custa R$ 35, com R$ 9.000 de custos fixos no mês. <b>Quantas unidades precisa vender para não ter prejuízo?</b> É o ponto em que o lucro é zero.</p>',
        task: 'Use o Atingir Meta: Definir célula <b>E5</b>, Para valor <code>0</code>, Alternando célula <b>E4</b>.',
        target: ['cell:D1:E5', '[data-menu="whatif"]'],
        check: (x) => near(x.val('E5'), 0, 0.01) && near(x.val('E4'), 200, 0.01),
        solve: (x) => seekGoal(x, 'E5', 0, 'E4'),
        success: '<b>200 unidades</b>: cada uma deixa R$ 45 (80 − 35) e 200 × 45 = R$ 9.000, exatamente os custos fixos. Acima disso, a loja tem lucro.',
      },
    ],
  });

  const gsDemo = (setRef, to, change) => [
    cell(setRef, `Clique em <b>${setRef}</b>, a célula com a fórmula`),
    menu('whatif', 'Na guia <b>Dados</b>, clique em <b>Teste de Hipóteses</b>'),
    item('goalseek', undefined, 'Escolha <b>Atingir Meta...</b>'),
    look('input[data-role="gs-set"]', `"Definir célula" já veio com <b>${setRef}</b>`),
    typeIn('input[data-role="gs-to"]', String(to), `Em <b>Para valor</b>, digite <code>${to}</code>`),
    typeIn('input[data-role="gs-change"]', change, `Em <b>Alternando célula</b>, digite <code>${change}</code>`),
    ok(),
    look('.ui-modal', 'O Excel encontrou uma solução'),
    el('[data-role="gs-ok"]', 'Clique em <b>OK</b> para manter o resultado'),
  ];
  R('a10-1', {
    'Quanto posso pegar?': gsDemo('B4', 900, 'B1'),
    'Desfaça': [key('Ctrl+Z', 'Pressione <kbd>Ctrl</kbd>+<kbd>Z</kbd>'), look('cell:B1', 'B1 voltou para R$ 30.000')],
    'Em quantos meses?': gsDemo('B4', 900, 'B3'),
    'Ponto de equilíbrio': gsDemo('E5', 0, 'E4'),
  });

  // ======================================================================
  // 10.2 Gerenciador de Cenários e Tabela de Dados
  // ======================================================================
  const PRECOS = [40, 45, 50, 55, 60, 65];
  const lucro = (p, q, cu, cf) => q * (p - cu) - cf;
  const scen = (x, n) => x.sheet.scenarios.find((s) => XL.normName(s.name) === XL.normName(n));
  XL.registerLesson({
    id: 'a10-2',
    title: 'Gerenciador de Cenários e Tabela de Dados',
    summary: [
      '<b>Tabela de Dados</b> (Teste de Hipóteses): testa vários valores de uma entrada de uma vez. Valores numa coluna → "Célula de entrada da coluna".',
      'A Tabela de Dados é <b>viva</b>: muda o modelo, ela recalcula (as células mostram <code>=TABELA(;$B$1)</code>).',
      '<b>Gerenciador de Cenários</b>: guarda conjuntos de valores com nome (Otimista, Pessimista...) e alterna entre eles com <b>Mostrar</b>.',
      '<b>Resumo</b> cria uma planilha comparando os cenários lado a lado.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Preço', fmt: { b: true } }, B1: { v: 50, fmt: moeda },
        A2: { v: 'Quantidade vendida', fmt: { b: true } }, B2: 1000,
        A3: { v: 'Custo por unidade', fmt: { b: true } }, B3: { v: 30, fmt: moeda },
        A4: { v: 'Custos fixos', fmt: { b: true } }, B4: { v: 12000, fmt: moeda },
        A6: { v: 'Lucro', fmt: { b: true } }, B6: { v: '=B2*(B1-B3)-B4', fmt: destaque },
        D1: { v: 'Preço testado', fmt: head }, E1: { v: 'Lucro', fmt: head },
        E2: { v: '=B6', fmt: { numFmt: 'moeda', color: '#808080' } },
      };
      PRECOS.forEach((p, i) => { m['D' + (i + 3)] = { v: p, fmt: moeda }; m['E' + (i + 3)] = { v: null, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 150, B: 110, C: 30, D: 110, E: 120 });
      x.app.sheet.scenarios.push(
        { name: 'Realista', cells: '$B$1:$B$2', values: ['50', '1000'], comment: 'Preço e volume atuais' },
        { name: 'Pessimista', cells: '$B$1:$B$2', values: ['45', '800'], comment: 'Concorrente baixou o preço' });
    },
    steps: [
      {
        title: 'E se...?',
        text: `<p>O lucro (B6) depende do preço, da quantidade e dos custos. Perguntas típicas: <b>"e se o preço for R$ 45? E R$ 60?"</b>, <b>"e no pior cenário?"</b>.</p>
          <p>Em vez de mudar B1 várias vezes e anotar, use a <b>Tabela de Dados</b> (vários preços de uma vez) e o <b>Gerenciador de Cenários</b> (conjuntos de valores com nome).</p>
          <p>A tabela em D2:E8 já está pronta: a fórmula <code>=B6</code> no topo (E2) e os preços a testar embaixo.</p>`,
        target: ['cell:A1:B6', 'cell:D1:E8'],
      },
      {
        title: 'Monte a Tabela de Dados',
        task: 'Selecione <b>D2:E8</b> e use <b>Dados › Teste de Hipóteses › Tabela de Dados</b>. Em <b>Célula de entrada da coluna</b>, digite <code>B1</code> (o preço) e clique em OK.',
        hint: 'Os preços estão numa coluna, por isso a entrada é "da coluna". A célula de entrada é onde o modelo lê o preço: B1.',
        target: ['cell:D2:E8', '[data-menu="whatif"]'],
        check: (x) => x.sheet.dataTables.length > 0 && PRECOS.every((p, i) => near(x.val('E' + (i + 3)), lucro(p, 1000, 30, 12000))),
        solve: (x) => x.app.createDataTable(XL.parseRange('D2:E8'), '', 'B1'),
        success: 'Seis preços testados de uma vez: a R$ 40 o lucro some, a R$ 65 passa de R$ 20 mil. Clique numa célula de E3:E8: a fórmula é <code>=TABELA(;$B$1)</code> — o Excel coloca cada preço em B1 "de mentirinha" e anota o lucro.',
      },
      {
        title: 'A tabela é viva',
        task: 'O fornecedor aumentou o custo: em <b>B3</b>, digite <code>32</code>. Veja a Tabela de Dados se atualizar.',
        target: 'cell:B3',
        check: (x) => x.val('B3') === 32 && PRECOS.every((p, i) => near(x.val('E' + (i + 3)), lucro(p, 1000, 32, 12000))),
        solve: (x) => commit(x, 'B3', 32),
        success: 'Todos os lucros caíram R$ 2.000 (R$ 2 a mais em cada uma das 1.000 unidades). Diferente de copiar valores, a Tabela de Dados acompanha o modelo.',
      },
      {
        title: 'Crie o cenário Otimista',
        text: '<p>Já existem dois cenários guardados: <b>Realista</b> (R$ 50 e 1.000 unidades) e <b>Pessimista</b> (R$ 45 e 800). Falta o otimista.</p>',
        task: 'Selecione <b>B1:B2</b>, abra <b>Teste de Hipóteses › Gerenciador de Cenários</b>, clique em <b>Adicionar</b>, dê o nome <code>Otimista</code> e clique em OK. Digite os valores <code>60</code> e <code>1500</code>, OK, e depois <b>Fechar</b>.',
        target: ['cell:B1:B2', '[data-menu="whatif"]'],
        check: (x) => { const s = scen(x, 'Otimista'); return !!s && Number(String(s.values[0]).replace(',', '.')) === 60 && Number(String(s.values[1]).replace(',', '.')) === 1500; },
        solve: (x) => { x.app.pushUndo(); x.sheet.scenarios.push({ name: 'Otimista', cells: '$B$1:$B$2', values: ['60', '1500'], comment: '' }); x.app.emit('scenario', { name: 'Otimista' }); },
        success: 'Três cenários guardados. Repare que B1 e B2 não mudaram: guardar um cenário não aplica os valores.',
      },
      {
        title: 'Mostre o cenário Pessimista',
        task: 'No <b>Gerenciador de Cenários</b>, escolha <b>Pessimista</b> na lista, clique em <b>Mostrar</b> e depois em <b>Fechar</b>.',
        target: '[data-menu="whatif"]',
        check: (x) => x.val('B1') === 45 && x.val('B2') === 800 && x.happened('scenarioshow'),
        solve: (x) => x.app.scenarioShow(scen(x, 'Pessimista')),
        success: `B1 e B2 receberam os valores do cenário e o lucro caiu para <b>${XL.formatValue(lucro(45, 800, 32, 12000), moeda).text}</b> — prejuízo! Bom saber antes de acontecer.`,
      },
      {
        title: 'Compare tudo num resumo',
        task: 'No <b>Gerenciador de Cenários</b>, clique em <b>Resumo...</b>, em <b>Células de resultado</b> digite <code>B6</code> e clique em OK.',
        target: '[data-menu="whatif"]',
        check: (x) => x.hasSheet('Resumo do Cenário') && x.happened('scenariosummary'),
        solve: (x) => x.app.scenarioSummary('B6'),
        success: 'Uma planilha nova, <b>Resumo do Cenário</b>, mostra lado a lado os valores e o lucro de cada cenário — pronta para levar à reunião.',
      },
    ],
  });

  const openScenarios = [menu('whatif', 'Na guia <b>Dados</b>, clique em <b>Teste de Hipóteses</b>'), item('scenario-manager', undefined, 'Escolha <b>Gerenciador de Cenários...</b>')];
  R('a10-2', {
    'Monte a Tabela de Dados': [
      drag('D2', 'E8', 'Selecione a tabela inteira: de <b>D2</b> até <b>E8</b>'),
      menu('whatif', 'Na guia <b>Dados</b>, clique em <b>Teste de Hipóteses</b>'),
      item('datatable', undefined, 'Escolha <b>Tabela de Dados...</b>'),
      typeIn('input[data-role="dt-col"]', 'B1', 'Em <b>Célula de entrada da coluna</b>, digite <code>B1</code>'),
      ok(),
    ],
    'A tabela é viva': write('B3', '32').concat([look('cell:E3:E8', 'Os lucros da tabela se atualizaram')]),
    'Crie o cenário Otimista': [
      drag('B1', 'B2', 'Selecione as células que mudam em cada cenário: <b>B1:B2</b>'),
      ...openScenarios,
      el('[data-role="sc-add"]', 'Clique em <b>Adicionar...</b>'),
      typeIn('input[data-role="sc-name"]', 'Otimista', 'Em <b>Nome do cenário</b>, digite <code>Otimista</code>'),
      look('input[data-role="sc-cells"]', 'As células variáveis já vieram da seleção: <b>$B$1:$B$2</b>'),
      ok(),
      typeIn('input[data-role="sc-val"][data-cell="B1"]', '60', 'Preço: <code>60</code>'),
      typeIn('input[data-role="sc-val"][data-cell="B2"]', '1500', 'Quantidade: <code>1500</code>'),
      ok(),
      el('[data-role="sc-close"]', 'Clique em <b>Fechar</b>'),
    ],
    'Mostre o cenário Pessimista': [
      ...openScenarios,
      set('select[data-role="sc-list"]', 'Pessimista', 'Na lista, escolha <b>Pessimista</b>'),
      el('[data-role="sc-show"]', 'Clique em <b>Mostrar</b>'),
      look('cell:B6', 'O lucro mudou na planilha, atrás da janela'),
      el('[data-role="sc-close"]', 'Clique em <b>Fechar</b>'),
    ],
    'Compare tudo num resumo': [
      ...openScenarios,
      el('[data-role="sc-summary"]', 'Clique em <b>Resumo...</b>'),
      typeIn('input[data-role="sc-result"]', 'B6', 'Em <b>Células de resultado</b>, digite <code>B6</code> (o lucro)'),
      ok(),
    ],
  });

  // ======================================================================
  // 10.3 Solver
  // ======================================================================
  const solveNow = (x) => {
    const S = x.sheet;
    const res = S.solve(S.solver);
    x.app.pushUndo();
    if (res.status === 'ok') res.vars.forEach((p, i) => S.setValue(p.r, p.c, res.x[i]));
    x.app.render();
    x.app.emit('solver', { status: res.status, kept: res.status === 'ok', x: res.x });
  };
  const model = (x) => x.sheet.solver || {};
  const sameRef = (a, b) => String(a || '').replace(/\$/g, '').toUpperCase() === String(b).toUpperCase();
  XL.registerLesson({
    id: 'a10-3',
    title: 'Solver',
    summary: [
      'O <b>Solver</b> encontra a <b>melhor</b> combinação (máximo ou mínimo) respeitando <b>restrições</b> — o Atingir Meta só acerta um valor.',
      'Três partes: <b>objetivo</b> (a célula a maximizar/minimizar), <b>células variáveis</b> (as decisões) e <b>restrições</b> (os limites).',
      'Restrições com intervalos: <code>$D$8:$D$9 &lt;= $E$8:$E$9</code> vale para cada linha.',
      'Mudou um dado? Rode o Solver de novo: o modelo fica salvo na planilha. No Excel, ative o suplemento em Arquivo › Opções › Suplementos.',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Planejamento da produção da marcenaria', fmt: { b: true, size: 13 } },
        B3: { v: 'Mesas', fmt: head }, C3: { v: 'Cadeiras', fmt: head },
        A4: { v: 'Quantidade a produzir', fmt: { b: true } }, B4: { v: 0, fmt: { fill: '#FFF2CC', b: true } }, C4: { v: 0, fmt: { fill: '#FFF2CC', b: true } },
        A5: { v: 'Lucro por unidade', fmt: { b: true } }, B5: { v: 120, fmt: moeda }, C5: { v: 50, fmt: moeda },
        A7: { v: 'Recursos', fmt: head }, B7: { v: 'Por mesa', fmt: head }, C7: { v: 'Por cadeira', fmt: head }, D7: { v: 'Usado', fmt: head }, E7: { v: 'Disponível', fmt: head },
        A8: 'Madeira (peças)', B8: 30, C8: 10, D8: '=SOMARPRODUTO(B8:C8;$B$4:$C$4)', E8: 1500,
        A9: 'Horas de trabalho', B9: 4, C9: 2, D9: '=SOMARPRODUTO(B9:C9;$B$4:$C$4)', E9: 240,
        A11: { v: 'Lucro total', fmt: { b: true } }, B11: { v: '=SOMARPRODUTO(B5:C5;B4:C4)', fmt: destaque },
      });
      x.widths({ A: 170, B: 95, C: 95, D: 80, E: 90 });
    },
    steps: [
      {
        title: 'O problema',
        text: `<p>Uma marcenaria faz <b>mesas</b> (lucro R$ 120) e <b>cadeiras</b> (R$ 50). Cada mesa gasta 30 peças de madeira e 4 horas; cada cadeira, 10 peças e 2 horas. No mês há <b>1.500 peças</b> e <b>240 horas</b>.</p>
          <p><b>Quantas de cada fazer para ganhar o máximo?</b> Só mesas? Só cadeiras? Uma mistura? Tentar no chute é lento — o <b>Solver</b> resolve.</p>
          <p>As quantidades ficam em B4:C4 (amarelo), o uso de recursos em D8:D9 e o lucro em B11.</p>`,
        target: ['cell:A3:E11'],
      },
      {
        title: 'Objetivo e variáveis',
        task: 'Abra <b>Dados › Solver</b>. Em <b>Definir Objetivo</b>, digite <code>$B$11</code>, deixe <b>Máx.</b> marcado e, em <b>Alterando Células Variáveis</b>, digite <code>$B$4:$C$4</code>. Clique em <b>Fechar</b> (o modelo fica salvo).',
        target: '.rb-btn[data-cmd="solver"]',
        check: (x) => { const m = model(x); return sameRef(m.obj, 'B11') && m.goal === 'max' && sameRef(m.vars, 'B4:C4'); },
        solve: (x) => { x.app.pushUndo(); x.sheet.solver = { obj: '$B$11', goal: 'max', target: '0', vars: '$B$4:$C$4', cons: [], nonneg: true, method: 'GRG Não Linear' }; x.app.emit('solvermodel', {}); },
        success: 'Objetivo: maximizar o lucro (B11) mudando as quantidades (B4:C4). Sem limites, o lucro seria infinito — faltam as restrições.',
      },
      {
        title: 'Restrições e solução',
        task: 'Abra o <b>Solver</b> de novo. Na linha de restrição, digite <code>$D$8:$D$9</code>, escolha <b>&lt;=</b>, digite <code>$E$8:$E$9</code> e clique em <b>Adicionar</b>. Depois clique em <b>Resolver</b> e em <b>OK</b> (Manter Solução).',
        hint: '"Usado" (D8:D9) tem que ser menor ou igual ao "Disponível" (E8:E9).',
        target: '.rb-btn[data-cmd="solver"]',
        check: (x) => x.val('B4') === 30 && x.val('C4') === 60,
        solve: (x) => { x.app.pushUndo(); x.sheet.solver.cons = [{ ref: '$D$8:$D$9', op: '<=', rhs: '$E$8:$E$9' }]; solveNow(x); },
        success: '<b>30 mesas e 60 cadeiras</b>: lucro de <b>R$ 6.600</b>, usando toda a madeira e todas as horas. Só mesas daria R$ 6.000; só cadeiras, também R$ 6.000. A mistura é melhor — difícil de achar no chute!',
      },
      {
        title: 'Chegou mais madeira',
        task: 'O fornecedor entregou mais madeira: em <b>E8</b>, digite <code>1800</code>.',
        target: 'cell:E8',
        check: (x) => x.val('E8') === 1800,
        solve: (x) => commit(x, 'E8', 1800),
        success: 'O plano antigo (30 e 60) ainda é possível, mas será que continua o melhor?',
      },
      {
        title: 'Rode o Solver de novo',
        task: 'Abra o <b>Solver</b> (o modelo continua lá) e clique em <b>Resolver</b> e depois em <b>OK</b>.',
        target: '.rb-btn[data-cmd="solver"]',
        check: (x) => x.val('B4') === 60 && x.val('C4') === 0 && x.val('E8') === 1800,
        solve: (x) => solveNow(x),
        success: 'Agora o melhor é <b>60 mesas e nenhuma cadeira</b> (R$ 7.200): com madeira sobrando, o gargalo virou as horas, e a mesa dá mais lucro por hora (R$ 30 contra R$ 25).',
      },
      {
        title: 'Um pedido mínimo',
        text: '<p>Um cliente encomendou <b>40 cadeiras</b>. Isso vira uma restrição: C4 precisa ser pelo menos 40.</p>',
        task: 'No <b>Solver</b>, adicione a restrição <code>$C$4</code> <b>&gt;=</b> <code>40</code>, clique em <b>Resolver</b> e em <b>OK</b>.',
        target: '.rb-btn[data-cmd="solver"]',
        check: (x) => x.val('B4') === 40 && x.val('C4') === 40 && (model(x).cons || []).some((k) => sameRef(k.ref, 'C4') && k.op === '>='),
        solve: (x) => { x.app.pushUndo(); x.sheet.solver.cons.push({ ref: '$C$4', op: '>=', rhs: '40' }); solveNow(x); },
        success: '<b>40 mesas e 40 cadeiras</b> (R$ 6.800): o Solver respeitou o pedido e encaixou o máximo de mesas nas horas que sobraram. Restrições são a forma de colocar a realidade no modelo.',
      },
    ],
  });

  const openSolver = btn('solver', undefined, 'Na guia <b>Dados</b>, clique em <b>Solver</b> (grupo Análise)');
  R('a10-3', {
    'Objetivo e variáveis': [
      openSolver,
      typeIn('input[data-role="sv-obj"]', '$B$11', 'Em <b>Definir Objetivo</b>, digite <code>$B$11</code> (o lucro total)'),
      el('input[data-role="sv-max"]', 'Deixe marcado <b>Máx.</b> — queremos o maior lucro'),
      typeIn('input[data-role="sv-vars"]', '$B$4:$C$4', 'Em <b>Alterando Células Variáveis</b>, digite <code>$B$4:$C$4</code> (as quantidades)'),
      el('[data-role="sv-close"]', 'Clique em <b>Fechar</b> — o modelo fica salvo'),
    ],
    'Restrições e solução': [
      openSolver,
      typeIn('input[data-role="sv-cref"]', '$D$8:$D$9', 'Na linha de restrição, digite <code>$D$8:$D$9</code> (o que foi usado)'),
      set('select[data-role="sv-cop"]', '<=', 'Escolha <b>&lt;=</b>'),
      typeIn('input[data-role="sv-crhs"]', '$E$8:$E$9', 'Digite <code>$E$8:$E$9</code> (o disponível)'),
      el('[data-role="sv-cadd"]', 'Clique em <b>Adicionar</b>'),
      el('[data-role="sv-solve"]', 'Clique em <b>Resolver</b>'),
      look('.sv-msg', 'O Solver encontrou a solução ótima'),
      el('[data-role="sv-ok"]', 'Clique em <b>OK</b> para manter a solução'),
    ],
    'Chegou mais madeira': write('E8', '1800'),
    'Rode o Solver de novo': [
      openSolver,
      el('[data-role="sv-solve"]', 'O modelo continua lá: clique em <b>Resolver</b>'),
      el('[data-role="sv-ok"]', 'Clique em <b>OK</b>'),
    ],
    'Um pedido mínimo': [
      openSolver,
      typeIn('input[data-role="sv-cref"]', '$C$4', 'Na linha de restrição, digite <code>$C$4</code> (as cadeiras)'),
      set('select[data-role="sv-cop"]', '>=', 'Escolha <b>&gt;=</b>'),
      typeIn('input[data-role="sv-crhs"]', '40', 'Digite <code>40</code>'),
      el('[data-role="sv-cadd"]', 'Clique em <b>Adicionar</b>'),
      el('[data-role="sv-solve"]', 'Clique em <b>Resolver</b>'),
      el('[data-role="sv-ok"]', 'Clique em <b>OK</b>'),
    ],
  });

  // ======================================================================
  // 10.4 Previsões e tendências
  // ======================================================================
  const MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const VENDAS = [52000, 54100, 53800, 57200, 58900, 58300, 61500, 62800, 62100, 65400, 67200, 68900];
  const reg = (() => {
    const n = 12, xs = VENDAS.map((_, i) => i + 1);
    const mx = xs.reduce((s, v) => s + v, 0) / n, my = VENDAS.reduce((s, v) => s + v, 0) / n;
    let sxy = 0, sxx = 0, syy = 0;
    xs.forEach((x, i) => { sxy += (x - mx) * (VENDAS[i] - my); sxx += (x - mx) ** 2; syy += (VENDAS[i] - my) ** 2; });
    const a = sxy / sxx;
    return { a, b: my - a * mx, r2: (sxy * sxy) / (sxx * syy) };
  })();
  const prev = (x) => reg.b + reg.a * x;
  const chartOf = (x) => x.sheet.charts.find((c) => XL.rangeToStr(c.g) === 'A1:B13') || null;
  XL.registerLesson({
    id: 'a10-4',
    title: 'Previsões e tendências',
    summary: [
      '<code>INCLINAÇÃO(vendas; períodos)</code> = quanto cresce, em média, a cada período.',
      '<code>PREVISÃO.LINEAR(x; vendas; períodos)</code> = o valor da reta no período x (ex.: o próximo mês).',
      '<code>TENDÊNCIA(vendas; períodos; novos_períodos)</code> prevê vários períodos de uma vez (o resultado derrama).',
      '<code>RQUAD</code> diz o quanto a reta explica os dados (perto de 1 = boa). Nos gráficos: <b>Design do Gráfico › Linha de Tendência</b>.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Mês', fmt: head }, B1: { v: 'Vendas', fmt: head }, C1: { v: 'Nº', fmt: head },
        E1: { v: 'Crescimento por mês', fmt: { b: true } }, F1: { v: null, fmt: destaque },
        E2: { v: 'Previsão para jan/26', fmt: { b: true } }, F2: { v: null, fmt: destaque },
        E3: { v: 'R² (qualidade da reta)', fmt: { b: true } }, F3: { v: null, fmt: { b: true, fill: '#FFF2CC', numFmt: 'numero', dec: 2 } } };
      VENDAS.forEach((v, i) => { const r = i + 2; m['A' + r] = MESES[i] + '/25'; m['B' + r] = { v, fmt: moeda }; m['C' + r] = i + 1; });
      ['jan/26', 'fev/26', 'mar/26'].forEach((lbl, i) => { const r = i + 14; m['A' + r] = { v: lbl, fmt: { i: true, color: '#595959' } }; m['B' + r] = { v: null, fmt: { numFmt: 'moeda', i: true, fill: '#EDEDED' } }; m['C' + r] = 13 + i; });
      x.set(m);
      x.widths({ A: 70, B: 110, C: 40, D: 24, E: 170, F: 120 });
    },
    steps: [
      {
        title: 'Para onde as vendas vão?',
        text: `<p>As vendas de 2025 sobem quase todo mês, com altos e baixos. A pergunta do planejamento: <b>quanto esperar para os próximos meses?</b></p>
          <p>A coluna <b>Nº</b> (1 a 12) numera os meses — as funções de previsão precisam de números no eixo do tempo. Vamos traçar uma <b>reta</b> que passa o mais perto possível de todos os pontos (regressão linear) e usá-la para prever.</p>`,
        target: 'cell:A1:C16',
      },
      {
        title: 'Quanto cresce por mês?',
        task: 'Em <b>F1</b>, digite <code>=INCLINAÇÃO(B2:B13;C2:C13)</code>.',
        target: 'cell:F1',
        check: (x) => x.uses('F1', 'INCLINAÇÃO') && near(x.val('F1'), reg.a, 1e-6),
        solve: (x) => commit(x, 'F1', '=INCLINAÇÃO(B2:B13;C2:C13)'),
        success: `Em média, as vendas crescem <b>${XL.formatValue(reg.a, moeda).text} por mês</b>. A ordem importa: primeiro o que você quer prever (vendas), depois o tempo (Nº).`,
      },
      {
        title: 'Previsão do próximo mês',
        task: 'Em <b>F2</b>, preveja o mês 13 (jan/26): <code>=PREVISÃO.LINEAR(13;B2:B13;C2:C13)</code>.',
        target: 'cell:F2',
        check: (x) => x.uses('F2', 'PREVISÃO.LINEAR') && near(x.val('F2'), prev(13), 1e-6),
        solve: (x) => commit(x, 'F2', '=PREVISÃO.LINEAR(13;B2:B13;C2:C13)'),
        success: `Cerca de <b>${XL.formatValue(prev(13), moeda).text}</b> em janeiro, se a tendência continuar. É uma estimativa, não uma certeza: use como referência para metas e estoque.`,
      },
      {
        title: 'Vários meses de uma vez',
        task: 'Em <b>B14</b>, digite <code>=TENDÊNCIA(B2:B13;C2:C13;C14:C16)</code> para prever jan, fev e mar de 2026.',
        target: 'cell:B14:B16',
        check: (x) => x.uses('B14', 'TENDÊNCIA') && [13, 14, 15].every((n, i) => near(x.val('B' + (14 + i)), prev(n), 1e-6)),
        solve: (x) => commit(x, 'B14', '=TENDÊNCIA(B2:B13;C2:C13;C14:C16)'),
        success: 'Três previsões com uma fórmula (o resultado derramou). Se as vendas de 2025 forem corrigidas, as previsões se ajustam sozinhas.',
      },
      {
        title: 'A reta é confiável?',
        text: '<p>Uma reta sempre pode ser traçada — mas será que ela representa bem os dados? O <b>R²</b> vai de 0 a 1: quanto mais perto de 1, mais os pontos ficam colados na reta.</p>',
        task: 'Em <b>F3</b>, digite <code>=RQUAD(B2:B13;C2:C13)</code>.',
        target: 'cell:F3',
        check: (x) => x.uses('F3', 'RQUAD') && near(x.val('F3'), reg.r2, 1e-9),
        solve: (x) => commit(x, 'F3', '=RQUAD(B2:B13;C2:C13)'),
        success: `R² de <b>${reg.r2.toFixed(2).replace('.', ',')}</b>: a reta explica a maior parte da variação — a tendência de alta é consistente. Com R² baixo (0,3, por exemplo), desconfie da previsão.`,
      },
      {
        title: 'Veja no gráfico',
        task: 'Selecione <b>A1:B13</b> e clique em <b>Inserir › Gráfico de Linhas</b>.',
        target: ['cell:A1:B13', '.rb-btn[data-cmd="chart-insert"][data-arg="line"]'],
        check: (x) => { const c = chartOf(x); return !!c && c.type === 'line'; },
        solve: (x) => { x.app.select('A1:B13'); x.app.insertChart('line'); },
      },
      {
        title: 'Linha de tendência',
        task: 'Com o gráfico selecionado, marque <b>Design do Gráfico › Linha de Tendência</b>.',
        target: ['.xl-chart', 'input[data-cmd="chart-trend"]'],
        check: (x) => { const c = chartOf(x); return !!c && !!c.trend; },
        solve: (x) => { const c = chartOf(x); x.app.selectChart(c.id); x.app.chartUpdate((q) => { q.trend = true; }); },
        success: 'A linha tracejada é a mesma reta das funções: passe o mouse sobre ela para ver a equação. Agora a tendência está visível para qualquer pessoa — o melhor jeito de apresentar uma previsão.',
      },
    ],
  });

  R('a10-4', {
    'Quanto cresce por mês?': write('F1', '=INCLINAÇÃO(B2:B13;C2:C13)', 'Digite <code>=INCLINAÇÃO(B2:B13;C2:C13)</code> — vendas primeiro, depois os meses'),
    'Previsão do próximo mês': write('F2', '=PREVISÃO.LINEAR(13;B2:B13;C2:C13)', 'Digite <code>=PREVISÃO.LINEAR(13;B2:B13;C2:C13)</code> — o 13 é jan/26'),
    'Vários meses de uma vez': write('B14', '=TENDÊNCIA(B2:B13;C2:C13;C14:C16)').concat([look('cell:B14:B16', 'Três meses previstos')]),
    'A reta é confiável?': write('F3', '=RQUAD(B2:B13;C2:C13)'),
    'Veja no gráfico': [
      drag('A1', 'B13', 'Selecione os meses e as vendas: de <b>A1</b> até <b>B13</b>'),
      btn('chart-insert', 'line', 'Na guia <b>Inserir</b>, clique em <b>Gráfico de Linhas</b>'),
    ],
    'Linha de tendência': [
      el('.xl-chart', 'Clique no gráfico para selecioná-lo'),
      tab('chartdesign', 'Clique na guia <b>Design do Gráfico</b>'),
      el('input[data-cmd="chart-trend"]', 'Em <b>Design do Gráfico</b>, marque <b>Linha de Tendência</b>'),
    ],
  });
})();
