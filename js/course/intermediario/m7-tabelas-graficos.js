/* Nível Intermediário · Módulo 7 — Tabelas e visualização (4 lições + demonstrações). */
(function () {
  'use strict';
  const XL = window.XL;
  const K = XL.SHEET_KEY;
  const head = { b: true, fill: '#DDEBF7' };
  const moeda = { numFmt: 'moeda' };
  const { cell, drag, point, type, key, tab, btn, menu, item, el, set, typeIn, look } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const ok = (say) => el('.ui-modal .btn-primary', say || 'Clique em <b>OK</b>');
  const g = (s) => XL.parseRange(s);
  const commit = (x, ref, text) => { const p = XL.parseRef(ref); x.app.select(ref); return x.app.commitInput(p.r, p.c, String(text)); };
  const sameRange = (a, s) => XL.rangeToStr(a) === s;
  const covers = (a, s) => { const b = g(s); return a.r1 <= b.r1 && a.c1 <= b.c1 && a.r2 >= b.r2 && a.c2 >= b.c2; };

  // ======================================================================
  // 7.1 Formatar como Tabela
  // ======================================================================
  const VENDAS = [['Ana', 'Notebook', 2, 3500], ['Bruno', 'Mouse', 10, 80], ['Carla', 'Monitor', 3, 1200],
    ['Diego', 'Teclado', 5, 150], ['Elisa', 'Notebook', 1, 3500], ['Fábio', 'Monitor', 2, 1200]];
  const table = (x) => x.sheet.tables[0] || null;
  XL.registerLesson({
    id: 'i7-1',
    title: 'Formatar como Tabela',
    summary: [
      '<kbd>Ctrl</kbd>+<kbd>T</kbd> (ou Inserir › Tabela) transforma um intervalo em <b>Tabela</b>: cores em tiras, filtros e a guia <b>Design da Tabela</b>.',
      'Dê um <b>nome</b> à tabela (ex.: <code>Vendas</code>) para usá-lo nas fórmulas.',
      'Dentro da tabela, <code>[@Qtd]</code> é "a Qtd desta linha". Uma fórmula numa coluna vazia vira <b>coluna calculada</b> e se preenche sozinha.',
      'Digitou logo abaixo? A tabela <b>cresce</b> e leva fórmulas e formatos junto.',
      'Fora da tabela: <code>=SOMASE(Vendas[Produto];"Notebook";Vendas[Total])</code> — referências que se ajustam quando a tabela cresce.',
    ],
    setup: (x) => {
      const m = { A1: 'Vendedor', B1: 'Produto', C1: 'Qtd', D1: 'Preço unit.', E1: 'Total', G1: { v: 'Total de notebooks', fmt: { b: true } } };
      VENDAS.forEach(([v, p, q, u], i) => {
        const r = i + 2;
        m['A' + r] = v; m['B' + r] = p; m['C' + r] = q; m['D' + r] = { v: u, fmt: moeda }; m['E' + r] = { v: null, fmt: moeda };
      });
      m.G2 = { v: null, fmt: moeda };
      x.set(m);
      x.widths({ A: 90, B: 95, C: 55, D: 100, E: 110, F: 30, G: 140 });
    },
    steps: [
      {
        title: 'Intervalo comum × Tabela',
        text: `<p>Esta lista de vendas é um <b>intervalo comum</b>: só células com dados. Quando você a transforma em <b>Tabela</b>, o Excel passa a entender que ela é um bloco único:</p>
          <ul><li>linhas em tiras e botões de filtro automáticos;</li><li>fórmulas que se copiam sozinhas;</li><li>a tabela cresce quando você digita logo abaixo;</li><li>um nome para usar nas fórmulas.</li></ul>`,
        target: 'cell:A1:E7',
      },
      {
        title: 'Crie a tabela',
        text: '<p>Não precisa selecionar tudo: basta clicar em <b>uma célula</b> dos dados. O Excel descobre o resto.</p>',
        task: 'Clique em qualquer célula da lista (ex.: <b>B3</b>), pressione <kbd>Ctrl</kbd>+<kbd>T</kbd> e confirme com <b>OK</b>.',
        hint: 'Também dá para usar <b>Inserir › Tabela</b> ou <b>Página Inicial › Formatar como Tabela</b>.',
        target: ['cell:A1:E7', '.rb-btn[data-cmd="table-create"]'],
        check: (x) => { const t = table(x); return !!t && sameRange(t.g, 'A1:E7'); },
        solve: (x) => { x.app.select('B3'); x.app.createTable('azul', 'A1:E7'); },
        success: 'Virou Tabela! Repare nas <b>linhas em tiras</b>, nos <b>botões de filtro</b> e na nova guia <b>Design da Tabela</b> lá no alto — ela só aparece quando você clica dentro da tabela.',
      },
      {
        title: 'Dê um nome à tabela',
        text: '<p>O Excel chamou a tabela de <code>Tabela1</code>. Um nome que diga o que ela guarda deixa as fórmulas muito mais fáceis de ler.</p>',
        task: 'Na guia <b>Design da Tabela</b>, troque o <b>Nome da Tabela</b> para <code>Vendas</code> e pressione <kbd>Enter</kbd>.',
        hint: 'Clique numa célula da tabela para a guia aparecer. O nome não pode ter espaços.',
        target: ['.xl-tab[data-tab="tabledesign"]', 'input[data-cmd="table-rename"]'],
        check: (x) => { const t = table(x); return !!t && t.name === 'Vendas'; },
        solve: (x) => { x.app.select('B3'); x.app.tableRename('Vendas'); },
      },
      {
        title: 'Coluna calculada com [@Coluna]',
        text: `<p>Dentro de uma tabela, ao clicar numa célula da <b>mesma linha</b>, o Excel escreve o <b>nome da coluna</b> em vez de "C2": <code>[@Qtd]</code> quer dizer "a Qtd desta linha".</p>
          <p>Quando o nome tem espaço ou ponto, ele vem com colchetes duplos: <code>[@[Preço unit.]]</code>.</p>`,
        task: 'Em <b>E2</b>, digite <code>=</code>, clique em <b>C2</b>, digite <code>*</code>, clique em <b>D2</b> e pressione <kbd>Enter</kbd>.',
        target: ['cell:E2', 'cell:C2:D2'],
        check: (x) => x.input('E2').includes('[@') && VENDAS.every(([, , q, u], i) => x.num('E' + (i + 2)) === q * u),
        solve: (x) => commit(x, 'E2', '=[@Qtd]*[@[Preço unit.]]'),
        success: 'A fórmula <code>=[@Qtd]*[@[Preço unit.]]</code> se preencheu <b>sozinha</b> até o fim da tabela — sem arrastar nada! Isso é uma <b>coluna calculada</b>.',
      },
      {
        title: 'A tabela cresce sozinha',
        text: '<p>Chegou uma venda nova. Em vez de ajustar fórmulas e formatos, é só digitar na primeira linha <b>logo abaixo</b> da tabela.</p>',
        task: 'Na linha 8, digite <code>Rui</code> em <b>A8</b>, <code>Teclado</code> em B8, <code>2</code> em C8 e <code>150</code> em D8 (use <kbd>Tab</kbd> para ir para a direita).',
        target: 'cell:A8:D8',
        check: (x) => { const t = table(x); return !!t && t.g.r2 >= 7 && x.val('A8') === 'Rui' && x.num('E8') === 300; },
        solve: (x) => { commit(x, 'A8', 'Rui'); commit(x, 'B8', 'Teclado'); commit(x, 'C8', 2); commit(x, 'D8', 150); },
        success: 'A tabela ganhou a linha 8 com cores, formato de moeda e o <b>Total</b> já calculado (R$ 300,00).',
      },
      {
        title: 'Linha de Totais',
        task: 'Clique dentro da tabela e, na guia <b>Design da Tabela</b>, marque <b>Linha de Totais</b>.',
        hint: 'Fica no grupo "Opções de Estilo de Tabela".',
        target: ['.xl-tab[data-tab="tabledesign"]', 'input[data-cmd="table-totals"]'],
        check: (x) => { const t = table(x); return !!t && t.totalRow && x.num('E9') === 18350; },
        solve: (x) => { x.app.select('B3'); x.app.tableTotals(true); },
        success: 'Apareceu a linha <b>Total</b> com <code>=SUBTOTAL(109;Vendas[Total])</code>. O SUBTOTAL soma só as linhas visíveis: se você filtrar a tabela, o total acompanha.',
      },
      {
        title: 'Mude o estilo',
        task: 'Em <b>Design da Tabela › Estilos de Tabela</b>, escolha o estilo <b>Verde</b>.',
        target: ['.xl-tab[data-tab="tabledesign"]', '.rb-btn[data-cmd="table-style"][data-arg="verde"]'],
        check: (x) => { const t = table(x); return !!t && t.style === 'verde'; },
        solve: (x) => { x.app.select('B3'); x.app.exec('table-style', 'verde'); },
      },
      {
        title: 'Usando o nome da tabela nas fórmulas',
        text: `<p>Fora da tabela, ao arrastar sobre uma coluna inteira de dados, o Excel escreve <code>Vendas[Produto]</code> em vez de <code>B2:B8</code>.</p>
          <p>A vantagem: quando a tabela crescer, a fórmula continua pegando <b>todas</b> as linhas.</p>`,
        task: 'Em <b>G2</b>, some o total dos notebooks: digite <code>=SOMASE(</code>, arraste de <b>B2</b> a <b>B8</b>, digite <code>;"Notebook";</code>, arraste de <b>E2</b> a <b>E8</b>, feche com <code>)</code> e <kbd>Enter</kbd>.',
        hint: 'O resultado deve ser R$ 10.500,00. Você também pode digitar direto: <code>=SOMASE(Vendas[Produto];"Notebook";Vendas[Total])</code>.',
        target: ['cell:G2', 'cell:B2:B8', 'cell:E2:E8'],
        check: (x) => x.uses('G2', 'SOMASE') && /VENDAS\[/i.test(x.input('G2')) && x.num('G2') === 10500,
        solve: (x) => commit(x, 'G2', '=SOMASE(Vendas[Produto];"Notebook";Vendas[Total])'),
        success: `Pronto: <code>=SOMASE(Vendas[Produto];"Notebook";Vendas[Total])</code> se lê quase como uma frase.
          <p>Dica: se um dia precisar das células "normais" de novo, use <b>Design da Tabela › Converter em Intervalo</b> — a formatação fica e as fórmulas viram referências comuns.</p>`,
      },
    ],
  });

  R('i7-1', {
    'Crie a tabela': [
      cell('B3', 'Clique em uma célula qualquer dos dados, por exemplo <b>B3</b>'),
      key('Ctrl+T', 'Pressione <kbd>Ctrl</kbd>+<kbd>T</kbd> (de <b>T</b>abela)'),
      look('input[data-role="tbl-range"]', 'O Excel encontrou sozinho o intervalo dos dados: <b>$A$1:$E$7</b>'),
      look('.ui-modal .sd-check', '"Minha tabela tem cabeçalhos": a linha 1 vira o cabeçalho'),
      ok(),
    ],
    'Dê um nome à tabela': [
      cell('B3', 'Clique dentro da tabela para a guia <b>Design da Tabela</b> aparecer'),
      tab('tabledesign', 'Clique na guia <b>Design da Tabela</b>'),
      typeIn('input[data-cmd="table-rename"]', 'Vendas', 'Clique na caixa <b>Nome da Tabela</b> e digite <code>Vendas</code>'),
      key('Enter'),
    ],
    'Coluna calculada com [@Coluna]': [
      cell('E2'),
      type('=', 'Digite <code>=</code> para começar a fórmula'),
      point('C2', 'Clique em <b>C2</b> — entra <code>[@Qtd]</code>, "a Qtd desta linha"'),
      type('*', 'Digite <code>*</code> (multiplicar)'),
      point('D2', 'Clique em <b>D2</b> — entra <code>[@[Preço unit.]]</code>'),
      enter('Pressione <kbd>Enter</kbd> e veja a coluna inteira se preencher'),
      look('cell:E2:E7', 'Uma fórmula só, copiada para todas as linhas da tabela'),
    ],
    'A tabela cresce sozinha': [
      cell('A8', 'Clique em <b>A8</b>, logo abaixo da tabela'),
      type('Rui'), key('Tab', 'Pressione <kbd>Tab</kbd> para ir para a direita — a tabela já cresceu'),
      type('Teclado'), key('Tab'),
      type('2'), key('Tab'),
      type('150'), enter(),
      look('cell:A8:E8', 'Linha nova com cores, moeda e o Total calculado'),
    ],
    'Linha de Totais': [
      cell('B3', 'Clique dentro da tabela'),
      tab('tabledesign', 'Clique na guia <b>Design da Tabela</b>'),
      el('input[data-cmd="table-totals"]', 'Marque a caixa <b>Linha de Totais</b>'),
      look('cell:A9:E9', 'A linha Total soma a coluna com <code>SUBTOTAL</code>'),
    ],
    'Mude o estilo': [
      cell('B3', 'Clique dentro da tabela'),
      btn('table-style', 'verde', 'Em <b>Estilos de Tabela</b>, clique no estilo <b>Verde</b>'),
    ],
    'Usando o nome da tabela nas fórmulas': [
      cell('G2'),
      type('=SOMASE(', 'Digite <code>=SOMASE(</code>'),
      point('B2:B8', 'Arraste de <b>B2</b> até <b>B8</b> — entra <code>Vendas[Produto]</code>'),
      type(';"Notebook";', 'Digite o critério: <code>;"Notebook";</code>'),
      point('E2:E8', 'Arraste de <b>E2</b> até <b>E8</b> — entra <code>Vendas[Total]</code>'),
      type(')', 'Feche o parêntese'),
      enter(),
    ],
  });

  // ======================================================================
  // 7.2 Gráficos
  // ======================================================================
  const REGIOES = [['Norte', 4200], ['Nordeste', 6100], ['Centro-Oeste', 3800], ['Sudeste', 9500], ['Sul', 5400]];
  const MESES = [['Jan', 12000], ['Fev', 13500], ['Mar', 12800], ['Abr', 15200], ['Mai', 16900], ['Jun', 18300]];
  const chartOf = (x, range) => x.sheet.charts.find((c) => sameRange(c.g, range)) || null;
  const selChart = (x, range) => { const c = chartOf(x, range); if (c) x.app.selectChart(c.id); return c; };
  const clickChart = (range, say) => el(`.xl-chart[data-range="${range}"]`, say || 'Clique no gráfico para selecioná-lo');
  XL.registerLesson({
    id: 'i7-2',
    title: 'Gráficos de colunas, linhas e pizza',
    summary: [
      '<b>Colunas/Barras</b>: comparar categorias (regiões, produtos). <b>Linhas</b>: evolução no tempo (meses, anos).',
      '<b>Pizza</b>: só para partes de um todo, com poucas fatias (até 5 ou 6). Para comparar, colunas são mais fáceis de ler.',
      'Selecione os dados <b>com os rótulos</b> e use <b>Inserir › Gráficos</b>. Com o gráfico selecionado, aparece a guia <b>Design do Gráfico</b>.',
      'O gráfico é ligado às células: mudou o número, o gráfico muda junto.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Região', fmt: head }, B1: { v: 'Vendas', fmt: head }, A9: { v: 'Mês', fmt: head }, B9: { v: 'Faturamento', fmt: head } };
      REGIOES.forEach(([n, v], i) => { m['A' + (i + 2)] = n; m['B' + (i + 2)] = { v, fmt: moeda }; });
      MESES.forEach(([n, v], i) => { m['A' + (i + 10)] = n; m['B' + (i + 10)] = { v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 100, B: 110 });
    },
    steps: [
      {
        title: 'Qual gráfico usar?',
        text: `<p>Antes de clicar, pergunte: <b>o que eu quero mostrar?</b></p>
          <ul><li><b>Comparar</b> categorias (qual região vendeu mais?) → <b>Colunas</b> ou <b>Barras</b>.</li>
          <li><b>Evolução no tempo</b> (o faturamento está subindo?) → <b>Linhas</b>.</li>
          <li><b>Partes de um todo</b> com poucas fatias → <b>Pizza</b> (use com moderação).</li></ul>
          <p>Aqui temos as duas situações: vendas por <b>região</b> (A1:B6) e faturamento por <b>mês</b> (A9:B15).</p>`,
        target: ['cell:A1:B6', 'cell:A9:B15'],
      },
      {
        title: 'Gráfico de colunas',
        task: 'Selecione <b>A1:B6</b> (com os rótulos) e clique em <b>Inserir › Gráfico de Colunas</b>.',
        hint: 'Arraste de A1 até B6. Na guia Inserir, grupo Gráficos, é o primeiro ícone pequeno (barrinhas em pé).',
        target: ['cell:A1:B6', '.rb-btn[data-cmd="chart-insert"][data-arg="col"]'],
        check: (x) => { const c = chartOf(x, 'A1:B6'); return !!c && c.type === 'col'; },
        solve: (x) => { x.app.select('A1:B6'); x.app.insertChart('col'); },
        success: 'O gráfico apareceu ao lado dos dados e a guia <b>Design do Gráfico</b> surgiu. Passe o mouse sobre as colunas para ver os valores. Dá para <b>arrastar</b> o gráfico e <b>redimensionar</b> pelo cantinho.',
      },
      {
        title: 'Um título que diga algo',
        task: 'Com o gráfico selecionado, clique em <b>Design do Gráfico › Título do Gráfico</b>, digite <code>Vendas por região</code> e confirme com <b>OK</b>.',
        hint: 'Se a guia Design do Gráfico sumiu, clique no gráfico de novo.',
        target: ['.xl-chart[data-range="A1:B6"]', '.rb-btn[data-cmd="chart-title"]'],
        check: (x) => { const c = chartOf(x, 'A1:B6'); return !!c && XL.normName(String(c.title || '')) === XL.normName('Vendas por região'); },
        solve: (x) => { selChart(x, 'A1:B6'); x.app.chartUpdate((c) => { c.title = 'Vendas por região'; }); },
      },
      {
        title: 'Rótulos de dados',
        text: '<p>Os <b>rótulos de dados</b> escrevem o valor em cima de cada coluna — ótimo quando poucas colunas precisam ser lidas com precisão.</p>',
        task: 'Em <b>Design do Gráfico</b>, marque <b>Rótulos de Dados</b>.',
        target: ['.xl-chart[data-range="A1:B6"]', 'input[data-cmd="chart-labels"]'],
        check: (x) => { const c = chartOf(x, 'A1:B6'); return !!c && c.labels; },
        solve: (x) => { selChart(x, 'A1:B6'); x.app.chartUpdate((c) => { c.labels = true; }); },
      },
      {
        title: 'Linhas para o tempo',
        text: '<p>Não sabe qual escolher? <b>Gráficos Recomendados</b> olha os seus dados e sugere um tipo. Com meses nos rótulos, a sugestão é <b>Linhas</b>.</p>',
        task: 'Selecione <b>A9:B15</b> e clique em <b>Inserir › Gráficos Recomendados</b>.',
        hint: 'Também vale clicar direto no ícone de <b>Linhas</b>.',
        target: ['cell:A9:B15', '.rb-btn[data-cmd="chart-insert"][data-arg="rec"]'],
        check: (x) => { const c = chartOf(x, 'A9:B15'); return !!c && c.type === 'line'; },
        solve: (x) => { x.app.select('A9:B15'); x.app.insertChart('line'); },
        success: 'A linha mostra a <b>tendência</b>: o faturamento cresceu quase todo mês, com uma pequena queda em março.',
      },
      {
        title: 'Trocando o tipo: pizza',
        text: '<p>As vendas por região somam o total da empresa, então dá para ver a <b>participação</b> de cada uma numa pizza (são só 5 fatias).</p>',
        task: 'Clique no gráfico de <b>regiões</b> e use <b>Design do Gráfico › Alterar Tipo de Gráfico › Pizza</b>.',
        target: ['.xl-chart[data-range="A1:B6"]', '[data-menu="chart-type"]'],
        check: (x) => { const c = chartOf(x, 'A1:B6'); return !!c && c.type === 'pie'; },
        solve: (x) => { selChart(x, 'A1:B6'); x.app.chartUpdate((c) => { c.type = 'pie'; }); },
        success: 'Agora cada fatia mostra a fatia do total. Mas repare: comparar Norte e Centro-Oeste é <b>mais difícil</b> na pizza do que nas colunas. Use pizza só quando a mensagem for "quanto cada parte representa do todo".',
      },
      {
        title: 'O gráfico acompanha os dados',
        task: 'Mude as vendas do <b>Nordeste</b> (<b>B3</b>) para <code>9800</code> e veja a pizza se ajustar.',
        target: 'cell:B3',
        check: (x) => x.num('B3') === 9800,
        solve: (x) => commit(x, 'B3', 9800),
        success: 'O gráfico é <b>ligado</b> às células: qualquer mudança nos números aparece na hora. Por isso vale montar gráficos em cima de dados organizados — de preferência numa <b>Tabela</b>.',
      },
    ],
  });

  R('i7-2', {
    'Gráfico de colunas': [
      drag('A1', 'B6', 'Arraste de <b>A1</b> até <b>B6</b> — com os rótulos (regiões) e os valores'),
      btn('chart-insert', 'col', 'Na guia <b>Inserir</b>, clique em <b>Gráfico de Colunas</b>'),
    ],
    'Um título que diga algo': [
      clickChart('A1:B6'),
      btn('chart-title', undefined, 'Na guia <b>Design do Gráfico</b>, clique em <b>Título do Gráfico</b>'),
      typeIn('input[data-role="chart-title"]', 'Vendas por região', 'Digite <code>Vendas por região</code>'),
      ok(),
    ],
    'Rótulos de dados': [
      clickChart('A1:B6'),
      tab('chartdesign', 'Clique na guia <b>Design do Gráfico</b>'),
      el('input[data-cmd="chart-labels"]', 'Marque <b>Rótulos de Dados</b>'),
    ],
    'Linhas para o tempo': [
      drag('A9', 'B15', 'Arraste de <b>A9</b> até <b>B15</b>'),
      btn('chart-insert', 'rec', 'Na guia <b>Inserir</b>, clique em <b>Gráficos Recomendados</b>'),
    ],
    'Trocando o tipo: pizza': [
      clickChart('A1:B6', 'Clique no gráfico das <b>regiões</b>'),
      menu('chart-type', 'Em <b>Design do Gráfico</b>, clique em <b>Alterar Tipo de Gráfico</b>'),
      item('chart-type', 'pie', 'Escolha <b>Pizza</b>'),
    ],
    'O gráfico acompanha os dados': [cell('B3'), type('9800'), enter(), look('.xl-chart[data-range="A1:B6"]', 'A fatia do Nordeste cresceu')],
  });

  // ======================================================================
  // 7.3 Formatação condicional
  // ======================================================================
  const ESTOQUE = [['Caneta', 120, 50, 3100], ['Caderno', 35, 40, 6200], ['Lápis', 200, 60, 1500], ['Borracha', 18, 30, 2800],
    ['Mochila', 22, 10, 8400], ['Régua', 75, 25, 4600], ['Estojo', 9, 15, 7300], ['Agenda', 40, 20, 5000]];
  const ruleOn = (x, kind, range) => x.sheet.cf.find((r) => r.kind === kind && sameRange(r.g, range)) || null;
  const cfAt = (x, ref) => { const p = XL.parseRef(ref); return x.sheet.conditionalStyles().get(p.r * K + p.c) || {}; };
  XL.registerLesson({
    id: 'i7-3',
    title: 'Formatação condicional',
    summary: [
      '<b>Página Inicial › Formatação Condicional</b> pinta as células <b>conforme o valor</b> — e repinta sozinha quando o valor muda.',
      '<b>Realçar Regras</b>: maior que, menor que, entre, igual, texto, duplicados. <b>Primeiros/Últimos</b>: top N, acima da média.',
      '<b>Barras de Dados</b> e <b>Escalas de Cor</b> mostram o tamanho dos números dentro da própria célula.',
      'Com <b>fórmula</b> (<code>=$B2&lt;$C2</code>) você pinta a linha inteira. O <code>$</code> trava a coluna.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Produto', fmt: head }, B1: { v: 'Estoque', fmt: head }, C1: { v: 'Mínimo', fmt: head }, D1: { v: 'Vendas (mês)', fmt: head } };
      ESTOQUE.forEach(([p, e, mn, v], i) => { const r = i + 2; m['A' + r] = p; m['B' + r] = e; m['C' + r] = mn; m['D' + r] = { v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 100, B: 75, C: 75, D: 115 });
    },
    steps: [
      {
        title: 'Cores que trabalham por você',
        text: `<p>Numa lista grande, achar "quem vendeu mais" ou "o que está acabando" exige ler número por número.</p>
          <p>A <b>formatação condicional</b> pinta as células automaticamente de acordo com uma <b>regra</b>. Se o número muda, a cor muda junto.</p>`,
        target: 'cell:A1:D9',
      },
      {
        title: 'Realce as melhores vendas',
        task: 'Selecione <b>D2:D9</b> e use <b>Formatação Condicional › É Maior do que...</b>: valor <code>5000</code>, com <b>Preenchimento Verde e Texto Verde Escuro</b>.',
        hint: 'Página Inicial › grupo Estilos › Formatação Condicional. Na janela, troque o valor sugerido por 5000 e escolha a opção verde na lista.',
        target: ['cell:D2:D9', '[data-menu="cf"]'],
        check: (x) => { const r = ruleOn(x, 'gt', 'D2:D9'); return !!r && x.sheet.cfValue(r.v1) === 5000 && r.style === 'verde'; },
        solve: (x) => { x.app.select('D2:D9'); x.app.addCF({ kind: 'gt', v1: '5000', style: 'verde' }); },
        success: 'Caderno, Mochila e Estojo ficaram verdes. A Agenda (exatamente R$ 5.000) não, porque a regra é <b>maior do que</b>.',
      },
      {
        title: 'A regra é viva',
        task: 'Mude as vendas do <b>Lápis</b> (<b>D4</b>) para <code>5500</code>.',
        target: 'cell:D4',
        check: (x) => x.num('D4') === 5500 && cfAt(x, 'D4').fill === XL.CF_STYLES.verde.fill,
        solve: (x) => commit(x, 'D4', 5500),
        success: 'O Lápis ficou verde sozinho. Você não pinta células: você define a <b>regra</b> e o Excel pinta.',
      },
      {
        title: 'Barras de dados',
        text: '<p>As <b>barras de dados</b> desenham uma barrinha dentro da célula, proporcional ao valor — um mini-gráfico sem sair da tabela.</p>',
        task: 'Selecione <b>B2:B9</b> (Estoque) e use <b>Formatação Condicional › Barras de Dados › Barra Azul</b>.',
        target: ['cell:B2:B9', '[data-menu="cf"]'],
        check: (x) => !!ruleOn(x, 'bar', 'B2:B9'),
        solve: (x) => { x.app.select('B2:B9'); x.app.addCF({ kind: 'bar', color: 'azul' }); },
        success: 'Dá para ver de relance que o Lápis tem o maior estoque e o Estojo, o menor. Experimente também as <b>Escalas de Cor</b> (do verde ao vermelho) em outros dados.',
      },
      {
        title: 'Pinte a linha inteira com uma fórmula',
        text: `<p>Queremos destacar a <b>linha toda</b> dos produtos com estoque <b>abaixo do mínimo</b>. Para isso, a regra é uma fórmula escrita para a <b>primeira linha</b> da seleção:</p>
          <p style="text-align:center"><code>=$B2&lt;$C2</code></p>
          <p>O <code>$</code> trava as colunas B e C: assim, em todas as colunas da linha, a comparação é sempre Estoque × Mínimo. A linha (2) fica livre para mudar.</p>`,
        task: 'Selecione <b>A2:D9</b>, use <b>Formatação Condicional › Nova Regra (usar uma fórmula)</b>, digite <code>=$B2&lt;$C2</code> e mantenha o preenchimento <b>vermelho</b>.',
        hint: 'Se só a coluna A ficou vermelha, faltou o $ antes do B e do C.',
        target: ['cell:A2:D9', '[data-menu="cf"]'],
        check: (x) => {
          if (!x.sheet.cf.some((r) => r.kind === 'formula' && covers(r.g, 'A2:D9'))) return false;
          const red = XL.CF_STYLES.vermelho.fill;
          const low = ['3', '5', '8'], high = ['2', '4', '6', '7', '9'];
          return low.every((r) => cfAt(x, 'A' + r).fill === red && cfAt(x, 'C' + r).fill === red) && high.every((r) => cfAt(x, 'A' + r).fill !== red);
        },
        solve: (x) => { x.app.select('A2:D9'); x.app.addCF({ kind: 'formula', formula: '=$B2<$C2', style: 'vermelho' }); },
        success: 'Caderno, Borracha e Estojo ficaram vermelhos de ponta a ponta: é hora de repor! Para remover regras, use <b>Formatação Condicional › Limpar Regras</b>.',
      },
    ],
  });

  R('i7-3', {
    'Realce as melhores vendas': [
      drag('D2', 'D9', 'Selecione as vendas: arraste de <b>D2</b> até <b>D9</b>'),
      menu('cf', 'Na guia <b>Página Inicial</b>, clique em <b>Formatação Condicional</b>'),
      item('cf', 'gt', 'Em <b>Realçar Regras das Células</b>, escolha <b>É Maior do que...</b>'),
      typeIn('input[data-role="cf-v1"]', '5000', 'Apague o valor sugerido e digite <code>5000</code>'),
      set('select[data-role="cf-style"]', 'verde', 'Na lista "com", escolha <b>Preenchimento Verde e Texto Verde Escuro</b>'),
      ok(),
    ],
    'A regra é viva': [cell('D4'), type('5500'), enter(), look('cell:D4', 'O Lápis passou de 5000 e ficou verde sozinho')],
    'Barras de dados': [
      drag('B2', 'B9', 'Arraste de <b>B2</b> até <b>B9</b>'),
      menu('cf', 'Clique em <b>Formatação Condicional</b>'),
      item('cf', 'bar-azul', 'Em <b>Barras de Dados</b>, escolha <b>Barra Azul</b>'),
    ],
    'Pinte a linha inteira com uma fórmula': [
      drag('A2', 'D9', 'Selecione todos os dados, sem o cabeçalho: de <b>A2</b> até <b>D9</b>'),
      menu('cf', 'Clique em <b>Formatação Condicional</b>'),
      item('cf', 'formula', 'Escolha <b>Nova Regra (usar uma fórmula)...</b>'),
      typeIn('input[data-role="cf-v1"]', '=$B2<$C2', 'Digite <code>=$B2&lt;$C2</code> — "o estoque desta linha é menor que o mínimo?"'),
      look('select[data-role="cf-style"]', 'O preenchimento vermelho já está escolhido'),
      ok(),
    ],
  });

  // ======================================================================
  // 7.4 Validação de dados
  // ======================================================================
  const PEDIDOS = [[101, 'Mercado Bom Preço'], [102, 'Padaria Sol'], [103, 'Loja Central'], [104, 'Farmácia Vida'], [105, 'Café Aroma'], [106, 'Posto Norte']];
  const dvOn = (x, type, range) => x.sheet.dv.find((r) => r.type === type && covers(r.g, range)) || null;
  const lower = (arr) => arr.map((s) => String(s).toLocaleLowerCase('pt-BR'));
  XL.registerLesson({
    id: 'i7-4',
    title: 'Validação de dados e listas suspensas',
    summary: [
      '<b>Dados › Validação de Dados</b> define o que pode ser digitado numa célula.',
      '<b>Lista</b> cria uma lista suspensa: itens separados por <code>;</code> ou um intervalo como <code>=$H$2:$H$5</code>. Abra com a setinha ou <kbd>Alt</kbd>+<kbd>↓</kbd>.',
      '<b>Número inteiro</b>, <b>Decimal</b>, <b>Data</b> e <b>Comprimento do texto</b> limitam valores (entre, maior que...).',
      'A <b>mensagem de entrada</b> orienta antes de digitar; o <b>alerta de erro</b> explica o que está errado.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Pedido', fmt: head }, B1: { v: 'Cliente', fmt: head }, C1: { v: 'Status', fmt: head }, D1: { v: 'Qtd (caixas)', fmt: head } };
      PEDIDOS.forEach(([p, c], i) => { const r = i + 2; m['A' + r] = p; m['B' + r] = c; });
      x.set(m);
      x.widths({ A: 65, B: 150, C: 100, D: 100 });
    },
    steps: [
      {
        title: 'Evitando erros de digitação',
        text: `<p>Se cada pessoa digitar o status do seu jeito ("pago", "Pago.", "PG", "quitado"), nenhum filtro ou CONT.SE vai funcionar direito.</p>
          <p>A <b>validação de dados</b> resolve isso: cria <b>listas suspensas</b> e recusa valores fora da regra.</p>`,
        target: 'cell:C2:D7',
      },
      {
        title: 'Crie uma lista suspensa',
        task: 'Selecione <b>C2:C7</b>, clique em <b>Dados › Validação de Dados</b>, em Permitir escolha <b>Lista</b> e na Fonte digite <code>Pendente;Pago;Enviado</code>. Confirme com <b>OK</b>.',
        hint: 'Os itens da lista são separados por ponto e vírgula, sem espaços.',
        target: ['cell:C2:C7', '.rb-btn[data-cmd="dv-dialog"]'],
        check: (x) => {
          const r = dvOn(x, 'list', 'C2:C7');
          if (!r) return false;
          const vals = lower(x.sheet.listValues(r));
          return ['pendente', 'pago', 'enviado'].every((v) => vals.includes(v));
        },
        solve: (x) => { x.app.select('C2:C7'); x.app.addDV({ type: 'list', op: 'between', source: 'Pendente;Pago;Enviado' }); },
        success: 'Clique em qualquer célula de C2 a C7: aparece uma <b>setinha</b> ao lado dela.',
      },
      {
        title: 'Escolha da lista',
        task: 'Clique em <b>C2</b>, abra a <b>setinha</b> e escolha <b>Pago</b>.',
        hint: 'Pelo teclado: selecione C2 e pressione <kbd>Alt</kbd>+<kbd>↓</kbd>.',
        target: 'cell:C2',
        check: (x) => x.val('C2') === 'Pago',
        solve: (x) => commit(x, 'C2', 'Pago'),
        success: 'Sem digitar nada — e sem erro de digitação. O atalho <kbd>Alt</kbd>+<kbd>↓</kbd> também abre a lista.',
      },
      {
        title: 'Tente burlar a regra',
        task: 'Em <b>C3</b>, digite <code>Cancelado</code> (que não está na lista) e pressione <kbd>Enter</kbd>. Depois, clique em <b>Cancelar</b> no aviso.',
        target: 'cell:C3',
        check: (x) => x.happened('dverror') && x.val('C3') !== 'Cancelado',
        solve: (x) => { x.app.select('C3'); x.app.grid.showDvError(x.sheet.validationAt(2, 2), 2, 2); XL.ui.closeAllModals(); },
        success: 'O Excel <b>recusou</b> o valor. Se "Cancelado" for um status válido, o certo é <b>incluí-lo na Fonte</b> da lista — e não digitar por fora.',
      },
      {
        title: 'Limite de números com alerta próprio',
        text: '<p>Validação não é só lista: dá para exigir um <b>número inteiro</b> dentro de um intervalo e escrever a sua própria mensagem de erro.</p>',
        task: 'Selecione <b>D2:D7</b> e abra <b>Validação de Dados</b>: Permitir <b>Número inteiro</b>, <b>está entre</b> <code>1</code> e <code>100</code>. No Alerta de erro, escreva a mensagem <code>Digite de 1 a 100 caixas.</code> e clique em <b>OK</b>.',
        target: ['cell:D2:D7', '.rb-btn[data-cmd="dv-dialog"]'],
        check: (x) => {
          const r = dvOn(x, 'whole', 'D2:D7');
          return !!r && r.op === 'between' && Number(r.v1) === 1 && Number(r.v2) === 100 && !!r.errorMsg;
        },
        solve: (x) => { x.app.select('D2:D7'); x.app.addDV({ type: 'whole', op: 'between', v1: '1', v2: '100', errorMsg: 'Digite de 1 a 100 caixas.' }); },
      },
      {
        title: 'Teste o alerta',
        task: 'Em <b>D2</b>, digite <code>500</code> e <kbd>Enter</kbd>. Leia a sua mensagem, clique em <b>Repetir</b>, digite <code>50</code> e <kbd>Enter</kbd>.',
        target: 'cell:D2',
        check: (x) => x.happened('dverror') && x.num('D2') === 50,
        solve: (x) => { x.app.select('D2'); x.app.grid.showDvError(x.sheet.validationAt(1, 3), 1, 3); XL.ui.closeAllModals(); commit(x, 'D2', 50); },
        success: `A sua mensagem apareceu no lugar do aviso genérico, e o 50 foi aceito.
          <p>Dicas: a aba de <b>Mensagem de entrada</b> mostra uma orientação assim que a célula é selecionada, e a Fonte da lista pode ser um intervalo (<code>=$H$2:$H$5</code>) — mudou a lista, mudam as opções.</p>`,
      },
    ],
  });

  R('i7-4', {
    'Crie uma lista suspensa': [
      drag('C2', 'C7', 'Selecione as células de status: de <b>C2</b> até <b>C7</b>'),
      btn('dv-dialog', undefined, 'Na guia <b>Dados</b>, clique em <b>Validação de Dados</b>'),
      set('select[data-role="dv-type"]', 'list', 'Em <b>Permitir</b>, escolha <b>Lista</b>'),
      typeIn('input[data-role="dv-source"]', 'Pendente;Pago;Enviado', 'Na <b>Fonte</b>, digite os itens separados por <code>;</code>'),
      ok(),
    ],
    'Escolha da lista': [
      cell('C2'),
      el('.xl-dvbtn', 'Clique na <b>setinha</b> ao lado da célula'),
      el('.ui-menu-item[data-dv="Pago"]', 'Escolha <b>Pago</b>'),
    ],
    'Tente burlar a regra': [
      cell('C3'),
      type('Cancelado', 'Digite <code>Cancelado</code>'),
      enter(),
      look('.dv-modal', 'O Excel mostra um aviso e não aceita o valor'),
      el('[data-role="dv-cancel"]', 'Clique em <b>Cancelar</b> para desistir'),
    ],
    'Limite de números com alerta próprio': [
      drag('D2', 'D7', 'Arraste de <b>D2</b> até <b>D7</b>'),
      btn('dv-dialog', undefined, 'Clique em <b>Dados › Validação de Dados</b>'),
      set('select[data-role="dv-type"]', 'whole', 'Em <b>Permitir</b>, escolha <b>Número inteiro</b>'),
      look('select[data-role="dv-op"]', '"está entre" já vem escolhido'),
      typeIn('input[data-role="dv-min"]', '1', 'Mínimo: <code>1</code>'),
      typeIn('input[data-role="dv-max"]', '100', 'Máximo: <code>100</code>'),
      typeIn('input[data-role="dv-errmsg"]', 'Digite de 1 a 100 caixas.', 'Em <b>Alerta de erro › Mensagem</b>, digite <code>Digite de 1 a 100 caixas.</code>'),
      ok(),
    ],
    'Teste o alerta': [
      cell('D2'),
      type('500'), enter(),
      look('.dv-modal', 'Olha a sua mensagem aí'),
      el('[data-role="dv-retry"]', 'Clique em <b>Repetir</b> — o valor fica selecionado para corrigir'),
      type('50', 'Digite <code>50</code>'),
      enter(),
    ],
  });
})();
