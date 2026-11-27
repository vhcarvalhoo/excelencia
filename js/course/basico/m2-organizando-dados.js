/* Nível Básico · Módulo 2 — Organizando dados (7 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };
  const moeda = { numFmt: 'moeda' };

  // Planilha de vendas usada nas lições de classificar e filtrar.
  const VENDAS = [
    ['Hugo', 'Nordeste', 'Monitor', 3600],
    ['Ana', 'Sudeste', 'Monitor', 2400],
    ['Karen', 'Sudeste', 'Notebook', 14000],
    ['Diego', 'Sul', 'Teclado', 900],
    ['Bruno', 'Nordeste', 'Mouse', 450],
    ['Lucas', 'Centro-Oeste', 'Headset', 840],
    ['Carla', 'Sul', 'Notebook', 7000],
    ['Íris', 'Sul', 'Mouse', 300],
    ['Fábio', 'Sudeste', 'Headset', 1260],
    ['João', 'Norte', 'Teclado', 1800],
    ['Elisa', 'Norte', 'Notebook', 10500],
    ['Gabi', 'Centro-Oeste', 'Webcam', 700],
  ];
  function vendas(x, top) {
    const m = {};
    ['Vendedor', 'Região', 'Produto', 'Valor', 'Comissão'].forEach((t, i) => { m[XL.numToCol(i + 1) + top] = { v: t, fmt: { b: true, fill: '#E2EFDA' } }; });
    VENDAS.forEach((row, i) => {
      const r = top + 1 + i;
      m['A' + r] = row[0]; m['B' + r] = row[1]; m['C' + r] = row[2];
      m['D' + r] = { v: row[3], fmt: moeda };
      m['E' + r] = { v: `=D${r}*5%`, fmt: moeda };
    });
    x.set(m);
    x.widths({ A: 90, B: 110, C: 95, D: 100, E: 100 });
  }
  // Confere se as linhas continuam "inteiras" (o valor ainda é o do vendedor certo).
  const intact = (x, top) => {
    const n = VENDAS.length;
    for (let i = 0; i < n; i++) {
      const r = top + 1 + i;
      const row = VENDAS.find((v) => v[0] === x.val('A' + r));
      if (!row || x.val('D' + r) !== row[3] || x.val('B' + r) !== row[1]) return false;
      if (!x.near(x.num('E' + r), row[3] * 0.05)) return false;
    }
    return true;
  };
  const colVals = (x, col, top) => x.col(col, top + 1, top + VENDAS.length);
  const sortedBy = (arr, desc) => arr.every((v, i) => i === 0 || (desc ? XL.Formula.compareValues(arr[i - 1], v) >= 0 : XL.Formula.compareValues(arr[i - 1], v) <= 0));

  // ======================================================================
  // 2.1 Classificar dados
  // ======================================================================
  XL.registerLesson({
    id: 'b2-1',
    title: 'Classificar dados',
    summary: [
      'Para classificar, clique em <b>uma célula</b> da coluna desejada e use <b>Classificar de A a Z</b> ou <b>de Z a A</b>.',
      'O Excel move a <b>linha inteira</b>: cada vendedor continua com o seu valor e suas fórmulas.',
      'Em <b>Dados › Classificar</b> você combina critérios: primeiro por Região, depois por Valor.',
      'Nunca selecione só uma coluna para classificar — isso embaralharia os dados.',
    ],
    setup: (x) => vendas(x, 1),
    steps: [
      {
        title: 'Colocando ordem na bagunça',
        text: `<p>Esta planilha tem 12 vendas, sem ordem nenhuma. <b>Classificar</b> (ou ordenar) organiza as linhas por uma coluna: nomes em ordem alfabética, valores do maior para o menor...</p>
          <p><b>Regra de ouro:</b> basta clicar em <b>uma única célula</b> da coluna que você quer usar. O Excel reconhece a tabela inteira sozinho e mantém cada linha unida.</p>`,
        target: 'cell:A1:E13',
      },
      {
        title: 'Ordem alfabética',
        task: 'Clique em qualquer nome da coluna <b>Vendedor</b> e use <b>Classificar de A a Z</b> (guia <b>Dados</b>).',
        hint: 'Clique, por exemplo, em A5. Depois vá na guia <b>Dados</b> e clique no botão <b>A→Z</b> (o de cima, no grupo Classificar e Filtrar). Também existe em Página Inicial › Classificar e Filtrar.',
        target: ['cell:A2:A13', '.xl-tab[data-tab="data"]', '.rb-btn[data-cmd="sort"][data-arg="asc"]'],
        check: (x) => x.happened('sort', (d) => d.keys[0].col === 'A' && !d.keys[0].desc) && sortedBy(colVals(x, 'A', 1)) && intact(x, 1),
        solve: (x) => { x.app.select('A5'); x.app.exec('sort', 'asc'); },
        success: 'Ana, Bruno, Carla... E repare: cada pessoa continua com <b>sua</b> região, produto e valor. A <b>Comissão</b> (fórmula) também acompanhou a linha.',
      },
      {
        title: 'Maiores vendas primeiro',
        task: 'Agora classifique pelo <b>Valor</b>, do <b>maior para o menor</b>.',
        hint: 'Clique numa célula da coluna D (Valor) e use o botão <b>Z→A</b> (Classificar do Maior para o Menor).',
        target: ['cell:D2:D13', '.rb-btn[data-cmd="sort"][data-arg="desc"]'],
        check: (x) => x.happened('sort', (d) => d.keys[0].col === 'D' && d.keys[0].desc) && x.num('D2') === 14000 && x.num('D13') === 300 && intact(x, 1),
        solve: (x) => { x.app.select('D3'); x.app.exec('sort', 'desc'); },
        success: 'A Karen lidera com R$ 14.000,00. Para números, A→Z significa "do menor para o maior" e Z→A, "do maior para o menor".',
      },
      {
        title: 'Dois critérios ao mesmo tempo',
        text: `<p>E se você quiser ver as vendas <b>agrupadas por região</b> e, dentro de cada região, <b>da maior para a menor</b>? Para isso existe a janela <b>Classificar</b>, que aceita vários níveis.</p>`,
        task: 'Use <b>Dados › Classificar</b>: classifique por <b>Região</b> (A a Z) e adicione um nível: <b>Valor</b> (do Maior para o Menor).',
        hint: 'Clique numa célula da tabela › guia <b>Dados</b> › botão <b>Classificar</b>. Em "Classificar por" escolha <b>Região</b>. Clique em <b>+ Adicionar Nível</b>, escolha <b>Valor</b> e ordem <b>Do Maior para o Menor</b>. Clique em OK.',
        target: ['.rb-btn[data-cmd="sort-dialog"]', '.sd-add'],
        check: (x) => x.happened('sort', (d) => d.keys.length >= 2 && d.keys[0].col === 'B' && !d.keys[0].desc && d.keys[1].col === 'D' && d.keys[1].desc)
          && x.val('B2') === 'Centro-Oeste' && x.num('D2') === 840 && x.num('D3') === 700 && intact(x, 1),
        solve: (x) => x.app.sortRange(XL.parseRange('A1:E13'), [{ col: 1, desc: false }, { col: 3, desc: true }], true),
        success: 'Centro-Oeste primeiro (Lucas 840, Gabi 700), depois Nordeste, Norte, Sudeste e Sul — cada grupo do maior para o menor. Isso é um relatório pronto para a reunião!',
      },
      {
        title: 'Cuidado com uma armadilha',
        text: `<p>Se você selecionar <b>só a coluna B</b> e classificar, um Excel desatento poderia ordenar apenas aquela coluna, separando cada região do seu vendedor — e os dados ficariam errados sem ninguém perceber.</p>
          <p>Por isso: clique em <b>uma célula</b> e deixe o Excel achar a tabela. Se errar, <kbd>Ctrl</kbd>+<kbd>Z</kbd> desfaz a classificação.</p>`,
      },
    ],
  });

  // ======================================================================
  // 2.2 Filtrar dados
  // ======================================================================
  const T2 = 4; // a tabela começa na linha 4 (as linhas 1–2 têm o resumo)
  XL.registerLesson({
    id: 'b2-2',
    title: 'Filtrar dados',
    summary: [
      '<b>Dados › Filtro</b> (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd>) coloca setinhas nos cabeçalhos.',
      'O filtro <b>esconde</b> linhas, não apaga nada. Números de linha azuis e a barra de status avisam que há filtro.',
      '<code>SOMA</code> soma tudo; <code>SUBTOTAL(9;...)</code> soma <b>só o que está visível</b>.',
      'Filtros de Número/Texto permitem condições como "maior do que 5000" ou "contém".',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Vendas do trimestre', fmt: { b: true, size: 14 } },
        F1: { v: 'Soma (tudo):', fmt: bold }, G1: { v: `=SOMA(D${T2 + 1}:D${T2 + 12})`, fmt: moeda },
        F2: { v: 'Subtotal (visível):', fmt: bold }, G2: { v: null, fmt: Object.assign({ fill: '#FFF2CC' }, moeda) },
      });
      vendas(x, T2);
      x.widths({ F: 125, G: 110 });
    },
    steps: [
      {
        title: 'Mostrar só o que interessa',
        text: `<p>Filtrar é mostrar <b>apenas as linhas</b> que atendem a uma condição — por exemplo, só as vendas da região Sul.</p>
          <p>Importante: o filtro <b>não apaga nada</b>. As outras linhas só ficam escondidas e voltam quando você limpa o filtro.</p>`,
        target: 'cell:A4:E16',
      },
      {
        title: 'Ligue o filtro',
        task: 'Clique em qualquer célula da tabela e ative o <b>Filtro</b> (guia <b>Dados</b>).',
        hint: 'Clique em A6, por exemplo. Na guia <b>Dados</b>, clique no funil <b>Filtro</b>. Atalho: <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd>.',
        target: ['.xl-tab[data-tab="data"]', '.rb-btn[data-cmd="filter-toggle"]'],
        check: (x) => !!x.sheet.filter && x.sheet.filter.g.r1 === T2 - 1,
        solve: (x) => { x.app.select('A6'); x.app.exec('filter-toggle'); },
        success: 'Apareceram <b>setinhas</b> em cada cabeçalho. É por elas que você filtra e classifica.',
      },
      {
        title: 'Só a região Sul',
        task: 'Use a setinha de <b>Região</b> para mostrar <b>apenas "Sul"</b>.',
        hint: 'Clique na setinha de B4. Desmarque <b>(Selecionar Tudo)</b>, marque só <b>Sul</b> e clique em <b>OK</b>. Dica: também dá para digitar "Sul" na caixa Pesquisar.',
        target: () => ['.xl-filterbtn[data-c="1"]', '.xl-fitem[data-value="Sul"]', '.xl-fmenu-ok'],
        check: (x) => { const c = x.sheet.filter && x.sheet.filter.crit[1]; return !!c && c.type === 'values' && c.values.length === 1 && c.values[0] === 'Sul' && x.sheet.filterStats().shown === 3; },
        solve: (x) => x.app.setFilterCriteria(1, { type: 'values', values: ['Sul'], blanks: false }),
        success: 'Só 3 vendas aparecem. Repare nos <b>números de linha em azul</b> (pulando linhas), no <b>funil</b> na setinha de Região e na barra de status: <b>3 de 12 registros encontrados</b>.',
      },
      {
        title: 'SOMA x SUBTOTAL',
        text: `<p>Olhe a célula <b>G1</b>: <code>=SOMA(D5:D16)</code> continua somando <b>todas</b> as 12 vendas, mesmo as escondidas.</p>
          <p>Para somar só o que aparece na tela, use <code>SUBTOTAL</code>. O primeiro argumento diz a operação: <b>9 = soma</b> (1 = média, 2 = contar números, 4 = máximo, 5 = mínimo).</p>`,
        task: 'Em <b>G2</b>, digite <code>=SUBTOTAL(9;D5:D16)</code>.',
        hint: 'Clique em G2 (a célula amarela) e digite a fórmula. Compare o resultado com G1.',
        target: ['cell:G1:G2'],
        check: (x) => x.uses('G2', 'SUBTOTAL') && x.near(x.num('G2'), 8200),
        solve: (x) => { x.app.setCell('G2', '=SUBTOTAL(9;D5:D16)'); x.app.select('G3'); },
        success: 'SOMA: <b>R$ 43.750,00</b> (tudo). SUBTOTAL: <b>R$ 8.200,00</b> (só o Sul). Dica: com um filtro ativo, o botão <b>Σ AutoSoma</b> já escreve SUBTOTAL para você!',
      },
      {
        title: 'Limpe o filtro da Região',
        task: 'Abra a setinha de <b>Região</b> e clique em <b>Limpar Filtro de “Região”</b>.',
        hint: 'É a terceira opção do menu da setinha. Todas as 12 linhas voltam.',
        target: ['.xl-filterbtn[data-c="1"]', '.ui-menu-item[data-cmd="fclear"]'],
        check: (x) => !!x.sheet.filter && !x.sheet.filter.crit[1] && x.sheet.filterStats().shown === 12,
        solve: (x) => x.app.setFilterCriteria(1, null),
      },
      {
        title: 'Filtro de número',
        text: '<p>Além de marcar valores, dá para filtrar por <b>condição</b>: maior que, menor que, entre... (em colunas de texto: contém, começa com...).</p>',
        task: 'Na coluna <b>Valor</b>, mostre só as vendas <b>maiores do que 5000</b>.',
        hint: 'Setinha de D4 › <b>Filtros de Número</b> › <b>É Maior do que...</b> › digite <b>5000</b> › OK.',
        target: ['.xl-filterbtn[data-c="3"]', '.ui-menu-item[data-cmd="fcondmenu"]', '.ui-menu-item[data-cmd="fcond"][data-arg=">"]'],
        check: (x) => { const c = x.sheet.filter && x.sheet.filter.crit[3]; return !!c && x.sheet.filterStats().shown === 3 && x.near(x.num('G2'), 31500); },
        solve: (x) => x.app.setFilterCriteria(3, { type: 'cond', op: '>', value: '5000' }),
        success: 'Karen, Elisa e Carla: as três grandes vendas, somando <b>R$ 31.500,00</b> no SUBTOTAL.',
      },
      {
        title: 'Desligue o filtro',
        task: 'Clique de novo em <b>Dados › Filtro</b> para desligar o filtro.',
        target: '.rb-btn[data-cmd="filter-toggle"]',
        check: (x) => !x.sheet.filter && x.hiddenCount() === 0,
        solve: (x) => { x.app.select('A6'); x.app.exec('filter-toggle'); },
        success: 'As setinhas sumiram e todas as linhas voltaram. Nenhum dado foi perdido. 🎯',
      },
    ],
  });

  // ======================================================================
  // 2.3 Inserir e excluir linhas e colunas
  // ======================================================================
  XL.registerLesson({
    id: 'b2-3',
    title: 'Inserir e excluir linhas e colunas',
    summary: [
      'Clique com o <b>botão direito</b> no número da linha (ou letra da coluna) › <b>Inserir</b> ou <b>Excluir</b>.',
      'A linha nova entra <b>acima</b> da selecionada; a coluna nova entra <b>à esquerda</b>.',
      'Fórmulas se ajustam sozinhas: <code>=SOMA(B2:B6)</code> vira <code>=SOMA(B2:B7)</code> quando você insere dentro do intervalo.',
      'Atalhos: <kbd>Ctrl</kbd>+<kbd>+</kbd> insere e <kbd>Ctrl</kbd>+<kbd>-</kbd> exclui (com a linha/coluna selecionada).',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Mês', fmt: bold }, B1: { v: 'Despesas', fmt: bold }, C1: { v: 'Rascunho', fmt: { b: true, color: '#9C0006' } },
        A2: 'Jan', B2: { v: 1200, fmt: moeda }, A3: 'Fev', B3: { v: 1350, fmt: moeda },
        A4: 'Abr', B4: { v: 1100, fmt: moeda }, A5: 'Mai', B5: { v: 1500, fmt: moeda },
        A6: { v: 'Teste', fmt: { color: '#9C0006' } }, B6: { v: 999, fmt: Object.assign({ color: '#9C0006' }, moeda) },
        A7: { v: 'Total', fmt: bold }, B7: { v: '=SOMA(B2:B6)', fmt: Object.assign({ b: true }, moeda) },
        C2: 'xxx', C3: 'apagar', C4: '???', C5: 'xxx',
      });
      x.widths({ B: 105, C: 90 });
    },
    steps: [
      {
        title: 'Esqueceram de março!',
        text: `<p>A lista pula de <b>Fev</b> para <b>Abr</b>. Precisamos abrir espaço para Março entre as linhas 3 e 4 — sem apagar nada e sem refazer o total.</p>
          <p>Ao inserir uma linha, ela entra <b>acima</b> da linha selecionada.</p>`,
        target: 'cell:A2:B7',
      },
      {
        title: 'Inserir uma linha',
        task: 'Clique com o <b>botão direito</b> no número da linha <b>4</b> e escolha <b>Inserir</b>.',
        hint: 'O número 4 fica na coluna cinza da esquerda. Botão direito nele › Inserir. Outra forma: selecione a linha 4 e use Página Inicial › Inserir › Inserir Linhas na Planilha.',
        target: ['.xl-rowh[data-r="3"]', '.ui-menu-item[data-cmd="insert-rows"]'],
        check: (x) => x.happened('insert', (d) => d.axis === 'row' && d.at === 3) && x.text('A5') === 'ABR',
        solve: (x) => { x.app.grid.setSel(3, 0, 3, x.app.sheet.cols - 1); x.app.exec('insert-rows'); },
        success: 'Uma linha vazia apareceu e tudo de baixo desceu uma posição.',
      },
      {
        title: 'Preencha março',
        task: 'Em <b>A4</b> digite <b>Mar</b> e em <b>B4</b> digite <b>1280</b>.',
        target: 'cell:A4:B4',
        check: (x) => x.text('A4') === 'MAR' && x.num('B4') === 1280,
        solve: (x) => { x.app.setCell('A4', 'Mar'); x.app.setCell('B4', '1280'); x.app.select('A5'); },
      },
      {
        title: 'O total se ajustou sozinho',
        task: 'Clique no total (<b>B8</b>) e olhe a fórmula na Barra de Fórmulas.',
        target: ['cell:B8', '.xl-finput'],
        check: (x) => x.active() === 'B8' && x.happened('select') && x.formula('B8') === '=SOMA(B2:B7)',
        solve: (x) => x.app.select('B8'),
        success: 'A fórmula virou <code>=SOMA(B2:B7)</code>: como a linha foi inserida <b>dentro</b> do intervalo, o Excel o ampliou e março já entrou no total.',
      },
      {
        title: 'Excluir uma linha',
        task: 'A linha <b>Teste</b> (linha 7) não deveria estar aí. Clique com o botão direito no número <b>7</b> e escolha <b>Excluir</b>.',
        hint: 'Atenção: <b>Excluir</b> remove a linha e puxa as de baixo para cima. Já <kbd>Delete</kbd> só apaga o conteúdo e deixa a linha vazia.',
        target: ['.xl-rowh[data-r="6"]', '.ui-menu-item[data-cmd="delete-rows"]'],
        check: (x) => x.happened('delete', (d) => d.axis === 'row') && !x.col('A', 1, 10).includes('Teste') && x.formula('B7') === '=SOMA(B2:B6)' && x.near(x.num('B7'), 6430),
        solve: (x) => { x.app.grid.setSel(6, 0, 6, x.app.sheet.cols - 1); x.app.exec('delete-rows'); },
        success: 'O "Teste" sumiu, o Total subiu para a linha 7 e a fórmula encolheu para <code>=SOMA(B2:B6)</code>. Total correto: <b>R$ 6.430,00</b>.',
      },
      {
        title: 'Excluir uma coluna',
        task: 'A coluna <b>C</b> (Rascunho) também é lixo. Clique com o botão direito na letra <b>C</b> e escolha <b>Excluir</b>.',
        target: ['.xl-colh[data-c="2"]', '.ui-menu-item[data-cmd="delete-cols"]'],
        check: (x) => x.happened('delete', (d) => d.axis === 'col' && d.at === 2) && x.empty('C1') && x.empty('C2'),
        solve: (x) => { x.app.grid.setSel(0, 2, x.app.sheet.rows - 1, 2); x.app.exec('delete-cols'); },
        success: 'Planilha limpa! Para inserir uma coluna é igual: botão direito na letra › Inserir (a nova entra à esquerda).',
      },
      {
        title: 'Cuidado com o #REF!',
        text: `<p>Se você excluir uma linha ou coluna que uma fórmula usa <b>diretamente</b> (por exemplo <code>=B3*2</code> e você exclui a linha 3), a fórmula vira <b>#REF!</b> — a referência deixou de existir.</p>
          <p>Fórmulas com <b>intervalos</b> (<code>=SOMA(B2:B6)</code>) são mais seguras: o intervalo só encolhe. Se acontecer um #REF!, <kbd>Ctrl</kbd>+<kbd>Z</kbd> desfaz.</p>`,
      },
    ],
  });

  // ======================================================================
  // 2.4 Congelar painéis
  // ======================================================================
  // rola a grade e avisa o guia (o evento de rolagem do navegador pode demorar um quadro)
  const scrollTo = (x, top, left) => {
    const sc = x.app.grid.scroller;
    if (top !== null) sc.scrollTop = top;
    if (left !== null) sc.scrollLeft = left;
    x.app.emit('scroll', { top: sc.scrollTop, left: sc.scrollLeft });
  };
  const MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  const CLIENTES = ['Padaria Sol', 'Mercado Bom Preço', 'Farmácia Vida', 'Loja Estrela', 'Café Central', 'Pet Shop Amigo', 'Ótica Visão', 'Livraria Saber'];
  XL.registerLesson({
    id: 'b2-4',
    title: 'Congelar painéis',
    summary: [
      '<b>Exibir › Congelar Painéis › Congelar Linha Superior</b> mantém o cabeçalho sempre visível.',
      'Para travar linhas <b>e</b> colunas, selecione a célula logo abaixo e à direita (ex.: <b>B2</b>) › Congelar Painéis.',
      'Para mudar o congelamento, primeiro use <b>Descongelar Painéis</b>.',
    ],
    setup: (x) => {
      // 24 meses (2 anos) para a tabela ficar larga em qualquer tela
      const MESES24 = [].concat(MESES.map((m) => m + '/25'), MESES.map((m) => m + '/26'));
      const m = { A1: { v: 'Cliente', fmt: { b: true, fill: '#DDEBF7' } } };
      MESES24.forEach((mes, j) => { m[XL.numToCol(j + 2) + '1'] = { v: "'" + mes, fmt: { b: true, fill: '#DDEBF7', align: 'center' } }; });
      for (let i = 0; i < 60; i++) {
        const r = i + 2;
        m['A' + r] = `${CLIENTES[i % CLIENTES.length]} ${Math.floor(i / CLIENTES.length) + 1}`;
        MESES24.forEach((_, j) => { m[XL.numToCol(j + 2) + r] = 200 + ((i * 37 + j * 53) % 900); });
      }
      x.set(m);
      x.widths({ A: 150 });
    },
    steps: [
      {
        title: 'Planilhas grandes',
        text: `<p>Esta planilha tem <b>60 clientes</b> e <b>24 meses</b> de compras. Quando você rola para baixo, o cabeçalho (os meses) some — e fica difícil saber o que cada número significa.</p>`,
        task: 'Role a planilha <b>para baixo</b> até a linha 40, mais ou menos.',
        hint: 'Use a rodinha do mouse sobre a grade, ou a barra de rolagem à direita.',
        target: '.xl-grid-host',
        check: (x) => x.happened('scroll', (d) => d.top > 350),
        solve: (x) => scrollTo(x, 700, null),
        success: 'Viu? Os meses sumiram lá em cima. Vamos resolver isso.',
      },
      {
        title: 'Congele a linha do cabeçalho',
        task: 'Na guia <b>Exibir</b>, clique em <b>Congelar Painéis › Congelar Linha Superior</b>.',
        target: ['.xl-tab[data-tab="view"]', '[data-menu="freeze"]', '.ui-menu-item[data-arg="top"]'],
        check: (x) => !!x.sheet.freeze && x.sheet.freeze.rows === 1 && x.sheet.freeze.cols === 0,
        solve: (x) => x.app.exec('freeze', 'top'),
        success: 'Apareceu uma linha mais escura embaixo da linha 1: é o limite do congelamento.',
      },
      {
        title: 'Teste rolando',
        task: 'Role para baixo de novo e veja o cabeçalho.',
        target: '.xl-grid-host',
        check: (x) => x.happened('scroll', (d) => d.top > 350),
        solve: (x) => scrollTo(x, 700, null),
        success: 'Agora os meses ficam <b>fixos no topo</b> enquanto os clientes passam por baixo. 🧊',
      },
      {
        title: 'E se rolar para o lado?',
        text: '<p>Rolando para a direita (para ver os meses de 2026), o que some é a coluna dos <b>clientes</b>. Dá para congelar a linha 1 <b>e</b> a coluna A ao mesmo tempo.</p><p>Para mudar o congelamento, primeiro é preciso <b>descongelar</b>.</p>',
        task: 'Use <b>Exibir › Congelar Painéis › Descongelar Painéis</b>.',
        target: ['.xl-tab[data-tab="view"]', '[data-menu="freeze"]', '.ui-menu-item[data-arg="none"]'],
        check: (x) => !x.sheet.freeze,
        solve: (x) => x.app.exec('freeze', 'none'),
      },
      {
        title: 'Congele linha e coluna',
        text: '<p>O truque: selecione a célula <b>logo abaixo</b> das linhas e <b>logo à direita</b> das colunas que você quer travar. Para travar a linha 1 e a coluna A, a célula é <b>B2</b>.</p>',
        task: 'Selecione <b>B2</b> e clique em <b>Congelar Painéis › Congelar Painéis</b>.',
        hint: 'Clique em B2 (talvez precise rolar para cima). Depois Exibir › Congelar Painéis › a primeira opção.',
        target: ['cell:B2', '[data-menu="freeze"]', '.ui-menu-item[data-arg="panes"]'],
        check: (x) => !!x.sheet.freeze && x.sheet.freeze.rows === 1 && x.sheet.freeze.cols === 1,
        solve: (x) => { x.app.select('B2'); x.app.exec('freeze', 'panes'); },
      },
      {
        title: 'Role para a direita',
        task: 'Role a planilha <b>para a direita</b> até ver <b>Dez/26</b>.',
        hint: 'Use a barra de rolagem horizontal embaixo da grade, ou <kbd>Shift</kbd> + rodinha do mouse.',
        target: '.xl-grid-host',
        check: (x) => x.happened('scroll', (d) => d.left > 150),
        solve: (x) => scrollTo(x, null, 1400),
        success: 'Os nomes dos clientes continuam à esquerda e os meses no topo. Agora dá para ler qualquer número sem se perder! 🏆',
      },
    ],
  });

  // ======================================================================
  // 2.5 Várias planilhas
  // ======================================================================
  const gastos = (vals) => ({
    A1: { v: 'Despesa', fmt: bold }, B1: { v: 'Valor', fmt: bold },
    A2: 'Aluguel', B2: { v: vals[0], fmt: moeda }, A3: 'Mercado', B3: { v: vals[1], fmt: moeda },
    A4: 'Transporte', B4: { v: vals[2], fmt: moeda }, A5: 'Lazer', B5: { v: vals[3], fmt: moeda },
    A6: { v: 'Total', fmt: bold }, B6: { v: '=SOMA(B2:B5)', fmt: Object.assign({ b: true }, moeda) },
  });
  const resumoLabels = (x) => {
    if (!x.hasSheet('Resumo')) return;
    x.app.setCells({
      A1: { v: 'Mês', fmt: bold }, B1: { v: 'Total', fmt: bold },
      A2: 'Janeiro', A3: 'Fevereiro', A4: { v: 'Bimestre', fmt: bold },
      B2: { v: null, fmt: moeda }, B3: { v: null, fmt: moeda }, B4: { v: null, fmt: Object.assign({ b: true }, moeda) },
    }, 'Resumo');
  };
  XL.registerLesson({
    id: 'b2-5',
    title: 'Várias planilhas e referências entre elas',
    summary: [
      'Cada <b>aba</b> lá embaixo é uma planilha. O <b>+</b> cria uma nova; <b>duplo clique</b> renomeia.',
      'Para usar uma célula de outra planilha: <code>=Janeiro!B6</code> (nome da planilha, <b>!</b> e a célula).',
      'Jeito fácil: digite <b>=</b>, clique na aba e na célula, e pressione <kbd>Enter</kbd>.',
      'Renomeou a planilha? As fórmulas se atualizam sozinhas.',
    ],
    setup: (x) => {
      const wb = x.app.wb;
      wb.renameSheet(wb.sheets[0], 'Janeiro');
      x.app.setCells(gastos([1500, 900, 300, 250]));
      x.app.setCells(gastos([1500, 980, 340, 180]), 'Fevereiro');
      x.app.setColWidths({ A: 110, B: 110 });
      x.app.setColWidths({ A: 110, B: 110 }, 'Fevereiro');
      x.app.showSheet(0);
    },
    steps: [
      {
        title: 'Uma pasta, várias planilhas',
        text: `<p>Um arquivo do Excel (a <b>pasta de trabalho</b>) pode ter várias <b>planilhas</b>. Elas aparecem como <b>abas</b> no rodapé — aqui temos <b>Janeiro</b> e <b>Fevereiro</b>.</p>
          <p>É ótimo para separar meses, filiais, dados e relatórios.</p>`,
        target: '.xl-sheettabs',
      },
      {
        title: 'Troque de planilha',
        task: 'Clique na aba <b>Fevereiro</b>.',
        hint: 'As abas ficam no rodapé, à esquerda. Atalho: <kbd>Ctrl</kbd>+<kbd>PageDown</kbd>.',
        target: '.xl-sheettab[data-sheet="Fevereiro"]',
        check: (x) => x.sheet.name === 'Fevereiro',
        solve: (x) => x.app.showSheet(1),
        success: 'Mesma estrutura, valores de fevereiro. Total: R$ 3.000,00.',
      },
      {
        title: 'Crie uma planilha nova',
        task: 'Clique no <b>+</b> ao lado das abas para criar uma planilha.',
        target: '.xl-sheetadd',
        check: (x) => x.sheetNames().length >= 3 && x.happened('sheetadd'),
        solve: (x) => x.app.exec('sheet-add'),
        success: 'Nasceu uma planilha em branco com um nome genérico.',
      },
      {
        title: 'Dê um nome de verdade',
        task: 'Dê um <b>duplo clique</b> na aba nova, digite <b>Resumo</b> e pressione <kbd>Enter</kbd>.',
        hint: 'Também funciona com o botão direito na aba › Renomear.',
        target: () => ['.xl-sheettab.active', '.xl-tabinput'],
        check: (x) => x.hasSheet('Resumo'),
        solve: (x) => { if (!x.hasSheet('Resumo')) { const s = x.app.wb.sheets.find((s2) => !['Janeiro', 'Fevereiro'].includes(s2.name)) || x.app.cmdSheetAdd(); x.app.renameSheet(s, 'Resumo'); } },
        success: 'Nomes claros (Janeiro, Fevereiro, Resumo) deixam a pasta fácil de entender.',
      },
      {
        title: 'Buscando dados de outra planilha',
        setup: (x) => { resumoLabels(x); if (x.hasSheet('Resumo')) x.app.showSheet(x.app.wb.byName('Resumo')); },
        text: `<p>Preparamos os rótulos do Resumo. Agora vamos trazer o total de janeiro <b>sem copiar o número</b> — com uma fórmula que aponta para a outra planilha: <code>=Janeiro!B6</code>.</p>`,
        task: 'Em <b>B2</b> do Resumo, digite <b>=</b>, clique na aba <b>Janeiro</b>, clique em <b>B6</b> e pressione <kbd>Enter</kbd>.',
        hint: 'Depois do <b>=</b>, ao clicar na aba Janeiro a fórmula continua sendo escrita na Barra de Fórmulas. Clique em B6 (o total) e tecle <kbd>Enter</kbd>: você volta sozinho para o Resumo.',
        target: ['cell:Resumo!B2', 'cell:Janeiro!B6', '.xl-sheettab[data-sheet="Janeiro"]'],
        check: (x) => x.formula('Resumo!B2') === '=JANEIRO!B6' && x.near(x.num('Resumo!B2'), 2950),
        solve: (x) => { const s = x.app.wb.byName('Resumo'); x.app.showSheet(s); x.app.setCell('B2', '=Janeiro!B6'); },
        success: 'A fórmula ficou <code>=Janeiro!B6</code>: nome da planilha, ponto de exclamação e a célula.',
      },
      {
        title: 'Agora fevereiro',
        task: 'Em <b>B3</b> do Resumo, traga o total de fevereiro: <code>=Fevereiro!B6</code>.',
        hint: 'Mesmo caminho: = › aba Fevereiro › B6 › Enter. Ou digite a fórmula direto.',
        target: ['cell:Resumo!B3', 'cell:Fevereiro!B6', '.xl-sheettab[data-sheet="Fevereiro"]'],
        check: (x) => x.formula('Resumo!B3') === '=FEVEREIRO!B6' && x.near(x.num('Resumo!B3'), 3000),
        solve: (x) => { x.app.showSheet(x.app.wb.byName('Resumo')); x.app.setCell('B3', '=Fevereiro!B6'); },
      },
      {
        title: 'Total do bimestre',
        task: 'Em <b>B4</b> do Resumo, some os dois meses: <code>=SOMA(B2:B3)</code>.',
        target: 'cell:Resumo!B4',
        check: (x) => x.uses('Resumo!B4', 'SOMA') && x.near(x.num('Resumo!B4'), 5950),
        solve: (x) => { x.app.showSheet(x.app.wb.byName('Resumo')); x.app.setCell('B4', '=SOMA(B2:B3)'); },
      },
      {
        title: 'Tudo conectado',
        task: 'Vá na planilha <b>Janeiro</b>, mude o <b>Lazer</b> (B5) para <b>400</b> e volte ao <b>Resumo</b>.',
        hint: 'Aba Janeiro › clique em B5 › digite 400 › Enter › aba Resumo.',
        target: ['cell:Janeiro!B5', '.xl-sheettab[data-sheet="Janeiro"]', '.xl-sheettab[data-sheet="Resumo"]'],
        check: (x) => x.num('Janeiro!B5') === 400 && x.near(x.num('Resumo!B4'), 6100) && x.sheet.name === 'Resumo',
        solve: (x) => { x.app.showSheet(0); x.app.setCell('B5', '400'); x.app.showSheet(x.app.wb.byName('Resumo')); },
        success: 'O Resumo foi para <b>R$ 6.100,00</b> sozinho. Uma alteração na origem atualiza todo o relatório.',
      },
      {
        title: 'Renomear não quebra nada',
        task: 'Renomeie a planilha <b>Janeiro</b> para <b>Jan</b> e confira a fórmula de <b>B2</b> no Resumo.',
        hint: 'Duplo clique na aba Janeiro › digite Jan › Enter. Depois clique em B2 do Resumo e olhe a Barra de Fórmulas.',
        target: '.xl-sheettab[data-sheet="Janeiro"]',
        check: (x) => x.hasSheet('Jan') && x.formula('Resumo!B2') === '=JAN!B6',
        solve: (x) => x.app.renameSheet(x.app.wb.byName('Janeiro'), 'Jan'),
        success: 'A fórmula virou <code>=Jan!B6</code> automaticamente. Dica: se o nome tiver espaço, o Excel usa aspas simples: <code>=\'Vendas 2025\'!B6</code>.',
      },
    ],
  });

  // ======================================================================
  // 2.6 Localizar, substituir e remover duplicatas
  // ======================================================================
  const CLIENTES6 = [
    ['Mariana Costa', 'S. Paulo', 120], ['Pedro Alves', 'Rio de Janeiro', 80], ['Carla Souza', 'Belo Horizonte', 200],
    ['Mariana Costa', 'S. Paulo', 120], ['João Lima', 'Curitiba', 95], ['Ana Ribeiro', 'S. Paulo', 310],
    ['Pedro Alves', 'Rio de Janeiro', 80], ['Lucas Rocha', 'Salvador', 150], ['Beatriz Nunes', 'S. Paulo', 60],
    ['Rafael Dias', 'Recife', 220], ['Carla Souza', 'Belo Horizonte', 200], ['Fernanda Melo', 'Porto Alegre', 175],
  ];
  XL.registerLesson({
    id: 'b2-6',
    title: 'Localizar, substituir e remover duplicatas',
    summary: [
      '<kbd>Ctrl</kbd>+<kbd>L</kbd> abre o <b>Localizar</b>; <kbd>Ctrl</kbd>+<kbd>U</kbd> abre o <b>Substituir</b>.',
      '<b>Substituir tudo</b> corrige um erro repetido na planilha inteira de uma vez.',
      '<b>Dados › Remover Duplicatas</b> apaga linhas repetidas, mantendo a primeira ocorrência.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Cliente', fmt: bold }, B1: { v: 'Cidade', fmt: bold }, C1: { v: 'Compras', fmt: bold } };
      CLIENTES6.forEach((c, i) => { m['A' + (i + 2)] = c[0]; m['B' + (i + 2)] = c[1]; m['C' + (i + 2)] = { v: c[2], fmt: moeda }; });
      x.set(m);
      x.widths({ A: 130, B: 125, C: 95 });
    },
    steps: [
      {
        title: 'Uma lista com problemas',
        text: `<p>Esta lista de clientes foi juntada de vários lugares e tem dois problemas comuns:</p>
          <ul><li>A cidade aparece como <b>"S. Paulo"</b> em vez de <b>"São Paulo"</b>;</li><li>Alguns clientes aparecem <b>repetidos</b>.</li></ul>
          <p>Em listas de milhares de linhas, corrigir na mão é impossível. Vamos usar as ferramentas certas.</p>`,
        target: 'cell:A1:C13',
      },
      {
        title: 'Localizar',
        task: 'Abra o <b>Localizar</b> (<kbd>Ctrl</kbd>+<kbd>L</kbd>), procure <b>Rafael</b> e clique em <b>Localizar próxima</b>.',
        hint: 'Também está em Página Inicial › <b>Localizar e Selecionar</b> › Localizar... Digite Rafael na caixa e clique em <b>Localizar próxima</b> (ou Enter).',
        target: ['[data-menu="findselect"]', '.xl-find input[data-role="find"]', '.xl-find button[data-role="findnext"]'],
        check: (x) => x.happened('find', (d) => d.found && d.ref === 'A11'),
        solve: (x) => { XL.Find.open(x.app, 'find'); XL.Find.findInput.value = 'Rafael'; XL.Find.findNext(); },
        success: 'O Excel pulou direto para <b>A11</b>. "Localizar tudo" lista todas as ocorrências de uma vez.',
      },
      {
        title: 'Substituir tudo',
        task: 'Na aba <b>Substituir</b>, troque <b>S. Paulo</b> por <b>São Paulo</b> e clique em <b>Substituir tudo</b>.',
        hint: 'Clique na aba "Substituir" da janelinha (ou <kbd>Ctrl</kbd>+<kbd>U</kbd>). Localizar: <b>S. Paulo</b> · Substituir por: <b>São Paulo</b> · <b>Substituir tudo</b>.',
        target: ['.xl-find .fr-tab[data-tab="replace"]', '.xl-find input[data-role="replace"]', '.xl-find button[data-role="replaceall"]'],
        check: (x) => x.happened('replace', (d) => d.count >= 4) && !x.col('B', 2, 13).includes('S. Paulo') && x.col('B', 2, 13).filter((v) => v === 'São Paulo').length === 4,
        solve: (x) => { XL.Find.open(x.app, 'replace'); XL.Find.findInput.value = 'S. Paulo'; XL.Find.replInput.value = 'São Paulo'; XL.Find.replaceAll(); },
        success: '"Concluído. Fizemos 4 substituições." Em um clique, a lista inteira ficou padronizada.',
      },
      {
        title: 'Remover duplicatas',
        setup: () => { if (XL.Find) XL.Find.close(); },
        text: '<p>Agora os repetidos: Mariana, Pedro e Carla aparecem duas vezes cada.</p>',
        task: 'Clique em uma célula da lista e use <b>Dados › Remover Duplicatas</b> › <b>OK</b>.',
        hint: 'Na janela, deixe todas as colunas marcadas (assim só linhas totalmente iguais são removidas) e "Meus dados contêm cabeçalhos" ligado.',
        target: ['.xl-tab[data-tab="data"]', '.rb-btn[data-cmd="dedupe"]'],
        check: (x) => x.happened('dedupe', (d) => d.removed === 3) && x.empty('A11') && !x.empty('A10'),
        solve: (x) => { x.app.select('A2'); x.app.removeDuplicates(XL.parseRange('A1:C13'), [0, 1, 2], true); },
        success: '3 duplicatas removidas, 9 clientes únicos permanecem. A primeira ocorrência de cada um foi mantida. 🧹',
      },
    ],
  });

  // ======================================================================
  // 2.7 Impressão e layout da página
  // ======================================================================
  XL.registerLesson({
    id: 'b2-7',
    title: 'Impressão e layout da página',
    summary: [
      '<b>Arquivo › Imprimir</b> (<kbd>Ctrl</kbd>+<kbd>P</kbd>) mostra exatamente como as páginas vão sair.',
      'Tabelas largas ficam melhores em <b>Paisagem</b> e com <b>Ajustar Todas as Colunas em Uma Página</b>.',
      '<b>Imprimir Títulos</b> repete o cabeçalho no topo de todas as páginas.',
      '<b>Área de Impressão</b> imprime só o pedaço que você escolher.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Código', fmt: { b: true, fill: '#E2EFDA' } }, B1: { v: 'Produto', fmt: { b: true, fill: '#E2EFDA' } }, C1: { v: 'Categoria', fmt: { b: true, fill: '#E2EFDA' } } };
      MESES.forEach((mes, j) => { m[XL.numToCol(j + 4) + '1'] = { v: mes, fmt: { b: true, fill: '#E2EFDA', align: 'center' } }; });
      const cats = ['Bebidas', 'Mercearia', 'Limpeza', 'Higiene', 'Padaria'];
      for (let i = 0; i < 60; i++) {
        const r = i + 2;
        m['A' + r] = 'P' + String(101 + i);
        m['B' + r] = 'Produto ' + (i + 1);
        m['C' + r] = cats[i % cats.length];
        MESES.forEach((_, j) => { m[XL.numToCol(j + 4) + r] = 50 + ((i * 29 + j * 41) % 450); });
      }
      x.set(m);
      x.widths({ A: 70, B: 110, C: 95 });
    },
    steps: [
      {
        title: 'Antes de imprimir, visualize',
        text: `<p>Imprimir direto costuma dar errado: colunas cortadas, cabeçalho só na primeira folha, páginas demais...</p>
          <p>O hábito profissional é sempre conferir a <b>Visualização de Impressão</b> primeiro.</p>`,
        task: 'Abra a visualização: guia <b>Arquivo</b> (ou <kbd>Ctrl</kbd>+<kbd>P</kbd>).',
        hint: 'Clique em <b>Arquivo</b>, a primeira guia da Faixa de Opções.',
        target: '.xl-tab[data-tab="file"]',
        check: (x) => x.happened('printpreview'),
        solve: (x) => x.app.exec('print-preview'),
        success: 'Veja o contador de páginas: a tabela não cabe na largura, então as colunas dos últimos meses foram parar em <b>outras páginas</b>. Use ◀ ▶ para conferir.',
      },
      {
        title: 'Deite a página',
        task: 'Nas configurações à esquerda, mude para <b>Orientação Paisagem</b>.',
        target: '.pp-overlay select[data-role="orientation"]',
        check: (x) => x.sheet.page.orientation === 'paisagem',
        solve: (x) => { if (!XL.Print.current) x.app.exec('print-preview'); x.app.setPage({ orientation: 'paisagem' }); },
        success: 'Mais largura por folha. Mas ainda pode sobrar coluna...',
      },
      {
        title: 'Todas as colunas em uma página',
        task: 'Em <b>Dimensionamento</b>, escolha <b>Ajustar Todas as Colunas em Uma Página</b>.',
        target: '.pp-overlay select[data-role="fit"]',
        check: (x) => ['cols', 'sheet'].includes(x.sheet.page.fit),
        solve: (x) => { if (!XL.Print.current) x.app.exec('print-preview'); x.app.setPage({ fit: 'cols' }); },
        success: 'O Excel reduziu a escala para os 12 meses caberem na largura. As linhas continuam em várias páginas — o que é normal para listas longas.',
      },
      {
        title: 'Volte para a planilha',
        task: 'Clique em <b>← Voltar à planilha</b> (ou pressione <kbd>Esc</kbd>).',
        target: '.pp-back',
        check: (x) => x.happened('printclose') || !XL.Print.current,
        solve: () => { if (XL.Print.current) XL.Print.current.close(); },
        success: 'Repare nas <b>linhas tracejadas</b> na grade: elas mostram onde cada página termina.',
      },
      {
        title: 'Repita o cabeçalho',
        text: '<p>Da segunda página em diante, ninguém sabe o que é cada coluna. <b>Imprimir Títulos</b> repete a linha de cabeçalho em todas as folhas.</p>',
        task: 'Na guia <b>Layout da Página</b>, clique em <b>Imprimir Títulos</b>, digite <b>$1:$1</b> e clique em OK.',
        target: ['.xl-tab[data-tab="layout"]', '.rb-btn[data-cmd="print-titles"]', 'input[data-role="titles"]'],
        check: (x) => !!x.sheet.page.titleRows && x.sheet.page.titleRows.r1 === 0 && x.sheet.page.titleRows.r2 === 0,
        solve: (x) => x.app.setPage({ titleRows: { r1: 0, r2: 0 } }),
        success: 'Pronto: a linha 1 vai aparecer no topo de todas as páginas.',
      },
      {
        title: 'Imprima só uma parte',
        task: 'Selecione <b>A1:F20</b> e use <b>Layout da Página › Área de Impressão › Definir Área de Impressão</b>.',
        hint: 'Arraste de A1 até F20 (ou digite A1:F20 na Caixa de Nome e Enter). Depois o botão <b>Área de Impressão</b>.',
        target: ['cell:A1:F20', '[data-menu="print-area"]', '.ui-menu-item[data-cmd="print-area"][data-arg="set"]'],
        check: (x) => { const a = x.sheet.page.printArea; return !!a && XL.rangeToStr(a) === 'A1:F20'; },
        solve: (x) => { x.app.select('A1:F20'); x.app.exec('print-area', 'set'); },
      },
      {
        title: 'Confira o resultado',
        task: 'Abra a visualização de impressão de novo (<kbd>Ctrl</kbd>+<kbd>P</kbd> ou Arquivo).',
        target: '.xl-tab[data-tab="file"]',
        check: (x) => x.happened('printpreview', (d) => d.pages === 1),
        solve: (x) => x.app.exec('print-preview'),
        success: 'Uma página só, com exatamente o pedaço escolhido. Para voltar a imprimir tudo: Área de Impressão › Limpar. Você concluiu o <b>Nível Básico</b>! 🏆',
      },
    ],
  });
})();
