/* Nível Intermediário · Módulo 4 — Contar e somar com critérios (3 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };
  const head = { b: true, fill: '#DDEBF7' };
  const res = { fill: '#FFF2CC', b: true }; // células de resultado
  const moeda = { numFmt: 'moeda' };

  // Registro de pedidos usado nas lições 4.1 e 4.2 (linhas 2 a 17).
  const PEDIDOS = [
    ['Ana', 'Sul', 'Notebook', 2, 7000], ['Bruno', 'Norte', 'Mouse', 10, 450], ['Ana', 'Sul', 'Monitor', 3, 2700],
    ['Carla', 'Sudeste', 'Notebook', 1, 3500], ['Bruno', 'Norte', 'Teclado', 5, 600], ['Diego', 'Sul', 'Headset', 4, 840],
    ['Carla', 'Sudeste', 'Monitor', 2, 1800], ['Ana', 'Sul', 'Mouse', 8, 360], ['Elisa', 'Nordeste', 'Notebook', 3, 10500],
    ['Diego', 'Sul', 'Notebook', 1, 3500], ['Bruno', 'Norte', 'Monitor', 1, 900], ['Elisa', 'Nordeste', 'Webcam', 6, 2100],
    ['Carla', 'Sudeste', 'Teclado', 3, 360], ['Ana', 'Sul', 'Headset', 2, 420], ['Diego', 'Sul', 'Mouse', 12, 540],
    ['Elisa', 'Nordeste', 'Monitor', 2, 1800],
  ];
  function pedidos(x) {
    const m = {};
    ['Vendedor', 'Região', 'Produto', 'Qtd', 'Valor'].forEach((t, i) => { m[XL.numToCol(i + 1) + '1'] = { v: t, fmt: head }; });
    PEDIDOS.forEach((p, i) => {
      const r = i + 2;
      m['A' + r] = p[0]; m['B' + r] = p[1]; m['C' + r] = p[2]; m['D' + r] = p[3]; m['E' + r] = { v: p[4], fmt: moeda };
    });
    x.set(m);
    x.widths({ A: 80, B: 90, C: 90, D: 45, E: 100, F: 24 });
  }

  // ======================================================================
  // 4.1 CONT.SE e CONT.SES
  // ======================================================================
  XL.registerLesson({
    id: 'i4-1',
    title: 'CONT.SE e CONT.SES',
    summary: [
      '<code>=CONT.SE(intervalo; critério)</code> conta as células que atendem a <b>uma</b> condição.',
      'Critérios de texto e com operadores vão <b>entre aspas</b>: <code>"Sul"</code>, <code>"&gt;2000"</code>, <code>"&lt;&gt;Sul"</code>.',
      '<code>=CONT.SES(intervalo1; critério1; intervalo2; critério2; ...)</code> exige <b>todas</b> as condições.',
    ],
    setup: (x) => {
      pedidos(x);
      x.set({
        G1: { v: 'Resumo', fmt: { b: true, size: 12 } },
        G2: 'Pedidos da Ana', G3: 'Pedidos da região Sul', G4: 'Pedidos acima de R$ 2.000',
        G5: 'Notebooks vendidos no Sul', G6: 'Pedidos do Sul com valor ≥ R$ 1.000',
        H2: { v: null, fmt: res }, H3: { v: null, fmt: res }, H4: { v: null, fmt: res }, H5: { v: null, fmt: res }, H6: { v: null, fmt: res },
      });
      x.widths({ G: 230, H: 60 });
    },
    steps: [
      {
        title: 'Contar com uma condição',
        text: `<p>Você já conhece o <code>CONT.NÚM</code>, que conta números. Mas perguntas reais são do tipo: <i>"quantos pedidos a Ana fez?"</i></p>
          <p>Para isso existe o <b>CONT.SE</b> — "conte <b>se</b>...":</p>
          <p style="text-align:center"><code class="big">=CONT.SE(A2:A17; "Ana")</code></p>
          <ul><li>1º argumento: <b>onde</b> procurar (o intervalo)</li><li>2º: <b>o que</b> contar (o critério)</li></ul>`,
        target: ['cell:A1:E17', 'cell:G2:H6'],
      },
      {
        title: 'Quantos pedidos da Ana?',
        task: 'Em <b>H2</b>, digite <code>=CONT.SE(A2:A17;"Ana")</code>.',
        hint: 'Digite <code>=CONT.SE(</code>, selecione A2 até A17 com o mouse, digite <code>;"Ana")</code> e Enter.',
        target: ['cell:H2', 'cell:A2:A17'],
        check: (x) => x.uses('H2', 'CONT.SE') && x.num('H2') === 4,
        solve: (x) => { x.app.setCell('H2', '=CONT.SE(A2:A17;"Ana")'); x.app.select('H3'); },
        success: '<b>4 pedidos</b>. Maiúsculas não importam: "ana" e "ANA" contariam igual.',
      },
      {
        title: 'Quantos pedidos do Sul?',
        task: 'Em <b>H3</b>, conte os pedidos da região <b>Sul</b>.',
        hint: 'Mesmo formato, olhando a coluna B: <code>=CONT.SE(B2:B17;"Sul")</code>.',
        target: ['cell:H3', 'cell:B2:B17'],
        check: (x) => x.uses('H3', 'CONT.SE') && x.num('H3') === 7,
        solve: (x) => { x.app.setCell('H3', '=CONT.SE(B2:B17;"Sul")'); x.app.select('H4'); },
        success: '<b>7 pedidos</b> no Sul (da Ana e do Diego).',
      },
      {
        title: 'Critérios com operadores',
        text: `<p>O critério pode ser uma <b>comparação</b>. O operador e o número vão <b>juntos, entre aspas</b>:</p>
          <table class="mini"><tr><th>Critério</th><th>Conta células...</th></tr>
          <tr><td><code>"&gt;2000"</code></td><td>maiores que 2000</td></tr>
          <tr><td><code>"&lt;=100"</code></td><td>menores ou iguais a 100</td></tr>
          <tr><td><code>"&lt;&gt;Sul"</code></td><td>diferentes de Sul</td></tr>
          <tr><td><code>"Sul"</code></td><td>iguais a Sul</td></tr></table>`,
      },
      {
        title: 'Pedidos grandes',
        task: 'Em <b>H4</b>, conte os pedidos com <b>Valor maior que 2000</b>: <code>=CONT.SE(E2:E17;"&gt;2000")</code>.',
        target: ['cell:H4', 'cell:E2:E17'],
        check: (x) => x.uses('H4', 'CONT.SE') && x.num('H4') === 6,
        solve: (x) => { x.app.setCell('H4', '=CONT.SE(E2:E17;">2000")'); x.app.select('H5'); },
        success: '<b>6 pedidos</b> passaram de R$ 2.000. O de R$ 1.800 ficou de fora, como esperado.',
      },
      {
        title: 'Duas condições: CONT.SES',
        text: `<p>E se a pergunta tiver <b>duas</b> condições? "Quantos <b>notebooks</b> foram vendidos <b>no Sul</b>?"</p>
          <p>O <b>CONT.SES</b> (com S no final, de "vários") recebe <b>pares</b> de intervalo e critério:</p>
          <p style="text-align:center"><code>=CONT.SES(B2:B17;"Sul"; C2:C17;"Notebook")</code></p>
          <p>Só conta a linha se <b>todas</b> as condições forem verdadeiras — como o E que você aprendeu.</p>`,
        task: 'Em <b>H5</b>, digite a fórmula acima.',
        target: ['cell:H5', 'cell:B2:C17'],
        check: (x) => x.uses('H5', 'CONT.SES') && x.num('H5') === 2,
        solve: (x) => { x.app.setCell('H5', '=CONT.SES(B2:B17;"Sul";C2:C17;"Notebook")'); x.app.select('H6'); },
        success: '<b>2 notebooks</b> no Sul: um da Ana e um do Diego.',
      },
      {
        title: 'Misturando texto e número',
        task: 'Em <b>H6</b>, conte os pedidos do <b>Sul</b> com valor <b>maior ou igual a 1000</b>.',
        hint: '<code>=CONT.SES(B2:B17;"Sul";E2:E17;"&gt;=1000")</code>',
        target: 'cell:H6',
        check: (x) => x.uses('H6', 'CONT.SES') && x.num('H6') === 3,
        solve: (x) => { x.app.setCell('H6', '=CONT.SES(B2:B17;"Sul";E2:E17;">=1000")'); x.app.select('H7'); },
        success: '<b>3 pedidos</b>. Você pode juntar quantos pares quiser no CONT.SES. Os intervalos precisam ter o <b>mesmo tamanho</b> (aqui, linhas 2 a 17).',
      },
    ],
  });

  // ======================================================================
  // 4.2 SOMASE e SOMASES
  // ======================================================================
  const REGIOES = [['Sul', 15360], ['Norte', 1950], ['Sudeste', 5660], ['Nordeste', 14400]];
  XL.registerLesson({
    id: 'i4-2',
    title: 'SOMASE e SOMASES',
    summary: [
      '<code>=SOMASE(onde_testar; critério; o_que_somar)</code> soma só as linhas que atendem ao critério.',
      'O critério pode vir de uma <b>célula</b> (<code>G2</code>) — com intervalos travados com <b>$</b>, dá para arrastar e montar um resumo.',
      '<code>=SOMASES(o_que_somar; intervalo1; critério1; ...)</code>: no SOMASES o intervalo da soma vem <b>primeiro</b>.',
    ],
    setup: (x) => {
      pedidos(x);
      const m = {
        G1: { v: 'Região', fmt: head }, H1: { v: 'Total vendido', fmt: head },
        G6: { v: 'Total', fmt: bold }, H6: { v: null, fmt: Object.assign({ b: true }, moeda) },
        J1: { v: 'Perguntas', fmt: head }, K1: { v: 'Resposta', fmt: head },
        J2: 'Faturamento da Ana (R$)', J3: 'Mouses vendidos (unidades)', J4: 'Notebooks no Sul (R$)',
        K2: { v: null, fmt: Object.assign({}, res, moeda) }, K3: { v: null, fmt: res }, K4: { v: null, fmt: Object.assign({}, res, moeda) },
      };
      REGIOES.forEach(([reg], i) => { m['G' + (i + 2)] = reg; m['H' + (i + 2)] = { v: null, fmt: moeda }; });
      x.set(m);
      x.widths({ G: 80, H: 105, I: 24, J: 190, K: 105 });
    },
    steps: [
      {
        title: 'Somar só o que interessa',
        text: `<p>O <b>SOMASE</b> é o primo do CONT.SE: em vez de contar, ele <b>soma</b> outra coluna — mas só nas linhas que passam no critério.</p>
          <p style="text-align:center"><code class="big">=SOMASE(A2:A17; "Ana"; E2:E17)</code></p>
          <p>Leia assim: <i>"olhe a coluna A; onde for <b>Ana</b>, some o valor da coluna E"</i>.</p>
          <ul><li>1º: onde <b>testar</b> · 2º: o <b>critério</b> · 3º: o que <b>somar</b></li></ul>`,
        target: ['cell:A1:E17', 'cell:J1:K4'],
      },
      {
        title: 'Quanto a Ana vendeu?',
        task: 'Em <b>K2</b>, digite <code>=SOMASE(A2:A17;"Ana";E2:E17)</code>.',
        hint: 'Os dois intervalos (A2:A17 e E2:E17) precisam ter o <b>mesmo tamanho</b>, para o Excel casar linha com linha.',
        target: ['cell:K2', 'cell:A2:A17', 'cell:E2:E17'],
        check: (x) => x.uses('K2', 'SOMASE') && x.near(x.num('K2'), 10480),
        solve: (x) => { x.app.setCell('K2', '=SOMASE(A2:A17;"Ana";E2:E17)'); x.app.select('K3'); },
        success: 'A Ana faturou <b>R$ 10.480,00</b> em 4 pedidos.',
      },
      {
        title: 'Somando outra coluna',
        task: 'Em <b>K3</b>, some a <b>Qtd</b> (coluna D) dos pedidos de <b>Mouse</b>.',
        hint: '<code>=SOMASE(C2:C17;"Mouse";D2:D17)</code> — o critério está na coluna C, a soma é da coluna D.',
        target: ['cell:K3', 'cell:C2:D17'],
        check: (x) => x.uses('K3', 'SOMASE') && x.num('K3') === 30,
        solve: (x) => { x.app.setCell('K3', '=SOMASE(C2:C17;"Mouse";D2:D17)'); x.app.select('K4'); },
        success: '<b>30 mouses</b>. Você pode somar qualquer coluna: valor, quantidade, horas...',
      },
      {
        title: 'Um resumo por região',
        text: `<p>Em vez de escrever <code>"Sul"</code> na fórmula, use a <b>célula</b> que já tem o nome da região (G2). Travando os intervalos com <b>$</b>, a mesma fórmula funciona para todas as regiões quando você arrasta:</p>
          <p style="text-align:center"><code>=SOMASE($B$2:$B$17; G2; $E$2:$E$17)</code></p>
          <p>Ao arrastar, só o <b>G2</b> muda (G3, G4, G5); os intervalos ficam travados.</p>`,
        task: 'Em <b>H2</b>, digite a fórmula acima e arraste até <b>H5</b>.',
        hint: 'Para colocar os $, pressione <kbd>F4</kbd> logo depois de cada intervalo (B2:B17 e E2:E17). O G2 fica <b>sem</b> $.',
        target: ['cell:H2:H5', 'cell:G2:G5'],
        check: (x) => x.uses('H2', 'SOMASE') && x.formula('H5').includes('G5') && REGIOES.every(([, v], i) => x.near(x.num('H' + (i + 2)), v)),
        solve: (x) => { x.app.setCell('H2', '=SOMASE($B$2:$B$17;G2;$E$2:$E$17)'); x.app.select('H2'); x.app.fill(XL.parseRange('H2'), XL.parseRange('H2:H5'), 'down'); },
        success: 'Um resumo por região com uma fórmula só! O Sul lidera com R$ 15.360,00.',
      },
      {
        title: 'Confira o total',
        task: 'Em <b>H6</b>, some as regiões: <code>=SOMA(H2:H5)</code>.',
        target: 'cell:H6',
        check: (x) => x.uses('H6', 'SOMA') && x.near(x.num('H6'), 37370),
        solve: (x) => { x.app.setCell('H6', '=SOMA(H2:H5)'); x.app.select('H7'); },
        success: '<b>R$ 37.370,00</b> — exatamente a soma de todos os pedidos. Conferir se o resumo bate com o total é um hábito que evita muitos erros.',
      },
      {
        title: 'Duas condições: SOMASES',
        text: `<p>Para somar com <b>várias</b> condições, use <b>SOMASES</b>. Atenção à ordem — ela é diferente do SOMASE:</p>
          <p style="text-align:center"><code>=SOMASES(<b>E2:E17</b>; B2:B17;"Sul"; C2:C17;"Notebook")</code></p>
          <p>No SOMASES, o intervalo que será <b>somado vem primeiro</b>, seguido dos pares intervalo/critério.</p>`,
        task: 'Em <b>K4</b>, some o valor dos <b>Notebooks</b> vendidos no <b>Sul</b>.',
        target: ['cell:K4'],
        check: (x) => x.uses('K4', 'SOMASES') && x.near(x.num('K4'), 10500),
        solve: (x) => { x.app.setCell('K4', '=SOMASES(E2:E17;B2:B17;"Sul";C2:C17;"Notebook")'); x.app.select('K5'); },
        success: '<b>R$ 10.500,00</b> (7.000 da Ana + 3.500 do Diego). Dica: muita gente usa SOMASES até com uma condição só, para não ter que lembrar duas ordens diferentes.',
      },
    ],
  });

  // ======================================================================
  // 4.3 MÉDIASE e critérios com curingas
  // ======================================================================
  const CATALOGO = [
    ['Monitor 24 pol', 899], ['Mouse sem fio', 89], ['Teclado mecânico', 349], ['Monitor 27 pol', 1299], ['Mouse com fio', 39],
    ['Headset', 259], ['Teclado sem fio', 179], ['Monitor 32 pol', 1899], ['Webcam HD', 219], ['Mouse gamer', 199],
  ];
  XL.registerLesson({
    id: 'i4-3',
    title: 'MÉDIASE e critérios com curingas',
    summary: [
      'Curingas: <code>*</code> = qualquer quantidade de caracteres, <code>?</code> = um caractere.',
      '<code>"Mouse*"</code> começa com Mouse · <code>"*sem fio"</code> termina com "sem fio" · <code>"*fio*"</code> contém "fio".',
      '<code>=MÉDIASE(onde_testar; critério; o_que_calcular_a_média)</code> funciona como o SOMASE.',
      'Critério vindo de célula com operador: <code>"&gt;"&amp;E1</code> (o <b>&amp;</b> junta o texto com o valor da célula).',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Produto', fmt: head }, B1: { v: 'Preço', fmt: head },
        D1: { v: 'Preço mínimo', fmt: bold }, E1: { v: 500, fmt: Object.assign({ fill: '#FFF2CC', b: true }, moeda) },
        D3: 'Quantos mouses?', D4: 'Quantos itens "sem fio"?', D5: 'Preço médio dos monitores', D6: 'Itens acima do preço mínimo',
        E3: { v: null, fmt: res }, E4: { v: null, fmt: res }, E5: { v: null, fmt: Object.assign({}, res, moeda) }, E6: { v: null, fmt: res },
      };
      CATALOGO.forEach(([p, v], i) => { m['A' + (i + 2)] = p; m['B' + (i + 2)] = { v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 140, B: 95, C: 24, D: 200, E: 110 });
    },
    steps: [
      {
        title: 'Quando o texto não é exato',
        text: `<p>Os produtos têm nomes compridos: "Mouse sem fio", "Mouse gamer"... Com <code>"Mouse"</code> o CONT.SE só acharia células escritas <b>exatamente</b> "Mouse".</p>
          <p>Os <b>curingas</b> resolvem:</p>
          <table class="mini"><tr><th>Critério</th><th>Encontra</th></tr>
          <tr><td><code>"Mouse*"</code></td><td>começa com "Mouse"</td></tr>
          <tr><td><code>"*sem fio"</code></td><td>termina com "sem fio"</td></tr>
          <tr><td><code>"*pol*"</code></td><td>contém "pol" em qualquer lugar</td></tr>
          <tr><td><code>"Monitor ?? pol"</code></td><td>"Monitor", 2 caracteres quaisquer, " pol"</td></tr></table>
          <p>O <code>*</code> vale "qualquer coisa" (até nada); o <code>?</code> vale exatamente <b>um</b> caractere.</p>`,
        target: 'cell:A1:B11',
      },
      {
        title: 'Começa com...',
        task: 'Em <b>E3</b>, conte os produtos que <b>começam com "Mouse"</b>: <code>=CONT.SE(A2:A11;"Mouse*")</code>.',
        target: ['cell:E3', 'cell:A2:A11'],
        check: (x) => x.uses('E3', 'CONT.SE') && x.num('E3') === 3,
        solve: (x) => { x.app.setCell('E3', '=CONT.SE(A2:A11;"Mouse*")'); x.app.select('E4'); },
        success: '<b>3 mouses</b>: sem fio, com fio e gamer.',
      },
      {
        title: 'Termina com...',
        task: 'Em <b>E4</b>, conte os itens que <b>terminam com "sem fio"</b>.',
        hint: 'Agora o asterisco vai <b>antes</b>: <code>=CONT.SE(A2:A11;"*sem fio")</code>.',
        target: ['cell:E4', 'cell:A2:A11'],
        check: (x) => x.uses('E4', 'CONT.SE') && x.num('E4') === 2,
        solve: (x) => { x.app.setCell('E4', '=CONT.SE(A2:A11;"*sem fio")'); x.app.select('E5'); },
        success: '<b>2 itens</b>: o mouse e o teclado sem fio.',
      },
      {
        title: 'MÉDIASE',
        text: '<p>O <b>MÉDIASE</b> tem os mesmos argumentos do SOMASE, mas calcula a <b>média</b> em vez da soma. E aceita curingas também.</p>',
        task: 'Em <b>E5</b>, calcule o preço médio dos <b>monitores</b>: <code>=MÉDIASE(A2:A11;"Monitor*";B2:B11)</code>.',
        target: ['cell:E5', 'cell:A2:B11'],
        check: (x) => x.uses('E5', 'MÉDIASE') && x.near(x.num('E5'), 4097 / 3, 1e-6),
        solve: (x) => { x.app.setCell('E5', '=MÉDIASE(A2:A11;"Monitor*";B2:B11)'); x.app.select('E6'); },
        success: 'Média dos 3 monitores: <b>R$ 1.365,67</b>. Existe também o <code>MÉDIASES</code>, para várias condições.',
      },
      {
        title: 'Critério que vem de uma célula',
        text: `<p>Queremos contar os itens acima do <b>preço mínimo</b> que está em <b>E1</b>. A tentação é escrever <code>"&gt;E1"</code> — mas aí o Excel procura o <b>texto</b> "E1"!</p>
          <p>O certo é deixar o operador entre aspas e <b>juntar</b> com a célula usando <code>&amp;</code>:</p>
          <p style="text-align:center"><code class="big">"&gt;"&amp;E1</code> &nbsp;→&nbsp; vira <code>"&gt;500"</code></p>`,
        task: 'Em <b>E6</b>, digite <code>=CONT.SE(B2:B11;"&gt;"&amp;E1)</code>.',
        hint: 'O <b>&amp;</b> fica entre a aspa de fechamento e o E1: <code>"&gt;"&amp;E1</code>.',
        target: ['cell:E6', 'cell:E1'],
        check: (x) => x.uses('E6', 'CONT.SE') && x.formula('E6').includes('&E1') && x.num('E6') === 3,
        solve: (x) => { x.app.setCell('E6', '=CONT.SE(B2:B11;">"&E1)'); x.app.select('E7'); },
        success: '<b>3 itens</b> acima de R$ 500 — os três monitores.',
      },
      {
        title: 'Mude o preço mínimo',
        task: 'Mude o preço mínimo em <b>E1</b> para <b>200</b> e veja a contagem.',
        target: 'cell:E1',
        check: (x) => x.num('E1') === 200 && x.num('E6') === 6,
        solve: (x) => { x.app.setCell('E1', '200'); x.app.select('E2'); },
        success: 'Agora são <b>6 itens</b> — o critério acompanhou a célula. Você concluiu o Módulo 4! 🏆',
      },
    ],
  });
})();
