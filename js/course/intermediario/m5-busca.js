/* Nível Intermediário · Módulo 5 — Buscando dados (3 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };
  const head = { b: true, fill: '#DDEBF7' };
  const headG = { b: true, fill: '#E2EFDA' };
  const res = { fill: '#FFF2CC', b: true };
  const moeda = { numFmt: 'moeda' };

  // ======================================================================
  // 5.1 PROCV
  // ======================================================================
  const PRODUTOS = [['P101', 'Caneta', 2.5], ['P102', 'Caderno', 15], ['P103', 'Mochila', 89.9], ['P104', 'Estojo', 12],
    ['P105', 'Agenda', 32], ['P106', 'Régua', 4], ['P107', 'Marca-texto', 6.5], ['P108', 'Calculadora', 45]];
  const PEDIDO = [['P103', 2], ['P101', 10], ['P108', 1], ['P105', 3], ['P102', 5], ['P110', 2]];
  XL.registerLesson({
    id: 'i5-1',
    title: 'PROCV',
    summary: [
      '<code>=PROCV(valor; tabela; nº_da_coluna; FALSO)</code> procura o valor na <b>1ª coluna</b> da tabela e traz o dado da coluna indicada.',
      'Use <b>FALSO</b> no último argumento para busca <b>exata</b> — quase sempre é o que você quer.',
      'Trave a tabela com <b>$</b> (<code>$G$2:$I$9</code>) para arrastar a fórmula.',
      '<b>#N/D</b> = não encontrado. Trate com <code>SEERRO</code>.',
      'Limite: o PROCV só procura na primeira coluna e só traz dados <b>à direita</b> dela.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Código', fmt: head }, B1: { v: 'Produto', fmt: head }, C1: { v: 'Preço unit.', fmt: head }, D1: { v: 'Qtd', fmt: head }, E1: { v: 'Total', fmt: head },
        G1: { v: 'Código', fmt: headG }, H1: { v: 'Produto', fmt: headG }, I1: { v: 'Preço', fmt: headG },
      };
      PEDIDO.forEach(([cod, q], i) => {
        const r = i + 2;
        m['A' + r] = cod; m['D' + r] = q;
        m['C' + r] = { v: null, fmt: moeda };
        m['E' + r] = { v: `=C${r}*D${r}`, fmt: moeda };
      });
      PRODUTOS.forEach(([cod, p, v], i) => { const r = i + 2; m['G' + r] = cod; m['H' + r] = p; m['I' + r] = { v, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 70, B: 120, C: 95, D: 45, E: 95, F: 24, G: 70, H: 100, I: 90 });
    },
    steps: [
      {
        title: 'O problema',
        text: `<p>À esquerda está um <b>pedido</b> só com os códigos. À direita (em verde), a <b>tabela de produtos</b> com nome e preço de cada código.</p>
          <p>Copiar na mão é lento e cheio de erro. O <b>PROCV</b> ("procurar na vertical") faz isso sozinho:</p>
          <p style="text-align:center"><code class="big">=PROCV(A2; $G$2:$I$9; 2; FALSO)</code></p>
          <ol><li><b>O que procurar</b>: o código em A2</li><li><b>Onde</b>: a tabela G2:I9 (o código precisa estar na <b>1ª coluna</b> dela)</li>
          <li><b>Qual coluna trazer</b>: 2 = a 2ª coluna da tabela (Produto)</li><li><b>FALSO</b>: só aceita o código <b>exato</b></li></ol>`,
        target: ['cell:A1:E7', 'cell:G1:I9'],
      },
      {
        title: 'Busque o nome do produto',
        task: 'Em <b>B2</b>, digite <code>=PROCV(A2;$G$2:$I$9;2;FALSO)</code>.',
        hint: 'Depois de selecionar G2:I9, pressione <kbd>F4</kbd> para virar <code>$G$2:$I$9</code> — assim a tabela não "escorrega" ao arrastar.',
        target: ['cell:B2', 'cell:G2:I9'],
        check: (x) => x.uses('B2', 'PROCV') && x.formula('B2').includes('$G$2:$I$9') && x.val('B2') === 'Mochila',
        solve: (x) => { x.app.setCell('B2', '=PROCV(A2;$G$2:$I$9;2;FALSO)'); x.app.select('B2'); },
        success: 'O Excel achou <b>P103</b> na primeira coluna da tabela e trouxe a 2ª coluna da mesma linha: <b>Mochila</b>.',
      },
      {
        title: 'Para o pedido todo',
        task: 'Arraste <b>B2</b> até <b>B7</b> com a alça de preenchimento.',
        target: ['cell:B2:B7', '.xl-handle'],
        check: (x) => x.val('B6') === 'Caderno' && x.isErr('B7', '#N/D'),
        solve: (x) => { x.app.select('B2'); x.app.fill(XL.parseRange('B2'), XL.parseRange('B2:B7'), 'down'); },
        success: 'Cinco produtos encontrados... e um <b>#N/D</b> na última linha. Por quê? Procure o código <b>P110</b> na tabela verde: ele não existe! <b>#N/D</b> significa "não disponível": o PROCV não achou o valor.',
      },
      {
        title: 'Agora o preço',
        task: 'Em <b>C2</b>, busque o <b>preço</b> (3ª coluna da tabela) e arraste até <b>C7</b>.',
        hint: '<code>=PROCV(A2;$G$2:$I$9;3;FALSO)</code> — só muda o número da coluna: 3.',
        target: ['cell:C2:C7', 'cell:I2:I9'],
        check: (x) => x.uses('C2', 'PROCV') && x.near(x.num('C2'), 89.9) && x.num('C6') === 15 && x.near(x.num('E2'), 179.8),
        solve: (x) => { x.app.setCell('C2', '=PROCV(A2;$G$2:$I$9;3;FALSO)'); x.app.select('C2'); x.app.fill(XL.parseRange('C2'), XL.parseRange('C2:C7'), 'down'); },
        success: 'Os totais da coluna E já calcularam sozinhos. Se o preço mudar na tabela verde, o pedido inteiro se atualiza.',
      },
      {
        title: 'Uma mensagem no lugar do #N/D',
        text: '<p>Você já conhece o <b>SEERRO</b>. Envolvendo o PROCV com ele, o #N/D vira uma mensagem clara.</p>',
        task: 'Em <b>B2</b>, escreva <code>=SEERRO(PROCV(A2;$G$2:$I$9;2;FALSO);"Não encontrado")</code> e arraste até <b>B7</b>.',
        target: ['cell:B2:B7'],
        check: (x) => x.uses('B2', 'SEERRO') && x.uses('B2', 'PROCV') && x.val('B7') === 'Não encontrado' && x.val('B2') === 'Mochila',
        solve: (x) => { x.app.setCell('B2', '=SEERRO(PROCV(A2;$G$2:$I$9;2;FALSO);"Não encontrado")'); x.app.select('B2'); x.app.fill(XL.parseRange('B2'), XL.parseRange('B2:B7'), 'down'); },
        success: 'Agora quem lê o pedido entende o problema na hora: o código está errado.',
      },
      {
        title: 'Corrija o código',
        task: 'O código certo era <b>P107</b>. Mude <b>A7</b> para <b>P107</b>.',
        target: 'cell:A7',
        check: (x) => x.text('A7') === 'P107' && x.val('B7') === 'Marca-texto' && x.near(x.num('C7'), 6.5),
        solve: (x) => { x.app.setCell('A7', 'P107'); x.app.select('A8'); },
        success: 'Nome, preço e total apareceram na hora: <b>Marca-texto</b>, R$ 6,50.',
      },
      {
        title: 'Por que FALSO? E o limite do PROCV',
        text: `<p><b>FALSO = busca exata.</b> Se você omitir ou usar VERDADEIRO, o PROCV faz uma busca <i>aproximada</i>: em vez de #N/D, pode trazer o item <b>errado</b> sem avisar. Use VERDADEIRO só em tabelas de faixas (ex.: faixas de comissão), ordenadas.</p>
          <p><b>O grande limite:</b> o PROCV procura só na <b>primeira coluna</b> da tabela e só traz colunas <b>à direita</b> dela. Se o código estivesse na última coluna, ele não conseguiria. A próxima lição resolve isso com o <b>PROCX</b>.</p>`,
      },
    ],
  });

  // ======================================================================
  // 5.2 PROCX
  // ======================================================================
  const FUNC = [['Ana Souza', 'Financeiro', 1021, 6500], ['Bruno Lima', 'TI', 1045, 8200], ['Carla Dias', 'RH', 1033, 5400],
    ['Diego Alves', 'Vendas', 1050, 4800], ['Elisa Rocha', 'TI', 1012, 9100], ['Fábio Melo', 'Financeiro', 1068, 5900],
    ['Gabi Nunes', 'Marketing', 1077, 6100], ['Hugo Prado', 'Vendas', 1003, 5100]];
  XL.registerLesson({
    id: 'i5-2',
    title: 'PROCX: a busca moderna',
    summary: [
      '<code>=PROCX(o_que_procurar; onde_procurar; o_que_trazer; [se_não_encontrar])</code>.',
      'Procura em <b>qualquer coluna</b> e traz de <b>qualquer lado</b> (inclusive à esquerda).',
      'Busca exata por padrão, sem número de coluna para contar.',
      'O 4º argumento substitui o SEERRO: <code>"Matrícula não existe"</code>.',
      'Disponível no Excel 2021 e Microsoft 365. Em versões antigas, use ÍNDICE + CORRESP.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Nome', fmt: head }, B1: { v: 'Departamento', fmt: head }, C1: { v: 'Matrícula', fmt: head }, D1: { v: 'Salário', fmt: head },
        F1: { v: 'Matrícula buscada:', fmt: bold }, G1: { v: 1045, fmt: { fill: '#FFF2CC', b: true, align: 'center' } },
        F3: 'Nome', F4: 'Departamento', F5: 'Salário',
        G3: { v: null, fmt: res }, G4: { v: null, fmt: res }, G5: { v: null, fmt: Object.assign({}, res, moeda) },
      };
      FUNC.forEach(([n, d, mat, s], i) => { const r = i + 2; m['A' + r] = n; m['B' + r] = d; m['C' + r] = mat; m['D' + r] = { v: s, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 110, B: 105, C: 80, D: 95, E: 24, F: 135, G: 150 });
    },
    steps: [
      {
        title: 'Quando o PROCV não dá conta',
        text: `<p>Aqui a <b>matrícula</b> está na coluna C e o <b>nome</b> na coluna A — à esquerda. O PROCV não consegue buscar "para trás".</p>
          <p>O <b>PROCX</b> (Excel 2021/365) separa as coisas: onde procurar e o que trazer são intervalos independentes:</p>
          <p style="text-align:center"><code class="big">=PROCX(G1; C2:C9; A2:A9)</code></p>
          <ol><li><b>O que procurar</b>: a matrícula em G1</li><li><b>Onde procurar</b>: a coluna de matrículas C2:C9</li><li><b>O que trazer</b>: a coluna de nomes A2:A9</li></ol>`,
        target: ['cell:A1:D9', 'cell:F1:G1'],
      },
      {
        title: 'Busque o nome',
        task: 'Em <b>G3</b>, digite <code>=PROCX(G1;C2:C9;A2:A9)</code>.',
        target: ['cell:G3', 'cell:C2:C9', 'cell:A2:A9'],
        check: (x) => x.uses('G3', 'PROCX') && x.val('G3') === 'Bruno Lima',
        solve: (x) => { x.app.setCell('G3', '=PROCX(G1;C2:C9;A2:A9)'); x.app.select('G4'); },
        success: 'Matrícula 1045 → <b>Bruno Lima</b>. O PROCX buscou na coluna C e trouxe da coluna A, que fica à esquerda.',
      },
      {
        title: 'Departamento e salário',
        task: 'Em <b>G4</b>, traga o <b>departamento</b> e, em <b>G5</b>, o <b>salário</b> da mesma matrícula.',
        hint: 'Só muda o 3º argumento: <code>=PROCX(G1;C2:C9;B2:B9)</code> e <code>=PROCX(G1;C2:C9;D2:D9)</code>.',
        target: ['cell:G4:G5'],
        check: (x) => x.uses('G4', 'PROCX') && x.uses('G5', 'PROCX') && x.val('G4') === 'TI' && x.num('G5') === 8200,
        solve: (x) => { x.app.setCell('G4', '=PROCX(G1;C2:C9;B2:B9)'); x.app.setCell('G5', '=PROCX(G1;C2:C9;D2:D9)'); x.app.select('G6'); },
        success: 'Sem contar colunas: você aponta direto para a coluna que quer. Se alguém inserir uma coluna no meio da tabela, o PROCX continua certo (o PROCV, com o número fixo, poderia trazer a coluna errada).',
      },
      {
        title: 'Uma ficha que se atualiza',
        task: 'Troque a matrícula em <b>G1</b> para <b>1077</b>.',
        target: 'cell:G1',
        check: (x) => x.num('G1') === 1077 && x.val('G3') === 'Gabi Nunes',
        solve: (x) => { x.app.setCell('G1', '1077'); x.app.select('G2'); },
        success: 'A ficha inteira mudou para a <b>Gabi Nunes</b>, do Marketing. Você acabou de criar uma pequena "tela de consulta".',
      },
      {
        title: 'Matrícula que não existe',
        task: 'Digite <b>9999</b> em <b>G1</b> e veja o que acontece.',
        target: 'cell:G1',
        check: (x) => x.num('G1') === 9999 && x.isErr('G3', '#N/D'),
        solve: (x) => { x.app.setCell('G1', '9999'); x.app.select('G2'); },
        success: 'O já conhecido <b>#N/D</b>. Mas o PROCX tem um jeito próprio de tratar isso...',
      },
      {
        title: 'O 4º argumento',
        text: '<p>O PROCX aceita um <b>4º argumento</b>: o que mostrar se não encontrar. Não precisa de SEERRO!</p>',
        task: 'Em <b>G3</b>, escreva <code>=PROCX(G1;C2:C9;A2:A9;"Matrícula não existe")</code>.',
        target: 'cell:G3',
        check: (x) => x.uses('G3', 'PROCX') && x.val('G3') === 'Matrícula não existe',
        solve: (x) => { x.app.setCell('G3', '=PROCX(G1;C2:C9;A2:A9;"Matrícula não existe")'); x.app.select('G4'); },
        success: `Resumo do PROCX contra o PROCV:
          <ul><li>busca para <b>qualquer lado</b></li><li><b>exata</b> por padrão (sem precisar do FALSO)</li><li>sem número de coluna para contar</li><li>trata o "não encontrado" sozinho</li></ul>
          Só tem um porém: arquivos abertos em Excel 2019 ou anterior mostram erro. Para esses casos, a próxima lição ensina a dupla clássica <b>ÍNDICE + CORRESP</b>.`,
      },
    ],
  });

  // ======================================================================
  // 5.3 ÍNDICE + CORRESP
  // ======================================================================
  const REG = ['Sul', 'Sudeste', 'Norte', 'Nordeste'];
  const FAIXAS = [['Até 1 kg', [15, 12, 25, 22]], ['Até 5 kg', [25, 20, 42, 38]], ['Até 10 kg', [38, 32, 65, 58]],
    ['Até 20 kg', [60, 50, 98, 90]], ['Até 30 kg', [85, 72, 140, 125]]];
  XL.registerLesson({
    id: 'i5-3',
    title: 'ÍNDICE + CORRESP',
    summary: [
      '<code>=ÍNDICE(tabela; linha; coluna)</code> devolve o valor numa posição da tabela.',
      '<code>=CORRESP(valor; lista; 0)</code> diz em que <b>posição</b> o valor está na lista (0 = exato).',
      'Juntos fazem uma busca em <b>duas direções</b>: <code>=ÍNDICE(B2:E6; CORRESP(...); CORRESP(...))</code>.',
      'Funcionam em <b>qualquer versão</b> do Excel.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Peso \\ Região', fmt: head }, G1: { v: 'Faixa de peso:', fmt: bold }, G2: { v: 'Região:', fmt: bold },
        H1: { v: 'Até 10 kg', fmt: { fill: '#FFF2CC', b: true } }, H2: { v: 'Norte', fmt: { fill: '#FFF2CC', b: true } },
        G4: 'Linha (CORRESP)', G5: 'Coluna (CORRESP)', G6: { v: 'Frete', fmt: bold }, G8: { v: 'Tudo numa fórmula', fmt: bold },
        H4: { v: null, fmt: res }, H5: { v: null, fmt: res }, H6: { v: null, fmt: Object.assign({}, res, moeda) }, H8: { v: null, fmt: Object.assign({}, res, moeda) } };
      REG.forEach((r, j) => { m[XL.numToCol(j + 2) + '1'] = { v: r, fmt: head }; });
      FAIXAS.forEach(([f, vals], i) => {
        m['A' + (i + 2)] = { v: f, fmt: bold };
        vals.forEach((v, j) => { m[XL.numToCol(j + 2) + (i + 2)] = { v, fmt: moeda }; });
      });
      x.set(m);
      x.widths({ A: 105, F: 24, G: 135, H: 105 });
    },
    steps: [
      {
        title: 'Uma tabela de duas entradas',
        text: `<p>Esta é uma tabela de <b>frete</b>: o preço depende da <b>faixa de peso</b> (linha) <b>e</b> da <b>região</b> (coluna). É uma busca em duas direções.</p>
          <p>A dupla que resolve isso em qualquer versão do Excel:</p>
          <ul><li><b>ÍNDICE</b>: "me dê o valor da linha X, coluna Y desta tabela"</li><li><b>CORRESP</b>: "em que posição está tal item nesta lista?"</li></ul>`,
        target: ['cell:A1:E6', 'cell:G1:H2'],
      },
      {
        title: 'ÍNDICE: valor por posição',
        task: 'Em <b>H6</b>, digite <code>=ÍNDICE(B2:E6;3;3)</code>.',
        hint: 'Lê-se: "na tabela B2:E6, pegue a 3ª linha e a 3ª coluna".',
        target: ['cell:H6', 'cell:B2:E6'],
        check: (x) => x.uses('H6', 'ÍNDICE') && x.num('H6') === 65,
        solve: (x) => { x.app.setCell('H6', '=ÍNDICE(B2:E6;3;3)'); x.app.select('H7'); },
        success: '3ª linha (Até 10 kg) e 3ª coluna (Norte): <b>R$ 65,00</b>. Mas digitar "3" na mão não serve — queremos que o Excel descubra as posições sozinho.',
      },
      {
        title: 'CORRESP: em que linha?',
        task: 'Em <b>H4</b>, descubra a posição da faixa: <code>=CORRESP(H1;A2:A6;0)</code>.',
        hint: 'O <b>0</b> no final pede correspondência <b>exata</b> (como o FALSO do PROCV).',
        target: ['cell:H4', 'cell:A2:A6'],
        check: (x) => x.uses('H4', 'CORRESP') && x.num('H4') === 3,
        solve: (x) => { x.app.setCell('H4', '=CORRESP(H1;A2:A6;0)'); x.app.select('H5'); },
        success: '"Até 10 kg" é o <b>3º item</b> da lista A2:A6.',
      },
      {
        title: 'CORRESP: em que coluna?',
        task: 'Em <b>H5</b>, descubra a posição da região na linha de cabeçalho: <code>=CORRESP(H2;B1:E1;0)</code>.',
        hint: 'O CORRESP funciona em listas na horizontal também.',
        target: ['cell:H5', 'cell:B1:E1'],
        check: (x) => x.uses('H5', 'CORRESP') && x.num('H5') === 3,
        solve: (x) => { x.app.setCell('H5', '=CORRESP(H2;B1:E1;0)'); x.app.select('H6'); },
        success: '"Norte" é o <b>3º</b> da lista B1:E1.',
      },
      {
        title: 'Junte as peças',
        task: 'Em <b>H6</b>, troque os números fixos pelas posições: <code>=ÍNDICE(B2:E6;H4;H5)</code>.',
        target: ['cell:H6', 'cell:H4:H5'],
        check: (x) => x.uses('H6', 'ÍNDICE') && x.refs('H6').includes('H4') && x.refs('H6').includes('H5') && x.num('H6') === 65,
        solve: (x) => { x.app.setCell('H6', '=ÍNDICE(B2:E6;H4;H5)'); x.app.select('H7'); },
      },
      {
        title: 'Teste com outra região',
        task: 'Troque a região em <b>H2</b> para <b>Sul</b>.',
        target: 'cell:H2',
        check: (x) => x.text('H2') === 'SUL' && x.num('H5') === 1 && x.num('H6') === 38,
        solve: (x) => { x.app.setCell('H2', 'Sul'); x.app.select('H3'); },
        success: 'O CORRESP achou "Sul" na posição 1 e o ÍNDICE trouxe <b>R$ 38,00</b>. Tudo se recalcula em cadeia.',
      },
      {
        title: 'Tudo numa fórmula só',
        text: '<p>Os profissionais costumam colocar os CORRESP <b>dentro</b> do ÍNDICE, sem células auxiliares:</p><p style="text-align:center"><code>=ÍNDICE(B2:E6; CORRESP(H1;A2:A6;0); CORRESP(H2;B1:E1;0))</code></p>',
        task: 'Em <b>H8</b>, digite a fórmula completa acima.',
        target: 'cell:H8',
        check: (x) => x.uses('H8', 'ÍNDICE') && x.uses('H8', 'CORRESP') && x.num('H8') === x.num('H6'),
        solve: (x) => { x.app.setCell('H8', '=ÍNDICE(B2:E6;CORRESP(H1;A2:A6;0);CORRESP(H2;B1:E1;0))'); x.app.select('H9'); },
        success: 'Mesmo resultado de H6. Você dominou as três formas de buscar dados no Excel: <b>PROCV</b>, <b>PROCX</b> e <b>ÍNDICE + CORRESP</b>. 🏆',
      },
    ],
  });
})();
