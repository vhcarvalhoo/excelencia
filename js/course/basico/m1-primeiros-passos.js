/* Nível Básico · Módulo 1 — Primeiros passos (6 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };

  // ======================================================================
  // 1.1 Conhecendo o Excel
  // ======================================================================
  XL.registerLesson({
    id: 'b1-1',
    title: 'Conhecendo o Excel',
    summary: [
      'A <b>Faixa de Opções</b> organiza os comandos em guias (Página Inicial, Fórmulas, Dados...).',
      'Colunas têm <b>letras</b>, linhas têm <b>números</b>, e cada célula tem um endereço como <b>C3</b>.',
      'A <b>Caixa de Nome</b> mostra a célula ativa e também leva você direto a qualquer célula.',
      'Um <b>intervalo</b> é escrito com dois-pontos: <code>B2:D6</code>.',
    ],
    steps: [
      {
        title: 'Bem-vindo ao Excel! 👋',
        text: `<p>O Excel é um programa de <b>planilhas</b>: uma grande grade onde você organiza dados, faz contas automáticas e cria relatórios.</p>
          <p>À esquerda está um <b>simulador do Excel</b> — ele funciona como o de verdade. Aqui no painel você recebe as instruções e o guia confere cada tarefa sozinho.</p>
          <p>Nesta primeira lição você vai conhecer as partes da tela. É rapidinho!</p>`,
        target: '.xl-grid-host',
      },
      {
        title: 'A Faixa de Opções',
        text: `<p>No topo fica a <b>Faixa de Opções</b>. Ela é dividida em <b>guias</b> (Página Inicial, Inserir, Fórmulas, Dados...) e cada guia agrupa botões parecidos.</p>
          <p>A guia <b>Página Inicial</b> tem os comandos mais usados no dia a dia: negrito, cores, formatos de número, AutoSoma...</p>
          <p class="note">Passe o mouse sobre um botão para ver o nome dele.</p>`,
        target: ['.xl-tabs', '.xl-ribbon'],
      },
      {
        title: 'Colunas, linhas e células',
        text: `<p>As <b>colunas</b> são identificadas por <b>letras</b> (A, B, C...) e as <b>linhas</b> por <b>números</b> (1, 2, 3...).</p>
          <p>O encontro de uma coluna com uma linha é uma <b>célula</b>. O endereço da célula é a letra da coluna seguida do número da linha.</p>
          <p>Exemplo: coluna <b>C</b> + linha <b>3</b> = célula <b>C3</b>.</p>`,
        target: ['.xl-colh[data-c="2"]', '.xl-rowh[data-r="2"]', 'cell:C3'],
      },
      {
        title: 'Selecione uma célula',
        task: 'Clique na célula <b>C3</b>.',
        hint: 'Ache a coluna <b>C</b> (a terceira) e desça até a linha <b>3</b>. Clique uma vez.',
        target: 'cell:C3',
        check: (x) => x.active() === 'C3' && x.happened('select'),
        solve: (x) => x.app.select('C3'),
        success: 'A célula ganhou uma <b>borda verde</b>: ela é a <b>célula ativa</b>, onde o que você digitar vai entrar. Repare que a letra C e o número 3 também ficaram destacados.',
      },
      {
        title: 'A Caixa de Nome',
        text: `<p>À esquerda, logo acima da grade, fica a <b>Caixa de Nome</b>. Ela sempre mostra o endereço da célula ativa — agora deve estar mostrando <b>C3</b>.</p>
          <p>Ao lado dela fica a <b>Barra de Fórmulas</b>, que mostra o conteúdo da célula ativa. Você vai usá-la muito a partir da lição 3.</p>`,
        target: ['.xl-namebox', '.xl-finput'],
      },
      {
        title: 'Navegue com o teclado',
        text: '<p>Usar o teclado é muito mais rápido que o mouse. As <b>setas</b> movem a célula ativa uma posição para cada lado.</p>',
        task: 'Usando as <kbd>setas</kbd> do teclado, vá até a célula <b>E5</b>.',
        hint: 'Saindo de C3: aperte <kbd>→</kbd> duas vezes e <kbd>↓</kbd> duas vezes.',
        target: 'cell:E5',
        check: (x) => x.active() === 'E5' && x.happened('select'),
        solve: (x) => x.app.select('E5'),
        success: 'Dica de ouro: <kbd>Enter</kbd> desce uma célula e <kbd>Tab</kbd> vai para a direita.',
      },
      {
        title: 'Pule direto para uma célula',
        text: '<p>A Caixa de Nome também serve para <b>ir</b> até uma célula. Em planilhas enormes, isso economiza muito tempo.</p>',
        task: 'Clique na <b>Caixa de Nome</b>, digite <b>H10</b> e pressione <kbd>Enter</kbd>.',
        hint: 'A Caixa de Nome é o campo branco no canto esquerdo, logo abaixo da Faixa de Opções. Clique nela (o texto fica selecionado), digite <b>H10</b> e tecle <kbd>Enter</kbd>.',
        target: '.xl-namebox',
        check: (x) => x.happened('namebox', (d) => d.ref === 'H10'),
        solve: (x) => x.app.gotoRef('H10'),
        success: 'A seleção pulou direto para H10.',
      },
      {
        title: 'Selecione um intervalo',
        text: `<p>Um <b>intervalo</b> é um bloco de células. Ele é escrito com <b>dois-pontos</b>: <code>B2:D6</code> quer dizer "de B2 até D6".</p>
          <p>Selecionar intervalos é o que você mais vai fazer no Excel: para formatar, somar, copiar...</p>`,
        task: 'Clique em <b>B2</b> e, segurando o botão do mouse, arraste até <b>D6</b>.',
        hint: 'Outra forma: clique em B2, segure <kbd>Shift</kbd> e clique em D6. Ou use <kbd>Shift</kbd> + setas.',
        target: 'cell:B2:D6',
        check: (x) => x.sel() === 'B2:D6',
        solve: (x) => x.app.select('B2:D6'),
        success: 'Enquanto você arrastava, a Caixa de Nome mostrou o tamanho da seleção (<b>5L x 3C</b> = 5 linhas por 3 colunas). A célula B2, onde você começou, continua sendo a célula ativa (branca).',
      },
      {
        title: 'Explore as guias',
        task: 'Clique na guia <b>Fórmulas</b> da Faixa de Opções.',
        target: '.xl-tab[data-tab="formulas"]',
        check: (x) => x.tab() === 'formulas',
        solve: (x) => x.app.ribbon.setTab('formulas'),
        success: 'Cada guia mostra um conjunto diferente de comandos. A guia Fórmulas reúne a <b>biblioteca de funções</b> do Excel — vamos usá-la bastante.',
      },
      {
        title: 'De volta à Página Inicial',
        task: 'Volte para a guia <b>Página Inicial</b>.',
        target: '.xl-tab[data-tab="home"]',
        check: (x) => x.tab() === 'home',
        solve: (x) => x.app.ribbon.setTab('home'),
        success: 'Pronto! Você já sabe se localizar no Excel. Na próxima lição você começa a digitar dados.',
      },
    ],
  });

  // ======================================================================
  // 1.2 Digitando e editando dados
  // ======================================================================
  XL.registerLesson({
    id: 'b1-2',
    title: 'Digitando e editando dados',
    summary: [
      'Digite e confirme com <kbd>Enter</kbd> (desce) ou <kbd>Tab</kbd> (vai para a direita). <kbd>Esc</kbd> cancela.',
      'No Brasil, números decimais usam <b>vírgula</b>: <code>2,50</code>.',
      'Números ficam à <b>direita</b>, textos à <b>esquerda</b> — é assim que você confere se o Excel entendeu.',
      '<kbd>F2</kbd> ou duplo clique editam sem apagar; <kbd>Delete</kbd> apaga; <kbd>Ctrl</kbd>+<kbd>Z</kbd> desfaz.',
    ],
    setup: (x) => x.widths({ A: 110 }),
    steps: [
      {
        title: 'Como digitar no Excel',
        text: `<p>É simples: <b>selecione</b> uma célula e <b>digite</b>. Para confirmar:</p>
          <ul><li><kbd>Enter</kbd> — confirma e desce para a célula de baixo</li>
          <li><kbd>Tab</kbd> — confirma e vai para a célula da direita</li>
          <li><kbd>Esc</kbd> — desiste do que digitou</li></ul>
          <p>Enquanto você digita, o canto inferior esquerdo (barra de status) muda de <b>Pronto</b> para <b>Digite</b>.</p>`,
        target: '.xl-status-mode',
      },
      {
        title: 'Seu primeiro dado',
        task: 'Na célula <b>A1</b>, digite <b>Produto</b> e pressione <kbd>Enter</kbd>.',
        hint: 'Clique em A1, digite a palavra Produto e aperte Enter.',
        target: 'cell:A1',
        check: (x) => x.text('A1') === 'PRODUTO',
        solve: (x) => { x.app.select('A1'); x.app.setCell('A1', 'Produto'); x.app.select('A2'); },
      },
      {
        title: 'Agora com Tab',
        task: 'Na célula <b>B1</b>, digite <b>Preço</b> e confirme.',
        hint: 'Clique em B1, digite Preço e tecle Enter. (Se estivesse em A1, bastaria usar Tab depois de digitar para ir direto a B1.)',
        target: 'cell:B1',
        check: (x) => x.text('B1') === 'PRECO',
        solve: (x) => { x.app.setCell('B1', 'Preço'); x.app.select('B2'); },
        success: 'Você acaba de criar os <b>cabeçalhos</b> de uma tabela. Toda boa tabela começa com eles.',
      },
      {
        title: 'Números com vírgula',
        text: `<p>No Excel em português, a parte decimal usa <b>vírgula</b>: <code>2,50</code>. Se você digitar com ponto (<code>2.50</code>), o Excel pode não entender como número.</p>`,
        task: 'Em <b>A2</b> digite <b>Caneta</b> e em <b>B2</b> digite <b>2,50</b>.',
        hint: 'Clique em A2, digite Caneta e aperte <kbd>Tab</kbd> — você vai parar em B2. Digite 2,50 e aperte <kbd>Enter</kbd>.',
        target: 'cell:A2:B2',
        check: (x) => x.text('A2') === 'CANETA' && x.near(x.num('B2'), 2.5),
        solve: (x) => { x.app.setCell('A2', 'Caneta'); x.app.setCell('B2', '2,5'); x.app.select('A3'); },
        success: 'Repare que o Excel mostra <b>2,5</b>: ele guarda o número, não os zeros à direita. Vamos formatar como dinheiro numa lição futura.',
      },
      {
        title: 'Complete a lista',
        task: 'Preencha: <b>A3</b> = Caderno, <b>B3</b> = 15, <b>A4</b> = Mochila, <b>B4</b> = 89,90.',
        hint: 'Use <kbd>Tab</kbd> para ir do produto para o preço e <kbd>Enter</kbd> no fim da linha. O Excel volta para a coluna onde você começou!',
        target: 'cell:A3:B4',
        check: (x) => x.text('A3') === 'CADERNO' && x.num('B3') === 15 && x.text('A4') === 'MOCHILA' && x.near(x.num('B4'), 89.9),
        solve: (x) => { x.set({ A3: 'Caderno', B3: 15, A4: 'Mochila', B4: 89.9 }); x.app.setCell('B4', '89,9'); x.app.select('A5'); },
      },
      {
        title: 'Texto à esquerda, número à direita',
        text: `<p>Olhe a coluna B: os preços ficaram <b>alinhados à direita</b>, e os nomes na coluna A ficaram <b>à esquerda</b>.</p>
          <p>Esse é um truque para conferir se o Excel entendeu o que você digitou: <b>se um número ficar à esquerda, o Excel o tratou como texto</b> e ele não vai entrar nas contas.</p>`,
        target: 'cell:B2:B4',
      },
      {
        title: 'Editar sem apagar tudo',
        text: `<p>Se você selecionar uma célula e começar a digitar, o conteúdo antigo é <b>substituído</b>.</p>
          <p>Para só <b>ajustar</b> o conteúdo, use <kbd>F2</kbd> ou dê um <b>duplo clique</b> na célula: o cursor aparece no fim do texto.</p>`,
        task: 'Mude o texto de <b>A2</b> para <b>Caneta azul</b> usando <kbd>F2</kbd> ou duplo clique.',
        hint: 'Selecione A2, aperte <kbd>F2</kbd>, digite um espaço e a palavra azul, e tecle <kbd>Enter</kbd>.',
        target: 'cell:A2',
        check: (x) => x.text('A2') === 'CANETA AZUL',
        solve: (x) => { x.app.setCell('A2', 'Caneta azul'); x.app.select('A3'); },
      },
      {
        title: 'Apagar o conteúdo',
        task: 'Selecione <b>B4</b> e apague o valor com a tecla <kbd>Delete</kbd>.',
        hint: 'Clique em B4 e aperte <kbd>Delete</kbd> (ou <kbd>Del</kbd>).',
        target: 'cell:B4',
        check: (x) => x.empty('B4'),
        solve: (x) => { x.app.select('B4'); x.app.exec('clear', 'contents'); },
        success: 'O valor sumiu. Mas calma, dá para voltar atrás...',
      },
      {
        title: 'Desfazer',
        text: '<p>Errou? Sem problema. <kbd>Ctrl</kbd>+<kbd>Z</kbd> <b>desfaz</b> a última ação. Pode apertar várias vezes para voltar vários passos. <kbd>Ctrl</kbd>+<kbd>Y</kbd> refaz.</p>',
        task: 'Pressione <kbd>Ctrl</kbd>+<kbd>Z</kbd> para trazer de volta o preço da mochila.',
        hint: 'Segure a tecla <kbd>Ctrl</kbd> e aperte <kbd>Z</kbd>.',
        target: 'cell:B4',
        check: (x) => x.happened('undo') && x.near(x.num('B4'), 89.9),
        solve: (x) => x.app.undo(),
        success: 'O <b>89,9</b> voltou! <kbd>Ctrl</kbd>+<kbd>Z</kbd> vai ser seu melhor amigo.',
      },
    ],
  });

  // ======================================================================
  // 1.3 Sua primeira fórmula
  // ======================================================================
  const dadosCompra = (x) => {
    x.set({
      A3: { v: 'Quantidade' }, B3: 4,
      A4: { v: 'Preço unitário' }, B4: 12.5,
      A5: { v: 'Total', fmt: bold },
    });
    x.widths({ A: 120 });
  };
  XL.registerLesson({
    id: 'b1-3',
    title: 'Sua primeira fórmula',
    summary: [
      'Toda fórmula começa com <b>=</b>. A célula mostra o <b>resultado</b>; a Barra de Fórmulas mostra a <b>fórmula</b>.',
      'Operadores: <code>+</code> <code>-</code> <code>*</code> <code>/</code> <code>^</code>. Parênteses vêm primeiro.',
      'Use <b>referências</b> (=B3*B4) em vez de números fixos: o resultado se atualiza sozinho.',
      'Erros começam com <b>#</b>: #DIV/0!, #NOME?, #VALOR!, #REF!.',
    ],
    steps: [
      {
        title: 'O segredo: o sinal de igual',
        text: `<p>Fórmulas são o que tornam o Excel poderoso. Toda fórmula começa com <b>=</b> (igual). Depois vem a conta:</p>
          <table class="mini"><tr><th>Operador</th><th>Faz</th><th>Exemplo</th></tr>
          <tr><td><code>+</code></td><td>soma</td><td><code>=10+5</code></td></tr>
          <tr><td><code>-</code></td><td>subtração</td><td><code>=10-5</code></td></tr>
          <tr><td><code>*</code></td><td>multiplicação</td><td><code>=10*5</code></td></tr>
          <tr><td><code>/</code></td><td>divisão</td><td><code>=10/5</code></td></tr>
          <tr><td><code>^</code></td><td>potência</td><td><code>=10^2</code></td></tr></table>`,
      },
      {
        title: 'Faça uma conta',
        task: 'Na célula <b>A1</b>, digite <code>=10+5</code> e pressione <kbd>Enter</kbd>.',
        hint: 'Clique em A1 e digite exatamente: sinal de igual, 10, sinal de mais, 5. Depois <kbd>Enter</kbd>.',
        target: 'cell:A1',
        check: (x) => x.isFormula('A1') && x.near(x.num('A1'), 15),
        solve: (x) => { x.app.setCell('A1', '=10+5'); x.app.select('A2'); },
        success: 'A célula mostra <b>15</b>. O Excel calculou para você!',
      },
      {
        title: 'Resultado x fórmula',
        task: 'Clique de volta na célula <b>A1</b> e olhe a <b>Barra de Fórmulas</b>.',
        target: ['cell:A1', '.xl-finput'],
        check: (x) => x.active() === 'A1' && x.happened('select'),
        solve: (x) => x.app.select('A1'),
        success: 'A <b>célula</b> mostra o resultado (15), mas a <b>Barra de Fórmulas</b> mostra a fórmula (<code>=10+5</code>). Sempre olhe a Barra de Fórmulas para saber o que há "por trás" de um número.',
      },
      {
        title: 'Fórmulas com referências',
        setup: dadosCompra,
        text: `<p>Colocar números fixos na fórmula funciona, mas o jeito profissional é usar <b>referências de células</b>.</p>
          <p>Em vez de <code>=4*12,5</code>, escrevemos <code>=B3*B4</code>: "multiplique o que está em B3 pelo que está em B4". Se os valores mudarem, o resultado muda junto!</p>
          <p>Preparamos uma pequena compra nas linhas 3 a 5.</p>`,
        target: 'cell:A3:B5',
      },
      {
        title: 'Calcule o total',
        setup: (x) => { if (x.empty('B3')) dadosCompra(x); },
        task: 'Em <b>B5</b>, digite <code>=B3*B4</code> e pressione <kbd>Enter</kbd>.',
        hint: 'Truque: digite <code>=</code>, <b>clique</b> em B3, digite <code>*</code>, <b>clique</b> em B4 e aperte <kbd>Enter</kbd>. Clicar nas células insere a referência (modo <b>Aponte</b>) e evita erros de digitação.',
        target: 'cell:B5',
        check: (x) => x.isFormula('B5') && x.refs('B5').includes('B3') && x.refs('B5').includes('B4') && x.near(x.num('B5'), 50),
        solve: (x) => { x.app.setCell('B5', '=B3*B4'); x.app.select('B6'); },
        success: 'Total: <b>50</b>. Enquanto você digitava, cada referência ganhou uma <b>cor</b> e a célula correspondente ficou contornada com a mesma cor. Isso ajuda a conferir a fórmula.',
      },
      {
        title: 'A mágica do recálculo',
        task: 'Mude a quantidade em <b>B3</b> para <b>10</b> e veja o total.',
        hint: 'Clique em B3, digite 10 e tecle <kbd>Enter</kbd>. Não precisa mexer em B5!',
        target: 'cell:B3',
        check: (x) => x.num('B3') === 10 && x.near(x.num('B5'), 125),
        solve: (x) => { x.app.setCell('B3', '10'); x.app.select('B4'); },
        success: 'O total virou <b>125</b> sozinho. É por isso que usamos referências: a planilha se atualiza automaticamente.',
      },
      {
        title: 'Ordem das operações',
        text: `<p>O Excel segue a ordem da matemática: primeiro <code>^</code>, depois <code>*</code> e <code>/</code>, e por último <code>+</code> e <code>-</code>.</p>
          <p><b>Parênteses</b> mudam a ordem: o que está dentro deles é calculado primeiro.</p>
          <p><code>=2+3*4</code> → 14 &nbsp;&nbsp;|&nbsp;&nbsp; <code>=(2+3)*4</code> → 20</p>`,
      },
      {
        title: 'Use parênteses',
        task: 'Em <b>D1</b>, digite <code>=(2+3)*4</code>.',
        target: 'cell:D1',
        check: (x) => x.isFormula('D1') && x.formula('D1').includes('(') && x.near(x.num('D1'), 20),
        solve: (x) => { x.app.setCell('D1', '=(2+3)*4'); x.app.select('D2'); },
      },
      {
        title: 'Agora sem parênteses',
        task: 'Em <b>D2</b>, digite <code>=2+3*4</code> e compare com D1.',
        target: 'cell:D2',
        check: (x) => x.isFormula('D2') && x.near(x.num('D2'), 14),
        solve: (x) => { x.app.setCell('D2', '=2+3*4'); x.app.select('D3'); },
        success: 'Sem parênteses o Excel faz 3×4 primeiro (12) e depois soma 2: <b>14</b>. Com parênteses: <b>20</b>.',
      },
      {
        title: 'Conhecendo os erros',
        text: '<p>Quando algo dá errado, o Excel mostra um código começando com <b>#</b>. Não é o fim do mundo — é uma pista!</p>',
        task: 'Em <b>D3</b>, digite <code>=10/0</code> para ver um erro de propósito.',
        target: 'cell:D3',
        check: (x) => x.isErr('D3', '#DIV/0!'),
        solve: (x) => { x.app.setCell('D3', '=10/0'); x.app.select('D4'); },
        success: `Não existe divisão por zero, então o Excel mostra <b>#DIV/0!</b>. Os erros mais comuns:
          <ul><li><b>#DIV/0!</b> — divisão por zero (ou por célula vazia)</li>
          <li><b>#NOME?</b> — nome de função digitado errado</li>
          <li><b>#VALOR!</b> — conta com texto onde deveria haver número</li>
          <li><b>#REF!</b> — referência a uma célula que não existe mais</li></ul>`,
      },
    ],
  });

  // ======================================================================
  // 1.4 Funções essenciais
  // ======================================================================
  XL.registerLesson({
    id: 'b1-4',
    title: 'Funções essenciais',
    summary: [
      'Função = fórmula pronta: <code>=NOME(argumentos)</code>. Argumentos são separados por <b>;</b>.',
      '<code>SOMA</code>, <code>MÉDIA</code>, <code>MÁXIMO</code>, <code>MÍNIMO</code> e <code>CONT.NÚM</code> trabalham com intervalos como <code>B2:B7</code>.',
      'O botão <b>Σ AutoSoma</b> (ou <kbd>Alt</kbd>+<kbd>=</kbd>) monta a soma sozinho.',
      'Selecionou números? A <b>barra de status</b> mostra Média, Contagem e Soma na hora.',
    ],
    setup: (x) => {
      const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun'];
      const vendas = [12500, 15800, 13200, 17900, 16400, 19300];
      const metas = [14000, 14000, 15000, 15000, 16000, 16000];
      const m = { A1: { v: 'Mês', fmt: bold }, B1: { v: 'Vendas', fmt: bold }, C1: { v: 'Meta', fmt: bold } };
      meses.forEach((mes, i) => { m['A' + (i + 2)] = mes; m['B' + (i + 2)] = vendas[i]; m['C' + (i + 2)] = metas[i]; });
      Object.assign(m, {
        A8: { v: 'Total', fmt: bold }, A9: { v: 'Média', fmt: bold }, A10: { v: 'Maior venda', fmt: bold }, A11: { v: 'Menor venda', fmt: bold },
        A13: { v: 'Qtde de meses', fmt: bold },
      });
      x.set(m);
      x.widths({ A: 115 });
    },
    steps: [
      {
        title: 'O que é uma função?',
        text: `<p>Uma <b>função</b> é uma fórmula pronta que o Excel já sabe fazer. Em vez de <code>=B2+B3+B4+B5+B6+B7</code>, você escreve:</p>
          <p style="text-align:center"><code class="big">=SOMA(B2:B7)</code></p>
          <ul><li><b>SOMA</b> é o nome da função</li><li>Entre parênteses vão os <b>argumentos</b> — aqui, o intervalo <b>B2:B7</b></li>
          <li>Se houver mais de um argumento, separe com <b>ponto e vírgula</b> (<code>;</code>) — é o padrão do Excel em português</li></ul>
          <p>Montamos uma planilha de vendas do semestre para você praticar.</p>`,
        target: 'cell:A1:C7',
      },
      {
        title: 'SOMA',
        task: 'Em <b>B8</b>, calcule o total de vendas: <code>=SOMA(B2:B7)</code>.',
        hint: 'Digite <code>=SO</code> e veja a lista de funções aparecer — aperte <kbd>Tab</kbd> para escolher SOMA. Depois <b>arraste</b> de B2 até B7 com o mouse, feche o parêntese e tecle <kbd>Enter</kbd>.',
        target: 'cell:B8',
        check: (x) => x.uses('B8', 'SOMA') && x.near(x.num('B8'), 95100),
        solve: (x) => { x.app.setCell('B8', '=SOMA(B2:B7)'); x.app.select('B9'); },
        success: 'Total do semestre: <b>95100</b>. Percebeu a <b>dica amarela</b> mostrando <code>SOMA(núm1; [núm2]; ...)</code> enquanto você digitava? Argumentos entre colchetes são opcionais.',
      },
      {
        title: 'MÉDIA',
        task: 'Em <b>B9</b>, calcule a média mensal: <code>=MÉDIA(B2:B7)</code>.',
        hint: 'Mesmo esquema: <code>=MÉDIA(</code>, selecione B2:B7, <kbd>Enter</kbd>. (Se esquecer o acento, o simulador corrige — mas no Excel de verdade o acento é obrigatório.)',
        target: 'cell:B9',
        check: (x) => x.uses('B9', 'MÉDIA') && x.near(x.num('B9'), 15850),
        solve: (x) => { x.app.setCell('B9', '=MÉDIA(B2:B7)'); x.app.select('B10'); },
      },
      {
        title: 'MÁXIMO',
        task: 'Em <b>B10</b>, encontre a maior venda: <code>=MÁXIMO(B2:B7)</code>.',
        target: 'cell:B10',
        check: (x) => x.uses('B10', 'MÁXIMO') && x.near(x.num('B10'), 19300),
        solve: (x) => { x.app.setCell('B10', '=MÁXIMO(B2:B7)'); x.app.select('B11'); },
      },
      {
        title: 'MÍNIMO',
        task: 'Em <b>B11</b>, encontre a menor venda: <code>=MÍNIMO(B2:B7)</code>.',
        target: 'cell:B11',
        check: (x) => x.uses('B11', 'MÍNIMO') && x.near(x.num('B11'), 12500),
        solve: (x) => { x.app.setCell('B11', '=MÍNIMO(B2:B7)'); x.app.select('B12'); },
        success: 'Total, média, maior e menor: o resumo clássico de qualquer relatório.',
      },
      {
        title: 'Contando números',
        text: '<p><code>CONT.NÚM</code> conta quantas células de um intervalo têm <b>números</b>. Útil para saber quantos registros existem.</p>',
        task: 'Em <b>B13</b>, conte quantos meses têm vendas: <code>=CONT.NÚM(B2:B7)</code>.',
        target: 'cell:B13',
        check: (x) => x.uses('B13', 'CONT.NÚM') && x.num('B13') === 6,
        solve: (x) => { x.app.setCell('B13', '=CONT.NÚM(B2:B7)'); x.app.select('B14'); },
      },
      {
        title: 'O atalho da barra de status',
        text: '<p>Quer só <b>consultar</b> uma soma rápida, sem criar fórmula? Selecione os números e olhe o canto inferior direito da tela.</p>',
        task: 'Selecione as metas de <b>C2</b> até <b>C7</b> e olhe a barra de status.',
        target: ['cell:C2:C7', '.xl-status-stats'],
        check: (x) => x.sel() === 'C2:C7',
        solve: (x) => x.app.select('C2:C7'),
        success: 'A barra de status mostra <b>Média</b>, <b>Contagem</b> e <b>Soma</b> (90.000) da seleção — sem escrever nada na planilha.',
      },
      {
        title: 'AutoSoma: o botão mágico',
        text: '<p>O botão <b>Σ AutoSoma</b>, na guia Página Inicial, escreve a função SOMA para você e já sugere o intervalo de números logo acima.</p>',
        task: 'Selecione <b>C8</b>, clique em <b>Σ AutoSoma</b> e confirme com <kbd>Enter</kbd>.',
        hint: 'O botão Σ fica no grupo <b>Edição</b>, à direita da guia Página Inicial. Ao clicar, o Excel escreve <code>=SOMA(C2:C7)</code> — só confirme com <kbd>Enter</kbd>. Atalho: <kbd>Alt</kbd>+<kbd>=</kbd>.',
        target: ['cell:C8', '.rb-btn[data-cmd="autosum"]'],
        check: (x) => x.uses('C8', 'SOMA') && x.near(x.num('C8'), 90000) && x.happened('autosum'),
        solve: (x) => { x.app.select('C8'); x.app.exec('autosum', 'SOMA'); x.app.grid.commit(1, 0); },
        success: 'A meta do semestre era <b>90.000</b> e as vendas foram <b>95.100</b>: meta batida! 🎯 A setinha ao lado do Σ ainda oferece Média, Contar Números, Máx. e Mín.',
      },
    ],
  });

  // ======================================================================
  // 1.5 Formatando a planilha
  // ======================================================================
  XL.registerLesson({
    id: 'b1-5',
    title: 'Formatando a planilha',
    summary: [
      'Formatar muda a <b>aparência</b>, nunca o <b>valor</b> guardado na célula.',
      'Negrito, cor de preenchimento e alinhamento deixam cabeçalhos claros.',
      'Use formatos de <b>Moeda/Contábil</b> e <b>Porcentagem</b> em vez de digitar "R$" ou "%".',
      'Bordas organizam visualmente a tabela.',
    ],
    setup: (x) => {
      x.set({
        A1: 'Controle de Gastos Mensais',
        A3: 'Categoria', B3: 'Valor', C3: '% do total',
        A4: 'Aluguel', B4: 1500, A5: 'Mercado', B5: 850, A6: 'Transporte', B6: 320, A7: 'Lazer', B7: 280,
        A8: 'Total', B8: '=SOMA(B4:B7)',
        C4: '=B4/$B$8', C5: '=B5/$B$8', C6: '=B6/$B$8', C7: '=B7/$B$8', C8: '=SOMA(C4:C7)',
      });
      x.widths({ A: 130, B: 100, C: 95 });
    },
    steps: [
      {
        title: 'Uma planilha sem graça',
        text: `<p>Esta planilha de gastos já tem as fórmulas prontas, mas está difícil de ler: os valores não parecem dinheiro e a coluna C mostra números como <b>0,508474576</b>.</p>
          <p>Vamos formatá-la. Lembre-se: <b>formatar muda só a aparência</b> — o valor guardado continua o mesmo.</p>`,
        target: 'cell:A1:C8',
      },
      {
        title: 'Cabeçalho em negrito',
        task: 'Selecione <b>A3:C3</b> e clique em <b>Negrito</b> (o botão <b>N</b>).',
        hint: 'Arraste de A3 até C3. O botão <b>N</b> fica no grupo Fonte. Atalho: <kbd>Ctrl</kbd>+<kbd>N</kbd>.',
        target: ['cell:A3:C3', '.rb-btn[data-cmd="bold"]'],
        check: (x) => x.all('A3:C3', (r) => x.fmt(r).b),
        solve: (x) => { x.app.select('A3:C3'); x.app.exec('bold'); },
      },
      {
        title: 'Cor de preenchimento',
        task: 'Com <b>A3:C3</b> selecionado, aplique uma <b>Cor de Preenchimento</b> (o balde de tinta). Escolha a cor que quiser na setinha ao lado.',
        hint: 'Clique na <b>setinha</b> ao lado do balde para abrir as cores e escolha uma clarinha, como um verde ou azul claro. Clicar no próprio balde aplica a última cor usada.',
        target: ['cell:A3:C3', '.rb-btn[data-cmd="fill"]', '[data-menu="fill"]'],
        check: (x) => x.all('A3:C3', (r) => !!x.fmt(r).fill),
        solve: (x) => { x.app.select('A3:C3'); x.app.exec('fill', '#C6E0B4'); },
      },
      {
        title: 'Centralizar',
        task: 'Ainda com <b>A3:C3</b> selecionado, clique em <b>Centralizar</b>.',
        hint: 'É o botão com linhas centralizadas, no grupo Alinhamento (segunda linha).',
        target: ['cell:A3:C3', '.rb-btn[data-cmd="align"][data-arg="center"]'],
        check: (x) => x.all('A3:C3', (r) => x.fmt(r).align === 'center'),
        solve: (x) => { x.app.select('A3:C3'); x.app.exec('align', 'center'); },
      },
      {
        title: 'Formato de dinheiro',
        task: 'Selecione <b>B4:B8</b> e aplique o formato <b>Contábil</b> (botão da nota de dinheiro) ou escolha <b>Moeda</b> na caixa de formato de número.',
        hint: 'No grupo <b>Número</b>: o botão verde de dinheiro aplica o formato Contábil (R$). A caixa que diz "Geral" tem a opção Moeda.',
        target: ['cell:B4:B8', '.rb-btn[data-cmd="currency"]', 'select[data-cmd="numfmt"]'],
        check: (x) => x.all('B4:B8', (r) => ['moeda', 'contabil'].includes(x.fmt(r).numFmt)),
        solve: (x) => { x.app.select('B4:B8'); x.app.exec('currency'); },
        success: 'Agora os valores aparecem como <b>R$ 1.500,00</b>. Nunca digite "R$" junto com o número: use o formato!',
      },
      {
        title: 'Porcentagem',
        task: 'Selecione <b>C4:C8</b> e clique em <b>Estilo de Porcentagem</b> (<b>%</b>).',
        target: ['cell:C4:C8', '.rb-btn[data-cmd="percent"]'],
        check: (x) => x.all('C4:C8', (r) => x.fmt(r).numFmt === 'porcentagem'),
        solve: (x) => { x.app.select('C4:C8'); x.app.exec('percent'); },
        success: '<b>0,508474576</b> virou <b>51%</b>. É o mesmo valor, só que mais fácil de ler.',
      },
      {
        title: 'Casas decimais',
        task: 'Com <b>C4:C8</b> selecionado, clique uma vez em <b>Aumentar Casas Decimais</b>.',
        hint: 'É o botão com <b>←0 ,00</b>, no grupo Número.',
        target: ['cell:C4:C8', '.rb-btn[data-cmd="dec-inc"]'],
        check: (x) => x.all('C4:C8', (r) => x.fmt(r).numFmt === 'porcentagem' && (x.fmt(r).dec || 0) >= 1),
        solve: (x) => { x.app.select('C4:C8'); x.app.exec('dec-inc'); },
        success: 'Agora aparece <b>50,8%</b>. O botão ao lado diminui as casas.',
      },
      {
        title: 'Um título de verdade',
        task: 'Deixe o título em <b>A1</b> em <b>negrito</b> e com tamanho de fonte <b>16</b>.',
        hint: 'Selecione A1, clique em <b>N</b> e escolha 16 na caixa de tamanho (onde está escrito 11), no grupo Fonte.',
        target: ['cell:A1', '.rb-btn[data-cmd="bold"]', 'select[data-cmd="size"]'],
        check: (x) => x.fmt('A1').b && (x.fmt('A1').size || 11) >= 14,
        solve: (x) => { x.app.select('A1'); if (!x.fmt('A1').b) x.app.exec('bold'); x.app.exec('size', 16); },
        success: 'Repare que o texto "transborda" para as células vizinhas porque elas estão vazias — o conteúdo continua todo em A1.',
      },
      {
        title: 'Bordas',
        task: 'Selecione <b>A3:C8</b> e aplique <b>Todas as Bordas</b>.',
        hint: 'Clique na <b>setinha</b> ao lado do botão de Bordas (grupo Fonte) e escolha <b>Todas as Bordas</b>.',
        target: ['cell:A3:C8', '[data-menu="border"]', '.ui-menu-item[data-arg="all"]'],
        check: (x) => x.all('A3:C8', (r) => { const b = x.fmt(r).border || {}; return b.t && b.r && b.b && b.l; }),
        solve: (x) => { x.app.select('A3:C8'); x.app.exec('border', 'all'); },
      },
      {
        title: 'Aparência x valor',
        text: `<p>Clique em <b>C4</b> e compare: a célula mostra <b>50,8%</b>, mas a Barra de Fórmulas mostra <code>=B4/$B$8</code> — a fórmula e o valor real continuam intactos.</p>
          <p>Ficou curioso com o <b>$</b> nessa fórmula? É o assunto da próxima lição! 😉</p>`,
        target: ['cell:C4', '.xl-finput'],
      },
    ],
  });

  // ======================================================================
  // 1.6 Alça de preenchimento e referências absolutas
  // ======================================================================
  XL.registerLesson({
    id: 'b1-6',
    title: 'Alça de preenchimento e o $',
    summary: [
      'A <b>alça de preenchimento</b> (quadradinho no canto da seleção) copia fórmulas e continua sequências.',
      'Referências <b>relativas</b> (B2) se ajustam ao copiar; <b>absolutas</b> ($H$1) ficam travadas.',
      '<kbd>F4</kbd> coloca e tira os <b>$</b> enquanto você edita a fórmula.',
      'Duplo clique na alça preenche até o fim dos dados.',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Produto', fmt: bold }, B1: { v: 'Qtd', fmt: bold }, C1: { v: 'Preço', fmt: bold }, D1: { v: 'Total', fmt: bold }, E1: { v: 'Com desconto', fmt: bold },
        A2: 'Teclado', B2: 3, C2: 120, A3: 'Mouse', B3: 5, C3: 45, A4: 'Monitor', B4: 2, C4: 899,
        A5: 'Headset', B5: 4, C5: 210, A6: 'Webcam', B6: 1, C6: 350,
        G1: { v: 'Desconto', fmt: bold }, H1: { v: 0.1, fmt: { numFmt: 'porcentagem', fill: '#FFF2CC' } },
      });
      x.widths({ A: 95, E: 105 });
    },
    steps: [
      {
        title: 'Primeiro, uma fórmula',
        task: 'Em <b>D2</b>, calcule o total do teclado: <code>=B2*C2</code>.',
        target: 'cell:D2',
        check: (x) => x.isFormula('D2') && x.refs('D2').includes('B2') && x.refs('D2').includes('C2') && x.near(x.num('D2'), 360),
        solve: (x) => { x.app.setCell('D2', '=B2*C2'); x.app.select('D2'); },
      },
      {
        title: 'A alça de preenchimento',
        setup: (x) => x.app.select('D2'),
        text: `<p>Precisamos da mesma conta para os outros produtos. Em vez de digitar 4 vezes, use a <b>alça de preenchimento</b>: o <b>quadradinho verde</b> no canto inferior direito da célula selecionada.</p>
          <p>Ao passar o mouse sobre ele, o cursor vira uma <b>cruz fina (+)</b>.</p>`,
        target: '.xl-handle',
      },
      {
        title: 'Arraste para copiar',
        task: 'Selecione <b>D2</b>, clique na alça de preenchimento e <b>arraste até D6</b>.',
        hint: 'Clique em D2, posicione o mouse exatamente no quadradinho verde do canto (o cursor vira +), segure e arraste até D6. Solte o botão.',
        target: ['cell:D2:D6', '.xl-handle'],
        check: (x) => ['D3', 'D4', 'D5', 'D6'].every((r) => x.isFormula(r)) && x.near(x.num('D4'), 1798) && x.near(x.num('D6'), 350),
        solve: (x) => x.app.fill(XL.parseRange('D2'), XL.parseRange('D2:D6'), 'down'),
        success: 'Todos os totais calculados de uma vez! Dica: um <b>duplo clique</b> na alça faz o mesmo, preenchendo até onde houver dados ao lado.',
      },
      {
        title: 'O que aconteceu com a fórmula?',
        task: 'Clique em <b>D4</b> e olhe a fórmula na Barra de Fórmulas.',
        target: ['cell:D4', '.xl-finput'],
        check: (x) => x.active() === 'D4' && x.happened('select'),
        solve: (x) => x.app.select('D4'),
        success: 'A fórmula virou <code>=B4*C4</code>! Ao copiar duas linhas para baixo, o Excel ajustou as referências também duas linhas. Isso se chama <b>referência relativa</b> — é o comportamento padrão.',
      },
      {
        title: 'Quando o ajuste atrapalha',
        text: `<p>Agora queremos o valor <b>com desconto</b>, usando o percentual que está em <b>H1</b> (10%).</p>
          <p>Se escrevermos <code>=D2*(1-H1)</code> e arrastarmos, a linha de baixo vira <code>=D3*(1-H2)</code> — mas <b>H2 está vazia</b>! O desconto precisa ficar <b>travado</b> em H1.</p>
          <p>Para travar, use o <b>$</b>: <code>$H$1</code> é uma <b>referência absoluta</b> — ela não muda ao copiar.</p>`,
        target: 'cell:H1',
      },
      {
        title: 'Trave a referência com $',
        task: 'Em <b>E2</b>, digite <code>=D2*(1-$H$1)</code>.',
        hint: 'Atalho profissional: digite <code>=D2*(1-H1</code> e, com o cursor logo após H1, aperte <kbd>F4</kbd> — o Excel coloca os $ para você. Depois feche o parêntese e tecle <kbd>Enter</kbd>.',
        target: 'cell:E2',
        check: (x) => x.formula('E2').includes('$H$1') && x.near(x.num('E2'), 324),
        solve: (x) => { x.app.setCell('E2', '=D2*(1-$H$1)'); x.app.select('E2'); },
        success: 'O teclado com desconto sai por <b>324</b>.',
      },
      {
        title: 'Arraste de novo',
        task: 'Use a alça de preenchimento para copiar <b>E2</b> até <b>E6</b>.',
        target: ['cell:E2:E6', '.xl-handle'],
        check: (x) => x.isFormula('E3') && x.near(x.num('E3'), 202.5) && x.near(x.num('E6'), 315),
        solve: (x) => { x.app.select('E2'); x.app.fill(XL.parseRange('E2'), XL.parseRange('E2:E6'), 'down'); },
        success: 'Clique em E4 e confira: <code>=D4*(1-$H$1)</code>. O D mudou (relativo), o $H$1 ficou travado (absoluto). 🔒',
      },
      {
        title: 'Mude o desconto',
        task: 'Altere o desconto em <b>H1</b> para <b>15%</b>.',
        hint: 'Clique em H1, digite 15% e tecle <kbd>Enter</kbd>.',
        target: 'cell:H1',
        check: (x) => x.near(x.num('H1'), 0.15) && x.near(x.num('E6'), 297.5),
        solve: (x) => { x.app.setCell('H1', '15%'); x.app.select('H2'); },
        success: 'Todos os preços com desconto mudaram de uma vez. Centralizar parâmetros (como o desconto) numa célula só é uma das melhores práticas do Excel.',
      },
      {
        title: 'Sequências automáticas',
        text: '<p>A alça também <b>continua sequências</b>. Se você der dois exemplos, o Excel descobre o padrão.</p>',
        task: 'Digite <b>1</b> em <b>J1</b> e <b>2</b> em <b>J2</b>. Selecione <b>J1:J2</b> e arraste a alça até <b>J10</b>.',
        hint: 'Depois de digitar os dois números, selecione os dois juntos (arraste de J1 a J2) e só então use a alça.',
        target: 'cell:J1:J10',
        check: (x) => x.num('J5') === 5 && x.num('J10') === 10 && x.happened('fill'),
        solve: (x) => { x.set({ J1: 1, J2: 2 }); x.app.select('J1:J2'); x.app.fill(XL.parseRange('J1:J2'), XL.parseRange('J1:J10'), 'down'); },
        success: 'Numeração de 1 a 10 em segundos. Funciona com qualquer passo: 5, 10, 15...',
      },
      {
        title: 'Meses e dias da semana',
        task: 'Digite <b>Jan</b> em <b>L1</b> e arraste a alça até <b>L6</b>.',
        target: 'cell:L1:L6',
        check: (x) => x.text('L1') === 'JAN' && x.text('L6') === 'JUN',
        solve: (x) => { x.app.setCell('L1', 'Jan'); x.app.select('L1'); x.app.fill(XL.parseRange('L1'), XL.parseRange('L1:L6'), 'down'); },
        success: 'O Excel conhece meses e dias da semana (<b>Seg, Ter...</b>). Você concluiu o Módulo 1! 🏆',
      },
    ],
  });
})();
