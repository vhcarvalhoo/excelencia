/* Nível Intermediário · Módulo 3 — Lógica e condições (4 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };
  const head = { b: true, fill: '#DDEBF7' };
  const moeda = { numFmt: 'moeda' };

  // ======================================================================
  // 3.1 Função SE
  // ======================================================================
  const ALUNOS = [['Ana', 8.5], ['Bruno', 5], ['Carla', 7], ['Diego', 6.5], ['Elisa', 9], ['Fábio', 4], ['Gabi', 7], ['Hugo', 3.5]];
  XL.registerLesson({
    id: 'i3-1',
    title: 'Função SE',
    summary: [
      '<code>=SE(teste; valor_se_verdadeiro; valor_se_falso)</code>: a fórmula toma uma decisão.',
      'O teste usa comparações: <code>=</code> <code>&lt;&gt;</code> <code>&gt;</code> <code>&lt;</code> <code>&gt;=</code> <code>&lt;=</code> e resulta em VERDADEIRO ou FALSO.',
      'Texto vai <b>entre aspas</b> (<code>"Aprovado"</code>); números e referências, não.',
      'Use uma célula de parâmetro com <b>$</b> (<code>$F$1</code>) para mudar a regra sem reescrever fórmulas.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Aluno', fmt: head }, B1: { v: 'Nota', fmt: head }, C1: { v: 'Situação', fmt: head },
        E1: { v: 'Média mínima', fmt: bold }, F1: { v: 7, fmt: { fill: '#FFF2CC', b: true } } };
      ALUNOS.forEach(([n, nota], i) => { m['A' + (i + 2)] = n; m['B' + (i + 2)] = nota; });
      x.set(m);
      x.widths({ C: 100, E: 110 });
    },
    steps: [
      {
        title: 'Fórmulas que tomam decisões',
        text: `<p>Até agora as fórmulas só faziam contas. Com a função <b>SE</b>, a fórmula <b>decide</b> o que mostrar, como a gente faz no dia a dia:</p>
          <p style="text-align:center"><i>"<b>Se</b> a nota for maior ou igual a 7, escreva Aprovado; <b>senão</b>, escreva Reprovado."</i></p>
          <p style="text-align:center"><code class="big">=SE(B2&gt;=7; "Aprovado"; "Reprovado")</code></p>
          <ul><li>1º argumento: o <b>teste lógico</b> (a pergunta)</li><li>2º: o que mostrar se for <b>verdade</b></li><li>3º: o que mostrar se for <b>falso</b></li></ul>`,
        target: 'cell:A1:C9',
      },
      {
        title: 'Primeiro, só o teste',
        text: `<p>Todo SE começa com uma <b>comparação</b>. Ela sempre responde VERDADEIRO ou FALSO:</p>
          <table class="mini"><tr><th>Operador</th><th>Significa</th></tr>
          <tr><td><code>=</code></td><td>igual a</td></tr><tr><td><code>&lt;&gt;</code></td><td>diferente de</td></tr>
          <tr><td><code>&gt;</code> / <code>&lt;</code></td><td>maior / menor que</td></tr><tr><td><code>&gt;=</code> / <code>&lt;=</code></td><td>maior ou igual / menor ou igual</td></tr></table>`,
        task: 'Em <b>C2</b>, digite <code>=B2&gt;=7</code> e veja a resposta.',
        hint: 'Clique em C2 e digite: sinal de igual, B2, maior, igual, 7. Depois <kbd>Enter</kbd>.',
        target: 'cell:C2',
        check: (x) => x.isFormula('C2') && x.val('C2') === true,
        solve: (x) => { x.app.setCell('C2', '=B2>=7'); x.app.select('C2'); },
        success: 'A Ana tirou 8,5, então a resposta é <b>VERDADEIRO</b>. Esse é o "teste lógico" que vai dentro do SE.',
      },
      {
        title: 'Agora com SE',
        task: 'Troque a fórmula de <b>C2</b> por <code>=SE(B2&gt;=7;"Aprovado";"Reprovado")</code>.',
        hint: 'Selecione C2 e digite por cima. Atenção às <b>aspas</b> nos textos e ao <b>ponto e vírgula</b> entre os argumentos. Enquanto digita, a dica amarela mostra em qual argumento você está.',
        target: 'cell:C2',
        check: (x) => x.uses('C2', 'SE') && x.val('C2') === 'Aprovado',
        solve: (x) => { x.app.setCell('C2', '=SE(B2>=7;"Aprovado";"Reprovado")'); x.app.select('C2'); },
        success: 'O VERDADEIRO virou "Aprovado". Repare: texto vai <b>entre aspas</b>; o 7 e o B2, não.',
      },
      {
        title: 'Para a turma toda',
        task: 'Use a <b>alça de preenchimento</b> para copiar <b>C2</b> até <b>C9</b>.',
        target: ['cell:C2:C9', '.xl-handle'],
        check: (x) => x.val('C3') === 'Reprovado' && x.val('C4') === 'Aprovado' && x.val('C9') === 'Reprovado',
        solve: (x) => { x.app.select('C2'); x.app.fill(XL.parseRange('C2'), XL.parseRange('C2:C9'), 'down'); },
        success: 'A Carla tirou exatamente 7 e foi <b>aprovada</b>, porque o teste é <b>&gt;=</b> (maior <i>ou igual</i>). Com <code>&gt;</code> ela teria reprovado — escolha o operador com cuidado!',
      },
      {
        title: 'Erro comum: esquecer as aspas',
        text: `<p>Se você escrever <code>=SE(B2&gt;=7;Aprovado;Reprovado)</code> sem aspas, o Excel acha que <b>Aprovado</b> é o nome de alguma coisa e mostra <b>#NOME?</b>.</p>
          <p>Regra: <b>texto entre aspas, sempre</b>. Números, referências (B2) e funções ficam sem aspas.</p>`,
      },
      {
        title: 'Uma regra fácil de mudar',
        text: '<p>O 7 está "escondido" dentro de 8 fórmulas. Se a escola mudar a média, você teria que editar todas. O jeito profissional: deixar a regra numa célula (<b>F1</b>) e apontar para ela com <b>$</b>.</p>',
        task: 'Em <b>C2</b>, escreva <code>=SE(B2&gt;=$F$1;"Aprovado";"Reprovado")</code> e arraste até <b>C9</b>.',
        hint: 'Dica: ao digitar F1, pressione <kbd>F4</kbd> para virar $F$1. Depois use a alça de preenchimento de novo.',
        target: ['cell:C2:C9', 'cell:F1'],
        check: (x) => x.formula('C2').includes('$F$1') && x.formula('C9').includes('$F$1') && x.val('C9') === 'Reprovado',
        solve: (x) => { x.app.setCell('C2', '=SE(B2>=$F$1;"Aprovado";"Reprovado")'); x.app.select('C2'); x.app.fill(XL.parseRange('C2'), XL.parseRange('C2:C9'), 'down'); },
      },
      {
        title: 'Mude a regra',
        task: 'A escola baixou a média para <b>6</b>. Mude <b>F1</b> para <b>6</b>.',
        target: 'cell:F1',
        check: (x) => x.num('F1') === 6 && x.val('C5') === 'Aprovado',
        solve: (x) => { x.app.setCell('F1', '6'); x.app.select('F2'); },
        success: 'O Diego (6,5) passou a <b>Aprovado</b> sem você tocar em nenhuma fórmula. 🎓',
      },
    ],
  });

  // ======================================================================
  // 3.2 E, OU e NÃO
  // ======================================================================
  const EQUIPE = [['Ana', 12000, 9], ['Bruno', 8500, 10], ['Carla', 15800, 7], ['Diego', 10500, 5], ['Elisa', 9800, 8], ['Fábio', 16200, 9]];
  XL.registerLesson({
    id: 'i3-2',
    title: 'E, OU e NÃO',
    summary: [
      '<code>E(teste1; teste2; ...)</code> é VERDADEIRO só se <b>todos</b> os testes forem verdadeiros.',
      '<code>OU(teste1; teste2; ...)</code> é VERDADEIRO se <b>pelo menos um</b> for verdadeiro.',
      '<code>NÃO(teste)</code> inverte a resposta.',
      'Combine com SE: <code>=SE(E(B2&gt;=10000;C2&gt;=8);"Sim";"Não")</code>.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Vendedor', fmt: head }, B1: { v: 'Vendas', fmt: head }, C1: { v: 'Nota do cliente', fmt: head },
        D1: { v: 'Bônus?', fmt: head }, E1: { v: 'Destaque?', fmt: head }, F1: { v: 'Treinamento?', fmt: head },
      };
      EQUIPE.forEach(([n, v, nota], i) => { const r = i + 2; m['A' + r] = n; m['B' + r] = { v, fmt: moeda }; m['C' + r] = nota; });
      x.set(m);
      x.widths({ B: 100, C: 110, D: 80, E: 90, F: 105 });
    },
    steps: [
      {
        title: 'Mais de uma condição',
        text: `<p>Muitas regras de negócio têm <b>duas ou mais condições</b>:</p>
          <ul><li><b>Bônus</b>: vendeu pelo menos R$ 10.000 <b>E</b> teve nota do cliente 8 ou mais.</li>
          <li><b>Destaque</b>: vendeu R$ 15.000 ou mais <b>OU</b> tirou nota 10.</li>
          <li><b>Treinamento</b>: <b>NÃO</b> atingiu nota 6.</li></ul>
          <table class="mini"><tr><th>Função</th><th>É VERDADEIRO quando...</th></tr>
          <tr><td><code>E</code></td><td><b>todas</b> as condições são verdadeiras</td></tr>
          <tr><td><code>OU</code></td><td><b>pelo menos uma</b> é verdadeira</td></tr>
          <tr><td><code>NÃO</code></td><td>a condição é <b>falsa</b> (inverte)</td></tr></table>`,
        target: 'cell:A1:F7',
      },
      {
        title: 'Testando o E',
        task: 'Em <b>D2</b>, digite <code>=E(B2&gt;=10000;C2&gt;=8)</code>.',
        hint: 'As duas condições vão dentro do E, separadas por <b>;</b>.',
        target: 'cell:D2',
        check: (x) => x.uses('D2', 'E') && x.val('D2') === true,
        solve: (x) => { x.app.setCell('D2', '=E(B2>=10000;C2>=8)'); x.app.select('D2'); },
        success: 'Ana vendeu R$ 12.000 (≥ 10.000 ✔) e tirou 9 (≥ 8 ✔): as duas são verdadeiras, então <b>VERDADEIRO</b>.',
      },
      {
        title: 'E dentro do SE',
        text: '<p>O E vira o <b>teste lógico</b> de um SE. Assim a resposta fica amigável: "Sim" ou "Não".</p>',
        task: 'Em <b>D2</b>, escreva <code>=SE(E(B2&gt;=10000;C2&gt;=8);"Sim";"Não")</code> e arraste até <b>D7</b>.',
        hint: 'Repare nos parênteses: o E abre e fecha <b>dentro</b> do primeiro argumento do SE.',
        target: 'cell:D2:D7',
        check: (x) => x.uses('D2', 'E') && x.uses('D2', 'SE') && x.val('D2') === 'Sim' && x.val('D3') === 'Não' && x.val('D7') === 'Sim' && x.val('D6') === 'Não',
        solve: (x) => { x.app.setCell('D2', '=SE(E(B2>=10000;C2>=8);"Sim";"Não")'); x.app.select('D2'); x.app.fill(XL.parseRange('D2'), XL.parseRange('D2:D7'), 'down'); },
        success: 'Só Ana e Fábio ganham bônus. A Elisa tem nota 8, mas vendeu R$ 9.800 — com <b>E</b>, falhar em uma condição já é "Não".',
      },
      {
        title: 'Agora o OU',
        task: 'Em <b>E2</b>, escreva <code>=SE(OU(B2&gt;=15000;C2=10);"Sim";"Não")</code> e arraste até <b>E7</b>.',
        hint: 'Mesmo formato do anterior, trocando E por OU e as condições.',
        target: 'cell:E2:E7',
        check: (x) => x.uses('E2', 'OU') && x.val('E3') === 'Sim' && x.val('E4') === 'Sim' && x.val('E5') === 'Não' && x.val('E7') === 'Sim',
        solve: (x) => { x.app.setCell('E2', '=SE(OU(B2>=15000;C2=10);"Sim";"Não")'); x.app.select('E2'); x.app.fill(XL.parseRange('E2'), XL.parseRange('E2:E7'), 'down'); },
        success: 'O Bruno vendeu pouco, mas tirou nota 10 — com <b>OU</b>, uma condição verdadeira basta. A Carla é destaque pelas vendas.',
      },
      {
        title: 'O NÃO inverte',
        task: 'Em <b>F2</b>, escreva <code>=SE(NÃO(C2&gt;=6);"Sim";"Não")</code> e arraste até <b>F7</b>.',
        hint: '<code>NÃO(C2&gt;=6)</code> é verdadeiro quando a nota <b>não</b> chegou a 6.',
        target: 'cell:F2:F7',
        check: (x) => x.uses('F2', 'NÃO') && x.val('F5') === 'Sim' && x.val('F2') === 'Não' && x.val('F7') === 'Não',
        solve: (x) => { x.app.setCell('F2', '=SE(NÃO(C2>=6);"Sim";"Não")'); x.app.select('F2'); x.app.fill(XL.parseRange('F2'), XL.parseRange('F2:F7'), 'down'); },
        success: 'Só o Diego (nota 5) vai para o treinamento. Dica: <code>NÃO(C2&gt;=6)</code> é o mesmo que <code>C2&lt;6</code> — o NÃO brilha em testes mais complexos, como <code>NÃO(E(...))</code>.',
      },
    ],
  });

  // ======================================================================
  // 3.3 SE aninhado e SES
  // ======================================================================
  const NOTAS3 = [['Ana', 9.5], ['Bruno', 7], ['Carla', 8], ['Diego', 4.5], ['Elisa', 5], ['Fábio', 6], ['Gabi', 10], ['Hugo', 7.5]];
  const CONCEITOS = ['A', 'B', 'B', 'D', 'C', 'C', 'A', 'B'];
  XL.registerLesson({
    id: 'i3-3',
    title: 'SE aninhado e SES',
    summary: [
      'Para mais de duas respostas, coloque um <b>SE dentro do outro</b> (no lugar do "se falso").',
      'Teste do <b>maior para o menor</b>: a primeira condição verdadeira vence.',
      '<code>SES(teste1; valor1; teste2; valor2; ...)</code> faz o mesmo sem tantos parênteses.',
      'No SES, termine com <code>VERDADEIRO; "valor padrão"</code> para cobrir o "senão".',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Aluno', fmt: head }, B1: { v: 'Nota', fmt: head }, C1: { v: 'Conceito (SE)', fmt: head }, D1: { v: 'Conceito (SES)', fmt: head },
        F1: { v: 'Tabela de conceitos', fmt: bold }, F2: '9 a 10 → A', F3: '7 a 8,9 → B', F4: '5 a 6,9 → C', F5: 'abaixo de 5 → D',
      };
      NOTAS3.forEach(([n, nota], i) => { m['A' + (i + 2)] = n; m['B' + (i + 2)] = nota; });
      x.set(m);
      x.widths({ C: 115, D: 120, F: 140 });
    },
    steps: [
      {
        title: 'Quando há mais de duas respostas',
        text: `<p>O SE escolhe entre <b>duas</b> respostas. Mas aqui precisamos de <b>quatro</b> conceitos (A, B, C, D), conforme a tabela em F1.</p>
          <p>A solução clássica é colocar <b>outro SE no lugar do "se falso"</b>:</p>
          <p><i>"Se ≥ 9 → A; <b>senão</b>, se ≥ 7 → B; <b>senão</b>, se ≥ 5 → C; <b>senão</b> → D."</i></p>`,
        target: ['cell:F1:F5', 'cell:A1:B9'],
      },
      {
        title: 'Monte o SE aninhado',
        task: 'Em <b>C2</b>, escreva <code>=SE(B2&gt;=9;"A";SE(B2&gt;=7;"B";SE(B2&gt;=5;"C";"D")))</code> e arraste até <b>C9</b>.',
        hint: 'São 3 SEs, então no final vêm <b>3 parênteses</b> fechando: <code>)))</code>. Se esquecer algum, o Excel completa ao pressionar Enter.',
        target: 'cell:C2:C9',
        check: (x) => x.uses('C2', 'SE') && CONCEITOS.every((c, i) => x.val('C' + (i + 2)) === c),
        solve: (x) => { x.app.setCell('C2', '=SE(B2>=9;"A";SE(B2>=7;"B";SE(B2>=5;"C";"D")))'); x.app.select('C2'); x.app.fill(XL.parseRange('C2'), XL.parseRange('C2:C9'), 'down'); },
        success: 'A Gabi (10) e a Ana (9,5) ficaram com A; o Diego (4,5) ficou com D. O Excel testa na ordem e para na <b>primeira</b> condição verdadeira.',
      },
      {
        title: 'A ordem importa',
        text: `<p>Imagine começar pelo menor: <code>=SE(B2&gt;=5;"C";SE(B2&gt;=7;"B";...))</code>. A Gabi tirou 10, que <b>também é ≥ 5</b> — ela pararia logo no primeiro teste e ganharia <b>C</b>!</p>
          <p>Regra: com faixas do tipo "maior ou igual", teste <b>do maior para o menor</b>.</p>`,
      },
      {
        title: 'O jeito moderno: SES',
        text: `<p>No Excel 2019/365 existe a função <b>SES</b>, que lista pares <b>teste; resultado</b> sem precisar aninhar:</p>
          <p style="text-align:center"><code>=SES(B2&gt;=9;"A"; B2&gt;=7;"B"; B2&gt;=5;"C"; VERDADEIRO;"D")</code></p>
          <p>O último par, <code>VERDADEIRO;"D"</code>, é o "senão": como VERDADEIRO é sempre verdadeiro, ele pega quem sobrou.</p>`,
        task: 'Em <b>D2</b>, escreva a fórmula com <b>SES</b> acima e arraste até <b>D9</b>.',
        hint: 'Sem o <code>VERDADEIRO;"D"</code> no final, quem tirou menos de 5 receberia o erro <b>#N/D</b> (nenhuma condição foi verdadeira).',
        target: 'cell:D2:D9',
        check: (x) => x.uses('D2', 'SES') && CONCEITOS.every((c, i) => x.val('D' + (i + 2)) === c),
        solve: (x) => { x.app.setCell('D2', '=SES(B2>=9;"A";B2>=7;"B";B2>=5;"C";VERDADEIRO;"D")'); x.app.select('D2'); x.app.fill(XL.parseRange('D2'), XL.parseRange('D2:D9'), 'down'); },
        success: 'Mesmo resultado, fórmula mais fácil de ler. Use <b>SES</b> quando seu Excel tiver; o SE aninhado funciona em qualquer versão. Mais à frente, você verá como o <b>PROCV</b> com uma tabela de faixas resolve isso de forma ainda mais elegante.',
      },
    ],
  });

  // ======================================================================
  // 3.4 Tratando erros com SEERRO
  // ======================================================================
  XL.registerLesson({
    id: 'i3-4',
    title: 'Tratando erros com SEERRO',
    summary: [
      'Erros "se espalham": uma fórmula que usa uma célula com erro também dá erro.',
      '<code>=SEERRO(fórmula; valor_se_der_erro)</code> troca o erro por algo amigável.',
      'Use <code>""</code> (vazio) quando o valor não deve entrar em médias; <b>0</b> distorceria a conta.',
      'SEERRO esconde <b>qualquer</b> erro — use só onde o erro é esperado.',
    ],
    setup: (x) => {
      x.set({
        A1: { v: 'Produto', fmt: head }, B1: { v: 'Faturamento', fmt: head }, C1: { v: 'Qtd vendida', fmt: head }, D1: { v: 'Preço médio', fmt: head },
        A2: 'Caneta', B2: { v: 250, fmt: moeda }, C2: 100,
        A3: 'Caderno', B3: { v: 900, fmt: moeda }, C3: 60,
        A4: 'Mochila', B4: { v: 0, fmt: moeda }, C4: 0,
        A5: 'Estojo', B5: { v: 360, fmt: moeda }, C5: 24,
        A6: 'Agenda', B6: { v: 480, fmt: moeda }, C6: 'n/d',
        A7: 'Régua', B7: { v: 120, fmt: moeda }, C7: 80,
        A8: { v: 'Média dos preços', fmt: bold },
        D2: { v: null, fmt: moeda }, D3: { v: null, fmt: moeda }, D4: { v: null, fmt: moeda }, D5: { v: null, fmt: moeda },
        D6: { v: null, fmt: moeda }, D7: { v: null, fmt: moeda }, D8: { v: null, fmt: Object.assign({ b: true }, moeda) },
      });
      x.widths({ A: 130, B: 105, C: 95, D: 105 });
    },
    steps: [
      {
        title: 'Calcule o preço médio',
        text: '<p>Preço médio = faturamento ÷ quantidade. Mas olhe a <b>Mochila</b> (não vendeu nada) e a <b>Agenda</b> (quantidade "n/d").</p>',
        task: 'Em <b>D2</b>, digite <code>=B2/C2</code> e arraste até <b>D7</b>.',
        target: 'cell:D2:D7',
        check: (x) => x.isErr('D4', '#DIV/0!') && x.isErr('D6', '#VALOR!') && x.near(x.num('D3'), 15),
        solve: (x) => { x.app.setCell('D2', '=B2/C2'); x.app.select('D2'); x.app.fill(XL.parseRange('D2'), XL.parseRange('D2:D7'), 'down'); },
        success: 'Apareceram dois erros: <b>#DIV/0!</b> (dividir por zero) e <b>#VALOR!</b> (dividir por um texto). Eles são esperados aqui — não há preço médio para quem não vendeu.',
      },
      {
        title: 'O erro se espalha',
        task: 'Em <b>D8</b>, calcule a média: <code>=MÉDIA(D2:D7)</code>.',
        target: 'cell:D8',
        check: (x) => x.uses('D8', 'MÉDIA') && x.isErr('D8'),
        solve: (x) => { x.app.setCell('D8', '=MÉDIA(D2:D7)'); x.app.select('D8'); },
        success: 'A média também virou erro! Qualquer conta que usa uma célula com erro "herda" o erro. Um único #DIV/0! pode estragar um relatório inteiro.',
      },
      {
        title: 'SEERRO ao resgate',
        text: `<p><code>SEERRO</code> tenta calcular; se der erro, mostra outra coisa:</p>
          <p style="text-align:center"><code class="big">=SEERRO(B2/C2; "")</code></p>
          <p>As aspas vazias <code>""</code> deixam a célula em branco quando a conta falha.</p>`,
        task: 'Em <b>D2</b>, escreva <code>=SEERRO(B2/C2;"")</code> e arraste até <b>D7</b>.',
        target: 'cell:D2:D7',
        check: (x) => x.uses('D2', 'SEERRO') && x.val('D4') === '' && x.val('D6') === '' && x.near(x.num('D8'), 8.5),
        solve: (x) => { x.app.setCell('D2', '=SEERRO(B2/C2;"")'); x.app.select('D2'); x.app.fill(XL.parseRange('D2'), XL.parseRange('D2:D7'), 'down'); },
        success: 'Os erros sumiram e a média em D8 voltou: <b>R$ 8,50</b>. A MÉDIA ignora os textos vazios.',
      },
      {
        title: 'Por que "" e não 0?',
        text: `<p>Se você usar <code>=SEERRO(B2/C2;0)</code>, a Mochila e a Agenda entram na média como <b>R$ 0,00</b> e a média cai para R$ 5,67 — um número <b>errado</b>, sem nenhum aviso.</p>
          <p>Pense no que o valor substituto significa: "sem preço" (<code>""</code>) é diferente de "preço zero" (<code>0</code>).</p>
          <p><b>Cuidado:</b> SEERRO esconde <i>qualquer</i> erro, inclusive de digitação (como #NOME?). Use só onde o erro é esperado. Para buscas com PROCV, existe ainda o <code>SENÃODISP</code>, que trata só o #N/D.</p>`,
      },
    ],
  });
})();
