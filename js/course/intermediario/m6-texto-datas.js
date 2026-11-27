/* Nível Intermediário · Módulo 6 — Texto e datas (4 lições). */
(function () {
  'use strict';
  const XL = window.XL;
  const bold = { b: true };
  const head = { b: true, fill: '#DDEBF7' };
  const data = { numFmt: 'data' };
  const fillDown = (x, from, to) => { x.app.select(from); x.app.fill(XL.parseRange(from), XL.parseRange(from + ':' + to), 'down'); };

  // ======================================================================
  // 6.1 Funções de texto
  // ======================================================================
  const NOMES = ['  ana   SOUZA ', 'BRUNO lima', 'carla DIAS  ', '  diego alves', 'ELISA ROCHA'];
  const ARRUMADOS = ['Ana Souza', 'Bruno Lima', 'Carla Dias', 'Diego Alves', 'Elisa Rocha'];
  const CODIGOS = ['SP-2026-0451', 'RJ-2025-0012', 'MG-2026-1033', 'PR-2024-0870', 'BA-2026-0215'];
  XL.registerLesson({
    id: 'i6-1',
    title: 'Funções de texto',
    summary: [
      '<code>ARRUMAR</code> tira espaços sobrando; <code>PRI.MAIÚSCULA</code>, <code>MAIÚSCULA</code> e <code>MINÚSCULA</code> padronizam letras.',
      '<code>ESQUERDA(texto; n)</code> e <code>DIREITA(texto; n)</code> pegam caracteres das pontas; <code>EXT.TEXTO(texto; início; n)</code>, do meio.',
      'O <code>&amp;</code> junta textos: <code>=C2&amp;" ("&amp;D2&amp;")"</code>.',
      'Funções de texto devolvem <b>texto</b>: use <code>VALOR</code> se precisar fazer conta com o resultado.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Nome (do sistema)', fmt: head }, B1: { v: 'Código', fmt: head }, C1: { v: 'Nome arrumado', fmt: head },
        D1: { v: 'UF', fmt: head }, E1: { v: 'Número', fmt: head }, F1: { v: 'Ano', fmt: head }, G1: { v: 'Rótulo', fmt: head },
      };
      NOMES.forEach((n, i) => { m['A' + (i + 2)] = n; m['B' + (i + 2)] = CODIGOS[i]; });
      x.set(m);
      x.widths({ A: 135, B: 105, C: 110, D: 45, E: 65, F: 55, G: 135 });
    },
    steps: [
      {
        title: 'Dados que vieram bagunçados',
        text: `<p>Dados exportados de sistemas costumam chegar assim: espaços sobrando, letras maiúsculas e minúsculas misturadas, informações grudadas num código só.</p>
          <p>O código <b>SP-2026-0451</b> tem três informações: a <b>UF</b> (SP), o <b>ano</b> (2026) e o <b>número</b> (0451). As funções de texto separam e arrumam tudo isso.</p>`,
        target: 'cell:A1:B6',
      },
      {
        title: 'Arrume os nomes',
        text: '<p><code>ARRUMAR</code> remove espaços extras (deixa só um entre as palavras) e <code>PRI.MAIÚSCULA</code> coloca a primeira letra de cada palavra em maiúscula. Uma função vai <b>dentro</b> da outra:</p>',
        task: 'Em <b>C2</b>, digite <code>=PRI.MAIÚSCULA(ARRUMAR(A2))</code> e arraste até <b>C6</b>.',
        target: 'cell:C2:C6',
        check: (x) => x.uses('C2', 'PRI.MAIÚSCULA') && x.uses('C2', 'ARRUMAR') && ARRUMADOS.every((n, i) => x.val('C' + (i + 2)) === n),
        solve: (x) => { x.app.setCell('C2', '=PRI.MAIÚSCULA(ARRUMAR(A2))'); fillDown(x, 'C2', 'C6'); },
        success: '"  ana   SOUZA " virou <b>Ana Souza</b>. Também existem <code>MAIÚSCULA</code> (TUDO GRANDE) e <code>MINÚSCULA</code> (tudo pequeno).',
      },
      {
        title: 'ESQUERDA: o começo do texto',
        task: 'Em <b>D2</b>, pegue os <b>2 primeiros</b> caracteres do código: <code>=ESQUERDA(B2;2)</code>, e arraste até <b>D6</b>.',
        target: 'cell:D2:D6',
        check: (x) => x.uses('D2', 'ESQUERDA') && x.val('D2') === 'SP' && x.val('D6') === 'BA',
        solve: (x) => { x.app.setCell('D2', '=ESQUERDA(B2;2)'); fillDown(x, 'D2', 'D6'); },
      },
      {
        title: 'DIREITA: o final do texto',
        task: 'Em <b>E2</b>, pegue os <b>4 últimos</b> caracteres: <code>=DIREITA(B2;4)</code>, e arraste até <b>E6</b>.',
        target: 'cell:E2:E6',
        check: (x) => x.uses('E2', 'DIREITA') && x.val('E2') === '0451' && x.val('E3') === '0012',
        solve: (x) => { x.app.setCell('E2', '=DIREITA(B2;4)'); fillDown(x, 'E2', 'E6'); },
        success: 'Repare que "0451" manteve o zero à esquerda e ficou alinhado à <b>esquerda</b>: o resultado é <b>texto</b>.',
      },
      {
        title: 'EXT.TEXTO: o meio do texto',
        text: '<p><code>EXT.TEXTO(texto; início; quantidade)</code> extrai um pedaço do meio. Em "SP-2026-0451", o ano começa no <b>4º</b> caractere e tem <b>4</b> caracteres.</p>',
        task: 'Em <b>F2</b>, digite <code>=EXT.TEXTO(B2;4;4)</code> e arraste até <b>F6</b>.',
        target: 'cell:F2:F6',
        check: (x) => x.uses('F2', 'EXT.TEXTO') && x.val('F2') === '2026' && x.val('F5') === '2024',
        solve: (x) => { x.app.setCell('F2', '=EXT.TEXTO(B2;4;4)'); fillDown(x, 'F2', 'F6'); },
        success: 'Os anos saíram como texto ("2026"). Para fazer contas com eles, envolva com <code>VALOR</code>: <code>=VALOR(EXT.TEXTO(B2;4;4))</code>.',
      },
      {
        title: 'Juntando textos com &',
        text: '<p>O <code>&amp;</code> "cola" textos. Espaços e parênteses fixos vão entre aspas:</p><p style="text-align:center"><code>=C2&amp;" ("&amp;D2&amp;")"</code> → Ana Souza (SP)</p>',
        task: 'Em <b>G2</b>, monte o rótulo <code>=C2&amp;" ("&amp;D2&amp;")"</code> e arraste até <b>G6</b>.',
        target: 'cell:G2:G6',
        check: (x) => x.val('G2') === 'Ana Souza (SP)' && x.val('G6') === 'Elisa Rocha (BA)',
        solve: (x) => { x.app.setCell('G2', '=C2&" ("&D2&")"'); fillDown(x, 'G2', 'G6'); },
        success: `Dados limpos e organizados! Outras funções úteis de texto:
          <ul><li><code>NÚM.CARACT</code> — quantos caracteres tem</li><li><code>SUBSTITUIR</code> — troca um pedaço do texto</li>
          <li><code>TEXTOANTES</code> / <code>TEXTODEPOIS</code> — o que vem antes/depois de um separador (ex.: o "@" de um e-mail)</li>
          <li><code>CONCAT</code> e <code>UNIRTEXTO</code> — juntam vários textos de uma vez</li></ul>`,
      },
    ],
  });

  // ======================================================================
  // 6.2 Datas e prazos
  // ======================================================================
  const PEDIDOS = [[1001, '03/11/2026', 10], [1002, '10/11/2026', 15], [1003, '20/11/2026', 7], [1004, '25/11/2026', 20], [1005, '27/11/2026', 5]];
  const workday = (s, n) => { while (n > 0) { s++; const wd = XL.serialToParts(s).wd; if (wd !== 0 && wd !== 6) n--; } return s; };
  XL.registerLesson({
    id: 'i6-2',
    title: 'Datas e prazos',
    summary: [
      'Para o Excel, uma data é um <b>número</b> (dias desde 1900). Por isso dá para somar e subtrair datas.',
      '<code>=B2+10</code> é 10 dias corridos depois; <code>=DIATRABALHO(B2;10)</code> pula sábados e domingos (e feriados, se informados).',
      'Data − data = número de dias entre elas.',
      '<code>HOJE()</code> sempre mostra a data atual. Outras úteis: <code>DIA</code>, <code>MÊS</code>, <code>ANO</code>, <code>DATADIF</code>, <code>FIMMÊS</code>.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Pedido', fmt: head }, B1: { v: 'Data do pedido', fmt: head }, C1: { v: 'Prazo (dias)', fmt: head },
        D1: { v: 'Entrega (corridos)', fmt: head }, E1: { v: 'Entrega (dias úteis)', fmt: head }, F1: { v: 'Faltam (dias)', fmt: head },
        H1: { v: 'Hoje:', fmt: bold }, I1: { v: '27/11/2026', fmt: Object.assign({ fill: '#FFF2CC', b: true }, data) },
      };
      PEDIDOS.forEach(([p, d, prazo], i) => {
        const r = i + 2;
        m['A' + r] = p; m['B' + r] = { v: d, fmt: data }; m['C' + r] = prazo;
        m['D' + r] = { v: null, fmt: data }; m['E' + r] = { v: null, fmt: data };
      });
      x.set(m);
      x.widths({ B: 110, C: 90, D: 125, E: 130, F: 95, G: 24, H: 50, I: 100 });
    },
    steps: [
      {
        title: 'O segredo das datas',
        text: `<p>Para o Excel, uma data é só um <b>número</b>: a quantidade de dias desde 1º de janeiro de 1900. O formato de data apenas "fantasia" esse número.</p>
          <p>Vamos tirar a fantasia para ver.</p>`,
        task: 'Selecione <b>B2</b> e mude o formato de número para <b>Geral</b> (caixa de formato, guia Página Inicial).',
        hint: 'Na guia Página Inicial, grupo Número, abra a caixa onde está escrito "Data Abreviada" e escolha <b>Geral</b>.',
        target: ['cell:B2', 'select[data-cmd="numfmt"]'],
        check: (x) => !x.fmt('B2').numFmt && typeof x.val('B2') === 'number',
        solve: (x) => { x.app.select('B2'); x.app.exec('numfmt', 'geral'); },
        success: 'A data virou <b>46329</b>: é o número de dias desde 1900. É por isso que dá para fazer contas com datas!',
      },
      {
        title: 'Volte para data',
        task: 'Coloque <b>B2</b> de volta no formato <b>Data Abreviada</b>.',
        hint: 'Mesma caixa, opção "Data Abreviada". Ou desfaça com <kbd>Ctrl</kbd>+<kbd>Z</kbd>.',
        target: ['cell:B2', 'select[data-cmd="numfmt"]'],
        check: (x) => x.fmt('B2').numFmt === 'data',
        solve: (x) => { x.app.select('B2'); x.app.exec('numfmt', 'data'); },
      },
      {
        title: 'Data + dias',
        task: 'Em <b>D2</b>, calcule a entrega em dias corridos: <code>=B2+C2</code>, e arraste até <b>D6</b>.',
        target: 'cell:D2:D6',
        check: (x) => x.isFormula('D2') && x.num('D2') === x.num('B2') + 10 && x.num('D6') === x.num('B6') + 5,
        solve: (x) => { x.app.setCell('D2', '=B2+C2'); fillDown(x, 'D2', 'D6'); },
        success: '03/11 + 10 dias = <b>13/11/2026</b>. Simples assim, porque datas são números.',
      },
      {
        title: 'Só dias úteis',
        text: '<p>Prazos de empresa costumam contar só <b>dias úteis</b>. <code>DIATRABALHO(data; dias)</code> pula sábados e domingos. Um 3º argumento opcional recebe uma lista de <b>feriados</b>.</p>',
        task: 'Em <b>E2</b>, digite <code>=DIATRABALHO(B2;C2)</code> e arraste até <b>E6</b>.',
        target: 'cell:E2:E6',
        check: (x) => x.uses('E2', 'DIATRABALHO') && [2, 3, 4, 5, 6].every((r) => x.num('E' + r) === workday(x.num('B' + r), x.num('C' + r))),
        solve: (x) => { x.app.setCell('E2', '=DIATRABALHO(B2;C2)'); fillDown(x, 'E2', 'E6'); },
        success: 'Compare as colunas D e E: em dias úteis a entrega fica mais longe, porque os fins de semana não contam. Para <b>contar</b> dias úteis entre duas datas, use <code>DIATRABALHOTOTAL</code>.',
      },
      {
        title: 'Quantos dias faltam?',
        text: '<p>Subtraindo uma data de outra, o resultado é a quantidade de dias entre elas. A data de hoje está em <b>I1</b> — trave com <b>$</b> para arrastar.</p>',
        task: 'Em <b>F2</b>, digite <code>=D2-$I$1</code> e arraste até <b>F6</b>.',
        target: ['cell:F2:F6', 'cell:I1'],
        check: (x) => x.formula('F2').includes('$I$1') && x.num('F2') === x.num('D2') - x.num('I1') && x.num('F6') === x.num('D6') - x.num('I1'),
        solve: (x) => { x.app.setCell('F2', '=D2-$I$1'); fillDown(x, 'F2', 'F6'); },
        success: 'Números negativos = a entrega já passou. Positivos = ainda faltam dias.',
      },
      {
        title: 'HOJE(): a data que se atualiza',
        text: '<p>Digitar a data de hoje na mão obriga você a mudar todo dia. A função <code>HOJE()</code> (sem nada entre os parênteses) sempre mostra a data atual do computador.</p>',
        task: 'Em <b>I1</b>, troque a data fixa por <code>=HOJE()</code>.',
        target: 'cell:I1',
        check: (x) => x.uses('I1', 'HOJE'),
        solve: (x) => { x.app.setCell('I1', '=HOJE()'); x.app.select('I2'); },
        success: `Agora a coluna "Faltam" se recalcula sozinha todo dia. Outras funções de data que valem a pena:
          <ul><li><code>DIA</code>, <code>MÊS</code>, <code>ANO</code> — pegam uma parte da data</li><li><code>DATA(ano; mês; dia)</code> — monta uma data</li>
          <li><code>DATADIF(nascimento; HOJE(); "a")</code> — idade em anos completos</li><li><code>FIMMÊS(data; 0)</code> — último dia do mês (ótimo para vencimentos)</li></ul>`,
      },
    ],
  });

  // ======================================================================
  // 6.3 Formatos personalizados e TEXTO
  // ======================================================================
  const D3 = [[7, '05/03/2026', 1234.5, 1250.5], [45, '18/07/2026', 89.9, 80], [123, '02/09/2026', 15000, 3.75], [9, '30/11/2026', 450, 960], [1002, '14/12/2026', 7.25, 15.2]];
  XL.registerLesson({
    id: 'i6-3',
    title: 'Formatos personalizados e TEXTO',
    summary: [
      '<b>Formatar Células › Personalizado</b> (<kbd>Ctrl</kbd>+<kbd>1</kbd>) cria formatos sob medida: <code>0000</code>, <code>#.##0,0 "kg"</code>, <code>dd/mm/aaaa</code>...',
      'No código: <b>0</b> = dígito obrigatório, <b>#</b> = opcional, <b>,</b> = decimal, <b>.</b> = milhar, texto entre <b>"aspas"</b>.',
      'O formato só muda a <b>aparência</b>; o valor continua sendo número.',
      '<code>TEXTO(valor; "formato")</code> aplica o mesmo código, mas devolve texto — ideal para montar frases com <code>&amp;</code>.',
    ],
    setup: (x) => {
      const m = {
        A1: { v: 'Código', fmt: head }, B1: { v: 'Data', fmt: head }, C1: { v: 'Valor', fmt: head }, D1: { v: 'Peso', fmt: head },
        E1: { v: 'Frase', fmt: head }, F1: { v: 'Mês', fmt: head }, G1: { v: 'Total formatado', fmt: head },
      };
      D3.forEach(([cod, d, v, p], i) => { const r = i + 2; m['A' + r] = cod; m['B' + r] = { v: d, fmt: data }; m['C' + r] = v; m['D' + r] = p; });
      x.set(m);
      x.widths({ B: 95, C: 80, D: 90, E: 180, F: 90, G: 160 });
    },
    steps: [
      {
        title: 'Zeros à esquerda',
        text: `<p>Os códigos deveriam ter sempre <b>4 dígitos</b> (0007, 0045...), mas o Excel descarta zeros à esquerda de números.</p>
          <p>A solução é um <b>formato personalizado</b>: o código <code>0000</code> diz "mostre sempre 4 dígitos, completando com zeros". O valor continua sendo o número 7.</p>`,
        task: 'Selecione <b>A2:A6</b>, abra <b>Formatar Células</b> (<kbd>Ctrl</kbd>+<kbd>1</kbd>), escolha <b>Personalizado</b>, digite <code>0000</code> e clique em OK.',
        hint: 'Também dá para abrir pelo botão <b>⋯</b> no grupo Número, pela opção "Personalizado..." da caixa de formato, ou com o botão direito › Formatar Células.',
        target: ['cell:A2:A6', '.fc-cat[data-cat="custom"]', 'input[data-role="code"]'],
        check: (x) => x.all('A2:A6', (r) => x.fmt(r).numFmt === 'custom' && x.fmt(r).code === '0000') && x.shown('A2') === '0007',
        solve: (x) => { x.app.select('A2:A6'); x.app.applyFmt({ numFmt: 'custom', code: '0000', dec: null, grouping: null }); },
        success: 'O 7 aparece como <b>0007</b>, mas a Barra de Fórmulas continua mostrando 7: o valor não mudou, só a aparência.',
      },
      {
        title: 'Unidade junto do número',
        text: '<p>Escrever "kg" na célula transformaria o número em texto e estragaria as contas. Com formato personalizado, o "kg" é só aparência: <code>#.##0,0 "kg"</code></p>',
        task: 'Aplique o formato personalizado <code>#.##0,0 "kg"</code> em <b>D2:D6</b>.',
        target: ['cell:D2:D6', 'input[data-role="code"]'],
        check: (x) => x.all('D2:D6', (r) => x.fmt(r).numFmt === 'custom' && /kg/.test(x.fmt(r).code || '')) && x.shown('D2') === '1.250,5 kg',
        solve: (x) => { x.app.select('D2:D6'); x.app.applyFmt({ numFmt: 'custom', code: '#.##0,0 "kg"', dec: null, grouping: null }); },
        success: '<b>1.250,5 kg</b>, com separador de milhar e uma casa decimal. Um <code>=SOMA(D2:D6)</code> continuaria funcionando normalmente.',
      },
      {
        title: 'Data no meio de uma frase',
        text: `<p>Tente juntar texto com data: <code>="Pedido de "&amp;B2</code> mostraria <b>"Pedido de 46086"</b> — o número por trás da data!</p>
          <p>A função <code>TEXTO(valor; "formato")</code> converte o valor em texto usando um código de formato:</p>
          <p style="text-align:center"><code>="Pedido de "&amp;TEXTO(B2;"dd/mm/aaaa")</code></p>`,
        task: 'Em <b>E2</b>, digite a fórmula acima e arraste até <b>E6</b>.',
        target: 'cell:E2:E6',
        check: (x) => x.uses('E2', 'TEXTO') && x.val('E2') === 'Pedido de 05/03/2026' && x.val('E6') === 'Pedido de 14/12/2026',
        solve: (x) => { x.app.setCell('E2', '="Pedido de "&TEXTO(B2;"dd/mm/aaaa")'); fillDown(x, 'E2', 'E6'); },
        success: 'No Excel em português, o ano é <b>aaaa</b> (de "ano"). Dia é <b>dd</b> e mês é <b>mm</b>.',
      },
      {
        title: 'O nome do mês',
        text: '<p>Códigos de data úteis: <code>mmm</code> = jan · <code>mmmm</code> = janeiro · <code>ddd</code> = seg · <code>dddd</code> = segunda-feira.</p>',
        task: 'Em <b>F2</b>, mostre o mês por extenso: <code>=TEXTO(B2;"mmmm")</code>, e arraste até <b>F6</b>.',
        target: 'cell:F2:F6',
        check: (x) => x.uses('F2', 'TEXTO') && x.val('F2') === 'março' && x.val('F6') === 'dezembro',
        solve: (x) => { x.app.setCell('F2', '=TEXTO(B2;"mmmm")'); fillDown(x, 'F2', 'F6'); },
        success: 'Ótimo para relatórios por mês. Com <code>"dddd"</code> você teria o dia da semana.',
      },
      {
        title: 'Valor formatado numa frase',
        task: 'Em <b>G2</b>, digite <code>="Total: "&amp;TEXTO(C2;"R$ #.##0,00")</code> e arraste até <b>G6</b>.',
        target: 'cell:G2:G6',
        check: (x) => x.uses('G2', 'TEXTO') && x.val('G2') === 'Total: R$ 1.234,50' && x.val('G4') === 'Total: R$ 15.000,00',
        solve: (x) => { x.app.setCell('G2', '="Total: "&TEXTO(C2;"R$ #.##0,00")'); fillDown(x, 'G2', 'G6'); },
        success: 'Lembre: o resultado do TEXTO é <b>texto</b>. Use-o para exibir, não para fazer contas. Para contas, formate a célula (Ctrl+1) e mantenha o número.',
      },
    ],
  });

  // ======================================================================
  // 6.4 Texto para Colunas e Preenchimento Relâmpago
  // ======================================================================
  const CAD = [['Ana Souza', 'ana.souza', 'São Paulo'], ['Bruno Lima', 'bruno.lima', 'Curitiba'], ['Carla Dias', 'carla.dias', 'Recife'],
    ['Diego Alves', 'diego.alves', 'Salvador'], ['Elisa Rocha', 'elisa.rocha', 'Belo Horizonte'], ['Fábio Melo', 'fabio.melo', 'Porto Alegre']];
  XL.registerLesson({
    id: 'i6-4',
    title: 'Texto para Colunas e Preenchimento Relâmpago',
    summary: [
      '<b>Dados › Texto para Colunas</b> divide um texto em várias colunas usando um separador (; , espaço...).',
      'O <b>Preenchimento Relâmpago</b> (<kbd>Ctrl</kbd>+<kbd>E</kbd>) aprende com um exemplo seu e completa a coluna.',
      'Ele reconhece padrões como "primeira palavra", "o que vem antes do @", junções e maiúsculas.',
      'Diferente de uma fórmula, o resultado é fixo: se a origem mudar, refaça o Ctrl+E.',
    ],
    setup: (x) => {
      const m = {
        A1: 'Nome;E-mail;Cidade',
        D1: { v: 'Primeiro nome', fmt: head }, E1: { v: 'Login', fmt: head }, F1: { v: 'Crachá', fmt: head },
      };
      CAD.forEach(([n, u, c], i) => { m['A' + (i + 2)] = `${n};${u}@email.com;${c}`; });
      x.set(m);
      x.widths({ A: 110, B: 175, C: 120, D: 110, E: 100, F: 130 });
    },
    steps: [
      {
        title: 'Tudo grudado numa coluna',
        text: `<p>Muitos sistemas exportam dados num formato chamado <b>CSV</b>: tudo numa coluna só, separado por <b>;</b> (ponto e vírgula).</p>
          <p>Antes de qualquer análise, é preciso separar cada informação na sua própria coluna.</p>`,
        target: 'cell:A1:A7',
      },
      {
        title: 'Texto para Colunas',
        task: 'Selecione <b>A1:A7</b> e use <b>Dados › Texto para Colunas</b>, com o delimitador <b>Ponto e vírgula</b>. Clique em <b>Concluir</b>.',
        hint: 'Na janela, confira se "Ponto e vírgula" está marcado — a visualização mostra como vai ficar.',
        target: ['cell:A1:A7', '.rb-btn[data-cmd="texttocols"]', 'input[data-delim="semicolon"]'],
        check: (x) => x.text('A2') === 'ANA SOUZA' && x.val('B2') === 'ana.souza@email.com' && x.val('C7') === 'Porto Alegre' && x.val('C1') === 'Cidade',
        solve: (x) => {
          x.app.pushUndo();
          const rows = [['Nome', 'E-mail', 'Cidade']].concat(CAD.map(([n, u, c]) => [n, u + '@email.com', c]));
          rows.forEach((row, i) => row.forEach((v, j) => x.app.sheet.setInput(i, j, v)));
          x.app.render();
          x.app.emit('texttocols', { cols: 3, rows: rows.length });
        },
        success: 'Nome, E-mail e Cidade, cada um na sua coluna. Funciona com qualquer separador: vírgula, espaço, tabulação ou outro caractere.',
      },
      {
        title: 'Preenchimento Relâmpago',
        text: `<p>Agora queremos só o <b>primeiro nome</b>. Em vez de fórmula, mostre ao Excel <b>um exemplo</b> e peça para ele adivinhar o padrão: é o <b>Preenchimento Relâmpago</b> (<kbd>Ctrl</kbd>+<kbd>E</kbd>).</p>`,
        task: 'Em <b>D2</b>, digite <b>Ana</b>, pressione <kbd>Enter</kbd> e depois <kbd>Ctrl</kbd>+<kbd>E</kbd>.',
        hint: 'O Ctrl+E precisa ser pressionado com a seleção na coluna D, logo abaixo do exemplo. Também está em Dados › Preenchimento Relâmpago.',
        target: ['cell:D2', '.rb-btn[data-cmd="flashfill"]'],
        check: (x) => x.happened('flashfill') && CAD.every(([n], i) => x.val('D' + (i + 2)) === n.split(' ')[0]),
        solve: (x) => { x.app.setCell('D2', 'Ana'); x.app.select('D3'); x.app.exec('flashfill'); },
        success: 'Um exemplo bastou: o Excel entendeu "primeira palavra do nome" e completou a coluna.',
      },
      {
        title: 'O login do e-mail',
        task: 'Em <b>E2</b>, digite <b>ana.souza</b> (a parte antes do @) e use <kbd>Ctrl</kbd>+<kbd>E</kbd>.',
        target: ['cell:E2'],
        check: (x) => x.happened('flashfill') && CAD.every(([, u], i) => x.val('E' + (i + 2)) === u),
        solve: (x) => { x.app.setCell('E2', 'ana.souza'); x.app.select('E3'); x.app.exec('flashfill'); },
        success: 'Ele percebeu que você queria "o que vem antes do @". Com fórmula, seria <code>=TEXTOANTES(B2;"@")</code>.',
      },
      {
        title: 'Reorganizando e mudando letras',
        task: 'Em <b>F2</b>, digite <b>SOUZA, Ana</b> (sobrenome em maiúsculas, vírgula, nome) e use <kbd>Ctrl</kbd>+<kbd>E</kbd>.',
        target: ['cell:F2'],
        check: (x) => x.happened('flashfill') && x.val('F3') === 'LIMA, Bruno' && x.val('F7') === 'MELO, Fábio',
        solve: (x) => { x.app.setCell('F2', 'SOUZA, Ana'); x.app.select('F3'); x.app.exec('flashfill'); },
        success: `Inverteu a ordem, colocou a vírgula e passou o sobrenome para maiúsculas. 🏆
          <br><br><b>Atenção:</b> o Preenchimento Relâmpago escreve <b>valores fixos</b>. Se um nome mudar na coluna A, a coluna D não muda sozinha — refaça o Ctrl+E ou use fórmulas quando os dados mudam com frequência.`,
      },
    ],
  });
})();
