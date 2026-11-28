/* Nível Avançado · Módulo 9 — Fórmulas modernas: matrizes dinâmicas, LET, LAMBDA e SOMARPRODUTO (4 lições + demonstrações). */
(function () {
  'use strict';
  const XL = window.XL;
  const head = { b: true, fill: '#DDEBF7' };
  const moeda = { numFmt: 'moeda' };
  const { cell, type, key, btn, el, typeIn, fill, look } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const ok = (say) => el('.ui-modal .btn-primary', say || 'Clique em <b>OK</b>');
  const esc = XL.escapeHtml;
  const write = (ref, formula, say) => [cell(ref), type(formula, say || `Digite <code>${esc(formula)}</code>`), enter()];
  const writeAndFill = (from, to, formula, say) => write(from, formula, say).concat([fill(from, to, `Selecione <b>${from}</b> de novo e arraste a alça de preenchimento até <b>${to}</b>`)]);
  const commit = (x, ref, text) => { const p = XL.parseRef(ref); x.app.select(ref); return x.app.commitInput(p.r, p.c, String(text)); };
  const fillDown = (x, from, to) => { x.app.select(from); x.app.fill(XL.parseRange(from), XL.parseRange(from + ':' + to), 'down'); };
  const near = (a, b) => typeof a === 'number' && Math.abs(a - b) < 1e-6;
  const spillOf = (x, ref) => { const p = XL.parseRef(ref); return x.sheet.spillAt(p.r, p.c); };

  // ======================================================================
  // 9.1 Matrizes dinâmicas: FILTRO, CLASSIFICAR, ÚNICO
  // ======================================================================
  const PED = [['Ana', 'Sul', 'Notebook', 3500], ['Bruno', 'Norte', 'Mouse', 240], ['Carla', 'Sudeste', 'Monitor', 2400], ['Diego', 'Nordeste', 'Teclado', 450],
    ['Elisa', 'Sul', 'Monitor', 1200], ['Fábio', 'Norte', 'Notebook', 7000], ['Gabi', 'Sudeste', 'Mouse', 160], ['Hugo', 'Nordeste', 'Notebook', 3500],
    ['Ana', 'Sudeste', 'Teclado', 300], ['Bruno', 'Sul', 'Notebook', 3500], ['Carla', 'Norte', 'Monitor', 1200], ['Diego', 'Sudeste', 'Monitor', 2400]];
  const REG_ORDEM = ['Sul', 'Norte', 'Sudeste', 'Nordeste'];
  const somaReg = (r) => PED.filter((p) => p[1] === r).reduce((s, p) => s + p[3], 0);
  const GRANDES = PED.filter((p) => p[3] >= 2000);
  XL.registerLesson({
    id: 'a9-1',
    title: 'Matrizes dinâmicas: FILTRO, CLASSIFICAR, ÚNICO',
    summary: [
      'No Excel 365, uma fórmula pode devolver <b>várias células</b>: o resultado "derrama" para baixo/para o lado, com uma borda azul em volta.',
      '<code>ÚNICO(intervalo)</code> tira as repetições; <code>CLASSIFICAR(matriz; coluna; 1 ou -1)</code> ordena; <code>FILTRO(matriz; condição; [se_vazio])</code> traz só as linhas que atendem à condição.',
      '<code>F2#</code> (com cerquilha) se refere ao resultado derramado inteiro, do tamanho que ele tiver.',
      '<code>#DESPEJAR!</code> aparece quando há algo no caminho do resultado: apague as células do caminho.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Vendedor', fmt: head }, B1: { v: 'Região', fmt: head }, C1: { v: 'Produto', fmt: head }, D1: { v: 'Valor', fmt: head },
        F1: { v: 'Regiões', fmt: head }, G1: { v: 'Total', fmt: head }, I1: { v: 'Vendas de R$ 2.000 ou mais', fmt: { b: true } } };
      PED.forEach(([v, r, p, val], i) => { const n = i + 2; m['A' + n] = v; m['B' + n] = r; m['C' + n] = p; m['D' + n] = { v: val, fmt: moeda }; });
      for (let n = 2; n <= 13; n++) { m['G' + n] = { v: null, fmt: moeda }; m['L' + n] = { v: null, fmt: moeda }; }
      x.set(m);
      x.widths({ A: 80, B: 80, C: 85, D: 95, E: 24, F: 85, G: 100, H: 24, I: 80, J: 80, K: 85, L: 95 });
    },
    steps: [
      {
        title: 'Uma fórmula, várias células',
        text: `<p>Até agora, cada fórmula preenchia <b>uma</b> célula. No Excel 365 existem as <b>matrizes dinâmicas</b>: você escreve a fórmula numa célula só e o resultado <b>derrama</b> pelas vizinhas.</p>
          <p>Vamos montar um mini-relatório destes pedidos sem arrastar nenhuma fórmula.</p>`,
        target: 'cell:A1:D13',
      },
      {
        title: 'ÚNICO: a lista sem repetições',
        task: 'Em <b>F2</b>, digite <code>=ÚNICO(B2:B13)</code> e pressione <kbd>Enter</kbd>.',
        target: 'cell:F2',
        check: (x) => x.uses('F2', 'ÚNICO') && REG_ORDEM.every((r, i) => x.val('F' + (i + 2)) === r) && x.empty('F3'),
        solve: (x) => commit(x, 'F2', '=ÚNICO(B2:B13)'),
        success: 'Apareceram as 4 regiões, cada uma uma vez. Só <b>F2</b> tem fórmula: clique em F3 e veja a fórmula <b>em cinza</b> na barra — ela vem de F2. A borda azul mostra até onde o resultado derramou.',
      },
      {
        title: 'F2#: o resultado inteiro',
        text: '<p>Para somar cada região, use o resultado do ÚNICO como critério. <code>F2#</code> quer dizer "todo o resultado que derramou a partir de F2" — se surgir uma região nova, ele cresce sozinho.</p>',
        task: 'Em <b>G2</b>, digite <code>=SOMASES(D2:D13;B2:B13;F2#)</code> e pressione <kbd>Enter</kbd>.',
        target: 'cell:G2',
        check: (x) => x.uses('G2', 'SOMASES') && /F2#/i.test(x.input('G2')) && REG_ORDEM.every((r, i) => near(x.val('G' + (i + 2)), somaReg(r))),
        solve: (x) => commit(x, 'G2', '=SOMASES(D2:D13;B2:B13;F2#)'),
        success: 'Um total para cada região, também derramado. Com a fórmula antiga você precisaria de uma fórmula por linha.',
      },
      {
        title: 'CLASSIFICAR em ordem alfabética',
        task: 'Volte em <b>F2</b> e troque a fórmula por <code>=CLASSIFICAR(ÚNICO(B2:B13))</code>.',
        target: 'cell:F2',
        check: (x) => x.uses('F2', 'CLASSIFICAR') && x.uses('F2', 'ÚNICO') && x.val('F2') === 'Nordeste' && x.val('F5') === 'Sul' && near(x.val('G2'), somaReg('Nordeste')),
        solve: (x) => commit(x, 'F2', '=CLASSIFICAR(ÚNICO(B2:B13))'),
        success: 'As regiões ficaram em ordem — e os totais em G <b>acompanharam</b>, porque G2 usa <code>F2#</code>. Para ordem decrescente: <code>=CLASSIFICAR(ÚNICO(B2:B13);1;-1)</code>.',
      },
      {
        title: 'FILTRO: só o que interessa',
        text: '<p><code>FILTRO(o_que_mostrar; condição)</code> devolve as <b>linhas inteiras</b> que atendem à condição — como o AutoFiltro, mas numa fórmula que se atualiza sozinha.</p>',
        task: 'Em <b>I2</b>, digite <code>=FILTRO(A2:D13;D2:D13&gt;=2000)</code> e pressione <kbd>Enter</kbd>.',
        target: 'cell:I2',
        check: (x) => { const sp = spillOf(x, 'I2'); return x.uses('I2', 'FILTRO') && !!sp && sp.ok && sp.h === GRANDES.length && sp.w === 4 && x.val('I2') === GRANDES[0][0]; },
        solve: (x) => commit(x, 'I2', '=FILTRO(A2:D13;D2:D13>=2000)'),
        success: `${GRANDES.length} pedidos de R$ 2.000 ou mais, com as 4 colunas. Se mudar um valor em D, a lista se refaz na hora. Dica: o 3º argumento mostra uma mensagem quando nada é encontrado: <code>=FILTRO(...;...;"Nenhum")</code>.`,
      },
      {
        title: 'O erro #DESPEJAR!',
        text: '<p>O resultado precisa de espaço livre. O que acontece se tiver algo no caminho?</p>',
        task: 'Em <b>I4</b> (dentro da área do FILTRO), digite <code>x</code> e pressione <kbd>Enter</kbd>.',
        target: 'cell:I4',
        check: (x) => x.isErr('I2', '#DESPEJAR!'),
        solve: (x) => commit(x, 'I4', 'x'),
        success: 'I2 virou <b>#DESPEJAR!</b> e a borda ficou tracejada, mostrando a área de que a fórmula precisa. Nada se perdeu: é só liberar o caminho.',
      },
      {
        title: 'Libere o caminho',
        task: 'Selecione <b>I4</b> e apague com a tecla <kbd>Delete</kbd>.',
        target: 'cell:I4',
        check: (x) => x.empty('I4') && !x.isErr('I2') && x.val('I4') === GRANDES[2][0],
        solve: (x) => { x.app.select('I4'); x.app.exec('clear', 'contents'); },
        success: 'O resultado voltou a derramar. Matrizes dinâmicas: ÚNICO, CLASSIFICAR, FILTRO, CLASSIFICARPOR e SEQUÊNCIA — uma fórmula, um relatório inteiro.',
      },
    ],
  });

  R('a9-1', {
    'ÚNICO: a lista sem repetições': write('F2', '=ÚNICO(B2:B13)').concat([look('cell:F2:F5', 'Uma fórmula só, quatro células: o resultado <b>derramou</b>')]),
    'F2#: o resultado inteiro': write('G2', '=SOMASES(D2:D13;B2:B13;F2#)', 'Digite <code>=SOMASES(D2:D13;B2:B13;F2#)</code> — o <b>#</b> pega todas as regiões de F'),
    'CLASSIFICAR em ordem alfabética': write('F2', '=CLASSIFICAR(ÚNICO(B2:B13))', 'Digite por cima: <code>=CLASSIFICAR(ÚNICO(B2:B13))</code>').concat([look('cell:F2:G5', 'As regiões ficaram em ordem e os totais acompanharam')]),
    'FILTRO: só o que interessa': write('I2', '=FILTRO(A2:D13;D2:D13>=2000)', 'Digite <code>=FILTRO(A2:D13;D2:D13&gt;=2000)</code>'),
    'O erro #DESPEJAR!': write('I4', 'x', 'Digite <code>x</code> no caminho do resultado').concat([look('cell:I2', 'I2 mostra <b>#DESPEJAR!</b>: falta espaço')]),
    'Libere o caminho': [cell('I4'), key('Delete', 'Pressione <kbd>Delete</kbd> para apagar')],
  });

  // ======================================================================
  // 9.2 LET: fórmulas legíveis
  // ======================================================================
  const ITENS = [['Cadeira', 450, 3], ['Mesa', 1200, 2], ['Armário', 980, 1], ['Estante', 320, 5], ['Sofá', 2900, 1], ['Luminária', 85, 10]];
  const final = (p, q) => { const t = p * q; return t > 1000 ? t * 0.9 : t; };
  XL.registerLesson({
    id: 'a9-2',
    title: 'LET: fórmulas legíveis',
    summary: [
      '<code>LET(nome; valor; cálculo)</code> dá um nome a um resultado intermediário e usa esse nome no cálculo.',
      'Vários nomes: <code>LET(total; B2*C2; desconto; SE(total&gt;1000; 10%; 0); total*(1-desconto))</code>.',
      'Vantagens: a fórmula fica legível, a conta repetida é feita <b>uma vez só</b> e, para mudar, você muda num lugar.',
      'LET combina com matrizes dinâmicas: dê nome a um FILTRO e use o nome em CLASSIFICAR, CONT.VALORES...',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Item', fmt: head }, B1: { v: 'Preço', fmt: head }, C1: { v: 'Qtd', fmt: head }, D1: { v: 'Sem LET', fmt: head }, E1: { v: 'Com LET', fmt: head }, F1: { v: 'Com 2 nomes', fmt: head }, H1: { v: 'Itens com desconto', fmt: { b: true } } };
      ITENS.forEach(([i, p, q], k) => { const n = k + 2; m['A' + n] = i; m['B' + n] = { v: p, fmt: moeda }; m['C' + n] = q; m['D' + n] = `=SE(B${n}*C${n}>1000;B${n}*C${n}*0,9;B${n}*C${n})`; m['E' + n] = { v: null, fmt: moeda }; m['F' + n] = { v: null, fmt: moeda }; });
      x.set(m);
      for (let n = 2; n <= 7; n++) x.app.sheet.setFmt(n - 1, 3, moeda);
      x.widths({ A: 85, B: 90, C: 50, D: 105, E: 105, F: 110, G: 24, H: 140 });
    },
    steps: [
      {
        title: 'A mesma conta, três vezes',
        text: `<p>A coluna D calcula o valor final com 10% de desconto acima de R$ 1.000. Clique em D2 e veja a fórmula:</p>
          <p style="text-align:center"><code>=SE(B2*C2&gt;1000;B2*C2*0,9;B2*C2)</code></p>
          <p><code>B2*C2</code> aparece <b>três vezes</b>. Se a regra mudar, é fácil esquecer uma. O <b>LET</b> resolve isso dando um nome à conta.</p>`,
        target: 'cell:D2:D7',
      },
      {
        title: 'Seu primeiro LET',
        text: '<p><code>LET(total; B2*C2; ...)</code> cria o nome <b>total</b> valendo B2*C2. Depois do último <code>;</code> vem o cálculo, que já pode usar <b>total</b>.</p>',
        task: 'Em <b>E2</b>, digite <code>=LET(total;B2*C2;SE(total&gt;1000;total*0,9;total))</code> e arraste até <b>E7</b>.',
        target: 'cell:E2:E7',
        check: (x) => x.uses('E2', 'LET') && ITENS.every(([, p, q], k) => near(x.val('E' + (k + 2)), final(p, q))),
        solve: (x) => { commit(x, 'E2', '=LET(total;B2*C2;SE(total>1000;total*0,9;total))'); fillDown(x, 'E2', 'E7'); },
        success: 'Os mesmos resultados da coluna D, mas a fórmula se lê como uma frase: "total é B2*C2; se total passar de 1000, dá 10% de desconto".',
      },
      {
        title: 'Vários nomes',
        text: '<p>Dá para criar quantos nomes quiser, sempre em pares <b>nome; valor</b>. Cada nome pode usar os anteriores.</p>',
        task: 'Em <b>F2</b>, digite <code>=LET(total;B2*C2;desconto;SE(total&gt;1000;10%;0);total*(1-desconto))</code> e arraste até <b>F7</b>.',
        target: 'cell:F2:F7',
        check: (x) => x.uses('F2', 'LET') && /DESCONTO/i.test(x.input('F2')) && ITENS.every(([, p, q], k) => near(x.val('F' + (k + 2)), final(p, q))),
        solve: (x) => { commit(x, 'F2', '=LET(total;B2*C2;desconto;SE(total>1000;10%;0);total*(1-desconto))'); fillDown(x, 'F2', 'F7'); },
        success: 'Agora a regra do desconto tem nome próprio. Se o desconto virar 15%, você muda <b>um</b> número.',
      },
      {
        title: 'LET com matrizes dinâmicas',
        text: '<p>LET também guarda <b>listas inteiras</b>. Vamos guardar os itens que tiveram desconto e mostrá-los em ordem.</p>',
        task: 'Em <b>H2</b>, digite <code>=LET(lista;FILTRO(A2:A7;F2:F7&lt;B2:B7*C2:C7);CLASSIFICAR(lista))</code>.',
        target: 'cell:H2',
        check: (x) => {
          const esperado = ITENS.filter(([, p, q]) => p * q > 1000).map((i) => i[0]).sort((a, b) => a.localeCompare(b, 'pt-BR'));
          return x.uses('H2', 'LET') && x.uses('H2', 'FILTRO') && esperado.every((nm, i) => x.val('H' + (i + 2)) === nm);
        },
        solve: (x) => commit(x, 'H2', '=LET(lista;FILTRO(A2:A7;F2:F7<B2:B7*C2:C7);CLASSIFICAR(lista))'),
        success: 'O nome <b>lista</b> guardou o resultado do FILTRO, e o CLASSIFICAR o ordenou. Em fórmulas grandes, o LET é o que separa "ninguém entende" de "dá para manter".',
      },
    ],
  });

  R('a9-2', {
    'Seu primeiro LET': writeAndFill('E2', 'E7', '=LET(total;B2*C2;SE(total>1000;total*0,9;total))',
      'Digite <code>=LET(total;B2*C2;SE(total&gt;1000;total*0,9;total))</code> — "total" vira um nome'),
    'Vários nomes': writeAndFill('F2', 'F7', '=LET(total;B2*C2;desconto;SE(total>1000;10%;0);total*(1-desconto))',
      'Digite <code>=LET(total;B2*C2;desconto;SE(total&gt;1000;10%;0);total*(1-desconto))</code>'),
    'LET com matrizes dinâmicas': write('H2', '=LET(lista;FILTRO(A2:A7;F2:F7<B2:B7*C2:C7);CLASSIFICAR(lista))',
      'Digite <code>=LET(lista;FILTRO(A2:A7;F2:F7&lt;B2:B7*C2:C7);CLASSIFICAR(lista))</code>'),
  });

  // ======================================================================
  // 9.3 LAMBDA: crie suas próprias funções
  // ======================================================================
  const VEND = [['Ana', 1800], ['Bruno', 650], ['Carla', 2400], ['Diego', 990], ['Elisa', 1000], ['Fábio', 300]];
  const LAMBDA_1 = '=LAMBDA(valor;SE(valor>=1000;valor*8%;valor*5%))';
  const LAMBDA_2 = '=LAMBDA(valor;SE(valor>=1000;valor*10%;valor*5%))';
  const com = (v, pct) => (v >= 1000 ? v * pct : v * 0.05);
  const nameOf = (x, n) => x.app.wb.names.find((m) => XL.normName(m.name) === XL.normName(n));
  XL.registerLesson({
    id: 'a9-3',
    title: 'LAMBDA: crie suas próprias funções',
    summary: [
      '<code>LAMBDA(parâmetro; cálculo)</code> cria uma função. Teste na célula chamando com parênteses: <code>=LAMBDA(x; x*2)(5)</code>.',
      'Dê um nome em <b>Fórmulas › Definir Nome</b> (Refere-se a: <code>=LAMBDA(...)</code>) e use como qualquer função: <code>=COMISSAO(B2)</code>.',
      'Mudou a regra? Edite no <b>Gerenciador de Nomes</b> e todas as células que usam a função se atualizam.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Vendedor', fmt: head }, B1: { v: 'Vendas', fmt: head }, C1: { v: 'Teste', fmt: head }, D1: { v: 'Comissão', fmt: head } };
      VEND.forEach(([n, v], k) => { const r = k + 2; m['A' + r] = n; m['B' + r] = { v, fmt: moeda }; m['C' + r] = { v: null, fmt: moeda }; m['D' + r] = { v: null, fmt: moeda }; });
      x.set(m);
      x.widths({ A: 85, B: 95, C: 95, D: 105 });
    },
    steps: [
      {
        title: 'Uma regra que se repete',
        text: `<p>A comissão é de <b>8%</b> para vendas de R$ 1.000 ou mais e <b>5%</b> abaixo disso. Essa regra aparece em várias planilhas da empresa — e, quando muda, alguém sempre esquece de atualizar uma.</p>
          <p>Com <b>LAMBDA</b> você cria uma função sua, <b>COMISSAO</b>, escrita uma vez só.</p>`,
        target: 'cell:A1:B7',
      },
      {
        title: 'Teste o LAMBDA na célula',
        text: '<p><code>LAMBDA(valor; cálculo)</code> cria a função; os parênteses no final <b>chamam</b> a função com um valor: <code>(B2)</code>.</p>',
        task: `Em <b>C2</b>, digite <code>${esc(LAMBDA_1)}(B2)</code> e pressione <kbd>Enter</kbd>.`,
        target: 'cell:C2',
        check: (x) => x.uses('C2', 'LAMBDA') && near(x.val('C2'), com(VEND[0][1], 0.08)),
        solve: (x) => commit(x, 'C2', LAMBDA_1 + '(B2)'),
        success: `R$ 144,00 = 8% de 1.800. A função funciona — mas escrever isso toda vez seria pior que antes. Vamos dar um <b>nome</b> a ela.`,
      },
      {
        title: 'Dê um nome à função',
        task: `Em <b>Fórmulas › Definir Nome</b>, digite o nome <code>COMISSAO</code> e, em <b>Refere-se a</b>, <code>${esc(LAMBDA_1)}</code>. Clique em OK.`,
        hint: 'Sem os parênteses do final: aqui você está guardando a função, não chamando.',
        target: '.rb-btn[data-cmd="name-define"]',
        check: (x) => { const n = nameOf(x, 'COMISSAO'); return !!n && /LAMBDA/i.test(n.ref) && near(x.app.wb.nameValue('COMISSAO', x.sheet) instanceof XL.Formula.Lambda ? 1 : 0, 1); },
        solve: (x) => { x.app.pushUndo(); x.app.wb.names.push({ name: 'COMISSAO', ref: XL.Formula.normalize(LAMBDA_1), comment: '' }); x.app.wb.invalidateAll(); x.app.render(); x.app.emit('name', { name: 'COMISSAO' }); },
        success: 'Pronto: a pasta de trabalho agora tem uma função chamada <b>COMISSAO</b>.',
      },
      {
        title: 'Use a sua função',
        task: 'Em <b>D2</b>, digite <code>=COMISSAO(B2)</code> e arraste até <b>D7</b>.',
        target: 'cell:D2:D7',
        check: (x) => /COMISSAO\(/i.test(x.input('D2')) && VEND.every(([, v], k) => near(x.val('D' + (k + 2)), com(v, 0.08))),
        solve: (x) => { commit(x, 'D2', '=COMISSAO(B2)'); fillDown(x, 'D2', 'D7'); },
        success: 'Uma fórmula curta e clara, como uma função nativa do Excel. Repare na Elisa: exatamente R$ 1.000 recebe 8% (a regra é "maior ou igual").',
      },
      {
        title: 'Mude a regra num lugar só',
        text: '<p>A diretoria aumentou a comissão de quem vende R$ 1.000 ou mais para <b>10%</b>. Em vez de corrigir célula por célula, edite a função.</p>',
        task: `Abra <b>Fórmulas › Gerenciador de Nomes</b>, clique em <b>COMISSAO</b>, depois em <b>Editar...</b>, troque a fórmula por <code>${esc(LAMBDA_2)}</code> e clique em OK.`,
        target: '.rb-btn[data-cmd="name-manager"]',
        check: (x) => VEND.every(([, v], k) => near(x.val('D' + (k + 2)), com(v, 0.1))),
        solve: (x) => { const n = nameOf(x, 'COMISSAO'); x.app.pushUndo(); n.ref = XL.Formula.normalize(LAMBDA_2); n._ast = null; x.app.wb.invalidateAll(); x.app.render(); x.app.emit('name', { name: 'COMISSAO' }); },
        success: 'Todas as comissões de R$ 1.000 ou mais foram recalculadas de uma vez. É assim que equipes mantêm regras de negócio: <b>uma definição, muitos usos</b>.',
      },
    ],
  });

  R('a9-3', {
    'Teste o LAMBDA na célula': write('C2', LAMBDA_1 + '(B2)', `Digite <code>${esc(LAMBDA_1)}(B2)</code> — o <b>(B2)</b> no final chama a função`),
    'Dê um nome à função': [
      btn('name-define', undefined, 'Na guia <b>Fórmulas</b>, clique em <b>Definir Nome</b>'),
      typeIn('input[data-role="nm-name"]', 'COMISSAO', 'Em <b>Nome</b>, digite <code>COMISSAO</code>'),
      typeIn('input[data-role="nm-ref"]', LAMBDA_1, 'Em <b>Refere-se a</b>, digite a função LAMBDA (sem chamar)'),
      ok(),
    ],
    'Use a sua função': writeAndFill('D2', 'D7', '=COMISSAO(B2)', 'Digite <code>=COMISSAO(B2)</code> — como qualquer função do Excel'),
    'Mude a regra num lugar só': [
      btn('name-manager', undefined, 'Na guia <b>Fórmulas</b>, clique em <b>Gerenciador de Nomes</b>'),
      el('.nm-table tr[data-name="COMISSAO"]', 'Clique em <b>COMISSAO</b>'),
      el('[data-role="nm-edit"]', 'Clique em <b>Editar...</b>'),
      typeIn('input[data-role="nm-ref"]', LAMBDA_2, 'Troque <b>8%</b> por <b>10%</b> na fórmula'),
      ok(),
      look('cell:D2:D7', 'Todas as comissões se atualizaram'),
    ],
  });

  // ======================================================================
  // 9.4 SOMARPRODUTO e cálculos com matrizes
  // ======================================================================
  const VND = [['Mouse', 'Sul', 12, 80], ['Teclado', 'Norte', 5, 150], ['Mouse', 'Norte', 8, 80], ['Monitor', 'Sul', 2, 1200],
    ['Mouse', 'Sul', 20, 75], ['Teclado', 'Sul', 4, 150], ['Monitor', 'Norte', 3, 1150], ['Mouse', 'Sudeste', 15, 80]];
  const fat = (pred) => VND.filter(pred).reduce((s, [, , q, p]) => s + q * p, 0);
  XL.registerLesson({
    id: 'a9-4',
    title: 'SOMARPRODUTO e cálculos com matrizes',
    summary: [
      '<code>=C2:C9*D2:D9</code> multiplica linha a linha e derrama os resultados: contas com <b>intervalos inteiros</b>.',
      '<code>SOMARPRODUTO(C2:C9;D2:D9)</code> multiplica par a par e soma tudo — o faturamento total numa fórmula.',
      'Condição vira número: <code>(B2:B9="Sul")</code> dá VERDADEIRO/FALSO, e ao multiplicar vira 1/0. Assim: <code>=SOMARPRODUTO((B2:B9="Sul")*C2:C9*D2:D9)</code>.',
      'Várias condições = vários parênteses multiplicados (E). Para OU, some as condições.',
    ],
    setup: (x) => {
      const m = { A1: { v: 'Produto', fmt: head }, B1: { v: 'Região', fmt: head }, C1: { v: 'Qtd', fmt: head }, D1: { v: 'Preço', fmt: head }, E1: { v: 'Total da linha', fmt: head },
        G2: { v: 'Faturamento total', fmt: { b: true } }, G3: { v: 'Faturamento do Sul', fmt: { b: true } }, G4: { v: 'Mouse no Sul', fmt: { b: true } } };
      VND.forEach(([pr, r, q, p], k) => { const n = k + 2; m['A' + n] = pr; m['B' + n] = r; m['C' + n] = q; m['D' + n] = { v: p, fmt: moeda }; m['E' + n] = { v: null, fmt: moeda }; });
      ['H2', 'H3', 'H4'].forEach((c) => { m[c] = { v: null, fmt: Object.assign({ fill: '#FFF2CC' }, moeda) }; });
      x.set(m);
      x.widths({ A: 80, B: 80, C: 50, D: 90, E: 110, F: 24, G: 140, H: 115 });
    },
    steps: [
      {
        title: 'Contas com intervalos inteiros',
        text: `<p>Quantidade × preço para cada linha: normalmente você escreveria <code>=C2*D2</code> e arrastaria. Com matrizes, uma fórmula faz <b>todas as linhas</b>: <code>=C2:C9*D2:D9</code> multiplica a 1ª pela 1ª, a 2ª pela 2ª...</p>`,
        task: 'Em <b>E2</b>, digite <code>=C2:C9*D2:D9</code> e pressione <kbd>Enter</kbd>.',
        target: 'cell:E2',
        check: (x) => VND.every(([, , q, p], k) => near(x.val('E' + (k + 2)), q * p)) && !x.isFormula('E3'),
        solve: (x) => commit(x, 'E2', '=C2:C9*D2:D9'),
        success: 'Oito resultados com uma fórmula, derramados para baixo. O Excel faz a conta "elemento por elemento".',
      },
      {
        title: 'SOMARPRODUTO: multiplica e soma',
        task: 'Em <b>H2</b>, calcule o faturamento total: <code>=SOMARPRODUTO(C2:C9;D2:D9)</code>.',
        target: 'cell:H2',
        check: (x) => x.uses('H2', 'SOMARPRODUTO') && near(x.val('H2'), fat(() => true)),
        solve: (x) => commit(x, 'H2', '=SOMARPRODUTO(C2:C9;D2:D9)'),
        success: `${XL.formatValue(fat(() => true), moeda).text}: a soma de todas as multiplicações, sem coluna auxiliar. É o mesmo que <code>=SOMA(E2#)</code>.`,
      },
      {
        title: 'Uma condição dentro da conta',
        text: `<p>E só o Sul? A comparação <code>(B2:B9="Sul")</code> devolve uma lista de VERDADEIRO/FALSO. Ao <b>multiplicar</b>, VERDADEIRO vira 1 e FALSO vira 0 — as linhas de fora do Sul são zeradas.</p>`,
        task: 'Em <b>H3</b>, digite <code>=SOMARPRODUTO((B2:B9="Sul")*C2:C9*D2:D9)</code>.',
        target: 'cell:H3',
        check: (x) => x.uses('H3', 'SOMARPRODUTO') && near(x.val('H3'), fat(([, r]) => r === 'Sul')),
        solve: (x) => commit(x, 'H3', '=SOMARPRODUTO((B2:B9="Sul")*C2:C9*D2:D9)'),
        success: 'Faturamento só do Sul. Repare que a condição vai entre parênteses e é <b>multiplicada</b>, não separada por ponto e vírgula.',
      },
      {
        title: 'Duas condições',
        text: '<p>Para "Sul <b>e</b> Mouse", multiplique as duas condições: só dá 1 quando as duas são verdadeiras (1×1).</p>',
        task: 'Em <b>H4</b>, digite <code>=SOMARPRODUTO((B2:B9="Sul")*(A2:A9="Mouse")*C2:C9*D2:D9)</code>.',
        target: 'cell:H4',
        check: (x) => x.uses('H4', 'SOMARPRODUTO') && near(x.val('H4'), fat(([pr, r]) => r === 'Sul' && pr === 'Mouse')),
        solve: (x) => commit(x, 'H4', '=SOMARPRODUTO((B2:B9="Sul")*(A2:A9="Mouse")*C2:C9*D2:D9)'),
        success: 'Mouse no Sul, com preços diferentes por linha — algo que SOMASES sozinho não resolve, porque precisa multiplicar Qtd × Preço antes de somar. Para "OU", <b>some</b> as condições: <code>((B2:B9="Sul")+(B2:B9="Norte")&gt;0)</code>.',
      },
    ],
  });

  R('a9-4', {
    'Contas com intervalos inteiros': write('E2', '=C2:C9*D2:D9', 'Digite <code>=C2:C9*D2:D9</code> — intervalos inteiros, sem arrastar').concat([look('cell:E2:E9', 'Oito resultados derramados')]),
    'SOMARPRODUTO: multiplica e soma': write('H2', '=SOMARPRODUTO(C2:C9;D2:D9)'),
    'Uma condição dentro da conta': write('H3', '=SOMARPRODUTO((B2:B9="Sul")*C2:C9*D2:D9)', 'Digite <code>=SOMARPRODUTO((B2:B9="Sul")*C2:C9*D2:D9)</code> — a condição entre parênteses, multiplicada'),
    'Duas condições': write('H4', '=SOMARPRODUTO((B2:B9="Sul")*(A2:A9="Mouse")*C2:C9*D2:D9)', 'Digite <code>=SOMARPRODUTO((B2:B9="Sul")*(A2:A9="Mouse")*C2:C9*D2:D9)</code>'),
  });
})();
