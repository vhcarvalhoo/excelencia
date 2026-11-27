/* Demonstrações "Mostre-me como" do Nível Intermediário. */
(function () {
  'use strict';
  const XL = window.XL;
  const { cell, point, type, key, fill, look } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  // escreve a fórmula na primeira célula e arrasta a alça até a última
  const writeAndFill = (from, to, formula, say) => [
    cell(from),
    type(formula, say || `Digite <code>${XL.escapeHtml(formula)}</code>`),
    enter(),
    fill(from, to, `Selecione <b>${from}</b> de novo e arraste a alça de preenchimento até <b>${to}</b>`),
  ];

  // ================= Módulo 3 =================
  R('i3-1', {
    'Primeiro, só o teste': [
      cell('C2'),
      type('=B2>=7', 'Digite <code>=B2&gt;=7</code> — "a nota em B2 é maior ou igual a 7?"'),
      enter('Pressione <kbd>Enter</kbd>: a resposta é VERDADEIRO ou FALSO'),
    ],
    'Agora com SE': [
      cell('C2', 'Clique em <b>C2</b> — vamos digitar por cima'),
      type('=SE(', 'Digite <code>=SE(</code> — a dica amarela mostra os 3 argumentos'),
      point('B2', 'Clique em <b>B2</b> (a nota)'),
      type('>=7;', 'Complete o teste com <code>&gt;=7</code> e passe para o próximo argumento com <code>;</code>'),
      type('"Aprovado";', 'Se for verdade: <code>"Aprovado"</code> (texto entre aspas) e <code>;</code>'),
      type('"Reprovado")', 'Se for falso: <code>"Reprovado"</code> e feche o parêntese'),
      enter(),
    ],
    'Para a turma toda': [fill('C2', 'C9', 'Arraste a <b>alça de preenchimento</b> de C2 até <b>C9</b>')],
    'Uma regra fácil de mudar': [
      cell('C2'),
      type('=SE(B2>=F1', 'Digite <code>=SE(B2&gt;=F1</code>'),
      key('F4', 'Pressione <kbd>F4</kbd>: F1 vira <code>$F$1</code> e fica travado ao arrastar'),
      type(';"Aprovado";"Reprovado")', 'Complete com <code>;"Aprovado";"Reprovado")</code>'),
      enter(),
      fill('C2', 'C9', 'Arraste a alça até <b>C9</b> para atualizar todos'),
    ],
    'Mude a regra': [cell('F1'), type('6'), enter(), look('cell:C5', 'O Diego (6,5) virou <b>Aprovado</b> sozinho')],
  });

  R('i3-2', {
    'Testando o E': [
      cell('D2'),
      type('=E(B2>=10000;C2>=8)', 'Digite <code>=E(B2&gt;=10000;C2&gt;=8)</code> — as duas condições separadas por <code>;</code>'),
      enter(),
    ],
    'E dentro do SE': writeAndFill('D2', 'D7', '=SE(E(B2>=10000;C2>=8);"Sim";"Não")',
      'Digite <code>=SE(E(B2&gt;=10000;C2&gt;=8);"Sim";"Não")</code> — o E é o teste do SE'),
    'Agora o OU': writeAndFill('E2', 'E7', '=SE(OU(B2>=15000;C2=10);"Sim";"Não")'),
    'O NÃO inverte': writeAndFill('F2', 'F7', '=SE(NÃO(C2>=6);"Sim";"Não")'),
  });

  R('i3-3', {
    'Monte o SE aninhado': [
      cell('C2'),
      type('=SE(B2>=9;"A";', 'Primeiro teste: <code>=SE(B2&gt;=9;"A";</code> — se for 9 ou mais, "A"'),
      type('SE(B2>=7;"B";', 'No lugar do "senão", outro SE: <code>SE(B2&gt;=7;"B";</code>'),
      type('SE(B2>=5;"C";"D")))', 'E mais um: <code>SE(B2&gt;=5;"C";"D")</code> — e feche os <b>3 parênteses</b>: <code>)))</code>'),
      enter(),
      fill('C2', 'C9'),
    ],
    'O jeito moderno: SES': [
      cell('D2'),
      type('=SES(B2>=9;"A";', 'Digite <code>=SES(</code> e o primeiro par: <code>B2&gt;=9;"A";</code>'),
      type('B2>=7;"B";B2>=5;"C";', 'Os próximos pares: <code>B2&gt;=7;"B";B2&gt;=5;"C";</code>'),
      type('VERDADEIRO;"D")', 'O "senão": <code>VERDADEIRO;"D"</code> e feche o parêntese'),
      enter(),
      fill('D2', 'D9'),
    ],
  });

  // ================= Módulo 4 =================
  // =FUNÇÃO( + selecionar o intervalo com o mouse + resto da fórmula
  const withRange = (ref, start, range, rest, sayRest) => [
    cell(ref),
    type(start, `Digite <code>${XL.escapeHtml(start)}</code>`),
    point(range, `Selecione <b>${range}</b> com o mouse — é <b>onde</b> o Excel vai procurar`),
    type(rest, sayRest || `Complete com <code>${XL.escapeHtml(rest)}</code>`),
    enter(),
  ];
  R('i4-1', {
    'Quantos pedidos da Ana?': withRange('H2', '=CONT.SE(', 'A2:A17', ';"Ana")', 'Digite <code>;"Ana")</code> — o critério, entre aspas, e feche o parêntese'),
    'Quantos pedidos do Sul?': withRange('H3', '=CONT.SE(', 'B2:B17', ';"Sul")'),
    'Pedidos grandes': withRange('H4', '=CONT.SE(', 'E2:E17', ';">2000")', 'Digite <code>;"&gt;2000")</code> — operador e número <b>juntos entre aspas</b>'),
    'Duas condições: CONT.SES': [
      cell('H5'),
      type('=CONT.SES(', 'Digite <code>=CONT.SES(</code>'),
      point('B2:B17', 'Primeiro par — o intervalo: selecione <b>B2:B17</b> (Região)'),
      type(';"Sul";', '...e o critério: <code>;"Sul";</code>'),
      point('C2:C17', 'Segundo par — o intervalo: selecione <b>C2:C17</b> (Produto)'),
      type(';"Notebook")', '...e o critério: <code>;"Notebook")</code>'),
      enter(),
    ],
    'Misturando texto e número': [cell('H6'), type('=CONT.SES(B2:B17;"Sul";E2:E17;">=1000")', 'Digite <code>=CONT.SES(B2:B17;"Sul";E2:E17;"&gt;=1000")</code>'), enter()],
  });

  R('i4-2', {
    'Quanto a Ana vendeu?': [
      cell('K2'),
      type('=SOMASE(', 'Digite <code>=SOMASE(</code>'),
      point('A2:A17', 'Onde <b>testar</b>: selecione <b>A2:A17</b> (Vendedor)'),
      type(';"Ana";', 'O <b>critério</b>: <code>;"Ana";</code>'),
      point('E2:E17', 'O que <b>somar</b>: selecione <b>E2:E17</b> (Valor)'),
      type(')', 'Feche o parêntese'),
      enter(),
    ],
    'Somando outra coluna': [cell('K3'), type('=SOMASE(C2:C17;"Mouse";D2:D17)', 'Digite <code>=SOMASE(C2:C17;"Mouse";D2:D17)</code> — testa a coluna C, soma a coluna D'), enter()],
    'Um resumo por região': [
      cell('H2'),
      type('=SOMASE(B2:B17', 'Digite <code>=SOMASE(B2:B17</code>'),
      key('F4', '<kbd>F4</kbd> trava o intervalo: <code>$B$2:$B$17</code>'),
      type(';', 'Digite <code>;</code>'),
      point('G2', 'Clique em <b>G2</b> — o critério vem da célula (sem $, para mudar ao arrastar)'),
      type(';E2:E17', 'Digite <code>;E2:E17</code>'),
      key('F4', '<kbd>F4</kbd> de novo: <code>$E$2:$E$17</code>'),
      type(')'),
      enter(),
      fill('H2', 'H5', 'Arraste a alça de <b>H2</b> até <b>H5</b>'),
    ],
    'Confira o total': [cell('H6'), type('=SOMA(H2:H5)'), enter()],
    'Duas condições: SOMASES': [
      cell('K4'),
      type('=SOMASES(E2:E17;', 'No SOMASES, o que <b>somar vem primeiro</b>: <code>=SOMASES(E2:E17;</code>'),
      type('B2:B17;"Sul";', 'Primeiro par: <code>B2:B17;"Sul";</code>'),
      type('C2:C17;"Notebook")', 'Segundo par: <code>C2:C17;"Notebook")</code>'),
      enter(),
    ],
  });

  R('i4-3', {
    'Começa com...': withRange('E3', '=CONT.SE(', 'A2:A11', ';"Mouse*")', 'Digite <code>;"Mouse*")</code> — o <b>*</b> no fim quer dizer "e qualquer coisa depois"'),
    'Termina com...': withRange('E4', '=CONT.SE(', 'A2:A11', ';"*sem fio")', 'Digite <code>;"*sem fio")</code> — o <b>*</b> no começo quer dizer "qualquer coisa antes"'),
    'MÉDIASE': [
      cell('E5'),
      type('=MÉDIASE(', 'Digite <code>=MÉDIASE(</code>'),
      point('A2:A11', 'Onde testar: <b>A2:A11</b>'),
      type(';"Monitor*";', 'Critério com curinga: <code>;"Monitor*";</code>'),
      point('B2:B11', 'Do que tirar a média: <b>B2:B11</b>'),
      type(')'),
      enter(),
    ],
    'Critério que vem de uma célula': [
      cell('E6'),
      type('=CONT.SE(B2:B11;', 'Digite <code>=CONT.SE(B2:B11;</code>'),
      type('">"&', 'O operador entre aspas e o <b>&amp;</b> para juntar: <code>"&gt;"&amp;</code>'),
      point('E1', 'Clique em <b>E1</b>, o preço mínimo'),
      type(')'),
      enter(),
    ],
    'Mude o preço mínimo': [cell('E1'), type('200'), enter(), look('cell:E6', 'A contagem acompanhou: agora são 6 itens')],
  });

  // ================= Módulo 5 =================
  R('i5-1', {
    'Busque o nome do produto': [
      cell('B2'),
      type('=PROCV(', 'Digite <code>=PROCV(</code>'),
      point('A2', '1º — o que procurar: clique no código em <b>A2</b>'),
      type(';', 'Digite <code>;</code>'),
      point('G2:I9', '2º — onde procurar: selecione a tabela verde <b>G2:I9</b> (o código fica na 1ª coluna)'),
      key('F4', '<kbd>F4</kbd> trava a tabela: <code>$G$2:$I$9</code>'),
      type(';2;', '3º — qual coluna trazer: <code>2</code> (Produto é a 2ª coluna da tabela)'),
      type('FALSO)', '4º — <code>FALSO</code> = busca exata. Feche o parêntese'),
      enter(),
    ],
    'Para o pedido todo': [fill('B2', 'B7', 'Arraste a alça de <b>B2</b> até <b>B7</b>'), look('cell:B7', 'O código P110 não existe na tabela: <b>#N/D</b>')],
    'Agora o preço': writeAndFill('C2', 'C7', '=PROCV(A2;$G$2:$I$9;3;FALSO)',
      'Digite <code>=PROCV(A2;$G$2:$I$9;3;FALSO)</code> — igual ao anterior, mas com a coluna <b>3</b> (Preço)'),
    'Uma mensagem no lugar do #N/D': [
      cell('B2'),
      type('=SEERRO(', 'Comece com <code>=SEERRO(</code>'),
      type('PROCV(A2;$G$2:$I$9;2;FALSO);', 'Dentro dele, o mesmo PROCV de antes, seguido de <code>;</code>'),
      type('"Não encontrado")', 'E a mensagem para quando der erro: <code>"Não encontrado")</code>'),
      enter(),
      fill('B2', 'B7'),
    ],
    'Corrija o código': [cell('A7'), type('P107'), enter(), look('cell:B7:C7', 'Produto e preço apareceram na hora')],
  });

  R('i5-2', {
    'Busque o nome': [
      cell('G3'),
      type('=PROCX(', 'Digite <code>=PROCX(</code>'),
      point('G1', '1º — o que procurar: clique na matrícula em <b>G1</b>'),
      type(';'),
      point('C2:C9', '2º — onde procurar: a coluna de matrículas <b>C2:C9</b>'),
      type(';'),
      point('A2:A9', '3º — o que trazer: a coluna de nomes <b>A2:A9</b> (à esquerda — o PROCV não conseguiria)'),
      type(')'),
      enter(),
    ],
    'Departamento e salário': [
      cell('G4'), type('=PROCX(G1;C2:C9;B2:B9)', 'Departamento: só troca o 3º argumento para <code>B2:B9</code>'), enter(),
      type('=PROCX(G1;C2:C9;D2:D9)', 'Salário: <code>D2:D9</code>'), enter(),
    ],
    'Uma ficha que se atualiza': [cell('G1'), type('1077'), enter(), look('cell:G3:G5', 'A ficha inteira trocou de pessoa')],
    'Matrícula que não existe': [cell('G1'), type('9999'), enter(), look('cell:G3', '#N/D: matrícula não encontrada')],
    'O 4º argumento': [
      cell('G3'),
      type('=PROCX(G1;C2:C9;A2:A9;', 'Digite o PROCX de antes e um <code>;</code> a mais...'),
      type('"Matrícula não existe")', '...com o 4º argumento: o que mostrar se não encontrar'),
      enter(),
    ],
  });

  R('i5-3', {
    'ÍNDICE: valor por posição': [
      cell('H6'),
      type('=ÍNDICE(', 'Digite <code>=ÍNDICE(</code>'),
      point('B2:E6', 'A tabela: selecione só os valores, <b>B2:E6</b>'),
      type(';3;3)', 'Linha <code>3</code> e coluna <code>3</code>, e feche o parêntese'),
      enter(),
    ],
    'CORRESP: em que linha?': [
      cell('H4'),
      type('=CORRESP(', 'Digite <code>=CORRESP(</code>'),
      point('H1', 'O que procurar: a faixa em <b>H1</b>'),
      type(';'),
      point('A2:A6', 'Em que lista: as faixas <b>A2:A6</b>'),
      type(';0)', '<code>;0)</code> — o 0 pede busca exata'),
      enter(),
    ],
    'CORRESP: em que coluna?': [
      cell('H5'),
      type('=CORRESP('),
      point('H2', 'O que procurar: a região em <b>H2</b>'),
      type(';'),
      point('B1:E1', 'Em que lista: o cabeçalho das regiões <b>B1:E1</b> (na horizontal)'),
      type(';0)'),
      enter(),
    ],
    'Junte as peças': [
      cell('H6'),
      type('=ÍNDICE(B2:E6;', 'Digite <code>=ÍNDICE(B2:E6;</code>'),
      point('H4', 'A linha vem de <b>H4</b>'),
      type(';'),
      point('H5', 'A coluna vem de <b>H5</b>'),
      type(')'),
      enter(),
    ],
    'Teste com outra região': [cell('H2'), type('Sul'), enter(), look('cell:H4:H6', 'A coluna virou 1 e o frete mudou para R$ 38,00')],
    'Tudo numa fórmula só': [
      cell('H8'),
      type('=ÍNDICE(B2:E6;', 'Digite <code>=ÍNDICE(B2:E6;</code>'),
      type('CORRESP(H1;A2:A6;0);', 'No lugar da linha, o CORRESP da faixa: <code>CORRESP(H1;A2:A6;0);</code>'),
      type('CORRESP(H2;B1:E1;0))', 'No lugar da coluna, o CORRESP da região, e feche o ÍNDICE: <code>CORRESP(H2;B1:E1;0))</code>'),
      enter(),
    ],
  });

  // ================= Módulo 6 =================
  const homeTab = XL.DemoDSL.when((x) => x.tab() !== 'home', XL.DemoDSL.tab('home'));
  const { set, el, typeIn, drag, btn } = XL.DemoDSL;
  // Ctrl+1 › Personalizado › código › OK
  const customFormat = (from, to, code) => [
    drag(from, to),
    key('Ctrl+1', 'Pressione <kbd>Ctrl</kbd>+<kbd>1</kbd> para abrir <b>Formatar Células</b>'),
    el('.fc-cat[data-cat="custom"]', 'Em <b>Categoria</b>, clique em <b>Personalizado</b>'),
    typeIn('input[data-role="code"]', code, `No campo <b>Tipo</b>, digite <code>${XL.escapeHtml(code)}</code> — veja o <b>Exemplo</b> mudar`),
    el('.ui-modal .btn-primary', 'Clique em <b>OK</b>'),
  ];
  const flash = (ref, example, say) => [
    cell(ref),
    type(example, say || `Digite o exemplo: <code>${XL.escapeHtml(example)}</code>`),
    enter(),
    key('Ctrl+E', 'Pressione <kbd>Ctrl</kbd>+<kbd>E</kbd>: o Excel aprende com o exemplo e completa a coluna'),
  ];

  R('i6-1', {
    'Arrume os nomes': writeAndFill('C2', 'C6', '=PRI.MAIÚSCULA(ARRUMAR(A2))',
      'Digite <code>=PRI.MAIÚSCULA(ARRUMAR(A2))</code> — o ARRUMAR limpa os espaços e o PRI.MAIÚSCULA ajeita as letras'),
    'ESQUERDA: o começo do texto': writeAndFill('D2', 'D6', '=ESQUERDA(B2;2)', 'Digite <code>=ESQUERDA(B2;2)</code> — 2 caracteres do começo'),
    'DIREITA: o final do texto': writeAndFill('E2', 'E6', '=DIREITA(B2;4)', 'Digite <code>=DIREITA(B2;4)</code> — 4 caracteres do final'),
    'EXT.TEXTO: o meio do texto': writeAndFill('F2', 'F6', '=EXT.TEXTO(B2;4;4)', 'Digite <code>=EXT.TEXTO(B2;4;4)</code> — começando no 4º caractere, pegue 4'),
    'Juntando textos com &': [
      cell('G2'),
      type('=', 'Digite <code>=</code>'),
      point('C2', 'Clique no nome arrumado, <b>C2</b>'),
      type('&" ("&', 'Cole um espaço e um parêntese: <code>&amp;" ("&amp;</code>'),
      point('D2', 'Clique na UF, <b>D2</b>'),
      type('&")"', 'E feche o parêntese como texto: <code>&amp;")"</code>'),
      enter(),
      fill('G2', 'G6'),
    ],
  });

  R('i6-2', {
    'O segredo das datas': [cell('B2'), homeTab, set('select[data-cmd="numfmt"]', 'geral', 'Na caixa de formato de número, escolha <b>Geral</b>')],
    'Volte para data': [cell('B2'), homeTab, set('select[data-cmd="numfmt"]', 'data', 'Na mesma caixa, escolha <b>Data Abreviada</b>')],
    'Data + dias': writeAndFill('D2', 'D6', '=B2+C2', 'Digite <code>=B2+C2</code> — data mais dias'),
    'Só dias úteis': writeAndFill('E2', 'E6', '=DIATRABALHO(B2;C2)', 'Digite <code>=DIATRABALHO(B2;C2)</code> — conta só de segunda a sexta'),
    'Quantos dias faltam?': [
      cell('F2'), type('=D2-I1', 'Digite <code>=D2-I1</code>'),
      key('F4', '<kbd>F4</kbd>: I1 vira <code>$I$1</code>, a data de hoje travada'),
      enter(), fill('F2', 'F6'),
    ],
    'HOJE(): a data que se atualiza': [cell('I1'), type('=HOJE()', 'Digite <code>=HOJE()</code> — sem nada entre os parênteses'), enter()],
  });

  R('i6-3', {
    'Zeros à esquerda': customFormat('A2', 'A6', '0000'),
    'Unidade junto do número': customFormat('D2', 'D6', '#.##0,0 "kg"'),
    'Data no meio de uma frase': writeAndFill('E2', 'E6', '="Pedido de "&TEXTO(B2;"dd/mm/aaaa")',
      'Digite <code>="Pedido de "&amp;TEXTO(B2;"dd/mm/aaaa")</code>'),
    'O nome do mês': writeAndFill('F2', 'F6', '=TEXTO(B2;"mmmm")', 'Digite <code>=TEXTO(B2;"mmmm")</code> — mmmm = mês por extenso'),
    'Valor formatado numa frase': writeAndFill('G2', 'G6', '="Total: "&TEXTO(C2;"R$ #.##0,00")'),
  });

  R('i6-4', {
    'Texto para Colunas': [
      drag('A1', 'A7', 'Selecione a coluna com os dados: arraste de <b>A1</b> até <b>A7</b>'),
      btn('texttocols', undefined, 'Clique em <b>Texto para Colunas</b> (guia Dados)'),
      look('input[data-delim="semicolon"]', 'O separador <b>Ponto e vírgula</b> já veio marcado'),
      look('.tc-preview', 'A visualização mostra como cada linha vai ser dividida'),
      el('.ui-modal .btn-primary', 'Clique em <b>Concluir</b>'),
    ],
    'Preenchimento Relâmpago': flash('D2', 'Ana', 'Digite só o primeiro exemplo: <code>Ana</code>'),
    'O login do e-mail': flash('E2', 'ana.souza', 'Digite o exemplo: <code>ana.souza</code> (o que vem antes do @)'),
    'Reorganizando e mudando letras': flash('F2', 'SOUZA, Ana', 'Digite o exemplo: <code>SOUZA, Ana</code>'),
  });

  R('i3-4', {
    'Calcule o preço médio': writeAndFill('D2', 'D7', '=B2/C2'),
    'O erro se espalha': [
      cell('D8'), type('=MÉDIA(D2:D7)'), enter(),
      look('cell:D8', 'A média herdou o erro das células de cima'),
    ],
    'SEERRO ao resgate': [
      cell('D2'),
      type('=SEERRO(', 'Digite <code>=SEERRO(</code>'),
      type('B2/C2;', 'Primeiro a conta que pode falhar: <code>B2/C2</code> e <code>;</code>'),
      type('"")', 'Depois o que mostrar se der erro: <code>""</code> (vazio) e feche o parêntese'),
      enter(),
      fill('D2', 'D7'),
      look('cell:D8', 'Sem erros na coluna, a média voltou a funcionar'),
    ],
  });
})();
