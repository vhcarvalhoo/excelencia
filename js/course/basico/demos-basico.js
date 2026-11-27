/* Demonstrações "Mostre-me como" do Nível Básico.
   Cada passo (pelo título) lista as ações que o cursor do guia executa, explicando uma a uma. */
(function () {
  'use strict';
  const XL = window.XL;
  const { cell, drag, point, type, key, tab, btn, menu, item, el, rclick, dbl, set, typeIn, fill, scroll, look, sheet, when } = XL.DemoDSL;
  const R = XL.registerDemos;
  const enter = (say) => key('Enter', say || 'Pressione <kbd>Enter</kbd> para confirmar');
  const homeTab = when((x) => x.tab() !== 'home', tab('home'));
  const hasPreview = () => !!document.querySelector('.pp-overlay');
  const openPreview = when(() => !hasPreview(), tab('file', 'Abra a visualização: clique em <b>Arquivo</b>'));

  // ================= Módulo 1 =================
  R('b1-1', {
    'Selecione uma célula': [cell('C3', 'Encontre a coluna <b>C</b> e a linha <b>3</b> e clique na célula onde elas se cruzam: <b>C3</b>')],
    'Navegue com o teclado': [
      cell('C3', 'Partindo da célula <b>C3</b>...'),
      key('ArrowRight', 'Seta para a <b>direita</b> (→): vamos para D3'),
      key('ArrowRight', 'Mais uma vez <b>→</b>: E3'),
      key('ArrowDown', 'Seta para <b>baixo</b> (↓): E4'),
      key('ArrowDown', 'Mais uma vez <b>↓</b>: chegamos em <b>E5</b>'),
    ],
    'Pule direto para uma célula': [
      el('.xl-namebox', 'Clique na <b>Caixa de Nome</b> (o campo que mostra o endereço da célula)'),
      typeIn('.xl-namebox', 'H10', 'Digite o endereço: <code>H10</code>'),
      enter('Pressione <kbd>Enter</kbd> — a seleção pula para H10'),
    ],
    'Selecione um intervalo': [drag('B2', 'D6')],
    'Explore as guias': [tab('formulas')],
    'De volta à Página Inicial': [tab('home')],
  });

  R('b1-2', {
    'Seu primeiro dado': [cell('A1'), type('Produto', 'Digite <code>Produto</code> — repare que a barra de status mostra "Digite"'), enter()],
    'Agora com Tab': [cell('B1'), type('Preço'), enter()],
    'Números com vírgula': [
      cell('A2'), type('Caneta'),
      key('Tab', 'Pressione <kbd>Tab</kbd>: confirma e vai para a direita (B2)'),
      type('2,50', 'Digite <code>2,50</code> — no Brasil, decimal é com <b>vírgula</b>'), enter(),
    ],
    'Complete a lista': [
      cell('A3'), type('Caderno'), key('Tab', '<kbd>Tab</kbd> para ir ao preço'), type('15'),
      enter('<kbd>Enter</kbd>: o Excel volta para a coluna A, na linha de baixo'),
      type('Mochila'), key('Tab'), type('89,90'), enter(),
    ],
    'Editar sem apagar tudo': [
      cell('A2'),
      key('F2', 'Pressione <kbd>F2</kbd>: entra no modo de edição com o cursor no fim do texto'),
      type(' azul', 'Digite um espaço e <code>azul</code> — o texto antigo continua lá'), enter(),
    ],
    'Apagar o conteúdo': [cell('B4'), key('Delete', 'Pressione <kbd>Delete</kbd>: apaga o conteúdo da célula')],
    'Desfazer': [key('Ctrl+Z', 'Segure <kbd>Ctrl</kbd> e aperte <kbd>Z</kbd>: desfaz a última ação')],
  });

  R('b1-3', {
    'Faça uma conta': [cell('A1'), type('=10+5', 'Digite <code>=10+5</code> — toda fórmula começa com <b>=</b>'), enter()],
    'Resultado x fórmula': [cell('A1'), look('.xl-finput', 'Olhe a <b>Barra de Fórmulas</b>: ela mostra <code>=10+5</code>, enquanto a célula mostra 15')],
    'Calcule o total': [
      cell('B5'), type('=', 'Digite <code>=</code> para começar a fórmula'),
      point('B3', 'Em vez de digitar, <b>clique em B3</b>: a referência entra sozinha (modo Aponte)'),
      type('*', 'Digite <code>*</code> (multiplicação)'),
      point('B4', 'Agora <b>clique em B4</b>'),
      enter('Pressione <kbd>Enter</kbd>: o resultado aparece'),
    ],
    'A mágica do recálculo': [cell('B3'), type('10'), enter(), look('cell:B5', 'Olhe o total em <b>B5</b>: mudou sozinho para 125')],
    'Use parênteses': [cell('D1'), type('=(2+3)*4'), enter()],
    'Agora sem parênteses': [cell('D2'), type('=2+3*4'), enter()],
    'Conhecendo os erros': [cell('D3'), type('=10/0', 'Digite <code>=10/0</code> (dividir por zero não existe)'), enter()],
  });

  const fn = (ref, name) => [
    cell(ref), type(`=${name}(`, `Digite <code>=${name}(</code> — a dica amarela mostra os argumentos`),
    point('B2:B7', 'Arraste de <b>B2</b> até <b>B7</b>: o intervalo entra na fórmula'),
    type(')', 'Feche o parêntese <code>)</code>'), enter(),
  ];
  R('b1-4', {
    'SOMA': [
      cell('B8'),
      type('=SO', 'Digite <code>=SO</code> — aparece a lista de funções que começam assim'),
      key('Tab', 'Com <b>SOMA</b> destacada, pressione <kbd>Tab</kbd> para escolher'),
      point('B2:B7', 'Arraste de <b>B2</b> até <b>B7</b>: o intervalo entra na fórmula'),
      type(')', 'Feche o parêntese <code>)</code>'), enter(),
    ],
    'MÉDIA': fn('B9', 'MÉDIA'),
    'MÁXIMO': fn('B10', 'MÁXIMO'),
    'MÍNIMO': fn('B11', 'MÍNIMO'),
    'Contando números': fn('B13', 'CONT.NÚM'),
    'O atalho da barra de status': [drag('C2', 'C7'), look('.xl-status-stats', 'Olhe o rodapé: <b>Média</b>, <b>Contagem</b> e <b>Soma</b> da seleção, sem fórmula nenhuma')],
    'AutoSoma: o botão mágico': [
      cell('C8'),
      btn('autosum', 'SOMA', 'Clique no botão <b>Σ AutoSoma</b> (grupo Edição)'),
      look('cell:C2:C7', 'O Excel escreveu <code>=SOMA(C2:C7)</code> e marcou o intervalo sugerido'),
      enter('Está certo? Pressione <kbd>Enter</kbd>'),
    ],
  });

  R('b1-5', {
    'Cabeçalho em negrito': [drag('A3', 'C3'), btn('bold', undefined, 'Clique em <b>N</b> (Negrito)')],
    'Cor de preenchimento': [
      drag('A3', 'C3'),
      menu('fill', 'Clique na <b>setinha</b> ao lado do balde de tinta'),
      el('.ui-pal-grid.tints .ui-sw:nth-child(10)', 'Escolha uma cor clara (aqui, um verde)'),
    ],
    'Centralizar': [drag('A3', 'C3'), btn('align', 'center', 'Clique em <b>Centralizar</b> (grupo Alinhamento)')],
    'Formato de dinheiro': [drag('B4', 'B8'), btn('currency', undefined, 'Clique no botão da <b>nota de dinheiro</b> (formato Contábil, R$)')],
    'Porcentagem': [drag('C4', 'C8'), btn('percent', undefined, 'Clique em <b>%</b> (Estilo de Porcentagem)')],
    'Casas decimais': [drag('C4', 'C8'), btn('dec-inc', undefined, 'Clique em <b>Aumentar Casas Decimais</b>')],
    'Um título de verdade': [
      cell('A1'),
      when((x) => !x.fmt('A1').b, btn('bold', undefined, 'Clique em <b>N</b> (Negrito)')),
      homeTab,
      set('select[data-cmd="size"]', '16', 'Na caixa de <b>tamanho da fonte</b>, escolha <b>16</b>'),
    ],
    'Bordas': [
      drag('A3', 'C8'),
      menu('border', 'Clique na <b>setinha</b> ao lado do botão de Bordas'),
      item('border', 'all', 'Escolha <b>Todas as Bordas</b>'),
    ],
  });

  R('b1-6', {
    'Primeiro, uma fórmula': [cell('D2'), type('='), point('B2'), type('*'), point('C2'), enter()],
    'Arraste para copiar': [fill('D2', 'D6')],
    'O que aconteceu com a fórmula?': [cell('D4'), look('.xl-finput', 'A fórmula virou <code>=B4*C4</code>: as referências desceram junto')],
    'Trave a referência com $': [
      cell('E2'),
      type('=D2*(1-H1', 'Digite <code>=D2*(1-H1</code>'),
      key('F4', 'Com o cursor logo depois de H1, pressione <kbd>F4</kbd>: vira <code>$H$1</code> (travado)'),
      type(')', 'Feche o parêntese'), enter(),
    ],
    'Arraste de novo': [fill('E2', 'E6')],
    'Mude o desconto': [cell('H1'), type('15%'), enter()],
    'Sequências automáticas': [
      cell('J1'), type('1'), enter(), type('2'), enter(),
      fill('J1:J2', 'J10', 'Com <b>J1:J2</b> selecionados, arraste a alça até <b>J10</b>: o Excel continua o padrão'),
    ],
    'Meses e dias da semana': [cell('L1'), type('Jan'), enter(), fill('L1', 'L6')],
  });

  // ================= Módulo 2 =================
  R('b2-1', {
    'Ordem alfabética': [
      cell('A5', 'Clique em <b>uma única célula</b> da coluna Vendedor (não selecione a coluna toda)'),
      tab('data'), btn('sort', 'asc', 'Clique em <b>A→Z</b> (Classificar de A a Z)'),
    ],
    'Maiores vendas primeiro': [cell('D3', 'Clique numa célula da coluna <b>Valor</b>'), btn('sort', 'desc', 'Clique em <b>Z→A</b> (do maior para o menor)')],
    'Dois critérios ao mesmo tempo': [
      cell('B3'),
      btn('sort-dialog', undefined, 'Clique no botão <b>Classificar</b> (guia Dados)'),
      set('.sd-levels select[data-level="0"][data-role="col"]', '1', 'Em <b>Classificar por</b>, escolha <b>Região</b> (ordem De A a Z)'),
      el('.sd-add', 'Clique em <b>+ Adicionar Nível</b>'),
      set('.sd-levels select[data-level="1"][data-role="col"]', '3', 'Em <b>E depois por</b>, escolha <b>Valor</b>'),
      set('.sd-levels select[data-level="1"][data-role="order"]', 'desc', 'Na ordem, escolha <b>Do Maior para o Menor</b>'),
      el('.ui-modal .btn-primary', 'Clique em <b>OK</b>'),
    ],
  });

  R('b2-2', {
    'Ligue o filtro': [cell('A6', 'Clique em qualquer célula da tabela'), btn('filter-toggle', undefined, 'Clique em <b>Filtro</b> (guia Dados) — atalho <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd>')],
    'Só a região Sul': [
      el('.xl-filterbtn[data-c="1"]', 'Clique na <b>setinha</b> do cabeçalho <b>Região</b>'),
      el('.xl-fitem.all input', 'Desmarque <b>(Selecionar Tudo)</b>'),
      el('.xl-fitem[data-value="Sul"] input', 'Marque só <b>Sul</b>'),
      el('.xl-fmenu-ok', 'Clique em <b>OK</b>'),
    ],
    'SOMA x SUBTOTAL': [cell('G2'), type('=SUBTOTAL(9;D5:D16)', 'Digite <code>=SUBTOTAL(9;D5:D16)</code> — o 9 significa "somar"'), enter()],
    'Limpe o filtro da Região': [
      el('.xl-filterbtn[data-c="1"]', 'Clique na setinha de <b>Região</b> (agora com um funil)'),
      el('.ui-menu-item[data-cmd="fclear"]', 'Clique em <b>Limpar Filtro de “Região”</b>'),
    ],
    'Filtro de número': [
      el('.xl-filterbtn[data-c="3"]', 'Clique na setinha de <b>Valor</b>'),
      el('.ui-menu-item[data-cmd="fcondmenu"]', 'Clique em <b>Filtros de Número</b>'),
      el('.ui-menu-item[data-cmd="fcond"][data-arg=">"]', 'Escolha <b>É Maior do que...</b>'),
      typeIn('.ui-modal input[data-role="val"]', '5000', 'Digite <code>5000</code>'),
      el('.ui-modal .btn-primary', 'Clique em <b>OK</b>'),
    ],
    'Desligue o filtro': [btn('filter-toggle', undefined, 'Clique de novo em <b>Filtro</b> para desligar')],
  });

  R('b2-3', {
    'Inserir uma linha': [
      rclick('.xl-rowh[data-r="3"]', 'Clique com o <b>botão direito</b> no número da linha <b>4</b>'),
      el('.ui-menu-item[data-cmd="insert-rows"]', 'Escolha <b>Inserir</b>'),
    ],
    'Preencha março': [cell('A4'), type('Mar'), key('Tab'), type('1280'), enter()],
    'O total se ajustou sozinho': [cell('B8'), look('.xl-finput', 'A fórmula agora é <code>=SOMA(B2:B7)</code>: o intervalo cresceu sozinho')],
    'Excluir uma linha': [
      rclick('.xl-rowh[data-r="6"]', 'Clique com o <b>botão direito</b> no número da linha <b>7</b> (a do Teste)'),
      el('.ui-menu-item[data-cmd="delete-rows"]', 'Escolha <b>Excluir</b>'),
    ],
    'Excluir uma coluna': [
      rclick('.xl-colh[data-c="2"]', 'Clique com o <b>botão direito</b> na letra da coluna <b>C</b>'),
      el('.ui-menu-item[data-cmd="delete-cols"]', 'Escolha <b>Excluir</b>'),
    ],
  });

  R('b2-4', {
    'Planilhas grandes': [scroll(700, null, 'Role <b>para baixo</b> com a rodinha do mouse... e o cabeçalho some')],
    'Congele a linha do cabeçalho': [tab('view'), menu('freeze', 'Clique em <b>Congelar Painéis</b>'), item('freeze', 'top', 'Escolha <b>Congelar Linha Superior</b>')],
    'Teste rolando': [scroll(700, null, 'Role para baixo: agora o cabeçalho fica parado no topo')],
    'E se rolar para o lado?': [menu('freeze', 'Clique em <b>Congelar Painéis</b>'), item('freeze', 'none', 'Escolha <b>Descongelar Painéis</b>')],
    'Congele linha e coluna': [
      cell('B2', 'Clique em <b>B2</b>: tudo acima (linha 1) e à esquerda (coluna A) vai ficar travado'),
      menu('freeze', 'Clique em <b>Congelar Painéis</b>'),
      item('freeze', 'panes', 'Escolha <b>Congelar Painéis</b>'),
    ],
    'Role para a direita': [scroll(null, 1400, 'Role <b>para a direita</b>: a coluna dos clientes não sai do lugar')],
  });

  R('b2-5', {
    'Troque de planilha': [sheet('Fevereiro', 'Clique na aba <b>Fevereiro</b>, no rodapé')],
    'Crie uma planilha nova': [el('.xl-sheetadd', 'Clique no <b>+</b> ao lado das abas')],
    'Dê um nome de verdade': [
      dbl('.xl-sheettab.active', 'Dê um <b>duplo clique</b> na aba nova'),
      typeIn('.xl-tabinput', 'Resumo', 'Digite <code>Resumo</code>'),
      enter(),
    ],
    'Buscando dados de outra planilha': [
      cell('B2'), type('=', 'Digite <code>=</code> para começar a fórmula'),
      sheet('Janeiro', 'Sem confirmar, clique na aba <b>Janeiro</b> — a fórmula continua aberta na barra'),
      point('B6', 'Clique no total, <b>B6</b>: a fórmula vira <code>=Janeiro!B6</code>'),
      enter('Pressione <kbd>Enter</kbd>: o Excel grava e volta para o Resumo'),
    ],
    'Agora fevereiro': [
      cell('B3'), type('='), sheet('Fevereiro', 'Clique na aba <b>Fevereiro</b>'), point('B6'),
      enter('Pressione <kbd>Enter</kbd>: volta para o Resumo'),
    ],
    'Total do bimestre': [cell('B4'), type('=SOMA('), point('B2:B3'), type(')'), enter()],
    'Tudo conectado': [sheet('Janeiro'), cell('B5', 'Clique no valor do <b>Lazer</b> (B5)'), type('400'), enter(), sheet('Resumo', 'Volte ao <b>Resumo</b> e veja os totais')],
    'Renomear não quebra nada': [
      dbl('.xl-sheettab[data-sheet="Janeiro"]', 'Dê um <b>duplo clique</b> na aba <b>Janeiro</b>'),
      typeIn('.xl-tabinput', 'Jan', 'Digite <code>Jan</code>'), enter(),
      sheet('Resumo'), cell('B2'),
      look('.xl-finput', 'A fórmula se atualizou sozinha: <code>=Jan!B6</code>'),
    ],
  });

  R('b2-6', {
    'Localizar': [
      key('Ctrl+L', 'Pressione <kbd>Ctrl</kbd>+<kbd>L</kbd> para abrir o <b>Localizar</b>'),
      typeIn('.xl-find input[data-role="find"]', 'Rafael', 'Digite <code>Rafael</code>'),
      el('.xl-find button[data-role="findnext"]', 'Clique em <b>Localizar próxima</b>'),
    ],
    'Substituir tudo': [
      when(() => !document.querySelector('.xl-find'), key('Ctrl+U', 'Pressione <kbd>Ctrl</kbd>+<kbd>U</kbd> para abrir o <b>Substituir</b>')),
      when(() => !!document.querySelector('.xl-find'), el('.xl-find .fr-tab[data-tab="replace"]', 'Clique na aba <b>Substituir</b>')),
      typeIn('.xl-find input[data-role="find"]', 'S. Paulo', 'Em Localizar, digite <code>S. Paulo</code>'),
      typeIn('.xl-find input[data-role="replace"]', 'São Paulo', 'Em Substituir por, digite <code>São Paulo</code>'),
      el('.xl-find button[data-role="replaceall"]', 'Clique em <b>Substituir tudo</b>'),
    ],
    'Remover duplicatas': [
      cell('A2', 'Clique em qualquer célula da lista'),
      btn('dedupe', undefined, 'Clique em <b>Remover Duplicatas</b> (guia Dados)'),
      el('.ui-modal .btn-primary', 'Deixe todas as colunas marcadas e clique em <b>OK</b>'),
      el('.ui-modal .btn-primary', 'Leia o resultado e clique em <b>OK</b>'),
    ],
  });

  R('b2-7', {
    'Antes de imprimir, visualize': [tab('file', 'Clique em <b>Arquivo</b> (atalho: <kbd>Ctrl</kbd>+<kbd>P</kbd>)')],
    'Deite a página': [openPreview, set('.pp-overlay select[data-role="orientation"]', 'paisagem', 'Em <b>Orientação</b>, escolha <b>Paisagem</b>')],
    'Todas as colunas em uma página': [openPreview, set('.pp-overlay select[data-role="fit"]', 'cols', 'Em <b>Dimensionamento</b>, escolha <b>Ajustar Todas as Colunas em Uma Página</b>')],
    'Volte para a planilha': [when(hasPreview, el('.pp-back', 'Clique em <b>← Voltar à planilha</b>'))],
    'Repita o cabeçalho': [
      btn('print-titles', undefined, 'Na guia <b>Layout da Página</b>, clique em <b>Imprimir Títulos</b>'),
      typeIn('input[data-role="titles"]', '$1:$1', 'Digite <code>$1:$1</code> (repetir a linha 1)'),
      el('.ui-modal .btn-primary', 'Clique em <b>OK</b>'),
    ],
    'Imprima só uma parte': [
      drag('A1', 'F20'),
      menu('print-area', 'Clique em <b>Área de Impressão</b> (guia Layout da Página)'),
      item('print-area', 'set', 'Escolha <b>Definir Área de Impressão</b>'),
    ],
    'Confira o resultado': [tab('file', 'Clique em <b>Arquivo</b> para ver a visualização')],
  });
})();
