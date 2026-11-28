/* Trilha completa do curso: do Básico ao Expert. Lições com "id" registrado em XL.Lessons ficam disponíveis;
   as demais aparecem como "em breve" para mostrar o caminho inteiro. */
(function () {
  'use strict';
  const XL = window.XL;

  XL.Course = {
    levels: [
      {
        id: 'basico', name: 'Básico', tag: 'Nível 1', color: '#107c41',
        desc: 'Do zero: a tela do Excel, digitar dados, primeiras fórmulas, formatação e organização.',
        modules: [
          {
            id: 'm1', title: 'Primeiros passos', desc: 'A base de tudo: navegar, digitar, calcular e formatar.',
            lessons: [
              { id: 'b1-1', title: 'Conhecendo o Excel', minutes: 5, desc: 'Faixa de Opções, células, Caixa de Nome e intervalos.' },
              { id: 'b1-2', title: 'Digitando e editando dados', minutes: 6, desc: 'Texto, números, editar, apagar e desfazer.' },
              { id: 'b1-3', title: 'Sua primeira fórmula', minutes: 7, desc: 'O sinal de =, operadores, referências e erros comuns.' },
              { id: 'b1-4', title: 'Funções essenciais', minutes: 8, desc: 'SOMA, MÉDIA, MÁXIMO, MÍNIMO, AutoSoma e barra de status.' },
              { id: 'b1-5', title: 'Formatando a planilha', minutes: 7, desc: 'Negrito, cores, moeda, porcentagem, casas decimais e bordas.' },
              { id: 'b1-6', title: 'Alça de preenchimento e o $', minutes: 9, desc: 'Copiar fórmulas arrastando, referências absolutas e séries.' },
            ],
          },
          {
            id: 'm2', title: 'Organizando dados', desc: 'Planilhas maiores, sem bagunça.',
            lessons: [
              { id: 'b2-1', title: 'Classificar dados', minutes: 6, desc: 'A→Z, maior→menor e classificação por vários critérios.' },
              { id: 'b2-2', title: 'Filtrar dados', minutes: 8, desc: 'AutoFiltro, filtros de número e SOMA x SUBTOTAL.' },
              { id: 'b2-3', title: 'Inserir e excluir linhas e colunas', minutes: 6, desc: 'Abrir espaço sem estragar fórmulas (e o #REF!).' },
              { id: 'b2-4', title: 'Congelar painéis', minutes: 5, desc: 'Cabeçalho e primeira coluna sempre visíveis.' },
              { id: 'b2-5', title: 'Várias planilhas e referências entre elas', minutes: 9, desc: 'Abas, renomear e fórmulas como =Janeiro!B6.' },
              { id: 'b2-6', title: 'Localizar, substituir e remover duplicatas', minutes: 6, desc: 'Corrigir e limpar listas em segundos.' },
              { id: 'b2-7', title: 'Impressão e layout da página', minutes: 7, desc: 'Paisagem, ajustar à página, títulos e área de impressão.' },
            ],
          },
        ],
      },
      {
        id: 'intermediario', name: 'Intermediário', tag: 'Nível 2', color: '#2563eb',
        desc: 'Fórmulas que tomam decisões, buscam dados e resumem informações. Tabelas e gráficos.',
        modules: [
          { id: 'm3', title: 'Lógica e condições', desc: 'Fórmulas que decidem sozinhas.', lessons: [
            { id: 'i3-1', title: 'Função SE', minutes: 8, desc: 'Teste lógico, aspas nos textos e regra numa célula com $.' },
            { id: 'i3-2', title: 'E, OU e NÃO', minutes: 7, desc: 'Várias condições: bônus, destaque e treinamento.' },
            { id: 'i3-3', title: 'SE aninhado e SES', minutes: 7, desc: 'Mais de duas respostas: conceitos A, B, C, D.' },
            { id: 'i3-4', title: 'Tratando erros com SEERRO', minutes: 6, desc: 'Evitar que um #DIV/0! estrague o relatório.' },
          ] },
          { id: 'm4', title: 'Contar e somar com critérios', desc: 'Responder perguntas sobre os dados.', lessons: [
            { id: 'i4-1', title: 'CONT.SE e CONT.SES', minutes: 7, desc: 'Quantos pedidos? Quantos acima de R$ 2.000? E no Sul?' },
            { id: 'i4-2', title: 'SOMASE e SOMASES', minutes: 8, desc: 'Totais por vendedor, produto e um resumo por região.' },
            { id: 'i4-3', title: 'MÉDIASE e critérios com curingas', minutes: 7, desc: '"Mouse*", "*sem fio" e critérios vindos de células.' },
          ] },
          { id: 'm5', title: 'Buscando dados', desc: 'Trazer informações de outra tabela.', lessons: [
            { id: 'i5-1', title: 'PROCV', minutes: 9, desc: 'Nome e preço pelo código, FALSO, $ e o #N/D.' },
            { id: 'i5-2', title: 'PROCX: a busca moderna', minutes: 7, desc: 'Buscar para qualquer lado e tratar o "não encontrado".' },
            { id: 'i5-3', title: 'ÍNDICE + CORRESP', minutes: 8, desc: 'Busca em duas direções numa tabela de frete.' },
          ] },
          { id: 'm6', title: 'Texto e datas', desc: 'Limpar, separar e formatar informações.', lessons: [
            { id: 'i6-1', title: 'Funções de texto', minutes: 8, desc: 'ARRUMAR, ESQUERDA, DIREITA, EXT.TEXTO e o &.' },
            { id: 'i6-2', title: 'Datas e prazos', minutes: 8, desc: 'Datas são números: prazos, dias úteis e HOJE().' },
            { id: 'i6-3', title: 'Formatos personalizados e TEXTO', minutes: 8, desc: 'Ctrl+1 › Personalizado: 0000, "kg", mmmm e TEXTO.' },
            { id: 'i6-4', title: 'Texto para Colunas e Preenchimento Relâmpago', minutes: 7, desc: 'Separar um CSV e completar colunas com Ctrl+E.' },
          ] },
          { id: 'm7', title: 'Tabelas e visualização', desc: 'Dados organizados, destacados e em gráficos.', lessons: [
            { id: 'i7-1', title: 'Formatar como Tabela', minutes: 9, desc: 'Ctrl+T, linha de totais, [@Coluna] e a tabela que cresce sozinha.' },
            { id: 'i7-2', title: 'Gráficos de colunas, linhas e pizza', minutes: 8, desc: 'Escolher o gráfico certo, título, rótulos e trocar o tipo.' },
            { id: 'i7-3', title: 'Formatação condicional', minutes: 8, desc: 'Realçar regras, barras de dados, escalas de cor e fórmulas.' },
            { id: 'i7-4', title: 'Validação de dados e listas suspensas', minutes: 7, desc: 'Listas para escolher, limites de valores e alertas de erro.' },
          ] },
        ],
      },
      {
        id: 'avancado', name: 'Avançado', tag: 'Nível 3', color: '#7c3aed',
        desc: 'Análise de verdade: tabelas dinâmicas, fórmulas de matriz dinâmica e dashboards.',
        modules: [
          { id: 'm8', title: 'Tabelas Dinâmicas', desc: 'Resumos poderosos sem fórmulas, filtros com botões e gráficos ligados.', lessons: [
            { id: 'a8-1', title: 'Criando sua primeira Tabela Dinâmica', minutes: 10, desc: 'Linhas, Colunas, Valores, formato de moeda e classificar pelo valor.' },
            { id: 'a8-2', title: 'Agrupamentos e campos calculados', minutes: 9, desc: 'Agrupar datas por mês, Lucro = Valor − Custo, % do total e Atualizar.' },
            { id: 'a8-3', title: 'Segmentação de dados e Linha do Tempo', minutes: 7, desc: 'Filtros em botões: regiões, seleção múltipla e períodos.' },
            { id: 'a8-4', title: 'Gráficos dinâmicos', minutes: 7, desc: 'Gráfico ligado à Tabela Dinâmica que muda com campos e filtros.' },
          ] },
          { id: 'm9', title: 'Fórmulas modernas', desc: 'Uma fórmula, um relatório inteiro: matrizes dinâmicas, LET e LAMBDA.', lessons: [
            { id: 'a9-1', title: 'Matrizes dinâmicas: FILTRO, CLASSIFICAR, ÚNICO', minutes: 9, desc: 'Resultados que derramam, a referência F2# e o erro #DESPEJAR!.' },
            { id: 'a9-2', title: 'LET: fórmulas legíveis', minutes: 7, desc: 'Dar nomes a contas repetidas e a listas inteiras.' },
            { id: 'a9-3', title: 'LAMBDA: crie suas próprias funções', minutes: 8, desc: 'A função COMISSAO com Definir Nome e Gerenciador de Nomes.' },
            { id: 'a9-4', title: 'SOMARPRODUTO e cálculos com matrizes', minutes: 7, desc: 'Contas com intervalos inteiros e condições que viram 1 e 0.' },
          ] },
          { id: 'm10', title: 'Análise e cenários', desc: 'Perguntas "e se?", a melhor decisão e previsões.', lessons: [
            { id: 'a10-1', title: 'Atingir Meta', minutes: 7, desc: 'Valor do empréstimo, prazo necessário e ponto de equilíbrio.' },
            { id: 'a10-2', title: 'Gerenciador de Cenários e Tabela de Dados', minutes: 9, desc: 'Vários preços de uma vez e cenários Otimista/Pessimista com resumo.' },
            { id: 'a10-3', title: 'Solver', minutes: 9, desc: 'O mix de produção que dá o maior lucro, com restrições.' },
            { id: 'a10-4', title: 'Previsões e tendências', minutes: 8, desc: 'INCLINAÇÃO, PREVISÃO.LINEAR, TENDÊNCIA, R² e linha de tendência.' },
          ] },
          { id: 'm11', title: 'Dashboards', desc: 'Painéis claros e interativos para quem decide.', lessons: [
            { id: 'a11-1', title: 'Princípios de um bom dashboard', minutes: 6, desc: 'Ordenar, escolher o gráfico, título com mensagem e menos ruído.' },
            { id: 'a11-2', title: 'KPIs e minigráficos', minutes: 8, desc: 'Total, variação, destaque do que piorou e gráficos dentro das células.' },
            { id: 'a11-3', title: 'Dashboard interativo completo', minutes: 9, desc: 'Gráficos dinâmicos e uma segmentação ligada a todo o painel.' },
          ] },
        ],
      },
      {
        id: 'expert', name: 'Expert', tag: 'Nível 4', color: '#b45309',
        desc: 'O que existe de mais poderoso: Power Query, linguagem M, Power Pivot, DAX e automação.',
        modules: [
          { id: 'm12', title: 'Power Query: fundamentos', lessons: [
            { id: 'e12-1', title: 'O que é ETL e o Editor do Power Query' }, { id: 'e12-2', title: 'Importando de CSV, Excel e pastas' },
            { id: 'e12-3', title: 'Limpeza: tipos, colunas, preencher, dividir' }, { id: 'e12-4', title: 'Etapas aplicadas e atualização com um clique' },
          ] },
          { id: 'm13', title: 'Power Query: transformações avançadas', lessons: [
            { id: 'e13-1', title: 'Mesclar consultas (joins)' }, { id: 'e13-2', title: 'Acrescentar consultas' },
            { id: 'e13-3', title: 'Transformar colunas em linhas (unpivot) e dinamizar' }, { id: 'e13-4', title: 'Agrupar por' },
            { id: 'e13-5', title: 'Colunas personalizadas e condicionais' }, { id: 'e13-6', title: 'Parâmetros e consultas como funções' },
          ] },
          { id: 'm14', title: 'Linguagem M', lessons: [
            { id: 'e14-1', title: 'Estrutura let ... in' }, { id: 'e14-2', title: 'Listas, registros e tabelas' },
            { id: 'e14-3', title: 'Tratamento de erros em M' }, { id: 'e14-4', title: 'Performance e query folding' },
          ] },
          { id: 'm15', title: 'Power Pivot e DAX', lessons: [
            { id: 'e15-1', title: 'Modelo de dados e relacionamentos' }, { id: 'e15-2', title: 'Esquema estrela' },
            { id: 'e15-3', title: 'Medidas DAX e CALCULATE' }, { id: 'e15-4', title: 'Inteligência de tempo' },
          ] },
          { id: 'm16', title: 'Automação', lessons: [
            { id: 'e16-1', title: 'Macros e VBA' }, { id: 'e16-2', title: 'Office Scripts' }, { id: 'e16-3', title: 'Python no Excel' },
          ] },
        ],
      },
    ],
  };

  // Lista linear de lições (para "próxima lição").
  XL.Course.flat = () => {
    const out = [];
    for (const lv of XL.Course.levels) for (const m of lv.modules) for (const l of m.lessons) out.push({ level: lv, module: m, lesson: l });
    return out;
  };
  XL.Course.find = (id) => XL.Course.flat().find((x) => x.lesson.id === id);
})();
