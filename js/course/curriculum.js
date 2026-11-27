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
          { id: 'm3', title: 'Lógica e condições', lessons: [
            { id: 'i3-1', title: 'Função SE' }, { id: 'i3-2', title: 'E, OU e NÃO' },
            { id: 'i3-3', title: 'SE aninhado e SES' }, { id: 'i3-4', title: 'Tratando erros com SEERRO' },
          ] },
          { id: 'm4', title: 'Contar e somar com critérios', lessons: [
            { id: 'i4-1', title: 'CONT.SE e CONT.SES' }, { id: 'i4-2', title: 'SOMASE e SOMASES' }, { id: 'i4-3', title: 'MÉDIASE e critérios com curingas' },
          ] },
          { id: 'm5', title: 'Buscando dados', lessons: [
            { id: 'i5-1', title: 'PROCV' }, { id: 'i5-2', title: 'PROCX: a busca moderna' }, { id: 'i5-3', title: 'ÍNDICE + CORRESP' },
          ] },
          { id: 'm6', title: 'Texto e datas', lessons: [
            { id: 'i6-1', title: 'Funções de texto (ESQUERDA, EXT.TEXTO, CONCAT...)' }, { id: 'i6-2', title: 'Datas e prazos' },
            { id: 'i6-3', title: 'Formatos personalizados e TEXTO' }, { id: 'i6-4', title: 'Texto para Colunas e Preenchimento Relâmpago' },
          ] },
          { id: 'm7', title: 'Tabelas e visualização', lessons: [
            { id: 'i7-1', title: 'Formatar como Tabela' }, { id: 'i7-2', title: 'Gráficos de colunas, linhas e pizza' },
            { id: 'i7-3', title: 'Formatação condicional' }, { id: 'i7-4', title: 'Validação de dados e listas suspensas' },
          ] },
        ],
      },
      {
        id: 'avancado', name: 'Avançado', tag: 'Nível 3', color: '#7c3aed',
        desc: 'Análise de verdade: tabelas dinâmicas, fórmulas de matriz dinâmica e dashboards.',
        modules: [
          { id: 'm8', title: 'Tabelas Dinâmicas', lessons: [
            { id: 'a8-1', title: 'Criando sua primeira Tabela Dinâmica' }, { id: 'a8-2', title: 'Agrupamentos e campos calculados' },
            { id: 'a8-3', title: 'Segmentação de dados e Linha do Tempo' }, { id: 'a8-4', title: 'Gráficos dinâmicos' },
          ] },
          { id: 'm9', title: 'Fórmulas modernas', lessons: [
            { id: 'a9-1', title: 'Matrizes dinâmicas: FILTRO, CLASSIFICAR, ÚNICO' }, { id: 'a9-2', title: 'LET: fórmulas legíveis' },
            { id: 'a9-3', title: 'LAMBDA: crie suas próprias funções' }, { id: 'a9-4', title: 'SOMARPRODUTO e cálculos com matrizes' },
          ] },
          { id: 'm10', title: 'Análise e cenários', lessons: [
            { id: 'a10-1', title: 'Atingir Meta' }, { id: 'a10-2', title: 'Gerenciador de Cenários e Tabela de Dados' },
            { id: 'a10-3', title: 'Solver' }, { id: 'a10-4', title: 'Previsões e tendências' },
          ] },
          { id: 'm11', title: 'Dashboards', lessons: [
            { id: 'a11-1', title: 'Princípios de um bom dashboard' }, { id: 'a11-2', title: 'KPIs e minigráficos' }, { id: 'a11-3', title: 'Dashboard interativo completo' },
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
