# Excelência — aprenda Excel praticando

Aplicativo que ensina Excel do básico ao Expert com um **simulador do Excel** e um **guia passo a passo** que confere cada tarefa automaticamente.

## Como abrir

Não precisa instalar nada: dê **duplo clique em `index.html`** (Chrome ou Edge).

Para desenvolver com recarregamento simples, também dá para servir a pasta:

```powershell
python -m http.server 8780
```

e abrir `http://localhost:8780`.

**Ao publicar mudanças**, troque o número de versão `?v=AAAAMMDD` de todos os arquivos no `index.html` (procure e substitua). Assim o navegador de quem já usou o app baixa os arquivos novos em vez de usar cópias antigas do cache — sem isso, lições novas podem aparecer bloqueadas.

## O que já existe

- **Simulador do Excel (pt-BR):** Faixa de Opções com as guias, Caixa de Nome, Barra de Fórmulas, grade com seleção por mouse/teclado, edição na célula, modo *Aponte* (clicar em células enquanto digita a fórmula), referências coloridas, autocompletar de funções, dica de sintaxe, F4 para `$`, alça de preenchimento (cópia, séries numéricas, meses, dias), desfazer/refazer, copiar/colar (inclusive com o Excel real), Pincel de Formatação, AutoSoma, Inserir Função, Mostrar Fórmulas, barra de status com Média/Contagem/Soma, formatos (Moeda, Contábil, %, Data...), bordas, cores, estilos de célula.
- **Organização de dados (Módulo 2):** várias planilhas (abas, renomear, excluir, referências como `=Janeiro!B6` — inclusive clicando na aba enquanto digita a fórmula), inserir/excluir linhas e colunas com ajuste automático de fórmulas (e `#REF!`), classificar (rápido e com vários níveis), AutoFiltro (valores, filtros de número/texto, linhas azuis, contagem na barra de status, `SUBTOTAL`), remover duplicatas, congelar painéis, Localizar/Substituir, menu do botão direito, Formatar Células (Ctrl+1) com formatos personalizados (`0000`, `#.##0,0 "kg"`, `dd/mm/aaaa`...), Texto para Colunas, Preenchimento Relâmpago (Ctrl+E), Layout da Página (margens, orientação, tamanho, área de impressão, imprimir títulos, ajustar à página) e Visualização de Impressão com contagem de páginas e quebras tracejadas na grade.
- **Tabelas e visualização (Módulo 7):** Formatar como Tabela (Ctrl+T, estilos, linhas em tiras, Linha de Totais, renomear, Converter em Intervalo), referências estruturadas (`[@Qtd]`, `[@[Preço unit.]]`, `Vendas[Total]`), colunas calculadas e tabela que cresce ao digitar logo abaixo; gráficos em SVG (colunas, barras, linhas, pizza, Gráficos Recomendados) que se movem, redimensionam e acompanham os dados, com título, legenda e rótulos; formatação condicional (realçar regras, primeiros/últimos, barras de dados, escalas de cor, regra com fórmula); validação de dados (lista suspensa com seta e Alt+↓, número inteiro/decimal/data/tamanho do texto, mensagem de entrada e alerta de erro). As guias contextuais **Design da Tabela** e **Design do Gráfico** aparecem só quando a tabela/o gráfico está selecionado.
- **Tabelas Dinâmicas (Módulo 8):** Inserir › Tabela Dinâmica (fonte: Tabela ou intervalo; nova planilha ou existente), Lista de Campos com arrastar e soltar entre Filtros, Colunas, Linhas e Valores, Configurações do Campo de Valor (Soma, Contagem, Média, Máx., Mín.; % do Total Geral/Coluna/Linha; formato do número), setinhas para filtrar e classificar (inclusive pelo valor), filtro de relatório, agrupar datas (meses, trimestres, anos) e números em faixas, Campo Calculado, Atualizar/Atualizar Tudo (a Tabela Dinâmica não muda sozinha quando a fonte muda), subtotais, totais gerais e estilos, Segmentação de Dados (seleção múltipla, limpar filtro, itens sem dados apagados), Linha do Tempo por meses e Gráfico Dinâmico. Guias contextuais **Análise da Tabela Dinâmica** e **Design**.
- **Fórmulas modernas (Módulo 9):** matrizes dinâmicas que "derramam" (borda azul, fórmula da origem em cinza na barra, erro `#DESPEJAR!` com borda tracejada, referência `F2#`), FILTRO, CLASSIFICAR, CLASSIFICARPOR, ÚNICO, SEQUÊNCIA, TRANSPOR, critérios em lista (`=SOMASES(D:D;B:B;F2#)`), contas com intervalos (`=C2:C9*D2:D9`), LET, LAMBDA (inclusive `=LAMBDA(x;x*2)(5)`) e nomes definidos (Fórmulas › Definir Nome e Gerenciador de Nomes) para criar funções próprias como `=COMISSAO(B2)`.
- **Análise e cenários (Módulo 10):** Dados › Teste de Hipóteses com **Atingir Meta** (com janela de status e Cancelar), **Tabela de Dados** viva de uma ou duas variáveis (`=TABELA(;$B$1)`) e **Gerenciador de Cenários** (adicionar, editar, mostrar, excluir e Resumo do Cenário numa planilha nova); **Solver** em Dados › Análise (objetivo Máx./Mín./Valor de, células variáveis, restrições <=, =, >= e inteiro, não negativas; resolve modelos lineares com Simplex e inteiros com branch and bound, e não lineares com busca Nelder-Mead; manter ou restaurar valores); previsão com PREVISÃO.LINEAR, TENDÊNCIA, INCLINAÇÃO, INTERCEPÇÃO, CORREL e RQUAD; e **Linha de Tendência** nos gráficos.
- **Dashboards (Módulo 11):** Minigráficos (Inserir › Minigráficos: Linha, Coluna e Ganho/Perda) desenhados dentro das células, com guia contextual **Minigráfico** (tipo, Ponto Alto, Ponto Baixo, Marcadores, Limpar); **Conexões de Relatório** para uma segmentação filtrar várias Tabelas Dinâmicas; segmentações posicionadas sem cobrir os gráficos; e o modo sem linhas de grade lembrado entre demonstração e "Agora é sua vez".
- **Motor de fórmulas:** operadores, precedência, `;` como separador, vírgula decimal, erros do Excel (`#DIV/0!`, `#NOME?`...), ~85 funções (SOMA, SUBTOTAL, TEXTO, DIATRABALHO, DATADIF, TEXTOANTES, SE, PROCV, PROCX, ÍNDICE/CORRESP, SOMASES, CONT.SES, datas, texto, financeiras...).
- **Guia de lições:** destaque do que clicar, verificação automática, dica, progresso salvo no navegador.
- **"👀 Mostre-me como":** um cursor animado executa a tarefa devagar na interface real, com um balão explicando cada clique, tecla e texto digitado. Depois, **"↺ Agora é sua vez"** desfaz a demonstração para a pessoa repetir sozinha (ou **"▶ Ver de novo"**).
- **Trilha completa** (Básico → Intermediário → Avançado → Expert com Power Query, M, Power Pivot/DAX e automação). O **Nível Básico está completo** (Módulos 1 e 2, 13 lições) e o **Nível Intermediário** já tem os Módulos 3 (Lógica e condições, 4 lições), 4 (Contar e somar com critérios, 3 lições) 5 (Buscando dados: PROCV, PROCX, ÍNDICE+CORRESP, 3 lições), 6 (Texto e datas, 4 lições) e 7 (Tabelas e visualização, 4 lições) — o **Nível Intermediário está completo**. O **Nível Avançado está completo**: Módulos 8 (Tabelas Dinâmicas), 9 (Fórmulas modernas), 10 (Análise e cenários) e 11 (Dashboards), com 15 lições. O restante aparece como "em breve".
- **Modo livre** com planilhas de exemplo.

## Estrutura

```
index.html                 carrega tudo (scripts clássicos, funciona via file://)
css/app.css                trilha, painel do guia, componentes
css/excel.css              visual do simulador
js/core/util.js            referências, DOM, números/datas pt-BR
js/core/formula.js         tokenização, parser, avaliação, F4, ajuste de referências
js/core/functions.js       biblioteca de funções (inclui FILTRO, ÚNICO, CLASSIFICAR, LET, LAMBDA)
js/core/sheet.js           modelo da planilha, cálculo, preenchimento, classificar, filtro, página
js/core/workbook.js        pasta de trabalho: várias planilhas, inserir/excluir linhas e colunas, desfazer
js/core/features.js        tabelas, formatação condicional, validação de dados e dados dos gráficos
js/core/whatif.js          cálculo com valores hipotéticos, Tabela de Dados, Atingir Meta e Solver
js/core/pivot.js           Tabela Dinâmica: fonte, filtros, agrupamentos, campos calculados e saída nas células
js/ui/ribbon.js            Faixa de Opções (cada botão tem data-cmd)
js/ui/grid.js              grade, seleção, edição, alça
js/ui/excel.js             ExcelApp: junta tudo e executa comandos (emite eventos)
js/ui/excel-data.js        comandos de abas, inserir/excluir, classificar, filtrar, congelar, página
js/ui/excel-viz.js         comandos de tabelas, gráficos, formatação condicional e validação
js/ui/charts.js            gráficos em SVG (colunas, barras, linhas, pizza) sobre a grade
js/ui/excel-pivot.js       comandos e janelas da Tabela Dinâmica, setinhas de filtro, Gráfico Dinâmico
js/ui/pivot-pane.js        Lista de Campos (arrastar e soltar), Segmentação de Dados e Linha do Tempo
js/ui/sparklines.js        Minigráficos dentro das células
js/ui/dialogs.js           janelas Classificar, menu do filtro, Remover Duplicatas, Imprimir Títulos
js/ui/dialogs-format.js    Formatar Células, Texto para Colunas, Preenchimento Relâmpago
js/ui/dialogs-viz.js       Criar Tabela, Título do Gráfico, regras condicionais, Validação de Dados
js/ui/dialogs-names.js     Definir Nome e Gerenciador de Nomes
js/ui/dialogs-whatif.js    Atingir Meta, Tabela de Dados, Gerenciador de Cenários e Solver
js/ui/find.js              Localizar e Substituir
js/ui/print.js             paginação e Visualização de Impressão
js/tutor/tutor.js          motor das lições guiadas
js/tutor/demo.js           demonstrações animadas ("Mostre-me como")
js/course/curriculum.js    trilha completa (níveis, módulos, lições)
js/course/basico/*.js      lições do Nível Básico (demos-basico.js = roteiros das demonstrações)
js/course/intermediario/   lições do Nível Intermediário (demos-intermediario.js; o Módulo 7 traz as demos no próprio arquivo)
js/course/avancado/        lições do Nível Avançado (com as demos no próprio arquivo)
js/app.js                  telas: trilha, lição, modo livre
```

## Como criar uma lição

1. Adicione a lição na trilha em `js/course/curriculum.js` (o `id` é a chave).
2. Registre o conteúdo com `XL.registerLesson({...})` num arquivo em `js/course/...` e inclua o `<script>` no `index.html`.

```js
XL.registerLesson({
  id: 'i3-1',
  title: 'Função SE',
  summary: ['O que a pessoa aprendeu (aparece ao concluir)'],
  setup: (x) => x.set({ A1: 'Nota', B1: 7 }),      // dados iniciais
  steps: [
    { title: 'Explicação', text: '<p>HTML livre…</p>', target: 'cell:A1:B1' },
    {
      title: 'Sua vez',
      task: 'Em <b>C1</b>, digite <code>=SE(B1>=6;"Aprovado";"Reprovado")</code>.',
      hint: 'Dica opcional',
      target: ['cell:C1', '.rb-btn[data-cmd="fx"]'],   // células e/ou seletores CSS
      check: (x) => x.uses('C1', 'SE') && x.val('C1') === 'Aprovado',
      solve: (x) => x.app.setCell('C1', '=SE(B1>=6;"Aprovado";"Reprovado")'),
      success: 'Mensagem ao acertar',
    },
  ],
});
```

Referências aceitam outra planilha: `x.val('Resumo!B2')`, e destaques também: `'cell:Resumo!B2'`.

Auxiliares do `x` (contexto): `val`, `num`, `text`, `input`, `isFormula`, `formula`, `uses`, `refs`, `fmt`, `all(intervalo, fn)`, `col(letra, r1, r2)`, `sheetNames`, `hasSheet`, `hiddenCount`, `active`, `sel`, `tab`, `happened(evento, fn)`, `near`, `isErr`, `empty`, `set`, `widths`, e `x.app` (o simulador: `setCell`, `select`, `exec(cmd, arg)`, `fill`...).

Eventos que o guia observa: `select`, `commit`, `command`, `tab`, `namebox`, `fill`, `undo`, `paste`, `format`, `autosum`, `insertfn`, `f4`, `sort`, `filter`, `insert`, `delete`, `freeze`, `scroll`, `sheetadd`, `sheetrename`, `sheetswitch`, `find`, `replace`, `dedupe`, `pagesetup`, `printpreview`, `table`, `tablerename`, `tabletotals`, `tableexpand`, `calccolumn`, `chart`, `chartchange`, `cf`, `dv`, `dvpick`, `dverror`, `pivot`, `pivotfield`, `pivotsettings`, `pivotgroup`, `pivotcalc`, `pivotsort`, `pivotfilter`, `pivotrefresh`, `pivotchart`, `slicer`, `slicerpick`, `timeline`, `timelinepick`, `name`, `namedelete`, `goalseek`, `datatable`, `scenario`, `scenarioshow`, `scenariosummary`, `solvermodel`, `solver`, `spark`, `sparkchange`, `slicerconn`...

## Próximos passos sugeridos

1. Nível Expert: simulador do **Editor do Power Query** (etapas aplicadas, mesclar, acrescentar, despivotar, linguagem M).

## Como criar a demonstração de um passo

Cada passo com tarefa deve ter um roteiro em `XL.registerDemos(idDaLição, { 'Título do passo': [...] })`. As ações usam a interface real (a tarefa é conferida pelos mesmos eventos de quando a pessoa faz sozinha):

```js
const { cell, drag, point, type, key, tab, btn, menu, item, el, rclick, dbl, set, typeIn, fill, scroll, look, sheet, when } = XL.DemoDSL;
XL.registerDemos('i3-1', {
  'Sua vez': [
    cell('C1'),                                   // clica numa célula
    type('=SE(', 'Digite <code>=SE(</code>'),      // texto opcional do balão
    point('B1'),                                  // clica numa célula durante a fórmula
    type('>=6;"Aprovado";"Reprovado")'),
    key('Enter'),
  ],
});
```

Outras ações: `btn(cmd, arg)` (troca de guia sozinha se precisar), `menu(id)` + `item(cmd, arg)`, `drag('A1','C3')`, `dnd(seletorOrigem, seletorDestino)` (arrastar e soltar, ex.: campo da Lista de Campos até uma área), `fill('D2','D6')`, `rclick(seletor)`, `set(seletor, valor)`, `typeIn(seletor, texto)`, `look(alvo, texto)`, `when(cond, ação)`. Para testar rápido: `XL.demoSpeed = 0.05`.
