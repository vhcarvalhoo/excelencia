# Excelência — aprenda Excel praticando

Aplicativo que ensina Excel do básico ao Expert com um **simulador do Excel** e um **guia passo a passo** que confere cada tarefa automaticamente.

## Como abrir

Não precisa instalar nada: dê **duplo clique em `index.html`** (Chrome ou Edge).

Para desenvolver com recarregamento simples, também dá para servir a pasta:

```powershell
python -m http.server 8780
```

e abrir `http://localhost:8780`.

## O que já existe

- **Simulador do Excel (pt-BR):** Faixa de Opções com as guias, Caixa de Nome, Barra de Fórmulas, grade com seleção por mouse/teclado, edição na célula, modo *Aponte* (clicar em células enquanto digita a fórmula), referências coloridas, autocompletar de funções, dica de sintaxe, F4 para `$`, alça de preenchimento (cópia, séries numéricas, meses, dias), desfazer/refazer, copiar/colar (inclusive com o Excel real), Pincel de Formatação, AutoSoma, Inserir Função, Mostrar Fórmulas, barra de status com Média/Contagem/Soma, formatos (Moeda, Contábil, %, Data...), bordas, cores, estilos de célula.
- **Organização de dados (Módulo 2):** várias planilhas (abas, renomear, excluir, referências como `=Janeiro!B6` — inclusive clicando na aba enquanto digita a fórmula), inserir/excluir linhas e colunas com ajuste automático de fórmulas (e `#REF!`), classificar (rápido e com vários níveis), AutoFiltro (valores, filtros de número/texto, linhas azuis, contagem na barra de status, `SUBTOTAL`), remover duplicatas, congelar painéis, Localizar/Substituir, menu do botão direito, Formatar Células (Ctrl+1) com formatos personalizados (`0000`, `#.##0,0 "kg"`, `dd/mm/aaaa`...), Texto para Colunas, Preenchimento Relâmpago (Ctrl+E), Layout da Página (margens, orientação, tamanho, área de impressão, imprimir títulos, ajustar à página) e Visualização de Impressão com contagem de páginas e quebras tracejadas na grade.
- **Tabelas e visualização (Módulo 7):** Formatar como Tabela (Ctrl+T, estilos, linhas em tiras, Linha de Totais, renomear, Converter em Intervalo), referências estruturadas (`[@Qtd]`, `[@[Preço unit.]]`, `Vendas[Total]`), colunas calculadas e tabela que cresce ao digitar logo abaixo; gráficos em SVG (colunas, barras, linhas, pizza, Gráficos Recomendados) que se movem, redimensionam e acompanham os dados, com título, legenda e rótulos; formatação condicional (realçar regras, primeiros/últimos, barras de dados, escalas de cor, regra com fórmula); validação de dados (lista suspensa com seta e Alt+↓, número inteiro/decimal/data/tamanho do texto, mensagem de entrada e alerta de erro). As guias contextuais **Design da Tabela** e **Design do Gráfico** aparecem só quando a tabela/o gráfico está selecionado.
- **Motor de fórmulas:** operadores, precedência, `;` como separador, vírgula decimal, erros do Excel (`#DIV/0!`, `#NOME?`...), ~85 funções (SOMA, SUBTOTAL, TEXTO, DIATRABALHO, DATADIF, TEXTOANTES, SE, PROCV, PROCX, ÍNDICE/CORRESP, SOMASES, CONT.SES, datas, texto, financeiras...).
- **Guia de lições:** destaque do que clicar, verificação automática, dica, progresso salvo no navegador.
- **"👀 Mostre-me como":** um cursor animado executa a tarefa devagar na interface real, com um balão explicando cada clique, tecla e texto digitado. Depois, **"↺ Agora é sua vez"** desfaz a demonstração para a pessoa repetir sozinha (ou **"▶ Ver de novo"**).
- **Trilha completa** (Básico → Intermediário → Avançado → Expert com Power Query, M, Power Pivot/DAX e automação). O **Nível Básico está completo** (Módulos 1 e 2, 13 lições) e o **Nível Intermediário** já tem os Módulos 3 (Lógica e condições, 4 lições), 4 (Contar e somar com critérios, 3 lições) 5 (Buscando dados: PROCV, PROCX, ÍNDICE+CORRESP, 3 lições), 6 (Texto e datas, 4 lições) e 7 (Tabelas e visualização, 4 lições) — o **Nível Intermediário está completo**. O restante aparece como "em breve".
- **Modo livre** com planilhas de exemplo.

## Estrutura

```
index.html                 carrega tudo (scripts clássicos, funciona via file://)
css/app.css                trilha, painel do guia, componentes
css/excel.css              visual do simulador
js/core/util.js            referências, DOM, números/datas pt-BR
js/core/formula.js         tokenização, parser, avaliação, F4, ajuste de referências
js/core/functions.js       biblioteca de funções
js/core/sheet.js           modelo da planilha, cálculo, preenchimento, classificar, filtro, página
js/core/workbook.js        pasta de trabalho: várias planilhas, inserir/excluir linhas e colunas, desfazer
js/core/features.js        tabelas, formatação condicional, validação de dados e dados dos gráficos
js/ui/ribbon.js            Faixa de Opções (cada botão tem data-cmd)
js/ui/grid.js              grade, seleção, edição, alça
js/ui/excel.js             ExcelApp: junta tudo e executa comandos (emite eventos)
js/ui/excel-data.js        comandos de abas, inserir/excluir, classificar, filtrar, congelar, página
js/ui/excel-viz.js         comandos de tabelas, gráficos, formatação condicional e validação
js/ui/charts.js            gráficos em SVG (colunas, barras, linhas, pizza) sobre a grade
js/ui/dialogs.js           janelas Classificar, menu do filtro, Remover Duplicatas, Imprimir Títulos
js/ui/dialogs-format.js    Formatar Células, Texto para Colunas, Preenchimento Relâmpago
js/ui/dialogs-viz.js       Criar Tabela, Título do Gráfico, regras condicionais, Validação de Dados
js/ui/find.js              Localizar e Substituir
js/ui/print.js             paginação e Visualização de Impressão
js/tutor/tutor.js          motor das lições guiadas
js/tutor/demo.js           demonstrações animadas ("Mostre-me como")
js/course/curriculum.js    trilha completa (níveis, módulos, lições)
js/course/basico/*.js      lições do Nível Básico (demos-basico.js = roteiros das demonstrações)
js/course/intermediario/   lições do Nível Intermediário (demos-intermediario.js; o Módulo 7 traz as demos no próprio arquivo)
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

Eventos que o guia observa: `select`, `commit`, `command`, `tab`, `namebox`, `fill`, `undo`, `paste`, `format`, `autosum`, `insertfn`, `f4`, `sort`, `filter`, `insert`, `delete`, `freeze`, `scroll`, `sheetadd`, `sheetrename`, `sheetswitch`, `find`, `replace`, `dedupe`, `pagesetup`, `printpreview`, `table`, `tablerename`, `tabletotals`, `tableexpand`, `calccolumn`, `chart`, `chartchange`, `cf`, `dv`, `dvpick`, `dverror`...

## Próximos passos sugeridos

1. Nível Avançado: Módulo 8 (Tabelas Dinâmicas) — o simulador ainda não tem Tabela Dinâmica nem Segmentação de Dados.
2. Nível Avançado: Módulo 9 (matrizes dinâmicas: FILTRO, CLASSIFICAR, ÚNICO, LET, LAMBDA) — o motor precisa "derramar" resultados.
3. Nível Expert: simulador do **Editor do Power Query** (etapas aplicadas, mesclar, acrescentar, despivotar, linguagem M).

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

Outras ações: `btn(cmd, arg)` (troca de guia sozinha se precisar), `menu(id)` + `item(cmd, arg)`, `drag('A1','C3')`, `fill('D2','D6')`, `rclick(seletor)`, `set(seletor, valor)`, `typeIn(seletor, texto)`, `look(alvo, texto)`, `when(cond, ação)`. Para testar rápido: `XL.demoSpeed = 0.05`.
