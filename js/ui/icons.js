/* Ícones SVG simples no estilo da Faixa de Opções do Excel. */
(function () {
  'use strict';
  const XL = window.XL;

  const G = '#107c41'; // verde Excel
  const BL = '#2b70c9';
  const RD = '#d13438';
  const OR = '#e8912d';

  const sm = (inner) =>
    `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;
  const lg = (inner) =>
    `<svg viewBox="0 0 32 32" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;
  const book = (color, glyph) =>
    lg(`<path d="M7 5h16a2 2 0 0 1 2 2v20H9a2 2 0 0 1-2-2z" fill="${color}" stroke="${color}"/><path d="M7 25a2 2 0 0 1 2-2h16" stroke="#fff"/><text x="16" y="18" fill="#fff" stroke="none" font-size="10" font-weight="700" text-anchor="middle" font-family="Segoe UI,Arial">${glyph}</text>`);

  XL.icons = {
    // Área de transferência
    paste: lg(`<rect x="6" y="6" width="17" height="22" rx="2" fill="#f3e3c3" stroke="#a5772c"/><rect x="10" y="3.5" width="9" height="5" rx="1.5" fill="#fff" stroke="#6b6b6b"/><rect x="13" y="13" width="14" height="16" rx="1.5" fill="#fff" stroke="#6b6b6b"/><path d="M16 18h8M16 21h8M16 24h5" stroke="#9a9a9a"/>`),
    cut: sm(`<circle cx="6.5" cy="17.5" r="3"/><circle cx="17.5" cy="17.5" r="3"/><path d="M8.5 15.2 17 3M15.5 15.2 7 3"/>`),
    copy: sm(`<rect x="8.5" y="8" width="11.5" height="13" rx="1.5"/><path d="M15.5 8V5a1.5 1.5 0 0 0-1.5-1.5H5.5A1.5 1.5 0 0 0 4 5v10.5A1.5 1.5 0 0 0 5.5 17h3"/>`),
    painter: sm(`<rect x="3.5" y="3" width="14" height="6" rx="1.2" fill="${OR}" stroke="${OR}"/><path d="M17.5 6h2.5v5.5h-8.5v2.5"/><rect x="10" y="14" width="3" height="7" rx="1"/>`),
    // Fonte
    borders: sm(`<rect x="3.5" y="3.5" width="17" height="17" stroke-dasharray="1.5 2" stroke="#8a8a8a"/><path d="M12 3.5v17M3.5 12h17" stroke-dasharray="1.5 2" stroke="#8a8a8a"/><path d="M3 20.5h18" stroke-width="2.2"/>`),
    fill: sm(`<path d="M5 11.5 11.5 5l6.5 6.5-6 6a1.5 1.5 0 0 1-2.1 0L5 12.6z"/><path d="M9 7.5 6.5 5"/><path d="M19.5 14s2 2.4 2 3.5a2 2 0 0 1-4 0c0-1.1 2-3.5 2-3.5z" fill="currentColor"/>`),
    // Alinhamento
    vTop: sm(`<path d="M4 4h16M9 8h6M9 11.5h6"/>`),
    vMid: sm(`<path d="M9 9.5h6M4 12h16M9 14.5h6"/>`),
    vBottom: sm(`<path d="M9 12.5h6M9 16h6M4 20h16"/>`),
    orient: sm(`<path d="M5 19 18 6M13 6h5v5"/><path d="M4 9.5 7.5 4l3 5.5" stroke-width="1.2"/>`),
    wrap: sm(`<path d="M4 6h16M4 12h12.5a3 3 0 0 1 0 6H12M14 16l-2 2 2 2M4 18h4"/>`),
    alignLeft: sm(`<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>`),
    alignCenter: sm(`<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>`),
    alignRight: sm(`<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>`),
    indentDec: sm(`<path d="M11 6h9M11 10h9M11 14h9M4 18h16M7 8.5 4 11l3 2.5"/>`),
    indentInc: sm(`<path d="M11 6h9M11 10h9M11 14h9M4 18h16M4 8.5 7 11l-3 2.5"/>`),
    merge: sm(`<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M7 12h10M9.5 9.5 7 12l2.5 2.5M14.5 9.5 17 12l-2.5 2.5"/>`),
    // Número
    currency: sm(`<rect x="2.5" y="6" width="19" height="12" rx="1.5" stroke="${G}"/><circle cx="12" cy="12" r="2.8" stroke="${G}"/><path d="M5.5 9v6M18.5 9v6" stroke="${G}"/>`),
    // Estilos
    condFormat: sm(`<rect x="3.5" y="3.5" width="17" height="17" rx="1"/><path d="M3.5 9h17M3.5 14.5h17M10 3.5v17"/><rect x="11" y="10" width="9" height="4" fill="${RD}" stroke="none"/><rect x="11" y="15.5" width="6" height="4.2" fill="${BL}" stroke="none"/>`),
    table: sm(`<rect x="3.5" y="3.5" width="17" height="17" rx="1"/><rect x="3.5" y="3.5" width="17" height="5" fill="${BL}" stroke="${BL}"/><path d="M3.5 13.5h17M10 8.5v12M15 8.5v12"/>`),
    cellStyles: sm(`<rect x="3.5" y="3.5" width="17" height="17" rx="1"/><rect x="5" y="5" width="6" height="6" fill="#c6efce" stroke="none"/><rect x="13" y="5" width="6" height="6" fill="#ffc7ce" stroke="none"/><rect x="5" y="13" width="6" height="6" fill="#ffeb9c" stroke="none"/><rect x="13" y="13" width="6" height="6" fill="#ddebf7" stroke="none"/>`),
    // Células
    insertCells: sm(`<rect x="3.5" y="5" width="13" height="14" rx="1"/><path d="M3.5 12h13M10 5v14"/><path d="M19.5 8v7M16 11.5h7" stroke="${G}" stroke-width="2"/>`),
    deleteCells: sm(`<rect x="3.5" y="5" width="13" height="14" rx="1"/><path d="M3.5 12h13M10 5v14"/><path d="m17 9 5 5M22 9l-5 5" stroke="${RD}" stroke-width="2"/>`),
    formatCells: sm(`<rect x="3.5" y="5" width="13" height="14" rx="1"/><path d="M3.5 12h13M10 5v14"/><path d="M17 14.5 21 6M19 18.5l-2-4" stroke="${OR}" stroke-width="2"/>`),
    // Edição
    fillDown: sm(`<rect x="5" y="3.5" width="14" height="17" rx="1"/><path d="M12 7v9M8.5 12.5 12 16l3.5-3.5" stroke="${BL}"/>`),
    clear: sm(`<path d="M8 20h12M5.2 14.8 13.5 6.5l5 5-6.3 6.3a2 2 0 0 1-2.8 0z"/><path d="m5.2 14.8 3.5 3.5" stroke="${RD}"/>`),
    sortFilter: lg(`<text x="5" y="13" fill="${BL}" stroke="none" font-size="11" font-weight="700" font-family="Segoe UI,Arial">A</text><text x="5" y="27" fill="${BL}" stroke="none" font-size="11" font-weight="700" font-family="Segoe UI,Arial">Z</text><path d="M12 5v20M9.5 22.5 12 25l2.5-2.5"/><path d="M16 8h12l-4.8 6.5V22l-2.4 1.6v-9.1z"/>`),
    find: lg(`<circle cx="18" cy="12" r="7.5"/><path d="M12.6 17.4 5 25" stroke-width="2.4"/>`),
    fx: lg(`<text x="5" y="22" fill="currentColor" stroke="none" font-size="16" font-style="italic" font-family="Cambria,Georgia,serif">fx</text>`),
    sigmaBig: lg(`<path d="M24 7H9l8 9-8 9h15" stroke-width="2"/>`),
    // Inserir
    pivot: lg(`<rect x="4" y="5" width="24" height="22" rx="1.5"/><path d="M4 11h24M11 5v22"/><path d="M15 15h9M15 19h6M15 23h9" stroke="${BL}"/><path d="M5.5 8h4" stroke="${G}" stroke-width="2"/>`),
    tableBig: lg(`<rect x="4" y="5" width="24" height="22" rx="1.5"/><rect x="4" y="5" width="24" height="6" fill="${BL}" stroke="${BL}"/><path d="M4 16.5h24M4 22h24M12 11v16M20 11v16"/>`),
    pictures: lg(`<rect x="4" y="6" width="24" height="20" rx="2"/><circle cx="11" cy="12.5" r="2.5" fill="${OR}" stroke="none"/><path d="m5 24 8-8 5 5 3-3 6 6" stroke="${G}"/>`),
    chartRec: lg(`<rect x="4" y="5" width="24" height="22" rx="1.5"/><rect x="8" y="16" width="4" height="8" fill="${BL}" stroke="none"/><rect x="14" y="11" width="4" height="13" fill="${G}" stroke="none"/><rect x="20" y="14" width="4" height="10" fill="${OR}" stroke="none"/>`),
    chartCol: sm(`<path d="M3.5 20.5h17"/><rect x="5" y="11" width="3.5" height="9" fill="${BL}" stroke="none"/><rect x="10.3" y="6" width="3.5" height="14" fill="${BL}" stroke="none"/><rect x="15.6" y="13" width="3.5" height="7" fill="${BL}" stroke="none"/>`),
    chartLine: sm(`<path d="M3.5 20.5h17M3.5 4v16.5"/><path d="m5 16 4.5-5 4 3.5L20 7" stroke="${BL}" stroke-width="2"/>`),
    chartPie: sm(`<circle cx="12" cy="12" r="8.5" fill="${BL}" stroke="${BL}"/><path d="M12 12V3.5A8.5 8.5 0 0 1 20.5 12z" fill="${OR}" stroke="#fff"/>`),
    chartBar: sm(`<path d="M3.5 3.5v17"/><rect x="4" y="5" width="12" height="3.5" fill="${BL}" stroke="none"/><rect x="4" y="10.3" width="16" height="3.5" fill="${BL}" stroke="none"/><rect x="4" y="15.6" width="8" height="3.5" fill="${BL}" stroke="none"/>`),
    chartScatter: sm(`<path d="M3.5 20.5h17M3.5 4v16.5"/><circle cx="8" cy="15" r="1.4" fill="${BL}" stroke="none"/><circle cx="12" cy="10" r="1.4" fill="${BL}" stroke="none"/><circle cx="16" cy="12" r="1.4" fill="${BL}" stroke="none"/><circle cx="18" cy="6" r="1.4" fill="${BL}" stroke="none"/>`),
    sparkline: lg(`<rect x="4" y="8" width="24" height="16" rx="1.5"/><path d="m7 19 4-6 4 4 4-7 6 5" stroke="${BL}" stroke-width="2"/>`),
    slicer: lg(`<rect x="6" y="4" width="20" height="24" rx="2"/><rect x="9" y="8" width="14" height="4" rx="1" fill="${BL}" stroke="none"/><rect x="9" y="14" width="14" height="4" rx="1"/><rect x="9" y="20" width="14" height="4" rx="1"/>`),
    // Fórmulas
    finance: book('#2e7d32', 'R$'),
    logic: book('#6a4fb3', '?'),
    text: book('#b4572a', 'A'),
    dateTime: book('#1f6fb2', '31'),
    lookup: book('#0f7d7d', '⇄'),
    math: book('#a2398a', 'θ'),
    more: book('#5a5a5a', '…'),
    names: lg(`<path d="M5 9h14l7 7-7 7H5z" fill="#fff6d5" stroke="#a5772c"/><circle cx="18" cy="16" r="1.6"/>`),
    tracePrec: sm(`<circle cx="5" cy="6" r="2" fill="${BL}" stroke="${BL}"/><path d="M6.5 7.5 17 17M14 17h3v-3" stroke="${BL}"/>`),
    traceDep: sm(`<circle cx="19" cy="18" r="2" fill="${BL}" stroke="${BL}"/><path d="M5 6l12 10.5M14 16.5h3v-3" stroke="${BL}"/>`),
    showFormulas: sm(`<rect x="3.5" y="4" width="17" height="16" rx="1"/><text x="6" y="16" fill="currentColor" stroke="none" font-size="9" font-style="italic" font-family="Cambria,Georgia,serif">fx</text>`),
    errorCheck: sm(`<path d="M12 3.5 21 19.5H3z" stroke="${OR}"/><path d="M12 9.5v4.5M12 16.8v.2" stroke="${OR}" stroke-width="2"/>`),
    evalFormula: sm(`<rect x="3.5" y="4" width="17" height="16" rx="1"/><path d="M7 9h4M9 7v4M13 9h4M7 15h4M13 14l3 3M16 14l-3 3"/>`),
    calc: lg(`<rect x="7" y="4" width="18" height="24" rx="2"/><rect x="10" y="7" width="12" height="5" fill="#ddebf7"/><path d="M11 16h.1M16 16h.1M21 16h.1M11 20h.1M16 20h.1M21 20h.1M11 24h.1M16 24h.1M21 24h.1" stroke-width="2.5"/>`),
    // Dados
    getData: lg(`<ellipse cx="11" cy="8" rx="6" ry="2.5" fill="#fff"/><path d="M5 8v12c0 1.4 2.7 2.5 6 2.5M17 8v5"/><rect x="14" y="15" width="14" height="12" rx="1" fill="#fff"/><path d="M14 19h14M19 15v12" /><rect x="14" y="15" width="14" height="4" fill="${G}" stroke="${G}"/>`),
    csv: sm(`<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M8.5 13h7M8.5 16h7" stroke="${G}"/>`),
    web: sm(`<circle cx="12" cy="12" r="8.5"/><path d="M3.5 12h17M12 3.5c2.5 2.4 3.7 5.2 3.7 8.5s-1.2 6.1-3.7 8.5c-2.5-2.4-3.7-5.2-3.7-8.5S9.5 5.9 12 3.5z"/>`),
    fromTable: sm(`<rect x="3.5" y="4" width="17" height="16" rx="1"/><path d="M3.5 9h17M3.5 14.5h17M10 4v16" /><path d="M3.5 4h17v5h-17z" fill="${G}" stroke="${G}"/>`),
    refresh: lg(`<path d="M25 12a10 10 0 1 0 1 7" stroke="${G}" stroke-width="2"/><path d="M26 5v7h-7" stroke="${G}" stroke-width="2"/>`),
    sortAZ: sm(`<text x="2.5" y="10" fill="${BL}" stroke="none" font-size="8.5" font-weight="700" font-family="Segoe UI,Arial">A</text><text x="2.5" y="21" fill="${BL}" stroke="none" font-size="8.5" font-weight="700" font-family="Segoe UI,Arial">Z</text><path d="M16 4v16M13 17l3 3 3-3"/>`),
    sortZA: sm(`<text x="2.5" y="10" fill="${BL}" stroke="none" font-size="8.5" font-weight="700" font-family="Segoe UI,Arial">Z</text><text x="2.5" y="21" fill="${BL}" stroke="none" font-size="8.5" font-weight="700" font-family="Segoe UI,Arial">A</text><path d="M16 4v16M13 17l3 3 3-3"/>`),
    sortBig: lg(`<rect x="4" y="5" width="15" height="22" rx="1.5"/><path d="M4 12h15M4 19.5h15"/><path d="M25 6v20M22 23l3 3 3-3" stroke="${BL}"/>`),
    filter: lg(`<path d="M5 6h22l-8.5 10.5V25l-5 2.5v-11z"/>`),
    textToCols: lg(`<rect x="4" y="6" width="10" height="20" rx="1"/><rect x="18" y="6" width="10" height="20" rx="1"/><path d="M14 16h4M16.5 14l2 2-2 2" stroke="${BL}"/>`),
    removeDup: lg(`<rect x="4" y="5" width="18" height="20" rx="1"/><path d="M4 11.5h18M4 18h18"/><path d="m20 18 8 8M28 18l-8 8" stroke="${RD}" stroke-width="2"/>`),
    validation: lg(`<rect x="4" y="5" width="18" height="20" rx="1"/><path d="M4 11.5h18M4 18h18"/><path d="m19 22 3.5 3.5L29 18" stroke="${G}" stroke-width="2.2"/>`),
    whatIf: lg(`<rect x="4" y="5" width="20" height="22" rx="1.5"/><path d="M4 12h20M4 19.5h20M12 5v22"/><text x="21" y="29" fill="${BL}" stroke="none" font-size="12" font-weight="700" font-family="Segoe UI,Arial">?</text>`),
    forecast: lg(`<path d="M4 26h24M4 5v21"/><path d="m6 22 5-6 4 3 5-7" stroke="${BL}" stroke-width="2"/><path d="m20 12 6-4" stroke="${OR}" stroke-width="2" stroke-dasharray="2 2"/>`),
    // Layout da Página
    margins: lg(`<rect x="7" y="4" width="18" height="24" rx="1.5" fill="#fff"/><path d="M11 4v24M21 4v24M7 8.5h18M7 23.5h18" stroke="${BL}" stroke-dasharray="2 1.6"/>`),
    orientation: lg(`<rect x="4" y="9" width="13" height="18" rx="1.5" fill="#fff"/><rect x="12" y="4" width="17" height="12" rx="1.5" fill="#fff" stroke="${BL}"/><path d="M7 14h7M7 18h7M7 22h5" stroke="#9a9a9a"/>`),
    paperSize: lg(`<rect x="8" y="4" width="16" height="24" rx="1.5" fill="#fff"/><path d="M5 7v18M3.5 8.5 5 7l1.5 1.5M3.5 23.5 5 25l1.5-1.5" stroke="${BL}"/><text x="16" y="20" fill="${BL}" stroke="none" font-size="8" font-weight="700" text-anchor="middle" font-family="Segoe UI,Arial">A4</text>`),
    printArea: lg(`<rect x="4" y="5" width="24" height="22" rx="1.5"/><path d="M4 11h24M4 17h24M12 5v22M20 5v22" stroke="#c8c8c8"/><rect x="8" y="9" width="14" height="12" stroke="${BL}" stroke-width="2" stroke-dasharray="3 2"/>`),
    printTitles: lg(`<rect x="5" y="4" width="15" height="18" rx="1" fill="#fff"/><rect x="12" y="10" width="15" height="18" rx="1" fill="#fff"/><rect x="5" y="4" width="15" height="4" fill="${BL}" stroke="${BL}"/><rect x="12" y="10" width="15" height="4" fill="${BL}" stroke="${BL}"/>`),
    preview: lg(`<rect x="6" y="4" width="16" height="22" rx="1.5" fill="#fff"/><path d="M9.5 9h9M9.5 13h9M9.5 17h5" stroke="#9a9a9a"/><circle cx="21" cy="21" r="5" fill="#fff" stroke="${G}" stroke-width="2"/><path d="m24.5 24.5 3.5 3.5" stroke="${G}" stroke-width="2.4"/>`),
    refreshSm: sm(`<path d="M19 9a7.5 7.5 0 1 0 .8 5" stroke="${G}" stroke-width="1.8"/><path d="M20 4v5h-5" stroke="${G}" stroke-width="1.8"/>`),
    // Exibir
    freeze: lg(`<rect x="4" y="5" width="24" height="22" rx="1.5"/><path d="M4 11h24M11 5v22" stroke="${BL}" stroke-width="2"/>`),
    zoom: lg(`<circle cx="14" cy="14" r="8"/><path d="M20 20l7 7M10.5 14h7M14 10.5v7" stroke-width="2"/>`),
    // Automatizar / Ajuda
    script: lg(`<rect x="5" y="4" width="22" height="24" rx="2"/><path d="M10 12l-3 3 3 3M22 12l3 3-3 3M17.5 10l-3 10" stroke="${G}"/>`),
    macro: lg(`<circle cx="16" cy="16" r="10"/><path d="M13 11.5v9l7.5-4.5z" fill="${G}" stroke="${G}"/>`),
    help: lg(`<circle cx="16" cy="16" r="11"/><path d="M12.5 12.5a3.5 3.5 0 1 1 5 3.2c-1 .5-1.5 1.2-1.5 2.3v.7M16 22.5v.2" stroke-width="2"/>`),
    keyboard: lg(`<rect x="3" y="9" width="26" height="15" rx="2"/><path d="M7 13h.1M11 13h.1M15 13h.1M19 13h.1M23 13h.1M7 17h.1M11 17h.1M15 17h.1M19 17h.1M23 17h.1M10 21h12" stroke-width="2"/>`),
    lock: sm(`<rect x="5" y="10.5" width="14" height="10" rx="1.5"/><path d="M8 10.5V8a4 4 0 0 1 8 0v2.5"/>`),
    caret: `<svg viewBox="0 0 10 10" width="8" height="8"><path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>`,
  };
})();
