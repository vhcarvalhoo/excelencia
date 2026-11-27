/* Aplicativo: trilha do curso (início), tela de lição com guia e modo livre. Navegação por #hash. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;

  const state = { excel: null, tutor: null };
  const progress = XL.store.get('progress', { done: {}, last: null });
  const saveProgress = () => XL.store.set('progress', progress);

  const lessonNum = (module, lesson) => `${module.id.replace(/\D/g, '')}.${module.lessons.indexOf(lesson) + 1}`;
  const available = (id) => !!XL.Lessons[id];

  function cleanup() {
    if (state.tutor) { state.tutor.destroy(); state.tutor = null; }
    if (state.excel) { state.excel.destroy(); state.excel = null; }
    XL.ui.closePopup();
    XL.ui.closeAllModals();
  }

  const logo = () => h('a.brand', { href: '#/' },
    h('span.brand-mark', { html: '<svg viewBox="0 0 32 32" width="30" height="30"><rect x="2" y="2" width="28" height="28" rx="7" fill="#107c41"/><path d="M9 10h14M9 16h14M9 22h14M16 10v12" stroke="#fff" stroke-width="2" opacity=".55"/><path d="M10 9l12 14M22 9 10 23" stroke="#fff" stroke-width="3" stroke-linecap="round"/></svg>' }),
    h('span.brand-name', 'Excel', h('b', 'ência')));

  // ======================== Início / Trilha ========================
  function showHome() {
    cleanup();
    document.title = 'Excelência — Aprenda Excel praticando';
    const root = document.getElementById('root');
    root.innerHTML = '';
    root.className = 'screen-home';
    const flat = XL.Course.flat();
    const ready = flat.filter((x) => available(x.lesson.id));
    const doneCount = ready.filter((x) => progress.done[x.lesson.id]).length;
    const nextItem = ready.find((x) => !progress.done[x.lesson.id]) || ready[0];
    const started = doneCount > 0;

    const header = h('header.home-top', logo(),
      h('nav.home-nav',
        h('a', { href: '#/livre' }, 'Modo livre'),
        h('span.pill', `${doneCount}/${ready.length} lições concluídas`)));

    const hero = h('section.hero',
      h('div.hero-text',
        h('div.hero-kicker', 'Curso prático e interativo'),
        h('h1', 'Aprenda Excel ', h('span.hl', 'praticando'), ', do zero ao Expert.'),
        h('p', 'Um simulador do Excel com um guia ao seu lado: você faz cada passo de verdade e o guia confere na hora. Do primeiro clique até Power Query, Power Pivot e DAX.'),
        h('div.hero-cta',
          h('a.btn.btn-primary.btn-lg', { href: '#/licao/' + nextItem.lesson.id }, started ? 'Continuar: ' + nextItem.lesson.title + ' ›' : 'Começar agora ›'),
          h('a.btn.btn-ghost.btn-lg', { href: '#/livre' }, 'Praticar no modo livre')),
        h('div.hero-feats',
          h('span', '✓ Simulador fiel ao Excel em português'),
          h('span', '✓ Verificação automática de cada passo'),
          h('span', '✓ Dicas e "Fazer por mim" quando travar'))),
      heroVisual());

    const levels = h('section.levels', h('h2.section-title', 'Sua trilha'),
      XL.Course.levels.map((lv) => levelCard(lv)));

    const foot = h('footer.home-foot', 'Excelência · simulador educacional. Excel é marca registrada da Microsoft; este app não é afiliado à Microsoft.');
    root.append(header, h('main.home-main', hero, levels), foot);
  }

  function heroVisual() {
    const rows = [['', 'A', 'B'], ['1', 'Mês', 'Vendas'], ['2', 'Jan', '12.500'], ['3', 'Fev', '15.800'], ['4', 'Mar', '13.200'], ['5', 'Total', '41.500']];
    const tbl = h('table.hero-sheet', rows.map((r, i) => h('tr', r.map((c, j) => (i === 0 || j === 0 ? h('th', c) : h('td' + (i === 5 && j === 2 ? '.act' : '') + (i === 1 || i === 5 ? '.b' : ''), c))))));
    return h('div.hero-visual',
      h('div.hero-card',
        h('div.hero-fbar', h('span.nb', 'B5'), h('span.fx', 'fx'), h('span.f', '=SOMA(B2:B4)')),
        tbl,
        h('div.hero-bubble', h('b', 'Guia: '), 'Agora clique em ', h('code', 'Σ AutoSoma'), ' e confirme com Enter. 🎯')));
  }

  function levelCard(lv) {
    const lessons = [].concat(...lv.modules.map((m) => m.lessons));
    const ready = lessons.filter((l) => available(l.id));
    const done = ready.filter((l) => progress.done[l.id]).length;
    const pct = ready.length ? Math.round((done / ready.length) * 100) : 0;
    const head = h('div.level-head',
      h('div.level-tag', { style: { background: lv.color } }, lv.tag),
      h('div.level-info', h('h3', lv.name), h('p', lv.desc)),
      ready.length
        ? h('div.level-prog', h('div.level-bar', h('div', { style: { width: pct + '%', background: lv.color } })), h('span', `${done}/${ready.length}`))
        : h('div.level-soon', 'Em produção'));
    const extra = lv.id === 'expert'
      ? h('div.level-chips', ['Power Query', 'Linguagem M', 'Power Pivot', 'DAX', 'VBA', 'Office Scripts', 'Python no Excel'].map((c) => h('span.chip', c)))
      : null;
    const mods = h('div.modules', lv.modules.map((m) => moduleCard(lv, m)));
    return h('div.level', { style: { '--lv': lv.color } }, head, extra, mods);
  }

  function moduleCard(lv, m) {
    const list = h('ol.lesson-list', m.lessons.map((l) => {
      const ok = available(l.id);
      const done = !!progress.done[l.id];
      const num = lessonNum(m, l);
      const inner = [
        h('span.ls-num', done ? '✓' : num),
        h('span.ls-body', h('span.ls-title', l.title), l.desc ? h('span.ls-desc', l.desc) : null),
        ok ? h('span.ls-meta', l.minutes ? l.minutes + ' min' : '') : h('span.ls-lock', { title: 'Em breve', html: XL.icons.lock }),
      ];
      return h('li', ok
        ? h('a.ls' + (done ? '.done' : ''), { href: '#/licao/' + l.id }, inner)
        : h('div.ls.soon', inner));
    }));
    return h('div.module', h('div.module-head', h('span.module-num', 'Módulo ' + m.id.replace(/\D/g, '')), h('h4', m.title), m.desc ? h('p', m.desc) : null), list);
  }

  // ======================== Lição ========================
  function showLesson(id) {
    const lesson = XL.Lessons[id];
    const info = XL.Course.find(id);
    if (!lesson || !info) { location.hash = '#/'; return; }
    cleanup();
    const num = lessonNum(info.module, info.lesson);
    document.title = `${num} ${lesson.title} — Excelência`;
    lesson.crumb = `Nível ${info.level.name} · Módulo ${info.module.id.replace(/\D/g, '')}: ${info.module.title}`;
    progress.last = id;
    saveProgress();

    const root = document.getElementById('root');
    root.innerHTML = '';
    root.className = 'screen-lesson';
    const host = h('div.excel-host');
    const top = h('header.lesson-top',
      h('a.btn.btn-ghost.btn-sm', { href: '#/' }, '‹ Trilha'),
      logo(),
      h('div.lesson-name', h('span.lesson-num', { style: { background: info.level.color } }, num), lesson.title),
      h('span.top-spacer'),
      h('button.btn.btn-ghost.btn-sm', { type: 'button', title: 'Recomeçar a lição do início', onclick: () => { if (confirm('Recomeçar esta lição do início?')) state.tutor.start(); } }, '↺ Recomeçar'));
    const main = h('main.lesson-main', host);
    root.append(top, main);

    state.excel = new XL.ExcelApp(host);
    state.tutor = new XL.Tutor(state.excel, lesson, { onComplete: () => completeLesson(id) });
    main.appendChild(state.tutor.el);
    state.tutor.start();
  }

  function completeLesson(id) {
    progress.done[id] = true;
    saveProgress();
    const lesson = XL.Lessons[id];
    const flat = XL.Course.flat();
    const i = flat.findIndex((x) => x.lesson.id === id);
    const next = flat.slice(i + 1).find((x) => available(x.lesson.id));
    const body = h('div.done-body',
      h('div.done-trophy', '🏆'),
      h('p.done-lead', 'Lição concluída! Você aprendeu:'),
      h('ul.done-list', (lesson.summary || []).map((s) => h('li', { html: s }))),
      next ? null : h('p.done-next', 'Você terminou todas as lições disponíveis por enquanto. Novas lições estão a caminho — até o nível Expert com Power Query!'));
    const buttons = [{ label: 'Voltar à trilha', onClick: () => { location.hash = '#/'; } }];
    if (next) buttons.push({ label: `Próxima: ${next.lesson.title} ›`, primary: true, onClick: () => { location.hash = '#/licao/' + next.lesson.id; } });
    else buttons.push({ label: 'Praticar no modo livre', primary: true, onClick: () => { location.hash = '#/livre'; } });
    buttons.unshift({ label: 'Continuar mexendo aqui', onClick: () => {} });
    XL.ui.modal({ title: lesson.title, body, buttons, width: 520, className: 'done-modal' });
  }

  // ======================== Modo livre ========================
  const SAMPLES = {
    vendas: {
      label: 'Vendas por região',
      cells: {
        A1: { v: 'Região', fmt: { b: true } }, B1: { v: 'Vendedor', fmt: { b: true } }, C1: { v: 'Produto', fmt: { b: true } }, D1: { v: 'Qtd', fmt: { b: true } }, E1: { v: 'Preço', fmt: { b: true } }, F1: { v: 'Total', fmt: { b: true } },
        A2: 'Sul', B2: 'Ana', C2: 'Notebook', D2: 3, E2: { v: 3500, fmt: { numFmt: 'moeda' } },
        A3: 'Sudeste', B3: 'Bruno', C3: 'Monitor', D3: 8, E3: { v: 899, fmt: { numFmt: 'moeda' } },
        A4: 'Nordeste', B4: 'Carla', C4: 'Teclado', D4: 15, E4: { v: 120, fmt: { numFmt: 'moeda' } },
        A5: 'Sul', B5: 'Diego', C5: 'Mouse', D5: 22, E5: { v: 45, fmt: { numFmt: 'moeda' } },
        A6: 'Norte', B6: 'Elisa', C6: 'Notebook', D6: 2, E6: { v: 3500, fmt: { numFmt: 'moeda' } },
        A7: 'Sudeste', B7: 'Fábio', C7: 'Headset', D7: 10, E7: { v: 210, fmt: { numFmt: 'moeda' } },
        A8: 'Centro-Oeste', B8: 'Gabi', C8: 'Webcam', D8: 6, E8: { v: 350, fmt: { numFmt: 'moeda' } },
      },
      widths: { A: 100, C: 90 },
    },
    gastos: {
      label: 'Orçamento pessoal',
      cells: {
        A1: { v: 'Orçamento do mês', fmt: { b: true, size: 14 } },
        A3: { v: 'Categoria', fmt: { b: true } }, B3: { v: 'Previsto', fmt: { b: true } }, C3: { v: 'Realizado', fmt: { b: true } }, D3: { v: 'Diferença', fmt: { b: true } },
        A4: 'Moradia', B4: 1800, C4: 1800, A5: 'Alimentação', B5: 1200, C5: 1340, A6: 'Transporte', B6: 450, C6: 380,
        A7: 'Saúde', B7: 300, C7: 410, A8: 'Lazer', B8: 400, C8: 250, A9: 'Educação', B9: 350, C9: 350,
      },
      widths: { A: 110 },
    },
  };

  function showSandbox() {
    cleanup();
    document.title = 'Modo livre — Excelência';
    const root = document.getElementById('root');
    root.innerHTML = '';
    root.className = 'screen-lesson';
    const host = h('div.excel-host');
    const top = h('header.lesson-top',
      h('a.btn.btn-ghost.btn-sm', { href: '#/' }, '‹ Trilha'),
      logo(),
      h('div.lesson-name', h('span.lesson-num', { style: { background: '#475569' } }, '∞'), 'Modo livre'),
      h('span.top-spacer'));
    const main = h('main.lesson-main', host);
    root.append(top, main);
    const app = (state.excel = new XL.ExcelApp(host));

    const load = (key) => {
      app.reset();
      const s = SAMPLES[key];
      app.setCells(s.cells);
      app.setColWidths(s.widths || {});
      XL.ui.toast(`Exemplo “${s.label}” carregado.`);
    };
    const panel = h('aside.tutor.sandbox',
      h('div.tutor-head', h('div.tutor-crumb', 'Sem regras, sem verificação'), h('div.tutor-title', 'Modo livre')),
      h('div.tutor-body',
        h('p', 'Experimente à vontade: tudo que você aprendeu nas lições funciona aqui.'),
        h('h3.tutor-step-title', 'Carregar um exemplo'),
        h('div.sample-list', Object.keys(SAMPLES).map((k) => h('button.btn.btn-soft.btn-block', { type: 'button', onclick: () => load(k) }, '📄 ' + SAMPLES[k].label))),
        h('button.btn.btn-ghost.btn-block', { type: 'button', onclick: () => { app.reset(); XL.ui.toast('Planilha limpa.'); } }, '🧹 Planilha em branco'),
        h('h3.tutor-step-title', 'Ideias para praticar'),
        h('ul.ideas',
          h('li', { html: 'No exemplo de vendas, calcule <code>=D2*E2</code> em F2 e arraste até F8.' }),
          h('li', { html: 'Use <code>=SOMASE(A2:A8;"Sul";F2:F8)</code> para somar só a região Sul.' }),
          h('li', { html: 'No orçamento, calcule a diferença e destaque os estouros com <b>Estilos de Célula › Ruim</b>.' }),
          h('li', { html: 'Teste <code>=PROCX("Carla";B2:B8;C2:C8)</code> para descobrir o que a Carla vendeu.' }),
          h('li', { html: 'Ative <b>Fórmulas › Mostrar Fórmulas</b> para ver todas as fórmulas de uma vez.' })),
        h('button.btn.btn-ghost.btn-block', { type: 'button', onclick: () => app.exec('shortcuts') }, '⌨️ Ver atalhos de teclado')));
    main.appendChild(panel);
    app.grid.focus();
  }

  // ======================== Rotas ========================
  function route() {
    const hash = location.hash || '#/';
    let m;
    if ((m = /^#\/licao\/([\w-]+)/.exec(hash))) showLesson(m[1]);
    else if (hash.startsWith('#/livre')) showSandbox();
    else showHome();
    window.scrollTo(0, 0);
  }
  window.addEventListener('hashchange', route);
  window.addEventListener('DOMContentLoaded', route);
  XL.App = { route, progress, state };
})();
