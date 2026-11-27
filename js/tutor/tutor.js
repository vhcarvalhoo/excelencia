/* Guia de lições: mostra cada passo, destaca onde clicar, verifica a tarefa automaticamente,
   oferece dica e "Fazer por mim". */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const PRAISE = ['Mandou bem!', 'Perfeito!', 'Isso aí!', 'Excelente!', 'Muito bom!', 'Acertou em cheio!'];
  // eventos que podem completar uma tarefa
  const CHECK_EVENTS = new Set(['select', 'commit', 'command', 'tab', 'namebox', 'fill', 'undo', 'redo', 'paste', 'copy', 'cut', 'format', 'clear', 'autosum', 'insertfn', 'colwidth', 'change', 'f4',
    'sort', 'filter', 'filtermenu', 'insert', 'delete', 'freeze', 'scroll', 'sheetadd', 'sheetrename', 'sheetswitch', 'sheetdelete',
    'find', 'findall', 'findopen', 'replace', 'dedupe', 'pagesetup', 'printpreview', 'printclose', 'print',
    'texttocols', 'flashfill', 'formatcells',
    'table', 'tablestyle', 'tablebanded', 'tabletotals', 'tablerename', 'tableconvert', 'tableexpand', 'calccolumn',
    'chart', 'chartselect', 'chartchange', 'chartdelete', 'chartmove', 'chartresize', 'cf', 'cfclear', 'dv', 'dvclear', 'dvpick', 'dvopen', 'dverror']);

  class Tutor {
    constructor(app, lesson, opts = {}) {
      this.app = app;
      this.lesson = lesson;
      this.opts = opts;
      this.idx = 0;
      this.done = false;
      this.rings = [];
      this.ctx = this.makeCtx();
      this.el = this.buildPanel();
      this.unsub = app.on('*', (evt, data) => this.onEvent(evt, data));
      this.raf = requestAnimationFrame(() => this.loop());
    }

    destroy() {
      this.unsub();
      cancelAnimationFrame(this.raf);
      this.rings.forEach((r) => r.remove());
      this.rings = [];
      this.app.grid.setGuides([]);
      this.el.remove();
    }

    start() {
      this.settingUp = true;
      this.app.reset();
      if (this.lesson.setup) this.lesson.setup(this.ctx);
      this.app.undoStack = [];
      this.settingUp = false;
      this.goto(0);
    }

    // ---------- Contexto usado nas verificações ----------
    makeCtx() {
      const app = this.app;
      // ref pode ser "B6" (planilha ativa) ou "Janeiro!B6"
      const loc = (ref) => {
        const i = ref.lastIndexOf('!');
        const S = i > 0 ? app.wb.byName(ref.slice(0, i).replace(/^'|'$/g, '')) : app.sheet;
        const p = XL.parseRef(i > 0 ? ref.slice(i + 1) : ref);
        return { S: S || app.sheet, r: p.r, c: p.c };
      };
      const ctx = {
        app,
        events: [],
        get sheet() { return app.sheet; },
        val(ref) { const p = loc(ref); return p.S.value(p.r, p.c); },
        num(ref) { const v = ctx.val(ref); return typeof v === 'number' ? v : NaN; },
        input(ref) { const p = loc(ref); const c = p.S.cell(p.r, p.c); return c ? c.input : ''; },
        text(ref) { return XL.normName(String(ctx.input(ref)).trim()); },
        shown(ref) { const p = loc(ref); return p.S.textOf(p.r, p.c); },
        empty(ref) { const p = loc(ref); return !p.S.has(p.r, p.c); },
        isFormula(ref) { const p = loc(ref); const c = p.S.cell(p.r, p.c); return !!c && c.kind === 'formula'; },
        formula(ref) { return ctx.isFormula(ref) ? ctx.input(ref).toUpperCase().replace(/\s+/g, '') : ''; },
        uses(ref, fn) { return XL.Formula.usesFunction(ctx.input(ref), fn); },
        refs(ref) { return XL.Formula.refsIn(ctx.input(ref)).map((x) => x.text); },
        fmt(ref) { const p = loc(ref); return p.S.fmt(p.r, p.c); },
        sheetNames() { return app.wb.sheets.map((s) => s.name); },
        hasSheet(name) { return !!app.wb.byName(name); },
        col(letter, r1, r2) { const out = []; for (let r = r1; r <= r2; r++) out.push(ctx.val(letter + r)); return out; },
        hiddenCount() { return app.sheet.hiddenRows.size; },
        lastEvent(evt) { for (let i = ctx.events.length - 1; i >= 0; i--) if (ctx.events[i].evt === evt) return ctx.events[i].data || {}; return null; },
        all(rangeStr, pred) {
          const g = XL.parseRange(rangeStr);
          let ok = true;
          XL.eachInRange(g, (r, c) => { if (!pred(XL.toRef(r, c))) ok = false; });
          return ok;
        },
        active() { return XL.toRef(app.grid.sel.ar, app.grid.sel.ac); },
        sel() { return XL.rangeToStr(app.grid.range()); },
        tab() { return app.ribbon.active; },
        happened(evt, pred) { return ctx.events.some((e) => e.evt === evt && (!pred || pred(e.data || {}))); },
        near(a, b, eps = 1e-6) { return typeof a === 'number' && Math.abs(a - b) < eps; },
        isErr(ref, code) { const v = ctx.val(ref); return !!v && v.isXLError && (!code || v.code === code); },
        editing() { return !!app.grid.editing; },
        set(map) { app.setCells(map); },
        widths(map) { app.setColWidths(map); },
      };
      return ctx;
    }

    // ---------- Painel ----------
    buildPanel() {
      const L = this.lesson;
      this.titleEl = h('h3.tutor-step-title');
      this.textEl = h('div.tutor-text');
      this.taskEl = h('div.tutor-task');
      this.hintEl = h('div.tutor-hint', { style: { display: 'none' } });
      this.successEl = h('div.tutor-success', { style: { display: 'none' } });
      this.progressBar = h('div.tutor-bar-fill');
      this.progressTxt = h('span.tutor-progress-txt');
      this.btnBack = h('button.btn.btn-ghost', { type: 'button', onclick: () => this.goto(this.idx - 1), title: 'Passo anterior' }, '‹ Voltar');
      this.btnHint = h('button.btn.btn-soft', { type: 'button', onclick: () => this.showHint() }, '💡 Dica');
      this.btnSolve = h('button.btn.btn-soft.btn-show', { type: 'button', onclick: () => this.showMe(), title: 'O guia faz a tarefa devagar, mostrando cada clique e tecla. Depois você pode refazer sozinho.' }, '👀 Mostre-me como');
      this.btnNext = h('button.btn.btn-primary', { type: 'button', onclick: () => this.next() }, 'Continuar ›');
      this.body = h('div.tutor-body', this.titleEl, this.textEl, this.taskEl, this.hintEl, this.successEl);
      return h('aside.tutor',
        h('div.tutor-head',
          h('div.tutor-crumb', L.crumb || ''),
          h('div.tutor-title', L.title),
          h('div.tutor-progress', h('div.tutor-bar', this.progressBar), this.progressTxt)),
        this.body,
        h('div.tutor-foot', this.btnBack, h('span.tutor-spacer'), this.btnHint, this.btnSolve, this.btnNext));
    }

    goto(i) {
      const steps = this.lesson.steps;
      if (i < 0 || i >= steps.length) return;
      this.idx = i;
      const st = steps[i];
      this.done = !st.check; // passos só de leitura já começam "feitos"
      this.demoed = false;
      this.demoSnap = null;
      if (this.app.grid.editing) this.app.grid.cancel();
      if (st.setup) {
        this.settingUp = true;
        st.setup(this.ctx);
        this.settingUp = false;
      }
      this.ctx.events = [];
      this.render();
      this.app.grid.focus();
    }

    render() {
      const steps = this.lesson.steps;
      const st = steps[this.idx];
      const isTask = !!st.check;
      this.progressBar.style.width = Math.round(((this.idx + (this.done && isTask ? 1 : 0)) / steps.length) * 100) + '%';
      this.progressTxt.textContent = `Passo ${this.idx + 1} de ${steps.length}`;
      this.titleEl.textContent = st.title || '';
      this.textEl.innerHTML = st.text || '';
      this.textEl.style.display = st.text ? '' : 'none';
      if (isTask) {
        this.taskEl.style.display = '';
        this.taskEl.className = 'tutor-task' + (this.done ? ' done' : '');
        this.taskEl.innerHTML = `<div class="tutor-task-head">${this.done ? '✅ Tarefa concluída' : '🎯 Sua tarefa'}</div><div class="tutor-task-body">${st.task || ''}</div>`;
      } else {
        this.taskEl.style.display = 'none';
      }
      this.hintEl.style.display = 'none';
      this.hintEl.innerHTML = st.hint ? '💡 ' + st.hint : '';
      if (this.done && isTask) {
        this.successEl.style.display = '';
        const head = this.demoed ? '👀 Foi assim que o guia fez' : `🎉 ${this.praise || PRAISE[0]}`;
        this.successEl.innerHTML = `<div class="tutor-success-head">${head}</div>${st.success ? `<div>${st.success}</div>` : ''}`;
        if (this.demoed && this.demoSnap) {
          this.successEl.appendChild(h('div.tutor-yourturn',
            h('div', 'Entendeu o caminho? O melhor jeito de aprender é repetir com as próprias mãos:'),
            h('button.btn.btn-primary.btn-sm', { type: 'button', onclick: () => this.tryYourself() }, '↺ Agora é sua vez'),
            h('button.btn.btn-ghost.btn-sm', { type: 'button', onclick: () => this.showMe(true) }, '▶ Ver de novo')));
        }
      } else this.successEl.style.display = 'none';
      this.btnBack.disabled = this.idx === 0;
      this.btnHint.style.display = isTask && st.hint && !this.done ? '' : 'none';
      this.btnSolve.style.display = isTask && (st.solve || this.demoFor(st)) && !this.done ? '' : 'none';
      this.btnNext.disabled = isTask && !this.done;
      this.btnNext.textContent = this.idx === steps.length - 1 ? 'Concluir lição ✓' : 'Continuar ›';
      this.btnNext.classList.toggle('pulse', this.done && isTask);
      this.updateTargets();
      this.body.scrollTop = 0;
    }

    showHint() {
      this.hintEl.style.display = '';
      this.hintEl.classList.remove('flash');
      void this.hintEl.offsetWidth;
      this.hintEl.classList.add('flash');
    }

    // Faz o passo na hora, sem animação (usado como reserva e nos testes automáticos).
    solve() {
      const st = this.lesson.steps[this.idx];
      if (!st.solve || this.done) return;
      if (this.app.grid.editing) this.app.grid.cancel();
      st.solve(this.ctx);
      this.check();
    }

    demoFor(st) {
      let d = st.demo || (XL.Demos[this.lesson.id] && XL.Demos[this.lesson.id][st.title]);
      if (typeof d === 'function') d = d(this.ctx);
      return d && d.length ? d : null;
    }

    // ---------- "Mostre-me como" ----------
    snapshot() {
      const app = this.app;
      return {
        wb: app.wb.snapshot(), undo: app.undoStack.slice(), tab: app.ribbon.active, sel: Object.assign({}, app.grid.sel),
        top: app.grid.scroller.scrollTop, left: app.grid.scroller.scrollLeft,
      };
    }
    restore(s) {
      const app = this.app;
      XL.ui.closePopup();
      XL.ui.closeAllModals();
      if (XL.Find) XL.Find.close();
      if (XL.Print && XL.Print.current) XL.Print.current.close(true);
      if (app.grid.editing) app.grid.cancel();
      this.settingUp = true;
      app.wb.restore(s.wb);
      app.undoStack = s.undo.slice();
      app.redoStack = [];
      app.afterWorkbookChange();
      app.ribbon.setTab(s.tab === 'file' ? 'home' : s.tab, true);
      app.grid.setSel(s.sel.ar, s.sel.ac, s.sel.fr, s.sel.fc, false);
      app.grid.scroller.scrollTop = s.top;
      app.grid.scroller.scrollLeft = s.left;
      this.settingUp = false;
    }

    async showMe(again) {
      const st = this.lesson.steps[this.idx];
      if (this.demoRunning) return;
      if (again && this.demoSnap) this.restore(this.demoSnap);
      else if (this.done) return;
      const demo = this.demoFor(st);
      const snap = again && this.demoSnap ? this.demoSnap : this.snapshot();
      this.done = false;
      this.ctx.events = [];
      if (!demo) { this.solve(); this.demoed = this.done; this.demoSnap = snap; this.render(); return; }
      this.demoRunning = true;
      this.el.classList.add('demo-running');
      this.render();
      if (this.app.grid.editing) this.app.grid.cancel();
      const finished = await new XL.Demo(this).run(demo, this.btnSolve);
      this.demoRunning = false;
      this.el.classList.remove('demo-running');
      if (!finished) {
        // parou no meio: volta como estava, para não deixar a tarefa pela metade
        this.restore(snap);
        this.done = false;
        this.ctx.events = [];
        this.render();
        return;
      }
      if (!this.done && st.solve) { st.solve(this.ctx); }
      this.check();
      this.demoed = this.done;
      this.demoSnap = snap;
      this.render();
    }

    // Desfaz a demonstração para a pessoa repetir sozinha.
    tryYourself() {
      if (!this.demoSnap) return;
      this.restore(this.demoSnap);
      this.done = false;
      this.demoed = false;
      this.ctx.events = [];
      this.render();
      this.app.grid.focus();
      XL.ui.toast('Sua vez! Siga a tarefa em amarelo — o guia confere quando você acertar. 💪');
    }

    onEvent(evt, data) {
      if (this.settingUp) return;
      this.ctx.events.push({ evt, data });
      if (CHECK_EVENTS.has(evt)) this.check();
      if (evt === 'colwidth' || evt === 'change' || evt === 'sheetswitch') this.updateTargets();
    }

    check() {
      const st = this.lesson.steps[this.idx];
      if (this.done || !st.check) return;
      let ok = false;
      try { ok = !!st.check(this.ctx); } catch (e) { console.error(e); }
      if (!ok) return;
      this.done = true;
      this.praise = PRAISE[Math.floor(Math.random() * PRAISE.length)];
      this.render();
      this.app.emit('stepdone', { idx: this.idx });
    }

    next() {
      const st = this.lesson.steps[this.idx];
      if (st.check && !this.done) return;
      if (this.idx === this.lesson.steps.length - 1) { this.opts.onComplete && this.opts.onComplete(); return; }
      this.goto(this.idx + 1);
    }

    // ---------- Destaques ----------
    updateTargets() {
      const st = this.lesson.steps[this.idx];
      let targets = st.target || [];
      if (typeof targets === 'function') targets = targets(this.ctx) || [];
      if (!Array.isArray(targets)) targets = [targets];
      if (this.done && st.check) targets = [];
      const guides = [];
      this.domTargets = [];
      for (const t of targets) {
        if (t.startsWith('cell:')) {
          // "cell:B2" (planilha ativa) ou "cell:Resumo!B2" (só aparece quando essa planilha está aberta)
          let spec = t.slice(5);
          const i = spec.lastIndexOf('!');
          if (i > 0) {
            if (XL.normName(spec.slice(0, i)) !== XL.normName(this.app.sheet.name)) continue;
            spec = spec.slice(i + 1);
          }
          const g = XL.parseRange(spec);
          if (g) guides.push(g);
        } else this.domTargets.push(t);
      }
      this.app.grid.setGuides(guides);
      if (guides.length && !this.done) this.app.grid.scrollToRange(guides[0]);
    }

    loop() {
      const sels = this.domTargets || [];
      const els = [];
      for (const s of sels) document.querySelectorAll(s).forEach((e) => els.push(e));
      while (this.rings.length < els.length) { const r = h('div.tutor-ring'); document.body.appendChild(r); this.rings.push(r); }
      this.rings.forEach((ring, i) => {
        const e = els[i];
        const rect = e && e.getBoundingClientRect();
        if (!rect || (!rect.width && !rect.height)) { ring.style.display = 'none'; return; }
        Object.assign(ring.style, {
          display: '', left: rect.left - 4 + 'px', top: rect.top - 4 + 'px', width: rect.width + 8 + 'px', height: rect.height + 8 + 'px',
        });
      });
      this.raf = requestAnimationFrame(() => this.loop());
    }
  }

  XL.Tutor = Tutor;

  // ---------- Registro de lições ----------
  XL.Lessons = {};
  XL.registerLesson = (def) => { XL.Lessons[def.id] = def; };
})();
