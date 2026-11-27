/* "Mostre-me como": o guia executa a tarefa devagar, com um cursor animado e um balão
   explicando cada clique, tecla e texto digitado — usando a interface de verdade. */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  // XL.demoSpeed < 1 acelera (usado nos testes automáticos)
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms * (XL.demoSpeed || 1)));
  const esc = XL.escapeHtml;

  // ---------- Pequena linguagem para descrever demonstrações ----------
  // Cada função devolve uma ação; "say" (opcional) troca o texto do balão.
  const D = (XL.DemoDSL = {
    cell: (ref, say) => ({ a: 'cell', ref, say }), // clicar numa célula
    drag: (from, to, say) => ({ a: 'drag', from, to, say }), // arrastar para selecionar
    point: (range, say) => ({ a: 'point', range, say }), // clicar/arrastar enquanto digita uma fórmula
    type: (text, say) => ({ a: 'type', text, say }), // digitar na célula/barra
    key: (k, say) => ({ a: 'key', k, say }), // tecla: 'Enter', 'Tab', 'F4', 'Ctrl+Z'...
    tab: (id, say) => ({ a: 'tab', id, say }), // guia da Faixa de Opções
    btn: (cmd, arg, say) => ({ a: 'btn', cmd, arg, say }), // botão da Faixa de Opções
    menu: (id, say) => ({ a: 'menu', id, say }), // abrir menu/setinha (data-menu)
    item: (cmd, arg, say) => ({ a: 'item', cmd, arg, say }), // item de um menu aberto
    el: (sel, say) => ({ a: 'el', sel, say }), // clicar em qualquer elemento
    rclick: (sel, say) => ({ a: 'rclick', sel, say }), // botão direito
    dbl: (sel, say) => ({ a: 'dbl', sel, say }), // duplo clique
    set: (sel, value, say) => ({ a: 'set', sel, value, say }), // escolher numa lista
    typeIn: (sel, text, say) => ({ a: 'typeIn', sel, text, say }), // digitar num campo
    fill: (from, to, say) => ({ a: 'fill', from, to, say }), // alça de preenchimento
    scroll: (top, left, say) => ({ a: 'scroll', top, left, say }),
    look: (target, say) => ({ a: 'look', target, say }), // apontar sem clicar
    sheet: (name, say) => ({ a: 'el', sel: `.xl-sheettab[data-sheet="${name}"]`, say: say || `Clique na aba <b>${esc(name)}</b>` }),
    when: (cond, action) => Object.assign({}, action, { when: cond }), // só executa se cond(x) for verdadeiro
  });

  const KEYNAMES = { ArrowRight: '→', ArrowLeft: '←', ArrowUp: '↑', ArrowDown: '↓', Delete: 'Delete', Escape: 'Esc' };
  const kbd = (k) => k.split('+').map((p) => `<kbd>${esc(KEYNAMES[p] || p)}</kbd>`).join('+');
  const labelOf = (el) => {
    const t = (el.getAttribute('title') || el.textContent || '').replace(/\s*\(mais opções\)/, '').replace(/\s+/g, ' ').trim();
    return esc(t);
  };
  const CURSOR = '<svg viewBox="0 0 24 24" width="26" height="26"><path d="M4 2.5 19.5 14l-7 1.2 4.1 7.3-3 1.6-4.1-7.4L4 21.5z" fill="#fff" stroke="#111" stroke-width="1.4" stroke-linejoin="round"/></svg>';

  class Demo {
    constructor(tutor) {
      this.tutor = tutor;
      this.app = tutor.app;
      this.stopped = false;
    }

    async run(actions, startEl) {
      const app = this.app;
      this.blocker = h('div.demo-block');
      this.cursor = h('div.demo-cursor', { html: CURSOR });
      this.ring = h('div.demo-ring');
      this.bubble = h('div.demo-bubble');
      this.stopBtn = h('button.demo-stop', { type: 'button', onclick: () => { this.stopped = true; } }, '⏹ Parar');
      document.body.append(this.blocker, this.ring, this.cursor, this.bubble);
      const r0 = startEl ? startEl.getBoundingClientRect() : { left: window.innerWidth / 2, top: window.innerHeight / 2, width: 0, height: 0 };
      this.pos = { x: r0.left + r0.width / 2, y: r0.top + r0.height / 2 };
      this.place(true);
      this.onKey = (e) => { if (e.key === 'Escape' && e.isTrusted) { this.stopped = true; } };
      document.addEventListener('keydown', this.onKey, true);
      const total = actions.filter((a) => !a.when || a.when(this.tutor.ctx)).length;
      let i = 0;
      try {
        await sleep(250);
        for (const a of actions) {
          if (this.stopped) break;
          if (a.when && !a.when(this.tutor.ctx)) continue;
          i++;
          this.count = `Passo ${i} de ${total}`;
          await this.exec(a);
          await sleep(350);
        }
        if (!this.stopped) await sleep(500);
      } catch (e) {
        console.error('Demonstração interrompida:', e);
      } finally {
        document.removeEventListener('keydown', this.onKey, true);
        [this.blocker, this.cursor, this.ring, this.bubble].forEach((x) => x.remove());
        app.render();
      }
      return !this.stopped;
    }

    // ---------- Cursor e balão ----------
    place(instant) {
      this.cursor.style.transition = instant ? 'none' : '';
      this.cursor.style.left = this.pos.x + 'px';
      this.cursor.style.top = this.pos.y + 'px';
      this.placeBubble();
    }
    placeBubble() {
      const b = this.bubble;
      let x = this.pos.x + 26, y = this.pos.y + 24;
      const w = b.offsetWidth || 280, hh = b.offsetHeight || 80;
      if (x + w > window.innerWidth - 10) x = this.pos.x - w - 14;
      if (y + hh > window.innerHeight - 10) y = this.pos.y - hh - 14;
      b.style.left = Math.max(8, x) + 'px';
      b.style.top = Math.max(8, y) + 'px';
    }
    say(html) {
      this.bubble.innerHTML = `<div class="demo-count">${this.count || ''}</div><div class="demo-text">${html}</div>`;
      this.bubble.appendChild(this.stopBtn);
      this.placeBubble();
    }
    async move(p, ms = 650) {
      this.pos = p;
      ms = ms * (XL.demoSpeed || 1);
      this.cursor.style.transitionDuration = ms + 'ms';
      this.cursor.style.left = p.x + 'px';
      this.cursor.style.top = p.y + 'px';
      this.bubble.style.transition = `left ${ms}ms ease, top ${ms}ms ease`;
      this.placeBubble();
      await new Promise((r) => setTimeout(r, ms + 60));
    }
    async click(double) {
      for (let n = 0; n < (double ? 2 : 1); n++) {
        this.cursor.classList.remove('press');
        void this.cursor.offsetWidth;
        this.cursor.classList.add('press');
        const rip = h('div.demo-ripple', { style: { left: this.pos.x + 'px', top: this.pos.y + 'px' } });
        document.body.appendChild(rip);
        setTimeout(() => rip.remove(), 700);
        await sleep(double ? 160 : 220);
      }
    }
    highlight(rect) {
      if (!rect) { this.ring.style.display = 'none'; return; }
      Object.assign(this.ring.style, { display: '', left: rect.left - 4 + 'px', top: rect.top - 4 + 'px', width: rect.width + 8 + 'px', height: rect.height + 8 + 'px' });
    }

    // ---------- Localizar coisas na tela ----------
    async waitFor(sel, ms = 2000) {
      const t0 = Date.now();
      for (;;) {
        const el = [...document.querySelectorAll(sel)].find((e) => { const r = e.getBoundingClientRect(); return r.width || r.height; });
        if (el) return el;
        if (Date.now() - t0 > ms) throw new Error('Elemento não encontrado: ' + sel);
        await sleep(50);
      }
    }
    center(el) { const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }
    cellRect(ref) {
      const rg = XL.parseRange(ref); // aceita "B5" ou "C2:C7"
      const g = this.app.grid;
      g.scrollIntoView(rg.r1, rg.c1);
      const a = g.tds[rg.r1][rg.c1].getBoundingClientRect();
      const b = g.tds[rg.r2][rg.c2].getBoundingClientRect();
      return { left: a.left, top: a.top, width: b.right - a.left, height: b.bottom - a.top, right: b.right, bottom: b.bottom };
    }
    cellPoint(ref) {
      const r = this.cellRect(ref);
      return { x: r.left + Math.min(r.width / 2, 40), y: r.top + r.height / 2 };
    }
    // Garante que o botão está visível, trocando de guia se preciso (mostrando o clique na guia).
    async ribbonEl(sel, pred) {
      let el = document.querySelector(sel);
      if (el && el.getBoundingClientRect().width) return el;
      const tabId = this.app.ribbon.findTabFor(pred);
      if (tabId && tabId !== this.app.ribbon.active) {
        await this.exec({ a: 'tab', id: tabId });
        await sleep(250);
      }
      return this.waitFor(sel);
    }

    // ---------- Ações ----------
    async exec(a) {
      const app = this.app;
      const g = app.grid;
      switch (a.a) {
        case 'cell': {
          this.say(a.say || `Clique na célula <b>${a.ref}</b>`);
          const p = this.cellPoint(a.ref);
          await this.move(p);
          await this.click();
          if (g.editing && !g.commit()) g.cancel();
          app.select(a.ref);
          return;
        }
        case 'drag': {
          this.say(a.say || `Clique em <b>${a.from}</b> e, segurando o botão do mouse, arraste até <b>${a.to}</b>`);
          await this.move(this.cellPoint(a.from));
          await this.click();
          if (g.editing && !g.commit()) g.cancel();
          const f = XL.parseRef(a.from), t = XL.parseRef(a.to);
          g.setSel(f.r, f.c, f.r, f.c, false);
          const end = this.cellPoint(a.to);
          await this.move(end, 900);
          g.setSel(f.r, f.c, t.r, t.c, true);
          return;
        }
        case 'point': {
          const rg = XL.parseRange(a.range);
          const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
          this.say(a.say || (multi
            ? `Com a fórmula aberta, arraste de <b>${XL.toRef(rg.r1, rg.c1)}</b> até <b>${XL.toRef(rg.r2, rg.c2)}</b> — o intervalo entra na fórmula`
            : `Com a fórmula aberta, clique em <b>${a.range}</b> — a referência entra na fórmula`));
          await this.move(this.cellPoint(XL.toRef(rg.r1, rg.c1)));
          await this.click();
          g.startPoint({ r: rg.r1, c: rg.c1 });
          if (multi) {
            await this.move(this.cellPoint(XL.toRef(rg.r2, rg.c2)), 800);
            g.updatePoint({ r: rg.r2, c: rg.c2 });
          }
          g.drag = null;
          g.focusActiveInput();
          return;
        }
        case 'type': {
          this.say(a.say || `Digite <code>${esc(a.text)}</code>`);
          if (!g.editing) {
            g.beginEdit('enter', '');
          }
          const inp = g.activeInput();
          await this.move(this.center(inp === g.editor && !g.isAway() ? g.editor : app.fbar), 450);
          inp.focus();
          for (const ch of a.text) {
            if (this.stopped) return;
            const s = inp.selectionStart, e2 = inp.selectionEnd;
            inp.setRangeText(ch, s, e2, 'end');
            if (inp === g.editor) g.onEditorInput(); else g.onBarInput();
            await sleep(95);
          }
          return;
        }
        case 'typeIn': {
          const el = await this.waitFor(a.sel);
          this.say(a.say || `Digite <code>${esc(a.text)}</code>`);
          await this.move(this.center(el));
          await this.click();
          el.focus();
          el.value = '';
          for (let k = 1; k <= a.text.length; k++) {
            if (this.stopped) return;
            el.value = a.text.slice(0, k);
            el.dispatchEvent(new Event('input', { bubbles: true }));
            await sleep(95);
          }
          return;
        }
        case 'key': {
          this.say(a.say || `Pressione ${kbd(a.k)}`);
          const parts = a.k.split('+');
          let key = parts.pop();
          const mods = { ctrlKey: parts.includes('Ctrl'), shiftKey: parts.includes('Shift'), altKey: parts.includes('Alt') };
          if (key.length === 1 && mods.ctrlKey) key = key.toLowerCase();
          const cap = h('div.demo-key', { html: kbd(a.k) });
          cap.style.left = this.pos.x + 'px';
          cap.style.top = this.pos.y - 46 + 'px';
          document.body.appendChild(cap);
          await sleep(500);
          let target = document.activeElement;
          const inside = target && target !== document.body && (app.root.contains(target) || target.closest('.xl-find, .ui-modal, .ui-menu'));
          if (!inside) { g.focus(); target = document.activeElement; }
          target.dispatchEvent(new KeyboardEvent('keydown', Object.assign({ key, bubbles: true, cancelable: true }, mods)));
          await sleep(450);
          cap.remove();
          return;
        }
        case 'tab': {
          const el = await this.waitFor(`.xl-tab[data-tab="${a.id}"]`);
          this.say(a.say || `Clique na guia <b>${esc(el.textContent)}</b>`);
          await this.move(this.center(el));
          await this.click();
          el.click();
          return;
        }
        case 'btn': {
          const sel = `.rb-btn[data-cmd="${a.cmd}"]` + (a.arg !== undefined ? `[data-arg="${a.arg}"]` : '');
          const el = await this.ribbonEl(sel, (it) => it.cmd === a.cmd && (a.arg === undefined || String(it.arg) === String(a.arg)));
          this.say(a.say || `Clique em <b>${labelOf(el)}</b>`);
          this.highlight(el.getBoundingClientRect());
          await this.move(this.center(el));
          await this.click();
          this.highlight(null);
          el.click();
          return;
        }
        case 'menu': {
          const el = await this.ribbonEl(`[data-menu="${a.id}"]`, (it) => (it.menuId || it.cmd) === a.id);
          this.say(a.say || `Clique em <b>${labelOf(el)}</b> para abrir as opções`);
          this.highlight(el.getBoundingClientRect());
          await this.move(this.center(el));
          await this.click();
          this.highlight(null);
          el.click();
          return;
        }
        case 'item': {
          const el = await this.waitFor(`.ui-menu-item[data-cmd="${a.cmd}"]` + (a.arg !== undefined ? `[data-arg="${a.arg}"]` : ''));
          this.say(a.say || `Escolha <b>${labelOf(el.querySelector('.ui-menu-lbl') || el)}</b>`);
          await this.move(this.center(el));
          await this.click();
          el.click();
          return;
        }
        case 'el': {
          const el = await this.waitFor(a.sel);
          this.say(a.say || `Clique em <b>${labelOf(el)}</b>`);
          this.highlight(el.getBoundingClientRect());
          await this.move(this.center(el));
          await this.click();
          this.highlight(null);
          el.click();
          return;
        }
        case 'rclick': {
          const el = await this.waitFor(a.sel);
          this.say(a.say || `Clique com o <b>botão direito</b> em <b>${labelOf(el)}</b>`);
          const p = this.center(el);
          await this.move(p);
          await this.click();
          el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: p.x, clientY: p.y }));
          return;
        }
        case 'dbl': {
          const el = await this.waitFor(a.sel);
          this.say(a.say || `Dê um <b>duplo clique</b> em <b>${labelOf(el)}</b>`);
          await this.move(this.center(el));
          await this.click(true);
          el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
          return;
        }
        case 'set': {
          const el = await this.waitFor(a.sel);
          const opt = [...el.options].find((o) => o.value === String(a.value));
          this.say(a.say || `Na lista, escolha <b>${esc(opt ? opt.textContent : a.value)}</b>`);
          this.highlight(el.getBoundingClientRect());
          await this.move(this.center(el));
          await this.click();
          this.highlight(null);
          el.value = String(a.value);
          el.dispatchEvent(new Event('change', { bubbles: true }));
          return;
        }
        case 'fill': {
          const src = XL.parseRange(a.from);
          const to = XL.parseRef(a.to);
          if (XL.rangeToStr(g.range()) !== XL.rangeToStr(src)) await this.exec(src.r1 === src.r2 && src.c1 === src.c2 ? { a: 'cell', ref: a.from } : { a: 'drag', from: XL.toRef(src.r1, src.c1), to: XL.toRef(src.r2, src.c2) });
          this.say(a.say || `Pegue a <b>alça de preenchimento</b> (o quadradinho verde no canto) e arraste até <b>${a.to}</b>`);
          await this.move(this.center(g.handle));
          await this.click();
          g.drag = { type: 'fill', src, dst: null, dir: null };
          const end = this.cellPoint(a.to);
          g.updateFill(to);
          await this.move(end, 1000);
          const d = g.drag;
          g.drag = null;
          g.finishFill(d);
          return;
        }
        case 'scroll': {
          this.say(a.say || (a.top !== null && a.top !== undefined ? 'Role a planilha <b>para baixo</b> (rodinha do mouse)' : 'Role a planilha <b>para o lado</b> (barra de rolagem de baixo)'));
          await this.move(this.center(g.el), 450);
          const sc = g.scroller;
          const t0 = sc.scrollTop, l0 = sc.scrollLeft;
          const t1 = a.top === null || a.top === undefined ? t0 : a.top;
          const l1 = a.left === null || a.left === undefined ? l0 : a.left;
          for (let k = 1; k <= 20; k++) {
            sc.scrollTop = t0 + ((t1 - t0) * k) / 20;
            sc.scrollLeft = l0 + ((l1 - l0) * k) / 20;
            await sleep(40);
          }
          app.emit('scroll', { top: sc.scrollTop, left: sc.scrollLeft });
          return;
        }
        case 'look': {
          this.say(a.say || 'Repare aqui');
          let rect;
          if (a.target.startsWith('cell:')) rect = this.cellRect(a.target.slice(5));
          else rect = (await this.waitFor(a.target)).getBoundingClientRect();
          this.highlight(rect);
          await this.move({ x: rect.left + Math.min(rect.width / 2, 60), y: rect.top + rect.height / 2 });
          await sleep(1600);
          this.highlight(null);
          return;
        }
      }
    }
  }

  XL.Demo = Demo;
  XL.Demos = {};
  XL.registerDemos = (lessonId, map) => { XL.Demos[lessonId] = Object.assign(XL.Demos[lessonId] || {}, map); };
})();
