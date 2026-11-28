/* Módulo 9: Definir Nome e Gerenciador de Nomes (intervalos, constantes e funções LAMBDA com nome). */
(function () {
  'use strict';
  const XL = window.XL;
  const h = XL.h;
  const proto = XL.ExcelApp.prototype;
  const baseExtra = proto.execExtra;
  const enterSubmits = (inp) => inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const b = inp.closest('.ui-modal').querySelector('.btn-primary'); if (b) b.click(); } });

  // texto curto do valor de um nome (para o Gerenciador)
  const preview = (app, n) => {
    const v = app.wb.nameValue(XL.normName(n.name), app.sheet);
    if (v instanceof XL.Formula.Lambda) return 'função LAMBDA (' + v.params.length + ' parâmetro' + (v.params.length === 1 ? '' : 's') + ')';
    if (v instanceof XL.Formula.Matrix) return '{' + v.flat().slice(0, 4).map((x) => XL.formatValue(x, {}).text).join('; ') + (v.h * v.w > 4 ? '; ...' : '') + '}';
    return XL.formatValue(v, {}).text;
  };

  Object.assign(proto, {
    execExtra(cmd, arg) {
      if (cmd === 'name-define') return this.defineNameDialog();
      if (cmd === 'name-manager') return this.nameManagerDialog();
      return baseExtra.call(this, cmd, arg);
    },

    defineNameDialog(existing) {
      const g = this.grid.range();
      const refDefault = '=' + XL.quoteSheet(this.sheet.name) + '!$' + XL.numToCol(g.c1 + 1) + '$' + (g.r1 + 1) +
        (g.r1 !== g.r2 || g.c1 !== g.c2 ? ':$' + XL.numToCol(g.c2 + 1) + '$' + (g.r2 + 1) : '');
      const name = h('input.fd-search', { type: 'text', value: existing ? existing.name : '', dataset: { role: 'nm-name' }, placeholder: 'Ex.: Comissao' });
      const comment = h('input.fd-search', { type: 'text', value: existing ? existing.comment || '' : '', dataset: { role: 'nm-comment' } });
      const ref = h('input.fd-search.nm-ref', { type: 'text', value: existing ? existing.ref : refDefault, dataset: { role: 'nm-ref' } });
      [name, comment, ref].forEach(enterSubmits);
      XL.ui.modal({
        title: existing ? 'Editar Nome' : 'Novo Nome', width: 480,
        body: h('div.pvd',
          h('label.fd-lbl', 'Nome:'), name,
          h('label.fd-lbl', 'Comentário:'), comment,
          h('label.fd-lbl', 'Refere-se a:'), ref,
          h('p.cd-note', 'Pode ser um intervalo (=Vendas!$C$2:$C$20), um valor (=0,15) ou uma função sua: =LAMBDA(valor; valor*0,15).')),
        buttons: [{
          label: 'OK', primary: true,
          onClick: () => {
            const nm = name.value.trim();
            const problem = this.wb.validateDefinedName(nm, existing);
            if (problem) { XL.ui.toast(problem); return false; }
            let f = ref.value.trim();
            if (!f) { XL.ui.toast('Preencha o campo "Refere-se a".'); return false; }
            if (f[0] !== '=') f = '=' + f;
            try { f = XL.Formula.normalize(f); XL.Formula.parse(f.slice(1)); } catch (e) { XL.ui.toast('Há um problema com a fórmula: ' + XL.escapeHtml(e.message)); return false; }
            this.pushUndo();
            if (existing) Object.assign(existing, { name: nm, ref: f, comment: comment.value.trim(), _ast: null });
            else this.wb.names.push({ name: nm, ref: f, comment: comment.value.trim() });
            this.wb.invalidateAll();
            this.render();
            this.updateFormulaBar();
            this.emit('name', { name: nm, ref: f, lambda: /^=\s*LAMBDA\s*\(/i.test(f) });
            this.emit('change', {});
          },
        }, { label: 'Cancelar' }],
        onClose: () => this.grid.focus(),
      });
      setTimeout(() => (existing ? ref : name).focus(), 0);
    },

    nameManagerDialog() {
      const names = this.wb.names;
      let pick = null;
      const body = h('table.nm-table', h('thead', h('tr', h('th', 'Nome'), h('th', 'Valor'), h('th', 'Refere-se a'), h('th', 'Comentário'))));
      const tb = h('tbody');
      names.forEach((n) => {
        const tr = h('tr', { dataset: { name: n.name }, tabindex: '0' }, h('td', n.name), h('td', preview(this, n)), h('td', h('code', n.ref)), h('td', n.comment || ''));
        tr.addEventListener('click', () => { pick = n; tb.querySelectorAll('tr').forEach((x) => x.classList.toggle('on', x === tr)); });
        tr.addEventListener('dblclick', () => { m.close(); this.defineNameDialog(n); });
        tb.appendChild(tr);
      });
      if (!names.length) tb.appendChild(h('tr', h('td.nm-empty', { colSpan: 4 }, 'Nenhum nome definido ainda. Clique em Novo...')));
      body.appendChild(tb);
      const m = XL.ui.modal({
        title: 'Gerenciador de Nomes', width: 640,
        body: h('div.nm', body),
        buttons: [
          { label: 'Novo...', role: 'nm-new', onClick: () => { setTimeout(() => this.defineNameDialog(), 0); } },
          { label: 'Editar...', role: 'nm-edit', onClick: () => { if (!pick) { XL.ui.toast('Clique em um nome da lista primeiro.'); return false; } const n = pick; setTimeout(() => this.defineNameDialog(n), 0); } },
          {
            label: 'Excluir', role: 'nm-delete', onClick: () => {
              if (!pick) { XL.ui.toast('Clique em um nome da lista primeiro.'); return false; }
              this.pushUndo();
              this.wb.names = this.wb.names.filter((x) => x !== pick);
              this.wb.invalidateAll();
              this.render();
              this.emit('namedelete', { name: pick.name });
              this.emit('change', {});
            },
          },
          { label: 'Fechar', primary: true },
        ],
        onClose: () => this.grid.focus(),
      });
    },
  });
})();
