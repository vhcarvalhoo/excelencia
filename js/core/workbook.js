/* Pasta de trabalho: várias planilhas, renomear/excluir com ajuste de fórmulas,
   inserir/excluir linhas e colunas e desfazer de tudo. */
(function () {
  'use strict';
  const XL = window.XL;
  const K = XL.SHEET_KEY;

  class Workbook {
    constructor(rows = 100, cols = 26) {
      this.rows = rows;
      this.cols = cols;
      this.reset();
    }

    reset() {
      this.names = []; // nomes definidos: { name, ref, comment }
      this.sheets = [];
      this.activeIndex = 0;
      this.addSheet('Planilha1');
    }

    get active() { return this.sheets[this.activeIndex]; }
    byName(name) {
      const k = XL.normName(String(name).trim());
      return this.sheets.find((s) => XL.normName(s.name) === k) || null;
    }
    nextName() {
      let i = 1;
      while (this.byName('Planilha' + i)) i++;
      return 'Planilha' + i;
    }
    addSheet(name, index) {
      const s = new XL.Sheet(this.rows, this.cols, name || this.nextName());
      s.wb = this;
      if (index === undefined || index > this.sheets.length) this.sheets.push(s);
      else this.sheets.splice(index, 0, s);
      this.invalidateAll();
      return s;
    }

    // Retorna a mensagem de erro (ou null se o nome for válido).
    validateName(name, except) {
      const n = String(name || '').trim();
      if (!n) return 'O nome da planilha não pode ficar vazio.';
      if (n.length > 31) return 'O nome da planilha pode ter no máximo 31 caracteres.';
      if (/[\\/?*[\]:]/.test(n)) return 'O nome não pode ter estes caracteres:  \\  /  ?  *  [  ]  :';
      if (n.startsWith("'") || n.endsWith("'")) return 'O nome não pode começar nem terminar com apóstrofo (\').';
      const other = this.byName(n);
      if (other && other !== except) return `Já existe uma planilha chamada "${other.name}". Escolha outro nome.`;
      return null;
    }

    forEachFormula(fn) {
      for (const s of this.sheets) for (const cell of s.cells.values()) if (cell.kind === 'formula') fn(s, cell);
    }

    renameSheet(sheet, name) {
      const old = sheet.name;
      sheet.name = String(name).trim();
      this.forEachFormula((s, cell) => {
        const ni = XL.Formula.renameSheetRefs(cell.input, old, sheet.name);
        if (ni !== cell.input) s.setFormulaText(cell, ni);
      });
      for (const n of this.names) { n.ref = XL.Formula.renameSheetRefs(n.ref, old, sheet.name); n._ast = null; }
      for (const s of this.sheets) for (const p of s.pivots || []) if (p.src.sheet && XL.normName(p.src.sheet) === XL.normName(old)) p.src.sheet = sheet.name;
      this.invalidateAll();
    }

    removeSheet(sheet) {
      if (this.sheets.length <= 1) return false;
      const i = this.sheets.indexOf(sheet);
      this.sheets.splice(i, 1);
      this.forEachFormula((s, cell) => {
        const ni = XL.Formula.dropSheetRefs(cell.input, sheet.name);
        if (ni !== cell.input) s.setFormulaText(cell, ni);
      });
      if (this.activeIndex >= this.sheets.length) this.activeIndex = this.sheets.length - 1;
      else if (i < this.activeIndex) this.activeIndex--;
      this.invalidateAll();
      return true;
    }

    invalidateAll() { for (const s of this.sheets) s.clearCache(); }

    // ---------- Nomes definidos (Gerenciador de Nomes) ----------
    nameByKey(k) { return this.names.find((n) => XL.normName(n.name) === k) || null; }
    nameValue(k, sheet) {
      const n = this.nameByKey(k);
      if (!n) return undefined;
      if (n._busy) return XL.Formula.err(XL.Formula.E.REF);
      n._busy = true;
      try {
        if (!n._ast) n._ast = XL.Formula.parse(XL.Formula.normalize(n.ref).slice(1));
        return XL.Formula.evaluate(n._ast, sheet.ctx);
      } catch (e) {
        return XL.Formula.err(XL.Formula.E.NAME);
      } finally { n._busy = false; }
    }
    // Retorna a mensagem de erro (ou null) para um nome novo
    validateDefinedName(name, except) {
      const n = String(name || '').trim();
      if (!n) return 'Digite um nome.';
      if (!/^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*$/.test(n)) return 'O nome deve começar com uma letra ou _ e não pode ter espaços nem símbolos (use _ ou ponto).';
      if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[RrCcLl]$/.test(n)) return `"${n}" parece um endereço de célula. Escolha outro nome.`;
      if (XL.Functions.get(n)) return `"${n}" já é o nome de uma função do Excel.`;
      const other = this.nameByKey(XL.normName(n));
      if (other && other !== except) return `Já existe um nome "${other.name}".`;
      return null;
    }

    // ---------- Inserir / excluir linhas e colunas ----------
    // kind: 'insert'|'delete', axis: 'row'|'col', at: índice 0-based, n: quantidade
    structural(sheet, kind, axis, at, n) {
      const max = axis === 'row' ? sheet.rows : sheet.cols;
      const moved = new Map();
      for (const [k, cell] of sheet.cells) {
        const r = Math.floor(k / K), c = k % K;
        let x = axis === 'row' ? r : c;
        if (kind === 'insert') { if (x >= at) x += n; }
        else { if (x >= at && x < at + n) continue; if (x >= at + n) x -= n; }
        if (x >= max) continue;
        moved.set(axis === 'row' ? x * K + c : r * K + x, cell);
      }
      sheet.cells = moved;
      const op = { kind, axis, at: at + 1, n, max };
      const target = XL.normName(sheet.name);
      this.forEachFormula((s, cell) => {
        const ni = XL.Formula.adjustStructure(cell.input, op, (tokSheet) => (tokSheet ? XL.normName(tokSheet) === target : s === sheet));
        if (ni !== cell.input) s.setFormulaText(cell, ni);
      });
      sheet.adjustState(op);
      if (sheet.freeze) {
        const f = sheet.freeze;
        const key = axis === 'row' ? 'rows' : 'cols';
        if (f[key] > at) f[key] = Math.max(0, kind === 'insert' ? Math.min(max - 1, f[key] + n) : f[key] - Math.min(n, f[key] - at));
      }
      this.invalidateAll();
    }

    // ---------- Desfazer ----------
    snapshot() {
      return {
        active: this.activeIndex, names: this.names.map((n) => ({ name: n.name, ref: n.ref, comment: n.comment || '' })),
        sheets: this.sheets.map((s) => ({ name: s.name, sel: s.sel ? Object.assign({}, s.sel) : null, data: s.snapshot() })),
      };
    }
    restore(snap) {
      this.sheets = snap.sheets.map((x) => {
        const s = new XL.Sheet(this.rows, this.cols, x.name);
        s.wb = this;
        s.sel = x.sel;
        return s;
      });
      snap.sheets.forEach((x, i) => this.sheets[i].restore(x.data));
      this.activeIndex = Math.min(snap.active, this.sheets.length - 1);
      this.names = (snap.names || []).map((n) => Object.assign({}, n));
      this.invalidateAll();
    }
  }

  XL.Workbook = Workbook;
})();
