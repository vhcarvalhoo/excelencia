/* Teste de Hipóteses: calcular a planilha com valores hipotéticos (sem alterá-la), Tabela de Dados (TABELA),
   Atingir Meta, Cenários e o Solver (programação linear/inteira com Simplex e, se não for linear, busca Nelder-Mead). */
(function () {
  'use strict';
  const XL = window.XL;
  const K = XL.SHEET_KEY;
  const P = XL.Sheet.prototype;
  const { Matrix, err, isErr, E } = XL.Formula;
  const isNum = (v) => typeof v === 'number' && isFinite(v);

  // ---------------- Avaliação com valores hipotéticos ----------------
  // overrides: Map(chave r*K+c → valor). Só a planilha atual recebe os valores trocados.
  P.evalWith = function (r, c, overrides) {
    const self = this;
    const cache = new Map();
    const busy = new Set();
    const ctx = Object.create(this.ctx);
    const mine = (s) => !s || XL.normName(s) === XL.normName(self.name);
    const val = (rr, cc) => {
      const k = rr * K + cc;
      if (overrides.has(k)) return overrides.get(k);
      if (cache.has(k)) return cache.get(k);
      const cell = self.cells.get(k);
      if (!cell || cell.kind !== 'formula') return self.value(rr, cc);
      if (busy.has(k)) return 0;
      busy.add(k);
      const prev = self.curCell;
      self.curCell = { r: rr, c: cc };
      let v;
      try { v = XL.Formula.evaluate(cell.ast, ctx); } catch (e) { v = err(E.VALUE); }
      self.curCell = prev;
      busy.delete(k);
      if (v instanceof Matrix) v = v.h && v.w ? v.rows[0][0] : err(E.CALC);
      if (v === null || v === undefined) v = 0;
      cache.set(k, v);
      return v;
    };
    ctx.cell = (rr, cc, s) => (mine(s) ? val(rr, cc) : self.ctx.cell(rr, cc, s));
    ctx.range = (r1, c1, r2, c2, s) => {
      if (!mine(s)) return self.ctx.range(r1, c1, r2, c2, s);
      if (r1 < 0 || c1 < 0 || r2 >= self.rows || c2 >= self.cols) return err(E.REF);
      const rows = [];
      for (let i = r1; i <= r2; i++) { const row = []; for (let j = c1; j <= c2; j++) row.push(val(i, j)); rows.push(row); }
      return new Matrix(rows);
    };
    return val(r, c);
  };
  const refKey = (ref) => { const p = XL.parseRef(String(ref).replace(/\$/g, '')); return p ? p.r * K + p.c : null; };
  const cellsOf = (rangeText) => {
    const out = [];
    for (const part of String(rangeText).split(/[;,]/)) {
      const g = XL.parseRange(part.trim().replace(/\$/g, ''));
      if (!g) return null;
      XL.eachInRange(g, (r, c) => out.push({ r, c }));
    }
    return out;
  };
  XL.whatIfCells = cellsOf;

  // ---------------- Tabela de Dados: =TABELA(entrada_linha; entrada_coluna) ----------------
  P.dataTableAt = function (r, c) {
    return (this.dataTables || []).find((t) => r > t.g.r1 && r <= t.g.r2 && c > t.g.c1 && c <= t.g.c2) || null;
  };
  P.dataTableValue = function (cur) {
    if (!cur) return err(E.REF);
    const t = this.dataTableAt(cur.r, cur.c);
    if (!t) return err(E.REF, 'TABELA só funciona dentro de uma Tabela de Dados (Dados › Teste de Hipóteses › Tabela de Dados).');
    const ov = new Map();
    const rk = t.rowIn ? refKey(t.rowIn) : null, ck = t.colIn ? refKey(t.colIn) : null;
    let fr, fc;
    if (rk !== null && ck !== null) { ov.set(rk, this.value(t.g.r1, cur.c)); ov.set(ck, this.value(cur.r, t.g.c1)); fr = t.g.r1; fc = t.g.c1; }
    else if (ck !== null) { ov.set(ck, this.value(cur.r, t.g.c1)); fr = t.g.r1; fc = cur.c; }
    else if (rk !== null) { ov.set(rk, this.value(t.g.r1, cur.c)); fr = cur.r; fc = t.g.c1; }
    else return err(E.REF);
    for (const [k, v] of ov) if (v === null) ov.set(k, 0);
    const f = this.cell(fr, fc);
    if (!f || f.kind !== 'formula') return this.value(fr, fc);
    return this.evalWith(fr, fc, ov);
  };

  // ---------------- Atingir Meta ----------------
  // Procura x na célula variável para que a célula de fórmula chegue ao valor desejado.
  P.goalSeek = function (setRef, target, changeRef) {
    const sk = refKey(setRef), ck = refKey(changeRef);
    const sr = Math.floor(sk / K), sc = sk % K;
    const f = (x) => { const v = this.evalWith(sr, sc, new Map([[ck, x]])); return isNum(v) ? v - target : NaN; };
    const cur = this.value(Math.floor(ck / K), ck % K);
    let x0 = isNum(cur) ? cur : 0;
    let f0 = f(x0);
    if (Math.abs(f0) < 1e-7) return { ok: true, x: x0 };
    let x1 = x0 === 0 ? 1 : x0 * 1.01, f1 = f(x1);
    // método da secante
    for (let i = 0; i < 100 && isNum(f0) && isNum(f1); i++) {
      if (Math.abs(f1) < 1e-7) return { ok: true, x: x1 };
      if (f1 === f0) break;
      const x2 = x1 - (f1 * (x1 - x0)) / (f1 - f0);
      if (!isFinite(x2)) break;
      x0 = x1; f0 = f1; x1 = x2; f1 = f(x2);
    }
    if (isNum(f1) && Math.abs(f1) < 1e-4) return { ok: true, x: x1 };
    // bisseção: procura um intervalo com troca de sinal
    const base = isNum(cur) ? cur : 0;
    let lo = null, hi = null;
    for (let span = 1; span < 1e12 && lo === null; span *= 2) {
      for (const [a, b] of [[base, base + span], [base - span, base]]) {
        const fa = f(a), fb = f(b);
        if (isNum(fa) && isNum(fb) && fa * fb <= 0) { lo = a; hi = b; break; }
      }
    }
    if (lo === null) return { ok: false };
    let flo = f(lo);
    for (let i = 0; i < 200; i++) {
      const mid = (lo + hi) / 2, fm = f(mid);
      if (Math.abs(fm) < 1e-9 || (hi - lo) / 2 < 1e-12) return { ok: true, x: mid };
      if (flo * fm <= 0) hi = mid; else { lo = mid; flo = fm; }
    }
    return { ok: true, x: (lo + hi) / 2 };
  };

  // ---------------- Solver ----------------
  // Simplex de duas fases (maximiza c·x, A x {<=,>=,=} b, x >= 0)
  function simplex(c, cons) {
    const n = c.length;
    const rows = cons.map((k) => (k.b < 0 ? { a: k.a.map((v) => -v), b: -k.b, op: k.op === '<=' ? '>=' : k.op === '>=' ? '<=' : '=' } : k));
    const m = rows.length;
    let nSlack = 0, nArt = 0;
    rows.forEach((k) => { if (k.op !== '=') nSlack++; if (k.op !== '<=') nArt++; });
    const N = n + nSlack + nArt;
    const T = [];
    const basis = [];
    let s = n, a = n + nSlack;
    const artCols = [];
    rows.forEach((k, i) => {
      const row = new Array(N + 1).fill(0);
      k.a.forEach((v, j) => { row[j] = v; });
      row[N] = k.b;
      if (k.op === '<=') { row[s] = 1; basis[i] = s++; }
      else if (k.op === '>=') { row[s] = -1; s++; row[a] = 1; artCols.push(a); basis[i] = a++; }
      else { row[a] = 1; artCols.push(a); basis[i] = a++; }
      T.push(row);
    });
    const pivot = (pr, pc) => {
      const pv = T[pr][pc];
      for (let j = 0; j <= N; j++) T[pr][j] /= pv;
      for (let i = 0; i < m; i++) if (i !== pr && Math.abs(T[i][pc]) > 1e-12) { const f = T[i][pc]; for (let j = 0; j <= N; j++) T[i][j] -= f * T[pr][j]; }
      basis[pr] = pc;
    };
    const run = (obj, allowed) => {
      for (let it = 0; it < 500; it++) {
        // custos reduzidos
        let pc = -1;
        for (let j = 0; j < N; j++) {
          if (!allowed(j)) continue;
          let rc = obj[j];
          for (let i = 0; i < m; i++) rc -= obj[basis[i]] * T[i][j];
          if (rc > 1e-9) { pc = j; break; } // regra de Bland
        }
        if (pc < 0) return 'ok';
        let pr = -1, best = Infinity;
        for (let i = 0; i < m; i++) if (T[i][pc] > 1e-9) { const ratio = T[i][N] / T[i][pc]; if (ratio < best - 1e-12) { best = ratio; pr = i; } }
        if (pr < 0) return 'unbounded';
        pivot(pr, pc);
      }
      return 'ok';
    };
    if (artCols.length) {
      const obj1 = new Array(N).fill(0);
      artCols.forEach((j) => { obj1[j] = -1; });
      run(obj1, () => true);
      const infeas = basis.reduce((sum, j, i) => sum + (artCols.includes(j) ? T[i][N] : 0), 0);
      if (infeas > 1e-7) return { status: 'infeasible' };
      // tira artificiais que ficaram na base com valor zero
      basis.forEach((j, i) => {
        if (!artCols.includes(j)) return;
        const pc = T[i].findIndex((v, k) => k < n + nSlack && Math.abs(v) > 1e-9);
        if (pc >= 0) pivot(i, pc);
      });
    }
    const obj2 = new Array(N).fill(0);
    c.forEach((v, j) => { obj2[j] = v; });
    const st = run(obj2, (j) => !artCols.includes(j));
    if (st === 'unbounded') return { status: 'unbounded' };
    const x = new Array(n).fill(0);
    basis.forEach((j, i) => { if (j < n) x[j] = T[i][N]; });
    return { status: 'ok', x, z: c.reduce((sum, v, j) => sum + v * x[j], 0) };
  }

  // Branch and bound para variáveis inteiras
  function milp(c, cons, ints) {
    let best = null, nodes = 0;
    const solve = (extra) => {
      if (++nodes > 3000) return;
      const r = simplex(c, cons.concat(extra));
      if (r.status !== 'ok') { if (!best && r.status === 'unbounded') best = r; return; }
      if (best && best.status === 'ok' && r.z <= best.z + 1e-9) return;
      const j = ints.find((k) => Math.abs(r.x[k] - Math.round(r.x[k])) > 1e-6);
      if (j === undefined) { best = r; return; }
      const v = r.x[j];
      const e = (op, b) => { const a = new Array(c.length).fill(0); a[j] = 1; return { a, op, b }; };
      solve(extra.concat([e('<=', Math.floor(v))]));
      solve(extra.concat([e('>=', Math.ceil(v))]));
    };
    solve([]);
    return best || { status: 'infeasible' };
  }

  // Nelder-Mead com penalidade (quando o modelo não é linear)
  function nelderMead(F, x0, iters = 800) {
    const n = x0.length;
    let pts = [x0.slice()];
    for (let i = 0; i < n; i++) { const p = x0.slice(); p[i] = p[i] ? p[i] * 1.2 : 1; pts.push(p); }
    let vals = pts.map(F);
    for (let it = 0; it < iters; it++) {
      const idx = vals.map((v, i) => i).sort((a, b) => vals[a] - vals[b]);
      pts = idx.map((i) => pts[i]); vals = idx.map((i) => vals[i]);
      if (Math.abs(vals[n] - vals[0]) < 1e-10) break;
      const cen = new Array(n).fill(0);
      for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) cen[j] += pts[i][j] / n;
      const at = (t) => cen.map((v, j) => v + t * (pts[n][j] - v));
      const xr = at(-1), fr = F(xr);
      if (fr < vals[0]) { const xe = at(-2), fe = F(xe); if (fe < fr) { pts[n] = xe; vals[n] = fe; } else { pts[n] = xr; vals[n] = fr; } }
      else if (fr < vals[n - 1]) { pts[n] = xr; vals[n] = fr; }
      else {
        const xc = at(0.5), fc = F(xc);
        if (fc < vals[n]) { pts[n] = xc; vals[n] = fc; }
        else for (let i = 1; i <= n; i++) { pts[i] = pts[i].map((v, j) => pts[0][j] + 0.5 * (v - pts[0][j])); vals[i] = F(pts[i]); }
      }
    }
    return pts[0];
  }

  // model: { obj, goal: 'max'|'min'|'value', target, vars, cons: [{ ref, op, rhs }], nonneg }
  P.solve = function (model) {
    const S = this;
    const vars = cellsOf(model.vars);
    if (!vars || !vars.length) return { status: 'error', msg: 'Informe as células variáveis (ex.: B2:C2).' };
    for (const v of vars) { const cell = S.cell(v.r, v.c); if (cell && cell.kind === 'formula') return { status: 'error', msg: `A célula variável ${XL.toRef(v.r, v.c)} tem fórmula. As células variáveis devem conter apenas números.` }; }
    const ok = refKey(model.obj);
    if (ok === null) return { status: 'error', msg: 'Informe a célula do objetivo (ex.: B8).' };
    const or = Math.floor(ok / K), oc = ok % K;
    const objCell = S.cell(or, oc);
    if (!objCell || objCell.kind !== 'formula') return { status: 'error', msg: 'A célula do objetivo deve conter uma fórmula que dependa das células variáveis.' };
    const n = vars.length;
    const keys = vars.map((v) => v.r * K + v.c);
    const ov = (x) => new Map(keys.map((k, i) => [k, x[i]]));
    const num = (v) => (isNum(v) ? v : NaN);
    const fObj = (x) => num(S.evalWith(or, oc, ov(x)));
    // restrições elemento a elemento
    const ints = new Set();
    const rows = [];
    for (const k of model.cons || []) {
      const lhs = cellsOf(k.ref);
      if (!lhs) return { status: 'error', msg: `Referência inválida na restrição: ${k.ref}` };
      if (k.op === 'int') { lhs.forEach((c) => { const i = keys.indexOf(c.r * K + c.c); if (i >= 0) ints.add(i); }); continue; }
      let rhsCells = null, rhsNum = null;
      const t = String(k.rhs).trim().replace(/^=/, '');
      const pn = XL.parseInput(t);
      if (pn.kind === 'number') rhsNum = pn.value;
      else { rhsCells = cellsOf(t); if (!rhsCells || (rhsCells.length !== 1 && rhsCells.length !== lhs.length)) return { status: 'error', msg: `Restrição inválida: ${k.ref} ${k.op} ${k.rhs}` }; }
      lhs.forEach((c, i) => {
        const rc = rhsCells ? rhsCells[rhsCells.length === 1 ? 0 : i] : null;
        rows.push({
          op: k.op,
          g: (x) => num(S.evalWith(c.r, c.c, ov(x))) - (rc ? num(S.evalWith(rc.r, rc.c, ov(x))) : rhsNum),
        });
      });
    }
    if (model.goal === 'value') rows.push({ op: '=', g: (x) => fObj(x) - Number(model.target) });
    const x0 = vars.map((v) => { const x = S.value(v.r, v.c); return isNum(x) ? x : 0; });
    const zero = new Array(n).fill(0);
    const unit = (j, s = 1) => { const x = zero.slice(); x[j] = s; return x; };
    // é linear? coeficientes por diferenças e conferência em pontos de teste
    const lin = (fn) => {
      const f0 = fn(zero);
      const a = zero.map((_, j) => fn(unit(j)) - f0);
      const okk = [x0, zero.map((_, j) => 3 + j * 1.7), zero.map((_, j) => 11 - j * 2.3)].every((p) => {
        const pred = f0 + a.reduce((s, v, j) => s + v * p[j], 0);
        const real = fn(p);
        return isNum(real) && Math.abs(real - pred) <= 1e-6 * (1 + Math.abs(real));
      });
      return okk && isNum(f0) ? { f0, a } : null;
    };
    const sign = model.goal === 'min' ? -1 : 1;
    const LO = lin(fObj);
    const Ls = rows.map((k) => lin(k.g));
    let res;
    if (LO && Ls.every(Boolean)) {
      const c = model.goal === 'value' ? zero.slice() : LO.a.map((v) => v * sign);
      const cons = rows.map((k, i) => ({ a: Ls[i].a, op: k.op, b: -Ls[i].f0 }));
      if (model.nonneg === false) {
        // variáveis livres: x = x⁺ − x⁻
        const c2 = c.concat(c.map((v) => -v));
        const cons2 = cons.map((k) => ({ a: k.a.concat(k.a.map((v) => -v)), op: k.op, b: k.b }));
        const r = ints.size ? milp(c2, cons2, [...ints].concat([...ints].map((i) => i + n))) : simplex(c2, cons2);
        res = r.status === 'ok' ? { status: 'ok', x: r.x.slice(0, n).map((v, i) => v - r.x[i + n]) } : r;
      } else {
        const r = ints.size ? milp(c, cons, [...ints]) : simplex(c, cons);
        res = r.status === 'ok' ? { status: 'ok', x: r.x } : r;
      }
      res.method = 'LP Simplex';
    } else {
      // não linear: penalidade + Nelder-Mead
      const proj = (x) => (model.nonneg === false ? x : x.map((v) => Math.max(0, v)));
      const F = (xx) => {
        const x = proj(xx);
        const f = fObj(x);
        if (!isNum(f)) return 1e30;
        let pen = 0;
        for (const k of rows) { const g = k.g(x); if (!isNum(g)) return 1e30; if (k.op === '<=' && g > 0) pen += g * g; else if (k.op === '>=' && g < 0) pen += g * g; else if (k.op === '=') pen += g * g; }
        return (model.goal === 'value' ? 0 : -sign * f) + 1e6 * pen;
      };
      let x = x0.slice();
      for (let round = 0; round < 4; round++) x = nelderMead(F, x);
      x = proj(x).map((v, i) => (ints.has(i) ? Math.round(v) : v));
      const feas = rows.every((k) => { const g = k.g(x); return k.op === '<=' ? g <= 1e-4 : k.op === '>=' ? g >= -1e-4 : Math.abs(g) <= 1e-4; });
      res = feas ? { status: 'ok', x } : { status: 'infeasible' };
      res.method = 'GRG Não Linear';
    }
    if (res.status === 'ok') res.x = res.x.map((v) => (Math.abs(v - Math.round(v)) < 1e-9 ? Math.round(v) : +v.toFixed(10)));
    res.vars = vars;
    res.x0 = x0;
    if (res.status === 'ok') res.value = fObj(res.x);
    return res;
  };
})();
