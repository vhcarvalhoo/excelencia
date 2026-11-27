/* Biblioteca de funções do Excel (nomes em português, com apelidos em inglês). */
(function () {
  'use strict';
  const XL = window.XL;
  const { Matrix, err, isErr, E, toNum, toStr, toBool, scalar, compareValues } = XL.Formula;

  const F = (XL.Functions = { map: new Map(), list: [] });
  F.get = (name) => F.map.get(XL.normName(name));
  F.search = (prefix) => {
    const p = XL.normName(prefix);
    return F.list.filter((d) => XL.normName(d.name).startsWith(p)).slice(0, 12);
  };
  F.CATEGORIES = ['Matemática e Trigonometria', 'Estatística', 'Lógica', 'Texto', 'Data e Hora', 'Pesquisa e Referência', 'Financeira'];
  F.byCategory = (cat) => F.list.filter((d) => d.cat === cat).sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));

  function def(name, en, cat, argSpec, desc, fn) {
    const args = argSpec ? argSpec.split(';').map((s) => s.trim()) : [];
    const min = args.filter((a) => !a.startsWith('[') && a !== '...').length;
    const max = args.includes('...') ? Infinity : args.length;
    const d = { name, en, cat, args, desc, fn, min, max, syntax: `${name}(${args.join('; ')})` };
    F.map.set(XL.normName(name), d);
    if (en) F.map.set(en.toUpperCase(), d);
    F.list.push(d);
  }

  // ---------- Auxiliares (lançam o erro para o motor tratar) ----------
  const S = (v) => scalar(v);
  const N = (v) => { const n = toNum(v); if (isErr(n)) throw n; return n; };
  const T = (v) => { const s = S(v); if (isErr(s)) throw s; const t = toStr(s); if (isErr(t)) throw t; return t; };
  const B = (v) => { const b = toBool(S(v)); if (isErr(b)) throw b; return b; };
  const has = (v) => v !== undefined && v !== null && !(v instanceof Matrix && v.h === 1 && v.w === 1 && v.rows[0][0] === null);
  const M = (v) => (v instanceof Matrix ? v : new Matrix([[v]]));

  // Números de vários argumentos: de intervalos só conta números; valores diretos são convertidos.
  function nums(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Matrix) {
        for (const v of a.flat()) {
          if (isErr(v)) throw v;
          if (typeof v === 'number') out.push(v);
        }
      } else if (a !== null && a !== undefined) {
        if (isErr(a)) throw a;
        out.push(N(a));
      }
    }
    return out;
  }
  function allValues(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Matrix) out.push(...a.flat());
      else out.push(a);
    }
    return out;
  }

  // Critérios no estilo CONT.SE: ">10", "<>Pago", "Jo*", 5 ...
  function wildcardRe(s) {
    let re = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '~' && (s[i + 1] === '*' || s[i + 1] === '?')) { re += '\\' + s[++i]; continue; }
      if (ch === '*') re += '[\\s\\S]*';
      else if (ch === '?') re += '[\\s\\S]';
      else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }
  function makeCriteria(c) {
    c = S(c);
    if (isErr(c)) throw c;
    if (typeof c === 'number') return (v) => typeof v === 'number' && v === c;
    if (typeof c === 'boolean') return (v) => v === c;
    if (c === null || c === undefined) return (v) => v === null || v === '';
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(String(c));
    const op = m[1] || '=';
    const rhs = m[2];
    const p = rhs !== '' ? XL.parseInput(rhs) : { kind: 'empty' };
    if (p.kind === 'number') {
      const n = p.value;
      const cmp = { '=': (v) => v === n, '<>': (v) => v !== n, '<': (v) => v < n, '>': (v) => v > n, '<=': (v) => v <= n, '>=': (v) => v >= n }[op];
      return (v) => (typeof v === 'number' ? cmp(v) : op === '<>');
    }
    if (op === '=' || op === '<>') {
      let eq;
      if (rhs === '') eq = (v) => v === null || v === '';
      else {
        const re = wildcardRe(rhs);
        eq = (v) => typeof v === 'string' && re.test(v);
      }
      return op === '=' ? eq : (v) => !eq(v);
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const c2 = v.localeCompare(rhs, 'pt-BR', { sensitivity: 'accent' });
      return { '<': c2 < 0, '>': c2 > 0, '<=': c2 <= 0, '>=': c2 >= 0 }[op];
    };
  }
  function lookupEq(a, b) {
    if (typeof a === 'string' && typeof b === 'string') {
      return /[*?]/.test(b) ? wildcardRe(b).test(a) : a.localeCompare(b, 'pt-BR', { sensitivity: 'accent' }) === 0;
    }
    return a === b;
  }
  const roundHalfAway = (x, d) => {
    const f = Math.pow(10, d);
    return (Math.sign(x) * Math.round(Math.abs(x) * f * (1 + Number.EPSILON))) / f;
  };

  // ================= Matemática =================
  const MAT = 'Matemática e Trigonometria';
  def('SOMA', 'SUM', MAT, 'núm1; [núm2]; ...', 'Soma todos os números de um intervalo de células.', (a) => nums(a).reduce((s, x) => s + x, 0));
  def('SUBTOTAL', 'SUBTOTAL', MAT, 'núm_função; ref1; [ref2]; ...', 'Calcula sobre as linhas VISÍVEIS: 9 = soma, 1 = média, 2 = contar números, 3 = contar valores, 4 = máximo, 5 = mínimo. Ideal com filtros.', (a) => {
    const code = Math.trunc(N(a[0])) % 100;
    const names = { 1: 'MÉDIA', 2: 'CONT.NÚM', 3: 'CONT.VALORES', 4: 'MÁXIMO', 5: 'MÍNIMO', 9: 'SOMA' };
    if (code === 6) {
      // produto
      const vals = nums(a.slice(1).map(visibleOnly));
      return vals.reduce((p, x) => p * x, 1);
    }
    const fn = names[code] && F.get(names[code]);
    if (!fn) throw err(E.VALUE);
    return fn.fn(a.slice(1).map(visibleOnly));
  });
  function visibleOnly(m) {
    if (!(m instanceof Matrix) || !m.hiddenRows) return m;
    const rows = m.rows.filter((_, i) => !m.hiddenRows[i]);
    return new Matrix(rows.length ? rows : [[null]]);
  }
  def('SOMARPRODUTO', 'SUMPRODUCT', MAT, 'matriz1; [matriz2]; ...', 'Multiplica as matrizes correspondentes e retorna a soma dos produtos.', (a) => {
    const ms = a.map(M);
    const h = ms[0].h, w = ms[0].w;
    if (ms.some((m) => m.h !== h || m.w !== w)) throw err(E.VALUE);
    let s = 0;
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
      let p = 1;
      for (const m of ms) { const v = m.rows[i][j]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : 0; }
      s += p;
    }
    return s;
  });
  def('ARRED', 'ROUND', MAT, 'núm; núm_dígitos', 'Arredonda um número até uma quantidade especificada de dígitos.', (a) => roundHalfAway(N(a[0]), Math.trunc(N(a[1]))));
  def('ARREDONDAR.PARA.CIMA', 'ROUNDUP', MAT, 'núm; núm_dígitos', 'Arredonda um número para cima, afastando-o de zero.', (a) => {
    const x = N(a[0]), f = Math.pow(10, Math.trunc(N(a[1])));
    return (Math.sign(x) * Math.ceil(Math.abs(x) * f - 1e-9)) / f;
  });
  def('ARREDONDAR.PARA.BAIXO', 'ROUNDDOWN', MAT, 'núm; núm_dígitos', 'Arredonda um número para baixo, aproximando-o de zero.', (a) => {
    const x = N(a[0]), f = Math.pow(10, Math.trunc(N(a[1])));
    return (Math.sign(x) * Math.floor(Math.abs(x) * f + 1e-9)) / f;
  });
  def('INT', 'INT', MAT, 'núm', 'Arredonda um número para baixo até o número inteiro mais próximo.', (a) => Math.floor(N(a[0])));
  def('ABS', 'ABS', MAT, 'núm', 'Retorna o valor absoluto de um número (sem sinal).', (a) => Math.abs(N(a[0])));
  def('RAIZ', 'SQRT', MAT, 'núm', 'Retorna a raiz quadrada de um número.', (a) => { const x = N(a[0]); if (x < 0) throw err(E.NUM); return Math.sqrt(x); });
  def('POTÊNCIA', 'POWER', MAT, 'núm; potência', 'Retorna o resultado de um número elevado a uma potência.', (a) => { const r = Math.pow(N(a[0]), N(a[1])); if (!isFinite(r)) throw err(E.NUM); return r; });
  def('MOD', 'MOD', MAT, 'núm; divisor', 'Retorna o resto da divisão.', (a) => { const n = N(a[0]), d = N(a[1]); if (d === 0) throw err(E.DIV0); return n - d * Math.floor(n / d); });
  def('PI', 'PI', MAT, '', 'Retorna o valor de Pi (3,14159...).', () => Math.PI);
  def('ALEATÓRIO', 'RAND', MAT, '', 'Retorna um número aleatório entre 0 e 1.', () => Math.random());
  def('ALEATÓRIOENTRE', 'RANDBETWEEN', MAT, 'inferior; superior', 'Retorna um número inteiro aleatório entre os números especificados.', (a) => {
    const lo = Math.ceil(N(a[0])), hi = Math.floor(N(a[1]));
    if (lo > hi) throw err(E.NUM);
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  });
  def('SOMASE', 'SUMIF', MAT, 'intervalo; critérios; [intervalo_soma]', 'Soma as células que atendem a um critério.', (a) => {
    const rng = M(a[0]), test = makeCriteria(a[1]), sum = has(a[2]) ? M(a[2]) : rng;
    let s = 0;
    for (let i = 0; i < rng.h; i++) for (let j = 0; j < rng.w; j++) {
      if (!test(rng.rows[i][j])) continue;
      const v = sum.rows[i] ? sum.rows[i][j] : null;
      if (isErr(v)) throw v;
      if (typeof v === 'number') s += v;
    }
    return s;
  });
  def('SOMASES', 'SUMIFS', MAT, 'intervalo_soma; intervalo_critérios1; critérios1; ...', 'Soma as células que atendem a vários critérios.', (a) => {
    const sum = M(a[0]);
    const pairs = criteriaPairs(a.slice(1), sum);
    let s = 0;
    for (let i = 0; i < sum.h; i++) for (let j = 0; j < sum.w; j++) {
      if (!pairs.every(([r, t]) => t(r.rows[i][j]))) continue;
      const v = sum.rows[i][j];
      if (isErr(v)) throw v;
      if (typeof v === 'number') s += v;
    }
    return s;
  });
  function criteriaPairs(list, shape) {
    if (list.length % 2 !== 0) throw err(E.VALUE);
    const out = [];
    for (let k = 0; k < list.length; k += 2) {
      const r = M(list[k]);
      if (shape && (r.h !== shape.h || r.w !== shape.w)) throw err(E.VALUE);
      out.push([r, makeCriteria(list[k + 1])]);
    }
    return out;
  }

  // ================= Estatística =================
  const EST = 'Estatística';
  def('MÉDIA', 'AVERAGE', EST, 'núm1; [núm2]; ...', 'Retorna a média aritmética dos argumentos.', (a) => {
    const n = nums(a);
    if (!n.length) throw err(E.DIV0);
    return n.reduce((s, x) => s + x, 0) / n.length;
  });
  def('MÁXIMO', 'MAX', EST, 'núm1; [núm2]; ...', 'Retorna o maior valor de um conjunto de valores.', (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; });
  def('MÍNIMO', 'MIN', EST, 'núm1; [núm2]; ...', 'Retorna o menor valor de um conjunto de valores.', (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; });
  def('MED', 'MEDIAN', EST, 'núm1; [núm2]; ...', 'Retorna a mediana (o número central) dos números indicados.', (a) => {
    const n = nums(a).sort((x, y) => x - y);
    if (!n.length) throw err(E.NUM);
    const m = Math.floor(n.length / 2);
    return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
  });
  def('MAIOR', 'LARGE', EST, 'matriz; k', 'Retorna o k-ésimo maior valor de um conjunto de dados.', (a) => {
    const n = nums([a[0]]).sort((x, y) => y - x), k = Math.trunc(N(a[1]));
    if (k < 1 || k > n.length) throw err(E.NUM);
    return n[k - 1];
  });
  def('MENOR', 'SMALL', EST, 'matriz; k', 'Retorna o k-ésimo menor valor de um conjunto de dados.', (a) => {
    const n = nums([a[0]]).sort((x, y) => x - y), k = Math.trunc(N(a[1]));
    if (k < 1 || k > n.length) throw err(E.NUM);
    return n[k - 1];
  });
  def('CONT.NÚM', 'COUNT', EST, 'valor1; [valor2]; ...', 'Conta quantas células contêm números.', (a) => {
    let c = 0;
    for (const x of a) {
      if (x instanceof Matrix) c += x.flat().filter((v) => typeof v === 'number').length;
      else if (x !== null && !isErr(toNum(x)) && typeof x !== 'boolean') c++;
    }
    return c;
  });
  def('CONT.VALORES', 'COUNTA', EST, 'valor1; [valor2]; ...', 'Conta quantas células não estão vazias.', (a) =>
    allValues(a).filter((v) => v !== null && v !== undefined).length);
  def('CONTAR.VAZIO', 'COUNTBLANK', EST, 'intervalo', 'Conta quantas células estão vazias em um intervalo.', (a) =>
    M(a[0]).flat().filter((v) => v === null || v === '').length);
  def('CONT.SE', 'COUNTIF', EST, 'intervalo; critérios', 'Conta as células de um intervalo que atendem a um critério.', (a) => {
    const t = makeCriteria(a[1]);
    return M(a[0]).flat().filter(t).length;
  });
  def('CONT.SES', 'COUNTIFS', EST, 'intervalo_critérios1; critérios1; ...', 'Conta as células que atendem a vários critérios.', (a) => {
    const pairs = criteriaPairs(a, M(a[0]));
    const base = pairs[0][0];
    let c = 0;
    for (let i = 0; i < base.h; i++) for (let j = 0; j < base.w; j++) if (pairs.every(([r, t]) => t(r.rows[i][j]))) c++;
    return c;
  });
  def('MÉDIASE', 'AVERAGEIF', EST, 'intervalo; critérios; [intervalo_média]', 'Retorna a média das células que atendem a um critério.', (a) => {
    const rng = M(a[0]), test = makeCriteria(a[1]), avg = has(a[2]) ? M(a[2]) : rng;
    let s = 0, c = 0;
    for (let i = 0; i < rng.h; i++) for (let j = 0; j < rng.w; j++) {
      if (!test(rng.rows[i][j])) continue;
      const v = avg.rows[i] ? avg.rows[i][j] : null;
      if (typeof v === 'number') { s += v; c++; }
    }
    if (!c) throw err(E.DIV0);
    return s / c;
  });
  def('MÉDIASES', 'AVERAGEIFS', EST, 'intervalo_média; intervalo_critérios1; critérios1; ...', 'Retorna a média das células que atendem a vários critérios.', (a) => {
    const avg = M(a[0]);
    const pairs = criteriaPairs(a.slice(1), avg);
    let s = 0, c = 0;
    for (let i = 0; i < avg.h; i++) for (let j = 0; j < avg.w; j++) {
      if (!pairs.every(([r, t]) => t(r.rows[i][j]))) continue;
      const v = avg.rows[i][j];
      if (typeof v === 'number') { s += v; c++; }
    }
    if (!c) throw err(E.DIV0);
    return s / c;
  });

  // ================= Lógica =================
  const LOG = 'Lógica';
  def('SE', 'IF', LOG, 'teste_lógico; [valor_se_verdadeiro]; [valor_se_falso]', 'Verifica se uma condição é verdadeira e retorna um valor se for VERDADEIRO e outro se for FALSO.', (a) => {
    const c = B(a[0]);
    if (c) return a.length > 1 ? (a[1] === null ? 0 : S(a[1])) : true;
    return a.length > 2 ? (a[2] === null ? 0 : S(a[2])) : false;
  });
  def('SES', 'IFS', LOG, 'teste_lógico1; valor_se_verdadeiro1; ...', 'Verifica várias condições e retorna o valor da primeira que for VERDADEIRA.', (a) => {
    if (a.length % 2) throw err(E.VALUE);
    for (let i = 0; i < a.length; i += 2) if (B(a[i])) return S(a[i + 1]);
    throw err(E.NA);
  });
  const bools = (a) => {
    const out = [];
    for (const x of a) {
      if (x instanceof Matrix) {
        for (const v of x.flat()) { if (isErr(v)) throw v; if (typeof v === 'boolean' || typeof v === 'number') out.push(!!v); }
      } else out.push(B(x));
    }
    if (!out.length) throw err(E.VALUE);
    return out;
  };
  def('E', 'AND', LOG, 'lógico1; [lógico2]; ...', 'Retorna VERDADEIRO se todos os argumentos forem verdadeiros.', (a) => bools(a).every(Boolean));
  def('OU', 'OR', LOG, 'lógico1; [lógico2]; ...', 'Retorna VERDADEIRO se pelo menos um argumento for verdadeiro.', (a) => bools(a).some(Boolean));
  def('NÃO', 'NOT', LOG, 'lógico', 'Inverte o valor lógico: VERDADEIRO vira FALSO e vice-versa.', (a) => !B(a[0]));
  def('SEERRO', 'IFERROR', LOG, 'valor; valor_se_erro', 'Retorna um valor alternativo se a expressão resultar em erro.', (a) => {
    const v = S(a[0]);
    return isErr(v) ? S(a[1]) : v === null ? 0 : v;
  });
  def('SENÃODISP', 'IFNA', LOG, 'valor; valor_se_nd', 'Retorna um valor alternativo se a expressão resultar em #N/D.', (a) => {
    const v = S(a[0]);
    return isErr(v) && v.code === E.NA ? S(a[1]) : v;
  });
  def('VERDADEIRO', 'TRUE', LOG, '', 'Retorna o valor lógico VERDADEIRO.', () => true);
  def('FALSO', 'FALSE', LOG, '', 'Retorna o valor lógico FALSO.', () => false);

  // ================= Texto =================
  const TXT = 'Texto';
  def('CONCATENAR', 'CONCATENATE', TXT, 'texto1; [texto2]; ...', 'Junta vários textos em um só.', (a) => a.map(T).join(''));
  def('CONCAT', null, TXT, 'texto1; [texto2]; ...', 'Junta textos e intervalos em um só texto.', (a) => allValues(a).map((v) => { if (isErr(v)) throw v; return toStr(v); }).join(''));
  def('UNIRTEXTO', 'TEXTJOIN', TXT, 'delimitador; ignorar_vazio; texto1; ...', 'Junta textos usando um delimitador entre eles.', (a) => {
    const d = T(a[0]), ign = B(a[1]);
    return allValues(a.slice(2)).map((v) => { if (isErr(v)) throw v; return toStr(v); }).filter((s) => !ign || s !== '').join(d);
  });
  def('ESQUERDA', 'LEFT', TXT, 'texto; [núm_caract]', 'Retorna os primeiros caracteres de um texto.', (a) => {
    const n = has(a[1]) ? N(a[1]) : 1;
    if (n < 0) throw err(E.VALUE);
    return T(a[0]).slice(0, n);
  });
  def('DIREITA', 'RIGHT', TXT, 'texto; [núm_caract]', 'Retorna os últimos caracteres de um texto.', (a) => {
    const n = has(a[1]) ? N(a[1]) : 1;
    if (n < 0) throw err(E.VALUE);
    const s = T(a[0]);
    return n === 0 ? '' : s.slice(-n);
  });
  def('EXT.TEXTO', 'MID', TXT, 'texto; núm_inicial; núm_caract', 'Retorna caracteres do meio de um texto, a partir de uma posição.', (a) => {
    const st = N(a[1]), n = N(a[2]);
    if (st < 1 || n < 0) throw err(E.VALUE);
    return T(a[0]).substr(st - 1, n);
  });
  def('NÚM.CARACT', 'LEN', TXT, 'texto', 'Retorna a quantidade de caracteres de um texto.', (a) => T(a[0]).length);
  def('MAIÚSCULA', 'UPPER', TXT, 'texto', 'Converte o texto em letras maiúsculas.', (a) => T(a[0]).toLocaleUpperCase('pt-BR'));
  def('MINÚSCULA', 'LOWER', TXT, 'texto', 'Converte o texto em letras minúsculas.', (a) => T(a[0]).toLocaleLowerCase('pt-BR'));
  def('PRI.MAIÚSCULA', 'PROPER', TXT, 'texto', 'Coloca a primeira letra de cada palavra em maiúscula.', (a) =>
    T(a[0]).toLocaleLowerCase('pt-BR').replace(/(^|[^A-Za-zÀ-ÿ])([A-Za-zÀ-ÿ])/g, (_, p, c) => p + c.toLocaleUpperCase('pt-BR')));
  def('ARRUMAR', 'TRIM', TXT, 'texto', 'Remove os espaços extras do texto, deixando só um espaço entre as palavras.', (a) => T(a[0]).replace(/ +/g, ' ').trim());
  def('SUBSTITUIR', 'SUBSTITUTE', TXT, 'texto; texto_antigo; novo_texto; [núm_da_ocorrência]', 'Substitui um texto antigo por um novo texto.', (a) => {
    const s = T(a[0]), old = T(a[1]), nw = T(a[2]);
    if (!old) return s;
    if (!has(a[3])) return s.split(old).join(nw);
    const k = N(a[3]);
    let idx = -1;
    for (let i = 0; i < k; i++) { idx = s.indexOf(old, idx + 1); if (idx < 0) return s; }
    return s.slice(0, idx) + nw + s.slice(idx + old.length);
  });
  def('PROCURAR', 'FIND', TXT, 'texto_procurado; no_texto; [núm_inicial]', 'Retorna a posição de um texto dentro de outro (diferencia maiúsculas).', (a) => {
    const i = T(a[1]).indexOf(T(a[0]), (has(a[2]) ? N(a[2]) : 1) - 1);
    if (i < 0) throw err(E.VALUE);
    return i + 1;
  });
  def('VALOR', 'VALUE', TXT, 'texto', 'Converte um texto que representa um número em número.', (a) => N(T(a[0])));
  def('TEXTO', 'TEXT', TXT, 'valor; formato_texto', 'Converte um número ou data em texto com o formato indicado, como "dd/mm/aaaa", "mmmm" ou "R$ #.##0,00".', (a) => {
    const v = S(a[0]);
    if (isErr(v)) throw v;
    const code = T(a[1]);
    if (typeof v === 'string') { const n = XL.parseInput(v); if (n.kind !== 'number') return v; return XL.formatCode(n.value, code); }
    return XL.formatCode(N(v), code);
  });
  def('LOCALIZAR', 'SEARCH', TXT, 'texto_procurado; no_texto; [núm_inicial]', 'Retorna a posição de um texto dentro de outro (não diferencia maiúsculas e aceita curingas).', (a) => {
    const find = T(a[0]), within = T(a[1]);
    const st = has(a[2]) ? N(a[2]) : 1;
    const re = wildcardRe(find);
    const src = re.source.slice(1, -1); // sem ^ e $
    const m = new RegExp(src, 'i').exec(within.slice(st - 1));
    if (!m) throw err(E.VALUE);
    return m.index + st;
  });
  const textPart = (before) => (a) => {
    const s = T(a[0]), d = T(a[1]);
    let n = has(a[2]) ? Math.trunc(N(a[2])) : 1;
    if (!d || n === 0) throw err(E.VALUE);
    let idx = -1;
    if (n > 0) { for (let i = 0; i < n; i++) { idx = s.indexOf(d, idx + 1); if (idx < 0) break; } }
    else { idx = s.length; for (let i = 0; i < -n; i++) { idx = s.lastIndexOf(d, idx - 1); if (idx < 0) break; } }
    if (idx < 0) throw err(E.NA);
    return before ? s.slice(0, idx) : s.slice(idx + d.length);
  };
  def('TEXTOANTES', 'TEXTBEFORE', TXT, 'texto; delimitador; [núm_instância]', 'Retorna o texto que vem antes do delimitador (ex.: o usuário antes do @ de um e-mail).', textPart(true));
  def('TEXTODEPOIS', 'TEXTAFTER', TXT, 'texto; delimitador; [núm_instância]', 'Retorna o texto que vem depois do delimitador (ex.: o domínio depois do @).', textPart(false));
  def('REPT', 'REPT', TXT, 'texto; núm_vezes', 'Repete um texto a quantidade de vezes indicada.', (a) => { const n = Math.trunc(N(a[1])); if (n < 0) throw err(E.VALUE); return T(a[0]).repeat(n); });

  // ================= Data e Hora =================
  const DT = 'Data e Hora';
  def('HOJE', 'TODAY', DT, '', 'Retorna a data de hoje.', () => XL.todaySerial());
  def('AGORA', 'NOW', DT, '', 'Retorna a data e a hora atuais.', () => {
    const n = new Date();
    return XL.todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
  });
  def('DATA', 'DATE', DT, 'ano; mês; dia', 'Monta uma data a partir do ano, mês e dia.', (a) => {
    let y = Math.trunc(N(a[0]));
    const m = Math.trunc(N(a[1])), d = Math.trunc(N(a[2]));
    if (y < 1900) y += 1900;
    const serial = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    if (!isFinite(serial) || serial < 0) throw err(E.NUM);
    return serial;
  });
  def('DIA', 'DAY', DT, 'núm_série', 'Retorna o dia de uma data (1 a 31).', (a) => XL.serialToParts(N(a[0])).d);
  def('MÊS', 'MONTH', DT, 'núm_série', 'Retorna o mês de uma data (1 a 12).', (a) => XL.serialToParts(N(a[0])).m);
  def('ANO', 'YEAR', DT, 'núm_série', 'Retorna o ano de uma data.', (a) => XL.serialToParts(N(a[0])).y);
  def('DIA.DA.SEMANA', 'WEEKDAY', DT, 'núm_série; [tipo_retorno]', 'Retorna o dia da semana (1 = domingo, por padrão).', (a) => {
    const wd = XL.serialToParts(N(a[0])).wd;
    const t = has(a[1]) ? N(a[1]) : 1;
    if (t === 2) return wd === 0 ? 7 : wd;
    if (t === 3) return wd === 0 ? 6 : wd - 1;
    return wd + 1;
  });
  def('DIAS', 'DAYS', DT, 'data_final; data_inicial', 'Retorna o número de dias entre duas datas.', (a) => Math.trunc(N(a[0])) - Math.trunc(N(a[1])));
  // dias úteis: segunda a sexta, sem os feriados informados
  const holidays = (v) => new Set(has(v) ? nums([v]).map(Math.trunc) : []);
  const isWorkday = (s, hol) => { const wd = XL.serialToParts(s).wd; return wd !== 0 && wd !== 6 && !hol.has(s); };
  def('DIATRABALHO', 'WORKDAY', DT, 'data_inicial; dias; [feriados]', 'Retorna a data depois de um número de dias úteis (pula fins de semana e feriados).', (a) => {
    let s = Math.trunc(N(a[0]));
    let n = Math.trunc(N(a[1]));
    const hol = holidays(a[2]);
    const step = n >= 0 ? 1 : -1;
    while (n !== 0) { s += step; if (isWorkday(s, hol)) n -= step; }
    return s;
  });
  def('DIATRABALHOTOTAL', 'NETWORKDAYS', DT, 'data_inicial; data_final; [feriados]', 'Conta os dias úteis entre duas datas (incluindo as duas).', (a) => {
    const s1 = Math.trunc(N(a[0])), s2 = Math.trunc(N(a[1]));
    const hol = holidays(a[2]);
    const lo = Math.min(s1, s2), hi = Math.max(s1, s2);
    let c = 0;
    for (let s = lo; s <= hi; s++) if (isWorkday(s, hol)) c++;
    return s1 <= s2 ? c : -c;
  });
  const addMonths = (serial, months, endOfMonth) => {
    const p = XL.serialToParts(serial);
    const mIdx = p.m - 1 + months;
    const last = new Date(Date.UTC(p.y, mIdx + 1, 0)).getUTCDate();
    const day = endOfMonth ? last : Math.min(p.d, last);
    return Math.round((Date.UTC(p.y, mIdx, day) - Date.UTC(1899, 11, 30)) / 86400000);
  };
  def('DATAM', 'EDATE', DT, 'data_inicial; meses', 'Retorna a data que fica alguns meses antes ou depois (ex.: vencimentos mensais).', (a) => addMonths(Math.trunc(N(a[0])), Math.trunc(N(a[1])), false));
  def('FIMMÊS', 'EOMONTH', DT, 'data_inicial; meses', 'Retorna o último dia do mês, alguns meses antes ou depois.', (a) => addMonths(Math.trunc(N(a[0])), Math.trunc(N(a[1])), true));
  def('DATADIF', 'DATEDIF', DT, 'data_inicial; data_final; unidade', 'Diferença entre datas em anos ("a"), meses ("m") ou dias ("d"). Também "md", "am", "ad".', (a) => {
    const s1 = Math.trunc(N(a[0])), s2 = Math.trunc(N(a[1]));
    const u = XL.normName(T(a[2])).replace('Y', 'A');
    if (s1 > s2) throw err(E.NUM);
    const p1 = XL.serialToParts(s1), p2 = XL.serialToParts(s2);
    let months = (p2.y - p1.y) * 12 + (p2.m - p1.m);
    if (p2.d < p1.d) months--;
    switch (u) {
      case 'D': return s2 - s1;
      case 'M': return months;
      case 'A': return Math.floor(months / 12);
      case 'AM': case 'YM': return months % 12;
      case 'MD': { const base = addMonths(s1, months, false); return s2 - base; }
      case 'AD': case 'YD': { const yrs = Math.floor(months / 12); return s2 - addMonths(s1, yrs * 12, false); }
    }
    throw err(E.NUM);
  });

  // ================= Pesquisa e Referência =================
  const PR = 'Pesquisa e Referência';
  def('PROCV', 'VLOOKUP', PR, 'valor_procurado; matriz_tabela; núm_índice_coluna; [procurar_intervalo]', 'Procura um valor na primeira coluna de uma tabela e retorna um valor da mesma linha, na coluna indicada.', (a) => {
    const v = S(a[0]);
    if (isErr(v)) throw v;
    const tb = M(a[1]);
    const col = Math.trunc(N(a[2]));
    if (col < 1) throw err(E.VALUE);
    if (col > tb.w) throw err(E.REF);
    const approx = has(a[3]) ? B(a[3]) : true;
    if (!approx) {
      for (const row of tb.rows) if (lookupEq(row[0], v)) return row[col - 1] === null ? 0 : row[col - 1];
      throw err(E.NA);
    }
    let found = -1;
    for (let i = 0; i < tb.h; i++) {
      const x = tb.rows[i][0];
      if (x === null) continue;
      if (compareValues(x, v) <= 0) found = i;
      else break;
    }
    if (found < 0) throw err(E.NA);
    const r = tb.rows[found][col - 1];
    return r === null ? 0 : r;
  });
  def('PROCX', 'XLOOKUP', PR, 'pesquisa_valor; pesquisa_matriz; matriz_retorno; [se_não_encontrada]', 'Procura um valor em um intervalo e retorna o item correspondente de outro intervalo.', (a) => {
    const v = S(a[0]);
    if (isErr(v)) throw v;
    const look = M(a[1]), ret = M(a[2]);
    const vertical = look.w === 1;
    const list = vertical ? look.rows.map((r) => r[0]) : look.rows[0];
    for (let i = 0; i < list.length; i++) {
      if (!lookupEq(list[i], v)) continue;
      if (vertical) {
        const row = ret.rows[i];
        if (!row) throw err(E.VALUE);
        return row.length === 1 ? row[0] : new Matrix([row]);
      }
      const colv = ret.rows.map((r) => r[i]);
      return colv.length === 1 ? colv[0] : new Matrix(colv.map((x) => [x]));
    }
    if (has(a[3])) return S(a[3]);
    throw err(E.NA);
  });
  def('ÍNDICE', 'INDEX', PR, 'matriz; núm_linha; [núm_coluna]', 'Retorna o valor de uma célula numa posição (linha e coluna) de um intervalo.', (a) => {
    const m = M(a[0]);
    let r = Math.trunc(N(a[1]));
    let c = has(a[2]) ? Math.trunc(N(a[2])) : 0;
    if (m.h === 1 && !has(a[2])) { c = r; r = 1; }
    if (r < 0 || c < 0 || r > m.h || c > m.w) throw err(E.REF);
    if (r === 0 && c === 0) return m;
    if (r === 0) return new Matrix(m.rows.map((row) => [row[c - 1]]));
    if (c === 0) return m.w === 1 ? m.rows[r - 1][0] : new Matrix([m.rows[r - 1]]);
    const v = m.rows[r - 1][c - 1];
    return v === null ? 0 : v;
  });
  def('CORRESP', 'MATCH', PR, 'valor_procurado; matriz_procurada; [tipo_correspondência]', 'Retorna a posição de um valor dentro de um intervalo.', (a) => {
    const v = S(a[0]);
    if (isErr(v)) throw v;
    const m = M(a[1]);
    const list = m.w === 1 ? m.rows.map((r) => r[0]) : m.rows[0];
    const type = has(a[2]) ? N(a[2]) : 1;
    if (type === 0) {
      const i = list.findIndex((x) => lookupEq(x, v));
      if (i < 0) throw err(E.NA);
      return i + 1;
    }
    let found = -1;
    for (let i = 0; i < list.length; i++) {
      if (list[i] === null) continue;
      const c = compareValues(list[i], v);
      if (type > 0 ? c <= 0 : c >= 0) found = i;
      else break;
    }
    if (found < 0) throw err(E.NA);
    return found + 1;
  });

  // ================= Financeira =================
  const FIN = 'Financeira';
  def('PGTO', 'PMT', FIN, 'taxa; nper; vp; [vf]; [tipo]', 'Calcula a parcela de um empréstimo com pagamentos e juros constantes.', (a) => {
    const r = N(a[0]), n = N(a[1]), pv = N(a[2]), fv = has(a[3]) ? N(a[3]) : 0, t = has(a[4]) ? N(a[4]) : 0;
    if (n === 0) throw err(E.NUM);
    if (r === 0) return -(pv + fv) / n;
    const q = Math.pow(1 + r, n);
    return -(r * (pv * q + fv)) / ((1 + r * t) * (q - 1));
  });
  def('VF', 'FV', FIN, 'taxa; nper; pgto; [vp]; [tipo]', 'Calcula o valor futuro de um investimento.', (a) => {
    const r = N(a[0]), n = N(a[1]), pmt = N(a[2]), pv = has(a[3]) ? N(a[3]) : 0, t = has(a[4]) ? N(a[4]) : 0;
    if (r === 0) return -(pv + pmt * n);
    const q = Math.pow(1 + r, n);
    return -(pv * q + (pmt * (1 + r * t) * (q - 1)) / r);
  });
  def('VP', 'PV', FIN, 'taxa; nper; pgto; [vf]; [tipo]', 'Calcula o valor presente de um investimento.', (a) => {
    const r = N(a[0]), n = N(a[1]), pmt = N(a[2]), fv = has(a[3]) ? N(a[3]) : 0, t = has(a[4]) ? N(a[4]) : 0;
    if (r === 0) return -(fv + pmt * n);
    const q = Math.pow(1 + r, n);
    return -(fv + (pmt * (1 + r * t) * (q - 1)) / r) / q;
  });

  F.list.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
})();
