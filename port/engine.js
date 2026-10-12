// isoharness preview port: the ea_irms 1.4.2 steps that depend on a user's choices
// (blank, linearity, drift, calibration, elemental content, uncertainty), in
// JavaScript so the page can redraw at once. The engine on the server stays the
// only source of saved or downloaded numbers; continuous integration holds this
// port to it (scripts/port_cases.py, port/test_port.js).
//
// runModel(D, dec) takes the "port" block of POST /api/process and the user's
// decisions. The configuration's settings that act inside these steps (the
// significance level, the monitor outlier limit, the check-standard offset, the
// blank multiple, k, the elemental gate and the first drift position) come with
// the block as D.settings, and each position says whether the server set its CO2
// peak aside as a reference pulse (rej);
// a block without them gets the package defaults. Settings that act before the
// block is built (the standards, the reference-pulse windows, the size limits,
// settling and carryover) are applied by the server, so a change to them goes
// back to the server.
//
// The methods below are the prototype's (logic.js lines 43 to 378), changed only
// to read D.settings in place of the fixed defaults (settings()).
class PortEngine {
  lgam(z) {
    const g = 7, c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028,
      771.32342877765313, -176.61502916214059, 12.507343278686905,
      -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    if (z < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * z))) - this.lgam(1 - z);
    z -= 1;
    let a = c[0];
    const t = z + g + 0.5;
    for (let i = 1; i < g + 2; i++) a += c[i] / (z + i);
    return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
  }

  betacf(a, b, x) {
    const tiny = 1e-300, qab = a + b, qap = a + 1, qam = a - 1;
    let c = 1, d = 1 - qab * x / qap;
    if (Math.abs(d) < tiny) d = tiny;
    d = 1 / d;
    let h = d;
    for (let m = 1; m < 300; m++) {
      const m2 = 2 * m;
      let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
      c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
      d = 1 / d; h *= d * c;
      aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
      c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
      d = 1 / d;
      const de = d * c;
      h *= de;
      if (Math.abs(de - 1) < 3e-16) break;
    }
    return h;
  }

  betainc(a, b, x) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    const lb = this.lgam(a + b) - this.lgam(a) - this.lgam(b);
    const front = Math.exp(lb + a * Math.log(x) + b * Math.log(1 - x));
    if (x < (a + 1) / (a + b + 2)) return front * this.betacf(a, b, x) / a;
    return 1 - front * this.betacf(b, a, 1 - x) / b;
  }

  tp(t, df) {
    if (df == null || df < 1 || t == null) return null;
    return this.betainc(df / 2, 0.5, df / (df + t * t));
  }

  fit(x, y) {
    const pr = [];
    for (let i = 0; i < x.length; i++) if (x[i] != null && y[i] != null) pr.push([x[i], y[i]]);
    const n = pr.length;
    if (n < 2) return null;
    let mx = 0, my = 0;
    for (const [a, b] of pr) { mx += a; my += b; }
    mx /= n; my /= n;
    let sxx = 0, sxy = 0, syy = 0;
    for (const [a] of pr) sxx += (a - mx) ** 2;
    if (sxx === 0) return null;
    for (const [a, b] of pr) sxy += (a - mx) * (b - my);
    const slope = sxy / sxx, inter = my - slope * mx;
    for (const [, b] of pr) syy += (b - my) ** 2;
    let res = 0;
    for (const [a, b] of pr) res += (b - (slope * a + inter)) ** 2;
    const r2 = syy > 0 ? 1 - res / syy : null;
    const sd = n > 2 ? Math.sqrt(res / (n - 2)) : null;
    const se = sd != null ? sd / Math.sqrt(sxx) : null;
    const t = se ? slope / se : null;
    const p = t != null ? this.tp(t, n - 2) : null;
    return { slope, intercept: inter, r2, sd, n, se, t, p, df: n - 2, mean_x: mx, sxx, sse: res };
  }

  sdev(v) {
    const a = v.filter((q) => q != null);
    if (a.length < 2) return null;
    const m = a.reduce((s, q) => s + q, 0) / a.length;
    return Math.sqrt(a.reduce((s, q) => s + (q - m) ** 2, 0) / (a.length - 1));
  }

  above(a, b) { return a - b > 1e-9 * Math.max(1, Math.abs(b)); }

  applies(mode, f, n, alpha) {
    if (mode === 'Off' || f == null || n < 3) return false;
    if (mode === 'On') return true;
    return f.p != null && f.p < alpha;
  }

  stud(f, x, y, inc) {
    if (f == null || x == null || y == null) return null;
    const n = f.n, e = y - (f.slope * x + f.intercept);
    const h = 1 / n + (x - f.mean_x) ** 2 / f.sxx;
    if (inc) {
      if (n < 5 || h >= 1) return null;
      const s2 = (f.sse - e * e / (1 - h)) / (n - 3);
      if (s2 <= 0) return null;
      return e / (Math.sqrt(s2) * Math.sqrt(1 - h));
    }
    if (f.sd == null || f.sd <= 0) return null;
    return e / (f.sd * Math.sqrt(1 + h));
  }

  // The settings runModel reads, as the API names them in D.settings; missing
  // keys keep the package defaults, and an explicit null blankMultiple turns that
  // flag off, as None does in ea_irms.
  settings(D) {
    const S = { alpha: 0.05, outlierT: 4.0, checkOffset: 0.5, blankMultiple: 3.0, k: 2.0, elemGate: 5.0,
      firstDrift: null };
    const given = D && D.settings ? D.settings : {};
    for (const key of Object.keys(S)) if (key in given) S[key] = given[key];
    return S;
  }

  runModel(D, dec) {
    const S = this.settings(D);
    const alpha = S.alpha, std = D.std;
    const P = D.pos.map((q) => {
      const s = q.si == null ? null : std[q.si];
      return Object.assign({}, q, {
        acc15: s ? s[3] : null, acc13: s ? s[2] : null, certN: s ? s[5] : null, certC: s ? s[4] : null,
        uacc15: s ? s[8] : null, uacc13: s ? s[7] : null, n: {}, c: {},
      });
    });
    const role = (r) => P.filter((q) => q.role === r);
    const M = { dec, settings: S, blank: {}, lin: {}, drift: {}, cal: {}, unc: {}, elem: {}, P };
    const E = [['n', 'aN', 'dN'], ['c', 'aC', 'dC']];
    for (const [el, ak, dk] of E) {
      const b = role('B').filter((q) => q[ak] != null);
      const blk = {
        n: b.length,
        area: b.length ? b.reduce((s, q) => s + q[ak], 0) / b.length : 0,
        d: b.length ? b.reduce((s, q) => s + q[dk], 0) / b.length : 0,
        sd: this.sdev(b.map((q) => q[dk])),
      };
      blk.inForce = dec.blank === 'Off' || blk.n < 1 ? false
        : dec.blank === 'On' ? true : blk.n >= 3 && blk.sd != null && blk.sd <= 2.0;
      M.blank[el] = blk;
      for (const q of P) {
        let v = q[dk];
        if (v != null && blk.inForce && q.role !== 'B' && q[ak] > blk.area) {
          v = (q[ak] * q[dk] - blk.area * blk.d) / (q[ak] - blk.area);
        }
        q[el].raw = q[dk];
        q[el].blank = v;
      }
    }
    const lin = role('L');
    for (const [el, ak] of E) {
      const xs = lin.map((q) => q[ak]), ys = lin.map((q) => q[el].blank);
      const f = this.fit(xs, ys);
      const good = xs.filter((v) => v != null);
      const mode = el === 'n' ? dec.linN : dec.linC;
      const L = {
        fit: f, pivot: good.length ? good.reduce((s, v) => s + v, 0) / good.length : 0,
        n: good.length, min: good.length ? Math.min(...good) : null,
        max: good.length ? Math.max(...good) : null, mode,
      };
      L.inForce = this.applies(mode, f, good.length, alpha);
      L.applied = L.inForce ? f.slope : 0;
      M.lin[el] = L;
      for (const q of P) {
        q[el].lin = q[el].blank == null ? null : q[el].blank - L.applied * (q[ak] - L.pivot);
      }
    }
    // Monitors before the configuration's first drift position stay out of the
    // drift fits altogether; the elemental correction still reads every monitor.
    const allMons = role('D'), first = S.firstDrift || 0;
    const mons = allMons.filter((q) => q.r >= first);
    const before = allMons.filter((q) => q.r < first).map((q) => q.r);
    for (const [el] of E) {
      const excl = new Set(el === 'n' ? dec.exclN : dec.exclC);
      const used = mons.filter((q) => !excl.has(q.r));
      const ys = used.map((q) => q[el].lin);
      const byRow = this.fit(used.map((q) => q.r), ys);
      const byBgd = this.fit(used.map((q) => q.rbgd), ys);
      let model;
      if (dec.model === 'Row' || dec.model === 'Background') model = dec.model;
      else if (byRow && byBgd && byRow.sd != null && byBgd.sd != null && byBgd.sd < byRow.sd) model = 'Background';
      else model = 'Row';
      const f = model === 'Row' ? byRow : byBgd;
      const xs = model === 'Row' ? used.map((q) => q.r) : used.map((q) => q.rbgd);
      const pr = xs.filter((v, i) => v != null && ys[i] != null);
      const Dr = {
        row: byRow, bgd: byBgd, model, fit: f,
        pivot: pr.length ? pr.reduce((s, v) => s + v, 0) / pr.length : 0,
        n: pr.length, excluded: [...excl].sort((a, b) => a - b), t: {}, before: before.slice(),
      };
      Dr.inForce = this.applies(dec.drift, f, pr.length, alpha);
      Dr.applied = Dr.inForce ? f.slope : 0;
      M.drift[el] = Dr;
      for (const q of P) {
        const v = q[el].lin;
        if (v == null) { q[el].drift = null; continue; }
        const x = model === 'Row' || q.rbgd == null ? q.r : q.rbgd;
        q[el].drift = v - Dr.applied * (x - Dr.pivot);
      }
      for (const q of mons) {
        const x = model === 'Row' ? q.r : q.rbgd;
        Dr.t[q.r] = this.stud(f, x, q[el].lin, !excl.has(q.r));
      }
      Dr.outliers = mons.filter((q) => Dr.t[q.r] != null && Math.abs(Dr.t[q.r]) > S.outlierT && !excl.has(q.r))
        .map((q) => q.r).sort((a, b) => a - b);
    }
    const anchors = role('A');
    for (const [el] of E) {
      const ak = el === 'n' ? 'acc15' : 'acc13';
      const xs = anchors.map((q) => q[el].drift), ys = anchors.map((q) => q[ak]);
      const f = this.fit(xs, ys) || { slope: null, intercept: null, r2: null, sd: null, n: 0 };
      const acc = ys.filter((a, i) => xs[i] != null && a != null);
      f.accMin = acc.length ? Math.min(...acc) : null;
      f.accMax = acc.length ? Math.max(...acc) : null;
      M.cal[el] = f;
      for (const q of P) {
        q[el].fin = q[el].drift == null || f.n < 2 ? null : f.slope * q[el].drift + f.intercept;
      }
    }
    for (const [el] of E) {
      const f = M.cal[el], ak = el === 'n' ? 'acc15' : 'acc13', uk = el === 'n' ? 'uacc15' : 'uacc13';
      const groups = new Map();
      for (const q of anchors) {
        const x = q[el].drift, a = q[ak];
        if (x == null || a == null) continue;
        if (!groups.has(q.si)) groups.set(q.si, { x: [], uAcc: q[uk], si: q.si });
        groups.get(q.si).x.push(x);
      }
      const B = { groups: [], pooled: null, uS: null, src: null, ok: false, df: 0 };
      if (f.n >= 3 && f.sxx) {
        const m = f.slope, mx = f.mean_x, sxx = f.sxx, n = f.n;
        let ss = 0, dfree = 0;
        for (const [idx, g] of groups) {
          const k = g.x.length, sd = this.sdev(g.x);
          g.n = k; g.mean = g.x.reduce((s, v) => s + v, 0) / k; g.sd = sd;
          g.uMean = sd != null ? sd / Math.sqrt(k) : 0;
          g.s1 = g.x.reduce((s, v) => s + (v - mx), 0);
          const acc = anchors.filter((q) => q.si === idx && q[el].drift != null)[0][ak];
          g.s2 = g.x.reduce((s, v) => s + ((acc - (m * v + f.intercept)) - m * (v - mx)), 0);
          if (sd != null) { ss += (k - 1) * sd * sd; dfree += k - 1; }
          B.groups.push(g);
        }
        B.pooled = dfree > 0 ? Math.sqrt(ss / dfree) : null;
        B.df = dfree;
        const chosen = el === 'n' ? dec.uN : dec.uC;
        if (chosen != null) { B.uS = chosen; B.src = 'configuration'; } else { B.uS = B.pooled; B.src = 'pooled anchor replicates'; }
        B.ok = true;
        const uS = B.uS;
        B.at = (x) => {
          let tot = (uS * m) ** 2;
          const parts = { rep: (uS * m) ** 2, acc: 0, mean: 0 };
          for (const g of B.groups) {
            const wAcc = g.n / n + g.s1 * (x - mx) / sxx;
            const wMeas = (x - mx) * g.s2 / sxx - g.n * m / n;
            const a = ((g.uAcc || 0) * wAcc) ** 2, b = (g.uMean * wMeas) ** 2;
            tot += a + b; parts.acc += a; parts.mean += b;
          }
          return { u: Math.sqrt(tot), parts };
        };
        for (const q of P) {
          const x = q[el].drift;
          if (x == null || uS == null) { q[el].u = null; q[el].U = null; continue; }
          const r = B.at(x);
          q[el].u = r.u; q[el].U = S.k * r.u; q[el].parts = r.parts;
        }
      } else {
        for (const q of P) { q[el].u = null; q[el].U = null; }
      }
      M.unc[el] = B;
    }
    for (const q of P) {
      q.pctNm = q.ugN == null || !q.amt ? null : q.ugN / (q.amt * 1000) * 100;
      q.pctCm = q.ugC == null || !q.amt ? null : q.ugC / (q.amt * 1000) * 100;
      q.cnm = q.ugN && q.ugC && q.ugN > 0 ? (q.ugC / D.atomic.c) / (q.ugN / D.atomic.n) : null;
    }
    const exE = new Set(dec.exclE);
    for (const [el, pm, ck] of [['n', 'pctNm', 'certN'], ['c', 'pctCm', 'certC']]) {
      const ms = [];
      for (const q of allMons) {
        q['bias' + el] = null;
        const meas = q[pm], cert = q[ck];
        if (meas == null || !cert) continue;
        const bias = meas / cert - 1;
        q['bias' + el] = bias;
        if (!exE.has(q.r)) ms.push([q.r, 1 + bias]);
      }
      ms.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      const maxBias = ms.length ? Math.max(...ms.map((m) => Math.abs(m[1] - 1))) : null;
      let inForce;
      if (dec.elem === 'Off' || ms.length < 2) inForce = false;
      else if (dec.elem === 'On') inForce = true;
      else inForce = this.above(maxBias * 100, S.elemGate);
      const at = (row) => {
        if (!inForce) return 1;
        if (row <= ms[0][0]) return ms[0][1];
        if (row >= ms[ms.length - 1][0]) return ms[ms.length - 1][1];
        for (let i = 0; i + 1 < ms.length; i++) {
          const [x0, f0] = ms[i], [x1, f1] = ms[i + 1];
          if (x0 <= row && row <= x1) return f0 + (row - x0) / (x1 - x0) * (f1 - f0);
        }
        return 1;
      };
      M.elem[el] = { mode: dec.elem, n: ms.length, maxBias, inForce, excluded: [...exE].sort((a, b) => a - b), mons: ms, at };
      for (const q of P) q['f' + el] = at(q.r);
    }
    for (const q of P) {
      q.pctN = q.pctNm == null ? null : q.pctNm / q.fn;
      q.pctC = q.pctCm == null ? null : q.pctCm / q.fc;
      q.cn = q.cnm == null ? null : q.cnm * q.fn / q.fc;
    }
    const L = D.limits, settled = D.settle.settled_from;
    for (const q of P) {
      const fl = [];
      if (q.aN == null) fl.push('no N peak');
      else if (L.min_n && q.aN < L.min_n) fl.push('N below linearity range');
      else if (L.max_n && q.aN > L.max_n) fl.push('N above linearity range');
      if (q.aC == null) {
        fl.push('no C peak');
        if (q.rej) fl.push('CO2 peak rejected as a reference pulse');
      } else if (L.min_c && q.aC < L.min_c) fl.push('C below linearity range');
      else if (L.max_c && q.aC > L.max_c) fl.push('C above linearity range');
      for (const [v, f, name] of [[q.n.fin, M.cal.n, 'd15N'], [q.c.fin, M.cal.c, 'd13C']]) {
        if (v != null && (q.role === 'S' || q.role === 'K') && f.accMin != null && !(f.accMin <= v && v <= f.accMax)) {
          fl.push(name + ' outside anchor range');
        }
      }
      if (q.role === 'K') {
        for (const [got, acc, name] of [[q.n.fin, q.acc15, 'd15N'], [q.c.fin, q.acc13, 'd13C']]) {
          if (got != null && acc != null && Math.abs(got - acc) > S.checkOffset) {
            fl.push('check ' + name + ' off by ' + (got - acc >= 0 ? '+' : '') + (got - acc).toFixed(2));
          }
        }
      }
      if (q.role === 'D') {
        for (const [el, nm] of [['n', 'N'], ['c', 'C']]) {
          if (M.drift[el].excluded.includes(q.r)) fl.push('excluded from ' + nm + ' drift fit');
          else if (M.drift[el].outliers.includes(q.r)) fl.push(nm + ' drift monitor outlier');
        }
        if (exE.has(q.r)) fl.push('excluded from elemental correction');
      }
      if (q.role !== 'B' && S.blankMultiple) {
        for (const [a, ab, nm] of [[q.aN, M.blank.n.area, 'N'], [q.aC, M.blank.c.area, 'C']]) {
          if (a != null && ab > 0 && a < ab * S.blankMultiple) fl.push(nm + ' peak only ' + (a / ab).toFixed(1) + 'x the blank');
        }
      }
      if (q.r < settled) fl.push('ran before the background settled');
      if (q.carry) fl.push('possible carryover from an oversized sample');
      q.flags = fl;
    }
    return M;
  }

  // ---------------------------------------------------------------- decisions
  dec0() {
    return { blank: 'Auto', linN: 'Auto', linC: 'Auto', drift: 'Auto', model: 'Row',
      exclN: [], exclC: [], elem: 'Auto', exclE: [], uSrcN: 'pooled', uSrcC: 'pooled' };
  }
}

const engine = new PortEngine();
if (typeof module !== 'undefined' && module.exports) module.exports = engine;
