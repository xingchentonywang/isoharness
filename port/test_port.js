// Hold the preview port to the engine: every value, flag list and fit result on
// every decision set in the cases file must agree within 1e-9 relative. A case
// with "settings" overrides the block's settings, and one with "std" its standards
// table, as port_cases.py did for the engine. From ea_irms 1.4.15 the values include
// each position's Kragten terms of the corrections in force, and the fits the floor
// with each step of the walk up the size series, the drift monitors' SD against the
// repeatability in use, the largest term of each correction, the acetanilide's
// whole-run check, and blank capsules that gave no peak.
//
//   node port/test_port.js PORT_INPUT.json CASES.json [ENGINE.js]
//
// Both JSON files come from scripts/port_cases.py, which runs ea_irms itself.
// Exits 1 on any mismatch.
const fs = require('fs');
const path = require('path');

const [inputPath, casesPath, enginePath] = process.argv.slice(2);
if (!inputPath || !casesPath) {
  console.error('usage: node port/test_port.js PORT_INPUT.json CASES.json [ENGINE.js]');
  process.exit(2);
}
const E = require(path.resolve(enginePath || path.join(__dirname, 'engine.js')));
const D = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
const cases = JSON.parse(fs.readFileSync(casesPath, 'utf8'));
const TOL = 1e-9;

let worst = 0, mism = 0, comps = 0, flagMism = 0, fitMism = 0;
const close = (a, b, tag) => {
  comps++;
  if (a == null && b == null) return;
  if (a == null || b == null || typeof a !== typeof b) { mism++; if (mism < 15) console.log('null mismatch', tag, a, b); return; }
  if (typeof a === 'boolean' || typeof a === 'string') { if (a !== b) { mism++; if (mism < 15) console.log('mismatch', tag, a, b); } return; }
  const rel = Math.abs(a - b) / Math.max(1, Math.abs(b));
  if (rel > worst) worst = rel;
  if (rel > TOL) { mism++; if (mism < 15) console.log('value mismatch', tag, a, b); }
};

for (const [ci, C] of cases.entries()) {
  const c = C.case;
  const dec = { blank: c.blank, linN: c.linN, linC: c.linC, drift: c.drift, model: c.model,
    exclN: c.exclN, exclC: c.exclC, elem: c.elem, exclE: c.exclE, uN: c.uN, uC: c.uC };
  const DC = Object.assign({}, D, c.settings ? { settings: Object.assign({}, D.settings, c.settings) } : {},
    c.std ? { std: c.std } : {});
  const M = E.runModel.call(E, DC, dec);
  const byRow = new Map(M.P.map((q) => [q.r, q]));
  const TERMS = ['blank', 'linearity', 'drift'];
  for (const row of C.rows) {
    const q = byRow.get(row[0]);
    if (!q) { mism++; console.log('missing row', ci, row[0]); continue; }
    const vals = [q.n.blank, q.n.lin, q.n.drift, q.n.fin, q.n.u, q.n.U, q.c.blank, q.c.lin, q.c.drift, q.c.fin, q.c.u, q.c.U,
      q.pctN, q.pctC, q.cn, q.fn, q.fc, q.role === 'D' ? M.drift.n.t[q.r] : null, q.role === 'D' ? M.drift.c.t[q.r] : null];
    for (let i = 0; i < vals.length; i++) close(vals[i] === undefined ? null : vals[i], row[i + 1], `case ${ci} row ${row[0]} col ${i}`);
    // the flags are last in a row
    const flags = row[row.length - 1];
    comps++;
    if (JSON.stringify(flags) !== JSON.stringify(q.flags)) { flagMism++; if (flagMism < 10) console.log('flags', ci, row[0], flags, q.flags); }
    // ea_irms 1.4.15: the Kragten terms of the corrections in force, blank, linearity, drift for N then C, before the flags
    if (row.length > 21) {
      const terms = [];
      for (const el of ['n', 'c']) for (const k of TERMS) terms.push(q[el].terms ? q[el].terms[k] : undefined);
      for (let i = 0; i < terms.length; i++) close(terms[i] === undefined ? null : terms[i], row[20 + i], `case ${ci} row ${row[0]} term ${i}`);
    }
  }
  const F = C.fits;
  const chk = (a, b, tag) => { const before = mism; close(a, b, tag); if (mism > before) fitMism++; };
  chk(M.blank.n.inForce, F.blankN, 'blankN'); chk(M.blank.c.inForce, F.blankC, 'blankC');
  chk(M.lin.n.inForce, F.linN[0], 'linN'); chk(M.lin.n.fit && M.lin.n.fit.p, F.linN[1], 'linNp');
  chk(M.lin.c.inForce, F.linC[0], 'linC'); chk(M.lin.c.fit && M.lin.c.fit.p, F.linC[1], 'linCp');
  for (const [el, k] of [['n', 'driftN'], ['c', 'driftC']]) {
    chk(M.drift[el].inForce, F[k][0], k); chk(M.drift[el].model, F[k][1], k + 'model');
    chk(M.drift[el].fit && M.drift[el].fit.p, F[k][2], k + 'p'); chk(M.drift[el].fit && M.drift[el].fit.slope, F[k][3], k + 'slope');
    comps++;
    if (JSON.stringify(M.drift[el].outliers) !== JSON.stringify(F[k][4])) { fitMism++; console.log('outliers', ci, el, M.drift[el].outliers, F[k][4]); }
  }
  for (const [el, k] of [['n', 'calN'], ['c', 'calC']]) { chk(M.cal[el].slope, F[k][0], k); chk(M.cal[el].intercept, F[k][1], k); chk(M.cal[el].sd, F[k][2], k); }
  for (const [el, k] of [['n', 'uncN'], ['c', 'uncC']]) { chk(M.unc[el].pooled, F[k][0], k); chk(M.unc[el].uS, F[k][1], k); }
  for (const [el, k] of [['n', 'elemN'], ['c', 'elemC']]) { chk(M.elem[el].inForce, F[k][0], k); chk(M.elem[el].maxBias, F[k][1], k); chk(M.elem[el].n, F[k][2], k); }
  // ea_irms 1.4.13: the elemental drift correction against the anchors (null where it is not in force)
  for (const [el, k] of [['n', 'elemCheckN'], ['c', 'elemCheckC']]) {
    const got = M.elem[el].check, want = F[k];
    if (want == null || got == null) {
      comps++;
      if ((want == null) !== (got == null)) { fitMism++; console.log('elemental check', ci, el, got, want); }
      continue;
    }
    chk(got.level, want[0], k + 'level'); chk(got.rmsMeas, want[1], k + 'rmsMeas'); chk(got.rmsRep, want[2], k + 'rmsRep');
    chk(got.added, want[3], k + 'added'); chk(got.n, want[4], k + 'n');
  }
  // ea_irms 1.4.15
  const same = (a, b, tag) => { comps++; if (JSON.stringify(a) !== JSON.stringify(b)) { fitMism++; if (fitMism < 15) console.log('fit', ci, tag, a, b); } };
  for (const [el, key] of [['n', 'N'], ['c', 'C']]) {
    const want = F['floor' + key];
    if (want) {
      const lim = M.limits[el], fl = M.floor[el];
      chk(lim.min, want[0], 'floor' + key + 'min'); chk(lim.src, want[1], 'floor' + key + 'src');
      chk(fl.floor, want[2], 'floor' + key + 'floor'); chk(fl.moved, want[3], 'floor' + key + 'moved');
      chk(fl.smallest, want[4], 'floor' + key + 'smallest');
      same(fl.steps.length, want[5].length, 'floor' + key + 'steps');
      fl.steps.forEach((s, i) => {
        if (!want[5][i]) return;
        chk(s.row, want[5][i][0], 'floor' + key + 'row'); chk(s.resid, want[5][i][1], 'floor' + key + 'resid');
        chk(s.limit, want[5][i][2], 'floor' + key + 'limit'); chk(s.moved, want[5][i][3], 'floor' + key + 'stepMoved');
      });
    }
    const mon = F['mon' + key];
    if (mon) {
      const m = M.unc[el].monitors || {};
      [m.n, m.sd, m.repeatability, m.df, m.p, m.larger, m.used].forEach((v, i) => chk(v === undefined ? null : v, mon[i], 'mon' + key + i));
    }
    const corr = F['corr' + key];
    if (corr) {
      const c2 = M.unc[el].corrections || {};
      ['blank', 'linearity', 'drift'].forEach((k, i) => chk(c2[k] === undefined ? null : c2[k], corr[i], 'corr' + key + k));
      same(Object.keys(c2).sort(), corr[3], 'corr' + key + 'keys');
    }
    const ace = F['ace' + key];
    if (ace) {
      const a = M.ace[el] || {};
      [a.n, a.sd, a.basis, a.df, a.p, a.over].forEach((v, i) => chk(v === undefined ? null : v, ace[i], 'ace' + key + i));
    }
  }
  if (F.blankBelow) {
    chk(M.blank.n.belowDetection, F.blankBelow[0], 'blankBelowN'); chk(M.blank.c.belowDetection, F.blankBelow[1], 'blankBelowC');
    chk(M.blank.n.capsules, F.blankBelow[2], 'blankCapsules');
  }
}

const nSet = cases.filter((C) => C.case.settings || C.case.std).length;
console.log(`${D.file}: cases ${cases.length} (${nSet} with non-default settings), comparisons ${comps}, value mismatches ${mism}, ` +
  `flag-list mismatches ${flagMism}, fit mismatches ${fitMism}, worst relative difference ${worst.toExponential(2)}`);
process.exit(mism || flagMism || fitMism ? 1 : 0);
