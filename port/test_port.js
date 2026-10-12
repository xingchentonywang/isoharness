// Hold the preview port to the engine: every value, flag list and fit result on
// every decision set in the cases file must agree within 1e-9 relative.
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
  const M = E.runModel.call(E, D, dec);
  const byRow = new Map(M.P.map((q) => [q.r, q]));
  for (const row of C.rows) {
    const q = byRow.get(row[0]);
    if (!q) { mism++; console.log('missing row', ci, row[0]); continue; }
    const vals = [q.n.blank, q.n.lin, q.n.drift, q.n.fin, q.n.u, q.n.U, q.c.blank, q.c.lin, q.c.drift, q.c.fin, q.c.u, q.c.U,
      q.pctN, q.pctC, q.cn, q.fn, q.fc, q.role === 'D' ? M.drift.n.t[q.r] : null, q.role === 'D' ? M.drift.c.t[q.r] : null];
    for (let i = 0; i < vals.length; i++) close(vals[i] === undefined ? null : vals[i], row[i + 1], `case ${ci} row ${row[0]} col ${i}`);
    comps++;
    if (JSON.stringify(row[20]) !== JSON.stringify(q.flags)) { flagMism++; if (flagMism < 10) console.log('flags', ci, row[0], row[20], q.flags); }
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
}

console.log(`${D.file}: cases ${cases.length}, comparisons ${comps}, value mismatches ${mism}, ` +
  `flag-list mismatches ${flagMism}, fit mismatches ${fitMism}, worst relative difference ${worst.toExponential(2)}`);
process.exit(mism || flagMism || fitMism ? 1 : 0);
