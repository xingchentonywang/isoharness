// Run the page's logic (web/logic.js) under node, outside a browser, for the tests.
//
//   node tests/page_harness.js < INPUT.json
//
// INPUT is { "bodies": [answers of POST /api/process, one per run], "code": "...", "data": ... }. Each answer
// becomes a run of the batch, as the page makes it (makeRun), with the preview port (port/engine.js) as the
// page's engine; the first run is shown. "code" is the body of a function of (L, engine, input), L being the
// page, and its return value is printed as JSON. The page is drawn nowhere: dc.js's DCLogic is stood in for by
// a class that only keeps state, and window holds no storage.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const engine = require(path.join(ROOT, 'port', 'engine.js'));

class DCLogic {
  constructor(props) { this.props = props || {}; this.state = null; }
  setState(next) {
    const add = typeof next === 'function' ? next(this.state || {}, this.props) : next;
    this.state = Object.assign({}, this.state || {}, add);
  }
  forceUpdate() {}
}

const context = vm.createContext({ engine, DCLogic, console, window: {}, setTimeout, clearTimeout });
const source = fs.readFileSync(path.join(ROOT, 'web', 'logic.js'), 'utf8');
vm.runInContext(source + '\n;globalThis.Component = Component;', context, { filename: 'web/logic.js' });

const input = JSON.parse(fs.readFileSync(0, 'utf8'));
const L = new context.Component({});
const runs = (input.bodies || []).map((body) => {
  const run = L.makeRun({ name: body.file.name }, null, body, []);
  run.check = L.verify(run);
  run.own = null;
  run.saved = { dec: run.decStart, sel: {}, filt: 'all', runDate: '', dateFrom: '', last: 'inst', runInfo: {} };
  return run;
});
if (runs.length) L.setState(Object.assign(L.restore(runs[0]), { runs, active: 0 }));
const fn = vm.runInContext('(function (L, engine, input) {\n' + input.code + '\n})', context);
process.stdout.write(JSON.stringify(fn(L, engine, input)));
