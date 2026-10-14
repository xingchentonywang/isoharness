// isoharness page logic: the approved prototype's views (its logic.js), on a real upload.
//
// Every value on the step screens comes from the preview port: the global `engine`
// that /engine.js defines (port/engine.js, the file continuous integration holds to
// ea_irms). The page calls engine.runModel on each choice; the workbook, the results
// table and the processed configuration come from the engine on the server.
//
// The upload screen posts each export, and a lab configuration if one is chosen, to
// /api/process. Its "port" block and configuration drive the instrument check, the eight
// steps and the results. The downloads post the same export again, kept in this tab only, with the
// configuration built from the decisions (configFor).
//
// Several exports make a batch (isoharness 0.2.0): each run is processed and corrected on its own,
// with the one lab configuration, and keeps its own decisions. The run on screen holds its working
// state (dec, sel, filt, runDate, dateFrom, last) in this.state; the others keep theirs in run.saved.
// The combined results post every run with its own configuration to /api/batch/results.xlsx.
//
// From isoharness 0.3.0 each run has its information (Operator, Instrument, Project, Notes), kept like
// its run date, and the batch may have one sample sheet. The service reads the sheet and matches the
// samples (/api/samplesheet/read); the page keeps the file in this tab and sends it, with each run's
// information, with every download.
'use strict';

const FLOW = ['inst', 'setup', 'blank', 'lin', 'drift', 'cal', 'elem', 'unc', 'qc', 'results'];
const NAMES = { upload: 'Upload', inst: 'Instrument', setup: 'Setup', blank: 'Blank', lin: 'Linearity', drift: 'Drift', cal: 'Calibration',
  elem: 'Elemental content', unc: 'Uncertainty', qc: 'Quality checks', results: 'Results' };
// The twelve columns ea_irms reads (ea_irms.parse.REQUIRED), by their exact names.
const COLUMNS = [
  ['Position', '(autosampler row)', 'Row'], ['Sample name', '', 'Identifier 1'], ['Peak number', '', 'Peak Nr'],
  ['Weight', '(mg)', 'Amount'], ['N2 amplitude', '(mV)', 'Ampl  28'], ['N2 background', '(mV)', 'BGD 28'],
  ['N2 peak area', '(Vs)', 'Area 28'], ['Reference-gas flag', '(1 marks a reference peak)', 'Is Ref _'],
  ['CO2 amplitude', '(mV)', 'Ampl  44'], ['CO2 peak area', '(Vs)', 'Area 44'],
  ['δ¹⁵N against the reference gas', '(‰)', 'd 15N/14N'], ['δ¹³C against the reference gas', '(‰)', 'd 13C/12C'],
];
const WORDS = ['no', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve'];
// The working-gas SD limit, in permil: the FIRMS guide (Dunn and Carter, 2018, 2nd edition,
// section 4.3.2) asks for an SD under 0.1 permil for CO2 and N2 in its zero-enrichment test of
// about ten pulses; Tony chose on 2026-10-12 to hold each run's pulses to it.
const REF_SD_LIMIT = 0.1;
const OK_ICON = { c: '#0B6E4F', d: 'M6 10.5l2.6 2.6L14 7.6' };
// The most runs in one batch, as the service combines them (ISOHARNESS_MAX_BATCH_RUNS, default 20).
const MAX_RUNS = 20;
// The run information a run carries (ea_irms.metadata.RUN_FIELDS).
const RUN_FIELDS = ['Operator', 'Instrument', 'Project', 'Notes'];
// The most rows of the sample-sheet match table drawn on the page; the downloads hold every sample.
const MAX_MATCH_ROWS = 600;
const WARN_ICON = { c: '#A33A00', d: 'M10 5.5v5.5M10 14v.1' };

class Component extends DCLogic {
  constructor(props) {
    super(props);
    this.state = { screen: 'upload', boot: null, run: null, runs: [], active: -1, cfg: null, runDate: '', dateFrom: '',
      busy: '', pending: '', error: null, toast: '', drag: false, dec: null, sel: {}, filt: 'all', last: 'inst',
      runInfo: {}, sheet: null };
    this._mcs = new Map();
    this._runs = 0;
    this._depth = 0;
  }

  // ---------------------------------------------------------------- the run
  data() { return this.state.run.port; }

  // The decisions a configuration starts from (linC from linearity_correction_c, or
  // from linearity_correction when that is empty, as the engine reads it).
  decisionsFrom(cfg) {
    const list = (v) => (Array.isArray(v) ? v.map(Number).filter((x) => Number.isFinite(x)).sort((a, b) => a - b) : []);
    return {
      blank: cfg.blank_correction, linN: cfg.linearity_correction,
      linC: cfg.linearity_correction_c ? cfg.linearity_correction_c : cfg.linearity_correction,
      drift: cfg.drift_correction, model: cfg.drift_model,
      exclN: list(cfg.exclude_drift_n), exclC: list(cfg.exclude_drift_c),
      elem: cfg.elemental_correction, exclE: list(cfg.exclude_elemental),
      uSrcN: cfg.u_sample_n == null ? 'pooled' : 'config', uSrcC: cfg.u_sample_c == null ? 'pooled' : 'config',
    };
  }

  // The configuration for the downloads: the one the run was processed with, with the
  // decisions written into their fields and every other field left as it was.
  configFor(d, run) {
    run = run || this.state.run;
    const M = this.model(d, run), out = {};
    for (const k of Object.keys(run.config)) out[k] = run.config[k];
    return Object.assign(out, {
      blank_correction: d.blank, linearity_correction: d.linN, linearity_correction_c: d.linC,
      drift_correction: d.drift, drift_model: d.model, exclude_drift_n: d.exclN.slice(), exclude_drift_c: d.exclC.slice(),
      elemental_correction: d.elem, exclude_elemental: d.exclE.slice(), u_sample_n: M.u.n, u_sample_c: M.u.c,
    });
  }

  // The port's model for a set of decisions. The repeatability source becomes u_sample_n
  // and u_sample_c: none for the pooled anchors, the configuration's value, or the SD of
  // the drift monitors in the fit (which needs the drift step first, hence a second run).
  model(d, run) {
    run = run || this.state.run;
    let mc = this._mcs.get(run.id);
    if (!mc) { mc = new Map(); this._mcs.set(run.id, mc); }
    const key = JSON.stringify(d);
    const hit = mc.get(key);
    if (hit) return hit;
    const D = run.port;
    const e = { blank: d.blank, linN: d.linN, linC: d.linC, drift: d.drift, model: d.model,
      exclN: d.exclN.slice(), exclC: d.exclC.slice(), elem: d.elem, exclE: d.exclE.slice(),
      uN: null, uC: null };
    const pick = (el, sd) => {
      const src = el === 'n' ? d.uSrcN : d.uSrcC;
      if (src === 'config') return run.uCfg[el] == null ? null : run.uCfg[el];
      if (src === 'monitors') return sd[el] == null ? null : sd[el];
      return null;
    };
    e.uN = pick('n', {});
    e.uC = pick('c', {});
    let M = engine.runModel(D, e);
    const sd = { n: this.monSD(M, e.exclN, 'n'), c: this.monSD(M, e.exclC, 'c') };
    if (d.uSrcN === 'monitors' || d.uSrcC === 'monitors') {
      e.uN = pick('n', sd);
      e.uC = pick('c', sd);
      M = engine.runModel(D, e);
    }
    M.monSD = sd;
    M.u = { n: e.uN, c: e.uC };
    if (mc.size > 80) mc.clear();
    mc.set(key, M);
    return M;
  }

  // The SD of the drift monitors in the fit, after the drift step.
  monSD(M, excluded, el) {
    // The monitors in the fit: neither left out nor before the first drift position.
    const ex = new Set(excluded.concat(M.drift[el].before || []));
    return engine.sdev(M.P.filter((q) => q.role === 'D' && !ex.has(q.r)).map((q) => q[el].drift));
  }

  // The repeatability source in force: the monitors only when two or more are in the
  // fit, the configuration's value only when it has one.
  uSrc(M, d, el) {
    const want = el === 'n' ? d.uSrcN : d.uSrcC;
    if (want === 'monitors') return M.monSD[el] != null ? 'monitors' : 'pooled';
    if (want === 'config') return M.u[el] != null ? 'config' : 'pooled';
    return 'pooled';
  }

  // The port's values on the configuration's own decisions against the engine's, at
  // every position: the corrections, U, %N, %C, C:N, the elemental factors, the flags.
  verify(run) {
    const M = this.model(run.dec0, run);
    const byRow = new Map(run.positions.map((p) => [p.row, p]));
    const same = (a, b) => (a == null || b == null ? a == null && b == null
      : Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)));
    const keys = [];
    for (const [el, iso] of [['n', 'd15n'], ['c', 'd13c']]) {
      for (const [k, s] of [['blank', '_blank'], ['lin', '_lin'], ['drift', '_drift'], ['fin', '_final']]) keys.push([(q) => q[el][k], iso + s]);
      keys.push([(q) => q[el].u, 'u_' + iso], [(q) => q[el].U, 'U_' + iso]);
    }
    keys.push([(q) => q.pctN, 'pct_n_reported'], [(q) => q.pctC, 'pct_c_reported'], [(q) => q.cn, 'cn_molar_reported'],
      [(q) => q.fn, 'elem_factor_n'], [(q) => q.fc, 'elem_factor_c']);
    let values = 0;
    const bad = [];
    for (const q of M.P) {
      const p = byRow.get(q.r);
      let ok = !!p;
      if (p) {
        for (const [get, key] of keys) { values++; if (!same(get(q), p[key])) ok = false; }
        if (JSON.stringify(q.flags) !== JSON.stringify(p.flags || [])) ok = false;
      }
      if (!ok) bad.push(q.r);
    }
    return { values, positions: M.P.length, bad };
  }

  makeRun(file, cfg, body) {
    const port = body.port, config = body.config;
    const run = {
      id: ++this._runs, file, name: body.file.name, bytes: body.file.bytes, version: body.ea_irms,
      port, config, positions: body.positions || [],
      cfgName: cfg ? cfg.name : null, cfgUnknown: cfg ? cfg.unknown : [],
      uCfg: { n: config.u_sample_n == null ? null : config.u_sample_n, c: config.u_sample_c == null ? null : config.u_sample_c },
      rejected: new Set(((body.fits || {}).ref_pulse_check || {}).rejected || []),
    };
    run.dec0 = this.decisionsFrom(config);
    return run;
  }

  // ---------------------------------------------------------------- the batch
  // The working state of the run on screen, as run.saved keeps it for the others.
  stash() {
    const st = this.state;
    if (st.run) st.run.saved = { dec: st.dec, sel: st.sel, filt: st.filt, runDate: st.runDate, dateFrom: st.dateFrom, last: st.last, runInfo: st.runInfo };
  }

  restore(run) {
    const s = run.saved || { dec: run.dec0, sel: {}, filt: 'all', runDate: '', dateFrom: '', last: 'inst', runInfo: {} };
    return { run, dec: s.dec, sel: s.sel, filt: s.filt, runDate: s.runDate, dateFrom: s.dateFrom, last: s.last, runInfo: s.runInfo || {} };
  }

  // The decisions and run date of any run in the batch, the one on screen included.
  decOf(run) { return run === this.state.run ? this.state.dec : (run.saved || {}).dec || run.dec0; }

  dateOf(run) { return run === this.state.run ? this.state.runDate : (run.saved || {}).runDate || ''; }

  // The run information of any run in the batch (isoharness 0.3.0), and the fields of it that hold text.
  infoOf(run) { return (run === this.state.run ? this.state.runInfo : (run.saved || {}).runInfo) || {}; }

  infoJSON(run) {
    const info = this.infoOf(run), out = {};
    for (const k of RUN_FIELDS) if (String(info[k] || '').trim()) out[k] = String(info[k]).trim();
    return out;
  }

  setInfo(k, v) { this.setState({ runInfo: Object.assign({}, this.state.runInfo, { [k]: v }) }); }

  // The run information of the run shown, copied to every run of the batch.
  infoToAll() {
    const st = this.state;
    if (!st.run || st.busy) return;
    for (const r of st.runs) if (r !== st.run) r.saved = Object.assign({}, r.saved || {}, { runInfo: Object.assign({}, st.runInfo) });
    this.setState({ toast: 'Every run of the batch now has the run information of ' + st.run.name + '.' });
  }

  // Show another run of the batch, on the same screen, with the decisions it was left with.
  switchRun(i) {
    const st = this.state;
    if (st.busy || i === st.active || !st.runs[i]) return;
    this.stash();
    this.setState(Object.assign(this.restore(st.runs[i]), { active: i, toast: '', error: null }));
  }

  removeRun(i) {
    const st = this.state;
    if (st.busy || !st.runs[i]) return;
    this.stash();
    const gone = st.runs[i], runs = st.runs.filter((_, j) => j !== i);
    this._mcs.delete(gone.id);
    if (!runs.length) {
      this.setState({ runs, active: -1, run: null, dec: null, sel: {}, filt: 'all', runDate: '', dateFrom: '', last: 'inst',
        runInfo: {}, error: null, toast: gone.name + ' was removed.' });
      this.go('upload');
      this.refreshSheet();
      return;
    }
    const active = st.active === i ? Math.min(i, runs.length - 1) : st.active > i ? st.active - 1 : st.active;
    this.setState(Object.assign(this.restore(runs[active]), { runs, active, error: null, toast: gone.name + ' was removed from the batch.' }));
    this.refreshSheet();
  }

  // ---------------------------------------------------------------- requests
  async boot() {
    const get = async (url) => {
      const r = await fetch(url);
      if (!r.ok) throw new Error(url + ' answered ' + r.status);
      return r.json();
    };
    try {
      const [v, d] = await Promise.all([get('api/version'), get('api/defaults')]);
      this.setState({ boot: { isoharness: v.isoharness, ea_irms: v.ea_irms, libreoffice: !!v.libreoffice, defaults: d.config } });
    } catch (err) {
      console.warn('isoharness: the service did not describe itself', err);
      this.setState({ boot: { failed: true } });
    }
  }

  async answer(resp) {
    let body = null;
    try { body = await resp.json(); } catch (err) { body = null; }
    return body;
  }

  // The service's own words when it refuses, as it wrote them.
  detail(resp, body) {
    if (body && typeof body.detail === 'string') return body.detail;
    if (body && Array.isArray(body.detail)) return body.detail.map((x) => (x && x.msg) || String(x)).join('; ');
    return 'The service answered ' + resp.status + (resp.statusText ? ' ' + resp.statusText : '') + ' without a message.';
  }

  // Exports chosen or dropped together are read one after another and added to the batch; an export
  // with the name of a run already in it replaces that run. The run date comes from each file's name;
  // a date typed before the first upload goes with it when its name has none.
  async addExports(files) {
    if (this.state.busy) return;
    const list = Array.from(files || []).filter(Boolean);
    if (!list.length) return;
    const fails = [];
    // Run information typed before the first upload goes with every run of that upload.
    const pre = this.state.runs.length ? null : Object.assign({}, this.state.runInfo || {});
    for (let k = 0; k < list.length; k++) {
      const file = list[k], st = this.state;
      const at = st.runs.findIndex((r) => r.name === file.name);
      if (at < 0 && st.runs.length >= MAX_RUNS) {
        fails.push({ name: file.name, detail: 'The batch already holds ' + MAX_RUNS + ' runs, the most the combined results take; remove one first.' });
        continue;
      }
      let runDate = this.dateFromName(file.name), dateFrom = runDate ? 'name' : '';
      if (!runDate && !st.runs.length && st.dateFrom === 'you' && st.runDate) { runDate = st.runDate; dateFrom = 'you'; }
      this.setState({ busy: 'process', pending: file.name + (list.length > 1 ? ' (' + (k + 1) + ' of ' + list.length + ')' : ''), error: null, toast: '' });
      const info = at >= 0 ? Object.assign({}, this.infoOf(st.runs[at])) : Object.assign({}, pre || {});
      const run = await this.send(file, st.cfg, runDate);
      if (run.fail) { fails.push({ name: file.name, detail: run.detail }); continue; }
      run.saved = { dec: run.dec0, sel: {}, filt: 'all', runDate, dateFrom, last: 'inst', runInfo: info };
      this.stash();
      const runs = this.state.runs.slice();
      if (at >= 0) { this._mcs.delete(runs[at].id); runs[at] = run; } else runs.push(run);
      this.setState(Object.assign(this.restore(run), { runs, active: at >= 0 ? at : runs.length - 1 }));
    }
    let error = null;
    if (fails.length === 1 && list.length === 1) error = { title: fails[0].name + ' was not processed', detail: fails[0].detail };
    else if (fails.length) {
      error = { title: (fails.length === list.length ? 'None of the ' : fails.length + ' of the ') + list.length + ' exports '
        + (fails.length === 1 ? 'was' : 'were') + ' processed', detail: fails.map((f) => f.name + ': ' + f.detail).join('\n') };
    }
    this.setState({ busy: '', pending: '', error });
    this.forceUpdate();
    if (this.state.sheet && list.length > fails.length) await this.refreshSheet();
  }

  // The 2026-09-15 test run, fetched from the service and then read like a dropped file
  // (dc.js applies setState at once, so busy is clear when addExports looks).
  loadExample() {
    if (this.state.busy) return;
    this.setState({ busy: 'process', pending: '20260915_raw.csv', error: null });
    fetch('api/example', { cache: 'no-store' })
      .then((r) => { if (!r.ok) throw new Error('the service answered ' + r.status); return r.blob(); })
      .then((b) => { this.setState({ busy: '', pending: '' }); this.addExports([new File([b], '20260915_raw.csv', { type: 'text/csv' })]); })
      .catch((e) => this.setState({ busy: '', pending: '', error: { title: 'The example run was not loaded', detail: String(e && e.message ? e.message : e) } }));
  }

  // Post one export, with the configuration, to /api/process. Returns the run, checked against the
  // engine's values, or { fail: true, detail } with the service's own words.
  async send(file, cfg, runDate) {
    const form = new FormData();
    form.append('file', file, file.name);
    if (cfg) form.append('config', JSON.stringify(cfg.data));
    if (runDate) form.append('run_date', runDate);
    let resp;
    try {
      resp = await fetch('api/process', { method: 'POST', body: form });
    } catch (err) {
      return { fail: true, detail: 'The service could not be reached (' + err.message + '). Check the connection and try again.' };
    }
    const body = await this.answer(resp);
    if (!resp.ok || !body || !body.port) {
      return { fail: true, detail: resp.ok ? 'The service sent an answer this page cannot read.' : this.detail(resp, body) };
    }
    try {
      const run = this.makeRun(file, cfg, body);
      run.check = this.verify(run);
      if (run.check.bad.length) console.warn('isoharness: the preview of ' + file.name + ' differs from the engine at positions', run.check.bad);
      return run;
    } catch (err) {
      console.error(err);
      return { fail: true, detail: 'The page could not read the service’s answer (' + err.message + ').' };
    }
  }

  // A new configuration processes every run again; if one is refused, the batch stays as it was.
  async reprocessAll(cfg) {
    this.stash();
    const old = this.state.runs.slice(), active = this.state.active, runs = [], fails = [];
    for (let k = 0; k < old.length; k++) {
      const prev = old[k], s = prev.saved || {};
      this.setState({ busy: 'process', pending: prev.name + (old.length > 1 ? ' (' + (k + 1) + ' of ' + old.length + ')' : ''), error: null, toast: '' });
      const run = await this.send(prev.file, cfg, s.runDate || '');
      if (run.fail) { fails.push({ name: prev.name, detail: run.detail }); continue; }
      run.saved = { dec: run.dec0, sel: {}, filt: 'all', runDate: s.runDate || '', dateFrom: s.dateFrom || '', last: 'inst', runInfo: s.runInfo || {} };
      runs.push(run);
    }
    if (fails.length) {
      const title = cfg ? cfg.name + ' was not accepted' : (old.length > 1 ? 'The runs were' : 'The run was') + ' not processed again';
      const detail = old.length === 1 ? fails[0].detail : fails.map((f) => f.name + ': ' + f.detail).join('\n')
        + '\nThe batch keeps the configuration it had.';
      this.setState(Object.assign(this.restore(old[active]), { busy: '', pending: '', runs: old, active, error: { title, detail } }));
      return;
    }
    for (const r of old) this._mcs.delete(r.id);
    const what = old.length > 1 ? 'The ' + old.length + ' runs were' : 'The run was';
    const toast = cfg ? what + ' processed again with ' + cfg.name + '; the decisions now start from it.'
      : what + ' processed again with the package defaults; the decisions now start from them.';
    this.setState(Object.assign(this.restore(runs[active]), { busy: '', pending: '', runs, active, cfg, error: null, toast }));
    this.forceUpdate();
    await this.refreshSheet();      // the configuration's standards table decides which positions are samples
  }

  async pickConfig(file) {
    if (this.state.busy) return;
    const fail = (detail) => this.setState({ error: { title: file.name + ' is not a configuration', detail } });
    if (/\.xlsx$/i.test(file.name)) {
      // The Excel lab configuration (ea_irms 1.4.6) is read by the service, which names any cell it cannot use.
      const form = new FormData();
      form.append('file', file, file.name);
      let resp;
      try {
        resp = await fetch('api/config/read', { method: 'POST', body: form });
      } catch (err) {
        fail('The service could not be reached to read it (' + err.message + ').');
        return;
      }
      const body = await this.answer(resp);
      if (!resp.ok || !body || !body.config) {
        fail(resp.ok ? 'The service sent an answer this page cannot read.' : this.detail(resp, body));
        return;
      }
      this.useConfig({ name: file.name, data: body.config, unknown: [], wrapped: false, notes: body.notes || [] });
      return;
    }
    let data;
    try {
      data = JSON.parse(await file.text());
    } catch (err) {
      fail('It could not be read as JSON (' + err.message + ').');
      return;
    }
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      fail('A configuration is a JSON object, such as the Configuration file downloaded from a run.');
      return;
    }
    const known = this.state.boot && this.state.boot.defaults ? Object.keys(this.state.boot.defaults) : null;
    let wrapped = false;
    if (known && data.config && typeof data.config === 'object' && !Array.isArray(data.config)
        && !Object.keys(data).some((k) => known.includes(k))) {
      data = data.config;    // the whole /api/defaults answer: its "config" is the configuration
      wrapped = true;
    }
    const keys = Object.keys(data);
    const unknown = known ? keys.filter((k) => !known.includes(k)) : [];
    if (known && keys.length && unknown.length === keys.length) {
      fail('None of its fields is one ea_irms reads (' + this.andList(unknown.slice(0, 5)) + (unknown.length > 5 ? ', and others' : '')
        + '); a configuration has fields such as standards and blank_correction.');
      return;
    }
    this.useConfig({ name: file.name, data, unknown, wrapped, notes: [] });
  }

  useConfig(cfg) {
    if (this.state.runs.length) this.reprocessAll(cfg);
    else this.setState({ cfg, error: null });
  }

  async downloadTemplate() {
    try {
      const resp = await fetch('api/config-template', { cache: 'no-store' });
      if (!resp.ok) throw new Error('the service answered ' + resp.status);
      this.save(await resp.blob(), this.fileName(resp.headers.get('Content-Disposition')) || 'lab_configuration.xlsx');
    } catch (err) {
      this.setState({ error: { title: 'The template was not downloaded', detail: String(err && err.message ? err.message : err) } });
    }
  }

  // ---------------------------------------------------------------- the sample sheet (isoharness 0.3.0)
  // The samples of each run as the page lists them, for the service to match: [{ name, samples: [[row, id]] }].
  sheetRuns() {
    return this.state.runs.map((r) => ({ name: r.name,
      samples: r.port.pos.filter((q) => q.role === 'S').map((q) => [q.r, String(q.id == null ? '' : q.id)]) }));
  }

  async readSheet(file, idColumn) {
    const form = new FormData();
    form.append('file', file, file.name);
    if (idColumn) form.append('id_column', idColumn);
    form.append('runs', JSON.stringify(this.sheetRuns()));
    let resp;
    try {
      resp = await fetch('api/samplesheet/read', { method: 'POST', body: form });
    } catch (err) {
      return { fail: true, detail: 'The service could not be reached (' + err.message + '). Check the connection and try again.' };
    }
    const body = await this.answer(resp);
    if (!resp.ok || !body || !body.sheet) return { fail: true, detail: resp.ok ? 'The service sent an answer this page cannot read.' : this.detail(resp, body) };
    return body;
  }

  async pickSheet(file, idColumn) {
    if (this.state.busy) return;
    this.setState({ busy: 'sheet', error: null, toast: '' });
    const body = await this.readSheet(file, idColumn);
    if (body.fail) {
      this.setState({ busy: '', error: { title: file.name + ' was not read as a sample sheet', detail: body.detail } });
      return;
    }
    this.setState({ busy: '', sheet: { file, name: file.name, idColumn: body.sheet.id_column, info: body } });
  }

  // The batch changed: match its samples to the sheet again.
  async refreshSheet() {
    const sh = this.state.sheet;
    if (!sh) return;
    const body = await this.readSheet(sh.file, sh.idColumn);
    if (body.fail) {
      this.setState({ error: { title: 'The samples were not matched to ' + sh.name + ' again', detail: body.detail } });
      return;
    }
    if (this.state.sheet && this.state.sheet.file === sh.file) this.setState({ sheet: Object.assign({}, this.state.sheet, { info: body }) });
  }

  clearSheet() {
    if (this.state.busy) return;
    this.setState({ sheet: null, error: null, toast: 'The sample sheet was removed; the downloads no longer carry it.' });
  }

  async downloadSheetTemplate() {
    const names = [];
    for (const r of this.state.runs) for (const q of r.port.pos) if (q.role === 'S') names.push(String(q.id == null ? '' : q.id));
    try {
      const resp = await fetch('api/samplesheet/template.xlsx', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ names }) });
      if (!resp.ok) throw new Error(this.detail(resp, await this.answer(resp)));
      this.save(await resp.blob(), this.fileName(resp.headers.get('Content-Disposition')) || 'sample_sheet.xlsx');
    } catch (err) {
      this.setState({ error: { title: 'The sample sheet was not downloaded', detail: String(err && err.message ? err.message : err) } });
    }
  }

  // The sample sheet, and a run's information when one is given, go with a download.
  addMeta(form, run) {
    const sh = this.state.sheet;
    if (sh) {
      form.append('sample_sheet', sh.file, sh.file.name);
      if (sh.idColumn) form.append('id_column', sh.idColumn);
    }
    if (run) {
      const info = this.infoJSON(run);
      if (Object.keys(info).length) form.append('run_info', JSON.stringify(info));
    }
  }

  clearConfig() {
    if (this.state.busy) return;
    if (this.state.runs.length) this.reprocessAll(null);
    else this.setState({ cfg: null, error: null });
  }

  async download(kind, i) {
    const st = this.state;
    if (!st.run || st.busy) return;
    if (kind === 'batch' || kind === 'batchcsv') {
      this.setState({ busy: kind, toast: '' });
      let toast = '';
      try {
        toast = await this.fetchBatch(kind);
      } catch (err) {
        console.error(err);
        toast = 'The combined results could not be saved (' + err.message + ').';
      }
      this.setState({ busy: '', toast });
      return;
    }
    if (i != null && st.runs[i] && st.runs[i] !== st.run) {
      // another run's workbook, with that run's decisions and run date
      const other = st.runs[i];
      this.setState({ busy: kind + ':' + i, toast: '' });
      let toast = '';
      try {
        toast = await this.fetchFile(kind, other, this.configFor(this.decOf(other), other), this.dateOf(other));
      } catch (err) {
        console.error(err);
        toast = 'The workbook could not be saved (' + err.message + ').';
      }
      this.setState({ busy: '', toast });
      return;
    }
    const run = st.run, config = this.configFor(st.dec);
    if (kind === 'config') {
      this.setState({ busy: kind, toast: '' });
      let toast = '';
      try {
        const resp = await fetch('api/config.xlsx', { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ config, name: this.stem(run.name) }) });
        if (!resp.ok) toast = 'The configuration was not written: ' + this.detail(resp, await this.answer(resp));
        else this.save(await resp.blob(), this.fileName(resp.headers.get('Content-Disposition')) || this.names(run).config);
      } catch (err) {
        toast = 'The configuration could not be saved (' + err.message + ').';
      }
      this.setState({ busy: '', toast });
      return;
    }
    this.setState({ busy: kind, toast: '' });
    let toast = '';
    try {
      toast = await this.fetchFile(kind, run, config, st.runDate);
    } catch (err) {
      console.error(err);
      toast = (kind === 'workbook' ? 'The workbook' : 'The results table') + ' could not be saved (' + err.message + ').';
    }
    this.setState({ busy: '', toast });
  }

  // POST the export and the configuration; save the answer under the service's name.
  // Returns the message to show, if any.
  async fetchFile(kind, run, config, runDate) {
    const form = new FormData();
    form.append('file', run.file, run.file.name);
    form.append('config', JSON.stringify(config));
    if (kind === 'workbook' && runDate) form.append('run_date', runDate);
    this.addMeta(form, kind === 'workbook' ? run : null);
    const url = kind === 'workbook' ? 'api/workbook' : 'api/results.csv';
    const what = kind === 'workbook' ? 'The workbook' : 'The results table';
    let resp;
    try {
      resp = await fetch(url, { method: 'POST', body: form });
    } catch (err) {
      return what + ' was not built: the service could not be reached (' + err.message + ').';
    }
    if (!resp.ok) return what + ' was not built: ' + this.detail(resp, await this.answer(resp));
    const blob = await resp.blob();
    const name = this.fileName(resp.headers.get('Content-Disposition'))
      || (kind === 'workbook' ? this.names(run).workbook : this.names(run).csv);
    this.save(blob, name);
    const formulas = kind === 'workbook' && resp.headers.get('X-Isoharness-Values') === 'formulas-only';
    return formulas ? 'This service has no LibreOffice, so ' + name + ' holds the formulas without their values; Excel or LibreOffice computes them when it opens the file.' : '';
  }

  // Post every run's export with its own configuration, its decisions written in, and its run date.
  async fetchBatch(kind) {
    const runs = this.state.runs, form = new FormData();
    for (const r of runs) form.append('files', r.file, r.file.name);
    form.append('configs', JSON.stringify(runs.map((r) => this.configFor(this.decOf(r), r))));
    form.append('run_dates', JSON.stringify(runs.map((r) => this.dateOf(r) || null)));
    form.append('run_infos', JSON.stringify(runs.map((r) => this.infoJSON(r))));
    this.addMeta(form, null);
    if (this.state.cfg) form.append('config_name', this.state.cfg.name);
    let resp;
    try {
      resp = await fetch(kind === 'batch' ? 'api/batch/results.xlsx' : 'api/batch/results.csv', { method: 'POST', body: form });
    } catch (err) {
      return 'The combined results were not built: the service could not be reached (' + err.message + ').';
    }
    if (!resp.ok) return 'The combined results were not built: ' + this.detail(resp, await this.answer(resp));
    const nm = this.batchNames();
    this.save(await resp.blob(), this.fileName(resp.headers.get('Content-Disposition')) || (kind === 'batch' ? nm.xlsx : nm.csv));
    return '';
  }

  batchNames() {
    const n = this.state.runs.length, v = (this.state.run && this.state.run.version) || '';
    const s = 'combined_' + n + '_run' + (n === 1 ? '' : 's') + '_results_ea_irms_v' + v;
    return { xlsx: s + '.xlsx', csv: s + '.csv' };
  }

  // Samples measured in more than one run of the batch, each run on its own decisions: the
  // identifier, without surrounding spaces and without regard to case, as the combined results match it.
  repeats() {
    const groups = new Map();
    this.state.runs.forEach((run, i) => {
      const M = this.model(this.decOf(run), run);
      for (const q of M.P) {
        if (q.role !== 'S' || !String(q.id || '').trim()) continue;
        const key = String(q.id).trim().toLowerCase();
        if (!groups.has(key)) groups.set(key, { name: String(q.id).trim(), runs: new Set(), n: [], c: [], count: 0 });
        const g = groups.get(key);
        g.runs.add(i);
        g.count++;
        if (q.n.fin != null) g.n.push(q.n.fin);
        if (q.c.fin != null) g.c.push(q.c.fin);
      }
    });
    return [...groups.values()].filter((g) => g.runs.size > 1).sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  }

  save(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  fileName(header) {
    const m = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(header || '');
    if (!m) return '';
    try { return decodeURIComponent(m[1]); } catch (err) { return m[1]; }
  }

  // The file stem as the service writes it (isoharness.app._stem).
  stem(name) {
    const base = String(name).replace(/^.*[\\/]/, '');
    const dot = base.lastIndexOf('.');
    const st = (dot > 0 ? base.slice(0, dot) : base).replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^[._]+|[._]+$/g, '');
    return st.slice(0, 100) || 'export';
  }

  names(run) {
    const s = this.stem(run.name), v = run.version;
    return { workbook: s + '_ea_irms_v' + v + '.xlsx', csv: s + '_results_ea_irms_v' + v + '.csv', config: s + '_config.xlsx' };
  }

  // A run date in the file name, when it can be read one way only: YYYYMMDD,
  // YYYY-MM-DD, or MMDDYYYY with a day past the 12th.
  dateFromName(name) {
    const ok = (y, m, d) => y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
    const iso = (y, m, d) => y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    let m = /(?:^|[^0-9])(\d{4})-(\d{2})-(\d{2})(?![0-9])/.exec(name);
    if (m && ok(+m[1], +m[2], +m[3])) return iso(+m[1], +m[2], +m[3]);
    m = /(?:^|[^0-9])(\d{8})(?![0-9])/.exec(name);
    if (!m) return '';
    const s = m[1];
    const a = [+s.slice(0, 4), +s.slice(4, 6), +s.slice(6, 8)], b = [+s.slice(4, 8), +s.slice(0, 2), +s.slice(2, 4)];
    const okA = ok(...a), okB = ok(...b) && b[2] > 12;
    if (okA && !okB) return iso(...a);
    if (okB && !okA) return iso(...b);
    return '';
  }

  // ---------------------------------------------------------------- navigation
  go(s) {
    if (s !== 'upload' && !this.state.run) s = 'upload';
    const hash = '#' + s;
    if (location.hash !== hash) location.hash = hash;
    else this.setState({ screen: s, toast: '' });
  }

  fromHash() {
    const h = location.hash.replace(/^#\/?/, '');
    let s = FLOW.includes(h) ? h : 'upload';
    if (s !== 'upload' && !this.state.run) {
      s = 'upload';
      if (h) history.replaceState(null, '', location.pathname + location.search);
    }
    this.setState({ screen: s, toast: s === this.state.screen ? this.state.toast : '' });
  }

  componentDidMount() {
    window.addEventListener('hashchange', () => this.fromHash());
    // A file dropped anywhere on the upload screen is taken as the export; elsewhere
    // the browser is kept from opening it in place of the page.
    window.addEventListener('dragenter', (e) => {
      if (this.state.screen !== 'upload' || this.state.busy) return;
      this._depth++;
      if (!this.state.drag) this.setState({ drag: true });
      e.preventDefault();
    });
    window.addEventListener('dragleave', () => {
      this._depth = Math.max(0, this._depth - 1);
      if (!this._depth && this.state.drag) this.setState({ drag: false });
    });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      this._depth = 0;
      if (this.state.drag) this.setState({ drag: false });
      const fs = e.dataTransfer && e.dataTransfer.files;
      if (fs && fs.length && this.state.screen === 'upload' && !this.state.busy) this.addExports(fs);
    });
    this.fromHash();
    this.boot();
    this._shown = this.state.screen;
  }

  componentDidUpdate() {
    const s = this.state.run ? this.state.screen : 'upload';
    if (s !== this._shown) {
      this._shown = s;
      if (FLOW.includes(s) && this.state.last !== s) this.setState({ last: s });
      try { window.scrollTo(0, 0); } catch (err) { /* a frame may refuse */ }
    }
    const title = s === 'upload' ? 'isoharness' : NAMES[s] + ' · isoharness';
    if (document.title !== title) document.title = title;
  }

  samples(M) { return M.P.filter((q) => q.role === 'S'); }

  // 1.4.5: every uncertainty shown is 1σ (u); the checks count against k × 1σ (U), named kσ
  kSig(M) { return this.num(M.settings.k) + 'σ'; }

  uSpan(S, el, d) {
    const u = S.map((q) => q[el].u).filter((v) => v != null);
    return u.length ? this.span(Math.min(...u), Math.max(...u), d) : 'none';
  }

  // changes in the samples' values between two runs of the corrections
  shift(Ma, Mb, el) {
    let lo = null, hi = null, n = 0, beyond = 0, big = null;
    for (let i = 0; i < Ma.P.length; i++) {
      const A = Ma.P[i], B = Mb.P[i];
      if (A.role !== 'S') continue;
      const a = A[el].fin, b = B[el].fin;
      if (a == null || b == null) continue;
      const d = b - a;
      n++;
      lo = lo == null ? d : Math.min(lo, d);
      hi = hi == null ? d : Math.max(hi, d);
      if (A[el].U != null && Math.abs(d) > A[el].U) beyond++;
      if (!big || Math.abs(d) > Math.abs(big.d)) big = { d, q: A };
    }
    return { lo, hi, n, beyond, big };
  }

  pctShift(Ma, Mb, key) {
    let lo = null, hi = null, n = 0;
    for (let i = 0; i < Ma.P.length; i++) {
      const A = Ma.P[i], B = Mb.P[i];
      if (A.role !== 'S' || A[key] == null || B[key] == null || A[key] === 0) continue;
      const d = (B[key] / A[key] - 1) * 100;
      n++;
      lo = lo == null ? d : Math.min(lo, d);
      hi = hi == null ? d : Math.max(hi, d);
    }
    return { lo, hi, n };
  }

  // ---------------------------------------------------------------- formatting
  fx(v, d) {
    if (v == null || !isFinite(v)) return '';
    const s = Math.abs(v).toFixed(d);
    return (v < 0 && +s !== 0 ? '−' : '') + s;
  }

  fs(v, d) {
    if (v == null || !isFinite(v)) return '';
    const s = Math.abs(v).toFixed(d);
    return +s === 0 ? s : (v < 0 ? '−' : '+') + s;
  }

  // a setting as the configuration wrote it: 0.05, 2, 2.5
  num(v) { return v == null ? '' : String(+v); }

  peq(p) { return p == null ? 'no p' : p < 0.001 ? 'p < 0.001' : 'p = ' + p.toFixed(3); }

  pv(p) { return p == null ? '' : p < 0.001 ? '< 0.001' : p.toFixed(3); }

  fdate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return iso || '';
    const mo = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+m[2] - 1];
    return +m[3] + ' ' + mo + ' ' + m[1];
  }

  span(lo, hi, d, signed) {
    const f = (v) => (signed ? this.fs(v, d) : this.fx(v, d));
    if (lo == null) return '';
    return Math.abs(hi - lo) < Math.pow(10, -d) / 2 ? f(lo) : f(lo) + ' to ' + f(hi);
  }

  bytes(n) {
    if (n < 1024) return n + ' bytes';
    if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
    return (n / 1024 / 1024).toFixed(1) + ' MB';
  }

  count(n, one, many) { return n + ' ' + (n === 1 ? one : many || one + 's'); }

  words(n) { return n >= 0 && n < WORDS.length ? WORDS[n] : String(n); }

  times(m) { return m === 2 ? 'twice' : Number.isInteger(m) && m > 2 && m < WORDS.length ? WORDS[m] + ' times' : this.num(m) + ' times'; }

  andList(items) {
    const a = items.map(String);
    if (a.length < 3) return a.join(' and ');
    return a.slice(0, -1).join(', ') + ', and ' + a[a.length - 1];
  }

  rowList(rows) {
    return rows.length <= 6 ? this.andList(rows) : rows.slice(0, 3).join(', ') + ', and ' + (rows.length - 3) + ' more';
  }

  who(q) { return (q.id || 'Unnamed') + ' (' + q.r + ')'; }

  union(a, b) { return [...new Set(a.concat(b))].sort((x, y) => x - y); }

  cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  // The material of a role's standards, from the standards table (acetanilide).
  material(M, role) {
    const D = this.data(), names = new Set();
    for (const q of M.P) {
      if (q.role !== role || q.si == null) continue;
      const m = D.std[q.si] && D.std[q.si][6];
      if (m) names.add(String(m).split(',')[0].trim());
    }
    return names.size === 1 ? [...names][0] : null;
  }

  anchorGroups(M) {
    const D = this.data(), g = new Map();
    for (const q of M.P) {
      if (q.role !== 'A') continue;
      if (!g.has(q.si)) g.set(q.si, { si: q.si, name: D.std[q.si][0], n: 0, acc15: q.acc15, acc13: q.acc13, uacc15: q.uacc15, uacc13: q.uacc13 });
      g.get(q.si).n++;
    }
    return [...g.values()];
  }

  roles() {
    return {
      S: { one: 'Sample', many: 'Samples', bg: '#8796A5', dk: '#8193A5' },
      A: { one: 'Anchor', many: 'Anchors', bg: '#13212B', dk: '#F1F5F8' },
      L: { one: 'Linearity', many: 'Linearity', bg: '#56B4E9', dk: '#56B4E9' },
      D: { one: 'Drift monitor', many: 'Drift monitors', bg: '#E69F00', dk: '#E69F00' },
      B: { one: 'Blank', many: 'Blanks', bg: '#AEB9C4', dk: '#5A6E82' },
      X: { one: 'Excluded', many: 'Excluded', bg: '#D3DBE2', dk: '#33475B' },
      K: { one: 'Check', many: 'Checks', bg: '#CC79A7', dk: '#CC79A7' },
    };
  }

  col(el) { return el === 'n' ? '#0072B2' : '#D55E00'; }

  tint(el) { return el === 'n' ? '#E3EEF7' : '#FBEADF'; }

  // ---------------------------------------------------------------- charts
  ticks(min, max, n) {
    const span = max - min;
    if (!(span > 0)) return [min];
    const raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw))), e = raw / mag;
    const step = (e >= 7.5 ? 10 : e >= 3.5 ? 5 : e >= 1.5 ? 2 : 1) * mag;
    const out = [], start = Math.ceil(min / step - 1e-9) * step;
    for (let k = 0; k < 60; k++) {
      const v = start + k * step;
      if (v > max + step * 1e-9) break;
      out.push(Math.abs(v) < step * 1e-9 ? 0 : v);
    }
    return out;
  }

  tdec(t) {
    if (t.length < 2) return 1;
    return Math.max(0, Math.ceil(-Math.log10(t[1] - t[0]) - 1e-9));
  }

  dom(vals, pad, n) {
    const v = vals.filter((q) => q != null && isFinite(q));
    if (!v.length) return { min: 0, max: 1 };
    let lo = Math.min(...v), hi = Math.max(...v);
    if (hi - lo < 1e-9) { lo -= 0.5; hi += 0.5; }
    const p = (hi - lo) * (pad == null ? 0.08 : pad);
    lo -= p; hi += p;
    const t = this.ticks(lo, hi, n || 5);
    const st = t.length > 1 ? t[1] - t[0] : hi - lo;
    return { min: Math.floor(lo / st + 1e-9) * st, max: Math.ceil(hi / st - 1e-9) * st };
  }

  // The x axis of a chart against position: the run's first to last row.
  rowAxis(M) {
    const rs = M.P.map((q) => q.r), lo = Math.min(...rs), hi = Math.max(...rs), span = hi - lo;
    const step = span <= 80 ? 10 : span <= 160 ? 20 : span <= 400 ? 50 : 100;
    const ticks = [lo];
    for (let v = Math.ceil((lo + 1) / step) * step; v <= hi; v += step) if (v - lo >= step * 0.6) ticks.push(v);
    return { min: lo - 1, max: hi + 1, ticks };
  }

  side(ax, x) { return x > ax.min + 0.8 * (ax.max - ax.min) ? 'left' : 'right'; }

  swatch(kind, color) {
    const dot = (r) => 'M' + (11 - r) + ' 6a' + r + ' ' + r + ' 0 1 0 ' + 2 * r + ' 0a' + r + ' ' + r + ' 0 1 0 ' + -2 * r + ' 0Z';
    if (kind === 'dot') return { d: dot(4.5), f: color, s: '#FFFFFF', w: 1, da: '' };
    if (kind === 'hollow') return { d: dot(4.2), f: '#FFFFFF', s: color, w: 1.8, da: '' };
    if (kind === 'ring') return { d: dot(4.6), f: 'none', s: '#B23A00', w: 1.8, da: '' };
    if (kind === 'small') return { d: dot(2.6), f: color, s: 'none', w: 0, da: '' };
    if (kind === 'band') return { d: 'M1 1H21V11H1Z', f: color, s: 'none', w: 0, da: '' };
    if (kind === 'dash') return { d: 'M1 6H21', f: 'none', s: color, w: 2, da: '4 3' };
    return { d: 'M1 6H21', f: 'none', s: color, w: 2, da: '' };
  }

  chart(o) {
    const W = o.W || 600, H = o.H || 320;
    const m = Object.assign({ l: 50, r: 14, t: 12, b: 30 }, o.m || {});
    const pw = W - m.l - m.r, ph = H - m.t - m.b;
    const xr = o.x, yr = o.y;
    const X = (v) => m.l + (v - xr.min) / (xr.max - xr.min) * pw;
    const Y = (v) => m.t + (1 - (v - yr.min) / (yr.max - yr.min)) * ph;
    const cx = (v) => Math.max(m.l, Math.min(W - m.r, X(v)));
    const cy = (v) => Math.max(m.t, Math.min(H - m.b, Y(v)));
    const r1 = (v) => Math.round(v * 10) / 10;
    const L = () => ({ d: '', f: 'none', s: 'none', w: 0, da: '', o: 1 });
    const S = [];
    for (let i = 0; i < 12; i++) S.push(L());
    const ok = (p) => p[0] != null && p[1] != null && isFinite(p[0]) && isFinite(p[1]);
    const rect = (b) => {
      const x0 = cx(b.x0 == null ? xr.min : b.x0), x1 = cx(b.x1 == null ? xr.max : b.x1);
      const y0 = cy(b.y1 == null ? yr.max : b.y1), y1 = cy(b.y0 == null ? yr.min : b.y0);
      return x1 > x0 && y1 > y0 ? 'M' + r1(x0) + ' ' + r1(y0) + 'H' + r1(x1) + 'V' + r1(y1) + 'H' + r1(x0) + 'Z' : '';
    };
    const poly = (pts) => pts.filter(ok).map((p, i) => (i ? 'L' : 'M') + r1(cx(p[0])) + ' ' + r1(cy(p[1]))).join('');
    const dot = (x, y, r) => 'M' + r1(x - r) + ' ' + r1(y) + 'a' + r + ' ' + r + ' 0 1 0 ' + 2 * r + ' 0a' + r + ' ' + r + ' 0 1 0 ' + -2 * r + ' 0Z';
    if (o.bandA) S[0] = { d: o.bandA.list.map(rect).join(''), f: o.bandA.fill, s: 'none', w: 0, da: '', o: 1 };
    if (o.bandB) S[1] = { d: o.bandB.list.map(rect).join(''), f: o.bandB.fill, s: 'none', w: 0, da: '', o: 1 };
    const xt = xr.ticks || this.ticks(xr.min, xr.max, xr.n || 6);
    const yt = yr.ticks || this.ticks(yr.min, yr.max, yr.n || 5);
    let g = '';
    for (const v of yt) if (v >= yr.min - 1e-9 && v <= yr.max + 1e-9) g += 'M' + m.l + ' ' + r1(Y(v)) + 'H' + (W - m.r);
    S[2] = { d: g, f: 'none', s: '#E1E7EC', w: 1, da: '', o: 1 };
    S[3] = { d: 'M' + m.l + ' ' + m.t + 'V' + (H - m.b) + 'H' + (W - m.r), f: 'none', s: '#8A99A5', w: 1, da: '', o: 1 };
    const R = o.r || 5;
    const pts = (o.pts || []).filter((p) => ok([p.x, p.y]));
    const small = pts.filter((p) => p.kind === 'small');
    if (small.length) S[4] = { d: small.map((p) => dot(cx(p.x), cy(p.y), 2.6)).join(''), f: o.smallColor || '#8D9AA7', s: 'none', w: 0, da: '', o: 1 };
    if (o.rug) {
      S[5] = { d: o.rug.xs.filter((v) => v != null && isFinite(v)).map((v) => 'M' + r1(cx(v)) + ' ' + (H - m.b) + 'v-9').join(''), f: 'none', s: o.rug.stroke, w: 1.4, da: '', o: 0.75 };
    } else if (o.errs) {
      S[5] = { d: o.errs.segs.filter((s) => ok([s[0], s[1]]) && ok([s[2], s[3]])).map((s) => 'M' + r1(cx(s[0])) + ' ' + r1(cy(s[1])) + 'L' + r1(cx(s[2])) + ' ' + r1(cy(s[3]))).join(''), f: 'none', s: o.errs.stroke, w: 1.2, da: '', o: 0.8 };
    }
    if (o.lineB) S[6] = { d: o.lineB.paths.map(poly).join(''), f: 'none', s: o.lineB.stroke, w: o.lineB.w || 1.5, da: o.lineB.dash || '', o: 1 };
    if (o.lineA) S[7] = { d: o.lineA.paths.map(poly).join(''), f: 'none', s: o.lineA.stroke, w: o.lineA.w || 2.2, da: o.lineA.dash || '', o: 1 };
    const fill = pts.filter((p) => p.kind === 'fill'), hol = pts.filter((p) => p.kind === 'hollow');
    if (fill.length) S[8] = { d: fill.map((p) => dot(cx(p.x), cy(p.y), R)).join(''), f: o.color, s: '#FFFFFF', w: 1.2, da: '', o: 1 };
    if (hol.length) S[9] = { d: hol.map((p) => dot(cx(p.x), cy(p.y), R - 0.4)).join(''), f: '#FFFFFF', s: o.color, w: 1.8, da: '', o: 1 };
    const rings = pts.filter((p) => p.ring);
    if (rings.length) S[10] = { d: rings.map((p) => dot(cx(p.x), cy(p.y), R + 4.5)).join(''), f: 'none', s: '#B23A00', w: 1.8, da: '', o: 1 };
    const sel = pts.filter((p) => p.sel);
    if (sel.length) S[11] = { d: sel.map((p) => dot(cx(p.x), cy(p.y), R + 8)).join(''), f: 'none', s: '#13212B', w: 2, da: '', o: 1 };
    const xf = xr.fmt || ((v) => this.fx(v, this.tdec(xt)));
    const yf = yr.fmt || ((v) => this.fx(v, this.tdec(yt)));
    const pc = (v, t) => Math.round(v / t * 10000) / 100;
    const xl = xt.filter((v) => v >= xr.min - 1e-9 && v <= xr.max + 1e-9).map((v) => ({ x: pc(X(v), W), y: pc(H - m.b, H), t: xf(v) }));
    const yl = yt.filter((v) => v >= yr.min - 1e-9 && v <= yr.max + 1e-9).map((v) => ({ x: pc(m.l, W), y: pc(Y(v), H), t: yf(v) }));
    const hits = pts.filter((p) => p.pick).map((p) => ({ x: pc(cx(p.x), W), y: pc(cy(p.y), H), label: p.label, pick: p.pick }));
    // Two notes on neighboring points would print over each other: put the higher one
    // above its point and the lower one below.
    const placed = (o.notes || []).filter((n) => ok([n.x, n.y])).map((n) => Object.assign({}, n, { px: cx(n.x), py: cy(n.y) }))
      .sort((a, b) => a.px - b.px);
    for (let i = 1; i < placed.length; i++) {
      const a = placed[i - 1], b = placed[i];
      if (Math.abs(b.px - a.px) < 40 && Math.abs(b.py - a.py) < 16) {
        const [hi, lo] = a.py <= b.py ? [a, b] : [b, a];
        hi.at = 'above';
        lo.at = 'below';
      }
    }
    const notes = placed.map((n) => ({
      x: pc(cx(n.x), W), y: pc(cy(n.y), H), t: n.t, c: n.c || '#3B4A55',
      tf: n.at === 'left' ? 'translate(calc(-100% - 10px), -50%)' : n.at === 'below' ? 'translate(-50%, 10px)'
        : n.at === 'above' ? 'translate(-50%, calc(-100% - 10px))' : 'translate(10px, -50%)',
    }));
    const legend = (o.legend || []).map((l) => Object.assign({ label: l.label }, this.swatch(l.kind, l.color)));
    return {
      title: o.title, sub: o.sub || '', hasBadge: !!o.badge, badge: o.badge || { text: '', bg: '#FFFFFF', fg: '#13212B' },
      vb: '0 0 ' + W + ' ' + H, ar: W + ' / ' + H,
      s0: S[0], s1: S[1], s2: S[2], s3: S[3], s4: S[4], s5: S[5], s6: S[6], s7: S[7], s8: S[8], s9: S[9], s10: S[10], s11: S[11],
      xl, yl, hits, notes, legend, hasLegend: legend.length > 0,
      xLabel: o.xLabel || '', yLabel: o.yLabel || '', full: !o.empty, empty: !!o.empty, emptyText: o.empty || '',
    };
  }

  // A chart with nothing to draw, saying why.
  none(title, color, text) {
    return this.chart({ title, color, x: { min: 0, max: 1 }, y: { min: 0, max: 1 }, empty: text });
  }

  badge(on, text) {
    return on ? { text: text || 'Applied', bg: '#DDF2EA', fg: '#0B6E4F' } : { text: text || 'Not applied', bg: '#EEF1F4', fg: '#3B4A55' };
  }

  status(s) {
    if (s === 'ok') return { label: 'OK', bg: '#DDF2EA', fg: '#0B6E4F', icon: 'M5 10.5l3.2 3.2L15 7' };
    if (s === 'note') return { label: 'Note', bg: '#FBEFD6', fg: '#8A5A00', icon: 'M10 9v5M10 6.2v.1' };
    if (s === 'info') return { label: 'Measured', bg: '#EEF1F4', fg: '#3B4A55', icon: 'M10 10v.1' };
    return { label: 'Caution', bg: '#FADFD2', fg: '#A33A00', icon: 'M10 5.5v5.5M10 14v.1' };
  }

  // ---------------------------------------------------------------- shared step parts
  iso(el) { return el === 'n' ? 'δ¹⁵N' : 'δ¹³C'; }

  ftext(s) {
    return s.replace(/d15N/g, 'δ¹⁵N').replace(/d13C/g, 'δ¹³C').replace(/(\d)x the blank/, '$1× the blank');
  }

  chip(s) {
    const t = [['N below linearity range', 'N2 small'], ['N above linearity range', 'N2 large'], ['C below linearity range', 'CO2 small'],
      ['C above linearity range', 'CO2 large'], ['d15N outside anchor range', 'δ¹⁵N extrapolated'],
      ['d13C outside anchor range', 'δ¹³C extrapolated'], ['possible carryover from an oversized sample', 'Carryover'],
      ['ran before the background settled', 'Unsettled'], ['no N peak', 'No N2 peak'], ['no C peak', 'No CO2 peak'],
      ['CO2 peak rejected as a reference pulse', 'CO2 was a pulse']];
    for (const [a, b] of t) if (s === a) return b;
    const m = /^(N|C) peak only ([\d.]+)x the blank$/.exec(s);
    if (m) return (m[1] === 'N' ? 'N2 ' : 'CO2 ') + m[2] + '× blank';
    return this.ftext(s);
  }

  seg(label, key, value, opts, H, lines) {
    return {
      label,
      opts: opts.map((o) => {
        const v = typeof o === 'string' ? o : o.v, t = typeof o === 'string' ? o : o.t, on = value === v;
        return { t, pick: () => H.set(key, v), bg: on ? '#13212B' : 'transparent', fg: on ? '#FFFFFF' : '#3B4A55', pressed: on ? 'true' : 'false' };
      }),
      lines: (lines || []).map((l) => ({ t: l.t, dot: l.on ? '#0B8A62' : '#8A99A5' })),
    };
  }

  focus(scr) {
    if (scr === 'inst') return () => true;
    if (scr === 'blank') return (q) => q.role === 'B' || q.flags.some((f) => / the blank$/.test(f));
    if (scr === 'lin') return (q) => q.role === 'L' || q.flags.some((f) => /linearity range/.test(f));
    if (scr === 'drift') return (q) => q.role === 'D';
    if (scr === 'cal') return (q) => q.role === 'A' || q.role === 'S';
    if (scr === 'elem') return (q) => q.role === 'D' || ((q.role === 'A' || q.role === 'L') && q.amt);
    if (scr === 'unc' || scr === 'results') return (q) => q.role === 'S' || q.role === 'A';
    if (scr === 'qc') return (q) => q.carry || q.over || q.role === 'D';
    return () => true;
  }

  // The position each screen opens on: the first one the screen is about.
  firstSel(M, scr) {
    const first = (f) => { const q = M.P.find(f); return q ? q.r : null; };
    const by = { inst: () => true, blank: (q) => q.role === 'B', lin: (q) => q.role === 'L', drift: (q) => q.role === 'D', elem: (q) => q.role === 'D',
      cal: (q) => q.role === 'A', qc: (q) => q.carry || q.over || q.role === 'D' };
    const r = by[scr] ? first(by[scr]) : null;
    return r != null ? r : first((q) => q.role === 'S') != null ? first((q) => q.role === 'S') : M.P[0].r;
  }

  ribbon(M, scr, sel, H) {
    const R = this.roles(), f = this.focus(scr), hi = Math.log10(200);
    return M.P.map((q) => ({
      h: q.aN == null || q.aN <= 1 ? 3 : Math.max(3, Math.round(4 + 52 * Math.min(1, Math.log10(q.aN) / hi))),
      bg: R[q.role].bg, op: f(q) ? 1 : 0.22, sbg: q.r === sel ? '#D5DEE7' : 'transparent',
      label: 'Position ' + q.r + ', ' + q.id + ', ' + R[q.role].one.toLowerCase() + (q.aN != null ? ', N2 peak ' + this.fx(q.aN, 1) + ' Vs' : ', no N2 peak'),
      pick: () => H.pick(q.r),
    }));
  }

  ribbonLegend(M) {
    const R = this.roles(), out = [];
    for (const k of ['S', 'A', 'L', 'D', 'B', 'X', 'K']) {
      const n = M.P.filter((q) => q.role === k).length;
      if (n) out.push({ bg: R[k].bg, t: R[k].many + ' ' + n });
    }
    return out;
  }

  inspect(M, row, scr) {
    const q = M.P.find((p) => p.r === row) || M.P[0];
    const R = this.roles();
    const st = [['Raw, as exported', 'raw'], ['After blank', 'blank'], ['After linearity', 'lin'], ['After drift', 'drift'], ['Calibrated', 'fin']];
    const chain = st.map(([label, key], i) => {
      const pk = i ? st[i - 1][1] : null;
      const dl = (el) => (!pk || q[el][key] == null || q[el][pk] == null ? '' : key === 'fin' ? 'scale' : this.fs(q[el][key] - q[el][pk], 3));
      return { stage: label, n: this.fx(q.n[key], 3), dn: dl('n'), c: this.fx(q.c[key], 3), dc: dl('c') };
    });
    const facts = [
      { k: 'Role', v: R[q.role].one },
      { k: 'Amount', v: q.amt ? this.fx(q.amt, 3) + ' mg' : 'not weighed' },
      { k: 'N2 peak', v: q.aN != null ? this.fx(q.aN, 2) + ' Vs' : 'none' },
      { k: 'CO2 peak', v: q.aC != null ? this.fx(q.aC, 2) + ' Vs' : 'none' },
      { k: 'BGD 28', v: q.bgd != null ? this.fx(q.bgd, 1) + ' mV' : 'none' },
    ];
    if (scr === 'inst') {
      facts.push(
        { k: 'BGD 28, reference', v: q.rbgd != null ? this.fx(q.rbgd, 1) + ' mV' : 'none' },
        { k: 'N2 pulse', v: q.pN != null ? this.fx(q.pN, 3) + '‰' : 'none' },
        { k: 'CO2 pulse', v: q.pC != null ? this.fx(q.pC, 3) + '‰' : 'none' },
        { k: 'N2 reference peak', v: q.raN != null ? this.fx(q.raN, 2) + ' Vs' : 'none' },
        { k: 'CO2 reference peak', v: q.raC != null ? this.fx(q.raC, 2) + ' Vs' : 'none' });
    }
    const tail = [
      { k: '1σ, δ¹⁵N', v: q.n.u != null ? this.fx(q.n.u, 3) + '‰' : '' },
      { k: '1σ, δ¹³C', v: q.c.u != null ? this.fx(q.c.u, 3) + '‰' : '' },
      { k: '%N measured, reported', v: q.pctNm != null ? this.fx(q.pctNm, 3) + ', ' + this.fx(q.pctN, 3) : '' },
      { k: '%C measured, reported', v: q.pctCm != null ? this.fx(q.pctCm, 2) + ', ' + this.fx(q.pctC, 2) : '' },
      { k: 'C:N, molar', v: q.cn != null ? this.fx(q.cn, 2) : '' },
    ].filter((t) => t.v !== '');
    return {
      title: q.id, sub: 'Position ' + q.r + ' of ' + M.P.length, facts, chain, tail,
      flags: q.flags.map((f) => ({ t: this.ftext(f) })), hasFlags: q.flags.length > 0, noFlags: q.flags.length === 0,
    };
  }

  // ---------------------------------------------------------------- the instrument
  // Step 1 reads what the mass spectrometer did before any correction: the working-gas
  // pulses, the reference peaks' size, and the background at mass 28. Nothing here
  // changes a value. The start of the run is looked at on its own, over the same five
  // positions the engine uses to scale the background's settling.
  pulseStart(M, k) {
    const pts = M.P.filter((q) => q[k] != null).sort((a, b) => a.r - b.r);
    const start = pts.slice(0, 5), rest = pts.slice(5);
    if (rest.length < 3) return null;
    const mean = rest.reduce((t, q) => t + q[k], 0) / rest.length;
    return { start, rest, mean, sdRest: engine.sdev(rest.map((q) => q[k])), devs: start.map((q) => q[k] - mean) };
  }

  settleSpan(M) {
    const D = this.data(), ax = this.rowAxis(M), sf = D.settle.settled_from;
    const off = !this.state.run.config.settled_threshold;
    return { sf, off, first: ax.min + 1, last: ax.max - 1, never: !off && sf > ax.max - 1, late: !off && sf > ax.min + 1 };
  }

  instChecks(M) {
    const D = this.data(), out = [], sp = this.settleSpan(M);
    const roleWords = { A: ['anchor analysis', 'anchor analyses'], L: ['linearity standard', 'linearity standards'],
      D: ['drift monitor', 'drift monitors'], K: ['check analysis', 'check analyses'], S: ['sample', 'samples'] };
    // the reference gas
    for (const [el, gas, k] of [['n', 'N2', 'pN'], ['c', 'CO2', 'pC']]) {
      const r = D.refgas[el];
      if (!r || !r.n) { out.push({ s: 'note', title: 'Reference gas, ' + gas, t: 'No ' + gas + ' reference pulse besides the designated one was exported, so the working gas was not followed.' }); continue; }
      const st = this.pulseStart(M, k);
      const high = r.sd != null && r.sd >= REF_SD_LIMIT;
      out.push({ s: r.trend || high ? 'note' : 'ok', title: 'Reference gas, ' + gas,
        t: r.n + ' pulses that Isodat did not designate read with SD ' + this.fx(r.sd, 3) + '‰'
          + (st ? ' over the run and ' + this.fx(st.sdRest, 3) + '‰ without the first five positions' : '')
          + (r.sd != null ? ', ' + (high ? 'at or above' : 'under') + ' the ' + this.num(REF_SD_LIMIT) + '‰ limit the FIRMS guide (Dunn and Carter, 2018, section 4.3.2) sets for working-gas pulses' : '') + '; '
          + (r.fit ? (r.trend ? 'they trend with position (' : 'no trend with position (') + this.peq(r.fit.p) + ').' : 'too few to test for a trend.')
          + (st ? ' The first five read ' + this.andList(st.devs.map((v) => this.fs(v, 3))) + '‰ against the mean of the rest.' : '') });
    }
    // the reference peaks' size
    const size = [];
    for (const [gas, k] of [['N2', 'raN'], ['CO2', 'raC']]) {
      const v = M.P.map((q) => q[k]).filter((x) => x != null).sort((a, b) => a - b);
      if (v.length < 3) continue;
      const med = v.length % 2 ? v[(v.length - 1) / 2] : (v[v.length / 2 - 1] + v[v.length / 2]) / 2;
      size.push(gas + ' ' + this.fx(v[0], 2) + ' to ' + this.fx(v[v.length - 1], 2) + ' Vs (' + this.fs((v[0] / med - 1) * 100, 1) + '% to ' + this.fs((v[v.length - 1] / med - 1) * 100, 1) + '% of its median)');
    }
    if (size.length) out.push({ s: 'info', title: 'Reference peak size', t: 'The area of the peak Isodat designates in each analysis ran ' + size.join('; ') + '.' });
    // the background at the start, and who ran before it settled
    const before = [];
    if (sp.late && !sp.never) {
      for (const code of ['A', 'L', 'D', 'K', 'S']) {
        const all = M.P.filter((q) => q.role === code), b = all.filter((q) => q.r < sp.sf).length;
        if (b) before.push((b === all.length ? (b > 1 ? 'all ' : '') + b : b + ' of ' + all.length) + ' ' + (all.length === 1 ? roleWords[code][0] : roleWords[code][1]));
      }
    }
    out.push({ s: sp.late ? 'note' : 'ok', title: 'Background at the start',
      t: sp.off ? 'The configuration turns the settling check off.' : sp.never ? 'The background never settled, so every position ran before it did.'
        : sp.late ? this.count(sp.sf - sp.first, 'position', 'positions') + ' ran before the background settled at position ' + sp.sf + (before.length ? ', among them ' + this.andList(before) : '') + '.'
          : 'Settled from the first position.' });
    // what the drift monitors did while it settled, before any correction
    if (sp.late && !sp.never) {
      const mons = M.P.filter((q) => q.role === 'D' && q.dN != null);
      const pre = mons.filter((q) => q.r < sp.sf), post = mons.filter((q) => q.r >= sp.sf);
      const fit = (g) => (g.length >= 3 ? engine.fit(g.map((q) => q.r), g.map((q) => q.dN)) : null);
      const fb = fit(pre), fa = fit(post);
      const part = (f, g, where) => (f ? where + ', ' + this.fs(f.slope, 3) + '‰ per position (' + this.peq(f.p) + ', ' + this.count(g.length, 'monitor') + ')' : where + ', ' + (g.length ? 'only ' + this.count(g.length, 'monitor') : 'no monitor'));
      if (mons.length) out.push({ s: 'info', title: 'Drift monitors while the background settled',
        t: 'Raw ' + this.iso('n') + ' of the drift monitors against position: ' + part(fb, pre, 'before position ' + sp.sf) + '; ' + part(fa, post, 'from position ' + sp.sf) + '.' });
    }
    const mb = D.monBgd;
    if (mb && mb.n) {
      const f = mb.fit;
      out.push({ s: mb.trend ? 'note' : 'ok', title: 'Background under the drift monitors',
        t: 'BGD 28 ran ' + this.fx(mb.min, 1) + ' to ' + this.fx(mb.max, 1) + ' mV across the ' + mb.n + ' monitors'
          + (f ? (mb.trend ? ', ' + (f.slope > 0 ? 'rising' : 'falling') + ' with position (' : ', with no trend (') + this.peq(f.p) + ').' : '.') });
    }
    const C = D.carry;
    if (C.threshold == null) out.push({ s: 'note', title: 'Carryover', t: 'The configuration turns the carryover check off.' });
    else {
      const cf = C.flagged || [];
      out.push({ s: cf.length ? 'note' : 'ok', title: 'Carryover',
        t: !C.oversized ? 'No position was larger than the size series, so none was watched for carryover.'
          : this.count(C.oversized, 'position was', 'positions were') + ' larger than the size series. ' + (cf.length
            ? 'The background under the next peak rose by more than ' + this.num(C.threshold) + ' mV at ' + this.andList(cf.map((r) => this.who(M.P.find((q) => q.r === r)))) + ', by up to ' + this.fx(C.maxRise, 1) + ' mV.'
            : 'No background rise above ' + this.num(C.threshold) + ' mV followed them.') });
    }
    return out.map((c) => Object.assign({}, c, this.status(c.s)));
  }

  vInst(M, d, sel, H) {
    const D = this.data(), charts = [], ax = this.rowAxis(M), sp = this.settleSpan(M);
    const before = sp.late && !sp.never ? { list: [{ x0: ax.min, x1: sp.sf - 0.5 }], fill: '#E6EBF0' } : null;
    for (const [el, gas, k] of [['n', 'N2', 'pN'], ['c', 'CO2', 'pC']]) {
      const r = D.refgas[el], col = this.col(el), pts = M.P.filter((q) => q[k] != null), title = gas + ' reference pulses';
      if (!pts.length) { charts.push(this.none(title, col, 'No ' + gas + ' reference pulse besides the designated one was exported, so the working gas cannot be followed.')); continue; }
      const st = this.pulseStart(M, k), early = new Set(st ? st.start.map((q) => q.r) : []), f = r.fit, lo = ax.min + 1, hi = ax.max - 1;
      const mean = st ? st.mean : r.mean;
      charts.push(this.chart({
        title, sub: 'The first pulse Isodat did not designate, ' + r.n + ' positions; SD ' + this.fx(r.sd, 3) + '‰' + (st ? ', ' + this.fx(st.sdRest, 3) + '‰ without the first five' : ''),
        badge: r.trend ? { text: 'Trend, ' + this.peq(f && f.p), bg: '#FBEFD6', fg: '#8A5A00' } : { text: 'No trend', bg: '#DDF2EA', fg: '#0B6E4F' }, color: col, r: 3.6,
        x: ax, y: this.dom(pts.map((q) => q[k]), 0.1), bandA: before,
        lineB: { paths: [[[ax.min, mean], [ax.max, mean]]], stroke: col, w: 1.2, dash: '5 4' },
        lineA: r.trend && f ? { paths: [[[lo, f.intercept + f.slope * lo], [hi, f.intercept + f.slope * hi]]], stroke: col, w: 2 } : null,
        pts: pts.map((q) => ({ x: q.r, y: q[k], kind: early.has(q.r) ? 'hollow' : 'fill', sel: q.r === sel,
          label: q.id + ', position ' + q.r + ': ' + this.fx(q[k], 3) + '‰' + (early.has(q.r) ? ', one of the first five' : ''), pick: () => H.pick(q.r) })),
        legend: [{ kind: 'dot', color: col, label: 'Pulse' }].concat(st ? [{ kind: 'hollow', color: col, label: 'First five positions' }] : [])
          .concat([{ kind: 'dash', color: col, label: (st ? 'Mean of the rest ' : 'Mean ') + this.fx(mean, 3) + '‰' }])
          .concat(before ? [{ kind: 'band', color: '#E6EBF0', label: 'Before the background settled' }] : []),
        xLabel: 'Position in the run', yLabel: this.iso(el) + ' of the pulse (‰)',
      }));
    }
    // the designated reference peaks' area, each as a percentage of its run median
    const rel = (k) => {
      const v = M.P.map((q) => q[k]).filter((x) => x != null).sort((a, b) => a - b);
      if (v.length < 3) return null;
      const med = v.length % 2 ? v[(v.length - 1) / 2] : (v[v.length / 2 - 1] + v[v.length / 2]) / 2;
      return M.P.filter((q) => q[k] != null).sort((a, b) => a.r - b.r).map((q) => [q.r, (q[k] / med - 1) * 100]);
    };
    const rN = rel('raN'), rC = rel('raC');
    if (rN || rC) {
      const ys = [].concat(rN || [], rC || []).map((p) => p[1]).concat([0]);
      charts.push(this.chart({
        title: 'Reference peak size', sub: 'Area of the peak Isodat designates in each analysis, against its median in this run',
        color: this.col('n'), x: ax, y: this.dom(ys, 0.15), bandA: before,
        lineA: rN ? { paths: [rN], stroke: this.col('n'), w: 2 } : null,
        lineB: rC ? { paths: [rC], stroke: this.col('c'), w: 2, dash: '6 3' } : null,
        legend: [].concat(rN ? [{ kind: 'line', color: this.col('n'), label: 'N2' }] : [], rC ? [{ kind: 'dash', color: this.col('c'), label: 'CO2' }] : [],
          before ? [{ kind: 'band', color: '#E6EBF0', label: 'Before the background settled' }] : []),
        xLabel: 'Position in the run', yLabel: 'Area against the median (%)',
      }));
    }
    const pts = M.P.filter((q) => q.bgd != null || q.rbgd != null), flagged = (D.carry.flagged || []).length;
    const vals = [].concat(M.P.map((q) => q.bgd), M.P.map((q) => q.rbgd));
    charts.push(pts.length ? this.chart({
      title: 'Background at m/z 28', sub: 'BGD 28 under each sample N2 peak and under the reference peak' + (sp.late && !sp.never ? '; settled from position ' + sp.sf : ''),
      badge: flagged ? { text: flagged + ' carryover', bg: '#FBEFD6', fg: '#8A5A00' } : sp.late ? { text: 'Settled late', bg: '#FBEFD6', fg: '#8A5A00' } : { text: 'Steady', bg: '#DDF2EA', fg: '#0B6E4F' },
      color: '#3B4A55', r: 3.6, smallColor: '#8D9AA7',
      x: ax, y: this.dom(vals, 0.1), bandA: before,
      pts: M.P.filter((q) => q.rbgd != null).map((q) => ({ x: q.r, y: q.rbgd, kind: 'small' }))
        .concat(M.P.filter((q) => q.bgd != null).map((q) => ({ x: q.r, y: q.bgd, kind: q.over ? 'hollow' : 'fill', ring: q.carry, sel: q.r === sel,
          label: q.id + ', position ' + q.r + ': ' + this.fx(q.bgd, 1) + ' mV under the sample peak' + (q.rbgd != null ? ', ' + this.fx(q.rbgd, 1) + ' mV under the reference' : '') + (q.over ? ', larger than the size series' : '') + (q.carry ? ', possible carryover' : ''), pick: () => H.pick(q.r) }))),
      legend: [{ kind: 'dot', color: '#3B4A55', label: 'Under the sample peak' }, { kind: 'small', color: '#8D9AA7', label: 'Under the reference peak' },
        { kind: 'hollow', color: '#3B4A55', label: 'Larger than the series' }, { kind: 'ring', color: '#3B4A55', label: 'Carryover flag' }]
        .concat(before ? [{ kind: 'band', color: '#E6EBF0', label: 'Before it settled' }] : []),
      xLabel: 'Position in the run', yLabel: 'BGD 28 (mV)',
    }) : this.none('Background at m/z 28', '#3B4A55', 'No position has an N2 peak, so there is no background to follow.'));
    return { ctrl: [], charts, checks: this.instChecks(M) };
  }

  // ---------------------------------------------------------------- the steps
  vSetup(M) {
    const D = this.data(), R = this.roles();
    const roleKey = { Anchor: 'A', Check: 'K', Drift: 'D', Linearity: 'L', Blank: 'B', Exclude: 'X' };
    const rows = D.std.map((s, i) => {
      const n = M.P.filter((q) => q.si === i).length;
      const u = s[7] == null && s[8] == null ? '' : (s[7] == null ? 'none' : this.fx(s[7], 2)) + ', ' + (s[8] == null ? 'none' : this.fx(s[8], 2));
      return {
        text: s[0], role: s[1], bg: R[roleKey[s[1]] || 'S'].bg,
        d13: s[2] == null ? '' : this.fx(s[2], 2), d15: s[3] == null ? '' : this.fx(s[3], 2),
        pc: s[4] == null ? '' : this.fx(s[4], 2), pn: s[5] == null ? '' : this.fx(s[5], 2),
        u, n: n ? String(n) : '', fw: n ? '600' : '400', op: n ? '1' : '0.62',
      };
    });
    const S = this.samples(M), lines = [];
    if (S.length) {
      const a = S[0].r, b = S[S.length - 1].r;
      lines.push({ t: S.length === 1 ? 'One position matches no row of the table and is treated as a sample (position ' + a + ').'
        : S.length + ' positions match no row of the table and are treated as samples' + (b - a + 1 === S.length ? ' (positions ' + a + ' to ' + b + ').' : ', from position ' + a + ' to ' + b + '.') });
    } else {
      lines.push({ t: 'Every position matches a row of the table, so the run has no samples to correct.' });
    }
    const ids = new Map();
    for (const q of S) {
      if (!ids.has(q.id)) ids.set(q.id, []);
      ids.get(q.id).push(q.r);
    }
    for (const [id, rs] of ids) {
      if (rs.length > 1) lines.push({ t: rs.length + ' samples share the identifier ' + (id || '(none)') + ' (positions ' + this.rowList(rs) + '), and all are used. Check the sequence if one was meant to carry another name.' });
    }
    const X = M.P.filter((q) => q.role === 'X');
    if (X.length) {
      const byText = new Map();
      for (const q of X) {
        const t = D.std[q.si] ? D.std[q.si][0] : q.id;
        if (!byText.has(t)) byText.set(t, []);
        byText.get(t).push(q.r);
      }
      lines.push({ t: (X.length === 1 ? 'One position matches an Exclude row and is' : X.length + ' positions match an Exclude row and are') + ' left out of every fit: '
        + [...byText].map(([t, rs]) => t + ' (' + this.rowList(rs) + ')').join('; ') + '.' });
    }
    const mats = this.anchorGroups(M).length;
    if (mats < 2) lines.push({ t: (mats ? 'Only one anchor material is in this run' : 'No anchor is in this run') + '; a calibration needs two that bracket the samples.' });
    return { rows, effect: { title: 'Notes', lines } };
  }

  blankChart(M, el, blanks, sel, H) {
    const blk = M.blank[el], col = this.col(el), gas = el === 'n' ? 'N2' : 'CO2', an = el === 'n' ? 'an N2' : 'a CO2';
    const ak = el === 'n' ? 'aN' : 'aC', dk = el === 'n' ? 'dN' : 'dC';
    const title = this.iso(el) + ' of the blank capsules';
    const withP = blanks.filter((q) => q[ak] != null && q[dk] != null);
    if (!withP.length) {
      const what = el === 'n' ? 'nitrogen' : 'carbon';
      return this.none(title, col, !blanks.length ? 'This run has no blank capsules, so there is no blank to subtract.'
        : (blanks.length === 1 ? 'The blank capsule gave no ' + gas + ' peak' : 'None of the ' + blanks.length + ' blank capsules gave ' + an + ' peak')
          + ', so there is no ' + what + ' blank to subtract.');
    }
    const ys = withP.map((q) => q[dk]), areas = withP.map((q) => q[ak]), ax = this.rowAxis(M);
    const yd = this.dom(ys.concat(blk.sd != null ? [blk.d - blk.sd, blk.d + blk.sd] : []), 0.12);
    return this.chart({
      title, sub: this.count(withP.length, 'capsule') + ', ' + gas + ' peaks ' + this.span(Math.min(...areas), Math.max(...areas), 2) + ' Vs',
      badge: this.badge(blk.inForce, blk.inForce ? 'Subtracted' : 'Not subtracted'), color: col, x: ax, y: yd,
      bandA: blk.sd != null ? { fill: this.tint(el), list: [{ y0: blk.d - blk.sd, y1: blk.d + blk.sd }] } : null,
      lineB: { paths: [[[ax.min, blk.d], [ax.max, blk.d]]], stroke: col, w: 1.5, dash: '5 4' },
      pts: withP.map((q) => ({ x: q.r, y: q[dk], kind: 'fill', sel: q.r === sel, label: q.id + ', position ' + q.r + ': ' + this.fx(q[dk], 2) + '‰, ' + this.fx(q[ak], 2) + ' Vs', pick: () => H.pick(q.r) })),
      notes: withP.map((q) => ({ x: q.r, y: q[dk], t: this.fx(q[ak], 2) + ' Vs', at: this.side(ax, q.r), c: '#56646F' })),
      legend: [{ kind: 'dot', color: col, label: 'Blank capsule' }, { kind: 'dash', color: col, label: 'Mean ' + this.fx(blk.d, 2) + '‰' }]
        .concat(blk.sd != null ? [{ kind: 'band', color: this.tint(el), label: '±1 SD' }] : []),
      xLabel: 'Position in the run', yLabel: this.iso(el) + ', raw (‰)',
    });
  }

  vBlank(M, d, sel, H) {
    const b = M.blank, mult = M.settings.blankMultiple, blanks = M.P.filter((q) => q.role === 'B'), S = this.samples(M);
    const why = (blk) => (d.blank !== 'Auto' || blk.inForce || !blk.n ? '' : blk.n < 3 ? ' (Auto needs three or more)'
      : blk.sd != null && blk.sd > 2 ? ' (Auto needs 2‰ or less)' : '');
    const verdict = (el) => {
      const blk = b[el], an = el === 'n' ? 'an N2' : 'a CO2';
      if (!blk.n) return { t: this.iso(el) + ': no blank capsule gave ' + an + ' peak', on: false };
      return { t: this.iso(el) + ': ' + (blk.inForce ? 'subtracted' : 'not subtracted') + ', ' + this.count(blk.n, 'capsule') + ' with ' + an + ' peak'
        + (blk.sd != null ? ', SD ' + this.fx(blk.sd, 2) + '‰' : '') + why(blk), on: blk.inForce };
    };
    const ctrl = [this.seg('Blank correction', 'blank', d.blank, ['Auto', 'On', 'Off'], H, [verdict('n'), verdict('c')])];
    const charts = [this.blankChart(M, 'n', blanks, sel, H), this.blankChart(M, 'c', blanks, sel, H)];
    const under = (el) => {
      const ak = el === 'n' ? 'aN' : 'aC', area = b[el].area;
      return mult && area > 0 ? S.filter((q) => q[ak] != null && q[ak] < mult * area) : [];
    };
    const uN = under('n'), uC = under('c');
    const stats = { rows: [
      { k: 'Capsules with a peak', a: String(b.n.n), b: String(b.c.n) },
      { k: 'Mean peak area (Vs)', a: b.n.n ? this.fx(b.n.area, 3) : '', b: b.c.n ? this.fx(b.c.area, 3) : '' },
      { k: 'Mean δ (‰)', a: b.n.n ? this.fx(b.n.d, 2) : '', b: b.c.n ? this.fx(b.c.d, 2) : '' },
      { k: 'SD (‰)', a: this.fx(b.n.sd, 2), b: this.fx(b.c.sd, 2) },
      mult ? { k: 'Samples under ' + this.num(mult) + '× the blank', a: b.n.area > 0 ? String(uN.length) : '', b: b.c.area > 0 ? String(uC.length) : '' }
        : { k: 'Samples close to the blank', a: 'not flagged', b: 'not flagged' },
      { k: 'Status', a: b.n.n ? (b.n.inForce ? 'Subtracted' : 'Not subtracted') : 'Nothing to subtract', b: b.c.n ? (b.c.inForce ? 'Subtracted' : 'Not subtracted') : 'Nothing to subtract' },
    ] };
    const lines = [];
    const alt = b.n.inForce || b.c.inForce ? 'Off' : 'On';
    const Ma = this.model(Object.assign({}, d, { blank: alt }));
    for (const el of ['n', 'c']) {
      if (Ma.blank[el].inForce === b[el].inForce) continue;
      const sh = this.shift(M, Ma, el);
      if (!sh.n) continue;
      const ak = el === 'n' ? 'aN' : 'aC';
      lines.push({ t: (Ma.blank[el].inForce ? 'Subtracting the blank' : 'Leaving the blank in') + ' would move sample ' + this.iso(el) + ' by ' + this.span(sh.lo, sh.hi, 2, true) + '‰, ' + sh.beyond + ' of ' + sh.n + ' by more than ' + this.kSig(M) + '.' });
      const big = sh.big;
      if (big && Math.abs(big.d) > 0.5) {
        lines.push({ t: 'The largest change is ' + this.who(big.q) + ' (' + this.fs(big.d, 2) + '‰), whose ' + (el === 'n' ? 'N2' : 'CO2') + ' peak is '
          + this.fx(big.q[ak] / b[el].area, 1) + ' times the blank. A sample this small relative to the blank cannot be corrected reliably.' });
      }
    }
    for (const [el, list, gas] of [['n', uN, 'N2'], ['c', uC, 'CO2']]) {
      if (!list.length) continue;
      const ak = el === 'n' ? 'aN' : 'aC';
      lines.push({ t: 'Under ' + this.times(mult) + ' the blank in ' + gas + ': ' + list.map((q) => q.id + ' (' + q.r + ', ' + this.fx(q[ak] / b[el].area, 1) + '×)').join(', ')
        + '. Their ' + this.iso(el) + ' is affected by the blank whether or not it is subtracted.' });
    }
    if (!lines.length) lines.push({ t: b.n.n || b.c.n ? 'The switch changes no sample in this run.' : 'No blank capsule gave a peak, so the switch changes nothing in this run.' });
    return { ctrl, charts, stats, effect: { title: 'Effect of this choice', lines } };
  }

  vLin(M, d, sel, H) {
    const L = M.lin, lin = M.P.filter((q) => q.role === 'L'), S = this.samples(M), mat = this.material(M, 'L');
    const verdict = (el) => {
      const l = L[el], f = l.fit;
      return { t: (l.inForce ? 'Applied' : 'Not applied') + ', ' + (f ? this.fs(f.slope, 5) + '‰ per Vs, ' + this.peq(f.p)
        : l.n ? 'no fit from ' + this.count(l.n, 'standard') : 'no size standards'), on: l.inForce };
    };
    const ctrl = [
      this.seg(this.iso('n'), 'linN', d.linN, ['Auto', 'On', 'Off'], H, [verdict('n')]),
      this.seg(this.iso('c'), 'linC', d.linC, ['Auto', 'On', 'Off'], H, [verdict('c')]),
    ];
    const charts = [], cnt = {};
    for (const el of ['n', 'c']) {
      const ak = el === 'n' ? 'aN' : 'aC', gas = el === 'n' ? 'N2' : 'CO2', l = L[el], f = l.fit, col = this.col(el);
      const over = l.max == null ? 0 : S.filter((q) => q[ak] != null && q[ak] > l.max).length;
      const under = l.min == null ? 0 : S.filter((q) => q[ak] != null && q[ak] < l.min).length;
      cnt[el] = [over, under];
      const title = this.iso(el) + ' against ' + gas + ' peak area';
      if (!l.n) { charts.push(this.none(title, col, 'No size standard gave a ' + gas + ' peak, so the dependence on peak size cannot be checked.')); continue; }
      const xmax = Math.max(l.max, ...S.map((q) => q[ak]).filter((v) => v != null));
      const ys = lin.map((q) => q[el].blank);
      const fitY = (x) => f.intercept + f.slope * x;
      const yd = this.dom(ys.concat(f ? [fitY(l.min), fitY(l.max)] : []), 0.25);
      const xd = { min: 0, max: this.dom([0, xmax], 0.02).max };
      charts.push(this.chart({
        title, sub: this.count(l.n, (mat ? mat + ' ' : '') + 'standard') + ', ' + this.span(l.min, l.max, 1) + ' Vs; ' + over + ' samples above, ' + under + ' below',
        badge: this.badge(l.inForce), color: col, x: xd, y: yd,
        bandA: { fill: this.tint(el), list: [{ x0: l.min, x1: l.max }] },
        lineA: f ? { paths: [[[l.min, fitY(l.min)], [l.max, fitY(l.max)]]], stroke: col, w: 2.2, dash: l.inForce ? '' : '6 4' } : null,
        lineB: f ? { paths: [[[0, fitY(0)], [l.min, fitY(l.min)]], [[l.max, fitY(l.max)], [xd.max, fitY(xd.max)]]], stroke: col, w: 1.2, dash: '2 4' } : null,
        rug: { xs: S.map((q) => q[ak]), stroke: '#6E7D8C' },
        pts: lin.map((q) => ({ x: q[ak], y: q[el].blank, kind: 'fill', sel: q.r === sel, label: q.id + ': ' + this.fx(q[ak], 1) + ' Vs, ' + this.fx(q[el].blank, 3) + '‰', pick: () => H.pick(q.r) })),
        legend: [{ kind: 'dot', color: col, label: 'Size standard' }, { kind: l.inForce ? 'line' : 'dash', color: col, label: l.inForce ? 'Fit, applied' : 'Fit, not applied' }, { kind: 'band', color: this.tint(el), label: 'Series range' }, { kind: 'line', color: '#6E7D8C', label: 'Sample peaks (ticks)' }],
        xLabel: gas + ' peak area (Vs)', yLabel: this.iso(el) + ' after blank (‰)',
      }));
    }
    const fN = L.n.fit, fC = L.c.fit;
    const stats = { rows: [
      { k: 'Slope (‰ per Vs)', a: this.fs(fN && fN.slope, 5), b: this.fs(fC && fC.slope, 5) },
      { k: 'Standard error', a: this.fx(fN && fN.se, 5), b: this.fx(fC && fC.se, 5) },
      { k: 'p', a: this.pv(fN && fN.p), b: this.pv(fC && fC.p) },
      { k: 'r²', a: this.fx(fN && fN.r2, 3), b: this.fx(fC && fC.r2, 3) },
      { k: 'Residual SD (‰)', a: this.fx(fN && fN.sd, 3), b: this.fx(fC && fC.sd, 3) },
      { k: 'Series range (Vs)', a: this.span(L.n.min, L.n.max, 1), b: this.span(L.c.min, L.c.max, 1) },
      { k: 'Samples above, below', a: cnt.n[0] + ', ' + cnt.n[1], b: cnt.c[0] + ', ' + cnt.c[1] },
      { k: 'Status', a: L.n.inForce ? 'Applied' : 'Not applied', b: L.c.inForce ? 'Applied' : 'Not applied' },
    ] };
    const lines = [];
    for (const el of ['n', 'c']) {
      const on = L[el].inForce, alt = on ? 'Off' : 'On';
      const Ma = this.model(Object.assign({}, d, el === 'n' ? { linN: alt } : { linC: alt }));
      if (Ma.lin[el].inForce === on) continue;
      const sh = this.shift(M, Ma, el);
      if (!sh.n) continue;
      lines.push({ t: (on ? 'Switching the ' + this.iso(el) + ' correction off' : 'Forcing the ' + this.iso(el) + ' correction on') + ' would move sample ' + this.iso(el) + ' by ' + this.span(sh.lo, sh.hi, 2, true) + '‰, ' + sh.beyond + ' of ' + sh.n + ' by more than ' + this.kSig(M) + '.' });
    }
    if (!lines.length) lines.push({ t: lin.length < 3 ? 'With fewer than three size standards, no slope can be applied, so the switches change nothing.' : 'The switches change no sample in this run.' });
    return { ctrl, charts, stats, effect: { title: 'Effect of this choice', lines } };
  }

  vDrift(M, d, sel, H) {
    const Dr = M.drift, mons = M.P.filter((q) => q.role === 'D'), T = this.num(M.settings.outlierT);
    const verdict = (el) => {
      const r = Dr[el], f = r.fit, out = r.outliers.map((x) => x + ' (t ' + this.fs(r.t[x], 2) + ')');
      return { t: this.iso(el) + ': ' + (r.inForce ? 'applied' : 'not applied') + ', ' + r.model + ', ' + (f ? this.peq(f.p) : 'no fit') + (out.length ? ', outlier ' + out.join(', ') : ''), on: r.inForce };
    };
    const ctrl = [
      this.seg('Drift correction', 'drift', d.drift, ['Auto', 'On', 'Off'], H, [verdict('n'), verdict('c')]),
      this.seg('Model', 'model', d.model, [{ t: 'Row', v: 'Row' }, { t: 'Background', v: 'Background' }, { t: 'Auto', v: 'Auto' }], H, [
        { t: 'Row fits drift against position. Background fits it against the background under the N2 reference peak. Auto uses whichever leaves less scatter.', on: false }]),
    ];
    const charts = [];
    for (const el of ['n', 'c']) {
      const r = Dr[el], f = r.fit, col = this.col(el), bg = r.model === 'Background';
      const title = this.iso(el) + ' of the drift monitors';
      const xs = mons.map((q) => (bg ? q.rbgd : q.r)), ys = mons.map((q) => q[el].lin);
      const good = xs.filter((v, i) => v != null && ys[i] != null);
      if (!good.length) { charts.push(this.none(title, col, mons.length ? 'No drift monitor has a ' + this.iso(el) + ' value' + (bg ? ' and a background' : '') + ' to fit.' : 'This run has no drift monitors, so drift cannot be followed.')); continue; }
      const x0 = Math.min(...good), x1 = Math.max(...good);
      const xd = bg ? this.dom(xs, 0.12) : this.rowAxis(M);
      const fitY = (x) => f.intercept + f.slope * x;
      const yd = this.dom(ys.concat(f ? [fitY(x0), fitY(x1)] : []), 0.2);
      const key = el === 'n' ? 'exclN' : 'exclC', early = new Set(r.before || []), side = (x) => this.side(xd, x);
      charts.push(this.chart({
        title, sub: r.n + ' of ' + mons.length + ' in the fit, ' + (bg ? 'against background' : 'against position') + ', ' + (f ? this.peq(f.p) : 'no fit'),
        badge: this.badge(r.inForce), color: col, x: xd, y: yd,
        lineA: f ? { paths: [[[x0, fitY(x0)], [x1, fitY(x1)]]], stroke: col, w: 2.2, dash: r.inForce ? '' : '6 4' } : null,
        pts: mons.map((q, i) => {
          const out = r.excluded.includes(q.r), pre = early.has(q.r);
          return {
            x: xs[i], y: ys[i], kind: out || pre ? 'hollow' : 'fill', ring: r.outliers.includes(q.r), sel: q.r === sel,
            label: q.id + ', position ' + q.r + ': ' + this.fx(ys[i], 3) + '‰' + (r.t[q.r] != null ? ', t ' + this.fs(r.t[q.r], 2) : '')
              + (pre && !out ? '. Before the configuration’s first drift position, so outside the fit.' : out ? '. Left out; click to put it back.' : '. Click to leave it out.'),
            pick: pre && !out ? null : () => H.toggle(key, q.r),
          };
        }),
        notes: mons.filter((q) => r.outliers.includes(q.r) || r.excluded.includes(q.r) || early.has(q.r)).map((q) => ({
          x: bg ? q.rbgd : q.r, y: q[el].lin, t: r.excluded.includes(q.r) ? 'left out' : early.has(q.r) ? 'before the fit' : 't ' + this.fs(r.t[q.r], 2),
          c: r.outliers.includes(q.r) ? '#A33A00' : '#56646F', at: side(bg ? q.rbgd : q.r),
        })),
        legend: [{ kind: 'dot', color: col, label: 'In the fit' }, { kind: 'hollow', color: col, label: 'Left out' }, { kind: 'ring', color: col, label: 'Outlier, |t| > ' + T }, { kind: r.inForce ? 'line' : 'dash', color: col, label: r.inForce ? 'Fit, applied' : 'Fit, not applied' }],
        xLabel: bg ? 'N2 reference-peak background (mV)' : 'Position in the run', yLabel: this.iso(el) + ' after linearity (‰)',
      }));
    }
    const unit = (el) => (Dr[el].model === 'Background' ? '‰ per mV' : '‰ per position');
    const stats = { rows: [
      { k: 'Model', a: Dr.n.model, b: Dr.c.model },
      { k: 'Slope', a: Dr.n.fit ? this.fs(Dr.n.fit.slope, 5) + ' ' + unit('n') : 'no fit', b: Dr.c.fit ? this.fs(Dr.c.fit.slope, 5) + ' ' + unit('c') : 'no fit' },
      { k: 'p', a: this.pv(Dr.n.fit && Dr.n.fit.p), b: this.pv(Dr.c.fit && Dr.c.fit.p) },
      { k: 'Residual SD (‰)', a: this.fx(Dr.n.fit && Dr.n.fit.sd, 3), b: this.fx(Dr.c.fit && Dr.c.fit.sd, 3) },
      { k: 'Monitors in the fit', a: Dr.n.n + ' of ' + mons.length, b: Dr.c.n + ' of ' + mons.length },
      { k: 'Left out', a: Dr.n.excluded.join(', ') || 'none', b: Dr.c.excluded.join(', ') || 'none' },
      { k: 'Outliers, |t| > ' + T, a: Dr.n.outliers.join(', ') || 'none', b: Dr.c.outliers.join(', ') || 'none' },
      { k: 'Status', a: Dr.n.inForce ? 'Applied' : 'Not applied', b: Dr.c.inForce ? 'Applied' : 'Not applied' },
    ] };
    if (Dr.n.before.length) stats.rows.splice(6, 0, { k: 'Before the first drift position', a: Dr.n.before.join(', '), b: Dr.c.before.join(', ') });
    const on = Dr.n.inForce || Dr.c.inForce, alt = on ? 'Off' : 'On';
    const Ma = this.model(Object.assign({}, d, { drift: alt }));
    const lines = [];
    for (const el of ['n', 'c']) {
      if (Ma.drift[el].inForce === Dr[el].inForce) continue;
      const sh = this.shift(M, Ma, el);
      if (!sh.n) continue;
      lines.push({ t: (alt === 'On' ? 'Forcing drift on' : 'Switching drift off') + ' would move sample ' + this.iso(el) + ' by ' + this.span(sh.lo, sh.hi, 2, true) + '‰, ' + sh.beyond + ' of ' + sh.n + ' by more than ' + this.kSig(M) + '.' });
    }
    if (Dr.n.outliers.length || Dr.c.outliers.length) lines.push({ t: 'With ' + this.words(mons.length) + ' monitors, leaving one out can make another look like an outlier. Leave a monitor out only when there is a known problem with its capsule or its weighing.' });
    if (!lines.length) lines.push({ t: Math.max(Dr.n.n, Dr.c.n) < 3 ? 'Fewer than three monitors are in the fit, so no drift slope can be applied.' : 'Reversing the switch changes no sample in this run.' });
    return { ctrl, charts, stats, effect: { title: 'Effect of this choice', lines } };
  }

  vCal(M, d, sel, H) {
    const A = M.P.filter((q) => q.role === 'A'), S = this.samples(M), K = M.P.filter((q) => q.role === 'K');
    const charts = [], lines = [];
    let outside = 0;
    for (const el of ['n', 'c']) {
      const f = M.cal[el], col = this.col(el), title = this.iso(el) + ' calibration';
      if (f.n < 2 || f.slope == null) {
        charts.push(this.none(title, col, 'Fewer than two anchor analyses have a ' + this.iso(el) + ' value, so ' + this.iso(el) + ' cannot be put on its scale.'));
        lines.push({ t: this.iso(el) + ': not calibrated.' });
        continue;
      }
      const pts = A.concat(S).filter((q) => q[el].drift != null && q[el].fin != null);
      const xs = pts.map((q) => q[el].drift), ys = pts.map((q) => q[el].fin);
      const xd = this.dom(xs, 0.06), yd = this.dom(ys.concat([f.accMin, f.accMax]), 0.06);
      const withV = S.filter((q) => q[el].fin != null);
      const inside = withV.filter((q) => q[el].fin >= f.accMin && q[el].fin <= f.accMax).length;
      const below = withV.filter((q) => q[el].fin < f.accMin), above = withV.filter((q) => q[el].fin > f.accMax);
      outside += below.length + above.length;
      const names = {};
      for (const q of A) (names[q.si] = names[q.si] || []).push(q);
      charts.push(this.chart({
        title, sub: 'Slope ' + this.fx(f.slope, 4) + ', intercept ' + this.fx(f.intercept, 3) + '‰, residual SD ' + this.fx(f.sd, 3) + '‰',
        badge: { text: inside + ' of ' + withV.length + ' inside the anchors', bg: '#EEF1F4', fg: '#3B4A55' }, color: col, x: xd, y: yd,
        bandB: { fill: '#FBF1DE', list: [{ y1: f.accMin }, { y0: f.accMax }] },
        lineA: { paths: [[[xd.min, f.slope * xd.min + f.intercept], [xd.max, f.slope * xd.max + f.intercept]]], stroke: col, w: 1.8 },
        pts: A.map((q) => ({ x: q[el].drift, y: q[el].fin, kind: 'fill', sel: q.r === sel, label: q.id + ', position ' + q.r + ': measured ' + this.fx(q[el].drift, 3) + ', accepted ' + this.fx(el === 'n' ? q.acc15 : q.acc13, 2) + '‰', pick: () => H.pick(q.r) }))
          .concat(withV.map((q) => ({ x: q[el].drift, y: q[el].fin, kind: 'small', sel: q.r === sel, label: q.id + ', position ' + q.r + ': ' + this.fx(q[el].fin, 2) + '‰', pick: () => H.pick(q.r) }))),
        notes: Object.keys(names).map((k) => {
          const g = names[k].filter((q) => q[el].drift != null);
          if (!g.length) return { x: null, y: null, t: '' };
          const mx = g.reduce((s, q) => s + q[el].drift, 0) / g.length;
          return { x: mx, y: el === 'n' ? g[0].acc15 : g[0].acc13, t: g[0].id, at: mx > (xd.min + xd.max) / 2 ? 'left' : 'right', c: '#13212B' };
        }),
        legend: [{ kind: 'dot', color: col, label: 'Anchor' }, { kind: 'small', color: '#8D9AA7', label: 'Sample' }, { kind: 'band', color: '#FBF1DE', label: 'Outside the anchors' }],
        xLabel: 'Measured, after drift (‰)', yLabel: this.iso(el) + ' calibrated (‰)',
      }));
      const lo = this.fx(f.accMin, 2), hi = this.fx(f.accMax, 2);
      if (!withV.length) { lines.push({ t: this.iso(el) + ': the anchors span ' + lo + ' to ' + hi + '‰; no sample has a value.' }); continue; }
      let t = this.iso(el) + ': ' + inside + ' of ' + withV.length + ' samples within ' + lo + ' to ' + hi + '‰.';
      if (above.length) t += ' ' + above.length + ' above, up to ' + this.fx(Math.max(...above.map((q) => q[el].fin)), 2) + '‰.';
      if (below.length) t += ' ' + below.length + ' below, down to ' + this.fx(Math.min(...below.map((q) => q[el].fin)), 2) + '‰.';
      lines.push({ t });
    }
    if (K.length) lines.push({ t: this.count(K.length, 'check analysis', 'check analyses') + ' in this run test the calibration; Quality checks compares them with their accepted values.' });
    else if (outside && this.anchorGroups(M).length < 3) lines.push({ t: 'A third anchor or a check standard near the samples would test the extrapolation; this run has neither.' });
    const n = M.cal.n, c = M.cal.c;
    const stats = { rows: [
      { k: 'Slope', a: this.fx(n.slope, 5), b: this.fx(c.slope, 5) },
      { k: 'Intercept (‰)', a: this.fx(n.intercept, 3), b: this.fx(c.intercept, 3) },
      { k: 'r²', a: this.fx(n.r2, 5), b: this.fx(c.r2, 5) },
      { k: 'Residual SD (‰)', a: this.fx(n.sd, 3), b: this.fx(c.sd, 3) },
      { k: 'Anchor analyses', a: String(n.n), b: String(c.n) },
    ] };
    return { ctrl: [], charts, stats, effect: { title: 'Extrapolation', lines } };
  }

  // the elemental factor along the run, from the monitors' factors
  interp(ms, row) {
    if (row <= ms[0][0]) return ms[0][1];
    if (row >= ms[ms.length - 1][0]) return ms[ms.length - 1][1];
    for (let i = 0; i + 1 < ms.length; i++) {
      const [x0, f0] = ms[i], [x1, f1] = ms[i + 1];
      if (x0 <= row && row <= x1) return f0 + (row - x0) / (x1 - x0) * (f1 - f0);
    }
    return 1;
  }

  vElem(M, d, sel, H) {
    const E = M.elem, D = this.data(), gate = M.settings.elemGate, mons = M.P.filter((q) => q.role === 'D'), S = this.samples(M);
    const mat = this.material(M, 'D');
    const big = (el) => {
      let b = null;
      for (const q of mons) if (!E[el].excluded.includes(q.r) && q['bias' + el] != null && (!b || Math.abs(q['bias' + el]) > Math.abs(b['bias' + el]))) b = q;
      return b;
    };
    const verdict = (el) => {
      const b = big(el);
      return { t: (el === 'n' ? '%N' : '%C') + ': ' + (E[el].inForce ? 'corrected' : 'not corrected') + (b ? ', largest monitor bias ' + this.fs(b['bias' + el] * 100, 1) + '% at ' + b.r : ', no monitor to read it from')
        + (E[el].excluded.length ? ', ' + this.andList(E[el].excluded) + ' left out' : ''), on: E[el].inForce };
    };
    const ctrl = [this.seg('Elemental correction', 'elem', d.elem, ['Auto', 'On', 'Off'], H, [verdict('n'), verdict('c')])];
    const ax = this.rowAxis(M), charts = [];
    for (const el of ['n', 'c']) {
      const e = E[el], col = this.col(el), label = el === 'n' ? '%N' : '%C';
      const title = label + ' of the drift monitors against certified';
      const ys = mons.map((q) => (q['bias' + el] == null ? null : q['bias' + el] * 100));
      if (!ys.some((v) => v != null)) {
        charts.push(this.none(title, col, mons.length ? 'No drift monitor has both a measured ' + label + ' and a certified one, so sensitivity drift cannot be read from them.' : 'This run has no drift monitors, so sensitivity drift cannot be read.'));
        continue;
      }
      const yd = this.dom(ys.concat([-(gate + 1), gate + 1]), 0.1);
      const curve = [];
      if (e.mons.length >= 2) for (let r = ax.min + 1; r <= ax.max - 1; r++) curve.push([r, (this.interp(e.mons, r) - 1) * 100]);
      const certs = [...new Set(mons.map((q) => (el === 'n' ? q.certN : q.certC)).filter((v) => v != null))];
      const cert = certs.length === 1 ? (mat ? this.cap(mat) + ', ' : '') + this.num(certs[0]) + '% ' + (el === 'n' ? 'N' : 'C') : 'certified values from the standards table';
      charts.push(this.chart({
        title, sub: cert + '; ' + e.n + ' of ' + mons.length + ' monitors used',
        badge: this.badge(e.inForce, e.inForce ? 'Corrected' : 'Not corrected'), color: col, x: ax, y: yd,
        bandA: { fill: '#E7F3EC', list: [{ y0: -gate, y1: gate }] },
        lineA: curve.length ? { paths: [curve], stroke: col, w: 2, dash: e.inForce ? '' : '6 4' } : null,
        pts: mons.map((q, i) => ({
          x: q.r, y: ys[i], kind: e.excluded.includes(q.r) ? 'hollow' : 'fill', sel: q.r === sel,
          label: q.id + ', position ' + q.r + ': ' + this.fs(ys[i], 1) + '%' + (e.excluded.includes(q.r) ? '. Left out; click to put it back.' : '. Click to leave it out.'),
          pick: () => H.toggle('exclE', q.r),
        })),
        notes: mons.filter((q, i) => ys[i] != null && Math.abs(ys[i]) > gate).map((q) => ({ x: q.r, y: q['bias' + el] * 100, t: this.fs(q['bias' + el] * 100, 1) + '%', at: this.side(ax, q.r), c: '#3B4A55' })),
        legend: [{ kind: 'dot', color: col, label: 'Monitor used' }, { kind: 'hollow', color: col, label: 'Left out' }, { kind: e.inForce ? 'line' : 'dash', color: col, label: e.inForce ? 'Factor applied' : 'Factor, not applied' }, { kind: 'band', color: '#E7F3EC', label: 'Within ' + this.num(gate) + '%' }],
        xLabel: 'Position in the run', yLabel: 'Measured over certified, minus 1 (%)',
      }));
    }
    const kn = D.k.n || {}, kc = D.k.c || {};
    const bn = big('n'), bc = big('c');
    const stats = { rows: [
      { k: 'k-factor (Vs per µg)', a: this.fx(kn.slope, 4), b: this.fx(kc.slope, 4) },
      { k: 'k-factor r²', a: this.fx(kn.r2, 4), b: this.fx(kc.r2, 4) },
      { k: 'Weighed standards in it', a: kn.n == null ? '' : String(kn.n), b: kc.n == null ? '' : String(kc.n) },
      { k: 'Monitors used', a: E.n.n + ' of ' + mons.length, b: E.c.n + ' of ' + mons.length },
      { k: 'Largest monitor bias', a: bn ? this.fs(bn.biasn * 100, 1) + '% at ' + bn.r : '', b: bc ? this.fs(bc.biasc * 100, 1) + '% at ' + bc.r : '' },
      { k: 'Status', a: E.n.inForce ? 'Corrected' : 'Not corrected', b: E.c.inForce ? 'Corrected' : 'Not corrected' },
    ] };
    const on = E.n.inForce || E.c.inForce, alt = on ? 'Off' : 'On';
    const Ma = this.model(Object.assign({}, d, { elem: alt }));
    const lines = [];
    if (Ma.elem.n.inForce !== E.n.inForce || Ma.elem.c.inForce !== E.c.inForce) {
      const parts = [];
      for (const [key, name] of [['pctN', '%N'], ['pctC', '%C'], ['cn', 'C:N']]) {
        const sh = this.pctShift(M, Ma, key);
        if (sh.n && (sh.lo !== 0 || sh.hi !== 0)) parts.push((parts.length ? '' : 'sample ') + name + ' by ' + this.span(sh.lo, sh.hi, 1, true) + '%');
      }
      lines.push({ t: (alt === 'On' ? 'Forcing the correction on' : 'Switching the correction off') + (parts.length ? ' would change ' + this.andList(parts) + ', relative to each value.' : ' would change no sample value, since no sample has %N, %C, or C:N.') });
    } else {
      lines.push({ t: 'Fewer than two monitors are used, so the switch changes nothing.' });
    }
    if (S.length && !S.some((q) => q.pctN != null || q.pctC != null)) lines.push({ t: 'The samples were not weighed, so the run reports no %N or %C for them' + (S.some((q) => q.cn != null) ? '; the correction still moves their C:N.' : '.') });
    lines.push({ t: 'Leave a monitor out when its weight is in doubt. A capsule weighed 10% light reads 10% high in both %N and %C.' });
    return { ctrl, charts, stats, effect: { title: 'Effect of this choice', lines } };
  }

  vUnc(M, d, sel, H) {
    const U = M.unc, run = this.state.run, S = this.samples(M), A = M.P.filter((q) => q.role === 'A');
    const src = (el) => {
      const b = U[el], want = el === 'n' ? d.uSrcN : d.uSrcC, now = this.uSrc(M, d, el);
      if (!b.ok) return { t: 'No uncertainty: the calibration has fewer than three anchor analyses', on: false };
      if (b.uS == null) return { t: 'No repeatability estimate: no anchor was analyzed more than once', on: false };
      const what = now === 'monitors' ? 'SD of the drift monitors in the fit' : now === 'config' ? 'set in the configuration'
        : 'pooled SD of the anchor replicates, ' + b.df + ' df' + (want === 'monitors' ? ', since fewer than two monitors are left in the drift fit' : '');
      return { t: '1σ = ' + this.fx(b.uS, 3) + '‰, ' + what, on: true };
    };
    const opts = (el) => [{ t: 'Pooled anchors', v: 'pooled' }, { t: 'Drift monitors', v: 'monitors' }]
      .concat(run.uCfg[el] != null ? [{ t: 'Configuration', v: 'config' }] : []);
    const ctrl = [
      this.seg(this.iso('n') + ' repeatability', 'uSrcN', d.uSrcN, opts('n'), H, [src('n')]),
      this.seg(this.iso('c') + ' repeatability', 'uSrcC', d.uSrcC, opts('c'), H, [src('c')]),
    ];
    const charts = [];
    for (const el of ['n', 'c']) {
      const b = U[el], f = M.cal[el], col = this.col(el), title = '1σ of ' + this.iso(el) + ' across the scale';
      const Us = S.map((q) => q[el].u).filter((v) => v != null);
      const xm = S.concat(A).map((q) => q[el].drift).filter((v) => v != null);
      if (!b.ok || b.uS == null || !Us.length || !xm.length) {
        charts.push(this.none(title, col, !b.ok ? 'The calibration has fewer than three anchor analyses, so it carries no uncertainty.'
          : b.uS == null ? 'No anchor was analyzed more than once and the configuration sets no repeatability, so 1σ cannot be computed.'
            : 'No sample has a ' + this.iso(el) + ' value.'));
        continue;
      }
      const lo = Math.min(...xm), hi = Math.max(...xm), curve = [];
      for (let i = 0; i <= 60; i++) {
        const x = lo + (hi - lo) * i / 60;
        curve.push([f.slope * x + f.intercept, b.at(x).u]);
      }
      const fin = S.map((q) => q[el].fin);
      const xd = this.dom(fin.concat([f.accMin, f.accMax]), 0.05);
      const yd = { min: 0, max: this.dom(Us.concat(curve.map((p) => p[1])), 0.15).max };
      const name = (acc) => { const q = A.find((p) => (el === 'n' ? p.acc15 : p.acc13) === acc); return q ? q.id : ''; };
      charts.push(this.chart({
        title, sub: 'Samples ' + this.span(Math.min(...Us), Math.max(...Us), 3) + '‰; least near the anchors’ mean',
        badge: this.badge(true, '1σ'), color: col, x: xd, y: yd,
        lineA: { paths: [curve], stroke: col, w: 1.8 },
        lineB: { paths: [[[f.accMin, 0], [f.accMin, yd.max]], [[f.accMax, 0], [f.accMax, yd.max]]], stroke: '#56646F', w: 1, dash: '3 4' },
        pts: S.map((q) => ({ x: q[el].fin, y: q[el].u, kind: 'fill', sel: q.r === sel, label: q.id + ', position ' + q.r + ': ' + this.fx(q[el].fin, 2) + ' ± ' + this.fx(q[el].u, 3) + '‰ (1σ)', pick: () => H.pick(q.r) })),
        // the anchors' names, one lower than the other when the anchors sit close on the axis
        notes: [{ x: f.accMin, y: yd.max * 0.94, t: name(f.accMin), at: 'right', c: '#3B4A55' },
          { x: f.accMax, y: yd.max * ((f.accMax - f.accMin) / (xd.max - xd.min) < 0.3 ? 0.82 : 0.94), t: name(f.accMax), at: 'left', c: '#3B4A55' }],
        legend: [{ kind: 'dot', color: col, label: 'Sample' }, { kind: 'line', color: col, label: '1σ along the line' }, { kind: 'dash', color: '#56646F', label: 'Anchor values' }],
        xLabel: this.iso(el) + ' calibrated (‰)', yLabel: '1σ (‰)',
      }));
    }
    const q = M.P.find((p) => p.r === sel) || S[0] || A[0] || M.P[0];
    const part = (el) => {
      const p = q[el].parts;
      if (!p) return null;
      const t = p.rep + p.acc + p.mean;
      return t > 0 ? [p.rep / t, p.acc / t, p.mean / t] : null;
    };
    const pn = part('n'), pc = part('c');
    const pct = (v) => (v == null ? '' : Math.round(v * 100) + '%');
    const w = (v) => (v == null ? '0%' : Math.max(1, Math.round(v * 100)) + '%');
    const budget = {
      title: this.who(q), rows: [
        { k: 'Single-analysis repeatability', n: pct(pn && pn[0]), c: pct(pc && pc[0]), wn: w(pn && pn[0]), wc: w(pc && pc[0]) },
        { k: 'Anchors’ certified values', n: pct(pn && pn[1]), c: pct(pc && pc[1]), wn: w(pn && pn[1]), wc: w(pc && pc[1]) },
        { k: 'Anchors’ measured means', n: pct(pn && pn[2]), c: pct(pc && pc[2]), wn: w(pn && pn[2]), wc: w(pc && pc[2]) },
      ],
    };
    const certs = (el) => {
      const vs = [...new Set(this.anchorGroups(M).map((g) => (el === 'n' ? g.uacc15 : g.uacc13)).filter((v) => v != null))];
      return vs.length ? vs.map((v) => this.fx(v, 2)).join(', ') : 'none';
    };
    const rows = [
      { k: 'Repeatability, 1σ (‰)', a: this.fx(U.n.uS, 3), b: this.fx(U.c.uS, 3) },
      { k: 'Pooled anchor SD (‰)', a: this.fx(U.n.pooled, 3), b: this.fx(U.c.pooled, 3) },
      { k: 'Drift monitor SD (‰)', a: this.fx(M.monSD.n, 3), b: this.fx(M.monSD.c, 3) },
    ];
    if (run.uCfg.n != null || run.uCfg.c != null) rows.push({ k: 'Configuration value (‰)', a: this.fx(run.uCfg.n, 3), b: this.fx(run.uCfg.c, 3) });
    rows.push({ k: 'Anchor certificates, 1σ (‰)', a: certs('n'), b: certs('c') }, { k: '1σ of the samples (‰)', a: this.uSpan(S, 'n', 3), b: this.uSpan(S, 'c', 3) });
    const lines = [];
    for (const el of ['n', 'c']) {
      const key = el === 'n' ? 'uSrcN' : 'uSrcC', now = this.uSrc(M, d, el), alt = now === 'monitors' ? 'pooled' : 'monitors';
      const dAlt = Object.assign({}, d, { [key]: alt }), Ma = this.model(dAlt);
      if (this.uSrc(Ma, dAlt, el) !== alt) continue;
      const a = this.samples(Ma).map((p) => p[el].u).filter((v) => v != null), b = S.map((p) => p[el].u).filter((v) => v != null);
      if (!a.length || !b.length) continue;
      lines.push({ t: 'With the ' + (alt === 'monitors' ? 'drift monitors’ SD (' + this.fx(M.monSD[el], 3) + '‰)' : 'pooled anchor SD (' + this.fx(U[el].pooled, 3) + '‰)')
        + ', 1σ(' + this.iso(el) + ') would be ' + this.span(Math.min(...a), Math.max(...a), 3) + '‰ instead of ' + this.span(Math.min(...b), Math.max(...b), 3) + '‰.' });
    }
    lines.push({ t: '1σ does not include the blank, linearity, or drift corrections, or differences between runs.' });
    return { ctrl, charts, stats: { rows }, effect: { title: 'Effect of this choice', lines }, budget: pn || pc ? budget : null };
  }

  qcChecks(M, d) {
    const D = this.data(), set = M.settings, S = this.samples(M), out = [], ax = this.rowAxis(M);
    if (S.length) {
      const rev = [];
      for (const [k, el, lab] of [['linN', 'n', 'linearity'], ['linC', 'c', 'linearity'], ['drift', 'n', 'drift'], ['drift', 'c', 'drift']]) {
        const on = k === 'drift' ? M.drift[el].inForce : M.lin[el].inForce;
        const Ma = this.model(Object.assign({}, d, { [k]: on ? 'Off' : 'On' }));
        rev.push({ el, lab, on, sh: this.shift(M, Ma, el) });
      }
      const hit = rev.filter((r) => r.sh.beyond > 0);
      out.push({ s: hit.length ? 'caution' : 'ok', title: 'Effect of the choices',
        t: 'If each choice were reversed, ' + rev.map((r, i) => (i === rev.length - 1 ? 'and ' : '') + this.iso(r.el) + ' ' + r.lab + ' (' + (r.on ? 'applied' : 'not applied') + ') would move ' + r.sh.beyond + (i ? '' : ' of ' + r.sh.n + ' samples')).join(', ') + ' by more than ' + this.kSig(M) + '.' });
    }
    const E = M.elem, mons = M.P.filter((q) => q.role === 'D');
    let worst = 0, at = null, el0 = 'n';
    for (const el of ['n', 'c']) {
      for (const q of mons) {
        if (E[el].excluded.includes(q.r) || q['bias' + el] == null) continue;
        if (Math.abs(q['bias' + el]) > Math.abs(worst) || at == null) { worst = q['bias' + el]; at = q.r; el0 = el; }
      }
    }
    if (at == null) {
      out.push({ s: 'note', title: 'Elemental content', t: 'No drift monitor has both a measured and a certified %N or %C, so sensitivity drift was not checked.' });
    } else {
      const left = E.n.excluded;
      const applied = E.n.inForce && E.c.inForce ? 'applied to both' : E.n.inForce ? 'applied to %N' : E.c.inForce ? 'applied to %C' : 'not applied';
      out.push({ s: Math.abs(worst) > 0.10 ? 'caution' : Math.abs(worst) * 100 > set.elemGate || left.length ? 'note' : 'ok', title: 'Elemental content',
        t: 'Largest bias among the monitors used is ' + this.fs(worst * 100, 1) + '% in ' + (el0 === 'n' ? '%N' : '%C') + ' at position ' + at + (left.length ? '; ' + this.andList(left) + ' left out' : '') + '. The correction is ' + applied + '.' });
    }
    if (mons.length) {
      const outs = [], lo = [];
      for (const el of ['n', 'c']) for (const r of M.drift[el].outliers) outs.push(this.who(M.P.find((q) => q.r === r)) + ' in ' + this.iso(el) + ', t ' + this.fs(M.drift[el].t[r], 2));
      for (const el of ['n', 'c']) for (const r of M.drift[el].excluded) lo.push(r + ' from the ' + this.iso(el) + ' fit');
      const tested = mons.some((q) => M.drift.n.t[q.r] != null || M.drift.c.t[q.r] != null);
      out.push({ s: outs.length || lo.length ? 'note' : 'ok', title: 'Drift monitors agree',
        t: (outs.length ? 'Outlier: ' + outs.join('; ') + '.' : tested ? 'No monitor is more than ' + this.num(set.outlierT) + ' t from the others.' : 'Fewer than five monitors are in the fit, so none is tested as an outlier.')
          + (lo.length ? ' Left out: ' + lo.join(', ') + '.' : '') });
    } else {
      out.push({ s: 'note', title: 'Drift monitors agree', t: 'This run has no drift monitors, so drift was not checked.' });
    }
    if (S.length) {
      const oa = (el) => S.filter((q) => q[el].fin != null && M.cal[el].accMin != null && (q[el].fin < M.cal[el].accMin || q[el].fin > M.cal[el].accMax)).length;
      out.push({ s: oa('n') || oa('c') ? 'note' : 'ok', title: 'Anchors bracket the samples',
        t: this.iso('n') + ' ' + oa('n') + ' of ' + S.length + ' and ' + this.iso('c') + ' ' + oa('c') + ' of ' + S.length + ' samples fall outside the anchors and are extrapolated.' });
      const cnt = (s) => S.filter((q) => q.flags.includes(s)).length;
      const nu = cnt('N below linearity range'), no = cnt('N above linearity range'), cu = cnt('C below linearity range'), co = cnt('C above linearity range');
      out.push({ s: nu + no + cu + co ? 'note' : 'ok', title: 'Sample size against the size series',
        t: 'N2 peaks of ' + nu + ' samples fall below the series and ' + no + ' above it; CO2 peaks of ' + cu + ' below and ' + co + ' above.' });
    }
    for (const c of this.instChecks(M)) if (c.s !== 'info') out.push(Object.assign({}, c, { step: 'inst' }));
    const K = M.P.filter((q) => q.role === 'K');
    if (K.length) {
      const off = K.filter((q) => q.flags.some((f) => /^check /.test(f)));
      out.push({ s: off.length ? 'note' : 'ok', title: 'Check standard',
        t: this.count(K.length, 'check analysis', 'check analyses') + ' in this run' + (off.length
          ? '; ' + off.length + ' differ from the accepted value by more than ' + this.num(set.checkOffset) + '‰: ' + off.map((q) => this.who(q) + ', ' + q.flags.filter((f) => /^check /.test(f)).map((f) => this.ftext(f).replace(/^check /, '')).join(', ')).join('; ') + '.'
          : '; each is within ' + this.num(set.checkOffset) + '‰ of its accepted value.') });
    } else {
      out.push({ s: 'note', title: 'Check standard', t: 'None in this run. The FIRMS guide (Dunn and Carter, 2018, section 6.3) recommends an independent check material in every run.' });
    }
    return out.map((c) => Object.assign({}, c, this.status(c.s)));
  }

  vQC(M, d, sel, H) {
    const checks = this.qcChecks(M, d).map((c) => (c.step === 'inst'
      ? Object.assign({}, c, { hasGo: true, goText: 'Step 1, ' + NAMES.inst, go: H.go('inst') }) : c));
    return { ctrl: [], charts: [], checks };
  }

  methods(M, d) {
    const D = this.data(), L = M.lin, Dr = M.drift, B = M.blank, E = M.elem, U = M.unc, set = M.settings, run = this.state.run;
    const S = this.samples(M), mons = M.P.filter((q) => q.role === 'D'), lin = M.P.filter((q) => q.role === 'L');
    const s = [];
    const groups = this.anchorGroups(M);
    let cal = 'Data were processed with ea_irms ' + D.version + '.';
    if (groups.length >= 2 && M.cal.n.n >= 2 && M.cal.c.n >= 2) {
      const points = groups.length === 2 ? 'two-point' : groups.length === 3 ? 'three-point' : groups.length + '-point';
      const vals = (g) => [g.acc15 != null ? this.fx(g.acc15, 2) + '‰' : null, g.acc13 != null ? this.fx(g.acc13, 2) + '‰' : null].filter(Boolean).join(' and ');
      const reps = groups.map((g) => g.n);
      const times = (n) => (n === 1 ? 'once' : n === 2 ? 'twice' : n === 3 ? 'in triplicate' : n + ' times');
      const repText = reps.every((n) => n === reps[0]) ? 'each analyzed ' + times(reps[0]) : 'analyzed ' + this.andList(reps.map(String)) + ' times';
      cal += ' δ¹⁵N and δ¹³C were normalized to the AIR and VPDB scales by a ' + points + ' calibration on ' + this.andList(groups.map((g) => g.name + ' (' + vals(g) + ')'))
        + ', ' + repText + ', regressing accepted on measured values (slopes ' + this.fx(M.cal.n.slope, 4) + ' and ' + this.fx(M.cal.c.slope, 4) + ').';
    } else {
      cal += ' The run could not be put on the AIR and VPDB scales, since it has fewer than two anchor materials with values for both isotopes.';
    }
    s.push(cal);
    const blankText = (el) => {
      const blk = B[el], what = el === 'n' ? 'nitrogen' : 'carbon', gas = el === 'n' ? 'N2' : 'CO2';
      if (!blk.n) return null;
      const caps = blk.n === 1 ? 'the blank capsule gave a ' + gas + ' peak of ' + this.fx(blk.area, 2) + ' Vs' : 'the ' + blk.n + ' blank capsules gave ' + gas + ' peaks of ' + this.fx(blk.area, 2) + ' Vs on average'
        + (blk.sd != null ? ' with a ' + this.iso(el) + ' SD of ' + this.fx(blk.sd, 2) + '‰' : '');
      return blk.inForce ? 'A ' + what + ' blank of ' + this.fx(blk.area, 2) + ' Vs (' + this.count(blk.n, 'capsule') + ', ' + this.iso(el) + ' ' + this.fx(blk.d, 2) + '‰) was subtracted by mass balance.'
        : 'No ' + what + ' blank was subtracted; ' + caps + '.';
    };
    const bt = ['n', 'c'].map(blankText).filter(Boolean);
    if (bt.length) s.push(...bt);
    else s.push(M.P.some((q) => q.role === 'B') ? 'No blank was subtracted, since no blank capsule gave a peak.' : 'No blank was subtracted; the run has no blank capsules.');
    const mat = this.material(M, 'L'), matD = this.material(M, 'D');
    const why = (l) => (l.mode === 'On' && !(l.fit && l.fit.p != null && l.fit.p < set.alpha) ? ', applied by choice'
      : l.mode === 'Off' && l.fit && l.fit.p != null && l.fit.p < set.alpha ? ', switched off by choice' : '');
    const lt = (el) => (L[el].fit ? this.fs(L[el].fit.slope, 5) + '‰ per Vs, ' + this.peq(L[el].fit.p) + why(L[el]) : 'no fit');
    const series = this.count(lin.length, (mat ? mat + ' ' : '') + 'standard');
    if (!lin.length) s.push('No peak-size correction was applied; the run has no size standards.');
    else if (L.n.inForce && L.c.inForce) s.push('A peak-size correction was applied to both δ¹⁵N (' + lt('n') + ') and δ¹³C (' + lt('c') + '), from ' + series + '.');
    else if (L.n.inForce || L.c.inForce) {
      const on = L.n.inForce ? 'n' : 'c', off = on === 'n' ? 'c' : 'n';
      s.push('A peak-size correction was applied to ' + this.iso(on) + ' (' + lt(on) + '; ' + series + ' of ' + this.span(L[on].min, L[on].max, 1) + ' Vs) but not to ' + this.iso(off) + ' (' + (L[off].fit ? this.peq(L[off].fit.p) + why(L[off]) : 'no fit') + ').');
    } else {
      s.push('No peak-size correction was applied (' + (L.n.fit ? this.peq(L.n.fit.p) + why(L.n) : 'no fit') + ' for δ¹⁵N and ' + (L.c.fit ? this.peq(L.c.fit.p) + why(L.c) : 'no fit') + ' for δ¹³C, from ' + series + ').');
    }
    const mdl = (el) => (Dr[el].model === 'Background' ? 'against the N2 reference-peak background' : 'against position');
    const dt = (el) => (Dr[el].fit ? this.peq(Dr[el].fit.p) : 'no fit') + (Dr[el].excluded.length ? ', position' + (Dr[el].excluded.length > 1 ? 's ' : ' ') + this.andList(Dr[el].excluded) + ' left out' : '');
    const first = run.config.first_drift_position;
    const monText = this.count(mons.length, (matD ? matD + ' ' : '') + 'monitor') + (Dr.n.before.length ? ' (those from position ' + first + ' on in the fit)' : '');
    if (!mons.length) s.push('No drift correction was applied; the run has no drift monitors.');
    else if (Dr.n.inForce || Dr.c.inForce) {
      const both = Dr.n.inForce && Dr.c.inForce, on = Dr.n.inForce ? 'n' : 'c';
      const how = both && Dr.n.model !== Dr.c.model ? 'Drift corrections were applied to both isotopes, δ¹⁵N ' + mdl('n') + ' and δ¹³C ' + mdl('c')
        : 'A drift correction ' + mdl(on) + ' was applied to ' + (both ? 'both isotopes' : this.iso(on) + ' only');
      s.push(how + ', from ' + monText + ' (' + this.iso('n') + ' ' + dt('n') + '; ' + this.iso('c') + ' ' + dt('c') + ').');
    } else s.push('No drift correction was applied (' + monText + '; ' + this.iso('n') + ' ' + dt('n') + ', ' + this.iso('c') + ' ' + dt('c') + ').');
    const kn = D.k.n || {}, kc = D.k.c || {};
    const weighed = S.some((q) => q.pctN != null || q.pctC != null), ratio = S.some((q) => q.cn != null);
    const big = (el) => Math.max(0, ...mons.filter((q) => !E[el].excluded.includes(q.r) && q['bias' + el] != null).map((q) => Math.abs(q['bias' + el]))) * 100;
    if (!weighed && !ratio) s.push('No elemental content is reported for the samples.');
    else {
      const fitted = kn.n === kc.n ? this.count(kn.n || 0, 'weighed standard') : (kn.n || 0) + ' and ' + (kc.n || 0) + ' weighed standards';
      let e = (weighed ? '%N, %C, and C:N were' : 'C:N was') + ' computed from k-factors fitted to ' + fitted;
      const left = E.n.excluded.length ? '; position' + (E.n.excluded.length > 1 ? 's ' : ' ') + this.andList(E.n.excluded) + ' left out' : '';
      if (E.n.inForce && E.c.inForce) e += ' and corrected for sensitivity drift with the drift monitors (largest bias ' + this.fx(big('n'), 1) + '% in %N and ' + this.fx(big('c'), 1) + '% in %C' + left + ').';
      else if (E.n.inForce || E.c.inForce) e += '; ' + (E.n.inForce ? 'nitrogen' : 'carbon') + ' was corrected for sensitivity drift with the drift monitors (largest bias ' + this.fx(big(E.n.inForce ? 'n' : 'c'), 1) + '%' + left + ') and ' + (E.n.inForce ? 'carbon' : 'nitrogen') + ' was not (' + this.fx(big(E.n.inForce ? 'c' : 'n'), 1) + '%).';
      else e += '; no sensitivity drift correction was applied (largest monitor bias ' + this.fx(big('n'), 1) + '% in %N and ' + this.fx(big('c'), 1) + '% in %C).';
      if (!weighed) e += ' The samples were not weighed, so no %N or %C is reported for them.';
      s.push(e);
    }
    if (U.n.ok && U.c.ok && U.n.uS != null && U.c.uS != null) {
      const srcName = (k) => (k === 'monitors' ? 'the SD of the drift monitors' : k === 'config' ? 'a value set in the configuration' : 'the pooled SD of the anchor replicates');
      const sN = this.uSrc(M, d, 'n'), sC = this.uSrc(M, d, 'c');
      const rep = sN === sC ? srcName(sN) + ', ' + this.fx(U.n.uS, 3) + '‰ for δ¹⁵N and ' + this.fx(U.c.uS, 3) + '‰ for δ¹³C'
        : srcName(sN) + ' for δ¹⁵N, ' + this.fx(U.n.uS, 3) + '‰, and ' + srcName(sC) + ' for δ¹³C, ' + this.fx(U.c.uS, 3) + '‰';
      s.push('Uncertainties are one standard deviation (1σ) and combine the anchors’ certified uncertainties, the standard errors of their measured means, and single-analysis repeatability (' + rep + '), propagated through the calibration; they do not include the blank, peak-size, or drift corrections.');
    } else {
      s.push('No uncertainty is reported, since the calibration has fewer than three anchor analyses or no repeatability estimate.');
    }
    return s.join(' ');
  }

  vResults(M, d, sel, H, filt) {
    const S = this.samples(M), st = this.state, run = st.run, nm = this.names(run);
    const size = (q) => q.flags.some((f) => /linearity range/.test(f));
    const sets = { all: S, inside: S.filter((q) => !size(q)), size: S.filter(size), carry: S.filter((q) => q.carry) };
    const list = sets[filt] || S;
    const rows = list.map((q) => ({
      r: String(q.r), id: q.id, amt: this.fx(q.amt, 3), aN: this.fx(q.aN, 1),
      dN: this.fx(q.n.fin, 2), UN: this.fx(q.n.u, 2), dC: this.fx(q.c.fin, 2), UC: this.fx(q.c.u, 2),
      pN: this.fx(q.pctN, 2), pC: this.fx(q.pctC, 2), cn: this.fx(q.cn, 1),
      chips: q.flags.map((f) => ({ t: this.chip(f), title: this.ftext(f) })),
      pick: () => H.pick(q.r), sbg: q.r === sel ? '#E9EFF5' : '#FFFFFF',
    }));
    const filters = [['all', 'All'], ['inside', 'Inside the size series'], ['size', 'Outside the size series'], ['carry', 'Carryover']].map(([k, t]) => ({
      t: t + ' ' + sets[k].length, pick: () => H.filter(k), bg: filt === k ? '#13212B' : '#FFFFFF', fg: filt === k ? '#FFFFFF' : '#3B4A55', bd: filt === k ? '#13212B' : '#C9D2DB',
    }));
    const both = S.filter((q) => q.c.fin != null && q.n.fin != null);
    const errs = [];
    for (const q of both) {
      if (q.c.u != null) errs.push([q.c.fin - q.c.u, q.n.fin, q.c.fin + q.c.u, q.n.fin]);
      if (q.n.u != null) errs.push([q.c.fin, q.n.fin - q.n.u, q.c.fin, q.n.fin + q.n.u]);
    }
    const chart = both.length ? this.chart({
      title: this.iso('n') + ' against ' + this.iso('c'), sub: both.length + ' samples, error bars 1σ; hollow where a peak is outside the size series', color: '#13212B', r: 4.4,
      x: this.dom(both.map((q) => q.c.fin), 0.06), y: this.dom(both.map((q) => q.n.fin), 0.06), errs: { segs: errs, stroke: '#8796A5' },
      pts: both.map((q) => ({ x: q.c.fin, y: q.n.fin, kind: size(q) ? 'hollow' : 'fill', ring: q.carry, sel: q.r === sel, label: q.id + ', position ' + q.r + ': ' + this.fx(q.n.fin, 2) + ', ' + this.fx(q.c.fin, 2) + '‰', pick: () => H.pick(q.r) })),
      legend: [{ kind: 'dot', color: '#13212B', label: 'Inside the size series' }, { kind: 'hollow', color: '#13212B', label: 'Outside it' }, { kind: 'ring', color: '#13212B', label: 'Carryover flag' }],
      xLabel: this.iso('c') + ' (‰ VPDB)', yLabel: this.iso('n') + ' (‰ AIR)', H: 380,
    }) : this.none(this.iso('n') + ' against ' + this.iso('c'), '#13212B', 'No sample has both a δ¹⁵N and a δ¹³C value.');
    const busy = st.busy, lo = st.boot && st.boot.libreoffice === false;
    const item = (kind, t, sub, note, wait) => ({
      t, sub, note: busy === kind ? wait : note, pick: () => this.download(kind),
      busy: !!busy, aria: busy === kind ? 'true' : 'false', cursor: busy ? (busy === kind ? 'progress' : 'default') : 'pointer',
    });
    const downloads = [
      item('workbook', 'Workbook', nm.workbook, lo ? 'Every step as a formula; Excel or LibreOffice fills in the values when it opens the file' : 'Every step as a formula, with the values filled in', 'Building the workbook on the server; this takes a few seconds'),
      item('csv', 'Results table', nm.csv, 'The workbook’s Results sheet as a CSV file, with every position at full precision', 'Building the results table on the server'),
      item('config', 'Configuration', nm.config, 'Your choices for this run as an Excel lab configuration, to edit and load next time', 'Writing the configuration on the server'),
    ];
    const many = st.runs.length > 1, bn = this.batchNames();
    if (many) {
      downloads.push(
        item('batch', 'Combined results, all ' + st.runs.length + ' runs', bn.xlsx, 'Every position of every run with a Run column first, one row per run, and the samples measured in more than one run, each run with its own choices', 'Correcting the ' + st.runs.length + ' runs on the server'),
        item('batchcsv', 'Combined results table', bn.csv, 'The combined Results sheet as a CSV file, at full precision', 'Correcting the ' + st.runs.length + ' runs on the server'));
    }
    const pair = (a, b) => (a && b ? 'on both' : a ? 'on ' + this.iso('n') : b ? 'on ' + this.iso('c') : 'off');
    const a = this.uSrc(M, d, 'n'), b = this.uSrc(M, d, 'c');
    const srcName = (k) => (k === 'monitors' ? 'the monitors' : k === 'config' ? 'the configuration' : 'the anchors');
    const chips = [
      'Blank ' + (M.blank.n.inForce || M.blank.c.inForce ? 'subtracted' : 'not subtracted'),
      'Linearity ' + pair(M.lin.n.inForce, M.lin.c.inForce),
      'Drift ' + pair(M.drift.n.inForce, M.drift.c.inForce),
      'Elemental ' + (M.elem.n.inForce || M.elem.c.inForce ? 'corrected' : 'not corrected'),
      'Repeatability from ' + (a === b ? srcName(a) : srcName(a) + ' for ' + this.iso('n') + ' and ' + srcName(b) + ' for ' + this.iso('c')),
    ].map((t) => ({ t }));
    return Object.assign({ rows, filters, chart, downloads, chips, methods: this.methods(M, d) }, this.batchView(H));
  }

  // The batch on the results screen: each run with its corrections and its own workbook, and the
  // samples measured in more than one run.
  batchView(H) {
    const st = this.state, runs = st.runs, busy = st.busy;
    if (runs.length < 2) return { multi: false, batch: { title: '', note: '', runs: [], reps: [], hasReps: false, noReps: '', repNote: '' } };
    const pair = (a, b) => (a && b ? 'both' : a ? this.iso('n') : b ? this.iso('c') : 'off');
    const list = runs.map((run, i) => {
      const M = this.model(this.decOf(run), run), S = this.samples(M), cur = run === st.run, k = 'workbook:' + i;
      const parts = ['Blank ' + (M.blank.n.inForce || M.blank.c.inForce ? 'subtracted' : 'not subtracted'),
        'linearity ' + pair(M.lin.n.inForce, M.lin.c.inForce), 'drift ' + pair(M.drift.n.inForce, M.drift.c.inForce),
        'elemental ' + (M.elem.n.inForce || M.elem.c.inForce ? 'corrected' : 'not corrected')];
      const chk = run.check;
      return {
        n: String(i + 1), name: run.name,
        meta: this.count(S.length, 'sample') + ' · ' + (this.dateOf(run) ? 'run of ' + this.fdate(this.dateOf(run)) : 'no run date') + (cur ? ' · shown now' : ''),
        how: parts.join(', '),
        status: chk && chk.bad.length ? 'Preview differs at ' + this.count(chk.bad.length, 'position') : 'Preview equals the server',
        sfg: chk && chk.bad.length ? '#A33A00' : '#0B6E4F',
        open: () => this.switchRun(i), openT: cur ? 'Shown' : 'Show', isCur: cur || !!busy,
        book: () => this.download('workbook', i), bookT: busy === (cur ? 'workbook' : k) ? 'Building…' : 'Workbook', bookBusy: !!busy,
        bg: cur ? '#F2F6F9' : '#FFFFFF',
      };
    });
    const reps = this.repeats(), rows = [];
    for (const g of reps) {
      const st2 = (v) => (v.length >= 2 ? { sd: this.fx(engine.sdev(v), 2), r: this.fx(Math.max(...v) - Math.min(...v), 2) } : { sd: '', r: '' });
      const mean = (v) => (v.length ? this.fx(v.reduce((a, b) => a + b, 0) / v.length, 2) : '');
      const a = st2(g.n), b = st2(g.c);
      rows.push({ name: g.name, runs: String(g.runs.size), n: String(g.count), mN: mean(g.n), sN: a.sd, rN: a.r, mC: mean(g.c), sC: b.sd, rC: b.r });
    }
    return {
      multi: true,
      batch: {
        title: 'All ' + runs.length + ' runs',
        note: 'Each run is corrected on its own, with the standards in that run and the choices made for it. Show a run to change its choices; the combined results take every run as it is left.',
        runs: list, reps: rows, hasReps: rows.length > 0,
        noReps: rows.length ? '' : 'No sample name appears in more than one run.',
        repNote: rows.length ? this.count(rows.length, 'sample') + ' measured in more than one run, matched by name. SD and range are over the analyses, in ‰; the combined results list each analysis with its difference from the mean.' : '',
      },
    };
  }

  railItems(M, d, scr, H) {
    const S = this.samples(M), A = M.P.filter((q) => q.role === 'A');
    const pair = (a, b, yes, no, n, c) => (a && b ? yes + ', both' : a ? yes + ', ' + n + ' only' : b ? yes + ', ' + c + ' only' : no);
    const checks = this.qcChecks(M, d);
    const nc = checks.filter((c) => c.s === 'caution').length, nn = checks.filter((c) => c.s === 'note').length;
    const sp = this.settleSpan(M);
    const list = [
      ['inst', 'Instrument', sp.off ? 'Settling check off' : sp.never ? 'Background never settled' : sp.late ? 'Settled at position ' + sp.sf : 'Steady from the start'],
      ['setup', 'Setup', this.count(S.length, 'sample') + ', ' + this.count(A.length, 'anchor analysis', 'anchor analyses')],
      ['blank', 'Blank', M.blank.n.inForce || M.blank.c.inForce ? 'Subtracted' : 'Not subtracted'],
      ['lin', 'Linearity', pair(M.lin.n.inForce, M.lin.c.inForce, 'Applied', 'Not applied', this.iso('n'), this.iso('c'))],
      ['drift', 'Drift', pair(M.drift.n.inForce, M.drift.c.inForce, 'Applied', 'Not applied', this.iso('n'), this.iso('c'))],
      ['cal', 'Calibration', M.cal.n.slope != null && M.cal.c.slope != null ? 'Slopes ' + this.fx(M.cal.n.slope, 3) + ' and ' + this.fx(M.cal.c.slope, 3) : 'Not calibrated'],
      ['elem', 'Elemental content', pair(M.elem.n.inForce, M.elem.c.inForce, 'Corrected', 'Not corrected', '%N', '%C')],
      ['unc', 'Uncertainty', '1σ(' + this.iso('n') + ') ' + this.uSpan(S, 'n', 2) + (this.uSpan(S, 'n', 2) === 'none' ? '' : '‰')],
      ['qc', 'Quality checks', (nc ? nc + ' caution' + (nc > 1 ? 's' : '') + ', ' : '') + nn + ' note' + (nn === 1 ? '' : 's')],
      ['results', 'Results', this.count(S.length, 'sample')],
    ];
    const ix = list.findIndex((x) => x[0] === scr);
    return list.map(([k, t, sub], i) => {
      const cur = k === scr, done = ix > i;
      return {
        n: String(i + 1), t, sub, pick: H.go(k), cur: cur ? 'step' : 'false',
        nbg: cur ? '#13212B' : done ? '#E3E9EF' : '#FFFFFF', nfg: cur ? '#FFFFFF' : '#13212B', nbd: cur ? '#13212B' : '#B8C3CD',
        rbg: cur ? '#FFFFFF' : 'transparent', rbd: cur ? '#D8DFE6' : 'transparent', tw: cur ? '600' : '500',
      };
    });
  }

  // ---------------------------------------------------------------- the upload screen
  upView(st, run, M, version) {
    const busy = !!st.busy, cfg = st.cfg;
    const pkg = 'Package defaults' + (version ? ', ea_irms ' + version : '');
    const up = {
      busy,
      dropTitle: st.busy === 'process' ? 'Reading ' + st.pending + '…' : run ? 'Drop more exports here to add them, or choose files' : 'Drop exports here, or choose files',
      dropNote: '.xls, .xlsx, .xlsm, .csv, .txt, or .tsv, as Isodat exported it. Several exports make a batch of up to ' + MAX_RUNS + ' runs, each corrected on its own.',
      dropBg: st.drag ? '#F2F7FB' : '#FFFFFF', dropBd: st.drag ? '#13212B' : '#9AA8B5', dropCursor: busy ? 'progress' : 'pointer',
      choose: () => { const el = document.getElementById('dc-export'); if (el && !this.state.busy) el.click(); },
      example: () => this.loadExample(), fresh: !run,
      picked: (e) => { const fs = Array.from(e.target.files || []); e.target.value = ''; if (fs.length) this.addExports(fs); },
      cfgName: cfg ? cfg.name : pkg, hasCfg: !!cfg, cfgButton: cfg ? 'Choose another configuration' : 'Choose a configuration',
      cfgNote: (cfg ? 'Its standards table, switches, and limits replace the package defaults' + (cfg.wrapped ? ' (read from its "config" field)' : '')
        + (cfg.unknown.length ? '. Fields ea_irms does not read are ignored: ' + this.andList(cfg.unknown) : '') + '.'
        + (cfg.notes && cfg.notes.length ? ' Note: ' + cfg.notes.join('; ') + '.' : '')
        : 'An Excel file with a standards table, switches, and limits: download the template, fill in its yellow cells, and choose it here. A .json configuration from an earlier run also works.')
        + (st.runs.length > 1 ? ' Choosing another processes all ' + st.runs.length + ' runs again.' : run ? ' Choosing another processes the run again.' : ''),
      chooseCfg: () => { const el = document.getElementById('dc-config'); if (el && !this.state.busy) el.click(); },
      pickedCfg: (e) => { const f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) this.pickConfig(f); },
      clearCfg: () => this.clearConfig(),
      template: () => this.downloadTemplate(),
      dateLabel: st.runs.length > 1 && run ? 'Run date of ' + run.name : 'Run date',
      dateValue: st.runDate || '',
      setDate: (e) => this.setState({ runDate: e.target.value, dateFrom: e.target.value ? 'you' : '' }),
      dateNote: st.dateFrom === 'name' && st.runDate ? 'Read from the file name; change it if it is wrong. It is printed on the workbook’s Setup sheet.'
        : 'Printed on the workbook’s Setup sheet; leave it empty if the run needs none.',
      infoLabel: st.runs.length > 1 && run ? 'Run information of ' + run.name : 'Run information',
      infoFields: RUN_FIELDS.map((k) => ({ k, label: k, id: 'dc-info-' + k.toLowerCase(), value: (st.runInfo || {})[k] || '',
        set: (e) => this.setInfo(k, e.target.value) })),
      infoNote: 'Written on the workbook’s Metadata sheet' + (st.runs.length > 1 ? ' and on the Runs sheet of the combined results. Each run keeps its own.'
        : ', and on the Runs sheet when several runs are combined.') + (run ? '' : ' Typed before an upload, it goes with each run of that upload.'),
      canInfoAll: st.runs.length > 1, infoAll: () => this.infoToAll(),
      sheet: this.sheetView(st, busy),
      hasError: !!st.error, errTitle: st.error ? st.error.title : '', errText: st.error ? st.error.detail : '',
      has: !!run, positions: '', date: st.runDate ? this.fdate(st.runDate) : 'no run date',
      heading: st.runs.length > 1 ? 'Correct ' + st.runs.length + ' runs' : 'Correct a run',
      multi: st.runs.length > 1, batchTitle: '', batchNote: '', runs: [],
    };
    if (!run) return up;
    if (st.runs.length > 1) {
      up.batchTitle = 'Batch of ' + st.runs.length + ' runs';
      up.batchNote = 'One lab configuration for all; each run keeps its own choices. The details below are for the run shown.';
      up.runs = st.runs.map((r, i) => {
        const cur = r === run, chk = r.check, date = this.dateOf(r);
        return {
          n: String(i + 1), name: r.name, cur: cur ? 'true' : 'false',
          meta: r.port.pos.length + ' positions · ' + this.count(r.port.pos.filter((q) => q.role === 'S').length, 'sample') + ' · ' + (date ? 'run of ' + this.fdate(date) : 'no run date'),
          status: chk && chk.bad.length ? 'Preview differs at ' + this.count(chk.bad.length, 'position') : 'Read without errors',
          sfg: chk && chk.bad.length ? '#A33A00' : '#0B6E4F', bg: cur ? '#F2F6F9' : '#FFFFFF',
          open: () => this.switchRun(i), openT: cur ? 'Shown' : 'Show', isCur: cur || busy,
          remove: () => this.removeRun(i), busy,
        };
      });
    }
    const S = this.samples(M), n = M.P.length, pc = run.port.pulseCheck || {}, chk = run.check;
    const icon = (good) => (good ? OK_ICON : WARN_ICON);
    const checks = [
      { t: pc.filtered ? 'A filtered export, without the reference peaks' + (pc.active ? '; ' + this.count(run.rejected.size, 'CO2 peak that matched', 'CO2 peaks that matched') + ' the designated reference pulses ' + (run.rejected.size === 1 ? 'was' : 'were') + ' set aside' : '')
        : 'A full export, with the N2 and CO2 reference peaks', good: true },
      { t: 'All twelve columns ea_irms reads were found by name (Columns, below)', good: true },
      { t: n + ' positions; the sample peak in each is found by its Is Ref flag', good: true },
      { t: this.count(S.length, 'sample') + ' and ' + (n - S.length) + ' standards, blanks, and other positions, matched by identifier against the ' + (run.cfgName ? 'configuration’s' : 'package') + ' standards table', good: true },
    ];
    if (chk) {
      checks.push(chk.bad.length
        ? { t: 'The values recalculated in this browser differ from the server’s at ' + this.count(chk.bad.length, 'position') + ' (' + this.rowList(chk.bad) + '); the workbook and the results table hold the server’s values', good: false }
        : { t: 'The values recalculated in this browser match the server’s at all ' + chk.positions + ' positions (' + chk.values + ' values and every flag)', good: true });
    }
    checks.push({ t: st.runDate ? 'Run date ' + this.fdate(st.runDate) + (st.dateFrom === 'name' ? ', read from the file name' : '') : 'No run date given; the workbook leaves it blank', good: true });
    return Object.assign(up, {
      name: run.name, positions: String(n),
      meta: this.bytes(run.bytes) + ' · ' + n + ' positions · ' + (st.runDate ? 'run of ' + this.fdate(st.runDate) : 'no run date'),
      badge: 'Read without errors',
      checks: checks.map((c) => Object.assign({ t: c.t }, icon(c.good))),
      mapNote: 'All twelve were found, so every column is read by name',
      mapFoot: 'ea_irms reads these twelve columns by their exact names and ignores the rest. An export that lacks one is refused with the names it lacks.',
      cols: COLUMNS.map(([field, unit, col]) => ({ field, unit, col, how: 'Found' })),
      cfgShort: run.cfgName || pkg,
      start: () => this.go(FLOW.includes(st.last) ? st.last : 'inst'),
    });
  }

  // The sample sheet on the upload screen: the file, its ID column, and how each sample matched.
  sheetView(st, busy) {
    const sh = st.sheet, multi = st.runs.length > 1;
    const v = {
      has: !!sh, name: sh ? sh.name : 'None chosen', meta: '', busy,
      button: sh ? 'Choose another sample sheet' : 'Choose a sample sheet',
      choose: () => { const el = document.getElementById('dc-sheet'); if (el && !this.state.busy) el.click(); },
      picked: (e) => { const f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) this.pickSheet(f); },
      clear: () => this.clearSheet(),
      template: () => this.downloadSheetTemplate(),
      templateLabel: st.runs.length ? 'Download a sheet with these sample names' : 'Download an empty sample sheet',
      note: 'A table with one row per sample: its ID, then any columns, such as project, site, species or material, fraction, '
        + 'collection date, depth, and notes. A sample matches a row when its name equals the ID, or begins with it followed by '
        + 'a space, a hyphen, or another character that is neither a letter nor a digit (CI 3-2 matches CI 3-2-S); the longest '
        + 'ID wins, and when no ID fits as written, spaces are ignored (PD1-H matches PD 1). The matched rows are written into '
        + 'the results tables and each workbook. Analyses of one sample in a run, '
        + 'with the same name or the same Replicate of, are compared, and a range above the critical range of ISO 5725-6 marks '
        + 'them for a rerun. The sheet stays in this tab and goes with each download.',
      cols: [], pickId: (e) => { if (sh && e.target.value !== sh.idColumn) this.pickSheet(sh.file, e.target.value); },
      summary: '', notes: '', hasNotes: false, written: '', rows: [], hasRows: false, more: '', hasMore: false,
      grid: multi ? 'minmax(0, 0.45fr) minmax(0, 0.7fr) minmax(0, 1.5fr) minmax(0, 1.3fr) minmax(0, 1.1fr) minmax(0, 0.8fr)'
        : 'minmax(0, 0.7fr) minmax(0, 1.5fr) minmax(0, 1.3fr) minmax(0, 1.1fr) minmax(0, 0.8fr)',
      runCell: multi ? 'block' : 'none',
    };
    if (!sh) return v;
    const info = sh.info, d = info.sheet;
    v.cols = d.columns.map((c) => ({ v: c, t: c, sel: c === sh.idColumn }));
    v.meta = this.count(d.rows, 'row') + ' · ID column ' + sh.idColumn + (d.sheet ? ' · sheet ' + d.sheet : '');
    let n = 0, ex = 0, pre = 0, sp = 0, groups = 0;
    const rows = [];
    info.runs.forEach((r, i) => {
      n += r.match.samples; ex += r.match.exact; pre += r.match.prefix; sp += r.match.spaces || 0; groups += r.groups.length;
      for (const q of r.samples) {
        const how = String(q.matched_by || '');
        rows.push({ run: String(i + 1), row: String(q.row), id: q.identifier, sid: q.sample_id || '–',
          how: !how ? 'Not in the sheet' : (how.startsWith('exact') ? 'Exact name' : 'ID at the start')
            + (how.endsWith('spaces ignored') ? ', spaces ignored' : ''),
          fg: how ? '#0B6E4F' : '#A33A00', group: q.group ? 'Group ' + q.group : '' });
      }
    });
    const none = n - ex - pre;
    v.summary = !st.runs.length ? 'Upload an export to match its samples to this sheet.'
      : (ex + pre) + ' of ' + this.count(n, 'sample') + ' matched' + (ex + pre ? ' (' + [ex ? ex + ' by the exact name' : '', pre ? pre + ' by an ID the name begins with' : ''].filter(Boolean).join(', ')
        + (sp ? '; ' + sp + ' of the ' + (ex + pre) + ' only with spaces ignored' : '') + ')' : '')
        + (none ? '; ' + none + ' not in the sheet' : '') + '. '
        + (groups ? this.count(groups, 'group') + ' of replicates in one run will be compared.' : 'No sample has replicates in one run.');
    v.notes = d.notes && d.notes.length ? 'Read with notes: ' + d.notes.join('; ') + '.' : '';
    v.hasNotes = !!v.notes;
    v.written = info.columns && info.columns.length ? 'Written after the results columns: ' + this.andList(info.columns) + '.' : '';
    v.rows = rows.slice(0, MAX_MATCH_ROWS);
    v.hasRows = rows.length > 0;
    v.more = rows.length > MAX_MATCH_ROWS ? 'The first ' + MAX_MATCH_ROWS + ' of ' + rows.length + ' samples are listed; the downloads hold all of them.' : '';
    v.hasMore = !!v.more;
    return v;
  }

  // ---------------------------------------------------------------- the view model
  renderVals() {
    const st = this.state, run = st.run, boot = st.boot || {};
    const version = (run && run.version) || boot.ea_irms || '';
    const scr = run ? st.screen : 'upload';
    const isFlow = FLOW.includes(scr);
    const is = { upload: scr === 'upload', flow: isFlow, step: isFlow && scr !== 'results' };
    for (const k of FLOW) is[k] = scr === k;
    const d = st.dec;
    const M = run ? this.model(d) : null;
    const go = (s) => () => this.go(s);
    const selMap = st.sel || {};
    const sel = M ? (selMap[scr] != null ? selMap[scr] : this.firstSel(M, scr)) : null;
    const H = {
      go,
      pick: (r) => this.setState({ sel: Object.assign({}, this.state.sel, { [scr]: r }) }),
      set: (k, v) => this.setState({ dec: Object.assign({}, this.state.dec, { [k]: v }) }),
      toggle: (k, r) => {
        const cur = this.state.dec[k], next = cur.includes(r) ? cur.filter((x) => x !== r) : cur.concat([r]).sort((a, b) => a - b);
        this.setState({ dec: Object.assign({}, this.state.dec, { [k]: next }), sel: Object.assign({}, this.state.sel, { [scr]: r }) });
      },
      filter: (k) => this.setState({ filt: k }),
    };
    const nav = [{ t: 'EA-IRMS', pick: () => this.go('upload'),
      fg: '#13212B', bd: '#13212B', fw: '600', cur: 'page' }];
    const up = this.upView(st, run, M, version);
    const set = M ? M.settings : engine.settings(null);
    const text = {
      alpha: this.num(set.alpha), k: this.num(set.k), elemGate: this.num(set.elemGate),
      blankFlag: set.blankMultiple ? 'Samples smaller than ' + this.times(set.blankMultiple) + ' the blank are flagged either way.' : 'The configuration turns off the flag for samples close to the blank.',
      linSeries: 'The size series', monitors: 'Drift monitors', anchors: 'The anchors',
    };
    let flow = { kicker: '', file: '', runLabel: 'Run', multi: false, runs: [], rail: [], ctrl: [], hasCtrl: false, ribbon: [], legend: [], ribbonStart: '', ribbonEnd: '', charts: [], hasCharts: false,
      stats: { rows: [] }, hasStats: false, statA: '', statB: '', effect: { title: '', lines: [] }, hasEffect: false, checks: [], hasChecks: false,
      budget: { title: '', rows: [] }, hasBudget: false, insp: { title: '', sub: '', facts: [], chain: [], tail: [], flags: [], hasFlags: false, noFlags: false },
      setupRows: [], prev: { t: '', pick: null }, next: { t: '', pick: null, busy: false }, stdSource: '', warn: false, warnText: '' };
    let res = { rows: [], filters: [], charts: [], downloads: [], chips: [], methods: '', multi: false,
      batch: { title: '', note: '', runs: [], reps: [], hasReps: false, noReps: '', repNote: '' } };
    if (M && isFlow) {
      const matL = this.material(M, 'L'), matD = this.material(M, 'D'), groups = this.anchorGroups(M);
      if (matL) text.linSeries = 'The ' + matL + ' size series';
      if (matD) text.monitors = this.cap(matD) + ' monitors';
      if (groups.length) text.anchors = this.andList(groups.map((g) => g.name));
      let view = { ctrl: [], charts: [], stats: null, effect: null, checks: [], budget: null };
      let setupRows = [];
      if (scr === 'inst') view = Object.assign(view, this.vInst(M, d, sel, H));
      else if (scr === 'setup') { const v = this.vSetup(M); setupRows = v.rows; view.effect = v.effect; }
      else if (scr === 'blank') view = Object.assign(view, this.vBlank(M, d, sel, H));
      else if (scr === 'lin') view = Object.assign(view, this.vLin(M, d, sel, H));
      else if (scr === 'drift') view = Object.assign(view, this.vDrift(M, d, sel, H));
      else if (scr === 'cal') view = Object.assign(view, this.vCal(M, d, sel, H));
      else if (scr === 'elem') view = Object.assign(view, this.vElem(M, d, sel, H));
      else if (scr === 'unc') view = Object.assign(view, this.vUnc(M, d, sel, H));
      else if (scr === 'qc') view = Object.assign(view, this.vQC(M, d, sel, H));
      else if (scr === 'results') {
        const r = this.vResults(M, d, sel, H, st.filt || 'all');
        res = Object.assign({}, r, { charts: r.chart ? [Object.assign({ num: '1' }, r.chart)] : [] });
      }
      const fi = FLOW.indexOf(scr), rows = M.P.map((q) => q.r);
      const chk = run.check;
      flow = {
        kicker: scr === 'results' ? 'After step ' + (FLOW.length - 1) : 'Step ' + (fi + 1) + ' of ' + (FLOW.length - 1), file: run.name,
        runLabel: st.runs.length > 1 ? 'Run ' + (st.active + 1) + ' of ' + st.runs.length : 'Run', multi: st.runs.length > 1,
        runs: st.runs.length > 1 ? st.runs.map((r, i) => ({ t: (i + 1) + '  ' + r.name, pick: () => this.switchRun(i), cur: r === run ? 'true' : 'false',
          bg: r === run ? '#13212B' : 'transparent', fg: r === run ? '#FFFFFF' : '#13212B', busy: !!st.busy })) : [],
        rail: this.railItems(M, d, scr, H),
        ctrl: view.ctrl, hasCtrl: view.ctrl.length > 0,
        ribbon: is.step ? this.ribbon(M, scr, sel, H) : [], legend: this.ribbonLegend(M),
        ribbonStart: 'Position ' + Math.min(...rows), ribbonEnd: String(Math.max(...rows)),
        charts: view.charts.map((c, i) => Object.assign({ num: String(i + 1) }, c)), hasCharts: view.charts.length > 0,
        stats: view.stats || { rows: [] }, hasStats: !!view.stats,
        statA: scr === 'elem' ? '%N' : this.iso('n'), statB: scr === 'elem' ? '%C' : this.iso('c'),
        effect: view.effect || { title: '', lines: [] }, hasEffect: !!view.effect,
        checks: view.checks.map((c) => Object.assign({ hasGo: false, goText: '', go: null }, c)), hasChecks: view.checks.length > 0,
        budget: view.budget || { title: '', rows: [] }, hasBudget: !!view.budget,
        insp: this.inspect(M, sel, scr), setupRows,
        prev: fi > 0 ? { t: 'Back to ' + NAMES[FLOW[fi - 1]], pick: go(FLOW[fi - 1]) } : { t: 'Back to upload', pick: go('upload') },
        next: scr === 'results'
          ? { t: st.busy === 'workbook' ? 'Building the workbook…' : 'Download the workbook', pick: () => this.download('workbook'), busy: !!st.busy }
          : { t: 'Continue to ' + NAMES[FLOW[fi + 1]], pick: go(FLOW[fi + 1]), busy: false },
        stdSource: run.cfgName ? 'From ' + run.cfgName + '.' : 'Package defaults, ea_irms ' + run.version + '.',
        warn: !!(chk && chk.bad.length),
        warnText: chk && chk.bad.length ? 'The values in this browser differ from the server’s at ' + this.count(chk.bad.length, 'position') + '; the downloads hold the server’s.' : '',
      };
    }
    return {
      is, nav, up, flow, res, text, version,
      footer: (boot.isoharness ? 'isoharness ' + boot.isoharness + (version ? ' with ea_irms ' + version : '') + '. ' : '')
        + 'Uploaded files are deleted as soon as they are processed, and no results are kept on the server. The values on these pages are recalculated in your browser; '
        + 'the workbook and the results table come from the server. MIT license.',
      toast: { show: !!st.toast, t: st.toast || '', close: () => this.setState({ toast: '' }) },
    };
  }
}

if (typeof window !== 'undefined' && window.DC && document.getElementById('dc-main')) {
  window.DC.mount(Component, document.getElementById('dc-main'), document.getElementById('app'));
}
