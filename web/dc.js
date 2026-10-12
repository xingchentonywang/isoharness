// isoharness page runtime: draws the prototype's Design-canvas template with Preact
// (web/vendor/preact, 10.29.8 as published on npm).
//
// The template is the markup in <template id="dc-main"> of index.html:
//
//   <sc-if value="{{path}}">          its children, when the value is truthy
//   <sc-for list="{{path}}" as="x">   its children once per item, the item named x
//   {{path}}                          a value in text or in an attribute, style included
//   onclick="{{path}}" (any on...)    a function of the view model, bound as a listener
//   style-hover="..."                 declarations applied while the pointer is over it
//
// A path is a dotted name looked up in the view model (the logic's renderVals()) and
// in the sc-for names around it. Nothing is evaluated, so the page runs under the
// service's Content-Security-Policy, which allows no 'unsafe-eval' and no inline
// script. The template is compiled once; after each setState the view model is
// built again and Preact brings the page up to date.
(function () {
  'use strict';
  const { h, render, Fragment } = window.preact;
  const HOLE = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const WHOLE = /^\{\{\s*([^{}]+?)\s*\}\}$/;
  // Attributes that are on when present, whatever their text.
  const FLAGS = new Set(['hidden', 'disabled', 'checked', 'multiple', 'required', 'readonly',
    'open', 'selected', 'autofocus', 'novalidate']);
  const hovers = new Map();

  function ref(text) {
    const t = text.trim();
    if (t === 'true' || t === 'false') return { lit: t === 'true' };
    return { keys: t.split('.') };
  }

  function holes(text) {
    const out = [];
    let last = 0;
    let m;
    HOLE.lastIndex = 0;
    while ((m = HOLE.exec(text))) {
      if (m.index > last) out.push(text.slice(last, m.index));
      out.push(ref(m[1]));
      last = HOLE.lastIndex;
    }
    if (last < text.length) out.push(text.slice(last));
    return out;
  }

  function look(scope, r) {
    if (!r.keys) return r.lit;
    let v = scope[r.keys[0]];
    for (let i = 1; i < r.keys.length && v != null; i++) v = v[r.keys[i]];
    return v;
  }

  function fill(scope, parts) {
    let s = '';
    for (const p of parts) {
      if (typeof p === 'string') s += p;
      else {
        const v = look(scope, p);
        if (v != null) s += v;
      }
    }
    return s;
  }

  function only(value, where) {
    const m = WHOLE.exec(value || '');
    if (!m) throw new Error('the template needs a single {{path}} in ' + where);
    return ref(m[1]);
  }

  function hoverClass(css) {
    if (!hovers.has(css)) hovers.set(css, 'dc-hover-' + hovers.size);
    return hovers.get(css);
  }

  // One rule per distinct style-hover. !important lets it win over the element's
  // own style attribute, which is where the prototype keeps every other style.
  function hoverSheet() {
    const rules = [];
    for (const [css, cls] of hovers) {
      const decl = css.split(';').map((d) => d.trim()).filter(Boolean).map((d) => d + ' !important').join('; ');
      rules.push('.' + cls + ':hover:not(:disabled) { ' + decl + ' }');
    }
    const el = document.createElement('style');
    el.textContent = rules.join('\n');
    document.head.appendChild(el);
  }

  function compile(node) {
    if (node.nodeType === 3) {
      const t = node.nodeValue;
      return t.indexOf('{{') < 0 ? { k: 'text', text: t } : { k: 'hole', parts: holes(t) };
    }
    if (node.nodeType !== 1) return null;
    const tag = node.localName;
    const kids = compileKids(node);
    if (tag === 'sc-if') return { k: 'if', test: only(node.getAttribute('value'), 'sc-if'), kids };
    if (tag === 'sc-for') {
      const as = node.getAttribute('as');
      if (!as) throw new Error('sc-for needs an as name');
      return { k: 'for', list: only(node.getAttribute('list'), 'sc-for'), as, kids };
    }
    const attrs = [];
    let hover = null;
    for (const a of Array.from(node.attributes)) {
      if (a.name.startsWith('hint-')) continue;
      if (a.name === 'style-hover') { hover = hoverClass(a.value); continue; }
      const whole = WHOLE.exec(a.value);
      if (whole) attrs.push({ name: a.name, whole: ref(whole[1]) });
      else if (a.value.indexOf('{{') >= 0) attrs.push({ name: a.name, parts: holes(a.value) });
      else attrs.push({ name: a.name, value: FLAGS.has(a.name) && a.value === '' ? true : a.value });
    }
    if (hover) {
      const cls = attrs.find((a) => a.name === 'class');
      if (!cls) attrs.push({ name: 'class', value: hover });
      else if (cls.parts) cls.parts.push(' ' + hover);
      else if (cls.whole) { cls.parts = [cls.whole, ' ' + hover]; delete cls.whole; }
      else cls.value += ' ' + hover;
    }
    return { k: 'el', tag, attrs, kids };
  }

  function compileKids(node) {
    const from = node.localName === 'template' ? node.content : node;
    const out = [];
    for (const c of Array.from(from.childNodes)) {
      const n = compile(c);
      if (n) out.push(n);
    }
    return out;
  }

  function build(n, scope, out) {
    if (n.k === 'text') { out.push(n.text); return; }
    if (n.k === 'hole') { out.push(fill(scope, n.parts)); return; }
    if (n.k === 'if') {
      if (look(scope, n.test)) for (const c of n.kids) build(c, scope, out);
      return;
    }
    if (n.k === 'for') {
      const list = look(scope, n.list);
      if (!list) return;
      for (const item of list) {
        const s = Object.create(scope);
        s[n.as] = item;
        for (const c of n.kids) build(c, s, out);
      }
      return;
    }
    const props = {};
    for (const a of n.attrs) {
      let v;
      if (a.whole) {
        v = look(scope, a.whole);
        if (v == null) continue;
      } else if (a.parts) v = fill(scope, a.parts);
      else v = a.value;
      props[a.name] = v;
    }
    const kids = [];
    for (const c of n.kids) build(c, scope, kids);
    out.push(h(n.tag, props, kids));
  }

  // The base class of the page's logic, in the shape the prototype was written
  // for: props, state, setState (merged, React style) and the two lifecycle hooks.
  class DCLogic {
    constructor(props) {
      this.props = props || {};
      this.state = null;
      this._redraw = null;
    }

    setState(next) {
      const add = typeof next === 'function' ? next(this.state || {}, this.props) : next;
      this.state = Object.assign({}, this.state || {}, add);
      if (this._redraw) this._redraw();
    }

    forceUpdate() {
      if (this._redraw) this._redraw();
    }
  }

  function mount(Logic, template, root, props) {
    const tree = compileKids(template);
    hoverSheet();
    const logic = new Logic(props || {});
    const paint = () => {
      try {
        const out = [];
        for (const n of tree) build(n, logic.renderVals(), out);
        render(h(Fragment, null, out), root);
      } catch (err) {
        console.error(err);
        render(h('div', { role: 'alert', style: 'max-width: 720px; margin: 48px auto; padding: 16px 20px; font: 15px/1.6 system-ui, sans-serif; color: #22303D; background: #FDF1EC; border: 1px solid #E3B3A0; border-radius: 10px' },
          'The page could not be drawn (' + (err && err.message ? err.message : String(err)) + '). Reload it to start again; nothing was stored.'), root);
      }
    };
    let queued = false;
    let before = logic.state;
    logic._redraw = () => {
      if (queued) return;
      queued = true;
      Promise.resolve().then(() => {
        queued = false;
        const prev = before;
        paint();
        before = logic.state;
        if (logic.componentDidUpdate) logic.componentDidUpdate(logic.props, prev);
      });
    };
    root.textContent = '';    // the loading note in index.html
    paint();
    before = logic.state;
    if (logic.componentDidMount) logic.componentDidMount();
    return logic;
  }

  window.DCLogic = DCLogic;
  window.DC = { mount, compile: compileKids };
})();
