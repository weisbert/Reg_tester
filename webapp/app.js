/* app.js — Reg_tester 信号流 GUI（M2 骨架 + M3 业务）。ES6，保守语法。
 * 数据来源：serve 模式走 /api/*；bundle 模式走 window.__BUNDLE__。
 * 预览序列用本地 Generator（与 gen_testcase.py 逐字节一致）；serve 模式导出走 Python 写盘。
 */
(function () {
  'use strict';
  var SVGNS = 'http://www.w3.org/2000/svg';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  // ---------------------------------------------------------------- backend
  var BUNDLE = window.__BUNDLE__ || null;
  var Backend = {
    bundle: !!BUNDLE,
    boot: function () {
      if (BUNDLE) {
        return Promise.resolve({
          project: BUNDLE.project, flowgraph: BUNDLE.flowgraph, regmap: BUNDLE.regmap,
          layout: BUNDLE.layout,
          modes: Object.keys(BUNDLE.modes || {}).filter(function (id) { return BUNDLE.modes[id]; }).map(function (id) {
            var m = BUNDLE.modes[id]; return { id: id, name: m.name || id, reg_group: m.reg_group || 'BT' };
          }), backend: 'bundle'
        });
      }
      return fetch('/api/bootstrap').then(function (r) { return r.json(); });
    },
    readMode: function (id) {
      if (BUNDLE) return Promise.resolve(BUNDLE.modes[id] || null);
      return fetch('/api/mode/' + encodeURIComponent(id)).then(function (r) { return r.ok ? r.json() : null; });
    },
    saveMode: function (id, obj) {
      if (BUNDLE) { BUNDLE.modes[id] = obj; return Promise.resolve({ ok: true, local: true }); }
      return fetch('/api/mode/' + encodeURIComponent(id), {
        method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(obj)
      }).then(function (r) { return r.json(); });
    },
    deleteMode: function (id) {
      if (BUNDLE) { delete BUNDLE.modes[id]; return Promise.resolve({ ok: true }); }
      return fetch('/api/mode/' + encodeURIComponent(id), { method: 'DELETE' }).then(function (r) { return r.json(); });
    },
    saveLayout: function (obj) {
      if (BUNDLE) { BUNDLE.layout = obj; return Promise.resolve({ ok: true, local: true }); }
      return fetch('/api/layout', {
        method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(obj)
      }).then(function (r) { return r.json(); });
    },
    exportSeq: function (id) {
      if (BUNDLE) return Promise.resolve(null); // bundle：客户端下载
      return fetch('/api/export/' + encodeURIComponent(id), { method: 'POST' }).then(function (r) { return r.json(); });
    },
    matching: function () {
      if (BUNDLE) return Promise.resolve(null); // bundle 只读：无匹配重建
      return fetch('/api/matching').then(function (r) { return r.ok ? r.json() : null; });
    },
    saveMatching: function (payload) {
      if (BUNDLE) return Promise.resolve({ ok: false, error: 'bundle 只读' });
      return fetch('/api/matching', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
      }).then(function (r) { return r.json(); });
    },
    // ---- P2.2 建库向导 ----
    _post: function (url, body) {
      if (BUNDLE) return Promise.resolve({ ok: false, error: 'bundle 只读，建库需 --serve' });
      return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body || {}) }).then(function (r) { return r.json(); });
    },
    saveConfig: function (cfg) { return this._post('/api/project/config', cfg); },
    importNetlist: function (path) { return this._post('/api/import/netlist', { path: path }); },
    importExcel: function (path, sheet, rowdump) { return this._post('/api/import/excel', { path: path, sheet: sheet, rowdump: rowdump }); },
    saveControlSignals: function (obj) { return this._post('/api/control_signals', obj); },
    build: function () { return this._post('/api/build', {}); },
    projectExport: function () {
      if (BUNDLE) return Promise.resolve({ project: BUNDLE.project, control_signals: {}, modes: BUNDLE.modes || {} });
      return fetch('/api/project/export').then(function (r) { return r.json(); });
    },
    pickFile: function (kind) {
      if (BUNDLE) return Promise.resolve({ path: null });
      return fetch('/api/pick-file?kind=' + encodeURIComponent(kind)).then(function (r) { return r.json(); });
    }
  };

  // ---------------------------------------------------------------- state
  var S = {
    project: null, fg: null, rm: null, layout: null, modeList: [],
    mode: null, modeId: null,
    view: { x: 40, y: 40, k: 0.9 },
    sel: {}, // selected node ids
    lay: null, // computed layout
    nodeById: {}, sigById: {},
    activeTab: 'inspector',
    recording: false,
    undo: [], redo: [],
    lastTc: null,
    match: null, matchLoading: false,
    needsSetup: false, wiz: null   // P2.2 建库向导本地态
  };

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.classList.remove('show'); }, 1800);
  }

  // ---------------------------------------------------------------- boot
  Backend.boot().then(function (b) {
    S.project = b.project; S.fg = b.flowgraph; S.rm = b.regmap; S.layout = b.layout || defaultLayout();
    S.modeList = b.modes || [];
    (S.fg.nodes || []).forEach(function (n) { S.nodeById[n.id] = n; });
    (S.rm.signals || []).forEach(function (s) { S.sigById[s.id] = s; });
    $('#proj-name').textContent = (S.project && S.project.name) ? ('  ·  ' + S.project.name + (Backend.bundle ? '  (bundle)' : '')) : '';
    S.needsSetup = !!b.needs_setup;
    initGroups();
    initModeSelect();
    bindUI();
    if (S.modeList.length) loadMode(S.modeList[0].id);
    else newModeFromScratch();
    relayout(true);
    fit();
    if (S.needsSetup && !Backend.bundle) { switchTab('project'); toast('新工程：在「工程」标签导入 netlist + Excel 建库'); }
  }).catch(function (e) { toast('加载失败: ' + e); console.error(e); });

  function defaultLayout() {
    return { schema_version: 'layout/1', positions: {}, collapsed: [], hidden: [], shown: [],
             symbol_override: {}, edge_flip: {}, notes: {}, expanded: [] };
  }

  function initGroups() {
    var sel = $('#group-select');
    (S.fg.reg_groups || ['BT', 'WL', 'WLT']).forEach(function (g) {
      var o = document.createElement('option'); o.value = g; o.textContent = g; sel.appendChild(o);
    });
  }
  function initModeSelect() {
    var sel = $('#mode-select'); sel.innerHTML = '';
    S.modeList.forEach(function (m) {
      var o = document.createElement('option'); o.value = m.id; o.textContent = m.name || m.id; sel.appendChild(o);
    });
  }

  // ---------------------------------------------------------------- mode load/save
  function loadMode(id) {
    return Backend.readMode(id).then(function (m) {
      if (!m) { toast('模式不存在: ' + id); return; }
      S.mode = normMode(m); S.modeId = id;
      $('#mode-select').value = id;
      $('#group-select').value = S.mode.reg_group;
      $('#order-select').value = (S.mode.order && S.mode.order.mode) || 'auto';
      pushUndo();
      render(); renderTabs(); toast('已载入 ' + id);
    });
  }
  function normMode(m) {
    m.schema_version = 'modes/1';
    m.enabled_nodes = m.enabled_nodes || [];
    m.baseline = m.baseline || {};
    m.mux_sel = m.mux_sel || {};
    m.order = m.order || { mode: 'auto', manual: [] };
    if (!m.order.manual) m.order.manual = [];
    m.extra_writes = m.extra_writes || [];
    m.reg_group = m.reg_group || 'BT';
    return m;
  }
  function newModeFromScratch() {
    S.mode = normMode({ id: 'NEW_MODE', name: 'New mode', reg_group: 'BT', enabled_nodes: [] });
    S.modeId = null; render(); renderTabs();
  }

  function saveAll() {
    S.mode.reg_group = $('#group-select').value;
    S.mode.order.mode = $('#order-select').value;
    var id = S.modeId || S.mode.id;
    if (!id || id === 'NEW_MODE') { promptNewId(); return; }
    Promise.all([Backend.saveMode(id, S.mode), Backend.saveLayout(S.layout)]).then(function () {
      S.modeId = id;
      if (!S.modeList.some(function (m) { return m.id === id; })) {
        S.modeList.push({ id: id, name: S.mode.name, reg_group: S.mode.reg_group }); initModeSelect(); $('#mode-select').value = id;
      }
      toast('已保存 ' + id + (Backend.bundle ? '（bundle：仅内存，用导出下载）' : ''));
    });
  }
  function promptNewId() {
    var id = window.prompt('模式 id（字母数字下划线）', S.mode.id === 'NEW_MODE' ? '' : S.mode.id);
    if (!id) return; id = id.replace(/[^A-Za-z0-9_.\-]/g, '_');
    S.mode.id = id; S.mode.name = window.prompt('模式显示名', S.mode.name || id) || id;
    S.modeId = id; saveAll();
  }

  // ---------------------------------------------------------------- undo/redo
  function snap() { return JSON.stringify({ mode: S.mode, layout: S.layout }); }
  function pushUndo() { S.undo.push(snap()); if (S.undo.length > 60) S.undo.shift(); S.redo = []; }
  function applySnap(s) { var o = JSON.parse(s); S.mode = o.mode; S.layout = o.layout; relayout(false); render(); renderTabs(); }
  function undo() { if (!S.undo.length) return; S.redo.push(snap()); applySnap(S.undo.pop()); toast('撤销'); }
  function redo() { if (!S.redo.length) return; S.undo.push(snap()); applySnap(S.redo.pop()); toast('重做'); }
  function mutate(fn) { pushUndo(); fn(); }

  // ---------------------------------------------------------------- layout
  function relayout(resetPositions) {
    if (resetPositions) S.layout.positions = {};
    S.lay = Layout.compute(S.fg, S.layout);
    render();
  }

  // ---------------------------------------------------------------- render graph
  function svgEl(tag, attrs) {
    var e = document.createElementNS(SVGNS, tag);
    if (attrs) for (var k in attrs) if (attrs.hasOwnProperty(k)) e.setAttribute(k, attrs[k]);
    return e;
  }
  var DEVICE_CLR = { dco: '--dco', div: '--div', buf: '--buf', inv: '--buf', mux: '--mux', logic: '--logic', route: '--route', blackbox: '--mux' };

  function resolveBox(id) {
    if (S.lay.nodes[id]) return S.lay.nodes[id];
    // 折到可见祖先（group box）
    var cur = id, seen = {};
    while (cur && !seen[cur]) {
      seen[cur] = true;
      if (S.lay.nodes[cur]) return S.lay.nodes[cur];   // 折叠后的 composite 以叶子形式存在 nodes 里
      var g = S.lay.groups.filter(function (x) { return x.id === cur; })[0];
      if (g) return g;
      var n = S.nodeById[cur]; cur = n ? n.parent : null;
    }
    return null;
  }

  function render() {
    if (!S.lay) return;
    var svg = $('#graph'); svg.innerHTML = '';
    var defs = svgEl('defs');
    defs.innerHTML = '<marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="var(--route)"/></marker>';
    svg.appendChild(defs);
    var root = svgEl('g', { id: 'vp' });
    root.setAttribute('transform', 'translate(' + S.view.x + ',' + S.view.y + ') scale(' + S.view.k + ')');
    svg.appendChild(root);

    var enabled = setOf(S.mode ? S.mode.enabled_nodes : []);

    // group boxes first
    S.lay.groups.forEach(function (g) {
      var r = svgEl('rect', { class: 'grp ' + g.kind, x: g.x, y: g.y, width: g.w, height: g.h, rx: 12 });
      root.appendChild(r);
      var collapsed = S.layout.collapsed && S.layout.collapsed.indexOf(g.id) >= 0;
      var t = svgEl('text', { class: 'grp-label', x: g.x + 12, y: g.y + 18, 'data-group': g.id });
      t.style.cursor = 'pointer';
      t.textContent = (g.kind === 'module' ? '▣ ' : (collapsed ? '▸ ' : '▾ ')) + g.label
        + (g.band ? '  [' + g.band + ']' : '') + (g.kind === 'composite' ? '  (双击折叠)' : '');
      root.appendChild(t);
    });

    // edges
    (S.fg.edges || []).forEach(function (e) {
      var a = resolveBox(e.from.node), bTo = e.to && e.to[0] ? resolveBox(e.to[0].node) : null;
      if (!a || !bTo) return;
      var flip = S.layout.edge_flip && S.layout.edge_flip[e.id];
      var src = flip ? bTo : a, dst = flip ? a : bTo;
      var p = edgePath(src, dst);
      var cls = 'edge';
      if (e.differential) cls += ' diff';
      if (e.cross_module) cls += ' cross';
      var onPath = enabled[e.from.node] && e.to.some(function (t) { return enabled[t.node]; });
      if (onPath) cls += ' path';
      var path = svgEl('path', { class: cls, d: p });
      path.setAttribute('data-edge', e.id);
      root.appendChild(path);
    });

    // nodes
    Object.keys(S.lay.nodes).forEach(function (id) {
      var n = S.nodeById[id]; if (!n) return;
      var b = S.lay.nodes[id];
      var g = svgEl('g', { class: nodeClass(n, enabled), 'data-id': id, transform: 'translate(' + b.x + ',' + b.y + ')' });
      g.appendChild(svgEl('rect', { class: 'box', width: b.w, height: b.h, rx: 9 }));
      g.appendChild(glyph(n));
      var lbl = svgEl('text', { class: 'lbl', x: 46, y: 22 }); lbl.textContent = short(n.inst_name || n.name || id.split('::').pop(), 13); g.appendChild(lbl);
      var sub = svgEl('text', { class: 'sub', x: 46, y: 37 }); sub.textContent = n.device + (n.inferred ? ' · 推断' : ''); g.appendChild(sub);
      var off = (n.off_controls || []).length;
      if (off) { var o2 = svgEl('text', { class: 'sub', x: 46, y: 50 }); o2.textContent = '⏻ ' + off + ' 门'; g.appendChild(o2); }
      if (n.warn || n.inferred) { var bd = svgEl('text', { class: 'badge', x: b.w - 14, y: 16 }); bd.textContent = n.warn ? '⚠' : '◌'; g.appendChild(bd); }
      root.appendChild(g);
    });

    updateZoomLabel(); renderMinimap();
  }

  function nodeClass(n, enabled) {
    var c = 'node';
    if (S.sel[n.id]) c += ' sel';
    if (enabled[n.id]) c += ' enabled path';
    if (n.inferred) c += ' inferred';
    return c;
  }
  function glyph(n) {
    var g = svgEl('g', { class: 'glyph', transform: 'translate(10,14)' });
    var clr = 'var(' + (DEVICE_CLR[n.device] || '--route') + ')';
    var d = n.device;
    if (d === 'buf' || d === 'inv') {
      g.appendChild(svgEl('path', { d: 'M2 2 L26 14 L2 26 Z', fill: clr, opacity: .85 }));
      if (d === 'inv') g.appendChild(svgEl('circle', { cx: 29, cy: 14, r: 3.2, fill: 'none', stroke: clr, 'stroke-width': 2 }));
    } else if (d === 'mux') {
      g.appendChild(svgEl('path', { d: 'M2 2 L26 8 L26 20 L2 26 Z', fill: clr, opacity: .85 }));
    } else if (d === 'div') {
      g.appendChild(svgEl('rect', { x: 2, y: 4, width: 26, height: 20, rx: 3, fill: clr, opacity: .85 }));
      var t = svgEl('text', { x: 15, y: 18, 'text-anchor': 'middle', 'font-size': 10, fill: '#0b0f14', 'font-weight': 700 }); t.textContent = '÷'; g.appendChild(t);
    } else if (d === 'dco') {
      g.appendChild(svgEl('circle', { cx: 15, cy: 14, r: 13, fill: clr, opacity: .85 }));
      g.appendChild(svgEl('path', { d: 'M7 14 q4 -7 8 0 t8 0', fill: 'none', stroke: '#0b0f14', 'stroke-width': 2 }));
    } else if (d === 'logic') {
      g.appendChild(svgEl('rect', { x: 2, y: 4, width: 26, height: 20, rx: 5, fill: clr, opacity: .7 }));
    } else {
      g.appendChild(svgEl('rect', { x: 2, y: 6, width: 26, height: 16, rx: 8, fill: clr, opacity: .7 }));
    }
    return g;
  }
  function edgePath(a, b) {
    var x1 = a.x + a.w, y1 = a.y + a.h / 2, x2 = b.x, y2 = b.y + b.h / 2;
    if (b.x < a.x) { x1 = a.x + a.w / 2; x2 = b.x + b.w / 2; y1 = a.y + a.h; y2 = b.y + b.h; }
    var dx = Math.max(40, Math.abs(x2 - x1) * 0.4);
    return 'M' + x1 + ' ' + y1 + ' C' + (x1 + dx) + ' ' + y1 + ' ' + (x2 - dx) + ' ' + y2 + ' ' + x2 + ' ' + y2;
  }
  function short(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
  function setOf(a) { var o = {}; (a || []).forEach(function (x) { o[x] = true; }); return o; }

  // ---------------------------------------------------------------- view transform
  function updateZoomLabel() { $('#zoom-label').textContent = Math.round(S.view.k * 100) + '%'; }
  function fit() {
    if (!S.lay || !S.lay.size) return;
    var wrap = $('#canvas-wrap'), W = wrap.clientWidth, H = wrap.clientHeight;
    var k = Math.min(W / (S.lay.size.w + 40), H / (S.lay.size.h + 40), 1.4);
    S.view.k = Math.max(0.15, k); S.view.x = 20; S.view.y = 20; render();
  }
  function zoom(f, cx, cy) {
    var wrap = $('#canvas-wrap');
    cx = cx === undefined ? wrap.clientWidth / 2 : cx; cy = cy === undefined ? wrap.clientHeight / 2 : cy;
    var k2 = Math.max(0.12, Math.min(3, S.view.k * f));
    S.view.x = cx - (cx - S.view.x) * (k2 / S.view.k);
    S.view.y = cy - (cy - S.view.y) * (k2 / S.view.k);
    S.view.k = k2; applyView();
  }
  function applyView() { var vp = $('#vp'); if (vp) vp.setAttribute('transform', 'translate(' + S.view.x + ',' + S.view.y + ') scale(' + S.view.k + ')'); updateZoomLabel(); renderMinimap(); }

  // ---------------------------------------------------------------- minimap
  function renderMinimap() {
    if (!S.lay) return;
    var mm = $('#minimap'); var sw = S.lay.size.w, sh = S.lay.size.h;
    var s = Math.min(180 / sw, 120 / sh);
    var parts = ['<svg viewBox="0 0 180 120">'];
    S.lay.groups.forEach(function (g) { if (g.kind === 'module') parts.push('<rect x="' + (g.x * s) + '" y="' + (g.y * s) + '" width="' + (g.w * s) + '" height="' + (g.h * s) + '" fill="none" stroke="var(--line)"/>'); });
    Object.keys(S.lay.nodes).forEach(function (id) { var b = S.lay.nodes[id]; parts.push('<rect x="' + (b.x * s) + '" y="' + (b.y * s) + '" width="' + (b.w * s) + '" height="' + (b.h * s) + '" fill="var(--mut)"/>'); });
    // viewport rect
    var wrap = $('#canvas-wrap');
    var vx = (-S.view.x / S.view.k) * s, vy = (-S.view.y / S.view.k) * s;
    var vw = (wrap.clientWidth / S.view.k) * s, vh = (wrap.clientHeight / S.view.k) * s;
    parts.push('<rect x="' + vx + '" y="' + vy + '" width="' + vw + '" height="' + vh + '" fill="rgba(110,168,254,.18)" stroke="var(--acc)"/>');
    parts.push('</svg>'); mm.innerHTML = parts.join('');
  }

  // ---------------------------------------------------------------- interactions
  function screenToWorld(px, py) { return { x: (px - S.view.x) / S.view.k, y: (py - S.view.y) / S.view.k }; }

  function bindUI() {
    var svg = $('#graph');
    var panning = false, dragNode = null, dragStart = null, movedNodes = null, boxSel = null, moved = false;

    svg.addEventListener('mousedown', function (ev) {
      hideCtx();
      var g = ev.target.closest ? ev.target.closest('.node') : null;
      var rect = svg.getBoundingClientRect();
      var mx = ev.clientX - rect.left, my = ev.clientY - rect.top;
      if (g) {
        var id = g.getAttribute('data-id');
        if (!S.sel[id]) { if (!ev.shiftKey) S.sel = {}; S.sel[id] = true; render(); onSelect(id); }
        else if (ev.shiftKey) { delete S.sel[id]; render(); }
        // start drag of selected nodes
        dragNode = id; moved = false; dragStart = screenToWorld(mx, my);
        movedNodes = Object.keys(S.sel);
        S._preDrag = JSON.stringify(S.layout.positions);
      } else {
        if (ev.shiftKey) { boxSel = { x0: mx, y0: my }; }
        else { panning = true; svg.classList.add('panning'); dragStart = { x: mx - S.view.x, y: my - S.view.y }; }
      }
    });
    window.addEventListener('mousemove', function (ev) {
      var rect = svg.getBoundingClientRect();
      var mx = ev.clientX - rect.left, my = ev.clientY - rect.top;
      if (panning) { S.view.x = mx - dragStart.x; S.view.y = my - dragStart.y; applyView(); }
      else if (dragNode) {
        var w = screenToWorld(mx, my); var dx = w.x - dragStart.x, dy = w.y - dragStart.y;
        if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
        movedNodes.forEach(function (id) {
          var b = S.lay.nodes[id]; if (!b) return;
          var base = S.layout.positions[id] || { x: b._auto.x, y: b._auto.y };
          // store target based on pre-drag start
        });
        // simpler: move all selected by delta from their current lay position
        movedNodes.forEach(function (id) {
          var b = S.lay.nodes[id]; if (!b) return;
          b.x = (S.layout.positions[id] ? S.layout.positions[id].x : b._auto.x) + dx;
          b.y = (S.layout.positions[id] ? S.layout.positions[id].y : b._auto.y) + dy;
        });
        render();
      } else if (boxSel) {
        boxSel.x1 = mx; boxSel.y1 = my; drawBoxSel(boxSel);
      }
    });
    window.addEventListener('mouseup', function (ev) {
      if (dragNode && moved) {
        mutate(function () {
          movedNodes.forEach(function (id) { var b = S.lay.nodes[id]; if (b) S.layout.positions[id] = { x: b.x, y: b.y }; });
        });
        saveLayoutQuiet();
      } else if (dragNode && !moved) { /* pure click already handled */ }
      if (boxSel && boxSel.x1 !== undefined) {
        var x0 = Math.min(boxSel.x0, boxSel.x1), x1 = Math.max(boxSel.x0, boxSel.x1);
        var y0 = Math.min(boxSel.y0, boxSel.y1), y1 = Math.max(boxSel.y0, boxSel.y1);
        if (!ev.shiftKey) S.sel = {};
        Object.keys(S.lay.nodes).forEach(function (id) {
          var b = S.lay.nodes[id];
          var sx = b.x * S.view.k + S.view.x, sy = b.y * S.view.k + S.view.y;
          if (sx > x0 && sx < x1 && sy > y0 && sy < y1) S.sel[id] = true;
        });
        render();
      }
      panning = false; dragNode = null; boxSel = null; svg.classList.remove('panning'); clearBoxSel();
    });

    svg.addEventListener('click', function (ev) {
      var g = ev.target.closest ? ev.target.closest('.node') : null;
      if (!g) {
        var grp = ev.target.closest ? ev.target.closest('[data-group]') : null;
        if (grp) { var gid = grp.getAttribute('data-group'); if (!ev.shiftKey) S.sel = {}; S.sel[gid] = true; render(); onSelect(gid); return; }
        if (!ev.shiftKey) { S.sel = {}; render(); onSelect(null); } return;
      }
      var id = g.getAttribute('data-id');
      onSelect(id);
      if (S.recording) recordClick(id);
      else if (S.activeTab === 'mode') toggleEnabled(id);
    });
    svg.addEventListener('dblclick', function (ev) {
      var g = ev.target.closest ? ev.target.closest('.node') : null;
      var grp = (!g && ev.target.closest) ? ev.target.closest('[data-group]') : null;
      var id = g ? g.getAttribute('data-id') : (grp ? grp.getAttribute('data-group') : null);
      if (!id) return;
      var n = S.nodeById[id];
      if (n && n.kind === 'composite') { toggleCollapse(id); }
    });
    svg.addEventListener('contextmenu', function (ev) {
      ev.preventDefault();
      var g = ev.target.closest ? ev.target.closest('.node') : null;
      if (g) { var id = g.getAttribute('data-id'); if (!S.sel[id]) { S.sel = {}; S.sel[id] = true; render(); onSelect(id); } showCtx(ev, id); }
    });
    svg.addEventListener('wheel', function (ev) {
      ev.preventDefault();
      var rect = svg.getBoundingClientRect();
      zoom(ev.deltaY < 0 ? 1.12 : 0.89, ev.clientX - rect.left, ev.clientY - rect.top);
    }, { passive: false });

    // toolbar
    $('#canvas-toolbar').addEventListener('click', function (ev) {
      var b = ev.target.closest('button'); if (!b) return;
      var a = b.getAttribute('data-act');
      if (a === 'fit') fit(); else if (a === 'zoom-in') zoom(1.2); else if (a === 'zoom-out') zoom(0.83);
      else if (a === 'relayout') { mutate(function () { relayout(true); }); toast('已自动重排'); }
      else if (a === 'toggle-logic') toggleLogic();
      else if (a === 'undo') undo(); else if (a === 'redo') redo();
      else if (a === 'export-svg') exportSVG();
    });

    // topbar
    $('#mode-select').addEventListener('change', function () { loadMode(this.value); });
    $('#group-select').addEventListener('change', function () { mutate(function () { S.mode.reg_group = $('#group-select').value; }); renderTabs(); render(); });
    $('#order-select').addEventListener('change', function () { mutate(function () { S.mode.order.mode = $('#order-select').value; }); renderTabs(); });
    $('#btn-record').addEventListener('click', toggleRecord);
    $('#btn-new-mode').addEventListener('click', function () { newModeFromScratch(); promptNewId(); });
    $('#btn-del-mode').addEventListener('click', delMode);
    $('#btn-save').addEventListener('click', saveAll);
    $('#btn-gen').addEventListener('click', function () { S.activeTab = 'seq'; switchTab('seq'); generatePreview(); });
    $('#search').addEventListener('input', function () { doSearch(this.value); });

    // side tabs
    $('#side-tabs').addEventListener('click', function (ev) {
      var b = ev.target.closest('button'); if (!b) return; switchTab(b.getAttribute('data-tab'));
    });
    // sidebar 动作按钮：用 data-act 委托代替 inline onclick，杜绝把 id 拼进 JS 串的注入面。
    $('#side').addEventListener('click', function (ev) {
      var b = ev.target.closest('button[data-act]'); if (!b) return;
      var act = b.getAttribute('data-act'), id = b.getAttribute('data-id'), val = b.getAttribute('data-val');
      if (act === 'toggle-en') toggleEnabled(id);
      else if (act === 'collapse') toggleCollapse(id);
      else if (act === 'setmux') setMux(id, parseInt(val, 10));
      else if (act === 'gen') { switchTab('seq'); generatePreview(); }
      else if (act === 'clear-manual') { mutate(function () { S.mode.order.manual = []; }); renderModeTab(); }
      else if (act === 'copy') seqCopy();
      else if (act === 'dl') seqDownload(val);
      else if (act === 'export') seqExport();
      else if (act === 'match-apply') applyMatching();
      else if (act === 'match-remove') matchSetAlias(id, null);
      else if (act === 'match-logic') matchToggleLogic(id);
      else if (act === 'match-reload') loadMatching(true);
      else if (act === 'wiz-save-config') wizSaveConfig();
      else if (act === 'wiz-pick-netlist') wizPick('netlist', 'netpath');
      else if (act === 'wiz-pick-excel') wizPick('excel', 'xlsxpath');
      else if (act === 'wiz-import-netlist') wizImportNetlist();
      else if (act === 'wiz-import-excel') wizImportExcel();
      else if (act === 'wiz-save-cs') wizSaveCS();
      else if (act === 'wiz-build') wizBuild();
      else if (act === 'wiz-export') wizExport();
    });

    // keyboard
    window.addEventListener('keydown', function (ev) {
      if (/input|select|textarea/i.test(ev.target.tagName)) return;
      if ((ev.ctrlKey || ev.metaKey) && ev.key === 'z') { ev.preventDefault(); ev.shiftKey ? redo() : undo(); }
      else if ((ev.ctrlKey || ev.metaKey) && ev.key === 'y') { ev.preventDefault(); redo(); }
      else if (ev.key === 'f') fit();
      else if (ev.key === 'h') { Object.keys(S.sel).forEach(hideNode); }
      else if (ev.key === 'e' && Object.keys(S.sel)[0]) toggleEnabled(Object.keys(S.sel)[0]);
      else if (ev.key === 'Escape') { hideCtx(); S.sel = {}; render(); }
    });

    window.addEventListener('resize', function () { render(); });
  }

  var boxSelEl = null;
  function drawBoxSel(bs) {
    if (bs.x1 === undefined) return;
    if (!boxSelEl) { boxSelEl = document.createElement('div'); boxSelEl.style.cssText = 'position:absolute;border:1px solid var(--acc);background:rgba(110,168,254,.12);z-index:4;pointer-events:none'; $('#canvas-wrap').appendChild(boxSelEl); }
    var x0 = Math.min(bs.x0, bs.x1), y0 = Math.min(bs.y0, bs.y1);
    boxSelEl.style.left = x0 + 'px'; boxSelEl.style.top = y0 + 'px';
    boxSelEl.style.width = Math.abs(bs.x1 - bs.x0) + 'px'; boxSelEl.style.height = Math.abs(bs.y1 - bs.y0) + 'px';
  }
  function clearBoxSel() { if (boxSelEl) { boxSelEl.remove(); boxSelEl = null; } }

  function saveLayoutQuiet() { Backend.saveLayout(S.layout); }

  // ---------------------------------------------------------------- edit ops
  function toggleEnabled(id) {
    var n = S.nodeById[id]; if (!n) return;
    if (!(n.off_controls || []).length && n.kind !== 'composite' && n.kind !== 'module') { toast(id + ' 无电流门，不作激活门'); }
    mutate(function () {
      var i = S.mode.enabled_nodes.indexOf(id);
      if (i >= 0) S.mode.enabled_nodes.splice(i, 1); else S.mode.enabled_nodes.push(id);
    });
    render(); renderTabs();
    toast((S.mode.enabled_nodes.indexOf(id) >= 0 ? '激活 ' : '关闭 ') + id);
  }
  function toggleCollapse(id) {
    mutate(function () {
      var arr = S.layout.collapsed = S.layout.collapsed || [];
      var i = arr.indexOf(id); if (i >= 0) arr.splice(i, 1); else arr.push(id);
      relayout(false);
    });
    toast((S.layout.collapsed.indexOf(id) >= 0 ? '收起 ' : '展开 ') + id);
  }
  function toggleLogic() {
    mutate(function () {
      var logicIds = (S.fg.nodes || []).filter(function (n) { return n.control_domain; }).map(function (n) { return n.id; });
      var exp = S.layout.expanded = S.layout.expanded || [];
      var anyHidden = logicIds.some(function (id) { return exp.indexOf(id) < 0; });
      if (anyHidden) logicIds.forEach(function (id) { if (exp.indexOf(id) < 0) exp.push(id); });
      else S.layout.expanded = exp.filter(function (id) { return logicIds.indexOf(id) < 0; });
      relayout(false);
    });
  }
  function hideNode(id) { mutate(function () { (S.layout.hidden = S.layout.hidden || []).push(id); delete S.sel[id]; relayout(false); }); }
  function flipEdge(eid) { mutate(function () { S.layout.edge_flip = S.layout.edge_flip || {}; S.layout.edge_flip[eid] = !S.layout.edge_flip[eid]; }); render(); }
  function setMux(id, val) { mutate(function () { S.mode.mux_sel[id] = val; }); renderTabs(); }

  function delMode() {
    if (!S.modeId) { toast('未保存的模式'); return; }
    if (!window.confirm('删除模式 ' + S.modeId + '?')) return;
    Backend.deleteMode(S.modeId).then(function () {
      S.modeList = S.modeList.filter(function (m) { return m.id !== S.modeId; }); initModeSelect();
      if (S.modeList.length) loadMode(S.modeList[0].id); else newModeFromScratch();
      toast('已删除');
    });
  }

  // ---------------------------------------------------------------- record & search
  function toggleRecord() {
    S.recording = !S.recording;
    $('#btn-record').classList.toggle('active', S.recording);
    if (S.recording) { mutate(function () { S.mode.order.mode = 'manual'; S.mode.order.manual = []; }); $('#order-select').value = 'manual'; toast('录制中：依次点击要关闭的模块'); }
    else toast('录制结束');
    renderTabs();
  }
  function recordClick(id) {
    if (S.mode.enabled_nodes.indexOf(id) < 0) { toast('该节点不在激活集，先激活它'); return; }
    if (S.mode.order.manual.indexOf(id) < 0) { mutate(function () { S.mode.order.manual.push(id); }); toast('记录 #' + S.mode.order.manual.length + ' ' + id); renderTabs(); }
  }
  function doSearch(q) {
    q = q.trim().toLowerCase(); if (!q) { S.sel = {}; render(); return; }
    var hit = null;
    Object.keys(S.lay.nodes).forEach(function (id) {
      var n = S.nodeById[id];
      var hay = (id + ' ' + (n.name || '') + ' ' + (n.inst_name || '') + ' ' + (n.controls || []).map(function (c) { return c.signal_ref; }).join(' ')).toLowerCase();
      if (hay.indexOf(q) >= 0) { S.sel[id] = true; if (!hit) hit = id; }
    });
    render();
    if (hit) { var b = S.lay.nodes[hit]; var wrap = $('#canvas-wrap'); S.view.x = wrap.clientWidth / 2 - (b.x + b.w / 2) * S.view.k; S.view.y = wrap.clientHeight / 2 - (b.y + b.h / 2) * S.view.k; applyView(); }
  }

  // ---------------------------------------------------------------- context menu
  function showCtx(ev, id) {
    var n = S.nodeById[id]; var m = $('#ctx-menu');
    var items = [];
    items.push(['⏻ ' + (S.mode.enabled_nodes.indexOf(id) >= 0 ? '取消激活' : '激活（模式开）'), function () { toggleEnabled(id); }]);
    if (n.kind === 'composite') items.push([(S.layout.collapsed.indexOf(id) >= 0 ? '展开黑盒' : '收起黑盒'), function () { toggleCollapse(id); }]);
    items.push(['隐藏节点', function () { hideNode(id); }]);
    items.push(['div', null]);
    // 翻转与该节点相连的边
    var eds = (S.fg.edges || []).filter(function (e) { return e.from.node === id || (e.to || []).some(function (t) { return t.node === id; }); });
    if (eds.length) items.push(['翻转相连边方向 (' + eds.length + ')', function () { eds.forEach(function (e) { flipEdge(e.id); }); }]);
    items.push(['备注/颜色…', function () { var t = window.prompt('备注', (S.layout.notes && S.layout.notes[id]) || ''); if (t !== null) { mutate(function () { S.layout.notes = S.layout.notes || {}; S.layout.notes[id] = t; }); } }]);
    m.innerHTML = ''; items.forEach(function (it) {
      if (it[0] === 'div') { var d = document.createElement('div'); d.className = 'div'; m.appendChild(d); return; }
      var b = document.createElement('button'); b.textContent = it[0]; b.onclick = function () { hideCtx(); it[1](); }; m.appendChild(b);
    });
    var rect = $('#canvas-wrap').getBoundingClientRect();
    m.style.left = (ev.clientX - rect.left) + 'px'; m.style.top = (ev.clientY - rect.top) + 'px'; m.style.display = 'block';
  }
  function hideCtx() { $('#ctx-menu').style.display = 'none'; }

  // ---------------------------------------------------------------- tabs / inspector
  function switchTab(t) {
    S.activeTab = t;
    $$('#side-tabs button').forEach(function (b) { b.classList.toggle('active', b.getAttribute('data-tab') === t); });
    $$('.tab-body').forEach(function (b) { b.classList.remove('active'); });
    $('#tab-' + t).classList.add('active');
    if (t === 'match' && !S.match && !S.matchLoading) loadMatching();
    if (t === 'project') renderProjectTab();
    renderTabs();
  }
  function onSelect(id) { if (S.activeTab === 'inspector') renderInspector(id); }
  // project(工程) tab 不进 renderTabs：它是表单，只在切入/操作后主动渲染，避免打断填写
  function renderTabs() { renderInspector(Object.keys(S.sel)[0]); renderModeTab(); renderMatchTab(); renderSeqTab(); }

  function variantOf(sig) {
    var s = S.sigById[sig]; if (!s) return null;
    var g = S.mode.reg_group;
    return (s.variants && (s.variants[g] || s.variants[s.single_copy ? 'COMMON' : 'COMMON'])) || (s.variants && s.variants.COMMON) || null;
  }

  function renderInspector(id) {
    var el = $('#tab-inspector');
    if (!id || !S.nodeById[id]) { el.innerHTML = '<div class="hint">点击画布中的节点查看其控制信号 → 寄存器映射。<br>提示：Inspector 里可编辑基线字段值；Mode 标签下点节点=激活/取消。</div>'; return; }
    var n = S.nodeById[id];
    var enabled = S.mode.enabled_nodes.indexOf(id) >= 0;
    var h = [];
    h.push('<h3 class="sec">节点</h3><div class="kv">');
    h.push('<div class="k">id</div><div class="v mono">' + esc(id) + '</div>');
    h.push('<div class="k">类型</div><div class="v">' + esc(n.kind) + ' · ' + esc(n.device) + (n.inferred ? ' <span class="pill">推断</span>' : '') + '</div>');
    if (n.inst_type) h.push('<div class="k">inst_type</div><div class="v mono">' + esc(n.inst_type) + '</div>');
    if (n.module) h.push('<div class="k">module</div><div class="v">' + esc((S.fg.module_tags && S.fg.module_tags[n.module]) || n.module) + '</div>');
    h.push('</div>');
    if (n.warn) h.push('<div class="card warn">⚠ ' + esc(n.warn) + '</div>');
    h.push('<div class="row-actions"><button data-act="toggle-en" data-id="' + esc(id) + '">' + (enabled ? '取消激活' : '⏻ 激活（模式开）') + '</button>');
    if (n.kind === 'composite') h.push('<button data-act="collapse" data-id="' + esc(id) + '">' + (S.layout.collapsed.indexOf(id) >= 0 ? '展开' : '收起') + '黑盒</button>');
    h.push('</div>');

    var ctrls = n.controls || [];
    if (ctrls.length) {
      h.push('<h3 class="sec">控制信号 → 寄存器（组 ' + esc(S.mode.reg_group) + '）</h3>');
      h.push('<table class="sig"><tr><th>信号 / pin</th><th>addr</th><th>bit</th><th>基线值</th></tr>');
      ctrls.forEach(function (c) {
        var sig = c.signal_ref; var v = variantOf(sig); var s = S.sigById[sig] || {};
        var isGate = (n.off_controls || []).some(function (o) { return o.signal_ref === sig; });
        var base = baselineValue(sig, id);
        h.push('<tr><td><b class="mono">' + esc(short(sig, 30)) + '</b>' + (c.shared ? ' <span class="pill shared">共用</span>' : '') + (isGate ? ' <span class="pill on">门</span>' : '') + '<br><span class="muted">' + esc(c.pin) + ' · ' + esc(c.role || '') + '</span></td>');
        if (v) {
          h.push('<td><span class="addr">' + esc((v.addr || '').toUpperCase()) + '</span><br><span class="muted">' + esc(v.reg_name || '') + '</span></td>');
          h.push('<td class="mono">' + esc(v.bit) + '<br><span class="muted">rst ' + esc(v.reset || '') + '</span></td>');
          h.push('<td><input data-sig="' + esc(sig) + '" value="' + esc(base) + '" ' + (isGate && enabled ? 'title="激活节点的 en 基线自动=on"' : '') + '></td>');
        } else {
          h.push('<td colspan="3" class="muted">未解析到寄存器（' + esc(s.category || '') + '）</td>');
        }
        h.push('</tr>');
      });
      h.push('</table><div class="hint">改值即写入模式基线（override）；enable 门留空则由激活状态决定 on/off。</div>');
    }
    var notes = S.layout.notes && S.layout.notes[id];
    if (notes) h.push('<div class="card">📝 ' + esc(notes) + '</div>');
    el.innerHTML = h.join('');
    $$('#tab-inspector input[data-sig]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var sig = this.getAttribute('data-sig'), val = this.value.trim();
        mutate(function () {
          if (val === '') delete S.mode.baseline[sig];
          else S.mode.baseline[sig] = parseVal(val);
        });
        toast('基线 ' + sig + ' = ' + val); renderSeqTab();
      });
    });
  }
  function baselineValue(sig, nodeId) {
    if (S.mode.baseline[sig] !== undefined) return S.mode.baseline[sig];
    var v = variantOf(sig); if (!v) return '';
    return ''; // 空=用默认/门自动
  }
  function parseVal(s) { s = String(s).trim(); return s.toLowerCase().indexOf('0x') === 0 ? parseInt(s, 16) : (parseInt(s, 10) || 0); }

  // ---------------------------------------------------------------- mode tab
  function renderModeTab() {
    var el = $('#tab-mode');
    var m = S.mode; if (!m) { el.innerHTML = ''; return; }
    var h = [];
    h.push('<h3 class="sec">模式</h3><div class="kv">');
    h.push('<div class="k">id</div><div class="v mono">' + esc(m.id) + (S.modeId ? '' : ' <span class="pill">未保存</span>') + '</div>');
    h.push('<div class="k">名称</div><div class="v"><input id="m-name" value="' + esc(m.name || '') + '" style="width:96%"></div>');
    h.push('<div class="k">寄存器组</div><div class="v">' + esc(m.reg_group) + '</div>');
    h.push('<div class="k">关闭顺序</div><div class="v">' + esc(m.order.mode) + '</div>');
    h.push('</div>');
    h.push('<div class="hint">在画布点节点=激活/取消（本标签下）。激活集 = 这条被点亮的 LO 通路。</div>');

    h.push('<h3 class="sec">激活节点 (' + m.enabled_nodes.length + ')</h3>');
    if (!m.enabled_nodes.length) h.push('<div class="hint">还没有激活节点，点画布里的 DCO/buffer/div 把通路点亮。</div>');
    m.enabled_nodes.forEach(function (id) {
      var n = S.nodeById[id] || {};
      var gates = (n.off_controls || []).length;
      h.push('<div class="card" style="padding:6px 10px"><b class="mono">' + esc(short(id, 34)) + '</b> <span class="pill">' + esc(n.device || '') + '</span> <span class="muted">' + gates + ' 门</span> <button data-act="toggle-en" data-id="' + esc(id) + '" style="float:right">✕</button></div>');
    });

    // mux sel
    var muxes = (S.fg.nodes || []).filter(function (n) { return n.device === 'mux' && m.enabled_nodes.indexOf(n.id) >= 0; });
    if (muxes.length) {
      h.push('<h3 class="sec">MUX 选择（0=上/1=下，约定可改）</h3>');
      muxes.forEach(function (n) {
        var v = m.mux_sel[n.id] === undefined ? 0 : m.mux_sel[n.id];
        h.push('<div class="kv"><div class="k mono">' + esc(short(n.id, 24)) + '</div><div class="v"><button data-act="setmux" data-id="' + esc(n.id) + '" data-val="0" ' + (v === 0 ? 'class="active"' : '') + '>0</button> <button data-act="setmux" data-id="' + esc(n.id) + '" data-val="1" ' + (v === 1 ? 'class="active"' : '') + '>1</button></div></div>');
      });
    }

    if (m.order.mode === 'manual') {
      h.push('<h3 class="sec">录制的关闭顺序 (' + m.order.manual.length + ')</h3>');
      m.order.manual.forEach(function (id, i) { h.push('<div class="card" style="padding:4px 10px">#' + (i + 1) + ' <span class="mono">' + esc(short(id, 30)) + '</span></div>'); });
      h.push('<div class="row-actions"><button data-act="clear-manual">清空顺序</button></div>');
    }
    h.push('<div class="row-actions"><button class="primary" data-act="gen">生成序列 ▸</button></div>');
    el.innerHTML = h.join('');
    var nm = $('#m-name'); if (nm) nm.addEventListener('change', function () { mutate(function () { S.mode.name = this.value; }.bind(this)); });
  }

  // ---------------------------------------------------------------- matching tab（P2：信号→寄存器半自动匹配）
  function shallowObj(o) { var r = {}; Object.keys(o || {}).forEach(function (k) { r[k] = o[k]; }); return r; }

  function loadMatching(force) {
    if (Backend.bundle) { renderMatchTab(); return; }
    if (S.matchLoading) return;
    if (S.match && !S.match.error && !force) { renderMatchTab(); return; }
    S.matchLoading = true; renderMatchTab();
    Backend.matching().then(function (d) {
      S.matchLoading = false;
      if (!d) { S.match = { error: '无响应' }; renderMatchTab(); return; }
      var srm = d.signal_reg_map || {}, mt = d.matching || {};
      var fields = [], seen = {};
      (srm.registers || []).forEach(function (reg) {
        (reg.fields || []).forEach(function (f) {
          if (f.name && !seen[f.name]) { seen[f.name] = 1; fields.push(f.name); }
        });
      });
      S.match = {
        signals: srm.signals || [], fields: fields, counts: srm.counts || {},
        alias: shallowObj(mt.alias), logic: (mt.logic_derived || []).slice(),
        baseAlias: shallowObj(mt.alias), baseLogic: (mt.logic_derived || []).slice()
      };
      renderMatchTab();
    }).catch(function (e) { S.matchLoading = false; S.match = { error: String(e) }; renderMatchTab(); });
  }
  function matchDirty() {
    if (!S.match || S.match.error) return false;
    return JSON.stringify(S.match.alias) !== JSON.stringify(S.match.baseAlias)
        || JSON.stringify(S.match.logic.slice().sort()) !== JSON.stringify(S.match.baseLogic.slice().sort());
  }
  function matchSetAlias(sig, field) {
    if (!S.match) return;
    if (field) { S.match.alias[sig] = field; var i = S.match.logic.indexOf(sig); if (i >= 0) S.match.logic.splice(i, 1); }
    else delete S.match.alias[sig];
    renderMatchTab();
  }
  function matchToggleLogic(sig) {
    if (!S.match) return;
    var i = S.match.logic.indexOf(sig);
    if (i >= 0) S.match.logic.splice(i, 1);
    else { S.match.logic.push(sig); delete S.match.alias[sig]; }
    renderMatchTab();
  }
  function matchEffStatus(s) {
    var sig = s.reg_net, M = S.match;
    if (s.match === 'exact' || s.match === 'case') return s.match;   // 名字直配，不受 alias 影响
    if (M.alias[sig]) return (M.baseAlias[sig] === M.alias[sig]) ? 'alias' : 'alias*';
    if (M.logic.indexOf(sig) >= 0) return (M.baseLogic.indexOf(sig) >= 0) ? 'logic' : 'logic*';
    return 'unresolved';
  }
  function applyMatching() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    if (!matchDirty()) { toast('无改动'); return; }
    toast('重建中…');
    Backend.saveMatching({ alias: S.match.alias, logic_derived: S.match.logic }).then(function (r) {
      if (!r || !r.ok) { toast('重建失败：' + ((r && r.error) || '?')); console.error(r); return; }
      S.fg = r.flowgraph; S.rm = r.regmap;
      S.nodeById = {}; (S.fg.nodes || []).forEach(function (n) { S.nodeById[n.id] = n; });
      S.sigById = {}; (S.rm.signals || []).forEach(function (s) { S.sigById[s.id] = s; });
      var srm = r.signal_reg_map || {}, mt = r.matching || {};
      S.match.signals = srm.signals || S.match.signals;
      S.match.counts = srm.counts || S.match.counts;
      S.match.alias = shallowObj(mt.alias); S.match.baseAlias = shallowObj(mt.alias);
      S.match.logic = (mt.logic_derived || []).slice(); S.match.baseLogic = (mt.logic_derived || []).slice();
      relayout(false); render(); renderTabs();
      var c = srm.counts || {};
      toast('已重建 · ' + (c.exact || 0) + ' exact/' + (c.alias || 0) + ' alias/' + (c.unresolved || 0) + ' 未匹配');
    }).catch(function (e) { toast('重建失败: ' + e); console.error(e); });
  }
  function renderMatchTab() {
    var el = $('#tab-match'); if (!el) return;
    if (Backend.bundle) { el.innerHTML = '<div class="hint">匹配编辑需 serve 模式（要跑 Python 重建 regmap/flowgraph）。bundle 是只读快照。</div>'; return; }
    if (S.matchLoading) { el.innerHTML = '<div class="hint">加载匹配状态…</div>'; return; }
    if (!S.match) { el.innerHTML = '<div class="hint">加载中…</div>'; return; }
    if (S.match.error) { el.innerHTML = '<div class="card warn">加载失败：' + esc(String(S.match.error)) + '</div>'; return; }
    var M = S.match;
    var todo = [], logic = [], resolved = [];
    M.signals.forEach(function (s) {
      var st = matchEffStatus(s);
      if (st === 'unresolved' || st === 'alias*') todo.push({ s: s, st: st });
      else if (st === 'logic' || st === 'logic*') logic.push({ s: s, st: st });
      else resolved.push({ s: s, st: st });
    });
    var h = [];
    h.push('<datalist id="fieldcat">');
    M.fields.forEach(function (f) { h.push('<option value="' + esc(f) + '">'); });
    h.push('</datalist>');
    h.push('<h3 class="sec">信号 → 寄存器匹配</h3>');
    h.push('<div class="hint">网表控制信号名与寄存器字段名经硅迭代漂移；自动匹配 exact/case，对不上的在此点选真字段（写回 <b>project.matching.alias</b>）或标为逻辑推导。确认后重建 regmap/flowgraph。</div>');
    var dirty = matchDirty();
    h.push('<div class="kv"><div class="k">状态</div><div class="v">' + resolved.length + ' 已匹配 · ' + todo.length + ' 待处理 · ' + logic.length + ' 逻辑推导</div></div>');
    h.push('<div class="row-actions"><button class="primary" data-act="match-apply"' + (dirty ? '' : ' disabled') + '>应用 &amp; 重建 ▸</button><button data-act="match-reload">↻ 重载</button></div>');
    if (dirty) h.push('<div class="card warn" style="font-size:11.5px">有未应用改动，应用会写 project.json 并重跑建库脚本。</div>');

    h.push('<h3 class="sec">需处理 (' + todo.length + ')</h3>');
    if (!todo.length) h.push('<div class="hint">没有待处理信号 🎉</div>');
    todo.forEach(function (x) {
      var s = x.s, sig = s.reg_net, picked = M.alias[sig] || '';
      h.push('<div class="card" style="padding:8px 10px"><b class="mono">' + esc(short(sig, 32)) + '</b>');
      if (s.category) h.push(' <span class="pill">' + esc(s.category) + '</span>');
      if (x.st === 'alias*') h.push(' <span class="pill on">待应用</span>');
      if ((s.drives || []).length) h.push('<div class="muted" style="font-size:11px">→ ' + esc(s.drives.join(', ')) + '</div>');
      h.push('<div style="margin-top:5px;display:flex;gap:4px">');
      h.push('<input list="fieldcat" data-matchsig="' + esc(sig) + '" placeholder="选/输真实字段名…" value="' + esc(picked) + '" style="flex:1">');
      if (picked) h.push('<button data-act="match-remove" data-id="' + esc(sig) + '" title="清除">✕</button>');
      h.push('</div>');
      h.push('<div class="row-actions" style="margin-top:4px"><button data-act="match-logic" data-id="' + esc(sig) + '">标为逻辑推导</button></div>');
      h.push('</div>');
    });

    if (logic.length) {
      h.push('<h3 class="sec">逻辑推导 (' + logic.length + ')</h3>');
      h.push('<div class="hint">这些不是直接寄存器位（组合逻辑/环路产物），不进 regmap 解析。</div>');
      logic.forEach(function (x) {
        var sig = x.s.reg_net;
        h.push('<div class="card" style="padding:6px 10px"><b class="mono">' + esc(short(sig, 30)) + '</b>' + (x.st === 'logic*' ? ' <span class="pill on">待应用</span>' : '') + ' <button data-act="match-logic" data-id="' + esc(sig) + '" style="float:right">取消</button></div>');
      });
    }

    h.push('<h3 class="sec">已匹配 (' + resolved.length + ')</h3>');
    h.push('<table class="sig"><tr><th>信号</th><th>字段 / addr</th><th>bit</th><th>来源</th></tr>');
    resolved.forEach(function (x) {
      var s = x.s;
      h.push('<tr><td><b class="mono">' + esc(short(s.reg_net, 24)) + '</b></td>');
      h.push('<td><span class="mono">' + esc(short(s.field_name || '', 22)) + '</span><br><span class="addr">' + esc((s.addr || '').toUpperCase()) + '</span></td>');
      h.push('<td class="mono">' + esc(s.bit || '') + '</td>');
      h.push('<td>' + esc(x.st) + (x.st === 'alias' ? ' <button data-act="match-remove" data-id="' + esc(s.reg_net) + '" title="取消别名">✕</button>' : '') + '</td></tr>');
    });
    h.push('</table>');
    el.innerHTML = h.join('');
    $$('#tab-match input[data-matchsig]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        matchSetAlias(this.getAttribute('data-matchsig'), this.value.trim() || null);
      });
    });
  }

  // ---------------------------------------------------------------- sequence tab
  function generatePreview() {
    try {
      S.mode.reg_group = $('#group-select').value; S.mode.order.mode = $('#order-select').value;
      S.lastTc = Generator.generate(S.fg, S.rm, S.mode);
      renderSeqTab();
    } catch (e) { toast('生成失败: ' + e); console.error(e); }
  }
  function renderSeqTab() {
    var el = $('#tab-seq');
    if (!S.lastTc) { el.innerHTML = '<div class="hint">点"生成序列"预览逐级关闭测试项。</div>'; return; }
    var tc = S.lastTc; var h = [];
    h.push('<div class="row-actions"><button class="primary" data-act="gen">重新生成</button>');
    h.push('<button data-act="copy">复制 ate.txt</button><button data-act="dl" data-val="ate">下载 ate.txt</button>');
    h.push('<button data-act="dl" data-val="html">下载 debug.html</button>');
    if (!Backend.bundle) h.push('<button data-act="export">写入 project/testcases</button>');
    h.push('</div>');
    h.push('<div class="kv"><div class="k">模式</div><div class="v">' + esc(tc.mode) + ' · 组 ' + esc(tc.reg_group) + ' · ' + esc(tc.order_mode) + '</div>');
    h.push('<div class="k">规模</div><div class="v">' + tc.stats.baseline_regs + ' baseline 寄存器 · ' + tc.stats.steps + ' 步 · ' + tc.stats.gates_off + ' 门关闭</div></div>');
    (tc.warnings || []).forEach(function (w) { h.push('<div class="card warn">⚠ ' + esc(w) + '</div>'); });

    h.push('<div class="step baseline"><h4>Baseline · 全开起始态</h4>');
    tc.baseline.writes.forEach(function (w) { h.push('<div class="w"><span class="addr">' + esc(w.addr) + '</span> <b>' + esc(w.value) + '</b> ' + esc(w.reg) + '</div>'); });
    h.push('</div>');
    tc.steps.forEach(function (st) {
      h.push('<div class="step"><h4>step ' + st.index + ' · OFF <span class="off">' + esc(st.off_label) + '</span> <span class="muted">(' + esc(st.device || '') + ')</span></h4>');
      (st.warnings || []).forEach(function (wn) { h.push('<div class="warn" style="font-size:11.5px">⚠ ' + esc(wn) + '</div>'); });
      if (st.note) h.push('<div class="muted" style="font-size:11.5px">· ' + esc(st.note) + '</div>');
      st.writes.forEach(function (w) {
        var ff = (w.fields || []).map(function (f) { return esc(f.signal) + '[' + esc(f.bit) + ']:' + f.before + '→' + f.after; }).join(', ');
        h.push('<div class="w"><span class="addr">' + esc(w.addr) + '</span> ' + esc(w.prev) + '→<b>' + esc(w.value) + '</b> · ' + ff + '</div>');
      });
      h.push('<div class="muted" style="font-size:11px;margin-top:3px">▸ ' + esc(st.measure) + '</div></div>');
    });
    if (tc.diagnostics.uncovered_off_gates.length) {
      h.push('<div class="card warn"><b>未覆盖门（人工补）</b>');
      tc.diagnostics.uncovered_off_gates.forEach(function (u) { h.push('<div>' + esc(u.signal) + '</div>'); });
      h.push('</div>');
    }
    h.push('<h3 class="sec">ate.txt 预览</h3><textarea readonly>' + esc(Generator.renderAte(tc)) + '</textarea>');
    el.innerHTML = h.join('');
  }
  function seqCopy() { if (!S.lastTc) return; var t = Generator.renderAte(S.lastTc); navigator.clipboard ? navigator.clipboard.writeText(t).then(function () { toast('已复制'); }) : toast('浏览器不支持剪贴板'); }
  function seqDownload(kind) {
    if (!S.lastTc) return;
    var name = (S.lastTc.mode || 'mode');
    if (kind === 'ate') download(name + '.ate.txt', Generator.renderAte(S.lastTc), 'text/plain');
    else download(name + '.debug.html', debugHtml(S.lastTc), 'text/html');
  }
  function seqExport() { Backend.exportSeq(S.modeId || S.mode.id).then(function (r) { toast(r && r.ok ? '已写入 ' + r.dir : '导出失败（先保存模式）'); }); }

  function download(fn, text, mime) {
    var blob = new Blob([text], { type: mime + ';charset=utf-8' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fn; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 500);
  }
  function exportSVG() {
    var svg = $('#graph').cloneNode(true);
    svg.setAttribute('xmlns', SVGNS);
    var s = new XMLSerializer().serializeToString(svg);
    download((S.mode ? S.mode.id : 'graph') + '.svg', s, 'image/svg+xml'); toast('已导出 SVG');
  }

  // minimal client debug.html (bundle 模式；serve 模式 Python 版更全)
  function debugHtml(tc) {
    var rows = tc.steps.map(function (st) {
      var ws = st.writes.map(function (w) { return w.addr + ' ' + w.prev + '→' + w.value; }).join('<br>');
      return '<tr><td>' + st.index + '</td><td>' + esc(st.off_label) + '</td><td class="mono">' + ws + '</td></tr>';
    }).join('');
    return '<!doctype html><meta charset=utf-8><title>debug ' + esc(tc.mode) + '</title>' +
      '<style>body{font-family:monospace;background:#0f1216;color:#e6e9ef;padding:20px}table{border-collapse:collapse}td,th{border:1px solid #2a2f3a;padding:4px 8px}.mono{font-family:monospace}</style>' +
      '<h2>' + esc(tc.mode) + ' · ' + esc(tc.reg_group) + '</h2><pre>' + esc(Generator.renderAte(tc)) + '</pre>' +
      '<table><tr><th>step</th><th>off</th><th>writes</th></tr>' + rows + '</table>';
  }

  // ---------------------------------------------------------------- 工程/建库向导 tab（P2.2）
  // 占位模板：键=make_mock 需要的字段名（固定契约），值=示例列号；用户按自己表的实际列索引改
  var COLMAP_TEMPLATE = { reg_name: 0, regtype: 1, offset: 2, width: 3, reset: 4, field_name: 5, bit: 6, attr: 7, default: 8, comment: 9 };
  var CS_HINT = '{"primary_current_related":[{"reg_net":"","category":"en","drives":["TAG.pin"],"desc":""}],"config_secondary":[]}';

  function initWiz() {
    if (S.wiz) return;
    var p = S.project || {}, nl = p.netlist || {}, rb = p.regbook || {};
    S.wiz = {
      name: p.name || '', root: nl.root_module || '',
      targets: (nl.target_modules || []).join('\n'),
      expand: (nl.expand_submodules || []).join('\n'),
      base: rb.base_address || '', sheet: rb.sheet_name || '',
      colmap: JSON.stringify(rb.column_map || COLMAP_TEMPLATE, null, 1),
      fgrules: p.flowgraph_rules ? JSON.stringify(p.flowgraph_rules, null, 1) : '',
      cs: '', netpath: '', xlsxpath: '', rowdump: '', st: {}, exportText: ''
    };
  }
  var WIZ_FIELDS = { name: 'w-name', root: 'w-root', targets: 'w-targets', expand: 'w-expand',
    base: 'w-base', sheet: 'w-sheet', colmap: 'w-colmap', fgrules: 'w-fgrules',
    cs: 'w-cs', netpath: 'w-netpath', xlsxpath: 'w-xlsxpath', rowdump: 'w-rowdump' };
  function wizReadForm() {
    initWiz();
    Object.keys(WIZ_FIELDS).forEach(function (k) { var e = $('#' + WIZ_FIELDS[k]); if (e) S.wiz[k] = e.value; });
  }
  function splitList(s) { return String(s || '').split(/[\n,]+/).map(function (x) { return x.trim(); }).filter(Boolean); }
  function tryParse(s, what) {
    try { return { ok: true, val: (s && s.trim()) ? JSON.parse(s) : null }; }
    catch (e) { toast(what + ' JSON 解析失败：' + e.message); return { ok: false }; }
  }

  function renderProjectTab() {
    initWiz();
    var w = S.wiz, el = $('#tab-project'); if (!el) return;
    var ro = Backend.bundle;
    var h = [];
    h.push('<div class="hint">air-gap 建库流程：①填配置 → ②导入 netlist → ③导入 Excel → ④确认控制信号 → ⑤建库 → ⑥导出配置发回。'
      + '<br>路径填<b>本机绝对路径</b>；<code>--serve</code> 后端在本机跑抽取脚本。' + (ro ? '<br><b>bundle 只读：建库需 <code>--serve</code></b>' : '') + '</div>');

    h.push('<h3 class="sec">① 芯片配置</h3><div class="kv">');
    h.push('<div class="k">name</div><div class="v"><input id="w-name" value="' + esc(w.name) + '"></div>');
    h.push('<div class="k">root_module</div><div class="v"><input id="w-root" value="' + esc(w.root) + '"></div>');
    h.push('<div class="k">base_address</div><div class="v"><input id="w-base" value="' + esc(w.base) + '" placeholder="0x..."></div>');
    h.push('<div class="k">sheet_name</div><div class="v"><input id="w-sheet" value="' + esc(w.sheet) + '" placeholder="寄存器 sheet 名"></div>');
    h.push('</div>');
    h.push('<label class="wlab">目标模块 target_modules（每行一个）</label><textarea id="w-targets" rows="3">' + esc(w.targets) + '</textarea>');
    h.push('<label class="wlab">展开子模块 expand_submodules（每行一个，可空）</label><textarea id="w-expand" rows="2">' + esc(w.expand) + '</textarea>');
    h.push('<label class="wlab">column_map（JSON：字段名→Excel 列号；键固定，值填你表的列索引）</label><textarea id="w-colmap" rows="5" class="mono">' + esc(w.colmap) + '</textarea>');
    h.push('<details><summary class="wsum">高级：flowgraph_rules（JSON，留空用通用默认；known_cross_edges/module_bands 等芯片专属规则在此）</summary>'
      + '<textarea id="w-fgrules" rows="6" class="mono" placeholder="留空=代码通用默认">' + esc(w.fgrules) + '</textarea></details>');
    h.push('<div class="row-actions"><button class="primary" data-act="wiz-save-config"' + (ro ? ' disabled' : '') + '>保存配置</button></div>');

    h.push('<h3 class="sec">② 导入 netlist</h3>');
    h.push('<label class="wlab">netlist 文件路径</label><div class="wrow"><input id="w-netpath" value="' + esc(w.netpath) + '" placeholder="点「浏览…」选，或手动填绝对路径">'
      + '<button data-act="wiz-pick-netlist"' + (ro ? ' disabled' : '') + ' title="打开系统文件对话框">浏览…</button></div>');
    h.push('<div class="row-actions"><button data-act="wiz-import-netlist"' + (ro ? ' disabled' : '') + '>抽取连接 + 生成控制信号候选</button></div>');
    if (w.st.netlist) h.push('<div class="card ' + (/失败/.test(w.st.netlist) ? 'warn' : '') + '">' + esc(w.st.netlist) + '</div>');

    h.push('<h3 class="sec">③ 导入 Excel 寄存器簿</h3>');
    h.push('<label class="wlab">.xlsm 路径</label><div class="wrow"><input id="w-xlsxpath" value="' + esc(w.xlsxpath) + '" placeholder="点「浏览…」选，或手动填绝对路径">'
      + '<button data-act="wiz-pick-excel"' + (ro ? ' disabled' : '') + ' title="打开系统文件对话框">浏览…</button></div>');
    h.push('<div class="kv"><div class="k">sheet</div><div class="v"><input id="w-xsheet" value="' + esc(w.sheet) + '" placeholder="默认取配置 sheet_name"></div>');
    h.push('<div class="k">行区间</div><div class="v"><input id="w-rowdump" value="' + esc(w.rowdump) + '" placeholder="START:END（先 --index/--schema 看结构）"></div></div>');
    h.push('<div class="row-actions"><button data-act="wiz-import-excel"' + (ro ? ' disabled' : '') + '>抽取寄存器行</button></div>');
    if (w.st.excel) h.push('<div class="card ' + (/失败/.test(w.st.excel) ? 'warn' : '') + '">' + esc(w.st.excel) + '</div>');

    h.push('<h3 class="sec">④ 控制信号（reg_net 清单）</h3>');
    h.push('<div class="hint">netlist 导入自动填候选（top_pin 输入控制脚）。<b>category/desc 需人工确认</b>；经内部逻辑门控的漏网脚需手工补。</div>');
    h.push('<textarea id="w-cs" rows="8" class="mono" placeholder=\'' + esc(CS_HINT) + '\'>' + esc(w.cs) + '</textarea>');
    h.push('<div class="row-actions"><button data-act="wiz-save-cs"' + (ro ? ' disabled' : '') + '>保存控制信号</button></div>');

    h.push('<h3 class="sec">⑤ 建库</h3>');
    h.push('<div class="hint">跑 make_mock_regmap → build_regmap → build_flowgraph（读工程包），生成派生物并载入图。</div>');
    h.push('<div class="row-actions"><button class="primary" data-act="wiz-build"' + (ro ? ' disabled' : '') + '>重建派生物 &amp; 载入图 ▸</button></div>');
    if (w.st.build) h.push('<div class="card ' + (/失败/.test(w.st.build) ? 'warn' : '') + '">' + esc(w.st.build) + '</div>');

    h.push('<h3 class="sec">⑥ 导出工程配置（发回）</h3>');
    h.push('<div class="hint">project.json + control_signals + 全部模式，打成一段文本复制发回本地归档。</div>');
    h.push('<div class="row-actions"><button data-act="wiz-export">复制工程配置文本</button></div>');
    if (w.exportText) h.push('<textarea readonly rows="8" class="mono">' + esc(w.exportText) + '</textarea>');

    el.innerHTML = h.join('');
  }

  function wizSaveConfig() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    wizReadForm();
    var cm = tryParse(S.wiz.colmap, 'column_map'); if (!cm.ok) return;
    var fr = tryParse(S.wiz.fgrules, 'flowgraph_rules'); if (!fr.ok) return;
    var p = S.project || {};
    var cfg = { name: S.wiz.name || 'chip' };
    cfg.netlist = Object.assign({}, p.netlist || {}, {
      root_module: S.wiz.root, target_modules: splitList(S.wiz.targets), expand_submodules: splitList(S.wiz.expand) });
    cfg.regbook = Object.assign({}, p.regbook || {}, {
      base_address: S.wiz.base, sheet_name: S.wiz.sheet, column_map: cm.val || COLMAP_TEMPLATE });
    if (fr.val) cfg.flowgraph_rules = fr.val;
    toast('保存配置…');
    Backend.saveConfig(cfg).then(function (r) {
      if (!r || !r.ok) { toast('保存失败'); return; }
      S.project = r.project; toast('配置已保存');
    }).catch(function (e) { toast('保存失败: ' + e); });
  }
  function wizPick(kind, field) {
    if (Backend.bundle) { toast('bundle 只读，手动填路径'); return; }
    wizReadForm();                       // 先存住其它已填字段，避免选完刷新丢失
    toast('打开文件对话框…（在运行 --serve 的本机弹出）');
    Backend.pickFile(kind).then(function (r) {
      if (r && r.path) { S.wiz[field] = r.path; renderProjectTab(); toast('已选：' + r.path); }
      else toast('未选（取消或对话框不可用，可手动填路径）');
    }).catch(function (e) { toast('对话框失败，手动填路径：' + e); });
  }
  function wizImportNetlist() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    wizReadForm();
    if (!S.wiz.netpath) { toast('先填 netlist 路径'); return; }
    toast('抽取 netlist…（跑 extract_ports，可能需几秒）');
    Backend.importNetlist(S.wiz.netpath).then(function (r) {
      if (!r || !r.ok) { S.wiz.st.netlist = '失败：' + ((r && r.error) || '?'); renderProjectTab(); return; }
      var c = r.candidate || {}, np = (c.primary_current_related || []).length, ns = (c.config_secondary || []).length;
      S.wiz.cs = JSON.stringify(c, null, 1);
      S.wiz.st.netlist = '成功：conn ' + r.conn_modules + ' 模块；控制信号候选 ' + (np + ns) + '（primary ' + np + ' / secondary ' + ns + '）'
        + (r.has_control_signals ? '｜已有 control_signals.json 未覆盖，候选见④，需要则替换后保存' : '｜候选已填入④，确认后保存');
      renderProjectTab(); toast('netlist 抽取完成');
    }).catch(function (e) { S.wiz.st.netlist = '失败：' + e; renderProjectTab(); });
  }
  function wizImportExcel() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    wizReadForm();
    var xsheet = ($('#w-xsheet') || {}).value;
    if (!S.wiz.xlsxpath) { toast('先填 .xlsm 路径'); return; }
    toast('抽取 Excel…（跑 explore_excel）');
    Backend.importExcel(S.wiz.xlsxpath, xsheet, S.wiz.rowdump).then(function (r) {
      if (!r || !r.ok) { S.wiz.st.excel = '失败：' + ((r && r.error) || '?'); renderProjectTab(); return; }
      S.wiz.st.excel = '成功：sheet ' + esc(r.sheet) + '，抽 ' + r.rows + ' 行 × ' + r.cols + ' 列 → pll_rows.json';
      renderProjectTab(); toast('Excel 抽取完成');
    }).catch(function (e) { S.wiz.st.excel = '失败：' + e; renderProjectTab(); });
  }
  function wizSaveCS() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    wizReadForm();
    var cs = tryParse(S.wiz.cs, 'control_signals'); if (!cs.ok) return;
    if (!cs.val) { toast('控制信号为空'); return; }
    Backend.saveControlSignals(cs.val).then(function (r) {
      toast(r && r.ok ? '控制信号已保存' : '保存失败');
    }).catch(function (e) { toast('保存失败: ' + e); });
  }
  function wizBuild() {
    if (Backend.bundle) { toast('bundle 只读'); return; }
    toast('建库中…（跑三件套，可能需几秒）');
    Backend.build().then(function (r) {
      if (!r || !r.ok) { S.wiz.st.build = '失败：' + ((r && r.error) || '?'); renderProjectTab(); console.error(r); toast('建库失败'); return; }
      S.fg = r.flowgraph; S.rm = r.regmap; S.needsSetup = !!r.needs_setup;
      S.nodeById = {}; (S.fg.nodes || []).forEach(function (n) { S.nodeById[n.id] = n; });
      S.sigById = {}; (S.rm.signals || []).forEach(function (s) { S.sigById[s.id] = s; });
      S.match = null;   // 强制匹配 tab 下次重载
      relayout(true); render(); fit(); renderTabs();
      var srm = r.signal_reg_map || {}, c = srm.counts || {};
      S.wiz.st.build = '成功：' + (S.fg.nodes || []).length + ' 节点 / ' + (S.rm.signals || []).length + ' 信号'
        + (c.unresolved !== undefined ? '（' + (c.exact || 0) + ' exact / ' + (c.alias || 0) + ' alias / ' + c.unresolved + ' 未匹配）' : '');
      renderProjectTab(); toast('已建库 · ' + (S.fg.nodes || []).length + ' 节点');
    }).catch(function (e) { S.wiz.st.build = '失败：' + e; renderProjectTab(); toast('建库失败: ' + e); });
  }
  function wizExport() {
    Backend.projectExport().then(function (d) {
      var text = JSON.stringify(d, null, 1);
      S.wiz.exportText = text; renderProjectTab();
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(
        function () { toast('已复制（' + text.length + ' 字符）'); },
        function () { toast('已生成，见下方文本框（剪贴板被拒）'); });
      else toast('已生成，手动复制下方文本框');
    }).catch(function (e) { toast('导出失败: ' + e); });
  }

  function esc(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]; }); }
})();
