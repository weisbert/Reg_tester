/* layout.js — 分层信号流布局（自研，无重依赖；替代 dagre，符合红区轻依赖纪律）。
 * 三个 module 各成一条水平 band，band 内按"距 DCO 源头深度"分层左→右，
 * expanded composite 画成子分组框。纯函数，ES6。
 * Layout.compute(flowgraph, layoutState) -> { nodes, groups, hidden }
 *   nodes: { id: {x,y,w,h} }（叶子符号）
 *   groups:[ {id,x,y,w,h,label,kind} ]（module / composite 分组框）
 *   hidden: { id:true }（被折叠/隐藏、不画的节点）
 */
(function (root) {
  'use strict';

  var NODE_W = 132, NODE_H = 56;
  var LAYER_DX = 188, ROW_DY = 82;
  var BAND_GAP = 70, TITLE_H = 30, PAD = 26, MOD_LEFT = 40, TOP = 40;
  var STAGE = { dco: 0, logic: 0, div: 1, mux: 2, buf: 2, inv: 2, blackbox: 2, route: 3, group: 0 };

  function compute(fg, ls) {
    ls = ls || {};
    var positions = ls.positions || {};
    var collapsed = toSet(ls.collapsed);      // composite id -> 收起（当叶子画）
    var expandedLogic = toSet(ls.expanded);   // 额外展开的 control_domain 节点
    var hiddenUser = toSet(ls.hidden);
    var shownUser = toSet(ls.shown);

    var nodeById = {};
    (fg.nodes || []).forEach(function (n) { nodeById[n.id] = n; });

    var modules = (fg.nodes || []).filter(function (n) { return n.kind === 'module'; });

    var hidden = {};
    function isHidden(n) {
      if (shownUser[n.id]) return false;
      if (hiddenUser[n.id]) return true;
      if (n.control_domain && !expandedLogic[n.id]) return true; // logic 默认隐藏
      return false;
    }

    // 判断哪些是叶子（画符号），哪些是分组框
    // - module：分组框
    // - composite：collapsed -> 叶子；否则 -> 分组框（其 inferred 子为叶子）
    // - 其它 primitive/inferred：叶子
    function childLeaves(mod) {
      var leaves = [];
      (fg.nodes || []).forEach(function (n) {
        if (n.module !== mod.id) return;
        if (n.kind === 'module') return;
        if (isHidden(n)) { hidden[n.id] = true; return; }
        if (n.kind === 'composite') {
          if (collapsed[n.id]) { leaves.push(n); }      // 收起当叶子
          // 展开时 composite 自身不是叶子（画成子框），子节点在下面加入
        } else if (n.parent && nodeById[n.parent] && nodeById[n.parent].kind === 'composite'
                   && collapsed[n.parent]) {
          hidden[n.id] = true;                          // 父 composite 收起 -> 子隐藏
        } else {
          leaves.push(n);
        }
      });
      return leaves;
    }

    // 层号：距源头深度
    function assignLayers(leaves, preds) {
      var idset = {}; leaves.forEach(function (n) { idset[n.id] = true; });
      var memo = {};
      function rank(id, stack) {
        if (memo[id] !== undefined) return memo[id];
        stack = stack || {};
        if (stack[id]) return STAGE[(nodeById[id] || {}).device] || 1;
        stack[id] = true;
        var n = nodeById[id] || {}, base = STAGE[n.device];
        if (base === undefined) base = 1;
        var best = base, ps = preds[id] ? Object.keys(preds[id]) : [];
        for (var i = 0; i < ps.length; i++) if (idset[ps[i]]) best = Math.max(best, 1 + rank(ps[i], stack));
        delete stack[id];
        memo[id] = best; return best;
      }
      var out = {};
      leaves.forEach(function (n) { out[n.id] = rank(n.id); });
      return out;
    }

    // 折叠边到可见叶子层，构造 preds（叶子 id -> {前驱叶子 id}）
    function leafPreds(leafSet) {
      function anc(id) {
        var cur = id, seen = {};
        while (cur && !seen[cur]) {
          seen[cur] = true;
          if (leafSet[cur]) return cur;
          var n = nodeById[cur]; cur = n ? n.parent : null;
        }
        return leafSet[id] ? id : null;
      }
      var preds = {};
      (fg.edges || []).forEach(function (e) {
        var f = anc(e.from.node);
        (e.to || []).forEach(function (t) {
          var to = anc(t.node);
          if (!f || !to || f === to) return;
          (preds[to] || (preds[to] = {}))[f] = true;
        });
      });
      return preds;
    }

    var nodes = {}, groups = [];
    var bandTop = TOP;

    modules.forEach(function (mod) {
      var leaves = childLeaves(mod);
      var leafSet = {}; leaves.forEach(function (n) { leafSet[n.id] = true; });
      var preds = leafPreds(leafSet);
      var layers = assignLayers(leaves, preds);

      // 分层分桶
      var byLayer = {};
      leaves.forEach(function (n) {
        var L = layers[n.id]; (byLayer[L] || (byLayer[L] = [])).push(n);
      });
      var layerKeys = Object.keys(byLayer).map(Number).sort(function (a, b) { return a - b; });

      // 桶内排序：先 inferred 分组连续（按父 composite），再按 id
      layerKeys.forEach(function (L) {
        byLayer[L].sort(function (a, b) {
          var pa = a.parent || '', pb = b.parent || '';
          if (pa !== pb) return pa < pb ? -1 : 1;
          return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0);
        });
      });

      var maxRows = 0;
      layerKeys.forEach(function (L) { maxRows = Math.max(maxRows, byLayer[L].length); });
      var innerTop = bandTop + TITLE_H + PAD;

      layerKeys.forEach(function (L, li) {
        var col = byLayer[L];
        // 垂直居中每列
        var colTop = innerTop + (maxRows - col.length) * ROW_DY / 2;
        col.forEach(function (n, ri) {
          var x = MOD_LEFT + PAD + li * LAYER_DX;
          var y = colTop + ri * ROW_DY;
          var p = positions[n.id];
          nodes[n.id] = {
            x: p ? p.x : x, y: p ? p.y : y, w: NODE_W, h: NODE_H,
            _auto: { x: x, y: y }
          };
        });
      });

      // module 分组框 = 叶子包围盒
      var box = bboxOf(leaves.map(function (n) { return nodes[n.id]; }));
      if (!box) box = { x: MOD_LEFT, y: bandTop + TITLE_H, w: 260, h: 80 };
      var gx = MOD_LEFT, gy = bandTop;
      var gw = Math.max((box.x + box.w) - MOD_LEFT + PAD, 300);
      var gh = (box.y + box.h) - bandTop + PAD;
      groups.push({
        id: mod.id, kind: 'module', label: (fg.module_tags && fg.module_tags[mod.id]) || mod.name || mod.id,
        band: mod.band, x: gx, y: gy, w: gw, h: gh
      });

      // expanded composite 子框
      (fg.nodes || []).forEach(function (c) {
        if (c.module !== mod.id || c.kind !== 'composite' || collapsed[c.id]) return;
        var kids = leaves.filter(function (n) { return n.parent === c.id; });
        if (!kids.length) return;
        var cb = bboxOf(kids.map(function (n) { return nodes[n.id]; }));
        if (!cb) return;
        groups.push({
          id: c.id, kind: 'composite', label: c.inst_name || c.name || c.id,
          x: cb.x - 14, y: cb.y - 22, w: cb.w + 28, h: cb.h + 34
        });
      });

      bandTop = bandTop + gh + BAND_GAP;
    });

    return { nodes: nodes, groups: groups, hidden: hidden,
             size: bboxAll(nodes, groups) };
  }

  function toSet(arr) { var s = {}; (arr || []).forEach(function (x) { s[x] = true; }); return s; }
  function bboxOf(boxes) {
    var xs = [], ys = [], xe = [], ye = [];
    boxes.forEach(function (b) { if (b) { xs.push(b.x); ys.push(b.y); xe.push(b.x + b.w); ye.push(b.y + b.h); } });
    if (!xs.length) return null;
    var x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
    return { x: x, y: y, w: Math.max.apply(null, xe) - x, h: Math.max.apply(null, ye) - y };
  }
  function bboxAll(nodes, groups) {
    var all = [];
    Object.keys(nodes).forEach(function (k) { all.push(nodes[k]); });
    groups.forEach(function (g) { all.push(g); });
    var b = bboxOf(all) || { x: 0, y: 0, w: 800, h: 600 };
    return { w: b.x + b.w + 60, h: b.y + b.h + 60 };
  }

  var api = { compute: compute, NODE_W: NODE_W, NODE_H: NODE_H };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Layout = api;
})(typeof window !== 'undefined' ? window : this);
