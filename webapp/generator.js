/* generator.js — 序列生成器（gen_testcase.py 的忠实移植，bundle/离线模式 + GUI 预览用）。
 * 算法权威定义见 SCHEMAS.md 第 5/6 节。与 Python 版产出的 testcase/1 结构逐字段一致，
 * ate.txt 逐字节一致（M4 用 node 交叉验证）。纯函数，无 DOM 依赖，可在 node 下跑。
 * ES6，不用 ?./?? 等新语法，兼容保守浏览器。
 */
(function (root) {
  'use strict';
  var SCHEMA = 'testcase/1';
  var DEVICE_STAGE = { dco: 0, logic: 0, div: 1, mux: 2, buf: 2, inv: 2, blackbox: 2, route: 3, group: 0 };

  function parseBit(bit) {
    bit = String(bit).trim();
    if (bit.indexOf(':') >= 0) {
      var a = bit.split(':'); return [parseInt(a[0], 10), parseInt(a[1], 10)];
    }
    var n = parseInt(bit, 10); return [n, n];
  }
  function fieldMask(bit) {
    var hl = parseBit(bit), hi = hl[0], lo = hl[1], width = hi - lo + 1;
    return [(((1 << width) - 1) << lo) >>> 0, lo, width];
  }
  function setField(word, bit, val) {
    var m = fieldMask(bit), mask = m[0], lo = m[1], width = m[2];
    val = (val & ((1 << width) - 1)) >>> 0;
    return ((word & ~mask) | (val << lo)) >>> 0;
  }
  function getField(word, bit) {
    var m = fieldMask(bit); return (word & m[0]) >>> m[1];
  }
  function hex4(v) {
    var s = (v & 0xFFFF).toString(16).toUpperCase();
    while (s.length < 4) s = '0' + s;
    return '0x' + s;
  }
  function hexAddr(a) {
    var s = String(a).toLowerCase().replace(/0x/g, '');
    var n = parseInt(s, 16);
    var h = (n >>> 0).toString(16).toUpperCase();
    while (h.length < 8) h = '0' + h;
    return '0x' + h;
  }
  function toInt(v, dflt) {
    if (v === null || v === undefined) return dflt || 0;
    if (typeof v === 'number') return v | 0;
    var s = String(v).trim();
    var n = s.toLowerCase().indexOf('0x') === 0 ? parseInt(s, 16) : parseInt(s, 10);
    return isNaN(n) ? (dflt || 0) : n;
  }

  function RegView(regmap, group) {
    this.common = regmap.common_group || 'COMMON';
    this.group = group;
    this.byId = {};
    var sigs = regmap.signals || [];
    for (var i = 0; i < sigs.length; i++) this.byId[sigs[i].id] = sigs[i];
  }
  RegView.prototype.variant = function (sigId) {
    var s = this.byId[sigId]; if (!s) return null;
    var v = s.variants || {};
    return v[this.group] || v[this.common] || null;
  };

  function collectGates(fg, gateOverride) {
    gateOverride = gateOverride || {};
    var gateNodes = {}, nodeGates = {};
    var nodes = fg.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i], gs = [];
      var ov = gateOverride[n.id];   // 指定了关断总闸则只用这些信号
      var ocs = n.off_controls || [];
      for (var j = 0; j < ocs.length; j++) {
        var oc = ocs[j], sig = oc.signal_ref;
        if (!sig) continue;
        if (ov && ov.indexOf(sig) < 0) continue;
        var g = {
          node: n.id, signal: sig, pin: oc.pin,
          off_value: oc.off_value === undefined ? 0 : oc.off_value,
          active_high: oc.active_high === undefined ? null : oc.active_high,
          polarity_inferred: !!oc.polarity_inferred, lane: oc.lane === undefined ? null : oc.lane
        };
        gs.push(g);
        (gateNodes[sig] || (gateNodes[sig] = {}))[n.id] = true;
      }
      if (gs.length) nodeGates[n.id] = gs;
    }
    return { gateNodes: gateNodes, nodeGates: nodeGates };
  }
  function gateNodeSet(gateNodes, sig) { return gateNodes[sig] ? Object.keys(gateNodes[sig]) : []; }
  function onValue(gate) { return gate.active_high === null ? 1 : (gate.active_high ? 1 : 0); }

  function buildEdgeMaps(fg, visible) {
    var nodeById = {};
    var nodes = fg.nodes || [];
    for (var i = 0; i < nodes.length; i++) nodeById[nodes[i].id] = nodes[i];
    function visAnc(nid) {
      var cur = nid, seen = {};
      while (cur !== null && cur !== undefined && !seen[cur]) {
        seen[cur] = true;
        if (visible[cur]) return cur;
        var n = nodeById[cur]; cur = n ? n.parent : null;
      }
      return visible[nid] ? nid : null;
    }
    var preds = {};
    var edges = fg.edges || [];
    for (var e = 0; e < edges.length; e++) {
      var frm = visAnc(edges[e].from.node);
      var tos = edges[e].to || [];
      for (var t = 0; t < tos.length; t++) {
        var to = visAnc(tos[t].node);
        if (!frm || !to || frm === to) continue;
        (preds[to] || (preds[to] = {}))[frm] = true;
      }
    }
    return { preds: preds, nodeById: nodeById };
  }
  function shutdownRank(nid, nodeById, preds, memo, stack) {
    if (memo[nid] !== undefined) return memo[nid];
    if (!stack) stack = {};
    if (stack[nid]) { var nn = nodeById[nid] || {}; return DEVICE_STAGE[nn.device] || 1; }
    stack[nid] = true;
    var n = nodeById[nid] || {}, base = DEVICE_STAGE[n.device];
    if (base === undefined) base = 1;
    var best = base, ps = preds[nid] ? Object.keys(preds[nid]) : [];
    for (var i = 0; i < ps.length; i++) best = Math.max(best, 1 + shutdownRank(ps[i], nodeById, preds, memo, stack));
    delete stack[nid];
    memo[nid] = best; return best;
  }

  // 把 present-but-undefined 归一成 null，使 JS 的 testcase JSON 与 Python(json.dump 写 null) 逐字段一致。
  function nullify(o) {
    if (Array.isArray(o)) { for (var i = 0; i < o.length; i++) { if (o[i] === undefined) o[i] = null; else nullify(o[i]); } }
    else if (o && typeof o === 'object') { for (var k in o) { if (o.hasOwnProperty(k)) { if (o[k] === undefined) o[k] = null; else nullify(o[k]); } } }
    return o;
  }

  function generate(fg, regmap, mode, gateOverride) {
    var group = mode.reg_group || regmap.primary_group || 'BT';
    var rv = new RegView(regmap, group);
    var enabled = {}; (mode.enabled_nodes || []).forEach(function (n) { enabled[n] = true; });
    var baselineOver = mode.baseline || {};
    var cg = collectGates(fg, gateOverride), gateNodes = cg.gateNodes, nodeGates = cg.nodeGates;

    function signalOn(sig) {
      var ns = gateNodeSet(gateNodes, sig);
      for (var i = 0; i < ns.length; i++) if (enabled[ns[i]]) return true;
      return false;
    }
    var warnings = [];

    // touched sigs
    var touched = {};
    Object.keys(gateNodes).forEach(function (s) { touched[s] = true; });
    Object.keys(baselineOver).forEach(function (s) { touched[s] = true; });
    (fg.nodes || []).forEach(function (n) {
      if (enabled[n.id]) (n.controls || []).forEach(function (c) { if (c.signal_ref) touched[c.signal_ref] = true; });
    });

    var images = {}, fieldsSet = {};
    function ensureImg(sig) {
      var v = rv.variant(sig);
      if (!v || !v.addr) return null;
      var addr = hexAddr(v.addr);
      if (!images[addr]) {
        images[addr] = { value: toInt(v.reset, 0), reg: v.reg_name, reset: toInt(v.reset, 0) };
        fieldsSet[addr] = [];
      }
      return [addr, v];
    }
    var unresolved = [];
    Object.keys(touched).sort().forEach(function (sig) { ensureImg(sig); });

    Object.keys(gateNodes).sort().forEach(function (sig) {
      var anyGate = null, ns = gateNodeSet(gateNodes, sig);
      for (var i = 0; i < ns.length && !anyGate; i++) {
        var gs = nodeGates[ns[i]] || [];
        for (var j = 0; j < gs.length; j++) if (gs[j].signal === sig) { anyGate = gs[j]; break; }
      }
      var r = ensureImg(sig);
      if (!r) { unresolved.push(sig); return; }
      var addr = r[0], v = r[1];
      var val = signalOn(sig) ? onValue(anyGate) : anyGate.off_value;
      images[addr].value = setField(images[addr].value, v.bit, val);
      fieldsSet[addr].push({ signal: sig, bit: v.bit, value: val, role: 'enable', on: !!signalOn(sig) });
    });

    // 3.5) mux_sel：GUI 的 MUX 选择写进 sel 字段（约定 0=上/1=下）。放在显式 baseline 之前。
    var muxSel = mode.mux_sel || {};
    var nodeByIdAll = {}; (fg.nodes || []).forEach(function (n) { nodeByIdAll[n.id] = n; });
    Object.keys(muxSel).sort().forEach(function (nid) {
      var node = nodeByIdAll[nid]; if (!node) return;
      var selCtrls = (node.controls || []).filter(function (c) {
        return c.role === 'sel' || (c.pin && c.pin.toLowerCase().indexOf('sel') >= 0);
      });
      if (selCtrls.length !== 1) {
        if (selCtrls.length) warnings.push('MUX ' + nid + ' 有多个 sel 控制，mux_sel 未落值（请在 baseline 里明确）');
        return;
      }
      var sig = selCtrls[0].signal_ref, r = ensureImg(sig);
      if (!r) return;
      var addr = r[0], v = r[1], val = toInt(muxSel[nid], 0);
      images[addr].value = setField(images[addr].value, v.bit, val);
      fieldsSet[addr].push({ signal: sig, bit: v.bit, value: val, role: 'mux_sel' });
    });

    Object.keys(baselineOver).forEach(function (sig) {
      var r = ensureImg(sig);
      if (!r) { unresolved.push(sig); return; }
      var addr = r[0], v = r[1], val = toInt(baselineOver[sig], 0);
      images[addr].value = setField(images[addr].value, v.bit, val);
      fieldsSet[addr].push({ signal: sig, bit: v.bit, value: val, role: 'override' });
    });

    if (unresolved.length) {
      var uniq = {}; unresolved.forEach(function (s) { uniq[s] = true; });
      warnings.push('基线里有未解析到寄存器的信号（跳过）：' + Object.keys(uniq).sort().join(', '));
    }

    var baselineWrites = [];
    Object.keys(images).sort().forEach(function (addr) {
      var im = images[addr];
      baselineWrites.push({ addr: addr, reg: im.reg, value: hex4(im.value), reset: hex4(im.reset), fields: fieldsSet[addr] });
    });

    // active nodes
    var activeNodes = [];
    Object.keys(enabled).forEach(function (nid) {
      var gs = (nodeGates[nid] || []).filter(function (g) { return signalOn(g.signal); });
      if (gs.length) activeNodes.push(nid);
    });

    var visible = {}; (fg.nodes || []).forEach(function (n) { visible[n.id] = true; });
    var em = buildEdgeMaps(fg, visible), preds = em.preds, nodeById = em.nodeById;
    var orderMode = (mode.order && mode.order.mode) || 'auto';
    var ordered, memo = {};
    // 次级键：enabled_nodes 位置（源→末端）→ 越靠后越先关。解决同 rank 同器件类相邻级被 id 字典序关反的问题。
    var enabledList = mode.enabled_nodes || [], eidx = {};
    enabledList.forEach(function (n, i) { eidx[n] = i; });
    function cmpKey(a, b) {
      var ra = shutdownRank(a, nodeById, preds, memo), rb = shutdownRank(b, nodeById, preds, memo);
      if (ra !== rb) return rb - ra;
      var ea = (eidx[a] === undefined ? -1 : eidx[a]), eb = (eidx[b] === undefined ? -1 : eidx[b]);
      if (ea !== eb) return eb - ea;
      return a < b ? -1 : (a > b ? 1 : 0);
    }
    if (orderMode === 'manual') {
      var manual = (mode.order && mode.order.manual) || [];
      var inMan = {};
      ordered = manual.filter(function (n) { return activeNodes.indexOf(n) >= 0; });
      ordered.forEach(function (n) { inMan[n] = true; });
      var rest = activeNodes.filter(function (n) { return !inMan[n]; });
      if (rest.length) warnings.push('manual 顺序未覆盖的激活节点按 auto 追加：' + rest.slice().sort().join(', '));
      rest.sort(cmpKey);
      ordered = ordered.concat(rest);
    } else {
      ordered = activeNodes.slice().sort(cmpKey);
    }
    // 顺序不确定性提示（与 Python 一致）
    var ambiguous = [];
    for (var oi = 0; oi < ordered.length - 1; oi++) {
      if (shutdownRank(ordered[oi], nodeById, preds, memo) === shutdownRank(ordered[oi + 1], nodeById, preds, memo))
        ambiguous.push([ordered[oi], ordered[oi + 1]]);
    }
    if (ambiguous.length) {
      warnings.push('以下相邻级的先后由拓扑无法判定，按 enabled_nodes 录入序（源→末端）兜底，请人工确认：'
        + ambiguous.map(function (p) { return p[0] + '→' + p[1]; }).join('; '));
    }

    var regImage = {}; Object.keys(images).forEach(function (a) { regImage[a] = images[a].value; });
    var signalsOff = {}, steps = [], sharedCollateral = [];
    ordered.forEach(function (nid, i) {
      var idx = i + 1, node = nodeById[nid] || {};
      var gs = (nodeGates[nid] || []).filter(function (g) { return signalOn(g.signal); });
      var writesByAddr = {}, stepGates = [], stepWarn = [];
      gs.forEach(function (g) {
        var sig = g.signal, v = rv.variant(sig);
        if (!v || !v.addr) return;
        var addr = hexAddr(v.addr);
        var others = gateNodeSet(gateNodes, sig).filter(function (x) { return x !== nid && enabled[x]; }).sort();
        stepGates.push({
          signal: sig, pin: g.pin, off_value: g.off_value,
          shared: gateNodeSet(gateNodes, sig).length > 1,
          polarity_inferred: !!g.polarity_inferred, collateral_nodes: others
        });
        if (signalsOff[sig]) { sharedCollateral.push({ step: idx, node: nid, signal: sig }); return; }
        var before = getField(regImage[addr], v.bit);
        var nw = setField(regImage[addr], v.bit, g.off_value);
        if (nw !== regImage[addr]) {
          var rec = writesByAddr[addr];
          if (!rec) { rec = writesByAddr[addr] = { addr: addr, reg: v.reg_name, prev: hex4(regImage[addr]), fields: [] }; }
          rec.fields.push({ signal: sig, bit: v.bit, before: before, after: g.off_value, role: 'enable' });
          regImage[addr] = nw; rec.value = hex4(nw);
        }
        signalsOff[sig] = true;
        if (others.length) stepWarn.push('共用位 ' + sig + ' 关闭同时波及：' + others.join(', ') + '（这些块的电流一并消失）');
      });
      var writes = Object.keys(writesByAddr).sort().map(function (a) { return writesByAddr[a]; });
      var note = writes.length ? null : '本级门已被前面共用位提前关掉（仍是一个测量点）';
      steps.push({
        index: idx, off_node: nid,
        off_label: node.inst_name || node.name || nid.split('::').pop(),
        device: node.device, measure: '关此级后测总电流',
        gates: stepGates, writes: writes, warnings: stepWarn, note: note
      });
    });

    var uncovered = [];
    ((fg.diagnostics || {}).uncovered_off_gates || []).forEach(function (u) {
      uncovered.push({ signal: u.signal, note: '真门但无可挂节点，序列不会自动关，需人工补' });
    });

    return nullify({
      schema_version: SCHEMA, mode: mode.id, mode_name: mode.name, reg_group: group,
      base_addr: regmap.base_addr, order_mode: orderMode,
      baseline: { note: '建立全开起始态（激活通路开、其余门关、tune/ictrl 取基线值）', writes: baselineWrites },
      steps: steps, extra_writes: mode.extra_writes || [], warnings: warnings,
      diagnostics: { uncovered_off_gates: uncovered, shared_collateral: sharedCollateral },
      stats: { baseline_regs: baselineWrites.length, steps: steps.length, gates_off: Object.keys(signalsOff).length }
    });
  }

  function renderAte(tc) {
    var L = [], eq = '';
    for (var i = 0; i < 66; i++) eq += '=';
    L.push('# ' + eq);
    L.push('# Test sequence: ' + tc.mode + '  (reg_group=' + tc.reg_group + ')');
    if (tc.mode_name) L.push('# ' + tc.mode_name);
    L.push('# Generated by Reg_tester gen_testcase (' + tc.schema_version + ')');
    L.push('# 语义：累积逐级关闭；每一步先发本段写、再测总电流；相邻步电流差=该级功耗。');
    L.push('# 数据行格式：ADDR VALUE MODULE  [; 行内注释]  ——ADDR=0x+大写8位、VALUE=0x+大写4位；');
    L.push('#            MODULE 后可有以 \' ; \' 起的行内注释；以 # 起头的整行为纯注释。');
    L.push('# ' + eq);
    L.push('');
    L.push('# --- baseline：建立全开起始态 ---');
    tc.baseline.writes.forEach(function (w) {
      var seg = [];
      (w.fields || []).forEach(function (f) {
        if (f.role === 'override') seg.push(f.signal + '=' + f.value);
        else if (f.on) seg.push(f.signal + '=on');
      });
      var cmt = seg.length ? ('  ; ' + seg.join(', ')) : '';
      L.push(w.addr + ' ' + w.value + '  baseline:' + (w.reg || '?') + cmt);
    });
    L.push('');
    tc.steps.forEach(function (st) {
      L.push('# --- step ' + st.index + ': OFF ' + st.off_label + ' (' + (st.device || '?') + ') → 测总电流 ---');
      (st.warnings || []).forEach(function (wn) { L.push('#   ⚠ ' + wn); });
      if (st.note) L.push('#   · ' + st.note);
      st.writes.forEach(function (w) {
        var ft = (w.fields || []).map(function (f) { return f.signal + '[' + f.bit + ']:' + f.before + '→' + f.after; }).join(', ');
        L.push(w.addr + ' ' + w.value + '  off:' + st.off_node + '  ; ' + w.prev + '→' + w.value + '  ' + ft);
      });
      L.push('');
    });
    if (tc.extra_writes && tc.extra_writes.length) {
      L.push('# --- extra writes（模式级额外写）---');
      tc.extra_writes.forEach(function (w) { L.push(hexAddr(w.addr) + ' ' + hex4(toInt(w.value, 0)) + '  extra  ; ' + (w.note || '')); });
      L.push('');
    }
    if (tc.diagnostics.uncovered_off_gates.length) {
      L.push('# --- ⚠ 未覆盖的门（需人工补写）---');
      tc.diagnostics.uncovered_off_gates.forEach(function (u) { L.push('#   ' + u.signal + ' : ' + u.note); });
    }
    return L.join('\n') + '\n';
  }

  var api = { generate: generate, renderAte: renderAte, setField: setField, getField: getField, hex4: hex4, hexAddr: hexAddr, shutdownRank: shutdownRank, buildEdgeMaps: buildEdgeMaps, collectGates: collectGates };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Generator = api;
})(typeof window !== 'undefined' ? window : this);
