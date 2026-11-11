/* 浏览器端文档对比引擎：忽略标点/格式，识别删除、新增、修改、移动。 */
(function (root) {
  const KEEP = new Set("+-×÷=≠≈≤≥<>√∑∏∫∂±∞^_%⋅−".split(""));
  const NOISE_RE = /(HYPERLINK|PAGEREF|_Toc\d+|TOC \\o|NUMPAGES|PAGE\s+\d+)/i;
  const MIN_MOVE_CHARS = 4;
  const MOVE_RATIO = 0.94;
  const MODIFY_RATIO = 0.42;
  const RELOCATE_RATIO = 0.55;

  const DEMO_LEFT = [
    "国泰海通证券投资基金招募说明书",
    "一、绪言",
    "本基金由国泰海通基金管理有限公司依照有关法律法规及本基金合同的规定募集。",
    "基金管理人应当恪尽职守，履行诚实信用、谨慎勤勉的义务。",
    "投资者购买基金份额，即表明承认并接受本招募说明书。",
    "二、释义",
    "基金合同：指《国泰海通证券投资基金基金合同》及对其的任何有效修订和补充。",
    "三、基金管理人",
    "基金管理人：国泰海通基金管理有限公司",
    "注册资本：人民币贰亿元整",
    "经营范围：公开募集证券投资基金管理。",
    "四、投资范围",
    "本基金投资于依法发行上市的股票、债券、货币市场工具。",
    "本基金股票投资比例为基金资产的 60%—95%。",
    "五、风险揭示",
    "投资有风险，投资者应认真阅读。",
    "基金份额净值增长率 = (期末基金份额净值 - 期初基金份额净值) / 期初基金份额净值",
  ];
  const DEMO_RIGHT = [
    "国泰海通证券投资基金招募说明书（更新）",
    "一、绪言",
    "本基金由国泰海通基金管理有限公司依照有关法律法规及本基金合同的规定募集。",
    "基金管理人必须恪尽职守，履行诚实信用、谨慎勤勉的义务。",
    "投资者申购基金份额，即表明承认并接受本招募说明书。",
    "三、基金管理人",
    "基金管理人：国泰海通基金管理有限公司",
    "注册资本：人民币叁亿元整",
    "经营范围：公开募集证券投资基金管理、基金销售。",
    "二、释义",
    "基金合同：指《国泰海通证券投资基金基金合同》及对其的任何有效修订和补充。",
    "四、投资范围",
    "本基金投资于依法发行上市的股票、债券、货币市场工具、资产支持证券。",
    "本基金股票投资比例为基金资产的 60%—95%。",
    "五、风险揭示",
    "投资有风险，投资者应认真阅读。",
    "新增：本基金可能面临流动性风险。",
    "基金份额净值增长率 = (期末基金份额净值 - 期初基金份额净值 + 期间分红) / 期初基金份额净值",
  ];

  function isIgnorable(ch) {
    if (!ch || /\s/.test(ch)) return true;
    if (KEEP.has(ch)) return false;
    try {
      return /^\p{P}$|^\p{S}$|^\p{Z}$/u.test(ch);
    } catch (e) {
      return /[^\w\u4e00-\u9fff]/.test(ch);
    }
  }

  function normToken(raw) {
    return raw.normalize("NFKC").toLowerCase();
  }

  function tokenize(text) {
    const tokens = [];
    let i = 0;
    const n = text.length;
    while (i < n) {
      const ch = text[i];
      if (isIgnorable(ch)) {
        let j = i + 1;
        while (j < n && isIgnorable(text[j])) j += 1;
        tokens.push({ text: text.slice(i, j), norm: "", isContent: false, start: i, end: j });
        i = j;
        continue;
      }
      if (/[A-Za-z0-9]/.test(ch)) {
        let j = i + 1;
        while (j < n && /[A-Za-z0-9]/.test(text[j])) j += 1;
        const raw = text.slice(i, j);
        tokens.push({ text: raw, norm: normToken(raw), isContent: true, start: i, end: j });
        i = j;
        continue;
      }
      tokens.push({ text: ch, norm: normToken(ch), isContent: true, start: i, end: i + 1 });
      i += 1;
    }
    return tokens;
  }

  function normalize(text) {
    if (!text) return "";
    return tokenize(text)
      .filter((t) => t.isContent)
      .map((t) => t.norm)
      .join("");
  }

  function compact(text) {
    return String(text || "")
      .replace(/[\u3000\xa0]/g, " ")
      .replace(/[\r\f]/g, "")
      .replace(/[ \t\u200b\u200c\u200d]+/g, " ")
      .trim();
  }

  function findLongestMatch(a, alo, ahi, b, blo, bhi) {
    const b2j = new Map();
    for (let j = blo; j < bhi; j++) {
      const k = b[j];
      if (!b2j.has(k)) b2j.set(k, []);
      b2j.get(k).push(j);
    }
    let besti = alo;
    let bestj = blo;
    let bestsize = 0;
    let j2len = new Map();
    for (let i = alo; i < ahi; i++) {
      const newj2len = new Map();
      const idxs = b2j.get(a[i]) || [];
      for (const j of idxs) {
        if (j < blo) continue;
        if (j >= bhi) break;
        const k = (j2len.get(j - 1) || 0) + 1;
        newj2len.set(j, k);
        if (k > bestsize) {
          besti = i - k + 1;
          bestj = j - k + 1;
          bestsize = k;
        }
      }
      j2len = newj2len;
    }
    return { a: besti, b: bestj, size: bestsize };
  }

  function matchingBlocks(a, b) {
    const queue = [[0, a.length, 0, b.length]];
    const matches = [];
    while (queue.length) {
      const [alo, ahi, blo, bhi] = queue.pop();
      const m = findLongestMatch(a, alo, ahi, b, blo, bhi);
      if (m.size) {
        matches.push(m);
        if (alo < m.a && blo < m.b) queue.push([alo, m.a, blo, m.b]);
        if (m.a + m.size < ahi && m.b + m.size < bhi) queue.push([m.a + m.size, ahi, m.b + m.size, bhi]);
      }
    }
    matches.sort((x, y) => x.a - y.a || x.b - y.b);
    const collapsed = [];
    let i1 = 0;
    let j1 = 0;
    let n1 = 0;
    for (const m of matches) {
      if (i1 + n1 === m.a && j1 + n1 === m.b) {
        n1 += m.size;
      } else {
        if (n1) collapsed.push({ a: i1, b: j1, size: n1 });
        i1 = m.a;
        j1 = m.b;
        n1 = m.size;
      }
    }
    if (n1) collapsed.push({ a: i1, b: j1, size: n1 });
    collapsed.push({ a: a.length, b: b.length, size: 0 });
    return collapsed;
  }

  function getOpcodes(a, b) {
    const blocks = matchingBlocks(a, b);
    const ops = [];
    let i = 0;
    let j = 0;
    for (const m of blocks) {
      if (i < m.a && j < m.b) ops.push(["replace", i, m.a, j, m.b]);
      else if (i < m.a) ops.push(["delete", i, m.a, j, m.b]);
      else if (j < m.b) ops.push(["insert", i, m.a, j, m.b]);
      if (m.size) ops.push(["equal", m.a, m.a + m.size, m.b, m.b + m.size]);
      i = m.a + m.size;
      j = m.b + m.size;
    }
    return ops;
  }

  function ratio(a, b) {
    if (!a.length && !b.length) return 1;
    let match = 0;
    for (const m of matchingBlocks(a, b)) match += m.size;
    return (2.0 * match) / (a.length + b.length);
  }

  function coalesceOpcodes(opcodes, glue) {
    glue = glue == null ? 2 : glue;
    const blocks = [];
    for (const [tag, i1, i2, j1, j2] of opcodes) {
      const kind = tag === "equal" ? "equal" : "edit";
      if (blocks.length && blocks[blocks.length - 1][0] === kind) {
        const p = blocks[blocks.length - 1];
        blocks[blocks.length - 1] = [kind, p[1], i2, p[3], j2];
      } else blocks.push([kind, i1, i2, j1, j2]);
    }
    let i = 0;
    while (i + 2 < blocks.length) {
      const a = blocks[i];
      const mid = blocks[i + 1];
      const c = blocks[i + 2];
      if (
        a[0] === "edit" &&
        mid[0] === "equal" &&
        c[0] === "edit" &&
        mid[2] - mid[1] <= glue &&
        mid[4] - mid[3] <= glue
      ) {
        blocks.splice(i, 3, ["edit", a[1], c[2], a[3], c[4]]);
        continue;
      }
      i += 1;
    }
    const out = [];
    for (const [kind, i1, i2, j1, j2] of blocks) {
      if (kind === "equal") out.push(["equal", i1, i2, j1, j2]);
      else if (i1 === i2) out.push(["insert", i1, i2, j1, j2]);
      else if (j1 === j2) out.push(["delete", i1, i2, j1, j2]);
      else out.push(["replace", i1, i2, j1, j2]);
    }
    return out;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
  }

  function mergeSegments(items) {
    const out = [];
    for (const seg of items) {
      if (!seg.text) continue;
      const last = out[out.length - 1];
      if (last && last.kind === seg.kind && last.change_id === seg.change_id) {
        last.text += seg.text;
      } else out.push({ text: seg.text, kind: seg.kind, change_id: seg.change_id || null });
    }
    return out;
  }

  function segmentsToHtml(segments, side) {
    return segments
      .map((seg) => {
        const body = escapeHtml(seg.text).replace(/\n/g, "<br>");
        if (seg.kind === "equal" || !seg.change_id) return body;
        return `<span class="hl hl-${seg.kind}" id="${seg.change_id}-${side}" data-cid="${seg.change_id}">${body}</span>`;
      })
      .join("");
  }

  function paraView(index, segments, side, box, pair, kind) {
    const cids = [];
    const seen = new Set();
    for (const s of segments) {
      if (s.change_id && !seen.has(s.change_id)) {
        seen.add(s.change_id);
        cids.push(s.change_id);
      }
    }
    return {
      index,
      html: segmentsToHtml(segments, side),
      text: segments.map((s) => s.text || "").join(""),
      box: box || null,
      change_ids: cids,
      pair: pair == null ? null : pair,
      kind: kind || "equal",
    };
  }

  function contentSpan(orig, content, i1, i2) {
    if (i1 >= i2) return "";
    return orig.slice(content[i1].start, content[i2 - 1].end);
  }

  function inlineDiff(oldT, newT, leftPara, rightPara, idGen, changes) {
    const tokA = tokenize(oldT);
    const tokB = tokenize(newT);
    const ca = tokA.filter((t) => t.isContent);
    const cb = tokB.filter((t) => t.isContent);
    if (ca.map((t) => t.norm).join("") === cb.map((t) => t.norm).join("")) {
      return [[{ text: oldT, kind: "equal", change_id: null }], [{ text: newT, kind: "equal", change_id: null }]];
    }
    const kindA = ca.map(() => "equal");
    const kindB = cb.map(() => "equal");
    const cidA = ca.map(() => null);
    const cidB = cb.map(() => null);
    for (const [tag, i1, i2, j1, j2] of coalesceOpcodes(getOpcodes(ca.map((t) => t.norm), cb.map((t) => t.norm)))) {
      if (tag === "equal") continue;
      const cid = idGen();
      const oldSpan = contentSpan(oldT, ca, i1, i2);
      const newSpan = contentSpan(newT, cb, j1, j2);
      let kind = "modify";
      if (tag === "delete") {
        kind = "delete";
        for (let i = i1; i < i2; i++) {
          kindA[i] = kind;
          cidA[i] = cid;
        }
      } else if (tag === "insert") {
        kind = "insert";
        for (let j = j1; j < j2; j++) {
          kindB[j] = kind;
          cidB[j] = cid;
        }
      } else {
        for (let i = i1; i < i2; i++) {
          kindA[i] = kind;
          cidA[i] = cid;
        }
        for (let j = j1; j < j2; j++) {
          kindB[j] = kind;
          cidB[j] = cid;
        }
      }
      changes.push({
        id: cid,
        kind,
        old_text: oldSpan,
        new_text: newSpan,
        left_para: leftPara,
        right_para: rightPara,
      });
    }
    function paint(orig, tokens, kinds, cids) {
      const segs = [];
      let contentI = -1;
      for (const t of tokens) {
        if (t.isContent) {
          contentI += 1;
          segs.push({ text: t.text, kind: kinds[contentI], change_id: cids[contentI] });
        } else segs.push({ text: t.text, kind: "equal", change_id: null });
      }
      if (!segs.length) segs.push({ text: orig, kind: "equal", change_id: null });
      return mergeSegments(segs);
    }
    return [paint(oldT, tokA, kindA, cidA), paint(newT, tokB, kindB, cidB)];
  }

  function bestPairs(leftIdx, rightIdx, aNorms, bNorms, minRatio) {
    const usedL = new Set();
    const usedR = new Set();
    const pairs = [];
    const byNorm = new Map();
    for (const j of rightIdx) {
      const n = bNorms[j];
      if (!byNorm.has(n)) byNorm.set(n, []);
      byNorm.get(n).push(j);
    }
    for (const i of leftIdx) {
      const na = aNorms[i];
      const cands = (byNorm.get(na) || []).filter((j) => !usedR.has(j));
      if (!na || !cands.length) continue;
      const j = cands.reduce((best, jj) => (Math.abs(jj - i) < Math.abs(best - i) ? jj : best), cands[0]);
      usedL.add(i);
      usedR.add(j);
      pairs.push([i, j, 1]);
    }
    const cand = [];
    const restL = leftIdx.filter((i) => !usedL.has(i));
    const restR = rightIdx.filter((j) => !usedR.has(j));
    for (const i of restL) {
      const na = aNorms[i];
      if (!na) continue;
      for (const j of restR) {
        const nb = bNorms[j];
        if (!nb) continue;
        const la = na.length;
        const lb = nb.length;
        if (Math.min(la, lb) / Math.max(la, lb) < minRatio * 0.8) continue;
        const r = ratio(na.split(""), nb.split(""));
        if (r >= minRatio) cand.push([r, i, j]);
      }
    }
    cand.sort((x, y) => y[0] - x[0] || Math.abs(x[1] - x[2]) - Math.abs(y[1] - y[2]));
    for (const [r, i, j] of cand) {
      if (usedL.has(i) || usedR.has(j)) continue;
      usedL.add(i);
      usedR.add(j);
      pairs.push([i, j, r]);
    }
    return pairs;
  }

  function compareTexts(leftParas, rightParas, leftName, rightName) {
    const leftItems = leftParas.map((p, i) => [i, p, normalize(p)]).filter((x) => x[2]);
    const rightItems = rightParas.map((p, i) => [i, p, normalize(p)]).filter((x) => x[2]);
    const aText = leftItems.map((x) => x[1]);
    const bText = rightItems.map((x) => x[1]);
    const aNorms = leftItems.map((x) => x[2]);
    const bNorms = rightItems.map((x) => x[2]);
    let n = 0;
    const idGen = () => {
      n += 1;
      return "c-" + n;
    };
    const changes = [];
    const nA = aText.length;
    const nB = bText.length;
    const statusL = Array(nA).fill("equal");
    const statusR = Array(nB).fill("equal");
    const pairRofL = Array(nA).fill(null);
    const pairLofR = Array(nB).fill(null);

    for (const [tag, i1, i2, j1, j2] of getOpcodes(aNorms, bNorms)) {
      if (tag === "equal") {
        for (let k = 0; k < i2 - i1; k++) {
          const i = i1 + k;
          const j = j1 + k;
          statusL[i] = "equal";
          statusR[j] = "equal";
          pairRofL[i] = j;
          pairLofR[j] = i;
        }
      } else if (tag === "delete") {
        for (let i = i1; i < i2; i++) statusL[i] = "delete";
      } else if (tag === "insert") {
        for (let j = j1; j < j2; j++) statusR[j] = "insert";
      } else {
        const blockL = [];
        const blockR = [];
        for (let i = i1; i < i2; i++) blockL.push(i);
        for (let j = j1; j < j2; j++) blockR.push(j);
        const paired = bestPairs(blockL, blockR, aNorms, bNorms, MODIFY_RATIO);
        const pairedL = new Set(paired.map((x) => x[0]));
        const pairedR = new Set(paired.map((x) => x[1]));
        for (const [i, j, r] of paired) {
          if (r >= 0.999 && aNorms[i] === bNorms[j]) {
            statusL[i] = "equal";
            statusR[j] = "equal";
          } else {
            statusL[i] = "modify";
            statusR[j] = "modify";
          }
          pairRofL[i] = j;
          pairLofR[j] = i;
        }
        for (const i of blockL) if (!pairedL.has(i)) statusL[i] = "delete";
        for (const j of blockR) if (!pairedR.has(j)) statusR[j] = "insert";
      }
    }

    const restL = statusL.map((st, i) => (st === "delete" ? i : -1)).filter((i) => i >= 0);
    const restR = statusR.map((st, j) => (st === "insert" ? j : -1)).filter((j) => j >= 0);
    for (const [i, j, r] of bestPairs(restL, restR, aNorms, bNorms, RELOCATE_RATIO)) {
      const la = aNorms[i].length;
      const lb = bNorms[j].length;
      if (r >= MOVE_RATIO && Math.min(la, lb) >= MIN_MOVE_CHARS) {
        statusL[i] = "move";
        statusR[j] = "move";
      } else {
        statusL[i] = "modify";
        statusR[j] = "modify";
      }
      pairRofL[i] = j;
      pairLofR[j] = i;
    }

    const leftViews = Array(nA).fill(null);
    const rightViews = Array(nB).fill(null);
    for (let i = 0; i < nA; i++) {
      const st = statusL[i];
      const j = pairRofL[i];
      if (st === "equal") {
        leftViews[i] = paraView(i, [{ text: aText[i], kind: "equal" }], "L", null, j, "equal");
      } else if (st === "delete") {
        const cid = idGen();
        changes.push({ id: cid, kind: "delete", old_text: aText[i], new_text: "", left_para: i, right_para: null });
        leftViews[i] = paraView(i, [{ text: aText[i], kind: "delete", change_id: cid }], "L", "delete", null, "delete");
      } else if (st === "move" && j != null) {
        const cid = idGen();
        changes.push({ id: cid, kind: "move", old_text: aText[i], new_text: bText[j], left_para: i, right_para: j });
        if (aNorms[i] === bNorms[j]) {
          leftViews[i] = paraView(i, [{ text: aText[i], kind: "move", change_id: cid }], "L", "move", j, "move");
          rightViews[j] = paraView(j, [{ text: bText[j], kind: "move", change_id: cid }], "R", "move", i, "move");
        } else {
          const [segsL, segsR] = inlineDiff(aText[i], bText[j], i, j, idGen, changes);
          segsL.forEach((s) => {
            if (s.kind === "equal") {
              s.kind = "move";
              s.change_id = cid;
            }
          });
          segsR.forEach((s) => {
            if (s.kind === "equal") {
              s.kind = "move";
              s.change_id = cid;
            }
          });
          leftViews[i] = paraView(i, segsL, "L", "move", j, "move");
          rightViews[j] = paraView(j, segsR, "R", "move", i, "move");
        }
      } else if (st === "modify" && j != null) {
        const [segsL, segsR] = inlineDiff(aText[i], bText[j], i, j, idGen, changes);
        if (segsL.every((s) => s.kind === "equal") && segsR.every((s) => s.kind === "equal")) {
          leftViews[i] = paraView(i, [{ text: aText[i], kind: "equal" }], "L", null, j, "equal");
          rightViews[j] = paraView(j, [{ text: bText[j], kind: "equal" }], "R", null, i, "equal");
        } else {
          leftViews[i] = paraView(i, segsL, "L", "modify", j, "modify");
          rightViews[j] = paraView(j, segsR, "R", "modify", i, "modify");
        }
      }
    }
    for (let j = 0; j < nB; j++) {
      if (rightViews[j]) continue;
      const st = statusR[j];
      if (st === "equal") {
        rightViews[j] = paraView(j, [{ text: bText[j], kind: "equal" }], "R", null, pairLofR[j], "equal");
      } else if (st === "insert") {
        const cid = idGen();
        changes.push({ id: cid, kind: "insert", old_text: "", new_text: bText[j], left_para: null, right_para: j });
        rightViews[j] = paraView(j, [{ text: bText[j], kind: "insert", change_id: cid }], "R", "insert", null, "insert");
      } else {
        rightViews[j] = paraView(j, [{ text: bText[j], kind: "equal" }], "R", null, pairLofR[j], st);
      }
    }

    let equalChars = 0;
    for (let i = 0; i < nA; i++) {
      if (statusL[i] === "equal") equalChars += aNorms[i].length;
      else if ((statusL[i] === "modify" || statusL[i] === "move") && pairRofL[i] != null) {
        const j = pairRofL[i];
        equalChars += Math.floor(
          ratio(aNorms[i].split(""), bNorms[j].split("")) * Math.min(aNorms[i].length, bNorms[j].length)
        );
      }
    }
    const total = aNorms.join("").length + bNorms.join("").length;
    let similarity = total ? (2.0 * equalChars) / total * 100 : 100;
    similarity = Math.max(0, Math.min(100, similarity));
    const counts = {
      all: changes.length,
      delete: changes.filter((c) => c.kind === "delete").length,
      insert: changes.filter((c) => c.kind === "insert").length,
      modify: changes.filter((c) => c.kind === "modify").length,
      move: changes.filter((c) => c.kind === "move").length,
    };
    return {
      similarity: Math.round(similarity * 100) / 100,
      left_name: leftName || "基准文档",
      right_name: rightName || "对照文档",
      counts,
      left: leftViews.filter(Boolean),
      right: rightViews.filter(Boolean),
      changes,
    };
  }

  function localName(el) {
    return (el.localName || el.nodeName || "").replace(/^.*:/, "").toLowerCase();
  }

  function xmlText(p) {
    const parts = [];
    function walk(node) {
      if (!node || node.nodeType !== 1) return;
      const tag = localName(node);
      if (tag === "fallback") return;
      if (tag === "omath" || tag === "omathpara") {
        const tex = compact(node.textContent || "");
        if (tex) parts.push(" ", tex, " ");
        return;
      }
      if (tag === "t" && node.textContent) parts.push(node.textContent);
      else if (tag === "tab") parts.push("\t");
      else if (tag === "br") parts.push("\n");
      for (const ch of node.childNodes) walk(ch);
    }
    walk(p);
    return parts.join("");
  }

  function cellText(tc) {
    const parts = [];
    function walk(parent) {
      for (const ch of parent.childNodes) {
        if (ch.nodeType !== 1) continue;
        const tag = localName(ch);
        if (tag === "p") {
          const t = compact(xmlText(ch));
          if (t) parts.push(t);
        } else if (tag === "tbl") {
          parts.push(...tableRows(ch));
        } else if (tag === "sdt") {
          walk(ch);
        } else if (tag === "sdtcontent" || tag === "tc") {
          walk(ch);
        }
      }
    }
    walk(tc);
    return parts.join(" ");
  }

  function tableRows(tbl) {
    const rows = [];
    for (const tr of tbl.childNodes) {
      if (tr.nodeType !== 1 || localName(tr) !== "tr") continue;
      const cells = [];
      for (const tc of tr.childNodes) {
        if (tc.nodeType !== 1 || localName(tc) !== "tc") continue;
        const t = cellText(tc);
        if (t) cells.push(t);
      }
      const line = cells.join(" | ").trim();
      if (line) rows.push(line);
    }
    return rows;
  }

  function extractFromBody(body) {
    const raw = [];
    function walk(parent) {
      for (const ch of parent.childNodes) {
        if (ch.nodeType !== 1) continue;
        const tag = localName(ch);
        if (tag === "p") {
          const text = compact(xmlText(ch));
          if (!text) continue;
          if (NOISE_RE.test(text) && !normalize(text)) continue;
          raw.push(text);
        } else if (tag === "tbl") {
          raw.push(...tableRows(ch));
        } else if (tag === "sdt" || tag === "sdtcontent") {
          walk(ch);
        }
      }
    }
    walk(body);
    return raw;
  }

  function readBuffer(file) {
    if (file && typeof file.arrayBuffer === "function") return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        resolve(fr.result);
      };
      fr.onerror = function () {
        reject(new Error("读取文件失败，请换 Microsoft Edge 或 Chrome 再试"));
      };
      fr.readAsArrayBuffer(file);
    });
  }

  function sniffExt(buf) {
    var u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    if (u8.length >= 5 && u8[0] === 0x25 && u8[1] === 0x50 && u8[2] === 0x44 && u8[3] === 0x46) {
      return "pdf";
    }
    if (
      u8.length >= 5 &&
      u8[0] === 0x7b &&
      u8[1] === 0x5c &&
      u8[2] === 0x72 &&
      u8[3] === 0x74 &&
      u8[4] === 0x66
    ) {
      return "rtf";
    }
    if (u8.length >= 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) {
      return "doc";
    }
    if (u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b) return "docx";
    return "";
  }

  function dvU16(dv, o) {
    return dv.getUint16(o, true);
  }
  function dvU32(dv, o) {
    return dv.getUint32(o, true);
  }

  function parseCfb(buf) {
    var data = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    if (data.length < 512 || data[0] !== 0xd0 || data[1] !== 0xcf) {
      throw new Error("不是 Word 复合文档");
    }
    var dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    var sectorSize = 1 << dvU16(dv, 0x1e);
    var miniSize = 1 << dvU16(dv, 0x20);
    var dirFirst = dvU32(dv, 0x30);
    var miniCutoff = dvU32(dv, 0x38);
    var miniFatFirst = dvU32(dv, 0x3c);
    var difatFirst = dvU32(dv, 0x44);
    var difatCount = dvU32(dv, 0x48);
    var difat = [];
    var i;
    for (i = 0; i < 109; i++) difat.push(dvU32(dv, 0x4c + 4 * i));
    var sec = difatFirst;
    var hops = 0;
    while (sec < 0xfffffff0 && hops < difatCount + 8) {
      var off = (sec + 1) * sectorSize;
      var n = sectorSize / 4;
      for (i = 0; i < n - 1; i++) {
        if (off + 4 * i + 4 <= data.length) difat.push(dvU32(dv, off + 4 * i));
      }
      sec = off + 4 * (n - 1) + 4 <= data.length ? dvU32(dv, off + 4 * (n - 1)) : 0xfffffffe;
      hops += 1;
    }
    var fat = [];
    for (i = 0; i < difat.length; i++) {
      var fatSec = difat[i];
      if (fatSec >= 0xfffffff0) continue;
      off = (fatSec + 1) * sectorSize;
      var cnt = sectorSize / 4;
      for (var j = 0; j < cnt; j++) {
        if (off + 4 * j + 4 <= data.length) fat.push(dvU32(dv, off + 4 * j));
      }
    }
    function readChain(start) {
      var out = [];
      var total = 0;
      var cur = start;
      var seen = {};
      var guard = 0;
      while (cur < 0xfffffff0 && guard++ < 1e6) {
        if (seen[cur]) break;
        seen[cur] = 1;
        var pos = (cur + 1) * sectorSize;
        var slice = data.subarray(pos, pos + sectorSize);
        out.push(slice);
        total += slice.length;
        if (cur >= fat.length) break;
        cur = fat[cur];
      }
      var all = new Uint8Array(total);
      var p = 0;
      for (var k = 0; k < out.length; k++) {
        all.set(out[k], p);
        p += out[k].length;
      }
      return all;
    }
    var dirData = readChain(dirFirst);
    var dirDv = new DataView(dirData.buffer, dirData.byteOffset, dirData.byteLength);
    var entries = [];
    var rootStart = 0;
    for (var di = 0; di + 128 <= dirData.length; di += 128) {
      var nameLen = dvU16(dirDv, di + 64);
      if (nameLen < 2) continue;
      var nameBytes = dirData.subarray(di, di + Math.min(64, nameLen - 2));
      var name = "";
      try {
        name = new TextDecoder("utf-16le").decode(nameBytes);
      } catch (eName) {
        name = "";
      }
      var typ = dirData[di + 66];
      var startSec = dvU32(dirDv, di + 116);
      var size = dvU32(dirDv, di + 120);
      if (typ === 5) rootStart = startSec;
      if (typ === 2 && name) entries.push({ name: name, start: startSec, size: size });
    }
    var miniStream = readChain(rootStart);
    var miniFatData = miniFatFirst < 0xfffffff0 ? readChain(miniFatFirst) : new Uint8Array(0);
    var miniFatDv = new DataView(miniFatData.buffer, miniFatData.byteOffset, miniFatData.byteLength);
    var miniFat = [];
    for (i = 0; i + 4 <= miniFatData.length; i += 4) miniFat.push(dvU32(miniFatDv, i));
    function readMini(start, size) {
      var chunks = [];
      var total = 0;
      var cur = start;
      var seen = {};
      var guard = 0;
      while (cur < 0xfffffff0 && total < size && guard++ < 1e6) {
        if (seen[cur]) break;
        seen[cur] = 1;
        var pos = cur * miniSize;
        var slice = miniStream.subarray(pos, pos + miniSize);
        chunks.push(slice);
        total += slice.length;
        if (cur >= miniFat.length) break;
        cur = miniFat[cur];
      }
      var all = new Uint8Array(Math.min(size, total));
      var p = 0;
      for (var k = 0; k < chunks.length && p < all.length; k++) {
        var ncopy = Math.min(chunks[k].length, all.length - p);
        all.set(chunks[k].subarray(0, ncopy), p);
        p += ncopy;
      }
      return all;
    }
    var streams = {};
    for (var ei = 0; ei < entries.length; ei++) {
      var e = entries[ei];
      var raw = e.size < miniCutoff ? readMini(e.start, e.size) : readChain(e.start).subarray(0, e.size);
      streams[e.name] = raw;
    }
    return streams;
  }

  var WORD_ANSI_MAP = {
    0x82: 0x201a,
    0x83: 0x0192,
    0x84: 0x201e,
    0x85: 0x2026,
    0x86: 0x2020,
    0x87: 0x2021,
    0x88: 0x02c6,
    0x89: 0x2030,
    0x8a: 0x0160,
    0x8b: 0x2039,
    0x8c: 0x0152,
    0x91: 0x2018,
    0x92: 0x2019,
    0x93: 0x201c,
    0x94: 0x201d,
    0x95: 0x2022,
    0x96: 0x2013,
    0x97: 0x2014,
    0x98: 0x02dc,
    0x99: 0x2122,
    0x9a: 0x0161,
    0x9b: 0x203a,
    0x9c: 0x0153,
    0x9f: 0x0178,
  };

  function decodeWordAnsi(bytes, lid) {
    var chinese = lid === 0x0804 || lid === 0x0004 || lid === 0x0800;
    if (chinese) {
      try {
        return new TextDecoder("gbk").decode(bytes);
      } catch (e1) {
        try {
          return new TextDecoder("gb18030").decode(bytes);
        } catch (e2) {}
      }
    }
    var out = "";
    for (var i = 0; i < bytes.length; i++) {
      var b = bytes[i];
      out += String.fromCharCode(WORD_ANSI_MAP[b] || b);
    }
    return out;
  }

  function decodeUtf16le(bytes) {
    try {
      return new TextDecoder("utf-16le").decode(bytes);
    } catch (e) {
      var s = "";
      for (var i = 0; i + 1 < bytes.length; i += 2) {
        s += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
      }
      return s;
    }
  }

  function cleanDocText(s) {
    var out = [];
    var i = 0;
    while (i < s.length) {
      var code = s.charCodeAt(i);
      if (code === 0x13) {
        i += 1;
        var depth = 1;
        while (i < s.length && depth > 0) {
          var c2 = s.charCodeAt(i);
          if (c2 === 0x13) depth += 1;
          else if (c2 === 0x14) {
            if (depth === 1) {
              i += 1;
              break;
            }
          } else if (c2 === 0x15) {
            depth -= 1;
            if (depth === 0) {
              i += 1;
              break;
            }
          }
          i += 1;
        }
        continue;
      }
      if (code === 0x14 || code === 0x15) {
        i += 1;
        continue;
      }
      if (code === 0x01 || code === 0x02 || code === 0x03 || code === 0x04 || code === 0x05 || code === 0x08) {
        i += 1;
        continue;
      }
      if (code === 0x07) out.push(" | ");
      else if (code === 0x0b || code === 0x0c || code === 0x0e || code === 0x0d) out.push("\n");
      else if (code >= 32 || code === 9 || code === 10) out.push(s.charAt(i));
      i += 1;
    }
    return out.join("");
  }

  function parasFromPlain(text) {
    return String(text || "")
      .split(/\n+/)
      .map(function (p) {
        return compact(p);
      })
      .filter(function (p) {
        if (!p) return false;
        if (NOISE_RE.test(p) && !normalize(p)) return false;
        return true;
      });
  }

  function extractDocBinary(buf) {
    var streams = parseCfb(buf);
    function findStream(want) {
      if (streams[want]) return streams[want];
      var keys = Object.keys(streams);
      var target = want.toLowerCase();
      for (var si = 0; si < keys.length; si++) {
        var kn = keys[si].replace(/^\u0005/, "").replace(/\u0000/g, "").toLowerCase();
        if (kn === target) return streams[keys[si]];
      }
      return null;
    }
    var word = findStream("WordDocument");
    if (!word || word.length < 32) throw new Error("不是有效的 Word 文件");
    var dv = new DataView(word.buffer, word.byteOffset, word.byteLength);
    if (dvU16(dv, 0) !== 0xa5ec) throw new Error("不是有效的 Word 文件");
    var flags = dvU16(dv, 0x0a);
    if (flags & 0x0100) throw new Error("文档已加密，请先去掉密码再对比");
    var lid = dvU16(dv, 0x06);
    var tableName = flags & 0x0200 ? "1Table" : "0Table";
    var table = findStream(tableName) || findStream("0Table") || findStream("1Table");
    var nFib = dvU16(dv, 2);
    var fcMin = dvU32(dv, 0x18);
    var fcMac = dvU32(dv, 0x1c);
    var text = "";
    var fcClx = 0;
    var lcbClx = 0;
    if (word.length > 426) {
      fcClx = dvU32(dv, 154 + 66 * 4);
      lcbClx = dvU32(dv, 154 + 67 * 4);
    }
    if (table && lcbClx > 0 && fcClx >= 0 && fcClx + lcbClx <= table.length) {
      var clx = table.subarray(fcClx, fcClx + lcbClx);
      var pcdt = null;
      var i = 0;
      while (i < clx.length) {
        var typ = clx[i];
        if (typ === 0x01) {
          var cb = i + 3 <= clx.length ? new DataView(clx.buffer, clx.byteOffset + i + 1, 2).getInt16(0, true) : 0;
          i += 3 + Math.max(cb, 0);
        } else if (typ === 0x02) {
          var lcb = i + 5 <= clx.length ? dvU32(new DataView(clx.buffer, clx.byteOffset, clx.byteLength), i + 1) : 0;
          pcdt = clx.subarray(i + 5, i + 5 + lcb);
          break;
        } else {
          pcdt = clx.subarray(i);
          break;
        }
      }
      if (pcdt && pcdt.length >= 16) {
        var n = Math.floor((pcdt.length - 4) / 12);
        var pdv = new DataView(pcdt.buffer, pcdt.byteOffset, pcdt.byteLength);
        var parts = [];
        var pcdOff = 4 * (n + 1);
        for (i = 0; i < n; i++) {
          var cp0 = dvU32(pdv, 4 * i);
          var cp1 = dvU32(pdv, 4 * (i + 1));
          var count = cp1 - cp0;
          if (count <= 0) continue;
          var fcComp = dvU32(pdv, pcdOff + 8 * i + 2);
          var fc = fcComp & 0x3fffffff;
          var compressed = !!(fcComp & 0x40000000);
          if (compressed) {
            var offc = Math.floor(fc / 2);
            parts.push(decodeWordAnsi(word.subarray(offc, offc + count), lid));
          } else {
            parts.push(decodeUtf16le(word.subarray(fc, fc + count * 2)));
          }
        }
        text = parts.join("");
      }
    }
    if (!text && fcMac > fcMin && fcMac <= word.length) {
      var raw = word.subarray(fcMin, fcMac);
      text = nFib >= 193 ? decodeUtf16le(raw) : decodeWordAnsi(raw, lid);
    }
    var paras = parasFromPlain(cleanDocText(text));
    if (!paras.length) throw new Error("读不到文字。若文件打不开，请另存为 .docx 再试");
    return paras;
  }

  function extractRtf(buf) {
    var u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    var src = "";
    for (var i = 0; i < u8.length; i++) src += String.fromCharCode(u8[i]);
    var start = src.indexOf("{\\rtf");
    if (start < 0) throw new Error("不是有效的 RTF");
    src = src.slice(start);
    var cpg = 1252;
    var m = src.match(/\\ansicpg(\d+)/);
    if (m) cpg = parseInt(m[1], 10);
    var decoder = null;
    var enc = cpg === 936 || cpg === 54936 ? "gbk" : cpg === 950 ? "big5" : cpg === 65001 ? "utf-8" : "windows-1252";
    try {
      decoder = new TextDecoder(enc);
    } catch (e) {
      try {
        decoder = new TextDecoder("gbk");
      } catch (e2) {
        decoder = null;
      }
    }
    var out = [];
    var i2 = 0;
    var ign = 0;
    while (i2 < src.length) {
      var ch = src.charAt(i2);
      if (ch === "\\") {
        if (src.charAt(i2 + 1) === "\\" || src.charAt(i2 + 1) === "{" || src.charAt(i2 + 1) === "}") {
          if (!ign) out.push(src.charAt(i2 + 1));
          i2 += 2;
          continue;
        }
        if (src.charAt(i2 + 1) === "'") {
          var hex = src.slice(i2 + 2, i2 + 4);
          var bytes = [];
          while (src.charAt(i2) === "\\" && src.charAt(i2 + 1) === "'") {
            bytes.push(parseInt(src.slice(i2 + 2, i2 + 4), 16) || 0);
            i2 += 4;
          }
          if (!ign && bytes.length) {
            var arr = new Uint8Array(bytes);
            out.push(decoder ? decoder.decode(arr) : String.fromCharCode.apply(null, bytes));
          }
          continue;
        }
        var wm = src.slice(i2 + 1).match(/^([a-zA-Z]+)(-?\d+)? ?/);
        if (!wm) {
          i2 += 1;
          continue;
        }
        var word = wm[1];
        var num = wm[2] ? parseInt(wm[2], 10) : 1;
        i2 += 1 + wm[0].length;
        if (word === "par" || word === "line" || word === "page" || word === "sect" || word === "row" || word === "pard") {
          if (!ign) out.push("\n");
        } else if (word === "tab") {
          if (!ign) out.push("\t");
        } else if (word === "u") {
          if (!ign && num) out.push(String.fromCharCode(num < 0 ? num + 65536 : num));
          if (src.charAt(i2) === "?") i2 += 1;
        } else if (word === "fonttbl" || word === "colortbl" || word === "stylesheet" || word === "info" || word === "pict" || word === "*") {
          ign += 1;
        }
        continue;
      }
      if (ch === "{") {
        i2 += 1;
        continue;
      }
      if (ch === "}") {
        if (ign) ign -= 1;
        i2 += 1;
        continue;
      }
      if (!ign && ch !== "\r") out.push(ch === "\n" ? "" : ch);
      i2 += 1;
    }
    var paras = parasFromPlain(out.join(""));
    if (!paras.length) throw new Error("读不到 RTF 文字，请另存为 .docx 再试");
    return paras;
  }

  function toU8Copy(buf) {
    var src = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    var copy = new Uint8Array(src.length);
    copy.set(src);
    return copy;
  }

  async function extractDocx(buf) {
    if (typeof JSZip === "undefined") throw new Error("缺少 JSZip");
    const zip = await JSZip.loadAsync(buf);
    const file = zip.file("word/document.xml");
    if (!file) throw new Error("不是有效的 .docx");
    const xml = await file.async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const body =
      doc.getElementsByTagName("w:body")[0] ||
      doc.getElementsByTagName("body")[0] ||
      doc.documentElement;
    return extractFromBody(body);
  }

  async function openPdf(buf) {
    if (typeof pdfjsLib === "undefined") throw new Error("缺少 PDF.js");
    var fileProto = typeof location !== "undefined" && location.protocol === "file:";
    function load(disableWorker) {
      return pdfjsLib.getDocument({
        data: toU8Copy(buf),
        disableWorker: !!disableWorker,
        isEvalSupported: false,
        disableAutoFetch: true,
        disableStream: true,
        verbosity: 0,
      }).promise;
    }
    if (fileProto) return await load(true);
    try {
      return await load(false);
    } catch (e1) {
      return await load(true);
    }
  }

  async function extractPdf(buf) {
    const pdf = await openPdf(buf);
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const lines = [];
      let curY = null;
      let bufLine = [];
      for (const it of content.items) {
        const y = it.transform ? it.transform[5] : 0;
        const str = (it.str || "").trim();
        if (!str) continue;
        if (curY != null && Math.abs(y - curY) > 4) {
          lines.push({ y: curY, text: compact(bufLine.join("")) });
          bufLine = [];
        }
        curY = y;
        bufLine.push(it.str);
      }
      if (bufLine.length) lines.push({ y: curY, text: compact(bufLine.join("")) });
      pages.push(lines.filter((x) => x.text));
    }
    const skip = new Set();
    if (pages.length >= 3) {
      const heads = {};
      const foots = {};
      for (const items of pages) {
        if (!items.length) continue;
        const h = items[0].text;
        const f = items[items.length - 1].text;
        if (h.length <= 60) heads[h] = (heads[h] || 0) + 1;
        if (f.length <= 60) foots[f] = (foots[f] || 0) + 1;
      }
      const thresh = Math.max(2, Math.floor(pages.length * 0.45));
      for (const bag of [heads, foots]) {
        for (const [t, c] of Object.entries(bag)) if (c >= thresh) skip.add(t);
      }
    }
    const raw = [];
    for (const items of pages) {
      for (const it of items) {
        if (skip.has(it.text)) continue;
        if (NOISE_RE.test(it.text) && !normalize(it.text)) continue;
        raw.push(it.text);
      }
    }
    return raw;
  }

  async function extractFile(file) {
    const name = file.name || "文档";
    let ext = (name.split(".").pop() || "").toLowerCase();
    const buf = await readBuffer(file);
    const sniffed = sniffExt(buf);
    if (sniffed) ext = sniffed;
    if (ext === "pdf") {
      return { paras: await extractPdf(buf), ext: ".pdf", name: name, blob: file };
    }
    if (ext === "docx") return { paras: await extractDocx(buf), ext: ".docx", name: name, blob: file };
    if (ext === "rtf") return { paras: extractRtf(buf), ext: ".rtf", name: name, blob: file };
    if (ext === "doc") {
      try {
        return { paras: extractDocBinary(buf), ext: ".doc", name: name, blob: file };
      } catch (e1) {
        try {
          return { paras: await extractDocx(buf), ext: ".doc", name: name, blob: file };
        } catch (e2) {
          try {
            return { paras: extractRtf(buf), ext: ".doc", name: name, blob: file };
          } catch (e3) {
            throw new Error(e1.message || "无法读取 .doc，请另存为 .docx 再试");
          }
        }
      }
    }
    throw new Error("仅支持 Word（.doc / .docx）、RTF 或 PDF");
  }

  function escXml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function redlineRuns(oldT, newT) {
    oldT = String(oldT || "");
    newT = String(newT || "");
    if (normalize(oldT) === normalize(newT)) return [{ text: newT || oldT, mark: "same" }];
    const ca = tokenize(oldT).filter((t) => t.isContent);
    const cb = tokenize(newT).filter((t) => t.isContent);
    if (!ca.length && !cb.length) return [{ text: newT || oldT, mark: "same" }];
    if (!ca.length) return [{ text: newT, mark: "ins" }];
    if (!cb.length) return [{ text: oldT, mark: "del" }];
    const runs = [];
    function push(text, mark) {
      if (!text) return;
      const last = runs[runs.length - 1];
      if (last && last.mark === mark) last.text += text;
      else runs.push({ text: text, mark: mark });
    }
    function take(orig, content, i1, i2) {
      if (i1 >= i2 || !content.length) return "";
      const start = i1 <= 0 ? 0 : content[i1 - 1].end;
      const end = i2 >= content.length ? orig.length : content[i2 - 1].end;
      return orig.slice(Math.max(0, start), end);
    }
    for (const [tag, i1, i2, j1, j2] of coalesceOpcodes(getOpcodes(ca.map((t) => t.norm), cb.map((t) => t.norm)))) {
      if (tag === "equal") push(take(oldT, ca, i1, i2), "same");
      else if (tag === "delete") push(take(oldT, ca, i1, i2), "del");
      else if (tag === "insert") push(take(newT, cb, j1, j2), "ins");
      else {
        push(take(oldT, ca, i1, i2), "del");
        push(take(newT, cb, j1, j2), "ins");
      }
    }
    if (!runs.length) runs.push({ text: newT || oldT, mark: "same" });
    return runs;
  }

  function manuscriptParagraphs(data) {
    const out = [];
    const counts = (data && data.counts) || {};
    function pushPara(runs) {
      const text = runs.map((r) => r.text || "").join("");
      if (!String(text).trim()) return;
      out.push(runs);
    }
    pushPara([{ text: "对比稿", mark: "title" }]);
    pushPara([{
      text: "基准：" + ((data && data.left_name) || "基准文档") + "　　对照：" + ((data && data.right_name) || "对照文档"),
      mark: "note",
    }]);
    pushPara([{
      text:
        "相似度 " + ((data && data.similarity) || 0) + "%。删除 " + (counts.delete || 0) +
        "，新增 " + (counts.insert || 0) + "，修改 " + (counts.modify || 0) +
        "，移动 " + (counts.move || 0) + "。红色删除线是删去的文字，蓝色下划线是新增的文字。",
      mark: "note",
    }]);

    const right = (data && data.right) || [];
    let rCursor = 0;
    function flushInserts(untilIndex) {
      while (rCursor < right.length) {
        const rv = right[rCursor];
        if (untilIndex != null && rv.index >= untilIndex) break;
        if (rv.kind === "insert") pushPara([{ text: rv.text || "", mark: "ins" }]);
        rCursor += 1;
      }
    }
    function advancePast(index) {
      while (rCursor < right.length && right[rCursor].index <= index) rCursor += 1;
    }
    for (const lv of (data && data.left) || []) {
      if (lv.pair != null) flushInserts(lv.pair);
      const oldT = lv.text || "";
      if (lv.kind === "delete") pushPara([{ text: oldT, mark: "del" }]);
      else if (lv.kind === "equal") pushPara([{ text: oldT, mark: "same" }]);
      else if ((lv.kind === "modify" || lv.kind === "move") && lv.pair != null) {
        const rv = right.filter((x) => x.index === lv.pair)[0];
        const newT = rv ? rv.text || "" : "";
        const runs = [];
        if (lv.kind === "move") runs.push({ text: "【位置有调整】", mark: "move" });
        if (normalize(oldT) === normalize(newT)) runs.push({ text: newT || oldT, mark: "same" });
        else runs.push.apply(runs, redlineRuns(oldT, newT));
        pushPara(runs);
        advancePast(lv.pair);
      } else pushPara([{ text: oldT, mark: "same" }]);
    }
    while (rCursor < right.length) {
      if (right[rCursor].kind === "insert") pushPara([{ text: right[rCursor].text || "", mark: "ins" }]);
      rCursor += 1;
    }
    if (out.length <= 3) pushPara([{ text: "两份文档没有发现文字差异。", mark: "note" }]);
    return out;
  }

  function runXml(run) {
    const mark = run.mark || "same";
    let rPr = "";
    if (mark === "del") rPr = '<w:rPr><w:strike/><w:color w:val="C00000"/></w:rPr>';
    else if (mark === "ins") rPr = '<w:rPr><w:u w:val="single"/><w:color w:val="1F6FB2"/></w:rPr>';
    else if (mark === "note") rPr = '<w:rPr><w:color w:val="666666"/><w:sz w:val="21"/></w:rPr>';
    else if (mark === "title") rPr = '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>';
    else if (mark === "move") rPr = '<w:rPr><w:b/><w:color w:val="B86E00"/></w:rPr>';
    return String(run.text || "")
      .split("\n")
      .map((part, i) => `<w:r>${rPr}${i ? "<w:br/>" : ""}<w:t xml:space="preserve">${escXml(part)}</w:t></w:r>`)
      .join("");
  }

  function paragraphXml(p) {
    if (Array.isArray(p)) {
      const runs = p.filter((r) => r && r.text);
      if (!runs.length || !runs.map((r) => r.text).join("").trim()) return "";
      return `<w:p>${runs.map(runXml).join("")}</w:p>`;
    }
    const text = String(p || "");
    if (!text.trim()) return "";
    return `<w:p><w:r><w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`;
  }

  async function buildDocx(paragraphs) {
    const body = (paragraphs || []).map(paragraphXml).filter(Boolean).join("");
    const docXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:body>" +
      body +
      "<w:sectPr/></w:body></w:document>";
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>"
    );
    zip.file(
      "_rels/.rels",
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        "</Relationships>"
    );
    zip.file("word/document.xml", docXml);
    zip.file(
      "word/_rels/document.xml.rels",
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    );
    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  }

  function utf16beHex(str) {
    var hex = "";
    var s = String(str || "");
    for (var i = 0; i < s.length; i++) {
      hex += ("000" + s.charCodeAt(i).toString(16)).slice(-4).toUpperCase();
    }
    return hex;
  }

  function wrapPdfLine(text, maxW, fontSize) {
    var raw = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    var out = [];
    var parts = raw.split("\n");
    for (var p = 0; p < parts.length; p++) {
      var line = parts[p];
      if (!line) {
        out.push("");
        continue;
      }
      var cur = "";
      var w = 0;
      for (var i = 0; i < line.length; i++) {
        var ch = line.charAt(i);
        var cw = (line.charCodeAt(i) < 128 ? 0.5 : 1) * fontSize;
        if (cur && w + cw > maxW) {
          out.push(cur);
          cur = ch;
          w = cw;
        } else {
          cur += ch;
          w += cw;
        }
      }
      if (cur) out.push(cur);
    }
    return out;
  }

  function buildPdf(paragraphs) {
    var pageW = 595.28;
    var pageH = 841.89;
    var margin = 54;
    var fontSize = 11;
    var lineH = 18;
    var maxW = pageW - margin * 2;
    var pages = [[]];
    var y = 72;

    function pushLine(ln) {
      if (y > pageH - 56) {
        pages.push([]);
        y = 72;
      }
      pages[pages.length - 1].push({ y: y, text: ln });
      y += lineH;
    }

    var list = paragraphs || [];
    for (var pi = 0; pi < list.length; pi++) {
      var lines = wrapPdfLine(list[pi], maxW, fontSize);
      if (!lines.length) {
        y += lineH * 0.3;
        continue;
      }
      for (var li = 0; li < lines.length; li++) pushLine(lines[li]);
      y += 8;
    }
    if (!pages[0].length) pages[0].push({ y: 72, text: " " });

    function pageStream(items) {
      var s = "BT\n/F1 " + fontSize + " Tf\n";
      var prevY = null;
      for (var i = 0; i < items.length; i++) {
        var yPdf = pageH - items[i].y - fontSize;
        if (prevY == null) s += margin + " " + yPdf.toFixed(2) + " Td\n";
        else s += "0 " + (yPdf - prevY).toFixed(2) + " Td\n";
        prevY = yPdf;
        if (items[i].text) s += "<" + utf16beHex(items[i].text) + "> Tj\n";
      }
      return s + "ET\n";
    }

    var n = pages.length;
    var objects = {};
    objects[1] = "<< /Type /Catalog /Pages 2 0 R /Lang (zh-CN) >>";
    var kids = [];
    var i;
    for (i = 0; i < n; i++) kids.push(6 + i + " 0 R");
    objects[2] = "<< /Type /Pages /Kids [" + kids.join(" ") + "] /Count " + n + " >>";
    objects[3] =
      "<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [4 0 R] >>";
    objects[4] =
      "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> /FontDescriptor 5 0 R /DW 1000 >>";
    objects[5] =
      "<< /Type /FontDescriptor /FontName /STSong-Light /Flags 6 /FontBBox [-25 -254 1000 880] /ItalicAngle 0 /Ascent 880 /Descent -254 /CapHeight 880 /StemV 44 >>";
    for (i = 0; i < n; i++) {
      var pageId = 6 + i;
      var streamId = 6 + n + i;
      objects[pageId] =
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " +
        pageW +
        " " +
        pageH +
        "] /Resources << /Font << /F1 3 0 R >> >> /Contents " +
        streamId +
        " 0 R >>";
      var stream = pageStream(pages[i]);
      objects[streamId] = "<< /Length " + stream.length + " >>\nstream\n" + stream + "endstream";
    }

    var maxId = 5 + 2 * n;
    var out = "%PDF-1.4\n";
    var offsets = [0];
    for (i = 1; i <= maxId; i++) {
      offsets[i] = out.length;
      out += i + " 0 obj\n" + objects[i] + "\nendobj\n";
    }
    var xrefPos = out.length;
    var xref = "xref\n0 " + (maxId + 1) + "\n0000000000 65535 f \n";
    for (i = 1; i <= maxId; i++) {
      xref += ("0000000000" + offsets[i]).slice(-10) + " 00000 n \n";
    }
    out += xref;
    out += "trailer\n<< /Size " + (maxId + 1) + " /Root 1 0 R >>\nstartxref\n" + xrefPos + "\n%%EOF\n";
    return new Blob([out], { type: "application/pdf" });
  }

  function parasUnchanged(oldP, newP) {
    const a = (oldP || []).map((x) => compact(x)).filter(Boolean);
    const b = (newP || []).map((x) => compact(x)).filter(Boolean);
    return a.length === b.length && a.every((x, i) => x === b[i]);
  }

  root.DocDiff = {
    DEMO_LEFT,
    DEMO_RIGHT,
    compact,
    normalize,
    compareTexts,
    extractFile,
    buildDocx,
    buildPdf,
    manuscriptParagraphs,
    parasUnchanged,
  };
})(window);
