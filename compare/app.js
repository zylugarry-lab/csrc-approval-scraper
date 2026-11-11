(() => {
  function resolveUrl(rel) {
    try {
      return new URL(rel, document.baseURI).href;
    } catch (e) {
      return rel;
    }
  }
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = resolveUrl("vendor/pdf.worker.min.js");
  }
  const $ = (id) => document.getElementById(id);

  const fileBase = $("fileBase");
  const fileComp = $("fileComp");
  const nameBase = $("nameBase");
  const nameComp = $("nameComp");
  const btnCompare = $("btnCompare");
  const btnDemo = $("btnDemo");
  const btnOpenEdit = $("btnOpenEdit");
  const btnNewBlank = $("btnNewBlank");
  const btnSwapFiles = $("btnSwapFiles");
  const btnReselect = $("btnReselect");
  const uploadPanel = $("uploadPanel");
  const resultPanel = $("resultPanel");
  const editPanel = $("editPanel");
  const editPane = $("editPane");
  const leftPane = $("leftPane");
  const rightPane = $("rightPane");
  const docsWrap = $("docsWrap");
  const svg = $("connectors");
  const changeList = $("changeList");
  const loading = $("loading");
  const toast = $("toast");
  const syncScroll = $("syncScroll");
  const swapDocs = $("swapDocs");
  const showLinks = $("showLinks");

  let token = "local";
  let data = null;
  let session = {
    demo: false,
    left: { paras: [], ext: ".docx", name: "", blob: null },
    right: { paras: [], ext: ".docx", name: "", blob: null },
  };
  let filterKind = "all";
  let selectedId = null;
  let syncing = false;
  let files = { base: null, comp: null };
  let anchors = [{ l: 0, r: 0 }];
  let editSession = { paras: [], ext: ".docx", name: "", blob: null };
  let activeEditPane = null;
  let activePara = null;

  const KIND_LABEL = {
    delete: "删除",
    insert: "新增",
    modify: "修改",
    move: "移动",
  };

  function showToast(msg) {
    toast.textContent = msg;
    toast.hidden = false;
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => {
      toast.hidden = true;
    }, 3200);
  }

  function setLoading(on, msg) {
    loading.hidden = !on;
    if (msg) $("loadingText").textContent = msg;
  }

  function updateCompareEnabled() {
    btnCompare.disabled = !(files.base && files.comp);
    btnOpenEdit.disabled = !(files.base || files.comp);
  }

  function bindDrop(el, which) {
    function setOver(e) {
      e.preventDefault();
      e.stopPropagation();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
      el.classList.add("over");
    }
    el.addEventListener("dragenter", setOver);
    el.addEventListener("dragover", setOver);
    el.addEventListener("dragleave", (e) => {
      if (!el.contains(e.relatedTarget)) el.classList.remove("over");
    });
    el.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      el.classList.remove("over");
      const list = (e.dataTransfer && e.dataTransfer.files) || [];
      const f = list[0];
      if (f) setFile(which, f);
    });
  }

  function markDrops() {
    $("dropBase").classList.toggle("ready", !!files.base);
    $("dropComp").classList.toggle("ready", !!files.comp);
  }

  function setFile(which, f) {
    files[which] = f;
    (which === "base" ? nameBase : nameComp).textContent = f.name;
    markDrops();
    updateCompareEnabled();
  }

  bindDrop($("dropBase"), "base");
  bindDrop($("dropComp"), "comp");
  fileBase.addEventListener("change", () => {
    if (fileBase.files[0]) setFile("base", fileBase.files[0]);
  });
  fileComp.addEventListener("change", () => {
    if (fileComp.files[0]) setFile("comp", fileComp.files[0]);
  });

  btnSwapFiles.addEventListener("click", () => {
    const a = files.base;
    files.base = files.comp;
    files.comp = a;
    nameBase.textContent = files.base ? files.base.name : "点击选择，或把文件拖到这里";
    nameComp.textContent = files.comp ? files.comp.name : "点击选择，或把文件拖到这里";
    markDrops();
  });

  function goHome() {
    resultPanel.hidden = true;
    editPanel.hidden = true;
    uploadPanel.hidden = false;
    btnReselect.hidden = true;
    $("btnExportBoth").hidden = true;
    $("btnManuscript").hidden = true;
    $("btnRecompare").hidden = true;
    $("btnEditSaveTop").hidden = true;
    $("btnEditSavePdfTop").hidden = true;
    $("btnSaveLeftTop").hidden = true;
    $("btnSaveLeftPdfTop").hidden = true;
    $("btnSaveRightTop").hidden = true;
    $("btnSaveRightPdfTop").hidden = true;
    $("saveBar").hidden = true;
    swapDocs.checked = false;
    svg.innerHTML = "";
    leftPane.closest(".pane").classList.remove("is-dirty");
    rightPane.closest(".pane").classList.remove("is-dirty");
  }
  btnReselect.addEventListener("click", goHome);

  function paraHtml(p, side) {
    const box = p.box ? ` box-${p.box}` : "";
    const cids = (p.change_ids || []).join(" ");
    const pair = p.pair == null ? "" : ` data-pair="${p.pair}"`;
    const kind = p.kind ? ` data-kind="${p.kind}"` : "";
    return `<p class="para${box}" id="para-${side}-${p.index}" data-index="${p.index}"${pair}${kind} data-cids="${cids}" contenteditable="true" spellcheck="false">${p.html || ""}</p>`;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[c]));
  }

  function preview(text, n) {
    const t = (text || "").replace(/\s+/g, " ").trim();
    if (!t) return "（无）";
    const cut = t.length > n ? t.slice(0, n) + "…" : t;
    return escapeHtml(cut);
  }

  function visibleChanges() {
    if (!data) return [];
    if (filterKind === "all") return data.changes;
    return data.changes.filter((c) => c.kind === filterKind);
  }

  function renderSidebar() {
    const list = visibleChanges();
    changeList.innerHTML = list
      .map((c, i) => {
        return `<article class="card${c.id === selectedId ? " on" : ""}" data-id="${c.id}">
          <div><span class="idx">${i + 1}.</span><span class="kind ${c.kind}">${KIND_LABEL[c.kind]}</span></div>
          <div class="row"><span class="lab">原文</span>${preview(c.old_text, 80)}</div>
          <div class="row"><span class="lab">差异</span>${preview(c.new_text, 80)}</div>
          <div class="card-actions">
            <button type="button" class="btn-apply" data-apply="${c.id}" data-side="left">改到左边</button>
            <button type="button" class="btn-apply alt" data-apply="${c.id}" data-side="right">改到右边</button>
          </div>
        </article>`;
      })
      .join("");
  }

  function renderResult(payload) {
    data = payload;
    token = payload.token;
    $("paneLeftName").textContent = payload.left_name;
    $("paneRightName").textContent = payload.right_name;
    $("similarity").textContent = payload.similarity.toFixed(2) + "%";
    $("cntAll").textContent = payload.counts.all;
    $("cntDelete").textContent = payload.counts.delete;
    $("cntInsert").textContent = payload.counts.insert;
    $("cntModify").textContent = payload.counts.modify;
    $("cntMove").textContent = payload.counts.move;
    leftPane.innerHTML = payload.left.map((p) => paraHtml(p, "L")).join("");
    rightPane.innerHTML = payload.right.map((p) => paraHtml(p, "R")).join("");
    selectedId = null;
    renderSidebar();
    uploadPanel.hidden = true;
    editPanel.hidden = true;
    resultPanel.hidden = false;
    btnReselect.hidden = false;
    $("btnExportBoth").hidden = false;
    $("btnManuscript").hidden = false;
    $("btnRecompare").hidden = false;
    $("btnEditSaveTop").hidden = true;
    $("btnEditSavePdfTop").hidden = true;
    $("btnSaveLeftTop").hidden = false;
    $("btnSaveLeftPdfTop").hidden = false;
    $("btnSaveRightTop").hidden = false;
    $("btnSaveRightPdfTop").hidden = false;
    showSaveBar("compare");
    requestAnimationFrame(() => {
      rebuildAnchors();
      drawConnectors();
      setTimeout(() => {
        rebuildAnchors();
        drawConnectors();
      }, 80);
    });
  }

  function contentTop(el, pane) {
    return el.getBoundingClientRect().top - pane.getBoundingClientRect().top + pane.scrollTop;
  }

  function rebuildAnchors() {
    const pts = [{ l: 0, r: 0 }];
    let lastL = -1;
    let lastR = -1;
    leftPane.querySelectorAll(".para").forEach((el) => {
      const kind = el.dataset.kind || "";
      if (kind === "move" || kind === "delete") return;
      if (el.dataset.pair == null || el.dataset.pair === "") return;
      const other = document.getElementById("para-R-" + el.dataset.pair);
      if (!other || other.dataset.kind === "move") return;
      const l = contentTop(el, leftPane);
      const r = contentTop(other, rightPane);
      if (l + 2 < lastL || r + 2 < lastR) return;
      lastL = l;
      lastR = r;
      pts.push({ l, r });
    });
    pts.push({ l: leftPane.scrollHeight, r: rightPane.scrollHeight });
    anchors = pts;
  }

  function mapAlongAnchors(y, fromKey, toKey) {
    const pts = anchors;
    if (!pts || pts.length < 2) return y;
    let i = 0;
    while (i < pts.length - 2 && pts[i + 1][fromKey] <= y) i += 1;
    const a = pts[i];
    const b = pts[i + 1];
    const span = b[fromKey] - a[fromKey];
    const t = span <= 1 ? 0 : (y - a[fromKey]) / span;
    return a[toKey] + t * (b[toKey] - a[toKey]);
  }

  function colorFor(kind) {
    return {
      delete: "#e53935",
      insert: "#2e7d32",
      modify: "#ef6c00",
      move: "#6a1b9a",
    }[kind] || "#1565c0";
  }

  function findPair(cid) {
    const a = document.getElementById(cid + "-L") || document.querySelector(`#leftPane [data-cid="${cid}"]`);
    const b = document.getElementById(cid + "-R") || document.querySelector(`#rightPane [data-cid="${cid}"]`);
    const ch = data && data.changes.find((x) => x.id === cid);
    const pa =
      a ||
      (ch && ch.left_para != null
        ? document.getElementById("para-L-" + ch.left_para)
        : null);
    const pb =
      b ||
      (ch && ch.right_para != null
        ? document.getElementById("para-R-" + ch.right_para)
        : null);
    return [pa, pb, ch];
  }

  function drawConnectors() {
    svg.innerHTML = "";
    if (!data || !showLinks.checked) return;
    const wrap = docsWrap.getBoundingClientRect();
    svg.setAttribute("viewBox", `0 0 ${wrap.width} ${wrap.height}`);
    svg.setAttribute("width", wrap.width);
    svg.setAttribute("height", wrap.height);

    const paneHeadEl = docsWrap.querySelector(".pane-head");
    const paneHeadH = paneHeadEl ? paneHeadEl.getBoundingClientRect().height : 0;

    const candidates = [];
    const kindsWanted =
      filterKind === "all"
        ? new Set(["modify", "move"])
        : new Set([filterKind]);

    for (const ch of data.changes) {
      if (!kindsWanted.has(ch.kind) && ch.id !== selectedId) continue;
      if (ch.kind === "delete" || ch.kind === "insert") {
        if (ch.id !== selectedId) continue;
      }
      const [elL, elR] = findPair(ch.id);
      if (!elL || !elR) continue;
      const rL = elL.getBoundingClientRect();
      const rR = elR.getBoundingClientRect();
      const visible =
        rL.bottom > wrap.top + paneHeadH &&
        rL.top < wrap.bottom &&
        rR.bottom > wrap.top + paneHeadH &&
        rR.top < wrap.bottom;
      if (!visible && ch.id !== selectedId) continue;
      const y1 = rL.top + rL.height / 2 - wrap.top;
      const y2 = rR.top + rR.height / 2 - wrap.top;
      const x1 = rL.right - wrap.left;
      const x2 = rR.left - wrap.left;
      if (x2 <= x1 + 8) continue;
      candidates.push({ ch, x1, y1, x2, y2, mid: Math.abs((y1 + y2) / 2 - wrap.height / 2) });
    }
    candidates.sort((a, b) => a.mid - b.mid);
    const keep = candidates.filter((c) => c.ch.id === selectedId);
    for (const c of candidates) {
      if (keep.length >= 36) break;
      if (!keep.includes(c)) keep.push(c);
    }
    const ns = "http://www.w3.org/2000/svg";
    for (const c of keep) {
      const path = document.createElementNS(ns, "path");
      const cx1 = c.x1 + Math.min(80, (c.x2 - c.x1) * 0.35);
      const cx2 = c.x2 - Math.min(80, (c.x2 - c.x1) * 0.35);
      path.setAttribute(
        "d",
        `M ${c.x1} ${c.y1} C ${cx1} ${c.y1}, ${cx2} ${c.y2}, ${c.x2} ${c.y2}`
      );
      path.setAttribute("stroke", colorFor(c.ch.kind));
      if (c.ch.id === selectedId) path.classList.add("sel");
      svg.appendChild(path);
    }
  }

  function clearActive() {
    document.querySelectorAll(".hl.active, .para.flash").forEach((el) => {
      el.classList.remove("active", "flash");
    });
  }

  function jumpTo(cid) {
    selectedId = cid;
    clearActive();
    const [elL, elR] = findPair(cid);
    syncing = true;
    scrollLeader = null;
    const focus = 0.28;
    if (elL) {
      elL.classList.add(elL.classList.contains("para") ? "flash" : "active");
      const y = contentTop(elL, leftPane);
      leftPane.scrollTop = Math.max(0, y - leftPane.clientHeight * focus);
    }
    if (elR) {
      elR.classList.add(elR.classList.contains("para") ? "flash" : "active");
      const y = contentTop(elR, rightPane);
      rightPane.scrollTop = Math.max(0, y - rightPane.clientHeight * focus);
    }
    renderSidebar();
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        syncing = false;
        drawConnectors();
      });
    });
  }

  changeList.addEventListener("click", (e) => {
    const applyBtn = e.target.closest("[data-apply]");
    if (applyBtn) {
      e.preventDefault();
      e.stopPropagation();
      applyChange(applyBtn.getAttribute("data-apply"), applyBtn.getAttribute("data-side") || "left");
      return;
    }
    const card = e.target.closest(".card");
    if (card) jumpTo(card.dataset.id);
  });

  docsWrap.addEventListener("click", (e) => {
    if (e.target.closest(".para")) return;
    const hl = e.target.closest("[data-cid]");
    if (hl) jumpTo(hl.getAttribute("data-cid"));
  });

  $("filters").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-kind]");
    if (!btn) return;
    filterKind = btn.dataset.kind;
    $("filters").querySelectorAll("button").forEach((b) => b.classList.toggle("on", b === btn));
    renderSidebar();
    drawConnectors();
  });

  function onScroll(from, to, fromSide) {
    if (!syncScroll.checked || syncing) {
      drawConnectors();
      return;
    }
    if (scrollLeader && from !== scrollLeader) {
      drawConnectors();
      return;
    }
    rebuildAnchors();
    syncing = true;
    const fromKey = fromSide === "L" ? "l" : "r";
    const toKey = fromSide === "L" ? "r" : "l";
    const focus = from.clientHeight * 0.28;
    const srcY = from.scrollTop + focus;
    const dstY = mapAlongAnchors(srcY, fromKey, toKey);
    const nextTop = Math.max(0, dstY - focus);
    to.scrollTop = nextTop;
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        syncing = false;
        drawConnectors();
      });
    });
  }
  let scrollLeader = null;
  function markLeader(el) {
    scrollLeader = el;
  }
  leftPane.addEventListener("wheel", () => markLeader(leftPane), { passive: true });
  rightPane.addEventListener("wheel", () => markLeader(rightPane), { passive: true });
  leftPane.addEventListener("pointerdown", () => markLeader(leftPane));
  rightPane.addEventListener("pointerdown", () => markLeader(rightPane));
  leftPane.addEventListener("scroll", () => onScroll(leftPane, rightPane, "L"), { passive: true });
  rightPane.addEventListener("scroll", () => onScroll(rightPane, leftPane, "R"), { passive: true });
  window.addEventListener("resize", () => {
    rebuildAnchors();
    drawConnectors();
  });
  function collectParas(pane) {
    return [...pane.querySelectorAll(".para")]
      .map((p) => (p.innerText || "").replace(/\u00a0/g, " ").replace(/\n+$/, ""))
      .filter((t) => t.trim());
  }

  function paraPlain(el) {
    return (el.innerText || "").replace(/\u00a0/g, " ").replace(/\n+$/, "");
  }

  function unwrapHighlights(para) {
    para.querySelectorAll("span.hl").forEach((sp) => {
      sp.replaceWith(document.createTextNode(sp.textContent || ""));
    });
    para.classList.remove("box-delete", "box-insert", "box-modify", "box-move");
    para.removeAttribute("data-cids");
    para.normalize();
  }

  function markDirty(pane) {
    const wrap = pane.closest ? pane.closest(".pane") : null;
    if (wrap) wrap.classList.add("is-dirty");
    if (pane === editPane) $("editMeta").textContent = collectParas(editPane).length + " 段";
    if (pane === leftPane || pane === rightPane) {
      rebuildAnchors();
      drawConnectors();
    }
  }

  function normalizeEditable(pane) {
    pane.querySelectorAll(":scope > p, :scope > div").forEach((el) => {
      el.classList.add("para");
      el.setAttribute("contenteditable", "true");
      el.setAttribute("spellcheck", "false");
    });
    markDirty(pane);
  }

  function makeParaEl(text) {
    const p = document.createElement("p");
    p.className = "para";
    p.setAttribute("contenteditable", "true");
    p.setAttribute("spellcheck", "false");
    p.textContent = text || "";
    return p;
  }

  function caretOffset(para) {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return 0;
    const range = sel.getRangeAt(0);
    const pre = document.createRange();
    pre.selectNodeContents(para);
    try {
      pre.setEnd(range.startContainer, range.startOffset);
    } catch (err) {
      return 0;
    }
    return pre.toString().length;
  }

  function placeCaret(el, offset) {
    const sel = window.getSelection();
    if (!sel) return;
    const range = document.createRange();
    let remaining = Math.max(0, offset || 0);
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    if (!node) {
      el.appendChild(document.createTextNode(""));
      node = el.firstChild;
      range.setStart(node, 0);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    while (node) {
      if (remaining <= node.length) {
        range.setStart(node, remaining);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        return;
      }
      remaining -= node.length;
      const next = walker.nextNode();
      if (!next) {
        range.setStart(node, node.length);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        return;
      }
      node = next;
    }
  }

  function splitPara(para) {
    unwrapHighlights(para);
    const offset = caretOffset(para);
    const text = paraPlain(para);
    const left = text.slice(0, offset);
    const right = text.slice(offset);
    para.textContent = left;
    const next = makeParaEl(right);
    para.after(next);
    next.focus();
    placeCaret(next, 0);
    markDirty(para.parentNode);
  }

  function mergePara(para) {
    const prev = para.previousElementSibling;
    if (!prev || (prev.className || "").indexOf("para") === -1) return false;
    unwrapHighlights(prev);
    unwrapHighlights(para);
    const keep = paraPlain(prev);
    prev.textContent = keep + paraPlain(para);
    para.parentNode.removeChild(para);
    prev.focus();
    placeCaret(prev, keep.length);
    markDirty(prev.parentNode);
    activePara = prev;
    return true;
  }

  function insertParaIn(pane) {
    const cur =
      activePara && pane.contains(activePara)
        ? activePara
        : pane.querySelector(".para:last-child");
    const p = makeParaEl("");
    if (cur) cur.after(p);
    else pane.appendChild(p);
    p.focus();
    placeCaret(p, 0);
    activePara = p;
    activeEditPane = pane;
    markDirty(pane);
  }

  function deleteParaIn(pane) {
    const cur = activePara && pane.contains(activePara) ? activePara : null;
    if (!cur) {
      showToast("请先点选要删除的段落");
      return;
    }
    const paras = pane.querySelectorAll(".para");
    if (paras.length <= 1) {
      cur.textContent = "";
      markDirty(pane);
      return;
    }
    const next = cur.nextElementSibling || cur.previousElementSibling;
    cur.parentNode.removeChild(cur);
    if (next && (next.className || "").indexOf("para") !== -1) {
      next.focus();
      placeCaret(next, paraPlain(next).length);
      activePara = next;
    }
    markDirty(pane);
  }

  function removeParaEl(el, pane) {
    if (!el) return;
    if (pane.querySelectorAll(".para").length <= 1) el.textContent = "";
    else el.parentNode.removeChild(el);
  }

  function insertFromOther(targetPane, sourceEl, sourceIndex, sourcePrefix) {
    const text = sourceEl ? paraPlain(sourceEl) : "";
    const p = makeParaEl(text);
    let anchor = null;
    if (sourceIndex != null && sourceIndex > 0) {
      const prev = document.getElementById("para-" + sourcePrefix + "-" + (sourceIndex - 1));
      if (prev && prev.getAttribute("data-pair")) {
        const destPrefix = sourcePrefix === "R" ? "L" : "R";
        anchor = document.getElementById("para-" + destPrefix + "-" + prev.getAttribute("data-pair"));
      }
    }
    if (anchor) anchor.after(p);
    else targetPane.appendChild(p);
    return p;
  }

  async function recompareCurrent() {
    if (!session.left || !session.right) return;
    await runCompare(
      {
        paras: collectParas(leftPane),
        ext: session.left.ext,
        name: session.left.name,
        blob: session.left.blob,
      },
      {
        paras: collectParas(rightPane),
        ext: session.right.ext,
        name: session.right.name,
        blob: session.right.blob,
      },
      session.demo
    );
  }

  async function applyChange(cid, side) {
    if (!data) return;
    const ch = data.changes.filter(function (x) {
      return x.id === cid;
    })[0];
    if (!ch) return;
    const toLeft = side !== "right";
    const pane = toLeft ? leftPane : rightPane;
    const elL = ch.left_para != null ? document.getElementById("para-L-" + ch.left_para) : null;
    const elR = ch.right_para != null ? document.getElementById("para-R-" + ch.right_para) : null;
    if (ch.kind === "modify" || ch.kind === "move") {
      const dest = toLeft ? elL : elR;
      const src = toLeft ? elR : elL;
      if (!dest) return;
      unwrapHighlights(dest);
      dest.textContent = src
        ? paraPlain(src)
        : toLeft
          ? ch.new_text || ""
          : ch.old_text || "";
    } else if (ch.kind === "insert") {
      if (toLeft) insertFromOther(leftPane, elR, ch.right_para, "R");
      else removeParaEl(elR, rightPane);
    } else if (ch.kind === "delete") {
      if (toLeft) removeParaEl(elL, leftPane);
      else insertFromOther(rightPane, elL, ch.left_para, "L");
    }
    markDirty(pane);
    await recompareCurrent();
    if (data && data.changes && data.changes.length) {
      jumpTo(data.changes[0].id);
      showToast("已改好这一条，继续下一条");
    } else {
      showToast("左右已经一致，可保存为 Word");
    }
  }

  function bindEditor(pane) {
    pane.addEventListener("focusin", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (para) {
        activeEditPane = pane;
        activePara = para;
      }
    });
    pane.addEventListener("input", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (para) unwrapHighlights(para);
      normalizeEditable(pane);
    });
    pane.addEventListener("keydown", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (!para) return;
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        splitPara(para);
      } else if (e.key === "Backspace") {
        const sel = window.getSelection();
        if (!sel || !sel.rangeCount || !sel.isCollapsed) return;
        if (caretOffset(para) !== 0) return;
        e.preventDefault();
        mergePara(para);
      }
    });
    pane.addEventListener("paste", (e) => {
      e.preventDefault();
      var clip = e.clipboardData || window.clipboardData;
      var text = "";
      if (clip) {
        text = clip.getData ? clip.getData("text/plain") || clip.getData("Text") || "" : "";
      }
      var blocks = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
      if (blocks.length > 1) {
        const para = e.target.closest && e.target.closest(".para");
        if (para) {
          unwrapHighlights(para);
          const offset = caretOffset(para);
          const cur = paraPlain(para);
          para.textContent = cur.slice(0, offset) + blocks[0];
          let last = para;
          for (let i = 1; i < blocks.length; i++) {
            const np = makeParaEl(blocks[i]);
            last.after(np);
            last = np;
          }
          last.focus();
          placeCaret(last, paraPlain(last).length);
          markDirty(pane);
          return;
        }
      }
      var ok = false;
      try {
        ok = document.execCommand("insertText", false, text);
      } catch (err) {
        ok = false;
      }
      if (ok) return;
      var sel = window.getSelection();
      if (!sel || !sel.rangeCount) return;
      var range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(document.createTextNode(text));
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    });
  }
  bindEditor(leftPane);
  bindEditor(rightPane);
  bindEditor(editPane);

  function updateEditMeta() {
    $("editFileName").textContent = editSession.name || "未命名.docx";
    $("editMeta").textContent = (editPane.querySelectorAll(".para").length || 0) + " 段";
  }

  function openEditor(rec) {
    editSession = rec;
    const paras = rec.paras && rec.paras.length ? rec.paras : [""];
    editPane.innerHTML = "";
    paras.forEach((t) => {
      editPane.appendChild(makeParaEl(t));
    });
    uploadPanel.hidden = true;
    resultPanel.hidden = true;
    editPanel.hidden = false;
    btnReselect.hidden = false;
    $("btnEditSaveTop").hidden = false;
    $("btnEditSavePdfTop").hidden = false;
    $("btnExportBoth").hidden = true;
    $("btnManuscript").hidden = true;
    $("btnRecompare").hidden = true;
    $("btnSaveLeftTop").hidden = true;
    $("btnSaveLeftPdfTop").hidden = true;
    $("btnSaveRightTop").hidden = true;
    $("btnSaveRightPdfTop").hidden = true;
    showSaveBar("edit");
    updateEditMeta();
    const first = editPane.querySelector(".para");
    if (first) {
      first.focus();
      placeCaret(first, paraPlain(first).length);
      activePara = first;
      activeEditPane = editPane;
    }
  }

  function showSaveBar(kind) {
    const bar = $("saveBar");
    bar.hidden = false;
    const isEdit = kind === "edit";
    $("saveBarText").textContent = isEdit
      ? "改完后可保存为 Word 或 PDF"
      : "左右都能改。可分别保存为 Word 或 PDF";
    $("btnSaveBarEdit").hidden = !isEdit;
    $("btnSaveBarEditPdf").hidden = !isEdit;
    $("btnSaveBarLeft").hidden = isEdit;
    $("btnSaveBarLeftPdf").hidden = isEdit;
    $("btnSaveBarRight").hidden = isEdit;
    $("btnSaveBarRightPdf").hidden = isEdit;
    $("btnSaveBarManuscript").hidden = isEdit;
  }

  function fileName(name, tag, fmt) {
    const stem = String(name || "文档").replace(/\.[^.]+$/, "") || "文档";
    const ext = fmt === "pdf" ? ".pdf" : ".docx";
    return stem + "-" + (tag || "已修改") + ext;
  }

  function downloadBlob(blob, filename) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, filename);
      return;
    }
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.rel = "noopener";
    a.style.cssText = "position:fixed;left:0;top:0;width:1px;height:1px;opacity:0;";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      if (a.parentNode) a.parentNode.removeChild(a);
      URL.revokeObjectURL(url);
    }, 40000);
  }

  async function pickSaveHandle(filename, fmt) {
    if (!window.showSaveFilePicker) return null;
    const isPdf = fmt === "pdf";
    try {
      return await window.showSaveFilePicker({
        suggestedName: filename,
        types: isPdf
          ? [{ description: "PDF 文档", accept: { "application/pdf": [".pdf"] } }]
          : [
              {
                description: "Word 文档",
                accept: {
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
                },
              },
            ],
      });
    } catch (err) {
      if (err && err.name === "AbortError") return "abort";
      return null;
    }
  }

  async function saveAsFile(paragraphs, filename, fmt) {
    if (!paragraphs || !paragraphs.length) {
      showToast("没有可保存的文字");
      return false;
    }
    const isPdf = fmt === "pdf";
    const name = filename.replace(/\.(docx|pdf|doc|rtf)$/i, "") + (isPdf ? ".pdf" : ".docx");
    const handle = await pickSaveHandle(name, isPdf ? "pdf" : "docx");
    if (handle === "abort") return false;
    setLoading(true, isPdf ? "正在生成 PDF…" : "正在生成 Word…");
    try {
      const blob = isPdf ? DocDiff.buildPdf(paragraphs) : await DocDiff.buildDocx(paragraphs);
      if (handle) {
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        showToast(isPdf ? "已保存为 PDF" : "已保存为 Word 文档");
      } else {
        downloadBlob(blob, name);
        showToast(isPdf ? "已下载 PDF" : "已下载 Word 文档（.docx）");
      }
      return true;
    } catch (err) {
      showToast(err.message || String(err));
      return false;
    } finally {
      setLoading(false);
    }
  }

  function preferredFmt(rec) {
    return rec && rec.ext === ".pdf" ? "pdf" : "docx";
  }

  async function saveEditor(fmt) {
    const use = fmt || preferredFmt(editSession);
    return saveAsFile(collectParas(editPane), fileName(editSession.name, "已修改", use), use);
  }

  async function saveCompareSide(side, fmt) {
    const rec = side === "right" ? session.right : session.left;
    const pane = side === "right" ? rightPane : leftPane;
    const tag = side === "right" ? "对照" : "基准";
    const use = fmt || preferredFmt(rec);
    const ok = await saveAsFile(collectParas(pane), fileName(rec && rec.name, tag + "-已修改", use), use);
    if (ok) pane.closest(".pane").classList.remove("is-dirty");
    return ok;
  }

  $("btnSaveLeft").addEventListener("click", () => saveCompareSide("left", "docx"));
  $("btnSaveLeftPdf").addEventListener("click", () => saveCompareSide("left", "pdf"));
  $("btnSaveRight").addEventListener("click", () => saveCompareSide("right", "docx"));
  $("btnSaveRightPdf").addEventListener("click", () => saveCompareSide("right", "pdf"));
  $("btnSaveLeftTop").addEventListener("click", () => saveCompareSide("left", "docx"));
  $("btnSaveLeftPdfTop").addEventListener("click", () => saveCompareSide("left", "pdf"));
  $("btnSaveRightTop").addEventListener("click", () => saveCompareSide("right", "docx"));
  $("btnSaveRightPdfTop").addEventListener("click", () => saveCompareSide("right", "pdf"));
  async function downloadManuscript() {
    if (!session.left || !session.right) {
      showToast("请先完成对比");
      return;
    }
    setLoading(true, "正在生成对比稿…");
    try {
      const payload = DocDiff.compareTexts(
        collectParas(leftPane),
        collectParas(rightPane),
        session.left.name,
        session.right.name
      );
      payload.token = "local";
      data = payload;
      renderResult(payload);
      const blob = await DocDiff.buildDocx(DocDiff.manuscriptParagraphs(payload));
      downloadBlob(blob, "对比稿.docx");
      showToast("已下载对比稿。红色删除线是删去的，蓝色下划线是新增的");
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  }

  $("btnManuscript").addEventListener("click", () => downloadManuscript());
  $("btnSaveBarManuscript").addEventListener("click", () => downloadManuscript());
  $("btnExportBoth").addEventListener("click", async () => {
    if (!data) {
      showToast("请先完成对比");
      return;
    }
    setLoading(true, "正在生成 Word 和 PDF…");
    try {
      const zip = new JSZip();
      const leftParas = collectParas(leftPane);
      const rightParas = collectParas(rightPane);
      zip.file(fileName(session.left.name, "基准", "docx"), await DocDiff.buildDocx(leftParas));
      zip.file(fileName(session.left.name, "基准", "pdf"), DocDiff.buildPdf(leftParas));
      zip.file(fileName(session.right.name, "对照", "docx"), await DocDiff.buildDocx(rightParas));
      zip.file(fileName(session.right.name, "对照", "pdf"), DocDiff.buildPdf(rightParas));
      downloadBlob(
        await zip.generateAsync({ type: "blob", mimeType: "application/zip" }),
        "文档对比-Word和PDF.zip"
      );
      showToast("已下载两侧 Word 和 PDF（压缩包）");
      leftPane.closest(".pane").classList.remove("is-dirty");
      rightPane.closest(".pane").classList.remove("is-dirty");
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  });

  async function runCompare(left, right, demo) {
    setLoading(true, "正在比对，请稍候…");
    try {
      const payload = DocDiff.compareTexts(left.paras, right.paras, left.name, right.name);
      payload.token = "local";
      session = { demo: !!demo, left, right };
      renderResult(payload);
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  }

  btnCompare.addEventListener("click", async () => {
    if (!files.base || !files.comp) return;
    setLoading(true, "正在读取文件…");
    try {
      const left = await DocDiff.extractFile(files.base);
      const right = await DocDiff.extractFile(files.comp);
      if (!left.paras.length && !right.paras.length) {
        throw new Error("没有读到文字。扫描版 PDF 无法识别，请另存为可复制文字的 Word 再比");
      }
      await runCompare(left, right, false);
    } catch (err) {
      setLoading(false);
      showToast(err.message || String(err));
    }
  });

  btnDemo.addEventListener("click", () => {
    runCompare(
      { paras: DocDiff.DEMO_LEFT, ext: ".docx", name: "示例-基准.docx", blob: null },
      { paras: DocDiff.DEMO_RIGHT, ext: ".docx", name: "示例-对照.docx", blob: null },
      true
    );
  });

  btnOpenEdit.addEventListener("click", async () => {
    const f = files.base || files.comp;
    if (!f) return;
    setLoading(true, "正在读取文件…");
    try {
      const rec = await DocDiff.extractFile(f);
      if (!rec.paras.length) {
        throw new Error("没有读到文字。扫描版 PDF 无法识别，请另存为可复制文字的 Word");
      }
      openEditor(rec);
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  });
  btnNewBlank.addEventListener("click", () => {
    openEditor({ paras: [""], ext: ".docx", name: "未命名.docx", blob: null });
  });
  $("btnEditInsert").addEventListener("click", () => insertParaIn(editPane));
  $("btnEditDelete").addEventListener("click", () => deleteParaIn(editPane));
  $("btnEditSave").addEventListener("click", () => saveEditor("docx"));
  $("btnEditSaveTop").addEventListener("click", () => saveEditor("docx"));
  $("btnEditSavePdf").addEventListener("click", () => saveEditor("pdf"));
  $("btnEditSavePdfTop").addEventListener("click", () => saveEditor("pdf"));
  $("btnSaveBarEdit").addEventListener("click", () => saveEditor("docx"));
  $("btnSaveBarEditPdf").addEventListener("click", () => saveEditor("pdf"));
  $("btnSaveBarLeft").addEventListener("click", () => saveCompareSide("left", "docx"));
  $("btnSaveBarLeftPdf").addEventListener("click", () => saveCompareSide("left", "pdf"));
  $("btnSaveBarRight").addEventListener("click", () => saveCompareSide("right", "docx"));
  $("btnSaveBarRightPdf").addEventListener("click", () => saveCompareSide("right", "pdf"));
  $("btnInsertLeft").addEventListener("click", () => insertParaIn(leftPane));
  $("btnDeleteLeft").addEventListener("click", () => deleteParaIn(leftPane));
  $("btnInsertRight").addEventListener("click", () => insertParaIn(rightPane));
  $("btnDeleteRight").addEventListener("click", () => deleteParaIn(rightPane));
  $("btnRecompare").addEventListener("click", () => recompareCurrent());
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || String(e.key).toLowerCase() !== "s") return;
    if (editPanel.hidden && resultPanel.hidden) return;
    e.preventDefault();
    if (!editPanel.hidden) saveEditor();
    else if (activeEditPane === rightPane) saveCompareSide("right");
    else saveCompareSide("left");
  });

  updateCompareEnabled();

  if (!window.DocDiff || !window.JSZip) {
    showToast("脚本未加载完整，请用 Microsoft Edge 或 Chrome 打开，并允许运行 JavaScript");
  }

  swapDocs.addEventListener("change", async () => {
    if (!data) {
      swapDocs.checked = false;
      return;
    }
    const tmp = session.left;
    session.left = session.right;
    session.right = tmp;
    await runCompare(session.left, session.right, session.demo);
  });
})();
