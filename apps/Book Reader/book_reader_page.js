/* Book Reader page script.
 *
 * Runs inside every chapter page, in the reader's own script world (the book's
 * own scripts are switched off). It lays the chapter out as pages one screen
 * wide (or two side by side, or one long scroll), keeps the place as a
 * character offset into the chapter's text so it survives any change of text
 * size or window size, and tells the window what happens through QWebChannel.
 */
(function () {
  "use strict";
  if (window.reader) { return; }
  const R = {};
  window.reader = R;

  // ------------------------------------------------------------ talking to the window
  let bridge = null;
  let generation = 0;          // which load of the window this page is; stale pages are ignored
  const outbox = [];
  function send(message) {
    message.gen = generation;
    const text = JSON.stringify(message);
    if (bridge) { bridge.event(text); } else { outbox.push(text); }
  }
  if (typeof QWebChannel !== "undefined" && window.qt && qt.webChannelTransport) {
    new QWebChannel(qt.webChannelTransport, function (channel) {
      bridge = channel.objects.bridge;
      while (outbox.length) { bridge.event(outbox.shift()); }
      // Say this page's script is here: the window starts it now if "page loaded" came too soon.
      bridge.event(JSON.stringify({ type: "hello", href: location.href, gen: -1 }));
    });
  }

  // ------------------------------------------------------------ looks
  const COLOURS = {
    white: { bg: "#ffffff", fg: "#000000", link: "#0645ad", select: "#90caf9" },
    sepia: { bg: "#f4ecd8", fg: "#000000", link: "#0645ad", select: "#90caf9" },
    grey: { bg: "#e0e0e0", fg: "#000000", link: "#0645ad", select: "#90caf9" },
    green: { bg: "#dcedc8", fg: "#000000", link: "#0645ad", select: "#90caf9" },
    night: { bg: "#121212", fg: "#e8e8e8", link: "#8ab4f8", select: "#3d5afe", dark: true },
    // warm letters on black: no bright white or blue light at night
    amber: { bg: "#121212", fg: "#ffb74d", link: "#ffd180", select: "#6d4c41", dark: true },
  };
  let style = { font: "Georgia", size: 22, line: 1.5, margin: 44, colours: "white", layout: "one", justify: true };
  let page = 0;
  let index = null;
  let lastOffset = 0;
  let anchor = 0;              // where the reader is reading; re-layouts keep it on screen
  let ready = false;

  // setTimeout never fires in this script world while the book's own scripts are
  // switched off; animation frames and promises do. So delays count frames.
  function after(ms, fn) {
    const handle = { cancelled: false };
    const until = performance.now() + ms;
    const tick = function () {
      if (handle.cancelled) { return; }
      if (performance.now() >= until) { fn(); } else { requestAnimationFrame(tick); }
    };
    requestAnimationFrame(tick);
    return handle;
  }
  function cancel(handle) { if (handle) { handle.cancelled = true; } }

  function paged() { return style.layout !== "scroll"; }
  function width() { return window.innerWidth; }
  function height() { return window.innerHeight; }
  function scroller() { return document.scrollingElement || document.documentElement; }
  function pad() { return Math.max(16, Math.round(style.margin * 0.6)); }

  function css() {
    const c = COLOURS[style.colours] || COLOURS.white;
    const w = width(), h = height(), m = style.margin, p = pad();
    let layout;
    if (paged()) {
      const two = style.layout === "two" && w >= 700;
      const gap = 2 * m;
      const column = two ? Math.floor(w / 2 - gap) : (w - gap);
      layout = `html { overflow: hidden !important; height: ${h}px !important; width: ${w}px !important;
                       margin: 0 !important; padding: 0 !important; }
        body { margin: 0 !important; padding: ${p}px ${m}px !important; box-sizing: border-box !important;
               height: ${h}px !important; width: ${w}px !important; max-width: none !important;
               min-width: 0 !important; column-width: ${column}px !important; column-gap: ${gap}px !important;
               column-fill: auto !important; column-count: auto !important; overflow: visible !important; }
        img, svg, video, object { max-height: ${h - 2 * p - 4}px !important; }`;
    } else {
      layout = `html { overflow-x: hidden !important; overflow-y: auto !important; height: auto !important; }
        body { margin: 0 auto !important; padding: ${p}px ${m}px ${Math.round(h / 2)}px !important;
               box-sizing: border-box !important; columns: auto !important; height: auto !important;
               max-width: ${Math.max(640, Math.min(w, 40 * style.size))}px !important; }`;
    }
    const font = style.font && style.font !== "The book's own"
      ? `body, body *:not(pre):not(code):not(kbd):not(samp):not(tt) { font-family: "${style.font}", serif !important; }`
      : "";
    return layout + `
      html, body { background: ${c.bg} !important; }
      body, body * { color: ${c.fg} !important; }
      body * { background-color: transparent !important; }
      body a, body a * { color: ${c.link} !important; }
      html { font-size: ${style.size}px !important; }
      body { font-size: ${style.size}px !important; line-height: ${style.line} !important; }
      body p, body li, body div, body blockquote, body dd, body dt, body td, body span {
        line-height: ${style.line} !important; }
      body p { text-align: ${style.justify ? "justify" : "left"} !important; hyphens: auto; }
      img, svg, video, object { max-width: 100% !important; height: auto; object-fit: contain; break-inside: avoid; }
      figure, table, pre { break-inside: avoid; }
      ::selection { background: ${c.select}; color: ${c.fg}; }
      ${font}
      ${extraCss(c)}`;
  }
  let extraCss = function () { return ""; };

  function applyStyle() {
    let el = document.getElementById("book-reader-style");
    if (!el) {
      el = document.createElement("style");
      el.id = "book-reader-style";
    }
    document.documentElement.appendChild(el);      // last in the page, so it wins
    el.textContent = css();
  }

  // ------------------------------------------------------------ text offsets
  function buildIndex() {
    const nodes = [], starts = [], map = new Map();
    let total = 0;
    if (document.body) {
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
        acceptNode(n) {
          const parent = n.parentNode && n.parentNode.nodeName.toLowerCase();
          return parent === "script" || parent === "style" || parent === "noscript"
            ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT;
        },
      });
      while (walker.nextNode()) {
        const n = walker.currentNode;
        map.set(n, nodes.length);
        nodes.push(n);
        starts.push(total);
        total += n.data.length;
      }
    }
    index = { nodes, starts, map, total, text: null };
  }

  function fullText() {
    if (index.text === null) { index.text = index.nodes.map((n) => n.data).join(""); }
    return index.text;
  }

  function locate(off) {
    const s = index.starts;
    if (!s.length) { return null; }
    off = Math.max(0, Math.min(off, index.total));
    let lo = 0, hi = s.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (s[mid] <= off) { lo = mid; } else { hi = mid - 1; }
    }
    let i = lo;
    while (i < s.length - 1 && off - s[i] >= index.nodes[i].data.length) { i++; }
    return { node: index.nodes[i], offset: Math.min(off - s[i], index.nodes[i].data.length) };
  }

  function boundaryOffset(container, off) {
    if (container.nodeType === Node.TEXT_NODE && index.map.has(container)) {
      return index.starts[index.map.get(container)] + off;
    }
    const nodes = index.nodes;
    let lo = 0, hi = nodes.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      const r = document.createRange();
      r.selectNodeContents(nodes[mid]);
      if (r.comparePoint(container, off) > 0) { lo = mid + 1; } else { hi = mid; }
    }
    return lo < nodes.length ? index.starts[lo] : index.total;
  }

  function rectAt(off) {
    const loc = locate(off);
    if (!loc) { return null; }
    const len = loc.node.data.length;
    if (!len) { return null; }
    const o = Math.min(loc.offset, len - 1);
    const r = document.createRange();
    r.setStart(loc.node, o);
    r.setEnd(loc.node, o + 1);
    for (const rect of r.getClientRects()) {
      if (rect.width > 0 || rect.height > 0) { return rect; }
    }
    return null;
  }

  function measurable(off, step) {
    const text = fullText();
    for (let k = 0; k < 400; k++) {
      const o = off + step * k;
      if (o < 0 || o >= index.total) { break; }
      const code = text.charCodeAt(o);
      if (code <= 32 || code === 160) { continue; }      // a space can sit at the end of the line before
      const rect = rectAt(o);
      if (rect) { return rect; }
    }
    return null;
  }

  function boxOf(off) { return measurable(off, 1) || measurable(off, -1); }

  function pageOf(off) {
    const rect = boxOf(off);
    if (!rect) { return 0; }
    if (paged()) { return Math.floor((rect.left + rect.width / 2 + scroller().scrollLeft) / width()); }
    return Math.floor((rect.top + scroller().scrollTop) / height());
  }

  function pageCount() {
    const el = scroller();
    if (paged()) { return Math.max(1, Math.ceil((el.scrollWidth - 4) / width())); }
    return Math.max(1, Math.ceil((el.scrollHeight - 1) / height()));
  }

  function firstVisible() {
    if (!index || !index.total) { return 0; }
    let lo = 0, hi = index.total - 1;
    if (paged()) {
      if (pageOf(hi) < page) { return hi; }
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (pageOf(mid) >= page) { hi = mid; } else { lo = mid + 1; }
      }
      return lo;
    }
    const below = function (o) { const r = boxOf(o); return r ? r.bottom > 1 : true; };
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (below(mid)) { hi = mid; } else { lo = mid + 1; }
    }
    return lo;
  }

  // ------------------------------------------------------------ moving
  function report() {
    if (!index) { return; }
    const pages = pageCount();
    if (!paged()) { page = Math.min(pages - 1, Math.floor(scroller().scrollTop / height() + 0.01)); }
    lastOffset = firstVisible();
    send({ type: "place", offset: lastOffset, page: page, pages: pages, total: index.total,
           chars: Math.round(index.total / pages) });
  }

  function goPage(n, quiet, keepAnchor) {
    const el = scroller();
    if (paged()) {
      page = Math.max(0, Math.min(n, pageCount() - 1));
      el.scrollTop = 0;
      el.scrollLeft = page * width();
    } else {
      el.scrollTop = Math.max(0, n) * height();
    }
    if (!quiet) { report(); }
    if (!keepAnchor) { anchor = firstVisible(); }
  }

  function goToOffset(off, quiet) {
    anchor = off;
    if (paged()) { goPage(pageOf(off), quiet, true); return; }
    const rect = boxOf(off);
    if (rect) { scroller().scrollTop += rect.top - pad(); }
    if (!quiet) { report(); }
  }

  function lineHeight() { return style.size * style.line; }

  R.next = function () {
    if (!ready) { return; }
    if (paged()) {
      if (page < pageCount() - 1) { goPage(page + 1); } else { send({ type: "edge", dir: "next" }); }
      return;
    }
    const el = scroller();
    if (el.scrollTop + height() >= el.scrollHeight - height() / 2 - 2) { send({ type: "edge", dir: "next" }); return; }
    el.scrollTop += height() - 2 * lineHeight();
    report();
    anchor = lastOffset;
  };

  R.prev = function () {
    if (!ready) { return; }
    if (paged()) {
      if (page > 0) { goPage(page - 1); } else { send({ type: "edge", dir: "prev" }); }
      return;
    }
    const el = scroller();
    if (el.scrollTop <= 1) { send({ type: "edge", dir: "prev" }); return; }
    el.scrollTop -= height() - 2 * lineHeight();
    report();
    anchor = lastOffset;
  };

  R.goToEnd = function () {
    if (paged()) { goPage(pageCount() - 1); return; }
    scroller().scrollTop = scroller().scrollHeight;
    report();
    anchor = lastOffset;
  };
  R.goToStart = function () { goPage(0); };
  R.goToOffset = function (off) { goToOffset(off); };
  R.goToAnchor = function (id) { place({ anchor: id }); report(); };
  R.isVisible = function (off) {
    if (paged()) { return pageOf(off) === page; }
    const rect = boxOf(off);
    return !!rect && rect.top >= -1 && rect.top < height();
  };
  R.info = function () {
    return { page: page, pages: pageCount(), offset: lastOffset, total: index ? index.total : 0 };
  };
  R.textAt = function (start, end) { return fullText().slice(start, end); };
  R.pageRange = function () {
    if (!index || !index.total) { return [0, 0]; }
    const start = lastOffset;
    let lo = start, hi = index.total;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      let after;
      if (paged()) { after = pageOf(mid) > page; } else { const r = boxOf(mid); after = !!r && r.top >= height(); }
      if (after) { hi = mid; } else { lo = mid + 1; }
    }
    return [start, lo];
  };
  R.pageInfo = function () {
    const range = R.pageRange();
    return { start: range[0], end: range[1], text: fullText().slice(range[0], Math.min(range[1], range[0] + 160)) };
  };

  function offsetOfElement(el) { return boundaryOffset(el, 0); }

  function place(target) {
    if (target.end) { R.goToEnd(); return; }
    if (target.anchor) {
      const el = document.getElementById(target.anchor) || document.getElementsByName(target.anchor)[0];
      if (el) { goToOffset(offsetOfElement(el), true); return; }
    }
    if (placeHooks.some((hook) => hook(target))) { return; }
    goToOffset(target.offset || 0, true);
  }
  const placeHooks = [];

  function whenSettled(done) {
    let finished = false;
    const finish = function () {
      if (finished) { return; }
      finished = true;
      requestAnimationFrame(() => requestAnimationFrame(done));
    };
    const fonts = document.fonts ? document.fonts.ready : Promise.resolve();
    fonts.then(function () {
      const waiting = Array.from(document.images).filter((i) => !i.complete);
      let left = waiting.length;
      if (!left) { finish(); return; }
      waiting.forEach(function (img) {
        const one = function () { if (--left <= 0) { finish(); } };
        img.addEventListener("load", one, { once: true });
        img.addEventListener("error", one, { once: true });
      });
      after(1500, finish);
    });
  }

  let startedFor = null;
  R.start = function (newStyle, target, gen) {
    if (startedFor === gen) { return; }       // asked twice for the same load: once is enough
    startedFor = gen;
    generation = gen || 0;
    style = Object.assign(style, newStyle || {});
    applyStyle();
    whenSettled(function () {
      applyStyle();                  // again: the window may have changed size while it loaded
      buildIndex();
      startHooks.forEach((hook) => hook());
      place(target || {});
      ready = true;
      send({ type: "ready", total: index.total });
      report();
    });
  };
  const startHooks = [];

  R.setStyle = function (newStyle) {
    const keep = anchor;
    style = Object.assign(style, newStyle || {});
    applyStyle();
    whenSettled(function () {
      goToOffset(keep, true);
      send({ type: "ready", total: index ? index.total : 0 });
      report();
    });
  };

  let resizeTimer = null, resizeKeep = 0;
  window.addEventListener("resize", function () {
    if (!ready) { return; }
    if (!resizeTimer) {
      resizeKeep = anchor;
      send({ type: "busy" });        // the window waits until the pages are laid out again
    }
    cancel(resizeTimer);
    resizeTimer = after(120, function () {
      resizeTimer = null;
      applyStyle();
      requestAnimationFrame(function () {
        goToOffset(resizeKeep);
        send({ type: "ready", total: index ? index.total : 0 });
      });
    });
  });

  let reportTimer = null;
  function reportSoon() {
    cancel(reportTimer);
    reportTimer = after(150, report);
  }
  window.addEventListener("scroll", function () {
    if (!ready || resizeTimer) { return; }
    if (!paged()) { reportSoon(); return; }
    const el = scroller(), want = page * width();
    if (Math.abs(el.scrollLeft - want) > 1 || el.scrollTop) {
      // Something else scrolled (a jump to an anchor, dragging a selection): snap to a page.
      page = Math.max(0, Math.min(Math.round(el.scrollLeft / width()), pageCount() - 1));
      el.scrollTop = 0;
      el.scrollLeft = page * width();
      anchor = firstVisible();
      reportSoon();
    }
  }, true);

  // ------------------------------------------------------------ clicks, keys, wheel
  function absolute(src) {
    try { return new URL(src, document.baseURI).href; } catch (e) { return ""; }
  }

  const clickHooks = [];
  document.addEventListener("click", function (e) {
    if (!ready || e.button !== 0) { return; }
    const selection = window.getSelection();
    if (selection && !selection.isCollapsed && String(selection).trim()) { return; }
    const link = e.target.closest && e.target.closest("a[href], area[href]");
    if (link) {
      e.preventDefault();
      send({ type: "link", href: link.href && link.href.baseVal !== undefined ? absolute(link.href.baseVal) : link.href });
      return;
    }
    const picture = e.target.closest && e.target.closest("img, svg");
    if (picture) {
      let src = picture.currentSrc || picture.src || "";
      if (!src && picture.querySelector) {
        const inner = picture.querySelector("image");
        if (inner) { src = inner.getAttribute("href") || inner.getAttributeNS("http://www.w3.org/1999/xlink", "href") || ""; }
      }
      if (src) { send({ type: "image", src: absolute(src) }); return; }
    }
    if (clickHooks.some((hook) => hook(e))) { return; }
    const w = width();
    if (e.clientX < w * 0.2) { R.prev(); } else if (e.clientX > w * 0.8) { R.next(); } else {
      send({ type: "tap", x: e.clientX, y: e.clientY });
    }
  }, true);

  document.addEventListener("mouseup", function (e) {
    if (!ready || e.button !== 0) { return; }
    requestAnimationFrame(function () {
      const selection = window.getSelection();
      if (!selection || selection.isCollapsed || !selection.rangeCount) { return; }
      const text = String(selection);
      if (!text.trim()) { return; }
      const range = selection.getRangeAt(0);
      const rect = range.getBoundingClientRect();
      send({ type: "selected", start: boundaryOffset(range.startContainer, range.startOffset),
             end: boundaryOffset(range.endContainer, range.endOffset), text: text,
             x: rect.left + rect.width / 2, y: rect.bottom, top: rect.top });
    });
  });

  let wheelAt = 0;
  document.addEventListener("wheel", function (e) {
    if (!paged() || e.ctrlKey) { return; }
    e.preventDefault();
    const now = Date.now();
    if (now - wheelAt < 350) { return; }
    wheelAt = now;
    if (e.deltaY > 0 || e.deltaX > 0) { R.next(); } else if (e.deltaY < 0 || e.deltaX < 0) { R.prev(); }
  }, { passive: false, capture: true });

  document.addEventListener("keydown", function (e) {
    if (e.ctrlKey || e.altKey || e.metaKey || !ready) { return; }
    const k = e.key;
    if (k === "ArrowRight" || k === "PageDown" || (k === " " && !e.shiftKey) || (paged() && k === "ArrowDown")) {
      e.preventDefault(); R.next();
    } else if (k === "ArrowLeft" || k === "PageUp" || (k === " " && e.shiftKey) || (paged() && k === "ArrowUp")) {
      e.preventDefault(); R.prev();
    } else if (k === "Home") {
      e.preventDefault(); R.goToStart();
    } else if (k === "End") {
      e.preventDefault(); R.goToEnd();
    }
  }, true);

  // ------------------------------------------------------------ for the other parts of this script
  R._internal = {
    send, after, cancel, locate, boundaryOffset, fullText, goToOffset, report, boxOf, paged, height,
    get index() { return index; }, get ready() { return ready; },
    addCss(fn) { const old = extraCss; extraCss = (c) => old(c) + fn(c); },
    onStart(fn) { startHooks.push(fn); }, onPlace(fn) { placeHooks.push(fn); }, onClick(fn) { clickHooks.push(fn); },
  };
})();


/* Highlights, search marks and the sentence being read aloud. They use the
 * browser's highlight painting (CSS.highlights), which never changes the
 * page's text, so places stay exact. */
(function () {
  "use strict";
  const R = window.reader;
  if (!R || R.setHighlights) { return; }
  const I = R._internal;
  const KINDS = ["yellow", "green", "blue", "pink"];
  const DAY = { yellow: "#fff176", green: "#b9f6ca", blue: "#b3e5fc", pink: "#f8bbd0" };
  const NIGHT = { yellow: "#6d5d00", green: "#1b5e20", blue: "#0d47a1", pink: "#880e4f" };
  let marks = [];
  let found = -1;

  I.addCss(function (c) {
    const night = !!c.dark;
    const set = night ? NIGHT : DAY;
    return KINDS.map((k) => `::highlight(br-${k}) { background-color: ${set[k]}; color: ${c.fg}; }`).join("\n")
      + "\n::highlight(br-find) { background-color: #ff9800; color: #000000; }"
      + `\n::highlight(br-speak) { background-color: ${night ? "#00695c" : "#b2ebf2"}; color: ${c.fg}; }`;
  });

  function rangeFor(start, end) {
    if (!I.index) { return null; }
    const a = I.locate(start), b = I.locate(end);
    if (!a || !b) { return null; }
    const r = document.createRange();
    try { r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); } catch (e) { return null; }
    return r;
  }
  I.rangeFor = rangeFor;

  function paint() {
    if (!window.CSS || !CSS.highlights || !I.index) { return; }
    for (const kind of KINDS) {
      const ranges = marks.filter((m) => m.colour === kind).map((m) => rangeFor(m.start, m.end)).filter(Boolean);
      if (ranges.length) { CSS.highlights.set("br-" + kind, new Highlight(...ranges)); } else { CSS.highlights.delete("br-" + kind); }
    }
  }
  R.setHighlights = function (list) { marks = Array.isArray(list) ? list : []; paint(); };
  I.onStart(paint);

  function offsetAtPoint(x, y) {
    let node = null, off = 0;
    if (document.caretPositionFromPoint) {
      const p = document.caretPositionFromPoint(x, y);
      if (p) { node = p.offsetNode; off = p.offset; }
    } else if (document.caretRangeFromPoint) {
      const r = document.caretRangeFromPoint(x, y);
      if (r) { node = r.startContainer; off = r.startOffset; }
    }
    return node ? I.boundaryOffset(node, off) : -1;
  }

  I.onClick(function (e) {
    if (!marks.length) { return false; }
    const off = offsetAtPoint(e.clientX, e.clientY);
    if (off < 0) { return false; }
    const hit = marks.find((m) => m.start <= off && off <= m.end);
    if (!hit) { return false; }
    const r = rangeFor(hit.start, hit.end);
    const on = r && Array.from(r.getClientRects()).some((b) =>
      e.clientX >= b.left - 2 && e.clientX <= b.right + 2 && e.clientY >= b.top - 2 && e.clientY <= b.bottom + 2);
    if (!on) { return false; }
    I.send({ type: "highlight", id: hit.id, x: e.clientX, y: e.clientY });
    return true;
  });

  function findOffset(query, n) {
    const text = I.fullText().toLowerCase(), q = String(query).toLowerCase();
    if (!q) { return -1; }
    let at = -1;
    for (let k = 0; k <= n; k++) {
      at = text.indexOf(q, at + 1);
      if (at < 0) { return -1; }
    }
    return at;
  }
  function markFound(query, n) {
    const at = findOffset(query, n);
    found = at;
    if (at >= 0 && window.CSS && CSS.highlights) {
      const r = rangeFor(at, at + String(query).length);
      if (r) { CSS.highlights.set("br-find", new Highlight(r)); }
    }
    return at;
  }
  I.onPlace(function (target) {
    if (!target.find) { return false; }
    const at = markFound(target.find[0], target.find[1] || 0);
    if (at < 0) { return false; }
    I.goToOffset(at, true);
    return true;
  });
  R.find = function (query, n) {
    const at = markFound(query, n);
    if (at >= 0) { I.goToOffset(at); }
    return at;
  };
  R.clearFind = function () {
    found = -1;
    if (window.CSS && CSS.highlights) { CSS.highlights.delete("br-find"); }
  };
  R.foundAt = function () { return found; };
  R.clearSelection = function () { const s = window.getSelection(); if (s) { s.removeAllRanges(); } };

  // ---- reading aloud: sentences end at . ! ? (and quotes after them) or at the edge of a paragraph
  const BLOCKS = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "TD", "TH",
    "DT", "DD", "PRE", "SECTION", "ARTICLE", "HEADER", "FOOTER", "FIGCAPTION", "TR", "TABLE", "ASIDE", "NAV",
    "CAPTION", "BODY", "UL", "OL", "DL", "FIGURE", "ADDRESS", "CENTER"]);
  function blockOf(node) {
    for (let el = node.parentNode; el && el.nodeType === 1; el = el.parentNode) {
      if (BLOCKS.has(el.nodeName.toUpperCase())) { return el; }
    }
    return null;
  }
  function blockBreaks() {
    const idx = I.index;
    if (idx.breaks) { return idx.breaks; }
    const breaks = [];
    let previous = null;
    for (let i = 0; i < idx.nodes.length; i++) {
      const block = blockOf(idx.nodes[i]);
      if (i && block !== previous) { breaks.push(idx.starts[i]); }
      previous = block;
    }
    idx.breaks = breaks;
    return breaks;
  }
  R.sentencesFrom = function (off) {
    if (!I.index) { return []; }
    const text = I.fullText(), total = text.length, out = [];
    let pos = Math.max(0, Math.min(off | 0, total));
    const breaks = blockBreaks();
    let b = 0;
    while (b < breaks.length && breaks[b] <= pos) { b++; }
    const ends = /[.!?…]+["'”’)\]]*(?=\s|$)/g;
    let pending = null;
    const push = function (a, z) {
      const piece = text.slice(a, z), t = piece.trim();
      if (t && /[\p{L}\p{N}]/u.test(t)) {
        const lead = piece.length - piece.trimStart().length;
        out.push({ start: a + lead, end: a + lead + t.length, text: t.replace(/\s+/g, " ") });
      }
    };
    while (pos < total && out.length < 3000) {
      const edge = b < breaks.length ? breaks[b] : total;
      if (!pending || pending.index < pos) {
        ends.lastIndex = pos;
        pending = ends.exec(text) || { index: total, 0: "" };
      }
      let end = Math.min(pending.index + pending[0].length, edge);
      if (end - pos > 400) {                      // very long: break at a space
        const space = text.lastIndexOf(" ", pos + 300);
        end = space > pos ? space : pos + 300;
      }
      if (end <= pos) { end = Math.min(total, edge > pos ? edge : pos + 1); }
      push(pos, end);
      pos = end;
      while (b < breaks.length && breaks[b] <= pos) { b++; }
    }
    return out;
  };
  R.markSentence = function (start, end) {
    if (window.CSS && CSS.highlights) {
      const r = rangeFor(start, end);
      if (r) { CSS.highlights.set("br-speak", new Highlight(r)); }
    }
    if (!R.isVisible(start)) { I.goToOffset(start); }
  };
  R.clearSpeech = function () { if (window.CSS && CSS.highlights) { CSS.highlights.delete("br-speak"); } };
})();
