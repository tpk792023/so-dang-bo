/* Xử lý bảng danh sách học sinh trong file Word (.docx).
   Cần: JSZip, DOMParser, XMLSerializer và SDB (core.js). */
(function (root) {
  "use strict";
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  const g = name => root[name] || (typeof globalThis !== "undefined" ? globalThis[name] : undefined);

  // ---------- Tiện ích XML ----------
  const isEl = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name);
  function kids(el, name) {
    const out = [];
    for (let n = el.firstChild; n; n = n.nextSibling) if (isEl(n, name)) out.push(n);
    return out;
  }
  const child = (el, name) => kids(el, name)[0] || null;
  const desc = (el, name) => Array.from(el.getElementsByTagNameNS(W, name));
  const getA = (el, name) => (el ? (el.getAttributeNS(W, name) || el.getAttribute("w:" + name) || null) : null);
  const setA = (el, name, v) => el.setAttributeNS(W, "w:" + name, String(v));

  // Thứ tự con của w:tcPr theo chuẩn OOXML
  const TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"];

  function tcPr(tc) {
    let p = child(tc, "tcPr");
    if (!p) { p = tc.ownerDocument.createElementNS(W, "w:tcPr"); tc.insertBefore(p, tc.firstChild); }
    return p;
  }
  function setTcProp(tc, name, val) {
    const pr = tcPr(tc);
    let el = child(pr, name);
    if (!el) {
      el = tc.ownerDocument.createElementNS(W, "w:" + name);
      const rank = TCPR_ORDER.indexOf(name);
      let before = null;
      for (let n = pr.firstChild; n; n = n.nextSibling)
        if (isEl(n) && TCPR_ORDER.indexOf(n.localName) > rank) { before = n; break; }
      pr.insertBefore(el, before);
    }
    if (val != null) setA(el, "val", val); else el.removeAttributeNS(W, "val");
    return el;
  }
  function removeTcProp(tc, name) {
    const pr = child(tc, "tcPr"); const el = pr && child(pr, name);
    if (el) pr.removeChild(el);
  }

  function cellText(tc) {
    return kids(tc, "p").map(p => {
      let s = "";
      for (const n of desc(p, "*")) {
        if (n.localName === "t") s += n.textContent;
        else if (n.localName === "tab") s += " ";
      }
      return s;
    }).join(" ").replace(/\s+/g, " ").trim();
  }

  function setCellText(tc, text) {
    const doc = tc.ownerDocument;
    let ps = kids(tc, "p");
    if (!ps.length) { const np = doc.createElementNS(W, "w:p"); tc.appendChild(np); ps = [np]; }
    const p = ps[0];
    ps.slice(1).forEach(x => tc.removeChild(x));
    const pPr = child(p, "pPr");
    let rPr = null;
    const firstRun = desc(p, "r").find(r => desc(r, "t").length) || desc(p, "r")[0];
    if (firstRun && child(firstRun, "rPr")) rPr = child(firstRun, "rPr").cloneNode(true);
    else if (pPr && child(pPr, "rPr")) rPr = child(pPr, "rPr").cloneNode(true);
    for (let n = p.firstChild; n;) { const nx = n.nextSibling; if (n !== pPr) p.removeChild(n); n = nx; }
    const r = doc.createElementNS(W, "w:r");
    if (rPr) r.appendChild(rPr);
    if (text) {
      const t = doc.createElementNS(W, "w:t");
      t.setAttributeNS(XML_NS, "xml:space", "preserve");
      t.appendChild(doc.createTextNode(text));
      r.appendChild(t);
    }
    p.appendChild(r);
  }

  function rowsOf(tbl) { return kids(tbl, "tr"); }

  // Các ô của một dòng kèm vị trí trên lưới (0-based)
  function rowCells(tr) {
    const trPr = child(tr, "trPr");
    let pos = +(getA(trPr && child(trPr, "gridBefore"), "val") || 0);
    const out = [];
    const take = tc => {
      const pr = child(tc, "tcPr");
      const span = +(getA(pr && child(pr, "gridSpan"), "val") || 1);
      const vmEl = pr && child(pr, "vMerge");
      const vm = vmEl ? (getA(vmEl, "val") || "continue") : null;
      out.push({ tc, start: pos, span, vm });
      pos += span;
    };
    for (let n = tr.firstChild; n; n = n.nextSibling) {
      if (isEl(n, "tc")) take(n);
      else if (isEl(n, "sdt")) { const c = child(n, "sdtContent"); c && kids(c, "tc").forEach(take); }
    }
    return out;
  }
  const cellAt = (cells, gi) => cells.find(c => c.start <= gi && gi < c.start + c.span) || null;
  const cellStarting = (cells, gi) => cells.find(c => c.start === gi && c.span === 1) || null;

  function tcWidth(tc, w) {
    const pr = child(tc, "tcPr"), el = pr && child(pr, "tcW");
    if (el && (getA(el, "type") || "dxa") === "dxa") setA(el, "w", Math.round(w));
  }

  // ---------- Đọc ----------
  async function loadDocx(buf) {
    const JSZip = g("JSZip");
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file("word/document.xml");
    if (!f) throw new Error("File Word không hợp lệ (thiếu word/document.xml).");
    const xml = await f.async("string");
    const doc = new (g("DOMParser"))().parseFromString(xml, "application/xml");
    return { zip, doc };
  }

  function tableTitle(tbl, k) {
    let n = tbl.previousSibling, steps = 0;
    while (n && steps < 8) {
      if (isEl(n, "p")) {
        steps++;
        const t = SDB().plain(paraText(n));
        const m = t.match(/\blop\s*:?\s*(\d{1,2}\s*[a-z]?\s*\d{0,2})\b/);
        if (m) return m[1].replace(/\s+/g, "").toUpperCase();
      } else if (isEl(n, "tbl")) break;
      n = n.previousSibling;
    }
    return "Bảng " + k;
  }
  function paraText(p) { return desc(p, "t").map(t => t.textContent).join(""); }

  function SDB() { return g("SDB"); }

  function readClassesDocx(pkg) {
    const S = SDB();
    const body = desc(pkg.doc, "body")[0];
    const tables = desc(body, "tbl").filter(t => { for (let a = t.parentNode; a; a = a.parentNode) if (isEl(a, "tbl")) return false; return true; });
    const classes = [];
    let k = 0;
    tables.forEach(tbl => {
      const trs = rowsOf(tbl);
      const cellsByRow = trs.map(rowCells);
      const maxG = Math.max(0, ...cellsByRow.map(cs => cs.length ? cs[cs.length - 1].start + cs[cs.length - 1].span : 0));
      // lưới 1-based cho hàm nhận diện tiêu đề
      const get = (r, c) => {
        const cs = cellsByRow[r - 1]; if (!cs) return "";
        const cell = cs.find(x => x.start === c - 1);
        return cell && cell.vm !== "continue" ? S.plain(cellText(cell.tc)) : "";
      };
      const h = S.locateNameHeader(get, Math.min(trs.length, 8), maxG);
      if (!h) return;
      k++;
      // chữ tiêu đề (kể cả ô gộp ngang/dọc) của cột c qua các dòng tiêu đề
      const hdrRows = [];
      for (let r = h.top; r <= h.row; r++) hdrRows.push(r);
      const nb = cellsByRow[h.row] && cellAt(cellsByRow[h.row], h.nameCol - 1);
      if (nb && nb.vm === "continue") hdrRows.push(h.row + 1); // tiêu đề 2 tầng
      const coverText = (r, c) => {
        const cs = cellsByRow[r - 1]; if (!cs) return "";
        let cell = cellAt(cs, c - 1);
        if (!cell) return "";
        if (cell.vm === "continue") { // lấy ô gốc phía trên
          for (let rr = r - 2; rr >= 0; rr--) {
            const up = cellAt(cellsByRow[rr] || [], c - 1);
            if (up && up.vm !== "continue") { cell = up; break; }
          }
        }
        return S.plain(cellText(cell.tc));
      };
      const cols = S.assignColumns(c => S.joinHeaderTexts(hdrRows.map(r => coverText(r, c))), maxG, h);
      const lastHdr = hdrRows[hdrRows.length - 1];
      const txtAt = (ri, c) => { const cell = c ? cellAt(cellsByRow[ri], c - 1) : null; return cell && cell.start === c - 1 ? cellText(cell.tc) : ""; };
      const nameAt = ri => cols.ten ? S.joinName(txtAt(ri, cols.name), txtAt(ri, cols.ten)) : txtAt(ri, cols.name);
      let ri = lastHdr; // dòng ngay sau tiêu đề (0-based)
      for (let s = 0; s < 3 && ri < trs.length; s++) {
        const nm = nameAt(ri), st = txtAt(ri, cols.stt);
        if ((!nm && !S.isNum(st)) || /^\(?\d+\)?$/.test(nm)) ri++; else break;
      }
      const rows = [];
      for (; ri < trs.length; ri++) {
        const st = txtAt(ri, cols.stt), name = nameAt(ri);
        if (st && !S.isNum(st)) break;
        rows.push({
          tr: trs[ri], name,
          gt: txtAt(ri, cols.gt), dt: txtAt(ri, cols.dt), ns: cols.ns ? S.fmtDate(txtAt(ri, cols.ns)) : "",
          noisinh: txtAt(ri, cols.noisinh), cho: txtAt(ri, cols.cho), chame: txtAt(ri, cols.chame), note: txtAt(ri, cols.note),
          cha: txtAt(ri, cols.cha), ngheCha: txtAt(ri, cols.ngheCha), me: txtAt(ri, cols.me), ngheMe: txtAt(ri, cols.ngheMe), lop: txtAt(ri, cols.lop),
          _stt: st,
        });
      }
      while (rows.length && !rows[rows.length - 1].name && !rows[rows.length - 1]._stt) rows.pop(); // dòng trống cuối bảng
      classes.push(S.finishClass({
        kind: "word", tbl, name: tableTitle(tbl, k), sheet: "bang" + k, header: h.row, headerTop: h.top, cols, rows, footer: null,
      }));
    });
    return S.markCrossDuplicates(classes);
  }

  // ---------- Xếp ABC ----------
  function sortDocx(classes, opts) {
    opts = opts || {};
    const S = SDB();
    for (const cl of classes) {
      if (!cl.rows.length) continue;
      const tbl = cl.tbl;
      const anchor = cl.rows[cl.rows.length - 1].tr.nextSibling;
      const named = cl.rows.filter(x => x.name).sort((a, b) => S.compareNames(a.name, b.name));
      const blank = cl.rows.filter(x => !x.name);
      cl.rows.forEach(x => tbl.removeChild(x.tr));
      const order = opts.dropBlank ? named : named.concat(blank);
      order.forEach((x, i) => {
        tbl.insertBefore(x.tr, anchor);
        const c = cellAt(rowCells(x.tr), cl.cols.stt - 1);
        if (c && c.start === cl.cols.stt - 1) setCellText(c.tc, String(i + 1));
      });
    }
  }

  // ---------- Tách / gộp cột họ tên ----------
  function gridCols(tbl) { const gr = child(tbl, "tblGrid"); return gr ? kids(gr, "gridCol") : []; }

  function splitDocx(cl) {
    if (cl.split) return false;
    const S = SDB(), tbl = cl.tbl, g0 = cl.cols.name - 1;
    const doc = tbl.ownerDocument;
    const grid = gridCols(tbl);
    let w1 = null, w2 = null;
    if (grid[g0]) {
      const w = +(getA(grid[g0], "w") || 0);
      if (w) {
        // cột họ ≥ ~4,2 cm, cột tên ≥ ~1,8 cm; thiếu thì lấy bớt từ các cột khác
        w2 = Math.max(Math.round(w * 0.32), 1000);
        w1 = Math.max(w - w2, 2400);
        let extra = w1 + w2 - w;
        if (extra > 0) {
          const MIN = 900;
          const others = grid.map((gc, i) => ({ gc, i, w: +(getA(gc, "w") || 0) })).filter(o => o.i !== g0 && o.w > MIN);
          const avail = others.reduce((s, o) => s + o.w - MIN, 0);
          const take = Math.min(extra, avail);
          others.forEach(o => setA(o.gc, "w", Math.round(o.w - take * (o.w - MIN) / avail)));
          w1 -= extra - take;
        }
        setA(grid[g0], "w", w1);
        const ng = doc.createElementNS(W, "w:gridCol"); setA(ng, "w", w2);
        grid[g0].parentNode.insertBefore(ng, grid[g0].nextSibling);
      } else {
        grid[g0].parentNode.insertBefore(grid[g0].cloneNode(true), grid[g0].nextSibling);
      }
    }
    const data = new Map(cl.rows.map(x => [x.tr, x]));
    const headerTr = rowsOf(tbl)[cl.header - 1];
    for (const tr of rowsOf(tbl)) {
      const cells = rowCells(tr), c = cellAt(cells, g0);
      if (!c) continue;
      if (c.start === g0 && c.span === 1) {
        const cp = c.tc.cloneNode(true);
        c.tc.parentNode.insertBefore(cp, c.tc.nextSibling);
        if (w1) { tcWidth(c.tc, w1); tcWidth(cp, w2); }
        if (tr === headerTr) { setCellText(c.tc, "Họ và chữ lót"); setCellText(cp, "Tên"); }
        else if (data.has(tr)) { const [ho, ten] = S.splitName(data.get(tr).name); setCellText(c.tc, ho); setCellText(cp, ten); }
        else setCellText(cp, "");
      } else {
        setTcProp(c.tc, "gridSpan", c.span + 1);
      }
    }
    syncWidths(tbl);
    return true;
  }

  // Đặt lại độ rộng từng ô theo lưới cột của bảng
  function syncWidths(tbl) {
    const ws = gridCols(tbl).map(gc => +(getA(gc, "w") || 0));
    if (!ws.length || ws.some(x => !x)) return;
    for (const tr of rowsOf(tbl))
      for (const c of rowCells(tr)) {
        let sum = 0;
        for (let i = c.start; i < c.start + c.span; i++) sum += ws[i] || 0;
        if (sum) tcWidth(c.tc, sum);
      }
  }

  function mergeDocx(cl) {
    if (!cl.split) return false;
    const tbl = cl.tbl, g0 = cl.cols.name - 1, g1 = cl.cols.ten - 1;
    const grid = gridCols(tbl);
    let wsum = null;
    if (grid[g0] && grid[g1]) {
      wsum = (+(getA(grid[g0], "w") || 0)) + (+(getA(grid[g1], "w") || 0));
      if (wsum) setA(grid[g0], "w", wsum);
      grid[g1].parentNode.removeChild(grid[g1]);
    }
    const data = new Map(cl.rows.map(x => [x.tr, x]));
    const trs = rowsOf(tbl);
    const headerTr = trs[cl.header - 1], topTr = trs[cl.headerTop - 1];
    const sub = headerTr !== topTr;
    for (const tr of trs) {
      const cells = rowCells(tr);
      const cover = cells.find(c => c.start <= g0 && c.start + c.span - 1 >= g1);
      if (cover) {
        if (cover.span - 1 <= 1) removeTcProp(cover.tc, "gridSpan"); else setTcProp(cover.tc, "gridSpan", cover.span - 1);
        if (sub && tr === topTr) setTcProp(cover.tc, "vMerge", "restart");
        continue;
      }
      const a = cellStarting(cells, g0), b = cellStarting(cells, g1);
      if (!a || !b) continue;
      if (tr === headerTr) {
        if (sub) { setCellText(a.tc, ""); setTcProp(a.tc, "vMerge", null); }
        else setCellText(a.tc, "Họ và tên");
      } else if (data.has(tr)) setCellText(a.tc, data.get(tr).name);
      b.tc.parentNode.removeChild(b.tc);
    }
    syncWidths(tbl);
    return true;
  }

  // ---------- Ghi ----------
  async function saveDocx(pkg) {
    let xml = new (g("XMLSerializer"))().serializeToString(pkg.doc);
    if (!/^<\?xml/.test(xml)) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
    pkg.zip.file("word/document.xml", xml);
    return pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  const API = { loadDocx, readClassesDocx, sortDocx, splitDocx, mergeDocx, saveDocx };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.SDBWord = API;
})(typeof window !== "undefined" ? window : globalThis);
