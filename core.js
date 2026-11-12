/* Xử lý danh sách lớp & sổ đăng bộ — dùng ExcelJS.
   Chạy được cả trên trình duyệt (window.SDB) lẫn Node (module.exports). */
(function (root) {
  "use strict";

  // ---------- Tiện ích chữ ----------
  const ALPHA = "aăâbcdđeêfghijklmnoôơpqrstuưvwxyz";
  const TONES = { "̀": 1, "̉": 2, "̃": 3, "́": 4, "̣": 5 }; // huyền, hỏi, ngã, sắc, nặng

  function plain(s) {
    return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase().replace(/\s+/g, " ").trim();
  }

  function wordKey(w) {
    const base = [], tones = [];
    for (const ch of w.toLowerCase()) {
      if (ch === "đ") { base.push(ALPHA.indexOf("đ")); tones.push(0); continue; }
      const d = ch.normalize("NFD");
      let letter = d[0], tone = 0;
      for (const m of d.slice(1)) {
        if (m === "̆" && letter === "a") letter = "ă";
        else if (m === "̂") letter = { a: "â", e: "ê", o: "ô" }[letter] || letter;
        else if (m === "̛") letter = { o: "ơ", u: "ư" }[letter] || letter;
        else if (TONES[m]) tone = TONES[m];
      }
      const i = ALPHA.indexOf(letter);
      base.push(i >= 0 ? i : 100 + letter.codePointAt(0));
      tones.push(tone);
    }
    return [base, tones];
  }

  // Xếp theo tên → họ → đệm; so chữ cái trước, dấu thanh sau.
  function nameKey(name) {
    const parts = String(name).trim().split(/\s+/);
    const ten = wordKey(parts[parts.length - 1]);
    const ho = parts.slice(0, -1).map(wordKey);
    return [ten[0], ho.map(k => k[0]), ten[1], ho.map(k => k[1])];
  }

  function cmp(a, b) {
    if (Array.isArray(a)) {
      for (let i = 0; i < Math.min(a.length, b.length); i++) {
        const c = cmp(a[i], b[i]);
        if (c) return c;
      }
      return a.length - b.length;
    }
    return a - b;
  }

  function compareNames(a, b) { return cmp(nameKey(a), nameKey(b)); }

  // ---------- Tiện ích ô Excel ----------
  function pad(n) { return String(n).padStart(2, "0"); }

  function fmtDate(v) {
    if (v instanceof Date) return pad(v.getUTCDate()) + "/" + pad(v.getUTCMonth() + 1) + "/" + v.getUTCFullYear();
    const s = txt(v);
    const m = s.match(/^\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
    return m ? pad(m[1]) + "/" + pad(m[2]) + "/" + m[3] : s;
  }

  function txt(v) {
    if (v == null) return "";
    if (v instanceof Date) return fmtDate(v);
    if (typeof v === "object") {
      if (v.richText) return v.richText.map(t => t.text).join("");
      if ("result" in v) return txt(v.result);
      if ("text" in v) return String(v.text);
      return "";
    }
    return String(v);
  }

  // Giá trị ô; ô nằm bên trong vùng gộp (không phải ô gốc) coi như trống.
  function val(ws, r, c) {
    const cell = ws.getCell(r, c);
    if (cell.isMerged && cell.master && cell.master.address !== cell.address) return null;
    return cell.value;
  }

  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }

  function isNum(v) { return typeof v === "number" || (typeof v === "string" && /^\s*\d+\s*$/.test(v)); }

  function lastCol(ws) { return Math.max(ws.columnCount || 0, ws.actualColumnCount || 0, 1); }

  // ---------- Đọc danh sách lớp ----------
  // ---- Nhận diện tiêu đề cột (dùng chung cho Excel và Word) ----
  // "Họ và tên", "Họ tên", "Họ và tên học sinh", "Họ và" (thiếu chữ tên)
  function isFullHeader(t) { return (/^ho (va )?ten( |$)/.test(t) && !/(dem|lot)$/.test(t)) || t === "ho va"; }
  // "Họ", "Họ và", "Họ đệm", "Họ và tên đệm", "Họ và chữ lót"
  function isHoHeader(t) { return t === "ho" || t === "ho va" || (/^ho /.test(t) && /(dem|lot)$/.test(t)); }
  function isTenHeader(t) { return t === "ten" || t === "ten hoc sinh"; }

  // Phân loại tiêu đề các cột còn lại
  function classifyHeader(h) {
    if (!h) return null;
    if (h.includes("stt") || h === "so tt" || h === "tt") return "stt";
    if (h.includes("gioi tinh") || h === "nam nu" || h === "nu") return "gt";
    if (h.includes("dan toc")) return "dt";
    if (h.includes("noi sinh")) return "noisinh";
    if (h.includes("sinh")) return "ns";
    if (h.includes("cho o") || h.includes("dia chi")) return "cho";
    if (/\bcha\b|\bme\b/.test(h)) return "chame";
    if (h.includes("ghi chu")) return "note";
    return null;
  }

  // Tìm cột họ tên trên một lưới ô: get(r, c) trả về chữ đã bỏ dấu.
  // Kết quả: { row (dòng tiêu đề cuối), top (dòng tiêu đề đầu), nameCol, tenCol? }
  function locateNameHeader(get, maxR, maxC) {
    for (let r = 1; r <= maxR; r++)
      for (let c = 1; c <= maxC; c++) {
        const t = get(r, c);
        if (!t || !t.startsWith("ho")) continue;
        const next = get(r, c + 1);
        if (isHoHeader(t) && isTenHeader(next)) return { row: r, top: r, nameCol: c, tenCol: c + 1 };
        if (isFullHeader(t) || isHoHeader(t)) {
          // "Họ và tên" gộp ngang, dòng dưới chia "Họ đệm" | "Tên"
          const a = get(r + 1, c), b = get(r + 1, c + 1);
          if (isHoHeader(a) && isTenHeader(b)) return { row: r + 1, top: r, nameCol: c, tenCol: c + 1 };
          return { row: r, top: r, nameCol: c };
        }
      }
    return null;
  }

  function findHeader(ws) {
    const get = (r, c) => plain(txt(val(ws, r, c)));
    return locateNameHeader(get, Math.min(ws.rowCount, 30), Math.min(lastCol(ws), 30));
  }

  function mapColumns(ws, h) {
    const cols = {};
    const maxC = lastCol(ws);
    for (let c = 1; c <= maxC; c++) {
      if (c === h.nameCol || c === h.tenCol) continue;
      // ô tiêu đề gộp dọc: lấy giá trị ô gốc
      const k = classifyHeader(plain(txt(ws.getCell(h.row, c).value))) || classifyHeader(plain(txt(ws.getCell(h.top, c).value)));
      if (k && !cols[k]) cols[k] = c;
    }
    if (!cols.stt) cols.stt = 1;
    cols.name = h.nameCol;
    if (h.tenCol) cols.ten = h.tenCol;
    return cols;
  }

  function splitName(full) {
    const p = String(full || "").trim().split(/\s+/).filter(Boolean);
    if (!p.length) return ["", ""];
    return [p.slice(0, -1).join(" "), p[p.length - 1]];
  }
  function joinName(ho, ten) { return (String(ho || "").trim() + " " + String(ten || "").trim()).replace(/\s+/g, " ").trim(); }

  // Kiểm tra & đóng gói một lớp (dùng chung Excel / Word)
  function finishClass(cl) {
    const named = cl.rows.filter(x => x.name);
    const w = [];
    const blank = cl.rows.length - named.length;
    if (blank) w.push(blank + " dòng trống tên");
    if (cl.footer && cl.footer.count !== named.length)
      w.push(`Ghi "${cl.footer.count} học sinh" nhưng có ${named.length} em`);
    if (cl.cols.gt || cl.cols.ns || cl.cols.dt) {
      const thieu = named.filter(x => (cl.cols.gt && !x.gt) || (cl.cols.ns && !x.ns) || (cl.cols.dt && !x.dt)).map(x => x.name);
      if (thieu.length) w.push("Thiếu giới tính/ngày sinh/dân tộc: " + thieu.join(", "));
    }
    const seen = {}, dup = [];
    named.forEach(x => { const k = plain(x.name) + "|" + x.ns; if (seen[k]) dup.push(x.name); seen[k] = 1; });
    if (dup.length) w.push("Có thể bị trùng: " + dup.join(", "));
    cl.warnings = w;
    cl.split = !!cl.cols.ten;
    cl.sorted = named.slice().sort((a, b) => compareNames(a.name, b.name));
    return cl;
  }

  // Trả về các lớp: { name, sheet, header, cols, first, last, rows:[{r, name, ...}], footer, warnings }
  function readClasses(wb) {
    const classes = [];
    wb.eachSheet(ws => {
      const h = findHeader(ws);
      if (!h) return;
      const cols = mapColumns(ws, h);
      const nameAt = r => cols.ten
        ? joinName(txt(val(ws, r, cols.name)), txt(val(ws, r, cols.ten)))
        : txt(val(ws, r, cols.name)).replace(/\s+/g, " ").trim();
      let r = h.row + 1;
      // bỏ qua dòng tiêu đề gộp (trống) và dòng đánh số cột (1) (2) (3)…
      for (let k = 0; k < 3; k++) {
        const nm = nameAt(r);
        if ((!nm && !isNum(val(ws, r, cols.stt))) || /^\(?\d+\)?$/.test(nm)) r++; else break;
      }
      const first = r, rows = [];
      for (; r <= ws.rowCount; r++) {
        const sttV = val(ws, r, cols.stt);
        const name = nameAt(r);
        const sttT = txt(sttV).trim();
        if (sttT && !isNum(sttV)) break;          // "Danh sách có ..."
        if (!sttT && !name) break;                 // hết danh sách
        const g = c => (c ? txt(val(ws, r, c)).trim() : "");
        rows.push({
          r, name,
          gt: g(cols.gt), dt: g(cols.dt),
          ns: cols.ns ? fmtDate(val(ws, r, cols.ns)) : "",
          noisinh: g(cols.noisinh), cho: g(cols.cho), chame: g(cols.chame),
          note: g(cols.note),
        });
      }
      const last = r - 1;
      // dòng "Danh sách có N học sinh"
      let footer = null;
      for (let fr = r; fr <= Math.min(r + 3, ws.rowCount); fr++) {
        const m = plain(txt(val(ws, fr, 1))).match(/co\s+(\d+)\s+hoc sinh/);
        if (m) { footer = { row: fr, count: +m[1] }; break; }
      }
      classes.push(finishClass({
        kind: "excel", name: ws.name.trim(), sheet: ws.name, header: h.row, headerTop: h.top,
        cols, first, last, rows, footer,
      }));
    });
    return classes;
  }

  // ---------- Chức năng 1: xếp ABC ngay trong file danh sách ----------
  function sortWorkbook(wb, classes, opts) {
    opts = opts || {};
    for (const cl of classes) {
      const ws = wb.getWorksheet(cl.sheet);
      const maxC = lastCol(ws);
      const snap = cl.rows.map(x => {
        const cells = [];
        for (let c = 1; c <= maxC; c++) {
          const cell = ws.getCell(x.r, c);
          cells.push({ v: clone(cell.value), s: clone(cell.style) });
        }
        return { x, cells, h: ws.getRow(x.r).height };
      });
      const named = snap.filter(o => o.x.name).sort((a, b) => compareNames(a.x.name, b.x.name));
      const blank = snap.filter(o => !o.x.name);
      const order = opts.dropBlank ? named : named.concat(blank);
      order.forEach((o, i) => {
        const r = cl.first + i;
        o.cells.forEach((cv, j) => {
          const cell = ws.getCell(r, j + 1);
          cell.value = (j + 1 === cl.cols.stt) ? i + 1 : reviveDate(cv.v);
          cell.style = cv.s || {};
        });
        if (o.h) ws.getRow(r).height = o.h;
      });
      if (opts.dropBlank) {
        for (let r = cl.first + order.length; r <= cl.last; r++)
          for (let c = 1; c <= maxC; c++) ws.getCell(r, c).value = null;
      }
      if (opts.fixFooter && cl.footer) {
        const fc = ws.getCell(cl.footer.row, 1);
        fc.value = txt(fc.value).replace(/(\d+)(\s*học sinh)/i, named.length + "$2");
      }
    }
  }

  function reviveDate(v) {
    if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) return new Date(v);
    return v;
  }

  // ---------- Tách / gộp cột họ tên (Excel) ----------
  function colLetter(n) { let s = ""; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

  // delta = +1: chèn cột trống tại `at`; delta = -1: xoá cột `at`.
  function shiftColumns(ws, at, delta) {
    const maxC = lastCol(ws), maxR = ws.rowCount;
    const merges = (ws.model.merges || []).map(parseMerge);
    merges.forEach(m => ws.unMergeCells(m.r1, m.c1, m.r2, m.c2));
    const widths = [];
    for (let c = 1; c <= maxC + 1; c++) widths[c] = ws.getColumn(c).width;
    const move = (from, to) => {
      for (let r = 1; r <= maxR; r++) {
        const s = ws.getCell(r, from), d = ws.getCell(r, to);
        d.value = clone(s.value) == null ? null : reviveDate(clone(s.value));
        d.style = clone(s.style) || {};
      }
      ws.getColumn(to).width = widths[from];
    };
    if (delta > 0) {
      for (let c = maxC; c >= at; c--) move(c, c + 1);
      for (let r = 1; r <= maxR; r++) ws.getCell(r, at).value = null;
    } else {
      for (let c = at; c < maxC; c++) move(c + 1, c);
      for (let r = 1; r <= maxR; r++) { const d = ws.getCell(r, maxC); d.value = null; d.style = {}; }
    }
    const out = [];
    for (const m of merges) {
      if (delta > 0) {
        if (m.c1 >= at) out.push({ ...m, c1: m.c1 + 1, c2: m.c2 + 1 });
        else if (m.c2 >= at) out.push({ ...m, c2: m.c2 + 1 });
        else {
          out.push(m);
          if (m.c1 === at - 1 && m.c2 === at - 1) out.push({ ...m, c1: at, c2: at }); // ô tiêu đề gộp dọc
        }
      } else {
        if (m.c1 === at && m.c2 === at) continue;
        if (m.c1 > at) out.push({ ...m, c1: m.c1 - 1, c2: m.c2 - 1 });
        else if (m.c2 >= at) { const n = { ...m, c2: m.c2 - 1 }; if (n.c1 !== n.c2 || n.r1 !== n.r2) out.push(n); }
        else out.push(m);
      }
    }
    out.forEach(m => ws.mergeCells(m.r1, m.c1, m.r2, m.c2));
    // vùng in
    const pa = ws.pageSetup && ws.pageSetup.printArea;
    const pm = pa && String(pa).match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (pm) {
      const toN = L => L.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      let c2 = toN(pm[3]);
      if (c2 >= at) c2 += delta;
      ws.pageSetup.printArea = `${pm[1]}${pm[2]}:${colLetter(c2)}${pm[4]}`;
    }
  }

  // Tách "Họ và tên" thành "Họ và chữ lót" | "Tên"
  function splitColumnExcel(wb, cl) {
    if (cl.split) return false;
    const ws = wb.getWorksheet(cl.sheet), nc = cl.cols.name;
    const w = ws.getColumn(nc).width || 25;
    shiftColumns(ws, nc + 1, +1);
    for (let r = 1; r <= ws.rowCount; r++) {
      const src = ws.getCell(r, nc);
      if (src.isMerged && src.master.address !== src.address && src.master.col !== nc) continue;
      ws.getCell(r, nc + 1).style = clone(src.style) || {};
    }
    ws.getCell(cl.header, nc).value = "Họ và chữ lót";
    ws.getCell(cl.header, nc + 1).value = "Tên";
    for (const x of cl.rows) {
      const [ho, ten] = splitName(x.name);
      ws.getCell(x.r, nc).value = ho || null;
      ws.getCell(x.r, nc + 1).value = ten || null;
    }
    ws.getColumn(nc).width = Math.max(Math.round(w * 0.7), 14);
    ws.getColumn(nc + 1).width = Math.max(Math.round(w * 0.4), 8);
    return true;
  }

  // Gộp "Họ và chữ lót" | "Tên" thành "Họ và tên"
  function mergeColumnExcel(wb, cl) {
    if (!cl.split) return false;
    const ws = wb.getWorksheet(cl.sheet), ho = cl.cols.name, ten = cl.cols.ten;
    const w = (ws.getColumn(ho).width || 16) + (ws.getColumn(ten).width || 8);
    for (const x of cl.rows) { ws.getCell(x.r, ho).value = x.name || null; ws.getCell(x.r, ten).value = null; }
    const sub = cl.header !== cl.headerTop;
    if (sub) { ws.getCell(cl.header, ho).value = null; ws.getCell(cl.header, ten).value = null; }
    else { ws.getCell(cl.header, ho).value = "Họ và tên"; ws.getCell(cl.header, ten).value = null; }
    shiftColumns(ws, ten, -1);
    if (sub) {
      ws.mergeCells(cl.headerTop, ho, cl.header, ho);
      ws.getCell(cl.headerTop, ho).value = "Họ và tên";
    }
    ws.getColumn(ho).width = w;
    return true;
  }

  // ---------- Chức năng 2: nhập sổ đăng bộ ----------
  function parseMerge(s) {
    const m = s.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    const col = L => L.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    return { c1: col(m[1]), r1: +m[2], c2: col(m[3]), r2: +m[4] };
  }

  function analyseTemplate(wb) {
    let tpl = null, hr = 0, hdr = null;
    wb.eachSheet(ws => {
      if (tpl) return;
      const h = findHeader(ws);
      if (h) { tpl = ws; hr = h.row; hdr = h; }
    });
    if (!tpl) throw new Error('Không tìm thấy sheet mẫu có cột "Họ và tên học sinh".');
    const maxC = lastCol(tpl);
    const cols = { stt: 1, name: 2, gt: 3, ns: 4, noisinh: 5, dt: 6, cho: 7, chame: 8, namhoc: 9, lop: 10, lydo: 11 };
    for (let c = 1; c <= maxC; c++) {
      const h = plain(txt(val(tpl, hr, c)));
      if (!h) continue;
      if (h.startsWith("so tt") || h.includes("stt")) cols.stt = c;
      else if (h.startsWith("ho va ten cha") || h.includes("nghe nghiep")) cols.chame = c;
      else if (/^ho (va )?ten( |$)|^ho va$/.test(h)) cols.name = c;
      else if (h.includes("nam nu") || h.includes("gioi tinh")) cols.gt = c;
      else if (h.includes("noi sinh")) cols.noisinh = c;
      else if (h.includes("ngay thang") || h.includes("ngay sinh")) cols.ns = c;
      else if (h.includes("dan toc")) cols.dt = c;
      else if (h.includes("cho o")) cols.cho = c;
      else if (h.includes("vao truong")) { cols.namhoc = c; cols.lop = c + 1; cols.lydo = c + 2; }
    }
    cols.name = hdr.nameCol;
    if (hdr.tenCol) cols.ten = hdr.tenCol;
    // dòng học sinh đầu tiên: cột STT là số, cột tên không phải số
    let start = 0;
    for (let r = hr + 1; r <= Math.min(tpl.rowCount, hr + 15); r++) {
      if (isNum(val(tpl, r, cols.stt)) && !isNum(val(tpl, r, cols.name))) { start = r; break; }
    }
    if (!start) throw new Error("Không xác định được dòng học sinh đầu tiên trong sheet mẫu.");
    // độ cao một khối = khoảng cách tới ô STT kế tiếp
    let size = 0;
    for (let r = start + 1; r <= Math.min(tpl.rowCount, start + 20); r++)
      if (txt(val(tpl, r, cols.stt)).trim()) { size = r - start; break; }
    if (!size) size = 1;
    const proto = (tpl.rowCount >= start + 2 * size - 1) ? start + size : start;
    const merges = (tpl.model.merges || []).map(parseMerge);
    const blockMerges = merges.filter(m => m.r1 >= proto && m.r2 < proto + size)
      .map(m => ({ c1: m.c1, c2: m.c2, dr1: m.r1 - proto, dr2: m.r2 - proto }));
    const headMerges = merges.filter(m => m.r2 < start);
    // dòng khóa học
    let khoaCell = null;
    for (let r = 1; r < hr; r++)
      for (let c = 1; c <= maxC; c++)
        if (plain(txt(val(tpl, r, c))).startsWith("khoa hoc")) khoaCell = { r, c };
    return { tpl, hr, cols, start, size, proto, blockMerges, headMerges, maxC, khoaCell };
  }

  function buildRegister(tplWb, classes, opts) {
    opts = opts || {};
    const T = analyseTemplate(tplWb);
    const { tpl, cols, start, size, proto, maxC } = T;
    const tplName = tpl.name;
    tpl.name = "__mau__" + Date.now();
    const lastL = colLetter(maxC);

    const protoCells = [], protoH = [];
    for (let dr = 0; dr < size; dr++) {
      protoH.push(tpl.getRow(proto + dr).height);
      const row = [];
      for (let c = 1; c <= maxC; c++) row.push(clone(tpl.getCell(proto + dr, c).style));
      protoCells.push(row);
    }

    const report = [];
    for (const cl of classes) {
      let sheetName = cl.name;
      if (tplWb.getWorksheet(sheetName)) tplWb.removeWorksheet(tplWb.getWorksheet(sheetName).id);
      const ws = tplWb.addWorksheet(sheetName, {
        properties: clone(tpl.properties),
        pageSetup: Object.assign(clone(tpl.pageSetup) || {}, { printArea: undefined, printTitlesRow: undefined }),
        views: clone(tpl.views),
        headerFooter: clone(tpl.headerFooter),
      });
      for (let c = 1; c <= maxC; c++) {
        const tc = tpl.getColumn(c), nc = ws.getColumn(c);
        if (tc.width) nc.width = tc.width;
        if (tc.hidden) nc.hidden = true;
      }
      // phần đầu trang
      for (let r = 1; r < start; r++) {
        const h = tpl.getRow(r).height;
        if (h) ws.getRow(r).height = h;
        for (let c = 1; c <= maxC; c++) {
          const s = tpl.getCell(r, c), d = ws.getCell(r, c);
          d.value = clone(val(tpl, r, c));
          d.style = clone(s.style) || {};
        }
      }
      for (const m of T.headMerges) ws.mergeCells(m.r1, m.c1, m.r2, m.c2);
      if (T.khoaCell && opts.khoaHoc) ws.getCell(T.khoaCell.r, T.khoaCell.c).value = "KHÓA HỌC " + opts.khoaHoc;

      const list = cl.sorted;
      list.forEach((hs, i) => {
        const top = start + i * size;
        for (let dr = 0; dr < size; dr++) {
          const row = ws.getRow(top + dr);
          if (protoH[dr]) row.height = protoH[dr];
          for (let c = 1; c <= maxC; c++) ws.getCell(top + dr, c).style = clone(protoCells[dr][c - 1]) || {};
        }
        for (const m of T.blockMerges) ws.mergeCells(top + m.dr1, m.c1, top + m.dr2, m.c2);
        const set = (c, v) => { if (c && v !== "" && v != null) ws.getCell(top, c).value = v; };
        set(cols.stt, i + 1);
        const nm = opts.upper === false ? hs.name : hs.name.toLocaleUpperCase("vi");
        if (cols.ten) { const [a, b] = splitName(nm); set(cols.name, a); set(cols.ten, b); }
        else set(cols.name, nm);
        set(cols.gt, hs.gt);
        set(cols.ns, hs.ns);
        set(cols.noisinh, hs.noisinh);
        set(cols.dt, hs.dt);
        set(cols.cho, hs.cho);
        set(cols.chame, hs.chame);
        set(cols.namhoc, opts.namHoc);
        set(cols.lop, cl.name);
        set(cols.lydo, plain(hs.note).includes("chuyen den") ? "Chuyển đến" : "Tuyển sinh");
      });
      const last = start + list.length * size - 1;
      ws.pageSetup.printArea = `A1:${lastL}${Math.max(last, start)}`;
      ws.pageSetup.printTitlesRow = `${T.hr}:${start - 1}`;
      report.push({ name: sheetName, count: list.length });
    }
    tplWb.removeWorksheet(tpl.id);

    // Bìa: cập nhật năm học
    if (opts.namHoc) {
      tplWb.eachSheet(ws => {
        const maxR = Math.min(ws.rowCount, 40), mc = Math.min(lastCol(ws), 20);
        for (let r = 1; r <= maxR; r++)
          for (let c = 1; c <= mc; c++) {
            const t = plain(txt(val(ws, r, c)));
            if (t === "nam hoc:" || t === "nam hoc") {
              const n = ws.getCell(r, c + 1);
              if (/\d{4}\s*-\s*\d{4}/.test(txt(n.value)) || !txt(n.value)) n.value = opts.namHoc;
            }
          }
      });
    }
    return { report, template: tplName, blockSize: size };
  }

  const API = {
    plain, compareNames, readClasses, sortWorkbook, buildRegister, analyseTemplate, fmtDate,
    splitColumnExcel, mergeColumnExcel, splitName, joinName, locateNameHeader, classifyHeader, finishClass, isNum,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.SDB = API;
})(typeof window !== "undefined" ? window : globalThis);
