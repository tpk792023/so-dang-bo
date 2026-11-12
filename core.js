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
  const PARENT_RE = /\bcha\b|\bme\b|\bbo\b|phu huynh|giam ho|do dau/;
  function isFullHeader(t) { return (/^ho (va )?ten( |$)/.test(t) && !/(dem|lot)$/.test(t) && !PARENT_RE.test(t)) || t === "ho va"; }
  // "Họ", "Họ và", "Họ đệm", "Họ và tên đệm", "Họ và chữ lót"
  function isHoHeader(t) { return t === "ho" || t === "ho va" || (/^ho /.test(t) && /(dem|lot)$/.test(t)); }
  function isTenHeader(t) { return t === "ten" || t === "ten hoc sinh"; }

  // Phân loại tiêu đề các cột còn lại
  function classifyHeader(h) {
    if (!h) return null;
    if (h.includes("dang bo")) return null; // "Sổ đăng bộ" (chữ "bộ" không phải "bố")
    if (h.length > 70) return null;         // câu hướng dẫn dài, không phải tiêu đề cột
    // thông tin cha / mẹ / người giám hộ
    if (PARENT_RE.test(h)) {
      if (/dien thoai|\bsdt\b|\bdt\b|nam sinh|ngay sinh|cccd|cmnd|can cuoc|email|zalo|dan toc|ton giao|quoc tich/.test(h)) return null;
      const cha = /\bcha\b|\bbo\b/.test(h), me = /\bme\b/.test(h);
      if (h.includes("nghe")) return cha && !me ? "ngheCha" : me && !cha ? "ngheMe" : "nghe";
      if (cha && me) return "chame";
      if (cha) return "cha";
      if (me) return "me";
      return "chame"; // phụ huynh / người giám hộ
    }
    if (h.includes("nghe nghiep") || h === "nghe") return "nghe";
    if (h.includes("stt") || h === "so tt" || h === "tt") return "stt";
    if (h.includes("gioi tinh") || h === "nam nu" || h === "nu") return "gt";
    if (h.includes("dan toc")) return "dt";
    if (h.includes("noi khai sinh")) return "noiks";
    if (h.includes("noi sinh")) return "noisinh";
    if (h.includes("que quan") || h.includes("ma hoc sinh") || h.includes("ma hs") || h.includes("vao truong") || h.includes("can cuoc")) return null;
    if (h.includes("sinh")) return "ns";
    if (h.includes("ho khau") || h.includes("thuong tru")) return "hokhau";
    if (/cho o|dia chi|noi o\b|tam tru|cu tru/.test(h)) return "cho";
    if (h.includes("ghi chu")) return "note";
    if (h === "lop" || h.startsWith("lop ")) return "lop";
    return null;
  }

  // Gán cột theo tiêu đề; "Nghề nghiệp" đứng sau cột cha/mẹ thì thuộc người đó.
  const MULTI_FIELDS = ["cho", "hokhau", "noisinh", "noiks"];
  function assignColumns(textOf, maxC, h) {
    const cols = {};
    Object.defineProperty(cols, "more", { value: {}, enumerable: false });
    let lastParent = null;
    for (let c = 1; c <= maxC; c++) {
      if (c === h.nameCol || c === h.tenCol) continue;
      let k = classifyHeader(textOf(c));
      if (k === "nghe") {
        if (lastParent === "cha") k = "ngheCha";
        else if (lastParent === "me") k = "ngheMe";
        else if (lastParent === "chame") k = cols.ngheCha ? "ngheMe" : "ngheCha";
        else k = null;
      }
      if (k === "cha" || k === "me" || k === "chame") lastParent = k;
      if (k && !cols[k]) cols[k] = c;
      // địa chỉ / nơi sinh chia nhiều cột liền nhau (ấp | xã | tỉnh) → ghép lại khi đọc
      else if (k && MULTI_FIELDS.includes(k) && (cols.more[k] || [cols[k]]).slice(-1)[0] === c - 1)
        (cols.more[k] = cols.more[k] || [cols[k]]).push(c);
    }
    cols.name = h.nameCol;
    if (h.tenCol) cols.ten = h.tenCol;
    return cols;
  }

  // Ghép chữ tiêu đề nhiều tầng: "Thông tin cha" + "Họ tên" → "thong tin cha ho ten"
  // Ghép các phần địa chỉ, bỏ ô trống và phần lặp: "Ấp Vĩnh Nghĩa, Xã Vĩnh Hậu, An Giang"
  function joinParts(list) {
    const out = [];
    list.forEach(s => { s = String(s || "").trim(); if (s && !out.some(o => plain(o) === plain(s))) out.push(s); });
    return out.join(", ");
  }

  function joinHeaderTexts(list) {
    const out = [];
    list.forEach(t => { if (t && out[out.length - 1] !== t) out.push(t); });
    return out.join(" ");
  }

  const PROFILE_FIELDS = ["noisinh", "cho", "cha", "ngheCha", "me", "ngheMe", "chame"];

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
    // các dòng tiêu đề: từ h.top tới h.row, thêm dòng dưới nếu ô họ tên ở đó thuộc vùng gộp (tiêu đề 2 tầng)
    const rows = [];
    for (let r = h.top; r <= h.row; r++) rows.push(r);
    const below = ws.getCell(h.row + 1, h.nameCol);
    if (below.isMerged && below.master.address !== below.address) rows.push(h.row + 1);
    // ô gộp: lấy giá trị ô gốc
    const textOf = c => joinHeaderTexts(rows.map(r => plain(txt(ws.getCell(r, c).value))));
    return assignColumns(textOf, lastCol(ws), h);
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
      // bảng không có cả STT lẫn ngày sinh (VD sheet "Hướng dẫn nhập" của vnEdu) → không phải danh sách học sinh
      if (!cols.stt && !cols.ns) return;
      const sttAt = r => (cols.stt ? val(ws, r, cols.stt) : null);
      const nameAt = r => cols.ten
        ? joinName(txt(val(ws, r, cols.name)), txt(val(ws, r, cols.ten)))
        : txt(val(ws, r, cols.name)).replace(/\s+/g, " ").trim();
      let r = h.row + 1;
      // bỏ qua dòng tiêu đề gộp (trống) và dòng đánh số cột (1) (2) (3)…
      for (let k = 0; k < 3; k++) {
        const nm = nameAt(r);
        if ((!nm && !isNum(sttAt(r))) || /^\(?\d+\)?$/.test(nm)) r++; else break;
      }
      const first = r, rows = [];
      for (; r <= ws.rowCount; r++) {
        const sttV = sttAt(r);
        const name = nameAt(r);
        const sttT = txt(sttV).trim();
        if (sttT && !isNum(sttV)) break;          // "Danh sách có ..."
        if (!sttT && !name) break;                 // hết danh sách
        const g = c => (c ? txt(val(ws, r, c)).trim() : "");
        const gm = k => joinParts((cols.more[k] || [cols[k]]).map(g));
        rows.push({
          r, name,
          gt: g(cols.gt), dt: g(cols.dt),
          ns: cols.ns ? fmtDate(val(ws, r, cols.ns)) : "",
          noisinh: gm("noisinh") || gm("noiks"), cho: gm("cho") || gm("hokhau"), chame: g(cols.chame),
          cha: g(cols.cha), ngheCha: g(cols.ngheCha), me: g(cols.me), ngheMe: g(cols.ngheMe),
          note: g(cols.note), lop: g(cols.lop),
        });
      }
      const last = r - 1;
      // không có cột STT thì phải có ngày sinh hợp lệ ở phần lớn các dòng (loại sheet hướng dẫn / danh mục)
      if (/huong dan|danh muc/.test(plain(ws.name))) return;
      if (!cols.stt && rows.filter(x => dateKey(x.ns)).length * 2 < rows.length) return;
      // dòng "Danh sách có N học sinh"
      let footer = null;
      for (let fr = r; fr <= Math.min(r + 3, ws.rowCount); fr++) {
        const m = plain(txt(val(ws, fr, 1))).match(/co\s+(\d+)\s+hoc sinh/);
        if (m) { footer = { row: fr, count: +m[1] }; break; }
      }
      // tên lớp: tên sheet nếu giống tên lớp (6A1), không thì tìm "Lớp: 6A1" ở phần đầu trang
      let lopName = ws.name.trim();
      if (!/^\d{1,2}\s*[a-z]\s*\d{0,2}$/i.test(lopName)) {
        for (let rr = 1; rr < h.top; rr++)
          for (let cc = 1; cc <= Math.min(lastCol(ws), 12); cc++) {
            const m = plain(txt(val(ws, rr, cc))).match(/\blop\s*:?\s*(\d{1,2}\s*[a-z]\s*\d{0,2})\b/);
            if (m) { lopName = m[1].replace(/\s+/g, "").toUpperCase(); rr = h.top; break; }
          }
      }
      classes.push(finishClass({
        kind: "excel", name: lopName, sheet: ws.name, header: h.row, headerTop: h.top,
        cols, first, last, rows, footer,
      }));
    });
    return markCrossDuplicates(classes);
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

    const report = [], issues = [];
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
      const miss = [];
      let matched = 0;
      list.forEach((hs0, i) => {
        let hs = hs0;
        if (opts.profiles) {
          const p = findProfile(opts.profiles, hs0, cl.name);
          if (p) {
            hs = mergeProfile(hs0, p); matched++;
            // ngày sinh theo nguồn lý lịch (VD vnEdu) khi được chọn
            if (opts.dobFromProfile && dateKey(p.ns) && p.ns !== hs0.ns) {
              hs.ns = p.ns;
              issues.push({ lop: cl.name, name: hs0.name, kind: "Ngày sinh khác", detail: `Danh sách: ${hs0.ns || "(trống)"} → lấy theo lý lịch: ${p.ns}` });
            } else if (p.ns && hs0.ns && p.ns !== hs0.ns)
              issues.push({ lop: cl.name, name: hs0.name, kind: "Ngày sinh khác", detail: `Danh sách: ${hs0.ns}, lý lịch: ${p.ns} (giữ theo danh sách)` });
            if (p._flag === "ten") issues.push({ lop: cl.name, name: hs0.name, kind: "Tên viết khác", detail: `Lý lịch ghi: ${p.name}` });
            if (p._lop && p._lop !== lopKey(cl.name)) issues.push({ lop: cl.name, name: hs0.name, kind: "Khác lớp", detail: `Lý lịch xếp ở lớp ${p._lop.toUpperCase()}` });
            for (const [f, label] of [["gt", "Giới tính"], ["dt", "Dân tộc"]])
              if (p[f] && hs0[f] && plain(p[f]) !== plain(hs0[f])) issues.push({ lop: cl.name, name: hs0.name, kind: label + " khác", detail: `Danh sách: ${hs0[f]}, lý lịch: ${p[f]}` });
          } else {
            miss.push(hs0.name);
            issues.push({ lop: cl.name, name: hs0.name, kind: "Chưa có lý lịch", detail: `Ngày sinh ${hs0.ns || "(trống)"}` });
          }
        }
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
        // cột cha mẹ: mỗi dòng một ý (tên cha / nghề cha / tên mẹ / nghề mẹ)
        parentLines(hs, size).forEach((line, dr) => {
          if (cols.chame && line) ws.getCell(top + dr, cols.chame).value = line;
        });
        set(cols.namhoc, opts.namHoc);
        set(cols.lop, cl.name);
        set(cols.lydo, plain(hs.note).includes("chuyen den") ? "Chuyển đến" : "Tuyển sinh");
      });
      const last = start + list.length * size - 1;
      ws.pageSetup.printArea = `A1:${lastL}${Math.max(last, start)}`;
      ws.pageSetup.printTitlesRow = `${T.hr}:${start - 1}`;
      report.push({ name: sheetName, count: list.length, matched, miss });
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
    return { report, issues, template: tplName, blockSize: size };
  }

  // ---------- Lý lịch học sinh: ghép theo họ tên + ngày sinh ----------
  function nameKeyStr(n) { return plain(n).replace(/[^a-z0-9 ]/g, "").replace(/\s+/g, " ").trim(); }
  function dateKey(d) {
    const m = String(d || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? m[3] + m[2] + m[1] : "";
  }
  const hasProfile = x => PROFILE_FIELDS.some(f => x[f]);
  // yyyymmdd: trùng nhau, hoặc chỉ đảo ngày ↔ tháng
  function sameDate(a, b) {
    return a === b || (a && b && a.slice(0, 4) === b.slice(0, 4) && a.slice(4, 6) === b.slice(6, 8) && a.slice(6, 8) === b.slice(4, 6));
  }

  const lopKey = s => plain(s).replace(/^lop\s*/, "").replace(/[^a-z0-9]/g, "");

  // classes: các lớp/bảng đọc từ file lý lịch (Excel hoặc Word). Lớp lấy từ cột "Lớp", không có thì lấy tên sheet/bảng.
  function buildProfileIndex(classes) {
    const byKey = new Map(), byName = new Map(), byLop = new Map(), seen = new Set();
    let n = 0;
    for (const cl of classes)
      for (const x0 of cl.rows) {
        if (!x0.name || !hasProfile(x0)) continue;
        const x = Object.assign({}, x0, { _lop: lopKey(x0.lop || cl.name) });
        const nk = nameKeyStr(x.name), dk = dateKey(x.ns);
        const uk = nk + "|" + dk + "|" + x._lop;
        if (!seen.has(uk)) { seen.add(uk); n++; }
        if (dk) { const k = nk + "|" + dk; byKey.set(k, (byKey.get(k) || []).concat([x])); }
        byName.set(nk, (byName.get(nk) || []).concat([x]));
        byLop.set(x._lop, (byLop.get(x._lop) || []).concat([x]));
      }
    return { byKey, byName, byLop, count: n };
  }

  // Tìm lý lịch: ưu tiên trùng họ tên + ngày sinh; nhiều em trùng thì xét thêm lớp;
  // chỉ dùng riêng họ tên khi không thể nhầm.
  function findProfile(idx, hs, lop) {
    const nk = nameKeyStr(hs.name), dk = dateKey(hs.ns), lk = lopKey(lop || "");
    // Nhiều dòng cùng một em (có trong nhiều file) → gộp thông tin lại.
    const combine = group => group.reduce((acc, x) => {
      for (const f of PROFILE_FIELDS.concat(["gt", "dt", "ns"])) if (!acc[f] && x[f]) acc[f] = x[f];
      return acc;
    }, Object.assign({}, group[0]));
    const pick = list => {
      if (!list || !list.length) return undefined;
      let group = lk ? list.filter(x => x._lop === lk) : [];
      if (!group.length) {
        if (new Set(list.map(x => x._lop)).size === 1) group = list;
        else return null; // có nhưng không phân biệt được em nào
      }
      return combine(group);
    };
    if (dk) {
      const a = pick(idx.byKey.get(nk + "|" + dk));
      if (a !== undefined) return a;
    }
    const b = idx.byName.get(nk);
    // khác ngày sinh → người khác (trừ trường hợp gõ đảo ngày/tháng, VD 07/01 ↔ 01/07)
    const cand = b && dk ? b.filter(x => !dateKey(x.ns) || sameDate(dateKey(x.ns), dk)) : b;
    const r = pick(cand);
    if (r) return r;
    if (!lk) return null;
    // Cùng lớp + cùng tên, chỉ một em → nhận, ghi chú "khác ngày sinh"
    const sameClassName = (b || []).filter(x => x._lop === lk);
    if (sameClassName.length === 1 && (idx.byLop.get(lk) || []).filter(x => nameKeyStr(x.name) === nk).length === 1)
      return Object.assign(combine(sameClassName), { _flag: "ngaysinh" });
    // Cùng lớp + cùng ngày sinh, tên chỉ khác cách viết (Nguyễn T ↔ Nguyễn Thị, A Ly ↔ Aly, đảo thứ tự)
    if (dk) {
      const loose = (idx.byLop.get(lk) || []).filter(x => sameDate(dateKey(x.ns), dk) && looseName(x.name, hs.name));
      if (loose.length === 1) return Object.assign(combine(loose), { _flag: "ten" });
    }
    return null;
  }

  function looseName(a, b) {
    const A = nameKeyStr(a).split(" "), B = nameKeyStr(b).split(" ");
    if (A.join("") === B.join("")) return true;                               // A Ly ↔ Aly
    if (A.slice().sort().join(" ") === B.slice().sort().join(" ")) return true; // đảo thứ tự
    return A.length === B.length && A.every((w, i) => w === B[i] ||           // Nguyễn T ↔ Nguyễn Thị
      (w.length === 1 && B[i][0] === w) || (B[i].length === 1 && w[0] === B[i]));
  }

  // Cảnh báo học sinh trùng họ tên + ngày sinh ở nhiều lớp
  function markCrossDuplicates(classes) {
    const seen = new Map();
    classes.forEach(cl => cl.rows.forEach(x => {
      if (!x.name || !dateKey(x.ns)) return;
      const k = nameKeyStr(x.name) + "|" + dateKey(x.ns);
      seen.set(k, (seen.get(k) || []).concat([[cl, x]]));
    }));
    for (const list of seen.values()) {
      const lops = [...new Set(list.map(([cl]) => cl.name))];
      if (lops.length < 2) continue;
      for (const [cl, x] of list) {
        const others = lops.filter(l => l !== cl.name).join(", ");
        cl.warnings.push(`${x.name} (${x.ns}) cũng có trong lớp ${others}`);
      }
    }
    return classes;
  }

  function mergeProfile(hs, p) {
    const out = Object.assign({}, hs);
    for (const f of PROFILE_FIELDS.concat(["gt", "dt", "ns"])) if (!out[f] && p[f]) out[f] = p[f];
    return out;
  }

  function parentLines(hs, size) {
    let lines;
    if (hs.cha || hs.ngheCha || hs.me || hs.ngheMe) lines = [hs.cha, hs.ngheCha, hs.me, hs.ngheMe].map(x => x || "");
    else if (hs.chame) lines = [hs.chame];
    else return [];
    if (lines.length <= size) return lines;
    if (size >= 2) {
      const j = (a, b) => [a, b].filter(Boolean).join(" - ");
      return [j(lines[0], lines[1]), j(lines[2], lines[3])];
    }
    return [lines.filter(Boolean).join("; ")];
  }

  // ---------- File mẫu lý lịch để giáo viên chủ nhiệm điền ----------
  const PROFILE_HEADERS = [
    ["STT", 6], ["Họ và tên", 26], ["Ngày sinh", 12], ["Giới tính", 9], ["Dân tộc", 9],
    ["Nơi sinh", 22], ["Chỗ ở hiện tại", 34], ["Họ tên cha", 22], ["Nghề nghiệp cha", 15],
    ["Họ tên mẹ", 22], ["Nghề nghiệp mẹ", 15], ["Ghi chú", 14],
  ];
  function makeProfileTemplate(wb, classes) {
    const thin = { style: "thin", color: { argb: "FF808080" } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };
    for (const cl of classes) {
      const ws = wb.addWorksheet(cl.name, { views: [{ state: "frozen", ySplit: 3, xSplit: 2 }],
        pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      ws.mergeCells(1, 1, 1, PROFILE_HEADERS.length);
      const t = ws.getCell(1, 1);
      t.value = "LÝ LỊCH HỌC SINH LỚP " + cl.name;
      t.font = { name: "Times New Roman", size: 14, bold: true };
      t.alignment = { horizontal: "center" };
      ws.getCell(2, 1).value = "Thầy cô điền các cột còn trống. Giữ nguyên Họ và tên và Ngày sinh để phần mềm ghép đúng học sinh.";
      ws.getCell(2, 1).font = { name: "Times New Roman", size: 11, italic: true, color: { argb: "FF7A5A00" } };
      PROFILE_HEADERS.forEach(([h, w], i) => {
        const c = ws.getCell(3, i + 1);
        c.value = h; ws.getColumn(i + 1).width = w;
        c.font = { name: "Times New Roman", size: 12, bold: true };
        c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE3EFE9" } };
        c.border = border;
      });
      ws.getRow(3).height = 32;
      cl.sorted.forEach((x, i) => {
        const v = [i + 1, x.name, x.ns, x.gt, x.dt, x.noisinh, x.cho, x.cha, x.ngheCha, x.me, x.ngheMe, ""];
        v.forEach((val, j) => {
          const c = ws.getCell(4 + i, j + 1);
          c.value = val || null;
          c.font = { name: "Times New Roman", size: 12, bold: j === 1 };
          c.alignment = { vertical: "middle", wrapText: j >= 5, horizontal: j === 0 || j === 2 ? "center" : "left" };
          c.border = border;
        });
      });
      ws.pageSetup.printTitlesRow = "3:3";
    }
    return wb;
  }

  const API = {
    plain, compareNames, readClasses, sortWorkbook, buildRegister, analyseTemplate, fmtDate,
    splitColumnExcel, mergeColumnExcel, splitName, joinName, locateNameHeader, classifyHeader, finishClass, isNum,
    assignColumns, joinHeaderTexts, buildProfileIndex, findProfile, mergeProfile, makeProfileTemplate, PROFILE_FIELDS,
    markCrossDuplicates,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.SDB = API;
})(typeof window !== "undefined" ? window : globalThis);
