/* Nhập đề từ file Word (.docx) / .txt ngay trên trình duyệt — chuyển từ quizz/extract.py, quizz/parser.py,
   quizz/moodle.py của bản Windows (giữ cùng quy tắc để hai bản cho cùng kết quả).
   Không dùng thư viện ngoài: giải nén .docx bằng DecompressionStream có sẵn của trình duyệt. */
(function (root) {
"use strict";

/* ============================================================ giải nén .docx (zip) */
async function unzip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  for (let k = u8.length - 22; k >= Math.max(0, u8.length - 65557); k--) {
    if (dv.getUint32(k, true) === 0x06054b50) { eocd = k; break; }
  }
  if (eocd < 0) throw new ImportError("File không phải định dạng .docx hợp lệ (có thể là file .doc cũ bị đổi đuôi). " +
    "Hãy mở bằng Word và chọn Save As → Word Document (.docx).");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = {};
  for (let n = 0; n < count && dv.getUint32(p, true) === 0x02014b50; n++) {
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
    entries[name] = { method, csize, local };
    p += 46 + nlen + xlen + clen;
  }
  return {
    has: name => name in entries,
    async text(name) {
      const e = entries[name];
      if (!e) return null;
      const start = e.local + 30 + dv.getUint16(e.local + 26, true) + dv.getUint16(e.local + 28, true);
      const raw = u8.subarray(start, start + e.csize);
      let data = raw;
      if (e.method === 8) {
        if (typeof DecompressionStream === "undefined")
          throw new ImportError("Trình duyệt này quá cũ để đọc file Word. Hãy cập nhật iOS / Chrome, hoặc nhập đề bằng máy tính.");
        const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        data = new Uint8Array(await new Response(ds).arrayBuffer());
      } else if (e.method !== 0) throw new ImportError("File .docx dùng kiểu nén lạ, không đọc được.");
      return new TextDecoder().decode(data);
    },
  };
}

/* ============================================================ đọc XML tối giản (cây phần tử) */
const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const unent = s => s.indexOf("&") < 0 ? s : s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
  e[0] === "#" ? String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10))
    : (ENT[e] != null ? ENT[e] : m));
function parseXml(src) {
  const top = { tag: "#root", attrs: {}, children: [] };
  const stack = [top];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<\?[\s\S]*?\?>|<!(?:[^>"']|"[^"]*"|'[^']*')*>|<(\/?)([^\s/>]+)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(src))) {
    const cur = stack[stack.length - 1];
    if (m[6] != null) { cur.children.push(unent(m[6])); continue; }
    if (m[1] != null) { cur.children.push(m[1]); continue; }
    if (!m[3]) continue;                     // chú thích, <?xml?>, <!DOCTYPE>
    if (m[2]) {                              // thẻ đóng
      for (let k = stack.length - 1; k > 0; k--) if (stack[k].tag === m[3]) { stack.length = k; break; }
      continue;
    }
    const el = { tag: m[3], attrs: {}, children: [] };
    m[4].replace(/([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g, (_, k, __, a, b) => { el.attrs[k] = unent(a != null ? a : b); });
    cur.children.push(el);
    if (!m[5]) stack.push(el);
  }
  return top.children.find(c => typeof c === "object") || top;
}
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// Tên thẻ trong file luôn có tiền tố (thường là "w:"); đọc tiền tố thật từ khai báo xmlns của phần tử gốc
function wPrefix(rootEl) {
  for (const [k, v] of Object.entries(rootEl.attrs)) if (v === W_NS && k.startsWith("xmlns:")) return k.slice(6) + ":";
  return "w:";
}
const kids = (el, tag) => el ? el.children.filter(c => typeof c === "object" && c.tag === tag) : [];
const kid = (el, tag) => kids(el, tag)[0] || null;
function* iter(el, tag) {   // giống lxml .iter(): duyệt theo thứ tự tài liệu, không gồm chính el
  for (const c of el.children) {
    if (typeof c !== "object") continue;
    if (c.tag === tag) yield c;
    yield* iter(c, tag);
  }
}

/* ============================================================ văn bản + mặt nạ đánh dấu */
class MarkedText {
  constructor() { this.parts = []; this.mask = []; }
  add(text, marked) {
    if (!text) return;
    this.parts.push(text);
    for (let k = 0; k < text.length; k++) this.mask.push(!!marked);
  }
  newline() { this.add("\n", false); }
  get text() { return this.parts.join(""); }
}

const isColorful = (r, g, b) => { const mx = Math.max(r, g, b), mn = Math.min(r, g, b); return mx - mn > 0.25 && mx > 0.25; };
function hexToRgb(v) {
  if (!v || v.length !== 6 || !/^[0-9A-Fa-f]{6}$/.test(v)) return null;
  return [0, 2, 4].map(i => parseInt(v.slice(i, i + 2), 16) / 255);
}
function toRoman(n) {
  const vals = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
                [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of vals) while (n >= v) { out += s; n -= v; }
  return out;
}
function fmtNum(n, fmt) {
  if (fmt === "upperLetter" || fmt === "lowerLetter") {
    let s = "";
    while (n > 0) { const r = (n - 1) % 26; n = Math.floor((n - 1) / 26); s = String.fromCharCode(65 + r) + s; }
    return fmt === "upperLetter" ? s : s.toLowerCase();
  }
  if (fmt === "upperRoman") return toRoman(n);
  if (fmt === "lowerRoman") return toRoman(n).toLowerCase();
  return String(n);
}

/* Tái tạo nhãn đánh số tự động của Word (ví dụ "A.", "Câu 1:") */
class DocxNumbering {
  constructor(W, numberingRoot, stylesRoot) {
    this.W = W; this.levels = new Map(); this.counters = new Map(); this.lastNumId = null; this.styleNum = new Map();
    try { if (stylesRoot) this.readStyleNumbering(stylesRoot); } catch (e) { /* bỏ qua */ }
    if (!numberingRoot) return;
    const abstract = new Map();
    for (const an of kids(numberingRoot, W + "abstractNum")) {
      const lv = new Map();
      for (const lvl of kids(an, W + "lvl")) lv.set(parseInt(lvl.attrs[W + "ilvl"] || "0", 10), this.readLvl(lvl));
      abstract.set(an.attrs[W + "abstractNumId"], lv);
    }
    for (const num of kids(numberingRoot, W + "num")) {
      const nid = num.attrs[W + "numId"], anEl = kid(num, W + "abstractNumId");
      if (!anEl) continue;
      const lv = new Map(abstract.get(anEl.attrs[W + "val"]) || []);
      for (const ov of kids(num, W + "lvlOverride")) {
        const il = parseInt(ov.attrs[W + "ilvl"] || "0", 10), lvl = kid(ov, W + "lvl");
        if (lvl) lv.set(il, this.readLvl(lvl));
        const so = kid(ov, W + "startOverride");
        if (so && lv.has(il)) { const [f, t] = lv.get(il); lv.set(il, [f, t, parseInt(so.attrs[W + "val"] || "1", 10)]); }
      }
      for (const [il, data] of lv) this.levels.set(nid + "|" + il, data);
    }
  }
  readStyleNumbering(stylesRoot) {
    const W = this.W, raw = new Map();
    for (const st of kids(stylesRoot, W + "style")) {
      const based = kid(st, W + "basedOn"), numpr = kid(kid(st, W + "pPr"), W + "numPr");
      let nid = null, il = null;
      if (numpr) {
        const n = kid(numpr, W + "numId"), i = kid(numpr, W + "ilvl");
        nid = n ? n.attrs[W + "val"] || null : null;
        il = i ? parseInt(i.attrs[W + "val"], 10) : null;
      }
      raw.set(st.attrs[W + "styleId"], [nid, il, based ? based.attrs[W + "val"] : null]);
    }
    for (const sid of raw.keys()) {
      let cur = sid, nid = null, il = null;
      const seen = new Set();
      while (cur && !seen.has(cur) && raw.has(cur)) {
        seen.add(cur);
        const [n, i, b] = raw.get(cur);
        nid = nid || n; il = il != null ? il : i; cur = b;
      }
      if (nid) this.styleNum.set(sid, [nid, il || 0]);
    }
  }
  readLvl(lvl) {
    const W = this.W, f = kid(lvl, W + "numFmt"), t = kid(lvl, W + "lvlText"), s = kid(lvl, W + "start");
    return [f ? f.attrs[W + "val"] : "decimal", (t && t.attrs[W + "val"]) || "", s ? parseInt(s.attrs[W + "val"], 10) : 1];
  }
  label(p) {
    const W = this.W, ppr = kid(p, W + "pPr"), numpr = kid(ppr, W + "numPr");
    let styleNid = null, styleIl = null;
    const ps = kid(ppr, W + "pStyle");
    if (ps && this.styleNum.has(ps.attrs[W + "val"])) [styleNid, styleIl] = this.styleNum.get(ps.attrs[W + "val"]);
    if (!numpr && styleNid == null) {   // đoạn thường: các danh sách chữ cái (A, B, C...) bắt đầu lại
      this.resetLetterLists(null); this.lastNumId = null; return "";
    }
    let nid = styleNid, il = styleIl || 0;
    if (numpr) {
      const n = kid(numpr, W + "numId"), i = kid(numpr, W + "ilvl");
      if (n) nid = n.attrs[W + "val"];
      if (i) il = parseInt(i.attrs[W + "val"], 10);
    }
    if (!nid || nid === "0" || !this.levels.has(nid + "|" + il)) return "";
    if (nid !== this.lastNumId) this.resetLetterLists(nid);
    this.lastNumId = nid;
    if (!this.counters.has(nid)) this.counters.set(nid, new Map());
    const cnt = this.counters.get(nid);
    const [fmt, txt, start] = this.levels.get(nid + "|" + il);
    cnt.set(il, (cnt.has(il) ? cnt.get(il) : start - 1) + 1);
    for (const k of [...cnt.keys()]) if (k > il) cnt.delete(k);
    if (fmt === "bullet" || fmt === "none" || !txt) return "";
    let out = txt;
    for (let k = 0; k < 9; k++) {
      if (out.includes("%" + (k + 1))) {
        const [kfmt, , kstart] = this.levels.get(nid + "|" + k) || ["decimal", "", 1];
        out = out.split("%" + (k + 1)).join(fmtNum(cnt.has(k) ? cnt.get(k) : kstart, kfmt));
      }
    }
    return out + " ";
  }
  resetLetterLists(keep) {
    for (const [nid, cnt] of this.counters) {
      if (nid === keep) continue;
      for (const il of [...cnt.keys()]) {
        const fmt = (this.levels.get(nid + "|" + il) || ["decimal"])[0];
        if (fmt === "upperLetter" || fmt === "lowerLetter") cnt.delete(il);
      }
    }
  }
}

function runMarked(W, r, opts) {
  const rpr = kid(r, W + "rPr");
  if (!rpr) return false;
  const val = (el, d) => el.attrs[W + "val"] != null ? el.attrs[W + "val"] : d;
  if (opts.highlight) {
    const hl = kid(rpr, W + "highlight");
    if (hl && !["none", "white"].includes(val(hl, "none"))) return true;
    const shd = kid(rpr, W + "shd");
    if (shd) { const fill = (shd.attrs[W + "fill"] || "").toUpperCase(); if (fill && fill !== "AUTO" && fill !== "FFFFFF") return true; }
  }
  if (opts.color) {
    const c = kid(rpr, W + "color");
    if (c) { const rgb = hexToRgb((c.attrs[W + "val"] || "").toUpperCase()); if (rgb && isColorful(...rgb)) return true; }
  }
  const on = (tag, d, off) => { const e = kid(rpr, W + tag); return !!e && !off.includes(val(e, d)); };
  if (opts.bold && on("b", "true", ["0", "false", "off"])) return true;
  if (opts.underline && on("u", "single", ["none", "0"])) return true;
  if (opts.italic && on("i", "true", ["0", "false", "off"])) return true;
  return false;
}

// Đường dẫn các phần trong .docx (theo file quan hệ _rels, mặc định word/...)
async function docxParts(zip) {
  const rels = async (path) => {
    const src = await zip.text(path);
    return src ? kids(parseXml(src), "Relationship").map(r => r.attrs) : [];
  };
  const join = (base, target) => {
    if (target.startsWith("/")) return target.slice(1);
    const parts = base.split("/").slice(0, -1);
    for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg !== ".") parts.push(seg); }
    return parts.join("/");
  };
  const top = (await rels("_rels/.rels")).find(r => /\/officeDocument$/.test(r.Type || ""));
  const doc = top ? join("", top.Target) : "word/document.xml";
  const relPath = doc.replace(/([^/]+)$/, "_rels/$1.rels");
  const docRels = await rels(relPath);
  const find = (kind, dflt) => { const r = docRels.find(x => new RegExp("/" + kind + "$").test(x.Type || "")); return r ? join(doc, r.Target) : dflt; };
  return { doc, numbering: find("numbering", "word/numbering.xml"), styles: find("styles", "word/styles.xml") };
}

async function readDocx(buf) {
  const zip = await unzip(buf);
  const parts = await docxParts(zip);
  const src = await zip.text(parts.doc);
  if (!src) throw new ImportError("File .docx không có nội dung văn bản.");
  const docRoot = parseXml(src), W = wPrefix(docRoot);
  const numSrc = await zip.text(parts.numbering), stySrc = await zip.text(parts.styles);
  const xml = { docRoot, W, numRoot: numSrc ? parseXml(numSrc) : null, styRoot: stySrc ? parseXml(stySrc) : null };
  return opts => {
    const numbering = new DocxNumbering(W, xml.numRoot, xml.styRoot);
    const out = new MarkedText();
    const body = kid(docRoot, W + "body") || docRoot;
    for (const p of iter(body, W + "p")) {
      out.add(numbering.label(p), false);
      for (const r of iter(p, W + "r")) {
        const marked = runMarked(W, r, opts);
        for (const c of r.children) {
          if (typeof c !== "object") continue;
          if (c.tag === W + "t") out.add(c.children.filter(x => typeof x === "string").join(""), marked);
          else if (c.tag === W + "tab") out.add("\t", marked);
          else if (c.tag === W + "br" || c.tag === W + "cr") out.newline();
        }
      }
      out.newline();
    }
    return out;
  };
}

function decodeTxt(buf) {
  const u8 = new Uint8Array(buf);
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder("utf-16le").decode(u8);
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder("utf-16be").decode(u8);
  try { return new TextDecoder("utf-8", { fatal: true }).decode(u8); } catch (e) { /* thử bảng mã cũ */ }
  try { return new TextDecoder("windows-1258").decode(u8); } catch (e) { return new TextDecoder("latin1").decode(u8); }
}

/* ============================================================ tách câu hỏi (quizz/parser.py) */
const WORD = "\\p{L}\\p{N}_";
const Q_WORD = /^[ \t]*(?:câu\s*hỏi|câu|cau\s*hoi|cau|question|q)\s*(\d{1,4})\s*[:.)\-–]?[ \t]*/gimu;
const Q_NUM = /^[ \t]*(\d{1,4})\s*[.)/:][ \t]+/gmu;
const OPT = new RegExp(`(?<![${WORD}*.])(\\*?)[ \\t]*\\(?([A-Ha-h])[.)](?=\\s|$|\\p{L})`, "gu");
const KEY = new RegExp("^[ \\t]*(?:đáp\\s*án(?:\\s*đúng)?|dap\\s*an(?:\\s*dung)?|đa|answer|ans|key)\\s*[:：\\-=]\\s*" +
  `([A-Ha-h](?:\\s*(?:,|;|&|và|and|\\s)\\s*[A-Ha-h])*)(?![${WORD}])[^\\n]*`, "gimu");
const KEY_TABLE = /^[ \t]*(?:bảng\s*)?(?:đáp\s*án|dap\s*an|answer\s*key|answers)[ \t:]*$/gimu;
const KEY_PAIR = new RegExp(`(?<!\\d)(\\d{1,4})\\s*[.\\-:)]?\\s*([A-Ha-h])(?![${WORD}])`, "gu");
const REF = /^[ \t]*(?:tham\s*khảo|tham\s*khao|tài\s*liệu|nguồn|reference|ref)\s*[:：][^\n]*/gimu;
const EXPLAIN = /^[ \t]*(?:giải\s*thích|giai\s*thich|lời\s*giải|loi\s*giai|hướng\s*dẫn(?:\s*giải)?|explanation)\s*[:：\-]\s*/imu;

// Giống Python finditer(text, lo, hi): nhìn ngược được trước lo, không thấy gì sau hi
function* findIter(re, text, lo = 0, hi = text.length) {
  const r = new RegExp(re.source, re.flags.includes("g") ? re.flags : re.flags + "g");
  const s = text.slice(0, hi);
  r.lastIndex = lo;
  let m;
  while ((m = r.exec(s))) {
    yield m;
    if (m[0] === "") r.lastIndex++;
  }
}
const pyStrip = (s, chars) => {
  let a = 0, b = s.length;
  while (a < b && chars.includes(s[a])) a++;
  while (b > a && chars.includes(s[b - 1])) b--;
  return s.slice(a, b);
};
const isSpace = ch => /\s/u.test(ch);
const isAlpha = ch => /\p{L}/u.test(ch);
const isUpper = ch => ch !== ch.toLowerCase();

function clean(s) {
  s = s.replace(/ /g, " ").replace(/\t/g, " ");
  s = s.replace(/ *\n */g, "\n").replace(/\n{2,}/g, "\n").replace(/ {2,}/g, " ");
  return pyStrip(s, " \n:;-");
}
const joinLines = s => clean(s).replace(/\s*\n\s*/gu, " ").trim();

function splitBlocks(text) {
  let starts = [...findIter(Q_WORD, text)].map(m => [m.index, m.index + m[0].length, m[1]]);
  if (starts.length < 2) starts = [...findIter(Q_NUM, text)].map(m => [m.index, m.index + m[0].length, m[1]]);
  if (starts.length < 2) return [];
  return starts.map(([, e, num], i) => [e, i + 1 < starts.length ? starts[i + 1][0] : text.length, num]);
}

function rfind(text, ch, lo, hi) { for (let k = hi - 1; k >= lo; k--) if (text[k] === ch) return k; return -1; }

function optionChain(text, lo, hi, firstOnly = false) {
  const cands = [];
  for (const m of findIter(OPT, text, lo, hi)) {
    const nl = rfind(text, "\n", lo, m.index);
    const lineStart = nl >= 0 ? nl + 1 : lo;
    cands.push([m.index, m.index + m[0].length, m[2], !!m[1], text.slice(lineStart, m.index).trim() === ""]);
  }
  let best = null, bestKey = null;
  const gt = (a, b) => { for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return a[k] > b[k]; return false; };
  for (let i = 0; i < cands.length; i++) {
    const c = cands[i];
    if (c[2] !== "A" && c[2] !== "a") continue;
    const chain = [c], upper = isUpper(c[2]);
    for (const d of cands.slice(i + 1)) {
      const want = String.fromCharCode(chain[chain.length - 1][2].charCodeAt(0) + 1);
      if (d[2] === want && isUpper(d[2]) === upper) chain.push(d);
      else if (d[2] === c[2] && d[4]) break;   // nhãn "A." mới ở đầu dòng -> sang câu khác
    }
    const key = [chain.length, chain.filter(x => x[4]).length, c[0]];
    if (chain.length >= 2 && (bestKey === null || gt(key, bestKey))) { best = chain; bestKey = key; }
    if (firstOnly && best) break;
  }
  return best;
}

function isMarked(mask, text, start, labelEnd, end) {
  for (let k = start; k < labelEnd; k++) if (mask[k] && isAlpha(text[k])) return true;
  let n = 0, on = 0;
  for (let k = labelEnd; k < end; k++) if (!isSpace(text[k])) { n++; if (mask[k]) on++; }
  return n > 0 && on / n >= 0.5;
}
function blankOut(chars, s, e) { for (let k = s; k < e; k++) if (chars[k] !== "\n") chars[k] = " "; }

let idSeq = 0;
const newId = () => Math.random().toString(36).slice(2, 8) + (idSeq++).toString(36);
const makeQ = (text, options, correct, explanation, number) =>
  ({ id: newId(), text, options, correct, explanation: explanation || "", number: String(number || ""), source: "" });

function parseBlock(text, mask, lo, hi, number, label, warnings) {
  const chars = text.split("");
  const seg = text.slice(lo, hi);
  let keyLetters = [];
  for (const m of findIter(KEY, seg)) {
    keyLetters = [...new Set((m[1].match(/[A-Ha-h]/g) || []).map(ch => ch.toUpperCase().charCodeAt(0) - 65))].sort((a, b) => a - b);
    blankOut(chars, lo + m.index, lo + m.index + m[0].length);
  }
  let explanation = "";
  const em = EXPLAIN.exec(seg);
  if (em) {
    explanation = joinLines(seg.slice(em.index + em[0].length));
    explanation = explanation.replace(new RegExp(KEY.source, "gimu"), "").trim();
    blankOut(chars, lo + em.index, hi);
  }
  const refs = [];
  const joined = chars.join("");
  for (const m of findIter(REF, joined, lo, hi)) {
    refs.push(joinLines(m[0]));
    blankOut(chars, m.index, m.index + m[0].length);
  }
  if (refs.length) explanation = [explanation, ...refs].filter(Boolean).join("\n");
  const text2 = chars.join("");

  const chain = optionChain(text2, lo, hi);
  if (!chain) {
    if (seg.trim()) warnings.push(`${label}: không tìm thấy các phương án A, B, C...`);
    return null;
  }
  let qText = joinLines(text2.slice(lo, chain[0][0]));
  const options = [], multiline = [];
  let correct = [];
  chain.forEach(([s, e, , star], i) => {
    let end = i + 1 < chain.length ? chain[i + 1][0] : hi;
    if (i === chain.length - 1 && multiline.length && !multiline.some(Boolean)) {
      // các phương án trước đều 1 dòng -> phương án cuối cũng chỉ lấy 1 dòng
      const nl = text2.indexOf("\n", e);
      if (nl !== -1 && nl < hi && text2.slice(nl, hi).trim()) end = nl;
    }
    const raw = text2.slice(e, end);
    multiline.push(clean(raw).includes("\n"));
    options.push(joinLines(raw));
    if (star || isMarked(mask, text2, s, e, end)) correct.push(i);
  });
  if (keyLetters.length) correct = keyLetters.filter(k => k < options.length);
  else if (correct.length === options.length && options.length > 2) {
    warnings.push(`${label}: tất cả phương án đều được đánh dấu — bỏ qua đánh dấu, hãy chọn lại đáp án.`);
    correct = [];
  }
  if (!qText) qText = "(Không có nội dung câu hỏi)";
  return makeQ(qText, options, correct, explanation, number);
}

function fallbackBlocks(text) {
  const blocks = [];
  let pos = 0, n = 0;
  while (pos < text.length) {
    const chain = optionChain(text, pos, text.length, true);
    if (!chain) break;
    const nl = text.indexOf("\n", chain[chain.length - 1][1]);
    const end = nl === -1 ? text.length : nl;
    blocks.push([pos, end, String(++n)]);
    pos = end + 1;
  }
  return blocks;
}

// Chữ Nga/Hy Lạp trông giống A, B, C... (hay gặp khi copy từ PDF / OCR)
const LOOK_FROM = "АВСЕНавсеΑΒΕΗ", LOOK_TO = "ABCEHabceABEH";
const LOOKALIKE_LABEL = new RegExp(`(?<![${WORD}.])[${LOOK_FROM}](?=[.)])`, "gu");

function parse(marked) {
  let text = marked.text.replace(LOOKALIKE_LABEL, ch => LOOK_TO[LOOK_FROM.indexOf(ch)]);
  let mask = marked.mask;
  const warnings = [];
  const table = {};
  const heads = [...findIter(KEY_TABLE, text)];
  if (heads.length) {
    const last = heads[heads.length - 1];
    const pairs = [...findIter(KEY_PAIR, text.slice(last.index + last[0].length))];
    if (pairs.length >= 2) {
      for (const m of pairs) (table[String(parseInt(m[1], 10))] = table[String(parseInt(m[1], 10))] || [])
        .push([m[2].toUpperCase().charCodeAt(0) - 65]);
      text = text.slice(0, last.index); mask = mask.slice(0, last.index);
    }
  }
  let blocks = splitBlocks(text);
  if (!blocks.length) blocks = fallbackBlocks(text);
  const questions = [], seen = {};
  for (const [lo, hi, num] of blocks) {
    const key = /^\d+$/.test(num) ? String(parseInt(num, 10)) : num;
    const occ = seen[key] = (seen[key] || 0) + 1;   // đề có thể trùng số câu
    const label = `Câu ${num}` + (occ > 1 ? ` thứ ${occ}` : "");
    const q = parseBlock(text, mask, lo, hi, num, label, warnings);
    if (!q) continue;
    const answers = table[key] || [];
    if (answers.length >= occ) q.correct = [...new Set(answers[occ - 1].filter(k => k < q.options.length))].sort((a, b) => a - b);
    q.source = label;
    questions.push(q);
  }
  if (!questions.length) {
    warnings.push("Không nhận diện được câu hỏi nào. Hãy kiểm tra định dạng: \"Câu 1: ...\" rồi các dòng \"A. ...\", \"B. ...\"");
    return { questions, warnings, note: "" };
  }
  return finish(questions, warnings);
}

const norm = s => s.normalize("NFC").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");

function dedupe(questions) {
  const kept = [], index = new Map(), warnings = [];
  let removed = 0;
  for (const q of questions) {
    const key = norm(q.text) + "\u0000" + [...new Set(q.options.map(norm))].sort().join("\u0000");
    const first = index.get(key);
    if (!first) { index.set(key, q); kept.push(q); continue; }
    const ansFirst = new Set(first.correct.map(i => norm(first.options[i])));
    const ansQ = new Set(q.correct.map(i => norm(q.options[i])));
    const same = ansFirst.size === ansQ.size && [...ansFirst].every(x => ansQ.has(x));
    if (ansFirst.size && ansQ.size && !same) {
      // trùng nội dung nhưng khác đáp án -> giữ cả hai để người dùng tự kiểm tra
      warnings.push(`${q.source} giống ${first.source} nhưng đáp án khác nhau — đã giữ cả hai, hãy kiểm tra lại.`);
      kept.push(q);
      continue;
    }
    if (!ansFirst.size && ansQ.size) first.correct = first.options.map((o, i) => ansQ.has(norm(o)) ? i : -1).filter(i => i >= 0);
    if (!first.explanation && q.explanation) first.explanation = q.explanation;
    removed++;
  }
  return [kept, removed, warnings];
}

function finish(questions, warnings, note = "") {
  const [kept, removed, dup] = dedupe(questions);
  warnings.push(...dup);
  kept.forEach((q, k) => {
    if (!q.correct.length) {
      const where = q.source !== `Câu ${k + 1}` ? ` (trong file ghi là "${q.source}")` : "";
      warnings.push(`Câu ${k + 1}${where}: chưa xác định được đáp án đúng.`);
    }
  });
  note = [note, removed ? `Đã bỏ ${removed} câu trùng lặp.` : ""].filter(Boolean).join(" ");
  return { questions: kept, warnings, note };
}

/* ============================================================ trang "xem lại bài làm" Moodle/LMS (quizz/moodle.py) */
const NW = `(?![${WORD}])`;
const M_HEAD = /^(?:câu\s*hỏi|question)\s*(\d{1,4})$/iu;
const M_TEXT_MARK = /^(?:đoạn\s*văn\s*câu\s*hỏi|question\s*text)$/iu;
const M_SELECT = new RegExp(`^(?:(?:câu\\s*hỏi|question)\\s*\\d{1,4}\\s*)?(?:chọn|select)${NW}[^:]{0,40}:$`, "iu");
const M_FEEDBACK = /^(?:phản\s*hồi|feedback)$/iu;
const M_YOUR = new RegExp(`^(?:câu\\s*trả\\s*lời\\s*của\\s*bạn|your\\s*answer\\s*is)${NW}`, "iu");
const M_RIGHT = /^(?:(?:các\s*)?câu\s*trả\s*lời\s*đúng\s*là|the\s*correct\s*answers?\s*(?:is|are))\s*:?\s*(.*)$/iu;
const M_KEY = /^(?:đáp\s*án\s*đúng\s*là|đáp\s*án\s*đúng|đáp\s*án)\s*:\s*(.*)$/iu;
const M_WHY = /^(?:vì|giải\s*thích|because)\s*:\s*/iu;
const M_STATUS = new RegExp("^(?:đúng|sai|đúng\\s*một\\s*phần|không\\s*trả\\s*lời|chưa\\s*trả\\s*lời|chưa\\s*hoàn\\s*thành|" +
  "đạt\\s*điểm.*|điểm\\s*:?.*|đặt\\s*cờ|bỏ\\s*cờ|xóa\\s*cờ|correct|incorrect|partially\\s*correct|" +
  "not\\s*answered|mark\\s*.*|flag\\s*question|remove\\s*flag)$", "iu");
const M_LABEL = /^\(?([a-hA-H])[.)]\s*/u;
const M_LABEL_ONLY = /^\(?([a-hA-H])[.)]\s*$/u;
const QUOTES = "'\"“”‘’";

function looksLikeMoodle(text) {
  const lines = text.split(/\r\n|\r|\n/).map(l => l.trim());
  return lines.some(l => M_HEAD.test(l)) && lines.some(l => M_TEXT_MARK.test(l) || M_SELECT.test(l)) &&
    lines.some(l => M_RIGHT.test(l));
}
const stripLabel = s => s.replace(M_LABEL, "").trim();
function answerText(s) {
  s = pyStrip(s.trim(), QUOTES).trim();
  const inner = s.slice(0, -1).trim();
  return s.endsWith(".") && QUOTES.includes(inner.slice(-1)) ? pyStrip(inner, QUOTES) : s;
}
function matchAnswer(options, answer) {
  const a = norm(stripLabel(answerText(answer)));
  if (!a) return [];
  const exact = options.map((o, i) => norm(o) === a ? i : -1).filter(i => i >= 0);
  if (exact.length) return exact.slice(0, 1);
  // nhiều đáp án đúng ("Các câu trả lời đúng là: X, Y") hoặc khác nhau chút ít về dấu câu
  let inside = options.map((o, i) => norm(o) && a.includes(norm(o)) ? i : -1).filter(i => i >= 0);
  inside = inside.filter(i => !inside.some(j => i !== j && norm(options[j]).includes(norm(options[i]))));
  if (inside.length) return inside;
  const contains = options.map((o, i) => norm(o).includes(a) ? i : -1).filter(i => i >= 0);
  return contains.length === 1 ? contains : [];
}

function parseMoodleOne(num, lines) {
  let i = lines.findIndex(l => M_TEXT_MARK.test(l));
  if (i < 0) { i = 0; while (i < lines.length && (!lines[i] || M_STATUS.test(lines[i]))) i++; } else i++;
  let sel = null, fb = lines.length;
  for (let k = i; k < lines.length; k++) if (M_SELECT.test(lines[k])) { sel = k; break; }
  for (let k = i; k < lines.length; k++) if (M_FEEDBACK.test(lines[k]) || M_RIGHT.test(lines[k])) { fb = k; break; }
  if (sel === null || sel > fb) return null;
  const qText = lines.slice(i, sel).filter(Boolean).join("\n");

  let options = [], pending = false;
  for (const ln of lines.slice(sel + 1, fb)) {
    if (!ln) continue;
    if (M_LABEL_ONLY.test(ln)) { pending = true; continue; }
    options.push(pending || !options.length || M_LABEL.test(ln) ? stripLabel(ln) : ln);
    pending = false;
  }
  options = options.filter(Boolean);
  if (options.length < 2) return null;

  const explanation = [];
  let right = null, key = null;
  for (let k = fb < lines.length && M_FEEDBACK.test(lines[fb]) ? fb + 1 : fb; k < lines.length; k++) {
    const ln = lines[k];
    const m = M_RIGHT.exec(ln);
    if (m) { right = m[1].trim() || lines.slice(k + 1).find(Boolean) || ""; break; }
    const km = M_KEY.exec(ln);
    if (km) key = km[1];
    else if (ln && !M_YOUR.test(ln)) explanation.push(ln.replace(M_WHY, ""));
  }
  let correct = right ? matchAnswer(options, right) : [];
  if (!correct.length && key) correct = matchAnswer(options, key);
  return makeQ(qText || "(Không có nội dung câu hỏi)", options, correct, explanation.join("\n"), num);
}

function parseMoodle(text) {
  if (!looksLikeMoodle(text)) return null;
  const lines = text.split(/\r\n|\r|\n/).map(l => l.replace(/ /g, " ").trim());
  const heads = [];
  lines.forEach((l, k) => { const m = M_HEAD.exec(l); if (m) heads.push([k, m[1]]); });
  const questions = [], warnings = [];
  heads.forEach(([start, num], n) => {
    const end = n + 1 < heads.length ? heads[n + 1][0] : lines.length;
    const q = parseMoodleOne(num, lines.slice(start + 1, end));
    if (!q) { warnings.push(`Câu hỏi ${num}: không đọc được phương án (có thể không phải câu chọn đáp án).`); return; }
    q.source = `Câu hỏi ${num}`;
    questions.push(q);
  });
  if (!questions.length) return null;
  return finish(questions, warnings, "Đã nhận dạng file bài làm Moodle/LMS (lấy đáp án và giải thích từ phần phản hồi).");
}

/* ============================================================ đọc file (giống parse_file) */
class ImportError extends Error {}
const answered = r => r.questions.filter(q => q.correct.length).length;
const OPTS = { highlight: true, color: true, bold: false, underline: false, italic: false };

// name: tên file; buf: ArrayBuffer. Trả về { questions, warnings, note }.
async function parseFile(name, buf) {
  const ext = (name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
  let extract;
  if (ext === ".docx") extract = await readDocx(buf);
  else if (ext === ".txt") { const t = decodeTxt(buf).replace(/\r\n?/g, "\n"); extract = () => { const m = new MarkedText(); m.add(t, false); return m; }; }
  else if (ext === ".doc") throw new ImportError("File .doc (Word cũ) chưa được hỗ trợ. Hãy mở bằng Word và lưu lại dạng .docx.");
  else if (ext === ".pdf") throw new ImportError("Bản web chưa đọc được PDF. Hãy nhập bằng Quizz trên máy tính rồi gửi sang, hoặc dùng file Word (.docx).");
  else throw new ImportError(`Định dạng ${ext || "này"} chưa được hỗ trợ. Hãy dùng .docx, .txt hoặc .json.`);
  const marked = extract(OPTS);
  const moodle = parseMoodle(marked.text);   // trang xem lại bài làm Moodle/LMS
  if (moodle) return moodle;
  const result = parse(marked);
  if (result.questions.length && answered(result) === 0) {
    // không thấy đáp án bôi màu -> tự thử in đậm / gạch chân / in nghiêng
    const names = { bold: "in đậm", underline: "gạch chân", italic: "in nghiêng" };
    for (const attr of ["bold", "underline", "italic"]) {
      const alt = parse(extract(Object.assign({}, OPTS, { [attr]: true })));
      if (answered(alt) > answered(result)) {
        alt.note = [`Không thấy đáp án bôi màu — đã tự nhận đáp án theo chữ ${names[attr]}.`, alt.note].filter(Boolean).join(" ");
        return alt;
      }
    }
  }
  return result;
}

const api = { parseFile, parse, parseMoodle, MarkedText, ImportError, readDocx, OPTS };
if (typeof module === "object" && module.exports) module.exports = api;
else root.QuizzImport = api;
})(typeof self !== "undefined" ? self : this);
