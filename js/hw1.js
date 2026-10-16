// hw1.js -- Homework 1 (Module 1, sheet 'quiz1#'): evidence engine, rubric, workbook writer.
// Same contract as the Quiz 1 pieces (engine.js / rubric.js / writer.js) so app.js and the
// agent treat both assignments alike.  Every expected value is recomputed from the student's
// own row (data_collection text, quantities, price, city) and from the student's own lookup
// tables, so sorted rows or edited tables never mark a right answer wrong.  No answers are
// stored here.
import {XDate, excelDate, colIndex as CIDX, colLetter as COL} from "./xlsx.js?v=4";
import {clean, num, close, show, pyRepr, isErr, funcs, short, canvasName} from "./engine.js?v=4";
import {Sheet, addFillStyles} from "./writer.js?v=4";

export const ID = "hw1";
export const TITLE = "Homework 1";
export const SHEETS = ["quiz1#", "quiz1", "quiz 1#", "quiz1 #"];
export const NQ = 10;
export const WEIGHTS = [0.10, 0.10, 0.05, 0.10, 0.10, 0.05, 0.15, 0.05, 0.10, 0.20];   // column AA of the template
const QTITLE = {1: "Extract the item", 2: "Extract the category", 3: "Extract the date (m/d/yy)",
  4: "Status from the first number (table S6:T8)", 5: "Due date = 10 working days later, with the holidays",
  6: "Total revenue Q1 and Q2 + top-3 highlighted in each column", 7: "Q evolution (Q2 - Q1) shown as up/down symbol",
  8: "Discount 1: 2% if Q evolution goes down, else 0%", 9: "Discount by location (table S23:T28)",
  10: "Weighted absolute discount (price x total discount) for New York and St. Louis"};
const HEAD = [["raw", "data_collection"], ["item", "item"], ["cat", "category"], ["date", "date"],
  ["status", "status2"], ["due", "due date"], ["q1", "quantity q1"], ["q2", "quantity q2"], ["price", "price"],
  ["j", "total revenue q1"], ["k", "total revenue q2"], ["l", "q evolution"], ["m", "discount 1"],
  ["n", "discount location 2"], ["o", "total discount"], ["city", "location"]];
const NICE = {item: "Item (B)", cat: "Category (C)", date: "Date (D)", status: "Status2 (E)", due: "Due date (F)",
  j: "Total revenue Q1 (J)", k: "Total revenue Q2 (K)", l: "Q evolution (L)", m: "Discount 1 (M)",
  n: "Discount location 2 (N)", o: "total discount (O)"};
const QCOLS = {1: ["item"], 2: ["cat"], 3: ["date"], 4: ["status"], 5: ["due"], 6: ["j", "k"], 7: ["l"], 8: ["m"], 9: ["n"], 10: []};
const RAW = /^\s*(\d+)\s*_(.+?)-(.+?):\s*(\d+(?:\.\d+)?)\s*$/;
const STATUS = {1: "Pending", 2: "Overdue", 3: "Completed"};
const DISCOUNT = {"New York": 0.01, "St. Louis": 0.02, "Los Angeles": 0.03, "Paris": 0.04, "Madrid": 0.05, "Tokyo": 0.06};
const HOLIDAYS = [44927, 44928, 44934];        // = D2, D3, D9 of the template (1/1/23, 1/2/23, 1/8/23)
const UP = new Set(["▲", "↑", "⬆", "up", "UP", "Up", "⇧", "△", "🔼"]), DOWN = new Set(["▼", "↓", "⬇", "down", "DOWN", "Down", "⇩", "▽", "🔽"]);
const REF = /(?<![A-Za-z_"!.])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])/g;
const range = (a, b) => Array.from({length: Math.max(0, b - a + 1)}, (_, i) => a + i);
const pct = x => Math.round(x * 100) + "%";
const round05 = x => Math.round(x * 2) / 2;

function workday(start, n, hol) {
  let d = Math.trunc(start), k = 0;
  while (k < n) { d += 1; const wd = excelDate(d).getUTCDay(); if (wd >= 1 && wd <= 5 && !hol.has(d)) k += 1; }
  return d;
}
function serialOf(v) {
  if (v instanceof XDate) return v.serial;
  if (typeof v === "number") return v;
  if (typeof v !== "string") return null;
  const t = clean(v).trim();
  let m = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})$/.exec(t);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return (Date.UTC(y, +m[1] - 1, +m[2]) - Date.UTC(1899, 11, 30)) / 86400000; }
  m = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
  if (m) return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000;
  if (/^\d{5}$/.test(t)) return +t;
  return null;
}
const dateTxt = s => { const d = excelDate(s); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${String(d.getUTCFullYear()).slice(2)}`; };
function r1c1(f, row, col) {
  return (f || "").replace(REF, (m, ca, c, ra, r) => (ra ? `R${r}` : `R[${+r - row}]`) + (ca ? `C${CIDX(c)}` : `C[${CIDX(c) - col}]`));
}
function rowOffsets(f, row, col) {       // row offsets of the relative references to the data block (cols A-P)
  const out = [];
  for (const m of (f || "").matchAll(REF)) if (!m[3] && CIDX(m[2]) <= 16) out.push(+m[4] - row);
  return out;
}
// arguments of the first call to fn(...) in a formula (strings and nested calls respected)
function callArgs(f, fn) {
  const F = (f || "").toUpperCase().replace(/_XLFN\./g, "");
  const m = new RegExp("(?<![A-Z0-9_.])" + fn.replace(".", "\\.") + "\\(").exec(F);
  if (!m) return null;
  const args = [];
  let depth = 0, cur = "", inStr = false;
  for (let i = m.index + m[0].length; i < F.length; i++) {
    const ch = F[i];
    if (ch === '"') inStr = !inStr;
    if (!inStr && (ch === "(" || ch === "{")) depth++;
    if (!inStr && (ch === ")" || ch === "}")) { if (depth === 0) { args.push(cur.trim()); return args; } depth--; }
    if (!inStr && ch === "," && depth === 0) { args.push(cur.trim()); cur = ""; continue; }
    cur += ch;
  }
  return args;
}
class Counter extends Map {
  inc(k, n = 1) { this.set(k, (this.get(k) || 0) + n); }
  g(k) { return this.get(k) || 0; }
  obj() { return Object.fromEntries(this); }
  top(n) { return [...this.entries()].sort((a, b) => b[1] - a[1]).slice(0, n); }
}

// --------------------------------------------------------------------------- conditional formatting
function sqrefRanges(sq) {
  return sq.trim().split(/\s+/).map(p => {
    const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(p);
    if (!m) return null;
    return {c1: CIDX(m[1]), r1: +m[2], c2: CIDX(m[3] || m[1]), r2: +(m[4] || m[2])};
  }).filter(Boolean);
}
function dxfColors(stylesXml) {
  const block = (stylesXml.match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/) || [, ""])[1];
  return [...block.matchAll(/<dxf>([\s\S]*?)<\/dxf>/g)].map(m => {
    const fill = (m[1].match(/<fill>[\s\S]*?<\/fill>/) || [""])[0];
    const c = fill.match(/<bgColor\b[^>]*?(rgb|theme|indexed)="([^"]*)"(?:[^>]*tint="([^"]*)")?/) || fill.match(/<fgColor\b[^>]*?(rgb|theme|indexed)="([^"]*)"/);
    const font = (m[1].match(/<font>[\s\S]*?<\/font>/) || [""])[0].match(/<color\b[^>]*?(rgb|theme|indexed)="([^"]*)"/);
    return (c ? `fill ${c[1]}:${c[2]}${c[3] ? "/" + c[3] : ""}` : "no fill") + (font ? ` font ${font[1]}:${font[2]}` : "");
  });
}
export function cfRules(book) {
  const xml = book.sheetXml || "", colors = dxfColors(book.stylesXml || ""), out = [];
  for (const m of xml.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const sq = (m[1].match(/sqref="([^"]*)"/) || [, ""])[1];
    for (const r of m[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const a = r[1], g = n => (a.match(new RegExp(`\\b${n}="([^"]*)"`)) || [, null])[1];
      const formula = [...(r[2] || "").matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map(x => x[1]).join(" ; ");
      const icon = ((r[2] || "").match(/<iconSet\b[^>]*iconSet="([^"]*)"/) || [, (r[2] || "").includes("<iconSet") ? "3TrafficLights1" : null])[1];
      out.push({sqref: sq, ranges: sqrefRanges(sq), type: g("type"), rank: g("rank") && +g("rank"), bottom: g("bottom") === "1",
        percent: g("percent") === "1", operator: g("operator"), formula, icon, color: g("dxfId") !== null ? colors[+g("dxfId")] || "?" : null});
    }
  }
  for (const m of xml.matchAll(/<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g)) {
    const sq = (m[1].match(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/) || [, ""])[1];
    for (const r of m[1].matchAll(/<x14:cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/x14:cfRule>)/g)) {
      const type = (r[1].match(/\btype="([^"]*)"/) || [, null])[1];
      const icon = ((r[2] || "").match(/<x14:iconSet\b[^>]*iconSet="([^"]*)"/) || [, type === "iconSet" ? "custom" : null])[1];
      out.push({sqref: sq, ranges: sqrefRanges(sq), type, rank: null, bottom: false, percent: false, operator: null,
        formula: [...(r[2] || "").matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map(x => x[1]).join(" ; "), icon, color: "x14"});
    }
  }
  return out;
}
const covers = (rule, col, r1 = 3, r2 = 101) => rule.ranges.some(g => g.c1 <= col && col <= g.c2 && g.r1 <= r2 && g.r2 >= r1 && (g.r2 - g.r1) >= 20);
function topRule(rule) {
  if (rule.type === "top10" && !rule.bottom) return rule.rank === 3 && !rule.percent ? "top 3" : `top ${rule.rank}${rule.percent ? "%" : ""}`;
  if (/LARGE\s*\(/i.test(rule.formula || "")) return /,\s*3\s*\)/.test(rule.formula) ? "LARGE(...,3) formula" : "LARGE formula";
  if (/RANK/i.test(rule.formula || "")) return "RANK formula";
  return null;
}
function fillsOf(stylesXml) {     // cellXfs index -> has a visible solid fill?
  const fills = [...((stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/) || [, ""])[1]).matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)]
    .map(m => /patternType="solid"/.test(m[1] || "") ? ((m[1].match(/<fgColor\b[^>]*?(rgb|theme|indexed)="([^"]*)"/) || [, "", "?"]).slice(1).join(":")) : null);
  const xfs = (stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [, ""])[1];
  return [...xfs.matchAll(/<xf\b[^>]*?(?:\/>|>)/g)].map(m => { const id = +((m[0].match(/fillId="(\d+)"/) || [, 0])[1]); return fills[id] || null; });
}

// --------------------------------------------------------------------------- the submission
export class HSub {
  constructor(book) {
    this.book = book; this.cells = book.cells; this.notes = [];
    if (book.sheetTitle && book.sheetTitle !== "quiz1#") this.notes.push(`sheet is called '${book.sheetTitle}'`);
    this.mapColumns(); this.findRows(); this.findTables();
  }
  cell(r, c) { return this.cells.get(r * 16384 + c); }
  F(r, c) { const x = this.cell(r, c); return x ? x.f : null; }
  V(r, c) { const x = this.cell(r, c); return x ? x.v : null; }
  NF(r, c) { const x = this.cell(r, c); return x ? x.nf : "General"; }
  S(r, c) { const x = this.cell(r, c); return x ? x.s : 0; }
  fv(r, c) { const x = this.cell(r, c); if (!x) return null; return x.f !== null ? x.f : x.v; }
  mapColumns() {
    this.col = {}; this.inserted = [];
    const names = HEAD.map(h => h[1]);
    let t = 0;
    for (const c of range(1, 30)) {
      if (t >= HEAD.length) break;
      const v = this.V(1, c), h = typeof v === "string" ? clean(v).trim().toLowerCase() : null;
      if (h === names[t]) { this.col[HEAD[t][0]] = c; t++; }
      else if (h !== null && names.slice(t + 1).includes(h)) { const k = names.indexOf(h, t + 1); this.col[HEAD[k][0]] = c; t = k + 1; }
      else if (h !== null && c <= 17) this.inserted.push(c);
    }
    HEAD.forEach(([k, h], i) => {
      if (!(k in this.col)) { this.col[k] = i + 1 + this.inserted.filter(c => c <= i + 1).length; if (i) this.notes.push(`header '${h}' not found; assumed column ${COL(this.col[k])}`); }
    });
    if (this.inserted.length) this.notes.push("extra column(s) " + this.inserted.map(COL).join(", ") + " inside the data block -- answers read from where the student put them");
    this.shift = this.col.city - 16;     // columns moved right by inserted columns
  }
  findRows() {
    this.rows = []; this.example = null;
    const A = this.col.raw;
    for (const r of range(2, Math.max(this.book.maxRow, 101))) {
      const m = RAW.exec(String(clean(this.V(r, A)) ?? ""));
      if (!m) continue;
      if (r === 2 || (this.example === null && /^\s*1_Item 1-/.test(this.V(r, A)))) { this.example = r; continue; }
      this.rows.push(r);
    }
    if (this.example === null) this.notes.push("worked-example row 2 not found");
    const items = this.rows.map(r => +RAW.exec(clean(this.V(r, A)))[2].replace(/\D/g, ""));
    if (items.some((x, i) => i && x < items[i - 1])) this.notes.push("the data rows are NOT in the original order (sorted/moved) -- each row is checked against its own data");
    if (this.rows.length !== 99) this.notes.push(`${this.rows.length} data rows found (template: 99, rows 3-101)`);
  }
  find(pred, rows = range(1, 60), cols = range(17 + this.shift, 32 + this.shift)) {
    for (const r of rows) for (const c of cols) { const v = this.V(r, c); if (pred(v, r, c)) return [r, c]; }
    return null;
  }
  findTables() {
    const s = this.shift, txt = v => typeof v === "string" ? clean(v).trim().toLowerCase() : "";
    // status table: codes 1,2,3 with the names to the right
    const st = this.find((v, r, c) => num(v) === 1 && num(this.V(r + 1, c)) === 2 && num(this.V(r + 2, c)) === 3 && typeof this.V(r, c + 1) === "string");
    this.statusAt = st; this.status = {};
    if (st) for (const i of [0, 1, 2]) this.status[i + 1] = clean(String(this.V(st[0] + i, st[1] + 1))).trim();
    else { Object.assign(this.status, STATUS); this.notes.push("status table (S6:T8) not found; default names used"); }
    const hol = this.find(v => txt(v) === "holidays");
    this.holAt = hol ? [hol[0] + 1, hol[1]] : [12, 19 + s];
    this.holidays = new Set(HOLIDAYS);
    this.holCells = [];
    for (const i of [0, 1, 2]) { const v = this.V(this.holAt[0] + i, this.holAt[1]); this.holCells.push(v); const x = serialOf(v); if (x) this.holidays.add(Math.trunc(x)); }
    const up = this.find(v => txt(v) === "up");
    this.symAt = up ? [up[0] + 1, up[1]] : [18, 19 + s];
    this.upSym = this.V(this.symAt[0], this.symAt[1]); this.downSym = this.V(this.symAt[0], this.symAt[1] + 1);
    const city = this.find(v => txt(v) === "city");
    this.discAt = city ? [city[0] + 1, city[1]] : [23, 19 + s];
    this.discount = {};
    for (const i of range(0, 7)) {
      const k = this.V(this.discAt[0] + i, this.discAt[1]), p = num(this.V(this.discAt[0] + i, this.discAt[1] + 1));
      if (typeof k === "string" && k.trim() && p !== null) this.discount[clean(k).trim()] = p;
    }
    if (Object.keys(this.discount).length < 6) { for (const [k, v] of Object.entries(DISCOUNT)) if (!(k in this.discount)) this.discount[k] = v; this.notes.push("discount table incomplete; default discounts used for the missing cities"); }
    const q10 = this.find(v => /^\s*10\s*\.-/.test(String(v ?? "")), range(1, 60), range(1, 40));
    this.q10At = q10;
    this.cityCells = {};
    for (const name of ["New York", "St. Louis"]) {
      const at = this.find(v => typeof v === "string" && clean(v).trim().toLowerCase() === name.toLowerCase(), range(q10 ? q10[0] : 25, (q10 ? q10[0] : 25) + 12), range(17 + s, 30 + s));
      this.cityCells[name] = at || (name === "New York" ? [32, 19 + s] : [33, 19 + s]);
    }
    const tbl = this.find(v => txt(v) === "question", range(1, 6), range(20, 45));
    this.table = tbl;
  }
}

// --------------------------------------------------------------------------- expected values
function rowExpected(sub, r) {
  const C = sub.col, m = RAW.exec(clean(sub.V(r, C.raw)));
  const code = +m[1], ser = +m[4];
  const g = num(sub.V(r, C.q1)), h = num(sub.V(r, C.q2)), price = num(sub.V(r, C.price));
  const city = clean(String(sub.V(r, C.city) ?? "")).trim();
  const j = g * price, k = h * price;
  const dir = k > j + 1e-9 ? "up" : k < j - 1e-9 ? "down" : "tie";
  const mm = dir === "down" ? 0.02 : 0, nn = sub.discount[city] ?? DISCOUNT[city] ?? null;
  return {item: m[2].trim(), cat: m[3].trim(), date: ser, status: sub.status[code] ?? STATUS[code], code,
    due: workday(ser, 10, sub.holidays), dueNoHol: workday(ser, 10, new Set()), j, k, dir, m: mm, n: nn, o: nn === null ? null : mm + nn, price, city};
}
export function buildReference(books) {
  // the given data must be the same in every file: keep the majority version of each row
  const votes = new Map();
  for (const b of books) {
    let sub; try { sub = new HSub(b); } catch (e) { continue; }
    for (const r of [sub.example, ...sub.rows].filter(x => x !== null)) {
      const key = [sub.V(r, sub.col.raw), sub.V(r, sub.col.q1), sub.V(r, sub.col.q2), sub.V(r, sub.col.price), sub.V(r, sub.col.city)].map(x => clean(String(x ?? "")).trim()).join("|");
      const id = clean(String(sub.V(r, sub.col.raw))).trim();
      if (!votes.has(id)) votes.set(id, new Counter());
      votes.get(id).inc(key);
    }
  }
  const data = new Map();
  for (const [id, c] of votes) data.set(id, c.top(1)[0][0]);
  // expected Q10 from the canonical data (every class gets the same file)
  const rows = [...data.values()].map(k => k.split("|"));
  const q10 = {};
  for (const name of ["New York", "St. Louis"]) {
    let all = 0, noEx = 0;
    for (const [raw, g, h, p, city] of rows) {
      if (city !== name) continue;
      const j = +g * +p, k = +h * +p, o = (k < j - 1e-9 ? 0.02 : 0) + (DISCOUNT[city] || 0), v = +p * o;
      all += v; if (!/^1_Item 1-/.test(raw)) noEx += v;
    }
    q10[name] = {without_row2: noEx, with_row2: all};
  }
  return {data, q10};
}

// --------------------------------------------------------------------------- per-cell judging
const isBlank = v => v === null || v === undefined || (typeof v === "string" && clean(v).trim() === "");
function symDir(v, sub) {
  if (typeof v === "string") {
    const t = clean(v).trim();
    if (t === clean(String(sub.upSym ?? "\u0000")).trim() || UP.has(t)) return "up";
    if (t === clean(String(sub.downSym ?? "\u0000")).trim() || DOWN.has(t)) return "down";
    if (t === "=" || t === "-" || t === "◄►" || t === "→" || t === "▶") return "tie";
    const n = num(t);
    if (n !== null) v = n; else return null;
  }
  if (typeof v === "number") {
    if (num(sub.upSym) !== null && v === num(sub.upSym)) return "up";
    if (num(sub.downSym) !== null && v === num(sub.downSym)) return "down";
    return v > 0 ? "up~" : v < 0 ? "down~" : "tie";
  }
  if (typeof v === "boolean") return v ? "up?" : "down?";
  return null;
}
// returns [status, note]; statuses: ok, cosmetic, format, carried_forward, typical, wrong, blank, error
function judge(key, sub, r, e) {
  const C = sub.col, v = sub.V(r, C[key]);
  if (isBlank(v)) return ["blank", ""];
  if (isErr(v)) return ["error", String(v)];
  const own = k => sub.V(r, C[k]);
  switch (key) {
    case "item": case "cat": case "status": {
      const want = e[key], got = String(clean(v));
      if (got === want) return ["ok", ""];
      if (got.trim().toLowerCase() === want.toLowerCase()) return ["cosmetic", "extra spaces / case"];
      if (key === "status" && num(got) === e.code) return ["typical", "returns the CODE, not the status name"];
      if (key === "status" && Object.values(sub.status).includes(got.trim())) return ["wrong", "a status name, but not this row's"];
      if (key === "item" && got.trim() === `${e.code}_${want}`) return ["typical", "keeps the 'n_' prefix"];
      return ["wrong", ""];
    }
    case "date": {
      const s = serialOf(v);
      if (s !== null && Math.trunc(s) === e.date) return typeof v === "string" ? ["format", "date as TEXT (accepted)"] : ["ok", ""];
      if (typeof v === "string" && clean(v).trim() === String(e.date)) return ["typical", "the serial number as text, not converted to a date"];
      if (typeof v === "number" && v === e.date) return ["ok", ""];
      return ["wrong", ""];
    }
    case "due": {
      const s = serialOf(v);
      if (s === null) return ["wrong", "not a date"];
      if (Math.trunc(s) === e.due) return ["ok", ""];
      const d = serialOf(own("date"));
      if (d !== null && Math.trunc(d) !== e.date && Math.trunc(s) === workday(d, 10, sub.holidays)) return ["carried_forward", "right formula on the student's own (wrong) date"];
      if (Math.trunc(s) === e.date + 10) return ["typical", "date + 10 calendar days (weekends not skipped)"];
      if (Math.trunc(s) === e.date + 14) return ["typical", "date + 14 calendar days"];
      if (Math.trunc(s) === workday(e.date, 9, sub.holidays) || Math.trunc(s) === workday(e.date, 11, sub.holidays)) return ["typical", "off by one working day"];
      return ["wrong", ""];
    }
    case "j": case "k": {
      const n = num(v);
      if (n !== null && close(n, e[key])) return ["ok", ""];
      return ["wrong", ""];
    }
    case "l": {
      const d = symDir(v, sub);
      const jj = num(own("j")), kk = num(own("k"));
      const ownDir = jj !== null && kk !== null ? (kk > jj + 1e-9 ? "up" : kk < jj - 1e-9 ? "down" : "tie") : null;
      if (typeof v === "number" && close(v, e.k - e.j) && e.dir !== "tie") return ["typical", "the difference K-J, not replaced by the symbol"];
      const base = d ? d.replace(/[~?]/, "") : null;
      const fits = dir => base === dir || (dir === "tie" && (base === "up" || base === "tie" || base === "down"));
      if (fits(e.dir)) return d.endsWith("~") ? ["format", "sign of the number (no up/down symbol value)"] : d.endsWith("?") ? ["format", "TRUE/FALSE instead of the symbol"] : ["ok", ""];
      if (ownDir && ownDir !== e.dir && fits(ownDir)) return ["carried_forward", "follows the student's own J/K"];
      if (base && base !== e.dir) return ["wrong", "direction reversed"];
      return ["wrong", ""];
    }
    case "m": {
      const n = num(v);
      if (n !== null && close(n, e.m)) return ["ok", ""];
      if (n !== null && close(n, e.m * 100)) return ["typical", "2 instead of 2% (0.02)"];
      const ld = symDir(own("l"), sub);
      if (ld && n !== null) { const want = ld.startsWith("down") ? 0.02 : 0; if (close(n, want) && (ld.startsWith("down") ? "down" : "up") !== e.dir) return ["carried_forward", "follows the student's own Q evolution"]; }
      if (n !== null && e.dir === "tie" && (close(n, 0) || close(n, 0.02))) return ["ok", "tie row"];
      return ["wrong", ""];
    }
    case "n": {
      const n = num(v);
      if (n !== null && e.n !== null && close(n, e.n)) return ["ok", ""];
      if (n !== null && e.n !== null && close(n, e.n * 100)) return ["typical", "1..6 instead of 1%..6%"];
      return ["wrong", ""];
    }
    case "o": {
      const n = num(v);
      if (n !== null && e.o !== null && close(n, e.o)) return ["ok", ""];
      const mm = num(own("m")), nn = num(own("n"));
      if (n !== null && mm !== null && nn !== null && close(n, mm + nn)) return ["carried_forward", "M+N of the student's own values"];
      return ["wrong", ""];
    }
  }
  return ["wrong", ""];
}

// --------------------------------------------------------------------------- column evidence
function columnStats(sub, key) {
  const c = sub.col[key], st = new Counter(), notes = new Counter(), shapes = new Counter(), shapeCell = {}, offsets = new Counter(), values = new Counter();
  const examples = [];
  let formulas = 0, typed = 0, typedOk = 0;
  for (const r of sub.rows) {
    const e = rowExpected(sub, r);
    const [s, note] = judge(key, sub, r, e);
    const f = sub.F(r, c), v = sub.V(r, c);
    if (s !== "blank") values.inc(show(v));
    if (f) {
      formulas++;
      const sh = r1c1(f, r, c); shapes.inc(sh); if (!(sh in shapeCell)) shapeCell[sh] = `${COL(c)}${r}`;
      for (const o of rowOffsets(f, r, c)) if (o !== 0) offsets.inc(o);
    } else if (s !== "blank") { typed++; if (["ok", "cosmetic", "format"].includes(s)) typedOk++; }
    const cls = !f && s !== "blank" ? "typed" : s;
    st.inc(cls); if (note) notes.inc(note);
    if (!["ok", "cosmetic", "format", "blank"].includes(cls) && examples.length < 4) {
      const want = key === "date" || key === "due" ? dateTxt(e[key]) : key === "l" ? `${e.dir} (up=${show(sub.upSym)}, down=${show(sub.downSym)})` : show(e[key]);
      examples.push({cell: `${COL(c)}${r}`, status: cls, got: show(v), expected: want, formula: f ? short(f, 110) : "typed", note});
    }
  }
  const n = sub.rows.length || 1;
  const patterns = shapes.top(3).map(([sh, k]) => ({count: k, cell: shapeCell[sh], formula: short(sub.F(+shapeCell[sh].replace(/\D/g, ""), c), 140)}));
  return {key, column: COL(c), n: sub.rows.length, stats: st.obj(), notes: notes.obj(), formulas, typed, typedOk,
    patterns, n_patterns: shapes.size, offsets: offsets.obj(), distinct: values.size, top_values: values.top(3).map(x => x[0]), examples,
    funcs: [...new Set(patterns.flatMap(p => funcs(p.formula)))], score: colScore(st, n)};
}
function colScore(st, n) {
  const good = st.g("ok") + st.g("cosmetic") + st.g("format") + 0.9 * st.g("carried_forward");
  let s = 10 * good / n;
  if (st.g("ok") + st.g("cosmetic") + st.g("format") >= 0.97 * n) s = 10;
  return round05(s);
}

// --------------------------------------------------------------------------- questions
function draft(q, score, verdict, comment, fix) {
  comment = comment.charAt(0).toUpperCase() + comment.slice(1);
  return {score: Math.max(0, Math.min(10, round05(score))), verdict, comment, fix};
}
// generic hints for the rules-only mode (no model formulas in this public file: the agent
// writes the exact corrected formula for each student)
const FIX = {
  1: "Find the positions of '_' and '-' with FIND and take the text between them with MID (or TEXTAFTER/TEXTBEFORE); fill down.",
  2: "Same idea as Q1: the text between '-' and ':' (FIND + MID, or TEXTAFTER/TEXTBEFORE).",
  3: "Take the number after ':' and turn it into a real date (VALUE or *1), then format the cell as m/d/yy.",
  4: "Take the leading number, convert it to a number and look it up in the status table with an exact-match lookup; lock the table with $.",
  5: "Use WORKDAY(date, 10, holidays) with the holiday range locked with $.",
  6: "Revenue = quantity x price for each quarter; then Conditional Formatting > Top 10 Items -> 3 on each column, a different colour per column.",
  7: "Compare Q2 with Q1 revenue in an IF and return the up/down symbol from the symbols table (locked with $).",
  8: "IF the Q evolution is down then 2% else 0 (write 2% or 0.02, not 2).",
  9: "Exact-match lookup of the city in the discount table (VLOOKUP with FALSE as 4th argument), table locked with $.",
  10: "For each city: sum of price x total discount over that city's rows only (SUMPRODUCT with a city condition, or SUMIF over a helper column); one formula per city in T32/T33."};
function columnQuestion(sub, q) {
  const cols = QCOLS[q].map(k => columnStats(sub, k));
  const extra = {};
  let score = cols.reduce((a, c) => a + c.score, 0) / cols.length;
  const issues = [];
  const all = new Counter(); for (const c of cols) for (const [k, v] of Object.entries(c.stats)) all.inc(k, v);
  const n = cols.reduce((a, c) => a + c.n, 0) || 1;
  if (q === 5) {
    const main = cols[0].patterns[0] ? sub.F(+cols[0].patterns[0].cell.replace(/\D/g, ""), sub.col.due) : null;
    const args = callArgs(main, "WORKDAY.INTL") || callArgs(main, "WORKDAY");
    extra.holidays_cells = sub.holCells.map(show);
    if (args) {
      const hol = (callArgs(main, "WORKDAY.INTL") ? args[3] : args[2]) || "";
      const hasHol = !!hol;
      const absHol = /^\$[A-Z]+\$\d+:\$[A-Z]+\$\d+$/.test(hol) || (/^[A-Z_][A-Z0-9_.]*$/.test(hol) && !/^[A-Z]{1,3}\d+$/.test(hol)) || /^\{/.test(hol);
      extra.workday_arguments = args;
      extra.holidays_argument = hasHol ? (absHol ? "present, absolute" : "present but RELATIVE (drifts when filled down)") : "MISSING (no holidays)";
      if (!hasHol) { score -= 2; issues.push("the holidays are not passed to WORKDAY"); }
      else if (!absHol) { score -= 1; issues.push("the holiday range is not absolute ($S$12:$S$14)"); }
    } else if (cols[0].formulas) {
      // D+14 equals 10 working days whenever the date is a weekday and no holiday falls in
      // between, so most rows look right: the method is still wrong
      extra.workday = "no WORKDAY in the formulas (calendar arithmetic)";
      score = Math.min(score, 3); issues.push("no WORKDAY: adding calendar days does not skip weekends and holidays (right only by coincidence on some rows)");
    }
    extra.note = "the holidays only change the result of the example row 2, so values alone cannot show whether they were used -- check the formula";
  }
  if (q === 6) {
    const rules = cfRules(sub.book);
    const fills = fillsOf(sub.book.stylesXml || "");
    const res = {};
    for (const k of ["j", "k"]) {
      const c = sub.col[k];
      const top = rules.filter(r => covers(r, c)).map(r => ({kind: topRule(r) || r.type, color: r.color, sqref: r.sqref})).filter(x => x.kind);
      // hand colouring = cells whose fill differs from the column's usual fill (the template
      // already shades every answer cell)
      const usual = new Counter(); for (const r of sub.rows) usual.inc(fills[sub.S(r, c)] || "none");
      const base = usual.top(1)[0]?.[0] || "none";
      const manual = sub.rows.filter(r => (fills[sub.S(r, c)] || "none") !== base).length;
      res[COL(c)] = {conditional_formatting: top, manually_filled_cells: manual};
    }
    extra.highlight = res;
    const okCol = Object.values(res).map(x => x.conditional_formatting.some(t => /top 3|LARGE\(\.\.\.,3\)/.test(t.kind)));
    const anyTop = Object.values(res).map(x => x.conditional_formatting.some(t => /top|LARGE|RANK/i.test(t.kind)));
    const manual = Object.values(res).map(x => x.manually_filled_cells >= 1 && x.manually_filled_cells <= 12);
    const hl = okCol.map((ok, i) => ok ? 3 : anyTop[i] ? 2 : manual[i] ? 1 : 0).reduce((a, b) => a + b, 0) / 2;    // 0..3
    const colours = Object.values(res).map(x => (x.conditional_formatting[0] || {}).color).filter(Boolean);
    let s = 0.7 * score + hl;
    if (hl === 3 && colours.length === 2 && colours[0] === colours[1] && colours[0] !== "x14") { s -= 0.5; issues.push("the same colour in both columns (the task asks for different colours)"); }
    if (hl < 3) issues.push(hl === 0 ? "no top-3 highlight (conditional formatting) in J and K" : "the top-3 highlight is incomplete or done by hand");
    score = s;
  }
  if (q === 7) {
    const rules = cfRules(sub.book).filter(r => covers(r, sub.col.l));
    const icons = rules.filter(r => r.icon || r.type === "iconSet");
    const fmts = new Counter(); for (const r of sub.rows) fmts.inc(sub.NF(r, sub.col.l));
    extra.symbols_table = `up = ${show(sub.upSym)}, down = ${show(sub.downSym)}`;
    extra.icon_set_on_L = icons.length ? icons.map(r => r.icon || "iconSet").join(", ") : "none";
    extra.number_formats_L = fmts.obj();
    const diff = cols[0].stats.typical || 0;
    const symbolicFormat = icons.length || Object.keys(fmts.obj()).some(f => /[▲▼↑↓]|up|down/i.test(f));
    if (diff && symbolicFormat) { score = Math.max(score, 10 * (diff + (cols[0].stats.ok || 0)) / n); issues.push("kept the difference but shows it with arrows (icon set / number format) -- accepted"); }
    else if (diff > 0.5 * n) { score = Math.min(score + 5 * diff / n, 5); issues.push("calculated the difference but did not substitute it with the up/down symbols"); }
    if ((cols[0].stats.format || 0) > 0.5 * n) { score = Math.min(score, 8); issues.push("shows the sign / TRUE-FALSE instead of the symbols of S18:T18"); }
  }
  if (q === 8 || q === 9) {
    // the only mistake is the scale (2 instead of 2%, 3 instead of 3%): -2, not a share of rows
    const c0 = cols[0], bad = c0.n - (c0.stats.ok || 0) - (c0.stats.cosmetic || 0) - (c0.stats.format || 0);
    if (bad > 0 && (c0.stats.typical || 0) === bad && Object.keys(c0.notes).every(k => /instead of/.test(k))) { score = 8; issues.push("percent written as a whole number (2 instead of 2% = 0.02)"); }
  }
  if (q === 9) {
    const approx = cols[0].patterns.some(p => { const a = callArgs(sub.F(+p.cell.replace(/\D/g, ""), sub.col.n), "VLOOKUP"); return a && (a.length < 4 || /^(TRUE|1)$/.test(a[3])); });
    if (approx) { extra.vlookup = "VLOOKUP WITHOUT the 4th argument (approximate match) -- wrong for an unsorted table"; }
    extra.total_discount_O = (({stats, column, examples, patterns}) => ({column, stats, formula: patterns[0]?.formula, examples: examples.slice(0, 2)}))(columnStats(sub, "o"));
  }
  // typed values: 0 for those cells (already in the column score); say it
  const typed = cols.reduce((a, c) => a + c.typed, 0), blank = all.g("blank");
  if (typed && typed >= 0.5 * n) issues.push(`${typed} of ${n} cells are typed values, not formulas (typed = 0)`);
  if (blank === n) return {q, title: QTITLE[q], kind: "column", columns: cols, extra, proposal: draft(q, 0, "blank", "Not answered.", FIX[q])};
  const off = cols.flatMap(c => Object.keys(c.offsets).map(Number)).filter(x => x !== 0);
  if (off.length && cols.some(c => Object.values(c.offsets).reduce((a, b) => a + b, 0) > 0.5 * c.formulas)) issues.push(`the formula reads another row (row offset ${[...new Set(off)].join(", ")})`);
  const wrongShare = (all.g("wrong") + all.g("typical") + all.g("error")) / n;
  let verdict = score >= 9.5 ? "correct" : all.g("carried_forward") > 0.3 * n ? "carried_forward" : typed >= 0.5 * n ? "hardcoded" : score >= 7 ? "minor" : score >= 3 ? "partial" : "major";
  if (all.g("error") > 0.5 * n) verdict = "error";
  const cw = cols.flatMap(c => c.examples.slice(0, 1).map(e => `${e.cell}: got ${e.got}, expected ${e.expected}${e.note ? " (" + e.note + ")" : ""}`));
  let comment;
  if (verdict === "correct" && !issues.length) comment = "Correct.";
  else comment = [issues.join("; "), wrongShare > 0 ? `${Math.round(wrongShare * 100)}% of the rows are wrong, e.g. ${cw.join("; ")}` : ""].filter(Boolean).join(". ") + ".";
  if (all.g("carried_forward") && verdict !== "correct") comment += " Part of the difference comes from your own earlier column (carried-forward error, small penalty only).";
  return {q, title: QTITLE[q], kind: "column", columns: cols, extra, proposal: draft(q, score, verdict, comment.replace(/^\. /, ""), verdict === "correct" ? "" : FIX[q])};
}

function singleRefs(f) { return [...(f || "").replace(/\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+/g, "").matchAll(REF)].map(m => m[2] + m[4]); }
function q10Question(sub, ref) {
  const s = sub.shift, out = {q: 10, title: QTITLE[10], kind: "single", cities: {}};
  const exp = ref.q10;
  // student's own O (carried forward) and typical variants
  const own = {}, variants = {};
  for (const name of ["New York", "St. Louis"]) {
    let o = 0, okOwn = true, onlyN = 0, price = 0, cnt = 0, avg = 0;
    for (const r of sub.rows) {
      const e = rowExpected(sub, r);
      if (e.city !== name) continue;
      const oo = num(sub.V(r, sub.col.o));
      if (oo === null) okOwn = false; else o += e.price * oo;
      onlyN += e.price * (e.n || 0); price += e.price; cnt++; avg += e.price * e.o;
    }
    own[name] = okOwn ? o : null;
    variants[name] = {"price x location discount only (no discount 1)": onlyN, "sum of prices only": price, "average instead of sum": cnt ? avg / cnt : null};
  }
  let found = 0, scoreSum = 0;
  const notes = [];
  for (const name of ["New York", "St. Louis"]) {
    const [cr, cc] = sub.cityCells[name];
    const cand = [];
    for (const c of range(cc + 1, cc + 6)) { const v = sub.V(cr, c), f = sub.F(cr, c); if (!isBlank(v) || f) cand.push({cell: `${COL(c)}${cr}`, v, f}); }
    const e = exp[name], accepted = [e.without_row2, e.with_row2];
    const a = cand.find(x => num(x.v) !== null) || cand[0] || null;
    const info = {city_cell: `${COL(cc)}${cr}`, expected: [...new Set(accepted.map(show))], answer: null};
    let sc = 0, verdict = "blank";
    if (a) {
      const n = num(a.v);
      info.answer = {cell: a.cell, value: show(a.v), formula: a.f ? short(a.f, 160) : "TYPED (no formula)"};
      const right = n !== null && accepted.some(x => close(n, x));
      const cfw = n !== null && own[name] !== null && close(n, own[name]);
      const fu = funcs(a.f), refs = singleRefs(a.f);
      const hand = a.f && !/:/.test(a.f) && refs.length >= 3 && !fu.some(x => /SUMIF|SUMPRODUCT|SUM$|FILTER|SUMIFS/.test(x));
      if (!a.f) { verdict = "hardcoded"; sc = 0; info.diagnosis = right ? "right number but TYPED (= 0)" : "typed and wrong"; }
      else if (hand) { verdict = "manual"; sc = 0; info.diagnosis = "hand-picked cells (" + refs.slice(0, 6).join("+") + "...) instead of a criteria formula (= 0)"; }
      else if (right) { verdict = "correct"; sc = 10; info.diagnosis = "correct"; }
      else if (cfw) { verdict = "carried_forward"; sc = 9; info.diagnosis = "right formula on the student's own total-discount column O (which is wrong)"; }
      else {
        const v = Object.entries(variants[name]).find(([k, x]) => x !== null && n !== null && close(n, x));
        verdict = v ? "partial" : "major"; sc = v ? 4 : 2;
        info.diagnosis = v ? `typical mistake: ${v[0]}` : "wrong value";
        if (isErr(a.v)) { verdict = "error"; sc = 1; info.diagnosis = "error " + a.v; }
      }
      if (cand.length > 1) info.other_cells = cand.filter(x => x !== a).slice(0, 3).map(x => `${x.cell}: ${show(x.v)} | ${x.f ? short(x.f, 100) : "typed"}`);
    }
    if (verdict !== "correct") {
      // the right value somewhere else in the sheet (outside the data block)?
      for (const [k, cell] of sub.cells) {
        const r = Math.floor(k / 16384), c = k % 16384, n = num(cell.v);
        if (c <= sub.col.city || n === null || (a && a.cell === `${COL(c)}${r}`)) continue;
        if (accepted.some(x => close(n, x)) || (own[name] !== null && close(n, own[name]))) { info.found_elsewhere = `${COL(c)}${r}: ${show(cell.v)} | ${cell.f ? short(cell.f, 120) : "typed"}`; break; }
      }
      if (!a) {
        info.diagnosis = info.found_elsewhere ? "no answer beside the city, but the expected value appears elsewhere" : "not answered";
        if (info.found_elsewhere) { verdict = /typed$/.test(info.found_elsewhere) ? "hardcoded" : "correct"; sc = verdict === "correct" ? 9.5 : 0; }
      }
    }
    info.verdict = verdict; info.points = sc;
    out.cities[name] = info;
    if (verdict !== "blank") found++;
    scoreSum += sc;
  }
  out.own_O_result = Object.fromEntries(Object.entries(own).map(([k, v]) => [k, v === null ? "column O incomplete" : show(v)]));
  out.typical_values = Object.fromEntries(Object.entries(variants).map(([k, v]) => [k, Object.fromEntries(Object.entries(v).map(([a, b]) => [a, show(b)]))]));
  const score = scoreSum / 2;
  const vs = Object.values(out.cities).map(c => c.verdict);
  const verdict = !found ? "blank" : vs.every(v => v === "correct") ? "correct" : vs.includes("hardcoded") ? "hardcoded" : vs.includes("manual") ? "manual" : vs.includes("carried_forward") && score >= 9 ? "carried_forward" : score >= 3 ? "partial" : "major";
  const parts = Object.entries(out.cities).map(([k, c]) => `${k}: ${c.answer ? c.answer.value : "no answer"} (${c.diagnosis}; expected ${c.expected[0]})`);
  const comment = verdict === "correct" ? "Correct." : verdict === "blank" ? "Not answered (put the result for each city in T32 and T33)." : parts.join("; ") + ".";
  out.proposal = draft(10, score, verdict, comment, verdict === "correct" ? "" : FIX[10]);
  return out;
}

function scratchWork(sub) {
  const skip = (r, c) => (r >= 2 && r <= 101 && c <= sub.col.city) || (sub.table && r <= sub.table[0] + 13 && c >= sub.table[1] && c <= sub.table[1] + 4)
    || (c === sub.holAt[1] && r >= sub.holAt[0] && r < sub.holAt[0] + 3) || (c === sub.discAt[1] && r >= sub.discAt[0] && r < sub.discAt[0] + 7);
  const out = [];
  const keys = [...sub.cells.keys()].sort((a, b) => a - b);
  for (const k of keys) {
    const r = Math.floor(k / 16384), c = k % 16384, x = sub.cells.get(k);
    if (!x.f || skip(r, c)) continue;
    const lab = sub.V(r, c - 1);
    out.push({cell: `${COL(c)}${r}`, label: typeof lab === "string" ? lab.slice(0, 40) : "", value: show(x.v), formula: short(x.f, 120)});
    if (out.length >= 25) break;
  }
  const extraCols = [];
  for (let c = sub.col.city + 1; c <= sub.col.city + 2; c++) {
    const n = sub.rows.filter(r => sub.F(r, c)).length;
    if (n >= 10) extraCols.push(`column ${COL(c)}: ${n} formulas, e.g. ${COL(c)}${sub.rows[0]} ${short(sub.F(sub.rows[0], c) || "", 80)}`);
  }
  return {cells: out, helper_columns: extraCols};
}

export function buildEvidence(book, ref, file) {
  const sub = new HSub(book);
  const {student, sid} = canvasName(file);
  const C = sub.col;
  const changed = sub.rows.filter(r => {
    const key = [sub.V(r, C.raw), sub.V(r, C.q1), sub.V(r, C.q2), sub.V(r, C.price), sub.V(r, C.city)].map(x => clean(String(x ?? "")).trim()).join("|");
    const want = ref.data.get(clean(String(sub.V(r, C.raw))).trim());
    return want && want !== key;
  });
  if (changed.length) sub.notes.push(`${changed.length} data row(s) differ from the original data (e.g. row ${changed[0]})`);
  const qs = [];
  for (let q = 1; q <= 9; q++) qs.push(columnQuestion(sub, q));
  qs.push(q10Question(sub, ref));
  return {student, id: sid, file, assignment: ID,
    layout: {notes: sub.notes, columns: Object.fromEntries(Object.entries(C).map(([k, c]) => [k, COL(c)])),
      symbols: `up=${show(sub.upSym)} down=${show(sub.downSym)}`, correction_table: sub.table ? `${COL(sub.table[1])}${sub.table[0]}` : null},
    questions: qs, scratch: scratchWork(sub), _sub: sub};
}
export function proposalTotal(ev) { return Math.round(ev.questions.reduce((a, q, i) => a + q.proposal.score * WEIGHTS[i], 0) * 10 * 100) / 100; }

export function evidenceMd(ev) {
  const L = [`# Homework 1 evidence -- ${ev.student} (${ev.id})`, `file: ${ev.file}`];
  L.push("layout: " + (ev.layout.notes.length ? ev.layout.notes.join("; ") : "template layout"));
  L.push(`symbols table: ${ev.layout.symbols} · correction table header at ${ev.layout.correction_table}`);
  for (const q of ev.questions) {
    const p = q.proposal;
    L.push("", `## Q${q.q} ${q.title} (weight ${Math.round(WEIGHTS[q.q - 1] * 100)}%)   [proposal ${p.score}/10 · ${p.verdict}]`);
    if (q.kind === "column") {
      for (const cs of q.columns) {
        const bits = Object.entries(cs.stats).filter(([k, v]) => v).map(([k, v]) => `${k} ${v}`);
        L.push(`- ${NICE[cs.key]} (rows checked n=${cs.n}): ` + bits.join(", ") + ` · formulas ${cs.formulas}, typed ${cs.typed}`);
        for (const pt of cs.patterns) L.push(`  - formula x${pt.count}: ${pt.cell} ${pt.formula}`);
        if (cs.n_patterns > 3) L.push(`  - (${cs.n_patterns} different formula shapes in the column)`);
        if (Object.keys(cs.offsets).length) L.push(`  - references to another row of the data: offsets ${pyRepr(cs.offsets)}`);
        if (Object.keys(cs.notes).length) L.push(`  - notes: ${pyRepr(cs.notes)}`);
        if (cs.distinct <= 3 && cs.n > 5) L.push(`  - only ${cs.distinct} distinct value(s): ${cs.top_values.join(", ")}`);
        for (const e of cs.examples) L.push(`  - ${e.cell} [${e.status}] got ${e.got} | expected ${e.expected} | ${e.formula}${e.note ? " | " + e.note : ""}`);
      }
      for (const [k, v] of Object.entries(q.extra || {})) L.push(`- ${k}: ${typeof v === "string" ? v : JSON.stringify(v)}`);
    } else {
      for (const [name, c] of Object.entries(q.cities)) {
        L.push(`- ${name} (city cell ${c.city_cell}) expected ${c.expected.join(" or ")} [without / with example row 2]`);
        if (c.answer) L.push(`  - answer ${c.answer.cell}: ${c.answer.value} | ${c.answer.formula}`);
        L.push(`  - diagnosis: ${c.diagnosis} -> ${c.points}/10`);
        if (c.found_elsewhere) L.push(`  - found elsewhere: ${c.found_elsewhere}`);
        for (const o of c.other_cells || []) L.push(`  - other cell beside the city: ${o}`);
      }
      L.push(`- with the student's own column O the result would be: ${JSON.stringify(q.own_O_result)}`);
      L.push(`- typical wrong values: ${JSON.stringify(q.typical_values)}`);
    }
    L.push(`- proposal comment: ${p.comment}`);
  }
  const sc = ev.scratch || {};
  if ((sc.cells || []).length || (sc.helper_columns || []).length) {
    L.push("", "## Other formulas in the sheet (student's side calculations)");
    for (const h of sc.helper_columns || []) L.push(`- ${h}`);
    for (const c of sc.cells || []) L.push(`- ${c.cell}${c.label ? " [" + c.label.trim() + "]" : ""}: ${c.value} | ${c.formula}`);
  }
  return L.join("\n") + "\n";
}
export function expectedSummary(ref) {
  const e = ref.q10;
  return `Expected Q10 values from this class's data: New York ${show(e["New York"].without_row2)} (rows 3-101) or ${show(e["New York"].with_row2)} (rows 2-101); `
    + `St. Louis ${show(e["St. Louis"].without_row2)} or ${show(e["St. Louis"].with_row2)}. Both are accepted. Row-level expected values are in each evidence packet.`;
}

// --------------------------------------------------------------------------- rubric (agent)
export const INSTRUCTIONS = `You are the grader of Homework 1 of BTM-2500, an Excel course at SLU Madrid. You mark like a careful instructor sitting in front of the student's workbook in Excel: you reward correct logic, penalise each mistake once, and tell every student exactly why they lost points and how to fix it.

For each student you receive an evidence packet produced by the grading engine. The engine has already recomputed every expected value of every row (rows 3-101 of sheet 'quiz1#') from that row's own data, compared all 99 rows of each answer column, separated typed values from formulas, detected carried-forward errors (a right formula fed by the student's own wrong earlier column), row offsets, typical mistakes, the conditional formatting, and the student's side calculations. Trust its numbers.

Work in two steps:
1. CHECK THE ANSWERS: start from the engine's proposal for each question and the counts of right / wrong / typed rows.
2. GO INTO THE WORKBOOK: like the Claude add-in inside Excel, use read_cells to open the cells yourself whenever the evidence leaves anything open -- a column with a mix of right and wrong rows, a formula that is right only by coincidence, an answer that is not where expected (look around the question, in columns Q-AD and below row 30), a nearly right formula (wrong reference, missing $, wrong range end). Give PARTIAL CREDIT when the answer is close to good: the right idea with a small slip earns most of the points; a carried-forward error is penalised once, in the question where the mistake was made.

Decide Q1-Q10 with the rubric below. Write \`basis\` for the instructor: one short sentence with the evidence behind the mark. Write \`comment\` for the student (English, 1-2 sentences, cite a cell, say what is wrong). Write \`fix\` = the correction: the formula or step that would make the answer right (empty string when the answer is correct). Return only the JSON object: "questions" is a list with exactly one item per question, q = 1 to 10.

===== RUBRIC =====
`;
export const RUBRIC = `# Homework 1 (Module 1) -- grading rubric

Sheet \`quiz1#\`. Data rows 3-101 (row 2 is the worked example). Each question is marked 0-10; the final grade is the weighted sum (column AA weights) x 10 = /100.

| Q | Weight | What the student must do | Where |
|---|---|---|---|
| 1 | 10% | extract the item ("Item 2") from data_collection | B |
| 2 | 10% | extract the category ("Category A") | C |
| 3 | 5% | extract the date (the number after ':' is a date serial), format m/d/yy | D |
| 4 | 10% | status name from the leading number via the table S6:T8 | E |
| 5 | 10% | due date: 10 working days after the date, with the holidays S12:S14 | F |
| 6 | 5% | total revenue Q1 (quantity Q1 x price) and Q2 (quantity Q2 x price) + highlight the 3 highest of each column with different colours | J, K + conditional formatting |
| 7 | 15% | Q evolution (Q2 - Q1) replaced by the up/down symbols of S17:T18 | L |
| 8 | 5% | Discount 1 = 2% if Q evolution goes down, otherwise 0% | M |
| 9 | 10% | discount by location from the table S23:T28 (exact match) | N (and O = M+N) |
| 10 | 20% | weighted absolute discount = sum of price x total discount of each city, for New York and St. Louis | T32, T33 |

{{EXPECTED}}

## Accepted variants (full marks)
- Q1/Q2: any function mix (MID/FIND, LEFT/RIGHT/LEN, TEXTBEFORE/TEXTAFTER, TEXTSPLIT, SUBSTITUTE) if every row is right. A fixed-position MID that only works for some rows is a real mistake: mark by the share of right rows.
- Q3: a real date with a date format, or TEXT(...,"mm/dd/yy") / "m/d/yy" text. Leaving the 5-digit serial unconverted is wrong (~3/10).
- Q4: VLOOKUP/XLOOKUP/INDEX-MATCH/nested IF/IFS/SWITCH/CHOOSE are all fine. Returning the code instead of the name, or reading the row above, is wrong.
- Q5: WORKDAY or WORKDAY.INTL. The holidays only change row 2, so check the formula: holidays missing -2, holiday range relative (no $) -1. Calendar days (D+10, D+14) are wrong (~2/10).
- Q6: 70% of the mark for J and K values, 30% for the highlight. Conditional formatting Top 10 -> 3 (or a LARGE(...,3) rule) on both columns = full; one column only or manual colouring = part; same colour in both columns -0.5.
- Q7: up/down taken from S18/T18 or symbols typed inside the IF are fine; K-J shown with an icon set / arrow number format is fine. Only the difference without any symbol = max 5. Ties (K = J) may be up, down or equal.
- Q8: 2% or 0.02 both fine; testing L or comparing K<J both fine. 2 instead of 2% = -2.
- Q9: VLOOKUP needs FALSE (exact match): without it the values are wrong for an unsorted table -- mark by the share of right rows. XLOOKUP/INDEX-MATCH fine. Column O (M+N) feeds Q10; a wrong O is penalised in Q10 only if Q10 uses it and it is wrong because of Q9 (then it is carried forward).
- Q10: SUMPRODUCT, SUMIF/SUMIFS over a helper column of I*O, or FILTER are fine; rows 3-101 or 2-101 both accepted. 5 points per city.

## Deductions
- **Typed by hand (no formula) = 0** for those cells, even if the number is right.
- **Hand-picked cells** (e.g. =I3*O3+I4*O4+... for one city, or sorting the data and summing a block) **= 0**.
- Carried-forward error (right method on the student's own wrong column): 9/10 in the later question; the mistake is penalised where it was made.
- A formula that reads another row (row offset, e.g. E4 =LEFT(A3,1)) is a real mistake even if a few rows match by chance.
- Small slips (missing $ that does not change values, extra spaces, one-off typo) -0.5 to -1. A right idea with a wrong range end or one wrong argument: 5-8. Wrong method: 0-3.
- #VALUE!/#N/A in most rows: 0-2 unless the evidence shows the idea is right (then up to 4).

## Comment rules
- \`comment\`: English, 1-2 sentences, specific (cite one cell and what it shows vs what was expected). For a correct answer: "Correct." plus at most a short tip.
- \`fix\`: the corrected formula for row 3 (or the step, e.g. "Home > Conditional Formatting > Top 10 Items -> 3"), empty when correct.
- Never say "the engine". Speak to the student.
`;

// --------------------------------------------------------------------------- writer
// final: {1..10: {score, verdict, comment, fix, basis}, summary}
export async function writeWorkbook(JSZip, data, sub, final) {
  const zip = await JSZip.loadAsync(data);
  const part = sub.book.sheetPath;
  const sx = new Sheet(await zip.file(part).async("string"));
  let tr, tc, made = false;
  if (sub.table) [tr, tc] = sub.table;
  else {
    made = true; tr = 1; tc = 25 + sub.shift;
    sx.setText(tr, tc, "Question"); sx.setText(tr, tc + 1, "Grade"); sx.setText(tr, tc + 2, "Weight");
    for (let i = 1; i <= NQ; i++) sx.setNumber(tr + i, tc, i);
    sx.setText(tr + NQ + 1, tc, "final grade");
    sx.setFormula(tr + NQ + 1, tc + 1, `SUMPRODUCT(${COL(tc + 1)}${tr + 1}:${COL(tc + 1)}${tr + NQ},${COL(tc + 2)}${tr + 1}:${COL(tc + 2)}${tr + NQ})`, 0);
  }
  const pc = tc + 1, wc = tc + 2;
  const filled = (r, c) => { const v = sub.V(r, c); return (v !== null && v !== "") || !!sub.F(r, c); };
  // weights: restore the official ones if the student changed them
  let weightsOk = true;
  for (let i = 0; i < NQ; i++) { const w = num(sub.V(tr + 1 + i, wc)); if (w === null || Math.abs(w - WEIGHTS[i]) > 1e-9) weightsOk = false; }
  if (!weightsOk || made) for (let i = 0; i < NQ; i++) sx.setNumber(tr + 1 + i, wc, WEIGHTS[i], made ? null : sx.style(tr + 1 + i, wc));
  // comment + fix columns: the first two free columns right of the weights
  let cc = wc + 1;
  while (range(tr, tr + NQ + 3).some(r => filled(r, cc) || filled(r, cc + 1))) cc++;
  const fc = cc + 1;
  const scores = Array.from({length: NQ}, (_, i) => final[i + 1].score);
  const grade10 = scores.reduce((a, s, i) => a + s * WEIGHTS[i], 0);
  const grade100 = Math.round(grade10 * 10 * 100) / 100;
  const st = addFillStyles(await zip.file("xl/styles.xml").async("string"), sx.style(tr + 1, pc));
  const [sg, sa, sr] = st.ids;
  const hs = sx.style(tr, pc);
  sx.setText(tr, cc, "Comment", hs); sx.setText(tr, fc, "Correction (how to fix it)", hs);
  for (let i = 1; i <= NQ; i++) {
    const s = final[i].score;
    sx.setNumber(tr + i, pc, s, s >= 9 ? sg : s < 3 ? sr : sa);
    sx.setText(tr + i, cc, final[i].comment);
    sx.setText(tr + i, fc, final[i].fix || "");
  }
  const g = tr + NQ + 1;
  if (!sx.updateCached(g, pc, grade10)) sx.setFormula(g, pc, `SUMPRODUCT(${COL(pc)}${tr + 1}:${COL(pc)}${tr + NQ},${COL(wc)}${tr + 1}:${COL(wc)}${tr + NQ})`, grade10);
  const P = COL(pc);
  const free = !filled(g + 1, tc) && !filled(g + 1, pc);
  if (free) { sx.setText(g + 1, tc, "GRADE /100", sx.style(g, tc)); sx.setFormula(g + 1, pc, `ROUND(${P}${g}*10,2)`, grade100, sx.style(g, pc)); }
  if (final.summary && !filled(g + 1, cc)) sx.setText(g + 1, cc, "Overall: " + final.summary);
  zip.file(part, sx.finish());
  zip.file("xl/styles.xml", st.xml);
  if (sx.droppedFormula && zip.file("xl/calcChain.xml")) {
    zip.remove("xl/calcChain.xml");
    const ct = await zip.file("[Content_Types].xml").async("string");
    zip.file("[Content_Types].xml", ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ""));
    const rl = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    zip.file("xl/_rels/workbook.xml.rels", rl.replace(/<Relationship[^>]*calcChain[^>]*\/>/, ""));
  }
  const bytes = await zip.generateAsync({type: "uint8array", compression: "DEFLATE"});
  return {bytes, info: {grade100, table: `${COL(tc)}${tr}`, points_col: P, comment_col: COL(cc), fix_col: COL(fc),
    grade_cell: free ? `${P}${g + 1}` : `${P}${g}`, weights_ok: weightsOk, created_table: made}};
}

export function grade(final) { return Math.round(Array.from({length: NQ}, (_, i) => final[i + 1].score * WEIGHTS[i]).reduce((a, b) => a + b, 0) * 10 * 100) / 100; }

export function justificationMd(ev, final, source, info, g, meta = {}) {
  const md = evidenceMd(ev), sections = {};
  for (const blk of md.split(/\n(?=## Q\d+ )/)) { const m = /^## Q(\d+) /.exec(blk); if (m) sections[+m[1]] = blk; }
  const L = [`# Homework 1 -- justification of the grade -- ${ev.student} (${ev.id})`, "", `- file: ${ev.file}`,
    `- **grade: ${g.toFixed(2)} / 100** (in the workbook: ${info.grade_cell}; marks in column ${info.points_col}, comments in ${info.comment_col}, corrections in ${info.fix_col || "-"})`,
    `- decided by: ${source}` + (meta.model ? ` -- ${meta.model} on ${meta.graded_at}` : ""),
    `- layout: ${ev.layout.notes.join("; ") || "template layout"}`, ""];
  if (final.summary) L.push(`**Overall:** ${final.summary}`, "");
  L.push("| Q | weight | mark /10 | verdict | told to the student | correction |", "|---|---|---|---|---|---|");
  for (let q = 1; q <= NQ; q++) L.push(`| ${q} | ${Math.round(WEIGHTS[q - 1] * 100)}% | ${final[q].score} | ${final[q].verdict || ""} | ${final[q].comment.replace(/\|/g, "/")} | ${(final[q].fix || "").replace(/\|/g, "/")} |`);
  for (let q = 1; q <= NQ; q++) {
    const f = final[q];
    L.push("", `## Q${q} -- ${f.score}/10 (${f.verdict || ""})`, `- **told to the student:** ${f.comment}`);
    if (f.fix) L.push(`- **correction:** ${f.fix}`);
    if (f.basis) L.push(`- **basis (grader):** ${f.basis}`);
    for (const c of (meta.calibration || []).filter(c => c.question === q)) L.push(`- **calibrated** across the class: ${c.old} -> ${c.new}`);
    L.push("- **engine evidence:**", "", "```", (sections[q] || "").trim(), "```");
  }
  return L.join("\n") + "\n";
}
