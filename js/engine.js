// engine.js -- the Quiz 1 evidence engine (JavaScript port of _tools/quiz1_agent.py).
// Recomputes the right answer of every cell from the raw import text, classifies every
// answer (ok / cosmetic / format / carried forward / row offset / typical mistake / typed /
// wrong / blank / #error), finds answers placed elsewhere, traces references, matches
// typical mistakes and writes a rule-based proposal per question.  No answers are stored
// here: they are computed from the submissions themselves.
import {XDate, excelDate, colIndex as CIDX, colLetter as COL} from "./xlsx.js?v=2";

export const SHEET = "Data&Q";
export const EXAMPLE = 4, FIRST = 5, LAST = 183;
const QTEXT_COL = 16;
const HOLIDAY_ROWS = [18, 19, 20, 21, 22, 23];
const LOOKUP_ROWS = [11, 12, 13, 14, 15];
const POS = {C: "Center", PG: "Point Guard", SG: "Shooting Guard", SF: "Small Forward", PF: "Power Forward"};
const HEAD = [["raw", "raw_import"], ["date", "date"], ["day", "day"], ["month", "month"],
  ["year", "year"], ["player", "player"], ["team", "team"], ["position", "position"],
  ["home", "home/away"], ["next", "next math"], ["result", "result"], ["win", "win"],
  ["tickets", "tickets_sold"], ["price", "avg_ticket_price"]];
const NICE = {date: "Date", day: "day", month: "month", year: "year", player: "player", team: "Team",
  position: "position", next: "next match", result: "Result", win: "Win"};
const UPSTREAM = {day: ["date", 1], month: ["date", 1], year: ["date", 1], team: ["player", 3],
  next: ["date", 1], win: ["result", 7]};
export const ROWQ = {1: ["Match date (format of row 4: mm/dd/yy)", ["date"]],
  2: ["Day / month / year", ["day", "month", "year"]], 3: ["Player", ["player"]],
  4: ["Team (no blank spaces)", ["team"]],
  5: ["Position name via the lookup table (add the missing code)", ["position"]],
  6: ["Next match date: 30 working days, skip weekends + holidays", ["next"]],
  7: ["Result (W/L) and Win flag", ["result", "win"]]};
export const AGG = {
  8: {title: "Total sales of all matches ($)", text_row: 26, ans_row: 28, money: true},
  9: {title: "Total sales for Hawks at home ($)", text_row: 30, ans_row: 31, money: true},
  10: {title: "Hawks sales, only matches they won at home ($)", text_row: 33, ans_row: 34, money: true},
  11: {title: "Tickets sold by Point Guards in Away matches", text_row: 36, ans_row: 37},
  12: {title: "Bulls sales in the matches they lost ($)", text_row: 39, ans_row: 40, money: true},
  13: {title: "Did the Hawks sell more tickets at Home than Away? (Yes/No)", text_row: 42, ans_row: 43, text: true},
  14: {title: "Rows with Tickets_Sold > 15,000 OR Avg_Ticket_Price > 50", text_row: 45, ans_row: 46}};
const QRE = /^\s*(\d{1,2})\s*\.-/;
const RAW = /^(\d{4})-(\d{2})-(\d{2}) \| (\S+) vs (\S+) @ (Home|Away) \| Player: (\S+) \((\w+)\) \|.*\| ([WL])\s*$/;
const ZW = /[\u200b\u200c\u200d\ufeff\u00a0]/g;
const REF = /(?<![A-Za-z_"!.])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])/g;
const RANGE = /\$?([A-Z]{1,3})\$?(\d{1,7}):\$?([A-Z]{1,3})\$?(\d{1,7})/g;
const CARRIED_FORWARD = 9.0, ROW_OFFSET = 7.0, TYPED_FACTOR = 0.0, MANUAL_RANGE = 0.0,
  ONE_MISTAKE = 5.0, MONEY_PENALTY = 0.5;
const range = (a, b) => Array.from({length: b - a + 1}, (_, i) => a + i);   // inclusive

// --------------------------------------------------------------------------- values
export function serial(v) {
  if (typeof v === "boolean" || v === null || v === undefined) return null;
  if (v instanceof XDate) return v.serial;
  if (typeof v === "number") return v;
  return null;
}
const dayOf = s => excelDate(Math.trunc(s));
const pad = n => String(n).padStart(2, "0");
const fmtDate = (d, f) => f.replace("%m", pad(d.getUTCMonth() + 1)).replace("%d", pad(d.getUTCDate()))
  .replace("%Y", String(d.getUTCFullYear())).replace("%y", pad(d.getUTCFullYear() % 100));
export const clean = v => typeof v === "string" ? v.replace(ZW, "") : v;
export function num(v) {
  if (typeof v === "boolean" || v === null || v === undefined || v instanceof XDate) return null;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const t = clean(v).trim().replace(/[$, ]/g, "");
    return /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t) ? parseFloat(t) : null;
  }
  return null;
}
function serialOfYMD(y, m, d) { return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; }
function parseDateText(t) {
  t = String(clean(t)).trim();
  let m;
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/.exec(t))) return valid(2000 + (+m[3] < 69 ? +m[3] : +m[3] - 100), +m[1], +m[2]);
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t))) return valid(+m[3], +m[1], +m[2]);
  if ((m = /^(\d{4})-(\d{2})-(\d{2})( 00:00:00)?$/.exec(t))) return valid(+m[1], +m[2], +m[3]);
  return null;
  function valid(y, mo, d) {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d));
    return dt.getUTCMonth() === mo - 1 ? serialOfYMD(y, mo, d) : null;
  }
}
const asSerial = v => serial(v) ?? (typeof v === "string" ? parseDateText(v) : null);
export const close = (a, b) => a !== null && b !== null && a !== undefined && b !== undefined && Math.abs(a - b) <= Math.max(0.011, 1e-9 * Math.abs(b));
const group = n => n.toLocaleString("en-US");
export function pyRepr(v) {
  if (v === null || v === undefined) return "None";
  if (typeof v === "boolean") return v ? "True" : "False";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  if (v instanceof XDate) return `datetime(${v})`;
  if (typeof v === "string") {
    const q = v.includes("'") && !v.includes('"') ? '"' : "'";
    let s = v.replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/\r/g, "\\r").replace(/\t/g, "\\t");
    if (q === "'") s = s.replace(/'/g, "\\'");
    s = s.replace(/[\u0000-\u001f\u007f-\u00a0\u00ad]/g, c => "\\x" + c.charCodeAt(0).toString(16).padStart(2, "0"))
      .replace(/[\u2000-\u200f\u2028-\u202f\u205f-\u206f\ufeff]/g, c => "\\u" + c.charCodeAt(0).toString(16).padStart(4, "0"));
    return q + s + q;
  }
  if (Array.isArray(v)) return "[" + v.map(pyRepr).join(", ") + "]";
  if (typeof v === "object") return "{" + Object.entries(v).map(([k, x]) => `${/^-?\d+$/.test(k) ? k : pyRepr(k)}: ${pyRepr(x)}`).join(", ") + "}";
  return String(v);
}
export function show(v) {
  if (v === null || v === undefined || v === "") return "(blank)";
  if (v instanceof XDate) return fmtDate(dayOf(v.serial), "%m/%d/%Y") + " (date)";
  if (typeof v === "number") return Number.isInteger(v) ? group(v)
    : v.toLocaleString("en-US", {minimumFractionDigits: 2, maximumFractionDigits: 2});
  if (typeof v === "string") return pyRepr(clean(v));
  if (typeof v === "boolean") return v ? "True" : "False";
  return String(v);
}
function workday(start, n, hol) {
  let d = Math.trunc(start), k = 0;
  while (k < n) {
    d += 1;
    const wd = excelDate(d).getUTCDay();
    if (wd >= 1 && wd <= 5 && !hol.has(d)) k += 1;
  }
  return d;
}
const ERRS = new Set(["VALUE", "N/A", "NAME", "REF", "DIV/0", "NUM", "NULL", "SPILL", "CALC"]);
export const isErr = v => typeof v === "string" && v.startsWith("#") && ERRS.has(v.replace(/[!?]+$/, "").replace(/^#/, "").toUpperCase());

// --------------------------------------------------------------------------- formulas
export function funcs(f) {
  return [...new Set([...(f || "").toUpperCase().matchAll(/([A-Z_][A-Z0-9_.]*)\(/g)]
    .map(m => m[1].replace("_XLFN.", "").replace("_XLWS.", "")))].sort();
}
function r1c1(f, row, col) {
  return (f || "").replace(REF, (m, ca, c, ra, r) => {
    const cc = CIDX(c), rr = +r;
    return (ra ? `R${rr}` : `R[${rr - row}]`) + (ca ? `C${cc}` : `C[${cc - col}]`);
  });
}
export function short(f, n = 150) {
  f = (f || "").replaceAll("_xlfn.", "").replaceAll("_xlws.", "");
  return f.length <= n ? f : f.slice(0, n - 3) + "...";
}
class Counter extends Map {
  inc(k, n = 1) { this.set(k, (this.get(k) || 0) + n); }
  g(k) { return this.get(k) || 0; }
  obj() { return Object.fromEntries(this); }
  mostCommon(n) { return [...this.entries()].map((e, i) => [e, i]).sort((a, b) => b[0][1] - a[0][1] || a[1] - b[1]).slice(0, n).map(x => x[0]); }
}

// --------------------------------------------------------------------------- reference
// Built from the submissions: the raw import text, tickets, prices and holidays are the
// given data of the quiz, identical in every file.
export function buildReference(books) {
  const holSets = new Map();
  let base = null;
  for (const b of books) {
    let s;
    try { s = new Sub(b, null); } catch (e) { continue; }
    const rows = {};
    let ok = true;
    for (const r of range(EXAMPLE, LAST)) {
      const a = clean(s.V(r, s.col.raw));
      const m = typeof a === "string" ? RAW.exec(a) : null;
      if (!m) { ok = false; break; }
      rows[r] = {a, m, tickets: num(s.V(r, s.col.tickets)), price: num(s.V(r, s.col.price))};
    }
    if (!ok) continue;
    const hol = HOLIDAY_ROWS.map(r => asSerial(s.V(r + s.qrowOff, s.qcol + 1))).filter(x => x !== null);
    if (hol.length === HOLIDAY_ROWS.length) {
      const key = hol.join(",");
      holSets.set(key, (holSets.get(key) || 0) + 1);
    }
    if (!base && rows[FIRST].tickets !== null) base = rows;
  }
  if (!base || !holSets.size) throw new Error("no submission keeps the original data (raw import, tickets, prices, holidays)");
  const holKey = [...holSets.entries()].sort((a, b) => b[1] - a[1])[0][0];
  const hol = new Set(holKey.split(",").map(Number));
  const rows = {};
  for (const r of range(EXAMPLE, LAST)) {
    const {a, m, tickets, price} = base[r];
    const [, y, mo, d, , , ha, player, code, res] = m;
    const ds = serialOfYMD(+y, +mo, +d);
    rows[r] = {raw: a, date: ds, day: +d, month: +mo, year: +y, player, team: player.split("-")[0],
      code, position: POS[code] ?? code, home: ha, next: workday(ds, 30, hol), result: res,
      win: res === "W" ? 1 : 0, tickets, price};
  }
  return {rows, hol, agg: aggregates(rows)};
}

export function aggregates(rows, starts = [EXAMPLE, FIRST]) {
  const eq = (a, b) => typeof a === "string" && a.toLowerCase() === b.toLowerCase();
  const S = (start, cond, val) => {
    let tot = 0;
    for (const r of range(start, LAST)) { const x = rows[r]; if (x && cond(x)) tot += val(x) || 0; }
    return Math.round(tot * 1e6) / 1e6;
  };
  const C = (start, cond) => range(start, LAST).filter(r => rows[r] && cond(rows[r])).length;
  const sales = x => (x.tickets || 0) * (x.price || 0), tk = x => x.tickets || 0, pr = x => x.price || 0;
  const T = () => true, hawks = x => eq(x.team, "Hawks"), bulls = x => eq(x.team, "Bulls");
  const home = x => eq(x.home, "Home"), away = x => eq(x.home, "Away");
  const won = x => eq(x.result, "W"), lost = x => eq(x.result, "L"), pg = x => eq(x.position, "Point Guard");
  const big = x => (x.tickets || 0) > 15000, dear = x => (x.price || 0) > 50;
  const out = {};
  const spec = [
    [8, s => S(s, T, sales), [
      ["sums only Tickets_Sold (no price)", s => S(s, T, tk)],
      ["adds the two columns (SUM(M:N)) instead of multiplying them", s => S(s, T, tk) + S(s, T, pr)],
      ["total tickets x average price", s => S(s, T, tk) * (S(s, T, pr) / Math.max(1, C(s, T)))]]],
    [9, s => S(s, x => hawks(x) && home(x), sales), [
      ["no Home filter (all Hawks matches)", s => S(s, hawks, sales)],
      ["Away instead of Home", s => S(s, x => hawks(x) && away(x), sales)],
      ["no team filter (all home matches)", s => S(s, home, sales)],
      ["tickets instead of sales (no price)", s => S(s, x => hawks(x) && home(x), tk)],
      ["also filters on wins", s => S(s, x => hawks(x) && home(x) && won(x), sales)]]],
    [10, s => S(s, x => hawks(x) && home(x) && won(x), sales), [
      ["no Win filter (= all Hawks home sales)", s => S(s, x => hawks(x) && home(x), sales)],
      ["no Home filter (Hawks wins anywhere)", s => S(s, x => hawks(x) && won(x), sales)],
      ["lost instead of won", s => S(s, x => hawks(x) && home(x) && lost(x), sales)],
      ["Away instead of Home", s => S(s, x => hawks(x) && away(x) && won(x), sales)],
      ["tickets instead of sales", s => S(s, x => hawks(x) && home(x) && won(x), tk)]]],
    [11, s => S(s, x => pg(x) && away(x), tk), [
      ["Home instead of Away", s => S(s, x => pg(x) && home(x), tk)],
      ["no Away filter (all Point Guard tickets)", s => S(s, pg, tk)],
      ["no position filter (all away tickets)", s => S(s, away, tk)],
      ["sales instead of tickets", s => S(s, x => pg(x) && away(x), sales)]]],
    [12, s => S(s, x => bulls(x) && lost(x), sales), [
      ["won instead of lost", s => S(s, x => bulls(x) && won(x), sales)],
      ["no result filter (all Bulls matches)", s => S(s, bulls, sales)],
      ["only home matches", s => S(s, x => bulls(x) && lost(x) && home(x), sales)],
      ["tickets instead of sales", s => S(s, x => bulls(x) && lost(x), tk)]]],
    [14, s => C(s, x => big(x) || dear(x)), [
      ["AND instead of OR (both conditions at once)", s => C(s, x => big(x) && dear(x))],
      ["double counts rows meeting both conditions (COUNTIF+COUNTIF without subtracting the overlap)", s => C(s, big) + C(s, dear)],
      ["only the tickets condition", s => C(s, big)],
      ["only the price condition", s => C(s, dear)],
      [">= instead of >", s => C(s, x => (x.tickets || 0) >= 15000 || (x.price || 0) >= 50)]]],
  ];
  for (const [q, accFn, variants] of spec) {
    const acc = [];
    for (const s of starts) { const v = accFn(s); if (!acc.some(a => close(v, a))) acc.push(v); }
    out[q] = {accepted: acc, variants: variants.map(([name, fn]) => [name, starts.map(fn)])};
  }
  const th = S(EXAMPLE, x => hawks(x) && home(x), tk), ta = S(EXAMPLE, x => hawks(x) && away(x), tk);
  const sh = S(EXAMPLE, x => hawks(x) && home(x), sales), sa = S(EXAMPLE, x => hawks(x) && away(x), sales);
  out[13] = {accepted: [th > ta ? "Yes" : "No"], by_sales: sh > sa ? "Yes" : "No", home: th, away: ta};
  return out;
}

// --------------------------------------------------------------------------- submission
export function canvasName(file) {
  const m = /^([^_]+)_(?:LATE_)?(\d+)_(\d+)_(.*)$/.exec(file);
  return m ? {student: m[1], sid: m[2]} : {student: file.replace(/\.xlsx$/i, ""), sid: ""};
}

export class Sub {
  constructor(book, ref) {
    this.book = book; this.ref = ref; this.notes = [];
    this.cells = book.cells;
    this.mapColumns(); this.findQuestions(); this.findTable();
  }
  cell(r, c) { return this.cells.get(r * 16384 + c); }
  F(r, c) { const x = this.cell(r, c); return x ? x.f : null; }
  V(r, c) { const x = this.cell(r, c); return x ? x.v : null; }
  NF(r, c) { const x = this.cell(r, c); return x ? x.nf : "General"; }
  fv(r, c) { const x = this.cell(r, c); if (!x) return null; return x.f !== null ? x.f : x.v; }  // openpyxl formula view
  mapColumns() {
    let hdr = 3;
    for (const r of range(1, 6)) {
      if (range(1, 3).some(c => typeof this.fv(r, c) === "string" && this.fv(r, c).trim().toLowerCase() === "raw_import")) { hdr = r; break; }
    }
    const names = HEAD.map(h => h[1]);
    this.col = {}; this.inserted = [];
    let t = 0;
    for (const c of range(1, 39)) {
      if (t >= HEAD.length) break;
      const v = this.fv(hdr, c);
      const h = typeof v === "string" ? v.trim().toLowerCase() : null;
      if (h === names[t]) { this.col[HEAD[t][0]] = c; t++; }
      else if (h !== null && names.slice(t + 1).includes(h)) { const k = names.indexOf(h, t + 1); this.col[HEAD[k][0]] = c; t = k + 1; }
      else this.inserted.push(c);
    }
    HEAD.forEach(([k, h], i) => {
      if (!(k in this.col)) {
        this.col[k] = i + 1 + this.inserted.filter(c => c <= i + 1).length;
        this.notes.push(`header '${h}' not found; assumed column ${COL(this.col[k])}`);
      }
    });
    if (this.inserted.length) this.notes.push("inserted column(s) " + this.inserted.map(COL).join(", ") + " inside the data block -- answers were read from where the student actually put them");
    if (this.ref) {
      const bad = range(EXAMPLE, LAST).filter(r => clean(this.V(r, this.col.raw)) !== this.ref.rows[r].raw);
      if (bad.length) this.notes.push(`${bad.length} rows of Raw_Import differ from the original data (e.g. row ${bad[0]}) -- rows may have been moved/edited`);
    }
  }
  findQuestions() {
    this.qpos = {};
    for (const r of range(1, 69)) for (const c of range(1, 44)) {
      const v = this.fv(r, c);
      if (typeof v === "string") { const m = QRE.exec(v); if (m && !(+m[1] in this.qpos)) this.qpos[+m[1]] = [r, c]; }
    }
    if (1 in this.qpos) { this.qcol = this.qpos[1][1]; this.qrowOff = this.qpos[1][0] - 5; }
    else { this.qcol = QTEXT_COL + this.inserted.length; this.qrowOff = 0; this.notes.push("question texts not found; assumed template positions"); }
  }
  findTable() {
    this.table = null;
    outer: for (const r of range(1, 39)) for (const c of range(17, 49)) {
      const v = this.fv(r, c);
      if (typeof v === "string" && v.trim().toLowerCase() === "question") { this.table = [r, c]; break outer; }
    }
  }
}

// --------------------------------------------------------------------------- per-row checks
function judge(kind, v, nf, e, qnum) {
  if (v === null || v === undefined || (typeof v === "string" && clean(v).trim() === "")) return ["blank", ""];
  if (isErr(v)) return ["error", String(v)];
  if (kind === "date") {
    const s = serial(v), d = dayOf(e);
    if (s !== null) {
      if (Math.trunc(s) !== Math.trunc(e)) return ["wrong", ""];
      const fmt = (nf || "General").toLowerCase().replaceAll("\\", "");
      if (fmt === "general" || ![..."dmy"].some(ch => fmt.includes(ch))) return ["cosmetic", "date shown as a serial number (no date format)"];
      const fd = fmt.indexOf("d"), fm = fmt.indexOf("m");
      if (fd >= 0 && fd < fm) return ["format", "displayed day-first (dd/mm), not mm/dd/yy as in row 4"];
      return ["ok", ""];
    }
    const t = String(clean(v)).trim();
    if (t === fmtDate(d, "%m/%d/%y")) return ["ok", ""];
    const M = d.getUTCMonth() + 1, D = d.getUTCDate(), Y = d.getUTCFullYear();
    if ([fmtDate(d, "%m/%d/%Y"), `${M}/${D}/${Y}`, `${M}/${D}/${pad(Y % 100)}`, fmtDate(d, "%Y-%m-%d")].includes(t)) return ["cosmetic", `written as ${pyRepr(t)}, row 4 uses mm/dd/yy`];
    if (D !== M && [fmtDate(d, "%d/%m/%y"), fmtDate(d, "%d/%m/%Y")].includes(t)) return ["format", `day/month swapped (${pyRepr(t)} is dd/mm, row 4 uses mm/dd/yy)`];
    return ["wrong", ""];
  }
  if (kind === "num") {
    const n = num(v);
    if (n !== null && close(n, +e)) return typeof v === "string" ? ["cosmetic", "number stored as text"] : ["ok", ""];
    return ["wrong", ""];
  }
  const t = String(clean(v));
  if (t === e) return ["ok", ""];
  if (t.trim() === e) return qnum === 4 ? ["spaces", "extra blank spaces (the question asks for none)"] : ["ok", ""];
  if (t.trim().toLowerCase() === e.toLowerCase()) return ["cosmetic", "different upper/lower case"];
  return ["wrong", ""];
}
const KIND = {date: "date", next: "date", day: "num", month: "num", year: "num", win: "num",
  player: "text", team: "text", position: "text", result: "text"};
const OKISH = new Set(["ok", "cosmetic", "format", "spaces"]);

function ecfExpected(sub, ref, key, r) {
  const up = UPSTREAM[key];
  if (!up) return null;
  const uv = sub.V(r, sub.col[up[0]]);
  if (["day", "month", "year", "next"].includes(key)) {
    const s = asSerial(uv);
    if (s === null) return null;
    const d = dayOf(s);
    return key === "next" ? workday(s, 30, ref.hol) : {day: d.getUTCDate(), month: d.getUTCMonth() + 1, year: d.getUTCFullYear()}[key];
  }
  if (key === "team") { const t = uv !== null && uv !== undefined ? String(clean(uv)) : ""; return t.includes("-") ? t.split("-")[0].trim() : null; }
  if (key === "win") { const k = uv !== null && uv !== undefined ? String(clean(uv)).trim().toUpperCase() : ""; return {W: 1, L: 0}[k] ?? null; }
  return null;
}

function columnStats(sub, ref, key, qnum) {
  const c = sub.col[key], kind = KIND[key], rows = ref.rows;
  const st = new Counter(), notes = new Counter(), offsets = new Counter(), fn = new Counter(), variant = new Counter(), seen = new Counter();
  const wrongEx = [], pats = new Map();
  for (const r of range(FIRST, LAST)) {
    const v = sub.V(r, c), f = sub.F(r, c), nf = sub.NF(r, c);
    seen.inc(show(v));
    const e = rows[r][key];
    let [status, note] = judge(kind, v, nf, e, qnum);
    if (f) {
      const p = r1c1(f, r, c);
      if (!pats.has(p)) pats.set(p, [0, `${COL(c)}${r}`, f]);
      pats.get(p)[0]++;
      for (const name of funcs(f)) fn.inc(name);
      for (const m of f.replaceAll("$", "").matchAll(REF)) if (CIDX(m[2]) === sub.col.raw) offsets.inc(+m[4] - r);
    } else if (status !== "blank") {
      st.inc("typed");
      if (OKISH.has(status)) { st.inc("typed_ok"); status = "typed_value"; }   // typed by hand: never counts
    }
    if (!["ok", "blank", "error", "typed_value"].includes(status)) {
      if (key === "win" && status === "wrong" && num(v) !== null && close(num(v), 1 - e)) {
        status = "alt"; note = "1 = loss, 0 = win (inverted, as in the row-4 example)";
      } else if (status === "wrong") {
        const ee = ecfExpected(sub, ref, key, r);
        if (ee !== null && OKISH.has(judge(kind, v, nf, ee, qnum)[0])) status = "ecf";
        else {
          for (const dr of [-1, 1]) {
            if (r + dr >= FIRST - 1 && r + dr <= LAST && rows[r + dr] && OKISH.has(judge(kind, v, nf, rows[r + dr][key], qnum)[0])) { status = `shift${dr > 0 ? "+" : ""}${dr}`; break; }
          }
          if (status === "wrong" && key === "next") {
            const s = asSerial(v);
            if (s !== null) {
              if (close(s, workday(rows[r].date, 30, new Set()))) variant.inc("WORKDAY without the holidays");
              else if (close(s, rows[r].date + 30)) variant.inc("30 calendar days (weekends not skipped)");
            }
          }
        }
      }
    }
    st.inc(status);
    if (note) notes.inc(note);
    if (["wrong", "error", "ecf"].includes(status) || status.startsWith("shift") || (["format", "spaces", "cosmetic"].includes(status) && wrongEx.length < 2)) {
      if (wrongEx.length < 6) {
        const expShow = kind === "date" ? fmtDate(dayOf(e), "%m/%d/%y") : e;
        wrongEx.push({cell: `${COL(c)}${r}`, formula: f ? short(f) : "typed", got: show(v), expected: show(expShow), status});
      }
    }
  }
  const n = LAST - FIRST + 1;
  st.set("n", n);
  const top = [...pats.values()].map((x, i) => [x, i]).sort((a, b) => b[0][0] - a[0][0] || a[1] - b[1]).slice(0, 3).map(x => x[0]);
  return {key, column: COL(c), stats: st.obj(), notes: notes.obj(), examples: wrongEx,
    patterns: top.map(([k, cell, f]) => ({count: k, cell, formula: short(f, 200)})),
    n_patterns: pats.size, raw_offsets: offsets.obj(), functions: fn.obj(), variants: variant.obj(),
    distinct_values: seen.size, top_values: Object.fromEntries(seen.mostCommon(3))};
}

function goodOf(cs) {
  const s = cs.stats;
  if (cs.key === "win") return Math.max(s.ok || 0, s.alt || 0) + (s.cosmetic || 0);
  return (s.ok || 0) + (s.cosmetic || 0);
}
const pct0 = x => Math.round(x * 100) + "%";
function scoreColumn(cs) {
  const s = cs.stats, n = s.n, g = k => s[k] || 0;
  const shift = Object.entries(s).filter(([k]) => k.startsWith("shift")).reduce((a, [, v]) => a + v, 0);
  const good = goodOf(cs);
  if (g("blank") === n) return [0, "blank", "not answered"];
  if (g("typed") > n / 2) return [Math.round(TYPED_FACTOR * (g("typed_ok") / n) * 2) / 2, "hardcoded", `${g("typed")}/${n} cells typed by hand, not formulas`];
  let acc = good / n;
  if (acc >= 0.98) return g("cosmetic") > n / 2 ? [9.5, "minor", "right values, presentation issue"] : [10, "correct", "all rows right"];
  if ((good + g("format")) / n >= 0.98) return [7, "minor", "required format not followed"];
  if ((good + g("spaces")) / n >= 0.98) return [7, "minor", "values keep blank spaces"];
  if ((good + g("ecf")) / n >= 0.98) return [CARRIED_FORWARD, "carried_forward", "own formula right; upstream column wrong"];
  if ((good + shift) / n >= 0.95 && shift >= 0.05 * n) return [ROW_OFFSET, "partial", "formula refers to the neighbouring row"];
  const vsum = Object.values(cs.variants).reduce((a, b) => a + b, 0);
  if (Object.keys(cs.variants).length && (good + vsum) / n >= 0.9) return [6, "partial", "one systematic mistake: " + Object.keys(cs.variants).join(", ")];
  acc = (good + g("ecf") + g("format") + g("spaces")) / n;
  for (const [lim, sc] of [[0.9, 8], [0.75, 7], [0.5, 5], [0.25, 3], [0.0001, 2]]) if (acc >= lim) return [sc, "partial", `${pct0(acc)} of rows right`];
  const attempted = cs.patterns.reduce((a, p) => a + p.count, 0) > 0;
  return [attempted ? 1 : 0, "major", "formula attempted, no row right"];
}

function rowQuestion(sub, ref, q) {
  const [title, keys] = ROWQ[q];
  const cols = keys.map(k => columnStats(sub, ref, k, q));
  const scores = cols.map(scoreColumn);
  const sc = Math.round(scores.reduce((a, x) => a + x[0], 0) / scores.length * 2) / 2;
  const order = ["blank", "major", "hardcoded", "partial", "carried_forward", "minor", "correct"];
  let verdict = scores.map(x => x[1]).sort((a, b) => order.indexOf(a) - order.indexOf(b))[0];
  if (verdict === "blank" && sc > 0) verdict = "partial";
  const extra = {};
  if (q === 5) { extra.lookup_table = lookupTable(sub); extra.by_code = byCode(sub, ref); }
  return {q, kind: "column", title, columns: cols, proposal: {score: sc, verdict, comment: draftRowComment(q, cols, scores, sc, extra)}, extra};
}
function lookupTable(sub) {
  const c = sub.qcol + 1, out = [];
  for (const r of LOOKUP_ROWS) {
    const a = sub.V(r + sub.qrowOff, c), b = sub.V(r + sub.qrowOff, c + 1);
    if (a !== null || b !== null) out.push(`${show(a)} -> ${show(b)}`);
  }
  return out;
}
function byCode(sub, ref) {
  const c = sub.col.position, tally = new Map();
  for (const r of range(FIRST, LAST)) {
    const code = ref.rows[r].code, v = sub.V(r, c);
    const ok = ["ok", "cosmetic"].includes(judge("text", v, null, ref.rows[r].position, 5)[0]);
    if (!tally.has(code)) tally.set(code, [0, 0, new Counter()]);
    const t = tally.get(code);
    t[0] += ok ? 1 : 0; t[1] += 1;
    if (!ok) t[2].inc(show(v));
  }
  const out = {};
  for (const [k, [a, n, w]] of tally) out[k] = `${a}/${n} right` + (w.size ? ` (got ${w.mostCommon(2).map(([x, m]) => `${x} x${m}`).join(", ")})` : "");
  return out;
}
function draftRowComment(q, cols, scores, sc, extra) {
  const parts = [];
  cols.forEach((cs, i) => {
    const [, verdict, why] = scores[i];
    const st = cs.stats, n = st.n, name = NICE[cs.key];
    const lead = cols.length > 1 ? `${name} (col ${cs.column}): ` : "";
    const good = goodOf(cs);
    if (verdict === "correct") {
      const inv = cs.key === "win" && (st.alt || 0) > (st.ok || 0);
      parts.push(lead + `all ${n} rows right with a formula` + (inv ? " (coded 1 = loss, as the row-4 example shows; accepted)." : "."));
      return;
    }
    if (verdict === "blank") { parts.push(lead + "not answered."); return; }
    if (verdict === "hardcoded") { parts.push(lead + `${st.typed || 0} of ${n} cells are typed by hand, not calculated with a formula, so they do not count.`); return; }
    if (verdict === "carried_forward") {
      const up = UPSTREAM[cs.key];
      parts.push(lead + `your formula is right; ${st.ecf || 0} rows differ only because your ${NICE[up[0]]} column (Q${up[1]}) is wrong (already penalised there).`);
      return;
    }
    const shifts = Object.entries(st).filter(([k]) => k.startsWith("shift"));
    if (shifts.length && why.startsWith("formula refers")) {
      const k = shifts.sort((a, b) => b[1] - a[1])[0][0];
      const where = k === "shift-1" ? "previous" : "next";
      const ex = cs.patterns[0];
      const eg = ex ? ` (e.g. ${ex.cell} ${ex.formula.slice(0, 70)})` : "";
      parts.push(lead + `the formula reads the ${where} row${eg}, so results are shifted by one row: ${n - good} of ${n} rows wrong.`);
      return;
    }
    let msg = lead + `${good} of ${n} rows right.`;
    if (cs.distinct_values === 1 && n > 1) msg += ` Every row gives the same value (${Object.keys(cs.top_values)[0]}), so the formula does not react to the data.`;
    else if (cs.key === "win" && st.alt && st.ok) msg += ` The flag mixes codings: ${st.ok} rows give 1 = win and ${st.alt} rows give 1 = loss.`;
    const nts = Object.keys(cs.notes).filter(k => !k.startsWith("1 = loss"));
    if (nts.length) msg += " " + nts.join("; ") + ".";
    if (Object.keys(cs.variants).length) msg += " Typical mistake: " + Object.keys(cs.variants).join(", ") + ".";
    const bad = cs.examples.filter(e => ["wrong", "error"].includes(e.status));
    if (bad.length) msg += ` E.g. ${bad[0].cell}: got ${bad[0].got}, expected ${bad[0].expected}.`;
    parts.push(msg);
  });
  if (q === 5 && extra.by_code && sc < 10) {
    const bad = Object.entries(extra.by_code).filter(([, v]) => !/^(\d+)\/\1 /.test(v)).map(([k, v]) => `${k} ${v}`);
    if (bad.length) parts.push("By position code: " + bad.join("; ") + ".");
  }
  const head = sc >= 10 ? "Correct." : sc >= 9 ? "Correct, minor issue:" : sc > 0 ? "Partially correct:" : "Incorrect:";
  if (sc === 0 && scores.every(x => x[1] === "blank")) return "Not answered.";
  return `${head} ` + parts.join(" ");
}

// --------------------------------------------------------------------------- aggregates
const LABEL = /^\s*(with(out)?\s+row\s*4.*|sum|answer.*)$/i;
function isHelper(sub, r, c) {
  const f = sub.F(r, c);
  if (!f) return false;
  const p = r1c1(f, r, c);
  return [-1, 1].some(d => { const g = sub.F(r + d, c); return g && r1c1(g, r + d, c) === p; });
}
function studentRows(sub) {
  const out = {};
  const s = v => { v = clean(v); return v === null || v === undefined ? null : (typeof v === "string" ? v : String(v)); };
  for (const r of range(EXAMPLE, LAST)) {
    out[r] = {team: s(sub.V(r, sub.col.team)), position: s(sub.V(r, sub.col.position)), home: s(sub.V(r, sub.col.home)),
      result: s(sub.V(r, sub.col.result)), tickets: num(sub.V(r, sub.col.tickets)), price: num(sub.V(r, sub.col.price))};
  }
  return out;
}
function singleRefs(f) {
  const out = [], seen = new Set();
  for (const m of (f || "").replaceAll("$", "").replace(RANGE, "").matchAll(REF)) {
    const k = `${CIDX(m[2])},${+m[4]}`;
    if (!seen.has(k)) { seen.add(k); out.push([CIDX(m[2]), +m[4]]); }
  }
  return out;
}
function trace(sub, f, depth = 2, seen = new Set()) {
  const out = [];
  const dataCols = new Set(["tickets", "price", "team", "position", "home", "result"].map(k => sub.col[k]));
  for (const [c, r] of singleRefs(f).slice(0, 6)) {
    const k = `${c},${r}`;
    if (seen.has(k)) continue;
    seen.add(k);
    if (dataCols.has(c) && r >= EXAMPLE && r <= LAST) { out.push({cell: `${COL(c)}${r}`, formula: null, value: show(sub.V(r, c)), role: "a data cell used as criterion"}); continue; }
    const pf = sub.F(r, c);
    out.push({cell: `${COL(c)}${r}`, formula: pf ? short(pf, 160) : null, value: show(sub.V(r, c))});
    if (pf && depth > 1) out.push(...trace(sub, pf, depth - 1, seen));
  }
  return out;
}
function rangesInfo(sub, formulas) {
  const partial = [], singles = [];
  const dataCols = new Set(["tickets", "price", "team", "position", "home", "result", "win"].map(k => sub.col[k]));
  for (const f of formulas) {
    if (!f) continue;
    for (const m of f.replaceAll("$", "").matchAll(RANGE)) {
      const [, c1, r1s, c2, r2s] = m, r1 = +r1s, r2 = +r2s;
      if (!range(CIDX(c1), CIDX(c2)).some(c => dataCols.has(c))) continue;
      if (r2 >= FIRST && r1 <= LAST && (r2 - Math.max(r1, FIRST) + 1) < 0.9 * (LAST - FIRST + 1)) partial.push(`${c1}${r1}:${c2}${r2}`);
    }
    for (const [c, r] of singleRefs(f)) if (r >= FIRST && r <= LAST && (c === sub.col.tickets || c === sub.col.price)) singles.push(`${COL(c)}${r}`);
  }
  const bits = [];
  if (partial.length) bits.push("sums only hand-picked rows (" + partial.slice(0, 3).join(", ") + ")");
  if (singles.length) bits.push("adds single rows by hand (" + singles.slice(0, 4).join(", ") + (singles.length > 4 ? "..." : "") + ")");
  return bits.join("; ") || null;
}
function iterCells(sub, maxRow, minCol, maxCol) {
  const out = [];
  for (const [k, x] of sub.cells) {
    const r = Math.floor(k / 16384), c = k % 16384;
    if (r <= maxRow && c >= minCol && c <= maxCol) out.push([r, c, x]);
  }
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}
function scratchWork(sub, shown) {
  const lastData = Math.max(...Object.values(sub.col));
  const tcol = sub.table ? sub.table[1] : 999;
  const cells = [], helpers = new Map();
  for (const [r, c, x] of iterCells(sub, Math.min(sub.book.maxRow, 400), lastData + 1, Math.min(sub.book.maxCol, 60))) {
    const f = x.f, coord = `${COL(c)}${r}`;
    if (!f || shown.has(coord) || (c >= tcol && c <= tcol + 3)) continue;
    if (isHelper(sub, r, c)) {
      const key = `${c}|${r1c1(f, r, c)}`;
      if (!helpers.has(key)) helpers.set(key, [c, r, r, f]);
      helpers.get(key)[2] = r;
      continue;
    }
    let label = null;
    for (let cc = c - 1; cc > Math.max(0, c - 4); cc--) { const v = sub.V(r, cc); if (typeof v === "string" && !QRE.test(v)) { label = v; break; } }
    cells.push({cell: coord, formula: short(f, 140), value: show(sub.V(r, c)), label});
  }
  const hl = [...helpers.values()].filter(([, , a, b]) => b - a >= 5).map(([c, , a, b, f]) => `helper column ${COL(c)} rows ${a}-${b}: ${short(helpers.get(`${c}|${r1c1(f, a, c)}`)?.[3] ?? f, 90)}`);
  return [cells.slice(0, 15), hl];
}
const REQ = {8: [["tickets"], ["price"]], 9: [["team"], ["home"], ["tickets"], ["price"]],
  10: [["team"], ["home"], ["result", "win"], ["tickets"], ["price"]], 11: [["position"], ["home"], ["tickets"]],
  12: [["team"], ["result", "win"], ["tickets"], ["price"]], 13: [["team"], ["home"], ["tickets"]], 14: [["tickets"], ["price"]]};
function columnsUsed(sub, formulas) {
  const inv = Object.fromEntries(Object.entries(sub.col).map(([k, c]) => [c, k]));
  const used = new Set();
  for (const f of formulas) {
    if (!f) continue;
    const g = f.replaceAll("$", "");
    for (const m of g.matchAll(RANGE)) for (const c of range(CIDX(m[1]), CIDX(m[3]))) if (inv[c]) used.add(inv[c]);
    for (const [c] of singleRefs(g)) if (inv[c]) used.add(inv[c]);
  }
  return used;
}
const DEPS = {8: [], 9: [4], 10: [4, 7], 11: [5], 12: [4, 7], 13: [4], 14: []};

function aggQuestion(sub, ref, q, ownAll) {
  const spec = AGG[q], R = ref.agg[q], text = !!spec.text;
  const tr = q in sub.qpos ? sub.qpos[q][0] : spec.text_row + sub.qrowOff;
  const qc = q in sub.qpos ? sub.qpos[q][1] : sub.qcol;
  const tmpl = [tr + spec.ans_row - spec.text_row, qc + 1];
  const cands = [];
  for (const r of range(tr, tr + 3)) for (const c of range(qc + 1, qc + 10)) {
    const v = sub.V(r, c), f = sub.F(r, c);
    if (([null, "", "="].includes(v) && !f) || (f || "").trim() === "=") continue;
    if (typeof v === "string" && (QRE.test(v) || LABEL.test(v) || v.length > 60)) continue;
    if (isHelper(sub, r, c)) continue;
    cands.push({cell: `${COL(c)}${r}`, r, c, formula: f ? short(f, 220) : null, value: v, typed: f === null, fmt: sub.NF(r, c)});
  }
  const acc = R.accepted;
  const matches = v => text ? typeof v === "string" && clean(v).trim().toLowerCase() === acc[0].toLowerCase() : acc.some(a => close(num(v), a));
  let chosen = cands.find(x => matches(x.value)) || cands.find(x => x.r === tmpl[0] && x.c === tmpl[1]);
  if (!chosen && cands.length) {
    const pool = cands.filter(x => !text || typeof x.value === "string");
    const P = pool.length ? pool : cands;
    chosen = P.reduce((best, x) => (Math.abs(x.r - tmpl[0]) * 3 + Math.abs(x.c - tmpl[1])) < (Math.abs(best.r - tmpl[0]) * 3 + Math.abs(best.c - tmpl[1])) ? x : best);
  }
  const elsewhere = [];
  if (!text) {
    const lastData = Math.max(...Object.values(sub.col));
    const inC = new Set(cands.map(x => `${x.r},${x.c}`));
    for (const [r, c, x] of iterCells(sub, Math.min(sub.book.maxRow, 400), lastData + 1, Math.min(sub.book.maxCol, 60))) {
      if (inC.has(`${r},${c}`)) continue;
      if (matches(x.v) && x.f) elsewhere.push({cell: `${COL(c)}${r}`, r, c, formula: short(x.f, 220), value: x.v, typed: false, fmt: x.nf});
    }
    if (!chosen && elsewhere.length) chosen = elsewhere[0];
  }
  const ev = {q, kind: "single", title: spec.title, template_cell: `${COL(tmpl[1])}${tmpl[0]}`, expected: acc.map(show),
    candidates: cands.slice(0, 6).map(({r, c, ...x}) => ({...x, value: show(x.value)}))};
  if (!chosen) { ev.answer = null; ev.proposal = {score: 0, verdict: "blank", comment: `Not answered. Expected ${ev.expected[0]}.`}; return ev; }
  const v = chosen.value, f = sub.F(chosen.r, chosen.c);
  const ans = {cell: chosen.cell, formula: f ? short(f, 300) : null, value: show(v), typed: f === null, number_format: chosen.fmt, functions: funcs(f)};
  if (!(chosen.r === tmpl[0] && chosen.c === tmpl[1])) ans.note = `answer found in ${chosen.cell} (template cell ${ev.template_cell})` + (elsewhere.includes(chosen) ? ", far from the question" : "");
  if (elsewhere.length) ev.expected_value_also_in = elsewhere.slice(0, 4).map(x => x.cell);
  const tr_ = f ? trace(sub, f) : [];
  if (tr_.length) ans.precedents = tr_;
  let picked = rangesInfo(sub, [f, ...tr_.map(t => t.formula)]);
  const crit = tr_.filter(t => t.role).map(t => t.cell);
  if (crit.length) picked = [picked, "takes its conditions from hand-picked data cells (" + crit.slice(0, 4).join(", ") + ")"].filter(Boolean).join("; ");
  ans.hand_picked = picked;
  ev.answer = ans;
  const diag = {match: "none"};
  const used = columnsUsed(sub, [f, ...tr_.map(t => t.formula)]);
  const missing = REQ[q].filter(s => !s.some(k => used.has(k))).map(s => [...s].sort().join(" or "));
  if (f) { diag.columns_used = [...used].sort(); if (missing.length) diag.columns_not_used = missing; }
  const logicOk = () => f !== null && !missing.length;     // no Excel in the browser: structural check
  if (isErr(v)) diag.match = "error";
  else if (matches(v)) diag.match = "correct";
  else if (text) {
    const own = ownAll[13];
    const sv = typeof v === "string" ? clean(v).trim().toLowerCase() : "";
    if (sv === "yes" || sv === "no") {
      if (sv === own.accepted[0].toLowerCase() && logicOk()) diag.match = "carried_forward";
      else if (sv === R.by_sales.toLowerCase()) { diag.match = "variant"; diag.variant = "compared sales instead of tickets"; }
    }
  } else {
    const nv = num(v), own = ownAll[q];
    diag.own_columns_value = own.accepted.map(show);
    if (logicOk() && own.accepted.some(a => close(nv, a))) diag.match = "carried_forward";
    else {
      const probes = [["", R, nv]];
      if (nv) probes.push([" (with your own helper columns)", own, nv]);
      outer: for (const [label, src, val] of probes) {
        for (const [name, vals] of src.variants) {
          if (vals.some(x => x && close(val, x))) { diag.match = "variant"; diag.variant = name + label; break outer; }
        }
      }
    }
  }
  if (q === 13 && ans.typed) diag.supporting_formulas = cands.filter(x => x.formula).map(x => x.cell);
  ev.diagnosis = diag;
  ev.proposal = proposeAgg(q, spec, ev, ans, diag);
  return ev;
}

function proposeAgg(q, spec, ev, ans, diag) {
  const exp = ev.expected.join(" or "), got = ans.value;
  const money = spec.money && !(ans.number_format || "").includes("$");
  if (diag.match === "correct") {
    if (ans.typed) {
      if (q === 13) {
        if ((diag.supporting_formulas || []).length) return {score: 10, verdict: "correct", comment: "Correct: 'Yes', backed by your calculation in " + diag.supporting_formulas.slice(0, 2).join(", ") + "."};
        return {score: 0, verdict: "hardcoded", comment: "Not valid: 'Yes' is typed by hand with no calculation behind it; it must come from a formula (e.g. =IF(SUMIFS(...Home)>SUMIFS(...Away),\"Yes\",\"No\"))."};
      }
      return {score: TYPED_FACTOR, verdict: "hardcoded", comment: `Not valid: the number ${got} is typed by hand, not calculated with a formula, so it does not count.`};
    }
    if (ans.hand_picked) return {score: MANUAL_RANGE, verdict: "manual", comment: `Not valid: the number (${got}) is right but your formula ${ans.hand_picked} (data sorted / rows selected by hand) instead of applying conditions to the whole column with SUMIFS / SUMPRODUCT.`};
    if (money) return {score: 10 - MONEY_PENALTY, verdict: "minor", money_missing: true, comment: `Correct value (${got}); the question asks for dollar format and the cell is not formatted as $ (-0.5, deducted only once in the quiz).`};
    return {score: 10, verdict: "correct", comment: `Correct (${got}).`};
  }
  if (ans.hand_picked && !ans.typed && diag.match !== "error") return {score: MANUAL_RANGE, verdict: "manual", comment: `Not valid: ${got} instead of ${exp}, and your calculation ${ans.hand_picked} (data sorted / rows selected by hand) instead of applying conditions to the whole column with SUMIFS / SUMPRODUCT.`};
  if (diag.match === "carried_forward") {
    const d = DEPS[q].map(x => `Q${x}`).join(", ") || "an earlier question";
    return {score: CARRIED_FORWARD, verdict: "carried_forward", comment: `Error carried forward: your formula is right; it gives ${got} because your helper column(s) from ${d} are wrong (expected ${exp}); that error was already penalised there.`};
  }
  if (diag.match === "variant") return {score: ans.typed ? 0 : ONE_MISTAKE, verdict: "partial", comment: `Partially correct: ${got} instead of ${exp}. Your calculation ${diag.variant}.`};
  if (diag.match === "error") return {score: ans.formula && ans.functions.length ? 1 : 0, verdict: "error", comment: `Your formula returns ${got} (it does not work), expected ${exp}.`};
  if (ans.typed) return {score: 0, verdict: "hardcoded", comment: `Incorrect: typed value ${got}, expected ${exp}, and no formula.`};
  const relevant = ans.functions.some(x => ["SUMPRODUCT", "SUMIFS", "SUMIF", "COUNTIFS", "COUNTIF", "SUM", "IF", "FILTER", "COUNT"].includes(x));
  return {score: relevant ? 3 : 1, verdict: "major", comment: `Incorrect: ${got} instead of ${exp}. Check the conditions and the columns used in your formula.`};
}

// --------------------------------------------------------------------------- evidence
export function buildEvidence(book, ref, file) {
  const sub = new Sub(book, ref);
  const {student, sid} = canvasName(file);
  const own = aggregates(studentRows(sub));
  const qs = [...Object.keys(ROWQ).map(q => rowQuestion(sub, ref, +q)), ...Object.keys(AGG).map(q => aggQuestion(sub, ref, +q, own))];
  const money = qs.filter(q => q.proposal.money_missing);
  for (const q of money.slice(1)) Object.assign(q.proposal, {score: 10, verdict: "correct", money_missing: false, comment: `Correct (${q.answer.value}); remember the dollar format (already deducted once, in Q${money[0].q}).`});
  const shown = new Set();
  for (const q of qs) {
    if (q.kind === "single") for (const c of q.candidates) shown.add(c.cell);
    if (q.answer) { shown.add(q.answer.cell); for (const t of q.answer.precedents || []) shown.add(t.cell); }
  }
  const [cells, helpers] = scratchWork(sub, shown);
  return {student, id: sid, file,
    layout: {notes: sub.notes, inserted_columns: sub.inserted.map(COL), columns: Object.fromEntries(Object.entries(sub.col).map(([k, c]) => [k, COL(c)])),
      question_text_column: COL(sub.qcol), correction_table: sub.table ? `${COL(sub.table[1])}${sub.table[0]}` : null},
    questions: qs, scratch: {cells, helper_columns: helpers}, _sub: sub};
}

export function proposalTotal(ev) { return Math.round(ev.questions.reduce((a, q) => a + q.proposal.score, 0) / ev.questions.length * 10 * 100) / 100; }

export function evidenceMd(ev) {
  const L = [`# Quiz 1 evidence -- ${ev.student} (${ev.id})`, `file: ${ev.file}`];
  const lay = ev.layout;
  L.push("layout: " + (lay.notes.length ? lay.notes.join("; ") : "template layout"));
  L.push(`correction table header at ${lay.correction_table}`);
  if (lay.inserted_columns.length) L.push("student columns: " + Object.entries(lay.columns).map(([k, v]) => `${k}=${v}`).join(", "));
  for (const q of ev.questions) {
    const p = q.proposal;
    L.push("", `## Q${q.q} ${q.title}   [proposal ${p.score}/10 · ${p.verdict}]`);
    if (q.kind === "column") {
      for (const cs of q.columns) {
        const st = cs.stats;
        const bits = Object.entries(st).filter(([k, v]) => k !== "n" && v).map(([k, v]) => `${k} ${v}`);
        L.push(`- ${NICE[cs.key]} col ${cs.column} (rows ${FIRST}-${LAST}, n=${st.n}): ` + bits.join(", "));
        for (const pt of cs.patterns) L.push(`  - formula x${pt.count}: ${pt.cell} ${pt.formula}`);
        if (cs.n_patterns > 3) L.push(`  - (${cs.n_patterns} different formula shapes in the column)`);
        if (Object.keys(cs.raw_offsets).some(k => +k !== 0)) L.push(`  - row offset of Raw_Import references: ${pyRepr(cs.raw_offsets)}`);
        if (Object.keys(cs.notes).length) L.push(`  - notes: ${pyRepr(cs.notes)}`);
        if (cs.distinct_values <= 3) L.push(`  - only ${cs.distinct_values} distinct value(s): ${pyRepr(cs.top_values)}`);
        if (Object.keys(cs.variants).length) L.push(`  - typical-mistake matches: ${pyRepr(cs.variants)}`);
        for (const e of cs.examples) L.push(`  - ${e.cell} [${e.status}] got ${e.got} | expected ${e.expected} | ${e.formula}`);
      }
      for (const [k, v] of Object.entries(q.extra || {})) L.push(`- ${k}: ${pyRepr(v)}`);
    } else {
      L.push(`- expected: ${q.expected.join(" or ")} (template cell ${q.template_cell})`);
      const a = q.answer;
      if (!a) L.push("- answer: none found near the question");
      else {
        L.push(`- answer ${a.cell}: ${a.value} | ` + (a.formula ? `formula ${a.formula}` : "TYPED (no formula)") + ` | format ${pyRepr(a.number_format)}`);
        for (const k of ["note", "hand_picked"]) if (a[k]) L.push(`- ${k}: ${a[k]}`);
        for (const pr of a.precedents || []) L.push(`- -> ${pr.cell}: ${pr.value} | ${pr.role || pr.formula || "typed"}`);
        L.push(`- diagnosis: ${pyRepr(q.diagnosis)}`);
      }
      if (q.expected_value_also_in) L.push(`- expected value also computed in: ${pyRepr(q.expected_value_also_in)}`);
      const others = q.candidates.filter(c => !a || c.cell !== a.cell);
      for (const c of others.slice(0, 4)) L.push(`- other cell near the question ${c.cell}: ${c.value} | ${c.formula || "typed"}`);
    }
    L.push(`- proposal comment: ${p.comment}`);
  }
  const sc = ev.scratch || {};
  if ((sc.cells || []).length || (sc.helper_columns || []).length) {
    L.push("", "## Other formulas in the sheet (student's side calculations)");
    for (const h of sc.helper_columns || []) L.push(`- ${h}`);
    for (const c of sc.cells || []) L.push(`- ${c.cell}${c.label ? " [" + String(c.label).trim() + "]" : ""}: ${c.value} | ${c.formula}`);
  }
  return L.join("\n") + "\n";
}

// Expected answers of the aggregate questions, injected into the prompt at run time
export function expectedSummary(ref) {
  const a = ref.agg, s = q => a[q].accepted.map(show).join(" / ");
  return `Expected values computed from this class's data (both "with row 4" and "without row 4" accepted where they differ): `
    + `Q8 ${s(8)} · Q9 ${s(9)} · Q10 ${s(10)} · Q11 ${s(11)} · Q12 ${s(12)} · Q13 ${a[13].accepted[0]} · Q14 ${s(14)}. `
    + `Typical wrong values for Q14: AND instead of OR = ${show(a[14].variants[0][1][0])}, COUNTIF+COUNTIF double counting = ${show(a[14].variants[1][1][0])}.`;
}
