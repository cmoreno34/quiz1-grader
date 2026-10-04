// writer.js -- hands back each student's ORIGINAL workbook with only the correction table
// edited (points 0-10 coloured, a 'comment' column right beside them, GRADE /100 and the
// overall summary).  String-level XML edits: every other part of the file is copied as is.
// Port of write_workbook() in _tools/quiz1_agent.py.
import {colLetter as COL, colIndex as CIDX, unescapeXml} from "./xlsx.js";
import {num, close, evidenceMd} from "./engine.js";

const GREEN = "FFC6EFCE", AMBER = "FFFFEB9C", RED = "FFFFC7CE";
const esc = s => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "")
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const attr = (s, n) => { const m = s.match(new RegExp(`\\b${n}="([^"]*)"`)); return m ? m[1] : null; };
const setAttr = (tag, n, v) => new RegExp(`\\b${n}="[^"]*"`).test(tag) ? tag.replace(new RegExp(`\\b${n}="[^"]*"`), `${n}="${v}"`)
  : tag.replace(/^<(\w+)/, `<$1 ${n}="${v}"`);
const delAttr = (tag, n) => tag.replace(new RegExp(`\\s${n}="[^"]*"`), "");
const fmtNum = v => String(Number.isInteger(v) ? v : +v.toPrecision(15));

class Sheet {
  constructor(xml) { this.xml = xml; this.maxCol = 0; this.droppedFormula = false; }
  rowSpan(r) {
    const re = new RegExp(`<row\\b(?=[^>]*\\sr="${r}")[^>]*?(\\/?)>`);
    const m = re.exec(this.xml);
    if (!m) return null;
    const start = m.index, openEnd = m.index + m[0].length;
    if (m[1] === "/") return {start, end: openEnd, open: m[0].replace(/\/>$/, ">"), inner: "", self: true};
    const close = this.xml.indexOf("</row>", openEnd);
    return {start, end: close + 6, open: m[0], inner: this.xml.slice(openEnd, close), self: false};
  }
  ensureRow(r) {
    let span = this.rowSpan(r);
    if (span) return span;
    const rows = [...this.xml.matchAll(/<row\b[^>]*\sr="(\d+)"/g)];
    const next = rows.find(m => +m[1] > r);
    const tag = `<row r="${r}"></row>`;
    if (next) this.xml = this.xml.slice(0, next.index) + tag + this.xml.slice(next.index);
    else if (/<sheetData\s*\/>/.test(this.xml)) this.xml = this.xml.replace(/<sheetData\s*\/>/, `<sheetData>${tag}</sheetData>`);
    else this.xml = this.xml.replace("</sheetData>", tag + "</sheetData>");
    return this.rowSpan(r);
  }
  cells(inner) {
    return [...inner.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map(m => {
      const ref = attr(m[0].match(/^<c\b[^>]*>/)[0], "r");
      return {xml: m[0], col: CIDX(ref.match(/^[A-Z]+/)[0])};
    });
  }
  getCell(r, c) {
    const span = this.rowSpan(r);
    if (!span) return null;
    const x = this.cells(span.inner).find(x => x.col === c);
    return x ? x.xml : null;
  }
  style(r, c) { const x = this.getCell(r, c); return x ? attr(x.match(/^<c\b[^>]*>/)[0], "s") : null; }
  putCell(r, c, build) {
    const span = this.ensureRow(r);
    const list = this.cells(span.inner);
    const i = list.findIndex(x => x.col === c);
    const old = i >= 0 ? list[i].xml : null;
    if (old && /<f\b/.test(old)) {
      if (/<f\b[^>]*\bt="shared"[^>]*\bref="/.test(old)) throw new Error(`${COL(c)}${r} is the master of a shared formula`);
      this.droppedFormula = true;
    }
    const oldStyle = old ? attr(old.match(/^<c\b[^>]*>/)[0], "s") : null;
    const xml = build(`${COL(c)}${r}`, oldStyle);
    if (i >= 0) list[i].xml = xml;
    else { const at = list.findIndex(x => x.col > c); list.splice(at < 0 ? list.length : at, 0, {xml, col: c}); }
    let open = span.open;
    const sp = attr(open, "spans");
    if (sp && sp.includes(":")) { const [lo, hi] = sp.split(":").map(Number); if (c > hi || c < lo) open = setAttr(open, "spans", `${Math.min(lo, c)}:${Math.max(hi, c)}`); }
    this.xml = this.xml.slice(0, span.start) + open + list.map(x => x.xml).join("") + "</row>" + this.xml.slice(span.end);
    this.maxCol = Math.max(this.maxCol, c);
  }
  setNumber(r, c, v, s = null) { this.putCell(r, c, (ref, os) => `<c r="${ref}"${(s ?? os) !== null ? ` s="${s ?? os}"` : ""}><v>${fmtNum(v)}</v></c>`); }
  setText(r, c, t, s = null) { this.putCell(r, c, (ref, os) => `<c r="${ref}"${(s ?? os) !== null ? ` s="${s ?? os}"` : ""} t="inlineStr"><is><t xml:space="preserve">${esc(t)}</t></is></c>`); }
  setFormula(r, c, f, cached, s = null) { this.putCell(r, c, (ref, os) => `<c r="${ref}"${(s ?? os) !== null ? ` s="${s ?? os}"` : ""}><f>${esc(f)}</f><v>${fmtNum(cached)}</v></c>`); }
  updateCached(r, c, cached) {
    const span = this.rowSpan(r);
    if (!span) return false;
    const list = this.cells(span.inner);
    const x = list.find(x => x.col === c);
    if (!x || !/<f\b/.test(x.xml)) return false;
    let cell = x.xml.replace(/^<c\b[^>]*>/, t => delAttr(t, "t"));
    cell = /<v>[\s\S]*?<\/v>/.test(cell) ? cell.replace(/<v>[\s\S]*?<\/v>/, `<v>${fmtNum(cached)}</v>`) : cell.replace(/<\/c>$/, `<v>${fmtNum(cached)}</v></c>`);
    x.xml = cell;
    this.xml = this.xml.slice(0, span.start) + span.open + list.map(x => x.xml).join("") + "</row>" + this.xml.slice(span.end);
    return true;
  }
  finish() {
    this.xml = this.xml.replace(/<dimension ref="([^"]*)"\s*\/>/, (m, ref) => {
      const [a, b0] = ref.split(":"), b = b0 || a;
      const bm = /^([A-Z]+)(\d+)$/.exec(b);
      if (bm && this.maxCol > CIDX(bm[1])) return `<dimension ref="${a}:${COL(this.maxCol)}${bm[2]}"/>`;
      return m;
    });
    return this.xml;
  }
}

function addFillStyles(stylesXml, baseXf) {
  const fills = stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
  let nFills = (fills[1].match(/<fill\b/g) || []).length;
  const newFills = [GREEN, AMBER, RED].map(rgb => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`);
  const fillIds = [nFills, nFills + 1, nFills + 2];
  nFills += 3;
  let xml = stylesXml.replace(fills[0], setAttr(fills[0].match(/^<fills\b[^>]*>/)[0], "count", nFills) + fills[1] + newFills.join("") + "</fills>");
  const xfs = xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  const list = [...xfs[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m => m[0]);
  const base = list[+(baseXf ?? 0)] || list[0];
  const clones = fillIds.map(id => base.replace(/^<xf\b[^>]*?(?=\/?>)/, t => setAttr(setAttr(t, "fillId", id), "applyFill", "1")));
  const ids = [list.length, list.length + 1, list.length + 2];
  xml = xml.replace(xfs[0], setAttr(xfs[0].match(/^<cellXfs\b[^>]*>/)[0], "count", list.length + 3) + xfs[1] + clones.join("") + "</cellXfs>");
  return {xml, ids};
}

// final: {1..14: {score, verdict, comment, basis}, summary}
export async function writeWorkbook(JSZip, data, sub, final) {
  const zip = await JSZip.loadAsync(data);
  const part = sub.book.sheetPath;
  const sx = new Sheet(await zip.file(part).async("string"));
  let tr, tc, made = false, weightsOk = true;
  if (!sub.table) {
    tr = 5 + sub.qrowOff; tc = sub.qcol + 9; made = true;
    sx.setText(tr, tc, "question"); sx.setText(tr, tc + 1, "weight"); sx.setText(tr, tc + 2, "points");
    for (let i = 1; i <= 14; i++) { sx.setNumber(tr + i, tc, i); sx.setFormula(tr + i, tc + 1, `1/COUNT($${COL(tc)}$${tr + 1}:$${COL(tc)}$${tr + 14})`, 1 / 14); }
    sx.setFormula(tr + 15, tc + 1, `SUM(${COL(tc + 1)}${tr + 1}:${COL(tc + 1)}${tr + 14})`, 1);
    sx.setText(tr + 16, tc + 1, "GRADE"); sx.setFormula(tr + 16, tc + 2, "0", 0);
  } else [tr, tc] = sub.table;
  const wc = tc + 1, pc = tc + 2;
  let cc = pc + 1;
  const filled = (r, c) => { const v = sub.V(r, c); return v !== null && v !== ""; };
  if (!made) while ([...Array(15).keys()].some(i => filled(tr + i, cc))) cc++;
  if (!made) weightsOk = [...Array(14).keys()].every(i => { const w = num(sub.V(tr + 1 + i, wc)); return w !== null && Math.abs(w - 1 / 14) < 1e-6; });
  const scores = Array.from({length: 14}, (_, i) => final[i + 1].score);
  const grade10 = scores.reduce((a, b) => a + b, 0) / 14;
  const grade100 = Math.round(grade10 * 10 * 100) / 100;
  const st = addFillStyles(await zip.file("xl/styles.xml").async("string"), sx.style(tr + 1, pc));
  const [sg, sa, sr] = st.ids;
  sx.setText(tr, cc, "comment", sx.style(tr, pc));
  for (let i = 1; i <= 14; i++) {
    const s = final[i].score;
    sx.setNumber(tr + i, pc, s, s >= 9 ? sg : s < 3 ? sr : sa);
    sx.setText(tr + i, cc, final[i].comment);
  }
  sx.updateCached(tr + 15, pc, scores.reduce((a, b) => a + b, 0));
  sx.updateCached(tr + 16, pc, weightsOk ? grade10 : scores.reduce((a, s, i) => a + (num(sub.V(tr + 1 + i, wc)) || 0) * s, 0));
  const P = COL(pc);
  const free = !filled(tr + 17, wc) && !filled(tr + 17, pc);
  if (final.summary && !filled(tr + 17, cc)) sx.setText(tr + 17, cc, "Overall: " + final.summary);
  if (free) {
    sx.setText(tr + 17, wc, "GRADE /100", sx.style(tr + 16, wc));
    sx.setFormula(tr + 17, pc, weightsOk ? `ROUND(${P}${tr + 16}*10,2)` : `ROUND(AVERAGE(${P}${tr + 1}:${P}${tr + 14})*10,2)`, grade100, sx.style(tr + 16, pc));
  }
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
  return {bytes, info: {grade100, table: `${COL(tc)}${tr}`, points_col: P, comment_col: COL(cc),
    grade_cell: free ? `${P}${tr + 17}` : `${P}${tr + 16}`, weights_ok: weightsOk, created_table: made}};
}

// ---------------------------------------------------------------------- reports
const csvCell = v => /[",\n\r]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
export const toCsv = rows => "﻿" + rows.map(r => r.map(csvCell).join(",")).join("\r\n") + "\r\n";

export function justificationMd(ev, final, source, info, grade, meta = {}) {
  const md = evidenceMd(ev);
  const sections = {};
  for (const blk of md.split(/\n(?=## Q\d+ )/)) { const m = /^## Q(\d+) /.exec(blk); if (m) sections[+m[1]] = blk; }
  const L = [`# Quiz 1 -- justification of the grade -- ${ev.student} (${ev.id})`, "",
    `- file: ${ev.file}`,
    `- **grade: ${grade.toFixed(2)} / 100** (in the workbook: ${info.grade_cell}; marks in column ${info.points_col}, comments in column ${info.comment_col})`,
    `- decided by: ${source}` + (meta.model ? ` -- ${meta.model} on ${meta.graded_at}` : ""),
    `- layout: ${ev.layout.notes.join("; ") || "template layout"}`, ""];
  if (final.summary) L.push(`**Overall:** ${final.summary}`, "");
  L.push("| Q | mark /10 | verdict | told to the student |", "|---|---|---|---|");
  for (let q = 1; q <= 14; q++) L.push(`| ${q} | ${final[q].score} | ${final[q].verdict || ""} | ${final[q].comment.replace(/\|/g, "/")} |`);
  for (let q = 1; q <= 14; q++) {
    const f = final[q];
    L.push("", `## Q${q} -- ${f.score}/10 (${f.verdict || ""})`, `- **told to the student:** ${f.comment}`);
    if (f.basis) L.push(`- **basis (grader):** ${f.basis}`);
    for (const c of (meta.calibration || []).filter(c => c.question === q)) L.push(`- **calibrated** across the class: ${c.old} -> ${c.new}`);
    L.push("- **engine evidence:**", "", "```", (sections[q] || "").trim(), "```");
  }
  return L.join("\n") + "\n";
}
