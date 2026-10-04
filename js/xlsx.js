// xlsx.js -- minimal read-only .xlsx reader: formulas (shared formulas expanded, array
// formulas in braces), cached values, number formats.  Mirrors what openpyxl gives the
// Python engine (formula view + data_only view) so both engines classify cells alike.

export class XDate {               // a number shown with a date format (openpyxl -> datetime)
  constructor(serial) { this.serial = serial; }
  toString() { return excelDate(this.serial).toISOString().slice(0, 19).replace("T", " "); }
}

export function excelDate(serial) {   // Excel serial -> JS Date (UTC)
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
}

const BUILTIN = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
  5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)', 7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)', 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?",
  13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yy h:mm", 37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
  44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)', 45: "mm:ss", 46: "[h]:mm:ss",
  47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

export function isDateFormat(fmt) {
  if (!fmt || fmt === "General") return false;
  let f = fmt.split(";")[0].replace(/"[^"]*"|\\.|\[(?!h\]|m\]|s\])[^\]]*\]/gi, "");
  return /(?<![_\\])[dmhysDMHYS]/.test(f);
}

export function unescapeXml(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (m, e) =>
    e === "lt" ? "<" : e === "gt" ? ">" : e === "amp" ? "&" : e === "quot" ? '"' : e === "apos" ? "'"
      : String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)))
    .replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
}

export function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function colLetter(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

const attr = (s, name) => { const m = s.match(new RegExp(`\\b${name}="([^"]*)"`)); return m ? m[1] : null; };

// shift the relative references of a formula (shared-formula expansion, like openpyxl)
export function translateFormula(f, dr, dc) {
  let out = "", i = 0;
  while (i < f.length) {
    if (f[i] === '"') { const j = f.indexOf('"', i + 1); const end = j < 0 ? f.length : j + 1; out += f.slice(i, end); i = end; continue; }
    const m = /^(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])/.exec(f.slice(i));
    const prev = i > 0 ? f[i - 1] : "";
    if (m && !/[A-Za-z_"!.\d]/.test(prev) || (m && prev === "!")) {
      const [all, ca, c, ra, r] = m;
      const nc = ca ? c : colLetter(colIndex(c) + dc);
      const nr = ra ? r : String(+r + dr);
      out += ca + nc + ra + nr; i += all.length; continue;
    }
    out += f[i]; i++;
  }
  return out;
}

async function text(zip, path) { const f = zip.file(path); return f ? await f.async("string") : null; }

export async function readWorkbook(JSZip, data, sheetName = "Data&Q") {
  const zip = await JSZip.loadAsync(data);
  const wb = await text(zip, "xl/workbook.xml");
  const rels = await text(zip, "xl/_rels/workbook.xml.rels");
  if (!wb || !rels) throw new Error("not an .xlsx workbook");
  const sheets = [...wb.matchAll(/<sheet\b[^>]*\/>/g)].map(m => m[0]);
  const names = sheets.map(t => unescapeXml(attr(t, "name") || ""));
  const tag = sheets.find(t => unescapeXml(attr(t, "name") || "") === sheetName);
  if (!tag) throw new Error(`no sheet '${sheetName}' (sheets: ${names.join(", ")})`);
  const rid = attr(tag, "r:id");
  const rel = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map(m => m[0]).find(t => attr(t, "Id") === rid);
  let target = attr(rel, "Target");
  const sheetPath = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\.\//, "");

  const ssXml = await text(zip, "xl/sharedStrings.xml");
  const shared = ssXml ? [...ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m =>
    unescapeXml([...m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, "").matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
      .map(x => x[1]).join(""))) : [];

  const stXml = (await text(zip, "xl/styles.xml")) || "";
  const custom = {};
  for (const m of stXml.matchAll(/<numFmt\b[^>]*\/>/g)) custom[attr(m[0], "numFmtId")] = unescapeXml(attr(m[0], "formatCode") || "");
  const xfsBlock = (stXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [, ""])[1];
  const xfFmt = [...xfsBlock.matchAll(/<xf\b[^>]*?(?:\/>|>)/g)].map(m => {
    const id = attr(m[0], "numFmtId") || "0";
    return custom[id] ?? BUILTIN[+id] ?? "General";
  });

  const sx = await text(zip, sheetPath);
  const cells = new Map();
  const masters = {};
  let maxRow = 0, maxCol = 0;
  for (const m of sx.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const a = m[1], inner = m[2] || "";
    const ref = attr(a, "r");
    if (!ref) continue;
    const mm = /^([A-Z]+)(\d+)$/.exec(ref);
    const c = colIndex(mm[1]), r = +mm[2];
    const s = attr(a, "s"), t = attr(a, "t");
    const nf = xfFmt[+(s || 0)] ?? "General";
    let f = null;
    const fm = inner.match(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/);
    if (fm) {
      const fa = fm[1], ftxt = fm[2] !== undefined ? unescapeXml(fm[2]) : "";
      const ft = attr(fa, "t"), si = attr(fa, "si");
      if (ft === "shared" && si !== null) {
        if (ftxt) { masters[si] = {r, c, f: ftxt}; f = "=" + ftxt; }
        else if (masters[si]) f = "=" + translateFormula(masters[si].f, r - masters[si].r, c - masters[si].c);
      } else if (ft === "array") f = "{=" + ftxt + "}";
      else if (ftxt) f = "=" + ftxt;
    }
    let v = null;
    const vm = inner.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/);
    if (t === "inlineStr") {
      const im = inner.match(/<is>([\s\S]*?)<\/is>/);
      v = im ? unescapeXml([...im[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x => x[1]).join("")) : "";
    } else if (vm) {
      const raw = vm[1];
      if (t === "s") v = shared[+raw] ?? "";
      else if (t === "str" || t === "e") v = unescapeXml(raw);
      else if (t === "b") v = raw === "1";
      else { const n = parseFloat(raw); v = Number.isNaN(n) ? unescapeXml(raw) : n; }
    }
    if (typeof v === "number" && isDateFormat(nf)) v = new XDate(v);
    if (f === null && v === null && !s) continue;
    cells.set(r * 16384 + c, {f, v, nf, t});
    if (f !== null || (v !== null && v !== "")) { if (r > maxRow) maxRow = r; if (c > maxCol) maxCol = c; }
  }
  return {zip, sheetPath, cells, maxRow, maxCol, sheetNames: names};
}
