// app.js -- page logic of the Quiz 1 grader (everything runs in this browser).
import Anthropic from "https://esm.sh/@anthropic-ai/sdk@0.131.0";
import {readWorkbook, colIndex, colLetter} from "./xlsx.js?v=2";
import {buildReference, buildEvidence, evidenceMd, expectedSummary, show, proposalTotal} from "./engine.js?v=2";
import {writeWorkbook, justificationMd, toCsv} from "./writer.js?v=2";
import {MODEL, systemBlocks, gradeStudent, calibrate, costOf} from "./agent.js?v=2";

const JSZip = window.JSZip;
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
const fmt = n => (Math.round(n * 100) / 100).toLocaleString("es-ES", {minimumFractionDigits: 2, maximumFractionDigits: 2});
const S = {zipName: "", items: [], ref: null, decisions: {}, results: [], usage: [], running: false, downloaded: true};
window.__q1 = S;                                   // handy for debugging in the console
window.__q1load = f => loadZip(f);

// ------------------------------------------------------------------ encrypted key storage
const KEY_SLOT = "q1web.key";
const enc = new TextEncoder(), dec = new TextDecoder();
const b64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
async function derive(pass, salt) {
  const base = await crypto.subtle.importKey("raw", enc.encode(pass), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey({name: "PBKDF2", salt, iterations: 310000, hash: "SHA-256"}, base, {name: "AES-GCM", length: 256}, false, ["encrypt", "decrypt"]);
}
async function saveKey(key, pass) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = await crypto.subtle.encrypt({name: "AES-GCM", iv}, await derive(pass, salt), enc.encode(key));
  localStorage.setItem(KEY_SLOT, JSON.stringify({salt: b64(salt), iv: b64(iv), ct: b64(ct)}));
}
async function loadKey(pass) {
  const o = JSON.parse(localStorage.getItem(KEY_SLOT));
  try { return dec.decode(await crypto.subtle.decrypt({name: "AES-GCM", iv: unb64(o.iv)}, await derive(pass, unb64(o.salt)), unb64(o.ct))); }
  catch (e) { throw new Error("Contraseña incorrecta"); }
}
const hasSavedKey = () => { try { return !!localStorage.getItem(KEY_SLOT); } catch (e) { return false; } };
let sessionKey = "";
function renderKey() {
  const saved = hasSavedKey();
  $("#keyNew").hidden = saved || !!sessionKey; $("#keySaved").hidden = !saved || !!sessionKey; $("#keyReady").hidden = !sessionKey;
}
$("#unlock").onclick = async () => {
  $("#keyErr").textContent = "";
  try { sessionKey = await loadKey($("#unlockPass").value); $("#unlockPass").value = ""; renderKey(); refreshButtons(); }
  catch (e) { $("#keyErr").textContent = e.message; }
};
$("#unlockPass").addEventListener("keydown", e => { if (e.key === "Enter") $("#unlock").click(); });
$("#forget").onclick = e => { e.preventDefault(); if (!confirm("¿Borrar la clave guardada en este navegador?")) return; localStorage.removeItem(KEY_SLOT); sessionKey = ""; renderKey(); refreshButtons(); };
$("#lock").onclick = e => { e.preventDefault(); sessionKey = ""; renderKey(); refreshButtons(); };
async function keyFromForm() {
  if (sessionKey) return sessionKey;
  const key = $("#key").value.trim(), pass = $("#pass").value;
  if (!/^sk-ant-/.test(key)) throw new Error("Pega tu clave: empieza por sk-ant-");
  if ($("#remember").checked) {
    if (pass.length < 6) throw new Error("Elige una contraseña de al menos 6 caracteres para guardar la clave cifrada");
    await saveKey(key, pass);
  }
  sessionKey = key; $("#key").value = ""; $("#pass").value = ""; renderKey();
  return key;
}

// ------------------------------------------------------------------ read the Canvas zip
const drop = $("#drop");
drop.onclick = () => $("#file").click();
$("#file").onchange = e => e.target.files[0] && loadZip(e.target.files[0]);
["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => { const f = e.dataTransfer.files[0]; if (f) loadZip(f); });

export async function loadZip(file) {
  $("#zipErr").textContent = ""; $("#dropTitle").textContent = "Leyendo " + file.name + "…";
  try {
    const zip = await JSZip.loadAsync(file);
    const entries = zip.file(/\.xlsx$/i).filter(e => !/(^|\/)(~\$|__MACOSX)/.test(e.name));
    if (!entries.length) throw new Error("El zip no contiene ningún Excel (.xlsx)");
    const items = [], errors = [];
    for (const e of entries) {
      const name = e.name.split("/").pop(), data = await e.async("uint8array");
      try { items.push({file: name, data, book: await readWorkbook(JSZip, data)}); }
      catch (err) { errors.push(`${name}: ${err.message}`); }
    }
    const ref = buildReference(items.map(x => x.book));
    for (const it of items) it.ev = buildEvidence(it.book, ref, it.file);
    Object.assign(S, {zipName: file.name, items, ref, decisions: {}, results: [], usage: []});
    $("#dropTitle").textContent = "✓ " + file.name;
    $("#dropSub").textContent = `${items.length} entregas leídas` + (errors.length ? ` · ${errors.length} no se pudieron leer` : "");
    if (errors.length) $("#zipErr").textContent = errors.join(" · ");
    buildOutputs("rules-preview");
  } catch (err) { $("#dropTitle").textContent = "Arrastra aquí el zip de Canvas"; $("#zipErr").textContent = err.message; }
  refreshButtons();
}

// ------------------------------------------------------------------ grading
function finalOf(it) {
  const d = S.decisions[it.ev.student];
  const final = {summary: d?.summary || ""};
  for (const q of it.ev.questions) {
    const x = d?.questions?.[String(q.q)];
    final[q.q] = x ? {score: x.score, verdict: x.verdict, comment: x.comment, basis: x.basis || ""}
      : {score: q.proposal.score, verdict: q.proposal.verdict, comment: q.proposal.comment, basis: ""};
  }
  return final;
}
function readCellsFor(it) {
  const sub = it.ev._sub;
  return rng => {
    const m = /^\s*\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?\s*$/.exec(String(rng).toUpperCase());
    if (!m) return "invalid range; use A1 notation like B145:B155";
    let c1 = colIndex(m[1]), r1 = +m[2], c2 = m[3] ? colIndex(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    [c1, c2] = [Math.min(c1, c2), Math.max(c1, c2)]; [r1, r2] = [Math.min(r1, r2), Math.max(r1, r2)];
    if ((c2 - c1 + 1) * (r2 - r1 + 1) > 300) return "range too large: at most 300 cells per call";
    const out = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const f = sub.F(r, c), v = sub.V(r, c);
      if (f || (v !== null && v !== "")) out.push(`${colLetter(c)}${r}: ${show(v)}` + (f ? ` | ${f}` : " | typed"));
    }
    return out.join("\n") || "all cells in the range are empty";
  };
}
async function pool(items, n, fn) {
  let i = 0;
  await Promise.all(Array.from({length: Math.min(n, items.length)}, async () => { while (i < items.length) { const it = items[i++]; await fn(it); } }));
}
function progress(step, pct, text) {
  $("#progCard").hidden = false;
  const steps = ["Leer entregas", "Agente", "Calibración", "Excel y justificaciones"];
  $("#steps").innerHTML = steps.map((s, i) => `<span class="pill ${i < step ? "done" : i === step ? "on" : ""}">${i < step ? "✓ " : ""}${s}</span>`).join("");
  $("#bar").style.width = pct + "%"; $("#status").innerHTML = text;
}

$("#goAgent").onclick = async () => {
  $("#gradeErr").textContent = "";
  let key;
  try { key = await keyFromForm(); } catch (e) { $("#gradeErr").textContent = e.message; return; }
  const client = window.__q1fakeClient || new Anthropic({apiKey: key, dangerouslyAllowBrowser: true, maxRetries: 4});  // fake client: tests only
  S.running = true; refreshButtons();
  try {
    progress(1, 5, "Comprobando la clave…");
    try { await client.models.retrieve(MODEL); }
    catch (e) { throw new Error("No se pudo conectar con Claude: revisa la clave sk-ant-… y que la cuenta tenga saldo. (" + (e.status || "") + " " + (e.message || "") + ")"); }
    const system = systemBlocks(expectedSummary(S.ref));
    const todo = S.items.filter(it => $("#regrade").checked || !S.decisions[it.ev.student]);
    let done = 0;
    progress(1, 8, `El agente está corrigiendo: 0 de ${todo.length} alumnos`);
    const failed = [];
    await pool(todo, 4, async it => {
      try {
        const d = await gradeStudent(client, {system, evidenceMd: evidenceMd(it.ev), file: it.file, readCells: readCellsFor(it)});
        S.decisions[it.ev.student] = d; S.usage.push(...d._meta.usage);
      } catch (e) { failed.push(`${it.ev.student}: ${e.message}`); }
      done++;
      progress(1, 8 + 80 * done / todo.length, `El agente está corrigiendo: ${done} de ${todo.length} alumnos · último: <b>${esc(it.ev.student)}</b> · coste aprox. ${costOf(S.usage).toFixed(2)} $`);
      buildOutputs("partial");
    });
    if (Object.keys(S.decisions).length > 1 && todo.length > failed.length) {
      progress(2, 90, "Revisando que el mismo error tenga la misma nota en toda la clase…");
      const evs = S.items.map(it => it.ev);
      const sec = {};
      for (const it of S.items) { sec[it.ev.student] = {}; for (const blk of evidenceMd(it.ev).split(/\n(?=## Q\d+ )/)) { const m = /^## Q(\d+) /.exec(blk); if (m) sec[it.ev.student][+m[1]] = blk; } }
      try {
        const cal = await calibrate(client, {system, evidences: evs, decisions: S.decisions, sectionsOf: st => sec[st]});
        if (cal.usage) S.usage.push(cal.usage);
        S.calibration = cal;
      } catch (e) { S.calibration = {changes: [], notes: "calibration skipped: " + e.message}; }
    }
    progress(3, 95, "Escribiendo los Excel corregidos y las justificaciones…");
    await buildOutputs("final");
    const cal = S.calibration?.changes?.length ? ` · ${S.calibration.changes.length} nota(s) igualadas en la calibración` : "";
    progress(4, 100, `<span class="ok-txt">✓ Corregidos ${Object.keys(S.decisions).length} alumnos${cal} · coste aprox. ${costOf(S.usage).toFixed(2)} $</span>` +
      (failed.length ? `<div class="err">Sin agente (propuesta automática): ${esc(failed.join(" · "))}</div>` : ""));
    S.downloaded = false;
  } catch (e) { progress(1, 100, `<span class="err">${esc(e.message)}</span>`); }
  S.running = false; refreshButtons();
};
$("#goRules").onclick = async () => {
  if (!confirm("Esto pone las notas automáticas del motor, sin el agente (gratis). ¿Continuar?")) return;
  S.decisions = {}; await buildOutputs("final"); progress(4, 100, '<span class="ok-txt">✓ Notas automáticas (sin IA) listas para revisar y descargar.</span>'); S.downloaded = false; refreshButtons();
};
function refreshButtons() {
  const ready = S.items.length > 0 && !S.running;
  $("#goAgent").disabled = !ready; $("#goRules").disabled = !ready;
  for (const id of ["dlExcel", "dlCsv", "dlAll"]) $("#" + id).disabled = !S.results.length || S.running;
  $("#goHint").textContent = S.running ? "Corrigiendo…" : S.items.length ? `${S.items.length} entregas listas` : "Primero arrastra el zip";
}

// ------------------------------------------------------------------ outputs
async function buildOutputs(mode) {
  const out = [];
  for (const it of S.items) {
    const final = finalOf(it);
    const scores = Array.from({length: 14}, (_, i) => final[i + 1].score);
    const grade = Math.round(scores.reduce((a, b) => a + b, 0) / 14 * 10 * 100) / 100;
    const source = S.decisions[it.ev.student] ? "agent" : "rules";
    const row = {student: it.ev.student, id: it.ev.id, file: it.file, score: grade, source, final, it};
    if (mode === "final") {
      const {bytes, info} = await writeWorkbook(JSZip, it.data, it.ev._sub, final);
      if (Math.abs(info.grade100 - grade) > 1e-9) throw new Error(`grade mismatch for ${it.ev.student}`);
      Object.assign(row, {bytes, info, just: justificationMd(it.ev, final, source, info, grade, S.decisions[it.ev.student]?._meta || {})});
    }
    out.push(row);
  }
  S.results = out; S.mode = mode;
  renderResults(); refreshButtons();
}
function cls(s, v) { if (v === "blank") return "nil"; return s >= 9 ? "ok" : s >= 3 ? "mid" : "bad"; }
function renderResults() {
  const rows = [...S.results].sort((a, b) => a.student.localeCompare(b.student));
  const sc = rows.map(r => r.score).sort((a, b) => a - b);
  $("#resNote").textContent = S.mode === "rules-preview" ? "Vista previa con las notas automáticas del motor (aún sin agente)." : S.mode === "partial" ? "Corrigiendo…" : "";
  $("#stats").innerHTML = sc.length ? [["Alumnos", sc.length], ["Media", fmt(sc.reduce((a, b) => a + b, 0) / sc.length)], ["Mediana", fmt(sc[Math.floor(sc.length / 2)])],
    ["Máxima", fmt(sc.at(-1))], ["Mínima", fmt(sc[0])], ["Aprobados (≥50)", sc.filter(x => x >= 50).length]].map(([k, v]) => `<div class="stat"><b>${v}</b><span>${k}</span></div>`).join("") : "";
  $("#rows").innerHTML = rows.length ? rows.map((r, i) => `<tr>
    <td class="name">${esc(r.student)}${r.source === "rules" ? ' <span class="badge b-nil" title="Propuesta automática del motor, sin agente">auto</span>' : ""}</td>
    <td><span class="grade badge ${r.score >= 70 ? "b-ok" : r.score >= 50 ? "b-mid" : "b-bad"}">${fmt(r.score)}</span></td>
    <td><div class="cells">${Array.from({length: 14}, (_, k) => { const q = r.final[k + 1]; return `<div class="c ${cls(q.score, q.verdict)}" data-s="${esc(r.student)}" data-q="${k + 1}">${q.verdict === "blank" ? "–" : String(q.score).replace(".", ",")}</div>`; }).join("")}</div></td>
    <td style="white-space:nowrap"><button class="ghost small" data-just="${esc(r.student)}">Justificación</button> <button class="ghost small" data-xl="${esc(r.student)}" ${r.bytes ? "" : "disabled"}>Excel</button></td></tr>`).join("")
    : '<tr><td colspan="4" class="empty">Arrastra el zip de Canvas para empezar.</td></tr>';
}
const byStudent = s => S.results.find(r => r.student === s);
const tip = $("#tip");
$("#rows").addEventListener("mouseover", e => {
  const c = e.target.closest(".c"); if (!c) { tip.style.opacity = 0; return; }
  const q = byStudent(c.dataset.s).final[c.dataset.q];
  tip.innerHTML = `<b>Q${c.dataset.q} · ${String(q.score).replace(".", ",")}/10</b><br>${esc(q.comment)}`;
  const r = c.getBoundingClientRect();
  tip.style.left = Math.min(r.left, innerWidth - 400) + "px"; tip.style.top = (r.bottom + 8) + "px"; tip.style.opacity = 1;
});
$("#rows").addEventListener("mouseleave", () => tip.style.opacity = 0);
$("#rows").addEventListener("click", e => {
  const c = e.target.closest(".c"), j = e.target.closest("[data-just]"), x = e.target.closest("[data-xl]");
  if (c) showJust(c.dataset.s, +c.dataset.q);
  if (j) showJust(j.dataset.just);
  if (x) { const r = byStudent(x.dataset.xl); if (r.bytes) download(new Blob([r.bytes]), r.file); }
});
function showJust(student, q) {
  const r = byStudent(student);
  const md = r.just || justificationMd(r.it.ev, r.final, r.source, {grade_cell: "-", points_col: "-", comment_col: "-"}, r.score, S.decisions[student]?._meta || {});
  $("#md").innerHTML = md2html(md); $("#modal").classList.add("open");
  if (q) setTimeout(() => { const h = document.getElementById("jq" + q); if (h) $("#modal").scrollTop = h.offsetTop - 20; }, 30);
}
$("#closeModal").onclick = () => $("#modal").classList.remove("open");
$("#modal").onclick = e => { if (e.target.id === "modal") $("#modal").classList.remove("open"); };
document.addEventListener("keydown", e => { if (e.key === "Escape") $("#modal").classList.remove("open"); });
function md2html(md) {
  const out = [], lines = md.split("\n"); let i = 0;
  const inline = s => esc(s).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/`([^`]+)`/g, "<code>$1</code>");
  while (i < lines.length) {
    const l = lines[i];
    if (l.startsWith("```")) { const buf = []; i++; while (i < lines.length && !lines[i].startsWith("```")) buf.push(lines[i++]); i++; out.push(`<pre>${esc(buf.join("\n"))}</pre>`); continue; }
    if (/^\|/.test(l)) { const rows = []; while (i < lines.length && /^\|/.test(lines[i])) rows.push(lines[i++]);
      const cells = r => r.replace(/^\||\|$/g, "").split("|").map(s => s.trim());
      out.push("<table><thead><tr>" + cells(rows[0]).map(c => `<th>${inline(c)}</th>`).join("") + "</tr></thead><tbody>" + rows.slice(2).map(r => "<tr>" + cells(r).map(c => `<td>${inline(c)}</td>`).join("") + "</tr>").join("") + "</tbody></table>"); continue; }
    let m;
    if ((m = l.match(/^## (Q(\d+).*)$/))) out.push(`<h2 id="jq${m[2]}">${inline(m[1])}</h2>`);
    else if ((m = l.match(/^# (.*)$/))) out.push(`<h1>${inline(m[1])}</h1>`);
    else if ((m = l.match(/^- (.*)$/))) { const buf = [m[1]]; while (i + 1 < lines.length && /^- /.test(lines[i + 1])) buf.push(lines[++i].slice(2)); out.push("<ul>" + buf.map(b => `<li>${inline(b)}</li>`).join("") + "</ul>"); }
    else if (l.trim()) out.push(`<p>${inline(l)}</p>`);
    i++;
  }
  return out.join("");
}

// ------------------------------------------------------------------ downloads
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
const assignment = () => $("#asgName").value.trim() || "Quiz 1";
const feedbackLines = r => [`${assignment()}: ${r.score.toFixed(2)}/100`, ...Array.from({length: 14}, (_, k) => `Q${k + 1} (${r.final[k + 1].score}/10): ${r.final[k + 1].comment}`),
  ...(r.final.summary ? ["Overall: " + r.final.summary] : []), "Your corrected workbook is attached: points and the reason for each question are in the correction table (columns 'points' and 'comment')."].join("\n");
function canvasCsv() {
  const id = $("#asgId").value.trim(), pts = +$("#asgPts").value || 100;
  const col = id ? `${assignment()} (${id})` : assignment();
  return toCsv([["Student", "ID", "SIS User ID", "SIS Login ID", "Section", col], ["Points Possible", "", "", "", "", pts],
    ...S.results.map(r => [r.student, r.id, "", "", "", Math.round(r.score / 100 * pts * 100) / 100])]);
}
$("#dlExcel").onclick = async () => {
  const z = new JSZip();
  for (const r of S.results) if (r.bytes) z.file(r.file, r.bytes);
  download(await z.generateAsync({type: "blob", compression: "DEFLATE"}), "Quiz_feedback_files.zip"); S.downloaded = true;
};
$("#dlCsv").onclick = () => download(new Blob([canvasCsv()], {type: "text/csv"}), "canvas_gradebook_import.csv");
$("#dlAll").onclick = async () => {
  const z = new JSZip();
  z.file("canvas_gradebook_import.csv", canvasCsv());
  z.file("grades.csv", toCsv([["student", "id", "score", ...Array.from({length: 14}, (_, k) => `Q${k + 1}`), "source", "file"],
    ...S.results.map(r => [r.student, r.id, r.score, ...Array.from({length: 14}, (_, k) => r.final[k + 1].score), r.source, r.file])]));
  z.file("grades_with_feedback.csv", toCsv([["student", "id", "score", "feedback"], ...S.results.map(r => [r.student, r.id, r.score, feedbackLines(r)])]));
  for (const r of S.results) {
    if (r.just) z.file(`justifications/${r.student}.md`, r.just);
    z.file(`evidence/${r.student}.md`, evidenceMd(r.it.ev));
    if (S.decisions[r.student]) z.file(`decisions/${r.student}.json`, JSON.stringify({student: r.student, ...S.decisions[r.student]}, null, 1));
  }
  const sc = S.results.map(r => r.score);
  z.file("report.md", `# ${assignment()} -- graded in the browser\n\n${S.results.length} students · mean ${(sc.reduce((a, b) => a + b, 0) / sc.length).toFixed(2)}\n\n| student | grade | source |\n|---|---|---|\n` +
    [...S.results].sort((a, b) => a.student.localeCompare(b.student)).map(r => `| ${r.student} | ${r.score.toFixed(2)} | ${r.source} |`).join("\n") + "\n");
  download(await z.generateAsync({type: "blob", compression: "DEFLATE"}), `${assignment().replace(/\W+/g, "_")}_informe.zip`); S.downloaded = true;
};
window.addEventListener("beforeunload", e => { if (!S.downloaded) { e.preventDefault(); e.returnValue = ""; } });

renderKey(); refreshButtons(); renderResults();
