// agent.js -- the grading agent: Claude (Anthropic API) called from the browser with the
// instructor's own key.  One conversation per student (evidence + rubric, read-only
// read_cells tool on the student's workbook, structured JSON answer), then one calibration
// call across the class.  Port of _tools/grade_quiz1.py.  Quiz 1 is the default; other
// assignments (Homework 1) pass cfg = {nq, sheet, instructions, rubric, withFix, example}.
import {RUBRIC} from "./rubric.js?v=4";

export const MODEL = "claude-opus-5-5";
const FALLBACK_BETA = "server-side-fallback-2026-07-01";
const VERDICTS = ["correct", "minor", "carried_forward", "partial", "major", "hardcoded", "manual", "error", "blank"];
// one compact item schema repeated in an array: a schema with 14 named properties compiles
// into a grammar the API rejects as too large
function decisionSchema(nq, withFix) {
  const props = {q: {type: "integer", enum: Array.from({length: nq}, (_, i) => i + 1)}, score: {type: "number"}, verdict: {type: "string", enum: VERDICTS},
    comment: {type: "string"}, basis: {type: "string"}};
  const req = ["q", "score", "verdict", "comment", "basis"];
  if (withFix) { props.fix = {type: "string"}; req.push("fix"); }
  return {type: "object", properties: {questions: {type: "array", items: {type: "object", properties: props, required: req, additionalProperties: false}}, summary: {type: "string"}},
    required: ["questions", "summary"], additionalProperties: false};
}
function calibSchema(withFix) {
  const props = {student: {type: "string"}, question: {type: "integer"}, score: {type: "number"}, verdict: {type: "string", enum: VERDICTS}, comment: {type: "string"}, basis: {type: "string"}};
  const req = ["student", "question", "score", "verdict", "comment", "basis"];
  if (withFix) { props.fix = {type: "string"}; req.push("fix"); }
  return {type: "object", properties: {changes: {type: "array", items: {type: "object", properties: props, required: req, additionalProperties: false}},
    notes: {type: "string"}}, required: ["changes", "notes"], additionalProperties: false};
}
const QUIZ1 = {nq: 14, sheet: "Data&Q", withFix: false, example: "'B145:B155' or 'Q26:X46'"};
const readTool = (sheet, example) => ({name: "read_cells",
  description: `Read the student's sheet '${sheet}' (read-only). Returns the formula and the cached value of every non-empty cell in an A1 range (at most 300 cells). Use it only to check something the evidence leaves unclear, e.g. ${example}.`,
  input_schema: {type: "object", properties: {range: {type: "string", description: "A1 range such as B145:B155"}}, required: ["range"], additionalProperties: false}});

const INSTRUCTIONS = `You are the grader of Quiz 1 (sports) of BTM-2500, an Excel course at SLU Madrid. You mark like a careful instructor: you reward correct logic, penalise each mistake once, and tell every student exactly why they lost points.

For each student you receive an evidence packet produced by the grading engine. The engine has already recomputed every expected value from the raw data, compared all 179 rows of each column, found answers placed in unexpected cells, traced references, matched typical mistakes, and checked whether a wrong value comes from the student's own wrong helper columns. Trust its numbers.

Decide Q1-Q14 with the rubric below. Start from the engine's proposal: keep it when it fits the rubric, change it when the evidence shows the rules misread the case (hand-tailored formula shapes, a fixed-position MID that is right only by coincidence, an IFERROR that hard-codes a value, a correct answer in a nearby cell, conditions borrowed from data cells, ...). If something is unclear, call read_cells on the student's workbook before deciding.

Write \`basis\` for the instructor: one short sentence with the evidence behind the mark (always; it is the justification kept on file). Write \`comment\` for the student exactly as the rubric's comment rules say. Return only the JSON object: "questions" is a list with exactly one item per question, q = 1 to 14.

===== RUBRIC =====
`;

export function systemBlocks(expected, cfg = null) {
  const instructions = cfg?.instructions ?? INSTRUCTIONS, rubric = cfg?.rubric ?? RUBRIC;
  return [{type: "text", text: instructions + rubric.replace("{{EXPECTED}}", expected), cache_control: {type: "ephemeral"}}];
}

async function call(client, {model, effort, schema, system, messages, tools}) {
  const body = {model, max_tokens: 16000, system, messages, betas: [FALLBACK_BETA], fallbacks: "default",
    output_config: {effort, format: {type: "json_schema", schema}}};
  if (tools) body.tools = tools;
  return client.beta.messages.create(body);
}

function validate(d, nq) {
  const qs = {};
  for (const it of Array.isArray(d.questions) ? d.questions : []) qs[String(it.q)] = it;
  d.questions = qs;
  for (const k of Array.from({length: nq}, (_, i) => String(i + 1))) {
    const q = qs[k];
    if (!q) throw new Error(`missing Q${k}`);
    const sc = Number(q.score);
    if (!(sc >= 0 && sc <= 10)) throw new Error(`Q${k} score ${q.score} outside 0-10`);
    q.score = Math.round(sc * 2) / 2;
    if (!String(q.comment || "").trim()) throw new Error(`Q${k} has no comment`);
  }
  return d;
}

export const usageOf = r => ({input_tokens: r.usage?.input_tokens || 0, output_tokens: r.usage?.output_tokens || 0,
  cache_read_input_tokens: r.usage?.cache_read_input_tokens || 0, cache_creation_input_tokens: r.usage?.cache_creation_input_tokens || 0});

export async function gradeStudent(client, {model = MODEL, effort = "high", system, evidenceMd, file, readCells, onEvent = () => {}, cfg = null}) {
  const {nq, sheet, withFix, example} = {...QUIZ1, ...(cfg || {})};
  const schema = decisionSchema(nq, withFix), tool = readTool(sheet, example);
  const first = {role: "user", content: `Student workbook: ${file}\n\n${evidenceMd}\n\nGrade Q1-Q${nq} now.`};
  let lastErr = null;
  const messages = [first], toolCalls = [], usage = [];
  let resp;
  for (let attempt = 0; attempt < 3; attempt++) {
    for (let turn = 0; turn < 10; turn++) {
      resp = await call(client, {model, effort, schema, system, messages, tools: [tool]});
      usage.push(usageOf(resp));
      if (resp.stop_reason === "refusal") throw new Error("the model declined this request");
      if (resp.stop_reason !== "tool_use") break;
      messages.push({role: "assistant", content: resp.content});
      const results = [];
      for (const b of resp.content) {
        if (b.type !== "tool_use") continue;
        const rng = String((b.input || {}).range || "");
        const out = readCells(rng);
        toolCalls.push({range: rng});
        onEvent({type: "tool", range: rng});
        results.push({type: "tool_result", tool_use_id: b.id, content: out});
      }
      messages.push({role: "user", content: results});
    }
    try {
      if (resp.stop_reason === "max_tokens") throw new Error("answer cut at max_tokens");
      const text = resp.content.find(b => b.type === "text")?.text;
      if (!text) throw new Error("no text block in the answer");
      const d = validate(JSON.parse(text), nq);
      d._meta = {model: resp.model, effort, graded_at: new Date().toISOString().slice(0, 19), tool_calls: toolCalls, usage};
      return d;
    } catch (e) {
      lastErr = e;
      // same conversation: show the model what was wrong and ask for the full answer
      messages.push({role: "assistant", content: resp.content});
      messages.push({role: "user", content: `That answer is not usable (${e.message}). Return the complete JSON again: "questions" must contain exactly ${nq} items, one for each q = 1 to ${nq}, plus "summary".`});
    }
  }
  throw new Error("no valid decision after 3 attempts: " + (lastErr?.message || ""));
}

export function calibrationGroups(evidences, decisions) {
  const groups = new Map();
  for (const ev of evidences) {
    if (!decisions[ev.student]) continue;
    for (const q of ev.questions) {
      const key = `${q.q}|${q.proposal.verdict}|${q.proposal.score}`;
      if (!groups.has(key)) groups.set(key, {q: q.q, verdict: q.proposal.verdict, score: q.proposal.score, students: []});
      groups.get(key).students.push(ev.student);
    }
  }
  return [...groups.values()].filter(g => new Set(g.students.map(s => decisions[s].questions[String(g.q)].score)).size > 1);
}

export async function calibrate(client, {model = MODEL, system, evidences, decisions, sectionsOf, cfg = null}) {
  const withFix = !!cfg?.withFix;
  const groups = calibrationGroups(evidences, decisions);
  if (!groups.length) return {changes: [], notes: "no inconsistent group", usage: null};
  const L = ["Calibration across the whole class. Each group below holds students whose answer to one question the engine classified the same way, but who got different marks. For each group decide whether the evidence really differs. Where it is the same mistake, align the marks to the rubric and return a change (new score, verdict, student comment, and a basis starting with 'calibration:'). Where the evidence differs, keep the marks and return nothing for them. Return only real changes."];
  for (const g of groups) {
    L.push(`\n### Q${g.q} -- engine class '${g.verdict}', proposal ${g.score}`);
    for (const st of g.students) {
      const dq = decisions[st].questions[String(g.q)];
      L.push(`\n#### ${st}: mark ${dq.score} (${dq.verdict})\ncomment: ${dq.comment}\nbasis: ${dq.basis || ""}\nevidence:\n${(sectionsOf(st)[g.q] || "").slice(0, 1800)}`);
    }
  }
  const resp = await call(client, {model, effort: "high", schema: calibSchema(withFix), system, messages: [{role: "user", content: L.join("\n")}]});
  if (resp.stop_reason === "refusal" || resp.stop_reason === "max_tokens") return {changes: [], notes: resp.stop_reason, usage: usageOf(resp)};
  const res = JSON.parse(resp.content.find(b => b.type === "text").text);
  const changes = [];
  for (const ch of res.changes || []) {
    const d = decisions[ch.student], q = String(ch.question);
    if (!d || !d.questions[q]) continue;
    const sc = Math.round(Number(ch.score) * 2) / 2;
    if (!(sc >= 0 && sc <= 10)) continue;
    const old = d.questions[q].score;
    Object.assign(d.questions[q], {score: sc, verdict: ch.verdict, comment: ch.comment, basis: ch.basis}, withFix ? {fix: ch.fix || ""} : {});
    (d._meta.calibration ||= []).push({question: +q, old, new: sc});
    changes.push({student: ch.student, question: +q, old, new: sc, basis: ch.basis});
  }
  return {changes, notes: res.notes || "", usage: usageOf(resp)};
}

// rough cost of the run (Claude Opus 5.5 prices, USD per million tokens)
export function costOf(usages) {
  let c = 0;
  for (const u of usages) c += (u.input_tokens * 4 + u.output_tokens * 20 + u.cache_read_input_tokens * 0.2 + u.cache_creation_input_tokens * 5) / 1e6;
  return c;
}
