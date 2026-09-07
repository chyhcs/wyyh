// English framework deck for the iFLYTEK / RNP technical meeting.
// 10 slides: logic, titles and section labels are filled in; content is left
// as bracketed placeholders for manual editing.
//
//   node build_framework_deck.js   ->  RNP-iFLYTEK-Framework-EN.pptx

const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.author = "iFLYTEK";
pres.title = "iFLYTEK scope, joint implementation and knowledge transfer";

const F = "Calibri";
const W = 13.333;
const H = 7.5;
const M = 0.5; // outer margin

// Palette: shared with the procurement-process deck
const NAVY = "10305F";
const BLUE = "1A5FBF";
const TEAL = "0E8577";
const TEAL_LT = "E7F3F0";
const AMBER = "B26A00";
const AMBER_LT = "FBF1E1";
const PANEL_BG = "F7F9FC";
const PANEL_LN = "DCE4EF";
const CARD_LN = "D9E2EC";
const GRAY = "4B5A6B";
const GRAY_L = "6B7A8B";
const PH = "8A96A6"; // placeholder text
const WHITE = "FFFFFF";
const ICE = "CADCFC";

const TOTAL = 10;
let slideNo = 0;

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "9FB0C4", blur: 5, offset: 1, angle: 90, opacity: 0.22 });

function newSlide(opts) {
  const s = pres.addSlide();
  slideNo += 1;
  s.background = { color: opts && opts.dark ? NAVY : WHITE };
  if (!(opts && opts.noFooter)) {
    s.addText("iFLYTEK  ·  RNP technical meeting  ·  Working draft for discussion", {
      x: M, y: H - 0.45, w: 8, h: 0.3, margin: 0, isTextBox: true,
      fontFace: F, fontSize: 9, color: opts && opts.dark ? ICE : GRAY_L,
    });
    s.addText(`${slideNo} / ${TOTAL}`, {
      x: W - M - 1.2, y: H - 0.45, w: 1.2, h: 0.3, margin: 0, isTextBox: true, align: "right",
      fontFace: F, fontSize: 9, color: opts && opts.dark ? ICE : GRAY_L,
    });
  }
  return s;
}

function title(s, text, sub) {
  s.addText(text, {
    x: M, y: 0.38, w: W - 2 * M, h: 0.62, margin: 0, isTextBox: true,
    fontFace: F, fontSize: 24, bold: true, color: NAVY, valign: "middle",
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.02, w: W - 2 * M, h: 0.36, margin: 0, isTextBox: true,
      fontFace: F, fontSize: 13, color: GRAY, valign: "middle",
    });
  }
}

function card(s, x, y, w, h, fill, line) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fill || WHITE },
    line: { color: line || CARD_LN, width: 0.75 },
    shadow: shadow(),
  });
}

function flat(s, x, y, w, h, fill, line) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fill || PANEL_BG },
    line: { color: line || PANEL_LN, width: 0.75 },
  });
}

function label(s, x, y, w, text, color, size) {
  s.addText(text, {
    x, y, w, h: 0.3, margin: 0, isTextBox: true,
    fontFace: F, fontSize: size || 12, bold: true, color: color || NAVY, valign: "middle",
  });
}

function ph(s, x, y, w, h, text, size) {
  s.addText(text, {
    x, y, w, h, margin: 0, isTextBox: true, valign: "top",
    fontFace: F, fontSize: size || 10.5, italic: true, color: PH,
  });
}

function body(s, x, y, w, h, text, size, color) {
  s.addText(text, {
    x, y, w, h, margin: 0, isTextBox: true, valign: "top",
    fontFace: F, fontSize: size || 11, color: color || GRAY,
  });
}

function badge(s, x, y, d, text, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: color || TEAL }, line: { color: color || TEAL, width: 0 } });
  s.addText(text, {
    x, y, w: d, h: d, margin: 0, isTextBox: true, align: "center", valign: "middle",
    fontFace: F, fontSize: 10, bold: true, color: WHITE,
  });
}

function chip(s, x, y, w, text, color) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.3, rectRadius: 0.15, fill: { color }, line: { color, width: 0 } });
  s.addText(text, {
    x, y, w, h: 0.3, margin: 0, isTextBox: true, align: "center", valign: "middle",
    fontFace: F, fontSize: 9, bold: true, color: WHITE, charSpacing: 0.5,
  });
}

function arrowRight(s, x, y, w, h) {
  s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill: { color: PANEL_LN }, line: { color: PANEL_LN, width: 0 } });
}

// step boxes in a horizontal flow
function flow(s, x, y, w, h, steps, opts) {
  const gap = 0.32;
  const bw = (w - gap * (steps.length - 1)) / steps.length;
  steps.forEach((st, i) => {
    const bx = x + i * (bw + gap);
    flat(s, bx, y, bw, h, i === 0 ? TEAL_LT : PANEL_BG);
    badge(s, bx + 0.12, y + 0.12, 0.3, String(i + 1), i === 0 ? TEAL : NAVY);
    s.addText(st.name, {
      x: bx + 0.5, y: y + 0.08, w: bw - 0.6, h: 0.38, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 11.5, bold: true, color: NAVY,
    });
    if (st.hint) ph(s, bx + 0.12, y + 0.55, bw - 0.24, h - 0.65, st.hint, 10);
    if (i < steps.length - 1) arrowRight(s, bx + bw + 0.06, y + h / 2 - 0.1, 0.2, 0.2);
  });
  return bw;
}

// table cell helpers
const th = (t, fill, color) => ({ text: t, options: { bold: true, color: color || WHITE, fill: { color: fill || NAVY }, fontSize: 10.5, valign: "middle" } });
const td = (t, opt) => ({ text: t, options: Object.assign({ color: PH, italic: true, fontSize: 10, valign: "top" }, opt || {}) });
const tb = (t, opt) => ({ text: t, options: Object.assign({ color: NAVY, bold: true, fontSize: 10.5, valign: "top" }, opt || {}) });

const tableBase = () => ({
  fontFace: F, fontSize: 10, color: GRAY,
  border: { type: "solid", pt: 0.5, color: PANEL_LN },
  margin: [0.05, 0.08, 0.05, 0.08],
  valign: "top",
});

// =====================================================================
// 1. COVER
// =====================================================================
{
  const s = newSlide({ dark: true, noFooter: true });

  s.addText("RNP  ×  iFLYTEK   ·   Technical meeting", {
    x: M + 0.3, y: 0.7, w: 8, h: 0.35, margin: 0, isTextBox: true,
    fontFace: F, fontSize: 12, bold: true, color: ICE, charSpacing: 1,
  });
  s.addText("Portuguese LLM, Application\nand Domain Models", {
    x: M + 0.3, y: 1.5, w: 8.2, h: 2.0, margin: 0, isTextBox: true, valign: "top",
    fontFace: F, fontSize: 36, bold: true, color: WHITE, lineSpacingMultiple: 1.05,
  });
  s.addText("iFLYTEK scope, joint implementation with RNP,\nand knowledge transfer", {
    x: M + 0.3, y: 3.65, w: 8.2, h: 1.0, margin: 0, isTextBox: true, valign: "top",
    fontFace: F, fontSize: 20, color: ICE, lineSpacingMultiple: 1.1,
  });

  chip(s, M + 0.3, 5.15, 2.9, "WORKING DRAFT FOR DISCUSSION", TEAL);
  s.addText(
    [
      { text: "Rio de Janeiro  ·  [date]", options: { breakLine: true } },
      { text: "Presented by iFLYTEK  ·  [presenter names and roles]", options: { breakLine: true } },
      { text: "Version [x]  ·  content to be confirmed against the SOW, quotation and technical plan" },
    ],
    {
      x: M + 0.3, y: 5.65, w: 9, h: 1.1, margin: 0, isTextBox: true, valign: "top",
      fontFace: F, fontSize: 12, color: ICE, lineSpacingMultiple: 1.25,
    }
  );

  // motif: three deliverable blocks + knowledge transfer band
  const bx = 9.55, by = 1.6, bw = 3.1, bh = 0.95, g = 0.22;
  ["Portuguese LLM", "Application (APP)", "Domain Models (DLM)"].forEach((t, i) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: bx, y: by + i * (bh + g), w: bw, h: bh, rectRadius: 0.1,
      fill: { color: "1B3F78" }, line: { color: "2E5AA8", width: 0.75 },
    });
    badge(s, bx + 0.2, by + i * (bh + g) + (bh - 0.36) / 2, 0.36, String(i + 1), TEAL);
    s.addText(t, {
      x: bx + 0.7, y: by + i * (bh + g), w: bw - 0.8, h: bh, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 15, bold: true, color: WHITE,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: bx, y: by + 3 * (bh + g), w: bw, h: 0.7, rectRadius: 0.1,
    fill: { color: TEAL }, line: { color: TEAL, width: 0 },
  });
  s.addText("Knowledge transfer · joint implementation", {
    x: bx + 0.2, y: by + 3 * (bh + g), w: bw - 0.4, h: 0.7, margin: 0, isTextBox: true, valign: "middle", align: "center",
    fontFace: F, fontSize: 12.5, bold: true, color: WHITE,
  });

  s.addNotes(
    "PURPOSE: Fix the scope of the meeting: iFLYTEK's model, application and domain-model work, joint implementation with RNP, and knowledge transfer. " +
    "RNP should see from the first page that the deck covers both deliverables and how their own people take part.\n\n" +
    "FILL: date, presenters, version/status (draft vs confirmed).\n\n" +
    "SOURCE: Mota call 01:03-02:29 (explain LLM/APP/DLM, responsibilities, RNP must allocate people and write its own plan)."
  );
}

// =====================================================================
// 2. WHAT RNP RECEIVES
// =====================================================================
{
  const s = newSlide();
  title(s, "What RNP receives from iFLYTEK", "Four outcomes. For each: the deliverable, what it is used for, and RNP's role in producing it.");

  const items = [
    { n: "1", name: "Portuguese LLM", color: TEAL },
    { n: "2", name: "Application (APP)", color: TEAL },
    { n: "3", name: "Domain Models (DLM)", color: TEAL },
    { n: "4", name: "In-project knowledge transfer", color: NAVY },
  ];
  const gap = 0.3;
  const cw = (W - 2 * M - gap * 3) / 4;
  const cy = 1.65, ch = 4.55;
  items.forEach((it, i) => {
    const cx = M + i * (cw + gap);
    card(s, cx, cy, cw, ch, WHITE);
    badge(s, cx + 0.22, cy + 0.22, 0.38, it.n, it.color);
    s.addText(it.name, {
      x: cx + 0.72, y: cy + 0.14, w: cw - 0.9, h: 0.6, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 14, bold: true, color: NAVY,
    });
    const rows = [
      ["Deliverable", "[What is handed over: model package, application, domain model, materials]"],
      ["Used for", "[The use it serves for RNP / government users]"],
      ["RNP's role", "[Actual work RNP performs, not only attendance at training]"],
    ];
    rows.forEach((r, k) => {
      const ry = cy + 0.85 + k * 1.0;
      label(s, cx + 0.22, ry, cw - 0.44, r[0], TEAL, 11);
      ph(s, cx + 0.22, ry + 0.32, cw - 0.44, 0.62, r[1], 10);
    });
    flat(s, cx + 0.22, cy + ch - 0.62, cw - 0.44, 0.4, PANEL_BG);
    s.addText("Status:  [included · optional · to confirm]", {
      x: cx + 0.32, y: cy + ch - 0.62, w: cw - 0.64, h: 0.4, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 9.5, italic: true, color: GRAY,
    });
  });

  body(s, M, 6.4, W - 2 * M, 0.4,
    "Examples, interim versions and final deliverables are labelled as such wherever they appear in this deck.", 10.5, GRAY_L);

  s.addNotes(
    "PURPOSE: Answer the first question before any detail: what iFLYTEK's part finally produces and what each item is for. " +
    "Gives every later page a fixed place.\n\n" +
    "MUST ANSWER per item: main deliverable, its use, RNP's participation (visible work, not just training attendance).\n\n" +
    "PREPARE: iFLYTEK technical plan, SOW and deliverable list. Mark whether shown material is a case, an interim version or the project deliverable.\n\n" +
    "LINKS: pages 3-5 expand items 1-3; page 8 expands item 4."
  );
}

// =====================================================================
// 3. PORTUGUESE LLM
// =====================================================================
{
  const s = newSlide();
  title(s, "Portuguese LLM: what it will do and how we build it together",
    "Target capability, build approach, data conditions, division of work, and how the result is checked.");

  // build approach flow
  label(s, M, 1.55, 6, "Build approach", NAVY, 12.5);
  flow(s, M, 1.9, W - 2 * M, 1.15, [
    { name: "Data preparation", hint: "[Sources, authorisation, data work and who does it]" },
    { name: "Adaptation / training", hint: "[Base model, method, why this route]" },
    { name: "Evaluation", hint: "[Benchmarks, test sets, RNP review]" },
    { name: "Delivery", hint: "[Model package, documents, acceptance]" },
  ]);

  // three columns
  const cy = 3.35, ch = 2.5, gap = 0.3;
  const cw = (W - 2 * M - 2 * gap) / 3;
  const cols = [
    { h: "Target capability & how it is measured", t: "[Use cases the model must support]\n\n[Evidence: samples, evaluation results, comparison]" },
    { h: "Data & conditions", t: "[Data needed and where it comes from]\n\n[Authorisation and data-processing responsibility]\n\n[Platform / compute dependencies]" },
  ];
  cols.forEach((c, i) => {
    const cx = M + i * (cw + gap);
    card(s, cx, cy, cw, ch, WHITE);
    label(s, cx + 0.22, cy + 0.18, cw - 0.44, c.h, NAVY, 12);
    ph(s, cx + 0.22, cy + 0.6, cw - 0.44, ch - 0.75, c.t, 10.5);
  });
  // who does what
  const cx = M + 2 * (cw + gap);
  card(s, cx, cy, cw, ch, WHITE);
  label(s, cx + 0.22, cy + 0.18, cw - 0.44, "Who does what", NAVY, 12);
  const half = (cw - 0.44 - 0.15) / 2;
  flat(s, cx + 0.22, cy + 0.6, half, ch - 0.8, TEAL_LT, TEAL_LT);
  flat(s, cx + 0.22 + half + 0.15, cy + 0.6, half, ch - 0.8, PANEL_BG);
  label(s, cx + 0.32, cy + 0.68, half - 0.2, "iFLYTEK", TEAL, 11);
  label(s, cx + 0.32 + half + 0.15, cy + 0.68, half - 0.2, "RNP", NAVY, 11);
  ph(s, cx + 0.32, cy + 1.02, half - 0.2, ch - 1.3, "[Technical work: data processing, training, evaluation, packaging]", 10);
  ph(s, cx + 0.32 + half + 0.15, cy + 1.02, half - 0.2, ch - 1.3, "[Requirements, data supply, evaluation, participation in each stage]", 10);

  // deliverables strip
  flat(s, M, 6.05, W - 2 * M, 0.75, PANEL_BG);
  label(s, M + 0.22, 6.12, 3, "Deliverables & acceptance", TEAL, 11.5);
  ph(s, M + 0.22, 6.42, W - 2 * M - 0.44, 0.35,
    "[Model delivery form]  ·  [Evaluation criteria and acceptance]  ·  [Accompanying documents: training data description, evaluation report, operating guide]", 10);

  s.addNotes(
    "PURPOSE: Expand 'Portuguese LLM' from page 2 so RNP understands how the model comes about, what iFLYTEK does, and what RNP must supply and take part in. " +
    "Serves both understanding of the result and RNP's own planning.\n\n" +
    "GUARDRAILS: do not presuppose training from scratch, and do not present technical selection rationale in detail - stages, division of work and results are what the audience needs. " +
    "Every capability and every data-processing task needs a named owner.\n\n" +
    "PREPARE: model plan, data requirements, evaluation material, delivery form, task split.\n\n" +
    "SOURCE: Mota call (how it is done, responsibilities); internal meeting 4 Sep (results-first, customer does not check the tech stack)."
  );
}

// =====================================================================
// 4. APPLICATION
// =====================================================================
{
  const s = newSlide();
  title(s, "Application: who uses it and what it delivers",
    "Seen from the user's side: the people, the task, the result — and what RNP provides for integration.");

  // left: visual frame
  const lx = M, ly = 1.6, lw = 5.4, lh = 4.1;
  card(s, lx, ly, lw, lh, PANEL_BG, PANEL_LN);
  s.addText("Screenshot  ·  demo video  ·  real interface", {
    x: lx + 0.3, y: ly + lh / 2 - 0.5, w: lw - 0.6, h: 0.5, margin: 0, isTextBox: true, align: "center", valign: "middle",
    fontFace: F, fontSize: 14, bold: true, color: GRAY_L,
  });
  s.addText("[Insert visual of an end-to-end user flow]", {
    x: lx + 0.3, y: ly + lh / 2 + 0.05, w: lw - 0.6, h: 0.4, margin: 0, isTextBox: true, align: "center", valign: "middle",
    fontFace: F, fontSize: 11, italic: true, color: PH,
  });
  flat(s, lx + 0.3, ly + lh - 0.62, lw - 0.6, 0.4, WHITE);
  s.addText("Label the material:  [existing case  ·  functional example  ·  project deliverable]", {
    x: lx + 0.4, y: ly + lh - 0.62, w: lw - 0.8, h: 0.4, margin: 0, isTextBox: true, valign: "middle",
    fontFace: F, fontSize: 9.5, italic: true, color: GRAY,
  });

  // right: four rows
  const rx = lx + lw + 0.4, rw = W - M - rx;
  const rows = [
    { h: "Users & scenarios", t: "[Target users and the real situations in which they use the app]" },
    { h: "Functions & results", t: "[What the app does and what the user gets at the end of the flow]" },
    { h: "Scale & performance", t: "[Users supported, concurrency, response requirements — as confirmed]" },
    { h: "Government use cases addressed", t: "[Which public-service tasks this solves; reference case if any, labelled]" },
  ];
  rows.forEach((r, i) => {
    const ry = 1.6 + i * 1.05;
    badge(s, rx, ry + 0.05, 0.34, String(i + 1), i === 3 ? NAVY : TEAL);
    label(s, rx + 0.5, ry + 0.02, rw - 0.5, r.h, NAVY, 12);
    ph(s, rx + 0.5, ry + 0.36, rw - 0.5, 0.6, r.t, 10.5);
  });

  // bottom strip: RNP to provide
  flat(s, M, 5.95, W - 2 * M, 0.85, TEAL_LT, TEAL_LT);
  label(s, M + 0.22, 6.02, 4, "RNP provides for integration", TEAL, 11.5);
  ph(s, M + 0.22, 6.33, W - 2 * M - 0.44, 0.45,
    "[Business system interfaces]  ·  [Data]  ·  [Testing and business confirmation]  ·  Integration scope to confirm: localisation, accounts & permissions, model calls, system interfaces", 10);

  s.addNotes(
    "PURPOSE: Let RNP understand the deliverable from the user's point of view: who uses it, what task it completes, what they get. " +
    "Functions belong inside a real usage flow. The audience does not want the technology stack.\n\n" +
    "MUST SHOW: a visual (screenshot, video or verified case) with its status labelled; functions and results; scale/performance as confirmed; the government use cases addressed.\n\n" +
    "KEEP ONE LINE on integration conditions so RNP's technical team can plan systems, data and tests (Mota: explain how the app is done and who does what).\n\n" +
    "SOURCE: internal meeting 4 Sep 34:06-34:46 (how many users, which functions, video or pictures, which government applications; a Chinese case may be shown as a reference)."
  );
}

// =====================================================================
// 5. DOMAIN MODELS
// =====================================================================
{
  const s = newSlide();
  title(s, "Domain Models (DLM): built together around real problems",
    "The right problem, usable data and experts who can judge the outcome — so both teams are on every step.");

  // candidate domain line
  flat(s, M, 1.6, W - 2 * M, 0.62, PANEL_BG);
  label(s, M + 0.22, 1.66, 2.4, "Candidate domain(s)", NAVY, 12);
  ph(s, M + 2.6, 1.7, 6.5, 0.45, "[Domain and objective, or the mechanism for selecting domains]", 10.5);
  chip(s, W - M - 3.0, 1.76, 2.8, "STATUS:  CANDIDATE  /  CONFIRMED", GRAY_L);

  // flow with iFLYTEK / RNP rows under each step
  const steps = ["Business problem", "Data & domain experts", "Development & adaptation", "Evaluation & acceptance", "Update & maintenance"];
  const gap = 0.28;
  const bw = (W - 2 * M - gap * 4) / 5;
  const fy = 2.5, fh = 0.62;
  steps.forEach((name, i) => {
    const bx = M + i * (bw + gap);
    flat(s, bx, fy, bw, fh, i === 0 ? TEAL_LT : PANEL_BG);
    badge(s, bx + 0.1, fy + 0.14, 0.32, String(i + 1), i === 0 ? TEAL : NAVY);
    s.addText(name, {
      x: bx + 0.5, y: fy, w: bw - 0.58, h: fh, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: NAVY,
    });
    if (i < 4) arrowRight(s, bx + bw + 0.04, fy + fh / 2 - 0.1, 0.2, 0.2);

    // iFLYTEK row
    const ry1 = fy + fh + 0.2;
    card(s, bx, ry1, bw, 1.05, WHITE);
    label(s, bx + 0.15, ry1 + 0.08, bw - 0.3, "iFLYTEK", TEAL, 10.5);
    ph(s, bx + 0.15, ry1 + 0.38, bw - 0.3, 0.6, "[Task at this step]", 10);
    // RNP row
    const ry2 = ry1 + 1.05 + 0.15;
    card(s, bx, ry2, bw, 1.05, WHITE);
    label(s, bx + 0.15, ry2 + 0.08, bw - 0.3, "RNP", NAVY, 10.5);
    ph(s, bx + 0.15, ry2 + 0.38, bw - 0.3, 0.6, "[Task at this step]", 10);
  });

  // outcomes strip
  const oy = 5.95;
  label(s, M, oy, 3, "Outcomes, listed separately", TEAL, 11.5);
  const ow = (W - 2 * M - 2 * 0.3) / 3;
  ["Model result", "Application result", "Documents & materials"].forEach((t, i) => {
    const ox = M + i * (ow + 0.3);
    flat(s, ox, oy + 0.32, ow, 0.55, PANEL_BG);
    s.addText([{ text: t + ":  ", options: { bold: true, color: NAVY, italic: false } }, { text: "[what is delivered and how it is accepted]", options: { italic: true, color: PH } }], {
      x: ox + 0.15, y: oy + 0.32, w: ow - 0.3, h: 0.55, margin: 0, isTextBox: true, valign: "middle",
      fontFace: F, fontSize: 10,
    });
  });

  s.addNotes(
    "PURPOSE: Show that domain results form around a concrete business problem and that RNP must organise data, experts and technical staff to take part. " +
    "This page is the main anchor of the joint-implementation relationship.\n\n" +
    "MUST ANSWER: domain(s) or the selection mechanism; business problem and expected result; data, experts and other conditions; iFLYTEK's development work; RNP's work; how results are evaluated, accepted and updated.\n\n" +
    "GUARDRAILS: mark candidate domains as candidates; number of domains and technical route follow the actual plan; keep model, application and documents as separate outcomes - not one vague 'solution'.\n\n" +
    "SOURCE: Mota call (DLM: how it is done, responsibilities); framework page 5."
  );
}

// =====================================================================
// 6. JOINT IMPLEMENTATION PLAN
// =====================================================================
{
  const s = newSlide();
  title(s, "Joint implementation plan: who does what, when",
    "One timeline, three parties — so RNP can build its own plan, and platform dependencies are not counted as iFLYTEK's.");

  const phases = ["Phase 1", "Phase 2", "Phase 3", "Phase 4", "Phase 5"];
  const header = [th("", NAVY)].concat(phases.map((p) => th(p + "\n[name]  ·  [timing]", NAVY)));
  const laneRows = [
    ["Objective of the phase", "[What this phase must achieve]"],
    ["iFLYTEK", "[Work iFLYTEK performs]"],
    ["RNP", "[Work RNP performs or joins]"],
    ["Huawei platform dependencies", "[Compute, environment, interfaces needed from the platform side]"],
    ["Phase deliverable / checkpoint", "[Output both sides use to judge completion]"],
    ["Preconditions to start", "[What must be in place before this phase begins]"],
  ];
  const rows = [header];
  laneRows.forEach((lr, i) => {
    const fill = lr[0] === "iFLYTEK" ? TEAL_LT : lr[0] === "RNP" ? "EEF2FA" : lr[0] === "Huawei platform dependencies" ? AMBER_LT : WHITE;
    const row = [tb(lr[0], { fill: { color: fill } })];
    phases.forEach(() => row.push(td(lr[1], { fill: { color: fill } })));
    rows.push(row);
  });
  const firstW = 2.35;
  const restW = (W - 2 * M - firstW) / 5;
  s.addTable(rows, Object.assign(tableBase(), {
    x: M, y: 1.6, w: W - 2 * M,
    colW: [firstW, restW, restW, restW, restW, restW],
    rowH: [0.55, 0.62, 0.72, 0.72, 0.72, 0.62, 0.62],
  }));

  body(s, M, 6.35, W - 2 * M, 0.45,
    "Phase names and count follow the actual plan. Data preparation, model and application build, evaluation, deployment and knowledge transfer must be visible in their relation to each other.", 10, GRAY_L);

  s.addNotes(
    "PURPOSE: Answer 'how are these pieces arranged together and when does each side need to do what'. Directly answers Mota's point that RNP must write its own participation plan; " +
    "an iFLYTEK-internal schedule alone does not meet that need.\n\n" +
    "THREE PARTIES: Mota framed responsibilities as ours / yours / theirs. Show Huawei platform dependencies (compute, environment, interfaces, Ascend stack) as a separate lane so RNP does not attribute them to iFLYTEK.\n\n" +
    "PREPARE: implementation plan, phase deliverables, responsibilities, resource dependencies, review points.\n\n" +
    "LINK: page 7 says which people and resources RNP must put in place to execute this plan."
  );
}

// =====================================================================
// 7. RNP TEAM & RESOURCES
// =====================================================================
{
  const s = newSlide();
  title(s, "What RNP needs to put in place",
    "Turn 'RNP cooperation' into things that can be scheduled: who takes part, doing what, from when, with which prerequisites.");

  // left table
  const roles = ["[e.g., Project lead]", "[e.g., Business / domain experts]", "[e.g., Data engineers]", "[e.g., Model engineers]", "[e.g., Application developers]", "[e.g., Operations]"];
  const rows = [[th("Role"), th("Tasks in the project"), th("Prerequisites"), th("Phases & time commitment"), th("Headcount")]];
  roles.forEach((r) => {
    rows.push([td(r, { color: GRAY, bold: true }), td("[ ]"), td("[language, programming, domain]"), td("[ ]"), td("[ ]")]);
  });
  const lw = 8.0;
  s.addTable(rows, Object.assign(tableBase(), {
    x: M, y: 1.6, w: lw,
    colW: [1.9, 2.05, 1.65, 1.4, 1.0],
    rowH: [0.42, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58],
  }));
  body(s, M, 5.75, lw, 0.5,
    "Distinguish people who carry project tasks from people who only attend courses. Roles follow the actual tasks — no fixed organisation is assumed.", 10, GRAY_L);

  // right: resources card
  const rx = M + lw + 0.35, rw = W - M - rx;
  card(s, rx, 1.6, rw, 4.5, WHITE);
  label(s, rx + 0.22, 1.78, rw - 0.44, "Resources", NAVY, 12.5);
  const res = [
    ["Data & authorisation", "[Datasets, ownership, licences]"],
    ["Environment & access", "[Accounts, network, test and experimentation environment]"],
    ["Interfaces & test conditions", "[Business systems, APIs, test data]"],
    ["Owner · ready-by · impact if late", "[For each item above]"],
  ];
  res.forEach((r, i) => {
    const ry = 2.2 + i * 0.92;
    badge(s, rx + 0.22, ry + 0.02, 0.3, String(i + 1), TEAL);
    label(s, rx + 0.65, ry, rw - 0.9, r[0], NAVY, 11);
    ph(s, rx + 0.65, ry + 0.32, rw - 0.9, 0.5, r[1], 10);
  });

  flat(s, M, 6.3, W - 2 * M, 0.5, TEAL_LT, TEAL_LT);
  s.addText("This page is the input for RNP's own participation plan: roles, headcount and timing, aligned with the Huawei capacity programme.", {
    x: M + 0.22, y: 6.3, w: W - 2 * M - 0.44, h: 0.5, margin: 0, isTextBox: true, valign: "middle",
    fontFace: F, fontSize: 10.5, bold: true, color: TEAL,
  });

  s.addNotes(
    "PURPOSE: RNP should leave with enough to organise its own people and resources - not just course enrolment requirements. This is the most direct answer to the Mota call " +
    "('they need to allocate people ... how many people ... they need to write down this program').\n\n" +
    "MUST ANSWER: roles by actual work, tasks per role, prerequisites (language, programming, degree), phase and time commitment, headcount; data and authorisation; environment and access; interfaces and test conditions; owner, ready-by date and impact if not ready.\n\n" +
    "GUARDRAILS: name shared work explicitly; do not hand data quality or implementation conditions wholesale to RNP. Keep the format compatible with the Huawei layer-1 talent tables so totals can be reconciled in the combined deck.\n\n" +
    "SOURCE: Mota call 01:03-02:29; internal meeting 4 Sep (entry requirements, language)."
  );
}

// =====================================================================
// 8. IN-PROJECT KNOWLEDGE TRANSFER
// =====================================================================
{
  const s = newSlide();
  title(s, "In-project knowledge transfer: learning by doing on real tasks");
  chip(s, M, 1.04, 4.6, "INCLUDED IN CURRENT PROPOSAL  ·  NO ADDITIONAL FEE", TEAL);
  s.addText("[confirm wording against the SOW and quotation]", {
    x: M + 4.75, y: 1.04, w: 6, h: 0.3, margin: 0, isTextBox: true, valign: "middle",
    fontFace: F, fontSize: 10, italic: true, color: PH,
  });

  const rows = [[
    th("Participants (role)"), th("Seats"), th("Project tasks they perform"),
    th("How iFLYTEK supports"), th("Capability outcome"), th("How it is verified"),
  ]];
  for (let i = 0; i < 4; i++) {
    rows.push([
      td("[role]", { color: GRAY, bold: true }), td("[n]"), td("[tasks]"),
      td("[briefing · demonstration · joint operation · guided practice]"), td("[what they can do afterwards]"), td("[task, experiment or output used as evidence]"),
    ]);
  }
  s.addTable(rows, Object.assign(tableBase(), {
    x: M, y: 1.55, w: W - 2 * M,
    colW: [2.0, 0.8, 2.5, 2.9, 2.2, 1.93],
    rowH: [0.45, 0.6, 0.6, 0.6, 0.6],
  }));

  // three boxes
  const by = 4.75, bh = 1.55, gap = 0.3;
  const bw = (W - 2 * M - 2 * gap) / 3;
  const boxes = [
    { h: "After the project, RNP's team can", t: "[Capability targets by role — what they can run, adapt and evaluate on their own]", fill: TEAL_LT, c: TEAL },
    { h: "Not covered by this transfer", t: "[Boundaries: e.g., work packages outside iFLYTEK's SOW; national software-stack R&D — confirm wording]", fill: PANEL_BG, c: NAVY },
    { h: "Language, materials & records", t: "[Working language and interpretation]  ·  [Materials handed over]  ·  [Practice records and technical outputs retained by RNP]", fill: PANEL_BG, c: NAVY },
  ];
  boxes.forEach((b, i) => {
    const bx = M + i * (bw + gap);
    flat(s, bx, by, bw, bh, b.fill, b.fill === TEAL_LT ? TEAL_LT : PANEL_LN);
    label(s, bx + 0.2, by + 0.14, bw - 0.4, b.h, b.c, 11.5);
    ph(s, bx + 0.2, by + 0.5, bw - 0.4, bh - 0.6, b.t, 10);
  });

  body(s, M, 6.42, W - 2 * M, 0.4,
    "Development, adaptation, testing and optimisation support tied to iFLYTEK deliverables is listed against concrete tasks. Time, seats and technical guidance effort follow the SOW.", 10, GRAY_L);

  s.addNotes(
    "PURPOSE: Answer 'once RNP's people join the project, how does iFLYTEK help them master the work, and what proves they did?'. " +
    "Knowledge transfer is tied to actual delivery: tasks, guidance, practice, capability result. Course names and hours are only part of it.\n\n" +
    "HEADLINE: this layer is already in the project quotation (internal meeting 4 Sep: 'already quoted, included in the project, free on the training side') - " +
    "verify against the SOW/quotation before presenting; do not extend it to 'all capability building is included'.\n\n" +
    "BOUNDARY: state plainly what the team will and will not be able to do. This page does not replace the whole national software-stack R&D support in the TR, and iFLYTEK does not by default cover all FP1-FP9.\n\n" +
    "FORMAT: keep the table compatible with the Huawei layer-1 course tables (audience, prerequisites, seats, content, outcome) so the combined talent deck reads as one programme.\n\n" +
    "SOURCE: internal meeting 4 Sep 23:13-24:16, 26:07, 44:41; Mota call ('to be with you and to learn')."
  );
}

// =====================================================================
// 9. OPTIONAL ADVANCED PROGRAMME
// =====================================================================
{
  const s = newSlide();
  title(s, "Optional: advanced programme for core talents");
  chip(s, M, 1.04, 3.0, "OPTIONAL  ·  SEPARATELY QUOTED", AMBER);
  s.addText("Adds to the in-project transfer for a small group with the right foundation. It does not replace anything already agreed.", {
    x: M + 3.2, y: 1.0, w: W - 2 * M - 3.2, h: 0.38, margin: 0, isTextBox: true, valign: "middle",
    fontFace: F, fontSize: 12, color: GRAY,
  });

  const cells = [
    { h: "Who & selection criteria", t: "[Profile, foundation required, how candidates are selected]" },
    { h: "Technical focus", t: "[Directions to go deeper: e.g., model adaptation, distillation, evaluation]" },
    { h: "Real task or research topic", t: "[The concrete task the group works on]" },
    { h: "Mentors, format & resources", t: "[Mentors, on-site / remote, environment and compute needed]" },
    { h: "Outcomes & how verified", t: "[Capability and tangible result; evidence used to check it]" },
    { h: "Size · duration · place · price", t: "[n people]  ·  [weeks]  ·  [location]  ·  [price and conditions]" },
  ];
  const gap = 0.3;
  const cw = (W - 2 * M - 2 * gap) / 3;
  const ch = 2.05;
  cells.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const cx = M + col * (cw + gap);
    const cy = 1.65 + row * (ch + 0.3);
    card(s, cx, cy, cw, ch, WHITE);
    badge(s, cx + 0.22, cy + 0.22, 0.34, String(i + 1), i === 5 ? AMBER : NAVY);
    label(s, cx + 0.7, cy + 0.2, cw - 0.9, c.h, NAVY, 12);
    ph(s, cx + 0.22, cy + 0.72, cw - 0.44, ch - 0.9, c.t, 10.5);
  });

  body(s, M, 6.4, W - 2 * M, 0.4,
    "A fee applies only to content and effort beyond the agreed scope. Knowledge transfer already agreed is not re-listed here because it is technically demanding.", 10, GRAY_L);

  s.addNotes(
    "PURPOSE: Beyond the in-project transfer, answer 'if RNP wants a few well-prepared people to go much deeper, what is added and what does it cost?'. " +
    "Corresponds to the optional layer discussed internally (a small, priced, tailor-made class for core talents who can later stand as visible results of the project).\n\n" +
    "MUST ANSWER: who qualifies; technical directions; the real task; mentors and format; outcomes and verification; size, duration, location, price, conditions.\n\n" +
    "GUARDRAIL: show clearly what is added relative to page 8 so RNP can decide whether to buy. Agreed transfer or R&D support cannot be moved into this paid item.\n\n" +
    "SOURCE: internal meeting 4 Sep 09:18-09:50, 24:59-26:43, 44:59."
  );
}

// =====================================================================
// 10. SCOPE AT A GLANCE & CONFIRMATIONS
// =====================================================================
{
  const s = newSlide();
  title(s, "Scope at a glance and what we confirm today",
    "Every item above mapped to included / optional / not in scope, then the decisions and inputs that let both sides move to planning.");

  // three scope columns
  const gap = 0.3;
  const cw = (W - 2 * M - 2 * gap) / 3;
  const cy = 1.6, ch = 2.0;
  const cols = [
    { h: "Included in the current proposal", c: TEAL, fill: TEAL_LT, t: "[LLM, APP, DLM deliverables]\n[In-project knowledge transfer]\n[Handover package: models, app, documents, licences, support period]" },
    { h: "Optional  ·  separately quoted", c: AMBER, fill: AMBER_LT, t: "[Advanced programme for core talents]\n[Other optional services]" },
    { h: "Not in scope  ·  third party", c: GRAY, fill: PANEL_BG, t: "[Work packages outside iFLYTEK's SOW]\n[Third-party licences and services]\n[Items RNP provides: data, environment, interfaces]" },
  ];
  cols.forEach((c, i) => {
    const cx = M + i * (cw + gap);
    flat(s, cx, cy, cw, ch, c.fill, c.fill === PANEL_BG ? PANEL_LN : c.fill);
    label(s, cx + 0.2, cy + 0.14, cw - 0.4, c.h, c.c, 12);
    ph(s, cx + 0.2, cy + 0.52, cw - 0.4, ch - 0.65, c.t, 10.5);
  });
  body(s, M, 3.68, W - 2 * M, 0.3, "Each item references the governing document: technical plan, SOW, quotation section and version.", 9.5, GRAY_L);

  // confirmation table
  const items = [
    "LLM, APP and DLM objectives and scope",
    "RNP team, resources and preparation",
    "Joint implementation plan",
    "Knowledge transfer & capability targets",
    "Optional advanced programme: refine?",
    "Handover, licences & support",
    "Scope and fee attribution",
  ];
  const rows = [[th("Item to confirm"), th("Status"), th("Conclusion / open question"), th("Owner"), th("Input needed"), th("Due")]];
  items.forEach((it, i) => {
    rows.push([tb(`${i + 1}.  ${it}`, { fontSize: 10 }), td("[aligned / open]"), td("[ ]"), td("[ ]"), td("[ ]"), td("[ ]")]);
  });
  s.addTable(rows, Object.assign(tableBase(), {
    x: M, y: 4.02, w: W - 2 * M,
    colW: [3.6, 1.8, 3.0, 1.3, 1.83, 0.8],
    rowH: [0.34, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33],
    fontSize: 9.5,
  }));

  s.addNotes(
    "PURPOSE: Make the meeting produce movement: what is aligned, what is still open, who supplies what by when. Also the second answer to RNP's need to write its own plan.\n\n" +
    "SCOPE BLOCK: content and value first (pages 2-9), then scope and fees here so RNP can see what each fee corresponds to and spot duplication or gaps. " +
    "Mark each item as included / separately quoted / optional / to confirm after checking the SOW, quotation and change records. " +
    "'Knowledge transfer is already in the quotation' must be verified in the documents and must not become 'everything is included' or 'the customer has accepted'.\n\n" +
    "HANDOVER (merged here): models, application, configs, scripts, interface documents, operating rights, licence conditions, RNP's tasks after handover, completion criteria, support and contact mechanism. Do not assume full source code or unlimited support.\n\n" +
    "CONFIRMATIONS: prepare the list before the meeting, record decisions during, turn into both sides' actions after. Distinguish confirmed / proposed / to discuss - appearance in the deck is not a commitment.\n\n" +
    "SOURCE: internal meeting 4 Sep 32:22-33:53 (this meeting effectively locks the SOW; send material in advance; aim for closure on the technical part)."
  );
}

pres.writeFile({ fileName: "RNP-iFLYTEK-Framework-EN.pptx" }).then((f) => console.log("wrote", f));
