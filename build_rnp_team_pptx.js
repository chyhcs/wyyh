// Iara's request (WhatsApp, 13 Sep 2026): "the skills we need on our side, the number of people"
//   -> 4-slide English deck in the style of the 9 Sep exchange deck (Calibri, red titles, light-blue boxes).
//      Content limited to her two questions; pre-filled from the 9 Sep deck (p.3, p.13, A3, A7, A10).
//      Yellow TBC tag on slide 2 + Chinese speaker notes on slide 1 are internal: remove / export PDF before sending.
//   node build_rnp_team_pptx.js  ->  RNP-Team-for-Technology-Transfer-Skills-and-Headcount.pptx
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "RNP Team for the Technology Transfer: Skills and Headcount";

const FONT = "Calibri";
const RED = "C00000", NAVY = "1F4E79", BLUE = "0070C0", INK = "1D1D1A", BODY = "333333", GRAY = "666666";
const LIGHTBLUE = "E8F0FA", LIGHTGRAY = "F2F2F2", LINE = "D9D9D9";

function title(slide, text) {
  slide.addText(text, { x: 0.52, y: 0.16, w: 11.42, h: 0.53, fontFace: FONT, fontSize: 18, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });
}
function intro(slide, text) {
  slide.addText(text, { x: 0.5, y: 0.82, w: 12.33, h: 0.5, fontFace: FONT, fontSize: 12.5, color: INK, margin: 0, valign: "top", isTextBox: true });
}
function meansBox(slide, text, y = 6.18) {
  slide.addShape(pres.ShapeType.roundRect, { x: 0.5, y, w: 12.33, h: 0.66, fill: { color: "FFFFFF" }, line: { color: NAVY, width: 0.75 }, rectRadius: 0.06 });
  slide.addText([
    { text: "What this means for RNP   ", options: { bold: true, color: NAVY } },
    { text, options: { bold: false, color: INK } },
  ], { x: 0.75, y: y + 0.04, w: 11.83, h: 0.58, fontFace: FONT, fontSize: 11, margin: 0, valign: "middle", isTextBox: true });
}
const BORDER = { type: "solid", pt: 0.5, color: LINE };
const hdr = t => ({ text: t, options: { bold: true, color: INK, fill: { color: LIGHTGRAY }, fontSize: 10 } });
const teamRow = (t, n) => [
  { text: t, options: { bold: true, color: NAVY, fill: { color: LIGHTBLUE }, fontSize: 10 } },
  { text: n, options: { bold: true, color: BLUE, fill: { color: LIGHTBLUE }, fontSize: 10, align: "center" } },
  { text: "", options: { fill: { color: LIGHTBLUE } } },
];
const roleRow = (r, n, s) => [
  { text: r, options: { bold: true, color: INK } },
  { text: n, options: { color: BLUE, bold: true, align: "center" } },
  { text: s, options: { color: BODY } },
];

// =====================================================================================
// Slide 1 · Title
// =====================================================================================
{
  const s = pres.addSlide();
  s.addText("RNP × iFLYTEK · Follow-up to the Technical Exchange of 9 September", { x: 0.8, y: 1.5, w: 8.6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: GRAY, margin: 0, isTextBox: true });
  s.addText("RNP Team for the Technology Transfer: Skills and Headcount", { x: 0.8, y: 1.98, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 27, bold: true, color: RED, margin: 0, valign: "top", isTextBox: true });
  s.addText("Which teams RNP needs on its side, what each must be able to do, and how many people", { x: 0.8, y: 3.25, w: 8.6, h: 0.6, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
  s.addText([
    { text: "Prepared for RNP at Iara's request of 13 September 2026", options: { breakLine: true } },
    { text: "Draws on the presentation of 9 September 2026 and its appendices A3, A7 and A10", options: { breakLine: true } },
    { text: "iFLYTEK · Draft v0.1 · 14 September 2026" },
  ], { x: 0.8, y: 4.1, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 13, color: BODY, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 4 });

  // Right: her request, quoted, as the scope of the deck
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 1.75, w: 2.9, h: 2.0, fill: { color: LIGHTGRAY }, line: { color: LIGHTGRAY }, rectRadius: 0.06 });
  s.addText("Your request, 13 September", { x: 10.15, y: 1.85, w: 2.45, h: 0.35, fontFace: FONT, fontSize: 11.5, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addText("“We are organizing RNP's side to receive the technology transfer and prepare to work together with you. Any important information about the skills we need on our side, the number of people, will be welcome.”", { x: 10.15, y: 2.22, w: 2.45, h: 1.45, fontFace: FONT, fontSize: 10, italic: true, color: BODY, margin: 0, valign: "top", isTextBox: true });
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 3.95, w: 2.9, h: 0.72, fill: { color: LIGHTBLUE }, line: { color: LIGHTBLUE }, rectRadius: 0.06 });
  s.addText("1  The skills we need on our side  →  slides 3–4", { x: 10.15, y: 3.95, w: 2.6, h: 0.72, fontFace: FONT, fontSize: 10.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 4.77, w: 2.9, h: 0.72, fill: { color: RED }, line: { color: RED }, rectRadius: 0.06 });
  s.addText("2  The number of people  →  slide 2", { x: 10.15, y: 4.77, w: 2.6, h: 0.72, fontFace: FONT, fontSize: 10.5, bold: true, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true });

  s.addNotes(
    "【内部说明，发送前处理】\n" +
    "1. 本稿只回答 Iara 9 月 13 日 WhatsApp 的两问：技能、人数。内容按 9 月 9 日英文 PPT 填（第 3、13 页与附录 A3、A7、A10）。\n" +
    "2. 人数口径要二选一：PPT 第 13 页写三批约 30 / 20 / 30 人，附录 A7 按岗位合计 53 / 25 / 33 人、总计 111。本稿按 A7 明细填。第 2 页右上角黄色标签是提醒，定稿后删除。\n" +
    "3. 发送前导出 PDF；备注不会进入 PDF。"
  );
}

// =====================================================================================
// Slide 2 · The number of people (team at a glance)
// =====================================================================================
{
  const s = pres.addSlide();
  title(s, "The number of people: eight RNP teams, 111 at peak");
  intro(s, "Each team is defined by what it does with iFLYTEK during the project; headcounts are peaks per function. Source: presentation of 9 September, p.13 and Appendix A7.");
  // internal TBC tag (delete before sending)
  s.addShape(pres.ShapeType.roundRect, { x: 8.55, y: 0.24, w: 4.3, h: 0.48, fill: { color: "FFFF00" }, line: { color: "FFC000" }, rectRadius: 0.05 });
  s.addText("TBC before sending: p.13 says about 30 / 20 / 30 per intake; A7 detail sums to 53 / 25 / 33. Confirm one set.", { x: 8.62, y: 0.24, w: 4.18, h: 0.48, fontFace: FONT, fontSize: 8.5, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });

  const rows = [
    [hdr("Team"), hdr("What the team does with iFLYTEK"), hdr("Core skills"), { ...hdr("People"), options: { ...hdr("People").options, align: "center" } }],
    roleRow4("1  Project and product", "Requirements, schedule, resources and delivery; coordinates Huawei compute and environment; organises acceptance; defines the National AI Application with iFLYTEK", "Project management; product definition; data-centre coordination", "9"),
    roleRow4("2  Corpus and data", "Organises data acquisition, licensing and full-volume production of the Portuguese corpus; runs the corpus pilot batch with iFLYTEK's specifications and sampling checks", "Data sourcing and licensing; data processing on the platform; annotation and quality control", "32"),
    roleRow4("3  Portuguese language experts", "Sets Portuguese quality, annotation rules, safety labels and speech standards; checks samples", "Brazilian Portuguese linguistics; annotation standards; speech data", "6"),
    roleRow4("4  Compliance and review", "Content, safety and legal review; secures corpus licences; runs the Content Safety Review Platform after acceptance", "Content safety policy; data and IP law; review operations", "6"),
    roleRow4("5  Business and domain experts", "Confirms the business scenarios for the three Domain Large Models; assesses domain results", "Domain knowledge for the chosen scenarios; requirements; evaluation of model output", "7"),
    roleRow4("6  Model and test engineering", "Joins national model training and evaluation in the Model Training Center; then trains the DLMs with iFLYTEK advising", "LLM training and evaluation on the platform; test and validation", "18"),
    roleRow4("7  Application and operations", "Takes over the National AI Application: architecture, integration, user approval, content and daily operations", "Software development; application operations; architecture", "27"),
    roleRow4("8  Platform O&M", "Runs the platform after acceptance: monitoring, deployment, incidents, escalation; assisted O&M with iFLYTEK for 12 months", "Cluster and platform operations; incident and change management", "6"),
    [
      { text: "Total, peak per function", options: { bold: true, color: NAVY, fill: { color: LIGHTBLUE } } },
      { text: "", options: { fill: { color: LIGHTBLUE } } },
      { text: "", options: { fill: { color: LIGHTBLUE } } },
      { text: "111", options: { bold: true, color: BLUE, fill: { color: LIGHTBLUE }, align: "center", fontSize: 11 } },
    ],
  ];
  s.addTable(rows, { x: 0.5, y: 1.4, w: 12.33, colW: [2.25, 5.2, 3.88, 1.0], fontFace: FONT, fontSize: 9.5, color: BODY, border: BORDER, valign: "middle", margin: [0.04, 0.08, 0.04, 0.08], rowH: 0.42, autoPage: false });
  meansBox(s, "Skills first, numbers second. A person can cover more than one activity, so these are not people on site at the same time; final numbers follow the split of work and the schedule agreed before implementation.", 6.42);
}
function roleRow4(t, does, skills, n) {
  return [
    { text: t, options: { bold: true, color: INK } },
    { text: does, options: { color: BODY } },
    { text: skills, options: { color: BODY } },
    { text: n, options: { bold: true, color: BLUE, align: "center", fontSize: 11 } },
  ];
}

// =====================================================================================
// Slide 3 · Skills by role (1 of 2)
// =====================================================================================
{
  const s = pres.addSlide();
  title(s, "The skills we need on our side (1 of 2): project, corpus, language and compliance teams");
  intro(s, "What each role must be able to do, with the number of people per role. Role composition from Appendix A7; activities from Appendices A10 and A3.");
  const rows = [
    [hdr("Role"), { ...hdr("People"), options: { ...hdr("People").options, align: "center" } }, hdr("Must be able to")],
    teamRow("Team 1  Project and product", "9"),
    roleRow("Project manager", "2", "Run the five-step schedule with iFLYTEK and Huawei; manage resources, risks and delivery; organise component and final acceptance"),
    roleRow("Product manager", "4", "Define the National AI Application with iFLYTEK: functions, user scope, content rules, Brazilian content and roadmap"),
    roleRow("PMO / data-centre expert", "3", "Coordinate Huawei compute, facilities, network and base environment; track T0 and the compute batches"),
    teamRow("Team 2  Corpus and data", "32"),
    roleRow("Data resource engineer", "6", "Identify data sources, secure licences and permitted use, organise acquisition; keep project data in the RNP-designated Brazilian environment"),
    roleRow("Data quality engineer", "6", "Apply the agreed counting, sampling, quality-threshold, batching and re-check rules; run the corpus pilot batch; operate the Corpus Construction Center (ingest, clean, de-duplicate, quality metrics)"),
    roleRow("Annotator", "20", "Annotate to the agreed standards: 200,000 supervised samples, 20,000 safety and values samples, transcription of 10,000 hours of speech"),
    teamRow("Team 3  Portuguese language experts", "6"),
    roleRow("Language expert", "6", "Set Brazilian Portuguese quality standards, annotation rules, safety labels and speech (ASR/TTS) standards; analyse samples and confirm rules with iFLYTEK"),
    teamRow("Team 4  Compliance and review", "6"),
    roleRow("Review expert", "2", "Define review policies, sensitive lexicons and risk samples on the Content Safety Review Platform; review model and application content"),
    roleRow("Legal expert", "2", "Clear data licences and permitted use; content, safety and legal compliance of corpus, models and application"),
    roleRow("Review lead", "2", "Own the review process and its records; escalate and approve"),
  ];
  s.addTable(rows, { x: 0.5, y: 1.4, w: 12.33, colW: [2.6, 0.9, 8.83], fontFace: FONT, fontSize: 9.5, color: BODY, border: BORDER, valign: "middle", margin: [0.03, 0.08, 0.03, 0.08], rowH: 0.33, autoPage: false });
}

// =====================================================================================
// Slide 4 · Skills by role (2 of 2)
// =====================================================================================
{
  const s = pres.addSlide();
  title(s, "The skills we need on our side (2 of 2): domain, model, application and O&M teams");
  intro(s, "What each role must be able to do, with the number of people per role. Role composition from Appendix A7; activities from Appendix A10.");
  const rows = [
    [hdr("Role"), { ...hdr("People"), options: { ...hdr("People").options, align: "center" } }, hdr("Must be able to")],
    teamRow("Team 5  Business and domain experts", "7"),
    roleRow("Business requirements", "3", "Define the business scenario for each of the three DLM services; confirm inputs and acceptance of the service"),
    roleRow("Domain requirements", "2", "Specify domain data and expected results; read evaluation reports and set content boundaries"),
    roleRow("Domain expert", "2", "Assess model results in the domain; take part in evaluation exercises with iFLYTEK"),
    teamRow("Team 6  Model and test engineering", "18"),
    roleRow("Algorithm engineer", "12", "Prepare datasets, configure and submit training jobs, read logs and evaluations, release models in the Model Training Center; train the DLMs on the platform with iFLYTEK advising"),
    roleRow("Test engineer", "6", "Run model evaluation and technical validation against the agreed evaluation rules; document results for acceptance"),
    teamRow("Team 7  Application and operations", "27"),
    roleRow("Software developer", "15", "Integrate the National AI Application (mobile, desktop, web) with RNP systems; develop Brazilian content and features on the delivered baseline; call the platform APIs"),
    roleRow("Operations engineer", "10", "Operate the application: configuration, user approval and permissions, API credentials and rate limits, usage analytics, feedback"),
    roleRow("Architect", "2", "Own application architecture and integration with RNP systems and the platform"),
    teamRow("Team 8  Platform O&M", "6"),
    roleRow("O&M engineer", "6", "Monitor the platform and cluster; handle incidents, changes and common faults; manage deployments and escalation; take over from assisted O&M and run the platform independently"),
  ];
  s.addTable(rows, { x: 0.5, y: 1.4, w: 12.33, colW: [2.6, 0.9, 8.83], fontFace: FONT, fontSize: 9.5, color: BODY, border: BORDER, valign: "middle", margin: [0.03, 0.08, 0.03, 0.08], rowH: 0.33, autoPage: false });
  meansBox(s, "These are the skills to look for when RNP assigns or recruits its people. Once RNP has a first team plan, iFLYTEK will review it against this and adjust it together with RNP.", 6.42);
}

const OUT = "RNP-Team-for-Technology-Transfer-Skills-and-Headcount.pptx";
pres.writeFile({ fileName: OUT }).then(f => console.log("wrote", f));
