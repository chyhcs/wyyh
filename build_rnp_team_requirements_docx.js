// Iara's request (WhatsApp, 13 Sep 2026): "the skills we need on our side, the number of people"
//   -> English Word note for RNP: teams, skills, headcount, timing. Pre-filled from the 9 Sep deck
//      (p.3, p.12, p.13, A3, A6, A7, A10). Anything the deck does not say is marked [TBC: ...] in yellow.
//   node build_rnp_team_requirements_docx.js  ->  RNP-Team-for-Technology-Transfer-Skills-and-Headcount.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
        ShadingType, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign } = require("docx");

const F = "Arial", FZ = "微软雅黑";
const NAVY = "1F3A5F", BLUE = "0057A8", GRAY = "6B6B6B", BLACK = "1D1D1A", RED = "C00000";
const D = [];

// ---- text helpers: **bold**, [TBC: ...] -> yellow highlight ----
function runs(t, o = {}) {
  const out = [];
  String(t).split(/(\*\*.+?\*\*|\[TBC[^\]]*\])/).forEach(seg => {
    if (!seg) return;
    if (seg.startsWith("**")) out.push(new TextRun({ text: seg.slice(2, -2), font: o.font || F, size: o.size || 20, bold: true, color: o.color || BLACK }));
    else if (seg.startsWith("[TBC")) out.push(new TextRun({ text: seg, font: o.font || F, size: o.size || 20, bold: true, color: RED, highlight: "yellow" }));
    else out.push(new TextRun({ text: seg, font: o.font || F, size: o.size || 20, bold: !!o.bold, color: o.color || BLACK, italics: !!o.italics }));
  });
  return out;
}
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || F, size: o.size || 20, bold: !!o.bold, color: o.color || BLACK, italics: !!o.italics });
const P = (t, o = {}) => D.push(new Paragraph({ children: runs(t, o), spacing: { after: o.after == null ? 120 : o.after, line: 300 }, ...(o.indent ? { indent: { left: o.indent } } : {}) }));
const H1 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 }, keepNext: true, children: [R(t, { size: 26, bold: true, color: NAVY })] }));
const H2 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 80 }, keepNext: true, children: [R(t, { size: 22, bold: true, color: BLUE })] }));
const B = t => D.push(new Paragraph({ children: runs(t), numbering: { reference: "b", level: 0 }, spacing: { after: 60, line: 300 } }));
const NOTE = t => D.push(new Paragraph({ children: runs(t, { size: 18, color: GRAY, italics: true }), spacing: { before: 60, after: 160, line: 280 } }));

// ---- tables ----
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
function cell(text, w, o = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map(line => new Paragraph({
    children: runs(line, { size: o.size || 18, bold: o.bold, color: o.color }),
    spacing: { after: 40, line: 260 },
    alignment: o.align || AlignmentType.LEFT,
  }));
  return new TableCell({ children: paras, width: { size: w, type: WidthType.DXA }, borders: BORDERS,
    margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: VerticalAlign.TOP,
    ...(o.fill ? { shading: { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } } : {}) });
}
function TBL(widths, header, rows, o = {}) {
  const hdr = new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: "DCE6F1", size: 18 })) });
  const body = rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { fill: (o.fillCol && i === o.fillCol.idx) ? o.fillCol.fill : (r.__fill || undefined), size: 18 })) }));
  D.push(new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: [hdr, ...body] }));
  D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
}
function BOX(lines, fill, font) {
  lines.forEach((t, i) => D.push(new Paragraph({
    children: runs(t, { size: 19, font: font || F }),
    spacing: { after: i === lines.length - 1 ? 200 : 60, line: 300 },
    indent: { left: 200, right: 200 },
    shading: { type: ShadingType.CLEAR, fill, color: "auto" },
  })));
}

// =====================================================================================
// 0. Internal fill-in guide (Chinese, delete before sending)
// =====================================================================================
BOX([
  "**【内部填写说明，发送前整段删除】**",
  "本稿只回答 Iara 9 月 13 日 WhatsApp 的两问：RNP 侧需要什么技能的人、需要多少人。内容先按 9 月 9 日英文 PPT 填充（第 3、12、13 页与附录 A3、A6、A7、A10）。PPT 没有的地方用黄底 [TBC: …] 标出，需要你手工填。",
  "1. **人数口径要二选一**：PPT 第 13 页写三批约 30 / 20 / 30 人，附录 A7 按岗位合计为 53 / 25 / 33 人。本稿正文按 A7 的岗位明细填，总览表下有一条黄底提示，定稿前统一成一套。",
  "2. **各岗位的背景要求**：PPT 只写了各角色参与什么工作，没有写学历、经验年限等要求，全部标 TBC。",
  "3. **移交后常设人数**：PPT 没有数字，会上口头说过 15 到 20 人，是否写入由你定。",
  "4. **到位提前量与 T0 日期**：PPT 只说 T0 为首批算力就绪、逐月日期实施前商定。",
  "5. **讯飞对口团队**：PPT 没有列，第 4 节留空待填。",
  "6. **可短期外借的岗位**：会上口头说过语言专家和领域专家只需半年到九个月，可请高校人员短期参与，是否写入由你定。",
], "FFF3CD", FZ);

// =====================================================================================
// Title
// =====================================================================================
D.push(new Paragraph({ spacing: { after: 60 }, children: [R("RNP Team for the Technology Transfer", { size: 34, bold: true, color: NAVY })] }));
D.push(new Paragraph({ spacing: { after: 40 }, children: [R("Skills and headcount RNP needs on its side, by team and by project step", { size: 24, bold: true, color: BLUE })] }));
D.push(new Paragraph({ spacing: { after: 280 }, children: [R("iFLYTEK input to RNP's team planning  ·  Draft v0.1  ·  14 September 2026  ·  Follows the presentation of 9 September 2026", { size: 18, color: GRAY })] }));

P("Further to your message of 13 September, this note sets out what RNP needs on its side to receive the technology transfer and to work with iFLYTEK during the project: **which teams, what skills, how many people, and from when**. It is drawn from the presentation of 9 September and its appendices, so RNP can plan from one consistent picture.");
P("All headcounts are indicative. They are iFLYTEK's early estimate of the peak per function; a person can cover more than one activity, so they are not people on site at the same time. Final numbers follow the SOW split of work and the month-by-month schedule agreed before implementation.");

// =====================================================================================
// 1. Summary
// =====================================================================================
H1("1.  The RNP team at a glance");
P("Eight teams. The first four start with Step 1 (platform and corpus), the next two with Step 2 (model training and evaluation), the last two with Step 3 (application and DLM). Source: deck p.13 and Appendix A7.");

const W1 = [1750, 3250, 2650, 900, 1050];
TBL(W1, ["Team", "What the team does with iFLYTEK", "Core skills", "People", "Needed from"], [
  ["**1  Project and product**", "Requirements, schedule, resources and delivery; coordinates Huawei compute and environment; organises acceptance; defines the National AI Application with iFLYTEK (functions, user scope, content rules)", "Project management; product definition; data-centre coordination", "9", "Step 1"],
  ["**2  Corpus and data**", "Organises data acquisition, licensing and full-volume production of the Portuguese corpus; runs the corpus pilot batch with iFLYTEK's specifications and sampling checks", "Data sourcing and licensing; data processing on the platform; annotation and quality control", "32", "Step 1"],
  ["**3  Portuguese language experts**", "Sets Portuguese quality, annotation rules, safety labels and speech standards; checks samples", "Brazilian Portuguese linguistics; annotation standards; speech data", "6", "Step 1"],
  ["**4  Compliance and review**", "Content, safety and legal review; secures corpus licences; runs the Content Safety Review Platform after acceptance", "Content safety policy; data and IP law; review operations", "6", "Step 1"],
  ["**5  Business and domain experts**", "Confirms the business scenarios for the three Domain Large Models; assesses domain results", "Domain knowledge for the chosen scenarios; requirements; evaluation of model output", "7", "Step 2"],
  ["**6  Model and test engineering**", "Joins national model training and evaluation in the Model Training Center; then trains the DLMs with iFLYTEK advising", "LLM training and evaluation on the platform; test and validation", "18", "Step 2"],
  ["**7  Application and operations**", "Takes over the National AI Application: architecture, integration, user approval, content and daily operations", "Software development; application operations; architecture", "27", "Step 3"],
  ["**8  Platform O&M**", "Runs the platform after acceptance: monitoring, deployment, incidents, escalation; assisted O&M with iFLYTEK for 12 months", "Cluster and platform operations; incident and change management", "6", "Step 3"],
  ["**Total (peak per function)**", "", "", "**111**", ""],
]);
P("[TBC: The deck headline (p.13) gives about 30 / 20 / 30 people per intake; the A7 detail above sums to 53 / 25 / 33. Confirm one set of figures before sending.]", { size: 18 });
P("[TBC: Permanent team RNP keeps after handover (development vs operations): not stated in the deck.]", { size: 18 });

// =====================================================================================
// 2. Skills by role
// =====================================================================================
H1("2.  Skills by role");
P("Each team broken down by role, with what each role must be able to do. Role composition from Appendix A7; activities from Appendix A10 and A3. Background requirements are open for RNP's HR to align with its own job profiles.");

const W2 = [2000, 800, 4100, 2700];
const HDR2 = ["Role", "People", "Must be able to", "Background"];

H2("Team 1  Project and product  (9)");
TBL(W2, HDR2, [
  ["Project manager", "2", "Run the five-step schedule with iFLYTEK and Huawei; manage resources, risks and delivery; organise component and final acceptance", "[TBC: e.g. large IT/infrastructure programme delivery]"],
  ["Product manager", "4", "Define the National AI Application with iFLYTEK: functions, user scope, content rules, Brazilian content and roadmap", "[TBC]"],
  ["PMO / data-centre expert", "3", "Coordinate Huawei compute, facilities, network and base environment; track T0 and compute batches", "[TBC]"],
]);

H2("Team 2  Corpus and data  (32)");
TBL(W2, HDR2, [
  ["Data resource engineer", "6", "Identify data sources, secure licences and permitted use, organise acquisition; store project data in the RNP-designated Brazilian environment", "[TBC]"],
  ["Data quality engineer", "6", "Apply the agreed counting, sampling, quality-threshold, batching and re-check rules; run the corpus pilot batch; operate the Corpus Construction Center (ingest, clean, de-duplicate, quality metrics)", "[TBC]"],
  ["Annotator", "20", "Annotate to the agreed standards: 200,000 supervised samples, 20,000 safety and values samples, transcription of 10,000 hours of speech", "[TBC]"],
]);

H2("Team 3  Portuguese language experts  (6)");
TBL(W2, HDR2, [
  ["Language expert", "6", "Set Brazilian Portuguese quality standards, annotation rules, safety labels and speech (ASR/TTS) standards; analyse samples and confirm rules with iFLYTEK", "[TBC: e.g. linguistics, corpus linguistics, phonetics]"],
]);

H2("Team 4  Compliance and review  (6)");
TBL(W2, HDR2, [
  ["Review expert", "2", "Define review policies, sensitive lexicons and risk samples on the Content Safety Review Platform; review model and application content", "[TBC]"],
  ["Legal expert", "2", "Clear data licences and permitted use; content, safety and legal compliance of corpus, models and application", "[TBC]"],
  ["Review lead", "2", "Own the review process and its records; escalate and approve", "[TBC]"],
]);

H2("Team 5  Business and domain experts  (7)");
TBL(W2, HDR2, [
  ["Business requirements", "3", "Define the business scenario for each of the three DLM services; confirm inputs and acceptance of the service", "[TBC]"],
  ["Domain requirements", "2", "Specify domain data and expected results; read evaluation reports and set content boundaries", "[TBC]"],
  ["Domain expert", "2", "Assess model results in the domain; take part in evaluation exercises with iFLYTEK", "[TBC]"],
]);

H2("Team 6  Model and test engineering  (18)");
TBL(W2, HDR2, [
  ["Algorithm engineer", "12", "Prepare datasets, configure and submit training jobs, read logs and evaluations, release models in the Model Training Center; train the DLMs on the platform with iFLYTEK advising", "[TBC: e.g. ML/LLM engineering, Python, distributed training]"],
  ["Test engineer", "6", "Run model evaluation and technical validation against the agreed evaluation rules; document results for acceptance", "[TBC]"],
]);

H2("Team 7  Application and operations  (27)");
TBL(W2, HDR2, [
  ["Software developer", "15", "Integrate the National AI Application (mobile, desktop, web) with RNP systems; develop Brazilian content and features on the delivered baseline; call the platform APIs", "[TBC: e.g. mobile/web/back-end development]"],
  ["Operations engineer", "10", "Operate the application: configuration, user approval and permissions, API credentials and rate limits, usage analytics, feedback", "[TBC]"],
  ["Architect", "2", "Own application architecture and integration with RNP systems and the platform", "[TBC]"],
]);

H2("Team 8  Platform O&M  (6)");
TBL(W2, HDR2, [
  ["O&M engineer", "6", "Monitor the platform and cluster, handle incidents, changes and common faults, manage deployments and escalation; take over from assisted O&M and run the platform independently", "[TBC: e.g. Linux, containers, cluster operations]"],
]);

// =====================================================================================
// 3. When each team is needed
// =====================================================================================
H1("3.  When each team is needed");
P("Timing follows the five implementation steps (deck p.12). T0 is the first compute readiness; month-by-month dates are agreed before implementation. Each team should be in place before its step starts.");

const W3 = [2300, 2900, 3300, 1100];
TBL(W3, ["Step", "Completion markers", "Teams that must be in place", "People"], [
  ["**1  Platform and corpus**", "Platform accepted; corpus pilot batch passes", "Project and product; Corpus and data; Portuguese language experts; Compliance and review", "53"],
  ["**2  Model training and evaluation**", "Text model passes; speech models pass", "Step 1 teams, plus Business and domain experts; Model and test engineering", "+ 25"],
  ["**3  Application and DLM**", "App Stage 1 delivered; DLM forms signed", "Steps 1–2 teams, plus Application and operations; Platform O&M", "+ 33"],
  ["**4  Acceptance**", "Component acceptance reports signed", "All teams; acceptance organised by Project and product", "—"],
  ["**5  Handover and support**", "Handover signed; assisted O&M begins (12 months on site, then 60 months remote)", "Application and operations; Platform O&M; Model engineering for DLMs. [TBC: permanent headcount after handover]", "[TBC]"],
]);
P("[TBC: Lead time: how many months before Step 1 the first teams should be recruited and in place; and the planned T0 date.]", { size: 18 });

// =====================================================================================
// 4. Working with iFLYTEK
// =====================================================================================
H1("4.  Working with iFLYTEK");
P("Each RNP team works with one iFLYTEK team throughout its steps, so that the transfer happens on the actual project work rather than in separate sessions.");
const W4 = [2600, 3400, 3600];
TBL(W4, ["RNP team", "iFLYTEK counterpart", "Where they work together"], [
  ["Project and product", "[TBC: iFLYTEK project management / product team]", "Schedule, acceptance, application definition"],
  ["Corpus and data; Portuguese language experts", "[TBC: iFLYTEK data / corpus team]", "Corpus specifications, pilot batch, sampling checks; Corpus Construction Center"],
  ["Compliance and review", "[TBC: iFLYTEK content safety team]", "Review policies, lexicons, risk samples"],
  ["Model and test engineering; Business and domain experts", "[TBC: iFLYTEK model training / DLM advisory team]", "National model training and evaluation; DLM training"],
  ["Application and operations", "[TBC: iFLYTEK application team]", "Application delivery, integration, operations"],
  ["Platform O&M", "[TBC: iFLYTEK O&M / support team]", "Assisted O&M, handover, remote support"],
]);

// =====================================================================================
// 5. Notes
// =====================================================================================
H1("5.  Notes");
B("Headcounts are peaks per function and can be reused across activities; they are not people on site at once.");
B("Under this SOW, RNP takes on data production, DLM training and daily operations, so the figures should be re-checked against the actual split of work once the schedule is fixed.");
B("[TBC: Which roles can be short-term or drawn from universities and partner institutions, and for how long.]");
B("iFLYTEK is glad to review RNP's team plan against this note once RNP has a first version.");

// =====================================================================================
const doc = new Document({
  creator: "iFLYTEK", title: "RNP Team for the Technology Transfer: Skills and Headcount",
  styles: { default: { document: { run: { font: F, size: 20, color: BLACK } } } },
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      R("RNP × iFLYTEK  ·  Team for the Technology Transfer  ·  Draft v0.1  ·  page ", { size: 16, color: GRAY }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 16, color: GRAY }) ] })] }) },
    children: D,
  }],
});
const OUT = "RNP-Team-for-Technology-Transfer-Skills-and-Headcount.docx";
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes"); });
