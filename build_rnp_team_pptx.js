// Iara's request (WhatsApp, 13 Sep 2026): "the skills we need on our side, the number of people"
//   -> 3-slide English deck in the style of the 9 Sep exchange deck:
//      1 cover · 2 overview (three intakes as on deck p.13, eight team cards, 111 at peak)
//      3 role profiles (one table: team / role / people / how the role works with iFLYTEK)
//      Pre-filled from the 9 Sep deck (p.13, A3, A7, A10). Yellow TBC tag on slide 2 + Chinese notes on slide 1
//      are internal: remove / export PDF before sending.
//   node build_rnp_team_pptx.js  ->  RNP-Team-for-Technology-Transfer-Skills-and-Headcount.pptx
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "RNP Team for the Technology Transfer: Skills and Headcount";

const FONT = "Calibri";
const RED = "C00000", NAVY = "1F4E79", BLUE = "0070C0", INK = "1D1D1A", BODY = "333333", GRAY = "666666";
const LIGHTBLUE = "E8F0FA", LIGHTGRAY = "F2F2F2", LINE = "D9D9D9";

function title(s, text, w = 8.0) {
  s.addText(text, { x: 0.52, y: 0.16, w, h: 0.5, fontFace: FONT, fontSize: 18, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });
}
function intro(s, text) {
  s.addText(text, { x: 0.5, y: 0.7, w: 12.33, h: 0.4, fontFace: FONT, fontSize: 10.5, color: INK, margin: 0, valign: "top", isTextBox: true });
}
function tbcTag(s) {
  s.addShape(pres.ShapeType.roundRect, { x: 8.7, y: 0.17, w: 4.15, h: 0.46, fill: { color: "FFFF00" }, line: { color: "FFC000" }, rectRadius: 0.05 });
  s.addText("TBC before sending: p.13 says about 30 / 20 / 30 per intake; A7 detail sums to 53 / 25 / 33. Confirm one set.", { x: 8.77, y: 0.17, w: 4.03, h: 0.46, fontFace: FONT, fontSize: 8, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });
}

// =====================================================================================
// Slide 1 · Cover
// =====================================================================================
{
  const s = pres.addSlide();
  s.addText("RNP × iFLYTEK · Follow-up to the Technical Exchange of 9 September", { x: 0.8, y: 1.5, w: 8.6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: GRAY, margin: 0, isTextBox: true });
  s.addText("RNP Team for the Technology Transfer: Skills and Headcount", { x: 0.8, y: 1.98, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 27, bold: true, color: RED, margin: 0, valign: "top", isTextBox: true });
  s.addText("The teams and roles RNP needs on its side, how each works with iFLYTEK, and how many people", { x: 0.8, y: 3.25, w: 8.6, h: 0.6, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
  s.addText([
    { text: "Prepared for RNP at Iara's request of 13 September 2026", options: { breakLine: true } },
    { text: "Draws on the presentation of 9 September 2026 and its appendices A3, A7 and A10", options: { breakLine: true } },
    { text: "iFLYTEK · Draft v0.1 · 14 September 2026" },
  ], { x: 0.8, y: 4.1, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 13, color: BODY, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 4 });

  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 1.75, w: 2.9, h: 2.0, fill: { color: LIGHTGRAY }, line: { color: LIGHTGRAY }, rectRadius: 0.06 });
  s.addText("Your request, 13 September", { x: 10.15, y: 1.85, w: 2.45, h: 0.35, fontFace: FONT, fontSize: 11.5, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addText("“We are organizing RNP's side to receive the technology transfer and prepare to work together with you. Any important information about the skills we need on our side, the number of people, will be welcome.”", { x: 10.15, y: 2.22, w: 2.45, h: 1.45, fontFace: FONT, fontSize: 10, italic: true, color: BODY, margin: 0, valign: "top", isTextBox: true });
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 3.95, w: 2.9, h: 0.72, fill: { color: LIGHTBLUE }, line: { color: LIGHTBLUE }, rectRadius: 0.06 });
  s.addText("Overview  ·  8 teams, 111 people at peak", { x: 10.15, y: 3.95, w: 2.6, h: 0.72, fontFace: FONT, fontSize: 10.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 4.77, w: 2.9, h: 0.72, fill: { color: RED }, line: { color: RED }, rectRadius: 0.06 });
  s.addText("Role profiles  ·  19 roles, how each works with iFLYTEK", { x: 10.15, y: 4.77, w: 2.6, h: 0.72, fontFace: FONT, fontSize: 10.5, bold: true, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true });

  s.addNotes(
    "【内部说明，发送前处理】\n" +
    "1. 本稿只回答 Iara 9 月 13 日 WhatsApp 的两问：技能、人数。第 2 页概览按 9 月 9 日 PPT 第 13 页的三批次布局，第 3 页岗位画像一张表。内容按 PPT 第 13 页与附录 A3、A7、A10 填。\n" +
    "2. 人数口径要二选一：PPT 第 13 页写三批约 30 / 20 / 30 人，附录 A7 按岗位合计 53 / 25 / 33 人、总计 111。本稿按 A7 明细填。第 2 页右上角黄色标签是提醒，定稿后删除。\n" +
    "3. 发送前导出 PDF；备注不会进入 PDF。"
  );
}

// =====================================================================================
// Slide 2 · Overview: three intakes, eight team cards
// =====================================================================================
{
  const s = pres.addSlide();
  title(s, "The RNP team at a glance: 8 teams in 3 intakes, 111 people at peak");
  intro(s, "Teams grouped as in the presentation of 9 September (p.13). Headcounts are peaks per function, reusable across activities, not people on site at once. Source: p.13 and Appendix A7.");
  tbcTag(s);

  const COLS = [
    { x: 0.5, name: "Intake 1 · Platform and corpus", n: 53, teams: [
      { name: "1  Project and product", n: 9, does: "Schedule, resources and acceptance; app definition with iFLYTEK", roles: "Project managers 2 · Product managers 4 · PMO / data-centre 3" },
      { name: "2  Corpus and data", n: 32, does: "Sources, licenses and produces the corpus to iFLYTEK's specs", roles: "Data resource 6 · Data quality 6 · Annotators 20" },
      { name: "3  Portuguese language experts", n: 6, does: "Sets Portuguese quality and annotation standards with iFLYTEK", roles: "Language experts 6" },
      { name: "4  Compliance and review", n: 6, does: "Content, safety and legal review; corpus licences", roles: "Review experts 2 · Legal experts 2 · Review leads 2" },
    ] },
    { x: 4.71, name: "Intake 2 · Model training and evaluation", n: 25, teams: [
      { name: "5  Business and domain experts", n: 7, does: "Confirms the scenarios for the three Domain Large Models and assesses domain results with iFLYTEK", roles: "Business requirements 3 · Domain requirements 2 · Domain experts 2" },
      { name: "6  Model and test engineering", n: 18, does: "Trains and evaluates the national model with iFLYTEK's engineers, then trains the DLMs with iFLYTEK advising", roles: "Algorithm engineers 12 · Test engineers 6" },
    ] },
    { x: 8.92, name: "Intake 3 · Application and DLM", n: 33, teams: [
      { name: "7  Application and operations", n: 27, does: "Takes over the National AI Application from iFLYTEK: architecture, integration, user approval, content and operations", roles: "Software developers 15 · Operations engineers 10 · Architects 2" },
      { name: "8  Platform O&M", n: 6, does: "Runs the platform with iFLYTEK during 12 months of assisted O&M, then independently", roles: "O&M engineers 6" },
    ] },
  ];
  const W = 3.91, Y0 = 1.25, PH = 4.85, CY = 1.92, CH_AREA = 4.08;
  COLS.forEach(c => {
    s.addShape(pres.ShapeType.roundRect, { x: c.x, y: Y0, w: W, h: PH, fill: { color: LIGHTBLUE }, line: { color: LIGHTBLUE }, rectRadius: 0.06 });
    s.addText(c.name, { x: c.x + 0.2, y: Y0 + 0.1, w: 2.75, h: 0.5, fontFace: FONT, fontSize: 11.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
    s.addText(`${c.n}`, { x: c.x + W - 1.05, y: Y0 + 0.08, w: 0.85, h: 0.5, fontFace: FONT, fontSize: 20, bold: true, color: BLUE, margin: 0, align: "right", valign: "middle", isTextBox: true });
    const k = c.teams.length, gap = 0.08, ch = (CH_AREA - gap * (k - 1)) / k;
    c.teams.forEach((t, i) => {
      const y = CY + i * (ch + gap), x = c.x + 0.16, w = W - 0.32;
      s.addShape(pres.ShapeType.roundRect, { x, y, w, h: ch, fill: { color: "FFFFFF" }, line: { color: LINE, width: 0.75 }, rectRadius: 0.05 });
      s.addText(t.name, { x: x + 0.14, y: y + 0.06, w: w - 0.9, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, color: INK, margin: 0, valign: "middle", isTextBox: true });
      s.addText(`${t.n}`, { x: x + w - 0.8, y: y + 0.04, w: 0.66, h: 0.34, fontFace: FONT, fontSize: 16, bold: true, color: BLUE, margin: 0, align: "right", valign: "middle", isTextBox: true });
      const tall = ch > 1.2;
      s.addText(t.does, { x: x + 0.14, y: y + 0.38, w: w - 0.28, h: tall ? 0.62 : 0.28, fontFace: FONT, fontSize: 9, color: BODY, margin: 0, valign: "top", isTextBox: true });
      s.addText(t.roles, { x: x + 0.14, y: y + (tall ? 0.98 : 0.68), w: w - 0.28, h: tall ? 0.5 : 0.26, fontFace: FONT, fontSize: 8.5, bold: true, color: GRAY, margin: 0, valign: "top", isTextBox: true });
    });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 6.28, w: 12.33, h: 0.55, fill: { color: "FFFFFF" }, line: { color: NAVY, width: 0.75 }, rectRadius: 0.06 });
  s.addText([
    { text: "Total, peak per function   ", options: { bold: true, color: NAVY } },
    { text: "111 people", options: { bold: true, color: BLUE, fontSize: 14 } },
    { text: "     across 8 teams and 19 roles; the roles are profiled on the next page.", options: { color: INK } },
  ], { x: 0.75, y: 6.3, w: 11.83, h: 0.5, fontFace: FONT, fontSize: 11, margin: 0, valign: "middle", isTextBox: true });
}

// =====================================================================================
// Slide 3 · Role profiles: one table
// =====================================================================================
{
  const s = pres.addSlide();
  title(s, "Role profiles: 19 roles, how each works with iFLYTEK, and how many people", 12.3);
  intro(s, "Role composition from Appendix A7; joint work from Appendices A10 and A3. Headcounts are peaks per function.");

  const FS = 8.5;
  const H = t => ({ text: t, options: { bold: true, color: INK, fill: { color: LIGHTGRAY }, fontSize: 9 } });
  const T = (name, n, span) => ({ text: [{ text: name, options: { breakLine: true } }, { text: `${n} people`, options: { color: BLUE } }], options: { rowspan: span, bold: true, color: NAVY, fill: { color: LIGHTBLUE }, fontSize: 9, valign: "middle" } });
  const role = r => ({ text: r, options: { bold: true, color: INK, fontSize: FS } });
  const num = n => ({ text: n, options: { bold: true, color: BLUE, align: "center", fontSize: 9.5 } });
  const does = t => ({ text: t, options: { color: BODY, fontSize: FS } });

  const rows = [
    [H("Team"), H("Role"), { ...H("People"), options: { ...H("People").options, align: "center" } }, H("How the role works with iFLYTEK")],
    [T("1  Project and product", 9, 3), role("Project manager"), num("2"), does("Runs the joint schedule with iFLYTEK's project team; aligns milestones, resources and acceptance at each of the five steps")],
    [role("Product manager"), num("4"), does("Defines the National AI Application with iFLYTEK's product team: functions, user scope, content rules, Brazilian content, roadmap")],
    [role("PMO / data-centre expert"), num("3"), does("Coordinates Huawei compute and environment readiness with iFLYTEK's deployment team; tracks T0 and the compute batches")],
    [T("2  Corpus and data", 32, 3), role("Data resource engineer"), num("6"), does("Sources and licenses Portuguese data to iFLYTEK's corpus specifications; agrees sources and permitted use with iFLYTEK")],
    [role("Data quality engineer"), num("6"), does("Runs the pilot batch and full-volume production on the Corpus Construction Center with iFLYTEK's sampling checks and guidance")],
    [role("Annotator"), num("20"), does("Annotates supervised, safety and speech data to standards agreed with iFLYTEK; iFLYTEK checks samples and feeds back")],
    [T("3  Portuguese language experts", 6, 1), role("Language expert"), num("6"), does("Sets Brazilian Portuguese quality, annotation, safety-label and speech standards with iFLYTEK; confirms rules on samples")],
    [T("4  Compliance and review", 6, 3), role("Review expert"), num("2"), does("Builds review policies, lexicons and risk samples on the Content Safety Review Platform with iFLYTEK's safety team")],
    [role("Legal expert"), num("2"), does("Clears data licences and permitted use with iFLYTEK's data team; reviews legal compliance of corpus, models and application")],
    [role("Review lead"), num("2"), does("Owns the review process; agrees escalation and approval rules with iFLYTEK")],
    [T("5  Business and domain experts", 7, 3), role("Business requirements"), num("3"), does("Defines the scenario for each of the three DLM services with iFLYTEK's DLM advisors; confirms inputs and service acceptance")],
    [role("Domain requirements"), num("2"), does("Specifies domain data and expected results; reviews evaluation reports with iFLYTEK and sets content boundaries")],
    [role("Domain expert"), num("2"), does("Assesses DLM results in the domain in joint evaluation exercises with iFLYTEK")],
    [T("6  Model and test engineering", 18, 2), role("Algorithm engineer"), num("12"), does("Trains and evaluates models on the Model Training Center with iFLYTEK's engineers; then trains the DLMs with iFLYTEK advising")],
    [role("Test engineer"), num("6"), does("Runs evaluation and validation against the rules agreed with iFLYTEK; documents results for acceptance")],
    [T("7  Application and operations", 27, 3), role("Software developer"), num("15"), does("Integrates the National AI Application with RNP systems; builds Brazilian content and features with iFLYTEK's app team")],
    [role("Operations engineer"), num("10"), does("Takes over application operations from iFLYTEK: configuration, user approval, API credentials and limits, usage analytics")],
    [role("Architect"), num("2"), does("Agrees application architecture and integration points with iFLYTEK's architects")],
    [T("8  Platform O&M", 6, 1), role("O&M engineer"), num("6"), does("Runs the platform with iFLYTEK during 12 months of assisted O&M, then independently: monitoring, incidents, changes, escalation")],
    [
      { text: "Total, peak per function", options: { colspan: 2, bold: true, color: NAVY, fill: { color: LIGHTBLUE }, fontSize: 9 } },
      { text: "111", options: { bold: true, color: BLUE, fill: { color: LIGHTBLUE }, align: "center", fontSize: 10 } },
      { text: "", options: { fill: { color: LIGHTBLUE } } },
    ],
  ];
  s.addTable(rows, { x: 0.5, y: 1.15, w: 12.33, colW: [2.05, 1.85, 0.7, 7.73], fontFace: FONT, fontSize: FS, color: BODY,
    border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.02, 0.07, 0.02, 0.07], rowH: 0.275, autoPage: false });
}

const OUT = "RNP-Team-for-Technology-Transfer-Skills-and-Headcount.pptx";
pres.writeFile({ fileName: OUT }).then(f => console.log("wrote", f));
