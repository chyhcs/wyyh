// Iara's request (WhatsApp, 13 Sep 2026): "the skills we need on our side, the number of people"
//   -> 2-slide English deck in the style of the 9 Sep exchange deck: cover + ONE table
//      (team / role / people / must be able to). Pre-filled from the 9 Sep deck (p.13, A3, A7, A10).
//      Yellow TBC tag on slide 2 + Chinese speaker notes on slide 1 are internal: remove / export PDF before sending.
//   node build_rnp_team_pptx.js  ->  RNP-Team-for-Technology-Transfer-Skills-and-Headcount.pptx
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "RNP Team for the Technology Transfer: Skills and Headcount";

const FONT = "Calibri";
const RED = "C00000", NAVY = "1F4E79", BLUE = "0070C0", INK = "1D1D1A", BODY = "333333", GRAY = "666666";
const LIGHTBLUE = "E8F0FA", LIGHTGRAY = "F2F2F2", LINE = "D9D9D9";

// =====================================================================================
// Slide 1 · Cover
// =====================================================================================
{
  const s = pres.addSlide();
  s.addText("RNP × iFLYTEK · Follow-up to the Technical Exchange of 9 September", { x: 0.8, y: 1.5, w: 8.6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: GRAY, margin: 0, isTextBox: true });
  s.addText("RNP Team for the Technology Transfer: Skills and Headcount", { x: 0.8, y: 1.98, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 27, bold: true, color: RED, margin: 0, valign: "top", isTextBox: true });
  s.addText("The roles RNP needs on its side, what each must be able to do, and how many people", { x: 0.8, y: 3.25, w: 8.6, h: 0.6, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
  s.addText([
    { text: "Prepared for RNP at Iara's request of 13 September 2026", options: { breakLine: true } },
    { text: "Draws on the presentation of 9 September 2026 and its appendices A3, A7 and A10", options: { breakLine: true } },
    { text: "iFLYTEK · Draft v0.1 · 14 September 2026" },
  ], { x: 0.8, y: 4.1, w: 8.6, h: 1.3, fontFace: FONT, fontSize: 13, color: BODY, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 4 });

  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 1.75, w: 2.9, h: 2.0, fill: { color: LIGHTGRAY }, line: { color: LIGHTGRAY }, rectRadius: 0.06 });
  s.addText("Your request, 13 September", { x: 10.15, y: 1.85, w: 2.45, h: 0.35, fontFace: FONT, fontSize: 11.5, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addText("“We are organizing RNP's side to receive the technology transfer and prepare to work together with you. Any important information about the skills we need on our side, the number of people, will be welcome.”", { x: 10.15, y: 2.22, w: 2.45, h: 1.45, fontFace: FONT, fontSize: 10, italic: true, color: BODY, margin: 0, valign: "top", isTextBox: true });
  s.addShape(pres.ShapeType.roundRect, { x: 9.9, y: 3.95, w: 2.9, h: 0.72, fill: { color: RED }, line: { color: RED }, rectRadius: 0.06 });
  s.addText("One table: 8 teams, 19 roles, 111 people at peak", { x: 10.15, y: 3.95, w: 2.6, h: 0.72, fontFace: FONT, fontSize: 10.5, bold: true, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true });

  s.addNotes(
    "【内部说明，发送前处理】\n" +
    "1. 本稿只回答 Iara 9 月 13 日 WhatsApp 的两问：技能、人数。一张表：团队 / 岗位 / 人数 / 必须会做什么。内容按 9 月 9 日英文 PPT 填（第 13 页与附录 A3、A7、A10）。\n" +
    "2. 人数口径要二选一：PPT 第 13 页写三批约 30 / 20 / 30 人，附录 A7 按岗位合计 53 / 25 / 33 人、总计 111。本稿按 A7 明细填。第 2 页右上角黄色标签是提醒，定稿后删除。\n" +
    "3. 发送前导出 PDF；备注不会进入 PDF。"
  );
}

// =====================================================================================
// Slide 2 · The one table
// =====================================================================================
{
  const s = pres.addSlide();
  s.addText("RNP team for the technology transfer: 8 teams, 19 roles, 111 people at peak", { x: 0.52, y: 0.16, w: 8.0, h: 0.5, fontFace: FONT, fontSize: 18, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });
  s.addText("What each role must be able to do, and how many people. Headcounts are peaks per function, reusable across activities, not people on site at once. Source: presentation of 9 September, p.13 and Appendices A3, A7, A10.", { x: 0.5, y: 0.7, w: 12.33, h: 0.4, fontFace: FONT, fontSize: 10.5, color: INK, margin: 0, valign: "top", isTextBox: true });
  // internal TBC tag (delete before sending)
  s.addShape(pres.ShapeType.roundRect, { x: 8.7, y: 0.17, w: 4.15, h: 0.46, fill: { color: "FFFF00" }, line: { color: "FFC000" }, rectRadius: 0.05 });
  s.addText("TBC before sending: p.13 says about 30 / 20 / 30 per intake; A7 detail sums to 53 / 25 / 33. Confirm one set.", { x: 8.77, y: 0.17, w: 4.03, h: 0.46, fontFace: FONT, fontSize: 8, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });

  const FS = 8.5;
  const H = t => ({ text: t, options: { bold: true, color: INK, fill: { color: LIGHTGRAY }, fontSize: 9 } });
  const T = (name, n, span) => ({ text: [{ text: name, options: { breakLine: true } }, { text: `${n} people`, options: { color: BLUE } }], options: { rowspan: span, bold: true, color: NAVY, fill: { color: LIGHTBLUE }, fontSize: 9, valign: "middle" } });
  const role = r => ({ text: r, options: { bold: true, color: INK, fontSize: FS } });
  const num = n => ({ text: n, options: { bold: true, color: BLUE, align: "center", fontSize: 9.5 } });
  const does = t => ({ text: t, options: { color: BODY, fontSize: FS } });

  const rows = [
    [H("Team"), H("Role"), { ...H("People"), options: { ...H("People").options, align: "center" } }, H("Must be able to")],
    [T("1  Project and product", 9, 3), role("Project manager"), num("2"), does("Run the five-step schedule with iFLYTEK and Huawei; manage resources, risks and delivery; organise acceptance")],
    [role("Product manager"), num("4"), does("Define the National AI Application with iFLYTEK: functions, user scope, content rules, Brazilian content, roadmap")],
    [role("PMO / data-centre expert"), num("3"), does("Coordinate Huawei compute, facilities, network and base environment; track T0 and the compute batches")],
    [T("2  Corpus and data", 32, 3), role("Data resource engineer"), num("6"), does("Identify data sources, secure licences and permitted use, organise acquisition; keep data in RNP's Brazilian environment")],
    [role("Data quality engineer"), num("6"), does("Apply the agreed sampling, quality-threshold and re-check rules; run the pilot batch; operate the Corpus Construction Center")],
    [role("Annotator"), num("20"), does("Annotate to the agreed standards: 200,000 supervised and 20,000 safety samples; transcribe 10,000 hours of speech")],
    [T("3  Portuguese language experts", 6, 1), role("Language expert"), num("6"), does("Set Brazilian Portuguese quality, annotation, safety-label and speech (ASR/TTS) standards; analyse samples with iFLYTEK")],
    [T("4  Compliance and review", 6, 3), role("Review expert"), num("2"), does("Define review policies, sensitive lexicons and risk samples on the Content Safety Review Platform; review content")],
    [role("Legal expert"), num("2"), does("Clear data licences and permitted use; legal compliance of corpus, models and application")],
    [role("Review lead"), num("2"), does("Own the review process and its records; escalate and approve")],
    [T("5  Business and domain experts", 7, 3), role("Business requirements"), num("3"), does("Define the business scenario for each of the three DLM services; confirm inputs and service acceptance")],
    [role("Domain requirements"), num("2"), does("Specify domain data and expected results; read evaluation reports; set content boundaries")],
    [role("Domain expert"), num("2"), does("Assess model results in the domain; take part in evaluation exercises with iFLYTEK")],
    [T("6  Model and test engineering", 18, 2), role("Algorithm engineer"), num("12"), does("Prepare datasets, run training jobs, read logs and evaluations, release models; train the DLMs with iFLYTEK advising")],
    [role("Test engineer"), num("6"), does("Run model evaluation and technical validation against the agreed rules; document results for acceptance")],
    [T("7  Application and operations", 27, 3), role("Software developer"), num("15"), does("Integrate the app (mobile, desktop, web) with RNP systems; build Brazilian content and features on the baseline; use the APIs")],
    [role("Operations engineer"), num("10"), does("Operate the application: configuration, user approval and permissions, API credentials and limits, usage analytics")],
    [role("Architect"), num("2"), does("Own application architecture and its integration with RNP systems and the platform")],
    [T("8  Platform O&M", 6, 1), role("O&M engineer"), num("6"), does("Monitor platform and cluster; handle incidents, changes, faults, deployments and escalation; run the platform after handover")],
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
