const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
// 32:9 ultrawide, single slide
pres.defineLayout({ name: "UW32x9", width: 20, height: 5.625 });
pres.layout = "UW32x9";

const F = "Microsoft YaHei";

const NAVY = "10305F";
const BLUE = "1A5FBF";
const BLUE_D = "0F3D80";
const TEAL = "0E8577";
const TEAL_D = "0A6659";
const TEAL_LT = "E7F3F0";
const PANEL_BG = "F7F9FC";
const PANEL_LN = "DCE4EF";
const CARD_LN = "D9E2EC";
const GRAY = "4B5A6B";
const GRAY_L = "6B7A8B";
const RED = "C62828";
const WHITE = "FFFFFF";

const slide = pres.addSlide();
slide.background = { color: WHITE };

const sh = () => ({ type: "outer", color: "9FB0C4", blur: 6, offset: 1, angle: 90, opacity: 0.28 });

function card(x, y, w, h, fill) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || WHITE },
    line: { color: CARD_LN, width: 0.75 },
    shadow: sh(),
  });
}

function badge(x, y, d, label, color) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color },
    line: { color, width: 0 },
  });
  slide.addText(label, {
    x, y, w: d, h: d, margin: 0, align: "center", valign: "middle",
    fontFace: F, fontSize: 9, bold: true, color: WHITE,
  });
}

function bullets(items, x, y, w, h, size, color) {
  slide.addText(
    items.map((t, i) => ({
      text: t,
      options: { breakLine: i < items.length - 1, bullet: { characterCode: "2013", indent: 9 } },
    })),
    {
      x, y, w, h, margin: 0, valign: "top",
      fontFace: F, fontSize: size, color, lineSpacingMultiple: 1.12,
      paraSpaceAfter: 1.5,
    }
  );
}

/* ───────────────────────── HEADER ───────────────────────── */

slide.addText("RNP-iFLYTEK 正式采购流程（1页概览）", {
  x: 0.4, y: 0.13, w: 8.7, h: 0.52, margin: 0, valign: "middle",
  fontFace: F, fontSize: 30, bold: true, color: NAVY, charSpacing: -0.3,
});
slide.addText("基于 2026/08/04 采购会议纪要（修正版）", {
  x: 0.43, y: 0.66, w: 8.7, h: 0.26, margin: 0, valign: "middle",
  fontFace: F, fontSize: 10.5, color: GRAY_L,
});

// 关键结论 card
card(9.4, 0.16, 10.25, 0.84, "F2F6FC");
slide.addShape(pres.ShapeType.roundRect, {
  x: 9.6, y: 0.06, w: 1.5, h: 0.3, rectRadius: 0.14,
  fill: { color: NAVY }, line: { color: NAVY, width: 0 }, shadow: sh(),
});
slide.addText("关键结论", {
  x: 9.6, y: 0.06, w: 1.5, h: 0.3, margin: 0, align: "center", valign: "middle",
  fontFace: F, fontSize: 10, bold: true, color: WHITE,
});

const concl = [
  ["1.", "A 阶段：A1 → A4 有先后顺序；A2 / A3 / A5 可并行"],
  ["2.", "部分需包含：讯飞自行聘请会计师事务所评估 CNAE / NBS 代码"],
  ["3.", "进入正式采购的触发条件仅两项：外部条件释放；Proposal 完全谈妥"],
  ["4.", "TR 编制前可无 CNPJ；最迟正式报价前需具备有效 CNPJ"],
];
const cCols = [{ x: 9.62, w: 4.45 }, { x: 14.25, w: 5.15 }];
concl.forEach((it, i) => {
  const col = cCols[i < 2 ? 0 : 1];
  const yy = 0.46 + (i % 2) * 0.25;
  slide.addText(
    [
      { text: it[0] + " ", options: { bold: true, color: BLUE } },
      { text: it[1], options: { color: "2B3A4A" } },
    ],
    { x: col.x, y: yy, w: col.w, h: 0.24, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5 }
  );
});

/* ───────────────────────── PANEL A ───────────────────────── */

const AX = 0.35, AW = 8.2, PY = 1.1, PH = 3.52;
const AIX = AX + 0.16, AIW = AW - 0.32;

slide.addShape(pres.ShapeType.roundRect, {
  x: AX, y: PY, w: AW, h: PH, rectRadius: 0.03,
  fill: { color: PANEL_BG }, line: { color: PANEL_LN, width: 0.75 },
});
slide.addShape(pres.ShapeType.roundRect, {
  x: AX, y: PY, w: AW, h: 0.38, rectRadius: 0.05,
  fill: { color: TEAL }, line: { color: TEAL, width: 0 },
});
slide.addText("A. 前期准备（部分串行 + 部分并行，尚未进入正式采购）", {
  x: AX, y: PY, w: AW, h: 0.38, margin: 0, align: "center", valign: "middle",
  fontFace: F, fontSize: 12.5, bold: true, color: WHITE,
});

slide.addText("顺序关系：先 A1，后 A4", {
  x: AIX, y: 1.52, w: AIW, h: 0.2, margin: 0, align: "center", valign: "middle",
  fontFace: F, fontSize: 10, bold: true, color: TEAL_D,
});

// A1 / A4
card(AIX, 1.74, 3.55, 0.78);
badge(AIX + 0.12, 1.84, 0.28, "A1", TEAL);
slide.addText("Proposal Draft 草案", {
  x: AIX + 0.48, y: 1.82, w: 2.95, h: 0.24, margin: 0, valign: "middle",
  fontFace: F, fontSize: 11, bold: true, color: NAVY,
});
bullets(["供 RNP 理解范围", "可先无 CNPJ"], AIX + 0.48, 2.06, 2.95, 0.40, 8.5, GRAY);

slide.addShape(pres.ShapeType.rightArrow, {
  x: AIX + 3.68, y: 2.04, w: 0.51, h: 0.18,
  fill: { color: TEAL }, line: { color: TEAL, width: 0 },
});

card(AIX + 4.33, 1.74, 3.55, 0.78);
badge(AIX + 4.45, 1.84, 0.28, "A4", TEAL);
slide.addText("RNP 提前分析", {
  x: AIX + 4.81, y: 1.82, w: 2.95, h: 0.24, margin: 0, valign: "middle",
  fontFace: F, fontSize: 11, bold: true, color: NAVY,
});
bullets(["初步理解 TR 范围", "预判 CAPEX / OPEX、NBS"], AIX + 4.81, 2.06, 2.95, 0.40, 8.5, GRAY);

slide.addText("可并行推进", {
  x: AIX, y: 2.54, w: AIW, h: 0.2, margin: 0, align: "center", valign: "middle",
  fontFace: F, fontSize: 10, bold: true, color: TEAL_D,
});

const parCards = [
  { id: "A2", t: "MoU / Proposal / 合同结构讨论", b: ["继续谈 MoU", "确认范围、结构、核心商务条款", "双语文本、服务、租赁结构"] },
  { id: "A3", t: "巴西主体 / CNPJ / 税务代码准备", b: ["公司设立", "CNPJ、CNAE、授权准备", "讯飞聘请会计师事务所评估 CNAE / NBS 代码"] },
  { id: "A5", t: "外部条件释放", b: ["政府 / 部门 / Casa Civil", "地点、项目条件明确"] },
];
const pcW = (AIW - 2 * 0.13) / 3;
parCards.forEach((c, i) => {
  const x = AIX + i * (pcW + 0.13);
  card(x, 2.76, pcW, 1.14);
  badge(x + 0.11, 2.86, 0.26, c.id, TEAL);
  slide.addText(c.t, {
    x: x + 0.43, y: 2.84, w: pcW - 0.56, h: 0.34, margin: 0, valign: "middle",
    fontFace: F, fontSize: 9.5, bold: true, color: NAVY, lineSpacingMultiple: 1.0,
  });
  bullets(c.b, x + 0.14, 3.20, pcW - 0.28, 0.66, 8, GRAY);
});

// trigger box
slide.addShape(pres.ShapeType.roundRect, {
  x: AIX, y: 3.94, w: AIW, h: 0.46, rectRadius: 0.05,
  fill: { color: TEAL_LT }, line: { color: TEAL, width: 0.75 },
});
slide.addText("进入正式采购阶段的触发条件", {
  x: AIX + 0.16, y: 3.97, w: 3.2, h: 0.19, margin: 0, valign: "middle",
  fontFace: F, fontSize: 9.5, bold: true, color: TEAL_D,
});
slide.addText([{ text: "外部条件释放 / 项目条件基本明确", options: { bullet: true } }], {
  x: AIX + 0.30, y: 4.17, w: 3.6, h: 0.19, margin: 0, valign: "middle",
  fontFace: F, fontSize: 9, color: "23414A", bullet: { indent: 9 },
});
slide.addText([{ text: "Proposal 完全谈妥", options: { bullet: true } }], {
  x: AIX + 4.05, y: 4.17, w: 3.4, h: 0.19, margin: 0, valign: "middle",
  fontFace: F, fontSize: 9, color: "23414A", bullet: { indent: 9 },
});

slide.addText(
  "注：A3（CNPJ / CNAE / NBS 评估）应尽快推进，但不是进入正式采购的硬触发条件；最迟需在正式报价前准备到位。",
  { x: AIX, y: 4.43, w: AIW, h: 0.16, margin: 0, valign: "middle", fontFace: F, fontSize: 7.5, color: GRAY_L }
);

/* ───────────────────────── DIVIDER ───────────────────────── */

slide.addShape(pres.ShapeType.line, {
  x: 8.975, y: PY + 0.05, w: 0, h: PH - 0.1,
  line: { color: RED, width: 1.25, dashType: "dash" },
});
slide.addShape(pres.ShapeType.ellipse, {
  x: 8.83, y: 2.05, w: 0.29, h: 0.29,
  fill: { color: RED }, line: { color: RED, width: 0 }, shadow: sh(),
});
slide.addShape(pres.ShapeType.rightArrow, {
  x: 8.885, y: 2.145, w: 0.18, h: 0.1,
  fill: { color: WHITE }, line: { color: WHITE, width: 0 },
});
slide.addText("从这里开始\n进入正式采购", {
  x: 8.55, y: 2.42, w: 0.85, h: 0.4, margin: 0, align: "center", valign: "top",
  fontFace: F, fontSize: 8, bold: true, color: RED, lineSpacingMultiple: 1.1,
});

/* ───────────────────────── PANEL B ───────────────────────── */

const BX = 9.4, BW = 10.25;
const BIX = BX + 0.16, BIW = BW - 0.32;

slide.addShape(pres.ShapeType.roundRect, {
  x: BX, y: PY, w: BW, h: PH, rectRadius: 0.03,
  fill: { color: PANEL_BG }, line: { color: PANEL_LN, width: 0.75 },
});
slide.addShape(pres.ShapeType.roundRect, {
  x: BX, y: PY, w: BW, h: 0.38, rectRadius: 0.05,
  fill: { color: BLUE }, line: { color: BLUE, width: 0 },
});
slide.addText("B. 正式采购流程（串行推进）", {
  x: BX, y: PY, w: BW, h: 0.38, margin: 0, align: "center", valign: "middle",
  fontFace: F, fontSize: 12.5, bold: true, color: WHITE,
});

const steps = [
  { t: "RNP 在 ERP\n发起需求", b: ["正式内部流程启动"] },
  { t: "正式编制 TR", b: ["明确范围、交付物、技术要求", "编制阶段本身可不依赖 CNPJ"] },
  { t: "RNP 正式索取报价 /\niFLYTEK 提交正式报价", b: ["需具备有效 CNPJ", "服务内容需匹配 CNAE"] },
  { t: "正式报价 + 内部审查\n+ 合同定稿", b: ["采购 / 项目 / 法务 / 控制部门审查", "重点：CNAE / NBS、CAPEX vs OPEX、软件权利"] },
  { t: "签署完成 →\n履约 / 开票 / 付款", b: ["签约后进入履约与付款", "开票通常仍需市政登记 / ISS 条件"] },
];
const stGap = 0.24;
const stW = (BIW - 4 * stGap) / 5;
steps.forEach((s, i) => {
  const x = BIX + i * (stW + stGap);
  card(x, 1.58, stW, 1.98);
  badge(x + 0.13, 1.70, 0.3, String(i + 1), BLUE_D);
  slide.addText(s.t, {
    x: x + 0.13, y: 2.06, w: stW - 0.26, h: 0.58, margin: 0, valign: "top",
    fontFace: F, fontSize: 10, bold: true, color: NAVY, lineSpacingMultiple: 1.05,
  });
  bullets(s.b, x + 0.13, 2.70, stW - 0.26, 0.76, 8, GRAY);
});

for (let i = 0; i < 4; i++) {
  const ax = BIX + i * (stW + stGap) + stW + 0.025;
  slide.addShape(pres.ShapeType.rightArrow, {
    x: ax, y: 2.47, w: 0.19, h: 0.17,
    fill: { color: BLUE }, line: { color: BLUE, width: 0 },
  });
}

const bbW = (BIW - 0.21) / 2;
card(BIX, 3.62, bbW, 0.94, "EEF4FC");
slide.addText("正式采购流程整体周期", {
  x: BIX + 0.16, y: 3.70, w: bbW - 0.32, h: 0.22, margin: 0, valign: "middle",
  fontFace: F, fontSize: 10, bold: true, color: BLUE_D,
});
slide.addText("按 1 个月测算。\n若 9 月初进入正式流程，则可按约 1 个月推进至签约完成。", {
  x: BIX + 0.16, y: 3.94, w: bbW - 0.32, h: 0.52, margin: 0, valign: "top",
  fontFace: F, fontSize: 8.5, color: GRAY, lineSpacingMultiple: 1.15,
});

card(BIX + bbW + 0.21, 3.62, bbW, 0.94, "EEF4FC");
slide.addText("9/30 倒排目标", {
  x: BIX + bbW + 0.37, y: 3.70, w: bbW - 0.32, h: 0.22, margin: 0, valign: "middle",
  fontFace: F, fontSize: 10, bold: true, color: BLUE_D,
});
slide.addText("8 月完成 A 区关键准备；9 月进入 B 流程；\n以 9/30 完成签约为目标。", {
  x: BIX + bbW + 0.37, y: 3.94, w: bbW - 0.32, h: 0.52, margin: 0, valign: "top",
  fontFace: F, fontSize: 8.5, color: GRAY, lineSpacingMultiple: 1.15,
});

/* ───────────────────────── TIMELINE ───────────────────────── */

slide.addShape(pres.ShapeType.roundRect, {
  x: 0.35, y: 4.70, w: 19.3, h: 0.34, rectRadius: 0.05,
  fill: { color: "EEF3FA" }, line: { color: "DCE6F3", width: 0.75 },
});
const tl = ["8 月：A 区关键准备", "9 月：B 流程（整体约 1 个月）", "9/30：签约完成"];
const tlColor = [TEAL_D, BLUE_D, NAVY];
tl.forEach((t, i) => {
  slide.addText(t, {
    x: 1.8 + i * 6.0, y: 4.70, w: 4.4, h: 0.34, margin: 0, align: "center", valign: "middle",
    fontFace: F, fontSize: 11, bold: true, color: tlColor[i],
  });
});
for (let i = 0; i < 2; i++) {
  slide.addShape(pres.ShapeType.rightArrow, {
    x: 6.45 + i * 6.0, y: 4.83, w: 1.1, h: 0.09,
    fill: { color: "9AB2CE" }, line: { color: "9AB2CE", width: 0 },
  });
}

/* ───────────────────────── BANNER ───────────────────────── */

slide.addShape(pres.ShapeType.roundRect, {
  x: 0.35, y: 5.08, w: 19.3, h: 0.42, rectRadius: 0.05,
  fill: { color: NAVY }, line: { color: NAVY, width: 0 },
});
slide.addText(
  "建议：尽快完成 Proposal Draft 与 Proposal 谈判，并同步推进 CNPJ / CNAE / NBS 准备，争取 9 月进入正式采购流程、9/30 完成签约。",
  { x: 0.35, y: 5.08, w: 19.3, h: 0.42, margin: 0, align: "center", valign: "middle",
    fontFace: F, fontSize: 11.5, bold: true, color: WHITE }
);

slide.addNotes(
  "RNP-iFLYTEK 正式采购流程一页概览（32:9）。左侧 A 区为尚未进入正式采购的前期准备：A1 Proposal Draft 先行，A4 RNP 提前分析随后；A2 合同结构讨论、A3 巴西主体与 CNPJ/CNAE/NBS 准备、A5 外部条件释放三线并行。进入正式采购的硬触发条件只有两项：外部条件释放与 Proposal 完全谈妥。右侧 B 区为串行的正式采购五步，整体按 1 个月测算，目标 9/30 完成签约。"
);

pres.writeFile({ fileName: "RNP-iFLYTEK-采购流程-32x9.pptx" }).then(() => console.log("done"));
