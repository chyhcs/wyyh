// 回应 Iara 9/13 WhatsApp（RNP 侧需要什么技能的人、需要多少人）—— 中文工作稿 docx  v0.2
//   第一部分：编号 / 分组 / 岗位 / 人数 / 岗位一句话（总览表，合计 111）
//   第二部分：逐岗位画像（编号 1.1 … 8.1）：人员要求（≤3 条）+ 与讯飞协作（1–2 条）
//   人数与分组按 9/9 PPT 附录 A7；协作形式依据附录 A10、A3、A5、A6；人员要求为讯飞起草建议
//   node build_rnp_team_cn_docx.js  ->  RNP侧团队岗位画像与人数建议-中文稿-20260914.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
        ShadingType, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, TabStopType } = require("docx");

const F = "微软雅黑";
const RED = "C00000", NAVY = "1F3A5F", BLUE = "0057A8", GRAY = "6B6B6B", BLACK = "1D1D1A", LIGHT = "F3F6FA", LINE = "C9D3E0";
const PAGE_W = 9638; // A4 内容宽度（DXA）
const D = [];

const R = (t, o = {}) => new TextRun({ text: t, font: F, size: o.size || 21, bold: !!o.bold, color: o.color || BLACK });
function runs(t, o = {}) { return String(t).split("**").map((seg, i) => seg ? R(seg, { ...o, bold: o.bold || i % 2 === 1 }) : null).filter(Boolean); }
const P = (t, o = {}) => D.push(new Paragraph({ children: runs(t, o), spacing: { after: o.after == null ? 120 : o.after, line: 320 }, ...(o.indent ? { indent: { left: o.indent } } : {}) }));
const H1 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 420, after: 160 }, keepNext: true, children: [R(t, { size: 28, bold: true, color: NAVY })] }));
function H2(num, name, n, intro) {
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 360, after: 40 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }],
    children: [R(`${num}　${name}`, { size: 25, bold: true, color: NAVY }), R(`\t${n} 人`, { size: 22, bold: true, color: BLUE })] }));
  D.push(new Paragraph({ spacing: { after: 140 }, keepNext: true, children: [R(intro, { size: 19, color: GRAY })] }));
}
function H3(num, name, en, n) {
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }],
    children: [R(`${num}　${name}`, { size: 22, bold: true, color: BLACK }), R(`　${en}`, { size: 18, color: GRAY }), R(`\t${n} 人`, { size: 21, bold: true, color: BLUE })] }));
}

// ---- 表格 ----
const TH = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function cell(children, w, o = {}) {
  return new TableCell({
    children, width: { size: w, type: WidthType.DXA },
    borders: { top: TH, bottom: TH, left: o.vlines ? TH : NONE, right: o.vlines ? TH : NONE },
    margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.CENTER,
    ...(o.rowSpan ? { rowSpan: o.rowSpan } : {}), ...(o.colSpan ? { columnSpan: o.colSpan } : {}),
    ...(o.fill ? { shading: { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } } : {}),
  });
}
const tp = (t, o = {}) => new Paragraph({ children: runs(t, { size: o.size || 19, bold: o.bold, color: o.color }), spacing: { after: 20, line: 270 }, alignment: o.align || AlignmentType.LEFT });
const bp = t => new Paragraph({ children: runs(t, { size: 19 }), numbering: { reference: "cellb", level: 0 }, spacing: { after: 30, line: 280 } });

// =====================================================================================
// 数据：8 个分组、19 个岗位
// =====================================================================================
const GROUPS = [
  { name: "项目与产品", n: 9, intro: "负责项目在 RNP 侧的组织、进度、资源与验收，以及国民 AI 应用的产品定义。第一批到位。", roles: [
    { name: "项目经理", en: "Project Manager", n: 2, one: "统筹 RNP 侧进度、资源与验收，对接讯飞与华为",
      req: ["大型 IT 或基础设施项目的交付管理经验，能协调 RNP、讯飞、华为与部委四方", "熟悉政府项目的合同与验收流程", "英语可作为工作语言"],
      col: ["对口讯飞项目组：共同维护总进度与里程碑，参加例会与阶段评审", "代表 RNP 组织并签署各项验收"] },
    { name: "产品经理", en: "Product Manager", n: 4, one: "定义国民 AI 应用的功能、用户范围与内容规则",
      req: ["面向公众或政务的互联网产品经验", "能把巴西本地内容与使用场景转化为产品需求", "了解大模型应用的能力与边界"],
      col: ["对口讯飞产品团队：在标准版基础上共同定义功能、用户范围、内容规则与路线图", "影随后接手应用配置与发布"] },
    { name: "PMO / 数据中心专家", en: "PMO / Data Centre Specialist", n: 3, one: "协调华为算力、机房、网络与基础环境就绪",
      req: ["熟悉数据中心与 AI 算力集群建设", "能与华为侧对接技术方案与进度", "有多项目并行的 PMO 经验"],
      col: ["对口讯飞部署团队：确认平台部署所需的环境与系统接口，同步算力批次进度"] },
  ] },
  { name: "语料与数据资源", n: 32, intro: "组织葡语语料的获取、授权、生产与质检，是第一批到位的主力。", roles: [
    { name: "数据资源工程师", en: "Data Resource Engineer", n: 6, one: "寻找数据来源，落实授权，组织语料获取",
      req: ["熟悉巴西公开与机构数据资源：政府、高校、媒体、科研数据库", "懂数据授权与许可，能审阅来源许可条款", "有数据采集与整理经验，葡语母语"],
      col: ["对口讯飞数据团队：按语料规范确定来源与配额，逐来源确认许可", "基础课程加影随，掌握数据源、许可与平台处理流程"] },
    { name: "数据质检工程师", en: "Data Quality Engineer", n: 6, one: "执行语料的计数、抽样、质量阈值与复检规则",
      req: ["数据处理或数据工程背景，会 Python 与 SQL", "能操作语料平台的清洗、去重与质量度量流程", "有质量管理意识，能维护质检记录"],
      col: ["对口讯飞数据团队：共同制定质检规则，先试点批次后全量生产", "RNP 生产，讯飞抽检并反馈"] },
    { name: "标注员", en: "Annotator", n: 20, one: "按规范完成文本标注、安全样本与语音转写",
      req: ["葡语母语，书面表达规范，细致稳定", "经培训后能按标注规范作业，转写岗需良好听写能力", "可部分由高校学生或外包承担，保留固定核心人员"],
      col: ["使用讯飞的标注规范、样例与工具作业，讯飞按批抽检反馈"] },
  ] },
  { name: "葡语语言专家", n: 6, intro: "定义巴西葡语的质量、标注与语音标准。", roles: [
    { name: "语言专家", en: "Portuguese Language Expert", n: 6, one: "制定葡语质量、标注规则、安全标签与语音标准",
      req: ["语言学、计算语言学或语音学背景", "熟悉巴西葡语的地区差异与当代用法", "有语料库或标注项目经验"],
      col: ["对口讯飞语言与语音团队：共同走查标注标准与安全标签，确认规则", "评价模型的葡语输出质量"] },
  ] },
  { name: "合规与审核", n: 6, intro: "负责内容、安全与法律合规，验收后承接内容安全审核平台。", roles: [
    { name: "审核专家", en: "Content Review Specialist", n: 2, one: "建立审核策略、敏感词库与风险样本",
      req: ["内容安全或信息审核经验", "了解巴西法律与社会语境下的敏感内容边界", "能定义并持续维护审核规则"],
      col: ["对口讯飞安全团队：在内容安全审核平台上共同配置策略、词库与样本", "验收后独立运营"] },
    { name: "法务专家", en: "Legal Counsel", n: 2, one: "落实数据许可与合规，审核语料、模型与应用的法律风险",
      req: ["数据保护法（LGPD）、知识产权与许可法律背景", "能审阅数据来源许可与开源许可条款", "有政府采购与合同经验者优先"],
      col: ["对口讯飞数据团队与商务：逐项确认数据来源许可，审阅权重、平台与应用的许可条款"] },
    { name: "审核负责人", en: "Review Lead", n: 2, one: "负责审核流程、记录与升级审批",
      req: ["合规或审核团队管理经验", "能定义流程、留痕与升级机制", "能与法务、业务和技术团队协同"],
      col: ["与讯飞约定审核升级与审批规则；作为审核平台验收与移交的 RNP 侧负责人"] },
  ] },
  { name: "业务与领域专家", n: 7, intro: "为三个领域大模型确定业务场景并评估领域结果。第二批到位。", roles: [
    { name: "业务需求专家", en: "Business Requirements Lead", n: 3, one: "为每个领域模型服务定义业务场景、输入与验收",
      req: ["熟悉对应领域的业务流程", "能把业务问题转化为模型可交付的场景", "能组织领域机构参与"],
      col: ["对口讯飞领域模型顾问：每个领域共同确定一个场景、输入数据与验收方式"] },
    { name: "领域需求专家", en: "Domain Requirements Specialist", n: 2, one: "明确领域数据与预期结果，设定内容边界",
      req: ["领域数据与知识背景", "能定义评测集与预期输出，读懂评测报告"],
      col: ["对口讯飞领域模型顾问：共同定义评测方法与内容边界，参与评测练习"] },
    { name: "领域专家", en: "Domain Expert", n: 2, one: "在领域内评价模型输出结果",
      req: ["领域内资深专业人员，可兼职或短期借用高校、机构专家", "能判断模型输出的正确性与适用性"],
      col: ["参加与讯飞的联合评测，出具领域评价意见"] },
  ] },
  { name: "模型与测试工程", n: 18, intro: "在模型训练中心参与国家模型的训练与评测，之后自主训练领域模型。第二批到位，是知识转移的核心受训对象。", roles: [
    { name: "算法工程师", en: "AI / Algorithm Engineer", n: 12, one: "在平台上训练与评测模型，承担领域模型训练",
      req: ["机器学习或深度学习背景，熟练 Python 与 PyTorch 等训练框架", "理解大模型预训练、微调与评测流程，能读训练日志并定位问题", "骨干需能修改训练脚本、适配新开源模型；有分布式训练经验者优先"],
      col: ["对口讯飞模型团队：影随参与国家模型训练与评测，指导实操下自主提交作业", "领域模型由 RNP 训练，讯飞提供顾问支持"] },
    { name: "测试工程师", en: "Model Test Engineer", n: 6, one: "按约定评测规则做模型评测与技术验证",
      req: ["软件测试或模型评测背景", "能构建评测集、执行评测脚本、分析指标", "文档与记录规范"],
      col: ["对口讯飞模型团队：共同确定评测规则；执行评测并记录结果，作为验收依据"] },
  ] },
  { name: "应用与运营", n: 27, intro: "接手国民 AI 应用的架构、集成、用户审批、内容与日常运营。第三批到位。", roles: [
    { name: "软件开发工程师", en: "Software Developer", n: 15, one: "集成国民 AI 应用与 RNP 系统，开发本地内容与功能",
      req: ["移动端（Android、iOS）与 Web 前后端开发能力", "系统集成与 API 开发经验", "能在交付的源码基础上二次开发"],
      col: ["对口讯飞应用团队：在交付基线与源码上共同开发本地功能，对接 RNP 系统接口", "影随后接手开发与发布"] },
    { name: "运营工程师", en: "Application Operations Engineer", n: 10, one: "运营应用：配置、用户审批、API 凭证与用量分析",
      req: ["互联网产品或政务服务的运营经验", "能处理用户审批与权限管理", "能读用量数据并反馈问题"],
      col: ["对口讯飞应用与支持团队：辅助运维期并行，之后独立运营"] },
    { name: "架构师", en: "Solution Architect", n: 2, one: "负责应用架构与 RNP 系统集成方案",
      req: ["企业级系统架构经验，熟悉云、容器、微服务与 API 网关", "了解模型服务调用与安全要求"],
      col: ["对口讯飞架构师：共同确定应用架构、集成点与安全方案，评审接口与部署方案"] },
  ] },
  { name: "平台运维", n: 6, intro: "验收后运行国家 AI 能力中心平台。第三批到位。", roles: [
    { name: "运维工程师", en: "Platform O&M Engineer", n: 6, one: "监控、事件、变更、部署与升级，运行平台",
      req: ["Linux、容器（Kubernetes）与集群运维经验", "熟悉监控告警、事件处理与变更管理", "了解 AI 算力集群与训练平台运维者优先"],
      col: ["对口讯飞运维支持团队：按运维手册在指导下操作并完成独立验证", "辅助运维期并行值守，之后独立运行"] },
  ] },
];
const TOTAL = GROUPS.reduce((a, g) => a + g.n, 0);
const NROLES = GROUPS.reduce((a, g) => a + g.roles.length, 0);

// =====================================================================================
// 封面
// =====================================================================================
D.push(new Paragraph({ spacing: { after: 100 }, children: [R("RNP 侧项目团队：岗位画像与人数建议", { size: 34, bold: true, color: RED })] }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [R("中文工作稿 · 供内部审定后翻译为英文发 Iara", { size: 23, bold: true, color: NAVY })] }));
D.push(new Paragraph({ spacing: { after: 320 }, children: [R("回应 Iara 2026 年 9 月 13 日 WhatsApp 请求　·　依据 9 月 9 日交流 PPT 第 13 页与附录 A3、A7、A10　·　草稿 v0.2 · 2026 年 9 月 14 日", { size: 18, color: GRAY })] }));
P(`本稿只回答 Iara 的两问：RNP 侧需要什么样的人，需要多少人。第一部分是一张总览表；第二部分逐岗位说明**人员要求**和**与讯飞的协作方式**。共 ${GROUPS.length} 个分组、${NROLES} 个岗位、${TOTAL} 人。人数为各职能的峰值，可跨活动复用，不是同时在岗人数。`);

// =====================================================================================
// 一、总览
// =====================================================================================
H1("一、岗位与人数总览");
const W = [640, 1660, 2060, 640, 4638];
const hdr = ["编号", "分组", "岗位", "人数", "岗位一句话"];
const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: hdr.map((h, i) => cell([tp(h, { bold: true, size: 19, color: NAVY, align: i === 3 || i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })], W[i], { fill: "DCE6F1" })) })];
GROUPS.forEach((g, gi) => {
  g.roles.forEach((r, ri) => {
    const c = [];
    c.push(cell([tp(`${gi + 1}.${ri + 1}`, { align: AlignmentType.CENTER, color: GRAY })], W[0]));
    if (ri === 0) c.push(cell([tp(`**${gi + 1}　${g.name}**`, { color: NAVY }), tp(`${g.n} 人`, { size: 18, color: BLUE })], W[1], { rowSpan: g.roles.length, fill: LIGHT }));
    c.push(cell([tp(`**${r.name}**`)], W[2]));
    c.push(cell([tp(`**${r.n}**`, { align: AlignmentType.CENTER, color: BLUE })], W[3]));
    c.push(cell([tp(r.one, { size: 18 })], W[4]));
    rows.push(new TableRow({ cantSplit: true, children: c }));
  });
});
rows.push(new TableRow({ cantSplit: true, children: [
  cell([tp("**合计（各职能峰值）**", { color: NAVY })], W[0] + W[1] + W[2], { colSpan: 3, fill: LIGHT }),
  cell([tp(`**${TOTAL}**`, { align: AlignmentType.CENTER, color: RED, size: 21 })], W[3], { fill: LIGHT }),
  cell([tp("")], W[4], { fill: LIGHT }),
] }));
D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: W, rows }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

// =====================================================================================
// 二、岗位画像
// =====================================================================================
H1("二、岗位画像");
P("每个岗位两行：**人员要求**说明要什么样的人；**与讯飞协作**说明这个岗位对口讯飞哪支团队、以什么方式一起工作。编号与总览表一致。", { after: 60 });

const WR = [1300, PAGE_W - 1300];
GROUPS.forEach((g, gi) => {
  H2(`${gi + 1}`, g.name, g.n, g.intro);
  g.roles.forEach((r, ri) => {
    H3(`${gi + 1}.${ri + 1}`, r.name, r.en, r.n);
    const t = new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: WR, rows: [
      new TableRow({ cantSplit: true, children: [
        cell([tp("人员要求", { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }),
        cell(r.req.map(bp), WR[1], { valign: VerticalAlign.TOP }),
      ] }),
      new TableRow({ cantSplit: true, children: [
        cell([tp("与讯飞协作", { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }),
        cell(r.col.map(bp), WR[1], { valign: VerticalAlign.TOP }),
      ] }),
    ] });
    D.push(t);
    D.push(new Paragraph({ spacing: { after: 40 }, children: [] }));
  });
});

// =====================================================================================
// 附：内部说明
// =====================================================================================
D.push(new Paragraph({ spacing: { before: 260, after: 60 }, keepNext: true, children: [R("附　内部说明（英文版不含）", { size: 19, bold: true, color: GRAY })] }));
[
  "1. Iara 原话：\"We are organizing RNP's side to receive the technology transfer and prepare to work together with you. Any important information about the skills we need on our side, the number of people, will be welcome.\"",
  `2. 分组、岗位与人数按 PPT 附录 A7，合计 ${TOTAL} 人（三批 53 / 25 / 33）。PPT 第 13 页写的三批约 30 / 20 / 30 为另一口径，定稿前统一为一套。`,
  "3. \"人员要求\"为讯飞起草的建议，PPT 中没有，请审定；\"与讯飞协作\"依据附录 A10、A3、A5、A6。",
].forEach(t => D.push(new Paragraph({ children: runs(t, { size: 17, color: GRAY }), spacing: { after: 30, line: 270 } })));

// =====================================================================================
const bullet = (ref, left, hanging) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });
const doc = new Document({
  creator: "iFLYTEK", title: "RNP 侧项目团队：岗位画像与人数建议",
  styles: { default: { document: { run: { font: F, size: 21, color: BLACK } } } },
  numbering: { config: [bullet("cellb", 260, 220)] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      R("RNP 侧项目团队：岗位画像与人数建议 · 中文工作稿 v0.2 · 第 ", { size: 16, color: GRAY }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 16, color: GRAY }), R(" 页", { size: 16, color: GRAY }) ] })] }) },
    children: D,
  }],
});
const OUT = "RNP侧团队岗位画像与人数建议-中文稿-20260914.docx";
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes"); });
