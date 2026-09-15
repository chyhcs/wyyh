// 回应 Iara 9/13 WhatsApp（RNP 侧需要什么技能的人、需要多少人）—— 中文工作稿 docx  v0.5
//   在总部研发 v0.4（26 个岗位、111 人）基础上，加入"核心团队 / 项目资源"两层逻辑：
//     核心团队 43 人 = RNP 编制，接知识转移、项目后常设；项目资源 68 人 = 津贴、合同、合作机构，按工作包进出
//   第一部分：总览表（编号 / 分组 / 岗位 / 核心 / 资源 / 岗位一句话）
//   第二部分：核心团队与项目资源（划分依据 + 外部人员由谁管理）
//   第三部分：逐岗位画像（人员要求 ≤3 条 + 与讯飞协作 ≤2 条），措辞沿用研发 v0.4
//   node build_rnp_team_cn_docx.js  ->  RNP侧团队岗位画像与人数建议-中文稿-v0.5-20260915.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
        ShadingType, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, TabStopType } = require("docx");

const F = "微软雅黑";
const RED = "C00000", NAVY = "1F3A5F", BLUE = "0057A8", GRAY = "6B6B6B", BLACK = "1D1D1A", LIGHT = "F3F6FA", LINE = "C9D3E0";
const PAGE_W = 9638;
const D = [];

const R = (t, o = {}) => new TextRun({ text: t, font: F, size: o.size || 21, bold: !!o.bold, color: o.color || BLACK });
function runs(t, o = {}) { return String(t).split("**").map((seg, i) => seg ? R(seg, { ...o, bold: o.bold || i % 2 === 1 }) : null).filter(Boolean); }
const P = (t, o = {}) => D.push(new Paragraph({ children: runs(t, o), spacing: { after: o.after == null ? 120 : o.after, line: 320 }, ...(o.indent ? { indent: { left: o.indent } } : {}) }));
const H1 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 420, after: 160 }, keepNext: true, children: [R(t, { size: 28, bold: true, color: NAVY })] }));
const split = (c, e) => `核心 ${c} · 资源 ${e}`;
function H2(num, name, c, e, intro) {
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 40 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }],
    children: [R(`${num}　${name}`, { size: 25, bold: true, color: NAVY }), R(`\t${split(c, e)}`, { size: 20, bold: true, color: BLUE })] }));
  D.push(new Paragraph({ spacing: { after: 140 }, keepNext: true, children: [R(intro, { size: 19, color: GRAY })] }));
}
function H3(num, name, en, c, e, note) {
  D.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: note ? 20 : 60 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W }],
    children: [R(`${num}　${name}`, { size: 22, bold: true, color: BLACK }), R(`　${en}`, { size: 18, color: GRAY }), R(`\t${split(c, e)}`, { size: 20, bold: true, color: BLUE })] }));
  if (note) D.push(new Paragraph({ spacing: { after: 80 }, keepNext: true, children: [R(note, { size: 18, color: GRAY })] }));
}

// ---- 表格 ----
const TH = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function cell(children, w, o = {}) {
  return new TableCell({
    children, width: { size: w, type: WidthType.DXA },
    borders: { top: TH, bottom: TH, left: NONE, right: NONE },
    margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.CENTER,
    ...(o.rowSpan ? { rowSpan: o.rowSpan } : {}), ...(o.colSpan ? { columnSpan: o.colSpan } : {}),
    ...(o.fill ? { shading: { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } } : {}),
  });
}
const tp = (t, o = {}) => new Paragraph({ children: runs(t, { size: o.size || 19, bold: o.bold, color: o.color }), spacing: { after: 20, line: 270 }, alignment: o.align || AlignmentType.LEFT });
const bp = t => new Paragraph({ children: runs(t, { size: 19 }), numbering: { reference: "cellb", level: 0 }, spacing: { after: 30, line: 280 } });
const C = AlignmentType.CENTER;

// =====================================================================================
// 数据：8 个分组、26 个岗位（研发 v0.4），加核心 / 资源划分
// =====================================================================================
const GROUPS = [
  { name: "项目与产品", intro: "负责项目在 RNP 侧的组织、进度、资源与验收，以及国民 AI 应用的产品定义。第一批到位。", roles: [
    { name: "项目经理", en: "Project Manager", c: 2, e: 0, one: "统筹 RNP 侧进度、资源与验收，对接讯飞与华为",
      note: "核心 2 人中 1 人兼知识转移与培训协调，对口 RNP 培训部门与 ESR",
      req: ["大型 IT 或基础设施项目的交付管理经验，能协调 RNP、讯飞、华为与部委四方", "熟悉政府项目的合同与验收流程", "英语可作为工作语言"],
      col: ["对口讯飞项目组：共同维护总进度与里程碑，参加例会与阶段评审", "组织 RNP 验收评审，由取得相应授权的人员签署验收文件"] },
    { name: "产品经理", en: "Product Manager", c: 2, e: 1, one: "定义国民 AI 应用的功能、用户范围与内容规则",
      note: "核心 2 人分管应用与 API 开放，资源 1 人协助",
      req: ["面向公众或政务的互联网产品经验", "能把巴西本地内容与使用场景转化为产品需求", "了解大模型应用的能力与边界"],
      col: ["对口讯飞产品团队：在已约定功能范围内确认用户范围、内容规则及配置要求；新增功能和路线图建议另行评估确认", "跟岗后接手应用配置与发布"] },
    { name: "PMO", en: "PMO", c: 0, e: 2, one: "维护项目计划、行动项、交付清单和文档，跟踪问题与决策落实",
      req: ["项目协调及计划跟踪经验", "熟悉行动项、问题台账和文档管理", "较强的组织与执行能力"],
      col: ["对口讯飞项目管理人员，维护会议纪要、行动项和交付清单", "跟踪双方待办事项，协调评审、验收和资料归档"] },
    { name: "基础设施专家", en: "Infrastructure Specialist", c: 2, e: 0, one: "协调平台运行所需的资源、网络、账号及环境条件，组织基础环境问题排查",
      note: "可由 RNP 现有架构与云团队承担",
      req: ["熟悉计算、存储和网络基础知识，了解 Linux、容器及 AI 平台部署环境", "能识别资源、权限、网络和软件依赖问题", "能协调基础设施及应用团队开展故障排查"],
      col: ["对口讯飞部署团队，确认资源、网络、账号及环境要求", "协调相关责任方解决基础环境问题，配合部署检查和联合定位"] },
  ] },
  { name: "语料与数据资源", intro: "组织葡语语料的获取、授权、生产与质检，是第一批到位的主力。", roles: [
    { name: "数据资源工程师", en: "Data Resource Engineer", c: 2, e: 2, one: "寻找数据来源，落实授权，组织语料获取",
      req: ["熟悉巴西公开与机构数据资源：政府、高校、媒体、科研数据库", "懂数据授权与许可，能审阅来源许可条款", "有数据采集与整理经验，葡语母语"],
      col: ["对口讯飞数据团队：按语料规范确认来源与配额，收集并登记授权材料，提交 RNP 法务确认许可条件", "基础课程与跟岗学习，掌握数据源、许可与平台处理流程"] },
    { name: "数据质检工程师", en: "Data Quality Engineer", c: 2, e: 2, one: "执行语料的计数、抽样、质量阈值与复检规则",
      note: "核心 2 人兼管标注人员与数据工程师",
      req: ["数据质量管理和抽样检查经验", "能理解数据格式、标注及质量规范", "能开展质量统计、缺陷分析和报告编制"],
      col: ["对口讯飞数据团队，确认抽样、质量判定和复检规则", "执行 RNP 侧质检和整改复核，配合讯飞约定的抽检及结果分析"] },
    { name: "数据工程师", en: "Data Engineer", c: 0, e: 4, one: "执行批量数据接收、清洗、转换和去重，维护处理流程及数据版本",
      req: ["熟练使用 Python、SQL 及数据处理工具", "批量清洗、转换、去重和流水线经验", "熟悉数据版本、元数据及来源追溯"],
      col: ["对口讯飞数据及平台团队，按分工执行数据处理任务并反馈异常", "接受平台操作指导，维护 RNP 负责的处理流程、配置和版本记录"] },
    { name: "标注员", en: "Annotator", c: 0, e: 20, one: "按规范完成文本标注、安全样本与语音转写",
      req: ["葡语母语，书面表达规范，细致稳定", "经培训后能按标注规范作业，转写岗需良好听写能力", "可由高校学生或合同人员承担，保留固定骨干"],
      col: ["使用双方确认的标注规范、样例与可用工具作业，配合讯飞在约定批次、抽样范围和复验轮次内开展质量复核"] },
  ] },
  { name: "葡语语言专家", intro: "定义巴西葡语的质量、标注与语音标准。", roles: [
    { name: "语言专家", en: "Portuguese Language Expert", c: 1, e: 5, one: "制定葡语质量、标注规则、安全标签与语音标准",
      note: "核心 1 人为组长，资源 5 人可来自高校",
      req: ["语言学、计算语言学或语音学背景", "熟悉巴西葡语的地区差异与当代用法", "有语料库或标注项目经验"],
      col: ["对口讯飞语言与语音团队：共同走查标注标准与安全标签，确认规则", "评价模型的葡语输出质量"] },
  ] },
  { name: "合规与审核", intro: "负责内容审核、数据及模型应用相关的法律合规，验收后承接内容安全审核平台；安全设备运行与技术告警由安全运维岗位负责。", roles: [
    { name: "审核专家", en: "Content Review Specialist", c: 1, e: 2, one: "维护风险标签、词库和审核策略，开展人工复核及误判分析",
      req: ["内容安全或信息审核经验", "了解巴西法律与社会语境下的敏感内容边界", "能定义并持续维护审核规则"],
      col: ["对口讯飞安全团队：在内容安全审核平台上共同配置策略、词库与样本", "验收后独立运营"] },
    { name: "法务专家", en: "Legal Counsel", c: 2, e: 0, one: "落实数据许可与合规，审核语料、模型与应用的法律风险",
      note: "可由 RNP 现有法务承担",
      req: ["数据保护法（LGPD）、知识产权与许可法律背景", "能审阅数据来源许可与开源许可条款", "有政府采购与合同经验者优先"],
      col: ["对口讯飞数据团队与商务：逐项确认数据来源许可，审阅权重、平台与应用的许可条款"] },
    { name: "审核负责人", en: "Review Lead", c: 1, e: 0, one: "组织审核规则确认、质量复核和争议升级，协调相关人员完成处置",
      req: ["合规或审核团队管理经验", "能定义流程、留痕与升级机制", "能与法务、业务和技术团队协同"],
      col: ["与讯飞约定审核升级与审批规则；作为审核平台验收与移交的 RNP 侧负责人"] },
  ] },
  { name: "业务与领域专家", intro: "为三个领域大模型确定业务场景并评估领域结果。第二批到位。", roles: [
    { name: "业务需求专家", en: "Business Requirements Lead", c: 3, e: 0, one: "为每个领域模型服务定义业务场景、输入与验收",
      note: "核心 3 人各负责一个领域，是领域模型在 RNP 侧的负责人",
      req: ["熟悉对应领域的业务流程", "能把业务问题转化为模型可交付的场景", "能组织领域机构参与"],
      col: ["对口讯飞领域模型顾问：每个 DLM 共同确定一个固定场景、输入数据及评测方法，区分 RNP 模型业务验证与讯飞辅助训练服务验收"] },
    { name: "领域需求专家", en: "Domain Requirements Specialist", c: 0, e: 2, one: "明确领域数据与预期结果，设定内容边界",
      req: ["领域数据与知识背景", "能定义评测集与预期输出，读懂评测报告"],
      col: ["对口讯飞领域模型顾问：共同定义评测方法与内容边界，参与评测练习"] },
    { name: "领域专家", en: "Domain Expert", c: 0, e: 2, one: "在领域内评价模型输出结果",
      req: ["领域内资深专业人员，可兼职或短期借用高校、机构专家", "能判断模型输出的正确性与适用性"],
      col: ["参加 RNP 组织的领域评测，出具专业评价意见，配合讯飞开展评测结果分析"] },
  ] },
  { name: "模型与测试工程", intro: "在模型训练中心参与国家模型训练与评测的技术交流和约定实操，并执行 RNP 负责的 DLM 训练与评测；不改变国家模型训练的合同责任分工。第二批到位，是知识转移的核心受训对象。", roles: [
    { name: "NLP/DLM 算法工程师", en: "NLP/DLM Algorithm Engineer", c: 6, e: 0, one: "执行 RNP 承担的 DLM 训练和实验，参与国家模型技术交流及效果分析",
      note: "其中 1 至 2 人可派驻中国参加 CPT 阶段",
      req: ["熟练使用 Python 和主流深度学习框架，理解微调、数据配比及实验设计", "能分析训练日志和模型效果，有实际训练经验者优先", "骨干应能修改训练脚本、适配新的开源模型"],
      col: ["对口讯飞模型团队：跟岗参与国家模型训练与评测，指导实操下自主提交作业", "领域模型由 RNP 训练，讯飞提供顾问支持"] },
    { name: "ASR/TTS 算法工程师", en: "ASR/TTS Algorithm Engineer", c: 0, e: 2, one: "开展语音数据检查、模型测试和问题定位，支持语音能力验证",
      req: ["音频处理和语音数据检查能力", "理解 ASR 或 TTS 原理及评测方法", "能开展模型调用、测试和问题复现"],
      col: ["对口讯飞语音模型团队，配合数据检查、服务调用及效果验证", "提供错误样例和复现信息，跟踪语音问题处理及回归结果"] },
    { name: "训练与推理工程师", en: "Training and Inference Engineer", c: 3, e: 1, one: "负责训练任务和推理验证的工程执行，处理资源、恢复及运行问题",
      req: ["熟悉 Linux、容器和分布式训练环境", "掌握资源配置、Checkpoint 恢复及任务重提", "能开展推理配置、日志分析和运行问题排查"],
      col: ["对口讯飞模型及平台团队，参与任务配置、资源检查和训练推理验证", "执行 RNP 负责的工程操作，提供日志及环境信息，配合复杂问题定位"] },
    { name: "模型评测工程师", en: "Model Evaluation Engineer", c: 2, e: 1, one: "管理独立评测集，执行效果、安全及回归测试并形成可复核结果",
      req: ["能编写或维护评测脚本", "熟悉独立评测集及数据隔离要求", "掌握基线比较、指标分析和结果复现"],
      col: ["对口讯飞模型评测人员，确认测试集、参数、指标和结果记录方式", "执行 RNP 侧评测及独立复核，提交差异样例并配合回归验证"] },
    { name: "平台/API/应用测试工程师", en: "Platform, API and Application Test Engineer", c: 0, e: 3, one: "验证平台、接口和多终端应用功能，跟踪缺陷并提供验收证据",
      req: ["软件测试和用例设计经验", "熟悉 API、认证及权限测试", "能开展多终端兼容、缺陷跟踪和回归验证"],
      col: ["对口讯飞测试及实施团队，确认产品、接口和应用验收用例", "执行 RNP 侧测试，提交可复现缺陷并确认整改结果"] },
  ] },
  { name: "应用与运营", intro: "承担国民 AI 应用与 RNP 系统的集成、用户审批、内容与日常运营；代码维护限于明确交付并授权修改的部分。第三批到位。", roles: [
    { name: "软件开发工程师", en: "Software Developer", c: 4, e: 11, one: "集成国民 AI 应用与 RNP 系统，开发本地内容与功能",
      note: "核心 4 人分别负责移动端、Web、后端与集成，项目后维护源码",
      req: ["移动端（Android、iOS）与 Web 前后端开发能力", "系统集成与 API 开发经验", "能基于文档化 API 开展集成"],
      col: ["对口讯飞应用团队：按确认的接口清单对接 RNP 系统；本地功能开发及源码使用按明确约定的范围执行", "跟岗后承接 RNP 侧集成、授权范围内的开发及发布"] },
    { name: "运营工程师", en: "Application Operations Engineer", c: 3, e: 7, one: "运营应用：配置、用户审批、API 凭证与用量分析",
      req: ["互联网产品或政务服务的运营经验", "能处理用户审批与权限管理", "能读用量数据并反馈问题"],
      col: ["对口讯飞应用与支持团队：辅助运维期并行，之后独立运营"] },
    { name: "架构师", en: "Solution Architect", c: 1, e: 1, one: "负责应用架构与 RNP 系统集成方案",
      req: ["企业级系统架构经验，熟悉云、容器、微服务与 API 网关", "了解模型服务调用与安全要求"],
      col: ["对口讯飞架构师：共同确定应用架构、集成点与安全方案，评审接口与部署方案"] },
  ] },
  { name: "平台运维", intro: "验收后运行国家 AI 能力中心平台。第三批到位。华为数据中心层面的运维培训人员不在此列。", roles: [
    { name: "运维负责人", en: "O&M Lead", c: 1, e: 0, one: "统筹运维流程、人员分工、故障升级和变更，保障服务有序运行",
      req: ["平台运维及 IT 服务管理经验", "熟悉事件、变更和故障升级流程", "了解容量、备份恢复及连续性管理"],
      col: ["对口讯飞运维及支持团队，确认服务分工、升级渠道和变更机制", "组织 RNP 运维人员承接操作，评审运行报告、遗留事项及移交结果"] },
    { name: "平台与应用运维工程师", en: "Platform and Application O&M Engineer", c: 3, e: 0, one: "执行平台、应用及模型服务的巡检、告警处理、配置备份和批准的维护操作",
      req: ["熟悉 Linux、容器及应用服务维护", "能分析日志、监控和常见告警", "掌握配置备份、恢复及补丁变更操作"],
      col: ["对口讯飞运维团队，参与巡检、配置检查、恢复演练和批准的维护操作", "承担 RNP 侧平台、应用及模型服务日常运维，按审批执行发布、配置和回退；提交故障信息并配合远程诊断及修复验证"] },
    { name: "安全运维工程师", en: "Security Operations Engineer", c: 2, e: 0, one: "运营安全产品，分析告警并跟踪漏洞整改和产品侧安全事件",
      note: "可由 RNP 现有安全团队承担",
      req: ["熟悉 WAF、堡垒机及审计类安全产品", "日志分析、漏洞扫描和整改跟踪能力", "有安全事件排查及协同处置经验"],
      col: ["对口讯飞安全产品及支持团队，参与部署配置、日志接入和功能验证", "承担 RNP 侧日常安全操作，提供事件信息并协同处理产品问题"] },
  ] },
];
GROUPS.forEach(g => { g.c = g.roles.reduce((a, r) => a + r.c, 0); g.e = g.roles.reduce((a, r) => a + r.e, 0); g.n = g.c + g.e; });
const CORE = GROUPS.reduce((a, g) => a + g.c, 0), EXT = GROUPS.reduce((a, g) => a + g.e, 0), TOTAL = CORE + EXT;
const NROLES = GROUPS.reduce((a, g) => a + g.roles.length, 0);

// =====================================================================================
// 封面
// =====================================================================================
D.push(new Paragraph({ spacing: { after: 100 }, children: [R("RNP 侧项目团队：核心团队与岗位画像", { size: 34, bold: true, color: RED })] }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [R("中文工作稿 · 供内部审定后翻译为英文发 Iara", { size: 23, bold: true, color: NAVY })] }));
D.push(new Paragraph({ spacing: { after: 320 }, children: [R("回应 Iara 2026 年 9 月 13 日 WhatsApp 请求　·　岗位与人数依据总部研发 v0.4　·　v0.5 · 2026 年 9 月 15 日", { size: 18, color: GRAY })] }));
P(`本稿回答 Iara 的两问：RNP 侧需要什么样的人，需要多少人。团队分两层。**核心团队 ${CORE} 人**是 RNP 的编制人员，承接知识转移，项目结束后原地转为常设运营团队；**项目资源 ${EXT} 人**通过科研津贴、固定期限合同和合作机构配置，按工作包投入和退出，由核心团队管理。两层合计 ${TOTAL} 人，${GROUPS.length} 个分组、${NROLES} 个岗位。`);
P("核心团队的划分依据是 RNP 在 9 月 9 日会上自己提出的两条：贯穿项目所有层面、走完华为与讯飞两级培训的人；项目结束后必须留住、把模型养住的人。", { after: 60 });

// =====================================================================================
// 一、总览
// =====================================================================================
H1("一、岗位与人数总览");
const W = [600, 1500, 1900, 620, 620, 4398];
const hdr = ["编号", "分组", "岗位", "核心", "资源", "岗位一句话"];
const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: hdr.map((h, i) => cell([tp(h, { bold: true, size: 19, color: NAVY, align: [0, 3, 4].includes(i) ? C : AlignmentType.LEFT })], W[i], { fill: "DCE6F1" })) })];
GROUPS.forEach((g, gi) => {
  g.roles.forEach((r, ri) => {
    const c = [];
    c.push(cell([tp(`${gi + 1}.${ri + 1}`, { align: C, color: GRAY })], W[0]));
    if (ri === 0) c.push(cell([tp(`**${gi + 1}　${g.name}**`, { color: NAVY }), tp(`核心 ${g.c} · 资源 ${g.e}`, { size: 17, color: BLUE })], W[1], { rowSpan: g.roles.length, fill: LIGHT }));
    c.push(cell([tp(`**${r.name}**`)], W[2]));
    c.push(cell([tp(r.c ? `**${r.c}**` : "–", { align: C, color: r.c ? NAVY : GRAY })], W[3]));
    c.push(cell([tp(r.e ? `${r.e}` : "–", { align: C, color: r.e ? BLUE : GRAY })], W[4]));
    c.push(cell([tp(r.one, { size: 18 })], W[5]));
    rows.push(new TableRow({ cantSplit: true, children: c }));
  });
});
rows.push(new TableRow({ cantSplit: true, children: [
  cell([tp("**合计**", { color: NAVY })], W[0] + W[1] + W[2], { colSpan: 3, fill: LIGHT }),
  cell([tp(`**${CORE}**`, { align: C, color: NAVY, size: 21 })], W[3], { fill: LIGHT }),
  cell([tp(`**${EXT}**`, { align: C, color: BLUE, size: 21 })], W[4], { fill: LIGHT }),
  cell([tp(`**核心团队 ${CORE} 人 + 项目资源 ${EXT} 人 = ${TOTAL} 人**`, { color: RED })], W[5], { fill: LIGHT }),
] }));
D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: W, rows }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

// =====================================================================================
// 二、核心团队与项目资源
// =====================================================================================
H1("二、核心团队与项目资源");
P(`**核心团队 ${CORE} 人**只放三类人：对外签字负责的，项目结束后要自己扛的，知识转移要一对一带的技术骨干。其中基础设施、法务、安全运维约 6 人可由 RNP 现有人员承担，需新增编制约 ${CORE - 6} 人。知识转移只覆盖核心团队，由核心团队里的组长再向项目资源人员传授。`);
P(`**项目资源 ${EXT} 人**按工作包配置：语料生产约 30 人，集中在第一批，语料完成即退出；领域与语言专家来自高校和合作机构，按阶段参与；应用开发与运营人员集中在第三批。每一群项目资源人员都由一个核心岗位管理：`);
const WM = [4200, 900, 4538];
const mrows = [new TableRow({ tableHeader: true, cantSplit: true, children: [
  cell([tp("项目资源人员", { bold: true, color: NAVY })], WM[0], { fill: "DCE6F1" }),
  cell([tp("人数", { bold: true, color: NAVY, align: C })], WM[1], { fill: "DCE6F1" }),
  cell([tp("由谁管理（核心团队岗位）", { bold: true, color: NAVY })], WM[2], { fill: "DCE6F1" }),
] })];
[
  ["标注员", 20, "数据质检工程师 2"],
  ["数据资源 2 · 数据质检 2 · 数据工程师 4", 8, "数据资源工程师 2 · 数据质检工程师 2"],
  ["语言专家", 5, "语言专家组长 1"],
  ["领域需求专家 2 · 领域专家 2", 4, "业务需求专家 3，每领域 1 人"],
  ["ASR/TTS 2 · 训练与推理 1 · 模型评测 1 · 测试 3", 7, "NLP/DLM 算法 6 · 训练与推理 3 · 模型评测 2"],
  ["软件开发 11 · 架构师 1", 12, "开发骨干 4 · 架构师 1"],
  ["运营工程师", 7, "运营工程师 3"],
  ["产品经理 1 · PMO 2", 3, "项目经理 2 · 产品经理 2"],
  ["审核专家", 2, "审核负责人 1 · 审核专家 1"],
].forEach(([a, n, b]) => mrows.push(new TableRow({ cantSplit: true, children: [
  cell([tp(a)], WM[0]), cell([tp(`${n}`, { align: C, color: BLUE, bold: true })], WM[1]), cell([tp(b)], WM[2]),
] })));
D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: WM, rows: mrows }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

// =====================================================================================
// 三、岗位画像
// =====================================================================================
H1("三、岗位画像");
P("每个岗位两行：**人员要求**说明要什么样的人；**与讯飞协作**说明这个岗位对口讯飞哪支团队、以什么方式一起工作。编号与总览表一致，标题右侧为核心与资源人数。", { after: 60 });
const WR = [1300, PAGE_W - 1300];
GROUPS.forEach((g, gi) => {
  H2(`${gi + 1}`, g.name, g.c, g.e, g.intro);
  g.roles.forEach((r, ri) => {
    H3(`${gi + 1}.${ri + 1}`, r.name, r.en, r.c, r.e, r.note);
    D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: WR, rows: [
      new TableRow({ cantSplit: true, children: [
        cell([tp("人员要求", { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }),
        cell(r.req.map(bp), WR[1], { valign: VerticalAlign.TOP }),
      ] }),
      new TableRow({ cantSplit: true, children: [
        cell([tp("与讯飞协作", { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }),
        cell(r.col.map(bp), WR[1], { valign: VerticalAlign.TOP }),
      ] }),
    ] }));
    D.push(new Paragraph({ spacing: { after: 10 }, children: [] }));
  });
});

// =====================================================================================
// 附：内部说明
// =====================================================================================
D.push(new Paragraph({ spacing: { before: 260, after: 60 }, keepNext: true, children: [R("附　内部说明（英文版不含）", { size: 19, bold: true, color: GRAY })] }));
[
  "1. Iara 原话：\"We are organizing RNP's side to receive the technology transfer and prepare to work together with you. Any important information about the skills we need on our side, the number of people, will be welcome.\"",
  `2. 分组、岗位、人数与措辞沿用总部研发 v0.4（26 个岗位、合计 111，三批 53 / 25 / 33）。本稿新增核心团队与项目资源的划分：核心 ${CORE}（含现有人员约 6），资源 ${EXT}。对外只讲"核心 ${CORE} 加资源 ${EXT}"，不用"峰值"一词。PPT 第 13 页的 30 / 20 / 30 仍须同步更新。`,
  "3. 相对 v0.4 的四处调整，均对应 RNP 会上原话：审核专家核心 1 人（其余 2 人为资源）；业务需求专家 3 人全部核心、每领域 1 人（副手要求 DLM 早启动、三领域与 Serpro 等机构协同）；项目经理 1 人兼知识转移与培训协调（Leandro 全程关注培训路径与 ESR）；NLP/DLM 核心 6 人恢复\"能修改训练脚本、适配开源模型\"一条（技术负责人 02:28 原话），并注明 1 至 2 人可赴华参加 CPT（02:30）。",
  "4. 需研发确认三点：项目资源人员使用源码与评测集的范围及保密约束（对应 v0.4 在 6.4、7.1 加的边界）；知识转移只覆盖核心 43 人；移交后常设团队规模对外如何表述（会上口头为 15 至 20 人，本稿核心团队含运营在内约 25 人）。",
].forEach(t => D.push(new Paragraph({ children: runs(t, { size: 17, color: GRAY }), spacing: { after: 30, line: 270 } })));

// =====================================================================================
const bullet = (ref, left, hanging) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });
const doc = new Document({
  creator: "iFLYTEK", title: "RNP 侧项目团队：核心团队与岗位画像",
  styles: { default: { document: { run: { font: F, size: 21, color: BLACK } } } },
  numbering: { config: [bullet("cellb", 260, 220)] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      R("RNP 侧项目团队：核心团队与岗位画像 · 中文工作稿 v0.5 · 第 ", { size: 16, color: GRAY }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 16, color: GRAY }), R(" 页", { size: 16, color: GRAY }) ] })] }) },
    children: D,
  }],
});
const OUT = "RNP侧团队岗位画像与人数建议-中文稿-v0.5-20260915.docx";
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes", "core", CORE, "ext", EXT, "total", TOTAL, "roles", NROLES); });
