// 科大讯飞致 RNP：RNP 侧项目团队配置建议（核心团队与岗位画像）—— 中文稿，供翻译为英文发 Iara
//   视角：科大讯飞向 RNP 提出建议；书面语；只建议岗位、人数、能力要求与协作方式，不涉及 RNP 内部的人事与经费操作
//   岗位与人数依据总部研发 v0.4（26 个岗位、111 人）；核心 43 / 扩展 68 的划分依据 9/15 内部讨论
//   node build_rnp_team_proposal_docx.js  ->  RNP侧项目团队配置建议-核心团队与岗位画像-中文稿-20260915.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
        ShadingType, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, TabStopType } = require("docx");

const F = "微软雅黑";
const RED = "C00000", NAVY = "1F3A5F", BLUE = "0057A8", GRAY = "6B6B6B", BLACK = "1D1D1A", LIGHT = "F3F6FA", LINE = "C9D3E0";
const PAGE_W = 9638;
const D = [];

const R = (t, o = {}) => new TextRun({ text: t, font: F, size: o.size || 21, bold: !!o.bold, color: o.color || BLACK });
function runs(t, o = {}) { return String(t).split("**").map((seg, i) => seg ? R(seg, { ...o, bold: o.bold || i % 2 === 1 }) : null).filter(Boolean); }
const P = (t, o = {}) => D.push(new Paragraph({ children: runs(t, o), spacing: { after: o.after == null ? 120 : o.after, line: 320 } }));
const H1 = t => D.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 420, after: 160 }, keepNext: true, children: [R(t, { size: 28, bold: true, color: NAVY })] }));
const split = (c, e) => `核心 ${c} · 扩展 ${e}`;
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
// 数据：8 个分组、26 个岗位
// =====================================================================================
const GROUPS = [
  { name: "项目与产品", intro: "负责项目在 RNP 侧的组织、进度、资源与验收，以及国民 AI 应用的产品定义。工作从项目第一阶段开始。", roles: [
    { name: "项目经理", en: "Project Manager", c: 2, e: 0, one: "统筹 RNP 侧进度、资源与验收，对接科大讯飞与华为",
      note: "建议其中 1 人兼任知识转移与培训的协调工作",
      req: ["大型 IT 或基础设施项目的交付管理经验，能协调 RNP、科大讯飞、华为与部委四方", "熟悉政府项目的合同与验收流程", "英语可作为工作语言"],
      col: ["对口科大讯飞项目组：共同维护总进度与里程碑，参加例会与阶段评审", "组织 RNP 侧验收评审，由授权人员签署验收文件"] },
    { name: "产品经理", en: "Product Manager", c: 2, e: 1, one: "定义国民 AI 应用的功能、用户范围与内容规则",
      req: ["面向公众或政务的互联网产品经验", "能把巴西本地内容与使用场景转化为产品需求", "了解大模型应用的能力与边界"],
      col: ["对口科大讯飞产品团队：在已约定的功能范围内确认用户范围、内容规则及配置要求；新增功能与路线图建议另行评估确认", "跟岗后接手应用配置与发布"] },
    { name: "PMO", en: "PMO", c: 0, e: 2, one: "维护项目计划、行动项、交付清单和文档，跟踪问题与决策落实",
      req: ["项目协调及计划跟踪经验", "熟悉行动项、问题台账和文档管理", "较强的组织与执行能力"],
      col: ["对口科大讯飞项目管理人员，维护会议纪要、行动项和交付清单", "跟踪双方待办事项，协调评审、验收和资料归档"] },
    { name: "基础设施专家", en: "Infrastructure Specialist", c: 2, e: 0, one: "协调平台运行所需的资源、网络、账号及环境条件，组织基础环境问题排查",
      req: ["熟悉计算、存储和网络基础知识，了解 Linux、容器及 AI 平台部署环境", "能识别资源、权限、网络和软件依赖问题", "能协调基础设施及应用团队开展故障排查"],
      col: ["对口科大讯飞部署团队，确认资源、网络、账号及环境要求", "协调相关责任方解决基础环境问题，配合部署检查和联合定位"] },
  ] },
  { name: "语料与数据资源", intro: "组织葡语语料的获取、授权、生产与质检，是项目第一阶段的主要力量。", roles: [
    { name: "数据资源工程师", en: "Data Resource Engineer", c: 2, e: 2, one: "寻找数据来源，落实授权，组织语料获取",
      req: ["熟悉巴西公开与机构数据资源：政府、高校、媒体、科研数据库", "懂数据授权与许可，能审阅来源许可条款", "有数据采集与整理经验，葡语母语"],
      col: ["对口科大讯飞数据团队：按语料规范确认来源与配额，收集并登记授权材料，完成许可条件的确认", "通过基础课程与跟岗学习，掌握数据源、许可与平台处理流程"] },
    { name: "数据质检工程师", en: "Data Quality Engineer", c: 2, e: 2, one: "执行语料的计数、抽样、质量阈值与复检规则",
      req: ["数据质量管理和抽样检查经验", "能理解数据格式、标注及质量规范", "能开展质量统计、缺陷分析和报告编制"],
      col: ["对口科大讯飞数据团队，确认抽样、质量判定和复检规则", "执行 RNP 侧质检和整改复核，配合科大讯飞按约定开展抽检及结果分析"] },
    { name: "数据工程师", en: "Data Engineer", c: 0, e: 4, one: "执行批量数据接收、清洗、转换和去重，维护处理流程及数据版本",
      req: ["熟练使用 Python、SQL 及数据处理工具", "批量清洗、转换、去重和流水线经验", "熟悉数据版本、元数据及来源追溯"],
      col: ["对口科大讯飞数据及平台团队，按分工执行数据处理任务并反馈异常", "接受平台操作指导，维护 RNP 负责的处理流程、配置和版本记录"] },
    { name: "标注员", en: "Annotator", c: 0, e: 20, one: "按规范完成文本标注、安全样本与语音转写",
      req: ["葡语母语，书面表达规范，细致稳定", "经培训后能按标注规范作业，转写岗位需良好的听写能力"],
      col: ["使用双方确认的标注规范、样例与工具作业，配合科大讯飞在约定批次、抽样范围和复验轮次内开展质量复核"] },
  ] },
  { name: "葡语语言专家", intro: "定义巴西葡语的质量、标注与语音标准。建议核心 1 人牵头。", roles: [
    { name: "语言专家", en: "Portuguese Language Expert", c: 1, e: 5, one: "制定葡语质量、标注规则、安全标签与语音标准",
      req: ["语言学、计算语言学或语音学背景", "熟悉巴西葡语的地区差异与当代用法", "有语料库或标注项目经验"],
      col: ["对口科大讯飞语言与语音团队：共同走查标注标准与安全标签，确认规则", "评价模型的葡语输出质量"] },
  ] },
  { name: "合规与审核", intro: "负责内容审核以及数据、模型与应用相关的法律合规，验收后承接内容安全审核平台；安全设备的运行与技术告警由安全运维岗位负责。", roles: [
    { name: "审核专家", en: "Content Review Specialist", c: 1, e: 2, one: "维护风险标签、词库和审核策略，开展人工复核及误判分析",
      req: ["内容安全或信息审核经验", "了解巴西法律与社会语境下的敏感内容边界", "能定义并持续维护审核规则"],
      col: ["对口科大讯飞安全团队：在内容安全审核平台上共同配置策略、词库与样本", "验收后独立运营"] },
    { name: "法务专家", en: "Legal Counsel", c: 2, e: 0, one: "落实数据许可与合规，审核语料、模型与应用的法律风险",
      req: ["数据保护法（LGPD）、知识产权与许可法律背景", "能审阅数据来源许可与开源许可条款", "有政府采购与合同经验者优先"],
      col: ["对口科大讯飞数据团队与商务团队：逐项确认数据来源许可，审阅权重、平台与应用的许可条款"] },
    { name: "审核负责人", en: "Review Lead", c: 1, e: 0, one: "组织审核规则确认、质量复核和争议升级，协调相关人员完成处置",
      req: ["合规或审核团队管理经验", "能定义流程、留痕与升级机制", "能与法务、业务和技术团队协同"],
      col: ["与科大讯飞约定审核升级与审批规则；作为审核平台验收与移交的 RNP 侧负责人"] },
  ] },
  { name: "业务与领域专家", intro: "为三个领域大模型确定业务场景并评估领域结果。工作从项目第二阶段开始。", roles: [
    { name: "业务需求专家", en: "Business Requirements Lead", c: 3, e: 0, one: "为每个领域模型服务定义业务场景、输入与验收",
      note: "建议 3 人各负责一个领域，作为该领域模型在 RNP 侧的负责人",
      req: ["熟悉对应领域的业务流程", "能把业务问题转化为模型可交付的场景", "能组织领域机构参与"],
      col: ["对口科大讯飞领域模型顾问：每个领域模型共同确定一个固定场景、输入数据及评测方法，区分 RNP 的模型业务验证与科大讯飞辅助训练服务的验收"] },
    { name: "领域需求专家", en: "Domain Requirements Specialist", c: 0, e: 2, one: "明确领域数据与预期结果，设定内容边界",
      req: ["领域数据与知识背景", "能定义评测集与预期输出，读懂评测报告"],
      col: ["对口科大讯飞领域模型顾问：共同定义评测方法与内容边界，参与评测练习"] },
    { name: "领域专家", en: "Domain Expert", c: 0, e: 2, one: "在领域内评价模型输出结果",
      req: ["领域内资深专业人员", "能判断模型输出的正确性与适用性"],
      col: ["参加 RNP 组织的领域评测，出具专业评价意见，配合科大讯飞开展评测结果分析"] },
  ] },
  { name: "模型与测试工程", intro: "在模型训练中心通过技术交流和约定的实操参与国家模型的训练与评测，并承担 RNP 负责的领域模型训练与评测；国家模型训练的合同责任分工不变。工作从项目第二阶段开始，是知识转移的核心受训对象。", roles: [
    { name: "NLP/DLM 算法工程师", en: "NLP/DLM Algorithm Engineer", c: 6, e: 0, one: "执行 RNP 承担的领域模型训练和实验，参与国家模型技术交流及效果分析",
      note: "建议其中 1 至 2 人参加在中国进行的持续预训练阶段工作",
      req: ["熟练使用 Python 和主流深度学习框架，理解微调、数据配比及实验设计", "能分析训练日志和模型效果，有实际训练经验者优先", "骨干应能修改训练脚本、适配新的开源模型"],
      col: ["对口科大讯飞模型团队：跟岗参与国家模型训练与评测，在指导下自主提交训练作业", "领域模型由 RNP 训练，科大讯飞提供顾问支持"] },
    { name: "ASR/TTS 算法工程师", en: "ASR/TTS Algorithm Engineer", c: 0, e: 2, one: "开展语音数据检查、模型测试和问题定位，支持语音能力验证",
      req: ["音频处理和语音数据检查能力", "理解 ASR 或 TTS 原理及评测方法", "能开展模型调用、测试和问题复现"],
      col: ["对口科大讯飞语音模型团队，配合数据检查、服务调用及效果验证", "提供错误样例和复现信息，跟踪语音问题处理及回归结果"] },
    { name: "训练与推理工程师", en: "Training and Inference Engineer", c: 3, e: 1, one: "负责训练任务和推理验证的工程执行，处理资源、恢复及运行问题",
      req: ["熟悉 Linux、容器和分布式训练环境", "掌握资源配置、Checkpoint 恢复及任务重提", "能开展推理配置、日志分析和运行问题排查"],
      col: ["对口科大讯飞模型及平台团队，参与任务配置、资源检查和训练推理验证", "执行 RNP 负责的工程操作，提供日志及环境信息，配合复杂问题定位"] },
    { name: "模型评测工程师", en: "Model Evaluation Engineer", c: 2, e: 1, one: "管理独立评测集，执行效果、安全及回归测试并形成可复核结果",
      req: ["能编写或维护评测脚本", "熟悉独立评测集及数据隔离要求", "掌握基线比较、指标分析和结果复现"],
      col: ["对口科大讯飞模型评测人员，确认测试集、参数、指标和结果记录方式", "执行 RNP 侧评测及独立复核，提交差异样例并配合回归验证"] },
    { name: "平台/API/应用测试工程师", en: "Platform, API and Application Test Engineer", c: 0, e: 3, one: "验证平台、接口和多终端应用功能，跟踪缺陷并提供验收证据",
      req: ["软件测试和用例设计经验", "熟悉 API、认证及权限测试", "能开展多终端兼容、缺陷跟踪和回归验证"],
      col: ["对口科大讯飞测试及实施团队，确认产品、接口和应用的验收用例", "执行 RNP 侧测试，提交可复现缺陷并确认整改结果"] },
  ] },
  { name: "应用与运营", intro: "承担国民 AI 应用与 RNP 系统的集成、用户审批、内容与日常运营，并维护已交付且授权修改的代码。工作从项目第三阶段开始。", roles: [
    { name: "软件开发工程师", en: "Software Developer", c: 4, e: 11, one: "集成国民 AI 应用与 RNP 系统，开发本地内容与功能",
      req: ["移动端（Android、iOS）与 Web 前后端开发能力", "系统集成与 API 开发经验", "能基于文档化的 API 开展集成"],
      col: ["对口科大讯飞应用团队：按确认的接口清单对接 RNP 系统；本地功能开发及源码使用按双方约定的范围执行", "跟岗后承接 RNP 侧集成以及授权范围内的开发与发布"] },
    { name: "运营工程师", en: "Application Operations Engineer", c: 3, e: 7, one: "运营应用：配置、用户审批、API 凭证与用量分析",
      req: ["互联网产品或政务服务的运营经验", "能处理用户审批与权限管理", "能读用量数据并反馈问题"],
      col: ["对口科大讯飞应用与支持团队：辅助运维期内并行运营，之后独立运营"] },
    { name: "架构师", en: "Solution Architect", c: 1, e: 1, one: "负责应用架构与 RNP 系统集成方案",
      req: ["企业级系统架构经验，熟悉云、容器、微服务与 API 网关", "了解模型服务调用与安全要求"],
      col: ["对口科大讯飞架构师：共同确定应用架构、集成点与安全方案，评审接口与部署方案"] },
  ] },
  { name: "平台运维", intro: "验收后运行国家 AI 能力中心平台。工作从项目第三阶段开始，全部为核心团队。华为数据中心层面的运维人员不在本建议范围内。", roles: [
    { name: "运维负责人", en: "O&M Lead", c: 1, e: 0, one: "统筹运维流程、人员分工、故障升级和变更，保障服务有序运行",
      req: ["平台运维及 IT 服务管理经验", "熟悉事件、变更和故障升级流程", "了解容量、备份恢复及连续性管理"],
      col: ["对口科大讯飞运维及支持团队，确认服务分工、升级渠道和变更机制", "组织 RNP 运维人员承接操作，评审运行报告、遗留事项及移交结果"] },
    { name: "平台与应用运维工程师", en: "Platform and Application O&M Engineer", c: 3, e: 0, one: "执行平台、应用及模型服务的巡检、告警处理、配置备份和经批准的维护操作",
      req: ["熟悉 Linux、容器及应用服务维护", "能分析日志、监控和常见告警", "掌握配置备份、恢复及补丁变更操作"],
      col: ["对口科大讯飞运维团队，参与巡检、配置检查、恢复演练和经批准的维护操作", "承担 RNP 侧平台、应用及模型服务的日常运维，按审批执行发布、配置和回退；提交故障信息并配合远程诊断及修复验证"] },
    { name: "安全运维工程师", en: "Security Operations Engineer", c: 2, e: 0, one: "运营安全产品，分析告警并跟踪漏洞整改和产品侧安全事件",
      req: ["熟悉 WAF、堡垒机及审计类安全产品", "日志分析、漏洞扫描和整改跟踪能力", "有安全事件排查及协同处置经验"],
      col: ["对口科大讯飞安全产品及支持团队，参与部署配置、日志接入和功能验证", "承担 RNP 侧日常安全操作，提供事件信息并协同处理产品问题"] },
  ] },
];
GROUPS.forEach(g => { g.c = g.roles.reduce((a, r) => a + r.c, 0); g.e = g.roles.reduce((a, r) => a + r.e, 0); g.n = g.c + g.e; });
const CORE = GROUPS.reduce((a, g) => a + g.c, 0), EXT = GROUPS.reduce((a, g) => a + g.e, 0), TOTAL = CORE + EXT;
const NROLES = GROUPS.reduce((a, g) => a + g.roles.length, 0);

// =====================================================================================
// 封面与引言
// =====================================================================================
D.push(new Paragraph({ spacing: { after: 100 }, children: [R("RNP 侧项目团队配置建议", { size: 34, bold: true, color: RED })] }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [R("核心团队与岗位画像", { size: 26, bold: true, color: NAVY })] }));
D.push(new Paragraph({ spacing: { after: 320 }, children: [R("科大讯飞提交 RNP 参考　·　2026 年 9 月 15 日　·　中文稿", { size: 18, color: GRAY })] }));

P("应 Iara 女士 9 月 13 日的请求，科大讯飞就 RNP 侧承接技术转移、并与科大讯飞共同实施本项目所需的团队提出以下建议：需要哪些岗位、各岗位的能力要求、建议人数，以及各岗位与科大讯飞的协作方式。");
P("本建议依据 9 月 9 日技术交流的内容和当前的工作范围划分测算，供 RNP 参考。科大讯飞愿在 RNP 确定团队方案后共同评审，并据此细化知识转移的安排。", { after: 60 });

// =====================================================================================
// 一、团队构成建议
// =====================================================================================
H1("一、团队构成建议");
P(`建议 RNP 侧团队分为两层，合计 ${TOTAL} 人、${GROUPS.length} 个分组、${NROLES} 个岗位。`);
P(`**核心团队（Equipe núcleo）${CORE} 人。** 贯穿项目全程的长期团队，是知识转移的主要对象，项目结束后继续运行平台、模型与应用。包括三类岗位：对项目成果负责和签署验收的岗位；项目结束后独立运行平台、模型与应用的岗位；承接知识转移并向其他人员传授的技术骨干。`);
P(`**扩展团队（Equipe complementar）${EXT} 人。** 承担语料生产、领域与语言专业支持、应用本地化开发与运营等阶段性、工作量集中的任务，随相应的工作阶段参与。建议每一类扩展团队岗位对应一个核心团队岗位作为对接负责人，以保证工作的连续性和成果的沉淀。`);
P("人数为按工作内容测算的建议配置，同一人可以承担多个岗位的工作，可根据各阶段的实际工作量调整。", { after: 60 });

// =====================================================================================
// 二、总览
// =====================================================================================
H1("二、岗位与人数总览");
const W = [600, 1500, 1900, 620, 620, 4398];
const hdr = ["编号", "分组", "岗位", "核心", "扩展", "岗位一句话"];
const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: hdr.map((h, i) => cell([tp(h, { bold: true, size: 19, color: NAVY, align: [0, 3, 4].includes(i) ? C : AlignmentType.LEFT })], W[i], { fill: "DCE6F1" })) })];
GROUPS.forEach((g, gi) => {
  g.roles.forEach((r, ri) => {
    const c = [];
    c.push(cell([tp(`${gi + 1}.${ri + 1}`, { align: C, color: GRAY })], W[0]));
    if (ri === 0) c.push(cell([tp(`**${gi + 1}　${g.name}**`, { color: NAVY }), tp(`核心 ${g.c} · 扩展 ${g.e}`, { size: 17, color: BLUE })], W[1], { rowSpan: g.roles.length, fill: LIGHT }));
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
  cell([tp(`**核心团队 ${CORE} 人 + 扩展团队 ${EXT} 人 = ${TOTAL} 人**`, { color: RED })], W[5], { fill: LIGHT }),
] }));
D.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: W, rows }));
D.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

// =====================================================================================
// 三、岗位画像
// =====================================================================================
H1("三、岗位画像");
P("每个岗位两部分：**人员要求**说明该岗位需要什么样的人；**协作方式**说明该岗位在项目中对口科大讯飞的哪支团队、以什么方式共同工作。编号与总览表一致，标题右侧为建议的核心与扩展人数。", { after: 60 });
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
        cell([tp("协作方式", { bold: true, size: 18, color: NAVY })], WR[0], { fill: LIGHT, valign: VerticalAlign.TOP }),
        cell(r.col.map(bp), WR[1], { valign: VerticalAlign.TOP }),
      ] }),
    ] }));
    D.push(new Paragraph({ spacing: { after: 10 }, children: [] }));
  });
});

// =====================================================================================
const bullet = (ref, left, hanging) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });
const doc = new Document({
  creator: "iFLYTEK", title: "RNP 侧项目团队配置建议：核心团队与岗位画像",
  styles: { default: { document: { run: { font: F, size: 21, color: BLACK } } } },
  numbering: { config: [bullet("cellb", 260, 220)] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      R("科大讯飞 · RNP 侧项目团队配置建议 · 2026 年 9 月 · 第 ", { size: 16, color: GRAY }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 16, color: GRAY }), R(" 页", { size: 16, color: GRAY }) ] })] }) },
    children: D,
  }],
});
const OUT = "RNP侧项目团队配置建议-核心团队与岗位画像-中文稿-20260915.docx";
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes", "core", CORE, "ext", EXT, "total", TOTAL, "roles", NROLES); });
