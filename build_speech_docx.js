// 演讲脚本（RNP 技术交流会议版，16 页）→ docx
// 标记：【详讲】红 = 回应 RNP 关注点，逐句讲；【略讲】灰 = 一句带过；【备查】蓝 = 问到才翻。
//   node build_speech_docx.js  ->  演讲脚本-会议版.docx
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, BorderStyle, LevelFormat, PageBreak,
} = require("docx");

const FONT = "微软雅黑";
const RED = "C00000", GRAY = "7F7F7F", BLUE = "0070C0", BLACK = "222222", MUTE = "666666";
const TAG = { 详讲: RED, 略讲: GRAY, 备查: BLUE };

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: o.size || 22, bold: !!o.bold, color: o.color || BLACK, italics: !!o.italics });
const p = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [run(children)], spacing: { after: o.after ?? 120, line: 340 }, alignment: o.align, indent: o.indent, shading: o.shading, border: o.border });
const bullet = (children) => new Paragraph({ children: Array.isArray(children) ? children : [run(children)], numbering: { reference: "b", level: 0 }, spacing: { after: 80, line: 340 } });
const h1 = (t) => new Paragraph({ text: t, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 } });

// 页标题：第 N 页 标题 【详讲 3 分】 关注点 ①②
function pageHead(n, title, tag, minutes, concerns) {
  const kids = [run(`${n}  ${title}`, { bold: true, size: 26, color: BLACK }), run("　"),
    run(`【${tag}${minutes ? " · " + minutes : ""}】`, { bold: true, size: 22, color: TAG[tag] })];
  if (concerns) kids.push(run(`　回应关注点 ${concerns}`, { size: 20, color: MUTE }));
  return new Paragraph({ children: kids, heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 120 } });
}
// 段落标签
const L = (label, text, color) => p([run(`【${label}】`, { bold: true, color: color || RED }), run(text)]);
const detail = (label, text) => L(label, text, RED);   // 详讲段
const brief = (label, text) => L(label, text, GRAY);   // 略讲段
const qa = (q, a) => p([run("问：", { bold: true, color: BLUE }), run(q, { bold: true }), run("　答：", { bold: true, color: BLUE }), run(a)]);
const say = (text) => p([run(text)]);

function cell(t, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(t, { bold: !!o.bold, color: o.color || BLACK, size: 20 })], spacing: { after: 0, line: 300 } })],
  });
}
function table(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: "F2F2F2" })) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => {
        const tagColor = TAG[c] ; return cell(c, widths[i], tagColor ? { bold: true, color: tagColor } : {});
      }) })),
    ],
  });
}

const box = (lines, fill) => lines.map((l, i) => new Paragraph({
  children: Array.isArray(l) ? l : [run(l)], spacing: { after: i === lines.length - 1 ? 200 : 60, line: 320 },
  shading: { type: ShadingType.CLEAR, fill, color: "auto" }, indent: { left: 200, right: 200 },
}));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, color: RED, font: FONT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, color: BLACK, font: FONT }, paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    children: [
      new Paragraph({ children: [run("RNP 技术交流演讲脚本", { bold: true, size: 40, color: RED })], spacing: { after: 80 } }),
      p([run("配套材料：《RNP-讯飞交流-会议版》16 页　·　讨论稿，数字以 SOW 为准", { color: MUTE, size: 20 })], { after: 240 }),

      h1("一、怎么用这份脚本"),
      p([run("RNP 会前确认的三个关注点，脚本里用 ①②③ 标出每页回应哪一个：", { bold: true })]),
      bullet([run("① 交付范围：", { bold: true }), run("葡语大模型、国家 APP、3 个行业模型、知识转移，讯飞到底交付什么。")]),
      bullet([run("② 责任、资源与边界：", { bold: true }), run("双方各负什么责、各出什么资源、SOW 的边界在哪。")]),
      bullet([run("③ 知识转移与培训方式：", { bold: true }), run("怎么教、教到什么程度，好让 RNP 编自己的 Capacity Program 和人力计划。")]),
      p([run("三种标记：", { bold: true })], { after: 60 }),
      bullet([run("【详讲】", { bold: true, color: RED }), run("直接回应关注点的内容，逐句讲，每页 3 到 4 分钟。")]),
      bullet([run("【略讲】", { bold: true, color: GRAY }), run("背景或讯飞自己的工作，一句带过，不超过 1 分钟。")]),
      bullet([run("【备查】", { bold: true, color: BLUE }), run("详情页，不主动讲，被问到才翻，脚本写成问答。")]),
      p([run("讲法上三条：", { bold: true })], { after: 60 }),
      bullet("用“你们出什么、我们出什么、对你们意味着什么”三句话讲每一页，不讲技术过程。"),
      bullet("数字只说六个：3000 亿参数、本地 50B 加通用 300B 语料、3 个行业模型、1000 万用户与 500 并发、280P 与 560P、6 个月一版。其余等对方问。"),
      bullet("蓝色内容一律说“按 SOW 确认”，不现场给数字，不现场承诺。"),

      h1("二、时间分配"),
      table(["页", "内容", "标记", "60 分钟版", "20 分钟版", "关注点"], [
        ["开场", "三个关注点即议程", "详讲", "1", "1", "①②③"],
        ["1", "封面", "略讲", "0.5", "—", ""],
        ["2", "一页看全", "详讲", "4", "4", "①"],
        ["3", "国家大模型", "详讲", "3", "—", "①②"],
        ["4、5", "详情：规格、语料、评测、数据主权", "备查", "问到再翻", "—", ""],
        ["6", "国家 APP", "略讲", "2（含演示）", "2（只放演示）", "①②"],
        ["7", "详情：服务、容量、接口", "备查", "问到再翻", "—", ""],
        ["8", "行业模型 DLM", "详讲", "3", "3", "①②③"],
        ["9", "详情：DLM 实施条件", "备查", "问到再翻", "—", ""],
        ["10", "整体路径", "略讲", "2", "—", "②"],
        ["11", "详情：交接、许可、运营", "备查", "问到再翻，要备好", "—", "②"],
        ["12", "RNP 需要投入什么", "详讲", "4", "3", "②③"],
        ["13", "详情：参与矩阵", "略讲", "0.5", "指一下", "③"],
        ["14", "RNP 能带走什么", "详讲", "3", "2", "③"],
        ["15", "详情：学习路径", "备查", "问到再翻", "—", "③"],
        ["16", "SOW 边界与确认", "详讲", "3", "2", "②"],
        ["收尾", "两个会后动作", "详讲", "1", "1", ""],
        ["讨论", "", "", "约 25", "约 3", ""],
      ], [900, 3300, 900, 1500, 1500, 1100]),
      p("", { after: 120 }),

      h1("三、逐页脚本"),

      pageHead("开场", "会前先说三件事", "详讲", "1 分钟", "①②③"),
      say("Yara，各位，谢谢安排今天的会。会前我们和你们确认过，你们最想知道三件事：第一，讯飞到底交付什么；第二，双方各自负责什么、要投入什么、SOW 的边界在哪；第三，我们的知识转移怎么做，好让你们编自己的 Capacity Program 和人力计划。今天就讲这三件事，材料也是按这三件事排的。"),
      say("材料已经提前发给你们。带“详情”标签的页今天不逐页讲，你们问到我再翻。我希望今天结束时，你们对讯飞的范围和分工没有疑问，并且拿到编制自己计划所需要的全部输入。"),

      pageHead("第 1 页", "封面", "略讲", "30 秒"),
      brief("一句带过", "这是讨论稿。黑字都和我们此前给你们的四份书面答复一致，4 月一份、5 月两份、8 月一份。蓝字是示例或待核，最终以 SOW 为准。有一处变化我先说：行业模型从 8 月答复里的 1 个变成 3 个，讲到那页时我会解释。"),

      pageHead("第 2 页", "一页看全", "详讲", "4 分钟", "①"),
      detail("先说结论", "讯飞交付三样东西，加一套能力。这一页就是全部范围。"),
      detail("第一样", "国家大模型。约 3000 亿参数，以讯飞星火为底座训练，加上巴西葡语的语音识别和语音合成。训练用的文本数据，我们叫语料：你们构建 50B 的本地葡语语料，我们免费提供 300B 以上的通用语料。约 6 个月出一个正式版本。我们主导训练，你们采集语料、参与训练和验收。"),
      detail("第二样", "国家 APP。以讯飞已经上线的星火 APP 为基础，做巴西本地化。面向公众，按约 1000 万用户、500 路并发规划。我们出产品和平台，你们定义场景、提供接口，上线后由你们验收和运营。"),
      detail("第三样", "3 个行业模型。每个解决 1 个具体场景的问题。方向由你们和 MCTI 提出，训练由你们自己做，我们做评估、给路线、做指导。"),
      detail("底下两行", "这些都跑在三大中心上。三大中心是三套软件平台，不是三座楼：语料库构建中心、模型训练中心、AI 应用创新中心。算力第一阶段约 280P，第二阶段合计约 560P。知识转移不是最后一门课，是你们的人从第一天就跟着做。"),
      detail("对你们意味着", "这一页可以直接拿去向上汇报。后面每样东西各展开一页。"),

      pageHead("第 3 页", "国家大模型：怎么建、谁来做", "详讲", "3 分钟", "①②"),
      detail("先说结论", "模型由我们主导训练，语料由双方各出一半：你们出本地的，我们出通用的。"),
      detail("接你们的问题", "你们 5 月问过模型多大、语料多少、谁负责。模型约 3000 亿参数，星火底座。语料 5 月 4 日的答复写过 500B 的总目标，5 月 8 日已经拆成两块：你们本地 50B，我们通用 300B 以上。今天按这个口径。"),
      detail("四步怎么走", "第一步你们采集本地语料，我们把原始语料加工成能直接训练的语料库。第二步在星火底座上训练、多轮迭代。第三步评测，用巴西本土的考试和基准，加国际通用基准，评测方案我们准备、你们审核。第四步交付，模型权重归巴西，加基础模型的源码、可导出的格式和一个 API 平台。"),
      detail("对你们意味着", "三件事。一，采集 50B 本地语料，并办好版权和 LGPD 授权，这是第一阶段最重的一项工作。二，你们的工程师跟着训练过程学。三，评测方案你们有审核权。"),
      brief("不主动讲", "参数、算力、基准名称、数据主权这些在后面两页，问到再翻。"),

      pageHead("第 4 页", "详情：规格与算力", "备查"),
      qa("底座是不是闭源的星火？", "是星火基础模型。星火基础模型的权重在双方约定范围内受限开放给巴西，参考 Qwen、Llama 的开源范围。这是 4 月的书面答复。"),
      qa("算力够不够、怎么分？", "训练 210P、语料构建 17.5P、语音 35P，都在第一阶段的 280P 里。全部算力统一池化调度，不是固定分区。"),
      qa("能不能在英伟达上跑？", "可以 ONNX 导出在英伟达上运行，性能低于昇腾。这是 5 月的答复。"),
      qa("包不包括图像、视频？", "VLM 不在范围。"),

      pageHead("第 5 页", "详情：语料、评测与数据主权", "备查"),
      qa("数据会不会出境？", "不会。采集、训练、微调、推理、运维全部在巴西境内，遵守 LGPD，禁止未授权跨境。你们的本土数据集你们完全控制；我们提供的通用语料是可用不可见。"),
      qa("用什么基准评测？", "本土候选 ENEM、BlueX、HateBR、OAB、ASSIN2；国际 M-MMLU、MGSM、Xwinogrande、Belebele、FLORES。达标线写在评测方案里，按 SOW。"),
      qa("语料任务量多大？", "有监督数据 20 万条以上，安全对齐 2 万条以上，语音识别 1 万小时以上，语音合成 100 小时以上。你们采集，我们出规范、做加工和质检。"),
      qa("合规怎么组织？", "建议你们指派一名 LGPD 合规专员和我们对接，这是 4 月的建议。"),

      pageHead("第 6 页", "国家 APP：给谁用、怎么建", "略讲", "2 分钟", "①②"),
      brief("一句带过", "你们 5 月问 APP 交付什么。答案是：讯飞星火 APP 已经在中国上线，问答、语音、文档理解、写作这些能力是现成的。本项目做的是巴西本地化。"),
      brief("放演示", "看 90 秒演示。说明这是基线产品的画面，不是本项目版本。"),
      detail("需要你们的", "本地化的第一步需要你们主导：和巴西伙伴一起定义场景和功能。之后你们提供本地系统的接口和测试数据，上线后由你们运营。接口清单按 SOW。"),

      pageHead("第 7 页", "详情：服务、容量与接口", "备查"),
      qa("能不能通过 API 调模型？", "可以。MaaS 平台提供 OpenAI 兼容的 REST API、Python、JS、Java 的 SDK、OpenAPI 文档，支持 Docker 和 K8s 部署，带认证、限流、监控、审计。"),
      qa("能不能接开源模型？", "可以。AI 能力中心已兼容 DeepSeek、Qwen 等，另含两次推理迁移适配。"),
      qa("其他本地应用谁做？", "可以由巴西本地企业和讯飞联合开发。性能指标按 SOW。"),

      pageHead("第 8 页", "行业模型 DLM：选什么领域、怎么共同开发", "详讲", "3 分钟", "①②③"),
      detail("先说结论", "3 个行业模型。方向你们定，训练你们做，我们辅助。这一项和你们的关系最直接。"),
      detail("先说变化", "8 月答复里是 1 个，最新方案是 3 个，每个面向 1 个边界清晰的固定场景。实施条件和分工不变，算力和数据规模按 3 个重估后随 SOW 确认。"),
      detail("四步怎么走", "第一步你们和 MCTI 确定行业、场景、目标和使用边界。第二步你们提供或协调授权数据、标注和独立评测数据，基座在 700 亿参数以内，星火或 DeepSeek、Qwen 都可以。第三步你们自主训练，我们辅助并持续咨询，这是我们建议的模式，也是知识转移最实在的一块。第四步基于双方确认的独立评测集验收，实施前签《DLM 场景及技术方案确认书》。"),
      detail("对你们意味着", "你们要确定场景、出数据和专家、自己训练、组织验收。讯飞在政务、法律、医疗、教育都做过行业模型，供你们选方向时参考。"),
      detail("需要你们", "三个方向什么时候能给。"),

      pageHead("第 9 页", "详情：DLM 实施条件与数据规格", "备查"),
      qa("数据要多少？", "LoRA 微调建议 5 千到 5 万条，SFT 微调建议 2 万到 20 万条以上，最终看数据质量和场景复杂度。"),
      qa("我们要出什么？", "训练、标注、独立评测数据的合法使用权；采集、清洗、标注、翻译、内容审核和 LGPD 合规；葡语、行业、安全、法律的本地专家。"),
      qa("算力够吗？", "8 月按 1 个行业模型规划了 70P，3 个需要重估。"),
      qa("自己训练搞不定怎么办？", "可以调整为讯飞进一步参与或直接承担训练。"),

      pageHead("第 10 页", "整体路径：两个 Stage，五个阶段", "略讲", "2 分钟", "②"),
      brief("一句带过", "项目分两个阶段：第一阶段建基础能力，约 280P；第二阶段全面运营，合计约 560P。五步从语料准备到交接运营，月份是示意，正式时间表随 SOW。"),
      detail("三行分工", "华为出算力、机房环境、网络和运维衔接，这一行单列，算力和环境的事不在讯飞。知识转移从第一步开始，不是最后才学。交接后你们主导三大中心的日常运营，我们持续支持。"),
      detail("对你们意味着", "你们的人第一步就要在场，采集语料、办授权是第一阶段的主要工作。许可、年费和长期路径在下一页。"),

      pageHead("第 11 页", "详情：交接、许可与运营支持", "备查", "要备好", "②"),
      p([run("这一页大概率会被问，对应你们对运营主权和供应商依赖的关切。", { italics: true, color: MUTE })]),
      qa("最后我们拿到什么？", "项目范围内永久授权：主权模型权重、本土语料、定制代码、平台与应用。权重可协商 100% 归巴西，在巴西本土数据中心存储维护。"),
      qa("代码能不能改、能不能给别人？", "定制代码有永久使用权和修改权，对外再分发需另行约定。三大中心和国民 APP 是讯飞核心资产，不开源；定制部分源码可开放；完全控制可付费买断。"),
      qa("怎么保证不永久依赖你们？", "三重保障：ONNX 导出和基础模型源码，可移植；五个知识包和现场培训，能力留在你们；两段式路径，1 到 3 年授权加增训加知识转移，3 到 5 年你们主导自主研发。"),
      qa("以后要付什么钱？", "运维阶段有软件运维年费和模型持续优化年费，另行商务约定，不影响软件使用。金额按报价。"),

      pageHead("第 12 页", "RNP 需要投入什么：人员与资源", "详讲", "4 分钟", "②③"),
      detail("先说结论", "这是今天最重要的一页。你们回去编人力计划，就照这页。八类人，做什么、什么阶段、大概几个。"),
      detail("逐行念", "项目负责人，统筹并确认方案、评测与验收。语料采集团队，执行本地语料采集，对接高校、SoberanIA、WideLabs、国家图书馆这些来源，第一阶段最重。葡语语言与标注专家，做语料审核、标注质量、发音人协调。LGPD 合规专员，负责数据来源分类、版权和合规审核，这是我们 4 月建议的岗位。行业专家，确定行业模型的场景、评价成果。模型工程师，参与国家大模型训练，自己做行业模型的训练和评测。应用开发，定义 APP 的本地场景，参与本地化和集成。运维工程师，第二阶段承担三大中心的日常运营。"),
      detail("资源四项", "本地语料 50B 及授权、标注与评测数据、实施环境、本地系统接口。到位时间按 SOW。"),
      detail("对你们意味着", "人数最终由你们定，我们给的是量级，SOW 的名额是参照。讯飞自己派几个人、驻场还是远程，会后随 SOW 一并给你们。"),
      detail("需要你们", "回填人数和到位时间。"),

      pageHead("第 13 页", "详情：参与矩阵", "略讲", "30 秒", "③"),
      brief("指一下", "这张表拿回去填：行是角色，列是五个阶段，格子是那个阶段同时投入的人数。填完就是你们 program 的骨架。同一个人可以跨阶段。"),

      pageHead("第 14 页", "RNP 能带走什么：知识转移两层安排", "详讲", "3 分钟", "③"),
      detail("先说结论", "两层。第一层已经在报价里，不另收费；第二层可选，单独报价。"),
      detail("第一层", "项目内知识转移。五个知识包：语料采集方法、数据工程、训练方法、模型评估、部署运维。方式是现场培训、工作坊、边做边学，行业模型由你们自己训练就是最大的一块实践。交付运维文档、产品文档、定制代码、集成脚本、接口文档、专属语料、工具链。学完的结果：你们可以自主研发行业模型和其他尺寸的模型，并主导三大中心运营。名额和授课语言按 SOW。"),
      detail("第二层", "可选的核心人才深入研修。面向少数有编程和模型基础的人，方向是训练调优、蒸馏、评测体系。要不要，你们评估后答复。"),
      detail("位置", "讯飞这一层是你们编 Capacity Program 的输入之一，华为的培训计划是另一层。软件栈的研发与验证算力在范围内，软件栈自研的主体工作由你们承担。"),

      pageHead("第 15 页", "详情：学习路径", "备查", "", "③"),
      qa("每类人具体怎么学、怎么证明学会了？", "这张表按角色写了：哪几个阶段参与什么任务，我们怎么支持，用什么验证。验证方式是示例，按 SOW 定。这张表也是给华为编整体能力计划的输入。"),

      pageHead("第 16 页", "SOW 边界与本次需确认事项", "详讲", "3 分钟", "②"),
      detail("三栏念一遍", "已包含：三项成果加知识转移；永久授权的权重、语料、定制代码与平台；开源模型导入和两次迁移适配；MaaS、导出与基础模型源码。可选另行约定：深入研修；清单外模型适配、微调、性能优化、定制开发；运维年费与模型优化年费；代码再分发；全量开源买断。不在范围：VLM、国家级技术验证实验室、全量非昇腾兼容测试。你们负责：本地语料与授权、LGPD 合规、场景确定、本地专家。"),
      detail("六项逐条要日期", "一，3 个行业模型的方向，你们和 MCTI 什么时候给，确认书什么时候签。二，参与人员、资源和语料采集计划，你们什么时候提交。三，实施计划，我们会后提交。四，知识转移名额和能力目标，按 SOW 核对后确认。五，可选研修要不要。六，范围与费用归属、许可与年费条款，讯飞和华为提交核对表。"),

      pageHead("收尾", "两个会后动作", "详讲", "1 分钟"),
      say("今天三件事：范围、分工与边界、知识转移，材料里都有对应页。你们回去编 Capacity Program 和人力计划，用第 12、13、15 页；向上汇报用第 2 页。"),
      say("我们会后一周内给实施计划、SOW 的名额与时间、讯飞的驻场安排，把蓝色数字补齐。请你们两周内给 3 个行业模型方向和人员名单。任何一页看不明白，随时找我们，不必等下次会。"),

      new Paragraph({ children: [new PageBreak()] }),
      h1("四、备用口径与准备事项"),
      p([run("会上三句备用口径", { bold: true })], { after: 60 }),
      bullet([run("行业模型 1 到 3：", { bold: true }), run("范围按最新方案扩为 3 个，实施条件不变，算力按 3 个重估后随 SOW 确认。不解释为什么变。")]),
      bullet([run("语料 500B 到 50B 加 300B：", { bold: true }), run("5 月 4 日写的是语料体系总目标，5 月 8 日已拆成你们本地 50B 与我们通用 300B。")]),
      bullet([run("人数与价格：", { bold: true }), run("结构今天定，数字按 SOW 和报价，不现场报。")]),
      p([run("四类敏感问题的边界", { bold: true })], { before: 200, after: 60 }),
      bullet("依赖与锁定：答第 11 页三重保障，不承诺全量开源。"),
      bullet("英伟达平台：可运行，性能低于昇腾，不承诺同等性能。"),
      bullet("人数与名额：只说结构，数字按 SOW。"),
      bullet("价格与年费：只说有哪两项，金额另行商务约定。"),
      p([run("会前三件事", { bold: true })], { before: 200, after: 60 }),
      bullet("讯飞 SOW 的名额和阶段时间，填进第 12、13 页。"),
      bullet("英文版材料，提前两天发给 Yara 团队。"),
      bullet("第 6 页的星火 APP 演示或截图。"),
      p([run("分工", { bold: true })], { before: 200, after: 60 }),
      bullet("主讲人（待定）讲开场、第 2、10、12、14、16 页和收尾；方案架构师讲第 3、6、8 页和全部详情页。"),
      bullet("华为的路易斯、郑丽雪在第 14 页接话，说明华为培训那一层和合稿安排。"),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("演讲脚本-会议版.docx", buf);
  console.log("wrote 演讲脚本-会议版.docx", buf.length, "bytes");
});
