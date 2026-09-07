// Build the 7-page RNP executive briefing deck from the material package script.
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const Fi = require('react-icons/fi');
const sharp = require('sharp');

const OUT = process.argv[2] || path.join(__dirname, 'RNP-iFLYTEK-高层汇报-7页-16x9.pptx');
const IMG = path.join(__dirname, 'assets');

const C = {
  NAVY: '0B2A5B', DEEP: '17408F', BLUE: '1F5EDB', BLUE2: '2F6FE4',
  TINT: 'EEF3FB', TINT2: 'DCE6F7', LINE: 'C9D6EA', GRAY: 'F4F6FA',
  TEXT: '1F2A3D', MUTED: '5F6F86', WHITE: 'FFFFFF',
  AMBER_BG: 'FFF3D6', AMBER_LN: 'E8CF93', AMBER_TX: '7A4E00',
};
const FONT = 'Arial';
const W = 13.333, ML = 0.55, CW = W - 2 * ML;

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'iFLYTEK RNP project team';
pres.company = 'iFLYTEK';
pres.title = 'RNP Executive Briefing';

// ---------- helpers ----------
const iconCache = {};
async function iconData(name, color, size = 256) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const Comp = Fi[name];
  if (!Comp) throw new Error('unknown icon ' + name);
  let svg = renderToStaticMarkup(React.createElement(Comp, { size, color: '#' + color }));
  svg = svg.replace(/currentColor/g, '#' + color);
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  iconCache[key] = 'image/png;base64,' + buf.toString('base64');
  return iconCache[key];
}
function shadow() { return { type: 'outer', color: '0B2A5B', blur: 6, offset: 1.5, angle: 90, opacity: 0.10 }; }

function box(s, x, y, w, h, o = {}) {
  const fill = o.fill || C.TINT;
  const opts = { x, y, w, h, fill: { color: fill }, line: { color: o.line || C.LINE, width: o.lw != null ? o.lw : 0.75 } };
  if (o.dash) opts.line.dashType = o.dash;
  if (o.round !== false) opts.rectRadius = o.r != null ? o.r : 0.08;
  if (o.shadow) opts.shadow = shadow();
  s.addShape(o.round === false ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, opts);
}

function text(s, str, x, y, w, h, o = {}) {
  const opts = {
    x, y, w, h, fontFace: FONT, fontSize: o.size || 14, color: o.color || C.TEXT,
    bold: !!o.bold, italic: !!o.italic, align: o.align || 'left', valign: o.valign || 'middle',
    margin: o.margin || [3, 6, 3, 6], isTextBox: true,
  };
  if (o.fill) opts.fill = { color: o.fill };
  if (o.line) opts.line = { color: o.line, width: o.lw != null ? o.lw : 0.75, dashType: o.dash || 'solid' };
  if (o.shape) opts.shape = o.shape;
  if (o.r != null) opts.rectRadius = o.r;
  if (o.rotate != null) opts.rotate = o.rotate;
  if (o.ls) opts.lineSpacingMultiple = o.ls;
  if (o.shadow) opts.shadow = shadow();
  s.addText(str, opts);
}

async function icon(s, name, x, y, d, o = {}) {
  const bg = o.bg || C.BLUE, fg = o.fg || C.WHITE;
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg, width: 0.5 } });
  const inset = d * (o.inset || 0.26);
  s.addImage({ data: await iconData(name, fg), x: x + inset, y: y + inset, w: d - 2 * inset, h: d - 2 * inset });
}

function arrow(s, x1, y1, x2, y2, o = {}) {
  const line = { color: o.color || C.BLUE, width: o.width || 1.5 };
  if (o.dash) line.dashType = o.dash;
  if (!o.noHead) line.endArrowType = 'triangle';
  if (o.beginHead) line.beginArrowType = 'triangle';
  s.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1, line,
  });
}

function header(s, n, title) {
  s.background = { color: C.WHITE };
  text(s, title, ML, 0.38, 11.0, 1.1, { size: 32, bold: true, color: C.NAVY, valign: 'top', margin: [0, 0, 0, 0] });
  text(s, String(n).padStart(2, '0') + ' / 07', W - ML - 1.0, 0.42, 1.0, 0.35, { size: 12, color: C.MUTED, align: 'right', margin: [0, 0, 0, 0] });
  text(s, 'iFLYTEK for RNP   |   Executive briefing   |   Scope basis: SOW V1', ML, 7.05, 9, 0.3, { size: 10, color: C.MUTED, margin: [0, 0, 0, 0] });
}

// ---------- speaker notes (Chinese, from the generation script) ----------
const NOTES = {
1: `【本页客户应记住什么】项目同时交付用户可使用的服务和RNP可管理的底层能力，并通过实施、培训及支持帮助落地。

【讲解词】这次建设要让RNP获得能够在巴西运行的葡语AI服务。用户看到的是应用，RNP同时获得数据处理、模型训练和服务管理工具。讯飞负责约定软件、模型、应用的交付和联调，并通过培训与支持帮助RNP承接。项目也需要RNP提供有授权的数据、本地专家和运营人员。后面几页将具体说明系统怎样工作、谁来使用，以及哪些条件决定上线。

【衔接下一页】先看RNP在后台怎样把数据和模型转成可使用的服务。

【来源与素材】S01第1章、§3.1、第8章；S02第4页（仅参考组织方式，已按SOW重排为可编辑对象，不使用旧图日期和未确认型号）。

【口径与边界】“交付”指SOW约定范围及目标，不表示已完成部署、验收或签约。HLD仅为历史方案参考，不覆盖SOW边界。安全审核与安全产品为贯穿能力：2台堡垒机、2台WAF、日志审计、数据库审计、漏洞扫描各1台，合计7台（S01 §8.8，合计为本稿加总）。SOW §5.2—5.4安全分工仍有“待与华为对齐”标记，整体安全责任由相关方统一，不宣称讯飞包揽整体安全治理。`,

2: `【本页客户应记住什么】RNP获得一套可以持续操作的模型工程和服务管理工具，同时收到可验收的软件、模型资产和操作资料。

【讲解词】第一个中心解决数据能否用于模型的问题，第二个中心支持训练、评测和版本管理，第三个中心把模型能力发布给应用和开发者。约定基础模型的实际权重及必要组件在巴西交付，并验证本地加载和推理；其他模型按各自模型包和许可范围交付。资料既可以用于训练，也可以进入知识库支撑检索问答。领域模型方面，当前基线是三个固定场景的辅助训练服务，由RNP执行训练、讯飞提供方案和指导。

【衔接下一页】这些后台能力最后会体现在用户每天能完成的任务里。

【来源与素材】S01 §3.1、§3.3—3.8、§3.13、§8.1—8.4、§9.3；S02第5、9、12、13页。

【备注与边界】专项服务数量只放备注：3项固定场景领域模型辅助训练服务（S01 §8.4.3），不画成三个已交付的行业应用；第三方模型推理适配为2项服务（§8.4.5），不表示所有模型均已兼容。基础模型开放权重不自动延伸为语音识别、语音合成及安全审核模型全部开放。模型权重、源码和所有权应按资产许可矩阵分别说明。领域模型结果责任需先消除原文冲突（§3.14与§8.4.3、§9.5.3、§11.8）。模型规格：1个巴西葡语文本基础模型（总参数不低于290B）、1个语音识别模型、1个语音合成模型（男声、女声各1套约定发音人）。`,

3: `【本页客户应记住什么】应用提供统一入口，帮助用户从问题、文档或语音得到可核对、可编辑的结果。

【讲解词】普通用户不需要理解模型训练。用户输入一个葡语问题，上传支持的文档，或者通过语音提问，就可以查找资料、提取要点、形成研究或办公草稿。引用帮助用户返回来源核对，专业判断和正式发布仍由人完成。RNP通过运营后台控制使用对象、权限、发布及反馈。

【衔接下一页】用一项完整工作任务，来看这些功能怎样组合起来。

【来源与素材】S01 §3.8—3.9、§8.5；S02第7—8页；S06官方产品指南只作中国产品演示参考。本页图片为S02第7页提取的中国星火桌面（639×352）和移动端（162×365）界面参考，仅用于功能讲解和构图参考；取得当前巴西葡语版本实图后应替换。

【必须保留的边界】葡语PPT生成及本地公文模板仍需适配验证，不能作为已经确认可用的主卖点。语音输入和朗读不等于同声传译。不得增加自动审批、付款、实验执行、任意业务系统操作或所有终端原生客户端。独立文档翻译系统亦不属于据此承诺的范围。画面中的国内工具和PPT入口不自动属于本期交付；SOW为Web加约定客户端，移动端小图不新增任何移动平台承诺。`,

4: `【本页客户应记住什么】把任务、输入资料、核对方法和最终成果连起来，客户就能判断应用是否适用于自己的团队。

【讲解词】我们建议从边界清晰的资料整理任务开始。研究人员给出问题和允许使用的材料，系统帮助整理观点和引用，用户再核对来源，最后得到可以用于汇报的草稿。同样的流程可以用于机构制度查询和办公材料整理。中科院的官方开通通知提供了一个真实的中国科研工作流参考，但巴西的资料、语言效果、权限和用户规模仍要在本项目中验证。

【衔接下一页】明确任务以后，再确定哪些用户先使用，以及服务容量怎样匹配推广规模。

【来源与素材】功能依据S01 §3.8—3.9；中国案例依据S05中科院文献情报中心官方开通通知（2023-11-03，https://www.las.cas.cn/news/fwcx/202311/t20231103_6915220.html）。本页RNP场景为建议方案，中国案例为已公布的历史事实。三份资料是演示设计数量，不是产品上限。

【演示安排】如有经核实的30—45秒录屏（分镜见材料03），可在本页播放，并准备三张真实关键帧；没有视频时按本页静态流程独立讲解，不放空白播放器或假视频链接。

【边界】不得展示无证据的节时比例、科研提升率、国内用户规模或“巴西已落地”字样。`,

5: `【本页客户应记住什么】总用户规划、同时调用的服务容量和真实应用体验需要分别管理；部署和推广应按已验证负载推进。

【讲解词】前期方案使用过1,000万总用户的覆盖规划，但这个数字不是日活人数，也不是今天已经验收的服务能力。SOW给出的是几类服务各自的并发基线，需要在约定模型、硬件和业务负载下验证。我们的建议是先确认首批机构、用户和任务组合，再开展端到端测试和分批推广。本地部署包括模型、平台及应用，联网搜索等外部服务仍要单独确认访问路径、授权及费用。一个研究任务可能触发多次检索、文本推理和审核，模型并发数不能直接换算成可以同时使用应用的人数。

【衔接下一页】系统运行之后，还要让RNP有人员、有方法管理它。

【来源与素材】S01 §3.1、§3.8—3.10、§8.5、§9.2；S02第19、21页；S03第3页。

【内部核对】旧HLD与资源稿的APP并发存在约1,000（S02第19、21页）和约500（S03第3页）两种口径，不同时作为本次承诺上屏；不把各服务并发相加成为用户数，不沿用旧带宽估算证明实际容量。若RNP不再认可1,000万规划，对外页可只展示用户类别，将该数字移到备注。`,

6: `【本页客户应记住什么】本地能力包括语言及内容适配、本地可用资产、RNP人员操作能力和持续支持渠道。

【讲解词】本地能力有三个实际表现：服务适配巴西葡语和约定内容，软件及模型资产在指定环境交付，RNP团队能够完成约定的操作与管理任务。培训共578小时，从基础学习到跟岗、指导实操，再到正式移交。第一年提供现场陪伴，五年内持续提供远程技术支持，二者前一年重叠。日常运营、内容审批和用户管理由RNP负责，因此固定参与的本地人员是项目成功的必要投入。

【衔接下一页】最后把这些交付对应到阶段成果，以及RNP现在需要落实的条件。

【来源与素材】S01 §3.1—3.2、第6章、§7.2—7.3、第11章；S02第15页（仅参考组织方法，不复用旧课时）。

【准确口径】四阶段30、360、20、168小时，合计578小时为本稿加总（S01 §6.1）。基础阶段30名核心学员，后续每阶段10名，不相加成60名不同人员。应至少在首年陪伴结束前3个月启动正式移交。五年支持不等于五年持续代运营或无限重训。另有模型持续优化服务：全部相关初验完成后连续12个月，文本3次、语音识别1次、语音合成1次、安全审核1次正式升级（S01 §3.11），其起点及内容与陪伴支持区分。正式起算日按SOW和书面确认填写。驻场人数、地点、语言安排仍待落实，不宣称已建成巴西研发中心或永久驻场团队。`,

7: `【本页客户应记住什么】项目有清楚的交付路径；关键风险可以转化为有责任人、有证据的推进条件。

【讲解词】第一阶段在环境就绪后部署平台、训练适配葡语模型并交付应用初版；第二阶段在新增算力就绪后接入资源、启用服务并逐步移交。最关键的投入是可用环境、合格数据、本地专家和运营人员。双方需要明确首批用户与场景，冻结终端、接口和容量测试方法，并落实持续费用和责任人。这样可以按实际证据判断阶段完成情况，再决定推广。

【风险与应对】
1. 环境不具备或兼容性不足，会阻断部署、训练和容量验证 → 环境就绪清单、兼容性记录、阶段资源接入确认；RNP及基础设施方提供环境，讯飞核对产品依赖并开展范围内验证。
2. 数据授权、质量或本地专家不足，会影响训练与葡语效果 → 数据权利和质量清单、代表性试批、独立评测材料、专家确认记录；RNP组织数据及专家，讯飞提供范围内工具、辅助服务和模型工作。
3. 功能、终端、接口和负载不明确，会造成范围扩大与体验预期偏差 → 功能正面清单、终端及API目录、混合任务测试方案、缺陷闭环；双方冻结范围并按批准方法验收。
4. 运营人员或持续费用缺位，会影响移交及后续可用性 → 承接岗位、实操结果、外部搜索/API及许可费用安排、剩余支持计划；RNP落实日常运营，双方确认支持和费用边界。

【验收目标】本地部署能够运行，并实际交付约定权重及必要组件、确认许可满足约定用途（巴西本地加载推理和部署测试记录为证）；用户能完成约定任务（业务用例和功能测试为证）；平台能管理访问、容量和风险（配置及验证记录为证）；RNP能完成约定操作（实操和移交确认材料为证）。四项为本稿对SOW验收要求的高层归纳，具体通过规则以批准方案为准，不能另加“零错误”“100%安全”等承诺。

【来源】S01第2章、§3.1、§3.8—3.11、第6—11章。满足验收或双方书面确认等约定起算条件后进入支持期，不能据此认定60个月服务已经完成（§7.2—7.3、§9.8.2）。当前SOW未提供冻结的日历计划，不沿用旧HLD中的日期。`,
};

// ---------- slide 1 ----------
async function slide1() {
  const s = pres.addSlide();
  header(s, 1, 'RNP Gains AI Services for Local Use and Operation');
  const labelX = ML, labelW = 1.3;
  const bx = ML + labelW + 0.12;
  const boxW = 5.75;
  const bandX = bx + boxW + 0.15;
  const bandW = 0.95;
  const fullW = boxW + 0.15 + bandW;
  const rx = bandX + bandW + 0.4;
  const rw = W - ML - rx;
  const rowLabel = (t, y, h) => text(s, t, labelX, y, labelW, h, { size: 11.5, color: C.MUTED, align: 'right', margin: [0, 4, 0, 0] });

  text(s, 'Brazilian Portuguese AI services for RNP-approved users', bx, 1.6, fullW, 0.68,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.NAVY, color: C.WHITE, size: 16, bold: true, align: 'center', shadow: true });
  rowLabel('Users receive', 1.6, 0.68);

  text(s, 'AI application and operations platform', bx, 2.42, boxW, 0.78,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.TINT, line: C.LINE, color: C.NAVY, size: 16, align: 'center' });
  rowLabel('Application layer', 2.42, 0.78);

  text(s, 'Text, speech recognition and speech synthesis models', bx, 3.34, boxW, 0.78,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.TINT, line: C.LINE, color: C.NAVY, size: 16, align: 'center' });
  rowLabel('Model layer', 3.34, 0.78);

  box(s, bx, 4.26, boxW, 1.05, { fill: C.TINT, line: C.LINE, r: 0.06 });
  const cw = (boxW - 0.4 - 0.3) / 3;
  ['Corpus construction', 'Model training', 'Application innovation'].forEach((t, i) => {
    text(s, t, bx + 0.2 + i * (cw + 0.15), 4.41, cw, 0.75,
      { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.WHITE, line: C.LINE, color: C.NAVY, size: 14, bold: true, align: 'center' });
  });
  rowLabel('Capability centers', 4.26, 1.05);

  text(s, 'Brazilian local environment provided by RNP and its infrastructure providers', bx, 5.45, fullW, 0.6,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.GRAY, line: C.NAVY, dash: 'dash', color: C.MUTED, size: 13.5, align: 'center' });
  rowLabel('Runs in Brazil', 5.45, 0.6);

  // vertical cross-cutting band (rotated text box)
  const bandY = 2.42, bandH = 5.31 - 2.42;
  const cx = bandX + bandW / 2, cy = bandY + bandH / 2;
  text(s, 'Content moderation and agreed security products', cx - bandH / 2, cy - bandW / 2, bandH, bandW,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.DEEP, color: C.WHITE, size: 13, bold: true, align: 'center', rotate: 270 });

  await icon(s, 'FiUsers', ML, 6.22, 0.5, { bg: C.TINT2, fg: C.NAVY });
  text(s, 'RNP provides approved data, local experts and operating teams', ML + 0.65, 6.17, 8.2, 0.6,
    { size: 15, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });

  const sideCard = async (y, iconName, label, copy) => {
    box(s, rx, y, rw, 1.85, { fill: C.WHITE, line: C.LINE, shadow: true });
    await icon(s, iconName, rx + 0.2, y + 0.22, 0.5);
    text(s, label, rx + 0.82, y + 0.2, rw - 1.0, 0.54, { size: 12.5, bold: true, color: C.MUTED, margin: [0, 0, 0, 0] });
    text(s, copy, rx + 0.2, y + 0.85, rw - 0.4, 0.9, { size: 15, color: C.NAVY, valign: 'top', margin: [0, 0, 0, 0] });
  };
  await sideCard(1.6, 'FiPackage', 'Delivery services', 'Deployment, integration, testing and operating documentation');
  await sideCard(3.6, 'FiBookOpen', 'Capability building', 'Knowledge transfer and ongoing support');
  text(s, 'Delivery here means the agreed SOW scope and targets. It does not mean that deployment, acceptance or signature has been completed.',
    rx, 5.6, rw, 1.2, { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.AMBER_BG, line: C.AMBER_LN, color: C.AMBER_TX, size: 12, margin: [6, 10, 6, 10] });
  s.addNotes(NOTES[1]);
}

// ---------- slide 2 ----------
async function slide2() {
  const s = pres.addSlide();
  header(s, 2, 'Three Centers Support Data Preparation, Model Training and Service Delivery');
  const y = 1.72, h = 2.5, gap = 0.36, cw = (CW - 2 * gap) / 3;
  const cards = [
    { name: 'Corpus Construction', icon: 'FiDatabase', action: 'Import, clean and check approved data', result: 'Datasets and quality records' },
    { name: 'Model Training', icon: 'FiCpu', action: 'Train, evaluate and manage model versions', result: 'Agreed model assets and evaluation reports' },
    { name: 'Application Innovation', icon: 'FiShare2', action: 'Publish services and control access and usage', result: 'Governed interfaces for applications' },
  ];
  for (let i = 0; i < 3; i++) {
    const c = cards[i], x = ML + i * (cw + gap);
    box(s, x, y, cw, h, { fill: C.WHITE, line: C.LINE, shadow: true });
    await icon(s, c.icon, x + 0.2, y + 0.2, 0.55);
    text(s, c.name, x + 0.88, y + 0.18, cw - 1.05, 0.6, { size: 16, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
    text(s, 'Main action', x + 0.2, y + 0.93, cw - 0.4, 0.28, { size: 10.5, bold: true, color: C.MUTED, margin: [0, 0, 0, 0] });
    text(s, c.action, x + 0.2, y + 1.2, cw - 0.4, 0.5, { size: 14, color: C.TEXT, valign: 'top', margin: [0, 0, 0, 0] });
    text(s, 'RNP receives', x + 0.2, y + 1.72, cw - 0.4, 0.28, { size: 10.5, bold: true, color: C.MUTED, margin: [0, 0, 0, 0] });
    text(s, c.result, x + 0.2, y + 2.0, cw - 0.4, 0.48, { size: 14, bold: true, color: C.NAVY, valign: 'top', margin: [0, 0, 0, 0] });
    if (i < 2) arrow(s, x + cw + 0.06, y + h / 2, x + cw + gap - 0.06, y + h / 2, { width: 2 });
  }
  const ky = y + h + 0.25;
  box(s, ML, ky, CW, 0.56, { fill: C.TINT, line: C.TINT });
  await icon(s, 'FiBookOpen', ML + 0.15, ky + 0.07, 0.42, { bg: C.TINT2, fg: C.NAVY });
  text(s, [
    { text: 'Knowledge base path: ', options: { bold: true, color: C.NAVY } },
    { text: 'approved documents can also support retrieval-based Q&A directly, without retraining each time.', options: { color: C.TEXT } },
  ], ML + 0.7, ky, CW - 0.9, 0.56, { size: 13.5, margin: [0, 0, 0, 0] });
  const c3x = ML + 2 * (cw + gap) + cw / 2;
  arrow(s, c3x, ky, c3x, y + h + 0.04, { width: 1.5 });

  const gy = ky + 0.68;
  await icon(s, 'FiShield', ML + 0.15, gy + 0.03, 0.36, { bg: C.WHITE, fg: C.MUTED });
  text(s, 'Access permissions, activity logs and content review apply across all three centers.', ML + 0.62, gy, CW - 0.7, 0.42,
    { size: 12.5, italic: true, color: C.MUTED, margin: [0, 0, 0, 0] });

  const ay = gy + 0.57;
  text(s, [
    { text: 'RNP receives the specified base model weights, required runtime components, licenses and operating guides for local use.', options: { fontSize: 14.5, bold: true, color: C.NAVY, breakLine: true } },
    { text: 'Model specification: one text model with at least 290B total parameters, one speech recognition model and one speech synthesis model with two agreed voices.', options: { fontSize: 11.5, color: C.MUTED } },
  ], ML, ay, CW, 1.05, { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.WHITE, line: C.NAVY, lw: 1.25, align: 'center', margin: [8, 14, 8, 14] });
  s.addNotes(NOTES[2]);
}

// ---------- slide 3 ----------
async function slide3() {
  const s = pres.addSlide();
  header(s, 3, 'Users Ask, Read, Research and Draft in Portuguese');
  const iy = 1.6, a1w = 4.9, a1h = a1w * 352 / 639;
  const a2h = a1h, a2w = a2h * 162 / 365;
  s.addShape(pres.shapes.RECTANGLE, { x: ML - 0.03, y: iy - 0.03, w: a1w + 0.06, h: a1h + 0.06, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
  s.addImage({ path: path.join(IMG, 'A1_HLD_p07_SparkDesk_China_reference.png'), x: ML, y: iy, w: a1w, h: a1h });
  const a2x = ML + a1w + 0.18;
  s.addShape(pres.shapes.RECTANGLE, { x: a2x - 0.03, y: iy - 0.03, w: a2w + 0.06, h: a2h + 0.06, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
  s.addImage({ path: path.join(IMG, 'A2_HLD_p07_MobileApp_China_reference.png'), x: a2x, y: iy, w: a2w, h: a2h });
  const leftW = a1w + 0.18 + a2w;
  text(s, 'China product interface reference (desktop and mobile). Brazilian features and the final client scope follow the agreed scope and localization validation.',
    ML, iy + a1h + 0.08, leftW, 0.5, { size: 10.5, italic: true, color: C.MUTED, valign: 'top', margin: [0, 0, 0, 0] });

  const oy = 5.53;
  box(s, ML, oy, leftW, 0.95, { fill: C.TINT, line: C.TINT });
  await icon(s, 'FiSliders', ML + 0.18, oy + 0.2, 0.55, { bg: C.NAVY });
  text(s, [
    { text: 'Operations console', options: { fontSize: 10.5, bold: true, color: C.MUTED, breakLine: true } },
    { text: 'RNP manages users, permissions, publishing, usage and feedback.', options: { fontSize: 14.5, bold: true, color: C.NAVY } },
  ], ML + 0.9, oy, leftW - 1.05, 0.95, { margin: [0, 0, 0, 0] });

  const rx = ML + leftW + 0.4, rw = W - ML - rx;
  const rows = [
    ['FiMessageSquare', 'Ask and Search', 'Ask questions and check available sources.'],
    ['FiFileText', 'Read Documents', 'Summarize supported files and ask follow-up questions.'],
    ['FiSearch', 'Research a Topic', 'Compare sources and create a report draft.'],
    ['FiEdit3', 'Draft and Organize', 'Prepare editable text, outlines and notes.'],
    ['FiMic', 'Speak and Listen', 'Use Portuguese voice input and answer readout.'],
  ];
  const rh = 0.88, rg = 0.12;
  for (let i = 0; i < rows.length; i++) {
    const [ic, label, copy] = rows[i], y = 1.6 + i * (rh + rg);
    box(s, rx, y, rw, rh, { fill: C.TINT, line: C.TINT });
    await icon(s, ic, rx + 0.18, y + (rh - 0.55) / 2, 0.55);
    text(s, label, rx + 0.9, y, 1.9, rh, { size: 15, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
    text(s, copy, rx + 2.85, y, rw - 3.0, rh, { size: 13.5, color: C.TEXT, margin: [0, 0, 0, 0] });
  }
  s.addNotes(NOTES[3]);
}

// ---------- slide 4 ----------
async function slide4() {
  const s = pres.addSlide();
  header(s, 4, 'Researchers Can Prepare a Source-Based Brief for Review');
  text(s, 'Proposed RNP scenario', ML, 1.68, 2.5, 0.38,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.19, fill: C.AMBER_BG, line: C.AMBER_LN, color: C.AMBER_TX, size: 12, bold: true, align: 'center', margin: [0, 0, 0, 0] });
  const steps = [
    ['User and task', 'An authorized researcher compares two campus AI deployment options.'],
    ['Input', 'A clear question and three approved documents'],
    ['AI assistance', 'Summarize findings, compare evidence and draft a brief'],
    ['Human review', 'Open the sources and correct unsupported claims'],
    ['Reusable output', 'An editable brief with references and open questions'],
  ];
  const y0 = 2.2, rh = 0.7, rg = 0.1, leftW = 8.25;
  s.addShape(pres.shapes.LINE, { x: ML + 0.25, y: y0 + rh / 2, w: 0, h: (steps.length - 1) * (rh + rg), line: { color: C.LINE, width: 2 } });
  for (let i = 0; i < steps.length; i++) {
    const y = y0 + i * (rh + rg);
    box(s, ML + 0.55, y, leftW - 0.55, rh, { fill: C.TINT, line: C.TINT });
    text(s, String(i + 1), ML, y + (rh - 0.5) / 2, 0.5, 0.5,
      { shape: pres.shapes.OVAL, fill: C.NAVY, line: C.WHITE, lw: 1.5, color: C.WHITE, size: 14, bold: true, align: 'center', margin: [0, 0, 0, 0] });
    text(s, steps[i][0], ML + 0.75, y, 1.95, rh, { size: 15, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
    text(s, steps[i][1], ML + 2.75, y, leftW - 2.9, rh, { size: 15, color: C.TEXT, margin: [0, 0, 0, 0] });
  }
  const ey = y0 + 5 * rh + 4 * rg + 0.22;
  await icon(s, 'FiBriefcase', ML + 0.05, ey + 0.03, 0.4, { bg: C.TINT2, fg: C.NAVY });
  text(s, 'Staff can query approved institutional guidance and prepare office drafts.', ML + 0.6, ey, leftW - 0.6, 0.46,
    { size: 13.5, color: C.TEXT, margin: [0, 0, 0, 0] });

  const rx = ML + leftW + 0.4, rw = W - ML - rx;
  box(s, rx, y0, rw, 3.5, { fill: C.WHITE, line: C.LINE, shadow: true });
  text(s, 'China reference', rx + 0.2, y0 + 0.2, 1.55, 0.32,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.16, fill: C.TINT2, color: C.NAVY, size: 11, bold: true, align: 'center', margin: [0, 0, 0, 0] });
  text(s, 'China reference: CAS and iFLYTEK jointly developed Spark Research Assistant. CAS announced internal access in November 2023.',
    rx + 0.2, y0 + 0.65, rw - 0.4, 1.35, { size: 14, color: C.TEXT, valign: 'top', margin: [0, 0, 0, 0] });
  text(s, 'A reference for institutional research workflows', rx + 0.2, y0 + 2.05, rw - 0.4, 0.55,
    { size: 13.5, italic: true, bold: true, color: C.NAVY, valign: 'top', margin: [0, 0, 0, 0] });
  text(s, 'Source: National Science Library, Chinese Academy of Sciences. Public notice dated 3 November 2023.',
    rx + 0.2, y0 + 2.7, rw - 0.4, 0.75, { size: 10.5, color: C.MUTED, valign: 'top', margin: [0, 0, 0, 0] });
  text(s, 'AI output stays a draft until a person has checked the sources.', rx, y0 + 3.62, rw, 0.6,
    { size: 11.5, italic: true, color: C.MUTED, valign: 'top', margin: [0, 0, 0, 0] });
  s.addNotes(NOTES[4]);
}

// ---------- slide 5 ----------
async function slide5() {
  const s = pres.addSlide();
  header(s, 5, 'User Rollout Depends on Workloads and Local Service Capacity');
  const ty = 1.7, lw = 3.6, ch = 1.22, dy = 1.42;
  const userCard = async (x, y, w, h, ic, title, entry, o = {}) => {
    box(s, x, y, w, h, { fill: o.fill || C.WHITE, line: o.line || C.LINE, dash: o.dash, shadow: !o.dash });
    await icon(s, ic, x + 0.18, y + (h - 0.55) / 2, 0.55, { bg: o.iconBg || C.BLUE, fg: o.iconFg || C.WHITE });
    text(s, title, x + 0.9, y + 0.1, w - 1.05, 0.6, { size: 14.5, bold: true, color: C.NAVY, valign: 'bottom', margin: [0, 0, 0, 0] });
    text(s, entry, x + 0.9, y + 0.74, w - 1.05, h - 0.78, { size: 12, color: C.MUTED, valign: 'top', margin: [0, 0, 0, 0] });
  };
  await userCard(ML, ty, lw, ch, 'FiUsers', 'RNP-approved users', 'Entry: Web and agreed clients');
  await userCard(ML, ty + dy, lw, ch, 'FiCode', 'Authorized applications and developers', 'Entry: Approved service interfaces');

  const bx = ML + lw + 0.4, bw = 4.4, bh = dy + ch;
  box(s, bx, ty, bw, bh, { fill: C.WHITE, line: C.NAVY, lw: 1.5, r: 0.08 });
  text(s, 'Brazilian deployment', bx + 0.2, ty + 0.12, bw - 0.4, 0.4, { size: 14, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  const iw = (bw - 0.4 - 0.3) / 3;
  ['Applications', 'Platforms', 'Models'].forEach((t, i) => text(s, t, bx + 0.2 + i * (iw + 0.15), ty + 0.62, iw, 0.8,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.06, fill: C.BLUE, color: C.WHITE, size: 13.5, bold: true, align: 'center', margin: [0, 0, 0, 0] }));
  text(s, 'Applications, models and platforms run in the agreed Brazilian environment.', bx + 0.2, ty + 1.55, bw - 0.4, 0.95,
    { size: 12.5, color: C.TEXT, valign: 'top', margin: [0, 0, 0, 0] });
  arrow(s, ML + lw + 0.05, ty + ch / 2, bx - 0.05, ty + ch / 2, { width: 1.75 });
  arrow(s, ML + lw + 0.05, ty + dy + ch / 2, bx - 0.05, ty + dy + ch / 2, { width: 1.75 });

  const rx = bx + bw + 0.4, rw = W - ML - rx;
  await userCard(rx, ty, rw, ch, 'FiUserCheck', 'RNP administrators', 'Entry: Operations and management tools', { iconBg: C.NAVY });
  arrow(s, bx + bw + 0.05, ty + ch / 2, rx - 0.05, ty + ch / 2, { width: 1.75, color: C.NAVY });
  await userCard(rx, ty + dy, rw, ch, 'FiGlobe', 'Approved external services', 'Internet search and third-party services',
    { fill: C.GRAY, line: C.MUTED, dash: 'dash', iconBg: C.TINT2, iconFg: C.NAVY });
  arrow(s, bx + bw + 0.05, ty + dy + ch / 2, rx - 0.05, ty + dy + ch / 2, { width: 1.25, color: C.MUTED, dash: 'dash', beginHead: true });

  const r1 = ty + bh + 0.3;
  text(s, 'Earlier planning reference', ML, r1 + 0.08, 2.75, 0.4,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.2, fill: C.GRAY, line: C.LINE, color: C.MUTED, size: 11.5, bold: true, align: 'center', margin: [0, 0, 0, 0] });
  text(s, 'Earlier HLD planning reference: 10 million total users', ML + 2.95, r1, CW - 2.95, 0.56, { size: 15, color: C.MUTED, margin: [0, 0, 0, 0] });

  const r2 = r1 + 0.66, th = 1.02;
  text(s, 'SOW capacity baseline', ML, r2 + 0.12, 2.75, 0.4,
    { shape: pres.shapes.ROUNDED_RECTANGLE, r: 0.2, fill: C.NAVY, color: C.WHITE, size: 11.5, bold: true, align: 'center', margin: [0, 0, 0, 0] });
  text(s, 'Concurrent requests', ML, r2 + 0.58, 2.75, 0.35, { size: 12, color: C.NAVY, bold: true, align: 'center', margin: [0, 0, 0, 0] });
  const tiles = [['1,000', 'text'], ['200', 'speech recognition'], ['200', 'speech synthesis'], ['400', 'content moderation']];
  const tx0 = ML + 2.95, tw = (CW - 2.95 - 3 * 0.2) / 4;
  tiles.forEach(([n, l], i) => {
    const x = tx0 + i * (tw + 0.2);
    box(s, x, r2, tw, th, { fill: C.TINT, line: C.TINT });
    text(s, n, x, r2 + 0.06, tw, 0.55, { size: 26, bold: true, color: C.NAVY, align: 'center', margin: [0, 0, 0, 0] });
    text(s, l, x, r2 + 0.6, tw, 0.38, { size: 12, color: C.MUTED, align: 'center', valign: 'top', margin: [0, 0, 0, 0] });
  });
  text(s, 'Rollout targets require confirmation. Capacity tests use agreed workloads. Peak capacities are not assumed to run simultaneously.',
    ML, r2 + th + 0.12, CW, 0.42, { size: 11, color: C.MUTED, margin: [0, 0, 0, 0] });
  s.addNotes(NOTES[5]);
}

// ---------- slide 6 ----------
async function slide6() {
  const s = pres.addSlide();
  header(s, 6, 'Training and Supported Practice Prepare RNP to Operate Locally');
  const ty = 1.7, th = 0.66;
  text(s, 'Local capability', ML, ty, 1.7, th, { size: 14, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  const caps = [['FiMessageCircle', 'Brazilian Portuguese adaptation'], ['FiMapPin', 'Local deployment'], ['FiUsers', 'RNP operating capability']];
  const cx0 = ML + 1.8, cgap = 0.2, cwid = (CW - 1.8 - 2 * cgap) / 3;
  for (let i = 0; i < 3; i++) {
    const x = cx0 + i * (cwid + cgap);
    box(s, x, ty, cwid, th, { fill: C.TINT, line: C.TINT });
    await icon(s, caps[i][0], x + 0.15, ty + (th - 0.46) / 2, 0.46, { bg: C.TINT2, fg: C.NAVY });
    text(s, caps[i][1], x + 0.75, ty, cwid - 0.85, th, { size: 14, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  }
  const sy = 2.56, sh = 1.0, sgap = 0.3, sw = (CW - 3 * sgap) / 4;
  const stages = [['Foundation', '30 hours', C.BLUE2], ['Shadowing', '360 hours', C.BLUE], ['Guided Practice', '20 hours', C.DEEP], ['Handover', '168 hours', C.NAVY]];
  stages.forEach(([n, hrs, col], i) => {
    const x = ML + i * (sw + sgap);
    box(s, x, sy, sw, sh, { fill: col, line: col, shadow: true });
    text(s, n, x + 0.2, sy + 0.1, sw - 0.4, 0.4, { size: 15, bold: true, color: C.WHITE, margin: [0, 0, 0, 0] });
    text(s, hrs, x + 0.2, sy + 0.48, sw - 0.4, 0.46, { size: 22, bold: true, color: C.WHITE, margin: [0, 0, 0, 0] });
    if (i < 3) arrow(s, x + sw + 0.05, sy + sh / 2, x + sw + sgap - 0.05, sy + sh / 2, { width: 2, color: C.NAVY });
  });
  const ry = sy + sh + 0.1;
  text(s, 'iFLYTEK leads, RNP shadows', ML, ry, 3.0, 0.34, { size: 12, italic: true, color: C.MUTED, margin: [0, 0, 0, 0] });
  arrow(s, ML + 3.1, ry + 0.17, W - ML - 3.1, ry + 0.17, { width: 1.25, color: C.MUTED });
  text(s, 'RNP leads, iFLYTEK guides', W - ML - 3.0, ry, 3.0, 0.34, { size: 12, italic: true, color: C.MUTED, align: 'right', margin: [0, 0, 0, 0] });
  text(s, '578 training hours support agreed operating tasks and handover.', ML, ry + 0.42, CW, 0.42, { size: 15, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });

  const ly = ry + 1.0;
  const ax0 = ML + 2.6, axLen = W - ML - ax0;
  const bar = (y, runs, months, col) => {
    text(s, runs, ML, y, 2.5, 0.44, { size: 12, color: C.NAVY, margin: [0, 0, 0, 0] });
    box(s, ax0, y, axLen * months / 60, 0.44, { fill: col, line: col, r: 0.05 });
  };
  bar(ly, [{ text: 'Onsite support', options: { bold: true, breakLine: true } }, { text: '12 months, 5×8', options: {} }], 12, C.BLUE);
  bar(ly + 0.56, [{ text: 'Remote technical support', options: { bold: true, breakLine: true } }, { text: '60 months, 7×24', options: {} }], 60, C.NAVY);
  const x12 = ax0 + axLen * 12 / 60;
  s.addShape(pres.shapes.LINE, { x: x12, y: ly - 0.1, w: 0, h: 1.2, line: { color: C.MUTED, width: 1, dashType: 'dash' } });
  text(s, 'The first 12 months overlap.', x12 + 0.12, ly, 3.2, 0.44, { size: 12, italic: true, color: C.MUTED, margin: [0, 0, 0, 0] });
  const tyk = ly + 1.1;
  s.addShape(pres.shapes.LINE, { x: ax0, y: tyk, w: axLen, h: 0, line: { color: C.LINE, width: 1 } });
  [0, 12, 24, 36, 48, 60].forEach(m => {
    const x = ax0 + axLen * m / 60;
    s.addShape(pres.shapes.LINE, { x, y: tyk - 0.05, w: 0, h: 0.1, line: { color: C.MUTED, width: 1 } });
    text(s, m === 0 ? 'Month 0' : String(m), x - 0.45, tyk + 0.06, 0.9, 0.28, { size: 10.5, color: C.MUTED, align: 'center', margin: [0, 0, 0, 0] });
  });
  const by = tyk + 0.48;
  await icon(s, 'FiUserCheck', ML, by + 0.03, 0.4, { bg: C.TINT2, fg: C.NAVY });
  text(s, 'RNP leads daily operations, content decisions and user management.', ML + 0.55, by, CW - 0.6, 0.46, { size: 14.5, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  s.addNotes(NOTES[6]);
}

// ---------- slide 7 ----------
async function slide7() {
  const s = pres.addSlide();
  header(s, 7, 'Delivery Milestones Depend on Resources, Data and Operating Readiness');
  const py = 1.72, ph = 1.1, pgap = 0.4, pw = (CW - pgap) / 2;
  const phases = [
    ['Phase 1', 'Deploy platforms, adapt models and deliver the initial application.', C.BLUE,
      'Evidence: environment readiness checklist, local model loading and inference records, deployment and functional test results.'],
    ['Phase 2', 'Connect additional resources, enable agreed services and transfer operations.', C.NAVY,
      'Evidence: resource connection confirmation, capacity and configuration validation records, hands-on results and handover confirmation.'],
  ];
  phases.forEach(([n, copy, col, ev], i) => {
    const x = ML + i * (pw + pgap);
    box(s, x, py, pw, ph, { fill: col, line: col, shadow: true });
    text(s, n, x + 0.25, py + 0.12, pw - 0.5, 0.35, { size: 12.5, bold: true, color: C.WHITE, margin: [0, 0, 0, 0] });
    text(s, copy, x + 0.25, py + 0.47, pw - 0.5, 0.6, { size: 15, bold: true, color: C.WHITE, valign: 'top', margin: [0, 0, 0, 0] });
    text(s, ev, x + 0.25, py + ph + 0.06, pw - 0.5, 0.5, { size: 11, color: C.MUTED, valign: 'top', margin: [0, 0, 0, 0] });
    if (i === 0) arrow(s, x + pw + 0.06, py + ph / 2, x + pw + pgap - 0.06, py + ph / 2, { width: 2, color: C.NAVY });
  });
  const tyy = py + ph + 0.7;
  const c1w = 3.9, hh = 0.42, rh = 0.5;
  box(s, ML, tyy, CW, hh, { fill: C.TINT2, line: C.TINT2, r: 0.04 });
  text(s, 'Key dependency', ML + 0.2, tyy, c1w - 0.2, hh, { size: 12.5, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  text(s, 'Required action', ML + c1w + 0.15, tyy, CW - c1w - 0.15, hh, { size: 12.5, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
  const deps = [
    ['FiServer', 'Infrastructure readiness', 'Confirm resources and compatibility.'],
    ['FiDatabase', 'Data and Portuguese quality', 'Approve data, experts and evaluation tasks.'],
    ['FiLayers', 'Scope and service capacity', 'Freeze channels, interfaces and test workloads.'],
    ['FiRepeat', 'Operating continuity', 'Assign owners and ongoing service budgets.'],
  ];
  for (let i = 0; i < deps.length; i++) {
    const y = tyy + hh + i * rh;
    if (i % 2 === 1) s.addShape(pres.shapes.RECTANGLE, { x: ML, y, w: CW, h: rh, fill: { color: C.GRAY }, line: { color: C.GRAY, width: 0 } });
    s.addShape(pres.shapes.LINE, { x: ML, y: y + rh, w: CW, h: 0, line: { color: C.LINE, width: 0.75 } });
    await icon(s, deps[i][0], ML + 0.15, y + (rh - 0.36) / 2, 0.36, { bg: C.WHITE, fg: C.BLUE });
    text(s, deps[i][1], ML + 0.65, y, c1w - 0.65, rh, { size: 14, bold: true, color: C.NAVY, margin: [0, 0, 0, 0] });
    text(s, deps[i][2], ML + c1w + 0.15, y, CW - c1w - 0.15, rh, { size: 14, color: C.TEXT, margin: [0, 0, 0, 0] });
  }
  const cy = tyy + hh + 4 * rh + 0.28;
  box(s, ML, cy, CW, 0.6, { fill: C.NAVY, line: C.NAVY, shadow: true });
  await icon(s, 'FiFlag', ML + 0.15, cy + 0.1, 0.4, { bg: C.BLUE, fg: C.WHITE });
  text(s, 'RNP confirms the first users and scenarios. Both teams agree owners, readiness gates and acceptance evidence.',
    ML + 0.7, cy, CW - 0.85, 0.6, { size: 14, bold: true, color: C.WHITE, margin: [0, 0, 0, 0] });
  s.addNotes(NOTES[7]);
}

(async () => {
  await slide1(); await slide2(); await slide3(); await slide4(); await slide5(); await slide6(); await slide7();
  await pres.writeFile({ fileName: OUT });
  console.log('written', OUT);
})().catch(e => { console.error(e); process.exit(1); });
