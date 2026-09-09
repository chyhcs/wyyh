// 主汇报讲稿（中英文对照）→ docx
// 站在 RNP 客户视角，讲清 idea 与核心内容，不展开细节。
//   node build_speech_bilingual.js  ->  主汇报讲稿-中英对照.docx
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, PageOrientation, Footer, PageNumber,
} = require("docx");

const CN = "微软雅黑", EN = "Calibri";
const RED = "C00000", BLACK = "1D1D1A", MUTE = "666666", BLUE = "0070C0";

// **粗体** 行内标记
function runs(t, font, o = {}) {
  return String(t).split("**").map((seg, i) => new TextRun({
    text: seg, font, size: o.size || 21, bold: i % 2 === 1, color: o.color || BLACK,
  })).filter(r => r);
}
const para = (t, font, o = {}) => new Paragraph({
  children: runs(t, font, o),
  spacing: { after: o.after ?? 120, line: o.line || 300 },
});

function cell(lines, font, w, o = {}) {
  const arr = Array.isArray(lines) ? lines : [lines];
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 90, bottom: 90, left: 130, right: 130 },
    children: arr.map((l, i) => para(l, font, {
      ...o, after: i === arr.length - 1 ? 0 : 110,
    })),
  });
}

const W = 7619;
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" };

function pageTable(cnTitle, enTitle, pairs, cnTake, enTake) {
  const rows = [
    new TableRow({
      tableHeader: true, cantSplit: true,
      children: [
        cell(cnTitle, CN, W, { fill: "F7F2F2", bold: true, color: RED, size: 22 }),
        cell(enTitle, EN, W, { fill: "F7F2F2", bold: true, color: RED, size: 22 }),
      ],
    }),
    ...pairs.map(([c, e]) => new TableRow({
      cantSplit: true,
      children: [cell(c, CN, W), cell(e, EN, W)],
    })),
  ];
  if (cnTake) {
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        cell("落点　" + cnTake, CN, W, { fill: "F2F2F2", size: 20 }),
        cell("In one line　" + enTake, EN, W, { fill: "F2F2F2", size: 20 }),
      ],
    }));
  }
  return new Table({
    width: { size: W * 2, type: WidthType.DXA },
    columnWidths: [W, W],
    borders: { top: thin, bottom: thin, left: noB, right: noB, insideHorizontal: thin, insideVertical: thin },
    rows,
  });
}

function head(n, cn, mins) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 300, after: 140 },
    children: [
      new TextRun({ text: n, font: CN, size: 26, bold: true, color: BLACK }),
      new TextRun({ text: "　" + cn, font: CN, size: 26, bold: true, color: BLACK }),
      new TextRun({ text: "　　" + mins, font: CN, size: 19, color: MUTE }),
    ],
  });
}

// ============================== 讲稿内容 ==============================
const DECK = [
  {
    n: "第 1 页", cn: "封面 · 开场", mins: "约 2 分钟",
    tCN: "RNP × 科大讯飞 · 技术交流",
    tEN: "RNP × iFLYTEK · Technical Exchange",
    pairs: [
      ["感谢 RNP 安排今天的交流。我们今天不讲公司介绍，直接讲三件事：**讯飞交付什么、双方怎么分工、能力怎么转移到你们手上**。",
       "Thank you for arranging today's session. We'll skip the company introduction and go straight to three things: **what iFLYTEK delivers, how the work is divided, and how the capability transfers to you**."],
      ["这个项目对 RNP 的意义，不是买到三个模型和一个应用，而是**建立起巴西自己的 AI 能力**——项目结束时，平台在你们手里、模型权重在你们手里、你们的人会自己训练模型。这是我们整个方案的出发点。",
       "What this project means for RNP is not buying three models and an app. It is **building Brazil's own AI capability** — when the project ends, the platform is yours, the weights are yours, and your people can train models themselves. That is the premise behind the whole design."],
      ["主汇报 14 页，大约 25 分钟。后面还有 15 页附录，是我们预判你们会问的 15 个问题，问到哪一页我们直接翻过去。",
       "The main deck is 14 pages, about 25 minutes. Behind it are 15 appendix pages — the 15 questions we expect you to ask. Whenever one comes up, we'll jump straight to it."],
    ],
    takeCN: "今天讲的是能力怎么留在巴西，不是产品清单。",
    takeEN: "Today is about how the capability stays in Brazil — not a product list.",
  },
  {
    n: "第 2 页", cn: "一页看全：四项成果", mins: "约 2 分钟",
    tCN: "一页看全：讯飞交付的四项成果",
    tEN: "The whole picture: four outcomes iFLYTEK delivers",
    pairs: [
      ["这一页是全貌。讯飞交付四项成果：**一个平台、一个国家模型、一个国民应用、三个领域模型**。",
       "This page is the whole picture. iFLYTEK delivers four outcomes: **one platform, one national model, one national application, and three domain models**."],
      ["四项之间的关系很简单：**平台是底座，也是工作台**——另外三项都在这个平台上做出来。所以平台不只是一个交付物，它是你们以后自己干活的地方。",
       "The relationship between them is simple: **the platform is both the foundation and the workbench** — the other three are all built on it. So the platform isn't just a deliverable; it's where your team will work afterwards."],
      ["另外有两条贯穿全程：知识转移分两轮进行；交付后我们现场陪你们 12 个月、远程支持 5 年。每一项后面都有专门一页，我们一项一项讲。",
       "Two things run across all of it: knowledge transfer in two rounds, and after delivery, twelve months on site with us plus five years of remote support. Each outcome has its own page — let's take them one at a time."],
    ],
    takeCN: "四项成果不是并列的四件商品，是一个平台加三样在它上面长出来的东西。",
    takeEN: "The four outcomes aren't four separate products — it's one platform, plus three things grown on it.",
  },
  {
    n: "第 3 页", cn: "国家 AI 能力中心", mins: "约 2 分钟",
    tCN: "国家 AI 能力中心：三项成果的共用平台，也是 RNP 未来的工作平台",
    tEN: "National AI Capability Center: the shared platform, and RNP's future workbench",
    pairs: [
      ["先讲平台。平台有四个模块：语料、训练、应用、安全审核。",
       "Let's start with the platform. It has four modules: corpus, training, application, and safety review."],
      ["我们想请你们注意的不是模块本身，而是每个模块下面那一行——**每个模块都对应你们的一支团队**。语料中心对应数据团队和葡语专家，训练中心对应算法工程师，应用中心对应产品与运营，审核平台对应合规团队。",
       "What we'd ask you to look at is not the modules themselves, but the line under each one — **every module maps to one of your teams**. Corpus to your data team and Portuguese experts; training to your ML engineers; application to product and operations; review to compliance."],
      ["验收之后，这四个模块由这四支团队自己管。所以**今天演示的每一步操作，都是后面跟岗和实操阶段你们的人要自己做的**。讯飞负责把平台建起来、接通、培训到能用；环境和接口由你们提供，验收后由你们接管日常运营。",
       "After acceptance, those four teams run those four modules. So **every step in today's walkthrough is something your own people will be doing** during the shadowing and hands-on stages. We build the platform, connect it and train your team until they can use it; you provide the environment and interfaces, and take over daily operations after acceptance."],
    ],
    takeCN: "平台不是交给你们一套软件，是交给你们四支团队各自的工位。",
    takeEN: "We're not handing you software — we're handing four of your teams their workstations.",
  },
  {
    n: "第 4 页", cn: "国家模型能获得什么", mins: "约 2 分钟",
    tCN: "巴西葡语国家模型最终能获得什么？",
    tEN: "What the Brazilian Portuguese National Model delivers",
    pairs: [
      ["接下来三个部分，我们都用同一个顺序：**先讲你们最后拿到什么，再讲怎么做出来、谁做哪一段**。",
       "For the next three parts we use the same order: **first what you end up with, then how it gets built and who does which part**."],
      ["国家模型最后给到你们的，是一个真正懂巴西的基础模型能力，四个方面：一是**语言**，听得懂、说得出巴西葡语，包括口语和区域差异；二是**知识**，装进了巴西本地的数据和机构语境；三是**价值对齐**，数据存在巴西境内、权限与审计归你们，我们配合 LGPD 合规；四是**自主**，权重交给你们本地部署，后面的领域模型也建在它上面。",
       "What the national model finally gives you is a foundation model that genuinely understands Brazil, in four respects. **Language** — it understands and speaks Brazilian Portuguese, including spoken usage and regional variation. **Knowledge** — it carries Brazilian data and institutional context. **Alignment** — data stays inside Brazil, access and audit belong to RNP, and we support your LGPD compliance. **Autonomy** — the weights are delivered to you for local deployment, and the domain models are built on top of it."],
      ["有一点要说明：**效果和覆盖范围，按双方共同确认的评测标准判定，不由讯飞单方面说了算**。",
       "One point to be clear about: **performance and coverage are judged against evaluation criteria we agree together**, not determined by iFLYTEK alone."],
    ],
    takeCN: "不是一个会说葡语的模型，是一个属于巴西、你们能自己往下走的底座。",
    takeEN: "Not a model that speaks Portuguese — a base that belongs to Brazil and that you can carry forward.",
  },
  {
    n: "第 5 页", cn: "国家模型怎么建", mins: "约 3 分钟",
    tCN: "巴西葡语国家模型：讯飞先行增强基座，RNP 提供本地语料",
    tEN: "National Model: iFLYTEK enhances the base first, RNP supplies the local corpus",
    pairs: [
      ["怎么做出来？这里有一个安排我们想特别说明。",
       "How is it built? There's one arrangement here we'd particularly like to explain."],
      ["**第 1 步是两条线同时走的。**你们在巴西建本地语料；同时我们在中国用自有算力采集公开的葡语语料，先把基座模型的通用葡语能力增强起来——**这一步不占用巴西的算力，也不需要等你们的语料齐备**。采集到的公开语料，我们也会一并提交给巴西方。",
       "**Step 1 runs on two parallel tracks.** You build the local corpus in Brazil; at the same time, in China, we use our own compute to collect public Portuguese corpus and strengthen the general Portuguese capability of the base model first — **that uses no Brazilian compute and doesn't wait for your corpus to be complete**. The public corpus we collect is handed over to Brazil as well."],
      ["等本地语料就绪、算力到位，我们就在这个已经增强过的基座上继续训练，直接进入第 2 步。好处对 RNP 很直接：**工期风险小了一大截**——不是等你们准备好我们才开始，而是我们已经先跑了一段。",
       "Once your corpus is ready and compute is in place, we continue training on that already-strengthened base and move straight into step 2. The benefit to RNP is direct: **it takes a significant chunk of schedule risk off the table**. We're not waiting for you to be ready before we start — we've already run part of the distance."],
      ["后面第 3 步评测、第 4 步本地交付：测试集和通过规则由双方共同确定，权重和模型卡交付给你们。",
       "Then step 3, evaluation — test sets and pass rules agreed jointly; and step 4, local delivery — we deploy in Brazil and hand over the weights and model cards."],
    ],
    takeCN: "我们已经提前起跑了，你们的语料一到位就能接上训练。",
    takeEN: "We've already started running. The moment your corpus lands, training picks up.",
  },
  {
    n: "第 6 页", cn: "国家 AI 应用能获得什么", mins: "约 2 分钟",
    tCN: "国家 AI 应用最终能获得什么？",
    tEN: "What the National AI Application delivers",
    pairs: [
      ["第二项成果，国家 AI 应用。最后给到你们的，是一个**面向你们批准的用户的统一 AI 入口**。",
       "The second outcome is the National AI Application. What you end up with is **a single AI entry point for the users you approve**."],
      ["从三个角度看：对**使用者**，是一个 APP——问答检索、文档分析与翻译、写作与笔记，还有巴葡的语音输入和朗读，手机、电脑、Web 都能用；对**管理者**，是一个数据运营平台——用户、租户、权限、日志、运营报表；对**合规**，是内置治理——接内容安全审核平台，数据和日志都在巴西境内。",
       "From three angles. For **users**, it's an app — Q&A and search, document analysis and translation, writing and notes, with Brazilian Portuguese speech input and read-aloud, on mobile, desktop and web. For **administrators**, a Data Operations Platform — users, tenants, permissions, logs, operational reporting. For **compliance**, governance is built in — it connects to the content safety review platform, and data and logs stay inside Brazil."],
      ["边界要讲清楚：**讯飞交付软件，不承担用户增长、内容生产、人工审核和业务托管运营**。这些是你们的产品、运营和审核团队的事。",
       "The boundary should be explicit: **iFLYTEK delivers the software. We do not take on user growth, content production, human moderation, or running the service as a managed business.** Those sit with your product, operations and review teams."],
    ],
    takeCN: "软件我们给，用户和内容归你们——这条线越早说清越好。",
    takeEN: "We provide the software; users and content are yours. Better said clearly now than later.",
  },
  {
    n: "第 7 页", cn: "国家 AI 应用怎么建", mins: "约 1.5 分钟",
    tCN: "国家 AI 应用：讯飞交付软件，RNP 负责内容、用户与运营",
    tEN: "National AI Application: iFLYTEK delivers the software, RNP owns content, users and operations",
    pairs: [
      ["建设分四步：明确需求、本地化、系统集成、上线运营。",
       "The build runs in four steps: define requirements, localise, integrate, then launch and operate."],
      ["对 RNP 最关键的是**人员到位的时间点**：产品团队**在需求阶段就要到位**，因为用户范围、内容规则、品牌这些由你们定；运营团队**要在 Stage 1 交付前到位**，交付当天就要接得住。",
       "The point that matters most for RNP is **when your people arrive**. The product team needs to be in place **at the requirements stage**, because user scope, content rules and branding are yours to define. The operations team needs to be in place **before the Stage 1 delivery**, so they can pick it up on day one."],
      ["上线之后，交付版本你们可以自主演进——不是锁死的。",
       "After launch, the delivered version is yours to evolve — it isn't locked."],
    ],
    takeCN: "这一页真正的交付物是时间点：你们的人什么时候必须到。",
    takeEN: "The real deliverable on this page is timing — when your people have to be there.",
  },
  {
    n: "第 8 页", cn: "领域模型能获得什么", mins: "约 1.5 分钟",
    tCN: "领域模型 DLM 最终能获得什么？",
    tEN: "What the Domain Large Models deliver",
    pairs: [
      ["第三项，领域模型。你们指定三个领域——目前是卫生、航天、生物多样性，每个领域聚焦一个边界清晰的固定场景。",
       "Third outcome: the domain models. You name three domains — currently Health, Space and Biodiversity — each with one clearly bounded scenario."],
      ["模型最后能做三件事：**读懂**这个领域的材料与规范；**回答**这个领域的问题；围绕约定场景做**分析与辅助判断**，归纳材料、协助起草报告。要说清楚的是：**它是辅助，不替代专业判断和决策**。",
       "What the models can finally do is three things: **read** the domain's material and standards; **answer** questions in that domain; and, within the agreed scenario, provide **analysis and decision support** — summarising material, helping draft reports. To be clear: **this supports professional judgement, it does not replace it**."],
      ["还有一个关键的不同——**这三个模型的权重，是由 RNP 自己训练产生的，归 RNP**。",
       "And one key difference — **the weights of these three models are produced by RNP's own training, and they belong to RNP**."],
    ],
    takeCN: "前两项是我们做给你们，这一项是你们自己做出来的。",
    takeEN: "The first two we build for you. This one you build yourselves.",
  },
  {
    n: "第 9 页", cn: "领域模型怎么建", mins: "约 2 分钟",
    tCN: "领域模型 DLM：RNP 训练，讯飞指导；3 项服务，每项 1 个场景",
    tEN: "Domain Large Models: RNP trains, iFLYTEK advises — 3 services, one scenario each",
    pairs: [
      ["这一页和国家模型**正好相反**：四个步骤都由 RNP 执行，讯飞在旁边指导。",
       "This page is **the reverse** of the national model: all four steps are executed by RNP, with iFLYTEK advising alongside."],
      ["这是 **RNP 第一次自己训练模型**。所以你们需要自己的算法工程师和领域专家——前面知识转移里的跟岗和指导实操，就是为这件事设计的。",
       "This is **RNP's first time training a model yourselves**. That's why you need your own ML engineers and domain experts — and it is exactly what the shadowing and guided hands-on stages are designed for."],
      ["我们提供的是：可行性评估、技术方案、训练路线建议、参数建议、查看训练日志、协助分析问题，最后出分析报告。**执行、数据和评测都在你们这边。**",
       "What we provide: feasibility assessment, the technical plan, recommended training routes, parameter recommendations, reviewing training logs, helping diagnose problems, and the final analysis report. **Execution, data and evaluation are all on your side.**"],
      ["边界：生产部署、API 开发和应用集成不在这项服务里，需要的话可以另行约定。",
       "The boundary: production deployment, API development and application integration are not part of this service. They can be arranged separately if needed."],
    ],
    takeCN: "方向盘在这一页交到你们手上，我们坐在副驾。",
    takeEN: "This is the page where the wheel passes to you. We're in the passenger seat.",
  },
  {
    n: "第 10 页", cn: "知识转移：学到什么、获得什么", mins: "约 2 分钟",
    tCN: "知识转移：RNP 学到什么、获得什么",
    tEN: "Knowledge transfer: what RNP learns and what RNP gets",
    pairs: [
      ["知识转移是这个项目里，我们认为**对 RNP 长期价值最大**的一块。",
       "We think knowledge transfer is the part of this project with **the greatest long-term value for RNP**."],
      ["学到什么，分三层：**基础课**，把从语料到服务的全流程打通；**12 门专题课**，分四个方向的高级研讨——这两层都在项目范围内，**不另收费**。第三层是**高端知识转移**，讲这套系统本身是怎么做出来的：集群调度、模型训练、语料、推理，另有专家研讨和来中国考察——这一层**需要另行预算**。",
       "What you learn comes in three levels. A **core course** connecting the whole pipeline from corpus to service; then **twelve specialist courses** in four tracks — both **inside project scope, at no extra cost**. The third level is the **premium transfer**: how this system itself was built — cluster scheduling, model training, corpus, inference — plus expert seminars and a study visit to China. That level **requires a separate budget**."],
      ["讲师不是外请的培训师，是**参与本项目交付的讯飞工程师和方案架构师**，研讨就在交付给你们的平台和模型上做。",
       "The instructors are not external trainers. They are **the iFLYTEK engineers and architects who deliver this project**, and the sessions run on the platform and models we deliver to you."],
      ["获得什么：模型权重、平台软件和定制代码、全部文档（英语或葡语），以及现场 12 个月加远程 60 个月的支持。",
       "What you get: the model weights, the platform software and custom code, all documentation in English or Portuguese, and support — twelve months on site plus sixty months remote."],
    ],
    takeCN: "教你们的人，就是给你们做这套系统的人。",
    takeEN: "The people who teach you are the people who built it for you.",
  },
  {
    n: "第 11 页", cn: "知识转移：四阶段、两轮", mins: "约 2.5 分钟",
    tCN: "知识转移：两个轮次，合计 1216 课时、约 158 个工作日",
    tEN: "Knowledge transfer: two rounds, 1,216 hours, about 158 working days",
    pairs: [
      ["怎么做？**四个阶段，做两轮**。逻辑是逐步交接方向盘。",
       "How is it done? **Four stages, run twice.** The logic is handing over the steering wheel step by step."],
      ["**奠基**是听讲；**跟岗**是我们操作、你们在旁边看和协助，而且看的是**真实任务**——国家模型的训练、12 个月内的模型升级、三项 DLM 的训练；**指导实操**反过来，你们操作、我们在旁边；**正式移交**交生产环境、交流程、交资产许可，最后验证你们能独立操作，双方签字。",
       "**Foundation** is classroom. In **shadowing**, we operate and you observe and assist — and you observe **real work**: the national model training, the model upgrades over twelve months, the three DLM trainings. In **guided hands-on** it reverses: you operate, we stand beside you. In **formal handover** we transfer the production environment, the processes and the asset licences, and verify that you can operate independently, signed off by both sides."],
      ["两轮合计 1216 课时、约 158 个工作日、**累计 260 人次**。说明两点：260 是两轮累计人次，同一个人参加多个阶段分别计入，**不是 260 名不同学员**；工作日是培训活动的天数，分散在实施期内，**不是连续工期**。",
       "Across two rounds: 1,216 hours, about 158 working days, **260 attendances in total**. Two clarifications: 260 is cumulative attendance over two rounds — one person joining several stages is counted each time, so it is **not 260 different people**; and working days are days of training activity spread across the implementation period, **not a continuous block**."],
      ["移交之后，你们要能独立做四件事：平台运维、训练提交、模型部署、版本发布。",
       "After handover you should be able to do four things on your own: run the platform, submit training jobs, deploy models, and release versions."],
    ],
    takeCN: "培训不是听课，是跟着真实任务把方向盘接过去。",
    takeEN: "This isn't classroom time — it's taking the wheel on real work.",
  },
  {
    n: "第 12 页", cn: "实施路径", mins: "约 2 分钟",
    tCN: "实施路径：两个 Stage、五个步骤、三方分工",
    tEN: "Implementation: two stages, five steps, three parties",
    pairs: [
      ["把前面几件事放到时间线上看：两个 Stage、五个步骤、三方分工。",
       "Now let's put all of that on a timeline: two stages, five steps, three parties."],
      ["三方分工一句话说清：**华为出算力和机房，讯飞出软件和模型，RNP 出数据和运营**。",
       "The split between the three parties in one line: **Huawei provides compute and facilities, iFLYTEK provides software and models, RNP provides data and operations**."],
      ["五个步骤**不是按日历推进的，是按条件推进的**——每个节点下面写的是完成标志。比如第 1 步的完成标志是平台验收加语料试批合格，达到了才走第 2 步。月份在实施前由双方确定，T0 是第一批算力就绪。",
       "The five steps **don't advance by calendar, they advance on preconditions** — under each node is its completion marker. Step 1, for example, completes on platform acceptance plus a passing corpus trial batch; only then does step 2 begin. The months are fixed jointly before implementation, with T0 being the first compute batch."],
      ["知识转移的四个阶段**贯穿全程**，不是最后才开始；正式移交在陪伴运维结束前至少 3 个月启动。",
       "The four knowledge-transfer stages run **across the whole project**, not at the end; and formal handover starts at least three months before the assisted O&M period ends."],
    ],
    takeCN: "进度不看日历看条件——这对三方都是保护。",
    takeEN: "Progress is gated by conditions, not dates. That protects all three parties.",
  },
  {
    n: "第 13 页", cn: "RNP 的投入", mins: "约 2 分钟",
    tCN: "RNP 的投入：SOW 规定的职责与三批到位的人员",
    tEN: "RNP's investment: SOW responsibilities and three staffing intakes",
    pairs: [
      ["这一页是**我们最希望你们带回去的一页**——RNP 这边要投入什么。",
       "This is **the page we'd most like you to take away** — what RNP needs to put in."],
      ["SOW 规定由 RNP 承担四件事：生产葡语语料并取得授权；自主训练和评测 DLM；审批用户并运营 APP；协调华为的算力和环境、组织验收。",
       "Under the SOW, four things sit with RNP: producing and licensing the Portuguese corpus; training and evaluating the DLMs yourselves; approving users and operating the app; and coordinating Huawei's compute and environment, and organising acceptance."],
      ["人员建议**分三批到位**：第一批是数据团队和葡语专家，**从第 1 步就要开始**；第二批是算法工程师和测试工程师，从第 2 步起，他们先参与国家模型的训练评测，然后接着自己训 DLM；第三批是应用运营和平台运维，从第 3 步起。",
       "We suggest staffing in **three intakes**. First, the data team and Portuguese experts — needed **from step 1**. Second, ML engineers and test engineers from step 2 — they first take part in the national model training and evaluation, then go on to train the DLMs themselves. Third, application operations and platform O&M, from step 3."],
      ["人数是参考量级，来自我们早期的估算；各职能峰值可以跨活动复用，**不是同时在岗人数**，最终按 SOW 的分工和工作量确认。知识转移的核心学员，就从这三批人里指定。",
       "The headcounts are indicative, from our early estimates. Peaks in each function can be reused across activities, so these are **not simultaneous headcounts** — final numbers follow the SOW split and workload. The core trainees for knowledge transfer are named from these three intakes."],
    ],
    takeCN: "整个方案里唯一没法由讯飞代劳的，就是这一页。",
    takeEN: "This page is the one part of the plan iFLYTEK cannot do on your behalf.",
  },
  {
    n: "第 14 页", cn: "对照 RNP 的 TR", mins: "约 1.5 分钟",
    tCN: "对照 RNP 的 TR：五层要求分别由谁落地",
    tEN: "Against RNP's TR: who delivers each of the five layers",
    pairs: [
      ["最后一页，回到你们的 TR。TR 把国家 AI 软件栈的要求分成五层，这张表是**承包商整体——华为加讯飞——的对应关系**。",
       "The last page comes back to your TR. It groups the national AI software stack into five layers, and this table shows how **the contractor side as a whole — Huawei plus iFLYTEK — maps to them**."],
      ["**第一层**底层计算软件，由华为承担，讯飞的平台运行在它之上。**第二层**数据、模型与应用平台，是讯飞的交付主体，就是前面第 3 到第 9 页的全部内容。**第三到第五层**——集成与运行保障、研发与人才能力、交付与长期自主——三方共同；其中人才和长期自主这两层，对应的正是刚才讲的知识转移和后续支持。",
       "**Layer one**, low-level computing software, is Huawei's; our platform runs on top of it. **Layer two**, the data, model and application platform, is iFLYTEK's main delivery — everything on pages 3 through 9. **Layers three to five** — integration and operational assurance, R&D and talent capability, delivery and long-term autonomy — are joint; and the talent and autonomy layers map directly to the knowledge transfer and support we just described."],
      ["我们的目的不是说\"都覆盖了\"，而是**让你们能一条一条对得上**。表里凡是需要逐项确认的，我们在实施前和你们一起过一遍。",
       "Our aim isn't to claim full coverage. It's to let you **check it line by line**. Wherever an item needs item-by-item confirmation, we go through it with you before implementation."],
    ],
    takeCN: "不是宣称覆盖，是请你们逐条核对。",
    takeEN: "Not a claim of coverage — an invitation to check it line by line.",
  },
  {
    n: "收尾", cn: "三句话 + 转入问答", mins: "约 1.5 分钟",
    tCN: "收尾与转场",
    tEN: "Closing and hand-off to Q&A",
    pairs: [
      ["三句话收一下。",
       "Let me close with three sentences."],
      ["**第一**，四项成果最后都落在你们手里——平台你们运营、权重你们持有、代码和文档都交给你们。",
       "**First**, all four outcomes end up in your hands — you operate the platform, you hold the weights, the code and documentation are handed over to you."],
      ["**第二**，国家模型是我们训、你们参与；领域模型反过来，是你们训、我们指导。到项目结束，你们要能自己走完全流程。",
       "**Second**, we train the national model with you taking part; the domain models are the reverse — you train, we advise. By the end of the project you should be able to run the whole pipeline yourselves."],
      ["**第三**，我们最需要 RNP 现在开始准备的是**人**——数据团队和葡语专家最早，算法工程师第二，运营团队第三。这是整个方案里唯一没法由我们代劳的部分。",
       "**Third**, what we most need RNP to start preparing now is **people** — the data team and Portuguese experts first, ML engineers second, operations third. That is the one part of this plan we cannot do on your behalf."],
      ["后面是 15 页附录，对应我们预判的 15 个问题。请开始提问。",
       "Behind this are 15 appendix pages covering the questions we expect. Please go ahead."],
    ],
    takeCN: "留给客户的最后印象：能力留在巴西，人要你们自己备。",
    takeEN: "The last impression to leave: the capability stays in Brazil — the people are yours to prepare.",
  },
];

// ============================== 组装文档 ==============================
const children = [
  new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: "RNP × 科大讯飞 · 主汇报讲稿（中英文对照）", font: CN, size: 34, bold: true, color: RED })],
  }),
  new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: "RNP × iFLYTEK · Speaking Script for the Main Deck (Chinese / English)", font: EN, size: 26, bold: true, color: BLACK })],
  }),
  new Paragraph({
    spacing: { after: 240 },
    children: [new TextRun({ text: "里约热内卢 · 2026 年 9 月 9 日　|　对应 14 页主汇报　|　全程约 28 分钟，附录留给问答", font: CN, size: 19, color: MUTE })],
  }),
  new Paragraph({
    spacing: { after: 300 },
    shading: { type: ShadingType.CLEAR, fill: "F7F7F7", color: "auto" },
    indent: { left: 160, right: 160 },
    children: [
      new TextRun({ text: "用法　", font: CN, size: 20, bold: true, color: RED }),
      new TextRun({ text: "左栏中文、右栏英文，逐段对照，可任选一种语言主讲、另一种备查。加粗处是需要说重、说清的地方。每页末尾的", font: CN, size: 20, color: MUTE }),
      new TextRun({ text: "「落点」", font: CN, size: 20, bold: true, color: MUTE }),
      new TextRun({ text: "是这一页只要客户记住一句话时的那一句。讲稿只讲 idea 与核心内容，具体数字与条款在 PPT 页面与 15 页附录上，问到再展开。", font: CN, size: 20, color: MUTE }),
    ],
  }),
];

for (const d of DECK) {
  children.push(head(d.n, d.cn, d.mins));
  children.push(pageTable(d.tCN, d.tEN, d.pairs, d.takeCN, d.takeEN));
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: CN, size: 21, color: BLACK }, paragraph: { spacing: { line: 300 } } } },
    paragraphStyles: [
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: CN, size: 26, bold: true, color: BLACK } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
        margin: { top: 720, bottom: 620, left: 800, right: 800 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: EN, size: 18, color: MUTE })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync("主汇报讲稿-中英对照.docx", b);
  console.log("saved 主汇报讲稿-中英对照.docx  pages:", DECK.length);
});
