# -*- coding: utf-8 -*-
"""
讯飞 × RNP 技术交流 PPT（中文精简版，9 页），基于华为简报模板。
每页一个问题；三项成果各一页，共用同一版式。
具体数字、名称、日期先填示例，以蓝色“示例：”标出，后续按实际替换。

用法：python3 build_cn_simple.py <template.pptx> <out.pptx>
"""
import sys

from pptx import Presentation

from pptx_helpers import (BLUET, BODY, CW, GREEN, GREEN_T, L, LINE, MUTE, ORANGE, ORANGE_T, PANEL,
                          PANEL2, R, RED, REDT, TXT, WHITE, arrow, body, card, chip, dot,
                          drop_first_slide, label, new_slide, note, notes, table, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "RNP-讯飞交流-中文简版.pptx"
RNP_C = "1F4E79"

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]


def slide(title):
    return new_slide(prs, LAYOUT, title)


def flow(s, y, h, steps, gap=0.25, x=L, w=CW, name_size=13, desc_size=11):
    """横向步骤：编号、名称、一句说明。"""
    n = len(steps)
    bw = (w - (n - 1) * gap) / n
    for i, (name, desc) in enumerate(steps):
        bx = x + i * (bw + gap)
        card(s, bx, y, bw, h, fill=REDT if i == 0 else PANEL)
        dot(s, bx + 0.18, y + 0.18, 0.36, str(i + 1))
        text(s, bx + 0.64, y + 0.14, bw - 0.8, 0.44, name, size=name_size, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, bx + 0.18, y + 0.7, bw - 0.36, h - 0.8, desc, size=desc_size)
        if i < n - 1:
            arrow(s, bx + bw + 0.02, y + 0.26)


def two_party(s, y, h, xf, rnp, x=L, w=CW, gap=0.3):
    half = (w - gap) / 2
    for i, (name, color, fill, items) in enumerate([("讯飞", RED, REDT, xf), ("RNP", RNP_C, BLUET, rnp)]):
        cx = x + i * (half + gap)
        card(s, cx, y, half, h, fill=fill)
        label(s, cx + 0.25, y + 0.12, 1.5, name, color=color, size=13)
        body(s, cx + 0.25, y + 0.5, half - 0.5, h - 0.6, items, size=11.5, bullet=True, bcolor=color, sa=3)


def band(s, y, h, title, content, fill=PANEL2, line=LINE, title_w=1.3):
    card(s, L, y, CW, h, fill=fill, line=line)
    label(s, L + 0.25, y + (h - 0.34) / 2, title_w, title, color=RED, size=12.5)
    body(s, L + 0.25 + title_w, y + 0.12, CW - 0.5 - title_w, h - 0.2, content, size=12, valign="m")


def deliverable_page(title, purpose, steps, xf, rnp, bottom_title, bottom, visual=None):
    s = slide(title)
    band(s, 1.05, 0.62, "用途", purpose)
    label(s, L, 1.85, 3, "怎么做", color=TXT, size=12.5)
    flow(s, 2.2, 1.5, steps)
    label(s, L, 3.9, 3, "谁做什么", color=TXT, size=12.5)
    if visual:
        vw = 3.6
        card(s, L, 4.25, vw, 1.4, fill=PANEL, line=LINE)
        text(s, L + 0.15, 4.25, vw - 0.3, 1.4, visual, size=11, color=MUTE, align="c", valign="m", margin=(0, 0, 0, 0), sa=2)
        two_party(s, 4.25, 1.4, xf, rnp, x=L + vw + 0.3, w=CW - vw - 0.3)
    else:
        two_party(s, 4.25, 1.4, xf, rnp)
    band(s, 5.9, 0.62, bottom_title, bottom, fill=PANEL, line=None, title_w=1.7)
    return s


# ============================== 1 封面 ==============================
s = slide("")
text(s, 0.8, 1.6, 8.5, 0.4, "RNP × 科大讯飞 · 技术交流", size=14, color=MUTE, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 2.1, 8.5, 1.0, "葡语大模型、应用与领域模型", size=36, color=RED, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 3.1, 8.5, 0.6, "讯飞交付什么、双方怎么分工、能力如何转移", size=20, color=TXT, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 4.3, 8.5, 1.2, [
    "里约热内卢 · {{示例：2026 年 9 月 9 日}}",
    "汇报：科大讯飞 {{示例：陈浩、方案架构师}}",
    "讨论稿，内容以 SOW、报价与技术方案为准",
], size=13, color=BODY, margin=(0, 0, 0, 0), ls=1.35, sa=4)
card(s, 0.8, 5.85, 8.5, 0.55, fill=BLUET)
text(s, 0.95, 5.85, 8.2, 0.55, "图例：{{蓝色文字}} = 示例内容，需按实际情况替换（日期、人名、数字、场景、口径）", size=11.5, color=BODY, valign="m", margin=(0, 0, 0, 0), sa=0)
bx, by, bw, bh, g = 9.8, 2.0, 2.9, 0.9, 0.22
for i, t in enumerate(["葡语基础模型", "应用 APP", "领域模型 DLM"]):
    card(s, bx, by + i * (bh + g), bw, bh, fill=PANEL)
    dot(s, bx + 0.2, by + i * (bh + g) + (bh - 0.38) / 2, 0.38, str(i + 1))
    text(s, bx + 0.75, by + i * (bh + g), bw - 0.85, bh, t, size=15, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, bx, by + 3 * (bh + g), bw, 0.75, fill=RED)
text(s, bx, by + 3 * (bh + g), bw, 0.75, "知识转移", size=15, color=WHITE, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
notes(s, "本次交流讲四件事：讯飞交付什么、怎么做、双方怎么分工、巴方人员怎样学会并接手。依据 Mota 电话中 RNP 的诉求。")

# ============================== 2 总览 ==============================
s = slide("讯飞交付什么：三项成果，一套能力")
items = [
    ("葡语基础模型", "葡语政务问答与文本处理的基础能力", "巴方提供语料并参与评测"),
    ("应用 APP", "政务人员直接使用的葡语智能应用\n场景：{{示例：政策咨询问答、公文辅助起草}}", "巴方提供系统接口并确认业务"),
    ("领域模型 DLM", "面向具体业务问题的专用模型\n领域：{{示例：教育}}", "巴方选定领域、提供数据与专家"),
]
gap = 0.35
cw = (CW - 2 * gap) / 3
cy, ch = 1.15, 3.5
for i, (name, use, rnp) in enumerate(items):
    cx = L + i * (cw + gap)
    card(s, cx, cy, cw, ch, fill=PANEL2, line=LINE)
    dot(s, cx + 0.25, cy + 0.28, 0.42, str(i + 1))
    text(s, cx + 0.8, cy + 0.22, cw - 1.0, 0.55, name, size=16, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    label(s, cx + 0.25, cy + 1.05, cw - 0.5, "用途", color=RED, size=12)
    body(s, cx + 0.25, cy + 1.42, cw - 0.5, 0.95, use, size=12)
    label(s, cx + 0.25, cy + 2.4, cw - 0.5, "巴方参与", color=RED, size=12)
    body(s, cx + 0.25, cy + 2.77, cw - 0.5, 0.6, rnp, size=12)
card(s, L, 4.95, CW, 1.0, fill=REDT)
label(s, L + 0.3, 5.05, 3, "项目内知识转移", color=RED, size=14, h=0.4)
body(s, L + 0.3, 5.47, CW - 0.6, 0.45, "巴方人员在参与以上三项工作的过程中，学会数据处理、模型训练与调优、应用开发与运行。", size=12.5)
note(s, 6.2, "以上四项均在当前方案范围内；范围口径以 SOW 与报价核对为准。第 3–5 页逐项展开。")
notes(s, "先回答最基本的问题：讯飞这部分最终形成什么，各有什么用，巴方要做什么。展示材料要标明是案例还是本项目交付。")

# ============================== 3 葡语基础模型 ==============================
s = deliverable_page(
    "葡语基础模型：怎么建、谁来做",
    "葡语政务问答与文本处理的基础能力，作为 APP 与领域模型的共同底座",
    [("数据准备", "收集与清洗葡语语料；巴方提供政务语料并落实授权"),
     ("训练与适配", "{{示例：在星火底座上做葡语增量训练与指令微调}}"),
     ("评测", "葡语通用基准 + 政务任务评测集，双方共同评审"),
     ("交付", "模型包、推理服务、评测报告与操作文档")],
    ["数据处理、训练与调优", "评测与报告", "打包、文档与交付"],
    ["提供语料并落实授权", "参与评测评审", "组织验收"],
    "交付物与验收", "{{示例：模型权重 + 推理服务镜像；评测达标后试运行 4 周，通过即验收}}",
)
notes(s, "让 RNP 明白模型怎样形成、讯飞负责什么、巴方要提供和参与什么。不预设从零训练，不展开技术选型依据。")

# ============================== 4 应用 APP ==============================
s = deliverable_page(
    "应用 APP：给谁用、怎么建",
    "政务人员直接使用的葡语智能应用；场景：{{示例：政策咨询问答、公文辅助起草、材料摘要}}",
    [("需求确认", "确认用户、场景与功能清单"),
     ("开发与本地化", "葡语界面、账号权限、模型调用"),
     ("系统对接", "对接巴方业务系统与数据"),
     ("测试上线", "联合测试、试运行、上线")],
    ["应用开发与本地化", "系统对接与上线", "使用培训与支持"],
    ["提供接口与测试数据", "确认业务流程", "组织测试与验收"],
    "功能与规模", "{{示例：问答、摘要、公文生成；支持 500 用户、并发 50；参考案例另附截图}}",
    visual="界面截图 / 演示视频\n{{示例位：现场演示一段完整操作}}",
)
notes(s, "从使用者角度讲：谁用、完成什么工作、得到什么。要有界面或视频，标明是案例还是本项目交付。听众不需要技术栈。")

# ============================== 5 领域模型 DLM ==============================
s = deliverable_page(
    "领域模型 DLM：选什么领域、怎么共同开发",
    "面向具体业务问题的专用模型；候选领域：{{示例：教育}}（按业务问题、数据可用性与专家投入选定）",
    [("选定领域", "双方共同确定业务问题与成功标准"),
     ("数据与专家", "巴方提供领域数据与授权，安排领域专家"),
     ("开发与适配", "基于葡语底座做领域适配与应用开发"),
     ("评价与验收", "专家评价、业务验收，移交更新方法")],
    ["数据规格与领域适配", "评测报告与问题修复", "移交更新方法与工具"],
    ["提出业务问题，指定负责人", "提供数据与授权，安排专家", "评价验收，接手后自行更新"],
    "交付物与验收", "{{示例：教育领域模型 + 领域应用 + 数据说明与更新指南；专家评价通过即验收}}",
)
notes(s, "领域是否选对、数据是否可用、专家能否参与评价，决定工作能否推进。候选领域标明候选状态。")

# ============================== 6 整体路径与时间线 ==============================
s = slide("整体路径：五个阶段，一条时间线")
steps = [
    ("数据准备", "语料收集清洗\n授权落实", "{{示例：M1–M2}}"),
    ("模型训练与评测", "训练葡语底座\n共同评审", "{{示例：M3–M5}}"),
    ("应用与领域开发", "APP 与领域模型\n巴方一起做", "{{示例：M5–M8}}"),
    ("部署与验收", "部署到华为平台\n试运行、验收", "{{示例：M9–M10}}"),
    ("交接与支持", "移交模型、应用、文档\n支持期响应", "{{示例：M11–M12}}"),
]
gap = 0.28
bw = (CW - 4 * gap) / 5
fy, fh = 1.15, 2.9
for i, (name, desc, when) in enumerate(steps):
    bx = L + i * (bw + gap)
    card(s, bx, fy, bw, fh, fill=REDT if i == 0 else PANEL)
    dot(s, bx + (bw - 0.5) / 2, fy + 0.28, 0.5, str(i + 1), size=14)
    text(s, bx + 0.1, fy + 0.9, bw - 0.2, 0.45, name, size=13.5, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, bx + 0.2, fy + 1.45, bw - 0.4, 0.85, desc, size=11.5, align="c")
    text(s, bx + 0.2, fy + 2.35, bw - 0.4, 0.4, when, size=12, color=BODY, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    if i < 4:
        arrow(s, bx + bw + 0.03, fy + 1.0)
band(s, 4.35, 0.62, "华为平台依赖", "训练与推理算力　·　部署环境与网络　·　数据中心运维衔接　　（单列，不计入讯飞责任）", title_w=1.8)
band(s, 5.2, 0.62, "关键里程碑", "{{示例：M2 语料到位　·　M5 底座通过评测　·　M8 应用联调通过　·　M10 验收　·　M12 交接完成}}", title_w=1.8)
note(s, 6.1, "阶段名称、数量与时间按实际方案确定。RNP 据此编制本方项目计划。")
notes(s, "回答“这些工作怎样安排在一起”。Mota 原话是三方责任，华为平台依赖单列，避免 RNP 把算力、环境的事算到讯飞头上。")

# ============================== 7 RNP 投入什么 ==============================
s = slide("RNP 需要投入的人员与资源")
rows = [
    ["项目负责人", "统筹巴方资源，对接讯飞与华为", "项目管理，英语沟通", "{{示例：1}}"],
    ["领域专家", "定义业务问题，评价领域成果", "领域业务经验", "{{示例：每领域 2}}"],
    ["数据工程师", "语料收集、清洗、标注", "Python，数据处理基础", "{{示例：3}}"],
    ["模型工程师", "参与训练、评测与调优", "Python，机器学习基础", "{{示例：3}}"],
    ["应用开发", "参与应用开发与系统对接", "软件开发经验", "{{示例：2}}"],
    ["运维工程师", "部署、运行与监控", "Linux 与容器基础", "{{示例：2}}"],
]
table(s, L, 1.2, [2.2, 4.4, 3.4, 2.33], ["角色", "做什么", "基础要求", "人数"], rows, row_h=0.5, header_h=0.45, size=12)
band(s, 5.15, 0.75, "资源", "数据与授权　·　环境与访问　·　业务系统接口　·　测试条件　　到位时间：{{示例：合同生效后 4 周内}}", title_w=1.1)
note(s, 6.15, "RNP 据此编制本方参与计划；角色按实际任务取舍。")
notes(s, "对 Mota 电话最直接的回应：巴方要安排多少人、什么人、什么时候投入。区分承担项目任务的人和只上课的人。")

# ============================== 8 知识转移 ==============================
s = slide("知识转移：两层安排")
gap = 0.4
cw = (CW - gap) / 2
cy, ch = 1.2, 4.3
card(s, L, cy, cw, ch, fill=GREEN_T)
chip(s, L + 0.3, cy + 0.3, 3.4, 0.36, "已包含在当前报价 · 不另收费", fill=GREEN)
text(s, L + 0.3, cy + 0.85, cw - 0.6, 0.5, "项目内知识转移", size=18, color="3B7A1A", bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
body(s, L + 0.3, cy + 1.5, cw - 0.6, ch - 1.7, [
    "方式：讲解、示范、共同操作、指导实践",
    "对象：参与项目的巴方人员，名额 {{示例：每阶段 10 人}}",
    "结果：能独立完成数据处理、模型微调、应用部署与运行",
    "语言：{{示例：英语授课，配葡语助教}}",
], size=12.5, bullet=True, bcolor="3B7A1A", sa=8)
cx = L + cw + gap
card(s, cx, cy, cw, ch, fill=ORANGE_T)
chip(s, cx + 0.3, cy + 0.3, 2.6, 0.36, "可选 · 单独报价", fill=ORANGE)
text(s, cx + 0.3, cy + 0.85, cw - 0.6, 0.5, "核心人才深入研修", size=18, color="B35400", bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
body(s, cx + 0.3, cy + 1.5, cw - 0.6, ch - 1.7, [
    "对象：少数有编程与模型基础的人员，人数 {{示例：6–8 人}}",
    "方向：模型训练与调优、蒸馏、评测",
    "形式与周期：{{示例：合肥现场 4 周 + 远程 8 周}}",
    "价格：{{示例：按人数与周期单独报价}}",
], size=12.5, bullet=True, bcolor="B35400", sa=8)
note(s, 5.8, "不覆盖：讯飞 SOW 以外的工作包。“已包含”口径以 SOW 与报价核对为准。")
notes(s, "第一层是项目内、已在报价中的知识转移；第二层是可选、单独报价的深入研修。学完能做什么要正面写清，避免客户期望学完就能自己做国家模型。")

# ============================== 9 本次确认 ==============================
s = slide("本次需要确认的事项")
rows = [
    ["1. 模型、应用与领域工作的目标与范围", "讯飞 + RNP", "{{示例：会后 1 周讯飞提交评测集与候选领域清单}}"],
    ["2. RNP 参与人员与资源", "RNP", "{{示例：会后 2 周 RNP 提交人员名单与到位时间}}"],
    ["3. 实施阶段与时间", "讯飞", "{{示例：会后 1 周提交实施计划}}"],
    ["4. 知识转移名额与能力目标", "讯飞 + RNP", "{{示例：按 SOW 核对名额，会后 1 周确认}}"],
    ["5. 可选研修是否需要", "RNP", "{{示例：RNP 内部评估后答复}}"],
    ["6. 范围与费用归属", "讯飞 + 华为", "{{示例：会后 1 周提交 SOW 与报价核对表}}"],
]
table(s, L, 1.2, [4.6, 1.8, 5.93], ["事项", "承接方", "结论 / 下一步"], rows, row_h=0.6, header_h=0.45, size=12.5)
note(s, 5.6, "会中记录结论，会后形成双方行动；材料中出现的内容不等于双方已承诺。", h=0.5)
notes(s, "让交流形成推进：哪些已对齐、哪些未定、谁负责补什么。这次会相当于锁定 SOW，争取技术部分当场基本谈定。")

drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
