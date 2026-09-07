# -*- coding: utf-8 -*-
"""
讯飞 × RNP 技术交流 PPT（中文精简版，7 页），基于华为简报模板。
每页一个问题、一个图形、不超过五条要点；具体数字、名称、日期留空，
以蓝色“待填”标出，后续人工补充。

用法：python3 build_cn_simple.py <template.pptx> <out.pptx>
"""
import sys

from pptx import Presentation

from pptx_helpers import (BLUET, BODY, CW, GRAY_T, GREEN, GREEN_T, L, LINE, MUTE, ORANGE, ORANGE_T,
                          PANEL, PANEL2, R, RED, RED2, REDT, TXT, WHITE, arrow, body, card, chip, dot,
                          drop_first_slide, label, new_slide, note, notes, table, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "RNP-讯飞交流-中文简版.pptx"

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]


def slide(title):
    return new_slide(prs, LAYOUT, title)


# ============================== 1 封面 ==============================
s = slide("")
text(s, 0.8, 1.6, 8.5, 0.4, "RNP × 科大讯飞 · 技术交流", size=14, color=MUTE, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 2.1, 8.5, 1.0, "葡语大模型、应用与领域模型", size=36, color=RED, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 3.1, 8.5, 0.6, "讯飞交付什么、双方怎么分工、能力如何转移", size=20, color=TXT, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 4.3, 8.5, 1.2, [
    "里约热内卢 · {{待填：日期}}",
    "汇报：科大讯飞 {{待填：汇报人}}",
    "讨论稿，内容以 SOW、报价与技术方案为准",
], size=13, color=BODY, margin=(0, 0, 0, 0), ls=1.35, sa=4)
card(s, 0.8, 5.85, 8.5, 0.55, fill=BLUET)
text(s, 0.95, 5.85, 8.2, 0.55, "图例：{{蓝色文字}} = 待填或待核实（日期、人名、数字、场景、口径）", size=11.5, color=BODY, valign="m", margin=(0, 0, 0, 0), sa=0)
bx, by, bw, bh, g = 9.8, 2.0, 2.9, 0.9, 0.22
for i, t in enumerate(["葡语基础模型", "应用 APP", "领域模型 DLM"]):
    card(s, bx, by + i * (bh + g), bw, bh, fill=PANEL)
    dot(s, bx + 0.2, by + i * (bh + g) + (bh - 0.38) / 2, 0.38, str(i + 1))
    text(s, bx + 0.75, by + i * (bh + g), bw - 0.85, bh, t, size=15, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, bx, by + 3 * (bh + g), bw, 0.75, fill=RED)
text(s, bx, by + 3 * (bh + g), bw, 0.75, "知识转移", size=15, color=WHITE, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
notes(s, "本次交流讲四件事：讯飞交付什么、怎么做、双方怎么分工、巴方人员怎样学会并接手。依据 Mota 电话中 RNP 的诉求。")

# ============================== 2 讯飞交付什么 ==============================
s = slide("讯飞交付什么：三项成果，一套能力")
items = [
    ("葡语基础模型", "葡语政务问答与文本处理的基础能力", "巴方提供语料并参与评测"),
    ("应用 APP", "政务人员直接使用的葡语智能应用\n场景：{{待填}}", "巴方提供系统接口并确认业务"),
    ("领域模型 DLM", "面向 {{待填：领域}} 具体业务问题的专用模型", "巴方选定领域、提供数据与专家"),
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
note(s, 6.2, "以上四项均在当前方案范围内；{{范围口径以 SOW 与报价核对为准}}。")
notes(s, "先回答最基本的问题：讯飞这部分最终形成什么，各有什么用，巴方要做什么。展示材料要标明是案例还是本项目交付。")

# ============================== 3 怎么做 ==============================
s = slide("怎么做：一条路径，五个阶段")
steps = [
    ("数据准备", "收集与清洗葡语语料\n巴方提供政务语料与授权"),
    ("模型训练与评测", "训练葡语底座\n双方共同评审评测结果"),
    ("应用与领域开发", "开发 APP 与领域模型\n巴方技术人员一起做"),
    ("部署与验收", "部署到华为平台\n试运行、验收"),
    ("交接与支持", "移交模型、应用与文档\n支持期内响应"),
]
gap = 0.28
bw = (CW - 4 * gap) / 5
fy, fh = 1.3, 3.0
for i, (name, desc) in enumerate(steps):
    bx = L + i * (bw + gap)
    card(s, bx, fy, bw, fh, fill=REDT if i == 0 else PANEL)
    dot(s, bx + (bw - 0.5) / 2, fy + 0.3, 0.5, str(i + 1), size=14)
    text(s, bx + 0.1, fy + 0.95, bw - 0.2, 0.5, name, size=14, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, bx + 0.2, fy + 1.55, bw - 0.4, 1.3, desc, size=11.5, align="c")
    if i < 4:
        arrow(s, bx + bw + 0.03, fy + fh / 2 - 0.11)
card(s, L, 4.7, CW, 1.1, fill=PANEL2, line=LINE)
label(s, L + 0.3, 4.8, 3, "时间与里程碑", color=RED, size=13)
body(s, L + 0.3, 5.18, CW - 0.6, 0.55, "各阶段时间：{{待填}}　　关键里程碑：{{待填}}", size=12.5)
note(s, 6.1, "阶段名称与数量按实际方案确定。RNP 据此编制本方项目计划。")
notes(s, "回答“这些工作怎样安排在一起”。只给讯飞内部排期不够，RNP 需要据此编制自己的计划。")

# ============================== 4 谁做什么 ==============================
s = slide("谁做什么：讯飞、RNP、华为平台")
cols = [
    ("讯飞", REDT, RED, ["数据处理、模型训练与评测", "APP 与领域模型开发", "交付模型、应用与文档", "在项目中带巴方人员一起做", "支持期内技术支持"]),
    ("RNP", BLUET, "1F4E79", ["提供葡语语料并落实授权", "指定联合团队，投入人员", "安排领域专家，确认业务", "组织测试与验收", "接手运行与后续更新"]),
    ("华为平台", ORANGE_T, "B35400", ["训练与推理算力", "部署环境与网络", "数据中心运维衔接"]),
]
gap = 0.35
cw = (CW - 2 * gap) / 3
cy, ch = 1.2, 4.3
for i, (name, fill, color, bullets) in enumerate(cols):
    cx = L + i * (cw + gap)
    card(s, cx, cy, cw, ch, fill=fill)
    text(s, cx + 0.3, cy + 0.25, cw - 0.6, 0.5, name, size=18, color=color, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, cx + 0.3, cy + 0.95, cw - 0.6, ch - 1.1, bullets, size=12.5, bullet=True, bcolor=color, sa=8)
note(s, 5.8, "共同处理的事项（数据质量、实施条件）由双方一起负责，不单方面归给 RNP。")
notes(s, "Mota 原话是三方责任：我们的、你们的、他们的。华为平台依赖单列，避免 RNP 把算力、环境的事算到讯飞头上。")

# ============================== 5 RNP 投入什么 ==============================
s = slide("RNP 需要投入的人员与资源")
rows = [
    ["项目负责人", "统筹巴方资源，对接讯飞与华为", "项目管理，英语沟通", "{{待填}}"],
    ["领域专家", "定义业务问题，评价领域成果", "领域业务经验", "{{待填}}"],
    ["数据工程师", "语料收集、清洗、标注", "Python，数据处理基础", "{{待填}}"],
    ["模型工程师", "参与训练、评测与调优", "Python，机器学习基础", "{{待填}}"],
    ["应用开发", "参与应用开发与系统对接", "软件开发经验", "{{待填}}"],
    ["运维工程师", "部署、运行与监控", "Linux 与容器基础", "{{待填}}"],
]
table(s, L, 1.2, [2.2, 4.6, 3.5, 2.03], ["角色", "做什么", "基础要求", "人数"], rows, row_h=0.5, header_h=0.45, size=12)
card(s, L, 5.15, CW, 0.75, fill=PANEL2, line=LINE)
label(s, L + 0.3, 5.22, 1.6, "资源", color=RED, size=13)
body(s, L + 1.6, 5.27, CW - 1.9, 0.5, "数据与授权　·　环境与访问　·　业务系统接口　·　测试条件　　到位时间：{{待填}}", size=12.5)
note(s, 6.15, "RNP 据此编制本方参与计划；角色按实际任务取舍。")
notes(s, "对 Mota 电话最直接的回应：巴方要安排多少人、什么人、什么时候投入。区分承担项目任务的人和只上课的人。")

# ============================== 6 知识转移 ==============================
s = slide("知识转移：两层安排")
gap = 0.4
cw = (CW - gap) / 2
cy, ch = 1.2, 4.3
# 左：项目内
card(s, L, cy, cw, ch, fill=GREEN_T)
chip(s, L + 0.3, cy + 0.3, 3.4, 0.36, "已包含在当前报价 · 不另收费", fill=GREEN)
text(s, L + 0.3, cy + 0.85, cw - 0.6, 0.5, "项目内知识转移", size=18, color="3B7A1A", bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
body(s, L + 0.3, cy + 1.5, cw - 0.6, ch - 1.7, [
    "方式：讲解、示范、共同操作、指导实践",
    "对象：参与项目的巴方人员，名额 {{待填}}",
    "结果：能独立完成数据处理、模型微调、应用部署与运行",
    "语言：{{待填}}",
], size=12.5, bullet=True, bcolor="3B7A1A", sa=8)
# 右：可选
cx = L + cw + gap
card(s, cx, cy, cw, ch, fill=ORANGE_T)
chip(s, cx + 0.3, cy + 0.3, 2.6, 0.36, "可选 · 单独报价", fill=ORANGE)
text(s, cx + 0.3, cy + 0.85, cw - 0.6, 0.5, "核心人才深入研修", size=18, color="B35400", bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
body(s, cx + 0.3, cy + 1.5, cw - 0.6, ch - 1.7, [
    "对象：少数有编程与模型基础的人员，人数 {{待填}}",
    "方向：模型训练与调优、蒸馏、评测",
    "形式与周期：{{待填}}",
    "价格：{{待填}}",
], size=12.5, bullet=True, bcolor="B35400", sa=8)
note(s, 5.8, "不覆盖：讯飞 SOW 以外的工作包。{{“已包含”口径以 SOW 与报价核对为准}}。")
notes(s, "第一层是项目内、已在报价中的知识转移；第二层是可选、单独报价的深入研修。学完能做什么要正面写清，避免客户期望学完就能自己做国家模型。")

# ============================== 7 本次确认 ==============================
s = slide("本次需要确认的事项")
rows = [
    ["1. 模型、应用与领域工作的目标与范围", "讯飞 + RNP", "{{待填}}"],
    ["2. RNP 参与人员与资源", "RNP", "{{待填}}"],
    ["3. 实施阶段与时间", "讯飞", "{{待填}}"],
    ["4. 知识转移名额与能力目标", "讯飞 + RNP", "{{待填}}"],
    ["5. 可选研修是否需要", "RNP", "{{待填}}"],
    ["6. 范围与费用归属", "讯飞 + 华为", "{{待填}}"],
]
table(s, L, 1.2, [5.3, 2.0, 5.03], ["事项", "承接方", "结论 / 下一步"], rows, row_h=0.6, header_h=0.45, size=12.5)
note(s, 5.6, "会中记录结论，会后形成双方行动；材料中出现的内容不等于双方已承诺。", h=0.5)
notes(s, "让交流形成推进：哪些已对齐、哪些未定、谁负责补什么。这次会相当于锁定 SOW，争取技术部分当场基本谈定。")

drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
