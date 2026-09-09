# -*- coding: utf-8 -*-
"""
v9 替换页：只做两页（外加一页可选的问答目录），供手工替换进 v8。
  第 1 页：主汇报第 10 页“知识转移：RNP 学到什么、获得什么”，能力层改为三层（基础课 / 12 门专题课 / 高端知识转移）。
  第 2 页：附录 A11“高端知识转移菜单”，四个方向 + 定制课程 + 专家交流 + 访问学者计划，全部标为建议项，不写价格。
  第 3 页：问答目录（可选替换），加入 A11 一行。
本文件不做页内跳转链接：链接指向的是本文件内的页，复制到 v8 后会失效。
用法：python3 build_v9_kt_pages.py <template.pptx> <out.pptx>
"""
import sys

from pptx import Presentation

from pptx_helpers import (BLUET, BODY, CW, L, LINE, MUTE, PANEL, PANEL2, R, RED, REDT, TXT, WHITE,
                          body, card, chip, drop_first_slide, label, new_slide, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "v9-替换页-第10页与A11.pptx"
RNP_C = "1F4E79"

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]


def slide(title, size=22, ref=None):
    s = new_slide(prs, LAYOUT, title, size=size)
    if ref:
        x = R
        for code in reversed([c.strip() for c in ref.split("·")]):
            x -= 1.08
            chip(s, x, 0.32, 1.0, 0.3, "附录 " + code, fill=WHITE, color=RNP_C, size=9.5, line=LINE)
    return s


def appendix_slide(title, question):
    s = slide(title)
    chip(s, R - 3.9, 0.3, 3.9, 0.32, question, fill=WHITE, color=TXT, size=10, line=LINE)
    chip(s, R - 5.05, 0.3, 1.05, 0.32, "返回目录", fill=WHITE, color=RNP_C, size=9.5, line=LINE)
    return s


def thesis(s, content, y=1.05, h=0.62):
    body(s, L, y, CW, h, content, size=14, color=TXT, valign="m", ls=1.25)


def two_boxes(s, y, h, xf, rnp, size=12):
    half = (CW - 0.3) / 2
    for i, (name, color, fill, items) in enumerate([("讯飞负责", RED, REDT, xf), ("RNP 怎么参与", RNP_C, BLUET, rnp)]):
        cx = L + i * (half + 0.3)
        card(s, cx, y, half, h, fill=fill)
        label(s, cx + 0.25, y + 0.1, 3.0, name, color=color, size=13)
        body(s, cx + 0.25, y + 0.5, half - 0.5, h - 0.56, items, size=size, bullet=True, bcolor=color, sa=3)


def means(s, y, content, h=0.6):
    card(s, L, y, CW, h, fill=WHITE, line=RNP_C)
    body(s, L + 0.25, y + 0.05, CW - 0.5, h - 0.1, "**对 RNP 意味着**　" + content, size=12.5, color=RNP_C, valign="m")


# ============================== 第 10 页（替换） ==============================
s = slide("知识转移：RNP 学到什么、获得什么", size=21, ref="A9 · A10 · A11 · A6")
thesis(s, "12 项课程主题归为 4 个方向，讲师是参与本项目交付的讯飞工程师与方案架构师，研讨在交付的平台和模型上进行")
lw = CW * 0.58
cy, ch = 1.9, 2.95
card(s, L, cy, lw, ch, fill=PANEL)
label(s, L + 0.3, cy + 0.12, lw - 0.6, "学到什么：三个层次", color=TXT, size=13)
tiers = [
    ("基础课 · 奠基阶段", "大模型基础、项目架构、语料到服务全流程、安全与运维基础", "计划内"),
    ("12 门专题课 · 4 个方向的高级研讨", "模型与数据评测 · 软件栈与生命周期 · 规模化训练与推理 · 检索、智能体与安全", "计划内"),
    ("高端知识转移 · 怎么把这套系统做出来", "集群调度、模型训练、语料、推理四个方向；定制课程、专家交流、访问学者计划", "{{建议 · 另行约定}}"),
]
ty = cy + 0.55
for i, (name, desc, status) in enumerate(tiers):
    yy = ty + i * 0.78
    card(s, L + 0.3, yy, lw - 0.6, 0.7, fill=WHITE, line=LINE)
    text(s, L + 0.45, yy + 0.04, lw - 2.4, 0.28, f"{i + 1}  {name}", size=11.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, L + lw - 1.95, yy + 0.06, 1.55, 0.24, status, size=9, color=MUTE, align="r", valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, L + 0.45, yy + 0.34, lw - 0.9, 0.32, desc, size=10.5, color=BODY)
rx = L + lw + 0.3
rw = CW - lw - 0.3
card(s, rx, cy, rw, ch, fill=PANEL)
label(s, rx + 0.3, cy + 0.12, rw - 0.6, "获得什么：权重、平台、文档、支持", color=TXT, size=13)
body(s, rx + 0.3, cy + 0.6, rw - 0.6, ch - 0.7, [
    "**模型权重**，DLM 权重归 RNP",
    "**平台软件**和定制代码",
    "**全部文档**，英语或葡语",
    "**支持**：现场 12 个月 + 远程 60 个月",
], size=11.5, bullet=True, bcolor=RED, sa=6)
two_boxes(s, 5.05, 1.05,
          ["按批准的课程计划授课，使用批准的样本", "研讨、带教与移交材料；高端课程按需另行安排"],
          ["指定核心学员，安排参加研讨与实操", "按六类岗位的学习任务组织本方人员，见附录 A10"], size=11.5)
means(s, 6.28, "基础课与专题课在批准的课程计划内；{{高端课程、专家交流与访问学者计划为建议项，另行约定，见附录 A11}}")

# ============================== A11 高端知识转移菜单（新增） ==============================
s = appendix_slide("A11 · 高端知识转移：可选菜单", "问：能否学到这套系统是怎么做出来的？")
text(s, L, 1.02, CW, 0.5, "面向 RNP 核心研发人员与高校课题组；内容可从架构与原理讲到工程经验与源码实践；均不在本项目范围与报价内，范围、课时、人数、形式与费用另行约定",
     size=11, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
label(s, L, 1.62, 6, "四个方向 · 每个方向约 5 课时", color=TXT, size=12)
menu = [
    ("千 P 以上集群的算力调度与管理", "集群构建与管理、算力调度核心技术解析与方法实践"),
    ("百 B 参数以上模型的长期稳定训练", "大规模训练的稳定性保障与优化经验"),
    ("TB 级语料的构建与处理", "大规模语料的采集、清洗、处理方法与流水线"),
    ("日产万亿 Token 的推理服务", "Token 网关与路由分发、低时延推理的极致性能优化与稳定性保障"),
]
gap = 0.25
mw = (CW - 3 * gap) / 4
my, mh = 1.98, 1.85
for i, (name, desc) in enumerate(menu):
    mx = L + i * (mw + gap)
    card(s, mx, my, mw, mh, fill=PANEL)
    text(s, mx + 0.2, my + 0.14, mw - 0.4, 0.62, f"{i + 1}  {name}", size=12, color=TXT, bold=True, valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.2)
    body(s, mx + 0.2, my + 0.82, mw - 0.4, 0.66, desc, size=10.5, color=BODY)
    text(s, mx + 0.2, my + mh - 0.36, mw - 0.4, 0.26, "{{建议 · 另行报价}}", size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
label(s, L, 4.0, 6, "三种补充形式", color=TXT, size=12)
forms = [
    ("定制课程", "按 RNP 提出的主题定制，例如模型迁移、平台架构设计与实现原理、在华为算力上的训练适配"),
    ("一对一专家交流", "针对 RNP 工程师提出的具体问题做设计评审与咨询，按时长约定"),
    ("访问学者计划", "{{拟议：5 到 10 名 RNP 研究人员或在读博士到讯飞参与前沿模型研发，一期约半年；讯飞不收取费用，人选需面试，差旅自理}}"),
]
fw = (CW - 2 * 0.3) / 3
fy, fh = 4.36, 1.35
for i, (name, desc) in enumerate(forms):
    fx = L + i * (fw + 0.3)
    card(s, fx, fy, fw, fh, fill=BLUET if i == 2 else PANEL)
    text(s, fx + 0.2, fy + 0.12, fw - 0.4, 0.32, name, size=12, color=RNP_C if i == 2 else TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, fx + 0.2, fy + 0.5, fw - 0.4, fh - 0.58, desc, size=10.5, color=BODY)
text(s, L, 5.9, CW, 0.5, "以上与基础知识转移（四阶段 578 课时）分开，不在当前 SOW 与报价范围内；具体范围、日程与费用另行约定。口径：方案与建议，非承诺。",
     size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)

# ============================== 问答目录（可选替换，加 A11） ==============================
s = slide("附录：12 个问题，按 RNP 的顺序讨论")
thesis(s, "主汇报讲完后从这页开始；每个问题对应一页附录，点击进入，页面右上角可返回本页")
text(s, L, 1.8, CW, 0.28, "附录表格“口径”列：SOW = 已写入工作范围　·　方案 = 讯飞技术方案的做法　·　{{待确认}} = 实施前双方共同确定，蓝字", size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
items = [
    ("A0", "能否提供一页总览？", "总览表：四项成果、怎么建、谁来做、关键数字"),
    ("A1", "平台容量多大、如何验收？", "能力中心：容量与验收"),
    ("A2", "模型多大、之后如何升级？", "国家模型：规格与升级"),
    ("A3", "语料需要多少、谁提供、存放在哪？", "语料：规模、质量与数据控制"),
    ("A4", "APP 谁能用、能接入什么？", "国家 AI 应用：服务、容量与接口"),
    ("A5", "DLM 由谁做、做到哪一步？", "DLM：范围、投入与服务验收"),
    ("A6", "是否会被锁定、后续有哪些费用？", "交接、许可与运营支持"),
    ("A7", "每类人员多少、何时到位？", "RNP 人员配置明细"),
    ("A8", "每个阶段做什么、多少天、如何签认？", "知识转移：四阶段明细"),
    ("A9", "具体学什么、有哪些形式？", "知识转移：课程主题与交流形式"),
    ("A10", "每类岗位学什么？", "知识转移：岗位学习路径"),
    ("A11", "能否学到这套系统是怎么做出来的？", "高端知识转移：可选菜单"),
]
mw = (CW - 2 * 0.3) / 3
mh = 0.92
for k, (code, q, topic) in enumerate(items):
    r, c = divmod(k, 3)
    mx, my = L + c * (mw + 0.3), 2.25 + r * (mh + 0.18)
    card(s, mx, my, mw, mh, fill=WHITE, line=LINE)
    text(s, mx + 0.2, my + 0.1, 0.7, 0.34, code, size=12, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, mx + 0.85, my + 0.1, mw - 1.05, 0.34, q, size=12, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, mx + 0.2, my + 0.5, mw - 0.4, 0.32, topic, size=10, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)

drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
