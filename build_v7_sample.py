# -*- coding: utf-8 -*-
"""
v7 视觉样张：只做第 4、7、9 三页，用于新旧版式对比。
规则：每页三层（结论句 / 主体图形 / 对 RNP 意味着），颜色只表达谁做（红=讯飞，蓝=RNP，灰=双方或他方），
保留“讯飞负责 / RNP 怎么参与”两框，取消标签带、脚注、主导方小标签、编号圆点，底线上移到 6.2。
用法：python3 build_v7_sample.py <template.pptx> <out.pptx>
"""
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import (BLUET, BODY, CW, GRAY_T, L, LINE, MUTE, PANEL, PANEL2, R, RED, REDT, TXT, WHITE,
                          arrow, body, box, card, chip, dot, drop_first_slide, label, new_slide, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "v7-sample.pptx"
RNP_C = "1F4E79"
XF, RN, BOTH = REDT, BLUET, PANEL

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]


def slide(title, size=22, ref=None):
    s = new_slide(prs, LAYOUT, title, size=size)
    if ref:
        chip(s, R - 2.3, 0.32, 2.3, 0.3, "详见附录 " + ref, fill=WHITE, color=MUTE, size=9.5, line=LINE)
    return s


def thesis(s, content, y=1.05):
    body(s, L, y, CW, 0.62, content, size=14, color=TXT, valign="m", ls=1.25)


def legend(s, y, x=None):
    x = R - 4.1 if x is None else x
    items = [(REDT, "讯飞主导"), (BLUET, "RNP 主导"), (PANEL, "双方")]
    for i, (f, t) in enumerate(items):
        bx = x + i * 1.35
        box(s, bx, y + 0.06, 0.22, 0.16, fill=f, line=LINE)
        text(s, bx + 0.28, y, 1.0, 0.28, t, size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)


def steps(s, y, h, items, gap=0.25):
    """items: (名称, 一句说明, 主导方底色)"""
    n = len(items)
    bw = (CW - (n - 1) * gap) / n
    for i, (name, desc, fill) in enumerate(items):
        bx = L + i * (bw + gap)
        card(s, bx, y, bw, h, fill=fill)
        text(s, bx + 0.2, y + 0.14, bw - 0.4, 0.4, f"{i + 1}  {name}", size=13, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, bx + 0.2, y + 0.62, bw - 0.4, h - 0.7, desc, size=12, ls=1.3)
        if i < n - 1:
            arrow(s, bx + bw + 0.02, y + 0.22, fill=LINE)


def two_boxes(s, y, h, xf, rnp, size=12):
    half = (CW - 0.3) / 2
    for i, (name, color, fill, items) in enumerate([("讯飞负责", RED, REDT, xf), ("RNP 怎么参与", RNP_C, BLUET, rnp)]):
        cx = L + i * (half + 0.3)
        card(s, cx, y, half, h, fill=fill)
        label(s, cx + 0.25, y + 0.1, 2.5, name, color=color, size=13)
        body(s, cx + 0.25, y + 0.5, half - 0.5, h - 0.56, items, size=size, bullet=True, bcolor=color, sa=3)


def means(s, y, content):
    card(s, L, y, CW, 0.6, fill=WHITE, line=RNP_C)
    body(s, L + 0.25, y + 0.05, CW - 0.5, 0.5, "**对 RNP 意味着**　" + content, size=12.5, color=RNP_C, valign="m")


# ============================== 样张 A：第 4 页 国家模型 ==============================
s = slide("巴西葡语国家模型：讯飞训练，RNP 提供语料，权重交付 RNP", size=21, ref="A2 · A3")
thesis(s, "文本大模型（≥290B 参数）、语音识别、语音合成各 1 个；权重、版本清单、模型卡、评测报告和部署说明交付 RNP")
legend(s, 1.78)
steps(s, 2.1, 1.5, [
    ("语料准备", "RNP 生产葡语语料；讯飞提供规范、工具、试批和抽样质检", RN),
    ("模型训练", "在讯飞已有的基础模型上，用葡语语料继续训练", XF),
    ("评测", "测试集与通过规则双方确定；讯飞出评测报告，RNP 专家评审", BOTH),
    ("本地交付", "部署在巴西环境，交付权重；12 个月内文本模型升级 3 次、语音各 1 次", XF),
])
two_boxes(s, 3.85, 1.45,
          ["训练、评测、交付", "交付后 12 个月持续优化"],
          ["生产葡语语料并取得授权，约 500 亿 Token", "专家参与评测评审；组织验收与发布"])
means(s, 5.55, "第一阶段工作量最大的是语料；数据团队和葡语专家需最早到位")

# ============================== 样张 B：第 7 页 实施路径 ==============================
s = slide("实施路径：两个 Stage、五个步骤、三方分工", size=21)
thesis(s, "五个步骤按条件推进：华为出算力，讯飞出软件和模型，RNP 出数据和运营")
n = 5
x0, x1 = L + 0.9, R - 0.9
xs = [x0 + i * (x1 - x0) / (n - 1) for i in range(n)]
ly = 3.3
# Stage 两条细横条（两行，避免边框切过节点文字）
card(s, xs[0] - 0.9, 1.85, xs[2] - xs[0] + 1.2, 0.3, fill=PANEL2, line=LINE)
text(s, xs[0] - 0.75, 1.85, 5.5, 0.3, "Stage 1 · 基础部署、模型建设与 APP 初交付", size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, xs[2] - 0.3, 2.2, xs[4] - xs[2] + 1.2, 0.3, fill=WHITE, line=LINE)
text(s, xs[2] - 0.15, 2.2, 5.5, 0.3, "Stage 2 · 第二批算力就绪后扩展运行", size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
# 知识转移括线
box(s, xs[0], 4.5, xs[4] - xs[0], 0.05, fill=RED)
text(s, xs[0], 4.57, xs[4] - xs[0], 0.28, "知识转移四阶段贯穿全程；正式移交在陪伴运维结束前至少 3 个月启动", size=10, color=RED, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
# 时间线
box(s, x0, ly - 0.02, x1 - x0, 0.05, fill=LINE)
nodes = [
    ("平台与语料准备", "平台验收，语料试批合格"),
    ("模型训练与评测", "文本与语音模型评测通过"),
    ("应用与 DLM", "APP Stage 1 交付；DLM 确认书签署"),
    ("验收", "分项验收报告签署"),
    ("移交与支持", "移交确认签署，进入陪伴运维"),
]
for i, (name, done) in enumerate(nodes):
    cx = xs[i]
    dot(s, cx - 0.22, ly - 0.22, 0.44, str(i + 1), size=12)
    text(s, cx - 1.15, ly - 0.72, 2.3, 0.4, name, size=12.5, color=TXT, bold=True, align="c", valign="b", margin=(0, 0, 0, 0), sa=0)
    body(s, cx - 1.15, ly + 0.32, 2.3, 0.6, "**完成标志**\n" + done, size=10, color=MUTE, align="c")
# 三方分工
party = [("华为", TXT, GRAY_T, "算力、机房、网络与基础环境，由 RNP 协调"),
         ("讯飞", RED, REDT, "平台、模型、应用、知识转移与支持"),
         ("RNP", RNP_C, BLUET, "语料与授权、DLM 训练、验收与日常运营")]
pw = (CW - 0.6) / 3
for i, (name, color, fill, items) in enumerate(party):
    px = L + i * (pw + 0.3)
    card(s, px, 5.0, pw, 0.72, fill=fill)
    label(s, px + 0.22, 5.0 + 0.19, 0.8, name, color=color, size=13)
    body(s, px + 1.0, 5.0 + 0.06, pw - 1.2, 0.6, items, size=11.5, valign="m")
means(s, 5.9, "各步骤月份实施前双方确定，T0 为第一批算力就绪；RNP 据此编制人力计划，三批到位顺序见下页")

# ============================== 样张 C：第 9 页 知识转移四阶段 ==============================
s = slide("知识转移的四个阶段：578 小时、约 74 个工作日", size=21, ref="A8")
thesis(s, "一支核心团队，移交后能独立完成约定的平台运维、训练提交、模型部署与发布；30 人建立基础，10 人深入到独立操作")
legend(s, 1.78)
phases = [
    ("奠基", "30", "约 5 个工作日 · 30 人", "大模型基础、项目架构、语料到服务全流程、安全与运维基础", XF),
    ("跟岗", "360", "约 45 个工作日 · 10 人", "讯飞操作，RNP 观察协助：数据集、训练任务、评测、发布、基础运维", XF),
    ("指导实操", "20", "约 3 个工作日 · 10 人", "RNP 操作：数据处理、训练提交、模型部署、监控与故障处理；双方签实操记录", RN),
    ("正式移交", "168", "约 21 个工作日 · 10 人", "移交资产、许可、配置、接口文档与运行手册；验证独立操作，双方签确认", RN),
]
gap = 0.25
pw = (CW - 3 * gap) / 4
py, ph = 2.1, 2.3
for i, (name, hours, meta, tasks, fill) in enumerate(phases):
    px = L + i * (pw + gap)
    card(s, px, py, pw, ph, fill=fill)
    text(s, px + 0.2, py + 0.12, pw - 0.4, 0.36, f"{i + 1}  {name}", size=13, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, px + 0.2, py + 0.5, 1.6, 0.6, hours, size=30, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, px + 0.2 + (1.05 if len(hours) == 3 else 0.75), py + 0.66, 1.2, 0.4, "小时", size=12, color=TXT, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, px + 0.2, py + 1.1, pw - 0.4, 0.3, meta, size=11, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, px + 0.2, py + 1.45, pw - 0.4, ph - 1.5, tasks, size=11, ls=1.25)
two_boxes(s, 4.65, 1.2,
          ["由参与本项目交付的讯飞工程师带教，检查关键步骤", "移交材料与独立操作验证"],
          ["指定稳定、具备基础的学员，阶段学员可重合", "安排时间，参与真实交付任务"], size=11.5)
means(s, 6.0, "74 个工作日是培训活动天数，分散在实施期内，不是连续工期")

drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
