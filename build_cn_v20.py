# -*- coding: utf-8 -*-
"""
讯飞 × RNP 技术交流 PPT v20：15 页主汇报 + 问答目录页 + 13 页附录，基于华为简报模板。
相对 v11（对齐华为 × 讯飞联合人才培养方案 RNP_AI_Talent_20260909）：
  · 新增附录 A12「产品使用培训」：6 个主题、每批 12 名运维工程师，随陪伴运维启动，不计入四阶段知识转移课时（SOW 7.5 / 8.3 / 9.4）。
  · A11 三种补充形式改为与联合方案一致：专家研讨与案例研究 / 一对一专家咨询 / 中国考察与技术交流；访问学者计划撤出材料，改为口头沟通。
  · A11 补两个准入条件：需先完成华为 AI Developer 全部课程；不在本项目范围与报价内，需另行预算。
  · 第 10 页三层状态由「计划内 / 建议·另行约定」改为更硬的「项目内已含 / 需另行预算」。
  · A9 三项交流形式改为指向 A11，中国考察注明与华为 6 天中国行分属不同安排，避免重复。
  · 第 4 页改为上下两个横带（并行两条线 / 汇合后三步），并新增 RNP 协调数据中心建成前语料处理算力的事项。
  · v13：第 4 页去掉与流程卡重复的底部两框。
  · v14：第 4、5、6 三页把建设步骤写实，每一步都写明讯飞做什么、RNP 做什么，取消与步骤重复的底部两框。
    v15：改回横向步骤卡（取消表格），每张卡内分「讯飞 / RNP / 产出」三格，文字压到短语级。
    v16：主汇报全面要点化——列表改自动编号（1. 2. 3.）。
    v17：句内并列项由圈码改为 1) 2) 3)（ASCII，投影更清楚，便于英文/葡语版沿用）；
    分号连接的并列动作全部拆成编号要点；第 9 页去掉图例行腾空间；第 10 页只做句内编号不拆行；
    附录每页底部增加「返回第 N 页」，直接回到引用该附录的主汇报页。
    v18：新增两页「最终能获得什么」——第 4 页国家模型、第 7 页 DLM，各自排在对应的「怎么建」页之前。
    两页同结构：能力格 + 关键词标签行 + 在本项目中的直接价值（价值链）+ 可验证成果 + 对 RNP 意味着。
    不使用「主权」表述，统一称国家模型；LGPD 由「符合」收为「配合 RNP 的合规要求，数据权限与审计归 RNP」；
    DLM 三个领域标明由 RNP 指定；DLM 的交付物按 v2 原文落到「可验证成果」行。主汇报 11 页 → 13 页。
    v19：按国家 APP 的 SOW（3.9 / 8.5 / 9.6 / 10.7）与 TR 软件栈响应补充。
    新增主汇报第 6 页「国家 AI 应用能获得什么」（学习 / 办公 / 生活三个使用方向）；
    新增主汇报第 14 页「对照 TR 五层」，第一层由华为承担；
    A4 由泛泛规格改为「功能清单与边界」（含文档翻译在范围内、独立语音翻译不含）；
    新增 A13「国家 AI 应用：交付物与验收」、A14「支持巴西自主研发：TR 要求对照」。
    附录 13 页 → 15 页，导航条相应收窄。主汇报 13 页 → 15 页。
    v20：主汇报重排为「价值块（第 3–7 页：四项成果 + 知识转移成果）+ 做法块（第 8–15 页）」，
    决策层在场时先讲完价值；价值页右上角加「怎么建 · 第 N 页」正向跳转。
    知识转移学员口径统一为「两轮累计 260 人次」。
相对 v10（A 档视觉优化，内容口径不变）：
  · 第 4 页改为“两条线并行 → 汇合”版式，并新增讯飞在中国境内先行采集公开葡语语料、增强基础模型通用葡语能力、
    采集语料一并提交巴西方（口径：讯飞方案，非 SOW 条款）；A3“通用语料”由待确认改为方案表述，A2 增加基座增强行。
  · 第 5 页改为“产品是什么 + 怎么落地”左右版式，第 6 页改为“按步骤的讯飞 / RNP 对照”版式，破除四页同模具。
  · 图例只在第 4 页出现一次（封面已有图例说明），第 5、6、9 页取消。
  · 附录去掉“口径”常量列（40 格 SOW / 2 格待确认），正文列加宽；非 SOW 行在“项目”里注明。
  · 全部 12 页附录底部加可点击导航条，当前页高亮，附录之间可直接互跳，不必回目录。
知识转移按赵祥确认口径：下列为单轮值，实施两个轮次，总课时与人数翻倍。
单轮：奠基 60 课时/10 天/50 人，跟岗 360 课时/45 天/20 人，指导实操 20 课时/3 天/20 人，
正式移交 168 课时/21 天/40 人，每轮合计 608 课时、约 79 个工作日。
两轮合计：1216 课时、约 158 个工作日、累计 260 人次（每轮 130 人次 × 2）。
另：第 5 页补 RNP 参与产品定义，第 9 页补真实训练任务与运维移交，第 10 页学到什么改三层，新增附录 A11 高端知识转移菜单。
相对 v7：按“几小时技术问答”的会议安排，主页右上标签可点击跳到对应附录页，附录页可点回问答目录；
分隔页改为 11 个问题的可点击目录；附录表格去竖线、浅表头；知识转移的“小时”改为“课时”。
版式规则（按三页样张确认）：
  · 每页三层：结论句（正文 14 磅，不加标签）/ 主体图形 / 底部一句“对 RNP 意味着”（蓝框）。
  · 颜色只表达谁做：红 = 讯飞主导，蓝 = RNP 主导，灰 = 双方或他方；不用绿、橙。
  · 保留“讯飞负责 / RNP 怎么参与”两框；取消标签带、脚注、主导方小标签、编号圆点；附录页码放标题右侧小标签。
  · 大数字不超过 22 磅。内容底线约 6.5，下方留白。
事实口径与 v6 相同（徐唯按最新 SOW 校验的 v2 + FINAL_TR_SOW 知识转移材料）。附录 11 页沿用 v6，仅 A9 状态色改为灰/蓝。备注留空。

用法：python3 build_cn_v20.py <template.pptx> <out.pptx>
"""
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import (BLUET, BODY, CW, GRAY, GRAY_T, L, LINE, MUTE, PANEL, PANEL2, R, RED, REDT, TXT, WHITE,
                          arrow, body, box, card, chip, dot, drop_first_slide, fill_tf, label, new_slide, note,
                          table, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "RNP-讯飞交流-v20.pptx"
RNP_C = "1F4E79"
SOW, PLAN, TBD = "SOW", "方案", "{{待确认}}"
XF, RN, BOTH = REDT, BLUET, PANEL

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]


LINKS = []          # (chip_shape, appendix_code)
MAIN_OF = {}        # 附录代号 -> 引用它的主汇报页（取第一处）
PAGENO = {}         # id(slide) -> 成品页码（Slide 不可哈希，模板首页最后会被删掉）
BACK2 = []          # (chip_shape, 主汇报页) 附录底部的「返回第 N 页」
FWD = []            # (chip_shape, key) 价值页右上角的「怎么建 · 第 N 页」
HOWTO = {}          # key -> 对应的「怎么建」页
APPX = {}           # code -> slide
BACK = []           # chip shapes on appendix pages -> index slide
INDEX = {}          # "index" -> divider slide


def slide(title, size=22, ref=None):
    s = new_slide(prs, LAYOUT, title, size=size)
    PAGENO[id(s)] = len(prs.slides) - 1
    if ref:
        codes = [c.strip() for c in ref.split("·")]
        x = R
        for code in reversed(codes):
            w = 1.0
            x -= w + 0.08
            c = chip(s, x, 0.32, w, 0.3, "附录 " + code, fill=WHITE, color=RNP_C, size=9.5, line=LINE)
            LINKS.append((c, code))
            MAIN_OF.setdefault(code, s)
    return s


APPX_CODES = ["A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14"]
NAV = []            # (cell_shape, appendix_code)


def appx_nav(s, cur):
    """附录页底部导航条：12 个页码，当前页红底，其余可点击直接互跳。"""
    y, h, gap = 6.60, 0.24, 0.055
    cw_ = 0.44
    x = L
    for code in APPX_CODES:
        if code == "A7":            # 交付内容 | 人与知识转移 之间留一个分组间隙
            x += 0.18
        cell = box(s, x, y, cw_, h, fill=RED if code == cur else WHITE,
                   line=None if code == cur else LINE, radius=0.35, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        fill_tf(cell.text_frame, code, size=8, color=WHITE if code == cur else MUTE,
                bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
        if code != cur:
            NAV.append((cell, code))
        x += cw_ + gap
    b = chip(s, 8.22, y, 1.30, h, "返回目录", fill=WHITE, color=RNP_C, size=9.5, line=LINE)
    BACK.append(b)
    tgt = MAIN_OF.get(cur)
    if tgt is not None:
        c = chip(s, 9.62, y, 1.30, h, "返回第 %d 页" % PAGENO[id(tgt)], fill=WHITE, color=RED, size=9.5, line=LINE)
        BACK2.append((c, tgt))


def howto(s, key, nref=0):
    """价值页右上角的正向跳转标签；目标页此时尚未创建，页码在末尾统一回填。"""
    w = 1.65
    c = chip(s, R - nref * 1.08 - w, 0.32, w, 0.3, "怎么建", fill=WHITE, color=RED, size=9.5, line=LINE)
    FWD.append((c, key))
    return c


def appendix_slide(title, question):
    code = title.split(" ·")[0].strip()
    s = slide(title)
    chip(s, R - 3.9, 0.3, 3.9, 0.32, question, fill=WHITE, color=TXT, size=10, line=LINE)
    APPX[code] = s
    appx_nav(s, code)
    return s


def thesis(s, content, y=1.05, h=0.62):
    body(s, L, y, CW, h, content, size=14, color=TXT, valign="m", ls=1.25)


def legend(s, y, x=None, extra=None):
    x = R - 4.1 if x is None else x
    for i, (f, t) in enumerate([(REDT, "讯飞主导"), (BLUET, "RNP 主导"), (PANEL, "双方")]):
        bx = x + i * 1.35
        box(s, bx, y + 0.06, 0.22, 0.16, fill=f, line=LINE)
        text(s, bx + 0.28, y, 1.0, 0.28, t, size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
    if extra:
        text(s, L, y, 7, 0.28, extra, size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)


def steps(s, y, h, items, gap=0.25, size=12):
    """items: (名称, 一句说明, 主导方底色)"""
    n = len(items)
    bw = (CW - (n - 1) * gap) / n
    for i, (name, desc, fill) in enumerate(items):
        bx = L + i * (bw + gap)
        card(s, bx, y, bw, h, fill=fill)
        text(s, bx + 0.2, y + 0.14, bw - 0.4, 0.4, f"{i + 1}  {name}", size=13, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, bx + 0.2, y + 0.62, bw - 0.4, h - 0.7, desc, size=size, ls=1.3)
        if i < n - 1:
            arrow(s, bx + bw + 0.02, y + 0.22, fill=LINE)


def two_boxes(s, y, h, xf, rnp, size=12, rnp_title="RNP 怎么参与"):
    half = (CW - 0.3) / 2
    for i, (name, color, fill, items) in enumerate([("讯飞负责", RED, REDT, xf), (rnp_title, RNP_C, BLUET, rnp)]):
        cx = L + i * (half + 0.3)
        card(s, cx, y, half, h, fill=fill)
        label(s, cx + 0.25, y + 0.1, 3.0, name, color=color, size=13)
        body(s, cx + 0.25, y + 0.5, half - 0.5, h - 0.56, items, size=size, bullet=True, num=True, bcolor=color, sa=3)



def step_cards(s, y, items, out_label="产出", gap=0.25, size=10.5, h=3.20):
    """横向步骤卡：卡内上红下蓝两格写双方动作，底部一行写产出 / 完成依据。内部高度按 h 自适应。"""
    n = len(items)
    bw = (CW - (n - 1) * gap) / n
    name_h, out_h, ig = 0.40, 0.50, 0.06
    blk = (h - name_h - out_h - ig - 0.10) / 2
    for i, (name, xf_t, rn_t, out_t) in enumerate(items):
        bx = L + i * (bw + gap)
        card(s, bx, y, bw, h, fill=WHITE, line=LINE)
        text(s, bx + 0.16, y + 0.05, bw - 0.32, 0.32, name, size=12.5, color=TXT, bold=True,
             valign="m", margin=(0, 0, 0, 0), sa=0)
        for k, (who, col, fl, t) in enumerate([("讯飞", RED, XF, xf_t), ("RNP", RNP_C, RN, rn_t)]):
            by = y + name_h + k * (blk + ig)
            card(s, bx + 0.14, by, bw - 0.28, blk, fill=fl)
            text(s, bx + 0.28, by + 0.03, 1.4, 0.22, who, size=9.5, color=col, bold=True,
                 valign="m", margin=(0, 0, 0, 0), sa=0)
            body(s, bx + 0.28, by + 0.26, bw - 0.56, blk - 0.30, t, size=size, ls=1.25,
                 bullet=isinstance(t, list), num=True, bcolor=col, sa=2)
        oy = y + name_h + 2 * blk + ig + 0.04
        text(s, bx + 0.28, oy, bw - 0.56, 0.20, out_label, size=8.5, color=MUTE,
             valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, bx + 0.28, oy + 0.20, bw - 0.56, 0.28, out_t, size=9.5, color=BODY)
        if i < n - 1:
            arrow(s, bx + bw + 0.015, y + name_h + blk * 0.5 - 0.11, fill=LINE)



def value_chain(s, y, h, items, label="在本项目中的直接价值"):
    """底部价值链：一条 3 段的链条，段间箭头。"""
    card(s, L, y, CW, h, fill=PANEL2, line=LINE)
    label_(s, L + 0.26, y + 0.08, 5.0, label)
    n = len(items)
    inner = CW - 0.5
    x0 = L + 0.25
    aw = 0.35
    w = (inner - (n - 1) * aw) / n
    cy = y + 0.44
    ch = h - 0.54
    for i, (t1, t2) in enumerate(items):
        cx = x0 + i * (w + aw)
        card(s, cx, cy, w, ch, fill=WHITE, line=LINE)
        text(s, cx + 0.16, cy + 0.05, w - 0.32, 0.28, t1, size=11, color=TXT, bold=True,
             valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, cx + 0.16, cy + 0.32, w - 0.32, ch - 0.38, t2, size=9.5, color=MUTE, ls=1.25)
        if i < n - 1:
            arrow(s, cx + w + 0.06, cy + ch / 2 - 0.11, fill=LINE)


def label_(s, x, y, w, t, color=TXT, size=12):
    text(s, x, y, w, 0.3, t, size=size, color=color, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)


def outcomes(s, y, items, label="可验证成果"):
    text(s, L, y, CW, 0.34, "**" + label + "**　" + "　·　".join(items),
         size=11, color=TXT, valign="m", margin=(0, 0, 0, 0), sa=0)

def means(s, y, content, h=0.6):
    card(s, L, y, CW, h, fill=WHITE, line=RNP_C)
    body(s, L + 0.25, y + 0.05, CW - 0.5, h - 0.1, "**对 RNP 意味着**　" + content, size=12.5, color=RNP_C, valign="m")


def band(s, y, h, title, content, fill=PANEL2, line=LINE, title_w=1.45, size=12, color=RED):
    card(s, L, y, CW, h, fill=fill, line=line)
    label(s, L + 0.25, y + (h - 0.34) / 2, title_w, title, color=color, size=12)
    body(s, L + 0.25 + title_w, y + 0.06, CW - 0.5 - title_w, h - 0.12, content, size=size, valign="m")


def spec_table(s, rows, y=1.15, colw=(2.35, 9.98), row_h=0.62, size=11, header=("项目", "内容")):
    """口径列已取消：全表默认 SOW 口径，非 SOW 行在“项目”中注明，待确认仍为蓝字。"""
    table(s, L, y, list(colw)[:2], list(header)[:2], [r[:2] for r in rows], row_h=row_h, header_h=0.4,
          size=size, header_size=11, header_fill="F2F2F2", header_color=TXT, vlines=False)

# ============================== 1 封面 ==============================
s = slide("")
text(s, 0.8, 1.5, 8.6, 0.4, "RNP × 科大讯飞 · 技术交流", size=14, color=MUTE, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 2.0, 8.6, 1.0, "国家模型、国家 AI 应用与领域模型", size=32, color=RED, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 3.0, 8.6, 0.6, "讯飞交付什么、双方怎么分工、能力如何转移", size=20, color=TXT, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 4.1, 8.6, 1.3, [
    "里约热内卢 · {{待确认：2026 年 9 月 X 日}}",
    "汇报：科大讯飞 {{待确认：主讲人、方案架构师}}",
    "讨论稿，内容以 SOW、BOQ 与技术方案为准；第 2 页一页看全，附录按需查阅",
], size=13, color=BODY, margin=(0, 0, 0, 0), ls=1.35, sa=4)
text(s, 0.8, 5.7, 8.6, 0.8, [
    "图例：{{蓝色文字}} = 待确认，实施前双方共同确定的数字或范围；黑色文字 = 已写入工作范围（SOW）或讯飞技术方案的做法",
    "版式：红底 = 讯飞主导，蓝底 = RNP 主导，灰底 = 双方",
], size=10.5, color=MUTE, margin=(0, 0, 0, 0), ls=1.3, sa=3)
bx, by, bw, bh, g = 9.9, 1.75, 2.9, 0.78, 0.18
for i, t in enumerate(["国家 AI 能力中心", "巴西葡语国家模型", "国家 AI 应用", "领域模型 DLM × 3"]):
    card(s, bx, by + i * (bh + g), bw, bh, fill=PANEL)
    text(s, bx + 0.3, by + i * (bh + g), bw - 0.5, bh, f"{i + 1}  {t}", size=13, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, bx, by + 4 * (bh + g), bw, 0.72, fill=RED)
text(s, bx, by + 4 * (bh + g), bw, 0.72, "知识转移", size=15, color=WHITE, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
# ============================== 2 一页看全 ==============================
s = slide("一页看全：讯飞交付的四项成果", ref="A0")
thesis(s, "四项成果都由讯飞交付，其中平台是三项成果的共用工作平台；每项一句话，细节见后面各页")
boxes = [
    ("国家 AI 能力中心", ["训练与应用开发平台", "三项成果都在其上完成", "未来 RNP 自主使用"], "3 个中心 + 审核平台"),
    ("巴西葡语国家模型", ["文本、语音识别与合成", "讯飞先行增强葡语基座", "权重交付 RNP"], "≥290B 参数"),
    ("国家 AI 应用", ["国民 APP 与运营平台", "面向 RNP 批准的用户", "讯飞交付，RNP 运营"], "1,000 路并发问答"),
    ("领域模型 DLM", ["3 个行业场景的模型", "RNP 训练，讯飞指导", "权重归 RNP"], "3 项 × 1 个场景"),
]
gap = 0.28
bw = (CW - 3 * gap) / 4
by, bh = 1.9, 2.72
for i, (name, one, num) in enumerate(boxes):
    bx = L + i * (bw + gap)
    card(s, bx, by, bw, bh, fill=PANEL)
    text(s, bx + 0.22, by + 0.16, bw - 0.44, 0.44, f"{i + 1}  {name}", size=14, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, bx + 0.22, by + 0.72, bw - 0.44, 1.1, one, size=11, ls=1.3, bullet=True, num=True, bcolor=RED, sa=4)
    text(s, bx + 0.22, by + 2.05, bw - 0.44, 0.5, num, size=15, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, L, 4.82, CW, 0.62, fill=PANEL2, line=LINE)
body(s, L + 0.25, 4.84, CW - 0.5, 0.58, "**知识转移**　两个轮次，合计 1216 课时、约 158 个工作日，累计 260 人次　　**交付后支持**　现场陪伴 12 个月 + 远程支持 60 个月", size=12, valign="m")
body(s, L, 5.66, CW, 0.7, [
    "术语　1) 文本大模型（NLP）处理文字的理解与生成　2) 语音识别（ASR）将语音转为文字　3) 语音合成（TTS）将文字转为语音　4) Token 是模型计量文本的单位，约等于一个词或词的一部分　5) Stage 1 / Stage 2 是两批算力对应的两个建设阶段",
], size=10.5, color=MUTE, ls=1.3)
# ============================== 3 成果一 · 国家 AI 能力中心 ==============================
s = slide("国家 AI 能力中心：三项成果的共用平台，也是 RNP 未来的工作平台", size=21, ref="A1 · A12")
thesis(s, "三项成果都在这个平台上完成；四个模块各对应 RNP 的一支团队，验收后由 RNP 自主管理，本次逐一演示约 10 分钟", h=0.48)
tops = ["巴西葡语国家模型", "国家 AI 应用", "领域模型 DLM"]
tw = 3.2
tx0 = L + (CW - (3 * tw + 2 * 0.4)) / 2
for i, t in enumerate(tops):
    chip(s, tx0 + i * (tw + 0.4), 1.62, tw, 0.34, t, fill=WHITE, color=TXT, size=11, line=LINE)
py, ph = 2.06, 2.26
card(s, L, py, CW, ph, fill=PANEL2, line=LINE)
label(s, L + 0.25, py + 0.08, 4, "国家 AI 能力中心（平台）", color=RED, size=12)
mods = [
    ("语料构建中心", "1) 导入 2) 清洗 3) 去重 4) 标注 5) 质量统计", "RNP 数据团队与葡语专家"),
    ("模型训练中心", "1) 数据集管理 2) 训练任务 3) 资源配置 4) 评测 5) 发布", "RNP 模型工程师；DLM 在此自主训练"),
    ("AI 应用创新中心", "1) 模型服务与 API 2) 应用配置 3) 运营统计", "RNP 应用与运营团队"),
    ("安全审核平台", "1) 审核策略 2) 敏感词库 3) 风险样例", "RNP 审核与合规团队"),
]
gap = 0.22
mw = (CW - 0.5 - 3 * gap) / 4
my, mh = py + 0.48, ph - 0.62
for i, (name, what, team) in enumerate(mods):
    mx = L + 0.25 + i * (mw + gap)
    card(s, mx, my, mw, mh, fill=WHITE, line=LINE)
    text(s, mx + 0.16, my + 0.1, mw - 0.32, 0.36, f"{i + 1}  {name}", size=12.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, mx + 0.16, my + 0.5, mw - 0.32, 0.62, what, size=10.5, ls=1.3)
    body(s, mx + 0.16, my + 1.16, mw - 0.32, 0.42, "**对应：**" + team, size=10.5, ls=1.3)
two_boxes(s, 4.48, 1.26,
          ["交付平台软件与授权、部署配置、API 说明、验收报告与运维文档", "部署平台、接通接口",
           "产品使用培训：6 个主题、每批 12 名运维工程师，见附录 A12"],
          ["提供环境和接口", "验收后自主管理，四个模块对应四支团队", "组织验收并接管日常运营"], size=11)
means(s, 5.90, "本次演示的每一步，都是跟岗和实操阶段 RNP 人员要自己完成的操作")
# ============================== 4 成果二 · 国家模型能获得什么 ==============================
s = slide("巴西葡语国家模型最终能获得什么？", size=22)
howto(s, "model")
thesis(s, "面向巴西语言、知识、价值与本地需求的国家级基础模型能力；权重、模型卡与评测报告交付 RNP", h=0.46)
qw, qgap = 4.55, 0.12
cw_c = CW - 2 * qw - 2 * qgap
qy, qh = 1.56, 1.36
cx_c = L + qw + qgap
card(s, cx_c, qy, cw_c, qh * 2 + qgap, fill=PANEL)
text(s, cx_c, qy + 0.42, cw_c, 0.42, "巴西葡语国家模型", size=15, color=RED, bold=True,
     align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
text(s, cx_c, qy + 0.86, cw_c, 0.30, "Brazilian Portuguese LLM", size=10.5, color=MUTE,
     align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
box(s, cx_c + 0.6, qy + 1.26, cw_c - 1.2, 0.02, fill=LINE)
body(s, cx_c + 0.3, qy + 1.38, cw_c - 0.6, 0.9,
     ["NLP 总参数 ≥290B", "语音识别、语音合成各 1 个", "本地语料 500 亿 Token"],
     size=10, color=BODY, align="c", ls=1.35)
quads = [
    (0, 0, "巴葡原生语言能力",
     ["理解巴西葡语、本地表达与{{区域差异}}", "提升葡语问答、生成与多轮交互能力"],
     "自然表达　|　口语理解　|　{{区域多样性}}"),
    (0, 1, "巴西本地知识与语境",
     ["结合巴西本地数据、知识与机构语境", "体现文化、社会与语言多样性"],
     "本地语料 500 亿 Token　|　巴西知识　|　文化多样性"),
    (1, 0, "价值与责任对齐",
     ["数据存放巴西境内，权限与审计归 RNP", "配合 LGPD 合规；安全与价值观数据 2 万条"],
     "偏见与安全评测　|　公平包容　|　透明可信"),
    (1, 1, "自主可控的 AI 底座",
     ["减少对外部通用模型的长期依赖", "权重交付 RNP，支撑本地数据控制与后续 DLM"],
     "权重归 RNP　|　本地部署　|　支撑国家优先领域"),
]
for k, (col, row, name, items, tags) in enumerate(quads):
    x = L if col == 0 else cx_c + cw_c + qgap
    y = qy + row * (qh + qgap)
    card(s, x, y, qw, qh, fill=PANEL2, line=LINE)
    text(s, x + 0.24, y + 0.10, qw - 0.48, 0.32, "%d  %s" % (k + 1, name), size=12.5, color=TXT,
         bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, x + 0.24, y + 0.46, qw - 0.48, 0.52, items, size=10.5, bullet=True, num=True, bcolor=RED, sa=2)
    box(s, x + 0.24, y + 1.02, qw - 0.48, 0.02, fill=LINE)
    text(s, x + 0.24, y + 1.06, qw - 0.48, 0.26, tags, size=9.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.52, 1.05, [
    ("巴西葡语国家模型", "更懂巴西的语言、知识与需求"),
    ("Health / Space / Biodiversity DLM", "RNP 指定的三个领域；基于国家模型或双方确认的基座"),
    ("面向巴西的 AI 应用与服务", "支持公共服务与国家优先领域"),
])
outcomes(s, 5.70, ["葡语能力测试", "本地知识与安全对齐测试", "模型服务 / API", "知识转移"])
means(s, 6.14, "效果与覆盖范围按双方确认的评测标准与边界确定，不由讯飞单方判定；权重交付后由 RNP 本地部署、自主演进", h=0.6)
# ============================== 5 成果三 · 国家 AI 应用能获得什么 ==============================
s = slide("国家 AI 应用最终能获得什么？", size=22)
howto(s, "app")
thesis(s, "面向 RNP 批准的授权用户的统一 AI 服务入口；讯飞交付 APP 与数据运营平台，内容、用户与运营由 RNP 负责", h=0.46)
aw = (CW - 2 * 0.3) / 3
ay, ah = 1.60, 2.45
acards = [
    ("国民 APP · 面向授权用户",
     ["统一 AI 服务入口：问答与检索、文档分析与翻译、写作笔记与演示",
      "巴葡语音输入与朗读；手机、电脑与 Web 多终端"],
     "统一入口　|　多终端　|　巴葡语音"),
    ("数据运营平台 · 面向 RNP 管理者",
     ["组织、用户、租户、角色与权限管理；应用注册、分类与发布",
      "使用日志查询与导出；访问量、活跃用户、使用趋势与反馈统计"],
     "用户与租户　|　日志　|　运营报表"),
    ("内置治理 · 面向合规与自主",
     ["接入本项目内容安全审核平台，在确认的交互环节识别风险并执行策略",
      "数据与日志在 RNP 指定的巴西环境；讯飞扫码登录不是本地运行的必需依赖"],
     "内容审核　|　权限与日志　|　本地运行"),
]
for i2, (name, items, tags) in enumerate(acards):
    ax = L + i2 * (aw + 0.3)
    card(s, ax, ay, aw, ah, fill=PANEL2, line=LINE)
    text(s, ax + 0.26, ay + 0.14, aw - 0.52, 0.36, "%d  %s" % (i2 + 1, name), size=12.5, color=TXT,
         bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, ax + 0.26, ay + 0.58, aw - 0.52, 1.30, items, size=10.5, bullet=True, num=True, bcolor=RED, sa=6, ls=1.3)
    box(s, ax + 0.26, ay + ah - 0.44, aw - 0.52, 0.02, fill=LINE)
    text(s, ax + 0.26, ay + ah - 0.38, aw - 0.52, 0.26, tags, size=9.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.25, 1.10, [
    ("模型服务与 API", "NLP、语音识别与合成、知识问答、内容安全审核"),
    ("国民 APP 与数据运营平台", "手机、电脑与 Web；问答 1,000 路并发"),
    ("面向公共服务的使用", "RNP 审批的授权用户，支撑公共服务"),
])
outcomes(s, 5.52, ["功能与配置说明", "多终端软件与授权", "数据运营平台", "接口集成测试记录", "测试报告与上线验收", "用户与管理员手册"])
means(s, 5.96, "讯飞不承担用户增长、内容生产、人工审核与业务托管运营；这些由 RNP 的产品、运营与审核团队承接", h=0.62)
# ============================== 6 成果四 · DLM 能获得什么 ==============================
s = slide("领域模型 DLM 最终能获得什么？", size=22)
howto(s, "dlm")
thesis(s, "RNP 指定三个领域，每项聚焦 1 个边界清晰的固定场景；{{场景由 RNP 提出，逐项签署确认书后启动}}", h=0.44)
dom_w = (CW - 2 * 0.3) / 3
for i4, t4 in enumerate(["Health　卫生", "Space　航天", "Biodiversity　生物多样性"]):
    chip(s, L + i4 * (dom_w + 0.3), 1.54, dom_w, 0.34, t4, fill=WHITE, color=TXT, size=11, line=LINE)
dw = (CW - 2 * 0.3) / 3
dy, dh = 1.98, 2.20
dcards = [
    ("读懂领域材料",
     ["理解该领域的专业术语与规范", "读懂该领域常用的文档结构"],
     "专业术语　|　领域规范　|　文档结构"),
    ("回答领域问题",
     ["基于领域知识与已接入资料回答", "检索、归纳与要点整理"],
     "知识问答　|　检索　|　要点整理"),
    ("支持领域工作",
     ["围绕约定场景提供分析与辅助判断", "材料归纳与报告草拟辅助"],
     "场景分析　|　辅助判断　|　报告草拟"),
]
for i, (name, items, tags) in enumerate(dcards):
    dx = L + i * (dw + 0.3)
    card(s, dx, dy, dw, dh, fill=PANEL2, line=LINE)
    text(s, dx + 0.26, dy + 0.16, dw - 0.52, 0.36, "%d  %s" % (i + 1, name), size=13.5, color=TXT,
         bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, dx + 0.26, dy + 0.62, dw - 0.52, 1.05, items, size=11.5, bullet=True, num=True, bcolor=RED, sa=6, ls=1.3)
    box(s, dx + 0.26, dy + dh - 0.44, dw - 0.52, 0.02, fill=LINE)
    text(s, dx + 0.26, dy + dh - 0.38, dw - 0.52, 0.26, tags, size=9.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.34, 1.10, [
    ("巴西葡语国家模型", "作为基座，或双方确认的其他基座"),
    ("Health / Space / Biodiversity DLM", "RNP 指定的三个领域，各 1 个固定场景"),
    ("面向公共服务的应用", "支撑 RNP 的公共服务与国家优先领域"),
])
outcomes(s, 5.62, ["DLM 方案", "指导记录", "训练与评测分析报告", "服务验收确认书", "模型权重归 RNP"])
means(s, 6.04, "训练由 RNP 执行、讯飞提供辅助训练技术服务；模型提供分析与辅助判断，不替代专业判断与决策", h=0.62)
# ============================== 7 成果五 · 知识转移学到什么、获得什么 ==============================
s = slide("知识转移：RNP 学到什么、获得什么", size=21, ref="A6 · A9 · A10 · A11")
howto(s, "kt", nref=4)
thesis(s, "12 项课程主题归为 4 个方向，讲师是参与本项目交付的讯飞工程师与方案架构师，研讨在交付的平台和模型上进行")
lw = CW * 0.58
cy, ch = 1.9, 2.95
card(s, L, cy, lw, ch, fill=PANEL)
label(s, L + 0.3, cy + 0.12, lw - 0.6, "学到什么：三个层次", color=TXT, size=13)
tiers = [
    ("基础课 · 奠基阶段", "1) 大模型基础 2) 项目架构 3) 语料到服务全流程 4) 安全与运维基础", "项目内已含"),
    ("12 门专题课 · 4 个方向的高级研讨", "1) 模型与数据评测 2) 软件栈与生命周期 3) 规模化训练与推理 4) 检索、智能体与安全", "项目内已含"),
    ("高端知识转移 · 怎么把这套系统做出来", "1) 集群调度 2) 模型训练 3) 语料 4) 推理；另有专家研讨、一对一咨询、中国考察", "{{需另行预算}}"),
]
ty = cy + 0.55
for i, (name, desc, status) in enumerate(tiers):
    yy = ty + i * 0.78
    card(s, L + 0.3, yy, lw - 0.6, 0.7, fill=WHITE, line=LINE)
    text(s, L + 0.45, yy + 0.04, lw - 2.4, 0.28, f"{i + 1}  {name}", size=11.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, L + lw - 1.95, yy + 0.06, 1.55, 0.24, status, size=9, color=MUTE, align="r", valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, L + 0.45, yy + 0.34, lw - 0.9, 0.32, desc, size=10, color=BODY)
rx = L + lw + 0.3
rw = CW - lw - 0.3
card(s, rx, cy, rw, ch, fill=PANEL)
label(s, rx + 0.3, cy + 0.12, rw - 0.6, "获得什么：权重、平台、文档、支持", color=TXT, size=13)
body(s, rx + 0.3, cy + 0.6, rw - 0.6, ch - 0.7, [
    "**模型权重**，DLM 权重归 RNP",
    "**平台软件**和定制代码",
    "**全部文档**，英语或葡语",
    "**支持**：现场 12 个月 + 远程 60 个月",
], size=11.5, bullet=True, num=True, bcolor=RED, sa=6)
two_boxes(s, 5.05, 1.05,
          ["按批准的课程计划授课，使用批准的样本", "研讨、带教与移交材料；高端课程按需另行安排"],
          ["指定核心学员，安排参加研讨与实操", "按六类岗位的学习任务组织本方人员，见附录 A10"], size=11.5)
means(s, 6.28, "基础课与专题课在项目范围内，不另收费；{{高端研修需先完成华为 AI Developer 全部课程并另行预算，见附录 A11}}")
# ============================== 8 国家模型：怎么建 ==============================
s = slide("巴西葡语国家模型：讯飞先行增强基座，RNP 提供本地语料", size=21, ref="A2 · A3")
HOWTO["model"] = s
thesis(s, "第 1 步两条线同时进行：讯飞在中国境内先行增强基座，RNP 在巴西建设本地语料；本地语料就绪、算力到位后进入第 2 步", h=0.52)
step_cards(s, 1.72, [
    ("1  语料与基座准备",
     ["中国境内采集公开葡语语料", "自有算力增强基座葡语能力", "公开语料一并提交巴西方"],
     ["巴西生产语料并取得授权", "规模约 500 亿 Token", "{{协调建成前的处理算力}}"],
     "增强后的基座 + 本地语料库"),
    ("2  模型训练",
     ["在增强后的基座上继续训练", "产出 NLP、ASR、TTS 模型", "按批准的技术方案与里程碑"],
     ["按讯飞规范持续供料", "配合抽样质检", "路线与次数书面确认"],
     "NLP ≥290B、ASR、TTS 各 1 个"),
    ("3  评测",
     ["出具评测报告", "覆盖回归、鲁棒性、偏见", "安全与文化适配"],
     ["共同定测试集与通过规则", "葡语与领域专家评审"],
     "评测报告，双方签认"),
    ("4  本地交付与持续优化",
     ["巴西环境部署", "交付权重与模型卡", "12 个月内 NLP 升 3 次", "语音各升 1 次"],
     ["组织验收与发布", "按 A2 的上限提供升级语料"],
     "权重、版本清单、模型卡、部署说明"),
], h=3.95)
means(s, 5.90, "基座增强先行，不占用巴西算力，也不必等本地语料齐备才开始；本地语料和算力到位即可进入训练，工期风险更小", h=0.6)
# ============================== 9 国家 AI 应用：怎么建 ==============================
s = slide("国家 AI 应用：讯飞交付软件，RNP 负责内容、用户与运营", size=21, ref="A4 · A13")
HOWTO["app"] = s
thesis(s, "面向 RNP 批准的授权用户；问答 1,000 路并发，语音识别与合成各 200 路，按服务分别配置并实测", h=0.48)
card(s, L, 1.56, CW, 1.06, fill=PANEL)
label(s, L + 0.26, 1.60, 4.0, "产品是什么", color=TXT, size=12.5)
for i, t in enumerate(["手机 APP", "电脑端", "数据运营平台"]):
    chip(s, L + 2.10 + i * 2.30, 1.62, 2.16, 0.30, t, fill=WHITE, color=TXT, size=10.5, line=LINE)
funcs = ["问答", "检索", "文档分析", "文档翻译", "语音输入与朗读", "笔记与写作"]
fw6 = (CW - 0.52 - 5 * 0.12) / 6
for i, t in enumerate(funcs):
    fx = L + 0.26 + i * (fw6 + 0.12)
    card(s, fx, 2.24, fw6, 0.32, fill=WHITE, line=LINE)
    text(s, fx, 2.24, fw6, 0.32, t, size=10.5, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
step_cards(s, 2.70, [
    ("1  明确需求",
     ["提供标准产品能力清单", "说明可配置项与边界"],
     ["参与产品定义与功能设计", "定用户范围、内容规则与品牌"],
     "需求与功能设计文档"),
    ("2  本地化",
     ["葡语界面与本地化配置", "接入 RNP 品牌与内容"],
     ["提供品牌素材与栏目内容", "{{葡语 PPT 与公文模板待验证}}"],
     "按指定浏览器与操作系统验证通过"),
    ("3  系统集成",
     ["完成接口对接与联调", "提供开发者文档与示例"],
     ["提供本地接口与账号凭证", "落实第三方 API 额度与许可"],
     "HTTP/REST/JSON/WebSocket 接口联调通过"),
    ("4  上线运营",
     ["Stage 1 初步交付", "Stage 2 扩展运行", "移交运维文档"],
     ["审批用户", "产品、运营与审核团队承接", "可在交付版本上自主演进"],
     "APP 与数据运营平台上线"),
], h=3.32)
means(s, 6.16, "产品团队在需求阶段即需到位，运营团队在 Stage 1 交付前到位；交付版本可由 RNP 自主演进", h=0.6)
# ============================== 10 DLM：怎么建 ==============================
s = slide("领域模型 DLM：RNP 训练，讯飞指导；3 项服务，每项 1 个场景", size=21, ref="A5")
HOWTO["dlm"] = s
thesis(s, "四个步骤都由 RNP 执行、讯飞在旁指导——这与国家模型相反，是 RNP 第一次自己训练模型", h=0.46)
band(s, 1.54, 0.46, "服务范围", "3 项行业模型辅助训练服务，每项 1 个边界清晰的固定场景；领域由 RNP 提出、讯飞评估可行性；模型权重由 RNP 训练产生并归 RNP", size=11)
step_cards(s, 2.08, [
    ("1  确定场景",
     ["评估领域可行性", "出具技术方案", "核实基座版本、规模与许可"],
     ["提出领域与场景", "明确用户、输入输出和边界"],
     "签署《DLM 场景及技术方案确认书》"),
    ("2  准备数据与路线",
     ["给出训练路线建议", "LoRA-SFT、SFT、GRPO 等", "每项确认 1 条主要路线"],
     ["提供训练、验证与安全数据", "提供独立评测数据与环境"],
     "数据与训练路线书面确认"),
    ("3  训练与评测",
     ["按约定工作量提供参数建议", "查看训练日志", "协助分析问题"],
     ["在模型训练中心执行训练", "完成评测"],
     "指导记录"),
    ("4  服务验收",
     ["基于 RNP 的实际结果", "完成分析报告"],
     ["核对指导记录", "完成业务验收"],
     "服务验收，与模型效果目标分开"),
], out_label="完成依据", h=3.55)
note(s, 5.72, "生产部署、API 开发与应用集成不在本服务内，可另行约定。", h=0.26, size=9.5)
means(s, 6.06, "这是 RNP 首次独立训练模型，需要自己的算法工程师和领域专家；跟岗和指导实操正是为此设计", h=0.6)
# ============================== 11 实施路径 ==============================
s = slide("实施路径：两个 Stage、五个步骤、三方分工", size=21)
thesis(s, "五个步骤按条件推进　1) 华为出算力　2) 讯飞出软件和模型　3) RNP 出数据和运营")
text(s, L, 1.62, CW, 0.26, "节点下方为各步骤的完成标志", size=9.5, color=MUTE, valign="m", align="r", margin=(0, 0, 0, 0), sa=0)
n = 5
x0, x1 = L + 1.05, R - 1.05
xs = [x0 + i * (x1 - x0) / (n - 1) for i in range(n)]
ly = 3.3
card(s, xs[0] - 0.9, 1.85, xs[2] - xs[0] + 1.2, 0.3, fill=PANEL2, line=LINE)
text(s, xs[0] - 0.75, 1.85, 5.5, 0.3, "Stage 1 · 基础部署、模型建设与 APP 初交付", size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, xs[2] - 0.3, 2.2, xs[4] - xs[2] + 1.2, 0.3, fill=WHITE, line=LINE)
text(s, xs[2] - 0.15, 2.2, 5.5, 0.3, "Stage 2 · 第二批算力就绪后扩展运行", size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
box(s, x0, ly - 0.02, x1 - x0, 0.05, fill=LINE)
nodes = [
    ("平台与语料准备", "1) 平台验收 2) 语料试批合格"),
    ("模型训练与评测", "1) 文本模型评测通过 2) 语音模型评测通过"),
    ("应用与 DLM", "1) APP Stage 1 交付 2) DLM 确认书签署"),
    ("验收", "分项验收报告签署"),
    ("移交与支持", "1) 移交确认签署 2) 进入陪伴运维"),
]
for i, (name, done) in enumerate(nodes):
    cx = xs[i]
    dot(s, cx - 0.22, ly - 0.22, 0.44, str(i + 1), size=12)
    text(s, cx - 1.1, ly - 0.72, 2.2, 0.4, name, size=12.5, color=TXT, bold=True, align="c", valign="b", margin=(0, 0, 0, 0), sa=0)
    body(s, cx - 1.1, ly + 0.32, 2.2, 0.6, done, size=11, color=MUTE, align="c")
box(s, xs[0], 4.5, xs[4] - xs[0], 0.05, fill=RED)
text(s, xs[0], 4.57, xs[4] - xs[0], 0.28, "知识转移四阶段贯穿全程；正式移交在陪伴运维结束前至少 3 个月启动", size=10, color=RED, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
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
# ============================== 12 RNP 的投入 ==============================
s = slide("RNP 的投入：SOW 规定的职责与三批到位的人员", size=21, ref="A7")
thesis(s, "SOW 规定由 RNP 承担　1) 生产葡语语料并取得授权　2) 自主训练和评测 DLM　3) 审批用户并运营 APP　4) 协调华为算力和环境，组织验收", h=0.7)
waves = [
    ("第 1 批 · 第 1 步起", "{{约 53 人}}", "1) 数据团队 2) 葡语专家 3) 合规审核 4) 项目统筹",
     "1) 采集与标注 2) 质检语料 3) 取得授权"),
    ("第 2 批 · 第 2 步起", "{{约 25 人}}", "1) 算法工程师 2) 测试工程师 3) 业务领域专家",
     "1) 参与国家模型训练评测 2) 随后自主训练 DLM"),
    ("第 3 批 · 第 3 步起", "{{约 33 人}}", "1) 应用开发与运营 2) 平台运维",
     "1) 承接 APP 运营 2) 用户审批 3) 平台运维"),
]
gap = 0.3
ww = (CW - 2 * gap) / 3
wy, wh = 2.0, 2.35
for i, (name, num, who, task) in enumerate(waves):
    wx = L + i * (ww + gap)
    card(s, wx, wy, ww, wh, fill=BLUET)
    text(s, wx + 0.25, wy + 0.15, ww - 0.5, 0.36, name, size=13, color=RNP_C, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, wx + 0.25, wy + 0.55, ww - 0.5, 0.5, num, size=20, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, wx + 0.25, wy + 1.12, ww - 0.5, 0.56, "**人员：**" + who, size=11, ls=1.3)
    body(s, wx + 0.25, wy + 1.74, ww - 0.5, 0.56, "**任务：**" + task, size=11, ls=1.3)
    if i < 2:
        arrow(s, wx + ww + 0.04, wy + 0.95, fill=LINE)
text(s, L, 4.5, CW, 0.5, "人数为参考量级，来自讯飞早期估算，各职能峰值可跨活动复用，不是同时在岗人数；最终按 SOW 分工与工作量确认。", size=10.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
means(s, 5.15, "首批组建数据团队与葡语专家，其次配置算法工程师，应用运营团队第三批到位；知识转移的核心学员从三批人员中指定：每轮奠基 50 人、跟岗与实操各 20 人、移交 40 人，共两轮，累计 260 人次", h=0.7)
# ============================== 13 知识转移：四个阶段 ==============================
s = slide("知识转移：两个轮次，合计 1216 课时、约 158 个工作日", size=21, ref="A8 · A12")
HOWTO["kt"] = s
thesis(s, "四个阶段每轮共 608 课时，实施两个轮次；移交后 RNP 能独立完成　1) 平台运维　2) 训练提交　3) 模型部署　4) 版本发布")
phases = [
    ("奠基", "60", "每轮 · 约 10 个工作日 · 50 人",
     ["大模型基础与项目架构", "语料到服务全流程", "安全与运维基础"], XF),
    ("跟岗", "360", "每轮 · 约 45 个工作日 · 20 人",
     ["讯飞操作，RNP 观察协助", "覆盖国家模型训练", "12 个月内的模型升级", "3 项 DLM 训练等真实任务"], XF),
    ("指导实操", "20", "每轮 · 约 3 个工作日 · 20 人",
     ["RNP 操作：数据处理、训练提交", "模型部署、监控与故障处理", "双方签实操记录"], RN),
    ("正式移交", "168", "每轮 · 约 21 个工作日 · 40 人",
     ["面向 RNP 运维团队", "交接生产环境与运维流程", "移交资产许可与运行手册", "验证独立操作，双方签确认"], RN),
]
gap = 0.25
pw = (CW - 3 * gap) / 4
py, ph = 1.80, 2.95
for i, (name, hours, meta, tasks, fill) in enumerate(phases):
    px = L + i * (pw + gap)
    card(s, px, py, pw, ph, fill=fill)
    text(s, px + 0.2, py + 0.12, pw - 0.4, 0.36, f"{i + 1}  {name}", size=13, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, px + 0.2, py + 0.5, pw - 0.4, 0.46, [f"**{hours}** 课时"], size=12, color=TXT, valign="m", margin=(0, 0, 0, 0), sa=0)
    r0 = s.shapes[-1].text_frame.paragraphs[0].runs[0]
    from pptx.util import Pt
    r0.font.size = Pt(22)
    r0.font.color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor.from_string(RED)
    text(s, px + 0.2, py + 0.98, pw - 0.4, 0.3, meta, size=11, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, px + 0.2, py + 1.30, pw - 0.4, ph - 1.36, tasks, size=10.5, ls=1.25, bullet=True, num=True, bcolor=RED, sa=2)
two_boxes(s, 4.88, 1.12,
          ["由参与本项目交付的讯飞工程师带教，检查关键步骤", "移交材料与独立操作验证"],
          ["指定稳定、具备基础的学员，阶段学员可重合", "安排时间，参与真实交付任务"], size=11.5)
means(s, 6.12, "卡片为单轮数字，两轮合计 1216 课时、约 158 个工作日；工作日是培训活动天数，分散在实施期内，不是连续工期。另有面向运维团队的产品使用培训，不计入这里的课时，见附录 A12", h=0.68)
# ============================== 14 对照 TR 五层 ==============================
s = slide("对照 RNP 的 TR：五层要求分别由谁落地", size=21, ref="A14")
thesis(s, "TR 把国家 AI 软件栈的要求分为五组；下表是承包商整体（华为 + 讯飞）的对应关系", h=0.46)
cols5 = [(L, 2.10), (2.70, 5.40), (8.20, 1.50), (9.80, 3.03)]
for (cx, cwd), nm, col in zip(cols5, ["层", "要求要点", "谁负责", "我们的对应"], [MUTE, MUTE, RED, RNP_C]):
    text(s, cx + 0.16, 1.66, min(cwd, 2.2), 0.28, nm, size=11, color=col, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
rows5 = [
    ("一  底层计算软件", "操作系统、固件与驱动；容器与运行环境；算子、数学库与编译工具；训练与推理框架；分布式通信、容错与编排",
     "华为", "华为底层软件栈；讯飞平台在其上运行"),
    ("二  数据模型与应用平台", "语料建设、模型训练、应用创新三个中心；国家模型；国家 AI 应用；模型服务与 API",
     "讯飞", "第 3–10 页全部内容；附录 A1–A5"),
    ("三  集成与运行保障", "接入 RNP 账号、安全与监控；第三方模型与工具迁移；基准测试；多租户隔离；故障恢复",
     "讯飞 + RNP", "附录 A1、A4、A6；{{接口与迁移范围逐项确认}}"),
    ("四  研发与人才能力", "研发合作计划；接口、工具与技术资料；研发技术支持；实践培训与知识转移；成果留存",
     "讯飞 + RNP", "第 7、13 页；附录 A8–A11、A14"),
    ("五  交付与长期自主", "资产与许可清单；五年维护与技术更新；期末移交；退出与迁移条款",
     "讯飞", "附录 A6；{{技术更新与期末移交条款对齐}}"),
]
for i3, (lay, req, who, ours) in enumerate(rows5):
    ry = 2.00 + i3 * 0.92
    card(s, cols5[0][0], ry, cols5[0][1], 0.84, fill=PANEL2, line=LINE)
    text(s, cols5[0][0] + 0.16, ry, cols5[0][1] - 0.3, 0.84, lay, size=11, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    card(s, cols5[1][0], ry, cols5[1][1], 0.84, fill=WHITE, line=LINE)
    body(s, cols5[1][0] + 0.16, ry + 0.06, cols5[1][1] - 0.32, 0.72, req, size=10, valign="m", ls=1.25)
    card(s, cols5[2][0], ry, cols5[2][1], 0.84, fill=GRAY_T if who == "华为" else (REDT if who == "讯飞" else PANEL))
    text(s, cols5[2][0], ry, cols5[2][1], 0.84, who, size=11, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    card(s, cols5[3][0], ry, cols5[3][1], 0.84, fill=RN)
    body(s, cols5[3][0] + 0.16, ry + 0.06, cols5[3][1] - 0.32, 0.72, ours, size=10, valign="m", ls=1.25)
# ============================== 15 需确认事项 ==============================
s = slide("本次需确认：范围边界与六项事项", size=21)
thesis(s, "验收依据　1) SOW　2) BOQ　3) 功能清单　4) 批准方案；新增需求按书面变更处理。会中记录结论，会后形成双方行动清单")
gap = 0.3
cw = (CW - 2 * gap) / 3
cy, ch = 1.85, 1.5
cols = [
    ("SOW 已列基线", ["平台、国家模型、AI 应用、安全产品", "3 项 DLM 辅助训练；2 项推理适配", "四阶段知识转移；12 + 60 个月支持"], PANEL2),
    ("范围外 · 另行约定", ["新增模型、场景、终端、接口", "额外语言、学员批次、专题培训", "扩容、续期、第三方许可费用"], WHITE),
    ("RNP 承担与前置条件", ["数据、授权、专家、环境", "接口与第三方 API 条件", "DLM 训练、验收、用户与运营"], BLUET),
]
for i, (h, items, f) in enumerate(cols):
    cx = L + i * (cw + gap)
    card(s, cx, cy, cw, ch, fill=f, line=LINE if f == WHITE else None)
    label(s, cx + 0.2, cy + 0.1, cw - 0.4, h, color=RNP_C if f == BLUET else TXT, size=12)
    body(s, cx + 0.2, cy + 0.46, cw - 0.4, ch - 0.5, items, size=11, bullet=True, num=True, bcolor=RNP_C if f == BLUET else GRAY, sa=3)
items = [
    ("资产、交付与许可矩阵", "RNP + 讯飞", "确定版本、数量、权利与费用"),
    ("模型与容量验收规则", "RNP + 讯飞", "确定测试集、通过规则与负载"),
    ("三项 DLM 场景", "RNP + 讯飞", "逐项确定场景、数据、路线，签署确认书"),
    ("资源与实施计划", "RNP / 华为 + 讯飞", "确定两批环境到位时间与里程碑月份"),
    ("培训与移交计划", "RNP + 讯飞", "确定课时、学员名单、语言与日程"),
    ("SOW 待对齐条款", "RNP + 讯飞 + 华为", "{{待确认：是否本次讨论}}"),
]
label(s, L, 3.5, 4, "六项确认事项", color=TXT, size=12.5)
iw = (CW - 2 * 0.3) / 3
iy, ih = 3.85, 1.1
for k, (name, who, nxt) in enumerate(items):
    r, c = divmod(k, 3)
    ix, yy = L + c * (iw + 0.3), iy + r * (ih + 0.2)
    card(s, ix, yy, iw, ih, fill=WHITE, line=LINE)
    text(s, ix + 0.2, yy + 0.1, iw - 0.4, 0.34, f"{k + 1}  {name}", size=12, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, ix + 0.2, yy + 0.44, iw - 0.4, 0.26, who, size=10, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, ix + 0.2, yy + 0.7, iw - 0.4, 0.36, nxt, size=11)
# ============================== 16 问答目录 ==============================
s = slide("附录：15 个问题，按 RNP 的顺序讨论")
INDEX["index"] = s
thesis(s, "主汇报讲完后从这页开始；每个问题对应一页附录，点击进入，页面右上角可返回本页")
text(s, L, 1.8, CW, 0.28, "附录内容除在“项目”中注明“方案”外，均为已写入工作范围（SOW）的口径；{{蓝字}} = 实施前双方共同确定的待确认项。每页底部导航条可在 15 页附录之间直接互跳", size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
menu = [
    ("A0", "能否提供一页总览？", "总览表：四项成果、怎么建、谁来做、关键数字"),
    ("A1", "平台容量多大、如何验收？", "能力中心：容量与验收"),
    ("A2", "模型多大、之后如何升级？", "国家模型：规格与升级"),
    ("A3", "语料需要多少、谁提供、存放在哪？", "语料：规模、质量与数据控制"),
    ("A4", "APP 有哪些功能、边界在哪？", "国家 AI 应用：功能清单与边界"),
    ("A5", "DLM 由谁做、做到哪一步？", "DLM：范围、投入与服务验收"),
    ("A6", "是否会被锁定、后续有哪些费用？", "交接、许可与运营支持"),
    ("A7", "每类人员多少、何时到位？", "RNP 人员配置明细"),
    ("A8", "每个阶段做什么、多少天、如何签认？", "知识转移：四阶段明细"),
    ("A9", "具体学什么、有哪些形式？", "知识转移：课程主题与交流形式"),
    ("A10", "每类岗位学什么？", "知识转移：岗位学习路径"),
    ("A11", "能否学到这套系统是怎么做出来的？", "高端知识转移：可选菜单"),
    ("A12", "平台交付后运维从哪里上手？", "产品使用培训：主题、天数与人数"),
    ("A13", "APP 交付什么、怎么验收？", "国家 AI 应用：交付物与验收"),
    ("A14", "TR 的研发支持要求怎么响应？", "支持巴西自主研发：要求对照"),
]
MENU = []
mw = (CW - 2 * 0.3) / 3
mh = 0.84
for k, (code, q, topic) in enumerate(menu):
    r, c = divmod(k, 3)
    mx, my = L + c * (mw + 0.3), 2.08 + r * (mh + 0.14)
    cardk = card(s, mx, my, mw, mh, fill=WHITE, line=LINE)
    MENU.append((cardk, code))
    text(s, mx + 0.2, my + 0.1, 0.7, 0.34, code, size=12, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, mx + 0.85, my + 0.1, mw - 1.05, 0.34, q, size=12, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, mx + 0.2, my + 0.46, mw - 0.4, 0.3, topic, size=10, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
# ============================== A0 总览表 ==============================
s = appendix_slide("A0 · 总览表：四项成果与关键数字", "问：能否提供一页总览？")
rows = [
    ["国家 AI 能力中心\n（平台）", "语料构建、模型训练、应用创新三个中心和一个安全审核平台；三项成果都在其上完成，未来由 RNP 自主使用",
     "软件实施、测试与培训", "基础环境、接口协调，验收后日常运营", "3 个中心\n1 个安全审核平台"],
    ["巴西葡语国家模型", "1 个 NLP 基础模型、1 个 ASR 模型、1 个 TTS 模型（男/女各 1 套发音人）；交付权重、模型卡、评测报告与部署说明",
     "训练、技术评测与交付；交付后 12 个月持续优化", "语料生产与授权、专家评审、验收", "NLP 总参数 ≥290B\n本地语料 50B Token"],
    ["国家 AI 应用", "国民 APP 多终端软件与数据运营平台，面向 RNP 批准的授权用户；问答、检索、文档、语音与写作",
     "应用交付、接口联调、平台配置", "内容、品牌、接口；用户审批与日常运营", "NLP 1,000 路\nASR / TTS 各 200 路\n安全审核 400 路"],
    ["领域模型 DLM", "3 项辅助训练服务，每项 1 个固定场景；模型由 RNP 训练，权重归 RNP",
     "方案、过程指导、结果分析", "数据、环境、执行训练与评测、业务验收", "3 项服务\n每项 1 个场景"],
]
table(s, L, 1.02, [1.75, 3.55, 2.35, 2.5, 2.18], ["成果", "RNP 获得什么", "讯飞负责", "RNP 负责", "关键数字"], rows,
      row_h=0.85, header_h=0.42, size=10.5, header_size=11, header_fill="F2F2F2", header_color=TXT, vlines=False)
band(s, 4.96, 0.52, "知识转移", "每轮 608 课时、约 79 个工作日：奠基 60、跟岗 360、指导实操 20、正式移交 168 课时，每轮学员 50 / 20 / 20 / 40 人；实施两个轮次，合计 1216 课时、约 158 个工作日", fill=REDT, line=None, size=11)
band(s, 5.60, 0.52, "交付后支持", "现场陪伴运维 12 个月（5×8）+ 远程技术支持 60 个月（7×24），首 12 个月重合", size=11)
note(s, 6.18, "范围以 SOW 与 BOQ 为准；容量按服务分别验证。", h=0.3, size=9.5)
# ============================== A1 能力中心容量 ==============================
s = appendix_slide("A1 · 国家 AI 能力中心：容量与验收", "问：平台容量多大、如何验收？")
spec_table(s, [
    ["软件组成", "语料构建、模型训练、应用创新三个中心及安全审核平台；配套租户、资源、队列和监控管理", SOW],
    ["资源纳管授权", "平台可纳管的加速卡授权上限 1,024 张；硬件数量以华为资源方案为准，扩容按 BOQ 确认", SOW],
    ["推理服务并发", "NLP 1,000 路、ASR 200 路、TTS 200 路、安全审核 400 路；按服务分别配置与实测，不相加为总并发", SOW],
    ["容量测试条件", "实施前确认硬件、模型版本、输入 / 输出长度、音频时长、请求分布、测试工具与通过规则", SOW],
    ["网络安全产品", "2 台堡垒机、1 台日志审计、1 台数据库审计、1 台漏洞扫描、2 台 WAF 及对应授权", SOW],
    ["验收方式", "核对 BOQ 与授权；验证部署、功能、容量、安全与许可；提交测试记录并整改阻断问题", SOW],
])
note(s, 5.4, "并发是按服务配置并实测的授权容量基线，不是同时峰值；1,024 张是平台纳管授权上限，不是本期硬件数量。", h=0.4, size=9.5)
# ============================== A2 模型规格与升级 ==============================
s = appendix_slide("A2 · 巴西葡语国家模型：规格与升级", "问：模型多大、之后如何升级？")
spec_table(s, [
    ["模型规模", "NLP 1 个，总参数 ≥290B；ASR 1 个，亿级参数；TTS 1 个，男 / 女各 1 套发音人", SOW],
    ["基座与权利", "交付基础模型实际权重及本地运行配套产物；基础权重与训练后权重分别标识，开放对象按资产与许可矩阵", SOW],
    ["基座通用葡语增强（方案）", "与本地语料生产并行：讯飞在中国境内投入自有算力，采集公开葡语语料并继续训练基础模型，提升其通用葡语能力，作为国家模型训练的起点；不占用巴西侧算力", PLAN],
    ["基座名称", "{{待确认：SOW 是否指明基座为讯飞星火基础模型；未指明则按“讯飞基础模型”表述}}", TBD],
    ["建设期训练", "按批准的模型技术方案执行；技术路线、次数、初始交付时间与里程碑实施前书面确认", SOW],
    ["持续优化期", "NLP、ASR、TTS 与安全审核平台全部初验后连续 12 个月；正式升级 NLP 3 次、ASR 1 次、TTS 1 次、安全平台 1 次", SOW],
    ["优化数据上限", "12 个月累计：NLP 文本 ≤10B Token、监督 ≤2 万条、安全对齐 ≤5 千条、ASR ≤1 千小时、TTS ≤10 小时；安全样例、词库、快修另有上限", SOW],
    ["导出与移植", "按已验证的模型、格式、框架与目标环境实施；受模型架构、算子和许可限制", SOW],
], row_h=0.55)
note(s, 6.02, "效果升级以 12 个月内约定次数为限，NLP 每次最多 3 项主要优化目标；建设期版本和缺陷整改不计入次数。", h=0.4, size=9.5)
# ============================== A3 语料 ==============================
s = appendix_slide("A3 · 巴西葡语语料：规模、质量与数据控制", "问：语料需要多少、谁提供、存放在哪？")
spec_table(s, [
    ["无监督文本", "50B Token，按批准 Tokenizer 统计有效数据，去重并剔除不合格内容", SOW],
    ["监督与安全数据", "监督调优 20 万条；安全与价值观 2 万条，按完整有效样本计量", SOW],
    ["语音数据", "ASR 1 万小时有效音频（含合格转写）；TTS 10 万条及不少于 100 小时高质量录音，配套发音词典与样本", SOW],
    ["分工与前置算力", "RNP 组织数据取得、授权和全量生产；讯飞提供平台、规范、试批指导及抽样质检。{{数据中心建成前用于语料采集与处理的算力，由 RNP 在巴西本地协调提供}}", SOW],
    ["质量规则", "批量生产前确认计量、抽样、质量阈值、批次与复验规则；不合格数据不计入有效规模", SOW],
    ["数据控制", "项目数据存储于 RNP 指定的巴西环境，按权限与审计管理；数据权利与使用范围遵循原始许可", SOW],
    ["讯飞公开语料投入（方案）", "讯飞在中国境内投入自有算力，采集并处理公开葡语语料，用于增强基础模型的通用葡语能力；该批公开语料一并提交巴西方。{{规模与提交方式实施前确认}}", PLAN],
], header=("项目", "内容（规模为最低参考量）", "口径"), row_h=0.56)
note(s, 5.6, "语料验收按有效数据定义与批准方案计量；模型效果按 A2 的评测规则另行验证。", h=0.4, size=9.5)
# ============================== A4 应用功能清单与边界 ==============================
s = appendix_slide("A4 · 国家 AI 应用：功能清单与边界", "问：APP 有哪些功能、边界在哪？")
rows = [
    ["对话与问答", "统一对话入口：文本输入、答案展示、重新生成、反馈与历史会话；通用问答与基于已接入知识库的问答，可显示来源引用；同一会话内多轮追问与历史续聊",
     "不保证回答所有问题或全部答案正确；行业专用对话机器人的定制开发不含；上下文受模型 Token 上限与会话策略限制"],
    ["检索", "在已接入的知识库、应用与授权文档内检索；可展示网络检索结果并跳转来源",
     "未接入系统、全部本地文件或任意互联网内容的检索不含；网络检索依赖外部服务"],
    ["文档分析与翻译", "附件解析与文档问答、摘要、提纲与要点提取；**对成功解析的文档文本提供翻译辅助**",
     "限于产品支持并成功解析的格式；超长文档整篇翻译、复杂版式保留、扫描件完整识别与专业译审不承诺；**独立语音翻译不含**"],
    ["写作与演示", "AI 写作、AI 笔记与 AI 演示创作，支持辅助生成、编辑、模板使用与导出",
     "{{葡语演示生成与本地公文模板需适配验证后纳入}}；专业设计制作与任意格式无损转换不含"],
    ["三个使用方向", "学习：问答与文档摘要；办公：笔记、写作、文档处理与应用快捷入口；生活：日常信息查询结合已接入的网络检索",
     "独立教学系统、课程体系与考试评分不含；OA 审批与业务系统自动化不含；预订、支付与生活服务对接不含"],
    ["巴葡语音与多终端", "接入本项目 ASR / TTS，提供语音输入与朗读；Web 端及本项目交付范围内的客户端，按指定浏览器与操作系统版本验证",
     "音频规格、时长、并发与发音人限于已交付语音服务范围；多终端不等于覆盖全部设备与操作系统，最终范围以 BOQ 为准"],
    ["认证与用户管理", "平台账号密码登录，支持经产品验证方式的外部认证；组织、用户、租户、角色与菜单功能权限管理",
     "**讯飞扫码登录不是巴西本地运行的必需依赖**；MFA 与任意身份系统适配默认不含；外部组织同步依赖已验证接口"],
    ["内容审核与安全", "接入本项目内容安全审核平台，在确认的文本交互环节识别风险并执行策略；账号认证、角色授权、租户访问控制与日志",
     "**星火云不替代独立的内容安全审核平台**；图片、音频、视频审核默认不含，不保证识别全部风险；不等同于完整网络安全体系或安全运营中心"],
    ["日志与运营报表", "产品实际记录的操作、客户端与服务日志查询与导出；访问量、活跃用户、使用趋势、收藏与反馈统计",
     "不承诺记录每一项业务活动；日志字段、留存期与容量在部署配置时确定；未接入系统统计、定制 BI 与计费结算不含"],
]
table(s, L, 1.02, [2.35, 5.45, 4.53], ["功能", "包含", "不含或限制"], rows,
      row_h=0.55, header_h=0.4, size=9.5, header_size=10.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
# ============================== A5 DLM 范围 ==============================
s = appendix_slide("A5 · 领域模型 DLM：范围、投入与服务验收", "问：DLM 由谁做、做到哪一步？")
spec_table(s, [
    ["数量与范围", "3 项 DLM 辅助训练技术服务；每项 1 个边界清晰的固定场景，领域由 RNP 提出、讯飞评估可行性", SOW],
    ["基础模型", "双方确认的既有基座；版本、规模、许可与工具链先核实", SOW],
    ["数据要求", "RNP 提供合法、足量、有代表性的训练、验证、安全及独立评测数据；规模按专项方案", SOW],
    ["训练路线", "每项确认 1 条主要路线，LoRA-SFT、SFT、GRPO 等按可行性选定", SOW],
    ["实施模式", "RNP 执行训练与评测；讯飞按约定工作量提供指导、分析与优化建议", SOW],
    ["权利与边界", "模型权重由 RNP 训练产生并归 RNP；讯飞交付方案、指导记录与分析报告；生产部署、API 开发与应用集成可另行约定", SOW],
    ["启动确认", "逐项签署《DLM 场景及技术方案确认书》，明确周期、工作量、交付形式与通过规则", SOW],
], row_h=0.56)
note(s, 5.6, "服务验收核对方案、指导活动和基于 RNP 实际结果的分析；模型效果目标与讯飞服务验收分开。", h=0.4, size=9.5)
# ============================== A6 交接许可运营 ==============================
s = appendix_slide("A6 · 交接、许可与运营支持", "问：是否会被锁定、后续有哪些费用？")
spec_table(s, [
    ["模型权利", "基础权重、训练后权重及 DLM 产物分别列明交付与许可；权重交付不含背景知识产权", SOW],
    ["平台与定制代码", "平台软件按 BOQ 与许可使用；定制开发部分按清单交付源码与修改权", SOW],
    ["永久使用", "标注永久使用的版本付清后永久有效，服务期结束不失效；维护、支持与升级按各自服务约定", SOW],
    ["导出与第三方适配", "按许可与技术可行性导出迁移，格式、框架与目标环境逐项确认；BOQ 列明 2 项第三方模型推理适配服务", SOW],
    ["现场与远程支持", "现场陪伴运维 12 个月（5×8）；远程技术支持 60 个月（7×24，L3/L4）；首 12 个月重合", SOW],
    ["首次响应", "P1 ≤30 分钟；P2 ≤60 分钟；P3 ≤2 小时；P4 ≤下 1 个工作日", SOW],
    ["后续费用与文档", "已列支持按 BOQ 执行；扩容、续期、第三方许可与范围外工作按明确条件另行确认；文档采用英语或巴西葡语", SOW],
], row_h=0.56)
note(s, 5.6, "响应时间是首次响应，不是恢复时限；远程支持后四年不含固定驻场。", h=0.4, size=9.5)
# ============================== A7 RNP 人员配置明细 ==============================
s = appendix_slide("A7 · RNP 人员配置明细：分三批到位", "问：每类人员多少、何时到位？")
rows = [
    ["项目与产品统筹", "项目经理 2、产品经理 4、PMO / 数据中心专家 3", "统筹需求、排期、资源和交付", "第 1 批 · 第 1 步起", "{{9}}"],
    ["语料生产与数据资源", "资源技术 6、数据质检 6、数据标注 20", "组织数据、标注及质量检查", "第 1 批 · 第 1 步起", "{{32}}"],
    ["葡语语言专家", "语言专家 6", "确认葡语质量、标注及语音规范", "第 1 批 · 第 1 步起", "{{6}}"],
    ["合规与审核", "审核专家 2、法律专家 2、审核负责人 2", "审核内容、安全及法律合规", "第 1 批 · 第 1 步起", "{{6}}"],
    ["业务与领域专家", "业务需求 3、领域需求 2、领域专家 2", "确认业务场景与领域效果", "第 2 批 · 第 2 步起", "{{7}}"],
    ["模型与测试工程", "算法工程师 12、测试工程师 6", "模型训练、评测及技术验证", "第 2 批 · 第 2 步起", "{{18}}"],
    ["应用开发与运营", "软件开发 15、运营工程师 10、架构师 2", "应用架构、开发集成及运营", "第 3 批 · 第 3 步起", "{{27}}"],
    ["平台运维", "O&M 工程师 6", "运行管理、部署及问题升级", "第 3 批 · 第 3 步起", "{{6}}"],
]
table(s, L, 1.15, [2.1, 3.6, 2.9, 2.0, 1.73], ["角色分组", "职能构成（参考名额）", "主要职责", "到位批次", "参考名额"], rows,
      row_h=0.46, header_h=0.4, size=10.5, header_size=11, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.4, ["三批小计：第 1 批 {{约 53}}、第 2 批 {{约 25}}、第 3 批 {{约 33}}，合计 {{111}} 为各职能峰值之和，可跨活动复用，不是同时在岗人数。",
              "参考名额来自讯飞早期按全流程交付模式的非约束估算；本 SOW 下 RNP 承担全量数据生产、DLM 训练与日常运营，人数按实际分工复核。"], h=0.7, size=9.5)
# ============================== A8 知识转移：四阶段明细 ==============================
s = appendix_slide("A8 · 知识转移：四阶段明细", "问：每个阶段做什么、多少天、如何签认？")
rows = [
    ["奠基\n（AI 基础）", "建立对机器学习、深度学习、自然语言处理和大模型的共同理解；覆盖语料准备、训练、评测、模型服务、应用创新、能力中心架构与运维基础",
     "讯飞授课", "60 课时\n约 10 个工作日", "50", "达到阶段要求\n（目标 ≥80%）"],
    ["跟岗\n（讯飞操作）", "在交付平台上观察并协助讯飞工程师完成数据集管理、训练任务、资源配置、模型评测、发布、服务与基础运维；内容跟随实际交付与运营任务",
     "讯飞操作\nRNP 观察协助", "360 课时\n约 45 个工作日", "20", "通过实操考核\n（目标 ≥70%）"],
    ["指导实操\n（RNP 操作）", "RNP 执行预先约定的任务：平台操作、数据处理、训练任务提交、模型部署、监控、告警与故障分级；讯飞指导并检查关键步骤",
     "RNP 操作\n讯飞指导检查", "20 课时\n约 3 个工作日", "20", "双方签署实操记录\n（目标 ≥50% 达标）"],
    ["正式移交\n（运营责任）", "按 SOW、BOQ 和最终验收清单实际交付的平台、许可资产、配置、权限和文档进行移交带教；验证 RNP 独立完成约定任务；确认遗留项与支持安排",
     "RNP 独立操作\n讯飞验证", "168 课时\n约 21 个工作日", "40", "双方签署移交确认"],
    ["合计", "课时跨主题只计一次；实施两个轮次，学员分批参加；阶段日期按批准的知识转移计划", "",
     "每轮 608 课时 / 约 79 工作日\n两轮 1216 课时 / 约 158 工作日", "每轮 50/20/20/40\n两轮累计 260 人次", ""],
]
table(s, L, 1.15, [1.45, 4.05, 1.5, 1.85, 1.25, 2.23], ["阶段", "学习目标与范围", "讯飞与 RNP 角色", "每轮课时 / 工作日", "每轮学员", "完成依据"], rows,
      row_h=[0.75, 0.75, 0.75, 0.75, 0.7], header_h=0.4, size=10, header_size=10.5, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.55, ["表中课时与人数为单轮口径，实施两个轮次；工作日是培训活动天数之和，不是连续工期；百分比是合格核心学员的培养目标，不是验收阈值；正式移交至少在一年陪伴运维结束前 3 个月启动，不等于 168 课时从该时点开始。",
              "两轮学员分批参加，各阶段不要求为同一批人；260 为两轮累计人次，同一人参加多个阶段分别计入，不是 260 名不同学员；正式移交面向运维团队，人数相应放宽；口径：SOW 与批准的知识转移计划。"], h=0.75, size=9.5)
# ============================== A9 知识转移：课程主题与交流形式 ==============================
s = appendix_slide("A9 · 知识转移：课程主题与交流形式", "问：具体学什么、有哪些形式？")
label(s, L, 1.02, 8, "四个方向 · 12 项 SOW 课程主题的分组，不增加课时", color=TXT, size=12)
tracks = [
    ("模型与数据评测", ["模型架构与训练路线：继续预训练、SFT、LoRA、偏好对齐", "葡语数据：来源、质量、血缘、隐私与生命周期治理", "评测：回归、鲁棒性、偏见、安全、幻觉与文化适配"], "目标：比较训练路线、检查数据质量、按约定基线评测模型"),
    ("软件栈与生命周期", ["软件层次与依赖；版本与兼容管理", "MLOps / LLMOps：实验跟踪、发布、监控与回滚", "软硬件与算法协同：负载与资源瓶颈"], "目标：理解依赖关系、生命周期运维和交付平台的能力"),
    ("规模化训练与推理", ["资源配置与检查点恢复；数据、张量、流水线与序列并行", "模型服务、批处理、KV 缓存；吞吐、时延与容量", "量化、剪枝、蒸馏；压缩与能效分析"], "目标：定位瓶颈，在效果、时延、吞吐与资源成本间权衡"),
    ("检索、智能体与安全", ["RAG 与图检索；混合检索与重排", "智能体：工具、规划、记忆与多智能体协作", "系统与模型安全：提示词、内容审核、供应链"], "目标：设计带来源核验和风险控制的知识服务与研究助手"),
]
gap = 0.25
tw = (CW - 3 * gap) / 4
ty, th = 1.4, 2.35
for i, (name, items, obj) in enumerate(tracks):
    tx = L + i * (tw + gap)
    card(s, tx, ty, tw, th, fill=PANEL)
    dot(s, tx + 0.16, ty + 0.14, 0.32, str(i + 1), size=10)
    text(s, tx + 0.56, ty + 0.1, tw - 0.7, 0.4, name, size=12, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, tx + 0.16, ty + 0.58, tw - 0.32, 1.25, items, size=9.5, bullet=True, bcolor=RED, sa=2)
    body(s, tx + 0.16, ty + 1.85, tw - 0.32, 0.46, obj, size=9.5, color=MUTE)
label(s, L, 3.9, 8, "四种交流形式", color=TXT, size=12)
formats = [
    ("高级 AI 研讨", "项目内已含", GRAY, "按批准的课程计划，在交付平台和模型上进行；内容与深度随课程计划"),
    ("中国 AI 考察", "{{见附录 A11}}", RNP_C, "访问讯飞团队与算力设施；华为方案已含 6 天中国行（100 人），此处为讯飞技术交流形式，避免重复安排"),
    ("一对一专家咨询", "{{见附录 A11}}", RNP_C, "针对 RNP 工程师提出的问题做设计评审；范围与时长事先约定，不含长期顾问"),
    ("AI 战略圆桌", "{{建议增加}}", RNP_C, "RNP 领导与讯飞架构师讨论模型路线、技术选型与能力规划；参与人与时间事先约定"),
]
fy, fh = 4.25, 1.3
for i, (name, status, color, desc) in enumerate(formats):
    fx = L + i * (tw + gap)
    card(s, fx, fy, tw, fh, fill=PANEL2 if i == 0 else BLUET)
    text(s, fx + 0.16, fy + 0.1, tw - 1.3, 0.36, name, size=11.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    chip(s, fx + tw - 1.1, fy + 0.13, 0.95, 0.28, status, fill=WHITE, color=TXT, size=8.5, line=color)
    body(s, fx + 0.16, fy + 0.52, tw - 0.32, 0.75, desc, size=9.5)
note(s, 5.7, "方向与主题对应：方向 1 = SOW 主题 1、7、8；方向 2 = 2、10、11；方向 3 = 3、4、9；方向 4 = 5、6、12。讲授方法不等于所有交付模型都支持该方法，也不包含新增优化任务；口径：SOW。交流形式中只有高级 AI 研讨在项目范围内，其余三项另行约定。", h=0.55, size=9.5)
# ============================== A10 知识转移：岗位学习路径 ==============================
s = appendix_slide("A10 · 知识转移：岗位学习路径", "问：每类岗位学什么？")
rows = [
    ["语料采集团队", "奠基、跟岗", "熟悉数据来源、授权、规范及平台处理流程", "教材、工具说明、代表性样例与处理指导", "课程、跟岗及问题记录"],
    ["语言与标注专家", "奠基、跟岗", "理解标注、语言质量、安全标签及语音规范", "规范讲解、样例分析、质量检查方法", "约定练习及规则确认"],
    ["模型工程师", "跟岗、指导实操", "提交训练任务、查看日志与评测，理解模型发布", "任务配置、日志分析、评测及发布指导", "约定实操结果与指导记录"],
    ["应用人员", "跟岗、指导实操", "掌握应用配置、用户权限、接口调用与运营统计", "应用平台、API 及运营功能使用指导", "约定功能及调用演练"],
    ["运维工程师", "指导实操、移交", "承接监控、事件、变更、常见故障与支持升级", "运行资料、代表性操作与自主验证带教", "自主操作及移交确认"],
    ["行业专家", "跟岗、指导实操", "理解场景、数据及模型效果评估方法", "评测方法、报告解读及内容边界说明", "参与记录和约定评测练习"],
]
table(s, L, 1.15, [1.8, 1.55, 3.4, 3.4, 2.18], ["角色", "培训阶段", "参与任务", "讯飞支持方式", "完成依据"], rows,
      row_h=0.58, header_h=0.4, size=10.5, header_size=11, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.2, "{{待确认：四阶段与实施五个步骤的对应关系}}；岗位分配和练习由批准的知识转移计划明确；专项 DLM 指导与实施培训不重复计入。", h=0.4, size=9.5)

# ============================== A11 高端知识转移：可选菜单 ==============================
s = appendix_slide("A11 · 高端知识转移：可选菜单", "问：能否学到这套系统是怎么做出来的？")
text(s, L, 0.98, CW, 0.56, "项目内的基线支持是文档开放与技术问题解答；以下为可选的主动式高端转移，面向 RNP 核心研发人员与高校课题组，内容可从架构与原理讲到工程经验与源码实践。**准入**：需先完成华为 AI Developer 全部课程　**预算**：不在本项目范围与报价内，需另行预算",
     size=10.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
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
    text(s, mx + 0.2, my + mh - 0.36, mw - 0.4, 0.26, "{{建议 · 另行约定}}", size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
label(s, L, 4.0, 6, "三种补充形式", color=TXT, size=12)
forms = [
    ("专家研讨与案例研究", "在选定方向上做专家研讨、工程案例复盘与实践讲解；内容结合 RNP 的实际环境定制，例如模型迁移、平台架构与实现原理、在华为算力上的训练适配"),
    ("一对一专家咨询", "针对 RNP 工程师提出的具体研发问题做技术问答、方案讨论与工程经验交流，按时长约定"),
    ("中国考察与技术交流", "可选形式，结合专家学习与现场技术交流；语言、行程与相关费用另行确定。{{与华为方案中的 6 天中国行（100 人）分属不同安排，避免重复}}"),
]
fw = (CW - 2 * 0.3) / 3
fy, fh = 4.36, 1.45
for i, (name, desc) in enumerate(forms):
    fx = L + i * (fw + 0.3)
    card(s, fx, fy, fw, fh, fill=BLUET if i == 2 else PANEL)
    text(s, fx + 0.2, fy + 0.12, fw - 0.4, 0.32, name, size=12, color=RNP_C if i == 2 else TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, fx + 0.2, fy + 0.5, fw - 0.4, fh - 0.58, desc, size=10, color=BODY)
text(s, L, 6.0, CW, 0.45, "以上与基础知识转移（四阶段两轮合计 1216 课时）分开，不在当前 SOW 与报价范围内；范围、课时、人数、形式与费用另行约定。口径：方案与建议，非承诺。",
     size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
# ============================== A12 产品使用培训（新增） ==============================
s = appendix_slide("A12 · 产品使用培训：交付后运维怎么上手", "问：平台交付后运维从哪里上手？")
text(s, L, 1.00, CW, 0.36, "讯飞对已交付的软件、模型与安全产品做操作演示、配置指导与移交说明；面向 RNP 运维团队，随陪伴运维启动，与四阶段知识转移并行、不重复计课时",
     size=11, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
rows = [
    ["语料与模型训练平台", "数据集流程、质量检查与血缘；容器镜像、训练任务、日志与模型产物；已交付的语音工具与断点续训", "2", "12"],
    ["算力管理与集群运维", "租户、角色、资源池、配额与调度；集群健康、节点状态、日志与告警；资源用量、事件信息与升级路径", "1", "12"],
    ["应用服务、API 与数据运营", "服务凭证、权限、限流与用量；已开通的知识问答与检索服务；API 调用与错误处理；门户配置、应用发布、用户与反馈", "1", "12"],
    ["基础模型与国家 AI 应用", "模型文件、加载、调用与配置；已知限制与运行检查；登录、问答、文档与语音交互；已开通的笔记、写作与演示功能", "2", "12"],
    ["内容安全与安全产品", "风险标签、审核规则、词库与快修；审核日志与已开通功能验证；所供安全产品的接入、策略与日志；基本操作与厂商支持渠道", "1", "12"],
    ["配置、文档与支持移交", "已部署版本、关键配置与文档；支持的备份与恢复方法；遗留问题、支持工单与升级路径", "2", "12"],
]
table(s, L, 1.42, [2.55, 7.53, 0.95, 1.30], ["主题", "培训内容", "工作日", "每批人数"], rows,
      row_h=0.58, header_h=0.4, size=10.5, header_size=10.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
band(s, 5.44, 0.46, "时间安排", "陪伴运维启动时开展产品使用培训，之后 12 个月陪伴运维期内穿插知识转移活动；正式移交在陪伴运维结束前至少 3 个月启动", size=10.5, title_w=1.2)
note(s, 5.98, ["六个主题是 SOW 第 7.5 节十二项产品使用主题的归并展示，不构成六门独立计量的课程；培训使用已交付版本与经授权的样本。",
               "{{批次、天数、参训人数、语言与线上/线下方式实施前双方确定；各主题工作日合计 9 天，联合材料横幅标注约 10 天，实施前确认}}　口径：SOW 7.5 / 8.3 / 9.4。"],
     h=0.5, size=9.5)

# ============================== A13 应用交付物与验收 ==============================
s = appendix_slide("A13 · 国家 AI 应用：交付物与验收", "问：APP 交付什么、怎么验收？")
rows = [
    ["应用范围与本地化确认", "国家 AI 应用功能与配置说明：实际功能、支持平台、界面本地化、已接服务与使用限制",
     "指定葡语界面与品牌配置完整，关键操作信息可理解、可使用"],
    ["多终端软件交付", "确认平台的软件部署包或客户端安装包、授权材料与版本说明；应用部署记录",
     "确认平台的软件完成部署或安装，可访问、可操作"],
    ["数据运营平台交付", "平台软件与授权材料；数据运营平台配置与使用手册",
     "应用配置、用户权限、日志、反馈与统计功能按用例验证，展示结果与实测数据一致"],
    ["接口与业务流程集成测试", "应用接口与集成测试记录，覆盖实际接入的模型、身份、知识等服务",
     "指定的模型与外部接口完成集成测试；认证、权限与已接内容审核流程通过验证"],
    ["测试与验收", "国家 AI 应用与数据运营平台测试报告；上线发布与验收确认材料",
     "明确清单内的功能通过批准的测试用例；代表性业务流程从用户输入到结果展示可跑通"],
    ["用户与管理员指导", "用户手册；管理员手册",
     "操作材料、测试结果、版本与发布记录已提交"],
]
table(s, L, 1.02, [2.55, 5.05, 4.73], ["服务范围", "交付物（SOW 8.5）", "最低验收标准（SOW 9.6）"], rows,
      row_h=0.72, header_h=0.4, size=9.5, header_size=10.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.86, ["验收流程：核对平台与版本 → 测试接口与功能 → 验证集成与权限 → 参照容量与模型测试 → 整改问题 → 确认应用。",
               "责任边界：RNP 提供品牌素材、知识内容与业务规则，确认葡语术语与界面，审批上线并运营用户、内容与业务；讯飞不承担用户增长、内容生产、人工审核与业务托管运营。口径：SOW 3.9 / 8.5 / 9.6 / 10.7。"],
     h=0.5, size=9)
# ============================== A14 研发支持要求对照 ==============================
s = appendix_slide("A14 · 支持巴西自主研发：TR 要求对照", "问：TR 的研发支持要求怎么响应？")
rows = [
    ["研发合作计划", "目标、支持的研究方向、参与机构、资源、责任、时间安排、指标与验收标准",
     "四个专题方向：集群构建与调度、大模型长时稳定训练、大规模语料构建与处理、高吞吐推理优化；{{与 RNP 优先任务对应后写入合作计划}}"],
    ["接口、工具与技术资料", "API、连接器、SDK 与工具链、数据格式与接口说明、日志、性能指标与使用文档",
     "平台已交付接口、开发者文档与示例，见附录 A4；{{按研究任务确认清单、版本、权限与使用条件}}"],
    ["研发技术支持", "支持开发、适配、测试与优化，涉及算子、软件库、运行环境、编译器与训练推理框架",
     "专家答疑、专题讲解、方案评审、操作演示、跟岗与定期技术讨论；{{项目内范围与另行约定的高端服务分开列明}}"],
    ["实践培训与知识转移", "面向 RNP 指定的研究人员、工程师与运维人员开展实践培训与知识转移",
     "四阶段知识转移，见第 7、13 页与附录 A8；课程主题见 A9，分岗位学习路径见 A10"],
    ["成果留存与进展报告", "有版本记录的技术资料库；定期报告研究进展",
     "课程材料、操作手册与案例分享；{{脚本、配置、测试结果与进展报告的归档位置与访问权限实施前确定}}"],
    ["开放对接与持续使用", "互操作与技术方案比较；专有依赖、许可限制、未来成本与可替代方案说明",
     "见附录 A6 交接、许可与运营支持；{{按组件说明开放范围、修改权与后续使用条件}}"],
]
table(s, L, 1.02, [2.35, 4.55, 5.43], ["TR 要求", "要求要点", "我们的对应"], rows,
      row_h=0.68, header_h=0.4, size=9.5, header_size=10.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.60, ["FP1–FP9 是 RNP 的研究方向（模型构建、模型服务、分布式运行、编译优化、集群管理、数据、跨机构协同、性能、系统安全）；本表说明承包商如何支持这些研究，不代表由讯飞独立承担全部方向。",
               "这组要求面向承包商整体：底层软件栈部分由华为承担，见第 14 页。口径：SOW 与实施前确认的研发合作计划。"],
     h=0.5, size=9)

for shp, code in LINKS + MENU + NAV:
    if code in APPX:
        shp.click_action.target_slide = APPX[code]
for shp in BACK:
    shp.click_action.target_slide = INDEX["index"]
for shp, tgt in BACK2:
    shp.click_action.target_slide = tgt
for shp, key in FWD:
    tgt = HOWTO[key]
    fill_tf(shp.text_frame, "怎么建 · 第 %d 页" % PAGENO[id(tgt)], size=9.5, color=RED, bold=True,
            align="c", valign="m", margin=(0.04, 0, 0.04, 0), sa=0)
    shp.click_action.target_slide = tgt
drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
