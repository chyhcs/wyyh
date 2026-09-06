# -*- coding: utf-8 -*-
"""
干部试用期转正述职 PPT 生成脚本 v4（两条线并列：战略专项·巴西 / 岗位本职·国际咨询）
用法：python3 build_review_deck.py <template.pptx> <out.pptx> [script.md]
"""
import copy
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "template.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "out.pptx"
SCRIPT_MD = sys.argv[3] if len(sys.argv) > 3 else "述职脚本30页.md"

# ───────────────────────── 设计令牌 ─────────────────────────
BLUE = "0C4BA1"
NAVY = "0A2A5E"
BLUE2 = "3F7BD6"
TINT = "EEF3FB"
TINT2 = "DCE7F8"
TXT = "2F3E4F"
MUTE = "6F7E8F"
LINE = "D3DCE8"
GREEN = "1F8F5F"
GREEN_T = "E4F4EC"
AMBER = "C57B14"
AMBER_T = "FBF0DD"
GRAYC = "7A8898"
GRAY_T = "F4F6F9"
WHITE = "FFFFFF"

SLIDE_W, SLIDE_H = 13.333, 7.5
L, R = 0.4, 12.93
CW = R - L
CT = 1.28

prs = Presentation(TEMPLATE)
LAYOUTS = {lay.name: lay for lay in prs.slide_layouts}
LAY_CONTENT = LAYOUTS["4_标题和内容"]

STATUS = {  # 统一状态标签
    "已形成": (GREEN, GREEN_T), "已提交": (GREEN, GREEN_T), "客户认可": (GREEN, GREEN_T),
    "推进中": (AMBER, AMBER_T), "待确认": (AMBER, AMBER_T), "计划": (BLUE2, TINT2),
    "已有项目使用": (GREEN, GREEN_T), "已有复用": (GREEN, GREEN_T), "研发输入": (BLUE2, TINT2),
    "试点中": (AMBER, AMBER_T), "样例积累": (AMBER, AMBER_T), "中止": (GRAYC, GRAY_T),
}


def rgb(h):
    return RGBColor.from_string(h)


# ───────────────────────── 脚本备注解析 ─────────────────────────
def load_script(path):
    try:
        txt = open(path, encoding="utf8").read()
    except OSError:
        return {}, ""
    txt = re.sub(r"fileciteturn\S+", "", txt)
    pages, cur, sec = {}, None, None
    for line in txt.split("\n"):
        m = re.match(r"^## 第(\d+)页｜", line)
        if m:
            cur = int(m.group(1))
            pages[cur] = {}
            sec = None
            continue
        if line.startswith("# ") and cur is not None:
            cur, sec = None, None
        if cur is None:
            continue
        m2 = re.match(r"^### (.+)$", line)
        if m2:
            sec = m2.group(1).strip()
            pages[cur][sec] = []
            continue
        if sec:
            pages[cur][sec].append(line)
    out = {p: {k: "\n".join(v).strip() for k, v in secs.items()} for p, secs in pages.items()}
    qa = ""
    if "# 五、管理层追问准备" in txt:
        qa = txt.split("# 五、管理层追问准备")[1].split("# 六、")[0].strip()
    return out, qa


SCRIPT, QA_TEXT = load_script(SCRIPT_MD)


def _clean(s):
    s = s.replace("**", "")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def script_notes(n, prefix=""):
    d = SCRIPT.get(n, {})
    parts = []
    if prefix:
        parts.append(prefix)
    if "完整讲述稿" in d:
        parts.append("【讲述稿】\n" + _clean(d["完整讲述稿"]))
    for k in d:
        if k.startswith("与") and "衔接" in k:
            parts.append("【衔接】" + _clean(d[k]))
    for k in d:
        if k.startswith("依据"):
            parts.append("【依据与归因（不上屏）】\n" + _clean(d[k]))
    return "\n\n".join(parts)


# ───────────────────────── 基础绘制 ─────────────────────────
def _parse_runs(s):
    parts = s.split("**")
    return [(p, i % 2 == 1) for i, p in enumerate(parts) if p]


def _norm_paras(content):
    if isinstance(content, str):
        return [{"text": t} for t in content.split("\n")]
    if isinstance(content, tuple) and len(content) == 2 and isinstance(content[1], dict):
        content = [content]
    out = []
    for item in content:
        if isinstance(item, str):
            out.append({"text": item})
        elif isinstance(item, tuple):
            d = dict(item[1])
            d["text"] = item[0]
            out.append(d)
        else:
            out.append(dict(item))
    return out


def _bullet(p, color):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.16))))
    pPr.set("indent", str(-int(Inches(0.16))))
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone"):
        for e in pPr.findall(qn(tag)):
            pPr.remove(e)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    c = etree.SubElement(buClr, qn("a:srgbClr"))
    c.set("val", color)
    etree.SubElement(pPr, qn("a:buSzPct")).set("val", "100000")
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def fill_tf(tf, content, size=11, color=TXT, bold=False, align="l", valign="t",
            margin=(0.06, 0.04, 0.06, 0.04), ls=1.12, sa=3, bullet=False, bcolor=None):
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margin]
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[valign]
    txBody = tf._txBody
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    for pd in _norm_paras(content):
        p = tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[pd.get("align", align)]
        p.line_spacing = pd.get("ls", ls)
        p.space_after = Pt(pd.get("sa", sa))
        if "sb" in pd:
            p.space_before = Pt(pd["sb"])
        for t, b in _parse_runs(pd["text"]):
            r = p.add_run()
            r.text = t
            r.font.size = Pt(pd.get("size", size))
            r.font.bold = bool(pd.get("bold", bold) or b)
            r.font.color.rgb = rgb(pd.get("color", color))
        if pd.get("bullet", bullet):
            _bullet(p, pd.get("bcolor", bcolor or BLUE))
    return tf


def text(slide, x, y, w, h, content, **kw):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_tf(tb.text_frame, content, **kw)
    return tb


def box(slide, x, y, w, h, fill=None, line=None, radius=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    return shp


def card(slide, x, y, w, h, fill=TINT, line=None, radius=0.05):
    return box(slide, x, y, w, h, fill=fill, line=line, radius=radius, shape=MSO_SHAPE.ROUNDED_RECTANGLE)


def chip(slide, x, y, w, h, label, fill=BLUE, color=WHITE, size=10, bold=True, line=None, radius=0.5):
    shp = box(slide, x, y, w, h, fill=fill, line=line, radius=radius, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    fill_tf(shp.text_frame, label, size=size, color=color, bold=bold, align="c", valign="m", margin=(0.04, 0, 0.04, 0), sa=0)
    return shp


def dot(slide, x, y, d, label, fill=BLUE, color=WHITE, size=11):
    shp = box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    fill_tf(shp.text_frame, label, size=size, color=color, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    return shp


def chevron(slide, x, y, w=0.16, h=0.32, fill=BLUE2):
    return box(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.CHEVRON)


def bar(slide, y, content, h=0.55, fill=NAVY, size=11.5, label=None):
    card(slide, L, y, CW, h, fill=fill, radius=0.12)
    if label:
        chip(slide, L + 0.15, y + (h - 0.3) / 2, 1.0, 0.3, label, fill=WHITE, color=fill, size=9.5)
        text(slide, L + 1.25, y, CW - 1.4, h, content, size=size, color=WHITE, bold=True, valign="m")
    else:
        text(slide, L + 0.2, y, CW - 0.4, h, content, size=size, color=WHITE, bold=True, valign="m")


def note(slide, y, content, h=0.36):
    text(slide, L, y, CW, h, content, size=9.5, color=MUTE, valign="m")


def stat(slide, x, y, w, h, big, label, fill=TINT, big_color=BLUE, big_size=18):
    card(slide, x, y, w, h, fill=fill)
    text(slide, x, y + 0.05, w, h * 0.55, big, size=big_size, color=big_color, bold=True, align="c", valign="m", margin=(0.04, 0, 0.04, 0), sa=0)
    text(slide, x, y + h * 0.56, w, h * 0.42, label, size=9.5, color=MUTE, align="c", valign="t", margin=(0.06, 0, 0.06, 0), sa=0)


def layer_card(slide, x, y, w, h, label, color, bullets, size=10.5):
    card(slide, x, y, w, h)
    chip(slide, x + 0.14, y + 0.14, 1.25, 0.32, label, fill=color, size=10)
    text(slide, x + 0.1, y + 0.56, w - 0.2, h - 0.62, [(b, {"bullet": True, "bcolor": color}) for b in bullets], size=size, sa=4)


def case_layers(slide, y, h, layers, size=10.5):
    n = len(layers)
    gap = 0.24
    w = (CW - gap * (n - 1)) / n
    for i, (label, color, bullets) in enumerate(layers):
        x = L + i * (w + gap)
        layer_card(slide, x, y, w, h, label, color, bullets, size=size)
        if i < n - 1:
            chevron(slide, x + w + 0.04, y + 0.16, fill=color)


def titled_card(slide, x, y, w, h, title, bullets, num=None, size=10.5, title_size=12.5, fill=TINT):
    card(slide, x, y, w, h, fill=fill)
    tx = x + 0.16
    if num is not None:
        dot(slide, x + 0.16, y + 0.16, 0.34, str(num), size=11)
        tx = x + 0.6
    text(slide, tx, y + 0.12, w - (tx - x) - 0.1, 0.42, title, size=title_size, color=NAVY, bold=True, valign="m", margin=(0.02, 0, 0.02, 0), sa=0)
    text(slide, x + 0.12, y + 0.6, w - 0.24, h - 0.66, [(b, {"bullet": True}) for b in bullets], size=size, sa=4)


def band(slide, x, y, w, h, label, bullets, cols=2, label_w=1.5, color=BLUE, size=10.5):
    card(slide, x, y, w, h)
    chip(slide, x + 0.14, y + (h - 0.36) / 2, label_w, 0.36, label, fill=color, size=10)
    cx = x + label_w + 0.35
    cw = w - label_w - 0.5
    if cols == 1:
        text(slide, cx, y + 0.08, cw, h - 0.14, [(b, {"bullet": True}) for b in bullets], size=size, sa=4, valign="m")
    else:
        half = (len(bullets) + 1) // 2
        colw = (cw - 0.2) / 2
        text(slide, cx, y + 0.08, colw, h - 0.14, [(b, {"bullet": True}) for b in bullets[:half]], size=size, sa=4, valign="m")
        text(slide, cx + colw + 0.2, y + 0.08, colw, h - 0.14, [(b, {"bullet": True}) for b in bullets[half:]], size=size, sa=4, valign="m")


def cell_borders(cell, bottom=LINE, top=None, left=None, right=None, w=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    for i, (tag, color) in enumerate([("a:lnL", left), ("a:lnR", right), ("a:lnT", top), ("a:lnB", bottom)]):
        ln = etree.Element(qn(tag))
        if color:
            ln.set("w", str(int(Pt(w))))
            ln.set("cap", "flat")
            ln.set("cmpd", "sng")
            ln.set("algn", "ctr")
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
            etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
        tcPr.insert(i, ln)


def table(slide, x, y, colw, header, rows, row_h=0.8, header_h=0.42, size=10, header_size=10.5,
          first_col_bold=True, zebra=True, valign="m", aligns=None, header_fill=BLUE, cell_pad=0.07, row_heights=None):
    nrows = len(rows) + 1
    ncols = len(colw)
    total_h = header_h + sum(row_heights) if row_heights else header_h + row_h * len(rows)
    gf = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(sum(colw)), Inches(total_h))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    for i, w in enumerate(colw):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(header_h)
    for r in range(1, nrows):
        tbl.rows[r].height = Inches(row_heights[r - 1] if row_heights else row_h)
    for c, h in enumerate(header):
        cell = tbl.cell(0, c)
        cell_borders(cell, bottom=None)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(header_fill)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        fill_tf(cell.text_frame, h, size=header_size, color=WHITE, bold=True, align="l", valign="m", margin=(cell_pad, 0.02, cell_pad, 0.02), sa=0)
    for r, row in enumerate(rows, start=1):
        rf = WHITE if (r % 2 == 1 or not zebra) else "F6F8FC"
        for c, content in enumerate(row):
            cell = tbl.cell(r, c)
            cell_borders(cell, bottom=LINE)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(rf)
            cell.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE}[valign]
            al = aligns[c] if aligns else "l"
            bold = first_col_bold and c == 0
            color = NAVY if bold else TXT
            fill_tf(cell.text_frame, content, size=size, color=color, bold=bold, align=al, valign=valign, margin=(cell_pad, 0.03, cell_pad, 0.03), sa=2, ls=1.08)
    return tbl


def role_tag(slide, label):
    w = min(4.0, 0.16 * len(label) + 0.5)
    chip(slide, R - w, 0.6, w, 0.36, label, fill=WHITE, color=BLUE, size=10, line=BLUE)


def page_no(slide, n):
    text(slide, 12.0, 7.02, 0.95, 0.3, str(n), size=9, color=MUTE, align="r", valign="m", margin=(0, 0, 0.05, 0), sa=0)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


PAGE = [1]  # P1 为模板封面


def content_slide(tag, title, role=None):
    s = prs.slides.add_slide(LAY_CONTENT)
    PAGE[0] += 1
    to_remove = []
    for ph in s.placeholders:
        y = ph.top / 914400
        if ph.placeholder_format.type == 1 or "标题" in ph.name:
            fill_tf(ph.text_frame, title, size=20, color=BLUE, bold=True, valign="m", margin=(0.05, 0, 0.05, 0), sa=0)
        elif y < 0.5:
            fill_tf(ph.text_frame, tag, size=14, color=WHITE, bold=True, valign="m", margin=(0.1, 0, 0.05, 0), sa=0)
        else:
            to_remove.append(ph)
    for ph in to_remove:
        ph._element.getparent().remove(ph._element)
    if role:
        role_tag(s, role)
    page_no(s, PAGE[0])
    return s


def replace_text(shape, lines, size=None):
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    rpr = copy.deepcopy(p0.runs[0]._r.get_or_add_rPr()) if p0.runs else None
    ppr = copy.deepcopy(p0._p.pPr) if p0._p.pPr is not None else None
    txBody = tf._txBody
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    for line in lines:
        p = etree.SubElement(txBody, qn("a:p"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        r = etree.SubElement(p, qn("a:r"))
        if rpr is not None:
            rp = copy.deepcopy(rpr)
            if size:
                rp.set("sz", str(int(size * 100)))
            r.append(rp)
        etree.SubElement(r, qn("a:t")).text = line


def walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from walk(sh.shapes)
        else:
            yield sh


SP = "战略专项"
BZ = "岗位本职"
cw2 = (CW - 0.24) / 2
w3 = (CW - 2 * 0.26) / 3
w4 = (CW - 3 * 0.2) / 4
w5 = (CW - 4 * 0.16) / 5

# ═════════════════════════ 模板原页：封面 / 目录 / 结尾 ═════════════════════════
slides0 = list(prs.slides)
cover, toc, end = slides0[0], slides0[1], slides0[9]

tbs = [sh for sh in cover.shapes if sh.has_text_frame]
title_box = min(tbs, key=lambda s: s.top)
who_box = max(tbs, key=lambda s: s.top)
replace_text(title_box, ["干部试用期转正述职"])
text(cover, 1.94, 3.28, 9.45, 0.6, "国家级战略业务突破与国际咨询能力建设", size=22, color=NAVY, align="c", valign="m", sa=0)
who_box.left, who_box.width = Inches(2.67), Inches(8.0)
who_box.top, who_box.height = Inches(4.75), Inches(1.15)
replace_text(who_box, ["汇报人：陈昊　｜　星火国际系统部　｜　市场咨询总监", "任职起点：2026年2月4日　　述职日期：　　年　　月　　日"], size=16)
text(cover, 1.94, 6.0, 9.45, 0.42, "重大业务持续牵引　｜　关键咨询亲自攻坚　｜　国际咨询队伍建设", size=13, color=BLUE, bold=True, align="c", valign="m", sa=0)
notes(cover, script_notes(1))

# ═════════════════════════ P2 双线成果总览 ═════════════════════════
s = content_slide("1-管理角色澄清", "双线成果总览：岗位本职与战略专项同步推进", role="看结果")
card(s, L, CT, CW, 0.5, fill=TINT2, radius=0.3)
chip(s, L + 0.12, CT + 0.09, 1.0, 0.32, "核心结论", size=9.5)
text(s, L + 1.22, CT, CW - 1.35, 0.5, "在国际业务与本地组织仍处建设阶段的条件下，同步推进巴西重大项目突破和多区域咨询能力建设。", size=11, color=NAVY, bold=True, valign="m", sa=0)


def line_cell(group, item, color):
    return [(group, {"size": 9, "color": color, "bold": True}), (item, {"size": 10.5, "color": NAVY, "bold": True})]


table(s, L, CT + 0.62, [1.75, 4.75, 2.95, 3.08], ["成果主线", "已形成的结果与产出", "公司获得的价值", "本人承担的关键责任"], [
    [line_cell("战略专项 · 巴西", "国家AI项目", BLUE),
     "总统公开宣布相关合作、三方合作文件签署；多轮交流及4批30多个问题应答后，主要技术路线获客户认可，方案内容体现到RNP技术需求书；采购合同、本地实体及交付前提持续推进",
     "在巴西国家级AI合作中形成明确合作位置，为后续签约与长期承接建立项目基础",
     "持续承担讯飞侧总体推进，补位客户与销售工作，亲自完成关键英语咨询，组织政府合作、华为及内部专业团队协同"],
    [line_cell("战略专项 · 巴西", "SERPRO应用合作", BLUE),
     "从多场景规划进入MGI-SGP招聘应用PoC；按8月更新已完成开发，继续推进部署、最终客户确认、验收及采购路径",
     "在国家级项目之外形成具体政务应用验证基础，为商业转化提供现实切入点",
     "牵引需求与方案、产品资源、伙伴协调及客户沟通；开发由研发团队完成"],
    [line_cell("岗位本职 · 国际咨询", "咨询组织建设", GREEN),
     "Q1形成共6人的团队配置，5个重点区域有主责咨询；持续培养骨干并跨国部署，徐唯逐步承担巴西主要方案及现场咨询",
     "将部分原来依赖个人的工作转为区域主责与骨干承接",
     "招募与识别、实战带教、任务分配、主备与现场衔接、关键质量把关"],
    [line_cell("岗位本职 · 国际咨询", "多区域业务支撑", GREEN),
     "持续形成国别方案、HLD、SoC、RFI、可研、演示、报价及合同范围材料；在埃及、SDEC、KNG、IGNITE等项目处理范围、竞争与交付风险",
     "提高重点机会的专业支撑与交易准备程度，积累可复用经验",
     "组织资源、参与策略判断、评审关键材料；主责咨询与专业团队落实具体产出"],
], row_h=0.98, size=9.5, first_col_bold=False, valign="m")
card(s, L, CT + 0.62 + 0.42 + 4 * 0.98 + 0.14, CW, 0.72, fill=NAVY, radius=0.12)
text(s, L + 0.2, CT + 0.62 + 0.42 + 4 * 0.98 + 0.14, CW - 0.4, 0.72, [
    ("**共同特征**　既亲自承担最关键的项目与专业问题，又持续组织团队形成跨区域产出。", {}),
    ("**当前经营状态**　重大阶段成果已形成；RNP正式商业合同、SERPRO最终验收及相关回款尚需继续完成。", {}),
], size=10.5, color=WHITE, valign="m", sa=2)
notes(s, script_notes(2))

# ═════════════════════════ P3 目录（模板页） ═════════════════════════
PAGE[0] += 1
label_map = {"管理角色澄清": "管理角色澄清与成果回顾", "Part5": "Part 5", "Part6": "Part 6"}
for sh in walk(toc.shapes):
    if sh.has_text_frame and sh.text_frame.text.strip() in label_map:
        replace_text(sh, [label_map[sh.text_frame.text.strip()]])
text(toc, 6.2, 1.84, 6.5, 0.42, "管理角色澄清表 · 同期投入与管理方式 · 战略专项：巴西 · 岗位本职：国际咨询 · 各项展开述职（角色自评）", size=10, color=MUTE, valign="t", sa=0)
for yy, pg in ((1.36, "P04—26"), (2.27, "P27"), (3.16, "P28—29"), (4.04, "P30"), (4.95, "P31"), (5.77, "P32")):
    text(toc, 10.9, yy, 1.4, 0.44, pg, size=10.5, color=MUTE, valign="m", align="r", sa=0)
notes(toc, script_notes(3))

# ═════════════════════════ P4 管理角色澄清表 ═════════════════════════
s = content_slide("1-管理角色澄清", "管理角色澄清表：四类角色的目标、投入与自评", role="看定位")
card(s, L, CT, CW, 0.62, fill=TINT2, radius=0.25)
chip(s, L + 0.12, CT + 0.15, 1.05, 0.32, "可衡量结果", size=9.5)
text(s, L + 1.25, CT, 7.0, 0.62, [("本职：重点区域主责配置、多类咨询产出、骨干承接", {}), ("专项：巴西合作位置、技术路线认可、采购与本地承接推进、SERPRO PoC", {})], size=10, color=NAVY, valign="m", sa=1)
chip(s, L + 8.4, CT + 0.15, 0.9, 0.32, "策略", size=9.5)
text(s, L + 9.4, CT, 3.1, 0.62, "重大项目亲自攻坚，咨询组织同步建设，两线并行推进", size=10, color=NAVY, valign="m", sa=0)


def self_eval(label, sentence, refs):
    fg, _ = STATUS.get(label, (AMBER, AMBER_T))
    return [(label, {"bold": True, "color": fg}), (sentence, {}), (refs, {"size": 8.5, "color": MUTE})]


table(s, L, CT + 0.74, [1.5, 2.3, 0.55, 2.7, 2.05, 0.85, 2.58],
      ["管理角色", "目标（实际承担的业绩责任）", "权重", "高回报活动（3项）", "管理机制", "时间投入", "完成情况自评"], [
          [["**岗位特性类**", ("关键客户咨询与复杂方案牵引", {"size": 8.5, "color": MUTE})],
           ["① 巴西主要技术路线获客户认可并进入采购需求", "② 重点区域形成国别方案与技术建议"], "",
           ["① 亲自承担RNP关键英语咨询与技术澄清〔专项〕", "② 国别方案与技术路线判断〔本职〕", "③ 重大效果、范围、权利边界把关〔双线〕"],
           ["联合准备、关键问题复核", "客户问答与材料迭代", "面向团队的标准口径材料"], "",
           self_eval("已形成", "RNP主要路线获认可，客户TR体现方案输入；合同边界仍需收口", "见 P8—P10、P19")],
          [["**目标管理类**", ("巴西重点项目总体推进与全球咨询目标统筹", {"size": 8.5, "color": MUTE})],
           ["① RNP进入采购流程并具备签约条件", "② SERPRO进入PoC验证", "③ 各区域商机按成熟度分级推进"], "",
           ["① 关键卡点识别与里程碑推进〔专项〕", "② 高层、伙伴与内部决策输入〔专项〕", "③ 人员、专业资源安排与商机取舍〔本职〕"],
           ["项目周报、阶段任务分解", "专题讨论与问题跟踪", "现场LTC"], "",
           self_eval("推进中", "重大合作、技术与交易准备取得阶段进展；最终经营闭环尚未完成", "见 P7、P11—P15、P22")],
          [["**人与团队类**", ("建队伍、育骨干、保持现场与区域支撑连续性", {"size": 8.5, "color": MUTE})],
           ["① 重点区域形成主责咨询配置", "② 骨干能独立承接项目主责"], "",
           ["① 招募与识别〔本职〕", "② 实战带教〔本职〕", "③ 跨区部署、交接与销售导入〔双线〕"],
           ["集中办公、常态辅导", "主责分工与关键事项升级"], "",
           self_eval("推进中", "区域配置形成，骨干逐步承担主责；少数关键任务仍依赖本人把关", "见 P16、P17")],
          [["**运营类**", ("连接客户、伙伴、方案、服务与法财工作", {"size": 8.5, "color": MUTE})],
           ["① 售前、合同、交付之间不出现断层", "② 本地承接条件明确"], "",
           ["① 关键材料评审〔本职〕", "② 采购事项书面化〔专项〕", "③ 专业接口与交付风险前置〔双线〕"],
           ["评审会、责任接口", "正式邮件与行动事项"], "",
           self_eval("已形成", "采购、合同与本地承接问题更加明确；仍需提高机制稳定性", "见 P12、P13、P18、P25")],
      ], row_h=1.1, size=9.5, header_size=9.5, valign="t", aligns=["l", "l", "c", "l", "l", "c", "l"], first_col_bold=False)
note(s, CT + 0.74 + 0.42 + 4 * 1.1 + 0.06, "权重与时间投入比例按实际任务由本人填写，不作事后推算；〔专项〕〔本职〕标注该活动所属主线。")
notes(s, script_notes(4, "本页与下一页共用脚本第4页讲述稿。"))

# ═════════════════════════ P5 同期投入与管理方式的变化 ═════════════════════════
s = content_slide("1-管理角色澄清", "同期投入与管理方式的变化：从本人补位走向骨干承接", role="看时间精力")
text(s, L, CT, 1.2, 0.5, "起点条件", size=11, color=NAVY, bold=True, valign="m", sa=0)
starts = ["本地无实体、无专业合同支撑", "无可套用的交易模式", "咨询队伍从0起步、人员轮换", "销售衔接不足，需本人补位"]
sw = (CW - 1.25 - 3 * 0.14) / 4
for i, t in enumerate(starts):
    chip(s, L + 1.25 + i * (sw + 0.14), CT + 0.08, sw, 0.36, t, fill=GRAY_T, color=NAVY, size=9.5, bold=True, line=LINE, radius=0.3)
hy = CT + 0.62
colw = [1.25, 3.4, 3.4, 4.1]
xs = [L, L + 1.37, L + 4.89, L + 8.41]
heads = ["阶段", "战略专项 · 巴西", "岗位本职 · 国际咨询", "本人与团队的分工"]
for i in range(4):
    box(s, xs[i], hy, colw[i], 0.38, fill=NAVY if i == 0 else (BLUE if i < 3 else GREEN))
    text(s, xs[i], hy, colw[i], 0.38, heads[i], size=10.5, color=WHITE, bold=True, valign="m", align="c" if i == 0 else "l")
rows = [
    ("补位期", "2—5月", ["RNP方案交流与4批技术澄清", "SERPRO医疗、教育、人事管理场景推进", "现场销售信息与客情导入"],
     ["埃及客户交流与售前向交付交底", "加纳、IGNITE、SDEC等方案支撑", "处理咨询离职、驻场与轮换压力"],
     ["本人直接面客、补位销售、把握方案细节", "徐唯进入巴西与埃及项目，联合准备与面客"]),
    ("分工期", "6—8月", ["服务化、总集与分签模式讨论；报价与付款设计", "SGP需求确认与PoC准备", "7月现场LTC，商法财到场"],
     ["埃及合同评审；非洲RFI、概算与可研", "东马现场部署；新咨询人员带教"],
     ["骨干承担主要方案与现场任务", "本人转向关键判断、复核、总体推进与法财接口"]),
    ("承接期", "9月起", ["政府间文件、采购接口、SOW", "本地实体注册与SGP部署验收条件"],
     ["KNG范围风险评审；SDEC五个子项目SoC", "哈萨克斯坦、泰国等方案与客户交流"],
     ["明确主备与授权边界；重大承诺先内部一致", "本人保留关键把关，减少单点依赖"]),
]
rh = 1.02
for r, (ph, mo, a, b, c) in enumerate(rows):
    y = hy + 0.46 + r * (rh + 0.08)
    card(s, xs[0], y, colw[0], rh, fill=TINT2)
    text(s, xs[0], y + 0.12, colw[0], 0.4, ph, size=12, color=NAVY, bold=True, align="c", valign="m", sa=0)
    text(s, xs[0], y + 0.5, colw[0], 0.35, mo, size=10, color=MUTE, align="c", valign="m", sa=0)
    for i, items in ((1, a), (2, b)):
        card(s, xs[i], y, colw[i], rh, fill=TINT)
        text(s, xs[i] + 0.08, y + 0.04, colw[i] - 0.16, rh - 0.08, [(t, {"bullet": True}) for t in items], size=9.5, valign="m", sa=2)
    card(s, xs[3], y, colw[3], rh, fill=GREEN_T)
    text(s, xs[3] + 0.08, y + 0.04, colw[3] - 0.16, rh - 0.08, [(t, {"bullet": True, "bcolor": GREEN}) for t in c], size=9.5, valign="m", sa=2)
sy = hy + 0.46 + 3 * (rh + 0.08) + 0.04
for i, (big, lab) in enumerate([("5次", "七个月内中巴往返"), ("11小时", "时差下的两地协同"), ("2地", "巴西、埃及同时需常驻与轮换"), ("本人接续", "现场人员变化时跨区支援")]):
    stat(s, L + i * (w4 + 0.2), sy, w4, 0.7, big, lab, big_size=16)
text(s, L, sy + 0.76, CW, 0.36, "专项与本职同期推进；管理方式从本人补位转向骨干承接，尚未消除单点依赖。三期划分为回顾归纳。", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
notes(s, script_notes(4, "本页承接脚本第4页的同期工作负荷与资源背景；三期划分为对周报事实的回顾归纳，讲述时说明。"))

# ═════════════════════════ P6 巴西成果全景 ═════════════════════════
s = content_slide("1-管理角色澄清", "巴西成果全景：合作地位、技术需求、交易准备与应用验证同步推进", role=f"{SP} ｜ 看结果")
c0, c1, c2, c3 = 1.45, 3.35, 4.45, 2.8
x0, x1, x2, x3 = L, L + c0 + 0.1, L + c0 + c1 + 0.2 + 0.28, L + c0 + c1 + c2 + 0.48 + 0.1
for x, w, h_, f in ((x0, c0, "维度", NAVY), (x1, c1, "早期状态 / 主要障碍", GRAYC), (x2, c2, "已取得的阶段变化", BLUE), (x3, c3, "对公司价值", GREEN)):
    box(s, x, CT, w, 0.38, fill=f)
    text(s, x, CT, w, 0.38, h_, size=10.5, color=WHITE, bold=True, valign="m")
rows = [
    ("合作位置", "需与华为共同争取客户对联合方案的认可；实施主体及合作方式逐步明确", "4月形成战略合作文件（意向性质）；RNP承担国家AI项目推进；8月20日总统公开宣布相关合作，三方MoU 8月完成签署", "讯飞在国家级AI合作中的角色由方案交流走向公开合作位置"),
    ("技术方案", "自主训练、数据、模型效果、技术依赖与知识转移存在大量待澄清问题", "联合完成多批问题应答，RNP总体接受主要方案与技术路线；相关内容体现在5月TR草稿", "咨询输出转为客户技术需求与后续采购讨论的输入"),
    ("交易准备", "联合体、总集、服务化、分签等路径反复讨论，范围与付款依赖交易结构", "形成阶段报价与付款里程碑；进入采购部门直接对接、TR／SoC／SOW及合同准备", "为合作转化为正式交易建立具体材料与工作路径"),
    ("应用合作", "SERPRO对医疗、教育和政务应用有兴趣，尚需明确客户与场景", "MGI-SGP招聘场景确认由讯飞开展PoC；8月完成开发并进入部署与验收沟通", "建立面向政府实际业务的应用验证基础"),
    ("本地承接", "本地实体、经营范围、专业合同支撑与履约接口尚不完备", "推动注册材料与经营范围确认、律师资源配置、现场LTC和采购接口建立", "为本地签约和履约准备提供基础"),
]
rh = 0.78
for r, (a, b, c, d) in enumerate(rows):
    y = CT + 0.46 + r * (rh + 0.07)
    card(s, x0, y, c0, rh, fill=TINT2)
    text(s, x0 + 0.06, y, c0 - 0.12, rh, a, size=10.5, color=NAVY, bold=True, valign="m", sa=0)
    card(s, x1, y, c1, rh, fill=GRAY_T)
    text(s, x1 + 0.08, y, c1 - 0.16, rh, b, size=9.5, color=MUTE, valign="m", sa=0)
    chevron(s, x1 + c1 + 0.06, y + (rh - 0.3) / 2, w=0.15, h=0.3)
    card(s, x2, y, c2, rh, fill=TINT)
    text(s, x2 + 0.08, y, c2 - 0.16, rh, c, size=9.5, color=TXT, valign="m", sa=0)
    card(s, x3, y, c3, rh, fill=GREEN_T)
    text(s, x3 + 0.08, y, c3 - 0.16, rh, d, size=9.5, color=TXT, valign="m", sa=0)
fy = CT + 0.46 + 5 * (rh + 0.07) + 0.02
text(s, L, fy, 2.2, 0.36, "本人主导的工作连接", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
flow = ["客户与高层推进", "关键咨询", "商业路径与材料", "本地与交付条件"]
fw = (CW - 2.3 - 3 * 0.3) / 4
for i, t in enumerate(flow):
    x = L + 2.3 + i * (fw + 0.3)
    chip(s, x, fy + 0.02, fw, 0.32, t, fill=TINT2, color=NAVY, size=10, radius=0.3)
    if i < 3:
        chevron(s, x + fw + 0.08, fy + 0.03, w=0.14, h=0.3)
note(s, fy + 0.42, "状态说明：上述为合作、技术、采购准备及开发成果；正式商业合同、回款和最终验收仍需推进。")
notes(s, script_notes(5))

# ═════════════════════════ P7 案例① 政府合作攻坚 ═════════════════════════
s = content_slide("1-管理角色澄清", "关键案例①　政府合作攻坚：将签约障碍转化为正式协调事项", role=f"{SP} ｜ 目标管理类 ｜ 跨层级推动")
case_layers(s, CT, 2.2, [
    ("原来的障碍", GRAYC, ["巴方提出直接采购需要政府间文件支持", "使馆签署涉及国内部委会商与授权，并要求书面说明项目情况与需求", "单靠供应商合同谈判无法解决"]),
    ("关键判断", BLUE2, ["把前线障碍整理成高层可处理的决策事项，而非一般性介绍项目", "伙伴内部协调受限时，讯飞侧单独提交、继续推进"]),
    ("本人行动", BLUE, ["推动8月19日再次向大使专题汇报，聚焦政府间协议、合规程序与签约时间", "当晚组织现场团队完成《当前进展和紧急需求的汇报》提交发改委", "后续持续提供项目具体信息"]),
    ("可见变化", GREEN, ["8月20日反馈原则同意签署", "8月21日文本发送巴西科创部", "8月25日巴方反馈意见，进入文本确认"]),
], size=10.5)
ty = CT + 2.4
nodes = [("8.19", "大使专题汇报"), ("当晚", "紧急需求报告提交发改委"), ("8.20", "原则同意签署"), ("8.21", "文本发送科创部"), ("8.25", "巴方反馈意见")]
box(s, L + 0.6, ty + 0.46, CW - 1.2, 0.03, fill=TINT2)
step = (CW - 1.2) / (len(nodes) - 1)
for i, (d, lab) in enumerate(nodes):
    cx = L + 0.6 + i * step
    dot(s, cx - 0.13, ty + 0.345, 0.26, "", fill=BLUE if i < 2 else GREEN)
    text(s, cx - 1.1, ty, 2.2, 0.32, d, size=11, color=NAVY, bold=True, align="c", valign="m", sa=0)
    text(s, cx - 1.1, ty + 0.66, 2.2, 0.32, lab, size=9.5, color=TXT, align="c", valign="t", sa=0)
vy = ty + 1.05
vals = [("推进方式改变", "从“有问题需协调”转为“有明确问题、书面材料、相关部门与后续动作”的持续跟踪"),
        ("总体责任体现", "伙伴内部流程受约束时保持讯飞侧行动连续，并与采购、合同及注册准备相衔接"),
        ("阶段边界", "原则同意与文本往返不等于最终签署，也不替代RNP自身的采购论证与内部审批")]
for i, (k, v) in enumerate(vals):
    x = L + i * (w3 + 0.26)
    card(s, x, vy, w3, 0.95, fill=TINT)
    chip(s, x + 0.12, vy + 0.1, 1.15, 0.28, k, fill=NAVY if i < 2 else GRAYC, size=9)
    text(s, x + 0.1, vy + 0.4, w3 - 0.2, 0.55, v, size=9.5, color=TXT, valign="t", sa=0)
card(s, L, vy + 1.05, CW, 0.5, fill=GRAY_T)
chip(s, L + 0.14, vy + 1.14, 1.0, 0.32, "同期对接", fill=NAVY, size=9.5)
text(s, L + 1.25, vy + 1.05, CW - 1.4, 0.5, "7月大使沟通 · MGI高层交流 · 8月总统府、外交部、科创部对接，分别对应文件、节奏、注册与本地经营承诺等实际议题", size=10, valign="m", sa=0)
notes(s, script_notes(6))

# ═════════════════════════ P8 技术方案成果 ═════════════════════════
s = content_slide("1-管理角色澄清", "技术方案成果：主要路线获认可，方案内容进入客户TR", role=f"{SP} ｜ 岗位特性类 ｜ 看结果")
stages = [
    ("面向新实施主体调整方案", "从LNCC相关方案转向RNP，需适应新主体的建设目标与采购关切", "按RNP需求调整主权AI方案，联合华为完成融合与汇报准备", "4月形成面向RNP的联合汇报材料并开展正式交流"),
    ("持续回应核心质疑", "客户围绕训练、语料、主权能力、评测与长期使用提出多批问题", "联合准备4批30多个问题应答；本人亲自承担关键英语技术与业务解释", "多轮现场与线上澄清，问题逐步转为明确方案与待讨论事项"),
    ("取得主要路线认可", "客户需判断当前方案能否支撑巴西模型与自主能力建设", "区分科研探索与项目交付，解释基于成熟底座持续训练与知识转移路径", "5月周报：RNP总体接受方案与技术路线，本阶段技术澄清关闭"),
    ("内容进入采购需求", "客户需将合作方向写成技术需求文件", "持续提交HLD及技术回复，并对客户TR进行分析", "5月15日TR草稿与HLD及澄清回复基本一致，形成采购讨论输入"),
    ("新增范围继续管理", "TR新增DLM、基础权重开放；后续还有内容安全、年费及算力问题", "联合服务、方案与伙伴讨论应答策略，形成SoC与后续澄清", "在路线认可后继续管理范围、费用与交付前提"),
]
for i, (t, *_rest) in enumerate(stages):
    x = L + i * (w5 + 0.16)
    box(s, x, CT, w5, 0.52, fill=BLUE if i < 4 else AMBER, shape=MSO_SHAPE.PENTAGON if i < 4 else MSO_SHAPE.RECTANGLE)
    dot(s, x + 0.1, CT + 0.11, 0.3, str(i + 1), fill=WHITE, color=BLUE if i < 4 else AMBER, size=10)
    text(s, x + 0.45, CT, w5 - 0.6, 0.52, t, size=10, color=WHITE, bold=True, valign="m", sa=0)
labels = [("障碍 / 任务", GRAYC), ("本人与团队工作", BLUE), ("产出或变化", GREEN)]
ry = CT + 0.62
rh = 0.98
for j, (lab, col) in enumerate(labels):
    y = ry + j * (rh + 0.06)
    for i, st in enumerate(stages):
        x = L + i * (w5 + 0.16)
        card(s, x, y, w5, rh, fill=(TINT if i < 4 else AMBER_T) if j != 0 else GRAY_T)
        text(s, x + 0.08, y + 0.02, w5 - 0.16, rh - 0.04, st[j + 1], size=9.5, color=MUTE if j == 0 else TXT, valign="m", sa=0)
    chip(s, L - 0.02, y - 0.13, 1.1, 0.24, lab, fill=col, size=8.5)
hy2 = ry + 3 * (rh + 0.06) + 0.06
hl = [("4批 · 30+", "客户问题联合应答"), ("5月", "客户总体接受方案与技术路线"), ("5.15", "TR草稿与HLD及澄清回复基本一致")]
hw = (CW - 3.3 - 2 * 0.2) / 3
for i, (b, lab) in enumerate(hl):
    stat(s, L + i * (hw + 0.2), hy2, hw, 0.7, b, lab, big_size=15)
card(s, L + 3 * (hw + 0.2), hy2, CW - 3 * (hw + 0.2), 0.7, fill=AMBER_T)
text(s, L + 3 * (hw + 0.2) + 0.12, hy2, CW - 3 * (hw + 0.2) - 0.24, 0.7, "**新增范围**　DLM、基础权重开放、内容安全、年费、算力：已认可与新增内容分别标识，继续评审，不视为无边界承诺", size=9.5, color=TXT, valign="m", sa=0)
note(s, hy2 + 0.78, "归因：客户侧文件由客户编制；主要路线认可与TR变化是多轮联合工作的阶段结果，不归因为某一次发言。")
notes(s, script_notes(7))

# ═════════════════════════ P9 案例② 关键咨询攻坚 ═════════════════════════
s = content_slide("1-管理角色澄清", "关键案例②　把自主能力诉求转化为可实施、可验收的路径", role=f"{SP} ｜ 岗位特性类 ｜ 专业判断")
case_layers(s, CT, 2.75, [
    ("客户诉求", GRAYC, ["能否先以小模型从零训练，掌握完整技术，随后再扩大模型与算力规模", "背景：项目转向RNP为实施主体，方案需与华为融合并首次正式面客"]),
    ("本人判断", BLUE2, ["“科研探索”与“当前面向公共服务的项目交付”必须区分", "后者要同时考虑效果、实施投入、可控性与供应商责任"]),
    ("本人行动", BLUE, ["用英语解释：先依托成熟底座完成本地增训、语料建设与知识转移，再基于吸收能力逐步探索从零训练", "不否定长期自主目标，也不把全部目标无条件放入当前项目"]),
    ("现场变化", GREEN, ["客户随后明确复述：当前可掌握持续训练与演进，从零训练是更高难度的后续方向", "讨论从“是否真正自主”转向能力阶段、实施条件与交付责任"]),
], size=10.5)
by = CT + 2.95
card(s, L, by, CW, 0.62, fill=NAVY, radius=0.12)
chip(s, L + 0.15, by + 0.16, 1.0, 0.3, "本人作用", fill=WHITE, color=NAVY, size=9.5)
text(s, L + 1.25, by, CW - 1.4, 0.62, "不仅解释技术“能不能做”，还把客户目标、实施条件和供应商承担结果的范围放到同一讨论框架。", size=11, color=WHITE, bold=True, valign="m", sa=0)
qy = by + 0.78
text(s, L, qy, 2.4, 0.34, "配套五问（见下页）", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
qs = ["本地数据不足如何开始", "语料规模为何不能与万亿级比较", "能否从昇腾迁移到其他生态", "如何验收葡语与本土能力", "如何兼顾自主使用与知识产权"]
qw = (CW - 2.5 - 4 * 0.12) / 5
for i, q in enumerate(qs):
    chip(s, L + 2.5 + i * (qw + 0.12), qy + 0.02, qw, 0.32, q, fill=TINT2, color=NAVY, size=9, radius=0.3)
notes(s, script_notes(8, "本页与下一页共用脚本第8页讲述稿。"))

# ═════════════════════════ P10 关键咨询问答 ═════════════════════════
s = content_slide("1-管理角色澄清", "关键咨询问答：配套五问的澄清及其对方案与交易的作用", role=f"{SP} ｜ 岗位特性类")
table(s, L, CT, [3.0, 5.6, 3.93], ["客户问题", "本人关键解释 / 建议", "对方案与交易的作用"], [
    ["本地数据不足，如何开始？", "先调查本地已有数据与机构资源，评估数量、类型和质量，再决定复用、质量处理或新建；讨论语料团队与训练团队并行", "将数据责任转为可启动的调查与建设安排"],
    ["本地语料规模为何不能直接与其他项目万亿级数据比较？", "区分从零预训练与基于成熟底座持续训练的统计口径，澄清本地语料与通用数据的不同用途", "为客户判断方案充分性提供一致的比较维度，不把不同口径直接对比"],
    ["能否从昇腾迁移到其他硬件生态？", "解释软件栈、算子与工程优化工作，区分基础可部署支持与额外性能优化服务", "回应开放性诉求，并把工程投入与服务边界纳入讨论"],
    ["如何验收葡语与本土能力？", "区分语言能力和巴西本地内容相关能力；提出初版基准、客户讨论及量化验收口径", "将主观效果目标转为可共同确认的评测安排"],
    ["如何兼顾自主使用与核心知识产权？", "区分模型／数据使用、定制代码、基础版本核心代码、持续升级及再分发范围", "为后续权利与服务条款提供讨论边界，而非作无条件转让承诺"],
], row_h=0.86, size=10)
note(s, CT + 0.42 + 5 * 0.86 + 0.15, "5月5日会议中有对应发言；5月8日语料比较内容为本人回忆补录。会议中的周期、性能、所有权说明不直接转成技术保证或合同条款。", h=0.5)
notes(s, script_notes(8, "配套五问页，追问技术细节时使用。"))

# ═════════════════════════ P11 商业模式与价值争取 ═════════════════════════
s = content_slide("1-管理角色澄清", "商业模式与价值争取：将技术范围、交易结构与付款安排衔接", role=f"{SP} ｜ 目标管理类")
table(s, L, CT, [2.1, 6.2, 4.23], ["阶段 / 问题", "本人组织或参与的工作", "形成的阶段结果"], [
    ["联合体或分签", "围绕客户希望整体承接的诉求，组织讨论各方合同界面、交付里程碑及付款方式", "将联合技术方案进一步转为交易结构讨论"],
    ["服务化与总集方案", "面对资产持有和OPEX采购要求，与华为一线研判讯飞担任总集的条件；组织多轮内部汇报，并将集团确定的态度与原则反馈一线", "形成明确的内部决策输入与对外原则，不在未明确条件下直接承诺"],
    ["本地承接能力顾虑", "围绕客户对知名度、本地组织及履约能力的担心，组织案例、资质和子公司相关说明准备", "针对影响交易选择的具体顾虑提供证明材料"],
    ["分签方向与材料调整", "随客户及华为反馈更新工作范围、报价与付款安排", "6月按新范围刷新讯飞侧报价，完成付款里程碑设计并提交客户"],
], row_h=0.78, size=10)
qy = CT + 0.42 + 4 * 0.78 + 0.2
text(s, L, qy, 3.0, 0.34, "本人重点关注的三个经营问题", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
qs = [("讯飞在联合方案中承担什么", "参与训练工具链与平台集成界面讨论，明确阶段分工；推理、容器与服务范围继续向伙伴确认"),
      ("项目分期与回款如何衔接", "组织讨论实施分期、一期份额与回款目标，避免交付计划与付款安排脱节；实际以合同确认"),
      ("责任与收入是否对应", "在总集、分包、资产持有与额外服务讨论中，将新增责任、资源投入与客户预期一起纳入内部判断")]
for i, (k, v) in enumerate(qs):
    x = L + i * (w3 + 0.26)
    card(s, x, qy + 0.4, w3, 1.1, fill=TINT)
    text(s, x + 0.12, qy + 0.46, w3 - 0.24, 0.34, k, size=11, color=BLUE, bold=True, valign="m", sa=0)
    text(s, x + 0.12, qy + 0.8, w3 - 0.24, 0.66, v, size=9.5, color=TXT, valign="t", sa=0)
note(s, qy + 1.58, "具体产出：集成方案讨论、总集原则与内部汇报、工作范围与报价更新、交付及付款里程碑。阶段报价金额与币种口径待核，不上屏。")
notes(s, script_notes(9))

# ═════════════════════════ P12 采购与合同攻坚 ═════════════════════════
s = content_slide("1-管理角色澄清", "采购与合同攻坚：将未知条件逐项明确并形成书面依据", role=f"{SP} ｜ 运营类")
c0, c1, c2, c3 = 2.15, 5.55, 3.55, 1.08
x0, x1, x2, x3 = L, L + c0 + 0.1, L + c0 + c1 + 0.2, L + c0 + c1 + c2 + 0.3
for x, w, h_, f in ((x0, c0, "关键工作", NAVY), (x1, c1, "本人及团队的具体行动", BLUE), (x2, c2, "已形成的产出 / 明确事项", BLUE), (x3, c3, "状态", BLUE)):
    box(s, x, CT, w, 0.38, fill=f)
    text(s, x, CT, w, 0.38, h_, size=10, color=WHITE, bold=True, valign="m")
rows = [
    ("配置本地专业支撑", "现场完成三家律所调研与交流，选定与RNP对接所需的葡语合同顾问律师", "为采购和合同沟通配置具体专业支持；不替代集团法务决策", "已形成"),
    ("组织采购部门首次直接对接", "8月4日与财务、法务、商务及外聘律师共同讨论主体资格、签约前提和节奏", "明确本地子公司承接方向、母公司业绩援用、采购资料与后续核查事项", "已形成"),
    ("推动准备工作并行", "询问公司尚在注册时能否先开展材料、报价和合同讨论", "RNP反馈前期准备可并行，正式签约须有效CNPJ；草稿提案可作需求与费用归类输入", "客户认可"),
    ("将沟通结论书面化", "会后向客户发送正式邮件，逐项求证分包报备、履约担保、用工等风险", "形成正式确认请求；建立项目负责人、采购组织、执行接口三层沟通通道", "已提交"),
    ("处理伙伴方案卡点", "两次与华为项目负责人深入交流，针对正式方案未提交推动联合SOW攻关并分解内部任务", "将伙伴侧依赖转为联合工作任务；提交与评审状态继续跟踪", "推进中"),
    ("持续管理未决问题", "汇总并跟进11项遗留事项的书面答复，关联方案、实体与专业责任人", "形成签约准备问题清单；未取得回复的事项保持未决", "待确认"),
]
rh = 0.68
for r, (a, b, c, st) in enumerate(rows):
    y = CT + 0.44 + r * (rh + 0.06)
    card(s, x0, y, c0, rh, fill=TINT2)
    text(s, x0 + 0.06, y, c0 - 0.12, rh, a, size=10, color=NAVY, bold=True, valign="m", sa=0)
    card(s, x1, y, c1, rh, fill=TINT)
    text(s, x1 + 0.08, y, c1 - 0.16, rh, b, size=9.5, valign="m", sa=0)
    card(s, x2, y, c2, rh, fill=TINT)
    text(s, x2 + 0.08, y, c2 - 0.16, rh, c, size=9.5, valign="m", sa=0)
    fg, bg = STATUS[st]
    card(s, x3, y, c3, rh, fill=bg)
    chip(s, x3 + 0.1, y + (rh - 0.3) / 2, c3 - 0.2, 0.3, st, fill=fg, size=9)
fy = CT + 0.44 + 6 * (rh + 0.06) + 0.02
card(s, L, fy, CW, 0.5, fill=GRAY_T)
text(s, L + 0.15, fy, CW - 0.3, 0.5, "**已明确的流程认识**　合作支持文件、采购需求、正式提案与客户内部审批相互关联但不能彼此替代；我方先完成可控准备，再持续跟进外部依赖。", size=10, color=TXT, valign="m", sa=0)
notes(s, script_notes(10))

# ═════════════════════════ P13 案例③ 本地承接与现场LTC ═════════════════════════
s = content_slide("1-管理角色澄清", "关键案例③　本地承接与现场LTC：使方案具备签约和履约条件", role=f"{SP} ｜ 运营类 ｜ 跨专业落地")
case_layers(s, CT, 3.1, [
    ("原来的障碍", GRAYC, ["RNP采购讨论主动提出经营范围与服务代码要求", "当前注册按服务类办理，暂不含贸易、设备租赁；纪要指出原安全设备OPEX安排不能按现有结构纳入签约范围", "CNAE、NBS分类影响提案与后续付款"]),
    ("关键判断", BLUE2, ["技术方案包含某项交付，不等于新设主体已具备以该方式承接、开票和履约的条件", "经营范围、设备安排与合同范围必须放在一起处理，不能只转给律师"]),
    ("本人行动", BLUE, ["7月起现场牵头LTC，商法财到场，围绕公司设立与合同准备协同", "组织方案侧与财务、律师对齐服务内容与CNAE、NBS（核定由会计师承担）", "按纪要行动项承担与华为沟通解决方案、细化方案、向RNP确认未决事项", "推动TR、SOW、BoQ、验收、责任矩阵与付款材料衔接；重大范围与资源承诺先在内部形成一致判断"]),
    ("可见变化", GREEN, ["问题进入签约前提清单并有明确责任人", "8月下旬完成注册材料与经营范围准备，继续推动注册申请", "总部人员、本地支持层级、数据、算力与培训对象等履约前提逐项明确"]),
], size=10)
sy = CT + 3.28
items = [("本地注册", "推进中"), ("客户资金", "待确认"), ("伙伴材料", "推进中"), ("数据中心建设", "待确认"), ("客户配合条件", "待确认")]
for i, (lab, st) in enumerate(items):
    x = L + i * (w5 + 0.16)
    fg, bg = STATUS[st]
    card(s, x, sy, w5, 0.6, fill=bg)
    text(s, x + 0.1, sy, w5 - 1.0, 0.6, lab, size=10.5, color=NAVY, bold=True, valign="m", sa=0)
    chip(s, x + w5 - 0.9, sy + 0.15, 0.78, 0.3, st, fill=fg, size=9)
card(s, L, sy + 0.72, CW, 0.62, fill=TINT)
chip(s, L + 0.14, sy + 0.87, 0.8, 0.32, "归因", fill=NAVY, size=9.5)
text(s, L + 1.05, sy + 0.72, CW - 1.2, 0.62, "问题由RNP在采购讨论中提出；财务及会计师核定经营范围与分类，法务落实文本；本人组织业务依赖与跨方推进。未关闭事项不写“已化解”。", size=10, valign="m", sa=0)
notes(s, script_notes(11))

# ═════════════════════════ P14 SERPRO 应用成果 ═════════════════════════
s = content_slide("1-管理角色澄清", "SERPRO应用成果：从多场景规划进入MGI-SGP招聘应用验证", role=f"{SP} ｜ 目标管理类")
stages = [
    ("建立合作入口", "主线", ["组织主权AI合作伙伴框架RFI阶段材料", "SERPRO高层来访与方案交流", "联合华为对接本地伙伴"]),
    ("联合驻场与需求规划", "主线", ["与华为组成联合团队驻场", "对接SERPRO业务与技术人员", "梳理可优先验证的应用场景"]),
    ("医疗、教育场景匹配", "支线", ["拉通医疗公司讨论移动医疗车影像辅助诊断", "配合对接教育部门，讨论中文学习等快赢场景", "多轮交流，明确偏好；不等同落地订单"]),
    ("MGI-SGP需求收敛", "主线", ["分析招聘和档案数字化需求", "组织产线准备Demo、调研清单和PoC方案", "6月招聘场景获客户认可并确认由讯飞开展PoC"]),
    ("PoC开发与部署", "成果", ["牵引需求细化、产品资源、客户反馈与伙伴协调", "推动周例会与过程演示", "8月PoC完成开发并持续优化，进入云环境部署和最终验收条件确认"]),
]
for i, (t, kind, b) in enumerate(stages):
    x = L + i * (w5 + 0.16)
    fill = TINT2 if kind == "成果" else (GRAY_T if kind == "支线" else TINT)
    card(s, x, CT, w5, 3.15, fill=fill)
    dot(s, x + 0.12, CT + 0.14, 0.32, str(i + 1), fill=GREEN if kind == "成果" else (GRAYC if kind == "支线" else BLUE), size=10)
    text(s, x + 0.5, CT + 0.1, w5 - 0.6, 0.4, t, size=11, color=NAVY, bold=True, valign="m", sa=0)
    chip(s, x + w5 - 0.62, CT + 0.52, 0.5, 0.22, kind, fill=GREEN if kind == "成果" else (GRAYC if kind == "支线" else BLUE2), size=8)
    text(s, x + 0.1, CT + 0.8, w5 - 0.2, 2.3, [(k, {"bullet": True, "bcolor": GRAYC if kind == "支线" else BLUE}) for k in b], size=9.5, sa=3)
    if i < 4:
        chevron(s, x + w5 + 0.01, CT + 0.2, w=0.14, h=0.3, fill=BLUE2)
cy = CT + 3.3
titled_card(s, L, cy, cw2, 1.15, "本人承担", ["需求与价值判断、跨产品线组织、客户和伙伴沟通、关键节奏与采购路径推进"], title_size=11.5, size=10)
titled_card(s, L + cw2 + 0.24, cy, cw2, 1.15, "团队承担", ["具体方案、Demo与系统开发、部署和技术测试；阶段效果由客户反馈和后续验收分别确认"], title_size=11.5, size=10)
note(s, cy + 1.22, "公司价值：在国家级基础能力项目之外，形成面向政府真实业务的应用验证基础，为商业转化积累场景、客户与实施经验。")
notes(s, script_notes(12))

# ═════════════════════════ P15 SERPRO 商业推进 ═════════════════════════
s = content_slide("1-管理角色澄清", "SERPRO商业推进：同步组织客户确认、研发验证、云资源与采购路径", role=f"{SP} ｜ 目标管理类")
table(s, L, CT, [1.7, 3.6, 4.0, 3.23], ["推进链条", "当时存在的关键问题", "本人与团队处理方式", "当前成果与边界"], [
    ["最终业务客户确认", "SERPRO认可方案，不等于MGI-SGP已确认全部PoC条件；后续存在局部调整", "通过周例会、确认函和持续沟通推进业务条件明确，收到变化后组织系统对齐", "过程反馈积极；正式确认及调整仍需逐项跟进"],
    ["开发与验证", "Demo认可需进一步转成明确功能和测试安排", "组织需求细化、产品经理及研发支持，持续开展过程演示和优化", "开发已形成阶段成果；不以开发完成替代最终验收"],
    ["云资源与部署", "华为云资源供给存在延迟，影响部署与交付节奏", "与华为、SERPRO和现场团队梳理资源卡点，推进申请、部署和后续安排", "进入部署阶段；资源和交付节奏仍有外部依赖"],
    ["采购路径", "PoC成功后如何通过现有合作关系采购，涉及上架主体和合同路径", "组织讨论通过华为云MarketPlace与既有MGI—SERPRO—华为框架衔接；保留香港主体、巴西代理等选项", "形成可继续论证的路径；尚不等于上架完成或订单获得"],
    ["销售承接", "前期项目背景与伙伴资源需要进入销售工作", "导入前期进展、华为侧接口和客户诉求，梳理后续推进策略", "推动商业信息衔接，避免应用团队与销售各自掌握不完整情况"],
], row_h=0.82, size=9.5)
bar(s, CT + 0.42 + 5 * 0.82 + 0.22, "不仅推进系统“做出来”，也推进客户“确认”、伙伴“能够部署”和后续“具备采购路径”；所有条件以实际确认和合同为准。", h=0.55, size=11)
notes(s, script_notes(13))

# ═════════════════════════ P16 国际咨询组织建设 ═════════════════════════
s = content_slide("1-管理角色澄清", "国际咨询组织建设：从0到1，从人员组建到重点区域主责配置", role=f"{BZ} ｜ 人与团队类")
cols3 = [
    ("人员配置", "2025年底启动招募与首轮带教；人员分散，临时跨区支撑", "Q1记录共6人，5个重点区域均配置主责咨询；随后经历离职、补充与跨区调度"),
    ("责任分工", "本人直接处理更多方案细节与客户问题", "骨干承担主责方案与现场工作；本人保留关键判断、复核与总体推进"),
    ("工作方法", "按单次任务交付，经验留在个人", "集中办公、实战复盘、关键材料评审、跨项目经验复用"),
]
for i, (t, before, after) in enumerate(cols3):
    x = L + i * (w3 + 0.26)
    text(s, x, CT, w3, 0.34, t, size=12, color=NAVY, bold=True, valign="m", sa=0)
    card(s, x, CT + 0.38, w3, 0.7, fill=GRAY_T)
    chip(s, x + 0.1, CT + 0.47, 0.55, 0.26, "起点", fill=GRAYC, size=8.5)
    text(s, x + 0.72, CT + 0.38, w3 - 0.8, 0.7, before, size=9.5, color=MUTE, valign="m", sa=0)
    box(s, x + w3 / 2 - 0.14, CT + 1.12, 0.28, 0.18, fill=BLUE2, shape=MSO_SHAPE.DOWN_ARROW)
    card(s, x, CT + 1.34, w3, 0.85, fill=TINT)
    chip(s, x + 0.1, CT + 1.47, 0.55, 0.26, "现在", fill=BLUE, size=8.5)
    text(s, x + 0.72, CT + 1.34, w3 - 0.8, 0.85, after, size=9.5, color=TXT, valign="m", sa=0)
ry = CT + 2.32
text(s, L, ry, 3.6, 0.3, "区域主责配置与代表项目（Q1阶段记录）", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
regions = [("拉丁美洲", "RNP、SERPRO，以及墨西哥等方案或Demo准备"), ("中东中亚", "IGNITE、KNG、沙特相关项目、哈萨克斯坦等"), ("东南亚", "SDEC、泰国、斯里兰卡等"), ("北部非洲", "埃及ISD、ACUD、SIDPEC等"), ("南部非洲等非洲机会", "加纳、埃塞俄比亚、坦桑尼亚、肯尼亚等")]
for i, (rg, pj) in enumerate(regions):
    x = L + i * (w5 + 0.16)
    card(s, x, ry + 0.34, w5, 0.9, fill=TINT)
    text(s, x + 0.1, ry + 0.37, w5 - 0.2, 0.3, rg, size=10.5, color=BLUE, bold=True, valign="m", sa=0)
    text(s, x + 0.1, ry + 0.66, w5 - 0.2, 0.55, pj, size=9, color=TXT, valign="t", sa=0)
ay = ry + 1.36
text(s, L, ay, 3.0, 0.3, "本人完成的组织工作", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
acts = [("选人与识别", "围绕海外面客、技术理解与跨文化沟通要求招募与判断"), ("任务与区域部署", "把人员放到具体项目与客户任务，按巴西、埃及、东马现场需要调整"), ("质量与资源统筹", "把关关键路线、投标边界、合同范围；人员紧张时决定本人介入或专业增援"), ("连续性保障", "处理离职、轮换与新人到岗衔接，持续推动招聘与带教")]
for i, (k, v) in enumerate(acts):
    x = L + i * (w4 + 0.2)
    card(s, x, ay + 0.34, w4, 0.8, fill=GREEN_T)
    text(s, x + 0.1, ay + 0.36, w4 - 0.2, 0.28, k, size=10, color=GREEN, bold=True, valign="m", sa=0)
    text(s, x + 0.1, ay + 0.62, w4 - 0.2, 0.5, v, size=9, color=TXT, valign="t", sa=0)
cy = ay + 1.26
card(s, L, cy, CW, 0.5, fill=AMBER_T)
chip(s, L + 0.14, cy + 0.1, 1.0, 0.3, "实际约束", fill=AMBER, size=9.5)
text(s, L + 1.25, cy, CW - 1.4, 0.5, "在项目不断进入、重点国家需要现场支持的同时建队；离职、双现场常驻与轮换、小颗粒商机并发，调度压力持续。", size=10, valign="m", sa=0)
notes(s, script_notes(14))

# ═════════════════════════ P17 骨干培养与业务连续性 ═════════════════════════
s = content_slide("1-管理角色澄清", "骨干培养与业务连续性：让关键能力逐步由团队承担", role=f"{BZ} ｜ 人与团队类")
lw = 5.7
text(s, L, CT, lw, 0.34, "徐唯：培养动作与责任递进", size=12, color=NAVY, bold=True, valign="m", sa=0)
steps = [("选人与进入项目", "招募识别；安排进入埃及、巴西等实际项目，导入业务背景、客户诉求与方案逻辑"),
         ("联合准备与面客", "共同准备关键客户交流；能区分决策层、业务层、技术层的关注点，组织分层材料"),
         ("承担主要方案与现场任务", "逐步承担材料编制、客户问题、伙伴沟通及现场主责；关键节点由本人复核"),
         ("本人转向关键把关", "保留重大技术、效果、范围及对外承诺判断；更多处理总体推进、本地实体与法财接口")]
for i, (t, d) in enumerate(steps):
    y = CT + 0.42 + i * 0.86
    card(s, L, y, lw, 0.76, fill=TINT if i < 3 else TINT2)
    dot(s, L + 0.14, y + 0.21, 0.34, str(i + 1), fill=BLUE if i < 3 else NAVY, size=10)
    text(s, L + 0.6, y + 0.04, lw - 0.7, 0.3, t, size=10.5, color=NAVY, bold=True, valign="m", sa=0)
    text(s, L + 0.6, y + 0.32, lw - 0.7, 0.42, d, size=9.5, color=TXT, valign="t", sa=0)
rx = L + lw + 0.3
rw = R - rx
table(s, rx, CT, [rw / 2, rw / 2], ["骨干主责（现在能独立承担）", "本人把关（仍由本人承担）"], [
    ["具体方案材料与常规客户问题应答", "重大技术路线与效果承诺"],
    ["现场主责咨询与日常客户沟通", "工作范围、权利边界与对外正式承诺"],
    ["伙伴跟进、联合准备与会后复盘", "交易模式、报价与付款设计"],
    ["主责交接与新人带教", "跨部门争议与升级事项"],
], row_h=0.6, size=9.5, first_col_bold=False)
text(s, rx, CT + 0.42 + 4 * 0.6 + 0.04, rw, 0.3, "此表为现有分工的概括，不代表已形成完整授权制度；以本人确认为准", size=8.5, color=MUTE, valign="m", sa=0)
my = CT + 3.95
mech = [("授权边界", "一般方案交给骨干；重大效果、范围、权利、对外承诺继续把关，避免“完全放手”与“全部自己做”两个极端"),
        ("现场连续性与主备", "人员阶段性离场或项目并发时，本人接续关键客户与方案工作，协调其他咨询支援并主持交接"),
        ("带教延伸到销售", "对新到现场销售导入项目背景、政府与客户接口、华为关系与现场目标；后续持续同步卡点与伙伴资源")]
for i, (t, b) in enumerate(mech):
    x = L + i * (w3 + 0.26)
    card(s, x, my, w3, 1.1, fill=TINT)
    text(s, x + 0.12, my + 0.08, w3 - 0.24, 0.3, t, size=11, color=NAVY, bold=True, valign="m", sa=0)
    text(s, x + 0.12, my + 0.38, w3 - 0.24, 0.7, b, size=9.5, color=TXT, valign="t", sa=0)
text(s, L, my + 1.18, CW, 0.36, "组织价值：关键任务开始由骨干承接，本人从全部亲自执行转向方向判断、质量复核和资源组织；尚未消除单点依赖。", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
notes(s, script_notes(15))

# ═════════════════════════ P18 埃及 ISD ═════════════════════════
s = content_slide("1-管理角色澄清", "埃及ISD：持续连接客户咨询、PoC、合同与交付准备", role=f"{BZ} ｜ 运营类")
bh = 1.45
band(s, L, CT, CW, bh, "本人判断与组织", [
    "参与AI Lab负责人与技术团队交流，梳理客户RAG现状与下一步策略", "向铁三角与交付团队交底前期HLD、背景与客户要求",
    "识别原HLD部分内容未完全响应RFP，提示后续可能被客户挑战", "参与法务合规讨论，审阅合同初稿并提出修改；提示主合同信息不足影响背靠背衔接"], color=BLUE, size=10)
band(s, L, CT + bh + 0.12, CW, bh, "团队与专业部门产出", [
    "客户调研、技术方案与高层演示材料", "场景演示准备、技术沟通与PoC问题跟踪",
    "基于方案的交付计划与后续报价输入", "协调咨询赴埃及，结合ISD、ACUD、SIDPEC及伙伴活动安排任务"], color=BLUE2, size=10)
band(s, L, CT + 2 * (bh + 0.12), CW, 1.2, "对项目的作用 / 状态", [
    "客户实际系统与我方方案衔接，不只做通用产品介绍", "潜在范围问题在交付前提出，不等到验收时才发现",
    "合同与交付范围更一致；法务合规结论由专业团队确认", "受限事项保留风险，不写成全部验证通过"], color=GREEN, size=10)
bar(s, CT + 2 * (bh + 0.12) + 1.2 + 0.2, "本人贡献重点：持续把握关键风险、组织前后场交接和合同评审；具体测试、实施与专业判断由对应团队承担。", h=0.55, size=11)
notes(s, script_notes(16))

# ═════════════════════════ P19 非洲多国 ═════════════════════════
s = content_slide("1-管理角色澄清", "非洲多国：形成语料、国别方案、RFI与可研等实际产出", role=f"{BZ} ｜ 岗位特性类")
table(s, L, CT, [1.5, 3.6, 4.3, 3.13], ["国家 / 机会", "本人的关键组织与判断", "已记录的团队产出", "阶段价值与未完成事项"], [
    ["加纳", "讨论本地大学语料摸排策略，导入其他国家经验，与巴西方案基线拉通；组织运营、配置及报价讨论", "现场语料摸排、运营方案与HLD、配置和报价更新；8月细化技术建议书，输出语料本土化方案与应用差异化说明", "从通用介绍进入本地资源与实施设计；预算与客户流程仍需跟踪"],
    ["埃塞俄比亚", "协调主责咨询及华为对齐预算、范围和方案口径，参与RFI及方案评审", "伙伴选择RFI初版、方案／概算、应用材料；8月与华为就基础方案达成一致", "形成具体合作与方案讨论输入，不等同客户最终采购接受"],
    ["坦桑尼亚", "跟进前期方案、投入测算与材料质量，组织评审", "一版概算，HLD持续更新", "提供项目规模与可行性讨论基础"],
    ["肯尼亚", "持续跟进可研修改及内部评审，协调资源回应材料需求", "可研材料迭代", "支撑项目可行性讨论，后续决策仍在推进"],
    ["南部非洲相关机会", "组织高阶训战、场景沟通和新机会梳理", "训战材料与现场交流、招采智能体PoC初步沟通等", "形成伙伴与场景输入，前期机会不计为已完成PoC"],
    ["加蓬等新增机会", "与华为及相关团队讨论智算项目和轻量方案方向", "项目沟通及后续材料准备安排", "作为新机会储备，不与成熟项目同口径呈现"],
], row_h=0.76, size=9.5)
note(s, CT + 0.42 + 6 * 0.76 + 0.12, "成果类型：语料调查、国别设计、运营说明、RFI、概算、可研、技术建议书与伙伴赋能。本人价值：把不同国家的需求、预算、数据与交付条件纳入方案判断，推动一线经验复用。", h=0.5)
notes(s, script_notes(17))

# ═════════════════════════ P20 中东中亚 ═════════════════════════
s = content_slide("1-管理角色澄清", "中东中亚：围绕竞争、范围和交付条件组织重点商机支撑", role=f"{BZ} ｜ 岗位特性类")
cases = [
    ("IGNITE（巴基斯坦）", ["新版标书存在无法满足项；伙伴方案与外部竞争均可能影响差异化", "与主责咨询分析标书和竞品，组织内部研判；讨论平台、语料与训练指导等最小工作范围"], ["参数修改建议、方案边界讨论与投标支持口径；内部建议不等同客户采纳"]),
    ("KNG（科威特）", ["客户仍依据初版SoC理解范围，尚未最终认可更新的HLD，存在合同与交付认知差异", "完成SoC评审并联合产线评估风险，推动与华为、总集和客户继续确认"], ["范围风险已识别并进入处理；SoC作为合同范围谈判的重要输入"]),
    ("沙特：MOI 与 SDAIA", ["MOI私有化PoC受硬件、接口适配条件影响，8月记录为中止；组织相关团队协同跟进遗留问题", "SDAIA新机会需明确平台、智能体、模型适配及服务边界"], ["已引入协作团队继续推动，不表述为PoC通过；SDAIA形成方案方向与后续任务"]),
]
for i, (t, j, o) in enumerate(cases):
    x = L + i * (w3 + 0.26)
    card(s, x, CT, w3, 2.95, fill=TINT)
    text(s, x + 0.14, CT + 0.08, w3 - 0.28, 0.36, t, size=12, color=NAVY, bold=True, valign="m", sa=0)
    chip(s, x + 0.14, CT + 0.5, 0.95, 0.26, "障碍与判断", fill=BLUE, size=8.5)
    text(s, x + 0.1, CT + 0.8, w3 - 0.2, 1.3, [(k, {"bullet": True}) for k in j], size=9.5, sa=3)
    chip(s, x + 0.14, CT + 2.1, 0.95, 0.26, "产出与状态", fill=GREEN, size=8.5)
    text(s, x + 0.1, CT + 2.4, w3 - 0.2, 0.55, [(k, {"bullet": True, "bcolor": GREEN}) for k in o], size=9.5, sa=3)
oy = CT + 3.12
card(s, L, oy, CW, 1.05, fill=GRAY_T)
text(s, L + 0.14, oy + 0.06, 2.0, 0.3, "其他产出", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
others = [("哈萨克斯坦", "MaaS与Agent平台客户汇报、对华为一线宣讲"), ("科威特其他机会", "政务、知识库及配置清单材料"), ("阿曼", "油气方案交流"), ("白俄罗斯", "客户接待与技术沟通")]
ow = (CW - 0.28 - 3 * 0.16) / 4
for i, (k, v) in enumerate(others):
    x = L + 0.14 + i * (ow + 0.16)
    text(s, x, oy + 0.38, ow, 0.28, k, size=10, color=BLUE, bold=True, valign="m", sa=0)
    text(s, x, oy + 0.64, ow, 0.38, v, size=9.5, color=TXT, valign="t", sa=0)
bar(s, oy + 1.2, "管理价值：在“客户提出需求”与“公司正式承诺”之间，持续开展竞争分析、范围评审和交付条件判断。", h=0.55, size=11)
notes(s, script_notes(18))

# ═════════════════════════ P21 东南亚 ═════════════════════════
s = content_slide("1-管理角色澄清", "东南亚：将应用方案推进到总集边界和采购材料", role=f"{BZ} ｜ 目标管理类")
lw = 8.7
text(s, L, CT, lw, 0.34, "SDEC（东马）阶段推进", size=12, color=NAVY, bold=True, valign="m", sa=0)
bands = [
    ("本人判断与组织", BLUE, ["华为提出讯飞担任总集并集成其他伙伴产品：提前识别利润与责任风险，讨论总集角色、边界与策略", "识别员工助手方案匹配卡点，推动一网通牵头整体方案整合", "与主责咨询讨论进入策略，组织材料刷新与内部对齐；安排咨询赴现场"], 1.6),
    ("团队与专业产出", BLUE2, ["第二轮方案材料，持续细化农业数字人、员工助手等需求", "HLD编写；SoC、项目范围及报价合理性材料评审"], 0.95),
    ("已形成的产出 / 状态", GREEN, ["8月完成细节对齐，拆为5个子项目；相应SoC已发送客户", "等待商务及客户内部流程，不等同签约"], 0.95),
]
y = CT + 0.4
for lab, col, items, h in bands:
    band(s, L, y, lw, h, lab, items, cols=1, color=col, label_w=1.55, size=10)
    y += h + 0.1
rx = L + lw + 0.28
rw = R - rx
titled_card(s, rx, CT, rw, 1.5, "泰国", ["根据预算约束调整方案和报价；与华为对齐MaaS、Agent平台及知识转移方向"], title_size=11.5, size=9.5)
titled_card(s, rx, CT + 1.62, rw, 1.25, "斯里兰卡", ["完成二期援助方案初版并内部对齐；不等同援助批准或实施"], title_size=11.5, size=9.5)
titled_card(s, rx, CT + 2.99, rw, 1.0, "其他线索", ["按需求和预算变化动态跟进，不与成熟采购项目同口径评价"], title_size=11.5, size=9.5)
bar(s, y + 0.08, "公司价值：让应用方案从“功能组合”转为明确范围、合作责任和采购输入，提前暴露总集与利润风险。", h=0.55, size=11)
notes(s, script_notes(19))

# ═════════════════════════ P22 咨询资源配置 ═════════════════════════
s = content_slide("1-管理角色澄清", "咨询资源配置：按机会成熟度和可实施性作投入取舍", role=f"{BZ} ｜ 目标管理类")
table(s, L, CT, [2.6, 6.0, 3.93], ["管理判断", "实际案例与动作", "对资源和经营的意义"], [
    ["无预算或无明确需求，不长期等量投入", "4月周报提出减少无预算、需求不清线索的消耗，聚焦有预算和落地可能的机会；与主责咨询分析乌兹TS AI并提出关闭建议", "为资源优先级提供依据；建议不等同于所有线索已正式关闭"],
    ["大项目条件不足，调整进入路径", "对沙特STC当年模型训练机会不足，提出先以小型应用或人才培训切入", "保留客户关系的同时，讨论与当期条件匹配的推进方式"],
    ["按预算和能力收敛工作范围", "IGNITE讨论最小工作范围；泰国依据预算调整平台和服务方案", "避免以宽泛承诺争取机会后由后端承担不可控投入"],
    ["总集与非标准内容，先审责任和成本", "SDEC总集及伙伴集成、KNG范围差异、埃及合同输入不足等问题进入评审", "将潜在风险前移，支持更清楚的参与和承诺决策"],
    ["有限人员优先保障关键节点", "巴西、埃及驻场和轮换并发时，结合项目阶段调度咨询，必要时本人接续或引入专业支撑", "保持关键任务连续性，并兼顾其他区域的成熟机会"],
], row_h=0.7, size=9.5)
py = CT + 0.42 + 5 * 0.7 + 0.18
text(s, L, py, 2.6, 0.32, "回顾归纳的投入原则", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
pr = ["客户是否明确", "预算与决策是否具备", "我方差异化是否成立", "承诺是否可交付", "当前最值得投入的下一步"]
pw = (CW - 2.7 - 4 * 0.12) / 5
for i, t in enumerate(pr):
    chip(s, L + 2.7 + i * (pw + 0.12), py + 0.02, pw, 0.3, t, fill=TINT2, color=NAVY, size=9, radius=0.3)
note(s, py + 0.42, "价值边界：已有记录证明进行了策略与资源判断；尚无统一统计支持投入减少比例、转化率提升或成本节约金额。")
notes(s, script_notes(20))

# ═════════════════════════ P23 方案资产与产品反馈 ═════════════════════════
s = content_slide("1-管理角色澄清", "方案资产与产品反馈：将项目经验用于其他国家和内部研发讨论", role=f"{BZ} ｜ 运营类")
rows = [
    ["国别定制方案", "巴西、埃塞俄比亚、巴基斯坦、加纳等国别方案，按预算、数据和客户阶段持续迭代", "明确价值逻辑与关键方向，组织产出与评审", "已有项目使用"],
    ["跨国方案基线", "加纳与巴西方案基线拉通；调用巴西、加纳经验支持IGNITE讨论", "推动不同现场的信息、范围和方法相互参照", "已有复用"],
    ["模型选型与表达口径", "针对埃及开闭源模型选择等问题，整理现场反馈形成团队口径材料", "复盘关键问题，提炼面向不同客户的表达重点", "已有复用"],
    ["LTC相关材料经验", "HLD、SoC、RFI、技术建议书、概算、验收及合同范围在各项目持续产出", "评审关键内容，推动材料与客户阶段、交易条件衔接", "样例积累"],
    ["产品与研发需求反馈", "参与开源大模型方案研发开工、三大中心国际版研发及相关评审讨论", "将海外客户诉求、竞争与实施约束带入内部讨论", "研发输入"],
    ["应用演示与新工作方式", "AI辅助场景Demo探索与试点；后续有AI Connection Demo等产出记录", "推动咨询侧尝试更直接的场景验证方式", "试点中"],
]
tbl_rows = [[r[0], r[1], r[2], (r[3], {"bold": True, "color": STATUS[r[3]][0]})] for r in rows]
table(s, L, CT, [2.0, 5.2, 3.9, 1.43], ["资产 / 输入类型", "已记录的工作与使用", "本人作用", "状态"], tbl_rows, row_h=0.72, size=9.5, aligns=["l", "l", "l", "c"])
card(s, L, CT + 0.42 + 6 * 0.72 + 0.16, CW, 0.62, fill=GRAY_T)
text(s, L + 0.15, CT + 0.42 + 6 * 0.72 + 0.16, CW - 0.3, 0.62, [("**组织价值**　让一次项目投入不仅服务当前客户，也为其他国家的价值呈现、范围讨论和材料准备留下可用经验。", {}), ("**建设边界**　复用已经发生；完整版本治理、使用反馈和统一知识库仍需继续完善。", {})], size=9.5, valign="m", sa=1)
notes(s, script_notes(21))

# ═════════════════════════ P24 伙伴赋能与高层交流 ═════════════════════════
s = content_slide("1-管理角色澄清", "伙伴赋能与高层交流：把公司能力转化为客户和伙伴可理解的方案", role=f"{BZ} ｜ 岗位特性类")
table(s, L, CT, [1.9, 4.6, 3.1, 2.93], ["工作场景", "具体工作与产出", "本人与团队分工", "对业务的作用"], [
    ["MWC等行业活动", "完成国家主权大模型圆桌方案分享与相关营销工作", "本人参与价值呈现，团队共同准备材料", "面向目标客户和伙伴的方案表达，不仅宣传单项产品"],
    ["华为区域与代表处赋能", "面向拉美、南部非洲、中亚及相关代表处开展主权AI、平台和场景方案培训；8月完成南部非洲高阶训战；北非智算周材料完成并通过伙伴评审", "本人统筹重点内容和资源，主责咨询承担部分现场分享", "帮助伙伴理解讯飞可承担的价值、场景与实施内容"],
    ["重点客户高层交流", "SERPRO来访及后续深圳交流；MGI高层接待远程协调与谈参准备；其他政企客户方案和技术交流", "根据对象组织背景、目标、方案和演示，各环节由对应人员承担", "将合作兴趣转为可继续对接的业务议题和材料"],
    ["集团级跨产线接待", "金光集团大模型板块接待：拉通工业智能、企业数字化、金融、行业大师等能力，准备展厅演示、汇报材料并现场汇报", "本人承担军团侧总体协调和相关汇报，产线提供专业演示与方案", "把分散能力组织为客户能够理解的整体价值"],
    ["区域行业与应用交流", "阿曼油气、白俄来访、哈萨克斯坦平台等交流形成方案和反馈", "组织咨询与销售、产线协同，团队面向客户交付", "为后续客户需求和项目推进建立信息基础"],
], row_h=0.8, size=9.5)
note(s, CT + 0.42 + 5 * 0.8 + 0.14, "本人重点：明确每次交流的对象、希望推动的事项、需要呈现的能力和会后动作。成果口径：交流、材料和赋能为已完成工作；商机金额与收入不推算。", h=0.5)
notes(s, script_notes(22))

# ═════════════════════════ P25 协同机制与工作方法 ═════════════════════════
s = content_slide("1-管理角色澄清", "协同机制与工作方法：已开展实践，持续提高稳定性", role=f"{BZ} ｜ 运营类")
table(s, L, CT, [1.8, 4.6, 3.2, 2.93], ["已开展实践", "具体记录与做法", "形成的作用", "后续需要完善"], [
    ["集中办公与辅导", "合肥3次集中办公共18天、广州骨干集中办公4天及常态化辅导；5月北京集中办公开展具体项目评审", "将问题讨论、方案打磨和带教结合，不只做一般培训", "把关键案例和复盘结论形成可复用记录"],
    ["前后场协同接口", "识别销售与咨询协同、跨过咨询直接调用后端等问题并推动处理", "减少信息分散，需求与资源通过更清楚的接口组织", "明确各类项目的主责、协同和升级安排"],
    ["LTC工作梳理", "参与国际系统部LTC协同机制和作战地图梳理，从咨询侧提出诉求并参与宣贯", "将实际协作问题转为共同工作讨论", "执行效果、责任更新和例外处理持续跟踪"],
    ["关键材料复核", "对国别方案、SoC、标书、报价及合同范围进行评审，发现风险后组织产线、法财或服务介入", "把质量与承诺问题前移", "形成统一的评审触发条件和状态留痕"],
    ["AI辅助Demo探索", "5月底开展快速构建Demo研究，6月进行咨询侧试点与推广", "尝试让咨询更直接地展示和验证场景", "继续验证适用范围、交付质量及实际投入效果"],
], row_h=0.78, size=9.5)
note(s, CT + 0.42 + 5 * 0.78 + 0.14, "管理价值：让经验不仅留在个人处理问题的过程里，也逐步进入团队可理解、可复用的工作方式。阶段判断：已有实践不等于全部制度化，提效探索不等于已量化。", h=0.5)
notes(s, script_notes(23))

# ═════════════════════════ P26 各项展开述职：角色自评汇总 ═════════════════════════
s = content_slide("1-管理角色澄清", "各项展开述职：四类角色的完成情况与可完善之处", role="承上启下")
table(s, L, CT, [1.7, 5.5, 3.5, 1.83], ["管理角色", "完成较好之处与关键举措", "可完善之处", "自评"], [
    ["岗位特性类", "关键咨询亲自上场；技术路线获认可并进入TR；多国国别方案形成（P8—P10、P19）", "标准口径材料覆盖不全，部分判断仍依赖本人", ("已形成", {"bold": True, "color": GREEN})],
    ["目标管理类", "政府间文件推进、交易与采购准备、SERPRO PoC、商机取舍（P7、P11—P15、P22）", "商业闭环未完成；前置事项可更早明确", ("推进中", {"bold": True, "color": AMBER})],
    ["人与团队类", "区域主责配置、骨干承接、销售导入（P16、P17）", "单点依赖仍在；主备机制待完善", ("推进中", {"bold": True, "color": AMBER})],
    ["运营类", "采购与合同书面化、本地承接矛盾处理、SoC与合同评审、机制建设（P12、P13、P18、P25）", "标准化与资产管理不足", ("已形成", {"bold": True, "color": GREEN})],
], row_h=0.9, size=10.5, aligns=["l", "l", "l", "c"])
ly = CT + 0.42 + 4 * 0.9 + 0.3
lenses = [("看定位", "是否准确 → P4 角色表、P2 总览"), ("看结果", "是否达到预期 → P6 巴西全景、P16 组织建设"), ("看过程动作", "是否扎实到位 → P7、P9、P13 案例"), ("看时间精力", "是否充分保障 → P5 同期投入")]
for i, (k, v) in enumerate(lenses):
    x = L + i * (w4 + 0.2)
    card(s, x, ly, w4, 0.8, fill=TINT2)
    text(s, x + 0.12, ly + 0.05, w4 - 0.24, 0.32, k, size=11, color=BLUE, bold=True, valign="m", sa=0)
    text(s, x + 0.12, ly + 0.37, w4 - 0.24, 0.4, v, size=9.5, color=TXT, valign="t", sa=0)
notes(s, "主讲1分钟。把两条线的案例收回四类角色，回应模板“各项展开述职”的要求：完成较好的原因与关键举措在中列，可完善之处在右列；没有良好完成的原因与改进见下一页。")

# ═════════════════════════ P27 不足与改进 ═════════════════════════
s = content_slide("2-不足分析与改进", "不足与改进：经营闭环、主备承接与成果口径仍需加强")
text(s, L, CT, 6, 0.32, "1. 当前经营状态与本人改进责任", size=11, color=NAVY, bold=True, valign="m", sa=0)
table(s, L, CT + 0.36, [1.6, 3.5, 4.1, 3.33], ["事项", "当前状态 / 客观依赖", "本人能够进一步改进的工作", "检验方式"], [
    ["RNP商业闭环", "正式合同、资金及客户内部流程尚未全部完成，部分事项依赖巴方与伙伴", "更早明确我方签约条件、关键文件状态和责任人；对外部依赖及时升级并保持备选安排", "关键事项是否逐项明确；我方材料与专业条件是否及时具备；未决问题是否透明"],
    ["SERPRO应用转化", "开发成果与最终客户确认、资源、验收和采购之间仍有距离", "将业务确认、测试、部署、预算和采购路径同步管理，而非只跟技术进度", "验收依据是否明确；关键资源和合同条件是否有责任与状态"],
], row_h=0.74, size=9.5)
oy = CT + 0.36 + 0.42 + 2 * 0.74 + 0.22
text(s, L, oy, 6, 0.32, "2. 组织能力仍需提升的方面", size=11, color=NAVY, bold=True, valign="m", sa=0)
table(s, L, oy + 0.36, [1.9, 5.1, 5.53], ["改进方向", "当前表现", "下一阶段举措"], [
    ["关键知识与主备", "部分项目完整背景和重大判断集中在本人及少数骨干", "明确主备、背景交接和关键承诺升级机制，逐步提高骨干独立承接"],
    ["成果与投入跟踪", "各区域材料、客户反馈和阶段状态需要更统一的口径，便于比较投入效果", "建立项目级成果记录，区分已完成产出、客户确认、商业结果与待决事项"],
    ["经验复用与版本管理", "已有跨项目复用，但部分仍依赖人员主动交流和临时组织", "明确基线材料责任、适用条件、版本更新与复用反馈"],
    ["授权和节奏", "亲自补位有助于当前任务，但可能挤占组织建设时间", "将必须本人介入的事项与骨干可自主处理的事项分层，固定关键评审与复盘节点"],
], row_h=0.56, size=9.5)
note(s, oy + 0.36 + 0.42 + 4 * 0.56 + 0.1, "改进原则：不以外部依赖取消自身责任，也不以个人持续超额投入替代组织能力建设。")
notes(s, script_notes(24))

# ═════════════════════════ P28 巴西下一阶段 ═════════════════════════
s = content_slide("3-工作规划与路径", "巴西下一阶段：围绕合同、应用验收和本地承接推进", role=f"{SP}")
OWN = {"本人统筹": BLUE, "专业确认": BLUE2, "外部决定": GRAYC}


def own_cell(*labels):
    return [(lab, {"bold": True, "color": OWN[lab], "size": 9}) for lab in labels]


table(s, L, CT, [1.6, 2.9, 4.2, 2.63, 1.2], ["工作方向", "目标结果", "关键动作与交付物", "主要依赖 / 资源", "责任属性"], [
    ["RNP签约条件收口", "我方合同材料和必要主体条件具备，客户可按其程序推进正式采购与签署", "统一TR应答、SOW、BoQ、验收、责任、实施计划及付款安排；形成未决事项清单和各专业确认记录", "RNP采购与审批、伙伴范围和材料、法务财务及方案服务支持", own_cell("本人统筹", "专业确认")],
    ["政府合作与预算事项", "支持文件、授权和资金相关状态持续明确，不成为信息盲区", "配合集团与伙伴提供所需材料，跟踪文本和审批进展，及时报告对签约节奏的影响", "政府部门、使馆与客户内部程序", own_cell("本人统筹", "外部决定")],
    ["本地实体及供应商准入", "满足本地签约和后续开票、支持所需条件", "跟进注册、CNPJ、经营范围与服务分类、准入材料、本地专业服务及行政接口", "财务、法务、律所、会计师和登记程序", own_cell("专业确认", "外部决定")],
    ["合同与交付衔接", "签约后可按已确认范围开展工作", "形成交付基线、客户配合条件、数据算力及人员安排、培训目标、变更和升级机制", "客户数据、算力、专家及受训人员，内部研发和服务团队", own_cell("本人统筹", "专业确认")],
    ["SERPRO应用转化", "完成最终业务确认、部署测试和PoC验收，并进入明确的采购讨论", "确认函、测试及验收清单、部署环境、预算与上架主体方案、商业条款输入", "MGI-SGP、SERPRO、华为云及内部产品研发和商务支持", own_cell("本人统筹", "外部决定")],
    ["后续经营基础", "在现有项目范围内建立可持续承接安排", "明确本地人员与总部支持分工，评估后续客户机会，不在现有能力外无边界扩张", "组织授权、人员预算和专业支持", own_cell("本人统筹", "专业确认")],
], row_h=0.66, size=9, header_size=9.5)
fy = CT + 0.42 + 6 * 0.66 + 0.14
card(s, L, fy, CW, 0.7, fill=TINT2)
text(s, L + 0.15, fy, CW - 0.3, 0.7, [("**节奏原则**　以签约前、签约后和PoC验收前后分别管理里程碑；外部条件变化时及时调整计划，保持我方可控事项准备程度。", {}), ("**需要组织支持**　关键节点集中投入的商法财与交付资源、现场主备和跨专业决策接口。", {})], size=9.5, valign="m", sa=1)
note(s, fy + 0.76, "本页为拟推进计划；所列交付物为拟形成结果，不代表均已形成或已获客户确认；角色分工与授权待组织确认。")
notes(s, script_notes(25))

# ═════════════════════════ P29 国际咨询下一阶段 ═════════════════════════
s = content_slide("3-工作规划与路径", "国际咨询下一阶段：提高骨干独立性、重点机会推进和经验复用", role=f"{BZ}")


def mod_cell(group, name, color):
    return [(group, {"size": 8.5, "color": color, "bold": True}), (name, {"size": 10, "color": NAVY, "bold": True})]


table(s, L, CT, [1.9, 3.2, 4.3, 3.13], ["工作模块", "拟达成的结果", "关键举措与交付物", "建议检验方式"], [
    [mod_cell("近期保障", "主责与备份", BLUE), "重点项目有明确主责、备份和关键事项升级安排", "核对区域与项目名册；形成客户背景交接、当前问题及材料清单；开展带教与复核", "主责离场后是否有明确接续；骨干可以独立完成哪些任务"],
    [mod_cell("近期保障", "重点机会分级投入", BLUE), "咨询资源与客户预算、决策阶段、我方能力和当期里程碑匹配", "对重点机会定期复核继续、调整或停止；明确下一步最有价值的产出", "是否减少重复无效讨论；客户需求、材料和阶段变化是否清楚"],
    [mod_cell("近期保障", "方案与承诺评审", BLUE), "非标准、总集、模型训练、私有化及重大权利承诺在对外前得到专业确认", "与方案、产品、服务、法财建立必要评审入口，保留结论及未决事项", "是否在报价或合同前识别并处理重大范围、成本和交付风险"],
    [mod_cell("组织效率", "案例与版本沉淀", GREEN), "形成可检索、可维护、可复用的关键材料和案例", "建立国别方案、技术问答、SoC、验收与合同问题样例；明确版本责任人和适用前提", "实际复用项目、采用反馈和更新记录，而非仅统计上传数量"],
    [mod_cell("组织效率", "客户与伙伴表达", GREEN), "伙伴更清楚讯飞的方案价值和实施边界，客户交流更有明确目的", "按对象维护高层版、业务版、技术版材料；交流后记录需求、反馈与下一步", "交流是否形成具体后续事项，关键材料是否进入客户讨论或需求"],
    [mod_cell("组织效率", "管理效果跟踪", GREEN), "对咨询工作形成统一但不过度复杂的成果口径", "区分本人、团队与专业部门贡献；记录产出、客户确认、商业阶段及当前风险", "用真实阶段推进评价支撑，不把文件数量直接等同收入"],
], row_h=0.68, size=9, header_size=9.5, first_col_bold=False)
fy = CT + 0.42 + 6 * 0.68 + 0.14
card(s, L, fy, CW, 0.6, fill=TINT2)
text(s, L + 0.15, fy, CW - 0.3, 0.6, "**实施节奏**　先完成重点项目与人员主备梳理，再选择高价值案例试运行统一记录和评审，随后按反馈扩大复用。具体周期与数量目标由部门确定。", size=9.5, valign="m", sa=0)
notes(s, script_notes(26))

# ═════════════════════════ P30 自我发展 ═════════════════════════
s = content_slide("4-自我发展", "自我发展：经营判断、授权育人与跨组织协同")
devs = [
    ("经营判断", "将技术方案、客户价值、工作范围、成本、付款和风险更系统地放在同一判断框架内", "深入参与重点项目报价与合同复盘；与商务、财务和服务共同检验关键假设；对重大方案明确收入、资源与责任关系", "对重大机会能形成清楚的方案选择依据、前提条件和风险建议，而不只给出技术判断"),
    ("授权与育人", "从本人多处补位，进一步转向骨干主责、必要复核和明确升级", "按事项重要性和成熟度分层授权；选择真实客户与项目任务带教；建立主备交接和关键评审节点", "骨干独立承担任务的范围扩大，关键节点质量可保持，本人能投入更多总体判断和组织建设"),
    ("跨组织协同", "降低对反复临时沟通的依赖，提高问题处理的可追踪性与决策效率", "把重复出现的问题整理成书面输入、责任接口和升级安排；持续复盘政府、伙伴、客户与内部协作方式", "关键问题有明确责任、状态和下一步，信息不因人员变化而丢失"),
]
for i, (t, g, a, c) in enumerate(devs):
    x = L + i * (w3 + 0.26)
    card(s, x, CT, w3, 4.55, fill=TINT)
    dot(s, x + 0.16, CT + 0.16, 0.36, str(i + 1))
    text(s, x + 0.62, CT + 0.14, w3 - 0.74, 0.4, t, size=13, color=NAVY, bold=True, valign="m", sa=0)
    for j, (lab, col, body, hh) in enumerate([("需提升之处", GRAYC, g, 0.95), ("关键行动", BLUE, a, 1.25), ("检验方式", GREEN, c, 1.0)]):
        yy = CT + 0.72 + sum([0.95, 1.25, 1.0][:j]) + j * 0.12
        chip(s, x + 0.16, yy, 0.95, 0.28, lab, fill=col, size=9)
        text(s, x + 0.16, yy + 0.3, w3 - 0.32, hh - 0.32, body, size=9.5, color=TXT, valign="t", sa=0)
note(s, CT + 4.65, "发展方向：保留关键问题亲自攻坚的能力，同时提升以经营结果为导向的判断、授权和组织协同水平。")
notes(s, script_notes(27))

# ═════════════════════════ P31 课程学习自检 ═════════════════════════
s = content_slide("5-课程学习自检", "课程学习自检")
text(s, L, CT, CW, 0.4, "以干部学习平台和考试记录为准，逐项确认必修课程及补齐安排。", size=11, color=NAVY, valign="m", sa=0)
table(s, L + 0.6, CT + 0.6, [3.6, 1.6, 2.0, 2.4, 1.73], ["课程名称", "课程类型", "学习完成时间", "考试 / 考核结果", "需补齐事项"], [
    ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""],
], row_h=0.5, size=10, first_col_bold=False)
note(s, CT + 0.6 + 0.42 + 5 * 0.5 + 0.2, "当前材料已记录参与集团干部培训，但尚不足以确认全部应知应会课程的学习与考试结果；正式提交前按平台记录补齐，未完成项如实列示。")
notes(s, script_notes(28))

# ═════════════════════════ P32 对公司和部门的建议 ═════════════════════════
s = content_slide("6-其他建议", "对公司和部门的建议：完善战略项目责任与出海专业支撑")
card(s, L, CT, CW, 0.5, fill=TINT2, radius=0.3)
chip(s, L + 0.12, CT + 0.09, 1.0, 0.32, "建议目标", size=9.5)
text(s, L + 1.22, CT, CW - 1.35, 0.5, "让重大海外项目既有持续的总体责任，又有清晰的专业接口和必要资源；减少每个项目反复依赖临时补位。", size=10.5, color=NAVY, bold=True, valign="m", sa=0)
table(s, L, CT + 0.64, [2.0, 3.6, 4.2, 2.73], ["建议", "来自实战的共性问题", "建议采取的组织安排", "预期价值"], [
    ["明确战略项目总体责任与决策接口", "项目横跨销售、咨询、方案、服务、法财及伙伴，部分关键问题没有单一专业部门可以独立关闭", "明确总体负责人、客户接口、各专业责任人及重大承诺决策链；责任与必要协调权限相匹配", "让问题有承接、决策有依据、对外口径一致，避免名义负责与实际责任脱节"],
    ["前置出海专业支撑", "本地签约主体、税务开票、用工、软件权利和伙伴合同等问题往往在临近签约时集中出现", "在重大机会达到相应成熟度时，组织法财、服务和本地顾问前置介入，沉淀可复用问题清单与材料样例", "提高交易准备效率，减少每个国家重复探索与后期返工"],
    ["将咨询资源与机会成熟度匹配", "重大项目驻场与多个区域小颗粒需求并发，队伍容易受到低成熟度机会消耗", "完善重点商机投入、跨区支援、轮换和主备安排；对长周期驻外及跨时区工作给予匹配保障", "稳定关键现场承接，提高资源使用针对性与人员可持续投入"],
], row_h=1.15, size=10)
note(s, CT + 0.64 + 0.42 + 3 * 1.15 + 0.15, "建议实施方式：先在重点项目中明确责任和专业介入节点，再复盘形成部门级可复用安排，不一次性增加过重流程。以上为组织建议，不是已批准制度或本人新增授权。", h=0.5)
notes(s, script_notes(29))

# ═════════════════════════ P33 总结与交流（模板结尾页） ═════════════════════════
PAGE[0] += 1
end_tb = [sh for sh in end.shapes if sh.has_text_frame][0]
end_tb.top, end_tb.height = Inches(0.95), Inches(0.95)
end_tb.left, end_tb.width = Inches(2.87), Inches(7.59)
replace_text(end_tb, ["总结与交流"], size=40)
text(end, 1.3, 1.95, 10.73, 0.5, "战略专项·巴西取得重大阶段突破　｜　岗位本职·国际咨询形成重点区域实际支撑", size=13, color=NAVY, bold=True, align="c", valign="m", sa=0)
contribs = [("持续承担总体责任", "在巴西客户、政府合作、华为伙伴和内部专业团队之间持续推进关键问题，将合作意向、技术方案和交易前提连接起来"),
            ("关键专业问题亲自攻坚", "直接用英语处理自主训练、语料、跨生态、效果验收和权利边界等问题，帮助客户理解可实施路径，并参与主要技术需求形成"),
            ("建设团队并组织跨区域产出", "招募培养骨干，安排现场与跨区支撑，持续评审重点机会和风险，将部分个人经验转为团队承接与复用")]
ex = 1.3
ew = (10.73 - 2 * 0.24) / 3
for i, (t, d) in enumerate(contribs):
    x = ex + i * (ew + 0.24)
    card(end, x, 2.6, ew, 1.5, fill=TINT)
    dot(end, x + 0.14, 2.74, 0.32, str(i + 1), size=10)
    text(end, x + 0.55, 2.7, ew - 0.65, 0.4, t, size=11.5, color=NAVY, bold=True, valign="m", sa=0)
    text(end, x + 0.14, 3.1, ew - 0.28, 0.95, d, size=9.5, color=TXT, valign="t", sa=0)
lw2 = (10.73 - 0.24) / 2
card(end, ex, 4.25, lw2, 1.6, fill=TINT2)
text(end, ex + 0.15, 4.3, lw2 - 0.3, 0.32, "下一阶段：三项我方可控的承诺", size=11, color=BLUE, bold=True, valign="m", sa=0)
text(end, ex + 0.15, 4.62, lw2 - 0.3, 1.2, [("RNP签约条件我方侧全部具备：材料、主体、专业确认", {"bullet": True}), ("SERPRO完成最终业务确认与PoC验收，并明确采购路径", {"bullet": True}), ("重点项目主备名单与授权边界清单落地，减少单点依赖", {"bullet": True})], size=9.5, sa=2)
card(end, ex + lw2 + 0.24, 4.25, lw2, 1.6, fill=GREEN_T)
text(end, ex + lw2 + 0.39, 4.3, lw2 - 0.3, 0.32, "需要组织的支持", size=11, color=GREEN, bold=True, valign="m", sa=0)
text(end, ex + lw2 + 0.39, 4.62, lw2 - 0.3, 1.2, [("关键节点集中投入的商法财、方案与服务资源", {"bullet": True, "bcolor": GREEN}), ("重大项目的总体责任人与各专业决策接口", {"bullet": True, "bcolor": GREEN}), ("现场主备、轮换与跨区支援的人员保障", {"bullet": True, "bcolor": GREEN})], size=9.5, sa=2)
text(end, 1.3, 6.0, 10.73, 0.4, "阶段成果已形成；正式商业闭环与组织稳定承接是下一阶段的责任。请各位领导指导。", size=11, color=NAVY, bold=True, align="c", valign="m", sa=0)
notes(end, script_notes(30) + "\n\n【管理层追问准备】\n" + _clean(QA_TEXT))


# ═════════════════════════ 垂直居中 / 删除模板说明页 / 排序 ═════════════════════════
def vcenter(slide, target_bottom=6.8):
    shapes = [sh for sh in slide.shapes if not sh.is_placeholder and CT - 0.05 <= sh.top / 914400 < 6.95]
    if not shapes:
        return
    bottom = max((sh.top + sh.height) / 914400 for sh in shapes)
    if bottom >= target_bottom - 0.35:
        return
    delta = (target_bottom - bottom) / 2
    for sh in shapes:
        sh.top = sh.top + Inches(delta)


for sl in list(prs.slides)[10:]:
    vcenter(sl)

sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
keep_tpl = {0, 1, 9}
for i in sorted(range(10), reverse=True):
    if i not in keep_tpl:
        prs.part.drop_rel(ids[i].rId)
        sldIdLst.remove(ids[i])
ids = list(sldIdLst)
cover_id, toc_id, end_id = ids[0], ids[1], ids[2]
new_ids = ids[3:]
order = [cover_id, new_ids[0], toc_id] + new_ids[1:] + [end_id]
for el in list(sldIdLst):
    sldIdLst.remove(el)
for el in order:
    sldIdLst.append(el)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
