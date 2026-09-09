# -*- coding: utf-8 -*-
"""
iFLYTEK x RNP Technical Exchange — English deck (mirrors RNP-讯飞交流-v22).
14 main slides + Q&A index + 15 appendix slides, on the Huawei briefing template.

Translation notes
  - Stage names follow the joint Huawei x iFLYTEK talent deck: Foundation / Shadowing /
    Guided Hands-on / Formal Handover.
  - 课时 -> hours, 工作日 -> working days, 人次 -> attendances.
  - Blue text = to be confirmed jointly before implementation (same convention as the CN deck).
  - Text is compressed relative to the CN deck to keep English pages readable.

Usage: python3 build_en.py <template.pptx> <out.pptx>
"""
import sys

import os

import pptx_helpers

# ---- 排版风格开关（可用环境变量覆盖，便于并排比稿）----
EN_FONT = os.environ.get("EN_FONT", "Calibri")        # Arial / Calibri / Segoe UI
LS_MIN = float(os.environ.get("EN_LS", "1.38"))       # 行距下限
SA_ADD = float(os.environ.get("EN_SA", "1"))         # 段后额外磅数
pptx_helpers.FONT = EN_FONT

_fill_tf = pptx_helpers.fill_tf


def _fill_tf_styled(tf, content, ls=1.15, sa=4, **kw):
    return _fill_tf(tf, content, ls=max(ls, LS_MIN), sa=sa + SA_ADD, **kw)


pptx_helpers.fill_tf = _fill_tf_styled

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import (BLUET, BODY, CW, GRAY, GRAY_T, L, LINE, MUTE, PANEL, PANEL2, R, RED, REDT, TXT, WHITE,
                          arrow, body, box, card, chip, dot, drop_first_slide, fill_tf, label, new_slide, note,
                          table, text)

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else "模板-华为简报.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "RNP-iFLYTEK-Exchange-EN.pptx"
RNP_C = "1F4E79"
SOW, PLAN, TBD = "SOW", "Approach", "{{TBC}}"
XF, RN, BOTH = REDT, BLUET, PANEL

prs = Presentation(TEMPLATE)
LAYOUT = prs.slide_layouts[0]

LINKS, APPX, BACK, INDEX = [], {}, [], {}
MAIN_OF, PAGENO, BACK2 = {}, {}, []
HOWTO = {}

APPX_CODES = ["A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14"]
NAV = []


def slide(title, size=20, ref=None):
    s = new_slide(prs, LAYOUT, title, size=size)
    PAGENO[id(s)] = len(prs.slides) - 1
    if ref:
        x = R
        for code in reversed([c.strip() for c in ref.split("·")]):
            w = 1.15
            x -= w + 0.08
            c = chip(s, x, 0.32, w, 0.3, "App. " + code, fill=WHITE, color=RNP_C, size=9.5, line=LINE)
            LINKS.append((c, code))
            MAIN_OF.setdefault(code, s)
    return s


def appx_nav(s, cur):
    y, h, gap = 6.60, 0.24, 0.055
    cw_ = 0.44
    x = L
    for code in APPX_CODES:
        if code == "A7":
            x += 0.18
        cell = box(s, x, y, cw_, h, fill=RED if code == cur else WHITE,
                   line=None if code == cur else LINE, radius=0.35, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        fill_tf(cell.text_frame, code, size=8, color=WHITE if code == cur else MUTE,
                bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
        if code != cur:
            NAV.append((cell, code))
        x += cw_ + gap
    b = chip(s, 8.22, y, 1.30, h, "Index", fill=WHITE, color=RNP_C, size=9.5, line=LINE)
    BACK.append(b)
    tgt = MAIN_OF.get(cur)
    if tgt is not None:
        c = chip(s, 9.62, y, 1.30, h, "Back to p.%d" % PAGENO[id(tgt)], fill=WHITE, color=RED, size=9, line=LINE)
        BACK2.append((c, tgt))


def appendix_slide(title, question):
    code = title.split(" ·")[0].strip()
    s = slide(title, size=19)
    chip(s, R - 4.6, 0.3, 4.6, 0.32, question, fill=WHITE, color=TXT, size=9.5, line=LINE)
    APPX[code] = s
    appx_nav(s, code)
    return s


def thesis(s, content, y=1.00, h=0.60):
    body(s, L, y, CW, h, content, size=13, color=TXT, valign="m", ls=1.25)


def legend(s, y, x=None):
    x = R - 4.6 if x is None else x
    for i, (f, t) in enumerate([(REDT, "iFLYTEK-led"), (BLUET, "RNP-led"), (PANEL, "Joint")]):
        bx = x + i * 1.55
        box(s, bx, y + 0.06, 0.22, 0.16, fill=f, line=LINE)
        text(s, bx + 0.28, y, 1.2, 0.28, t, size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)


def step_cards(s, y, items, gap=0.25, size=10, h=3.20):
    n = len(items)
    bw = (CW - (n - 1) * gap) / n
    name_h, ig = 0.42, 0.08
    blk = (h - name_h - ig - 0.12) / 2
    for i, (name, xf_t, rn_t) in enumerate(items):
        bx = L + i * (bw + gap)
        card(s, bx, y, bw, h, fill=WHITE, line=LINE)
        text(s, bx + 0.16, y + 0.05, bw - 0.32, 0.34, name, size=11.5, color=TXT, bold=True,
             valign="m", margin=(0, 0, 0, 0), sa=0)
        for k, (who, col, fl, t) in enumerate([("iFLYTEK", RED, XF, xf_t), ("RNP", RNP_C, RN, rn_t)]):
            by = y + name_h + k * (blk + ig)
            card(s, bx + 0.14, by, bw - 0.28, blk, fill=fl)
            text(s, bx + 0.28, by + 0.03, 1.6, 0.22, who, size=9, color=col, bold=True,
                 valign="m", margin=(0, 0, 0, 0), sa=0)
            body(s, bx + 0.28, by + 0.26, bw - 0.56, blk - 0.30, t, size=size, ls=1.25,
                 bullet=isinstance(t, list), num=True, bcolor=col, sa=2)
        if i < n - 1:
            arrow(s, bx + bw + 0.015, y + name_h + blk * 0.5 - 0.11, fill=LINE)


def two_boxes(s, y, h, xf, rnp, size=10.5, rnp_title="How RNP takes part"):
    half = (CW - 0.3) / 2
    for i, (name, color, fill, items) in enumerate([("iFLYTEK delivers", RED, REDT, xf), (rnp_title, RNP_C, BLUET, rnp)]):
        cx = L + i * (half + 0.3)
        card(s, cx, y, half, h, fill=fill)
        label(s, cx + 0.25, y + 0.1, 3.2, name, color=color, size=12)
        body(s, cx + 0.25, y + 0.5, half - 0.5, h - 0.56, items, size=size, bullet=True, num=True, bcolor=color, sa=3)


def value_chain(s, y, h, items, lab="Direct value in this project"):
    card(s, L, y, CW, h, fill=PANEL2, line=LINE)
    text(s, L + 0.26, y + 0.08, 5.0, 0.3, lab, size=11.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    n = len(items)
    inner, x0, aw = CW - 0.5, L + 0.25, 0.35
    w = (inner - (n - 1) * aw) / n
    cy, ch = y + 0.44, h - 0.54
    for i, (t1, t2) in enumerate(items):
        cx = x0 + i * (w + aw)
        card(s, cx, cy, w, ch, fill=WHITE, line=LINE)
        text(s, cx + 0.16, cy + 0.05, w - 0.32, 0.28, t1, size=10.5, color=TXT, bold=True,
             valign="m", margin=(0, 0, 0, 0), sa=0)
        body(s, cx + 0.16, cy + 0.30, w - 0.32, ch - 0.34, t2, size=9, color=MUTE, ls=1.2)
        if i < n - 1:
            arrow(s, cx + w + 0.06, cy + ch / 2 - 0.11, fill=LINE)


def outcomes(s, y, items, lab="Verifiable outputs"):
    text(s, L, y, CW, 0.34, "**" + lab + "**   " + "   ·   ".join(items),
         size=10.5, color=TXT, valign="m", margin=(0, 0, 0, 0), sa=0)


def means(s, y, content, h=0.6):
    card(s, L, y, CW, h, fill=WHITE, line=RNP_C)
    body(s, L + 0.25, y + 0.05, CW - 0.5, h - 0.1, "**What this means for RNP**   " + content,
         size=11.5, color=RNP_C, valign="m")


def band(s, y, h, title, content, fill=PANEL2, line=LINE, title_w=1.7, size=10.5, color=RED):
    card(s, L, y, CW, h, fill=fill, line=line)
    label(s, L + 0.25, y + (h - 0.34) / 2, title_w, title, color=color, size=11.5)
    body(s, L + 0.25 + title_w, y + 0.06, CW - 0.5 - title_w, h - 0.12, content, size=size, valign="m")


def spec_table(s, rows, y=1.15, colw=(2.6, 9.73), row_h=0.62, size=10, header=("Item", "Content")):
    table(s, L, y, list(colw)[:2], list(header)[:2], [r[:2] for r in rows], row_h=row_h, header_h=0.4,
          size=size, header_size=10.5, header_fill="F2F2F2", header_color=TXT, vlines=False)


# ============================== 1 Cover ==============================
s = slide("")
text(s, 0.8, 1.5, 8.6, 0.4, "RNP × iFLYTEK · Technical Exchange", size=14, color=MUTE, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 1.98, 8.6, 1.30, "National Model, National AI Application and Domain Models", size=27, color=RED, bold=True, margin=(0, 0, 0, 0), ls=1.1)
text(s, 0.8, 3.15, 8.6, 0.6, "What iFLYTEK delivers, how the work is split, how capability transfers", size=15, color=TXT, bold=True, margin=(0, 0, 0, 0))
text(s, 0.8, 4.1, 8.6, 1.3, [
    "Rio de Janeiro · 9 September 2026",
    "Presented by iFLYTEK: CHEN Hao / William Xu / ZHAO Xiang",
], size=13, color=BODY, margin=(0, 0, 0, 0), ls=1.35, sa=4)
bx, by, bw, bh, g = 9.9, 1.75, 2.9, 0.78, 0.18
for i, t in enumerate(["National AI Capability Center", "Brazilian Portuguese National Model",
                       "National AI Application", "Domain Large Models × 3"]):
    card(s, bx, by + i * (bh + g), bw, bh, fill=PANEL)
    text(s, bx + 0.28, by + i * (bh + g), bw - 0.45, bh, f"{i + 1}  {t}", size=11.5, color=TXT, bold=True,
         valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.15)
card(s, bx, by + 4 * (bh + g), bw, 0.72, fill=RED)
text(s, bx, by + 4 * (bh + g), bw, 0.72, "Knowledge Transfer", size=14, color=WHITE, bold=True,
     align="c", valign="m", margin=(0, 0, 0, 0), sa=0)

# ============================== 2 Overview ==============================
s = slide("The whole picture: four outcomes iFLYTEK delivers", ref="A0")
thesis(s, "All four are delivered by iFLYTEK; the platform is the shared workbench for the other three. One line each — details follow.")
boxes = [
    ("National AI Capability Center", ["Training and application platform", "All three outcomes built on it", "Run by RNP afterwards"], "3 centers + review platform"),
    ("Brazilian Portuguese National Model", ["Text, speech recognition, speech synthesis", "iFLYTEK enhances the base model first", "Weights delivered to RNP"], "≥290B parameters"),
    ("National AI Application", ["National App and operations platform", "For users approved by RNP", "iFLYTEK delivers, RNP operates"], "1,000 concurrent Q&A"),
    ("Domain Large Models (DLM)", ["Models for 3 industry scenarios", "RNP trains, iFLYTEK advises", "Weights belong to RNP"], "3 services × 1 scenario"),
]
gap = 0.28
bw = (CW - 3 * gap) / 4
by, bh = 1.86, 2.80
for i, (name, one, num) in enumerate(boxes):
    bx = L + i * (bw + gap)
    card(s, bx, by, bw, bh, fill=PANEL)
    text(s, bx + 0.22, by + 0.14, bw - 0.44, 0.62, f"{i + 1}  {name}", size=12.5, color=TXT, bold=True,
         valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    body(s, bx + 0.22, by + 0.86, bw - 0.44, 1.15, one, size=10, ls=1.3, bullet=True, num=True, bcolor=RED, sa=4)
    text(s, bx + 0.22, by + 2.15, bw - 0.44, 0.5, num, size=13, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, L, 4.90, CW, 0.62, fill=PANEL2, line=LINE)
body(s, L + 0.25, 4.92, CW - 0.5, 0.58,
     "**Knowledge transfer**   two rounds, 1,216 hours, about 158 working days, 260 attendances in total"
     "     **Post-delivery support**   12 months on site + 60 months remote", size=11, valign="m")
body(s, L, 5.70, CW, 0.7, [
    "Terms   1) Large language model (NLP) understands and generates text   2) ASR turns speech into text   "
    "3) TTS turns text into speech   4) A token is how a model counts text, roughly a word or part of one   "
    "5) Stage 1 / Stage 2 are the two build phases tied to the two compute batches",
], size=9.5, color=MUTE, ls=1.3)

# ============================== 3 Capability Center ==============================
s = slide("National AI Capability Center: the shared platform, and RNP's future workbench", size=18, ref="A1 · A12")
thesis(s, "All three outcomes are built on this platform. Each of the four modules maps to one RNP team. RNP manages it after acceptance. A 10-minute walkthrough follows.", h=0.5)
tops = ["Brazilian Portuguese National Model", "National AI Application", "Domain Large Models"]
tw = 3.7
tx0 = L + (CW - (3 * tw + 2 * 0.4)) / 2
for i, t in enumerate(tops):
    chip(s, tx0 + i * (tw + 0.4), 1.62, tw, 0.34, t, fill=WHITE, color=TXT, size=10, line=LINE)
py, ph = 2.02, 2.30
card(s, L, py, CW, ph, fill=PANEL2, line=LINE)
label(s, L + 0.25, py + 0.08, 5, "National AI Capability Center (platform)", color=RED, size=11.5)
mods = [
    ("Corpus Construction Center", "1) Ingest 2) Clean 3) De-duplicate 4) Annotate 5) Quality metrics", "RNP data team and Portuguese experts"),
    ("Model Training Center", "1) Datasets 2) Training jobs 3) Resources 4) Evaluation 5) Release", "RNP model engineers; DLMs trained here"),
    ("AI Application Innovation Center", "1) Model services and APIs 2) App configuration 3) Usage analytics", "RNP application and operations team"),
    ("Content Safety Review Platform", "1) Review policies 2) Sensitive lexicons 3) Risk samples", "RNP review and compliance team"),
]
gap = 0.22
mw = (CW - 0.5 - 3 * gap) / 4
my, mh = py + 0.48, ph - 0.62
for i, (name, what, team) in enumerate(mods):
    mx = L + 0.25 + i * (mw + gap)
    card(s, mx, my, mw, mh, fill=WHITE, line=LINE)
    text(s, mx + 0.16, my + 0.08, mw - 0.32, 0.46, f"{i + 1}  {name}", size=11, color=TXT, bold=True,
         valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    body(s, mx + 0.16, my + 0.58, mw - 0.32, 0.62, what, size=9.5, ls=1.3)
    body(s, mx + 0.16, my + 1.24, mw - 0.32, 0.42, "**Maps to:** " + team, size=9.5, ls=1.3)
two_boxes(s, 4.46, 1.60,
          ["Platform software and licences, deployment configuration, API guide, acceptance report, O&M documentation",
           "Deploy the platform and connect interfaces",
           "Product operations training: 6 topics, 12 O&M engineers per batch — see Appendix A12"],
          ["Provide the environment and interfaces",
           "Manage the platform after acceptance; four modules, four teams",
           "Organise acceptance and take over daily operations"], size=10)
means(s, 6.18, "Every step in this demo is something RNP staff will do themselves during shadowing and guided hands-on.")


# ============================== 4 National Model: what it delivers ==============================
s = slide("What the Brazilian Portuguese National Model delivers", size=20)
thesis(s, "A national foundation model for Brazil's language, knowledge, values and local needs. Weights, model card and evaluation report are delivered to RNP.", h=0.46)
qw, qgap = 4.55, 0.12
cw_c = CW - 2 * qw - 2 * qgap
qy, qh = 1.56, 1.36
cx_c = L + qw + qgap
card(s, cx_c, qy, cw_c, qh * 2 + qgap, fill=PANEL)
text(s, cx_c, qy + 0.40, cw_c, 0.42, "Brazilian Portuguese", size=13, color=RED, bold=True,
     align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
text(s, cx_c, qy + 0.72, cw_c, 0.42, "National Model", size=13, color=RED, bold=True,
     align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
box(s, cx_c + 0.6, qy + 1.22, cw_c - 1.2, 0.02, fill=LINE)
body(s, cx_c + 0.25, qy + 1.34, cw_c - 0.5, 0.95,
     ["NLP: ≥290B parameters", "ASR and TTS: one model each", "Local corpus: 50B tokens"],
     size=9.5, color=BODY, align="c", ls=1.35)
quads = [
    (0, 0, "Native Brazilian Portuguese",
     ["Understands Brazilian Portuguese and local usage", "Portuguese Q&A, generation, multi-turn dialogue"],
     "Natural phrasing  |  Spoken language  |  {{Regional variety}}"),
    (0, 1, "Brazilian knowledge and context",
     ["Built on Brazilian data and institutional context", "Reflects cultural, social and linguistic diversity"],
     "50B-token local corpus  |  Cultural diversity"),
    (1, 0, "Values and accountability",
     ["Data held in Brazil; access and audit belong to RNP", "Supports LGPD compliance; 20,000 safety and values samples"],
     "Bias and safety testing  |  Fairness and inclusion"),
    (1, 1, "A base RNP controls",
     ["Reduces long-term dependence on external models", "Weights delivered to RNP; supports the DLMs"],
     "Weights to RNP  |  Deployed in Brazil"),
]
for k, (col, row, name, items, tags) in enumerate(quads):
    x = L if col == 0 else cx_c + cw_c + qgap
    y = qy + row * (qh + qgap)
    card(s, x, y, qw, qh, fill=PANEL2, line=LINE)
    text(s, x + 0.24, y + 0.08, qw - 0.48, 0.30, "%d  %s" % (k + 1, name), size=11.5, color=TXT,
         bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, x + 0.24, y + 0.42, qw - 0.48, 0.58, items, size=9.5, bullet=True, num=True, bcolor=RED, sa=2)
    box(s, x + 0.24, y + 1.02, qw - 0.48, 0.02, fill=LINE)
    text(s, x + 0.24, y + 1.06, qw - 0.48, 0.26, tags, size=8.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.60, 1.32, [
    ("Brazilian Portuguese National Model", "Closer to Brazil's language, knowledge and needs"),
    ("Health / Space / Biodiversity DLMs", "Three domains named by RNP; built on this model or another agreed base"),
    ("AI applications and services for Brazil", "Supporting public services and national priorities"),
])
means(s, 6.05, "Performance and coverage are set against jointly agreed evaluation criteria — not decided by iFLYTEK alone. After delivery RNP deploys and evolves the model itself.", h=0.6)

# ============================== 5 National Model: how it is built ==============================
s = slide("National Model: iFLYTEK enhances the base first, RNP supplies the local corpus", size=18, ref="A2 · A3")
HOWTO["model"] = s
thesis(s, "Step 1 runs on two parallel tracks: iFLYTEK enhances the base model in China while RNP builds the local corpus in Brazil. Step 2 starts once corpus and compute are ready.", h=0.58)
step_cards(s, 1.70, [
    ("1  Corpus and base model",
     ["Collect public Portuguese corpus in China", "Strengthen the base model with own compute", "Public corpus handed to Brazil too"],
     ["Produce and license the corpus in Brazil", "About 50 billion tokens", "{{Arrange compute before the data center opens}}"]),
    ("2  Model training",
     ["Continue training on the enhanced base", "Produce the NLP, ASR and TTS models", "Follow the approved plan and milestones"],
     ["Keep supplying data to iFLYTEK's spec", "Support sampling quality checks", "Confirm route and rounds in writing"]),
    ("3  Evaluation",
     ["Issue the evaluation report", "Regression, robustness, bias", "Safety and cultural fit"],
     ["Agree test sets and pass rules jointly", "Portuguese and domain experts review"]),
    ("4  Local delivery and tuning",
     ["Deploy in the Brazilian environment", "Deliver weights and model cards", "NLP upgraded 3× within 12 months", "ASR and TTS once each"],
     ["Organise acceptance and release", "Supply upgrade data within the A2 caps"]),
], h=4.14)
means(s, 6.08, "Enhancing the base first uses no Brazilian compute and does not wait for the corpus to be complete. Training can start as soon as corpus and compute are in place — less schedule risk.", h=0.6)

# ============================== 6 National AI Application: what it delivers ==============================
s = slide("What the National AI Application delivers", size=20)
thesis(s, "A single AI entry point for users approved by RNP. iFLYTEK delivers the app and the Data Operations Platform; content, users and operations belong to RNP.", h=0.46)
aw = (CW - 2 * 0.3) / 3
ay, ah = 1.62, 2.55
acards = [
    ("National App · for approved users",
     ["Q&A and search, document analysis and translation, writing, notes and presentations",
      "Portuguese speech input and read-aloud; mobile, desktop and web"],
     "Single entry  |  Multi-device  |  Portuguese speech"),
    ("Data Operations Platform · for RNP administrators",
     ["Organisations, users, tenants and permissions; app registration and release",
      "Usage logs, visits, active users and feedback statistics"],
     "Users and tenants  |  Logs  |  Operations reports"),
    ("Built-in governance · compliance and autonomy",
     ["Connected to the Content Safety Review Platform, enforcing policy at agreed points",
      "Data and logs stay in Brazil; iFLYTEK QR-code login is not required to run locally"],
     "Content review  |  Access and logs  |  Runs locally"),
]
for i2, (name, items, tags) in enumerate(acards):
    ax = L + i2 * (aw + 0.3)
    card(s, ax, ay, aw, ah, fill=PANEL2, line=LINE)
    text(s, ax + 0.26, ay + 0.12, aw - 0.52, 0.50, "%d  %s" % (i2 + 1, name), size=11.5, color=TXT,
         bold=True, valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    body(s, ax + 0.26, ay + 0.70, aw - 0.52, 1.25, items, size=10, bullet=True, num=True, bcolor=RED, sa=6, ls=1.3)
    box(s, ax + 0.26, ay + ah - 0.44, aw - 0.52, 0.02, fill=LINE)
    text(s, ax + 0.26, ay + ah - 0.38, aw - 0.52, 0.26, tags, size=8.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.40, 1.32, [
    ("Model services and APIs", "NLP, speech recognition and synthesis, knowledge Q&A, content safety review"),
    ("National App and Data Operations Platform", "Mobile, desktop and web; 1,000 concurrent Q&A sessions"),
    ("Use in public services", "Users approved by RNP, supporting public services"),
])
means(s, 6.00, "iFLYTEK does not take on user growth, content production, human moderation or managed business operations — those sit with RNP's product, operations and review teams.", h=0.62)


# ============================== 7 National AI Application: how it is built ==============================
s = slide("National AI Application: iFLYTEK delivers the software, RNP owns content, users and operations", size=17, ref="A4 · A13")
HOWTO["app"] = s
thesis(s, "For users approved by RNP. 1,000 concurrent Q&A sessions; 200 each for speech recognition and synthesis, configured and tested per service.", h=0.46)
card(s, L, 1.54, CW, 1.06, fill=PANEL)
label(s, L + 0.26, 1.58, 4.6, "What the product is", color=TXT, size=11.5)
for i, t in enumerate(["Mobile app", "Desktop", "Data Operations Platform"]):
    chip(s, L + 2.60 + i * 2.55, 1.60, 2.40, 0.30, t, fill=WHITE, color=TXT, size=10, line=LINE)
funcs = ["Q&A", "Search", "Document analysis", "Document translation", "Speech in / read-aloud", "Notes and writing"]
fw6 = (CW - 0.52 - 5 * 0.12) / 6
for i, t in enumerate(funcs):
    fx = L + 0.26 + i * (fw6 + 0.12)
    card(s, fx, 2.20, fw6, 0.34, fill=WHITE, line=LINE)
    text(s, fx, 2.20, fw6, 0.34, t, size=9, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
step_cards(s, 2.64, [
    ("1  Define requirements",
     ["Provide the standard capability list", "Explain configurable items and limits"],
     ["Join product definition and design", "Set user scope, content rules and branding"]),
    ("2  Localisation",
     ["Portuguese interface and configuration", "Apply RNP branding and content"],
     ["Supply branding and editorial content", "{{Portuguese slides and official templates to be verified}}"]),
    ("3  System integration",
     ["Connect and test the interfaces", "Provide developer docs and samples"],
     ["Supply local interfaces and credentials", "Arrange third-party API quota and licences"]),
    ("4  Launch and operate",
     ["Stage 1 initial delivery", "Stage 2 extended operation", "Hand over O&M documentation"],
     ["Approve users", "Product, operations and review teams take over", "Evolve on the delivered version"]),
], h=3.28)
means(s, 6.14, "The product team is needed from the requirements stage and the operations team before Stage 1 delivery. RNP can evolve the delivered version on its own.", h=0.6)

# ============================== 8 DLM: what it delivers ==============================
s = slide("What the Domain Large Models deliver", size=20)
thesis(s, "RNP names three domains, each with one clearly bounded scenario. {{Scenarios are proposed by RNP and start once each confirmation form is signed.}}", h=0.44)
dom_w = (CW - 2 * 0.3) / 3
for i4, t4 in enumerate(["Health", "Space", "Biodiversity"]):
    chip(s, L + i4 * (dom_w + 0.3), 1.52, dom_w, 0.34, t4, fill=WHITE, color=TXT, size=11, line=LINE)
dw = (CW - 2 * 0.3) / 3
dy, dh = 1.96, 2.24
dcards = [
    ("Reads the domain's material",
     ["Understands the domain's terminology and standards", "Handles the document formats the domain uses"],
     "Terminology  |  Standards  |  Document structure"),
    ("Answers the domain's questions",
     ["Answers from domain knowledge and connected sources", "Search, synthesis and key-point extraction"],
     "Knowledge Q&A  |  Search  |  Key points"),
    ("Supports the domain's work",
     ["Analysis and decision support within the agreed scenario", "Synthesis of material and draft reporting"],
     "Scenario analysis  |  Decision support  |  Drafting"),
]
for i, (name, items, tags) in enumerate(dcards):
    dx = L + i * (dw + 0.3)
    card(s, dx, dy, dw, dh, fill=PANEL2, line=LINE)
    text(s, dx + 0.26, dy + 0.14, dw - 0.52, 0.36, "%d  %s" % (i + 1, name), size=12.5, color=TXT,
         bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, dx + 0.26, dy + 0.58, dw - 0.52, 1.10, items, size=10.5, bullet=True, num=True, bcolor=RED, sa=6, ls=1.3)
    box(s, dx + 0.26, dy + dh - 0.44, dw - 0.52, 0.02, fill=LINE)
    text(s, dx + 0.26, dy + dh - 0.38, dw - 0.52, 0.26, tags, size=8.5, color=MUTE,
         valign="m", margin=(0, 0, 0, 0), sa=0)
value_chain(s, 4.32, 1.26, [
    ("Brazilian Portuguese National Model", "As the base model, or another base agreed by both parties"),
    ("Health / Space / Biodiversity DLMs", "Three domains named by RNP, one fixed scenario each"),
    ("Applications for public services", "Supporting RNP's public services and national priorities"),
])
outcomes(s, 5.66, ["DLM plan", "Advisory records", "Training and evaluation report", "Service acceptance form", "Weights owned by RNP"])
means(s, 6.02, "RNP performs the training; iFLYTEK provides advisory training services. The models offer analysis and decision support — they do not replace professional judgement.", h=0.62)

# ============================== 9 DLM: how it is built ==============================
s = slide("Domain Large Models: RNP trains, iFLYTEK advises — 3 services, one scenario each", size=18, ref="A5")
HOWTO["dlm"] = s
thesis(s, "All four steps are executed by RNP with iFLYTEK advising — the reverse of the national model, and RNP's first time training a model itself.", h=0.44)
band(s, 1.48, 0.58, "Service scope",
     "Three advisory training services for industry models, one clearly bounded scenario each. RNP proposes the domain, iFLYTEK assesses feasibility. Weights are produced by RNP's training and belong to RNP.", size=10)
step_cards(s, 2.14, [
    ("1  Define the scenario",
     ["Assess feasibility of the domain", "Issue the technical plan", "Verify base model version, size, licence"],
     ["Propose the domain and scenario", "Define users, inputs, outputs and limits"]),
    ("2  Data and training route",
     ["Recommend the training route", "LoRA-SFT, SFT, GRPO and similar", "One main route confirmed per service"],
     ["Supply training, validation and safety data", "Supply independent test data and environment"]),
    ("3  Training and evaluation",
     ["Parameter advice within the agreed effort", "Review training logs", "Help analyse issues"],
     ["Run training in the Model Training Center", "Complete the evaluation"]),
    ("4  Service acceptance",
     ["Based on RNP's actual results", "Produce the analysis report"],
     ["Check the advisory records", "Complete business acceptance"]),
], h=3.52)
note(s, 5.78, "Production deployment, API development and application integration are outside this service and can be agreed separately.", h=0.26, size=9)
means(s, 6.10, "This is RNP's first time training a model independently. It needs RNP's own algorithm engineers and domain experts — which is exactly what shadowing and guided hands-on are for.", h=0.6)


# ============================== 10 Knowledge Transfer: what RNP gains ==============================
s = slide("Knowledge transfer: what RNP learns and what RNP gets", size=19, ref="A6 · A9 · A10 · A11")
thesis(s, "Twelve SOW course topics grouped into four tracks. Instructors are the iFLYTEK engineers and architects delivering this project; sessions run on the delivered platform and models.")
lw = CW * 0.58
cy, ch = 1.86, 3.05
card(s, L, cy, lw, ch, fill=PANEL)
label(s, L + 0.3, cy + 0.10, lw - 0.6, "What RNP learns: three levels", color=TXT, size=12)
tiers = [
    ("Core course · Foundation stage", "1) LLM basics 2) Project architecture 3) Corpus to service 4) Safety and O&M", "In scope"),
    ("12 specialist courses · four tracks", "1) Models and evaluation 2) Software stack and lifecycle 3) Training and inference at scale 4) Retrieval, agents and safety", "In scope"),
    ("Premium transfer · how the system was built", "1) Cluster scheduling 2) Model training 3) Corpus 4) Inference; plus seminars, consulting and a China visit", "{{Separate budget}}"),
]
ty = cy + 0.52
for i, (name, desc, status) in enumerate(tiers):
    yy = ty + i * 0.82
    card(s, L + 0.3, yy, lw - 0.6, 0.74, fill=WHITE, line=LINE)
    text(s, L + 0.45, yy + 0.04, lw - 2.5, 0.28, f"{i + 1}  {name}", size=11, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, L + lw - 2.05, yy + 0.06, 1.65, 0.24, status, size=8.5, color=MUTE, align="r", valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, L + 0.45, yy + 0.32, lw - 0.9, 0.40, desc, size=9, color=BODY, ls=1.25)
rx = L + lw + 0.3
rw = CW - lw - 0.3
card(s, rx, cy, rw, ch, fill=PANEL)
label(s, rx + 0.3, cy + 0.10, rw - 0.6, "What RNP gets", color=TXT, size=12)
body(s, rx + 0.3, cy + 0.56, rw - 0.6, ch - 0.66, [
    "**Model weights** — DLM weights belong to RNP",
    "**Platform software** and custom code",
    "**All documentation**, in English or Portuguese",
    "**Support**: 12 months on site + 60 months remote",
], size=10.5, bullet=True, num=True, bcolor=RED, sa=6, ls=1.25)
two_boxes(s, 5.08, 1.08,
          ["Teach to the approved course plan, using approved samples",
           "Seminars, coaching and handover materials; premium courses arranged separately"],
          ["Name the core trainees and schedule them for sessions and hands-on",
           "Organise staff by the six role learning paths — see Appendix A10"], size=10)
means(s, 6.24, "Core and specialist courses are in scope at no extra cost. {{Premium transfer requires all Huawei AI Developer courses first, plus a separate budget — see A11.}}", h=0.62)

# ============================== 11 Knowledge Transfer: four stages ==============================
s = slide("Knowledge transfer: two rounds, 1,216 hours, about 158 working days", size=19, ref="A8 · A12")
HOWTO["kt"] = s
thesis(s, "608 hours per round across four stages, delivered twice. After handover RNP can independently run  1) platform O&M  2) training submission  3) model deployment  4) release")
phases = [
    ("Foundation", "60", "per round · ~10 working days · 50 people",
     ["LLM basics and project architecture", "Corpus to service, end to end", "Safety and O&M basics"], XF),
    ("Shadowing", "360", "per round · ~45 working days · 20 people",
     ["iFLYTEK operates, RNP observes and assists", "Covers national model training", "Model upgrades within 12 months", "3 DLM trainings and other live work"], XF),
    ("Guided hands-on", "20", "per round · ~3 working days · 20 people",
     ["RNP operates: data, training jobs", "Deployment, monitoring, fault triage", "Both parties sign the hands-on record"], RN),
    ("Formal handover", "168", "per round · ~21 working days · 40 people",
     ["For RNP's O&M team", "Production environment and O&M process", "Asset licences and runbooks", "Independent operation verified and signed"], RN),
]
gap = 0.25
pw = (CW - 3 * gap) / 4
py, ph = 1.78, 3.16
for i, (name, hours, meta, tasks, fill) in enumerate(phases):
    px = L + i * (pw + gap)
    card(s, px, py, pw, ph, fill=fill)
    text(s, px + 0.2, py + 0.10, pw - 0.4, 0.34, f"{i + 1}  {name}", size=12, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, px + 0.2, py + 0.46, pw - 0.4, 0.46, [f"**{hours}** hours"], size=11, color=TXT, valign="m", margin=(0, 0, 0, 0), sa=0)
    r0 = s.shapes[-1].text_frame.paragraphs[0].runs[0]
    from pptx.util import Pt
    r0.font.size = Pt(21)
    r0.font.color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor.from_string(RED)
    text(s, px + 0.2, py + 0.94, pw - 0.4, 0.3, meta, size=9, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, px + 0.2, py + 1.26, pw - 0.4, ph - 1.32, tasks, size=9.5, ls=1.25, bullet=True, num=True, bcolor=RED, sa=2)
two_boxes(s, 5.06, 1.08,
          ["Coaching by the iFLYTEK engineers delivering this project, checking the key steps",
           "Handover materials and verification of independent operation"],
          ["Name stable trainees with the right background; the same people may span stages",
           "Allocate their time and join real delivery work"], size=10)
means(s, 6.26, "Card figures are per round; two rounds total 1,216 hours and ~158 working days. Working days are training days spread across the project, not a continuous period. Product operations training is separate — see A12.", h=0.6)

# ============================== 12 Implementation path ==============================
s = slide("Implementation: two stages, five steps, three parties", size=19)
thesis(s, "Five steps advance on preconditions:  1) Huawei provides compute  2) iFLYTEK provides software and models  3) RNP provides data and operations")
text(s, L, 1.58, CW, 0.26, "Completion markers shown under each node", size=9, color=MUTE, valign="m", align="r", margin=(0, 0, 0, 0), sa=0)
n = 5
x0, x1 = L + 1.15, R - 1.15
xs = [x0 + i * (x1 - x0) / (n - 1) for i in range(n)]
ly = 3.34
card(s, xs[0] - 1.0, 1.84, xs[2] - xs[0] + 1.4, 0.30, fill=PANEL2, line=LINE)
text(s, xs[0] - 0.85, 1.84, 6.4, 0.30, "Stage 1 · base deployment, model build, first app delivery", size=9.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
card(s, xs[2] - 0.4, 2.20, xs[4] - xs[2] + 1.5, 0.30, fill=WHITE, line=LINE)
text(s, xs[2] - 0.25, 2.20, 6.4, 0.30, "Stage 2 · extended operation once the second compute batch is ready", size=9.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
box(s, x0, ly - 0.02, x1 - x0, 0.05, fill=LINE)
nodes = [
    ("Platform and corpus", "1) Platform accepted  2) Corpus pilot batch passes"),
    ("Model training and evaluation", "1) Text model passes  2) Speech models pass"),
    ("Application and DLM", "1) App Stage 1 delivered  2) DLM forms signed"),
    ("Acceptance", "Component acceptance reports signed"),
    ("Handover and support", "1) Handover signed  2) Assisted O&M begins"),
]
for i, (name, done) in enumerate(nodes):
    cx = xs[i]
    dot(s, cx - 0.22, ly - 0.22, 0.44, str(i + 1), size=12)
    text(s, cx - 1.2, ly - 0.80, 2.4, 0.46, name, size=11, color=TXT, bold=True, align="c", valign="b", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    body(s, cx - 1.2, ly + 0.32, 2.4, 0.66, done, size=9, color=MUTE, align="c", ls=1.25)
box(s, xs[0], 4.56, xs[4] - xs[0], 0.05, fill=RED)
text(s, xs[0], 4.63, xs[4] - xs[0], 0.28, "Knowledge transfer runs across all four stages; formal handover starts at least 3 months before assisted O&M ends", size=9, color=RED, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
party = [("Huawei", TXT, GRAY_T, "Compute, facilities, network and base environment, coordinated by RNP"),
         ("iFLYTEK", RED, REDT, "Platform, models, applications, knowledge transfer and support"),
         ("RNP", RNP_C, BLUET, "Corpus and licences, DLM training, acceptance and daily operations")]
pw = (CW - 0.6) / 3
for i, (name, color, fill, items) in enumerate(party):
    px = L + i * (pw + 0.3)
    card(s, px, 5.04, pw, 0.76, fill=fill)
    label(s, px + 0.22, 5.04 + 0.21, 1.1, name, color=color, size=12)
    body(s, px + 1.25, 5.04 + 0.06, pw - 1.45, 0.64, items, size=9.5, valign="m", ls=1.25)
means(s, 5.94, "Month-by-month dates are agreed before implementation, with T0 at first compute readiness. RNP plans staffing from this — the three intakes are on the next page.")


# ============================== 13 RNP's investment ==============================
s = slide("RNP's investment: SOW responsibilities and three staffing intakes", size=19, ref="A7")
thesis(s, "Under the SOW, RNP is responsible for  1) producing and licensing the Portuguese corpus  2) training and evaluating the DLMs  3) approving users and operating the app  4) coordinating Huawei compute and environment, and organising acceptance", h=0.66)
waves = [
    ("Intake 1 · from step 1", "{{about 30}}", "1) Data team 2) Portuguese experts 3) Compliance review 4) Project coordination",
     "1) Collect and annotate 2) Quality-check the corpus 3) Secure licences"),
    ("Intake 2 · from step 2", "{{about 20}}", "1) Algorithm engineers 2) Test engineers 3) Business domain experts",
     "1) Join national model training and evaluation 2) Then train the DLMs"),
    ("Intake 3 · from step 3", "{{about 30}}", "1) Application development and operations 2) Platform O&M",
     "1) Take over app operations 2) User approval 3) Platform O&M"),
]
gap = 0.3
ww = (CW - 2 * gap) / 3
wy, wh = 1.96, 2.45
for i, (name, num, who, task) in enumerate(waves):
    wx = L + i * (ww + gap)
    card(s, wx, wy, ww, wh, fill=BLUET)
    text(s, wx + 0.25, wy + 0.14, ww - 0.5, 0.34, name, size=12, color=RNP_C, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, wx + 0.25, wy + 0.52, ww - 0.5, 0.5, num + " people", size=17, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, wx + 0.25, wy + 1.08, ww - 0.5, 0.66, "**Roles:** " + who, size=10, ls=1.3)
    body(s, wx + 0.25, wy + 1.80, ww - 0.5, 0.56, "**Tasks:** " + task, size=10, ls=1.3)
    if i < 2:
        arrow(s, wx + ww + 0.04, wy + 0.95, fill=LINE)
text(s, L, 4.54, CW, 0.5, "Headcounts are indicative, from iFLYTEK's early estimate. They are peaks per function that can be reused across activities, not people on site at the same time; final numbers follow the SOW split and workload.", size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
means(s, 5.08, "Data team and Portuguese experts come first, then algorithm engineers, with application and operations in the third intake. Knowledge-transfer trainees are named from these three intakes: 50 Foundation, 20 each Shadowing and Guided hands-on, 40 Handover, per round, twice — 260 attendances.", h=0.86)

# ============================== 14 TR five layers ==============================
s = slide("Against RNP's TR: who delivers each of the five layers", size=19, ref="A14")
thesis(s, "The TR groups the national AI software stack into five sets of requirements. The table below shows how the contractor side (Huawei + iFLYTEK) maps to them.", h=0.46)
cols5 = [(L, 2.30), (2.90, 5.10), (8.20, 1.60), (9.90, 2.93)]
for (cx, cwd), nm, col in zip(cols5, ["Layer", "Requirement", "Who delivers", "Where we cover it"], [MUTE, MUTE, RED, RNP_C]):
    text(s, cx + 0.16, 1.64, min(cwd, 2.4), 0.28, nm, size=10.5, color=col, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
rows5 = [
    ("1  Base compute software", "OS, firmware and drivers; containers and runtimes; operators, math libraries and compilers; training and inference frameworks; distributed communication and orchestration",
     "Huawei", "Huawei's base software stack; the iFLYTEK platform runs on it"),
    ("2  Data, model and application platform", "Corpus, training and application innovation centers; national model; national AI application; model services and APIs",
     "iFLYTEK", "Pages 3–9; Appendices A1–A5"),
    ("3  Integration and operations", "RNP identity, security and monitoring; third-party model and tool migration; benchmarking; multi-tenant isolation; fault recovery",
     "All three", "Appendices A1, A4, A6; {{interface and migration scope confirmed item by item}}"),
    ("4  R&D and talent", "R&D cooperation plan; interfaces, tools and technical materials; R&D support; practical training and knowledge transfer; retained results",
     "All three", "Pages 10–11; Appendices A8–A11, A14"),
    ("5  Delivery and long-term autonomy", "Asset and licence list; five-year maintenance and technology refresh; end-of-term handover; exit and migration clauses",
     "All three", "Appendix A6; {{refresh and end-of-term handover clauses to align}}"),
]
for i3, (lay, req, who, ours) in enumerate(rows5):
    ry = 1.98 + i3 * 0.92
    card(s, cols5[0][0], ry, cols5[0][1], 0.84, fill=PANEL2, line=LINE)
    text(s, cols5[0][0] + 0.16, ry, cols5[0][1] - 0.3, 0.84, lay, size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    card(s, cols5[1][0], ry, cols5[1][1], 0.84, fill=WHITE, line=LINE)
    body(s, cols5[1][0] + 0.16, ry + 0.06, cols5[1][1] - 0.32, 0.72, req, size=8.5, valign="m", ls=1.25)
    card(s, cols5[2][0], ry, cols5[2][1], 0.84, fill=GRAY_T if who == "Huawei" else (REDT if who == "iFLYTEK" else PANEL))
    text(s, cols5[2][0], ry, cols5[2][1], 0.84, who, size=10.5, color=TXT, bold=True, align="c", valign="m", margin=(0, 0, 0, 0), sa=0)
    card(s, cols5[3][0], ry, cols5[3][1], 0.84, fill=RN)
    body(s, cols5[3][0] + 0.16, ry + 0.06, cols5[3][1] - 0.32, 0.72, ours, size=8.5, valign="m", ls=1.25)

# ============================== 15 Q&A index ==============================
s = slide("Appendix: 15 questions, in RNP's order", size=20)
INDEX["index"] = s
thesis(s, "We start here once the main deck is done. Each question has one appendix page — click to open, and use the bar at the bottom of that page to come back.")
text(s, L, 1.76, CW, 0.28, "Unless an item is marked \"Approach\", appendix content is written into the scope of work (SOW). {{Blue}} = to be confirmed jointly before implementation. The bar at the bottom of every appendix page jumps between all 15 pages.", size=9, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
menu = [
    ("A0", "Can we see it all on one page?", "Overview: four outcomes, build, roles, key figures"),
    ("A1", "How large is the platform, how is it accepted?", "Capability Center: capacity and acceptance"),
    ("A2", "How big is the model, how is it upgraded?", "National Model: specification and upgrades"),
    ("A3", "How much corpus, from whom, held where?", "Corpus: volume, quality and data control"),
    ("A4", "What can the app do, where are the limits?", "National AI Application: functions and limits"),
    ("A5", "Who builds the DLMs, and how far?", "DLM: scope, inputs and service acceptance"),
    ("A6", "Will we be locked in, what costs follow?", "Handover, licensing and operational support"),
    ("A7", "How many people of each type, and when?", "RNP staffing detail"),
    ("A8", "What happens in each stage, how many days?", "Knowledge transfer: four stages in detail"),
    ("A9", "What exactly is taught, in what formats?", "Knowledge transfer: topics and formats"),
    ("A10", "What does each role learn?", "Knowledge transfer: learning path by role"),
    ("A11", "Can we learn how this system was built?", "Premium knowledge transfer: optional menu"),
    ("A12", "How does O&M get started after delivery?", "Product operations training: topics, days, people"),
    ("A13", "What is delivered for the app, how is it accepted?", "National AI Application: deliverables and acceptance"),
    ("A14", "How do you answer the TR's R&D requirements?", "Supporting Brazil's own R&D: requirement mapping"),
]
MENU = []
mw = (CW - 2 * 0.3) / 3
mh = 0.90
for k, (code, q, topic) in enumerate(menu):
    r, c = divmod(k, 3)
    mx, my = L + c * (mw + 0.3), 2.04 + r * (mh + 0.06)
    cardk = card(s, mx, my, mw, mh, fill=WHITE, line=LINE)
    MENU.append((cardk, code))
    text(s, mx + 0.2, my + 0.08, 0.75, 0.32, code, size=11.5, color=RED, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    text(s, mx + 0.88, my + 0.04, mw - 1.05, 0.50, q, size=10.5, color=TXT, bold=True, valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.1)
    text(s, mx + 0.2, my + 0.56, mw - 0.4, 0.3, topic, size=9, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)


# ============================== A0 Overview table ==============================
s = appendix_slide("A0 · Overview: four outcomes and key figures", "Q: Can we see it all on one page?")
rows = [
    ["National AI Capability\nCenter (platform)", "Corpus construction, model training and application innovation centers plus a content safety review platform; all three outcomes are built on it and RNP runs it afterwards",
     "Software implementation, testing and training", "Base environment, interface coordination, daily operations after acceptance", "3 centers\n1 review platform"],
    ["Brazilian Portuguese\nNational Model", "One NLP base model, one ASR model, one TTS model (one male and one female voice); weights, model card, evaluation report and deployment guide delivered",
     "Training, technical evaluation and delivery; 12 months of tuning afterwards", "Corpus production and licensing, expert review, acceptance", "NLP ≥290B parameters\n50B-token local corpus"],
    ["National AI Application", "Multi-device National App and Data Operations Platform for users approved by RNP; Q&A, search, documents, speech and writing",
     "Application delivery, interface testing, platform configuration", "Content, branding, interfaces; user approval and daily operations", "NLP 1,000 sessions\nASR / TTS 200 each\nSafety review 400"],
    ["Domain Large Models", "Three advisory training services, one fixed scenario each; RNP trains the models and owns the weights",
     "Plan, process advice, results analysis", "Data, environment, training and evaluation, business acceptance", "3 services\n1 scenario each"],
]
table(s, L, 1.02, [1.95, 3.75, 2.35, 2.5, 1.78], ["Outcome", "What RNP gets", "iFLYTEK", "RNP", "Key figures"], rows,
      row_h=0.90, header_h=0.42, size=9, header_size=10, header_fill="F2F2F2", header_color=TXT, vlines=False)
band(s, 5.02, 0.58, "Knowledge transfer",
     "608 hours and about 79 working days per round: Foundation 60, Shadowing 360, Guided hands-on 20, Formal handover 168; 50 / 20 / 20 / 40 trainees per round. Two rounds: 1,216 hours, about 158 working days.",
     fill=REDT, line=None, size=9.5)
band(s, 5.68, 0.50, "Post-delivery support",
     "12 months assisted O&M on site (5×8) plus 60 months remote technical support (7×24); the first 12 months overlap.", size=9.5)
note(s, 6.22, "Scope follows the SOW and BOQ; capacity is verified per service.", h=0.3, size=9)

# ============================== A1 Capability Center capacity ==============================
s = appendix_slide("A1 · National AI Capability Center: capacity and acceptance", "Q: How large is the platform, how is it accepted?")
spec_table(s, [
    ["Software components", "Corpus construction, model training and application innovation centers plus the content safety review platform; with tenant, resource, queue and monitoring management", SOW],
    ["Managed accelerator licence", "The platform is licensed to manage up to 1,024 accelerator cards; hardware quantity follows Huawei's resource plan, expansion per the BOQ", SOW],
    ["Inference concurrency", "NLP 1,000, ASR 200, TTS 200, safety review 400 sessions; configured and tested per service — these are not added into one total", SOW],
    ["Capacity test conditions", "Hardware, model version, input/output length, audio duration, request distribution, test tooling and pass rules agreed before implementation", SOW],
    ["Network security products", "2 bastion hosts, 1 log audit, 1 database audit, 1 vulnerability scanner, 2 WAFs, with licences", SOW],
    ["Acceptance method", "Check BOQ and licences; verify deployment, function, capacity, security and licensing; submit test records and clear blocking issues", SOW],
])
note(s, 5.40, "Concurrency is a licensed capacity baseline configured and tested per service, not a simultaneous peak. 1,024 is the platform's management licence ceiling, not this phase's hardware count.", h=0.4, size=9)

# ============================== A2 Model specification ==============================
s = appendix_slide("A2 · Brazilian Portuguese National Model: specification and upgrades", "Q: How big is the model, how is it upgraded?")
spec_table(s, [
    ["Model size", "One NLP model, ≥290B total parameters; one ASR model, hundreds of millions of parameters; one TTS model with one male and one female voice", SOW],
    ["Base model and rights", "Actual base model weights and the artefacts needed to run locally; base and post-training weights identified separately, access per the asset and licence matrix", SOW],
    ["Base Portuguese enhancement (Approach)", "In parallel with local corpus production, iFLYTEK uses its own compute in China to collect public Portuguese corpus and continue training the base model, as the starting point for the national model — no Brazilian compute used", PLAN],
    ["Base model name", "{{TBC: whether the SOW names iFLYTEK Spark as the base; if not, it is referred to as the iFLYTEK base model}}", TBD],
    ["Build-phase training", "Executed per the approved model plan; technical route, number of rounds, first delivery date and milestones confirmed in writing before implementation", SOW],
    ["Continuous tuning period", "12 consecutive months after initial acceptance of NLP, ASR, TTS and the safety platform; formal upgrades: NLP 3, ASR 1, TTS 1, safety platform 1", SOW],
    ["Tuning data caps", "Across 12 months: NLP text ≤10B tokens, supervised ≤20,000 samples, safety alignment ≤5,000, ASR ≤1,000 hours, TTS ≤10 hours; separate caps for safety samples, lexicons and hot fixes", SOW],
    ["Export and portability", "Per verified model, format, framework and target environment; subject to model architecture, operator and licence constraints", SOW],
], row_h=0.55)
note(s, 6.02, "Performance upgrades are limited to the agreed number within 12 months, with at most 3 main objectives per NLP upgrade. Build-phase versions and defect fixes are not counted.", h=0.4, size=9)

# ============================== A3 Corpus ==============================
s = appendix_slide("A3 · Brazilian Portuguese corpus: volume, quality and data control", "Q: How much corpus, from whom, held where?")
spec_table(s, [
    ["Unsupervised text", "50B tokens of valid data counted with the approved tokenizer, de-duplicated and screened", SOW],
    ["Supervised and safety data", "200,000 supervised tuning samples; 20,000 safety and values samples, counted as complete valid samples", SOW],
    ["Speech data", "ASR: 10,000 hours of valid audio with qualified transcription; TTS: 100,000 items and at least 100 hours of high-quality recording, with lexicon and samples", SOW],
    ["Split of work and early compute", "RNP organises data acquisition, licensing and full-volume production; iFLYTEK provides the platform, specifications, pilot-batch guidance and sampling quality checks. {{Compute for corpus collection and processing before the data center opens is arranged locally by RNP.}}", SOW],
    ["Quality rules", "Counting, sampling, quality thresholds, batching and re-check rules agreed before volume production; data that fails is not counted", SOW],
    ["Data control", "Project data is stored in the Brazilian environment designated by RNP, managed by access rights and audit; data rights and permitted use follow the original licences", SOW],
    ["iFLYTEK public corpus (Approach)", "iFLYTEK uses its own compute in China to collect and process public Portuguese corpus to strengthen the base model's general Portuguese, and hands that corpus to Brazil as well. {{Volume and handover method confirmed before implementation.}}", PLAN],
], header=("Item", "Content (volumes are minimum references)"), row_h=0.56)
note(s, 5.60, "Corpus acceptance is measured by the definition of valid data and the approved plan; model performance is verified separately under the A2 evaluation rules.", h=0.4, size=9)


# ============================== A4 Application functions and limits ==============================
s = appendix_slide("A4 · National AI Application: functions and limits", "Q: What can the app do, where are the limits?")
rows = [
    ["Chat and Q&A", "Single entry point: text input, answers, regeneration, feedback, history; general Q&A and Q&A over connected knowledge bases with source citation; multi-turn follow-up",
     "No guarantee of an answer to every question or of correctness; custom industry chatbots excluded; context limited by token limits and session policy"],
    ["Search", "Search across connected knowledge bases, applications and authorised documents; web results shown with links to the source",
     "Search of unconnected systems, all local files or arbitrary web content excluded; web search depends on an external service"],
    ["Documents and translation", "Attachment parsing, document Q&A, summaries, outlines and key points; **translation assistance for parsed document text**",
     "Limited to supported, successfully parsed formats; very long whole-document translation, complex layout retention, full scan recognition and professional review not committed; **standalone speech translation excluded**"],
    ["Writing and presentations", "AI writing, AI notes and AI presentation creation, with assisted generation, editing, templates and export",
     "{{Portuguese slide generation and official templates included only after adaptation is verified}}; professional design services and lossless format conversion excluded"],
    ["Three usage directions", "Learning: Q&A and document summaries. Work: notes, writing, documents, app shortcuts. Daily life: everyday queries with web search",
     "Teaching systems, curricula and exam scoring excluded; OA approvals and process automation excluded; bookings, payments and lifestyle services excluded"],
    ["Portuguese speech, multi-device", "This project's ASR and TTS for voice input and read-aloud; web client plus the clients in this project's delivery scope, verified on specified browsers and OS versions",
     "Audio spec, duration, concurrency and voices limited to the delivered speech service; multi-device is not every device or OS — final scope per the BOQ"],
    ["Authentication and users", "Platform account login, plus external authentication through product-verified methods; organisation, user, tenant, role and menu permissions",
     "**iFLYTEK QR-code login is not required to run locally in Brazil**; MFA and arbitrary identity systems not included by default; directory sync needs verified interfaces"],
    ["Content review and security", "Connected to this project's content safety review platform, applying policy at the agreed text interaction points; authentication, role authorisation, tenant access control, logging",
     "**Spark Cloud does not replace the standalone review platform**; image, audio and video review not included by default; not all risks guaranteed to be caught; not a full security framework or SOC"],
    ["Logs and operations reports", "Query and export of the operation, client and service logs the product records; visits, active users, usage trends, favourites and feedback statistics",
     "Not every business activity is logged; log fields, retention and capacity set at deployment; unconnected-system statistics, custom BI and billing excluded"],
]
table(s, L, 0.98, [2.25, 5.40, 4.68], ["Function", "Included", "Excluded or limited"], rows,
      row_h=0.55, header_h=0.38, size=8, header_size=9.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)

# ============================== A5 DLM scope ==============================
s = appendix_slide("A5 · Domain Large Models: scope, inputs and service acceptance", "Q: Who builds the DLMs, and how far?")
spec_table(s, [
    ["Number and scope", "Three DLM advisory training services; one clearly bounded fixed scenario each. RNP proposes the domain, iFLYTEK assesses feasibility", SOW],
    ["Base model", "An existing base model agreed by both parties; version, size, licence and toolchain verified first", SOW],
    ["Data requirements", "RNP provides lawful, sufficient and representative training, validation, safety and independent evaluation data; volumes per the specific plan", SOW],
    ["Training route", "One main route confirmed per service — LoRA-SFT, SFT, GRPO or similar, selected on feasibility", SOW],
    ["Delivery model", "RNP performs training and evaluation; iFLYTEK provides guidance, analysis and optimisation advice within the agreed effort", SOW],
    ["Rights and boundaries", "Weights are produced by RNP's training and belong to RNP; iFLYTEK delivers the plan, advisory records and analysis report. Production deployment, API development and application integration can be agreed separately", SOW],
    ["Start confirmation", "A DLM Scenario and Technical Plan Confirmation is signed for each service, setting duration, effort, deliverable form and pass rules", SOW],
], row_h=0.56)
note(s, 5.60, "Service acceptance checks the plan, the advisory activity and the analysis based on RNP's actual results. Model performance targets are separate from iFLYTEK's service acceptance.", h=0.4, size=9)

# ============================== A6 Handover, licensing, support ==============================
s = appendix_slide("A6 · Handover, licensing and operational support", "Q: Will we be locked in, what costs follow?")
spec_table(s, [
    ["Model rights", "Base weights, post-training weights and DLM outputs are listed separately for delivery and licensing; weight delivery does not include background intellectual property", SOW],
    ["Platform and custom code", "Platform software used per the BOQ and its licences; for custom development, source code and modification rights are delivered per the list", SOW],
    ["Perpetual use", "Versions marked perpetual remain valid after payment and do not expire when the service period ends; maintenance, support and upgrades follow their own service terms", SOW],
    ["Export and third-party adaptation", "Export and migration per licence and technical feasibility, with format, framework and target environment confirmed item by item; the BOQ lists 2 third-party model inference adaptation services", SOW],
    ["On-site and remote support", "12 months assisted O&M on site (5×8); 60 months remote technical support (7×24, L3/L4); the first 12 months overlap", SOW],
    ["First response", "P1 ≤30 minutes; P2 ≤60 minutes; P3 ≤2 hours; P4 ≤next working day", SOW],
    ["Later costs and documents", "Listed support follows the BOQ; expansion, renewal, third-party licences and out-of-scope work are confirmed separately on stated terms; documentation in English or Brazilian Portuguese", SOW],
], row_h=0.56)
note(s, 5.60, "These are first-response times, not restoration times. The last four years of remote support do not include permanent on-site staff.", h=0.4, size=9)

# ============================== A7 RNP staffing ==============================
s = appendix_slide("A7 · RNP staffing detail: three intakes", "Q: How many people of each type, and when?")
rows = [
    ["Project and product", "Project managers 2, product managers 4, PMO / data center experts 3", "Requirements, schedule, resources, delivery", "Intake 1 · from step 1", "{{9}}"],
    ["Corpus and data resources", "Resource engineers 6, data QC 6, annotators 20", "Organise data, annotation and quality checks", "Intake 1 · from step 1", "{{32}}"],
    ["Portuguese language experts", "Language experts 6", "Portuguese quality, annotation and speech standards", "Intake 1 · from step 1", "{{6}}"],
    ["Compliance and review", "Review experts 2, legal experts 2, review leads 2", "Content, safety and legal compliance review", "Intake 1 · from step 1", "{{6}}"],
    ["Business and domain experts", "Business requirements 3, domain requirements 2, domain experts 2", "Confirm business scenarios and domain results", "Intake 2 · from step 2", "{{7}}"],
    ["Model and test engineering", "Algorithm engineers 12, test engineers 6", "Model training, evaluation and technical validation", "Intake 2 · from step 2", "{{18}}"],
    ["Application and operations", "Software developers 15, operations engineers 10, architects 2", "Application architecture, integration and operations", "Intake 3 · from step 3", "{{27}}"],
    ["Platform O&M", "O&M engineers 6", "Run management, deployment and escalation", "Intake 3 · from step 3", "{{6}}"],
]
table(s, L, 1.15, [2.25, 3.7, 3.0, 1.85, 1.53], ["Role group", "Composition (indicative)", "Main responsibility", "Intake", "Indicative"], rows,
      row_h=0.46, header_h=0.4, size=9.5, header_size=10, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.40, ["Intake subtotals: intake 1 {{about 53}}, intake 2 {{about 25}}, intake 3 {{about 33}}; the total of {{111}} is the sum of peaks per function, reusable across activities, not people on site at once.",
               "These indicative numbers come from iFLYTEK's early, non-binding estimate for a full-delivery model. Under this SOW RNP takes on full data production, DLM training and daily operations, so headcount should be re-checked against the actual split."], h=0.7, size=9)


# ============================== A8 Knowledge transfer: four stages ==============================
s = appendix_slide("A8 · Knowledge transfer: the four stages in detail", "Q: What happens in each stage, how many days?")
rows = [
    ["Foundation\n(AI basics)", "A shared understanding of machine learning, deep learning, NLP and large models; covers corpus preparation, training, evaluation, model services, application innovation, Capability Center architecture and O&M basics",
     "iFLYTEK teaches", "60 hours\n~10 working days", "50", "Stage requirements met\n(target ≥80%)"],
    ["Shadowing\n(iFLYTEK operates)", "Observe and assist iFLYTEK engineers on the delivered platform: dataset management, training jobs, resource configuration, evaluation, release, serving and basic O&M; content follows live delivery and operations work",
     "iFLYTEK operates\nRNP observes, assists", "360 hours\n~45 working days", "20", "Hands-on assessment passed\n(target ≥70%)"],
    ["Guided hands-on\n(RNP operates)", "RNP performs the pre-agreed task list: platform operation, data processing, training job submission, model deployment, monitoring, alarms and fault triage; iFLYTEK guides and checks the key steps",
     "RNP operates\niFLYTEK guides, checks", "20 hours\n~3 working days", "20", "Hands-on record signed by both\n(target ≥50% met)"],
    ["Formal handover\n(operational ownership)", "Handover coaching on the platform, licensed assets, configuration, access rights and documents actually delivered under the SOW, BOQ and final acceptance list; RNP's independent operation is verified; open items and support arrangements confirmed",
     "RNP operates alone\niFLYTEK verifies", "168 hours\n~21 working days", "40", "Handover confirmation signed by both"],
    ["Total", "Hours are counted once across topics; two rounds are delivered with trainees taking part in batches; stage dates follow the approved knowledge transfer plan", "",
     "Per round 608 h / ~79 days\nTwo rounds 1,216 h / ~158 days", "Per round 50/20/20/40\nTwo rounds: 260 attendances", ""],
]
table(s, L, 1.15, [1.75, 4.05, 1.6, 1.7, 1.2, 2.03], ["Stage", "Learning objectives and scope", "Roles", "Hours / days per round", "Trainees", "Completion evidence"], rows,
      row_h=[0.75, 0.75, 0.75, 0.75, 0.7], header_h=0.4, size=8.5, header_size=9.5, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.55, ["Hours and headcount in the table are per round and delivered twice. Working days are the sum of training days, not a continuous period. Percentages are development targets for qualified core trainees, not acceptance thresholds. Formal handover starts at least 3 months before the year of assisted O&M ends — the 168 hours do not begin at that point.",
               "Trainees take part in batches and need not be the same people across stages. 260 is the total attendances across two rounds — one person joining several stages is counted each time, so it is not 260 distinct trainees. Formal handover targets the O&M team, so its headcount is wider. Basis: SOW and the approved knowledge transfer plan."], h=0.88, size=8.5)

# ============================== A9 Topics and formats ==============================
s = appendix_slide("A9 · Knowledge transfer: topics and exchange formats", "Q: What exactly is taught, in what formats?")
label(s, L, 1.02, 9, "Four tracks · a grouping of the 12 SOW course topics, adding no hours", color=TXT, size=11.5)
tracks = [
    ("Models and data evaluation", ["Model architecture and training routes: continued pre-training, SFT, LoRA, preference alignment", "Portuguese data: source, quality, lineage, privacy and lifecycle governance", "Evaluation: regression, robustness, bias, safety, hallucination, cultural fit"], "Goal: compare training routes, check data quality, evaluate against agreed baselines"),
    ("Software stack and lifecycle", ["Software layers and dependencies; version and compatibility management", "MLOps / LLMOps: experiment tracking, release, monitoring and rollback", "Hardware-software-algorithm fit: load and resource bottlenecks"], "Goal: understand dependencies, lifecycle operations and what the platform can do"),
    ("Training and inference at scale", ["Resource allocation and checkpoint recovery; data, tensor, pipeline and sequence parallelism", "Model serving, batching, KV cache; throughput, latency and capacity", "Quantisation, pruning, distillation; compression and energy analysis"], "Goal: find bottlenecks and trade off quality, latency, throughput and cost"),
    ("Retrieval, agents and safety", ["RAG and graph retrieval; hybrid retrieval and re-ranking", "Agents: tools, planning, memory and multi-agent collaboration", "System and model safety: prompts, content review, supply chain"], "Goal: design knowledge services and research assistants with source checking and risk control"),
]
gap = 0.25
tw = (CW - 3 * gap) / 4
ty, th = 1.38, 2.62
for i, (name, items, obj) in enumerate(tracks):
    tx = L + i * (tw + gap)
    card(s, tx, ty, tw, th, fill=PANEL)
    dot(s, tx + 0.16, ty + 0.14, 0.32, str(i + 1), size=10)
    text(s, tx + 0.56, ty + 0.08, tw - 0.7, 0.44, name, size=10.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.1)
    body(s, tx + 0.16, ty + 0.56, tw - 0.32, 1.55, items, size=8.5, bullet=True, bcolor=RED, sa=2, ls=1.25)
    body(s, tx + 0.16, ty + 2.14, tw - 0.32, 0.44, obj, size=8.5, color=MUTE, ls=1.2)
label(s, L, 4.10, 6, "Four exchange formats", color=TXT, size=11.5)
formats = [
    ("Advanced AI seminars", "In scope", GRAY, "Per the approved course plan, on the delivered platform and models; content and depth follow that plan"),
    ("China AI study visit", "{{See A11}}", RNP_C, "Visit iFLYTEK teams and compute facilities. Huawei's plan already includes a 6-day China journey for 100 people — this is the iFLYTEK technical variant, arranged so the two do not overlap"),
    ("One-to-one expert consulting", "{{See A11}}", RNP_C, "Design review on questions raised by RNP engineers; scope and duration agreed in advance, no long-term advisory role"),
    ("AI strategy round table", "{{Proposed}}", RNP_C, "RNP leadership and iFLYTEK architects discuss model roadmap, technology choices and capability planning; participants and timing agreed in advance"),
]
fy, fh = 4.44, 1.36
for i, (name, status, color, desc) in enumerate(formats):
    fx = L + i * (tw + gap)
    card(s, fx, fy, tw, fh, fill=PANEL2 if i == 0 else BLUET)
    text(s, fx + 0.16, fy + 0.06, tw - 1.25, 0.42, name, size=9.5, color=TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.1)
    chip(s, fx + tw - 1.05, fy + 0.11, 0.90, 0.28, status, fill=WHITE, color=TXT, size=8, line=color)
    body(s, fx + 0.16, fy + 0.52, tw - 0.32, 0.84, desc, size=8.5, ls=1.25)
note(s, 5.92, "Track-to-topic mapping: track 1 = SOW topics 1, 7, 8; track 2 = 2, 10, 11; track 3 = 3, 4, 9; track 4 = 5, 6, 12. Teaching a method does not mean every delivered model supports it, and no new optimisation work is implied. Basis: SOW. Of the four formats, only the advanced seminars are in project scope; the other three are agreed separately.", h=0.5, size=8.5)

# ============================== A10 Learning path by role ==============================
s = appendix_slide("A10 · Knowledge transfer: learning path by role", "Q: What does each role learn?")
rows = [
    ["Corpus collection team", "Foundation, Shadowing", "Data sources, licensing, standards and platform processing flow", "Materials, tool guides, representative samples and processing guidance", "Course, shadowing and issue records"],
    ["Language and annotation experts", "Foundation, Shadowing", "Annotation, language quality, safety labels and speech standards", "Standards walkthrough, sample analysis, quality check methods", "Agreed exercises and rule confirmation"],
    ["Model engineers", "Shadowing, Guided hands-on", "Submit training jobs, read logs and evaluations, understand model release", "Job configuration, log analysis, evaluation and release guidance", "Agreed hands-on results and advisory records"],
    ["Application staff", "Shadowing, Guided hands-on", "App configuration, user permissions, API calls and usage analytics", "Guidance on the application platform, APIs and operations features", "Agreed feature and API walkthroughs"],
    ["O&M engineers", "Guided hands-on, Handover", "Monitoring, incidents, changes, common faults and escalation", "Runbooks, representative operations and independent-verification coaching", "Independent operation and handover confirmation"],
    ["Domain experts", "Shadowing, Guided hands-on", "Scenarios, data and how model results are assessed", "Evaluation methods, report reading and content boundaries", "Participation records and agreed evaluation exercises"],
]
table(s, L, 1.15, [2.1, 1.85, 3.3, 3.15, 1.93], ["Role", "Stages", "What they take part in", "How iFLYTEK supports", "Completion evidence"], rows,
      row_h=0.60, header_h=0.4, size=9, header_size=10, header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.30, "{{TBC: how the four stages map to the five implementation steps.}} Role allocation and exercises are set by the approved knowledge transfer plan; dedicated DLM advisory work and implementation training are not double-counted.", h=0.4, size=9)


# ============================== A11 Premium knowledge transfer ==============================
s = appendix_slide("A11 · Premium knowledge transfer: optional menu", "Q: Can we learn how this system was built?")
text(s, L, 0.98, CW, 0.56, "Baseline support inside the project is open documentation and answers to technical questions. Below is optional, proactive premium transfer for RNP's core R&D staff and university groups, from architecture and principles through to engineering experience and source-code practice. **Entry**: all Huawei AI Developer courses completed first.   **Budget**: outside this project's scope and price — separate budget required",
     size=9.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
label(s, L, 1.62, 7, "Four directions · about 5 hours each", color=TXT, size=11.5)
menu = [
    ("Scheduling and managing clusters above 1,000 PFLOPS", "Cluster build and management; core scheduling techniques and practice"),
    ("Long, stable training of 100B+ parameter models", "Stability assurance and optimisation experience at scale"),
    ("Building and processing TB-scale corpora", "Collection, cleaning and processing methods and pipelines at scale"),
    ("Inference serving at a trillion tokens a day", "Token gateway and routing; low-latency optimisation and stability"),
]
gap = 0.25
mw = (CW - 3 * gap) / 4
my, mh = 1.98, 1.90
for i, (name, desc) in enumerate(menu):
    mx = L + i * (mw + gap)
    card(s, mx, my, mw, mh, fill=PANEL)
    text(s, mx + 0.2, my + 0.12, mw - 0.4, 0.78, f"{i + 1}  {name}", size=10.5, color=TXT, bold=True, valign="t", margin=(0, 0, 0, 0), sa=0, ls=1.15)
    body(s, mx + 0.2, my + 0.96, mw - 0.4, 0.60, desc, size=9, color=BODY, ls=1.25)
    text(s, mx + 0.2, my + mh - 0.32, mw - 0.4, 0.26, "{{Proposed · priced separately}}", size=8.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0)
label(s, L, 4.06, 7, "Three supporting formats", color=TXT, size=11.5)
forms = [
    ("Expert seminars and case studies", "Seminars, engineering case reviews and practice walkthroughs on selected directions, tailored to RNP's actual environment — for example model migration, platform architecture and implementation principles, training adaptation on Huawei compute"),
    ("One-to-one expert consulting", "Technical Q&A, solution discussion and engineering experience exchange on specific R&D problems raised by RNP engineers, agreed by duration"),
    ("China study visit and technical exchange", "Optional format combining expert learning with on-site technical exchange; language, itinerary and related costs agreed separately. {{Distinct from the 6-day China journey for 100 people in Huawei's plan, so the two do not overlap.}}"),
]
fw = (CW - 2 * 0.3) / 3
fy, fh = 4.42, 1.42
for i, (name, desc) in enumerate(forms):
    fx = L + i * (fw + 0.3)
    card(s, fx, fy, fw, fh, fill=BLUET if i == 2 else PANEL)
    text(s, fx + 0.2, fy + 0.10, fw - 0.4, 0.32, name, size=11, color=RNP_C if i == 2 else TXT, bold=True, valign="m", margin=(0, 0, 0, 0), sa=0)
    body(s, fx + 0.2, fy + 0.46, fw - 0.4, fh - 0.54, desc, size=8.5, color=BODY, ls=1.25)
text(s, L, 6.02, CW, 0.45, "All of the above is separate from the core knowledge transfer (four stages, 1,216 hours across two rounds) and outside the current SOW and price. Scope, hours, headcount, format and fees are agreed separately. Basis: proposal and suggestion, not a commitment.",
     size=8.5, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)

# ============================== A12 Product operations training ==============================
s = appendix_slide("A12 · Product operations training: getting O&M started", "Q: How does O&M get started after delivery?")
text(s, L, 0.98, CW, 0.48, "iFLYTEK demonstrates operation, gives configuration guidance and explains handover for the delivered software, models and security products. For RNP's O&M team, starting with assisted O&M — parallel to the four-stage knowledge transfer, with no double-counted hours.",
     size=10, color=MUTE, valign="m", margin=(0, 0, 0, 0), sa=0, ls=1.25)
rows = [
    ["Corpus and model training platforms", "Dataset workflows, quality checks and lineage; container images, training jobs, logs and model outputs; delivered speech tools and checkpoint recovery", "2", "12"],
    ["Compute management and cluster O&M", "Tenants, roles, resource pools, quotas and scheduling; cluster health, node status, logs and alarms; resource usage, incidents and escalation", "1", "12"],
    ["Application services, APIs and data operations", "Service credentials, permissions, rate limits and usage; enabled knowledge Q&A and search; API calls and error handling; portal setup, app publishing, users and feedback", "1", "12"],
    ["Foundation models and the National AI Application", "Model files, loading, calls and configuration; known limits and operational checks; login, Q&A, documents and voice; enabled notes, writing and presentation features", "2", "12"],
    ["Content safety and security products", "Risk labels, review rules, lexicons and quick fixes; review logs and verification of enabled functions; access, policy and logs for the supplied security products; basic operation and vendor support channels", "1", "12"],
    ["Configuration, documents and support handover", "Deployed versions, key settings and documentation; supported backup and recovery; open issues, support tickets and escalation", "2", "12"],
]
table(s, L, 1.42, [3.15, 6.93, 0.95, 1.30], ["Topic", "Training content", "Days", "Per batch"], rows,
      row_h=0.58, header_h=0.4, size=9, header_size=10,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
band(s, 5.42, 0.58, "Timing", "Product operations training runs when assisted O&M begins; knowledge transfer activity continues through the following 12 months of assisted O&M. Formal handover starts at least 3 months before assisted O&M ends.", size=9.5, title_w=1.0)
note(s, 6.04, ["The six topics are a consolidated view of the twelve product-use themes in SOW section 7.5; they are not six separately quantified courses. Training uses delivered versions and authorised samples.",
               "{{Batches, days, participants, language and on-site/remote arrangements are agreed before implementation. The topic days total 9; the joint material states about 10 — to be confirmed.}}   Basis: SOW 7.5 / 8.3 / 9.4."],
     h=0.5, size=8.5)

# ============================== A13 Application deliverables and acceptance ==============================
s = appendix_slide("A13 · National AI Application: deliverables and acceptance", "Q: What is delivered, how is it accepted?")
rows = [
    ["Scope and localisation confirmation", "National AI Application function and configuration guide: actual functions, supported platforms, interface localisation, connected services and usage limits",
     "Specified Portuguese interfaces and branding complete; key operating information understandable and usable"],
    ["Multi-device software delivery", "Deployment packages or client installers for the confirmed platforms, licensing materials and release notes; application deployment records",
     "Software for the confirmed platforms is deployed or installed and can be accessed and operated"],
    ["Data Operations Platform delivery", "Platform software and licensing materials; Data Operations Platform configuration and user guide",
     "App configuration, user permissions, logging, feedback and statistics verified against test cases; displayed results match the test data"],
    ["Interface and workflow integration testing", "Application interface and integration test records covering the models, identity and knowledge services actually connected",
     "Specified model and external interfaces pass integration testing; authentication, permissions and connected review workflows verified"],
    ["Testing and acceptance", "National AI Application and Data Operations Platform test report; go-live release and acceptance confirmation materials",
     "Functions on the explicit list pass the approved test cases; representative workflows run from user input to displayed result"],
    ["User and administrator guidance", "User guide; administrator guide",
     "Operating materials, test results, versions and release records submitted"],
]
table(s, L, 1.02, [2.75, 4.85, 4.73], ["Service scope", "Deliverables (SOW 8.5)", "Minimum acceptance criteria (SOW 9.6)"], rows,
      row_h=0.72, header_h=0.4, size=8.5, header_size=9.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.86, ["Acceptance flow: check platforms and versions → test interfaces and functions → verify integration and permissions → reference capacity and model tests → clear issues → confirm the application.",
               "Responsibility: RNP provides branding, knowledge content and business rules, confirms Portuguese terminology and interfaces, approves go-live and operates users, content and services. iFLYTEK does not take on user growth, content production, human moderation or managed business operations. Basis: SOW 3.9 / 8.5 / 9.6 / 10.7."],
     h=0.62, size=8.5)

# ============================== A14 TR R&D support mapping ==============================
s = appendix_slide("A14 · Supporting Brazil's own R&D: TR requirement mapping", "Q: How do you answer the TR's R&D requirements?")
rows = [
    ["R&D cooperation plan", "Objectives, research directions supported, participating institutions, resources, responsibilities, schedule, indicators and acceptance criteria",
     "Four specialist directions: cluster build and scheduling, long stable training of large models, large-scale corpus construction, high-throughput inference optimisation; {{written into the cooperation plan once matched to RNP's priority tasks}}"],
    ["Interfaces, tools and materials", "APIs, connectors, SDKs and toolchains, data formats and interface specs, logs, performance metrics and usage documentation",
     "Platform interfaces, developer documentation and samples are delivered — see Appendix A4; {{list, versions, permissions and terms of use confirmed against the research tasks}}"],
    ["R&D technical support", "Support for development, adaptation, testing and optimisation across operators, libraries, runtimes, compilers and training/inference frameworks",
     "Expert Q&A, specialist walkthroughs, design review, demonstrations, shadowing and regular technical discussion; {{in-scope support and separately agreed premium services listed apart}}"],
    ["Practical training and knowledge transfer", "Practical training and knowledge transfer for the researchers, engineers and O&M staff RNP designates",
     "Four-stage knowledge transfer — see pages 10–11 and Appendix A8; topics in A9, learning path by role in A10"],
    ["Retained results and progress reports", "A version-controlled technical repository; regular reporting of research progress",
     "Course materials, operating manuals and case sharing; {{archive location and access rights for scripts, configurations, test results and progress reports confirmed before implementation}}"],
    ["Interoperability and continued use", "Interoperability and comparison of technical options; proprietary dependencies, licence limits, future cost and alternatives explained",
     "See Appendix A6 on handover, licensing and operational support; {{open scope, modification rights and terms of continued use explained component by component}}"],
]
table(s, L, 1.02, [2.6, 4.3, 5.43], ["TR requirement", "What it asks for", "How we answer it"], rows,
      row_h=0.68, header_h=0.4, size=8.5, header_size=9.5,
      header_fill="F2F2F2", header_color=TXT, vlines=False)
note(s, 5.60, ["FP1–FP9 are RNP's research fronts (model construction, model serving, distributed execution, compiler optimisation, cluster management, data, cross-institution collaboration, performance, system security). This table shows how the contractor supports that research; it does not mean iFLYTEK covers every front alone.",
               "These requirements address the contractor side as a whole: the base software stack is Huawei's, as shown on page 14. Basis: SOW and the R&D cooperation plan confirmed before implementation."],
     h=0.62, size=8.5)

for shp, code in LINKS + MENU + NAV:
    if code in APPX:
        shp.click_action.target_slide = APPX[code]
for shp in BACK:
    shp.click_action.target_slide = INDEX["index"]
for shp, tgt in BACK2:
    shp.click_action.target_slide = tgt
drop_first_slide(prs)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
