# -*- coding: utf-8 -*-
"""
When Does User-Generated Content Build Brand Resonance?
The Roles of Content Credibility, Consumer Involvement, and Brand Trust

GitHub  ->  Google Colab  ->  Word (.docx, journal layout)  ->  Telegram

HOW TO RUN IN COLAB
-------------------
1. Push this file to GitHub.
2. In Colab, open the left "Secrets" (key icon) panel and add:
       TELEGRAM_BOT_TOKEN   = <your NEW bot token from @BotFather>
       TELEGRAM_CHAT_ID     = 1262412842          (optional - this is the default)
   and switch "Notebook access" ON.
3. Run these cells:

       !pip -q install python-docx scipy matplotlib pandas numpy requests
       !wget -q https://raw.githubusercontent.com/<USER>/<REPO>/main/ugc_brand_resonance_article.py -O ugc.py
       %run ugc.py

OPTIONAL: upload your real survey file as  survey_data.csv  (columns listed in
the README block below). If the file is absent the script runs on SIMULATED
demo data so the pipeline can be tested - replace it before submission.

survey_data.csv columns:
   CC1..CC4  (content credibility)   CI1..CI4 (consumer involvement)
   BT1..BT4  (brand trust)           BR1..BR4 (brand resonance)
   Platform, Gender, AgeGroup        (all 1-5 Likert for the items)
"""

import os
import getpass
import warnings

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import requests
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

warnings.filterwarnings("ignore")

# =====================================================================
# 1. JOURNAL / MANUSCRIPT SETTINGS  (edit freely)
# =====================================================================
TITLE = ("When Does User-Generated Content Build Brand Resonance? "
         "The Roles of Content Credibility, Consumer Involvement, and Brand Trust")
JOURNAL = "Journal of Marketing & Social Research"
ISSN = "ISSN (Online): 3008-0711"
VOLUME = "Volume: 02 | Issue 10 | 2025"
HOMEPAGE = "https://jmsr-online.com/"
PAGES = "130\u2013139"
START_PAGE = 130

AUTHORS = [("Sunil G E", "1*"), ("Dr. Sowmya Erappa K", "2"), ("Dhanalakshmi R", "3"),
           ("Thriveni H R", "4"), ("Pavithra H B", "5")]
AFFILIATIONS = [
    ("1*", "Assistant Professor, DoS in Commerce, STG First Grade College, University of Mandya, Mandya District, Karnataka 571455 India."),
    ("2", "Principal, Seshadripuram Degree College, University of Mysore, Mysuru, Karnataka 570016 India."),
    ("3", "Assistant Professor, DoS in Commerce, Geetha Shishu Shikshana Sangha (R) Simha Subbamahalakshmi First Grade College, Metagalli, Mysore College Code 1052 Karnataka, India."),
    ("4", "Assistant professor and HoD of Commerce, Sri Adichunchanagiri College of Arts Commerce and Science, Nagamangala, Karnataka 571432 India."),
    ("5", "Assistant Professor, DoS in Commerce, STG First Grade College, University of Mandya, Mandya District 571455 Karnataka, India."),
]
RECEIVED, REVISION, ACCEPTED, PUBLISHED = "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY"
CORR_NAME, CORR_EMAIL = "Sunil G E", "sunilge2000@gmail.com"

KEYWORDS = ("User-generated content, Content credibility, Brand trust, "
            "Consumer involvement, Brand resonance, Moderated mediation.")
FONT = "Times New Roman"
BOOT = 5000
SEED = 2025

# =====================================================================
# 2. CONSTRUCTS
# =====================================================================
CONSTRUCTS = {
    "CC": ("Content credibility", "Ismagilova et al. (2020)",
           ["The brand-related content posted by other users is believable.",
            "The users who post about this brand seem honest.",
            "The experiences shared by other users appear genuine.",
            "I can rely on what other users say about this brand online."]),
    "CI": ("Consumer involvement", "Zaichkowsky (1985)",
           ["This product category is important to me.",
            "I take a keen interest in this product category.",
            "Choosing a brand in this category matters to me.",
            "This category is relevant to my daily life."]),
    "BT": ("Brand trust", "Chaudhuri & Holbrook (2001)",
           ["I trust this brand.",
            "I rely on this brand.",
            "This is an honest brand.",
            "This brand would make any effort to satisfy me."]),
    "BR": ("Brand resonance", "Keller (2001), adapted",
           ["I feel a strong sense of loyalty to this brand.",
            "I feel this brand is a part of who I am.",
            "I follow and talk about this brand in online communities.",
            "I am proud to tell others that I use this brand."]),
}
KEYS = ["CC", "CI", "BT", "BR"]

# =====================================================================
# 3. DATA  (real CSV if present, otherwise simulated demo data)
# =====================================================================
def simulate(n=320, seed=SEED):
    rng = np.random.default_rng(seed)
    cc = rng.normal(0, 1, n)
    ci = rng.normal(0, 1, n)
    bt = 0.45 * cc + 0.20 * ci + 0.15 * cc * ci + rng.normal(0, 0.75, n)
    br = 0.40 * bt + 0.20 * cc + 0.25 * ci + rng.normal(0, 0.70, n)
    d = {}
    for pre, lat in (("CC", cc), ("CI", ci), ("BT", bt), ("BR", br)):
        z = (lat - lat.mean()) / lat.std()
        for i in range(1, 5):
            v = 3.6 + 0.64 * z + rng.normal(0, 0.55, n)
            d[f"{pre}{i}"] = np.clip(np.rint(v), 1, 5).astype(int)
    df = pd.DataFrame(d)
    df["Platform"] = rng.choice(["Instagram", "YouTube", "Facebook", "X (Twitter)", "Online review sites"],
                                size=n, p=[.30, .25, .20, .10, .15])
    df["Gender"] = rng.choice(["Male", "Female", "Other"], size=n, p=[.48, .50, .02])
    df["AgeGroup"] = rng.choice(["18-24", "25-34", "35-44", "45+"], size=n, p=[.34, .36, .20, .10])
    return df


def load_data():
    for path in (os.environ.get("DATA_CSV", ""), "survey_data.csv"):
        if path and os.path.exists(path):
            print(f"Using real survey data: {path}")
            return pd.read_csv(path), False
    print("!! survey_data.csv not found -> using SIMULATED demo data (replace before submission) !!")
    return simulate(), True


# =====================================================================
# 4. STATISTICS (numpy/scipy only)
# =====================================================================
def ols(xs, y):
    n = len(y)
    X = np.column_stack([np.ones(n)] + list(xs))
    k = X.shape[1]
    XtXi = np.linalg.inv(X.T @ X)
    b = XtXi @ X.T @ y
    res = y - X @ b
    sse = float(res @ res)
    sst = float(((y - y.mean()) ** 2).sum())
    df = n - k
    se = np.sqrt(np.diag(XtXi) * sse / df)
    t = b / se
    p = 2 * stats.t.sf(np.abs(t), df)
    return dict(b=b, se=se, t=t, p=p, r2=1 - sse / sst, sse=sse, df=df, k=k, n=n)


def f_change(full, restr):
    q = full["k"] - restr["k"]
    F = ((restr["sse"] - full["sse"]) / q) / (full["sse"] / full["df"])
    return F, float(stats.f.sf(F, q, full["df"]))


def cronbach(items):
    a = np.asarray(items, float)
    k = a.shape[1]
    return k / (k - 1) * (1 - a.var(axis=0, ddof=1).sum() / a.sum(axis=1).var(ddof=1))


def paths(x, w, m, y):
    one = np.ones(len(x))
    a = np.linalg.lstsq(np.column_stack([one, x, w, x * w]), m, rcond=None)[0]
    b = np.linalg.lstsq(np.column_stack([one, x, m, w]), y, rcond=None)[0]
    return a, b


def effects(a, b, levels):
    out = [(a[1] + a[3] * w) * b[2] for w in levels]
    return out + [a[3] * b[2], b[1]]


def analyse(df):
    R = {"n": len(df)}
    comp = pd.DataFrame({k: df[[f"{k}{i}" for i in range(1, 5)]].mean(axis=1) for k in KEYS})
    R["comp"] = comp
    R["desc"] = {k: (comp[k].mean(), comp[k].std(),
                     cronbach(df[[f"{k}{i}" for i in range(1, 5)]])) for k in KEYS}

    # correlations
    corr = np.zeros((4, 4)); cp = np.zeros((4, 4))
    for i, a in enumerate(KEYS):
        for j, b in enumerate(KEYS):
            corr[i, j], cp[i, j] = stats.pearsonr(comp[a], comp[b])
    R["corr"], R["corr_p"] = corr, cp

    # centred variables
    x = (comp.CC - comp.CC.mean()).values
    w = (comp.CI - comp.CI.mean()).values
    m = (comp.BT - comp.BT.mean()).values
    y = comp.BR.values
    R["sd_w"] = w.std(ddof=1)

    # hierarchical regression on brand resonance
    M1 = ols([x], y)
    M2 = ols([x, w, m], y)
    M3 = ols([x, w, m, x * w], y)
    R["M"] = (M1, M2, M3)
    R["dR2"] = [(M1["r2"], None, None),
                (M2["r2"] - M1["r2"],) + f_change(M2, M1),
                (M3["r2"] - M2["r2"],) + f_change(M3, M2)]

    # path models (Hayes Model 7 logic)
    A = ols([x, w, x * w], m)
    B = ols([x, m, w], y)
    R["A"], R["B"] = A, B
    a, b = paths(x, w, m, y)
    levels = [-R["sd_w"], 0.0, R["sd_w"]]
    R["levels"] = levels
    est = effects(a, b, levels)
    rng = np.random.default_rng(SEED)
    n = len(y)
    bs = np.empty((BOOT, 5))
    for r in range(BOOT):
        idx = rng.integers(0, n, n)
        ab, bb = paths(x[idx], w[idx], m[idx], y[idx])
        bs[r] = effects(ab, bb, levels)
    R["boot"] = [(est[i], bs[:, i].std(ddof=1), *np.percentile(bs[:, i], [2.5, 97.5])) for i in range(5)]

    # simple slopes of CC -> BT
    R["slopes"] = [(lv, a[1] + a[3] * lv) for lv in levels]
    R["a_coef"] = a

    # ANOVA by platform
    R["anova"] = None
    if "Platform" in df:
        rows = []
        for k in ("CC", "BT", "BR"):
            g = [comp.loc[df.Platform == p, k].values for p in df.Platform.unique()]
            allv = comp[k].values
            gm = allv.mean()
            ssb = sum(len(v) * (v.mean() - gm) ** 2 for v in g)
            ssw = sum(((v - v.mean()) ** 2).sum() for v in g)
            d1, d2 = len(g) - 1, len(allv) - len(g)
            F = (ssb / d1) / (ssw / d2)
            rows.append((k, ssb, d1, ssb / d1, F, float(stats.f.sf(F, d1, d2)),
                         float(stats.f.ppf(.95, d1, d2)), ssw, d2, ssw / d2, ssb + ssw, d1 + d2))
        R["anova"] = rows
        R["plat_mean"] = comp.groupby(df.Platform)["BR"].mean().sort_values()

    # hypotheses
    c = M1["b"][1]; pc = M1["p"][1]
    bo = R["boot"]
    R["H"] = [
        ("H1", "Content credibility \u2192 Brand resonance (total effect)", f"B = {c:.3f}, p {pf(pc)}", c > 0 and pc < .05),
        ("H2", "Content credibility \u2192 Brand trust", f"B = {A['b'][1]:.3f}, p {pf(A['p'][1])}", A["b"][1] > 0 and A["p"][1] < .05),
        ("H3", "Brand trust \u2192 Brand resonance", f"B = {B['b'][2]:.3f}, p {pf(B['p'][2])}", B["b"][2] > 0 and B["p"][2] < .05),
        ("H4", "Brand trust mediates credibility \u2192 resonance", f"Indirect = {bo[1][0]:.3f}, 95% CI [{bo[1][2]:.3f}, {bo[1][3]:.3f}]", bo[1][2] > 0),
        ("H5", "Involvement moderates credibility \u2192 trust", f"B = {A['b'][3]:.3f}, p {pf(A['p'][3])}", A["b"][3] > 0 and A["p"][3] < .05),
        ("H6", "Indirect effect stronger at higher involvement", f"Index = {bo[3][0]:.3f}, 95% CI [{bo[3][2]:.3f}, {bo[3][3]:.3f}]", bo[3][2] > 0),
    ]
    return R


# ---------- small text helpers ----------
def pf(p):
    return "< 0.001" if p < .001 else f"= {p:.3f}"


def stars(p):
    return "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""


def sup(flag, yes, no):
    return yes if flag else no


def hs(R, i):
    return "supported" if R["H"][i][3] else "not supported"


# =====================================================================
# 5. FIGURES
# =====================================================================
def make_figures(R):
    f1 = "fig1_platform_resonance.png"
    s = R["plat_mean"] if R.get("anova") else pd.Series(dtype=float)
    if len(s):
        fig, ax = plt.subplots(figsize=(8, 3.8))
        ax.barh(s.index, s.values, color="#87CEEB")
        for i, v in enumerate(s.values):
            ax.text(v + .03, i, f"{v:.2f}", va="center", fontsize=9)
        ax.set_xlim(0, 5); ax.set_xlabel("Mean brand resonance (1\u20135)")
        ax.set_title("Brand Resonance by UGC Platform \u2013 Survey Respondents", fontsize=11)
        ax.grid(axis="x", alpha=.3); fig.tight_layout(); fig.savefig(f1, dpi=200); plt.close(fig)
    else:
        f1 = None
    f2 = "fig2_simple_slopes.png"
    a = R["a_coef"]; sd = R["sd_w"]
    xs = np.linspace(-1.5, 1.5, 20)
    fig, ax = plt.subplots(figsize=(8, 3.8))
    for lab, lv, col in (("Low involvement (\u22121 SD)", -sd, "#d62728"),
                         ("Mean involvement", 0, "#7f7f7f"), ("High involvement (+1 SD)", sd, "#1f77b4")):
        ax.plot(xs, a[0] + (a[1] + a[3] * lv) * xs + a[2] * lv, label=lab, color=col, lw=2)
    ax.set_xlabel("Content credibility (mean-centred)"); ax.set_ylabel("Predicted brand trust")
    ax.set_title("Conditional Effect of Content Credibility on Brand Trust", fontsize=11)
    ax.legend(frameon=False); ax.grid(alpha=.3); fig.tight_layout(); fig.savefig(f2, dpi=200); plt.close(fig)
    return f1, f2


# =====================================================================
# 6. WORD HELPERS
# =====================================================================
J, L, C = WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER


def style_run(r, size=10, bold=False, italic=False, color=None, sup_=False, underline=False):
    r.font.name = FONT
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    r.font.size = Pt(size); r.bold = bold; r.italic = italic; r.underline = underline
    r.font.superscript = sup_
    if color:
        r.font.color.rgb = RGBColor.from_string(color)


def para(cont, content="", size=10, bold=False, italic=False, align=J, color=None,
         after=4, before=0, keep=False, style=None, first=False):
    p = cont.paragraphs[0] if first else (cont.add_paragraph(style=style) if style else cont.add_paragraph())
    p.alignment = align
    pf_ = p.paragraph_format
    pf_.space_after, pf_.space_before, pf_.line_spacing = Pt(after), Pt(before), 1.0
    pf_.keep_with_next = keep
    if isinstance(content, str):
        content = [(content, {})]
    for text, o in content:
        r = p.add_run(text)
        style_run(r, o.get("size", size), o.get("bold", bold), o.get("italic", italic),
                  o.get("color", color), o.get("sup", False), o.get("ul", False))
    return p


def shade(el_pr, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    el_pr.append(shd)


def border_par(p, side="bottom", sz=6, color="000000"):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    e = OxmlElement(f"w:{side}")
    e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "1"); e.set(qn("w:color"), color)
    b.append(e); pPr.append(b)


def h1(d, t):
    return para(d, t.upper(), size=12, bold=True, align=L, before=6, after=3, keep=True)


def h2(d, t):
    return para(d, t, size=10.5, bold=True, align=L, before=4, after=2, keep=True)


def bullet(d, t):
    p = para(d, t, style="List Bullet", after=2)
    return p


def add_field(run, instr):
    for typ, txt in (("begin", None), (None, instr), ("end", None)):
        if typ:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), typ)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt
        run._r.append(e)


def set_cols(section, num):
    cols = section._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols"); section._sectPr.append(cols)
    cols.set(qn("w:num"), str(num)); cols.set(qn("w:space"), "425")


def make_footer(ft):
    ft.is_linked_to_previous = False
    p = ft.paragraphs[0]
    p.alignment = C
    border_par(p, "top", 6)
    r = p.add_run(); style_run(r, 10); add_field(r, "PAGE")
    para(ft, [("Name: ", {"bold": True}), (CORR_NAME, {"bold": True})], size=9, align=L, after=0)
    para(ft, [("Email: ", {"bold": True}), (CORR_EMAIL, {"bold": True, "color": "0000FF", "ul": True})], size=9, align=L, after=0)


def caption(d, t):
    return para(d, t, size=10, bold=True, align=C, before=6, after=2, keep=True)


def source(d, t):
    return para(d, [("Source: ", {"bold": True}), (t, {})], size=9, align=L, after=6)


def table(d, header, rows, widths, size=8.5, bold_rows=(), shade_rows=(), left_cols=(0,)):
    t = d.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); t._tbl.tblPr.append(lay)
    for i, wd_ in enumerate(widths):
        t.columns[i].width = Cm(wd_)

    def fill(cell, text, bold=False, al=C, wd=None):
        cell.paragraphs[0].text = ""
        para(cell, str(text), size=size, bold=bold, align=al, after=1, first=True)
        if wd:
            cell.width = Cm(wd)

    for i, h in enumerate(header):
        fill(t.rows[0].cells[i], h, True, C, widths[i])
        shade(t.rows[0].cells[i]._tc.get_or_add_tcPr(), "D9D9D9")
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            fill(cells[i], v, ri in bold_rows, L if i in left_cols else C, widths[i])
            if ri in shade_rows:
                shade(cells[i]._tc.get_or_add_tcPr(), "F2F2F2")
    return t


# =====================================================================
# 7. THE ARTICLE
# =====================================================================
def build_doc(R, figs, simulated, out="UGC_Brand_Resonance_Article.docx"):
    n = R["n"]
    A, B = R["A"], R["B"]
    M1, M2, M3 = R["M"]
    bo = R["boot"]
    H = R["H"]
    d = Document()
    st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    s0 = d.sections[0]
    s0.page_width, s0.page_height = Cm(21), Cm(29.7)
    for a_ in ("left_margin", "right_margin"):
        setattr(s0, a_, Cm(1.9))
    s0.top_margin, s0.bottom_margin = Cm(1.8), Cm(2.0)
    pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), str(START_PAGE))
    s0._sectPr.find(qn("w:cols")).addprevious(pg)
    s0.different_first_page_header_footer = True
    make_footer(s0.footer); make_footer(s0.first_page_footer)
    s0.first_page_header.is_linked_to_previous = False
    hd = s0.header; hd.is_linked_to_previous = False
    hp = para(hd, [("How to Cite: ", {"bold": True}),
                   (f"Sunil G E, et, al. {TITLE}. {JOURNAL} vol. 2, no. 10, 2025, pp. {PAGES}.", {})],
              size=9, align=L, after=2, first=True)
    border_par(hp, "bottom", 6)

    # ---------- front matter ----------
    para(d, JOURNAL, size=13, bold=True, align=L, after=2)
    para(d, ISSN, size=11, bold=True, align=L, after=2)
    para(d, VOLUME, size=9, bold=True, align=L, after=0)
    para(d, [("Journal homepage: ", {}), (HOMEPAGE, {"color": "0000FF", "ul": True})], size=8.5, align=L, after=4)
    bar = d.add_table(rows=1, cols=2); bar.autofit = False
    c0, c1 = bar.rows[0].cells
    c0.width, c1.width = Cm(3.0), Cm(14.2)
    shade(c0._tc.get_or_add_tcPr(), "00A651"); shade(c1._tc.get_or_add_tcPr(), "EEEAD9")
    para(c0, "Research Article", size=12, bold=True, color="FFFFFF", align=L, after=1, first=True)
    para(c1, "", first=True, after=1)
    para(d, "", after=2)
    para(d, TITLE, size=17, bold=True, align=J, after=6)
    runs = []
    for i, (nm, sp) in enumerate(AUTHORS):
        sep = " and " if i == len(AUTHORS) - 1 and i > 0 else (", " if i else "")
        runs += [(sep, {}), (nm, {"bold": True}), (sp, {"bold": True, "sup": True})]
    para(d, runs, size=10, align=L, after=2)
    for k, t in AFFILIATIONS:
        para(d, [(k, {"sup": True}), (t, {})], size=8.5, align=L, after=0)
    p = para(d, [(f"Received: {RECEIVED};", {}), (f"          Revision: {REVISION};", {}),
                 (f"          Accepted: {ACCEPTED};", {}), (f"          Published: {PUBLISHED}", {})],
             size=9, bold=True, color="C55A11", align=L, before=3, after=2)
    border_par(p, "top", 6)
    para(d, [("*Corresponding author: ", {"bold": True}), (f"{CORR_NAME} (", {"bold": True}),
             (CORR_EMAIL, {"bold": True, "italic": True, "color": "0000FF", "ul": True}), (")", {"bold": True})],
         size=10.5, align=L, before=4, after=6)

    # ---------- abstract box ----------
    ab = (f"Brand resonance, the deepest form of the consumer\u2013brand relationship, is increasingly shaped by what other "
          f"consumers say about brands online. This study examines when user-generated content (UGC) builds brand resonance by "
          f"testing a moderated mediation model in which content credibility influences brand resonance through brand trust, "
          f"while consumer involvement conditions the credibility\u2013trust link. Survey data from {n} social media users who "
          f"recently engaged with brand-related UGC were analysed using reliability analysis, hierarchical regression, "
          f"bootstrapped conditional indirect effects ({BOOT:,} resamples) and one-way ANOVA. Content credibility had a "
          f"{sup(M1['b'][1] > 0, 'positive', 'negative')}, {sup(M1['p'][1] < .05, 'statistically significant', 'statistically non-significant')} "
          f"total effect on brand resonance (B = {M1['b'][1]:.2f}), and brand trust "
          f"{sup(bo[1][2] > 0, 'significantly mediated', 'did not significantly mediate')} this relationship "
          f"(indirect effect = {bo[1][0]:.2f}, 95% CI [{bo[1][2]:.2f}, {bo[1][3]:.2f}]). Consumer involvement "
          f"{sup(H[4][3], 'strengthened', 'did not significantly strengthen')} the effect of credibility on trust "
          f"(B = {A['b'][3]:.2f}, p {pf(A['p'][3])}), and the index of moderated mediation was "
          f"{sup(H[5][3], 'significant', 'not significant')}. "
          + (f"A one-way ANOVA showed that brand resonance "
             f"{sup(R['anova'][2][5] < .05, 'differed significantly', 'did not differ significantly')} across UGC platforms "
             f"(p {pf(R['anova'][2][5])}). " if R["anova"] else "")
          + "The findings highlight that credibility alone is not sufficient; trust is the mechanism through which credible UGC "
            "converts into resonance, and the strength of this conversion depends on how involved consumers are. The study "
            "recommends that brands curate credible UGC, build trust-based engagement, and tailor UGC strategies to consumer involvement.")
    box = d.add_table(rows=1, cols=1); box.style = "Table Grid"; box.autofit = False
    bc = box.rows[0].cells[0]; bc.width = Cm(17.2)
    shade(bc._tc.get_or_add_tcPr(), "F2F2F2")
    para(bc, [("Abstract: ", {"bold": True, "color": "1F4E9E"}), (ab, {})], size=9.5, after=3, first=True)
    para(bc, [("Keywords: ", {"bold": True, "color": "1F4E9E"}), (KEYWORDS, {})], size=9.5, after=2)

    # ---------- two-column body: Introduction -> Methodology ----------
    s1 = d.add_section(WD_SECTION.CONTINUOUS); s1.different_first_page_header_footer = False; set_cols(s1, 2)
    h1(d, "Introduction")
    for t in [
        "Consumers no longer rely only on brand-controlled communication. Reviews, posts, videos, comments and testimonials "
        "created by other consumers now form a large share of the information people see before and after purchase "
        "(Muntinga et al., 2011). This user-generated content (UGC) is perceived as independent of the marketer, which gives it "
        "unusual persuasive power, but also makes its quality uneven and its effect on brands conditional.",
        "A brand\u2019s ultimate objective is not merely awareness or a favourable attitude but resonance: an intense, active "
        "psychological bond in which consumers feel in sync with the brand (Keller, 2001). Resonance is reflected in loyalty, "
        "attachment, a sense of community and active engagement. Existing research shows that social media communication, "
        "including UGC, can shape brand perceptions (Schivinski & D\u0105browski, 2016), yet it remains unclear when and why "
        "exposure to UGC deepens the consumer\u2013brand relationship rather than leaving it unchanged.",
        "Three constructs are likely to matter. First, the credibility of the content determines whether consumers accept what "
        "other users say (Ismagilova et al., 2020). Second, brand trust, the willingness to rely on a brand\u2019s ability and "
        "intentions (Chaudhuri & Holbrook, 2001), is a plausible mechanism by which credible content is converted into a "
        "stronger relationship. Third, consumer involvement determines how carefully people process brand information "
        "(Petty et al., 1983; Zaichkowsky, 1985), and may therefore decide how strongly credible content influences trust.",
        "The present study therefore asks when UGC builds brand resonance. It tests a moderated mediation model in which content "
        "credibility affects brand resonance through brand trust and consumer involvement moderates the credibility\u2013trust "
        "link. The results help brand managers decide which UGC deserves amplification and for which consumers it is most "
        "likely to work."]:
        para(d, t)

    h1(d, "Review of Literature")
    for t in [
        "UGC refers to brand-related content created and shared by consumers rather than by firms. Muntinga et al. (2011) "
        "describe consumers\u2019 brand-related online activities as ranging from consuming to contributing to creating content, "
        "each driven by different motivations. Because the content is not paid for by the brand, readers treat it as more "
        "authentic, which is the main reason why UGC is expected to be influential. Reviews of electronic word-of-mouth research "
        "emphasise that its effect depends on characteristics of the message, the source and the receiver (Cheung & Thadani, 2012).",
        "Credibility is the most frequently examined source- and message-related factor. A meta-analysis of source credibility "
        "characteristics shows that expertise, trustworthiness and related cues influence consumer behaviour (Ismagilova et al., "
        "2020). In the review-helpfulness literature, the diagnosticity and perceived quality of content determine whether it "
        "is adopted (Filieri, 2015). Credible UGC reduces perceived risk, which is expected to raise trust in the brand being discussed.",
        "Brand trust is the confident expectation that the brand will deliver on its promises. Trust is a core ingredient of "
        "durable relationships (Morgan & Hunt, 1994) and a driver of brand loyalty and attitudinal commitment (Chaudhuri & "
        "Holbrook, 2001). In Keller\u2019s (2001) brand-equity pyramid, judgements and feelings about the brand lead to "
        "resonance at the apex. Trust is therefore expected to link credible content with resonance.",
        "Involvement reflects the personal relevance of a product category (Zaichkowsky, 1985). In the elaboration likelihood "
        "framework, highly involved receivers scrutinise message content, whereas low-involvement receivers rely on peripheral cues "
        "(Petty et al., 1983). For UGC, this implies that highly involved consumers should be more sensitive to the credibility "
        "of the content when forming trust.",
        "Hypotheses development. Based on this reasoning, content credibility is expected to enhance brand resonance (H1) and brand "
        "trust (H2); brand trust is expected to enhance resonance (H3) and to mediate the credibility\u2013resonance link (H4); "
        "involvement is expected to strengthen the credibility\u2013trust relationship (H5), making the indirect effect conditional "
        "on involvement (H6)."]:
        para(d, t)

    h2(d, "Research Gap")
    para(d, "Prior studies mostly examine UGC as a driver of purchase intention, attitude or brand equity, and often treat "
            "credibility, trust and involvement separately. Few studies place trust as the explanatory mechanism and involvement as "
            "the boundary condition in one model of brand resonance, the strongest relational outcome. There is also limited "
            "evidence combining the mechanism (how) and the boundary condition (when) with a platform-wise comparison. This "
            "study addresses this gap.")
    h2(d, "Objectives of the study")
    bullet(d, "To examine the effect of content credibility of UGC on brand resonance.")
    bullet(d, "To test the mediating role of brand trust in the relationship between content credibility and brand resonance.")
    bullet(d, "To test whether consumer involvement moderates the credibility\u2013trust link and the resulting indirect effect.")
    bullet(d, "To compare credibility, trust and resonance across the UGC platforms used by consumers.")
    h2(d, "Research Questions")
    bullet(d, "Does content credibility of UGC build brand resonance, and to what extent does brand trust explain this effect?")
    bullet(d, "Does the strength of this process depend on the level of consumer involvement?")
    h2(d, "Hypotheses of the study")
    for t in ["H1: Content credibility has a positive effect on brand resonance.",
              "H2: Content credibility has a positive effect on brand trust.",
              "H3: Brand trust has a positive effect on brand resonance.",
              "H4: Brand trust mediates the effect of content credibility on brand resonance.",
              "H5: Consumer involvement moderates the effect of content credibility on brand trust, such that the effect is stronger at higher involvement.",
              "H6: The indirect effect of content credibility on brand resonance through brand trust is conditional on consumer involvement (moderated mediation)."]:
        para(d, t, after=2)

    h1(d, "Research Methodology")
    para(d, "This study adopts a descriptive and analytical, cross-sectional survey design to examine how content credibility, "
            "brand trust and consumer involvement jointly explain brand resonance among consumers exposed to brand-related UGC.")
    h2(d, "7.1 Data Source")
    para(d, f"The study is based on primary data collected through a structured questionnaire from {n} social media users who had "
            "viewed brand-related UGC (reviews, posts, videos or comments) in the previous three months. Respondents answered "
            "with reference to a brand they had recently encountered in UGC."
            + (" [NOTE: this version of the manuscript was generated with SIMULATED data to demonstrate the analytical pipeline; "
               "replace with the actual survey file before submission.]" if simulated else ""))
    h2(d, "7.2 Sample Design")
    para(d, "Respondents were selected through purposive sampling with a screening question on recent exposure to brand-related UGC. "
            "The sample covers different genders, age groups and UGC platforms (Table 1), which allows platform-wise comparison.")
    h2(d, "7.3 Measurement Framework")
    para(d, "Four constructs were measured with four items each on a five-point Likert scale (1 = strongly disagree, 5 = strongly "
            "agree): content credibility, consumer involvement, brand trust and brand resonance (Table 2). Items were adapted from "
            "established scales and averaged to create construct scores. Internal consistency was assessed with Cronbach\u2019s alpha "
            "(Cronbach, 1951), with 0.70 as the acceptable threshold (Hair et al., 2019).")
    h2(d, "7.4 Scale Coding")
    for t in ["1 = Strongly disagree", "2 = Disagree", "3 = Neutral", "4 = Agree", "5 = Strongly agree"]:
        para(d, t, after=0, align=L)
    para(d, "Independent and moderating variables were mean-centred before computing the interaction term.", before=3)
    h2(d, "7.5 Statistical Analysis")
    para(d, "Descriptive statistics, reliability and Pearson correlations were computed first. Hierarchical regression tested the "
            "effects on brand resonance. The mediation and moderated-mediation model, in which involvement moderates the first stage "
            f"(credibility \u2192 trust), was estimated using the conditional-process logic of Hayes (2018), with {BOOT:,} bootstrap "
            "resamples and 95% percentile confidence intervals for conditional indirect effects at \u22121 SD, the mean and +1 SD of "
            "involvement. A one-way ANOVA compared the constructs across UGC platforms, with platform as the independent variable.")

    # ---------- single column: results ----------
    s2 = d.add_section(WD_SECTION.CONTINUOUS); s2.different_first_page_header_footer = False; set_cols(s2, 1)
    h1(d, "Data Analysis and Results")
    para(d, "This section presents the analysis of the relationships among content credibility, brand trust, consumer involvement and "
            "brand resonance. It reports the sample profile, reliability, correlations, hierarchical regression, the moderated "
            "mediation model and platform-wise differences.")
    comp = R["comp"]

    # Table 1
    para(d, "The profile of the respondents is summarised in Table 1.")
    caption(d, "Table No: 1 Profile of respondents")
    rows = R["profile_rows"]
    table(d, ["Variable", "Category", "n", "%"], rows, [4.5, 6.2, 3.0, 3.0])
    source(d, "Primary survey data." + (" (Simulated demonstration data.)" if simulated else ""))

    # Table 2
    para(d, "Table 2 lists the constructs, a sample item, the source of the scale and the reliability of each construct.")
    caption(d, "Table No: 2 Constructs, scale sources and reliability")
    rows = [[CONSTRUCTS[k][0], "4", CONSTRUCTS[k][2][0], CONSTRUCTS[k][1], f"{R['desc'][k][2]:.3f}"] for k in KEYS]
    table(d, ["Construct", "Items", "Sample item", "Adapted from", "Cronbach\u2019s \u03b1"], rows, [3.0, 1.2, 7.2, 3.6, 2.2])
    source(d, "Authors\u2019 questionnaire and reliability analysis.")
    para(d, [("Interpretation: ", {"bold": True}),
             (f"All constructs have alpha values of {min(R['desc'][k][2] for k in KEYS):.2f} or higher, "
              + sup(min(R['desc'][k][2] for k in KEYS) >= .70, "above the 0.70 benchmark, indicating acceptable internal consistency.",
                    "and at least one construct falls below the 0.70 benchmark, so results for that construct should be read with caution."), {})])

    # Table 3
    para(d, "Means, standard deviations and inter-construct correlations are presented in Table 3.")
    caption(d, "Table No: 3 Descriptive statistics and correlation matrix")
    rows = []
    for i, k in enumerate(KEYS):
        r_ = [CONSTRUCTS[k][0], f"{R['desc'][k][0]:.2f}", f"{R['desc'][k][1]:.2f}"]
        for j in range(4):
            r_.append("\u2013" if j == i else (f"{R['corr'][i, j]:.3f}{stars(R['corr_p'][i, j])}" if j < i else ""))
        rows.append(r_)
    table(d, ["Construct", "Mean", "SD", "1", "2", "3", "4"], rows, [4.4, 1.7, 1.7, 2.35, 2.35, 2.35, 2.35])
    source(d, "Authors\u2019 analysis. 1 = Content credibility, 2 = Consumer involvement, 3 = Brand trust, 4 = Brand resonance. "
              "* p < 0.05, ** p < 0.01, *** p < 0.001.")
    para(d, [("Interpretation: ", {"bold": True}),
             (f"Brand resonance is correlated with content credibility (r = {R['corr'][0, 3]:.2f}), involvement (r = {R['corr'][1, 3]:.2f}) "
              f"and brand trust (r = {R['corr'][2, 3]:.2f}), and brand trust is correlated with credibility (r = {R['corr'][0, 2]:.2f}), "
              "which supports further testing of the hypothesised model.", {})])

    # Table 4
    para(d, "Table 4 reports the hierarchical regression with brand resonance as the dependent variable.")
    caption(d, "Table No: 4 Hierarchical regression for brand resonance (B, SE in parentheses)")
    names = ["(Constant)", "Content credibility", "Consumer involvement", "Brand trust", "Credibility \u00d7 Involvement"]
    idx = {1: [0, 1], 2: [0, 1, 2, 3], 3: [0, 1, 2, 3, 4]}
    rows = []
    for j, nm in enumerate(names):
        r_ = [nm]
        for mi, M in enumerate((M1, M2, M3), 1):
            if j in idx[mi]:
                q = idx[mi].index(j)
                r_.append(f"{M['b'][q]:.3f}{stars(M['p'][q])} ({M['se'][q]:.3f})")
            else:
                r_.append("\u2013")
        rows.append(r_)
    rows.append(["R\u00b2", f"{M1['r2']:.3f}", f"{M2['r2']:.3f}", f"{M3['r2']:.3f}"])
    rows.append(["\u0394R\u00b2 (F-change, p)", "\u2013",
                 f"{R['dR2'][1][0]:.3f} (F = {R['dR2'][1][1]:.2f}, p {pf(R['dR2'][1][2])})",
                 f"{R['dR2'][2][0]:.3f} (F = {R['dR2'][2][1]:.2f}, p {pf(R['dR2'][2][2])})"])
    table(d, ["Predictor", "Model 1", "Model 2", "Model 3"], rows, [4.8, 4.0, 4.2, 4.2], bold_rows=(5, 6))
    source(d, "Authors\u2019 analysis; predictors mean-centred. * p < 0.05, ** p < 0.01, *** p < 0.001.")
    para(d, [("Interpretation: ", {"bold": True}),
             (f"Content credibility alone explains {M1['r2'] * 100:.1f}% of the variance in brand resonance (H1 is {hs(R, 0)}). "
              f"Adding involvement and brand trust raises R\u00b2 to {M2['r2']:.3f}, and brand trust remains "
              f"{sup(M2['p'][3] < .05, 'a significant', 'a non-significant')} predictor (B = {M2['b'][3]:.3f}). The interaction term "
              f"{sup(R['dR2'][2][2] < .05, 'adds significant explanatory power', 'does not add significant explanatory power')} "
              f"to the resonance model (\u0394R\u00b2 = {R['dR2'][2][0]:.3f}).", {})])

    # Table 5
    para(d, "The path coefficients of the moderated mediation model are shown in Table 5.")
    caption(d, "Table No: 5 Path coefficients of the moderated mediation model")
    rows = [["Mediator model: DV = Brand trust", "", "", "", ""]]
    for q, nm in enumerate(["(Constant)", "Content credibility (a)", "Consumer involvement", "Credibility \u00d7 Involvement (a3)"]):
        rows.append([nm, f"{A['b'][q]:.3f}", f"{A['se'][q]:.3f}", f"{A['t'][q]:.2f}", f"{A['p'][q]:.3f}"])
    rows.append(["R\u00b2", f"{A['r2']:.3f}", "", "", ""])
    rows.append(["Outcome model: DV = Brand resonance", "", "", "", ""])
    for q, nm in enumerate(["(Constant)", "Content credibility (direct, c\u2032)", "Brand trust (b)", "Consumer involvement"]):
        rows.append([nm, f"{B['b'][q]:.3f}", f"{B['se'][q]:.3f}", f"{B['t'][q]:.2f}", f"{B['p'][q]:.3f}"])
    rows.append(["R\u00b2", f"{B['r2']:.3f}", "", "", ""])
    table(d, ["Predictor", "B", "SE", "t", "p"], rows, [6.6, 2.7, 2.7, 2.6, 2.6], bold_rows=(0, 6), shade_rows=(0, 6))
    source(d, "Authors\u2019 analysis (Hayes, 2018, Model 7 logic); predictors mean-centred.")
    para(d, [("Interpretation: ", {"bold": True}),
             (f"Content credibility has a {sup(A['b'][1] > 0, 'positive', 'negative')} and "
              f"{sup(A['p'][1] < .05, 'significant', 'non-significant')} effect on brand trust (a = {A['b'][1]:.3f}), and brand trust has a "
              f"{sup(B['p'][2] < .05, 'significant', 'non-significant')} effect on brand resonance (b = {B['b'][2]:.3f}). The "
              f"credibility \u00d7 involvement interaction is {sup(A['p'][3] < .05, 'significant', 'not significant')} "
              f"(a3 = {A['b'][3]:.3f}, p {pf(A['p'][3])}).", {})])

    # Table 6
    para(d, f"Table 6 presents the bootstrapped direct, conditional indirect and moderated-mediation effects ({BOOT:,} resamples).")
    caption(d, "Table No: 6 Bootstrapped conditional indirect effects of content credibility on brand resonance via brand trust")
    labs = ["Indirect effect at low involvement (\u22121 SD)", "Indirect effect at mean involvement",
            "Indirect effect at high involvement (+1 SD)", "Index of moderated mediation", "Direct effect (c\u2032)"]
    rows = [[labs[i], f"{bo[i][0]:.3f}", f"{bo[i][1]:.3f}", f"{bo[i][2]:.3f}", f"{bo[i][3]:.3f}",
             "Yes" if (bo[i][2] > 0 or bo[i][3] < 0) else "No"] for i in range(5)]
    table(d, ["Effect", "Estimate", "Boot SE", "LLCI", "ULCI", "CI excludes 0"], rows, [6.2, 2.0, 2.0, 2.0, 2.0, 3.0])
    source(d, "Authors\u2019 bootstrap analysis (percentile 95% confidence intervals).")
    para(d, [("Interpretation: ", {"bold": True}),
             (f"The indirect effect at the mean level of involvement is {bo[1][0]:.3f} (95% CI [{bo[1][2]:.3f}, {bo[1][3]:.3f}]), so H4 is "
              f"{hs(R, 3)}. The index of moderated mediation is {bo[3][0]:.3f} (95% CI [{bo[3][2]:.3f}, {bo[3][3]:.3f}]), so H6 is {hs(R, 5)}. "
              f"The conditional effect of credibility on trust is {R['slopes'][0][1]:.3f} at low, {R['slopes'][1][1]:.3f} at mean and "
              f"{R['slopes'][2][1]:.3f} at high involvement (Figure 1).", {})])
    d.add_picture(figs[1], width=Cm(12.5)); d.paragraphs[-1].alignment = C
    para(d, "Figure 1: Conditional effect of content credibility on brand trust by involvement level", size=10, bold=True, align=C, after=0)
    source(d, "Created by authors based on survey data analysis.")

    # Table 7 hypotheses
    para(d, "A summary of hypothesis testing is given in Table 7.")
    caption(d, "Table No: 7 Summary of hypothesis testing")
    table(d, ["Hypothesis", "Relationship", "Test statistic", "Decision"],
          [[h[0], h[1], h[2], "Supported" if h[3] else "Not supported"] for h in H], [1.8, 6.8, 5.8, 2.8])
    source(d, "Authors\u2019 analysis based on Tables 4\u20136.")

    # Table 8 ANOVA
    if R["anova"]:
        para(d, "To examine whether the constructs differ across UGC platforms, a one-way ANOVA was conducted. Table 8 shows the results.")
        caption(d, "Table No: 8 Platform-wise ANOVA for the study constructs")
        rows = []
        for (k, ssb, d1, msb, F, p, fc, ssw, d2, msw, sst, dft) in R["anova"]:
            rows += [[CONSTRUCTS[k][0], "Between groups", f"{ssb:.3f}", d1, f"{msb:.3f}", f"{F:.3f}", f"{p:.3f}", f"{fc:.3f}"],
                     ["", "Within groups", f"{ssw:.3f}", d2, f"{msw:.3f}", "", "", ""],
                     ["", "Total", f"{sst:.3f}", dft, "", "", "", ""]]
        table(d, ["Construct", "Source", "Sum of Squares", "df", "Mean Square", "F", "Sig. (p)", "F Crit"], rows,
              [3.0, 2.6, 2.4, 1.2, 2.2, 1.9, 1.9, 1.9])
        source(d, "Authors\u2019 one-way ANOVA with UGC platform as the independent variable.")
        sig_any = [CONSTRUCTS[r_[0]][0].lower() for r_ in R["anova"] if r_[5] < .05]
        para(d, [("Interpretation: ", {"bold": True}),
                 ("The ANOVA results show " + (f"statistically significant platform differences for {', '.join(sig_any)}. "
                  if sig_any else "no statistically significant platform differences in content credibility, brand trust or brand resonance "
                                  "(all p > 0.05), so the null hypothesis of equal means across platforms is retained. ")
                  + "This suggests that " + sup(not sig_any, "the process operates similarly across platforms.", "platform characteristics matter for UGC outcomes."), {})])
        if figs[0]:
            d.add_picture(figs[0], width=Cm(12.5)); d.paragraphs[-1].alignment = C
            para(d, "Figure 2: Mean brand resonance by UGC platform", size=10, bold=True, align=C, after=0)
            source(d, "Created by authors based on survey data analysis.")

    # Findings
    h1(d, "Findings")
    for t in [
        f"The study finds that content credibility is {sup(M1['p'][1] < .05, 'a significant', 'not a significant')} predictor of brand "
        f"resonance (B = {M1['b'][1]:.3f}, p {pf(M1['p'][1])}), so H1 is {hs(R, 0)}. Credibility also {sup(H[1][3], 'significantly increases', 'does not significantly increase')} "
        f"brand trust (H2 {hs(R, 1)}), and brand trust {sup(H[2][3], 'significantly increases', 'does not significantly increase')} resonance (H3 {hs(R, 2)}).",
        f"Brand trust {sup(H[3][3], 'mediates', 'does not mediate')} the credibility\u2013resonance relationship (indirect effect = {bo[1][0]:.3f}; "
        f"95% CI [{bo[1][2]:.3f}, {bo[1][3]:.3f}]; H4 {hs(R, 3)}). The direct effect of credibility after controlling for trust is {bo[4][0]:.3f}, "
        + sup(bo[4][2] > 0 or bo[4][3] < 0, "indicating partial mediation.", "indicating that the effect operates mainly through trust.") ,
        f"Consumer involvement {sup(H[4][3], 'significantly strengthens', 'does not significantly strengthen')} the credibility\u2013trust link (H5 {hs(R, 4)}), "
        f"and the moderated-mediation index is {sup(H[5][3], 'significant', 'not significant')} (H6 {hs(R, 5)}). The conditional indirect effect "
        f"ranges from {bo[0][0]:.3f} at low involvement to {bo[2][0]:.3f} at high involvement.",
        "Overall, UGC builds brand resonance when it is credible and when that credibility is converted into brand trust; the strength "
        "of the conversion depends on the consumer\u2019s level of involvement to the extent shown by the interaction results. "
        "Platform-wise differences are reported in Table 8."]:
        para(d, t)

    para(d, "Based on the results (Tables 4\u20138), Table 9 presents targeted recommendations for brand managers and platforms.")
    caption(d, "Table 9. Recommendations for Building Brand Resonance through User-Generated Content")
    recs = [
        ("1. Curate and Amplify Credible UGC", "Brands should prioritise verified, detailed and balanced customer content (e.g., verified-purchase tags, disclosure of incentives) because credibility is the starting point of the trust process."),
        ("2. Build Trust Beyond the Content", "Respond transparently to UGC, including negative content, and keep brand promises, so that credible content is converted into brand trust."),
        ("3. Segment by Consumer Involvement", "Use detailed, evidence-rich UGC (comparisons, long reviews) for highly involved consumers and simple, peer-based cues for low-involvement consumers."),
        ("4. Move from Content to Community", "Invite trusting consumers into brand communities, challenges and co-creation, which are the behavioural markers of resonance."),
        ("5. Platform-Specific Execution", "Adapt the UGC format to each platform while maintaining the same credibility standards across platforms."),
        ("6. Monitor Credibility and Trust Metrics", "Track perceived credibility, trust and resonance as separate indicators in brand-tracking studies rather than relying on engagement counts alone."),
        ("7. Platform and Regulatory Safeguards", "Platforms and regulators should enforce disclosure rules for sponsored or incentivised content to protect the credibility of UGC."),
    ]
    table(d, ["Particulars", "Recommendations"], [list(r_) for r_ in recs], [5.4, 11.8], size=9, left_cols=(0, 1))

    # ---------- two-column end ----------
    s3 = d.add_section(WD_SECTION.CONTINUOUS); s3.different_first_page_header_footer = False; set_cols(s3, 2)
    h1(d, "Conclusion")
    for t in [
        "This study examined when UGC builds brand resonance by testing the roles of content credibility, brand trust and consumer "
        f"involvement among {n} social media users. Credibility of UGC mattered for resonance mainly through brand trust, and the "
        f"conditional-process results show that {sup(H[5][3], 'the strength of this pathway rises with involvement.', 'the pathway does not differ significantly by involvement level in this sample.')}",
        "For practice, the findings imply that brands should treat credibility as an asset to be protected and trust as the bridge to "
        "resonance, rather than counting UGC volume. For theory, the study combines a mechanism (trust) and a boundary condition "
        "(involvement) within one model of brand resonance.",
        "Limitations include the cross-sectional design, self-reported measures and a single-country sample; causal claims should "
        "be avoided. Future studies may use longitudinal or experimental designs, other moderators such as brand familiarity or "
        "platform type, and behavioural outcome measures."]:
        para(d, t)
    h1(d, "Data Availability Statement")
    para(d, "The data that support the findings of this study are derived from the questionnaire survey described in Section 7. "
            "Processed data and analysis outputs are available from the corresponding author upon reasonable request. "
            "Data sharing is subject to ethical and confidentiality considerations, where applicable.")
    h1(d, "Conflict of Interest")
    para(d, "The author declares no conflict of interest regarding the publication of this research.")
    h1(d, "References")
    refs = [
        "Chaudhuri, A., & Holbrook, M. B. (2001). The chain of effects from brand trust and brand affect to brand performance: The role of brand loyalty. Journal of Marketing, 65(2), 81\u201393.",
        "Cheung, C. M. K., & Thadani, D. R. (2012). The impact of electronic word-of-mouth communication: A literature analysis and integrative model. Decision Support Systems, 54(1), 461\u2013470.",
        "Cronbach, L. J. (1951). Coefficient alpha and the internal structure of tests. Psychometrika, 16(3), 297\u2013334.",
        "Filieri, R. (2015). What makes online reviews helpful? A diagnosticity-adoption framework to explain informational and normative influences in e-WOM. Journal of Business Research, 68(6), 1261\u20131270.",
        "Hair, J. F., Black, W. C., Babin, B. J., & Anderson, R. E. (2019). Multivariate Data Analysis (8th ed.). Cengage Learning.",
        "Hayes, A. F. (2018). Introduction to Mediation, Moderation, and Conditional Process Analysis (2nd ed.). Guilford Press.",
        "Ismagilova, E., Slade, E., Rana, N. P., & Dwivedi, Y. K. (2020). The effect of characteristics of source credibility on consumer behaviour: A meta-analysis. Journal of Retailing and Consumer Services, 53, 101736.",
        "Keller, K. L. (2001). Building customer-based brand equity: A blueprint for creating strong brands. Marketing Science Institute Report No. 01-107.",
        "Morgan, R. M., & Hunt, S. D. (1994). The commitment-trust theory of relationship marketing. Journal of Marketing, 58(3), 20\u201338.",
        "Muntinga, D. G., Moorman, M., & Smit, E. G. (2011). Introducing COBRAs: Exploring motivations for brand-related social media use. International Journal of Advertising, 30(1), 13\u201346.",
        "Petty, R. E., Cacioppo, J. T., & Schumann, D. (1983). Central and peripheral routes to advertising effectiveness: The moderating role of involvement. Journal of Consumer Research, 10(2), 135\u2013146.",
        "Schivinski, B., & D\u0105browski, D. (2016). The effect of social media communication on consumer perceptions of brands. Journal of Marketing Communications, 22(2), 189\u2013214.",
        "Zaichkowsky, J. L. (1985). Measuring the involvement construct. Journal of Consumer Research, 12(3), 341\u2013352.",
    ]
    for i, r_ in enumerate(refs, 1):
        para(d, f"{i}. {r_}", size=9.5, after=2, align=L)

    d.save(out)
    return out


def profile_rows(df):
    rows = []
    for col, lab in (("Gender", "Gender"), ("AgeGroup", "Age group"), ("Platform", "Main UGC platform")):
        if col in df:
            vc = df[col].value_counts()
            for cat, cnt in vc.items():
                rows.append([lab if cat == vc.index[0] else "", cat, int(cnt), f"{cnt / len(df) * 100:.1f}"])
    rows.append(["Total", "", len(df), "100.0"])
    return rows


# =====================================================================
# 8. TELEGRAM
# =====================================================================
def get_secret(name, default=None):
    try:
        from google.colab import userdata
        v = userdata.get(name)
        if v:
            return v
    except Exception:
        pass
    v = os.environ.get(name) or default
    if v:
        return v
    return getpass.getpass(f"Enter {name}: ")


def tg_message(token, chat_id, text):
    r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                      data={"chat_id": chat_id, "text": text[:4000]}, timeout=60)
    print("sendMessage:", r.status_code, r.text[:200])
    return r.ok


def tg_document(token, chat_id, path, caption=""):
    with open(path, "rb") as f:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                          data={"chat_id": chat_id, "caption": caption[:1000]},
                          files={"document": (os.path.basename(path), f)}, timeout=180)
    print("sendDocument:", r.status_code, r.text[:200])
    return r.ok


# =====================================================================
# 9. MAIN
# =====================================================================
def main():
    df, simulated = load_data()
    R = analyse(df)
    R["profile_rows"] = profile_rows(df)
    figs = make_figures(R)
    out = build_doc(R, figs, simulated)
    print("Word file created:", out)

    summary = ["UGC & Brand Resonance - results"]
    summary.append(f"N = {R['n']}" + ("  (SIMULATED DEMO DATA)" if simulated else ""))
    for h in R["H"]:
        summary.append(f"{h[0]}: {'Supported' if h[3] else 'Not supported'} | {h[2]}")
    summary = "\n".join(summary)
    print(summary)

    if os.environ.get("SKIP_TELEGRAM") == "1":
        return
    token = get_secret("8837048625:AAFaq30r4xWFAKbUif-fXSjjuAZ54N6PQeU")
    chat_id = get_secret("TELEGRAM_CHAT_ID", default="1262412842")
    tg_message(token, chat_id, summary)
    tg_document(token, chat_id, out, caption="Sector-format article: UGC, Credibility, Involvement, Trust & Brand Resonance")
    for f in figs:
        if f:
            tg_document(token, chat_id, f, caption=f)


if __name__ == "__main__":
    main()
