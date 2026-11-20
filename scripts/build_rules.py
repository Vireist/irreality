#!/usr/bin/env python3
"""
build_rules.py — сборка красивых PDF и DOCX из Markdown-правил Irreality.

Зависимости:  pip install markdown weasyprint python-docx   (+ pandoc для DOCX)
Запуск:       python build_rules.py rules/*.md --outdir build --format a4
"""

import argparse
import glob
import re
import shutil
import subprocess
import sys
import urllib.request
from html import escape
from pathlib import Path

# ----------------------------------------------------------------------------- профили

PROFILES = {
    "a4": dict(page="A4", margin="19mm 17mm 18mm", body="10.5pt", table="9pt",
               code="8.6pt", h2="16pt", h3="12.5pt", cover="30pt", coverfoot="95mm",
               docx_body=11, docx_h2=16, docx_h3=13, docx_title=28,
               mm=(210, 297, 19, 17, 18)),
    "a5": dict(page="A5", margin="14mm 12mm 14mm", body="10pt", table="8.2pt",
               code="7.8pt", h2="14pt", h3="11.5pt", cover="24pt", coverfoot="52mm",
               docx_body=10, docx_h2=14.5, docx_h3=12, docx_title=24,
               mm=(148, 210, 14, 12, 14)),
}

INK, MUTED, LINE = "#26201E", "#6B625A", "#D8CFC4"


# --- Пути к внешним программам. Правьте, если положили не туда. ---
PANDOC_PATH = ""      # пусто → ищем в PATH
WEASYPRINT_PATH = ""      # пусто → ищем в scripts\weasyprint\ и в PATH


def find_exe(configured, name, local_subdir=None):
    """Ищем программу: указанный путь → папка проекта → PATH → типичные места."""
    import os
    if configured and Path(configured).is_file():
        return str(configured)
    if local_subdir:
        p = Path(__file__).resolve().parent / local_subdir / f"{name}.exe"
        if p.is_file():
            return str(p)
    found = shutil.which(name)
    if found:
        return found
    la = os.environ.get("LOCALAPPDATA", "")
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    cands = []
    if la:
        cands += [Path(la) / "Pandoc" / f"{name}.exe",
                  Path(la) / "Programs" / name.capitalize() / f"{name}.exe",
                  Path(la) / "Microsoft" / "WinGet" / "Links" / f"{name}.exe"]
        wg = Path(la) / "Microsoft" / "WinGet" / "Packages"
        if wg.is_dir():
            cands += list(wg.glob(f"*{name.capitalize()}*/**/{name}.exe"))
    cands += [Path(pf) / name.capitalize() / f"{name}.exe",
              Path(r"C:\Tools") / name / f"{name}.exe"]
    for c in cands:
        try:
            if c.is_file():
                return str(c)
        except OSError:
            pass
    return None

FONT_SPECS = [
    ("ptserif/PT_Serif-Web-Regular.ttf",    "PT Serif", 400, "normal"),
    ("ptserif/PT_Serif-Web-Bold.ttf",       "PT Serif", 700, "normal"),
    ("ptserif/PT_Serif-Web-Italic.ttf",     "PT Serif", 400, "italic"),
    ("ptserif/PT_Serif-Web-BoldItalic.ttf", "PT Serif", 700, "italic"),
    ("ptsans/PT_Sans-Web-Regular.ttf",      "PT Sans",  400, "normal"),
    ("ptsans/PT_Sans-Web-Bold.ttf",         "PT Sans",  700, "normal"),
    ("ptsans/PT_Sans-Web-Italic.ttf",       "PT Sans",  400, "italic"),
    ("ptsans/PT_Sans-Web-BoldItalic.ttf",   "PT Sans",  700, "italic"),
]
FONT_BASE = "https://raw.githubusercontent.com/google/fonts/main/ofl/"

# ----------------------------------------------------------------------------- CSS

CSS = """
@page {
  size: __PAGE__; margin: __MARGIN__;
  @top-left  { content: string(doctitle); font: 8pt 'PT Sans','Segoe UI',sans-serif;
               color: __MUTED__; letter-spacing: .4pt; }
  @top-right { content: string(chapter);  font: 8pt 'PT Sans','Segoe UI',sans-serif;
               color: __MUTED__; }
  @bottom-center { content: counter(page) " / " counter(pages);
               font: 8.5pt 'PT Sans','Segoe UI',sans-serif; color: __MUTED__; }
}
@page:first { @top-left{content:none} @top-right{content:none} @bottom-center{content:none} }

html { lang: ru; }
body { font-family: 'PT Serif', Georgia, 'DejaVu Serif', 'Times New Roman', serif;
       font-size: __BODY__; line-height: 1.42; color: __INK__; }
p { margin: 0 0 2.4mm; text-align: justify; hyphens: auto; widows: 2; orphans: 2; }
strong { font-weight: 700; }
a { color: inherit; text-decoration: none; }
ul, ol { margin: 0 0 2.5mm; padding-left: 5.5mm; }
li { margin: 0 0 1mm; }
img { max-width: 100%; }
hr { border: none; border-top: .6pt solid __LINE__; margin: 5mm 25%; }

/* обложка */
.cover { padding-top: 48mm; break-after: page; }
.cover .brand { font: 9pt 'PT Sans','Segoe UI',sans-serif; letter-spacing: 2.5pt;
                text-transform: uppercase; color: __MUTED__; }
.cover .title { font-family: 'PT Sans','Segoe UI',sans-serif; font-size: __COVER__;
                font-weight: 700; line-height: 1.12; margin: 6mm 0 0;
                string-set: doctitle content(); }
.cover .rule { width: 34mm; border-top: 2.5pt solid __ACCENT__; margin: 7mm 0; }
.cover .subtitle { font-size: 12pt; color: __MUTED__; text-align: left; }
.cover .cover-foot { margin-top: __COVERFOOT__; font: 9pt 'PT Sans',sans-serif;
                     color: __MUTED__; }

/* оглавление */
nav.toc { break-after: page; }
nav.toc h2 { font-family: 'PT Sans',sans-serif; font-size: __H2__; margin: 0 0 5mm;
             border-bottom: 2pt solid __ACCENT__; padding-bottom: 1.8mm; }
nav.toc ul { list-style: none; padding: 0; margin: 0; }
nav.toc li.l1 { margin: 2.4mm 0 0; font: 700 10pt 'PT Sans',sans-serif; }
nav.toc li.l2 { margin: .8mm 0 .8mm 6mm; font: 400 9.5pt 'PT Serif',serif; }
nav.toc a::after { content: leader('. ') target-counter(attr(href), page);
                   font-weight: 400; color: __MUTED__; }

/* главы и заголовки */
article h2 { break-before: page; string-set: chapter content();
             font-family: 'PT Sans',sans-serif; font-size: __H2__; font-weight: 700;
             margin: 0 0 4mm; padding-bottom: 1.8mm;
             border-bottom: 2pt solid __ACCENT__; }
article h3 { font-family: 'PT Sans',sans-serif; font-size: __H3__; font-weight: 700;
             margin: 6mm 0 2mm; break-after: avoid; }
article h4 { font-family: 'PT Sans',sans-serif; font-size: 10.5pt; font-weight: 700;
             color: __MUTED__; margin: 4mm 0 1.5mm; break-after: avoid; }

/* таблицы */
table { width: 100%; border-collapse: collapse; font-size: __TABLE__; margin: 3mm 0 4mm; }
thead { display: table-header-group; }
th { font-family: 'PT Sans',sans-serif; text-align: left; background: __SOFT__;
     border-bottom: 1.2pt solid __ACCENT__; padding: 1.6mm 2mm; }
td { padding: 1.4mm 2mm; border-bottom: .4pt solid __LINE__; vertical-align: top; }
th, td { overflow-wrap: break-word; }
tr { break-inside: avoid; }
tbody tr:nth-child(even) { background: __ZEBRA__; }

/* цитаты-принципы и боксы */
blockquote { margin: 3mm 0; padding: 2.5mm 3.5mm; background: __SOFT__;
             border-left: 2.5pt solid __ACCENT__; break-inside: avoid; }
blockquote p { text-align: left; margin: 0 0 1.5mm; }
blockquote p:last-child { margin: 0; }
.note { margin: 3mm 0; padding: 2.5mm 3.5mm; border-left: 2.5pt solid;
        break-inside: avoid; }
.note p { text-align: left; margin: 0 0 1.5mm; }
.note p:last-child { margin: 0; }
.note-warn    { border-color: __ACCENT__; background: __SOFT__; }
.note-tip     { border-color: #4A6B57; background: #EFF3EF; }
.note-example { border-color: #8A7A55; background: #F5F1E6; }

/* код и шпаргалки */
pre { font-family: 'DejaVu Sans Mono', Consolas, 'Courier New', monospace;
      font-size: __CODE__; line-height: 1.35; background: #F6F3EE;
      border: .5pt solid __LINE__; border-left: 2.5pt solid __ACCENT__;
      padding: 3mm; margin: 3mm 0; white-space: pre-wrap; word-wrap: break-word; }
code { font-family: 'DejaVu Sans Mono', Consolas, monospace; font-size: .92em;
       background: #F1EDE6; padding: 0 .6mm; }
pre code { background: none; padding: 0; font-size: inherit; }
"""

# ----------------------------------------------------------------------------- утилиты


def tint(hex_color: str, k: float) -> str:
    """Подмешать цвет к белому: k — доля цвета."""
    h = hex_color.lstrip("#")
    rgb = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    return "#%02X%02X%02X" % tuple(round(c * k + 255 * (1 - k)) for c in rgb)


def prepare_fonts(fdir: Path, download: bool) -> str:
    fdir.mkdir(parents=True, exist_ok=True)
    faces, missing = [], False
    for rel, family, weight, style in FONT_SPECS:
        fname = rel.split("/")[-1]
        p = fdir / fname
        if not p.exists() and download:
            try:
                with urllib.request.urlopen(FONT_BASE + rel, timeout=25) as r:
                    p.write_bytes(r.read())
            except Exception as e:
                print(f"  ! не скачался {fname}: {e}")
        if p.exists():
            faces.append(
                f"@font-face {{ font-family:'{family}'; src:url('{p.resolve().as_uri()}');"
                f" font-weight:{weight}; font-style:{style}; }}"
            )
        else:
            missing = True
    if missing:
        print("  ! часть шрифтов недоступна — использую системные (Georgia/Segoe UI/DejaVu)."
              "\n    Можно вручную положить TTF в", fdir)
    return "\n".join(faces)


def split_head(text: str):
    """Первый H1 → заголовок обложки; следующая строка → подзаголовок."""
    lines = text.splitlines()
    title, subtitle = "Irreality", ""
    for i, ln in enumerate(lines):
        if ln.startswith("# "):
            title = ln[2:].strip()
            rest = lines[i + 1:]
            for j, nxt in enumerate(rest):
                if nxt.strip():
                    if "**" in nxt or "·" in nxt:
                        subtitle = re.sub(r"\*\*", "", nxt).strip()
                        rest = rest[j + 1:]
                    break
            return title, subtitle, "\n".join(rest).lstrip("\n")
    return title, subtitle, text


def render_body(body_md: str):
    import markdown
    from markdown.extensions.toc import slugify_unicode

    md = markdown.Markdown(
        extensions=["tables", "fenced_code", "toc", "attr_list"],
        extension_configs={"toc": {"slugify": slugify_unicode, "toc_depth": "2-3"}},
    )
    html = md.convert(body_md)

    # «Важно!» / «Учтите!» → цветные боксы
    kinds = {"Важно!": "warn", "Внимание!": "warn", "Учтите!": "tip",
             "Совет": "tip", "Пример": "example"}
    pat = re.compile(
        r"<blockquote>\s*<p><strong>(" + "|".join(map(re.escape, kinds)) +
        r")\s*</strong>(.*?)</blockquote>", re.S)
    html = pat.sub(
        lambda m: f'<div class="note note-{kinds[m.group(1)]}">'
                  f'<p><strong>{m.group(1)}</strong>{m.group(2)}</div>', html)

    def walk(items, depth=1):
        if not items or depth > 2:
            return ""
        cls = "l1" if depth == 1 else "l2"
        out = ["<ul>"]
        for t in items:
            out.append(f'<li class="{cls}"><a href="#{t["id"]}">{escape(t["name"])}</a>')
            out.append(walk(t.get("children", []), depth + 1))
            out.append("</li>")
        out.append("</ul>")
        return "".join(out)

    return html, walk(md.toc_tokens)


def assemble(title, subtitle, body_html, toc_html, css):
    nav = f'<nav class="toc"><h2>Содержание</h2>{toc_html}</nav>' if toc_html else ""
    sub = f'<p class="subtitle">{escape(subtitle)}</p>' if subtitle else ""
    return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<title>{escape(title)}</title><style>{css}</style></head><body>
<section class="cover">
  <div class="brand">Настольная ролевая система</div>
  <h1 class="title">{escape(title)}</h1>
  <div class="rule"></div>{sub}
  <div class="cover-foot">Irreality · правила для игроков и мастера</div>
</section>
{nav}
<article>{body_html}</article>
</body></html>"""


def make_css(profile, accent, fonts_css):
    soft, zebra = tint(accent, 0.12), tint(accent, 0.05)
    css = CSS
    repl = {"__PAGE__": profile["page"], "__MARGIN__": profile["margin"],
            "__BODY__": profile["body"], "__TABLE__": profile["table"],
            "__CODE__": profile["code"], "__H2__": profile["h2"], "__H3__": profile["h3"],
            "__COVER__": profile["cover"], "__COVERFOOT__": profile["coverfoot"],
            "__ACCENT__": accent, "__SOFT__": soft, "__ZEBRA__": zebra,
            "__INK__": INK, "__MUTED__": MUTED, "__LINE__": LINE}
    for k, v in repl.items():
        css = css.replace(k, v)
    return fonts_css + "\n" + css


# ----------------------------------------------------------------------------- PDF


def build_pdf(full_html, md_path: Path, out_pdf: Path):
    exe = find_exe(WEASYPRINT_PATH, "weasyprint", "weasyprint")
    if exe:
        tmp = md_path.with_name(f"_{md_path.stem}_tmp.html")
        tmp.write_text(full_html, encoding="utf-8")
        try:
            subprocess.run([exe, str(tmp), str(out_pdf)], check=True)
        finally:
            tmp.unlink(missing_ok=True)
        return
    from weasyprint import HTML
    HTML(string=full_html, base_url=str(md_path.parent.resolve())).write_pdf(str(out_pdf))


# ----------------------------------------------------------------------------- DOCX


def make_reference(pandoc: str, work: Path, accent: str, profile: dict) -> Path:
    from docx import Document
    from docx.shared import Pt, Mm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn


    base = subprocess.run([pandoc, "--print-default-data-file", "reference.docx"],
                          capture_output=True, check=True).stdout
    bp = work / "_ref_base.docx"
    bp.write_bytes(base)
    doc = Document(str(bp))

    def S(name):
        try:
            return doc.styles[name]
        except KeyError:
            return None

    def font(st, name=None, size=None, bold=None, italic=None, color=None):
        if st is None:
            return
        f = st.font
        if name:
            f.name = name
            rpr = st.element.get_or_add_rPr()
            rpr.get_or_add_rFonts().set(qn("w:cs"), name)
        if size is not None:
            f.size = Pt(size)
        if bold is not None:
            f.bold = bold
        if italic is not None:
            f.italic = italic
        if color:
            f.color.rgb = RGBColor.from_string(color.lstrip("#"))

    def shade(pPr, fill):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), fill.lstrip("#"))
        pPr.append(shd)

    def field(p, instr, placeholder="1"):
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), instr)
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = placeholder
        r.append(t)
        fld.append(r)
        p._p.append(fld)

    acc = accent.lstrip("#")
    soft, zebra = tint(accent, 0.12).lstrip("#"), tint(accent, 0.05).lstrip("#")
    body = profile["docx_body"]

    st = S("Normal")
    font(st, "PT Serif", body)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "ru-RU")
    st.element.get_or_add_rPr().append(lang)

    heads = {"Heading 1": (profile["docx_h2"] + 3, acc, False),
             "Heading 2": (profile["docx_h2"], "26201E", True),
             "Heading 3": (profile["docx_h3"], "26201E", False),
             "Heading 4": (body + 0.5, "6B625A", False)}
    for name, (sz, col, pb) in heads.items():
        st = S(name)
        font(st, "PT Sans", sz, bold=True, color=col)
        if st is not None:
            st.paragraph_format.space_before = Pt(0 if pb else 12)
            st.paragraph_format.space_after = Pt(6)
            st.paragraph_format.page_break_before = pb
            st.paragraph_format.keep_with_next = True

    st = S("Title")
    font(st, "PT Sans", profile["docx_title"], bold=True, color=acc)
    if st is not None:
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.space_after = Pt(6)
    st = S("Subtitle")
    font(st, "PT Sans", 12, color="6B625A")
    if st is not None:
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.space_after = Pt(24)

    st = S("Block Text")
    font(st, size=body - 0.5, italic=True, color="3C3732")
    if st is not None:
        st.paragraph_format.left_indent = Mm(4)
        pPr = st.element.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        for k, v in (("val", "single"), ("sz", "12"), ("space", "6"), ("color", acc)):
            left.set(qn("w:" + k), v)
        bdr.append(left)
        pPr.append(bdr)
        shade(pPr, soft)

    st = S("Source Code")
    font(st, "Consolas", max(8, body - 2))
    if st is not None:
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.line_spacing = 1.0
        shade(st.element.get_or_add_pPr(), "F6F3EE")
    font(S("Verbatim Char"), "Consolas", body - 1)
    for cap in ("Caption", "Table Caption", "Image Caption"):
        font(S(cap), size=9, italic=True, color="6B625A")
    font(S("Header"), "PT Sans", 8.5, color="6B625A")
    font(S("Footer"), "PT Sans", 8.5, color="6B625A")

    st = S("Table")
    if st is not None:
        el = st.element
        tblPr = el.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            el.append(tblPr)
        for old in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(old)
        borders = OxmlElement("w:tblBorders")
        for side, sz, val, col in (("top", "4", "single", "D8CFC4"),
                                   ("bottom", "4", "single", "D8CFC4"),
                                   ("insideH", "4", "single", "D8CFC4"),
                                   ("left", "0", "none", "auto"),
                                   ("right", "0", "none", "auto"),
                                   ("insideV", "0", "none", "auto")):
            e = OxmlElement("w:" + side)
            e.set(qn("w:val"), val)
            e.set(qn("w:sz"), sz)
            e.set(qn("w:color"), col)
            borders.append(e)
        tblPr.append(borders)
        mar = OxmlElement("w:tblCellMar")
        for side, w in (("top", "30"), ("bottom", "30"), ("left", "80"), ("right", "80")):
            e = OxmlElement("w:" + side)
            e.set(qn("w:w"), w)
            e.set(qn("w:type"), "dxa")
            mar.append(e)
        tblPr.append(mar)

        def cond(name, fill=None, bold=False, color=None):
            sp = OxmlElement("w:tblStylePr")
            sp.set(qn("w:type"), name)
            if bold or color:
                rPr = OxmlElement("w:rPr")
                if bold:
                    rPr.append(OxmlElement("w:b"))
                if color:
                    c = OxmlElement("w:color")
                    c.set(qn("w:val"), color)
                    rPr.append(c)
                sp.append(rPr)
            if fill:
                tcPr = OxmlElement("w:tcPr")
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), fill)
                tcPr.append(shd)
                sp.append(tcPr)
            el.append(sp)

        cond("firstRow", fill=soft, bold=True, color=acc)
        cond("band1Horizontal", fill=zebra)

        sec = doc.sections[0]
    w, h, mt, mr, mb = profile["mm"]
    sec.page_width, sec.page_height = Mm(w), Mm(h)
    sec.top_margin, sec.bottom_margin = Mm(mt), Mm(mb)
    sec.left_margin, sec.right_margin = Mm(mr + 2), Mm(mr)

    # Титульная страница — без колонтитулов: там нет заголовка главы,
    # и STYLEREF выдавал бы ошибку Word.
    sec.different_first_page_header_footer = True
    sec.first_page_header.paragraphs[0].text = ""
    sec.first_page_footer.paragraphs[0].text = ""

    hp = sec.header.paragraphs[0]
    hp.text = ""
    field(hp, ' STYLEREF "Заголовок 2" \\* MERGEFORMAT ', "Irreality")
    fp = sec.footer.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run("— ")
    field(fp, " PAGE ")
    fp.add_run(" —")

    out = work / "reference.docx"
    doc.save(str(out))
    return out


def build_docx(pandoc, md_path, title, subtitle, body, out_docx, work, accent, profile):
    ref = make_reference(pandoc, work, accent, profile)
    meta = ('---\n'
            f'title: "{title.replace(chr(34), chr(92) + chr(34))}"\n'
            f'subtitle: "{subtitle.replace(chr(34), chr(92) + chr(34))}"\n'
            'lang: ru\n---\n\n')
    tmp = work / f"_tmp_{md_path.stem}.md"
    tmp.write_text(meta + body, encoding="utf-8")
    try:
        subprocess.run([pandoc, str(tmp), "-f", "markdown", "-t", "docx",
                        "--toc", "--toc-depth=3",
                        "--reference-doc", str(ref), "-o", str(out_docx)], check=True)
    finally:
        tmp.unlink(missing_ok=True)
        for junk in (ref, work / "_ref_base.docx"):
            junk.unlink(missing_ok=True)


# ----------------------------------------------------------------------------- main


def main():
    ap = argparse.ArgumentParser(description="Markdown → PDF + DOCX для правил Irreality")
    ap.add_argument("files", nargs="+", help=".md файлы (можно маски: rules/*.md)")
    ap.add_argument("--outdir", default="build")
    ap.add_argument("--format", choices=("a4", "a5"), default="a4")
    ap.add_argument("--accent", default="#7A1E2B", help="акцентный цвет, hex")
    ap.add_argument("--pdf", action="store_true", help="только PDF")
    ap.add_argument("--docx", action="store_true", help="только DOCX")
    ap.add_argument("--fonts-dir", default=None)
    ap.add_argument("--no-font-download", action="store_true")
    ap.add_argument("--keep-html", action="store_true", help="сохранить HTML для отладки")
    args = ap.parse_args()

    if not re.fullmatch(r"#[0-9A-Fa-f]{6}", args.accent):
        sys.exit("Цвет --accent должен быть вида #7A1E2B")
    do_pdf, do_docx = (args.pdf or not args.docx), (args.docx or not args.pdf)

    paths = []
    for pat in args.files:
        hits = sorted(glob.glob(pat))
        paths.extend(Path(h) for h in hits) if hits else paths.append(Path(pat))
    paths = [p for p in paths if p.suffix.lower() == ".md"]
    if not paths:
        sys.exit("Не найдено ни одного .md")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    profile = PROFILES[args.format]
    fdir = Path(args.fonts_dir) if args.fonts_dir else Path(__file__).resolve().parent / "fonts"
    fonts_css = prepare_fonts(fdir, not args.no_font_download) if do_pdf else ""
    css = make_css(profile, args.accent, fonts_css)

    pandoc = find_exe(PANDOC_PATH, "pandoc")
    if do_docx and not pandoc:
        print("! pandoc не найден — проверьте PANDOC_PATH в начале скрипта.")
        do_docx = False

    errors = 0
    for p in paths:
        if not p.exists():
            print(f"! нет файла {p}")
            errors += 1
            continue
        title, subtitle, body = split_head(p.read_text(encoding="utf-8"))
        ver = re.search(r"\d+\.\d+\.\d+", subtitle)
        stem = p.stem + (f"-{ver.group()}" if ver and ver.group() not in p.stem else "")
        stem = f"{stem}-{args.format}"
        print(f"→ {p.name}: «{title}»")

        if do_pdf:
            try:
                body_html, toc_html = render_body(body)
                full = assemble(title, subtitle, body_html, toc_html, css)
                if args.keep_html:
                    (outdir / f"{stem}.html").write_text(full, encoding="utf-8")
                build_pdf(full, p, outdir / f"{stem}.pdf")
                print(f"   PDF: {outdir / (stem + '.pdf')}")
            except ImportError:
                print("! pip install markdown weasyprint")
                errors += 1
            except Exception as e:
                print(f"! PDF не собрался: {e}")
                errors += 1
        if do_docx:
            try:
                build_docx(pandoc, p, title, subtitle, body,
                           outdir / f"{stem}.docx", outdir, args.accent, profile)
                print(f"   DOCX: {outdir / (stem + '.docx')}")
            except Exception as e:
                print(f"! DOCX не собрался: {e}")
                errors += 1

    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()