#!/usr/bin/env python3
"""
build_site.py — готовит исходники сайта из rules/*.md.
Нарезает книги по главам (H2), красит модули по паспорту, генерирует mkdocs.yml.
Запуск:  py scripts/build_site.py   затем   mkdocs serve
"""
import re
import shutil
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
RULES = ROOT / "rules"
OUT = ROOT / "site-src"
FONTS = ROOT / "scripts" / "fonts"
REPO_URL = "https://github.com/Vireist/irreality"
SITE_URL = "https://vireist.github.io/irreality/"
DEFAULT_ACCENT = "#7A1E2B"

TR = str.maketrans({
    "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"e","ж":"zh","з":"z",
    "и":"i","й":"y","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r",
    "с":"s","т":"t","у":"u","ф":"f","х":"h","ц":"ts","ч":"ch","ш":"sh",
    "щ":"sch","ъ":"","ы":"y","ь":"","э":"e","ю":"yu","я":"ya",
})


def slugify(t):
    s = t.lower().translate(TR)
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-") or "stranica"


def mix(hx, o, k):
    h, o = hx.lstrip("#"), o.lstrip("#")
    a = [int(h[i:i+2], 16) for i in (0, 2, 4)]
    b = [int(o[i:i+2], 16) for i in (0, 2, 4)]
    return "#%02X%02X%02X" % tuple(round(x*k + y*(1-k)) for x, y in zip(a, b))


def parse_front_matter(text):
    m = re.match(r"\A---\s*\n(.*?)\n---\s*\n", text, re.S)
    if not m:
        return {}, text
    meta = {}
    for line in m.group(1).splitlines():
        line = line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        k, v = line.split(":", 1)
        meta[k.strip()] = v.strip().strip('"').strip("'")
    return meta, text[m.end():]


def split_head(text):
    lines = text.splitlines()
    title, subtitle = "Irreality", ""
    for i, ln in enumerate(lines):
        if ln.startswith("# "):
            title = ln[2:].strip()
            rest = lines[i+1:]
            for j, nxt in enumerate(rest):
                if nxt.strip():
                    if "**" in nxt or "·" in nxt:
                        subtitle = re.sub(r"\*\*", "", nxt).strip()
                        rest = rest[j+1:]
                    break
            return title, subtitle, "\n".join(rest).lstrip("\n")
    return title, subtitle, text


def split_chapters(body):
    """H2 вне блоков кода — границы глав."""
    chapters, preamble, cur, fence = [], [], None, False
    for line in body.split("\n"):
        if line.lstrip().startswith("```"):
            fence = not fence
        if (not fence) and line.startswith("## "):
            if cur is not None:
                chapters.append(cur)
            cur = [line[3:].strip(), []]
        elif cur is None:
            preamble.append(line)
        else:
            cur[1].append(line)
    if cur is not None:
        chapters.append(cur)
    return "\n".join(preamble).strip(), [(t, "\n".join(c).strip()) for t, c in chapters]


NOTE_KINDS = {"Важно!": "warning", "Внимание!": "warning",
              "Учтите!": "tip", "Совет": "tip", "Пример": "example"}


def convert_notes(md):
    """Цитаты «Важно!» → admonition-блоки Material."""
    lines = md.split("\n")
    out, i = [], 0
    while i < len(lines):
        m = re.match(r"^>\s*\*\*(" + "|".join(map(re.escape, NOTE_KINDS)) + r")\*\*\s*(.*)$", lines[i])
        if m:
            kind, first = m.group(1), m.group(2)
            block = [first] if first.strip() else []
            i += 1
            while i < len(lines) and lines[i].startswith(">"):
                block.append(re.sub(r"^>\s?", "", lines[i]))
                i += 1
            out.append(f'!!! {NOTE_KINDS[kind]} "{kind}"')
            out += ["    " + b for b in block]
            out.append("")
        else:
            out.append(lines[i])
            i += 1
    return "\n".join(out)


def fm(d):
    return "---\n" + yaml.safe_dump(d, allow_unicode=True, sort_keys=False) + "---\n\n"


def write_css_and_fonts():
    ss = OUT / "stylesheets"
    (ss / "fonts").mkdir(parents=True, exist_ok=True)
    files = ["PT_Serif-Web-Regular.ttf", "PT_Serif-Web-Bold.ttf",
             "PT_Serif-Web-Italic.ttf", "PT_Serif-Web-BoldItalic.ttf",
             "PT_Sans-Web-Regular.ttf", "PT_Sans-Web-Bold.ttf"]
    faces = []
    for f in files:
        src = FONTS / f
        if src.is_file():
            shutil.copy2(src, ss / "fonts" / f)
            fam = "PT Serif" if "Serif" in f else "PT Sans"
            weight = 700 if "Bold" in f else 400
            style = "italic" if "Italic" in f else "normal"
            faces.append(f"@font-face {{ font-family:'{fam}'; src:url('fonts/{f}');"
                         f" font-weight:{weight}; font-style:{style}; font-display:swap; }}")
        else:
            print(f"  ! нет шрифта {f} — сайт использует запасные")
    a = DEFAULT_ACCENT
    css = "\n".join(faces) + f"""
:root {{
  --md-text-font: 'PT Serif', Georgia, serif;
  --md-primary-fg-color: {a};
  --md-primary-fg-color--light: {mix(a, 'FFFFFF', 0.75)};
  --md-primary-fg-color--dark: {mix(a, '000000', 0.75)};
  --md-accent-fg-color: {a};
}}
.md-typeset {{ font-size: .82rem; line-height: 1.65; }}
.md-typeset h1, .md-typeset h2, .md-typeset h3, .md-typeset h4,
.md-header__title, .md-tabs__link {{
  font-family: 'PT Sans', 'Segoe UI', sans-serif;
}}
.md-typeset h1 {{ font-weight: 700; }}
.md-typeset h2 {{ border-bottom: 2px solid var(--md-primary-fg-color); padding-bottom: .2em; }}
.md-typeset table:not([class]) th {{
  background: var(--md-primary-fg-color--light);
  color: #fff; font-family: 'PT Sans', sans-serif;
}}
.md-typeset .md-button {{ margin: .2em .4em .2em 0; }}
"""
    (ss / "extra.css").write_text(css, encoding="utf-8")


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    books = []
    for p in sorted(RULES.glob("*.md")):
        meta, rest = parse_front_matter(p.read_text(encoding="utf-8"))
        title, subtitle, body = split_head(rest)
        slug = meta.get("slug") or p.stem
        pre, chapters = split_chapters(body)
        accent = meta.get("accent") or DEFAULT_ACCENT
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", accent):
            accent = DEFAULT_ACCENT
        colors = {"accent": accent,
                  "accent_light": mix(accent, "FFFFFF", 0.75),
                  "accent_dark": mix(accent, "000000", 0.75)}
        bdir = OUT / slug
        bdir.mkdir()
        pages = []
        for idx, (ct, cc) in enumerate(chapters, 1):
            fname = f"{idx:02d}-{slugify(ct)}.md"
            (bdir / fname).write_text(
                fm(dict(title=ct, **colors)) + "# " + ct + "\n\n" + convert_notes(cc) + "\n",
                encoding="utf-8")
            pages.append((ct, f"{slug}/{fname}"))
        idx_md = fm(dict(title=title, **colors)) + "# " + title + "\n\n" + (subtitle or "") + "\n\n"
        if meta.get("requires"):
            idx_md += ('!!! info "Совместимость"\n'
                       f'    Модуль работает с ядром **{meta["requires"]}**.\n\n')
        if pre:
            idx_md += convert_notes(pre) + "\n\n"
        idx_md += "## Главы\n\n"
        for ct, rel in pages:
            idx_md += f"- [{ct}]({rel.split('/', 1)[1]})\n"
        (bdir / "index.md").write_text(idx_md, encoding="utf-8")
        btype = meta.get("type") or "core"
        tab = meta.get("tab") or ("Ядро" if btype != "module" else title)
        books.append(dict(stem=slug, title=title, subtitle=subtitle,
                          type=btype, tab=tab, pages=pages, chapters=chapters))
        print(f"→ {p.name}: {len(pages)} глав [{accent}]")

    core_books = [b for b in books if b["type"] != "module"]
    mods = [b for b in books if b["type"] == "module"]

    # Главная: кнопки + «Ядро за одну страницу» из самой книги
    home = "# Irreality\n\nНастольная ролевая система о людях под давлением.\n\n"
    home += (f"[Скачать PDF и DOCX]({REPO_URL}/releases/latest)"
             "{ .md-button .md-button--primary } "
             f"[Исходники на GitHub]({REPO_URL})" "{ .md-button }\n\n")
    quick = None
    for b in core_books:
        for ct, cc in b["chapters"]:
            if "за одну страницу" in ct.lower():
                quick = (ct, cc)
    if quick:
        home += "## " + quick[0] + "\n\n" + convert_notes(quick[1]) + "\n"
    (OUT / "index.md").write_text(home, encoding="utf-8")

    (OUT / "modules").mkdir()
    mi = "# Модули\n\n"
    if mods:
        for b in mods:
            mi += f'- [{b["title"]}](../{b["stem"]}/index.md) — {b["subtitle"] or ""}\n'
    else:
        mi += "Модули сеттингов появятся здесь. Первый — «Изотопный Ярус».\n"
    (OUT / "modules" / "index.md").write_text(mi, encoding="utf-8")

    nav = [{"Главная": "index.md"}]
    for b in core_books:
        nav.append({b["tab"]: [b["stem"] + "/index.md"] + [{t: r} for t, r in b["pages"]]})
    if mods:
        sec = ["modules/index.md"]
        for b in mods:
            sec.append({b["title"]: [b["stem"] + "/index.md"] + [{t: r} for t, r in b["pages"]]})
        nav.append({"Модули": sec})
    else:
        nav.append({"Модули": "modules/index.md"})
    nav.append({"Скачать": REPO_URL + "/releases/latest"})

    write_css_and_fonts()
    cfg = dict(
        site_name="Irreality",
        site_url=SITE_URL,
        site_description="Настольная ролевая система: ядро и модули сеттингов",
        repo_url=REPO_URL,
        repo_name="Vireist/irreality",
        copyright="CC BY 4.0 · Vireist",
        docs_dir="site-src",
        site_dir="site",
        theme=dict(
            name="material", variant="classic", language="ru", custom_dir="overrides", font=False,
            palette=[
                dict(scheme="default", primary="custom", accent="custom",
                     toggle=dict(icon="material/brightness-7", name="Тёмная тема")),
                dict(scheme="slate", primary="custom", accent="custom",
                     toggle=dict(icon="material/brightness-4", name="Светлая тема")),
            ],
            features=["navigation.tabs", "navigation.sections", "navigation.indexes",
                      "navigation.top", "navigation.footer", "toc.follow",
                      "search.highlight", "search.suggest", "content.code.copy"],
        ),
        markdown_extensions=["admonition", "attr_list", "md_in_html", "tables",
                             "pymdownx.details", "pymdownx.superfences",
                             dict(toc=dict(permalink=True))],
        extra_css=["stylesheets/extra.css"],
        extra=dict(version=dict(provider="mike", default="latest", alias=True)),
        plugins=["search", {"mike": dict(alias_type="copy",
                                         canonical_version="latest")}],
        nav=nav,
    )
    (ROOT / "mkdocs.yml").write_text(
        yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False), encoding="utf-8")
    print("mkdocs.yml готов. Дальше: mkdocs serve")


if __name__ == "__main__":
    main()