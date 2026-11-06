"""Rebuild the animations site from the Chapter 17 PowerPoint deck.

Usage:
    python update.py                  # uses the default deck path below
    python update.py path/to/deck.pptx

What it does:
    1. Finds every animated GIF in the deck (ppt/media) and the slide it is on
       (via the slide .rels files), in slide order.
    2. Names each GIF from its slide title (see NAMES and auto_name below) and
       writes it to gifs/<name>.gif. GIFs no longer in the deck are removed.
    3. Writes index.html with one section per GIF, anchored by its name.

The deck is copied to a temporary file first, so it can stay open in PowerPoint.
Only the last *saved* version is read.
"""

import html
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_DECK = HERE.parent.parent / "PPT" / "Dynamics_Chap17_PlanarKinematicsOfRigidBodies.pptx"
GIF_DIR = HERE / "gifs"
GA_ID = "G-M3V744VBMY"  # same Google Analytics property as the other teaching tools
PAGE_TITLE ="Dynamics Chapter 17: Planar Kinematics of Rigid Bodies – Animations"
PAGE_DESC = "Animations from the Chapter 17 lecture slides: rigid-body motion, instantaneous centers, rolling, and worked problems. By Dr. Nuha Aljuneidi."
SITE_URL = "https://nuhaaljuneidi.github.io/dynamics-chapter17-animations/"

# Hand-picked names, keyed by (slide title, picture name in PowerPoint).
# Any GIF not listed here gets an automatic name from auto_name().
NAMES = {
    ("Watch the Mass Center: Its Path Alone vs. Rotation Too", "Motion Animation"): "mass-center-path",
    ("One Machine, All Three Motions", "Picture 22"): "engine-three-motions",
    ("Three Problem Types: What Do You Know About B?", "Anim Surfaces"): "problem-types-surfaces",
    ("Three Problem Types: What Do You Know About B?", "Anim Fixed pin"): "problem-types-pin",
    ("Three Problem Types: What Do You Know About B?", "Anim Linkage"): "problem-types-linkage",
    ("The Instantaneous Center (IC)", "Anim IC rotation"): "ic-rotation",
    ("Locating the IC", "Anim Locate IC"): "locating-ic",
    ("Locating the IC: parallel velocities", "Anim Case 2"): "ic-parallel-case2",
    ("Locating the IC: parallel velocities", "Anim Case 3"): "ic-parallel-case3",
    ("Rolling Without Slipping", "Anim Rolling wheel"): "rolling-without-slipping",
    ("Rolling Wheels and the IC", "Anim Wheel IC"): "rolling-wheel-ic",
    ("Same Three Problem Types: Where Is the IC?", "Anim IC surfaces"): "ic-cases-surfaces",
    ("Same Three Problem Types: Where Is the IC?", "Anim IC rolling"): "ic-cases-rolling",
    ("Same Three Problem Types: Where Is the IC?", "Anim IC linkage"): "ic-cases-linkage",
}

STOPWORDS = {"a", "an", "and", "the", "of", "to", "vs", "is", "its", "it", "on", "in",
             "with", "what", "where", "do", "you", "about", "how", "too", "alone"}

REL = re.compile(r"<Relationship\b[^>]*/>")


def attr(tag, name):
    m = re.search(r'\b%s="([^"]*)"' % name, tag)
    return html.unescape(m.group(1)) if m else ""


def slug(text, max_words=4):
    words = [w for w in re.findall(r"[a-z0-9]+", text.lower()) if w not in STOPWORDS]
    return "-".join(words[:max_words]) or "animation"


def auto_name(title, pic_name, shared_slide):
    """Problem slides become problemN; others use the title's key words.
    When a slide holds several GIFs, the picture name is added as a suffix."""
    m = re.search(r"problem\s*(\d+)", title, re.I)
    base = f"problem{m.group(1)}" if m else slug(title)
    if shared_slide:
        extra = re.sub(r"^(anim(ation)?|picture|graphic)\s*", "", pic_name, flags=re.I)
        base += "-" + (slug(extra, 2) if extra and not extra.isdigit() else "part")
    return base


def slide_title(xml):
    """Text of the shape named 'Title' (or a title placeholder)."""
    for sp in re.findall(r"<p:sp>.*?</p:sp>", xml, re.S):
        name = attr(re.search(r"<p:cNvPr\b[^>]*>", sp).group(0), "name")
        if name.lower().startswith("title") or re.search(r'<p:ph\b[^>]*type="(title|ctrTitle)"', sp):
            paras = re.findall(r"<a:p>.*?</a:p>", sp, re.S)
            text = " ".join("".join(re.findall(r"<a:t>([^<]*)</a:t>", p)) for p in paras)
            return html.unescape(re.sub(r"\s+", " ", text)).strip()
    return ""


def read_deck(deck):
    """Return a list of dicts (slide, title, picture name, alt text, media path), in deck order."""
    z = zipfile.ZipFile(deck)
    rels = {attr(r, "Id"): attr(r, "Target")
            for r in REL.findall(z.read("ppt/_rels/presentation.xml.rels").decode("utf-8"))}
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    slides = [rels[rid] for rid in re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', pres)]

    found = []
    for number, target in enumerate(slides, 1):
        part = posixpath.normpath(posixpath.join("ppt", target))
        rel_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rel_part not in z.namelist():
            continue
        gif_rels = {attr(r, "Id"): posixpath.normpath(posixpath.join(posixpath.dirname(part), attr(r, "Target")))
                    for r in REL.findall(z.read(rel_part).decode("utf-8"))
                    if attr(r, "Target").lower().endswith(".gif")}
        if not gif_rels:
            continue
        xml = z.read(part).decode("utf-8")
        title = slide_title(xml) or f"Slide {number}"
        pics = []
        for pic in re.findall(r"<p:pic>.*?</p:pic>", xml, re.S):
            rid = re.search(r'r:embed="([^"]+)"', pic)
            if not rid or rid.group(1) not in gif_rels:
                continue
            nv = re.search(r"<p:cNvPr\b[^>]*>", pic).group(0)
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', pic)
            pos = (int(off.group(2)) // 500000, int(off.group(1))) if off else (0, 0)  # rows, then left to right
            pics.append((pos, attr(nv, "name"), attr(nv, "descr"), gif_rels[rid.group(1)]))
        pics.sort()
        for _, pic_name, descr, media in pics:
            found.append({"slide": number, "title": title, "pic": pic_name, "alt": descr, "media": media,
                          "shared": len(pics) > 1})
    return z, found


def assign_names(items):
    used = set()
    for it in items:
        name = NAMES.get((it["title"], it["pic"])) or auto_name(it["title"], it["pic"], it["shared"])
        base, n = name, 2
        while name in used:
            name, n = f"{base}-{n}", n + 1
        used.add(name)
        it["name"] = name


def case_label(it):
    if not it["shared"]:
        return ""
    label = re.sub(r"^(anim(ation)?|picture)\s*", "", it["pic"], flags=re.I).strip()
    return label[:1].upper() + label[1:] if label and not label.isdigit() else ""


def write_page(items):
    e = html.escape
    toc = "\n".join(
        f'        <li><a href="#{it["name"]}">{e(it["title"])}{" – " + e(case_label(it)) if case_label(it) else ""}</a></li>'
        for it in items)
    sections = "\n".join(f"""    <section id="{it['name']}">
      <h2>{e(it['title'])}</h2>
      <p class="meta">{e(case_label(it)) + ' · ' if case_label(it) else ''}<a href="#{it['name']}">#{it['name']}</a></p>
      <img src="gifs/{it['name']}.gif" alt="{e(it['alt'] or it['title'])}" loading="lazy">
    </section>""" for it in items)
    page = f"""<!doctype html>
<html lang="en">
<head>
  <!-- Google tag (gtag.js) -->
  <script async src="https://www.googletagmanager.com/gtag/js?id={GA_ID}"></script>
  <script>
    window.dataLayer = window.dataLayer || [];
    function gtag(){{dataLayer.push(arguments);}}
    gtag('js', new Date());
    gtag('config', '{GA_ID}');
  </script>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{e(PAGE_TITLE)}</title>
  <meta name="description" content="{e(PAGE_DESC)}">
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="Dr. Nuha Aljuneidi">
  <meta property="og:title" content="{e(PAGE_TITLE)}">
  <meta property="og:description" content="{e(PAGE_DESC)}">
  <meta property="og:url" content="{SITE_URL}">
  <meta name="twitter:card" content="summary">
  <style>
    :root {{ --navy: #10243E; --blue: #0070C0; --rule: #D5DCE5; }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: Arial, Helvetica, sans-serif; color: var(--navy); background: #fff; line-height: 1.5; }}
    header {{ background: var(--navy); color: #fff; padding: 28px 16px; }}
    header h1 {{ margin: 0 auto; max-width: 900px; font-size: 1.6rem; }}
    header p {{ margin: 8px auto 0; max-width: 900px; color: #C9D6E6; }}
    header .author {{ display: block; color: #fff; font-size: 1.9rem; font-weight: bold; }}
    main {{ max-width: 900px; margin: 0 auto; padding: 24px 16px 48px; }}
    nav h2 {{ font-size: 1.1rem; margin: 0 0 8px; }}
    nav ol {{ margin: 0 0 8px; padding-left: 22px; }}
    nav li {{ margin: 4px 0; }}
    a {{ color: var(--blue); }}
    section {{ border-top: 1px solid var(--rule); padding-top: 24px; margin-top: 28px; scroll-margin-top: 12px; }}
    section h2 {{ color: var(--navy); font-size: 1.3rem; margin: 0; border-left: 4px solid var(--blue); padding-left: 10px; }}
    .meta {{ margin: 6px 0 14px; color: #5B6B80; font-size: .9rem; }}
    .meta a {{ text-decoration: none; }}
    img {{ display: block; max-width: 100%; height: auto; border: 1px solid var(--rule); }}
    footer {{ max-width: 900px; margin: 0 auto; padding: 0 16px 32px; color: #5B6B80; font-size: .85rem; }}
  </style>
</head>
<body>
  <header>
    <h1>{e(PAGE_TITLE)}</h1>
    <p><span class="author">Dr. Nuha Aljuneidi</span>Dynamics</p>
  </header>
  <main>
    <nav>
      <h2>Animations</h2>
      <ol>
{toc}
      </ol>
    </nav>
{sections}
  </main>
  <footer>{len(items)} animations from the Chapter 17 lecture slides.<br>&copy; {date.today().year} Dr. Nuha Aljuneidi. All rights reserved.</footer>
</body>
</html>
"""
    (HERE / "index.html").write_text(page, encoding="utf-8")


def copy_deck(src, dst):
    """Copy the deck even while PowerPoint has it open (Python's open() is refused then,
    but the Windows CopyFile call is allowed)."""
    try:
        shutil.copyfile(src, dst)
    except PermissionError:
        import ctypes
        if not ctypes.windll.kernel32.CopyFileW(str(src), str(dst), False):
            raise PermissionError(f"Could not read {src}. Save and close it in PowerPoint, then run again.")


def main():
    deck = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_DECK
    with tempfile.TemporaryDirectory() as tmp:
        copy = Path(tmp) / "deck.pptx"
        copy_deck(deck, copy)
        z, items = read_deck(copy)
        assign_names(items)
        GIF_DIR.mkdir(exist_ok=True)
        keep = set()
        for it in items:
            out = GIF_DIR / f"{it['name']}.gif"
            out.write_bytes(z.read(it["media"]))
            keep.add(out.name)
        z.close()
    for old in GIF_DIR.glob("*.gif"):
        if old.name not in keep:
            old.unlink()
    write_page(items)
    print(f"{len(items)} GIFs from {deck.name}")
    for it in items:
        print(f"  slide {it['slide']:>3}  {it['name'] + '.gif':32s} {it['title']}")


if __name__ == "__main__":
    main()
