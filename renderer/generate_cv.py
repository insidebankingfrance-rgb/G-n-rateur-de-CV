"""Render a normalized CV JSON to an Inside Circle .pptx (FR + EN slides).

Usage:
    python renderer/generate_cv.py <input.json> [--out outputs/<name>.pptx]
    python renderer/generate_cv.py data/ --out outputs/    # batch mode
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Iterable

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

from theme import (
    BLUE_ROYAL,
    CYAN_ACCENT,
    CYAN_SOFT,
    FONT_FAMILY,
    FOOTER_H,
    FOOTER_TEXT,
    FOOTER_TEXT_EN,
    FS_BODY,
    FS_BODY_SMALL,
    FS_DOMAIN,
    FS_FOOTER,
    FS_INITIALS,
    FS_SECTION,
    FS_TITLE,
    HEADER_H,
    MARGIN,
    NAVY_DEEP,
    NAVY_MID,
    SIDEBAR_W,
    SLIDE_H,
    SLIDE_W,
    WHITE,
    WHITE_SOFT,
)
from i18n import LABELS


# ─── Low-level helpers ───────────────────────────────────────────────────────
def _set_fill(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def _add_rect(slide, left, top, width, height, color: RGBColor):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    _set_fill(shp, color)
    shp.shadow.inherit = False
    return shp


def _add_textbox(slide, left, top, width, height):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    # Clear default paragraph
    tf.paragraphs[0].text = ""
    return tb, tf


def _run(paragraph, text: str, *, bold=False, italic=False,
         size=FS_BODY, color=WHITE, font=FONT_FAMILY):
    r = paragraph.add_run()
    r.text = text
    f = r.font
    f.name = font
    f.size = size
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    return r


def _new_para(tf, *, align=PP_ALIGN.LEFT, space_after=Pt(2), level=0):
    """Add a new paragraph (skipping the empty default one on first call)."""
    if not tf.paragraphs[0].runs and tf.paragraphs[0].text == "":
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.space_after = space_after
    p.level = level
    return p


def _section_title(tf, label: str):
    p = _new_para(tf, space_after=Pt(4))
    _run(p, label, bold=True, size=FS_SECTION, color=CYAN_ACCENT)
    # underline-like accent: small cyan bar below — implemented as next paragraph
    return p


def _bullet(tf, text: str, *, size=FS_BODY, color=WHITE, level=0,
            bold=False, italic=False):
    p = _new_para(tf, space_after=Pt(2), level=level)
    _run(p, "•  ", bold=True, size=size, color=CYAN_ACCENT)
    _run(p, text, size=size, color=color, bold=bold, italic=italic)
    return p


# ─── Anonymization helper ────────────────────────────────────────────────────
def initials_from_name(first: str, last: str) -> str:
    f = (first or "").strip()
    l = (last or "").strip()
    if not f and not l:
        return "?.?."
    return f"{f[:1].upper()}.{l[:1].upper()}."


# ─── Slide builders ──────────────────────────────────────────────────────────
def _build_background(slide):
    """Solid navy background + cyan accent in bottom-right corner."""
    # Full background
    bg = _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY_DEEP)
    # Diagonal accent — large soft cyan oval in bottom-right
    accent = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        SLIDE_W - Emu(3500000), SLIDE_H - Emu(2800000),
        Emu(5500000), Emu(5500000),
    )
    _set_fill(accent, CYAN_SOFT)
    accent.fill.transparency = 0  # python-pptx ignores; visual effect from color
    accent.line.fill.background()
    # Layer a navy oval slightly offset to fake a soft gradient
    soft = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        SLIDE_W - Emu(4500000), SLIDE_H - Emu(3800000),
        Emu(5500000), Emu(5500000),
    )
    _set_fill(soft, NAVY_MID)
    soft.line.fill.background()


def _build_sidebar(slide):
    """Sidebar panel — left column, semi-distinct fill."""
    _add_rect(slide, 0, 0, SIDEBAR_W, SLIDE_H, NAVY_MID)


def _build_header(slide, initials: str, title: str, domain: str):
    """Initials (top-left) + title and domain (right of initials)."""
    # Initials block — inside the sidebar
    tb, tf = _add_textbox(
        slide,
        MARGIN, Emu(280000),
        SIDEBAR_W - 2 * MARGIN, Emu(900000),
    )
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _run(p, initials, bold=True, size=FS_INITIALS, color=CYAN_ACCENT)

    # Title + domain — to the right of the sidebar
    tb, tf = _add_textbox(
        slide,
        SIDEBAR_W + MARGIN, Emu(300000),
        SLIDE_W - SIDEBAR_W - 2 * MARGIN, Emu(900000),
    )
    p = tf.paragraphs[0]
    _run(p, title or "", bold=True, size=FS_TITLE, color=WHITE)
    if domain:
        p2 = tf.add_paragraph()
        p2.space_before = Pt(2)
        _run(p2, domain, size=FS_DOMAIN, color=CYAN_ACCENT, italic=True)


def _build_footer(slide, lang: str):
    text = FOOTER_TEXT if lang == "fr" else FOOTER_TEXT_EN
    tb, tf = _add_textbox(
        slide,
        MARGIN, SLIDE_H - FOOTER_H,
        SLIDE_W - 2 * MARGIN, FOOTER_H,
    )
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, text, italic=True, size=FS_FOOTER, color=WHITE_SOFT)


def _build_sidebar_content(slide, cv: dict, lang: str):
    L = LABELS[lang]
    top = HEADER_H + Emu(100000)
    height = SLIDE_H - top - FOOTER_H - Emu(100000)
    tb, tf = _add_textbox(
        slide,
        MARGIN, top,
        SIDEBAR_W - 2 * MARGIN, height,
    )
    tf.word_wrap = True
    first = True

    def section(label, items_renderer):
        nonlocal first
        if first:
            _section_title(tf, label)
            first = False
        else:
            # add some space before next section
            spacer = _new_para(tf, space_after=Pt(2))
            _run(spacer, " ", size=Pt(4), color=NAVY_MID)
            _section_title(tf, label)
        items_renderer()

    # Expertise
    expertise = cv.get("expertise", [])
    if expertise:
        def _render():
            for e in expertise:
                _bullet(tf, e, size=FS_BODY)
        section(L["expertise"], _render)

    # Languages
    languages = cv.get("languages", [])
    if languages:
        def _render():
            for lng in languages:
                name = lng.get("name", "") if isinstance(lng, dict) else str(lng)
                level = lng.get("level", "") if isinstance(lng, dict) else ""
                txt = f"{name} — {level}" if level else name
                _bullet(tf, txt, size=FS_BODY)
        section(L["languages"], _render)

    # Education & certifications (merged)
    edu = cv.get("education", [])
    if edu:
        def _render():
            for item in edu:
                if isinstance(item, dict):
                    year = item.get("year", "")
                    school = item.get("school", "")
                    degree = item.get("degree", "")
                    p = _new_para(tf, space_after=Pt(1))
                    if year:
                        _run(p, f"{year}  ", bold=True, size=FS_BODY_SMALL,
                             color=CYAN_ACCENT)
                    if school:
                        _run(p, school, bold=True, size=FS_BODY_SMALL,
                             color=WHITE)
                    if degree:
                        p2 = _new_para(tf, space_after=Pt(3))
                        _run(p2, degree, size=FS_BODY_SMALL, color=WHITE_SOFT,
                             italic=True)
                else:
                    _bullet(tf, str(item), size=FS_BODY_SMALL)
        section(L["education"], _render)


def _build_main_content(slide, cv: dict, lang: str):
    L = LABELS[lang]
    left = SIDEBAR_W + MARGIN
    top = HEADER_H + Emu(100000)
    width = SLIDE_W - SIDEBAR_W - 2 * MARGIN
    height = SLIDE_H - top - FOOTER_H - Emu(100000)
    tb, tf = _add_textbox(slide, left, top, width, height)
    tf.word_wrap = True
    first = True

    def section(label, items_renderer):
        nonlocal first
        if first:
            _section_title(tf, label)
            first = False
        else:
            spacer = _new_para(tf, space_after=Pt(2))
            _run(spacer, " ", size=Pt(4), color=NAVY_DEEP)
            _section_title(tf, label)
        items_renderer()

    # Summary
    summary = cv.get("summary", "").strip()
    if summary:
        def _render():
            p = _new_para(tf, space_after=Pt(3))
            _run(p, summary, size=FS_BODY, color=WHITE)
        section(L["summary"], _render)

    # Experience
    experiences = cv.get("experiences", [])
    if experiences:
        def _render():
            for exp in experiences:
                # Header line: employer — role  + duration (right side, but
                # python-pptx tabs are flaky → put duration in italic at end).
                employer = exp.get("employer", "")
                role = exp.get("role", "")
                duration = exp.get("duration", "")
                p = _new_para(tf, space_after=Pt(1))
                _run(p, "▸ ", bold=True, size=FS_BODY, color=CYAN_ACCENT)
                if employer:
                    _run(p, employer, bold=True, size=FS_BODY, color=WHITE)
                if role:
                    _run(p, f" — {role}", size=FS_BODY, color=WHITE)
                if duration:
                    _run(p, f"  ({duration})", italic=True,
                         size=FS_BODY_SMALL, color=CYAN_ACCENT)
                # Achievements / bullets
                for bul in exp.get("achievements", []):
                    _bullet(tf, bul, size=FS_BODY_SMALL, level=1)
        section(L["experience"], _render)

    # References
    refs = cv.get("references", [])
    if refs:
        def _render():
            text = ", ".join(refs)
            p = _new_para(tf, space_after=Pt(2))
            _run(p, text, size=FS_BODY, color=WHITE, italic=True)
        section(L["references"], _render)

    # Engagements & extras
    engagements = cv.get("engagements", [])
    if engagements:
        def _render():
            for e in engagements:
                if isinstance(e, dict):
                    title = e.get("title", "")
                    desc = e.get("description", "")
                    p = _new_para(tf, space_after=Pt(1))
                    _run(p, "• ", bold=True, size=FS_BODY_SMALL,
                         color=CYAN_ACCENT)
                    if title:
                        _run(p, title, bold=True, size=FS_BODY_SMALL,
                             color=WHITE)
                    if desc:
                        _run(p, f" — {desc}", size=FS_BODY_SMALL,
                             color=WHITE_SOFT)
                else:
                    _bullet(tf, str(e), size=FS_BODY_SMALL)
        section(L["engagements"], _render)

    # Hobbies
    hobbies = cv.get("hobbies", [])
    if hobbies:
        def _render():
            txt = "  •  ".join(hobbies)
            p = _new_para(tf, space_after=Pt(1))
            _run(p, txt, size=FS_BODY_SMALL, color=WHITE)
        section(L["hobbies"], _render)


# ─── Top-level render ────────────────────────────────────────────────────────
def _build_slide(prs: Presentation, cv: dict, lang: str):
    blank = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(blank)
    _build_background(slide)
    _build_sidebar(slide)

    initials = cv.get("initials") or initials_from_name(
        cv.get("first_name", ""), cv.get("last_name", "")
    )
    title = cv.get(f"title_{lang}") or cv.get("title", "")
    domain = cv.get(f"domain_{lang}") or cv.get("domain", "")
    _build_header(slide, initials, title, domain)

    payload = cv.get(lang, cv)  # if cv has 'fr'/'en' keys, use lang-specific
    _build_sidebar_content(slide, payload, lang)
    _build_main_content(slide, payload, lang)

    _build_footer(slide, lang)


def render(cv_json: dict, out_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    _build_slide(prs, cv_json, "fr")
    _build_slide(prs, cv_json, "en")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path


def _slug(s: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9._-]+", "_", s.strip())
    return s.strip("_") or "cv"


def _iter_jsons(path: Path) -> Iterable[Path]:
    if path.is_dir():
        return sorted(p for p in path.glob("*.json"))
    return [path]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path, help="JSON file or directory")
    parser.add_argument("--out", type=Path, default=Path("outputs"))
    args = parser.parse_args()

    out = args.out
    multiple = args.input.is_dir() or out.is_dir() or out.suffix != ".pptx"
    if multiple and out.suffix == ".pptx":
        raise SystemExit("In batch mode, --out must be a directory")

    for j in _iter_jsons(args.input):
        cv = json.loads(j.read_text(encoding="utf-8"))
        if multiple:
            out_file = (out if out.is_dir() else Path("outputs")) / f"{_slug(j.stem)}.pptx"
        else:
            out_file = out
        rendered = render(cv, out_file)
        print(f"✓ {j.name} → {rendered}")


if __name__ == "__main__":
    main()
