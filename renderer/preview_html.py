"""Render a CV JSON to a single HTML page showing the FR + EN slides.

Used to validate the visual layout without opening PowerPoint. The HTML is a
faithful facsimile of what the .pptx slides look like (same colors, same
fonts, same truncation rules).

Usage:
    python renderer/preview_html.py <input.json> --out preview.html
"""

from __future__ import annotations

import argparse
import base64
import html
import json
from pathlib import Path

from layout import paginate, source_lang
from theme import (
    LOGO_FILENAME,
    MAX_BULLETS_PER_EXP,
    MAX_ENGAGEMENTS,
    MAX_EXPERIENCES,
    MAX_EXPERTISE,
    MAX_HOBBIES,
)


# ─── Tokens (mirror theme.py) ────────────────────────────────────────────────
NAVY_SIDEBAR = "#061A5E"
CYAN_ACC    = "#3AEDE5"
WHITE       = "#FFFFFF"
WHITE_SOFT  = "#CCD6E8"


SECTION_LABELS = {
    "fr": {
        "summary":      "RÉSUMÉ",
        "expertise":    "DOMAINES D'EXPERTISE",
        "languages":    "LANGUES",
        "education":    "FORMATION & CERTIFICATIONS",
        "hobbies":      "CENTRES D'INTÉRÊT",
        "experience":   "PRINCIPALES EXPÉRIENCES PROFESSIONNELLES",
        "references":   "RÉFÉRENCES",
        "engagements":  "ENGAGEMENTS & RÉALISATIONS",
        "footer":       "Les informations de ce document sont strictement confidentielles — Ce document ne peut être partagé qu'avec l'accord de son propriétaire",
        "lang_tag":     "Slide {n} — Français",
        "continued":    "(suite)",
    },
    "en": {
        "summary":      "SUMMARY",
        "expertise":    "MAIN AREAS OF EXPERTISE",
        "languages":    "LANGUAGES",
        "education":    "EDUCATION & CERTIFICATIONS",
        "hobbies":      "INTERESTS",
        "experience":   "MAIN PROFESSIONAL EXPERIENCE",
        "references":   "MAIN REFERENCES",
        "engagements":  "ENGAGEMENTS & ACHIEVEMENTS",
        "footer":       "The information in this document is strictly confidential — This document may only be shared with the owner's consent",
        "lang_tag":     "Slide {n} — English",
        "continued":    "(cont.)",
    },
}


def _esc(s) -> str:
    return html.escape(str(s) if s is not None else "")


def _estimate_sidebar_pt_html(payload: dict) -> float:
    """Lightweight mirror of the renderer's sidebar height estimator."""
    line_pt, sec_pt, gap_pt, cpl = 14, 19, 1, 42
    def wrap(s):
        return max(1, (len(s) + cpl - 1) // cpl) if s else 0
    pt = 0.0
    if payload.get("expertise"):
        pt += sec_pt + len(payload["expertise"]) * line_pt
    if payload.get("languages"):
        pt += sec_pt + len(payload["languages"]) * line_pt
    if payload.get("education"):
        pt += sec_pt
        for e in payload["education"]:
            year = (e.get("year", "") if isinstance(e, dict) else "")
            school = (e.get("school", "") if isinstance(e, dict) else str(e))
            degree = (e.get("degree", "") if isinstance(e, dict) else "")
            pt += wrap(f"{year} {school}") * line_pt + wrap(degree) * line_pt
    if payload.get("hobbies"):
        pt += sec_pt + sum(wrap(h) for h in payload["hobbies"]) * line_pt
    if payload.get("engagements"):
        pt += sec_pt
        for e in payload["engagements"]:
            if isinstance(e, dict):
                line = f"{e.get('title','')} {e.get('description','')}"
            else:
                line = str(e)
            pt += wrap(line) * line_pt + gap_pt
    return pt


_AVAIL_SIDEBAR_HTML = 409
_AVAIL_MAIN_HTML = 419


def _estimate_main_pt_html(payload: dict) -> float:
    """Lightweight mirror of the renderer's main-column height estimator."""
    line_pt, sec_pt, gap_pt, cpl = 14, 19, 1, 80
    def wrap(s):
        return max(1, (len(s) + cpl - 1) // cpl) if s else 0
    pt = 0.0
    if (payload.get("summary") or "").strip():
        pt += sec_pt + wrap(payload["summary"]) * line_pt
    if payload.get("experiences"):
        pt += sec_pt
        for exp in payload["experiences"]:
            header = (f"{exp.get('employer','')} {exp.get('role','')} "
                      f"{exp.get('duration','')}")
            pt += wrap(header) * line_pt + gap_pt
            for b in exp.get("achievements", []):
                pt += wrap(b) * line_pt + gap_pt
    if payload.get("references"):
        ref = ", ".join(payload["references"])
        pt += sec_pt + wrap(ref) * line_pt
    return pt


def _truncate(payload: dict) -> dict:
    """Mirror the renderer's MAX_* caps so preview = output."""
    out = dict(payload)
    exp = out.get("experiences", [])[:MAX_EXPERIENCES]
    exp = [dict(e) for e in exp]
    for e in exp:
        if len(e.get("achievements", [])) > MAX_BULLETS_PER_EXP:
            e["achievements"] = e["achievements"][:MAX_BULLETS_PER_EXP]
    out["experiences"] = exp
    for key, cap in (("expertise", MAX_EXPERTISE),
                     ("hobbies", MAX_HOBBIES),
                     ("engagements", MAX_ENGAGEMENTS)):
        if key in out:
            out[key] = out[key][:cap]

    # Adaptive sidebar trim: mirror the renderer's eviction order.
    if _estimate_sidebar_pt_html(out) > _AVAIL_SIDEBAR_HTML and out.get("hobbies"):
        out["hobbies"] = []
    if _estimate_sidebar_pt_html(out) > _AVAIL_SIDEBAR_HTML and out.get("engagements"):
        out["engagements"] = [
            ({**e, "description": ""} if isinstance(e, dict) else e)
            for e in out["engagements"]
        ]
    if _estimate_sidebar_pt_html(out) > _AVAIL_SIDEBAR_HTML and out.get("engagements"):
        out["engagements"] = out["engagements"][:2]

    # Adaptive main trim: drop references when main is near/over budget
    # (marge de sécurité de 20pt).
    if _estimate_main_pt_html(out) > _AVAIL_MAIN_HTML - 20 and out.get("references"):
        out["references"] = []
    return out


def _section_title(label: str) -> str:
    return f'<h3 class="section">{_esc(label)}</h3>'


def _bullet_list(items, *, sub=False) -> str:
    items_html = "".join(
        f'<li class="bullet"><span class="dot">•</span><span>{_esc(it)}</span></li>'
        for it in items
    )
    cls = "bullet-list sub" if sub else "bullet-list"
    return f'<ul class="{cls}">{items_html}</ul>'


def _render_sidebar(payload, labels) -> str:
    parts = []

    if payload.get("expertise"):
        parts.append(_section_title(labels["expertise"]))
        parts.append(_bullet_list(payload["expertise"]))

    if payload.get("languages"):
        parts.append(_section_title(labels["languages"]))
        items = []
        for lng in payload["languages"]:
            if isinstance(lng, dict):
                name = lng.get("name", "")
                level = lng.get("level", "")
                items.append(f"{name} — {level}" if level else name)
            else:
                items.append(str(lng))
        parts.append(_bullet_list(items))

    if payload.get("education"):
        parts.append(_section_title(labels["education"]))
        rows = []
        for e in payload["education"]:
            if isinstance(e, dict):
                year = _esc(e.get("year", ""))
                school = _esc(e.get("school", ""))
                degree = _esc(e.get("degree", ""))
                rows.append(
                    f'<div class="edu-row">'
                    f'<div class="edu-line">'
                    + (f'<span class="edu-year">{year}</span> ' if year else "")
                    + f'<span class="edu-school">{school}</span>'
                    + '</div>'
                    + (f'<div class="edu-degree">{degree}</div>' if degree else "")
                    + '</div>'
                )
            else:
                rows.append(f'<div class="edu-row"><div class="edu-school">{_esc(e)}</div></div>')
        parts.append(f'<div class="edu-list">{"".join(rows)}</div>')

    if payload.get("hobbies"):
        parts.append(_section_title(labels["hobbies"]))
        parts.append(_bullet_list(payload["hobbies"]))

    if payload.get("engagements"):
        parts.append(_section_title(labels["engagements"]))
        rows = []
        for e in payload["engagements"]:
            if isinstance(e, dict):
                t = _esc(e.get("title", ""))
                d = _esc(e.get("description", ""))
                rows.append(
                    f'<li class="bullet"><span class="dot">•</span>'
                    f'<span><strong>{t}</strong>'
                    + (f' — <em class="muted">{d}</em>' if d else "")
                    + "</span></li>"
                )
            else:
                rows.append(f'<li class="bullet"><span class="dot">•</span><span>{_esc(e)}</span></li>')
        parts.append(f'<ul class="bullet-list">{"".join(rows)}</ul>')

    return "\n".join(parts)


def _render_main(payload, labels) -> str:
    parts = []

    summary = (payload.get("summary") or "").strip()
    if summary:
        parts.append(_section_title(labels["summary"]))
        parts.append(f'<p class="summary">{_esc(summary)}</p>')

    if payload.get("experiences"):
        exp_label = labels["experience"]
        if payload.get("continued"):
            exp_label = f'{exp_label} {labels["continued"]}'
        parts.append(_section_title(exp_label))
        exp_html = []
        for exp in payload["experiences"]:
            employer = _esc(exp.get("employer", ""))
            role = _esc(exp.get("role", ""))
            duration = _esc(exp.get("duration", ""))
            header = (
                f'<div class="exp-header">'
                f'<span class="exp-marker">▸</span>'
                f'<span class="exp-employer">{employer}</span>'
                + (f' <span class="exp-sep">—</span> <span class="exp-role">{role}</span>' if role else "")
                + (f' <span class="exp-duration">({duration})</span>' if duration else "")
                + '</div>'
            )
            achs = exp.get("achievements", [])
            ach_html = _bullet_list(achs, sub=True) if achs else ""
            exp_html.append(f'<div class="experience">{header}{ach_html}</div>')
        parts.append("\n".join(exp_html))

    if payload.get("references"):
        parts.append(_section_title(labels["references"]))
        parts.append(f'<p class="references">{_esc(", ".join(payload["references"]))}</p>')

    return "\n".join(parts)


def _logo_data_uri() -> str | None:
    """Locate assets/inside_circle_logo.png and embed as base64 data URI."""
    here = Path(__file__).resolve().parent
    path = here.parent / "assets" / LOGO_FILENAME
    if not path.is_file():
        return None
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{data}"


def _render_slide(cv: dict, lang: str, logo_uri: str | None,
                  page: dict, index: int) -> str:
    labels = SECTION_LABELS[lang]
    first_name = _esc(cv.get("first_name", ""))
    last_name = _esc(cv.get("last_name", "")).upper()
    initials = _esc(cv.get("initials", ""))
    title = _esc(cv.get(f"title_{lang}") or cv.get("title", ""))
    domain = _esc(cv.get(f"domain_{lang}") or cv.get("domain", ""))
    if page.get("continued"):
        title = f'{title} {_esc(labels["continued"])}'
    sidebar_payload = page["sidebar"]
    main_payload = page["main"]

    logo_html = (
        f'<img class="logo" src="{logo_uri}" alt="Inside Circle">'
        if logo_uri else ""
    )

    if first_name or last_name:
        header_html = (
            '<div class="header-name">'
            + (f'<div class="first-name">{first_name}</div>' if first_name else "")
            + (f'<div class="last-name">{last_name}</div>' if last_name else "")
            + '</div>'
        )
    else:
        header_html = f'<div class="header-initials">{initials}</div>'

    tag = labels["lang_tag"].format(n=index)
    return f"""
<section class="slide-wrap" aria-label="{tag}">
  <div class="slide-meta">{tag}</div>
  <article class="slide">
    <aside class="sidebar">
      {header_html}
      <div class="sidebar-inner">{_render_sidebar(sidebar_payload, labels)}</div>
    </aside>
    <main class="content">
      <header class="content-header">
        <h1>{title}</h1>
        <div class="domain">{domain}</div>
      </header>
      <div class="content-inner">{_render_main(main_payload, labels)}</div>
    </main>
    {logo_html}
    <footer class="confidential">{_esc(labels['footer'])}</footer>
  </article>
</section>
"""


PAGE_CSS = f"""
:root {{
  --navy-sidebar: {NAVY_SIDEBAR};
  --cyan-acc: {CYAN_ACC};
  --white: {WHITE};
  --white-soft: {WHITE_SOFT};
  --body-font: "Alegreya Sans", -apple-system, BlinkMacSystemFont,
    "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
}}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; font-family: var(--body-font); }}
body {{
  background: #eef0f4;
  padding: 36px 20px 60px;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 28px;
  color: #20242c;
  -webkit-font-smoothing: antialiased;
}}
.page-title {{
  margin: 0 0 4px;
  font-size: 13px;
  letter-spacing: 0.2em;
  text-transform: uppercase;
  color: #6b7280;
  font-weight: 600;
}}
.slide-wrap {{ width: 100%; max-width: 1280px; }}
.slide-meta {{
  font-size: 11px;
  letter-spacing: 0.18em;
  text-transform: uppercase;
  color: #6b7280;
  margin-bottom: 8px;
  padding-left: 4px;
}}

/* The slide. 16:9 box. Fond uniforme navy ; séparateur cyan vertical
   matérialise la frontière sidebar / main column. */
.slide {{
  position: relative;
  width: 100%;
  aspect-ratio: 16 / 9;
  background: var(--navy-sidebar);
  color: var(--white);
  display: grid;
  grid-template-columns: 35.2% 64.8%;
  overflow: hidden;
  border-radius: 4px;
  box-shadow: 0 12px 38px rgba(8, 18, 60, 0.20);
}}
.sidebar {{
  padding: 22px 22px 30px;
  display: flex;
  flex-direction: column;
  position: relative;
  border-right: 1.5px solid var(--cyan-acc);
}}
.header-initials {{
  font-size: 64px;
  font-weight: 800;
  color: var(--cyan-acc);
  letter-spacing: 0.02em;
  line-height: 1;
  margin-bottom: 16px;
}}
.header-name {{
  margin-bottom: 18px;
  color: var(--cyan-acc);
  line-height: 1.05;
}}
.header-name .first-name {{
  font-size: 36px;
  font-weight: 400;
}}
.header-name .last-name {{
  font-size: 42px;
  font-weight: 800;
  letter-spacing: 0.01em;
}}
.sidebar-inner {{
  font-size: 15.5px;        /* ≥ PPT 12pt visually */
  line-height: 1.35;
}}
.sidebar-inner h3 {{ margin: 12px 0 4px; }}
.sidebar-inner h3:first-child {{ margin-top: 0; }}
.content {{
  padding: 22px 28px 30px;
  display: flex;
  flex-direction: column;
  position: relative;
  min-height: 0;
}}
.content-header h1 {{
  margin: 0;
  padding-right: 160px;        /* reserve for the horizontal logo */
  font-size: 30px;
  font-weight: 700;
  color: var(--white);
  letter-spacing: 0.01em;
  text-wrap: balance;
}}
.logo {{
  position: absolute;
  top: 14px;                   /* almost flush with the top edge */
  right: 24px;
  height: 40px;                /* logo content cropped (aspect 3.1) */
  width: auto;
  z-index: 2;
  pointer-events: none;
}}
.content-header .domain {{
  margin-top: 2px;
  color: var(--cyan-acc);
  font-style: italic;
  font-size: 18px;
}}
.content-inner {{
  margin-top: 10px;
  font-size: 15.5px;
  line-height: 1.35;
  flex: 1;
  min-height: 0;
}}
.section {{
  font-size: 14.5px;
  font-weight: 800;
  letter-spacing: 0.08em;
  color: var(--cyan-acc);
  margin: 10px 0 4px;
  padding-bottom: 3px;
  border-bottom: 1px solid rgba(58, 237, 229, 0.4);
}}
.content-inner .section:first-child,
.sidebar-inner .section:first-child {{ margin-top: 0; }}

.bullet-list {{ list-style: none; padding: 0; margin: 0; }}
.bullet {{
  display: grid;
  grid-template-columns: 12px 1fr;
  gap: 4px;
  align-items: start;
  padding: 1px 0;
}}
.bullet .dot {{ color: var(--cyan-acc); font-weight: 700; }}
.bullet-list.sub {{ padding-left: 14px; margin-top: 2px; }}

.summary {{ margin: 0 0 4px; }}
.experience {{ margin-top: 6px; }}
.experience:first-child {{ margin-top: 0; }}
.exp-header {{ font-size: 15.5px; }}
.exp-marker {{ color: var(--cyan-acc); font-weight: 700; margin-right: 4px; }}
.exp-employer {{ font-weight: 700; }}
.exp-sep {{ color: var(--white-soft); }}
.exp-duration {{
  color: var(--cyan-acc); font-style: italic; font-size: 14px;
  margin-left: 4px;
}}
.references {{ margin: 0; font-style: italic; color: var(--white-soft); }}
.muted {{ color: var(--white-soft); }}
.edu-list {{ display: flex; flex-direction: column; gap: 3px; }}
.edu-row {{ display: grid; gap: 0; }}
.edu-line {{ }}
.edu-year {{ color: var(--cyan-acc); font-weight: 700; margin-right: 4px; }}
.edu-school {{ font-weight: 700; }}
.edu-degree {{ color: var(--white-soft); font-style: italic; font-size: 14.5px; }}
.confidential {{
  position: absolute;
  bottom: 0; left: 0; right: 0;
  background: var(--navy-sidebar);  /* masks any textbox overflow */
  text-align: center;
  font-size: 11px;
  color: var(--white-soft);
  font-style: italic;
  padding: 6px 28px;
  z-index: 3;
}}
@media (max-width: 900px) {{
  .slide {{ font-size: 0.92em; }}
  .header-initials {{ font-size: 46px; }}
  .content-header h1 {{ font-size: 20px; }}
}}
"""


def render_html(cv: dict, langs: list[str] | None = None) -> str:
    initials = cv.get("initials", "CV")
    logo_uri = _logo_data_uri()
    langs = langs or [source_lang(cv)]
    slides, n = [], 0
    for lang in langs:
        for page in paginate(cv.get(lang, cv)):
            n += 1
            slides.append(_render_slide(cv, lang, logo_uri, page, n))
    body = "\n".join(slides)

    return f"""<meta charset="utf-8">
<title>CV {html.escape(initials)} — Aperçu Inside Circle</title>
<meta name="description" content="Aperçu visuel du CV harmonisé Inside Circle.">
<style>{PAGE_CSS}</style>
<div class="page-title">CV {html.escape(initials)} — Aperçu Inside Circle ({n} slide(s))</div>
{body}
"""


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--lang", choices=["auto", "fr", "en", "both"],
                        default="auto")
    args = parser.parse_args()
    langs = None if args.lang == "auto" else (
        ["fr", "en"] if args.lang == "both" else [args.lang])

    cv = json.loads(args.input.read_text(encoding="utf-8"))
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(render_html(cv, langs), encoding="utf-8")
    print(f"✓ {args.input.name} → {args.out}")


if __name__ == "__main__":
    main()
