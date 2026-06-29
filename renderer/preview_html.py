"""Render a CV JSON to a single HTML page showing the FR + EN slides.

Used to validate the visual layout without opening PowerPoint. The HTML is a
faithful facsimile of what the .pptx slides look like.

Usage:
    python renderer/preview_html.py <input.json> --out preview.html
"""

from __future__ import annotations

import argparse
import html
import json
from pathlib import Path


# ─── Tokens (mirror renderer/theme.py) ───────────────────────────────────────
NAVY_DEEP   = "#061A4D"
NAVY_MID    = "#0B2A8A"
BLUE_ROYAL  = "#1E3FA8"
CYAN_ACC    = "#3AEDE5"
CYAN_SOFT   = "#1FCBD9"
WHITE       = "#FFFFFF"
WHITE_SOFT  = "#CCD6E8"


SECTION_LABELS = {
    "fr": {
        "summary":      "RÉSUMÉ",
        "expertise":    "DOMAINES D'EXPERTISE",
        "languages":    "LANGUES",
        "education":    "FORMATION & CERTIFICATIONS",
        "experience":   "EXPÉRIENCE PROFESSIONNELLE",
        "references":   "RÉFÉRENCES",
        "engagements":  "ENGAGEMENTS & RÉALISATIONS",
        "hobbies":      "CENTRES D'INTÉRÊT",
        "footer":       "Les informations de ce document sont strictement confidentielles — Ce document ne peut être partagé qu'avec l'accord de son propriétaire",
        "lang_tag":     "Slide 1 — Français",
    },
    "en": {
        "summary":      "SUMMARY",
        "expertise":    "MAIN AREAS OF EXPERTISE",
        "languages":    "LANGUAGES",
        "education":    "EDUCATION & CERTIFICATIONS",
        "experience":   "PROFESSIONAL EXPERIENCE",
        "references":   "MAIN REFERENCES",
        "engagements":  "ENGAGEMENTS & ACHIEVEMENTS",
        "hobbies":      "INTERESTS",
        "footer":       "The information in this document is strictly confidential — This document may only be shared with the owner's consent",
        "lang_tag":     "Slide 2 — English",
    },
}


def _esc(s) -> str:
    return html.escape(str(s) if s is not None else "")


def _section_title(label: str) -> str:
    return f'<h3 class="section">{_esc(label)}</h3>'


def _bullets(items, *, cls="bullet") -> str:
    out = []
    for it in items:
        out.append(f'<li class="{cls}"><span class="dot">•</span><span>{_esc(it)}</span></li>')
    return f'<ul class="bullet-list">{"".join(out)}</ul>'


def _render_sidebar(payload, lang_labels) -> str:
    parts = []

    if payload.get("expertise"):
        parts.append(_section_title(lang_labels["expertise"]))
        parts.append(_bullets(payload["expertise"]))

    if payload.get("languages"):
        parts.append(_section_title(lang_labels["languages"]))
        items = []
        for lng in payload["languages"]:
            if isinstance(lng, dict):
                name = lng.get("name", "")
                level = lng.get("level", "")
                items.append(f"{name} — {level}" if level else name)
            else:
                items.append(str(lng))
        parts.append(_bullets(items))

    if payload.get("education"):
        parts.append(_section_title(lang_labels["education"]))
        rows = []
        for e in payload["education"]:
            if isinstance(e, dict):
                year = _esc(e.get("year", ""))
                school = _esc(e.get("school", ""))
                degree = _esc(e.get("degree", ""))
                rows.append(
                    f'<div class="edu-row">'
                    f'<div class="edu-year">{year}</div>'
                    f'<div class="edu-school">{school}</div>'
                    f'<div class="edu-degree">{degree}</div>'
                    f'</div>'
                )
            else:
                rows.append(f'<div class="edu-row"><div class="edu-school">{_esc(e)}</div></div>')
        parts.append(f'<div class="edu-list">{"".join(rows)}</div>')

    return "\n".join(parts)


def _render_main(payload, lang_labels) -> str:
    parts = []

    summary = (payload.get("summary") or "").strip()
    if summary:
        parts.append(_section_title(lang_labels["summary"]))
        parts.append(f'<p class="summary">{_esc(summary)}</p>')

    if payload.get("experiences"):
        parts.append(_section_title(lang_labels["experience"]))
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
                + f'</div>'
            )
            achs = exp.get("achievements", [])
            ach_html = ""
            if achs:
                ach_items = "".join(
                    f'<li class="bullet sub"><span class="dot">•</span><span>{_esc(a)}</span></li>'
                    for a in achs
                )
                ach_html = f'<ul class="bullet-list sub">{ach_items}</ul>'
            exp_html.append(f'<div class="experience">{header}{ach_html}</div>')
        parts.append("\n".join(exp_html))

    if payload.get("references"):
        parts.append(_section_title(lang_labels["references"]))
        parts.append(f'<p class="references">{_esc(", ".join(payload["references"]))}</p>')

    if payload.get("engagements"):
        parts.append(_section_title(lang_labels["engagements"]))
        rows = []
        for e in payload["engagements"]:
            if isinstance(e, dict):
                t = _esc(e.get("title", ""))
                d = _esc(e.get("description", ""))
                rows.append(
                    f'<li class="bullet"><span class="dot">•</span>'
                    f'<span><strong>{t}</strong>'
                    + (f' — <span class="muted">{d}</span>' if d else "")
                    + f'</span></li>'
                )
            else:
                rows.append(f'<li class="bullet"><span class="dot">•</span><span>{_esc(e)}</span></li>')
        parts.append(f'<ul class="bullet-list">{"".join(rows)}</ul>')

    if payload.get("hobbies"):
        parts.append(_section_title(lang_labels["hobbies"]))
        parts.append(f'<p class="hobbies">{_esc("  •  ".join(payload["hobbies"]))}</p>')

    return "\n".join(parts)


def _render_slide(cv: dict, lang: str) -> str:
    labels = SECTION_LABELS[lang]
    initials = _esc(cv.get("initials", ""))
    title = _esc(cv.get(f"title_{lang}") or cv.get("title", ""))
    domain = _esc(cv.get(f"domain_{lang}") or cv.get("domain", ""))
    payload = cv.get(lang, cv)

    return f"""
<section class="slide-wrap" aria-label="{labels['lang_tag']}">
  <div class="slide-meta">{labels['lang_tag']}</div>
  <article class="slide">
    <div class="bg-accent"></div>
    <aside class="sidebar">
      <div class="header-initials">{initials}</div>
      <div class="sidebar-inner">{_render_sidebar(payload, labels)}</div>
    </aside>
    <main class="content">
      <header class="content-header">
        <h1>{title}</h1>
        <div class="domain">{domain}</div>
      </header>
      <div class="content-inner">{_render_main(payload, labels)}</div>
    </main>
    <footer class="confidential">{_esc(labels['footer'])}</footer>
  </article>
</section>
"""


PAGE_CSS = """
:root {
  --navy-deep: #061A4D;
  --navy-mid: #0B2A8A;
  --blue-royal: #1E3FA8;
  --cyan-acc: #3AEDE5;
  --cyan-soft: #1FCBD9;
  --white: #FFFFFF;
  --white-soft: #CCD6E8;
  --body-font: "Alegreya Sans", -apple-system, BlinkMacSystemFont,
    "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; font-family: var(--body-font); }
body {
  background: #f4f5f8;
  padding: 36px 20px 60px;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 28px;
  color: #20242c;
  -webkit-font-smoothing: antialiased;
}
.page-title {
  text-align: center;
  margin: 0 0 8px;
  font-size: 20px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #4a5266;
}
.slide-wrap { width: 100%; max-width: 1280px; }
.slide-meta {
  font-size: 11px;
  letter-spacing: 0.18em;
  text-transform: uppercase;
  color: #6b7280;
  margin-bottom: 8px;
  padding-left: 4px;
}
.slide {
  position: relative;
  width: 100%;
  aspect-ratio: 16 / 9;
  background: var(--navy-deep);
  color: var(--white);
  display: grid;
  grid-template-columns: 30% 70%;
  overflow: hidden;
  border-radius: 4px;
  box-shadow: 0 10px 40px rgba(0,0,0,0.18);
}
.bg-accent {
  position: absolute;
  right: -8%;
  bottom: -14%;
  width: 55%;
  aspect-ratio: 1;
  background: radial-gradient(circle at 40% 40%, var(--cyan-soft) 0%, transparent 65%);
  pointer-events: none;
  z-index: 0;
}
.sidebar {
  background: var(--navy-mid);
  padding: 28px 26px 18px;
  display: flex;
  flex-direction: column;
  position: relative;
  z-index: 1;
}
.header-initials {
  font-size: 56px;
  font-weight: 800;
  color: var(--cyan-acc);
  letter-spacing: 0.02em;
  line-height: 1;
  margin-bottom: 18px;
}
.sidebar-inner {
  font-size: 11.5px;
  line-height: 1.45;
}
.sidebar-inner h3 { margin: 14px 0 6px; }
.sidebar-inner h3:first-child { margin-top: 0; }
.content {
  padding: 24px 30px 18px;
  display: flex;
  flex-direction: column;
  position: relative;
  z-index: 1;
  min-height: 0;
}
.content-header h1 {
  margin: 0;
  font-size: 26px;
  font-weight: 700;
  color: var(--white);
  letter-spacing: 0.01em;
  text-wrap: balance;
}
.content-header .domain {
  margin-top: 2px;
  color: var(--cyan-acc);
  font-style: italic;
  font-size: 14px;
}
.content-inner {
  margin-top: 12px;
  font-size: 11.5px;
  line-height: 1.45;
  flex: 1;
  min-height: 0;
}
.section {
  font-size: 12px;
  font-weight: 800;
  letter-spacing: 0.08em;
  color: var(--cyan-acc);
  margin: 12px 0 5px;
  padding-bottom: 3px;
  border-bottom: 1px solid rgba(58, 237, 229, 0.4);
}
.sidebar .section { font-size: 11px; }
.content-inner .section:first-child,
.sidebar-inner .section:first-child { margin-top: 0; }
.bullet-list { list-style: none; padding: 0; margin: 0; }
.bullet {
  display: grid;
  grid-template-columns: 14px 1fr;
  align-items: start;
  gap: 4px;
  padding: 1px 0;
}
.bullet .dot { color: var(--cyan-acc); font-weight: 700; }
.bullet-list.sub { padding-left: 16px; margin-top: 2px; }
.summary { margin: 0 0 4px; }
.experience { margin-top: 8px; }
.experience:first-child { margin-top: 0; }
.exp-header { font-size: 12px; }
.exp-marker { color: var(--cyan-acc); font-weight: 700; margin-right: 4px; }
.exp-employer { font-weight: 700; }
.exp-sep { color: var(--white-soft); }
.exp-duration { color: var(--cyan-acc); font-style: italic; font-size: 10.5px; margin-left: 4px; }
.references { margin: 0; font-style: italic; color: var(--white-soft); }
.hobbies { margin: 0; }
.muted { color: var(--white-soft); }
.edu-list { display: flex; flex-direction: column; gap: 4px; }
.edu-row { display: grid; gap: 1px; }
.edu-year { color: var(--cyan-acc); font-weight: 700; font-size: 10.5px; }
.edu-school { font-weight: 700; }
.edu-degree { color: var(--white-soft); font-style: italic; font-size: 10.5px; }
.confidential {
  position: absolute;
  bottom: 6px;
  left: 0; right: 0;
  text-align: center;
  font-size: 9px;
  color: var(--white-soft);
  font-style: italic;
  padding: 0 28px;
  z-index: 2;
}
@media (max-width: 900px) {
  .slide { font-size: 0.9em; }
  .header-initials { font-size: 42px; }
  .content-header h1 { font-size: 20px; }
}
"""


def render_html(cv: dict) -> str:
    initials = cv.get("initials", "CV")
    fr = _render_slide(cv, "fr")
    en = _render_slide(cv, "en")

    return f"""<meta charset="utf-8">
<title>CV {html.escape(initials)} — Aperçu Inside Circle</title>
<meta name="description" content="Aperçu visuel du CV harmonisé Inside Circle (FR + EN).">
<style>{PAGE_CSS}</style>
<h1 class="page-title">CV {html.escape(initials)} — Aperçu Inside Circle</h1>
{fr}
{en}
"""


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()

    cv = json.loads(args.input.read_text(encoding="utf-8"))
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(render_html(cv), encoding="utf-8")
    print(f"✓ {args.input.name} → {args.out}")


if __name__ == "__main__":
    main()
