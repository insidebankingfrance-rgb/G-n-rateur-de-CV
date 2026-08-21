"""Layout & pagination — partagé par le renderer PPTX et l'aperçu HTML.

Règles de restitution (v7) :
- Le CV est produit dans LA LANGUE DU DOCUMENT SOURCE uniquement
  (champ `source_lang`, "fr" par défaut). Une traduction n'est générée que
  sur demande explicite.
- Le niveau de détail s'adapte à chaque CV, dans la limite de
  MAX_PAGES_PER_LANG slides (5) pour la langue produite.
- Le visuel est identique d'un slide à l'autre : fond navy, séparateur cyan,
  logo en haut à droite, bandeau de confidentialité.

Répartition du contenu :
    Slide 1 : nom + titre/domaine ; sidebar : expertise, langues
              main : résumé + expériences (autant que la page en tient)
    Slide 2+ : même cadre, mention « (suite) » ; la sidebar poursuit avec
              formation, engagements, centres d'intérêt ; la colonne
              principale poursuit les expériences, puis les références.

Quand tout tient sur un seul slide, le rendu est exactement celui d'avant.
Aucune troncature mid-string ("…") : on ne déplace ou ne retire que des
éléments entiers.
"""

from __future__ import annotations

from theme import (
    FOOTER_H,
    HEADER_H,
    MAX_BULLETS_PER_EXP,
    MAX_ENGAGEMENTS,
    MAX_EXPERIENCES,
    MAX_EXPERTISE,
    MAX_HOBBIES,
    SLIDE_H,
)

# ─── Métrique d'estimation (calibrée sur Alegreya Sans 12pt) ─────────────────
CPL_SIDEBAR = 42          # caractères par ligne, colonne de gauche
CPL_MAIN = 80             # caractères par ligne, colonne principale
CPL_MAIN_BULLET = CPL_MAIN - 4
LINE_PT = 14              # 12pt + ~17 % d'interligne
SECTION_PT = 19           # titre de section + filet
PARA_GAP_PT = 1

AVAIL_MAIN = (SLIDE_H - HEADER_H - FOOTER_H - 160000) / 12700
AVAIL_SIDEBAR = AVAIL_MAIN - 10
SAFETY_PT = 20            # marge : l'estimateur peut sous-évaluer de quelques pt

MAX_PAGES_PER_LANG = 5    # plafond demandé : 5 slides par langue

DEFAULT_SOURCE_LANG = "fr"


def source_lang(cv: dict) -> str:
    """Langue du document source — seule langue produite par défaut."""
    lang = (cv.get("source_lang") or DEFAULT_SOURCE_LANG).lower()
    return lang if lang in ("fr", "en") else DEFAULT_SOURCE_LANG


def wrapped_lines(text: str, cpl: int) -> int:
    if not text:
        return 0
    return max(1, (len(text) + cpl - 1) // cpl)


# ─── Hauteurs par bloc ───────────────────────────────────────────────────────
def experience_pt(exp: dict) -> float:
    header = (f"{exp.get('employer','')} {exp.get('role','')} "
              f"{exp.get('duration','')}")
    pt = wrapped_lines(header, CPL_MAIN) * LINE_PT + PARA_GAP_PT
    for b in exp.get("achievements", []):
        pt += wrapped_lines(b, CPL_MAIN_BULLET) * LINE_PT + PARA_GAP_PT
    return pt


def summary_pt(summary: str) -> float:
    summary = (summary or "").strip()
    return SECTION_PT + wrapped_lines(summary, CPL_MAIN) * LINE_PT if summary else 0.0


def references_pt(refs: list) -> float:
    if not refs:
        return 0.0
    return SECTION_PT + wrapped_lines(", ".join(refs), CPL_MAIN) * LINE_PT


def main_pt(payload: dict) -> float:
    pt = summary_pt(payload.get("summary"))
    exps = payload.get("experiences") or []
    if exps:
        pt += SECTION_PT + sum(experience_pt(e) for e in exps)
    pt += references_pt(payload.get("references") or [])
    return pt


def _edu_pt(items) -> float:
    pt = SECTION_PT
    for e in items:
        year = e.get("year", "") if isinstance(e, dict) else ""
        school = e.get("school", "") if isinstance(e, dict) else str(e)
        degree = e.get("degree", "") if isinstance(e, dict) else ""
        pt += wrapped_lines(f"{year} {school}", CPL_SIDEBAR) * LINE_PT
        if degree:
            pt += wrapped_lines(degree, CPL_SIDEBAR) * LINE_PT
    return pt


def _engagements_pt(items) -> float:
    pt = SECTION_PT
    for e in items:
        line = (f"{e.get('title','')} {e.get('description','')}"
                if isinstance(e, dict) else str(e))
        pt += wrapped_lines(line, CPL_SIDEBAR) * LINE_PT + PARA_GAP_PT
    return pt


def _block_pt(key: str, items) -> float:
    if not items:
        return 0.0
    if key == "education":
        return _edu_pt(items)
    if key == "engagements":
        return _engagements_pt(items)
    if key == "languages":
        return SECTION_PT + len(items) * LINE_PT
    return SECTION_PT + sum(
        wrapped_lines(i, CPL_SIDEBAR) for i in items
    ) * LINE_PT


SIDEBAR_ORDER = ("expertise", "languages", "education", "hobbies", "engagements")


def sidebar_pt(payload: dict) -> float:
    return sum(_block_pt(k, payload.get(k) or []) for k in SIDEBAR_ORDER)


def _empty_sidebar() -> dict:
    return {k: [] for k in SIDEBAR_ORDER}


# ─── Plafonds : on ne retire que des éléments entiers ────────────────────────
def apply_caps(payload: dict, warn=None) -> dict:
    def _w(msg):
        if warn:
            warn(msg)

    out = dict(payload)
    exps = out.get("experiences") or []
    if len(exps) > MAX_EXPERIENCES:
        dropped = [e.get("employer", "?") for e in exps[MAX_EXPERIENCES:]]
        _w(f"{len(exps) - MAX_EXPERIENCES} expérience(s) retirée(s) "
           f"(les plus anciennes) : {', '.join(dropped)}")
        exps = exps[:MAX_EXPERIENCES]
    exps = [dict(e) for e in exps]
    for e in exps:
        ach = e.get("achievements") or []
        if len(ach) > MAX_BULLETS_PER_EXP:
            _w(f"{e.get('employer','?')} — {len(ach) - MAX_BULLETS_PER_EXP} "
               f"bullet(s) retirée(s)")
            e["achievements"] = ach[:MAX_BULLETS_PER_EXP]
    out["experiences"] = exps

    for key, cap in (("expertise", MAX_EXPERTISE),
                     ("hobbies", MAX_HOBBIES),
                     ("engagements", MAX_ENGAGEMENTS)):
        if out.get(key) and len(out[key]) > cap:
            _w(f"{key} plafonné à {cap} (était {len(out[key])})")
            out[key] = out[key][:cap]
    return out


def _trim_block(key: str, items, budget: float, warn=None):
    """Un bloc de sidebar plus grand qu'une page entière : on retire des
    éléments entiers en fin de liste jusqu'à ce qu'il tienne."""
    def _w(msg):
        if warn:
            warn(msg)
    items = list(items)
    while items and _block_pt(key, items) > budget:
        items.pop()
    if items:
        _w(f"{key} réduit pour tenir dans la colonne de gauche")
    return items


# ─── Pagination ──────────────────────────────────────────────────────────────
def _page(sidebar: dict, main: dict, continued: bool) -> dict:
    return {"sidebar": sidebar, "main": main, "continued": continued}


def _pack_sidebar(all_side: dict, n_pages: int, warn=None) -> list[dict]:
    """Répartit les blocs de sidebar sur les pages, dans l'ordre canonique.

    Une page de suite ne doit jamais avoir une colonne de gauche vide : si
    tous les blocs ont déjà été placés, on y rappelle les domaines
    d'expertise (et les langues), comme un bandeau de rappel.
    """
    pages = [_empty_sidebar() for _ in range(n_pages)]
    idx, used = 0, 0.0
    for key in SIDEBAR_ORDER:
        items = all_side.get(key) or []
        if not items:
            continue
        h = _block_pt(key, items)
        if h > AVAIL_SIDEBAR:                      # bloc plus grand qu'une page
            items = _trim_block(key, items, AVAIL_SIDEBAR, warn)
            h = _block_pt(key, items)
        if used + h > AVAIL_SIDEBAR and idx + 1 < n_pages:
            idx += 1
            used = 0.0
        if used + h > AVAIL_SIDEBAR:               # plus de page disponible
            items = _trim_block(key, items, AVAIL_SIDEBAR - used, warn)
            h = _block_pt(key, items)
            if not items:
                continue
        pages[idx][key] = items
        used += h

    # Rappel d'expertise sur les pages de suite restées vides.
    recall_keys = [k for k in ("expertise", "languages") if all_side.get(k)]
    for i in range(1, n_pages):
        if any(pages[i][k] for k in SIDEBAR_ORDER):
            continue
        budget = AVAIL_SIDEBAR
        for k in recall_keys:
            items = all_side[k]
            if _block_pt(k, items) <= budget:
                pages[i][k] = items
                budget -= _block_pt(k, items)
    return pages


def _hard_budgets(n: int, summary_h: float, refs_h: float) -> list[float]:
    hard = [AVAIL_MAIN - SAFETY_PT - SECTION_PT for _ in range(n)]
    hard[0] -= summary_h
    hard[-1] -= refs_h
    return hard


def _fill(exps: list[dict], hard: list[float], cap: float
          ) -> list[list[dict]] | None:
    """Remplissage séquentiel : chaque page reçoit au plus `cap` (et jamais
    plus que sa limite physique). Retourne None si le contenu ne tient pas
    dans le nombre de pages disponibles. Une expérience n'est jamais coupée."""
    pages: list[list[dict]] = [[]]
    idx, used = 0, 0.0
    for e in exps:
        h = experience_pt(e)
        limit = min(cap, hard[idx])
        if used + h > limit:
            # page pleine, ou item plus grand que le budget (page 1 réduite par
            # le résumé) : on passe à la page suivante.
            idx += 1
            if idx >= len(hard):
                return None
            pages.append([])
            used = 0.0
            limit = min(cap, hard[idx])
        if used + h > limit:
            return None                   # une seule expérience dépasse la page
        pages[idx].append(e)
        used += h
    return pages


def _pack_balanced(exps: list[dict], summary_h: float, refs_h: float
                   ) -> list[list[dict]] | None:
    """Répartit `exps` sur le plus petit nombre de pages possible, puis
    équilibre la charge entre ces pages (on minimise la page la plus remplie).
    """
    if not exps:
        return None
    total = sum(experience_pt(e) for e in exps)
    biggest = max(experience_pt(e) for e in exps)

    # 1) plus petit nombre de pages qui tienne
    n_min = None
    for n in range(1, MAX_PAGES_PER_LANG + 1):
        hard = _hard_budgets(n, summary_h, refs_h)
        if min(hard) <= 0:
            continue
        pages = _fill(exps, hard, total)
        if pages is not None and len(pages) <= n:
            n_min = len(pages)
            break
    if n_min is None:
        return None

    # 2) équilibrage : découpe en n_min tranches contiguës qui égalise le
    #    remplissage (on minimise la somme des carrés de l'espace libre, ce qui
    #    évite à la fois les pages surchargées et les pages presque vides).
    hard = _hard_budgets(n_min, summary_h, refs_h)
    hs = [experience_pt(e) for e in exps]
    n_items = len(hs)
    INF = float("inf")

    from functools import lru_cache

    @lru_cache(maxsize=None)
    def best(i: int, p: int):
        """Coût minimal pour placer les items i.. sur les pages p..n_min-1."""
        pages_left = n_min - p
        if pages_left == 0:
            return (0.0, ()) if i == n_items else (INF, ())
        if n_items - i < pages_left:        # pas assez d'items → page vide
            return (INF, ())
        best_cost, best_cut = INF, ()
        used = 0.0
        for j in range(i, n_items):
            used += hs[j]
            if used > hard[p]:
                break
            free = hard[p] - used
            sub_cost, sub_cut = best(j + 1, p + 1)
            if sub_cost == INF:
                continue
            cost = free * free + sub_cost
            if cost < best_cost:
                best_cost, best_cut = cost, (j + 1,) + sub_cut
        return (best_cost, best_cut)

    cost, cuts = best(0, 0)
    if cost == INF:
        return _fill(exps, hard, total)
    pages, start = [], 0
    for cut in cuts:
        pages.append(exps[start:cut])
        start = cut
    return pages


def paginate(payload: dict, warn=None) -> list[dict]:
    """Découpe un payload de langue en 1 à MAX_PAGES_PER_LANG pages."""
    def _w(msg):
        if warn:
            warn(msg)

    p = apply_caps(payload, warn)
    all_side = {k: p.get(k) or [] for k in SIDEBAR_ORDER}
    summary = (p.get("summary") or "").strip()
    exps = list(p.get("experiences") or [])
    refs = list(p.get("references") or [])

    # ── Cas 1 : tout tient sur un seul slide → rendu historique inchangé.
    single = {"summary": summary, "experiences": exps, "references": refs}
    if (main_pt(single) <= AVAIL_MAIN - SAFETY_PT
            and sidebar_pt(all_side) <= AVAIL_SIDEBAR):
        return [_page(all_side, {**single, "continued": False}, False)]

    # ── Cas 2 : découpe des expériences sur plusieurs pages, ÉQUILIBRÉE.
    # On cherche le plus petit nombre de pages qui tienne, puis on répartit le
    # contenu à hauteur égale entre les pages pour éviter les grands vides.
    groups = _pack_balanced(exps, summary_pt(summary), references_pt(refs))
    if groups is None:
        # Même au plafond de pages, tout ne tient pas : on retire les
        # expériences les plus anciennes (jamais de coupe à l'intérieur).
        kept = list(exps)
        while kept and groups is None:
            dropped = kept.pop()
            _w(f"plafond de {MAX_PAGES_PER_LANG} slides atteint — "
               f"expérience retirée : {dropped.get('employer','?')}")
            groups = _pack_balanced(kept, summary_pt(summary),
                                    references_pt(refs))
        if groups is None:
            groups = [exps[:1]]

    side_pages = _pack_sidebar(all_side, len(groups), warn)
    pages = []
    for i, group in enumerate(groups):
        is_last = i == len(groups) - 1
        pages.append(_page(
            side_pages[i],
            {"summary": summary if i == 0 else "",
             "experiences": group,
             "references": refs if is_last else [],
             "continued": i > 0},
            i > 0,
        ))
    return pages
