"""Inside Circle visual theme — colors, font, dimensions.

Couleurs prélevées sur le slide "Tarification" Inside Circle (screen de
référence client) : dégradé diagonal bleu royal profond → cyan turquoise.
"""

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor


# Slide dimensions — 16:9, classic widescreen
SLIDE_W = Emu(12192000)  # 13.333"
SLIDE_H = Emu(6858000)   # 7.5"


# ─── Colors ──────────────────────────────────────────────────────────────────
# Gradient — calibré sur le slide tarification Inside Circle :
# le fond reste NAVY_DEEP sur ~70 % de la diagonale, puis bascule
# vers CYAN_SOFT uniquement dans le coin bas-droite.
NAVY_DEEP     = RGBColor(0x0C, 0x2A, 0x8E)   # top-left, bleu royal profond
CYAN_SOFT     = RGBColor(0x1F, 0xCB, 0xD9)   # bottom-right, cyan turquoise

# Sidebar — un cran plus sombre pour la lisibilité du texte clair
NAVY_SIDEBAR  = RGBColor(0x07, 0x1F, 0x66)

# Cyan d'accent — titres, KPI, initiales, puces
CYAN_ACCENT   = RGBColor(0x3A, 0xED, 0xE5)

# Texte
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_SOFT    = RGBColor(0xCC, 0xD6, 0xE8)


# Hex (for XML gradient + HTML preview parity)
NAVY_DEEP_HEX    = "0C2A8E"
NAVY_SIDEBAR_HEX = "071F66"
CYAN_SOFT_HEX    = "1FCBD9"
CYAN_ACCENT_HEX  = "3AEDE5"

# Gradient stops (positions in 0-100000) — mostly navy, cyan corner only.
GRADIENT_STOPS = [
    (0,      NAVY_DEEP_HEX),
    (70000,  NAVY_DEEP_HEX),   # hold navy until ~70% diagonal
    (100000, CYAN_SOFT_HEX),
]


# ─── Font ────────────────────────────────────────────────────────────────────
FONT_FAMILY = "Alegreya Sans"


# ─── Type scale ──────────────────────────────────────────────────────────────
# Règle : aucune typo de contenu en-dessous de 10pt.
FS_INITIALS    = Pt(44)   # X.X. dans le header
FS_TITLE       = Pt(20)   # poste principal
FS_DOMAIN      = Pt(13)   # domaine d'expertise
FS_SECTION     = Pt(11)   # titres de section
FS_BODY        = Pt(10)   # corps de texte — plancher
FS_BODY_SMALL  = Pt(10)   # même plancher : la hiérarchie passe par couleur/poids
FS_FOOTER      = Pt(9)    # bandeau confidentialité (mention légale, exception)

FS_FLOOR       = Pt(10)   # plancher contrôlé par _check_min_font()


# ─── Layout (en EMU — 914400 = 1 pouce) ──────────────────────────────────────
MARGIN          = Emu(228600)         # 0.25"
SIDEBAR_W       = Emu(4300000)        # ~4.70" (35.2% — wider to host engagements)
HEADER_H        = Emu(1100000)        # ~1.20"
FOOTER_H        = Emu(280000)         # ~0.31"

CONTENT_TOP     = HEADER_H + Emu(50000)
CONTENT_BOTTOM  = SLIDE_H - FOOTER_H


# ─── Capacity / truncation ───────────────────────────────────────────────────
# Pour garantir 1 page : on cape le contenu à ces valeurs.
MAX_EXPERIENCES        = 4
MAX_BULLETS_PER_EXP    = 2     # 2 bullets per experience to guarantee 1 page
MAX_EXPERTISE          = 5
MAX_HOBBIES            = 4
MAX_ENGAGEMENTS        = 3
MAX_SUMMARY_CHARS      = 220
MAX_BULLET_CHARS       = 160   # achievement bullet — trim with ellipsis above
MAX_ENGAGEMENT_DESC_CHARS = 40 # engagement description — short tag only
MAX_DEGREE_CHARS       = 45    # education degree line (single sidebar line)


# ─── Confidentiality footer text ─────────────────────────────────────────────
FOOTER_TEXT = (
    "Les informations de ce document sont strictement confidentielles — "
    "Ce document ne peut être partagé qu'avec l'accord de son propriétaire"
)
FOOTER_TEXT_EN = (
    "The information in this document is strictly confidential — "
    "This document may only be shared with the owner's consent"
)
