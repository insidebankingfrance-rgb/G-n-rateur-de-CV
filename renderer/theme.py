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
# le fond reste NAVY_DEEP sur ~85 % de la diagonale, puis bascule
# vers CYAN_SOFT uniquement dans le coin bas-droite. Le navy a été assombri
# pour garantir le contraste avec le texte blanc.
NAVY_DEEP     = RGBColor(0x07, 0x22, 0x7A)   # top-left, bleu royal profond
CYAN_SOFT     = RGBColor(0x1F, 0xCB, 0xD9)   # bottom-right, cyan turquoise

# Sidebar — un cran plus sombre pour la lisibilité du texte clair
NAVY_SIDEBAR  = RGBColor(0x06, 0x1A, 0x5E)

# Cyan d'accent — titres, KPI, initiales, puces
CYAN_ACCENT   = RGBColor(0x3A, 0xED, 0xE5)

# Texte
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_SOFT    = RGBColor(0xCC, 0xD6, 0xE8)


# Hex (for XML gradient + HTML preview parity)
NAVY_DEEP_HEX    = "07227A"
NAVY_SIDEBAR_HEX = "061A5E"
CYAN_SOFT_HEX    = "1FCBD9"
CYAN_ACCENT_HEX  = "3AEDE5"

# Gradient stops (positions in 0-100000) — navy holds until 85 %, cyan only
# in the bottom-right corner. Mirrors the Inside Circle template feel.
GRADIENT_STOPS = [
    (0,      NAVY_DEEP_HEX),
    (85000,  NAVY_DEEP_HEX),
    (100000, CYAN_SOFT_HEX),
]


# ─── Font ────────────────────────────────────────────────────────────────────
FONT_FAMILY = "Alegreya Sans"


# ─── Type scale ──────────────────────────────────────────────────────────────
# Règle : aucune typo de contenu en-dessous de 12pt.
FS_INITIALS    = Pt(48)   # X.X. dans le header
FS_TITLE       = Pt(24)   # poste principal
FS_DOMAIN      = Pt(15)   # domaine d'expertise
FS_SECTION     = Pt(13)   # titres de section
FS_BODY        = Pt(12)   # corps de texte — plancher
FS_BODY_SMALL  = Pt(12)   # même plancher
FS_FOOTER      = Pt(10)   # bandeau confidentialité (mention légale, exception)

FS_FLOOR       = Pt(12)   # plancher contrôlé par _check_min_font()


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
MAX_SUMMARY_CHARS      = 200
MAX_BULLET_CHARS       = 140   # achievement bullet — trim with ellipsis above
MAX_ENGAGEMENT_DESC_CHARS = 35 # engagement description — short tag only
MAX_DEGREE_CHARS       = 40    # education degree line (single sidebar line)


# ─── Confidentiality footer text ─────────────────────────────────────────────
FOOTER_TEXT = (
    "Les informations de ce document sont strictement confidentielles — "
    "Ce document ne peut être partagé qu'avec l'accord de son propriétaire"
)
FOOTER_TEXT_EN = (
    "The information in this document is strictly confidential — "
    "This document may only be shared with the owner's consent"
)
