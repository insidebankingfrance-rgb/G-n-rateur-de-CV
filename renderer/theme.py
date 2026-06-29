"""Inside Circle visual theme — colors, font, dimensions."""

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor


# Slide dimensions — 16:9, classic widescreen
SLIDE_W = Emu(12192000)  # 13.333"
SLIDE_H = Emu(6858000)   # 7.5"


# ─── Colors ──────────────────────────────────────────────────────────────────
# Sampled from the Inside Circle deck (gradient navy → cyan)
NAVY_DEEP     = RGBColor(0x06, 0x1A, 0x4D)   # background top-left
NAVY_MID      = RGBColor(0x0B, 0x2A, 0x8A)   # background mid
BLUE_ROYAL    = RGBColor(0x1E, 0x3F, 0xA8)   # card / sidebar fill
CYAN_ACCENT   = RGBColor(0x3A, 0xED, 0xE5)   # titles, KPIs, accents, initials
CYAN_SOFT     = RGBColor(0x1F, 0xCB, 0xD9)   # background bottom-right (gradient end)

WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_SOFT    = RGBColor(0xCC, 0xD6, 0xE8)   # secondary text on dark bg


# ─── Font ────────────────────────────────────────────────────────────────────
# Charte client : Alegreya Sans (Google Fonts).
# Fallback automatique géré par PowerPoint si la police n'est pas installée.
FONT_FAMILY = "Alegreya Sans"


# ─── Type scale ──────────────────────────────────────────────────────────────
FS_INITIALS    = Pt(40)   # X.X. dans le header
FS_TITLE       = Pt(18)   # poste principal
FS_DOMAIN      = Pt(12)   # domaine d'expertise
FS_SECTION     = Pt(11)   # titres de section ("EXPÉRIENCE", "FORMATION", …)
FS_BODY        = Pt(8.5)  # corps de texte
FS_BODY_SMALL  = Pt(7.5)  # bullets sous-niveau, dates
FS_FOOTER      = Pt(7)    # bandeau confidentialité


# ─── Layout (en EMU — 914400 = 1 pouce) ──────────────────────────────────────
MARGIN          = Emu(228600)         # 0.25"
SIDEBAR_W       = Emu(3600000)        # ~3.94"
HEADER_H        = Emu(1200000)        # ~1.31"
FOOTER_H        = Emu(280000)         # ~0.31"

CONTENT_TOP     = HEADER_H + Emu(50000)
CONTENT_BOTTOM  = SLIDE_H - FOOTER_H


# ─── Confidentiality footer text ─────────────────────────────────────────────
FOOTER_TEXT = (
    "Les informations de ce document sont strictement confidentielles — "
    "Ce document ne peut être partagé qu'avec l'accord de son propriétaire"
)
FOOTER_TEXT_EN = (
    "The information in this document is strictly confidential — "
    "This document may only be shared with the owner's consent"
)
