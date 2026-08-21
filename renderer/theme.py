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
# Fond uniforme bleu foncé sur tout le slide (le gradient cyan a été retiré
# pour maximiser le contraste du texte blanc). Le séparateur vertical en cyan
# marque visuellement la frontière sidebar / colonne principale.
NAVY_SIDEBAR  = RGBColor(0x06, 0x1A, 0x5E)   # fond uniforme du slide

# Cyan d'accent — titres, KPI, initiales, puces, séparateur
CYAN_ACCENT   = RGBColor(0x3A, 0xED, 0xE5)

# Texte
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_SOFT    = RGBColor(0xCC, 0xD6, 0xE8)


# Hex pour la parité HTML preview
NAVY_SIDEBAR_HEX = "061A5E"
CYAN_ACCENT_HEX  = "3AEDE5"

# Séparateur vertical entre sidebar et main (largeur EMU + couleur cyan).
SEPARATOR_W   = Emu(15000)   # ~0.016" — trait fin


# ─── Font ────────────────────────────────────────────────────────────────────
FONT_FAMILY = "Alegreya Sans"


# ─── Type scale ──────────────────────────────────────────────────────────────
# Règle : aucune typo de contenu en-dessous de 12pt.
FS_FIRST_NAME  = Pt(28)   # Prénom (sidebar, casse d'origine)
FS_LAST_NAME   = Pt(32)   # NOM (sidebar, bold, uppercase)
FS_INITIALS    = Pt(48)   # fallback si first/last absents
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

# Logo (top-right corner) — calé sur l'exemple manuel du client :
# petit, à hauteur du titre, ancré quasi au coin supérieur droit.
# Le PNG embarqué est croppé à son bbox de contenu (aspect 3.1:1).
LOGO_FILENAME   = "inside_circle_logo.png"
LOGO_H          = Emu(420000)         # ~0.46" — compact, à hauteur du titre
LOGO_TOP        = Emu(130000)         # ~0.14" — quasi flush avec le haut
LOGO_RIGHT_PAD  = Emu(230000)         # ~0.25"
# Logo réel aspect 3.1 → largeur ≈ 1.42" pour une hauteur de 0.46".
LOGO_RESERVE_W  = Emu(1700000)        # ~1.86" — reserve incluant le padding

CONTENT_TOP     = HEADER_H + Emu(50000)
CONTENT_BOTTOM  = SLIDE_H - FOOTER_H


# ─── Capacity ────────────────────────────────────────────────────────────────
# Pour garantir 1 page, on plafonne le NOMBRE d'items. Le texte de chaque
# item n'est JAMAIS tronqué avec "…" — il doit être écrit à la bonne
# longueur en phase d'extraction. Si la mise en page déborde, on drop
# des éléments entiers (hobbies, engagements, références) via les
# fonctions _adaptive_*_trim.
MAX_EXPERIENCES        = 14
MAX_BULLETS_PER_EXP    = 6
MAX_EXPERTISE          = 10
MAX_HOBBIES            = 8
MAX_ENGAGEMENTS        = 6


# ─── Confidentiality footer text ─────────────────────────────────────────────
FOOTER_TEXT = (
    "Les informations de ce document sont strictement confidentielles — "
    "Ce document ne peut être partagé qu'avec l'accord de son propriétaire"
)
FOOTER_TEXT_EN = (
    "The information in this document is strictly confidential — "
    "This document may only be shared with the owner's consent"
)
