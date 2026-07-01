# Générateur de CV — Inside Circle

Agent qui harmonise des CV hétérogènes (PDF / Word / PPT) au format Inside Circle :
**1 slide FR + 1 slide EN** par CV, anonymisé, prêt à envoyer à un client final.

## Architecture

Pipeline en 2 étapes :

1. **Extraction (Claude)** — lecture du CV source et production d'un JSON normalisé
   selon le schéma `agent/schema.json`. Le prompt système est dans
   `agent/system_prompt.md` (anonymisation, champs interdits retirés, traduction
   FR↔EN automatique selon la langue de l'input).

2. **Rendu (Python)** — `renderer/generate_cv.py` consomme le JSON et produit le
   `.pptx` au template Inside Circle (couleurs charte, police Alegreya Sans,
   layout sidebar / main, footer confidentialité).

## Workflow

```bash
# 1. Déposer le ou les CV sources
cp mon_cv.pdf inputs/

# 2. Demander à Claude de produire le JSON normalisé (suit agent/system_prompt.md)
#    → résultat dans data/<slug>.json

# 3a. Générer un PPT par CV
python renderer/generate_cv.py data/<slug>.json --out outputs/<slug>.pptx

# 3b. Ou générer UN SEUL PPT contenant tous les CVs (mode livraison batch)
python renderer/generate_cv.py data/ --merged --out outputs/all_cvs.pptx
```

Modes du renderer :
- **Single** : `<file>.json` → `<file>.pptx` (2 slides FR + EN)
- **Batch séparé** : `data/` → `outputs/` (un PPT par CV)
- **Batch fusionné** : `data/ --merged` → un `.pptx` unique de N × 2 slides

## Règles métier (résumé)

- **Anonymisation** : initiale prénom + initiale nom (ex. `Richard Michaud` → `R.M.`)
- **Champs systématiquement retirés** : adresse, email, téléphone, date de naissance,
  nationalité, lien LinkedIn, photo, mention de genre
- **Sortie bilingue** : input FR → traduit EN ; input EN → traduit FR ;
  input bilingue → utilise chaque version telle quelle
- **Contenu** : pas de reformulation ni d'invention, info manquante = vide,
  quantification conservée quand présente dans la source
- **Hobbies** : conservés (élément différenciant côté client)
- **Engagements / réalisations extra** : conservés quand pertinents pour un livrable
  consulting financier
- **Tronquage** : si trop dense pour 1 page, on coupe les expériences les plus
  anciennes en priorité

## Structure

```
agent/
  system_prompt.md   # prompt extraction
  schema.json        # schéma JSON normalisé
renderer/
  generate_cv.py     # CLI json → pptx
  theme.py           # couleurs / fontes / dimensions
  i18n.py            # labels FR/EN
inputs/              # CV sources
data/                # JSON intermédiaires
outputs/             # PPT finaux
```
