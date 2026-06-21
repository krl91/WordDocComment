---
name: word-comment-merge
description: "Transfert et comparaison de commentaires Word (.docx). Use when: transferring Word comments between document versions, merging annotations from Doc1 to Doc2, exporting Word comments to CSV, comparing comments across 2 or 3 Word document versions, anchoring comments to paragraphs, detecting unique Doc2 comments, migrating OOXML annotations. Scripts: transfer_comments.py, export_comments.py."
argument-hint: "Mode souhaité : 'fusion' pour transférer des commentaires, 'export' pour CSV, 'comparer' pour multi-versions"
---

# Word Comment Merge — Skill

Skill de référence pour les deux scripts du projet.
Charger ce skill pour connaître la syntaxe exacte des commandes,
le format du CSV, les seuils de similarité et les conseils de dépannage.

---

## Scripts disponibles

| Script | Rôle |
|--------|------|
| [`transfer_comments.py`](../../transfer_comments.py) | Transfère les commentaires de Doc1 → Doc2, produit Doc_Final |
| [`export_comments.py`](../../export_comments.py) | Exporte les commentaires d'un ou plusieurs .docx en CSV |

---

## Workflow 1 — Fusion de commentaires

### Commande

```bash
python3 transfer_comments.py <source.docx> <content.docx> [output.docx] [--threshold 0.75]
```

### Arguments

| Argument | Requis | Défaut | Description |
|----------|--------|--------|-------------|
| `source` | ✅ | — | Fichier contenant les commentaires à transférer |
| `content` | ✅ | — | Fichier dont le texte est conservé dans le résultat |
| `output` | ❌ | `Doc_Final.docx` | Nom du fichier résultat |
| `--threshold` | ❌ | `0.75` | Seuil de similarité (0 à 1) pour l'ancrage automatique |

### Étapes internes

```
[1/4] Copie content.docx → output.docx
[2/4] Détection des commentaires uniques à Doc2
[3/4] Suppression des marqueurs de commentaires existants
[4/4] Réinsertion et réancrage de tous les commentaires
```

### Seuils de similarité — Étape 2 (détection Doc2)

| Symbole | Score | Comportement |
|---------|-------|-------------|
| `[=]` | ≥ 90 % | Déjà dans Doc1 → ignoré automatiquement |
| `[?]` | 60–90 % | Ambigu → **demande à l'utilisateur** |
| `[+]` | < 60 % | Unique à Doc2 → conservé automatiquement |

### Seuils de similarité — Étape 4 (réancrage)

| Score | Comportement |
|-------|-------------|
| ≥ 90 % | Ancrage automatique (phrase exacte ou approchée) |
| 70–90 % | Propose 5 candidats → **choix utilisateur** |
| < 70 % | Propose le début du chapitre ou choix manuel |
| Aucune correspondance | Propose début de chapitre ou ignorer |

### Propriétés préservées

- Auteur, initiales, horodatage d'origine
- Statut résolu / non résolu (`[RESOLU]`)
- Texte intégral du commentaire

---

## Workflow 2 — Export CSV

### Commande

```bash
# Fichier unique → <nom>_comments.csv
python3 export_comments.py Doc1.docx

# Plusieurs fichiers → comments_export.csv (ou nom personnalisé)
python3 export_comments.py Doc1.docx Doc2.docx Doc_Final.docx --output comparaison.csv
```

### Colonnes du CSV (séparateur `;`)

| Colonne | Type | Description |
|---------|------|-------------|
| `source_file` | texte | Nom du fichier `.docx` d'origine |
| `id` | entier | Identifiant interne du commentaire |
| `author` | texte | Nom de l'auteur |
| `initials` | texte | Initiales de l'auteur |
| `date` | `YYYY-MM-DD` | Date du commentaire |
| `resolved` | `oui` / `non` | Commentaire marqué comme résolu |
| `chapter` | texte | Chemin de chapitre (ex : `Intro > Section 1`) |
| `anchor_text` | texte | Phrase sur laquelle le commentaire est ancré |
| `comment_text` | texte | Contenu du commentaire (sauts de ligne aplatis) |

> **Encodage :** UTF-8 avec BOM (`utf-8-sig`) — s'ouvre correctement dans Excel.

---

## Workflow 3 — Comparaison multi-versions

1. Exporter les 3 versions dans un seul CSV :

   ```bash
   python3 export_comments.py V1.docx V2.docx V3.docx --output comparaison.csv
   ```

2. Lire le CSV et regrouper les commentaires par similarité de texte (≥ 85 %).

3. Construire un tableau de présence :

   ```
   | Auteur | Date       | Texte                  | Chapitre | V1 | V2 | V3 |
   |--------|------------|------------------------|----------|----|----|----|
   | Dupont | 2024-03-15 | "Vérifier ce point"    | Intro    | ✅ | ✅ | ✅ |
   | Martin | 2024-04-01 | "À reformuler"         | Chap 2   | ❌ | ✅ | ✅ |
   ```

4. Calculer le résumé :
   - Commentaires présents dans toutes les versions
   - Commentaires disparus entre V1→V2 (ou V2→V3)
   - Commentaires nouveaux en V2 (ou V3)
   - Commentaires résolus dans au moins une version

---

## Dépannage

| Symptôme | Cause probable | Solution |
|----------|---------------|----------|
| `ImportError: transfer_comments` | `export_comments.py` lancé depuis un autre dossier | `cd` dans le répertoire du projet avant d'exécuter |
| `word/comments.xml` absent | Le fichier .docx n'a aucun commentaire | Vérifier le document source dans Word |
| Commentaire non ancré | Texte trop modifié entre les versions | Abaisser `--threshold` (ex : `0.60`) |
| CSV illisible dans Excel | Encodage BOM manquant | Le script utilise déjà `utf-8-sig` ; ouvrir via *Données > Depuis texte/CSV* |
| Score toujours 0 % | Commentaires vides (texte `""`) | Comportement normal — commentaires sans texte sont conservés tels quels |

---

## Détails techniques

Voir [`./references/ooxml.md`](./references/ooxml.md) pour les détails sur le format
OOXML (namespaces, balises `w:comment`, `w:commentRangeStart`, `commentsExtended.xml`).
