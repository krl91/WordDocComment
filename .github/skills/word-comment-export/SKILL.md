---
name: word-comment-export
description: "Export Word document comments to CSV. Use when: exporting Word comments to a CSV file, extracting annotations from a .docx file, comparing comments between 2 or 3 versions of a Word document, building a comment comparison table, listing all comments with author date chapter anchor text, opening Word comments in Excel or a spreadsheet."
argument-hint: "Fichier(s) .docx à exporter, et nom du CSV de sortie"
---

# Export de commentaires Word vers CSV

Utilise le script [`export_comments.py`](../../export_comments.py) pour extraire
les commentaires d'un ou plusieurs fichiers `.docx` dans un fichier CSV (séparateur `;`).

---

## Quand utiliser ce skill

- Exporter les commentaires d'un document pour les relire dans Excel
- Comparer les commentaires de 2 ou 3 versions côte à côte
- Produire un rapport de révision (auteur, date, texte, chapitre)
- Préparer une analyse avant de lancer la fusion (`transfer_comments.py`)

---

## Commandes

### Export d'un seul fichier

```bash
python3 export_comments.py Doc1.docx
# → produit : Doc1_comments.csv
```

### Export multi-fichiers (comparaison de versions)

```bash
python3 export_comments.py Doc1.docx Doc2.docx --output comparaison.csv
python3 export_comments.py Doc0.docx Doc1.docx Doc2.docx --output comparaison.csv
```

### Nom de sortie personnalisé

```bash
python3 export_comments.py Doc1.docx --output rapport_revisions.csv
```

---

## Colonnes du CSV

| Colonne | Type | Description |
|---------|------|-------------|
| `source_file` | texte | Nom du fichier `.docx` d'origine |
| `id` | entier | Identifiant interne du commentaire |
| `author` | texte | Nom complet de l'auteur |
| `initials` | texte | Initiales de l'auteur |
| `date` | `YYYY-MM-DD` | Date du commentaire |
| `resolved` | `oui` / `non` | Commentaire marqué comme résolu |
| `chapter` | texte | Chemin de chapitre (ex : `Intro > Section 1`) |
| `anchor_text` | texte | Phrase sur laquelle le commentaire est ancré |
| `comment_text` | texte | Contenu du commentaire (retours à la ligne aplatis) |

> **Encodage :** UTF-8 avec BOM (`utf-8-sig`).
> Ouvrir dans Excel via *Données → À partir d'un fichier texte/CSV*, séparateur `;`.

---

## Comparaison multi-versions

Quand plusieurs fichiers sont passés, tous les commentaires sont regroupés dans
un seul CSV avec la colonne `source_file` pour distinguer les versions.

Pour construire un tableau de présence ✅/❌ :

1. Lancer l'export :
   ```bash
   python3 export_comments.py V1.docx V2.docx V3.docx --output comp.csv
   ```

2. Regrouper les lignes par `comment_text` (similarité ≥ 85 %) et `author`.

3. Résumé attendu :
   - Commentaires présents dans toutes les versions
   - Commentaires disparus entre V1→V2 (ou V2→V3)
   - Commentaires nouveaux en V2 (ou V3)
   - Commentaires résolus dans au moins une version

---

## Dépannage

| Symptôme | Solution |
|----------|----------|
| `ImportError: transfer_comments` | Lancer depuis le répertoire du projet |
| CSV vide (en-tête seulement) | Le fichier `.docx` n'a pas de commentaires |
| Caractères illisibles dans Excel | Ouvrir via *Données → À partir d'un fichier texte/CSV* (pas double-clic) |
