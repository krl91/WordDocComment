# transfer_comments

Transfère les commentaires d'un document Word (.docx) vers un autre en créant un document final fusionné.

## Cas d'usage

Vous avez deux versions d'un document :

- **Doc1.docx** — version ancienne, contenant des commentaires (certains ajoutés après la création de Doc2)
- **Doc2.docx** — version plus récente, avec son propre contenu et ses propres commentaires

Le script produit **Doc_Final.docx** contenant :
- le **texte de Doc2** (contenu le plus à jour)
- **tous les commentaires de Doc1** (anciens et nouveaux), réancrés sur les phrases correspondantes
- les **commentaires propres à Doc2** qui n'existent pas dans Doc1

## Prérequis

- Python 3.9 ou supérieur
- `lxml` (installé automatiquement si absent)

```bash
pip install lxml
```

## Installation

```bash
git clone <repo>
cd WordDocComment
```

## Utilisation

```
usage: transfer_comments [-h] [--threshold FLOAT] source content [output]
```

### Arguments

| Argument      | Description |
|---------------|-------------|
| `source`      | Document source contenant les commentaires à transférer (ex : `Doc1.docx`) |
| `content`     | Document dont le texte est conservé dans le résultat final (ex : `Doc2.docx`) |
| `output`      | *(optionnel)* Document de sortie — défaut : `Doc_Final.docx` |
| `--threshold` | *(optionnel)* Seuil de similarité textuelle pour l'ancrage automatique, entre `0` et `1` — défaut : `0.75` |

### Exemples

```bash
# Utilisation standard (noms par défaut)
python transfer_comments.py Doc1.docx Doc2.docx

# Avec un nom de sortie personnalisé
python transfer_comments.py Doc1.docx Doc2.docx resultat.docx

# Avec un seuil de similarité plus strict
python transfer_comments.py Doc1.docx Doc2.docx --threshold 0.85
```

## Logique de traitement

```
[1/4] Copie Doc2 → Doc_Final
[2/4] Analyse des commentaires de Doc2 (détection des nouveaux)
[3/4] Suppression des commentaires existants de Doc_Final
[4/4] Réinsertion de tous les commentaires
```

### Étape 2 — Détection des commentaires propres à Doc2

Chaque commentaire de Doc2 est comparé aux commentaires de Doc1 par similarité de texte, d'auteur et de date :

| Symbole | Score | Action |
|---------|-------|--------|
| `[=]`   | ≥ 90 % | Présent dans Doc1 → sera réinséré via Doc1, ignoré ici |
| `[?]`   | 60–90 % | Ambigu → **demande à l'utilisateur** |
| `[+]`   | < 60 % | Unique à Doc2 → conservé automatiquement |

### Étape 4 — Réancrage des commentaires

Pour chaque commentaire de Doc1 (et les uniques de Doc2), le script cherche le meilleur emplacement dans Doc_Final :

| Score de correspondance | Comportement |
|-------------------------|--------------|
| ≥ 90 % | Ancrage automatique sur la phrase exacte ou approximative |
| 70–90 % | **Demande à l'utilisateur** de choisir parmi les 5 meilleures options |
| < 70 % | Propose d'ancrer au début du chapitre, ou choix manuel |
| Aucune | Propose d'ancrer au début du chapitre, ou d'ignorer |

L'ancrage tente d'abord une **correspondance exacte de la phrase** dans le même chapitre, puis une correspondance approchée, puis un repli sur le chapitre entier.

### Propriétés des commentaires préservées

- Auteur et horodatage d'origine
- Statut résolu / non résolu (`[RESOLU]`)
- Texte du commentaire

## Résumé affiché en fin d'exécution

```
==============================================================
  Resume du traitement
  Commentaires Doc1 inseres    : 11
  Commentaires Doc2 conserves  : 2
  Effectivement inseres        : 12
  Ignores                      : 1
  Fichier cree                 : Doc_Final.docx

  Liste des commentaires ignores :
    #9 | Karel REDON (2024-03-15) | raison: ignore par utilisateur
         ancre    : "texte introuvable"
         chapitre : "Introduction"
         texte    : "Ce commentaire n'a pas pu etre place"
==============================================================
```

## Fusion depuis plusieurs versions antérieures (Doc0 + Doc1 + Doc2)

Si les commentaires à récupérer se trouvent dans **deux versions source** (par exemple Doc0.docx *et* Doc1.docx), il suffit de **chaîner le script deux fois** :

```bash
# Étape 1 : fusionner Doc0 dans Doc1 → fichier intermédiaire
python3 transfer_comments.py Doc0.docx Doc1.docx Doc1_merged.docx

# Étape 2 : transférer tous les commentaires vers Doc2
python3 transfer_comments.py Doc1_merged.docx Doc2.docx Doc_Final.docx
```

**Résultat de `Doc_Final.docx` :**
- Texte de Doc2 (le plus récent)
- Commentaires de Doc0 (réancrés passe 1 puis passe 2)
- Commentaires propres à Doc1 (réancrés passe 2)
- Commentaires propres à Doc2 conservés


> **Conseil issu des tests** : quand le texte de Doc2 diffère légèrement de Doc1 (typos, reformulations),
> les commentaires de Doc0 peuvent obtenir une confiance de 88 % lors de la passe 2 et déclencher
> une intervention manuelle. Pour limiter ces questions, ajouter `--threshold 0.85` :
>
> ```bash
> python3 transfer_comments.py Doc1_merged.docx Doc2.docx Doc_Final.docx --threshold 0.85
> ```

Ce principe se généralise à N versions : `Doc0 → Doc1_merged → … → Doc_Final`.

---

## Export CSV des commentaires

Le script `export_comments.py` extrait les commentaires d'un ou plusieurs `.docx` dans un fichier CSV (séparateur `;`).

```bash
# Un seul fichier → Doc1_comments.csv
python3 export_comments.py Doc1.docx

# Plusieurs fichiers (comparaison multi-versions)
python3 export_comments.py Doc1.docx Doc2.docx Doc_Final.docx --output comparaison.csv
```

### Colonnes exportées

| Colonne | Description |
|---------|-------------|
| `source_file` | Nom du fichier `.docx` source |
| `id` | Identifiant du commentaire |
| `author` / `initials` | Auteur et initiales |
| `date` | Date (`YYYY-MM-DD`) |
| `resolved` | `oui` / `non` |
| `chapter` | Chemin de chapitre (ex : `Intro > Section 1`) |
| `anchor_text` | Phrase annotée |
| `comment_text` | Texte du commentaire |

> L'encodage est **UTF-8 avec BOM** (`utf-8-sig`) pour une ouverture correcte dans Excel.

---

## Intégration VS Code (GitHub Copilot)

Ce projet inclut des automatisations pour **GitHub Copilot** dans VS Code.
Elles permettent de piloter les scripts en langage naturel, sans taper de commandes.

### Prérequis

| Composant | Version minimale |
|-----------|-----------------|
| Visual Studio Code | 1.90+ |
| Extension GitHub Copilot | Dernière version (abonnement requis) |
| Python | 3.9+ |
| `lxml` | installé automatiquement au premier lancement |

### Installation

Aucune installation supplémentaire n'est nécessaire : les fichiers `.github/` sont
détectés automatiquement par VS Code dès que le dossier est ouvert comme workspace.

```bash
git clone <repo>
code WordDocComment   # ouvre le dossier dans VS Code
```

> Si Copilot ne détecte pas les agents après ouverture, recharger la fenêtre :
> **Cmd+Shift+P → Developer: Reload Window**.

### Agents disponibles

Les agents s'utilisent dans la fenêtre **Copilot Chat** (`Ctrl+Alt+I` / `Cmd+Alt+I`).
Cliquez sur le sélecteur d'agent (icône ✦ ou menu `@`) pour en choisir un.

| Agent | Langue | Description |
|-------|--------|-------------|
| **Fusion commentaires Word** | 🇫🇷 Français | Fusion + comparaison de commentaires |
| **Word Comment Merge** | 🇬🇧 Anglais | Merge + multi-version comparison |
| **Qualite document Word** | 🇫🇷 Français | Audit qualite `.docx`, copie corrigee legere, version structurelle optionnelle |

**Exemples de messages d'activation :**

```
Fusionne les commentaires de Doc1.docx dans Doc2.docx
Transfère les annotations Word entre deux fichiers
Compare les commentaires de Doc1.docx, Doc2.docx et Doc_Final.docx
Exporte les commentaires en CSV
Verifie la qualite de mise en page de MonDocument.docx
```

**Ce que l'agent fait automatiquement :**
1. Demande les chemins de fichiers via des boutons dans le chat
2. Lance le script en arrière-plan
3. Intercepte chaque question interactive et la reformule dans le chat
4. Affiche un résumé formaté à la fin

### Skills (référence technique)

Le skill `word-comment-merge` est chargé automatiquement par les agents
lorsqu'ils ont besoin de détails techniques.  
Il peut aussi être invoqué manuellement dans le chat avec `/word-comment-merge`.

```
/word-comment-merge export  → syntaxe et format CSV
/word-comment-merge fusion  → seuils et comportements de transfer_comments.py
/word-comment-merge ooxml   → détails techniques du format OOXML
```

Le skill `word-doc-quality` sert de référence à l'agent **Qualite document Word**.
Il décrit les contrôles de lisibilité, d'homogénéité des tableaux et d'encadrés,
avec une règle stricte : corriger seulement les écarts locaux qui ne nécessitent
pas de refonte du document. Il peut aussi produire une seconde version structurelle
qui remplace certains tableaux à cellule unique par un titre suivi du contenu.

---

## Structure du projet

```
WordDocComment/
├── transfer_comments.py          # Script de fusion des commentaires
├── export_comments.py            # Script d'export CSV
├── .pylintrc                     # Configuration pylint (score : 9.88/10)
├── Doc1.docx                     # Document source (commentaires)
├── Doc2.docx                     # Document contenu
├── Doc_Final.docx                # Document généré (créé par le script)
└── .github/
    ├── agents/
    │   ├── word-comment-merge.agent.md      # Agent Copilot (français)
    │   ├── word-comment-merge-en.agent.md   # Agent Copilot (anglais)
    │   └── word-doc-quality.agent.md        # Agent Copilot qualité Word
    └── skills/
        ├── word-comment-merge/
        │   ├── SKILL.md                     # Skill Copilot (référence)
        │   └── references/
        │       └── ooxml.md                 # Détails format OOXML
        └── word-doc-quality/
            └── SKILL.md                     # Skill contrôle qualité Word
```
