---
name: word-doc-quality
description: "Controle qualite et correction legere de documents Word (.docx). Use when: auditing Word document layout, detecting isolated headings, inconsistent tables, nested tables, one-cell callout boxes, inconsistent information boxes, inconsistent colors or borders, producing a corrected copy without heavy rewrite."
argument-hint: "Chemin du .docx, mode audit ou correction, et fichier de sortie"
---

# Word Doc Quality - Skill

Ce skill guide l'analyse qualite d'un document Word `.docx` et la production d'une
copie corrigee quand les corrections sont locales et sures.

L'objectif est de traiter les petits problemes qui degradent la lisibilite :
titres mal solidaires du texte, tableaux heterogenes, encadres incoherents,
couleurs non homogenes, tableaux imbriques utilises comme mise en page.

---

## Principe central

Corriger seulement les problemes localises qui peuvent etre resolus sans refonte.

Une correction automatique est acceptable si :

- le motif correct existe deja dans le document ;
- l'ecart est minoritaire ou evident ;
- la correction preserve le texte, les commentaires, les champs et la structure generale ;
- le document source reste intact.

Sinon, produire une recommandation dans le rapport.

---

## Sorties attendues

| Sortie | Description |
|--------|-------------|
| `<nom>_qualite_corrigee.docx` | Copie Word corrigee, jamais le fichier source |
| `<nom>_structure_simplifiee.docx` | Seconde version optionnelle, avec certains encadres-tableaux convertis en titre + contenu |
| `<nom>_qualite_rapport.md` | Rapport de detection, corrections et points a verifier |

---

## Controle 1 - Titres et lisibilite

### Points a detecter

- Titre suivi d'un saut de page manuel.
- Titre suivi de plusieurs paragraphes vides.
- Titre de style `Heading`, `Titre`, `Titre 1`, `Titre 2`, etc. sans `keepNext`.
- Titre detecte en bas de page avec debut du contenu sur la page suivante si un rendu PDF est disponible.

### Corrections recommandees

- Ajouter `w:keepNext` sur les paragraphes de titre concernes.
- Supprimer les paragraphes vides consecutifs uniquement s'ils sont entre un titre et son contenu.
- Eviter de deplacer du contenu entre sections.

### OOXML utile

Un paragraphe a conserver avec le suivant contient :

```xml
<w:pPr>
  <w:keepNext/>
</w:pPr>
```

Un saut de page manuel apparait souvent sous la forme :

```xml
<w:br w:type="page"/>
```

---

## Controle 2 - Tableaux

### Points a detecter

- Tableaux imbriques : `w:tbl` descendant de `w:tc`.
- Tableaux a une seule cellule : souvent des encadres.
- Premiere ligne qui ressemble a un en-tete mais n'a pas `w:tblHeader`.
- Couleur de fond incoherente dans une ligne d'en-tete.
- Bordures ou marges incoherentes entre tableaux similaires.

### Corrections recommandees

- Harmoniser l'en-tete d'un tableau avec le motif dominant du meme tableau.
- Ajouter `w:tblHeader` uniquement si la premiere ligne est clairement un en-tete.
- Ne pas aplatir les tableaux imbriques complexes.
- Ne pas changer la structure des cellules fusionnees.
- Garder la correction de tableaux a cellule unique dans une version structurelle separee, jamais dans la copie corrigee standard.

### OOXML utile

Ligne d'en-tete repetee :

```xml
<w:trPr>
  <w:tblHeader/>
</w:trPr>
```

Couleur de cellule :

```xml
<w:tcPr>
  <w:shd w:fill="FFFF00"/>
</w:tcPr>
```

Bordures de cellule :

```xml
<w:tcPr>
  <w:tcBorders>...</w:tcBorders>
</w:tcPr>
```

---

## Controle 3 - Encadres

Dans beaucoup de documents, les encadres sont en fait des tableaux a une cellule.
Ils doivent etre conserves comme tels s'ils font partie de la charte du document.

### Classification

Classer les encadres par :

- mots cles : `information`, `info`, `note`, `attention`, `important`, `avertissement`, `conseil`, `exemple` ;
- couleur de fond ;
- bordure ;
- style du premier paragraphe ;
- libelle ou premiers mots ;
- proximite avec d'autres encadres similaires.

### Regle de format dominant

Pour un meme type, utiliser le format majoritaire observe dans le document.

Exemple :

| Type | Format dominant | Action |
|------|-----------------|--------|
| Information | fond jaune + bordure fine | appliquer aux encadres Information minoritaires |
| Attention | fond rouge clair + bordure rouge | appliquer aux encadres Attention minoritaires |
| Note | pas de majorite claire | signaler, ne pas corriger |

Ne jamais inventer une couleur cible. Si le document utilise le jaune pour les encadres
"Information", le jaune est la reference. Si le document utilise le bleu, le bleu est
la reference.

---

## Controle 4 - Version structurelle sans tableaux a cellule unique

Cette sortie est optionnelle et plus interventionniste que la copie corrigee standard.
Elle sert quand l'utilisateur veut eviter les tableaux imbriques et remplacer les encadres
a cellule unique par du contenu Word normal.

### Candidats a convertir

Un tableau peut etre converti automatiquement si :

- il contient exactement une ligne et une cellule ;
- il est principalement textuel ;
- il ne contient pas d'image, graphique, equation, champ complexe ou tableau imbrique complexe ;
- son premier paragraphe contient un titre probable ;
- le titre probable est court, en gras, ou termine par `:`.

Ne pas convertir si le tableau sert clairement de mise en page complexe ou si le titre
probable n'est pas identifiable.

### Extraction du titre

Ordre de preference :

1. Premier texte en gras au debut de la cellule.
2. Premier paragraphe court.
3. Premier libelle termine par `:`.

Si aucun titre n'est fiable, signaler le tableau dans `A verifier` et le laisser intact.

### Transformation

Dans `<nom>_structure_simplifiee.docx` seulement :

1. Remplacer le tableau par un paragraphe de titre.
2. Ajouter `keepNext` au titre.
3. Inserer le reste du contenu de la cellule apres le titre.
4. Conserver les paragraphes, listes et commentaires existants autant que possible.
5. Ne pas modifier les tableaux multi-cellules.

Cette version doit etre presentee comme une proposition de simplification, pas comme
la correction principale.

---

## Controle 5 - Rapport

Le rapport doit etre actionnable et court. Il ne doit pas devenir un audit exhaustif
de toutes les imperfections du document.

### Categories

- **Corrige** : modification appliquee dans la copie.
- **A verifier** : doute raisonnable, intervention humaine recommandee.
- **Ignore** : hors perimetre ou correction trop lourde.

### Format recommande

```markdown
# Rapport qualite Word

## Resume

| Type de point | Detectes | Corriges | A verifier |
|---------------|----------|----------|------------|
| Titres        | 0        | 0        | 0          |
| Tableaux      | 0        | 0        | 0          |
| Encadres      | 0        | 0        | 0          |
| Version structurelle | 0 | 0 | 0 |

## Corrections appliquees

- Page/section si disponible - description courte.

## Points a verifier manuellement

- Description courte et raison du doute.

## Fichiers produits

- Document corrige : ...
- Version structurelle : ...
- Rapport : ...
```

---

## Garde-fous

- Toujours travailler sur une copie.
- Toujours conserver le texte.
- Ne pas supprimer de tableau, image, commentaire, note de bas de page ou champ.
- Ne pas normaliser tout le document si seuls quelques elements posent probleme.
- Ne pas convertir un encadre en paragraphe si le document utilise des tableaux a cellule unique comme standard.
- Ne convertir les tableaux a cellule unique que dans la version structurelle optionnelle.
- Ne pas faire de correction automatique quand le motif dominant n'est pas clair.
