---
name: word-doc-review
description: "Relecture editoriale de documents Word (.docx) chapitre par chapitre. Use when: proofreading a Word document, proposing chapter-by-chapter corrections, fixing spelling grammar clarity terminology, creating a corrected Word copy, producing a correction summary table with severity."
argument-hint: "Chemin du .docx, niveau de relecture, fichier corrige"
---

# Word Doc Review - Skill

Ce skill guide la relecture editoriale d'un document Word `.docx`.
Il sert a proposer des corrections chapitre par chapitre, produire une copie corrigee,
puis generer un bilan final des corrections avec leur gravite.

---

## Objectif

Ameliorer le document sans le re-ecrire.

La relecture doit rester utile et concrete :

- corriger les fautes ;
- clarifier les phrases localement ambigues ;
- harmoniser les termes ;
- signaler les incoherences ;
- verifier les references Confluence et Jira si un MCP Atlassian est disponible ;
- produire un document corrige et un bilan exploitable.

---

## Sorties attendues

| Sortie | Description |
|--------|-------------|
| `<nom>_relu_corrige.docx` | Copie Word corrigee, jamais le fichier source |
| `<nom>_relecture_bilan.md` | Tableau bilan des corrections et problemes restants |

---

## Niveaux de relecture

| Niveau | Corrections autorisees |
|--------|------------------------|
| Legere | Orthographe, grammaire, ponctuation, typographie |
| Standard | Legere + clarte locale, repetitions, terminologie |
| Approfondie | Standard + incoherences apparentes, titres faibles, points a clarifier |

Le niveau `Standard` est le choix par defaut.

---

## Segmentation par chapitre

### Sources de segmentation

Utiliser dans l'ordre :

1. Styles de titres Word : `Heading`, `Titre`, `Titre 1`, `Titre 2`, etc.
2. Paragraphes numerotes qui ressemblent a des titres : `1.`, `1.1`, `2.3.4`.
3. Paragraphes courts en gras ou en majuscules.
4. Blocs successifs de taille raisonnable si aucun chapitre fiable n'est detecte.

### Regle

Une correction doit toujours etre rattachee a un chapitre ou, a defaut, a un bloc.
Le bilan final doit permettre de savoir ou chaque correction a ete appliquee.

---

## Types de problemes

| Type | Exemples |
|------|----------|
| Orthographe | faute, accent manquant, accord simple |
| Grammaire | accord sujet-verbe, temps incoherent |
| Typographie | espace avant `:`, guillemets, ponctuation |
| Clarite | phrase trop longue, reference vague, formulation ambigue |
| Terminologie | deux termes pour le meme concept, sigle non stable |
| Coherence | contradiction locale, exigence incompatible avec une autre phrase |
| Reference externe | lien Confluence inaccessible, ticket Jira introuvable, statut Jira incompatible |
| Style | repetition, lourdeur, titre peu explicite |

---

## Gravite

| Gravite | Definition | Action |
|---------|------------|--------|
| Critique | Contresens probable, contradiction forte, exigence incomprehensible | Ne pas corriger sans validation explicite |
| Majeure | Ambiguite significative ou incoherence qui peut gener la comprehension | Proposer, valider avant application si le sens change |
| Mineure | Fautes, grammaire, typographie, repetition locale | Peut etre appliquee apres validation du chapitre |
| Suggestion | Amelioration de style non indispensable | Appliquer seulement si validee |

---

## Regles de correction

### Corrections sures

Peuvent etre proposees comme "corrections sures" :

- fautes d'orthographe evidentes ;
- accords grammaticaux evidents ;
- ponctuation francaise ;
- double espace ;
- incoherence de casse sur un terme clairement identique ;
- repetition immediate sans valeur stylistique.

### Corrections a valider

Doivent rester dans `A valider` si elles modifient le sens ou ajoutent une interpretation :

- reformulation technique ;
- changement d'un terme metier ;
- correction issue d'une page Confluence ou d'un ticket Jira ;
- phrase contradictoire ;
- exigence incomplete ;
- titre renomme ;
- suppression d'une phrase.

### Interdictions

- Ne pas inventer d'information.
- Ne pas ajouter une decision d'architecture non presente.
- Ne pas supprimer une contrainte technique.
- Ne pas changer un terme metier si le document ne permet pas de determiner le bon.
- Ne pas transformer le ton general du document.

---

## Verification Atlassian optionnelle

Quand un MCP Atlassian est disponible, l'agent doit verifier les references Confluence
et Jira citees dans le document.

### References a detecter

- URL Confluence.
- URL Jira.
- Cle Jira de type `ABC-123`.
- Mention explicite d'une page Confluence ou d'un ticket Jira.

### Verifications utiles

- La page Confluence existe et est accessible.
- Le ticket Jira existe et est accessible.
- Le statut Jira ne contredit pas le texte.
- Le titre ou resume Jira correspond a la mention dans le document.
- La page Confluence ne contredit pas explicitement une affirmation du document.

### Warnings obligatoires

Si le MCP Atlassian n'est pas disponible :

```text
warning: verification Atlassian non effectuee - MCP indisponible
```

Si une page Confluence ou un ticket Jira cite n'est pas accessible :

```text
warning: reference Atlassian inaccessible - <reference>
```

Ces warnings doivent apparaitre dans le bilan final, meme si la relecture locale du
document a reussi.

Les informations Atlassian ne doivent pas etre appliquees comme corrections automatiques
si elles changent le sens du document. Elles doivent etre classees dans `A valider`.

---

## Application dans le `.docx`

La copie corrigee doit preserver autant que possible :

- styles ;
- commentaires ;
- images ;
- tableaux ;
- listes ;
- champs ;
- notes ;
- relations OOXML.

Approche recommandee :

1. Copier le fichier source.
2. Modifier seulement `word/document.xml` quand c'est suffisant.
3. Remplacer les textes dans les runs cibles.
4. Eviter les remplacements globaux aveugles.
5. Valider que le `.docx` de sortie reste une archive ZIP ouvrable.

Si un texte a corriger traverse plusieurs runs avec des styles differents, appliquer
la correction seulement si le remplacement conserve la mise en forme essentielle.
Sinon, placer la proposition dans le bilan comme point a valider.

---

## Bilan final

Le bilan doit etre court, mais assez precis pour auditer le travail.

### Format recommande

```markdown
# Bilan de relecture Word

## Resume

| Gravite | Detectes | Appliques | A valider |
|---------|----------|-----------|-----------|
| Critique | 0 | 0 | 0 |
| Majeure | 0 | 0 | 0 |
| Mineure | 0 | 0 | 0 |
| Suggestion | 0 | 0 | 0 |

## Bilan par chapitre

| Chapitre | Critique | Majeure | Mineure | Suggestion | Statut |
|----------|----------|---------|---------|------------|--------|
| ... | ... | ... | ... | ... | ... |

## Corrections appliquees

| Chapitre | Gravite | Type | Avant | Apres |
|----------|---------|------|-------|-------|
| ... | ... | ... | ... | ... |

## Points a valider manuellement

| Chapitre | Gravite | Probleme | Proposition |
|----------|---------|----------|-------------|
| ... | ... | ... | ... |

## Warnings

- warning: ...
```

---

## Garde-fous

- Toujours travailler sur une copie.
- Toujours valider chapitre par chapitre.
- Ne pas appliquer les corrections de sens par defaut.
- Ne pas afficher de longs extraits du document dans le chat.
- Ne pas produire un bilan exhaustif de micro-suggestions inutiles.
- Preferer peu de corrections pertinentes a une relecture bruyante.
