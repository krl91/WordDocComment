# Référence OOXML — Commentaires Word

Les fichiers `.docx` sont des archives ZIP. Les commentaires sont stockés dans
plusieurs fichiers XML à l'intérieur.

## Fichiers impliqués

| Fichier dans le ZIP | Rôle |
|---------------------|------|
| `word/comments.xml` | Corps des commentaires (`<w:comment>`) |
| `word/document.xml` | Corps du document ; contient les ancres (`<w:commentRangeStart>`, `<w:commentRangeEnd>`, `<w:commentReference>`) |
| `word/commentsExtended.xml` | Statut résolu / non résolu (`w15:done="1"`) |
| `[Content_Types].xml` | Déclare les parts XML présentes |
| `word/_rels/document.xml.rels` | Relations entre `document.xml` et `comments.xml` |

## Namespaces clés

| Préfixe | URI |
|---------|-----|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` |

## Structure d'un commentaire (`word/comments.xml`)

```xml
<w:comment w:id="3" w:author="Karel REDON" w:date="2024-03-15T10:00:00Z" w:initials="KR">
  <w:p w14:paraId="A1B2C3D4">
    <w:r><w:t>Texte du commentaire</w:t></w:r>
  </w:p>
</w:comment>
```

## Ancres dans `word/document.xml`

```xml
<w:commentRangeStart w:id="3"/>
<w:r><w:t>phrase annotée</w:t></w:r>
<w:commentRangeEnd w:id="3"/>
<w:r><w:commentReference w:id="3"/></w:r>
```

## Statut résolu (`word/commentsExtended.xml`)

```xml
<w15:commentsEx>
  <w15:commentEx w15:paraId="A1B2C3D4" w15:done="1"/>
</w15:commentsEx>
```

`w15:paraId` correspond au `w14:paraId` du `<w:p>` enfant du `<w:comment>`.
`w15:done="1"` signifie résolu, `"0"` (ou absent) signifie non résolu.
