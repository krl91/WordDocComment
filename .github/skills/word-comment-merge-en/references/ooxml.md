# OOXML Reference — Word Comments

`.docx` files are ZIP archives. Comments are stored across several XML files inside.

## Relevant files

| File in the ZIP | Role |
|-----------------|------|
| `word/comments.xml` | Comment bodies (`<w:comment>`) |
| `word/document.xml` | Document body; contains anchors (`<w:commentRangeStart>`, `<w:commentRangeEnd>`, `<w:commentReference>`) |
| `word/commentsExtended.xml` | Resolved / unresolved status (`w15:done="1"`) |
| `[Content_Types].xml` | Declares the XML parts present in the package |
| `word/_rels/document.xml.rels` | Relationships between `document.xml` and `comments.xml` |

## Key namespaces

| Prefix | URI |
|--------|-----|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` |

## Comment structure (`word/comments.xml`)

```xml
<w:comment w:id="3" w:author="Karel REDON" w:date="2024-03-15T10:00:00Z" w:initials="KR">
  <w:p w14:paraId="A1B2C3D4">
    <w:r><w:t>Comment text</w:t></w:r>
  </w:p>
</w:comment>
```

## Anchors in `word/document.xml`

```xml
<w:commentRangeStart w:id="3"/>
<w:r><w:t>annotated phrase</w:t></w:r>
<w:commentRangeEnd w:id="3"/>
<w:r><w:commentReference w:id="3"/></w:r>
```

## Resolved status (`word/commentsExtended.xml`)

```xml
<w15:commentsEx>
  <w15:commentEx w15:paraId="A1B2C3D4" w15:done="1"/>
</w15:commentsEx>
```

`w15:paraId` matches the `w14:paraId` of the `<w:p>` child of `<w:comment>`.
`w15:done="1"` means resolved; `"0"` (or absent) means unresolved.
