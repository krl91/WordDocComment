---
name: word-comment-merge-en
description: "Transfer and comparison of Word document comments (.docx). Use when: transferring Word comments between document versions, merging annotations from Doc1 to Doc2, exporting Word comments to CSV, comparing comments across 2 or 3 Word document versions, anchoring comments to paragraphs, detecting unique Doc2 comments, migrating OOXML annotations. Scripts: transfer_comments.py, export_comments.py."
argument-hint: "Desired mode: 'merge' to transfer comments, 'export' for CSV, 'compare' for multi-version comparison"
---

# Word Comment Merge — Skill

Reference skill for both project scripts.
Load this skill to get exact command syntax, CSV format details,
similarity thresholds, and troubleshooting tips.

---

## Available scripts

| Script | Role |
|--------|------|
| [`transfer_comments.py`](../../transfer_comments.py) | Transfers comments from Doc1 → Doc2, produces Doc_Final |
| [`export_comments.py`](../../export_comments.py) | Exports comments from one or more .docx files to CSV |

---

## Workflow 1 — Comment merge

### Command

```bash
python3 transfer_comments.py <source.docx> <content.docx> [output.docx] [--threshold 0.75]
```

### Arguments

| Argument | Required | Default | Description |
|----------|----------|---------|-------------|
| `source` | ✅ | — | File containing the comments to transfer |
| `content` | ✅ | — | File whose text is preserved in the result |
| `output` | ❌ | `Doc_Final.docx` | Output filename |
| `--threshold` | ❌ | `0.75` | Text similarity threshold (0 to 1) for automatic anchoring |

### Internal steps

```
[1/4] Copy content.docx → output.docx
[2/4] Detect comments unique to Doc2
[3/4] Strip existing comment markers
[4/4] Re-insert and re-anchor all comments
```

### Similarity thresholds — Step 2 (Doc2 detection)

| Symbol | Score | Behaviour |
|--------|-------|-----------|
| `[=]` | ≥ 90 % | Already in Doc1 → skipped automatically |
| `[?]` | 60–90 % | Ambiguous → **prompts the user** |
| `[+]` | < 60 % | Unique to Doc2 → kept automatically |

### Similarity thresholds — Step 4 (re-anchoring)

| Score | Behaviour |
|-------|-----------|
| ≥ 90 % | Automatic anchoring (exact or approximate phrase) |
| 70–90 % | Presents 5 candidates → **user chooses** |
| < 70 % | Offers chapter-start anchor or manual selection |
| No match | Offers chapter-start anchor or skip |

### Preserved comment properties

- Author, initials, and original timestamp
- Resolved / unresolved status (`[RESOLU]`)
- Full comment text

---

## Workflow 2 — CSV export

### Command

```bash
# Single file → <name>_comments.csv
python3 export_comments.py Doc1.docx

# Multiple files (multi-version comparison)
python3 export_comments.py Doc1.docx Doc2.docx Doc_Final.docx --output comparison.csv
```

### CSV columns (separator `;`)

| Column | Type | Description |
|--------|------|-------------|
| `source_file` | text | Name of the source `.docx` file |
| `id` | integer | Internal comment identifier |
| `author` | text | Author name |
| `initials` | text | Author initials |
| `date` | `YYYY-MM-DD` | Comment date |
| `resolved` | `oui` / `non` | Whether the comment is marked as resolved |
| `chapter` | text | Chapter path (e.g. `Intro > Section 1`) |
| `anchor_text` | text | Phrase the comment is anchored to |
| `comment_text` | text | Comment body (line breaks flattened) |

> **Encoding:** UTF-8 with BOM (`utf-8-sig`) — opens correctly in Excel.

---

## Workflow 3 — Multi-version comparison

1. Export all versions into a single CSV:

   ```bash
   python3 export_comments.py V1.docx V2.docx V3.docx --output comparison.csv
   ```

2. Read the CSV and group comments by text similarity (≥ 85 %).

3. Build a presence table:

   ```
   | Author | Date       | Comment text           | Chapter  | V1 | V2 | V3 |
   |--------|------------|------------------------|----------|----|----|----|
   | Smith  | 2024-03-15 | "Check this point"     | Intro    | ✅ | ✅ | ✅ |
   | Jones  | 2024-04-01 | "Needs rewording"      | Chap 2   | ❌ | ✅ | ✅ |
   ```

4. Calculate the summary:
   - Comments present in all versions
   - Comments that disappeared between V1→V2 (or V2→V3)
   - Comments newly added in V2 (or V3)
   - Comments resolved in at least one version

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `ImportError: transfer_comments` | `export_comments.py` run from a different directory | `cd` to the project directory before running |
| `word/comments.xml` missing | The `.docx` file has no comments | Verify the source document in Word |
| Comment not anchored | Text changed too much between versions | Lower `--threshold` (e.g. `0.60`) |
| CSV unreadable in Excel | Encoding issue | Script already uses `utf-8-sig`; open via *Data > From Text/CSV* |
| Score always 0 % | Empty comment text | Normal — comments with no text are preserved as-is |

---

## Technical details

See [`./references/ooxml.md`](./references/ooxml.md) for details about the
OOXML format (namespaces, `w:comment` tags, `w:commentRangeStart`,
`commentsExtended.xml`).
