#!/usr/bin/env python3
"""Export Word document comments to a semicolon-separated CSV file.

Accepts one or more .docx files.  When several files are given, all
comments are merged into a single CSV with a ``source_file`` column so
that an agent (or a spreadsheet) can compare versions side-by-side.

Usage
-----
    python export_comments.py tests/fixtures/Doc1.docx
    python export_comments.py tests/fixtures/Doc1.docx tests/fixtures/Doc2.docx --output comparison.csv
    python export_comments.py tests/fixtures/Doc1.docx tests/fixtures/Doc2.docx Doc3.docx
    python export_comments.py --help
"""

import argparse
import csv
import os
import sys

# Re-use the extraction logic already present in transfer_comments.py.
# The guard ``if __name__ == "__main__":`` in that module ensures that
# its main() is never executed on import.
try:
    from transfer_comments import load_doc_comments
except ImportError as exc:
    print(
        "[ERREUR] Impossible d'importer transfer_comments.py. "
        "Assurez-vous qu'il est dans le meme repertoire.",
        file=sys.stderr,
    )
    raise SystemExit(1) from exc

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CSV_FIELDS = [
    "source_file",
    "id",
    "author",
    "initials",
    "date",
    "resolved",
    "chapter",
    "anchor_text",
    "comment_text",
]


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------


def _default_output(docx_paths: list) -> str:
    """Return a sensible default CSV filename from the input file list."""
    if len(docx_paths) == 1:
        base = os.path.splitext(os.path.basename(docx_paths[0]))[0]
        return f"{base}_comments.csv"
    return "comments_export.csv"


def _build_rows(docx_paths: list) -> list:
    """Extract comments from every DOCX file and return a flat row list."""
    rows = []
    for docx_path in docx_paths:
        if not os.path.isfile(docx_path):
            print(
                f"[ERREUR] Fichier introuvable : {docx_path}",
                file=sys.stderr,
            )
            raise SystemExit(1)

        comments, _, _ = load_doc_comments(docx_path)
        source_name = os.path.basename(docx_path)

        print(f"  {source_name} : {len(comments)} commentaire(s) trouve(s).")

        for cmt in comments:
            rows.append(
                {
                    "source_file": source_name,
                    "id": cmt["id"],
                    "author": cmt["author"],
                    "initials": cmt["initials"],
                    # Keep only the date part (YYYY-MM-DD) from the ISO timestamp
                    "date": (cmt["date"] or "")[:10],
                    "resolved": "oui" if cmt["resolved"] else "non",
                    "chapter": cmt["chapter_path"],
                    "anchor_text": cmt["anchor_text"],
                    # Flatten multi-line comment text for CSV readability
                    "comment_text": (cmt["comment_text"] or "").replace("\n", " "),
                }
            )
    return rows


def export_to_csv(docx_paths: list, output_path: str) -> None:
    """Export comments from *docx_paths* to a semicolon-delimited *output_path*.

    The output file uses UTF-8 with BOM (utf-8-sig) so that Excel on
    Windows opens it without encoding issues.
    """
    print("[1/2] Extraction des commentaires...")
    rows = _build_rows(docx_paths)

    print(f"[2/2] Ecriture de {len(rows)} ligne(s) dans : {output_path}")
    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=CSV_FIELDS,
            delimiter=";",
            quoting=csv.QUOTE_ALL,
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nFichier CSV cree : {output_path}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    """Parse and return command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Exporte les commentaires d'un ou plusieurs fichiers Word (.docx) "
            "dans un fichier CSV (separateur : ';')."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Exemples :\n"
            "  python export_comments.py tests/fixtures/Doc1.docx\n"
            "  python export_comments.py tests/fixtures/Doc1.docx tests/fixtures/Doc2.docx --output comparaison.csv\n"
            "  python export_comments.py tests/fixtures/Doc1.docx tests/fixtures/Doc2.docx Doc3.docx\n"
        ),
    )
    parser.add_argument(
        "docx_files",
        nargs="+",
        metavar="FICHIER.docx",
        help="Un ou plusieurs fichiers Word (.docx)",
    )
    parser.add_argument(
        "--output",
        "-o",
        metavar="SORTIE.csv",
        default=None,
        help=(
            "Fichier CSV de sortie "
            "(defaut : <source>_comments.csv ou comments_export.csv)"
        ),
    )
    return parser.parse_args()


def main() -> None:
    """Entry point."""
    args = parse_args()
    output = args.output or _default_output(args.docx_files)
    export_to_csv(args.docx_files, output)


if __name__ == "__main__":
    main()
