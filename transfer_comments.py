#!/usr/bin/env python3
"""Transfer Word document comments between two DOCX files.

Steps
-----
1. Copy the *content* document (Doc2) to the output file.
2. Identify comments in Doc2 that are absent from Doc1.
3. Remove all existing comments from the output file.
4. Re-insert all Doc1 comments and unique Doc2 comments, anchoring each
   to matching text in the correct chapter when possible.

Usage
-----
    python transfer_comments.py source.docx content.docx [output.docx]
    python transfer_comments.py --help
"""

import argparse
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy
from difflib import SequenceMatcher

try:
    from lxml import etree
except ImportError:
    import subprocess
    print("Installation de lxml en cours...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "lxml"])
    from lxml import etree  # pylint: disable=ungrouped-imports

# =============================================================================
# XML namespaces
# =============================================================================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
XML_NS = "http://www.w3.org/XML/1998/namespace"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# =============================================================================
# OOXML content-types and relationship types (module-level constants)
# =============================================================================
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.comments+xml"
)
COMMENTS_EXT_CT = (
    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.commentsExtended+xml"
)
COMMENTS_RT = (
    "http://schemas.openxmlformats.org/officeDocument"
    "/2006/relationships/comments"
)
COMMENTS_EXT_RT = (
    "http://schemas.microsoft.com/office/2011"
    "/relationships/commentsExtended"
)

# =============================================================================
# OOXML element tags  (Clark notation: {namespace}local)
# =============================================================================
T_BODY = f"{{{W_NS}}}body"
T_P = f"{{{W_NS}}}p"
T_R = f"{{{W_NS}}}r"
T_T = f"{{{W_NS}}}t"
T_PPR = f"{{{W_NS}}}pPr"
T_PSTYLE = f"{{{W_NS}}}pStyle"
T_RPR = f"{{{W_NS}}}rPr"
T_CRS = f"{{{W_NS}}}commentRangeStart"
T_CRE = f"{{{W_NS}}}commentRangeEnd"
T_CREF = f"{{{W_NS}}}commentReference"
T_CMT = f"{{{W_NS}}}comment"
T_CMT_EX = f"{{{W15_NS}}}commentEx"
T_CMTS_EX = f"{{{W15_NS}}}commentsEx"

# =============================================================================
# OOXML attribute names  (Clark notation)
# =============================================================================
A_ID = f"{{{W_NS}}}id"
A_VAL = f"{{{W_NS}}}val"
A_AUTHOR = f"{{{W_NS}}}author"
A_DATE = f"{{{W_NS}}}date"
A_INITIALS = f"{{{W_NS}}}initials"
A_PARA_ID = f"{{{W14_NS}}}paraId"
A_EX_PARA_ID = f"{{{W15_NS}}}paraId"
A_EX_DONE = f"{{{W15_NS}}}done"
A_SPACE = f"{{{XML_NS}}}space"

# Heading style pattern (multi-language)
_HEADING_RE = re.compile(
    r"^(Heading|heading|Titre|titre|Head|head|"
    r"Uberschrift|uberschrift|Titolo|titolo|Titulo|titulo)\s*(\d*)$"
)

# Default empty relationships XML (used when the file is absent)
_DEFAULT_RELS = (
    b'<Relationships xmlns="'
    b"http://schemas.openxmlformats.org/package/2006/relationships\"/>"
)


# =============================================================================
# Text utilities
# =============================================================================


def get_para_text(para) -> str:
    """Return the concatenated plain text of a <w:p> element."""
    return "".join(elem.text or "" for elem in para.iter(T_T))


def get_para_style(para) -> str:
    """Return the style identifier of a <w:p> (e.g. 'Heading1'), or ''."""
    para_props = para.find(T_PPR)
    if para_props is not None:
        style_elem = para_props.find(T_PSTYLE)
        if style_elem is not None:
            return style_elem.get(A_VAL, "")
    return ""


def is_heading_style(style: str) -> bool:
    """Return True when *style* matches a known heading style name."""
    return bool(_HEADING_RE.match(style))


def get_heading_level(style: str) -> int:
    """Return the numeric level of a heading style (defaults to 1)."""
    match = _HEADING_RE.match(style)
    if match:
        num = match.group(2)
        return int(num) if num.isdigit() else 1
    return 1


def compute_similarity(text_a: str, text_b: str) -> float:
    """Return the SequenceMatcher ratio between *text_a* and *text_b*."""
    return SequenceMatcher(None, text_a.lower(), text_b.lower()).ratio()


# =============================================================================
# Document structure
# =============================================================================


def build_doc_structure(doc_root) -> list:
    """Walk a document body and return a list of paragraph info dicts.

    Each dict has the keys:
        elem          - the <w:p> lxml element
        text          - plain-text content
        style         - paragraph style identifier
        chapter_path  - ' > '-joined heading hierarchy
        heading_stack - list of (level, title) tuples
    """
    body = doc_root.find(f".//{T_BODY}")
    if body is None:
        return []

    result = []
    heading_stack = []  # [(level: int, title: str), ...]

    def _add(para, chapter, stack):
        result.append(
            {
                "elem": para,
                "text": get_para_text(para),
                "style": get_para_style(para),
                "chapter_path": chapter,
                "heading_stack": list(stack),
            }
        )

    for child in body:
        if child.tag == T_P:
            text = get_para_text(child)
            style = get_para_style(child)
            if is_heading_style(style) and text.strip():
                level = get_heading_level(style)
                heading_stack = [
                    (lvl, t) for lvl, t in heading_stack if lvl < level
                ]
                heading_stack.append((level, text.strip()))
            chapter = " > ".join(t for _, t in heading_stack)
            _add(child, chapter, heading_stack)
        else:
            # Tables, SDT, etc. – collect all nested paragraphs
            chapter = " > ".join(t for _, t in heading_stack)
            for nested_para in child.iter(T_P):
                _add(nested_para, chapter, heading_stack)

    return result


# =============================================================================
# Comment extraction
# =============================================================================


def _get_anchor_text(para, comment_id: str) -> str:
    """Extract text between commentRangeStart/End for *comment_id* in *para*."""
    collecting = False
    parts = []
    for elem in para.iter():
        if elem.tag == T_CRS and elem.get(A_ID) == comment_id:
            collecting = True
        elif elem.tag == T_CRE and elem.get(A_ID) == comment_id:
            collecting = False
        elif collecting and elem.tag == T_T and elem.text:
            parts.append(elem.text)
    return "".join(parts)


def _build_resolved_set(cmts_root, raw_ext) -> set:
    """Return the set of comment IDs that are marked as resolved."""
    para_id_to_cid = {}
    for comment_elem in cmts_root.iter(T_CMT):
        cid = comment_elem.get(A_ID, "")
        for para in comment_elem.iter(T_P):
            pid = para.get(A_PARA_ID, "")
            if pid:
                para_id_to_cid[pid] = cid

    resolved = set()
    if raw_ext:
        ext_root = etree.fromstring(raw_ext)
        for ext_entry in ext_root.iter(T_CMT_EX):
            if ext_entry.get(A_EX_DONE, "0") == "1":
                pid = ext_entry.get(A_EX_PARA_ID, "")
                if pid in para_id_to_cid:
                    resolved.add(para_id_to_cid[pid])
    return resolved


def load_doc_comments(zip_path: str) -> tuple:
    """Load all comments from a DOCX file with anchor text and chapter context.

    Returns
    -------
    (comments_list, raw_comments_xml_bytes, raw_commentsExtended_xml_bytes)
    The last two elements may be None if the files are absent.
    """
    with zipfile.ZipFile(zip_path) as zfile:
        namelist = zfile.namelist()
        if "word/comments.xml" not in namelist:
            return [], None, None
        raw_cmts = zfile.read("word/comments.xml")
        raw_ext = (
            zfile.read("word/commentsExtended.xml")
            if "word/commentsExtended.xml" in namelist
            else None
        )
        doc_bytes = zfile.read("word/document.xml")

    cmts_root = etree.fromstring(raw_cmts)
    doc_root = etree.fromstring(doc_bytes)
    structure = build_doc_structure(doc_root)
    resolved = _build_resolved_set(cmts_root, raw_ext)

    comments = []
    for comment_elem in cmts_root.iter(T_CMT):
        cid = comment_elem.get(A_ID, "")
        cmt_text = "".join(e.text or "" for e in comment_elem.iter(T_T))
        anchor, chapter, h_stack = "", "", []

        for para_info in structure:
            for crs in para_info["elem"].iter(T_CRS):
                if crs.get(A_ID) == cid:
                    anchor = (
                        _get_anchor_text(para_info["elem"], cid) or para_info["text"]
                    )
                    chapter = para_info["chapter_path"]
                    h_stack = para_info["heading_stack"]
                    break
            if anchor or chapter:
                break

        comments.append(
            {
                "id": cid,
                "author": comment_elem.get(A_AUTHOR, ""),
                "date": comment_elem.get(A_DATE, ""),
                "initials": comment_elem.get(A_INITIALS, ""),
                "comment_text": cmt_text,
                "elem": deepcopy(comment_elem),
                "resolved": cid in resolved,
                "anchor_text": anchor,
                "chapter_path": chapter,
                "heading_stack": h_stack,
            }
        )
    return comments, raw_cmts, raw_ext


# =============================================================================
# Comment removal
# =============================================================================


def remove_comment_markers(doc_root) -> None:
    """Remove all comment range markers from *doc_root* in place."""
    for tag in (T_CRS, T_CRE):
        for elem in list(doc_root.iter(tag)):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

    for ref in list(doc_root.iter(T_CREF)):
        run = ref.getparent()
        if run is not None and run.tag == T_R:
            if not any(child.tag == T_T for child in run):
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)


# =============================================================================
# Comment insertion helpers
# =============================================================================


def _make_text_run(run_props, text: str):
    """Return a <w:r> element with optional run properties and a <w:t> child."""
    run = etree.Element(T_R)
    if run_props is not None:
        run.append(deepcopy(run_props))
    text_elem = etree.SubElement(run, T_T)
    text_elem.text = text
    if text and (text[0] == " " or text[-1] == " "):
        text_elem.set(A_SPACE, "preserve")
    return run


def _make_comment_ref_run(comment_id: str, run_props=None):
    """Return a <w:r><w:commentReference w:id="…"/></w:r> element."""
    run = etree.Element(T_R)
    if run_props is not None:
        run.append(deepcopy(run_props))
    ref = etree.SubElement(run, T_CREF)
    ref.set(A_ID, comment_id)
    return run


def anchor_comment_in_para(para, comment_id: str, anchor_text: str) -> bool:
    """Insert comment range markers into *para*.

    If *anchor_text* is found within a single run the comment is anchored
    precisely to that phrase.  Otherwise the whole paragraph is wrapped.

    Returns True for an exact phrase anchor, False for a paragraph-level wrap.
    """
    if anchor_text and anchor_text.strip():
        for child in list(para):
            if child.tag != T_R:
                continue
            run_text = "".join(e.text or "" for e in child.iter(T_T))
            if anchor_text not in run_text:
                continue
            run_props = child.find(T_RPR)
            insert_pos = list(para).index(child)
            before = run_text[: run_text.index(anchor_text)]
            after = run_text[run_text.index(anchor_text) + len(anchor_text) :]

            new_elems = []
            if before:
                new_elems.append(_make_text_run(run_props, before))
            crs = etree.Element(T_CRS)
            crs.set(A_ID, comment_id)
            new_elems.append(crs)
            new_elems.append(_make_text_run(run_props, anchor_text))
            cre = etree.Element(T_CRE)
            cre.set(A_ID, comment_id)
            new_elems.append(cre)
            new_elems.append(_make_comment_ref_run(comment_id, run_props))
            if after:
                new_elems.append(_make_text_run(run_props, after))

            para.remove(child)
            for offset, elem in enumerate(new_elems):
                para.insert(insert_pos + offset, elem)
            return True

    # Fallback: wrap all runs in the paragraph
    runs = [child for child in para if child.tag == T_R]
    if not runs:
        return False

    idx_first = list(para).index(runs[0])
    idx_last = list(para).index(runs[-1])
    crs = etree.Element(T_CRS)
    crs.set(A_ID, comment_id)
    para.insert(idx_first, crs)
    idx_last += 1  # shifted by the insertion above
    cre = etree.Element(T_CRE)
    cre.set(A_ID, comment_id)
    para.insert(idx_last + 1, cre)
    para.insert(idx_last + 2, _make_comment_ref_run(comment_id))
    return False


# =============================================================================
# Paragraph matching
# =============================================================================


def find_best_matches(
    anchor: str,
    chapter: str,
    structure: list,
    threshold: float = 0.75,
) -> list:
    """Return candidate paragraphs sorted by match score (descending).

    Each entry is a (para_info, score, match_type) tuple.
    """
    results = []
    for para_info in structure:
        text = para_info["text"].strip()
        if not text:
            continue
        same_chapter = para_info["chapter_path"] == chapter
        if anchor and anchor.strip():
            if anchor in para_info["text"]:
                score = 1.0 if same_chapter else 0.85
                mtype = (
                    "exact_meme_chapitre" if same_chapter else "exact_autre_chapitre"
                )
                results.append((para_info, score, mtype))
                continue
            ratio = compute_similarity(anchor, text)
            if ratio >= threshold:
                score = min(ratio * (1.1 if same_chapter else 0.9), 1.0)
                mtype = (
                    "approx_meme_chapitre" if same_chapter else "approx_autre_chapitre"
                )
                results.append((para_info, score, mtype))
        elif same_chapter:
            results.append((para_info, 0.4, "meme_chapitre_sans_ancre"))

    results.sort(key=lambda x: x[1], reverse=True)
    return results


def find_chapter_first_para(chapter: str, structure: list):
    """Return the first paragraph belonging to *chapter*, or None."""
    for para_info in structure:
        if para_info["chapter_path"] == chapter:
            return para_info
    if chapter:
        last_title = chapter.split(" > ")[-1]
        for para_info in structure:
            if (
                para_info["heading_stack"]
                and para_info["heading_stack"][-1][1] == last_title
            ):
                return para_info
    return None


# =============================================================================
# Comment comparison (Doc1 vs Doc2)
# =============================================================================


def compare_comment(candidate: dict, references: list) -> tuple:
    """Find the closest match for *candidate* in *references*.

    Comparison is based on comment-text similarity with bonuses for matching
    author and date.

    Returns
    -------
    (best_reference, score)  or  (None, 0.0)
    """
    best = None
    best_score = 0.0
    for ref in references:
        cand_text = candidate["comment_text"] or ""
        ref_text = ref["comment_text"] or ""
        if cand_text and ref_text:
            text_sim = compute_similarity(cand_text, ref_text)
        else:
            text_sim = 1.0 if not cand_text and not ref_text else 0.0
        score = text_sim
        if candidate["author"] == ref["author"]:
            score = min(score * 1.05, 1.0)
        if candidate["date"] == ref["date"]:
            score = min(score * 1.05, 1.0)
        if score > best_score:
            best_score = score
            best = ref
    return best, best_score


# =============================================================================
# User interaction
# =============================================================================


def ask_numbered_choice(question: str, choices: list) -> int:
    """Display numbered choices and wait for user input.

    Returns the 0-based index of the chosen item, or -1 to skip.
    """
    print(f"\n{question}")
    for idx, (label, _) in enumerate(choices):
        print(f"  [{idx + 1}] {label}")
    print("  [0] Ignorer ce commentaire")
    while True:
        raw = input("Votre choix > ").strip()
        if raw == "0":
            return -1
        try:
            num = int(raw)
            if 1 <= num <= len(choices):
                return num - 1
        except ValueError:
            pass
        print("  Choix invalide, entrez un numero.")


def ask_yes_no(question: str) -> bool:
    """Ask a yes/no question in French and return True for yes."""
    print(f"\n{question} (o/n)")
    while True:
        raw = input("> ").strip().lower()
        if raw in ("o", "oui", "y", "yes"):
            return True
        if raw in ("n", "non", "no"):
            return False
        print("  Repondez o ou n.")


# =============================================================================
# XML helpers  (content types and relationships)
# =============================================================================


def add_content_type_override(ct_root, part_name: str, content_type: str) -> None:
    """Add an <Override> entry to [Content_Types].xml when absent."""
    for override in ct_root.findall(f"{{{CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            return
    override = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
    override.set("PartName", part_name)
    override.set("ContentType", content_type)


def remove_content_type_override(ct_root, part_name: str) -> None:
    """Remove the <Override> entry for *part_name* from [Content_Types].xml."""
    for override in list(ct_root.findall(f"{{{CT_NS}}}Override")):
        if override.get("PartName") == part_name:
            ct_root.remove(override)


def add_relationship(rels_root, target: str, rel_type: str) -> str:
    """Add a relationship to document.xml.rels when absent. Returns its Id."""
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Target") == target:
            return rel.get("Id", "")
    existing_ids = {
        r.get("Id", "")
        for r in rels_root.findall(f"{{{REL_NS}}}Relationship")
    }
    idx = 1
    while f"rId{idx}" in existing_ids:
        idx += 1
    new_id = f"rId{idx}"
    rel = etree.SubElement(rels_root, f"{{{REL_NS}}}Relationship")
    rel.set("Id", new_id)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    return new_id


def remove_relationship_for(rels_root, target: str) -> None:
    """Remove the relationship pointing to *target*."""
    for rel in list(rels_root.findall(f"{{{REL_NS}}}Relationship")):
        if rel.get("Target") == target:
            rels_root.remove(rel)


# =============================================================================
# Step-2 helpers: identify unique Doc2 comments
# =============================================================================


def _renumber_doc2_ids(doc2_only: list, doc1_comments: list) -> None:
    """Assign non-conflicting IDs to *doc2_only* comments (mutates in place)."""
    doc1_ids = set()
    for cmt in doc1_comments:
        try:
            doc1_ids.add(int(cmt["id"]))
        except ValueError:
            pass
    next_id = (max(doc1_ids) + 1) if doc1_ids else 0
    for cmt in doc2_only:
        cmt["elem"].set(A_ID, str(next_id))
        cmt["id"] = str(next_id)
        next_id += 1


def _prompt_ambiguous_doc2(
    d2_cmt: dict,
    match,
    score: float,
    doc2_only: list,
) -> None:
    """Prompt the user about an ambiguous Doc2 comment and update doc2_only."""
    date_str = (d2_cmt["date"] or "")[:10]
    preview = (d2_cmt["comment_text"] or "")[:55].replace("\n", " ")
    resolved = " [RESOLU]" if d2_cmt["resolved"] else ""
    print(f"  [?] Commentaire Doc2 ambigu{resolved} :")
    print(f"       Auteur : {d2_cmt['author']} ({date_str})")
    print(f'       Texte  : "{preview}"')
    if match:
        mp = (match["comment_text"] or "")[:55].replace("\n", " ")
        print(f"       Meilleure corresp. Doc1 ({score:.0%}) : \"{mp}\"")
    if ask_yes_no("Ce commentaire Doc2 existe deja dans Doc1 (ne pas conserver) ?"):
        print("       -> Ignore (present dans Doc1).\n")
    else:
        doc2_only.append(d2_cmt)
        print("       -> Marque comme unique a Doc2, sera conserve.\n")


def _classify_doc2_comment(
    d2_cmt: dict, doc1_comments: list, doc2_only: list
) -> None:
    """Classify a single Doc2 comment and append to doc2_only when unique."""
    match, score = compare_comment(d2_cmt, doc1_comments)
    date_str = (d2_cmt["date"] or "")[:10]
    preview = (d2_cmt["comment_text"] or "")[:55].replace("\n", " ")
    resolved = " [RESOLU]" if d2_cmt["resolved"] else ""

    if score >= 0.90:
        print(f'  [=] #{d2_cmt["id"]}{resolved} "{preview}"')
        print(f"       -> Correspond a Doc1 ({score:.0%}), pas besoin de conserver.\n")
    elif score >= 0.60:
        _prompt_ambiguous_doc2(d2_cmt, match, score, doc2_only)
    else:
        print(f'  [+] #{d2_cmt["id"]}{resolved} "{preview}"')
        print(f"       Auteur : {d2_cmt['author']} ({date_str})")
        print("       -> Unique a Doc2, sera conserve.\n")
        doc2_only.append(d2_cmt)


def identify_doc2_unique_comments(doc1_comments: list, doc2_path: str) -> tuple:
    """Identify comments in *doc2_path* absent from *doc1_comments*.

    Ambiguous cases (similarity 60-90%) trigger a user prompt.

    Returns
    -------
    (doc2_only_list, raw_commentsExtended_bytes_or_None)
    """
    doc2_comments, _raw, doc2_raw_ext = load_doc_comments(doc2_path)
    print(f"      {len(doc2_comments)} commentaire(s) dans Doc2.\n")
    print("[2/4] Analyse des commentaires de Doc2 (recherche des nouveaux)...\n")

    doc2_only = []
    for d2_cmt in doc2_comments:
        _classify_doc2_comment(d2_cmt, doc1_comments, doc2_only)

    if doc2_only:
        print(f"  => {len(doc2_only)} commentaire(s) unique(s) a Doc2 a conserver.\n")
    else:
        print("  => Aucun commentaire unique a Doc2.\n")

    _renumber_doc2_ids(doc2_only, doc1_comments)
    return doc2_only, doc2_raw_ext


# =============================================================================
# XML merge helpers
# =============================================================================


def build_combined_comments_xml(raw_cmts_xml, doc2_only: list) -> bytes:
    """Return merged comments.xml bytes: Doc1 comments + *doc2_only*."""
    if raw_cmts_xml is None:
        root = etree.Element(f"{{{W_NS}}}comments", nsmap={"w": W_NS})
    else:
        root = etree.fromstring(raw_cmts_xml)
    for cmt in doc2_only:
        root.append(cmt["elem"])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_combined_ext_xml(raw_ext_xml, doc2_only: list, doc2_raw_ext) -> bytes | None:
    """Return merged commentsExtended.xml bytes, or *raw_ext_xml* unchanged.

    Every Doc2-only comment needs its own <w15:commentEx> entry carried over,
    not just resolved ones: that entry is also where the reply parent/child
    link (w15:paraIdParent) lives, so filtering by resolved status alone
    silently drops thread relationships for unresolved replies.
    """
    if not doc2_only or not doc2_raw_ext:
        return raw_ext_xml

    doc2_only_para_ids = set()
    for cmt in doc2_only:
        for para in cmt["elem"].iter(T_P):
            pid = para.get(A_PARA_ID, "")
            if pid:
                doc2_only_para_ids.add(pid)

    if not doc2_only_para_ids:
        return raw_ext_xml

    ext_base = (
        etree.fromstring(raw_ext_xml)
        if raw_ext_xml
        else etree.Element(T_CMTS_EX, nsmap={"w15": W15_NS, "w14": W14_NS})
    )
    for ext_entry in etree.fromstring(doc2_raw_ext).iter(T_CMT_EX):
        if ext_entry.get(A_EX_PARA_ID, "") in doc2_only_para_ids:
            ext_base.append(deepcopy(ext_entry))
    return etree.tostring(
        ext_base, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# =============================================================================
# Insertion loop – target selection helpers
# =============================================================================


def _handle_no_match(cmt: dict, chapter: str, structure: list):
    """Handle a comment with zero text matches. Returns (para_info, anchor) or None."""
    print("  |   Aucune correspondance trouvee.")
    ch_para = find_chapter_first_para(chapter, structure)
    if ch_para and chapter:
        if ask_yes_no(f'Ancrer au debut du chapitre "{chapter}" ?'):
            return ch_para, ""
        return None
    if ask_yes_no(f'Aucune cible pour #{cmt["id"]}. Ignorer ?'):
        return None
    return None


def _handle_medium_match(cmt: dict, matches: list, anchor: str):
    """Prompt user for a medium-confidence match. Returns (para_info, anchor) or None."""
    best_score = matches[0][1]
    print(f"  |   Confiance moyenne ({best_score:.0%}). Intervention necessaire.")
    choices = [
        (
            f'[{s:.0%}] "{pi["text"][:65]}" (ch: {pi["chapter_path"] or "--"})',
            pi,
        )
        for pi, s, _ in matches[:5]
    ]
    idx = ask_numbered_choice(
        f'Commentaire #{cmt["id"]} : choisissez le paragraphe cible :', choices
    )
    if idx < 0:
        return None
    chosen = choices[idx][1]
    print(f'  |   -> Choix manuel : "{chosen["text"][:65]}"')
    return chosen, anchor


def _handle_low_match(
    cmt: dict, matches: list, anchor: str, chapter: str, structure: list
):
    """Prompt user for a low-confidence match. Returns (para_info, anchor) or None."""
    print(f"  |   Texte non trouve avec precision (score max: {matches[0][1]:.0%}).")
    ch_para = find_chapter_first_para(chapter, structure)
    if ch_para and chapter:
        if ask_yes_no(f'Ancrer au debut du chapitre "{chapter}" ?'):
            return ch_para, ""
    choices = [
        (f'[{s:.0%}] "{pi["text"][:65]}"', pi) for pi, s, _ in matches[:5]
    ]
    idx = ask_numbered_choice(
        f'Choisir manuellement pour commentaire #{cmt["id"]} :', choices
    )
    if idx < 0:
        return None
    return choices[idx][1], anchor


def _choose_target_para(cmt: dict, structure: list, threshold: float):
    """Find or interactively select the target paragraph for a comment.

    Returns (para_info, anchor_text) or None if the comment should be skipped.
    """
    anchor = cmt["anchor_text"]
    chapter = cmt["chapter_path"]
    matches = find_best_matches(anchor, chapter, structure, threshold)

    if not matches:
        return _handle_no_match(cmt, chapter, structure)

    best_pi, best_score, best_type = matches[0]
    if best_score >= 0.90:
        print(f"  |   -> Ancrage auto ({best_type}, {best_score:.0%}):")
        print(f'  |      "{best_pi["text"][:65]}"')
        return best_pi, anchor
    if best_score >= 0.70:
        return _handle_medium_match(cmt, matches, anchor)
    return _handle_low_match(cmt, matches, anchor, chapter, structure)


# =============================================================================
# Insertion loop
# =============================================================================


def run_insertion_loop(
    all_comments: list, structure: list, threshold: float
) -> tuple:
    """Insert all comments into the document structure interactively.

    Returns
    -------
    (inserted_count, skipped_details_list)
    """
    inserted = 0
    skipped_list = []

    for cmt in all_comments:
        cid = cmt["id"]
        date_str = (cmt["date"] or "")[:10]
        preview = (cmt["comment_text"] or "")[:70].replace("\n", " ")
        resolved = " [RESOLU]" if cmt["resolved"] else ""

        print(f'  +-- Commentaire #{cid}{resolved}')
        print(f'  |   Auteur   : {cmt["author"]} ({date_str})')
        print(f'  |   Texte    : "{preview}"')
        if cmt["anchor_text"]:
            print(f'  |   Ancre    : "{cmt["anchor_text"][:80]}"')
        if cmt["chapter_path"]:
            print(f'  |   Chapitre : "{cmt["chapter_path"]}"')

        result = _choose_target_para(cmt, structure, threshold)
        if result is None:
            print("  +-- Ignore.\n")
            skipped_list.append(
                (
                    cid,
                    cmt["author"],
                    date_str,
                    cmt["comment_text"],
                    cmt["anchor_text"],
                    cmt["chapter_path"],
                    "ignore par utilisateur",
                )
            )
            continue

        chosen_pi, anchor_used = result
        exact = anchor_comment_in_para(chosen_pi["elem"], cid, anchor_used)
        kind = "ancre precise" if exact else "ancre paragraphe entier"
        print(f"  +-- Insere ({kind}).\n")
        inserted += 1

    return inserted, skipped_list


# =============================================================================
# ZIP output writer
# =============================================================================


def write_output_docx(
    all_files: dict,
    output: str,
    new_doc_bytes: bytes,
    new_cmts_bytes: bytes,
    final_ext_bytes,
) -> None:
    """Rewrite the output DOCX ZIP with the updated XML files."""
    ct_root = etree.fromstring(all_files["[Content_Types].xml"])
    rels_root = etree.fromstring(
        all_files.get("word/_rels/document.xml.rels", _DEFAULT_RELS)
    )

    add_content_type_override(ct_root, "/word/comments.xml", COMMENTS_CT)
    if final_ext_bytes:
        add_content_type_override(
            ct_root, "/word/commentsExtended.xml", COMMENTS_EXT_CT
        )
    else:
        remove_content_type_override(ct_root, "/word/commentsExtended.xml")

    add_relationship(rels_root, "comments.xml", COMMENTS_RT)
    if final_ext_bytes:
        add_relationship(rels_root, "commentsExtended.xml", COMMENTS_EXT_RT)
    else:
        remove_relationship_for(rels_root, "commentsExtended.xml")

    new_ct_bytes = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    new_rels_bytes = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            if name == "word/document.xml":
                zout.writestr(name, new_doc_bytes)
            elif name == "word/comments.xml":
                zout.writestr(name, new_cmts_bytes)
            elif name == "word/commentsExtended.xml":
                if final_ext_bytes:
                    zout.writestr(name, final_ext_bytes)
                # absent from output when final_ext_bytes is None
            elif name == "[Content_Types].xml":
                zout.writestr(name, new_ct_bytes)
            elif name == "word/_rels/document.xml.rels":
                zout.writestr(name, new_rels_bytes)
            else:
                zout.writestr(name, data)

        if "word/comments.xml" not in all_files:
            zout.writestr("word/comments.xml", new_cmts_bytes)
        if final_ext_bytes and "word/commentsExtended.xml" not in all_files:
            zout.writestr("word/commentsExtended.xml", final_ext_bytes)
        if "word/_rels/document.xml.rels" not in all_files:
            zout.writestr("word/_rels/document.xml.rels", new_rels_bytes)


# =============================================================================
# Summary report
# =============================================================================


def _print_skipped_details(skipped_list: list) -> None:
    """Print per-comment details for every skipped entry."""
    if not skipped_list:
        return
    print()
    print("  Liste des commentaires ignores :")
    for cid, author, date_str, cmt_text, anchor, chapter, reason in skipped_list:
        preview = (cmt_text or "")[:60].replace("\n", " ")
        print(f"    #{cid} | {author} ({date_str}) | raison: {reason}")
        if anchor:
            print(f'         ancre    : "{anchor[:70]}"')
        if chapter:
            print(f'         chapitre : "{chapter}"')
        print(f'         texte    : "{preview}"')


def print_report(
    sep: str,
    doc1_comments: list,
    doc2_only: list,
    inserted: int,
    skipped_list: list,
    output: str,
) -> None:
    """Print the final transfer summary to stdout."""
    print(sep)
    print("  Resume du traitement")
    print(f"  Commentaires Doc1 inseres    : {len(doc1_comments)}")
    print(f"  Commentaires Doc2 conserves  : {len(doc2_only)}")
    print(f"  Effectivement inseres        : {inserted}")
    print(f"  Ignores                      : {len(skipped_list)}")
    print(f"  Fichier cree                 : {output}")
    _print_skipped_details(skipped_list)
    print(f"{sep}\n")


# =============================================================================
# Main transfer orchestrator
# =============================================================================


def transfer(
    doc1: str,
    doc2: str,
    output: str,
    threshold: float = 0.75,
) -> None:
    """Transfer comments from *doc1* to a copy of *doc2*, writing to *output*.

    Parameters
    ----------
    doc1:       Path to the source document (comments to transfer).
    doc2:       Path to the content document (text to keep).
    output:     Path of the output document to create.
    threshold:  Minimum similarity score for automatic text matching (0–1).
    """
    sep = "=" * 62
    print(f"\n{sep}")
    print("  Transfert de commentaires Word")
    print(f"  Source (commentaires) : {doc1}")
    print(f"  Contenu               : {doc2}")
    print(f"  Document final        : {output}")
    print(f"{sep}\n")

    if os.path.exists(output):
        if not ask_yes_no(f'Le fichier "{output}" existe deja. L\'ecraser ?'):
            print("Annule.")
            sys.exit(0)

    # Step 1: copy doc2 -> output
    shutil.copy2(doc2, output)
    print("[1/4] Copie de Doc2 vers Doc_Final... OK\n")

    doc1_comments, raw_cmts_xml, raw_ext_xml = load_doc_comments(doc1)
    print(f"      {len(doc1_comments)} commentaire(s) charge(s) depuis Doc1.")

    with zipfile.ZipFile(output, "r") as zfile:
        all_files = {name: zfile.read(name) for name in zfile.namelist()}

    # Step 2: identify unique Doc2 comments
    doc2_only, doc2_raw_ext = identify_doc2_unique_comments(doc1_comments, doc2)

    # Step 3: remove existing comment markers from output
    doc_root = etree.fromstring(all_files["word/document.xml"])
    remove_comment_markers(doc_root)
    print("[3/4] Commentaires existants supprimes de Doc_Final... OK\n")

    structure = build_doc_structure(doc_root)

    # Step 4: insert all comments
    all_comments = doc1_comments + doc2_only
    if not all_comments:
        print("Aucun commentaire a inserer. Fin du traitement.")
        return

    print(
        f"[4/4] Insertion de {len(all_comments)} commentaire(s) "
        f"(Doc1: {len(doc1_comments)}, Doc2-unique: {len(doc2_only)})...\n"
    )
    inserted, skipped_list = run_insertion_loop(all_comments, structure, threshold)

    new_doc_bytes = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    new_cmts_bytes = build_combined_comments_xml(raw_cmts_xml, doc2_only)
    final_ext_bytes = build_combined_ext_xml(raw_ext_xml, doc2_only, doc2_raw_ext)

    write_output_docx(all_files, output, new_doc_bytes, new_cmts_bytes, final_ext_bytes)
    print_report(sep, doc1_comments, doc2_only, inserted, skipped_list, output)


# =============================================================================
# CLI
# =============================================================================


def parse_args() -> argparse.Namespace:
    """Parse and return command-line arguments."""
    parser = argparse.ArgumentParser(
        prog="transfer_comments",
        description="Transfert de commentaires entre deux documents Word (.docx).",
        epilog=(
            "Exemple : python transfer_comments.py tests/fixtures/Doc1.docx tests/fixtures/Doc2.docx Doc_Final.docx"
        ),
    )
    parser.add_argument(
        "source",
        help="Document source contenant les commentaires a transferer",
    )
    parser.add_argument(
        "content",
        help="Document contenu dont le texte sera conserve dans le resultat",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default="Doc_Final.docx",
        help="Document de sortie (defaut: Doc_Final.docx)",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.75,
        metavar="FLOAT",
        help="Seuil de similarite pour la correspondance de texte (defaut: 0.75)",
    )
    return parser.parse_args()


def main() -> None:
    """Entry point: parse CLI arguments and run the comment transfer."""
    args = parse_args()
    for path in (args.source, args.content):
        if not os.path.exists(path):
            sys.exit(f"Erreur: fichier introuvable: {path}")
    transfer(args.source, args.content, args.output, args.threshold)


if __name__ == "__main__":
    main()
