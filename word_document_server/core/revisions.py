"""
Tracked-changes (revisions) support for WaDocx MCP.

python-docx has no native API for Word's tracked changes, so this module works
directly with the underlying OOXML. It provides the full revision lifecycle:

* authoring tracked insertions (``w:ins``) and deletions (``w:del``),
* a tracked find-and-replace (deletion + insertion),
* accepting and rejecting all changes,
* toggling the document's "track changes" setting.

Insertions wrap runs in ``<w:ins>``; deletions wrap runs in ``<w:del>`` and
convert their ``<w:t>`` text into ``<w:delText>`` as Word requires. Accept/reject
walk the document body (including tables) and resolve every revision element.
"""
import datetime
from typing import Optional

from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _now_iso() -> str:
    """Return a Word-compatible UTC timestamp (no microseconds)."""
    return datetime.datetime.now().replace(microsecond=0).isoformat() + "Z"


def _next_revision_id(doc) -> int:
    """Return the next free w:id for a revision element."""
    ids = [0]
    for el in doc.element.body.iter():
        tag = el.tag.split("}")[-1]
        if tag in ("ins", "del"):
            try:
                ids.append(int(el.get(qn("w:id"), "0")))
            except (TypeError, ValueError):
                pass
    return max(ids) + 1


def _clone_rpr(run_element):
    """Return a deep copy of a run's w:rPr element, or None."""
    rpr = run_element.find(qn("w:rPr"))
    if rpr is None:
        return None
    import copy

    return copy.deepcopy(rpr)


def _make_run(text: str, rpr_template=None, deleted: bool = False):
    """Build a ``w:r`` element holding ``text`` (as w:t or w:delText)."""
    run = OxmlElement("w:r")
    if rpr_template is not None:
        import copy

        run.append(copy.deepcopy(rpr_template))
    text_tag = "w:delText" if deleted else "w:t"
    text_el = OxmlElement(text_tag)
    text_el.set(qn("xml:space"), "preserve")
    text_el.text = text
    run.append(text_el)
    return run


def _wrap_revision(tag: str, runs, rev_id: int, author: str, date: str):
    """Wrap run elements in a ``w:ins`` or ``w:del`` revision container."""
    container = OxmlElement(f"w:{tag}")
    container.set(qn("w:id"), str(rev_id))
    container.set(qn("w:author"), author or "WaDocx")
    container.set(qn("w:date"), date or _now_iso())
    for run in runs:
        container.append(run)
    return container


def _iter_paragraph_elements(doc):
    """Yield all ``w:p`` elements in the body, including those inside tables."""
    for p in doc.element.body.iter(qn("w:p")):
        yield p


def tracked_replace_in_paragraph(
    p_element, old_text: str, new_text: str, rev_id_start: int,
    author: str, date: str,
) -> int:
    """Replace ``old_text`` with a tracked deletion + insertion in a paragraph.

    Returns the number of replacements performed. Formatting is taken from the
    first run of the paragraph and applied uniformly to the rebuilt segments.
    """
    runs = p_element.findall(qn("w:r"))
    if not runs:
        return 0
    full_text = ""
    for r in runs:
        for t in r.findall(qn("w:t")):
            full_text += t.text or ""
    if not old_text or old_text not in full_text:
        return 0

    rpr_template = _clone_rpr(runs[0])
    count = full_text.count(old_text)

    # Remove all existing content runs (keep w:pPr and any bookmarks intact by
    # only removing w:r elements).
    for r in runs:
        p_element.remove(r)

    # Find an insertion anchor: after w:pPr if present, else at the start.
    ppr = p_element.find(qn("w:pPr"))

    new_children = []
    segments = full_text.split(old_text)
    for idx, segment in enumerate(segments):
        if segment:
            new_children.append(_make_run(segment, rpr_template))
        if idx < len(segments) - 1:
            rid = rev_id_start + idx * 2
            del_run = _make_run(old_text, rpr_template, deleted=True)
            new_children.append(_wrap_revision("del", [del_run], rid, author, date))
            if new_text:
                ins_run = _make_run(new_text, rpr_template)
                new_children.append(
                    _wrap_revision("ins", [ins_run], rid + 1, author, date)
                )

    if ppr is not None:
        anchor = ppr
        for child in new_children:
            anchor.addnext(child)
            anchor = child
    else:
        for child in reversed(new_children):
            p_element.insert(0, child)

    return count


def add_tracked_insertion_paragraph(doc, text: str, author: str, date: str,
                                    style: Optional[str] = None):
    """Append a new paragraph whose text is a tracked insertion."""
    paragraph = doc.add_paragraph()
    if style:
        try:
            paragraph.style = style
        except KeyError:
            pass
    rev_id = _next_revision_id(doc)
    ins_run = _make_run(text, None)
    ins = _wrap_revision("ins", [ins_run], rev_id, author, date)
    paragraph._p.append(ins)
    return paragraph


def mark_paragraph_text_deleted(doc, old_text: str, author: str, date: str) -> int:
    """Mark every occurrence of ``old_text`` as a tracked deletion."""
    total = 0
    rev_id = _next_revision_id(doc)
    for p in _iter_paragraph_elements(doc):
        made = tracked_replace_in_paragraph(p, old_text, "", rev_id, author, date)
        if made:
            total += made
            rev_id = _next_revision_id(doc)
    return total


def tracked_replace(doc, old_text: str, new_text: str, author: str, date: str) -> int:
    """Tracked find-and-replace across the whole document body."""
    total = 0
    rev_id = _next_revision_id(doc)
    for p in _iter_paragraph_elements(doc):
        made = tracked_replace_in_paragraph(p, old_text, new_text, rev_id, author, date)
        if made:
            total += made
            rev_id = _next_revision_id(doc)
    return total


def _unwrap(element):
    """Replace ``element`` with its children, preserving order."""
    parent = element.getparent()
    if parent is None:
        return
    index = list(parent).index(element)
    for child in list(element):
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def accept_all_changes(doc) -> int:
    """Accept every tracked change: keep insertions, drop deletions."""
    resolved = 0
    body = doc.element.body
    changed = True
    while changed:
        changed = False
        for el in list(body.iter()):
            tag = el.tag.split("}")[-1]
            if tag == "ins":
                _unwrap(el)
                resolved += 1
                changed = True
                break
            if tag == "del":
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                    resolved += 1
                    changed = True
                    break
    return resolved


def reject_all_changes(doc) -> int:
    """Reject every tracked change: drop insertions, restore deletions."""
    resolved = 0
    body = doc.element.body
    changed = True
    while changed:
        changed = False
        for el in list(body.iter()):
            tag = el.tag.split("}")[-1]
            if tag == "ins":
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                    resolved += 1
                    changed = True
                    break
            if tag == "del":
                # Convert w:delText back to w:t, then unwrap.
                for del_text in el.iter(qn("w:delText")):
                    new_t = OxmlElement("w:t")
                    new_t.set(qn("xml:space"), "preserve")
                    new_t.text = del_text.text
                    del_text.getparent().replace(del_text, new_t)
                _unwrap(el)
                resolved += 1
                changed = True
                break
    return resolved


def set_track_changes(doc, enabled: bool) -> None:
    """Toggle Word's document-level 'track changes' setting."""
    settings = doc.settings.element
    existing = settings.find(qn("w:trackChanges"))
    if enabled:
        if existing is None:
            existing = OxmlElement("w:trackChanges")
            settings.append(existing)
    else:
        if existing is not None:
            settings.remove(existing)


def count_revisions(doc) -> dict:
    """Return counts of tracked insertions and deletions in the body."""
    ins = del_ = 0
    for el in doc.element.body.iter():
        tag = el.tag.split("}")[-1]
        if tag == "ins":
            ins += 1
        elif tag == "del":
            del_ += 1
    return {"insertions": ins, "deletions": del_}
