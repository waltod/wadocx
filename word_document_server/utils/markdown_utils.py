"""
Markdown import/export utilities for WaDocx MCP.
"""
import os
import re
from typing import Any, Dict, List, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table

from word_document_server.utils.document_utils import (
    delete_block_under_header,
    get_body_elements,
    get_paragraph_from_element,
    is_paragraph_element,
    is_table_element,
    insert_content_blocks_after_element,
    normalize_paragraph_text,
)


HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
ORDERED_LIST_RE = re.compile(r"^\s*(\d+)\.\s+(.*)$")
BULLET_LIST_RE = re.compile(r"^\s*[-*+]\s+(.*)$")


def _is_markdown_table_line(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 2


def _split_markdown_row(line: str) -> List[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _is_table_separator_row(cells: List[str]) -> bool:
    if not cells:
        return False
    for cell in cells:
        normalized = cell.replace(":", "").replace("-", "").strip()
        if normalized:
            return False
    return True


def parse_markdown_blocks(markdown_text: str) -> List[Dict[str, Any]]:
    """Parse a markdown string into simple document blocks."""
    blocks: List[Dict[str, Any]] = []
    paragraph_lines: List[str] = []
    code_lines: Optional[List[str]] = None
    lines = markdown_text.splitlines()
    i = 0

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        if paragraph_lines:
            text = " ".join(line.strip() for line in paragraph_lines if line.strip())
            if text:
                blocks.append({"type": "paragraph", "text": text})
            paragraph_lines = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith("```"):
            if code_lines is None:
                flush_paragraph()
                code_lines = []
            else:
                blocks.append({"type": "paragraph", "text": "\n".join(code_lines)})
                code_lines = None
            i += 1
            continue

        if code_lines is not None:
            code_lines.append(line)
            i += 1
            continue

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        heading_match = HEADING_RE.match(stripped)
        if heading_match:
            flush_paragraph()
            blocks.append(
                {
                    "type": "heading",
                    "level": len(heading_match.group(1)),
                    "text": heading_match.group(2).strip(),
                }
            )
            i += 1
            continue

        if _is_markdown_table_line(stripped):
            flush_paragraph()
            table_lines = []
            while i < len(lines) and _is_markdown_table_line(lines[i].strip()):
                table_lines.append(lines[i].strip())
                i += 1

            rows = [_split_markdown_row(table_line) for table_line in table_lines]
            if len(rows) > 1 and _is_table_separator_row(rows[1]):
                rows.pop(1)
            if rows:
                blocks.append({"type": "table", "rows": rows})
            continue

        ordered_match = ORDERED_LIST_RE.match(line)
        bullet_match = BULLET_LIST_RE.match(line)
        if ordered_match or bullet_match:
            flush_paragraph()
            ordered = ordered_match is not None
            items = []
            while i < len(lines):
                current_line = lines[i]
                current_match = ORDERED_LIST_RE.match(current_line) if ordered else BULLET_LIST_RE.match(current_line)
                if not current_match:
                    break
                items.append(current_match.group(2 if ordered else 1).strip())
                i += 1
            blocks.append({"type": "list", "ordered": ordered, "items": items})
            continue

        paragraph_lines.append(line)
        i += 1

    flush_paragraph()
    if code_lines:
        blocks.append({"type": "paragraph", "text": "\n".join(code_lines)})
    return blocks


def _paragraph_list_kind(para) -> Optional[str]:
    """Infer whether a paragraph belongs to a bullet or numbered list."""
    if para.style:
        style_name = para.style.name.lower()
        if "list number" in style_name or style_name == "numbered list":
            return "ordered"
        if "list bullet" in style_name or style_name == "bullet list":
            return "unordered"

    p_pr = para._element.find(qn("w:pPr"))
    if p_pr is None:
        return None
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is None:
        return None
    num_id = num_pr.find(qn("w:numId"))
    if num_id is not None and num_id.get(qn("w:val")) == "2":
        return "ordered"
    return "unordered"


def document_to_markdown(doc_path: str) -> str:
    """Export a Word document to a simple markdown representation."""
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"

    doc = Document(doc_path)
    blocks: List[Dict[str, Any]] = []

    for el in get_body_elements(doc):
        if is_paragraph_element(el):
            para = get_paragraph_from_element(doc, el)
            if para is None:
                continue
            text = para.text.rstrip()
            if not text:
                continue

            style_name = para.style.name if para.style else ""
            if style_name.startswith("Heading "):
                try:
                    level = int(style_name.split(" ")[1])
                except (IndexError, ValueError):
                    level = 1
                blocks.append({"type": "heading", "level": level, "text": text})
                continue

            list_kind = _paragraph_list_kind(para)
            if list_kind:
                ordered = list_kind == "ordered"
                if blocks and blocks[-1]["type"] == "list" and blocks[-1]["ordered"] == ordered:
                    blocks[-1]["items"].append(text)
                else:
                    blocks.append({"type": "list", "ordered": ordered, "items": [text]})
                continue

            blocks.append({"type": "paragraph", "text": text})
            continue

        if is_table_element(el):
            table = Table(el, doc._body)
            rows = []
            for row in table.rows:
                rows.append([normalize_paragraph_text(cell.text) for cell in row.cells])
            if rows:
                blocks.append({"type": "table", "rows": rows})

    rendered_blocks: List[str] = []
    for block in blocks:
        if block["type"] == "heading":
            rendered_blocks.append(f"{'#' * block['level']} {block['text']}")
            continue

        if block["type"] == "list":
            prefix = "1." if block["ordered"] else "-"
            lines = []
            for index, item in enumerate(block["items"], start=1):
                marker = f"{index}." if block["ordered"] else prefix
                lines.append(f"{marker} {item}")
            rendered_blocks.append("\n".join(lines))
            continue

        if block["type"] == "table":
            rows = block["rows"]
            if not rows:
                continue
            width = max(len(row) for row in rows)

            def pad(row: List[str]) -> List[str]:
                return row + [""] * (width - len(row))

            header = pad(rows[0])
            lines = [
                f"| {' | '.join(header)} |",
                f"| {' | '.join(['---'] * width)} |",
            ]
            for row in rows[1:]:
                lines.append(f"| {' | '.join(pad(row))} |")
            rendered_blocks.append("\n".join(lines))
            continue

        rendered_blocks.append(block["text"])

    return "\n\n".join(rendered_blocks)


def export_document_markdown(doc_path: str, output_path: Optional[str] = None) -> str:
    """Export a document to markdown on disk and return the output path."""
    markdown = document_to_markdown(doc_path)
    if markdown.startswith("Document ") and markdown.endswith(" does not exist"):
        return markdown

    if output_path is None:
        base_name, _ = os.path.splitext(doc_path)
        output_path = f"{base_name}.md"

    with open(output_path, "w", encoding="utf-8") as markdown_file:
        markdown_file.write(markdown)

    return output_path


def clear_document_body(doc) -> None:
    """Remove all body elements except section properties."""
    body = doc.element.body
    for child in list(body.iterchildren()):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def replace_document_with_markdown(doc_path: str, markdown_text: str) -> Dict[str, Any]:
    """Replace the body of a document with parsed markdown blocks."""
    doc = Document(doc_path) if os.path.exists(doc_path) else Document()
    clear_document_body(doc)

    blocks = parse_markdown_blocks(markdown_text)
    anchor = doc.add_paragraph("")
    inserted = insert_content_blocks_after_element(doc, anchor._element, blocks)
    anchor._element.getparent().remove(anchor._element)
    doc.save(doc_path)

    return {"inserted_blocks": inserted, "block_count": len(blocks)}


def replace_section_with_markdown(doc_path: str, header_text: str, markdown_text: str) -> Dict[str, Any]:
    """Replace a section body with parsed markdown blocks."""
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    doc = Document(doc_path)
    header_el, removed_count = delete_block_under_header(doc, header_text)
    if header_el is None:
        return {"error": f"Header '{header_text}' not found in document."}

    blocks = parse_markdown_blocks(markdown_text)
    inserted = insert_content_blocks_after_element(doc, header_el, blocks)
    doc.save(doc_path)
    return {"inserted_blocks": inserted, "removed_blocks": removed_count, "block_count": len(blocks)}

