import asyncio
from pathlib import Path

from docx import Document

from word_document_server.tools.format_tools import format_text
from word_document_server.tools.markdown_tools import (
    get_document_markdown,
    replace_document_with_markdown,
    replace_section_with_markdown,
)
from word_document_server.utils.document_utils import (
    replace_block_between_manual_anchors,
    replace_paragraph_block_below_header,
)


def _paragraph_texts(doc_path: Path):
    return [paragraph.text for paragraph in Document(doc_path).paragraphs if paragraph.text]


def test_replace_paragraph_block_below_header_is_stable_across_iterations(tmp_path: Path):
    doc_path = tmp_path / "iterative-header.docx"
    doc = Document()
    doc.add_heading("Scope", level=1)
    doc.add_paragraph("Old clause 1")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Legacy"
    table.cell(0, 1).text = "Block"
    table.cell(1, 0).text = "Should"
    table.cell(1, 1).text = "Go"
    doc.add_paragraph("Old clause 2")
    doc.add_heading("Next Section", level=1)
    doc.add_paragraph("Keep me")
    doc.save(doc_path)

    for iteration in range(6):
        message = replace_paragraph_block_below_header(
            str(doc_path),
            "Scope",
            [f"Clause revision {iteration}", f"Subclause revision {iteration}"],
        )
        assert "Replaced content under 'Scope'" in message

    final_doc = Document(doc_path)
    assert [paragraph.text for paragraph in final_doc.paragraphs if paragraph.text] == [
        "Scope",
        "Clause revision 5",
        "Subclause revision 5",
        "Next Section",
        "Keep me",
    ]
    assert len(final_doc.tables) == 0


def test_replace_block_between_manual_anchors_is_stable_across_iterations(tmp_path: Path):
    doc_path = tmp_path / "iterative-anchors.docx"
    doc = Document()
    doc.add_paragraph("START")
    doc.add_paragraph("Legacy paragraph")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Legacy table"
    doc.add_paragraph("Another legacy paragraph")
    doc.add_paragraph("END")
    doc.add_paragraph("Trailing content")
    doc.save(doc_path)

    for iteration in range(5):
        message = replace_block_between_manual_anchors(
            str(doc_path),
            "START",
            [f"Anchor revision {iteration}", f"Anchor notes {iteration}"],
            end_anchor_text="END",
        )
        assert "Replaced content between 'START' and 'END'" in message

    final_doc = Document(doc_path)
    assert [paragraph.text for paragraph in final_doc.paragraphs if paragraph.text] == [
        "START",
        "Anchor revision 4",
        "Anchor notes 4",
        "END",
        "Trailing content",
    ]
    assert len(final_doc.tables) == 0


def test_markdown_round_trip_and_section_replacement(tmp_path: Path):
    doc_path = tmp_path / "markdown-flow.docx"
    initial_markdown = """# Standard Draft

Intro paragraph.

## Scope

- Item A
- Item B

| Clause | Value |
| --- | --- |
| A1 | Draft |
"""

    replace_message = asyncio.run(replace_document_with_markdown(str(doc_path), initial_markdown))
    assert "replaced from markdown" in replace_message

    exported_markdown = asyncio.run(get_document_markdown(str(doc_path)))
    assert "# Standard Draft" in exported_markdown
    assert "## Scope" in exported_markdown
    assert "- Item A" in exported_markdown
    assert "| Clause | Value |" in exported_markdown

    section_markdown = """Scope paragraph updated.

### Clause Details

1. First clause
2. Second clause
"""

    section_message = asyncio.run(
        replace_section_with_markdown(str(doc_path), "Scope", section_markdown)
    )
    assert "Section 'Scope' replaced from markdown" in section_message

    final_markdown = asyncio.run(get_document_markdown(str(doc_path)))
    assert "Scope paragraph updated." in final_markdown
    assert "### Clause Details" in final_markdown
    assert "1. First clause" in final_markdown
    assert "2. Second clause" in final_markdown


def test_format_text_preserves_plain_text_content(tmp_path: Path):
    doc_path = tmp_path / "format-text.docx"
    doc = Document()
    doc.add_paragraph("Alpha Beta Gamma")
    doc.save(doc_path)

    message = asyncio.run(format_text(str(doc_path), 0, 6, 10, bold=True, color="red"))
    assert "formatted successfully" in message

    final_doc = Document(doc_path)
    assert final_doc.paragraphs[0].text == "Alpha Beta Gamma"
