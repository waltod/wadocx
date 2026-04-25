import asyncio
import base64
import hashlib
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


def _write_png(path: Path) -> None:
    png_base64 = (
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Z0mQAAAAASUVORK5CYII="
    )
    path.write_bytes(base64.b64decode(png_base64))


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


def test_markdown_image_round_trip(tmp_path: Path):
    doc_path = tmp_path / "markdown-image.docx"
    image_path = tmp_path / "risk-flow.png"
    _write_png(image_path)

    markdown = f"""# Visual Draft

![Figure 5 - Risk and degraded-mode decision flow for landslide screening support](<{image_path}>)

Image paragraph.
"""

    replace_message = asyncio.run(replace_document_with_markdown(str(doc_path), markdown))
    assert "replaced from markdown" in replace_message

    doc = Document(doc_path)
    assert len(doc.inline_shapes) == 1
    assert doc.paragraphs[0].text == "Visual Draft"
    assert doc.paragraphs[-1].text == "Image paragraph."

    exported_markdown = asyncio.run(get_document_markdown(str(doc_path)))
    assert "![Figure 5 - Risk and degraded-mode decision flow for landslide screening support]" in exported_markdown
    assert "Image paragraph." in exported_markdown


def test_markdown_fidelity_bundle_restores_exact_docx(tmp_path: Path):
    source_doc_path = tmp_path / "source.docx"
    restored_doc_path = tmp_path / "restored.docx"

    doc = Document()
    doc.add_heading("Scope", level=1)
    doc.add_paragraph("Exact round-trip text.")
    doc.add_table(rows=2, cols=2)
    doc.tables[0].cell(0, 0).text = "A"
    doc.tables[0].cell(0, 1).text = "B"
    doc.tables[0].cell(1, 0).text = "C"
    doc.tables[0].cell(1, 1).text = "D"
    doc.save(source_doc_path)

    exported_markdown = asyncio.run(get_document_markdown(str(source_doc_path)))
    assert "<!-- wadocx:fidelity-bundle" in exported_markdown
    assert "# Scope" in exported_markdown

    restore_message = asyncio.run(
        replace_document_with_markdown(str(restored_doc_path), exported_markdown)
    )
    assert "restored exactly from a wadocx fidelity bundle" in restore_message

    source_hash = hashlib.sha256(source_doc_path.read_bytes()).hexdigest()
    restored_hash = hashlib.sha256(restored_doc_path.read_bytes()).hexdigest()
    assert source_hash == restored_hash


def test_section_replacement_rejects_fidelity_bundle(tmp_path: Path):
    doc_path = tmp_path / "section-bundle.docx"
    doc = Document()
    doc.add_heading("Scope", level=1)
    doc.add_paragraph("Original content")
    doc.save(doc_path)

    exported_markdown = asyncio.run(get_document_markdown(str(doc_path)))
    message = asyncio.run(
        replace_section_with_markdown(str(doc_path), "Scope", exported_markdown)
    )
    assert "does not accept a wadocx fidelity bundle" in message


def test_markdown_can_use_base_template_md_with_page_breaks(tmp_path: Path):
    template_doc_path = tmp_path / "template.docx"
    template_md_path = tmp_path / "template.md"
    output_doc_path = tmp_path / "derived.docx"

    template = Document()
    template.sections[0].header.paragraphs[0].text = "Template header"
    template.add_paragraph("Template cover")
    template.add_section()
    template.sections[1].footer.paragraphs[0].text = "Template footer"
    template.add_paragraph("Template body")
    template.save(template_doc_path)

    exported_template_markdown = asyncio.run(get_document_markdown(str(template_doc_path)))
    template_md_path.write_text(exported_template_markdown, encoding="utf-8")

    derived_markdown = f"""<!-- wadocx:base-template-md
path: {template_md_path}
-->

# Cover

Cover text.

<!-- PAGE BREAK -->

# Body

Body text.
"""

    replace_message = asyncio.run(
        replace_document_with_markdown(str(output_doc_path), derived_markdown)
    )
    assert "replaced from markdown" in replace_message

    derived = Document(output_doc_path)
    assert len(derived.sections) == 2
    assert derived.sections[0].header.paragraphs[0].text == "Template header"
    assert derived.sections[1].footer.paragraphs[0].text == "Template footer"
    assert "Cover" in [p.text for p in derived.paragraphs]
    assert "Body" in [p.text for p in derived.paragraphs]
