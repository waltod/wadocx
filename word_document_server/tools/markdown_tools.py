"""
Markdown drafting and review tools for Word documents.
"""
import os
from typing import Optional

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.markdown_utils import (
    document_to_markdown,
    export_document_markdown,
    replace_document_with_markdown as replace_document_with_markdown_impl,
    replace_section_with_markdown as replace_section_with_markdown_impl,
)


async def get_document_markdown(filename: str) -> str:
    """Return the document as markdown for review, diffing, or draft edits."""
    filename = ensure_docx_extension(filename)
    return document_to_markdown(filename)


async def export_document_markdown_to_file(filename: str, output_filename: Optional[str] = None) -> str:
    """Export the document to a markdown file and return the output path."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if output_filename and not output_filename.lower().endswith(".md"):
        output_filename = f"{output_filename}.md"

    output_path = export_document_markdown(filename, output_filename)
    if output_path.startswith("Document ") and output_path.endswith(" does not exist"):
        return output_path
    return f"Markdown exported to {output_path}"


async def replace_document_with_markdown(filename: str, markdown_text: str) -> str:
    """Replace the document body with markdown-rendered content."""
    filename = ensure_docx_extension(filename)

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    result = replace_document_with_markdown_impl(filename, markdown_text)
    if result.get("restored_exact_docx"):
        return (
            f"Document {filename} restored exactly from a wadocx fidelity bundle "
            f"(sha256={result['sha256']})."
        )
    return (
        f"Document {filename} replaced from markdown with "
        f"{result['inserted_blocks']} body element(s) across {result['block_count']} parsed block(s)."
    )


async def replace_section_with_markdown(filename: str, header_text: str, markdown_text: str) -> str:
    """Replace the content below a section heading using markdown-rendered blocks."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    result = replace_section_with_markdown_impl(filename, header_text, markdown_text)
    if "error" in result:
        return result["error"]

    return (
        f"Section '{header_text}' replaced from markdown with "
        f"{result['inserted_blocks']} body element(s); removed {result['removed_blocks']} existing element(s)."
    )
