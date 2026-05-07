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
    """Return markdown for review, diffing, draft edits, or exact restoration.

    The returned markdown starts with a wadocx:fidelity-bundle comment that can
    restore the exact DOCX bytes through replace_document_with_markdown. Embedded
    images are exported to a sibling *_media directory and referenced as markdown
    image links.
    """
    filename = ensure_docx_extension(filename)
    return document_to_markdown(filename)


async def export_document_markdown_to_file(filename: str, output_filename: Optional[str] = None) -> str:
    """Export markdown to a file and return the output path.

    The file includes a wadocx:fidelity-bundle for byte-for-byte DOCX restore.
    Embedded images are exported to a sibling *_media directory.
    """
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
    """Replace the document body with markdown-rendered content.

    Supported editable markdown:
    - Headings (# through ######), paragraphs, bullet lists, numbered lists,
      tables, and local markdown images.
    - Alignment blocks: <div align="left|center|right|justify"> ... </div>.
    - Page breaks: <!-- PAGE BREAK -->.
    - Native Word table-of-contents fields: <!-- TOC -->.
    - Configured TOC fields:
      <!-- wadocx:toc
      title: Contents
      max_level: 3
      style: dotted
      page_break_after: true
      -->
      Supported styles are dotted, page_numbers/plain, and links/web.

    If markdown_text begins with a wadocx:fidelity-bundle exported by wadocx,
    the original DOCX bytes are restored exactly instead of rebuilding content.
    """
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
    """Replace content below a section heading using markdown-rendered blocks.

    Supports the same editable markdown blocks as replace_document_with_markdown,
    including <!-- TOC --> and configured wadocx:toc directives. Fidelity bundles
    are rejected here because exact restore is only valid for whole documents.
    """
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
