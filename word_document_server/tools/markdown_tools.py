"""
Markdown drafting and review tools for Word documents.
"""
import os
from typing import Optional

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.markdown_utils import (
    document_to_markdown,
    export_document_markdown as export_document_markdown_impl,
    replace_document_with_markdown as replace_document_with_markdown_impl,
    replace_section_with_markdown as replace_section_with_markdown_impl,
)


async def get_document_markdown(filename: str, include_fidelity_bundle: bool = True) -> str:
    """Return markdown for review, diffing, draft edits, or exact restoration.

    When include_fidelity_bundle is true, the returned markdown starts with a
    wadocx:fidelity-bundle comment that can restore the exact DOCX bytes through
    replace_document_with_markdown. Embedded images are exported to a sibling
    *_media directory and referenced as markdown image links.
    """
    filename = ensure_docx_extension(filename)
    return document_to_markdown(filename, include_fidelity_bundle=include_fidelity_bundle)


async def export_document_markdown_to_file(
    filename: str,
    output_filename: Optional[str] = None,
    include_fidelity_bundle: bool = True,
) -> str:
    """Export markdown to a file and return the output path.

    When include_fidelity_bundle is true, the file includes a
    wadocx:fidelity-bundle for byte-for-byte DOCX restore. Embedded images are
    exported to a sibling *_media directory.
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if output_filename and not output_filename.lower().endswith(".md"):
        output_filename = f"{output_filename}.md"

    try:
        output_path = export_document_markdown_impl(
            filename,
            output_filename,
            include_fidelity_bundle=include_fidelity_bundle,
        )
    except OSError as e:
        return f"Markdown export failed: {str(e)}"
    if output_path.startswith("Document ") and output_path.endswith(" does not exist"):
        return output_path
    return f"Markdown exported to {output_path}"


async def export_document_markdown(
    filename: str,
    output_filename: Optional[str] = None,
    include_fidelity_bundle: bool = True,
) -> str:
    """Alias for the MCP-facing markdown export tool name."""
    return await export_document_markdown_to_file(
        filename,
        output_filename,
        include_fidelity_bundle=include_fidelity_bundle,
    )


async def replace_document_with_markdown(
    filename: str,
    markdown_text: str,
    mode: str = "auto",
    source_base_dir: Optional[str] = None,
) -> str:
    """Replace the document body with markdown-rendered content.

    Supported editable markdown:
    - Headings (# through ######), paragraphs, bullet lists, numbered lists,
      tables, and local markdown images.
    - Alignment blocks: <div align="left|center|right|justify"> ... </div>.
    - Page breaks: <!-- PAGE BREAK -->.
    - Template section boundaries: <!-- SECTION BREAK -->.
    - Native Word table-of-contents fields: <!-- TOC -->.
    - Configured TOC fields:
      <!-- wadocx:toc
      title: Contents
      max_level: 3
      style: dotted
      page_break_after: true
      -->
      Supported styles are dotted, page_numbers/plain, and links/web.

    Modes:
    - auto: restore exactly when a fidelity bundle is present, otherwise rebuild.
    - exact_restore: require a fidelity bundle and restore exact DOCX bytes.
    - editable_rebuild: ignore any leading fidelity bundle and rebuild content.
    - editable_rebuild_with_template: use a leading fidelity bundle as the base
      template, then rebuild the editable markdown body.

    source_base_dir resolves relative image and wadocx:base-template-md paths.
    A base template can be supplied with:
      <!-- wadocx:base-template-md
      path: C:\\path\\to\\template-export.md
      -->
    The referenced template markdown must include a wadocx:fidelity-bundle,
    which is produced by the default markdown export mode.
    """
    filename = ensure_docx_extension(filename)

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    try:
        result = replace_document_with_markdown_impl(
            filename,
            markdown_text,
            mode=mode,
            source_base_dir=source_base_dir,
        )
    except FileNotFoundError as e:
        return f"Markdown import failed: {str(e)}"
    except Exception as e:
        return f"Markdown import failed: {str(e)}"
    if "error" in result:
        return result["error"]
    if result.get("restored_exact_docx"):
        return (
            f"Document {filename} restored exactly from a wadocx fidelity bundle "
            f"(sha256={result['sha256']})."
        )
    return (
        f"Document {filename} replaced from markdown with "
        f"{result['inserted_blocks']} body element(s) across {result['block_count']} parsed block(s)."
    )


async def replace_document_with_markdown_file(
    filename: str,
    markdown_path: str,
    mode: str = "auto",
) -> str:
    """Replace a document using markdown read from a file.

    Relative markdown image paths and wadocx:base-template-md paths are resolved
    against the markdown file's directory. This is the preferred one-call MCP
    path for file-backed Markdown drafts.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(markdown_path):
        return f"Markdown file {markdown_path} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    try:
        with open(markdown_path, "r", encoding="utf-8") as markdown_file:
            markdown_text = markdown_file.read()
    except Exception as e:
        return f"Failed to read markdown file {markdown_path}: {str(e)}"

    return await replace_document_with_markdown(
        filename,
        markdown_text,
        mode=mode,
        source_base_dir=os.path.dirname(os.path.abspath(markdown_path)),
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
