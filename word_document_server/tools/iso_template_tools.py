"""
Tools for compiling markdown drafts into Word templates with fixed front matter.
"""
from __future__ import annotations

from scripts.compile_iso_template_draft import compile_iso_draft


def compile_iso_template_draft(
    markdown_path: str,
    template_docx_path: str,
    output_docx_path: str,
) -> str:
    """
    Compile a markdown draft into an ISO-style Word template document.

    This is intended for templates where the final document must preserve
    template front matter, section breaks, headers, footers, and body styles
    while replacing the draft content from markdown.
    """
    result_path = compile_iso_draft(markdown_path, template_docx_path, output_docx_path)
    return (
        "ISO template draft compiled successfully.\n"
        f"Markdown: {markdown_path}\n"
        f"Template: {template_docx_path}\n"
        f"Output: {result_path}"
    )
