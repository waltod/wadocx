# WaDocx (MCP) vs. Anthropic's `docx` Skill

A practical comparison of the two ways to produce/edit Word documents with an AI
assistant: **WaDocx**, the MCP server in this repository, and **Anthropic's
`docx` skill** (from `anthropics/skills`, `skills/docx`).

They solve the same end-goal — let an agent create and edit `.docx` files — but
take fundamentally different architectural approaches.

---

## TL;DR

| | **WaDocx (this repo)** | **Anthropic `docx` skill** |
|---|---|---|
| **Form factor** | MCP server (76 tools over stdio/SSE/HTTP) | A *skill*: instructions + helper scripts the agent runs |
| **Core engine** | `python-docx` + targeted `lxml`/OOXML | docx-js (create) + raw OOXML unpack/edit/repack (edit) + pandoc + LibreOffice |
| **Editing model** | Call a named, typed tool (`add_paragraph`, `set_page_setup`, …) | Agent edits unpacked `document.xml` with the Edit tool, then repacks |
| **Tracked changes / comments** | Comment *reading* only; no tracked-changes authoring | First-class: insert/delete with author attribution, threaded comments |
| **Markdown round-trip** | Yes — strong (`get/replace_document_with_markdown`, fidelity bundle) | Via pandoc, less structured |
| **PDF / legacy `.doc`** | `convert_to_pdf` (docx2pdf) | LibreOffice (`soffice`) convert + render |
| **Determinism** | High — fixed tool surface, validated args | High control, but depends on the agent editing XML correctly |
| **Setup** | Install server + Python deps, configure MCP client | Needs Node (docx-js), pandoc, LibreOffice, Python |
| **Best for** | Automated/repeatable pipelines, drafting/revision loops, non-experts | Surgical edits, legal redlines (tracked changes), maximal OOXML control |

---

## How each one works

### WaDocx
A Model Context Protocol server exposing **76 typed tools**. The agent calls a
tool (`add_heading`, `add_table`, `add_live_table_of_contents`,
`set_page_setup`, `search_and_replace`, …); the server manipulates the document
with `python-docx`, dropping to raw `lxml`/OOXML only where python-docx can't
reach (field codes for TOC/PAGE/SEQ, bookmarks, hyperlinks, equations). The
contract is the tool signature — the agent never touches XML.

**Strengths**
- **Low cognitive load / low error surface.** Named tools with validated
  arguments; the agent can't emit malformed XML.
- **Repeatable & embeddable.** Any MCP client (Claude Code, Codex, Desktop) or
  script can drive identical operations — ideal for batch/automated generation.
- **Markdown round-tripping** with a fidelity bundle, plus stable block
  replacement (`replace_block_between_manual_anchors`,
  `replace_paragraph_block_below_header`) for revision loops.
- **Native fields done for you:** live (now *pre-rendered*) TOC, Table of
  Figures, SEQ captions, PAGE numbers, OMML equations.

**Limitations**
- **No tracked-changes authoring** (insertions/deletions with author) and only
  comment *reading*, not creation — a gap for legal/redline review.
- Bounded by the tool surface: anything without a tool isn't directly reachable
  (no arbitrary XML edits).
- `python-docx` semantics leak through for a few operations (e.g. replacing a
  table cell's text collapses multiple paragraphs).

### Anthropic `docx` skill
Not a server but a **skill**: a set of instructions plus helper scripts. It
treats `.docx` as a ZIP of XML and uses the right tool per job:
- **docx-js** to build new documents programmatically.
- **Unpack → edit `document.xml` with the Edit tool → repack/validate** for
  editing existing files (with run-merging + pretty-print to make string edits
  reliable, and auto-repair on repack).
- **pandoc** for extraction/conversion (`--track-changes=all`).
- **LibreOffice** for `.doc` conversion and PDF rendering.
- Python utilities: `accept_changes.py`, `comment.py`, `validate.py`.

**Strengths**
- **Tracked changes & threaded comments** as first-class citizens — the killer
  feature for document review/redlining.
- **Maximal control.** Direct XML editing reaches anything OOXML can express.
- Handles legacy `.doc` and PDF rendering via LibreOffice.

**Limitations**
- **Higher error surface.** Correctness depends on the agent editing raw XML
  correctly; malformed structure isn't auto-repaired.
- **Heavier toolchain** (Node + pandoc + LibreOffice + Python).
- **Less repeatable for automation** — it's an interactive agent workflow, not a
  stable callable API; harder to embed in a non-agent pipeline.
- Known sharp edges (explicit page-size config; landscape requires passing
  portrait dims; tables need dual width specs).

---

## When to use which

- **Choose WaDocx** when you want *programmatic, repeatable* document generation
  driven by tool calls — drafting pipelines, report factories, markdown→docx,
  or when the operator isn't an OOXML expert. Lowest chance of a corrupt file.
- **Choose the Anthropic skill** when you need **tracked changes / comment
  threads** (legal redlines, editorial review), surgical edits to a *specific*
  existing document, or OOXML features no tool exposes — and you have the
  toolchain installed.
- **They compose.** WaDocx can generate/assemble the document and round-trip
  markdown; the skill can then layer tracked-change review on top. WaDocx's main
  capability gap vs. the skill is exactly **tracked-changes authoring + comment
  creation** — the most valuable area for a future WaDocx tool.

---

## Feature matrix

| Capability | WaDocx | Anthropic `docx` |
|---|---|---|
| Create from scratch | ✅ tools | ✅ docx-js |
| Edit existing | ✅ tools / markdown blocks | ✅ raw XML edit |
| Headings / paragraphs / runs styling | ✅ | ✅ |
| Tables (incl. merge, shading, widths) | ✅ | ✅ |
| Images | ✅ | ✅ |
| Headers / footers + page numbers | ✅ | ✅ |
| Page setup (size/orientation/margins) | ✅ `set_page_setup` | ✅ |
| Table of contents (live field) | ✅ **pre-rendered + live** | ✅ (manual XML) |
| Table of figures / captions (SEQ) | ✅ `add_caption` / `add_table_of_figures` | ✅ (manual XML) |
| External hyperlinks | ✅ `add_hyperlink` | ✅ |
| Bookmarks / internal links | ✅ | ✅ |
| Footnotes / endnotes | ✅ (rich) | ✅ |
| Equations (OMML) | ✅ | ✅ |
| Document statistics | ✅ `get_document_statistics` | ➖ (ad hoc) |
| Markdown round-trip | ✅ (fidelity bundle) | ➖ via pandoc |
| **Tracked changes authoring** | ❌ | ✅ |
| **Comment creation / threads** | ❌ (read only) | ✅ |
| Read comments | ✅ | ✅ |
| PDF export | ✅ docx2pdf | ✅ LibreOffice |
| Legacy `.doc` import | ➖ | ✅ LibreOffice |
| Callable from non-agent code | ✅ (MCP) | ➖ (agent workflow) |

✅ supported · ➖ partial/indirect · ❌ not supported
