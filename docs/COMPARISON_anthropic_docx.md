# WaDocx (MCP) vs. Anthropic's `docx` Skill

A practical comparison of the two ways to produce/edit Word documents with an AI
assistant: **WaDocx**, the MCP server in this repository, and **Anthropic's
`docx` skill** (from `anthropics/skills`, `skills/docx`).

They solve the same end-goal — let an agent create and edit `.docx` files — but
take fundamentally different architectural approaches.

---

## Verdict (as of WaDocx 1.4.0)

**WaDocx now wins overall for most practical use.** The two capabilities that
previously made the Anthropic skill the only choice — **tracked-changes
authoring** and **comment threads** — are now first-class, validated WaDocx tools
(plus legacy `.doc` import). WaDocx keeps its structural advantages on top:
typed/validated tools (lower error surface than hand-editing XML), a stable
callable API embeddable in non-agent pipelines, and strong markdown
round-tripping. The skill retains an edge only in *unbounded* OOXML reach — it
can touch anything the format expresses, whereas WaDocx is bounded by its 87
tools — and in reading tracked-inserted text inline (it uses pandoc;
python-docx's `.text` skips `w:ins` content).

## TL;DR

| | **WaDocx (this repo)** | **Anthropic `docx` skill** |
|---|---|---|
| **Form factor** | MCP server (87 tools over stdio/SSE/HTTP) | A *skill*: instructions + helper scripts the agent runs |
| **Core engine** | `python-docx` + targeted `lxml`/OOXML | docx-js (create) + raw OOXML unpack/edit/repack (edit) + pandoc + LibreOffice |
| **Editing model** | Call a named, typed tool (`add_paragraph`, `set_page_setup`, …) | Agent edits unpacked `document.xml` with the Edit tool, then repacks |
| **Tracked changes** | ✅ author + accept + reject + counts (`tracked_search_and_replace`, …) | ✅ via raw XML + `accept_changes.py` |
| **Comments** | ✅ create + threaded reply + resolve + read | ✅ create + threaded reply (`comment.py`) |
| **Markdown round-trip** | Yes — strong (`get/replace_document_with_markdown`, fidelity bundle) | Via pandoc, less structured |
| **PDF / legacy `.doc`** | `convert_to_pdf` (docx2pdf) + `convert_doc_to_docx` (LibreOffice) | LibreOffice (`soffice`) convert + render |
| **Determinism** | High — fixed tool surface, validated args | High control, but depends on the agent editing XML correctly |
| **Unbounded OOXML reach** | Bounded by the 87-tool surface | ✅ anything the format can express |
| **Setup** | Install server + Python deps, configure MCP client | Needs Node (docx-js), pandoc, LibreOffice, Python |
| **Best for** | Automated/repeatable pipelines, drafting + **review/redline** loops, non-experts | Exotic OOXML features no tool exposes; inline reading of tracked text |

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
- Bounded by the tool surface: anything without a tool isn't directly reachable
  (no arbitrary XML edits) — the skill's main remaining edge.
- `python-docx`'s `.text` skips `w:ins` content, so `get_document_text` won't
  show *tracked-inserted* text inline (accept the changes first, or the skill's
  pandoc `--track-changes=all` reads it directly).
- `python-docx` semantics leak through for a few operations (e.g. replacing a
  table cell's text collapses multiple paragraphs).

> **Update (1.4.0):** tracked-changes authoring (`tracked_search_and_replace`,
> `mark_text_as_deleted`, `add_tracked_insertion`, `accept_all_changes`,
> `reject_all_changes`) and comment authoring (`add_comment`, `reply_to_comment`,
> `resolve_comment`) are now built in — the former biggest gap is closed.

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

- **Choose WaDocx** for almost everything: *programmatic, repeatable* generation
  (drafting pipelines, report factories, markdown→docx) **and** document
  **review/redlining** (tracked changes + comment threads), with the lowest
  chance of a corrupt file because every operation is a validated tool call.
- **Choose the Anthropic skill** only when you need OOXML features no WaDocx tool
  exposes (exotic field codes, unusual structures) or must read tracked-inserted
  text inline via pandoc — and you already have the Node/pandoc/LibreOffice
  toolchain installed.
- **They still compose.** WaDocx can generate/assemble + redline a document; the
  skill can drop down to raw XML for anything off the beaten path. As of 1.4.0
  the former capability gap (tracked changes + comment authoring) is closed.

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
| **Tracked changes authoring** | ✅ insert/delete/replace | ✅ raw XML |
| **Accept / reject changes** | ✅ `accept_all_changes` / `reject_all_changes` | ✅ `accept_changes.py` |
| **Comment creation / threads** | ✅ create + reply + resolve | ✅ create + reply |
| Read comments | ✅ | ✅ |
| Inline read of tracked-inserted text | ➖ (python-docx skips `w:ins`) | ✅ pandoc `--track-changes` |
| PDF export | ✅ docx2pdf | ✅ LibreOffice |
| Legacy `.doc` import | ✅ `convert_doc_to_docx` | ✅ LibreOffice |
| Unbounded raw-OOXML reach | ➖ (bounded by tool surface) | ✅ |
| Callable from non-agent code | ✅ (MCP) | ➖ (agent workflow) |

✅ supported · ➖ partial/indirect · ❌ not supported
