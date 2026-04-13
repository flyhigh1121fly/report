# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Materials failure analysis report generation tool (失效报告生成助手). A Flask web app that accepts project info and images via a browser UI, uses AI (Qwen VLM/LLM via OpenAI-compatible API) to analyze failure images, and generates professional Word reports from a `.docx` template.

## Commands

```bash
# Run via batch script (creates venv, installs deps, starts app)
run.bat

# Or manually with a working venv:
python report_tool.py
# → starts at http://127.0.0.1:7860

# Install dependencies
pip install -r requirements.txt
```

No test framework, linter, or build system is configured.

## Architecture

Backend (`report_tool.py` + `section_builders.py`) + single-file frontend (`index.html`):

**Backend (Flask)**
- `/` — serves `index.html`
- `/api/generate` (POST) — accepts form data + images, returns SSE stream for real-time progress
- `/api/models` (GET) — proxies model list from AI API
- `/api/kb-list` (GET) — proxies knowledge base list from KB-Matrix API
- `/api/download` (GET) — serves generated `.docx` files
- Uses a background thread per request to stream SSE events through a `queue.Queue`

**Frontend (SPA, `index.html`)**
- Vanilla JS, no framework. Accordion UI with drag-and-drop image upload per slot
- `imageConfig` JS array holds all section/slot definitions — users can add/remove sections and slots dynamically
- `fileStore` object caches File objects keyed by slot ID to survive re-renders
- SSE consumption via `ReadableStream` reader for real-time status updates

### Report generation pipeline (dynamic assembly, the default path)

1. Load cover template (`模版_cover.docx`) — contains cover page, tables, and tail placeholders only
2. Replace cover page placeholders (`{{项目名称}}`, `{{背景概述}}`)
3. Process uploaded images (aspect-fit to 782×591px for sections 1-4, original for 5-6)
4. **Per section**: detect type via `_TYPE_KEYWORDS` priority list, build type-specific VLM prompt from `VLM_TYPE_PROMPTS`, call VLM, store result
5. **Table extraction**: for slots whose label contains "表", a second VLM call extracts structured table data (headers + rows)
6. **Chapter construction** via `section_builders.py`:
   - First accordion section → "1 概述" with figure caption "图1 [groupTitle]"
   - Remaining sections → "2.1, 2.2..." sub-sections under "2 分析过程及结果"
   - Each section: bold title (14pt) → intro text (type-driven `INTRO_TYPE_TEMPLATES`) → image layout → figure/table caption → VLM analysis text
   - Single images get centered paragraphs; multiple images get borderless side-by-side tables (2 per row)
7. Optionally retrieve knowledge base content to augment LLM context
8. Call LLM to generate discussion, conclusion, abstract, keywords
9. Replace tail placeholders in template (`{{分析与讨论}}`, `{{结论}}`, `{{报告摘要}}`, `{{关键词}}`)
10. Count pages via Word COM (`win32com`) and insert on cover
11. Save to `~/Desktop/gradio_report_jobs/{timestamp}_{project_name}/`

**Feature flag:** `USE_DYNAMIC_ASSEMBLY` (default `true`) — when `false`, uses the legacy full-template placeholder replacement approach via `模版.docx`. Legacy path uses `PLACEHOLDER_MAP` to map section IDs to template placeholders and has a 3-layer image layout per slot.

### Section type detection

`_TYPE_KEYWORDS` is a priority-ordered list mapping Chinese keywords to analysis types: `metallography`, `chemical`, `mechanical`, `sem`, `macro`, `overview`. The first section always defaults to `overview`. Each type drives a specific VLM prompt (`VLM_TYPE_PROMPTS`) and intro text template (`INTRO_TYPE_TEMPLATES`).

### Figure/table numbering

Global counters (`figure_counter`, `table_counter`) increment across all sections. Sections with "表" in slot labels get "表N" labels; others get "图N". Pre-computed `section_fig_labels` dict maps section IDs to their labels for use in VLM prompts.

## Key Data Structures

- `imageConfig` (frontend JS, sent as `image_config` JSON) — array of `{id, title, groupTitle, slots: [{id, label}]}`, fully dynamic
- `_AnchorProxy` (in `section_builders.py`) — wraps XML elements to unify paragraph/table manipulation, enabling `addnext()` chaining for sequential XML insertion
- `analysis_results: Dict[str, str]` — keyed by section ID, stores VLM analysis text
- `slot_table_data: Dict[str, dict]` — keyed by slot ID, stores extracted table data for table-type slots
- `PLACEHOLDER_MAP` — maps section IDs to Word template text placeholders (legacy path only)

## Configuration

- All config in `.env` (see `.env.example` for template): `BASE_URL`, `API_KEY`, `USE_DYNAMIC_ASSEMBLY`, plus KB credentials (`KB_AUTH_URL`, `KB_API_URL`, `KB_USERNAME`, `KB_PASSWORD`)
- Cover template: `模版_cover.docx` (dynamic mode) or `模版.docx`/`模板.docx` (legacy mode), searched in: script directory → Desktop → macOS iCloud Desktop
- Output: `~/Desktop/gradio_report_jobs/`
- Font: 宋体 10.5pt (五号) for body text, 14pt bold for section titles

## AI Integration

- OpenAI-compatible API (default: SiliconFlow `api.siliconflow.cn/v1`, configurable via `BASE_URL`)
- VLM models: `Qwen/Qwen3-VL-32B-Instruct` (default), `Qwen/Qwen2.5-VL-72B-Instruct`, `Pro/Qwen/Qwen2.5-VL-7B-Instruct`
- LLM models: `Qwen/Qwen3-32B` (default), `deepseek-ai/DeepSeek-R1-Distill-Qwen-32B`, `Qwen/Qwen2.5-72B-Instruct`
- Knowledge Base: KB-Matrix service (separate API with login auth) — augments LLM context with retrieved domain knowledge

## Dependencies

`flask`, `requests`, `pillow`, `python-docx`, `python-dotenv` — no version pins in requirements.txt. `win32com` is used for page counting on Windows but is not listed in requirements.txt.
