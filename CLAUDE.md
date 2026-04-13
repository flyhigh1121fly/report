# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Materials failure analysis report generation tool (失效报告生成助手). A Flask web app that accepts project info and images via a browser UI, uses AI (Qwen VLM/LLM via SiliconFlow API) to analyze failure images, and generates professional Word reports from a `.docx` template.

## Commands

```bash
# Run the app (auto-creates .venv and installs deps on first run)
run.bat

# Or manually:
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
- `/api/download` (GET) — serves generated `.docx` files
- Uses a background thread per request to stream SSE events through a `queue.Queue`

**Report generation pipeline (dynamic assembly):**
1. Load cover template (`模版_cover.docx`) — contains cover page, tables, and tail placeholders only
2. Replace cover page placeholders (`{{项目名称}}`, `{{背景概述}}`)
3. Process uploaded images (aspect-fit to 782×591px for sections 1-4, original for 5-6)
4. Call VLM (vision model) per section to describe images
5. Programmatic section building via `section_builders.py` — builds title + image slots + analysis text for each section, inserting before the "分析与讨论" tail section
6. Call LLM to generate discussion, conclusion, abstract, keywords
7. Replace tail placeholders in template
8. Save final `.docx` to `~/Desktop/gradio_report_jobs/{timestamp}_{project_name}/`

**Feature flag:** `USE_DYNAMIC_ASSEMBLY` (default `true`) — when `false`, uses the old full-template placeholder replacement approach via `模版.docx`

**Frontend (SPA, `index.html`)**
- Accordion UI with drag-and-drop image upload per slot
- SSE consumption for real-time status updates
- Dynamic section/slot configuration passed as JSON to backend

**AI Integration:**
- API: SiliconFlow OpenAI-compatible endpoint (`api.siliconflow.cn/v1`)
- VLM models: `Qwen/Qwen3-VL-32B-Instruct` (default), `Qwen/Qwen2.5-VL-72B-Instruct`, `Pro/Qwen/Qwen2.5-VL-7B-Instruct`
- LLM models: `Qwen/Qwen3-32B` (default), `deepseek-ai/DeepSeek-R1-Distill-Qwen-32B`, `Qwen/Qwen2.5-72B-Instruct`

## Key Data Structures

- `image_config` (from frontend) — JSON array of sections, each with `id`, `title`, and `slots` (each with `id` and `label`), allowing fully dynamic reconfiguration
- `section_builders.build_complete_section()` — builds a complete section (title + image slots + analysis text) and inserts into document
- `PLACEHOLDER_MAP` — maps section IDs to Word template text placeholders (used by legacy path only)

## Configuration

- `API_KEY` read from `.env` via `dotenv`
- `USE_DYNAMIC_ASSEMBLY` — feature flag, default `true` (set `false` in `.env` to use legacy template approach)
- Cover template: `模版_cover.docx` (dynamic mode) or `模版.docx`/`模板.docx` (legacy mode), placed in project dir or Desktop
- Output goes to `~/Desktop/gradio_report_jobs/`
- Font: 宋体 10.5pt (五号) for body text, 14pt bold for section titles

## Dependencies

`flask`, `requests`, `pillow`, `python-docx` — no version pins in requirements.txt
