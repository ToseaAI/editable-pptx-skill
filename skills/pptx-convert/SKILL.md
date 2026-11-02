---
name: tosea-pptx-convert
description: Convert slide images or image-based PDFs into fully editable PowerPoint (.pptx) via the Reslide API by Tosea. Use when the user wants an editable PPT from screenshots, slide images, or NotebookLM-style PDFs.
---

# Reslide Editable PPTX Convert

Turn slide images or image-based PDFs into a **fully editable** PowerPoint
(`.pptx`) — real text boxes, vector shapes and pictures on their own layers, not
a screenshot pasted onto a slide. This skill drives the Reslide REST API from
a single, dependency-free Python script.

## When to use

Use this skill when the user wants an editable PowerPoint from:

- a **PDF** that is essentially images/slides (a report, an exported deck),
- one or more **slide images** or **screenshots** (PNG/JPG/WebP),
- a **NotebookLM** slide export or PDF that they want to edit and rebrand.

Do **not** use it to author brand-new slide content from a text prompt — this
converts existing visual pages into an editable deck.

## Prerequisites

Set these environment variables before running:

| Variable | Required | Default | Notes |
| --- | --- | --- | --- |
| `TOSEA_PPTX_API_KEY` | yes | — | Bearer API key. Create one at <https://reslide.tosea.ai/keys> |
| `TOSEA_PPTX_API_BASE` | no | `https://reslideapi.tosea.ai` | Override only for self-hosted / staging |

```bash
export TOSEA_PPTX_API_KEY=tpx_live_...
```

The script uses only the Python standard library (Python 3.8+). No `pip install`.

## Usage

Run `scripts/convert.py`. It uploads/creates the task, polls until done, and
downloads the finished `.pptx`. Progress is printed to **stderr**; the final
saved file path is the only line printed to **stdout** (easy to capture).

```bash
# Local PDF -> editable .pptx  (uploaded, then converted)
python scripts/convert.py deck.pdf

# Local slide images -> one editable .pptx  (uploaded in slide order)
python scripts/convert.py slide1.png slide2.png slide3.png

# Remote PDF by URL
python scripts/convert.py --pdf-url https://example.com/deck.pdf

# Remote images by URL (one per slide, in order)
python scripts/convert.py --image-urls https://ex.com/1.png https://ex.com/2.png

# Options
python scripts/convert.py deck.pdf --out ./out/deck.pptx --title "Q3 board deck"

# Supply your own stable job ID when an external system already has one
python scripts/convert.py deck.pdf --idempotency-key "customer-job-123"
```

- `--out PATH` — where to write the result (default: `<input-name>.pptx`, or
  `output.pptx` for URL inputs).
- `--title "…"` — optional deck title stored on the task.
- `--idempotency-key "…"` — stable caller job ID. By default the script derives
  one from the ordered input content/URLs and title, so rerunning the same job
  returns the existing task instead of charging twice.
- `--force` — atomically replace an existing output. Without it, the script
  refuses to overwrite a file.

### What the script does

1. **Local files** → `POST /v1/uploads` (multipart) to get an `upload_id`
   (one per file, uploaded in the order given so slide order is preserved).
2. **Create task** → `POST /v1/tasks` with `{ operation, input }` and an
   `Idempotency-Key`:
   - `pdf_to_pptx` with `{ upload_id | pdf_url, title? }`
   - `images_to_pptx` with `{ upload_ids | image_urls, title? }`
3. **Poll** → `GET /v1/tasks/{task_id}` every 3 seconds, honoring a `429`
   `Retry-After` header, until the status is `succeeded`, `failed` or `canceled`.
4. **Download** → save `result.pptx_url` to `--out`.

All API responses use the envelope `{ code, message, data }`, where `code == 0`
means success.

## Billing

- **1 page = 1 credit.** A 12-page PDF costs 12 credits; 5 images cost 5 credits.
- Credits are **reserved** when a task is created and **settle only on success**.
- A **failed or canceled task has zero net billable credits**. Its reservation
  is reconciled against the original paid and bonus pools subject to those
  pools' lifecycle rules; do not promise that a later balance will increase.
- The web workspace and the API share **one credit pool** per account.

## Error handling

The script exits non-zero and prints a clear message on failure. Common cases:

| Status | Meaning | What to do |
| --- | --- | --- |
| `401` | API key missing or invalid | Check `TOSEA_PPTX_API_KEY`; create/rotate a key at <https://reslide.tosea.ai/keys> |
| `402` | Not enough credits | Top up at <https://reslide.tosea.ai/billing> (1 page = 1 credit) |
| `429` | Rate limited | The poller waits for `Retry-After` automatically; if it happens on task creation, wait and retry |
| task `failed` | Conversion could not complete | The task has zero net billable credits; inspect the printed `error`, and retry with cleaner/higher-resolution inputs |

Exit codes: `0` = success (path printed to stdout), `1` = error (message on
stderr), `130` = interrupted.
