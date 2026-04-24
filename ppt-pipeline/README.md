# PDF to Video Pipeline

Convert a source PDF into a narrated presentation video through a 7-stage pipeline:

1. Parse slides from PDF
2. Build logical section structure
3. Generate content and narration notes
4. Build themed PPTX output
5. Export slide images
6. Generate audio narration
7. Render final MP4

The pipeline is checkpoint-driven, so each stage can be rerun independently and downstream stages can reuse upstream outputs.

## What This Project Is Good At

- Turning document-like PDFs into presentation-like narrated outputs
- Recovering from intermittent AI provider failures with provider fallback
- Producing usable content for image or diagram heavy slides
- Detecting empty body slides in Stage 3 audit output
- Keeping title, agenda, and conclusion aligned with generated body content

## Architecture Overview

The API server is Flask-based and orchestrates stage modules in the `pipeline/` package.

| Stage | Module | Input | Output checkpoint |
|---|---|---|---|
| 1 | `stage1_parser.py` | uploaded PDF | `checkpoints/stage1_parsed/<name>.json` |
| 2 | `stage2_structurer.py` | stage1 checkpoint | `checkpoints/stage2_structured/<name>.json` |
| 3 | `stage3_content.py` | stage1 + stage2 | `checkpoints/stage3_content/<name>.json` |
| 4 | `stage4_builder.py` | stage1 + stage3 | `checkpoints/stage4_pptx/<name>.json` and `.pptx` |
| 5 | `stage5_images.py` | stage4 PPTX or revised PPTX | `checkpoints/stage5_images/<name>.json` + PNGs |
| 6 | `stage6_audio.py` | stage4/revised PPTX notes | `checkpoints/stage6_audio/<name>.json` + MP3s |
| 7 | `stage7_video.py` | stage5 images + stage6 audio | `checkpoints/stage7_video/<name>.json` + MP4 |

## Prerequisites

- Python 3.12 (required, enforced at runtime)
- Windows is the primary target environment for this repo
- At least one LLM API key (Groq, Gemini, or OpenRouter)
- One of the following for Stage 5 image export:
	- Microsoft PowerPoint (COM automation), or
	- LibreOffice (`soffice`)

Recommended:

- 8 GB+ RAM
- Stable internet for AI and fallback TTS providers

## New System Setup (Windows)

Use this sequence on a fresh machine.

1. Open PowerShell at workspace root (folder that contains `ppt-pipeline/`).
2. Create Python 3.12 virtual environment at workspace root as `.venv312`.
3. Install dependencies from `ppt-pipeline/requirements.txt`.
4. Start the app from inside `ppt-pipeline/`.

```powershell
# 1) from workspace root (example)
cd C:\Users\mdsam\content_auto

# 2) create and activate venv
py -3.12 -m venv .venv312
.\.venv312\Scripts\Activate.ps1

# 3) install dependencies for this app
pip install -r .\ppt-pipeline\requirements.txt

# 4) move to app folder and run
cd .\ppt-pipeline
.\start_app.ps1
```

App URL: `http://localhost:5000`

## Quick Start

### 1) Create and activate Python 3.12 environment

```powershell
# Run from workspace root (parent of ppt-pipeline)
py -3.12 -m venv .venv312
.\.venv312\Scripts\activate
```

### 2) Install dependencies

```powershell
# If currently at workspace root:
pip install -r .\ppt-pipeline\requirements.txt

# If currently inside ppt-pipeline:
pip install -r requirements.txt
```

### 3) Configure environment

Create `.env` in the project root (`ppt-pipeline/.env`):

```env
# Required: configure at least one provider
GROQ_API_KEY=your_groq_api_key
GEMINI_API_KEY=your_gemini_api_key
OPENROUTER_API_KEY=your_openrouter_api_key

# Optional model overrides
GROQ_MODEL=llama-3.3-70b-versatile
GEMINI_MODEL=gemini-2.0-flash
MODEL=anthropic/claude-sonnet-4-5

# Optional provider and transport controls
AI_PROVIDER_ORDER=gemini,groq,openrouter
AI_HTTP_TIMEOUT_SECONDS=45
AI_HTTP_MAX_RETRIES=2
AI_HTTP_RETRY_BACKOFF_SECONDS=1.5

# Optional TTS defaults
DEFAULT_VOICE=af_heart
SILENCE_DURATION_MS=1500

# Stage checkpoint reuse: 1 enabled (default), 0 disabled
PIPELINE_USE_CACHE=1
```

### 4) Start the app

Recommended on Windows:

```powershell
# Run from inside ppt-pipeline/
./start_app.ps1
```

Alternative:

```powershell
python app.py
```

Open: http://localhost:5000

## End-to-End Usage

1. Upload a PDF (`/upload`)
2. Run Stage 2 structure (`/pipeline/structure/<filename>`)
3. Run Stage 3 content (`/pipeline/content/<filename>`)
4. Run Stage 4 build (`/pipeline/build/<filename>`)
5. Optionally upload a human-revised PPTX (`/upload-revised-pptx`)
6. Run Stage 5 images (`/pipeline/images/<filename>`)
7. Run Stage 6 audio (`/pipeline/audio/<filename>`)
8. Run Stage 7 video (`/pipeline/video/<filename>`)
9. Download outputs:
	 - PPTX: `/download-pptx/<filename>`
	 - MP4: `/download-video/<filename>`

Detailed endpoint contracts are in `docs/API_REFERENCE.md`.

## Stage 3 Quality and Reliability Notes

Stage 3 is the quality-critical stage. Current behavior includes:

- Deterministic typed blueprint output
- Mandatory title recovery for image-dominant slides where needed
- Optional vision enrichment for sparse slides (configurable)
- Rate-limit-safe behavior to avoid complete failure under provider throttling
- Empty body slide detection surfaced in:
	- `audit.content_gaps`
	- `stage3_runtime.empty_body_slides`
	- `stage3_runtime.empty_body_count`
- Agenda and conclusion generation grounded in generated body content

Useful Stage 3 environment overrides:

```env
STAGE3_PROVIDER_ORDER=gemini,groq,openrouter
STAGE3_AI_TIMEOUT_SECONDS=45
STAGE3_AI_MAX_RETRIES=2
STAGE3_AI_RETRY_BACKOFF_SECONDS=1.5

STAGE3_ENABLE_VISION_ENRICHMENT=1
STAGE3_MAX_VISION_SLIDES=1
STAGE3_VISION_IMAGE_ONLY=1
STAGE3_RATE_LIMIT_SAFE_MODE=1

STAGE3_MAX_REBALANCE_ATTEMPTS=1
STAGE3_REBALANCE_MARGIN_WORDS=80
```

## Stage 4 Rendering Notes

Stage 4 uses a theme scaffold deck at `theme/reference.pptx` and maps generated content into appropriate placeholders.

Current builder policy includes dynamic body placeholder selection for mixed layouts, which improves rendering reliability for image and diagram-oriented slides.

## Voice and Audio

Available Kokoro voice IDs:

- `af_heart`
- `af_sarah`
- `am_adam`
- `am_michael`
- `bf_emma`
- `bm_george`
- `af_nova`

Voice preview endpoint uses Kokoro first, then Piper, then gTTS fallback.

## Checkpoints, Cache, and Reruns

- Checkpoints are stored under `checkpoints/`.
- By default, successful checkpoints are reused (`PIPELINE_USE_CACHE=1`).
- To force fresh recomputation globally:

```powershell
$env:PIPELINE_USE_CACHE="0"
python app.py
```

- To force Stage 7 rebuild only, call video endpoint with `{"force": true}`.

More operations recipes are in `docs/OPERATIONS_AND_QUALITY.md`.

## Troubleshooting

### App exits with Python version error

The server enforces Python 3.12. Use `.venv312` and run via `start_app.ps1`.

### Stage 2/3 provider failures

- Ensure at least one API key is valid
- Adjust provider order (`AI_PROVIDER_ORDER` or `STAGE3_PROVIDER_ORDER`)
- Increase timeout and retries

### Stage 5 image export fails

- Install Microsoft PowerPoint (COM route) or LibreOffice (`soffice`)
- Confirm `soffice` is in PATH or default install path
- Check returned error details for the failing command

### Stage 6 mostly silent output

- Inspect `engine_summary` in Stage 6 result
- If many `silence` entries appear, TTS engines failed for those slides
- Verify Kokoro dependencies and Python 3.12 runtime

### Stage 7 ffmpeg or concat issues

- Re-run Stage 7 with `force=true`
- Check stage5 images and stage6 audio outputs exist and are complete

## API and Operations Guides

- `docs/API_REFERENCE.md`
- `docs/OPERATIONS_AND_QUALITY.md`

## Repository Structure

```text
ppt-pipeline/
	app.py
	start_app.ps1
	requirements.txt
	config/
		pipeline_config.json
	pipeline/
		checkpoint.py
		config.py
		stage1_parser.py
		stage2_structurer.py
		stage3_content.py
		stage4_builder.py
		stage5_images.py
		stage6_audio.py
		stage7_video.py
	static/
		index.html
	theme/
		reference.pptx
	uploads/
	checkpoints/
	docs/
		API_REFERENCE.md
		OPERATIONS_AND_QUALITY.md
```
