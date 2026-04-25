# Operations and Quality Guide

This guide focuses on practical runbooks for operating and validating the pipeline.

## 0) Fresh Machine Bootstrap (Windows)

Run these commands on a new system from workspace root (the parent folder of `ppt-pipeline/`).

```powershell
cd C:\Users\mdsam\content_auto
py -3.12 -m venv .venv312
.\.venv312\Scripts\Activate.ps1
pip install -r .\ppt-pipeline\requirements.txt
Copy-Item .\ppt-pipeline\.env.example .\ppt-pipeline\.env
cd .\ppt-pipeline
.\start_app.ps1
```

Important:

- `start_app.ps1` expects Python at `../.venv312/Scripts/python.exe` relative to `ppt-pipeline/`.
- Start the server from inside `ppt-pipeline/` when using `start_app.ps1`.
- Add provider keys in `ppt-pipeline/.env` (copied from `.env.example`) before first run.
- Default local URL is `http://localhost:5000`.

## 1) Recommended Daily Workflow

1. Start the app in Python 3.12 (`start_app.ps1`).
2. Upload PDF and run Stages 2 to 4.
3. Review Stage 3 quality output:
   - `audit.content_gaps`
   - `stage3_runtime.empty_body_count`
   - `stage3_runtime.vision_rate_limited`
4. If slide text/layout needs human edits, upload revised PPTX.
5. Run Stages 5 to 7.
6. Download PPTX and MP4.

## 2) Checkpoint and Cache Strategy

Checkpoint root: `checkpoints/`

By default, stages reuse valid checkpoints (`PIPELINE_USE_CACHE=1`).

Use fresh recompute mode when validating quality changes:

```powershell
$env:PIPELINE_USE_CACHE="0"
python app.py
```

Use cache mode for normal repeated runs:

```powershell
$env:PIPELINE_USE_CACHE="1"
python app.py
```

## 3) Clearing Checkpoints for Clean Tests

If you need completely fresh test runs, remove stage folders for a specific file or clear all stage outputs.

PowerShell example (all stage outputs for one file key):

```powershell
$k="notheme"
Remove-Item "checkpoints/stage1_parsed/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage2_structured/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage3_content/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage4_pptx/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage4_pptx/$k.pptx" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage5_images/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage6_audio/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage7_video/$k.json" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage7_video/$k.mp4" -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage1_parsed/$k" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage5_images/$k" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "checkpoints/stage6_audio/$k" -Recurse -Force -ErrorAction SilentlyContinue
```

## 4) Stage 3 Quality Controls

Stage 3 supports targeted controls through environment variables.

### Provider and retry behavior

```env
STAGE3_PROVIDER_ORDER=gemini,groq,openrouter
STAGE3_AI_TIMEOUT_SECONDS=45
STAGE3_AI_MAX_RETRIES=2
STAGE3_AI_RETRY_BACKOFF_SECONDS=1.5
```

### Vision enrichment behavior

```env
STAGE3_ENABLE_VISION_ENRICHMENT=1
STAGE3_MAX_VISION_SLIDES=1
STAGE3_VISION_IMAGE_ONLY=1
STAGE3_RATE_LIMIT_SAFE_MODE=1
```

### Narration rebalancing

```env
STAGE3_MAX_REBALANCE_ATTEMPTS=1
STAGE3_REBALANCE_MARGIN_WORDS=80
```

## 5) Diagnosing Content Quality Regressions

Symptoms and where to inspect:

- Image/diagram slides are weak:
  - Check `stage3_runtime.vision_rate_limited`
  - Check `stage3_runtime.diagram_context_fallback_count`
  - Check `stage3_runtime.vision_enriched_slides`
- Missing body content not detected:
  - Confirm `audit.content_gaps` is present
  - Confirm `stage3_runtime.empty_body_count` reflects reality
- Agenda mismatch with body slides:
  - Inspect Stage 3 `typed_blueprint` sequence and agenda entries
- Slide rendered blank though bullets exist:
  - Rebuild Stage 4 and verify `build_policy_version` in Stage 4 checkpoint

## 6) Stage 5 Export Reliability

Stage 5 priority:

1. PowerPoint COM (Windows)
2. LibreOffice headless fallback

If Stage 5 fails:

1. Verify PowerPoint is installed, or install LibreOffice.
2. Confirm `soffice` is available in PATH or default install path.
3. Re-run Stage 5 and read command-level stderr details from API/UI error output.

## 7) Stage 6 Audio Reliability

Engine order for each slide note:

1. Sarvam (multilingual-capable)
2. Kokoro (English)
3. Piper (English)
4. gTTS (English)
5. Silence fallback

Routing note:

- For non-English language requests, unsupported providers are auto-routed to Sarvam.

What to monitor in Stage 6 result:

- `engine_summary`
- `fallback_used`
- `primary_engine`
- `provider_routing`
- `transcripts.files`
- `transcript_language_mode`
- `matched_keywords`

If too many silence outputs:

1. Confirm Python 3.12 runtime
2. Confirm `SARVAM_API_KEY` is configured and valid
3. Confirm Kokoro dependencies are installed
4. Verify internet access for TTS providers when needed

Transcript policy checks:

1. Verify `transcript_en.json` and `transcript_en.srt` exist in Stage 6 output directory.
2. Verify transcript language mode is `always_english` for multilingual narration runs.
3. Verify protected technical terms remain English in transcript (`matched_keywords`).

## 8) Stage 7 Video Reliability

Stage 7 can reuse cache when outputs are still valid.

Use forced rebuild when:

- audio changed
- slide images changed
- previous run produced partially valid files

```json
{
  "force": true
}
```

If concat fails, Stage 7 uses fallback behavior; still inspect logs and rerun after confirming Stage 5 and Stage 6 outputs.

## 9) Uploading Human-Revised PPTX

Use `/upload-revised-pptx` when human edits are required after Stage 4.

Downstream behavior:

- Stage 5 exports images from revised PPTX
- Stage 6 reads notes from revised PPTX
- Stage 7 uses those generated assets

This is the preferred correction path when visual or narrative tuning is easier in PowerPoint than in pipeline prompts/config.

## 10) Operational Baseline Checklist

Before marking a run successful:

1. Stage 2 returns non-empty groups and no unassigned slides.
2. Stage 3 returns `empty_body_count = 0` (or known accepted gaps).
3. Stage 4 PPTX opens cleanly and key slides are populated.
4. Stage 5 slide count equals Stage 4 total slides.
5. Stage 6 engine summary shows expected speech coverage.
6. Stage 6 transcript artifacts (`transcript_en.json`, `transcript_en.srt`) are generated.
7. Stage 6 provider routing matches requested language expectations.
8. Stage 7 duration is plausible for total narration words.
9. Downloaded MP4 and PPTX files are readable.
