# API Reference

Base URL: http://localhost:5000

All JSON error responses follow this shape:

```json
{
  "success": false,
  "error": {
    "code": "ERROR_CODE",
    "message": "Human readable message",
    "traceback": "... optional ..."
  }
}
```

Notes:

- `filename` route params refer to the checkpoint key (usually the uploaded PDF basename without extension).
- Stage calls are idempotent with checkpoint reuse when `PIPELINE_USE_CACHE=1`.

## Health and Static

### GET /

Serves the UI (`static/index.html`).

## Upload and Checkpoints

### POST /upload

Upload a PDF and run Stage 1 parse.

Request:

- `multipart/form-data`
- field `file`: `.pdf`

Success response (Stage 1 result):

```json
{
  "filename": "notheme",
  "page_count": 12,
  "slides": [
    {
      "slide_num": 1,
      "title": "...",
      "title_from_image": false,
      "raw_text": "...",
      "bullets": ["..."],
      "word_count": 123,
      "content_type_hint": "text_heavy",
      "slide_width_px": 1920,
      "slide_height_px": 1080,
      "image_path": ".../checkpoints/stage1_parsed/notheme/slide_01.png"
    }
  ],
  "parse_policy_version": 3
}
```

Errors:

- `UPLOAD_MISSING_FILE` (400)
- `UPLOAD_INVALID_FILE` (400)
- `UPLOAD_PARSE_FAILED` (500)

### GET /pipeline/status/<filename>

Returns checkpoint existence per stage.

```json
{
  "stage1_parsed": true,
  "stage2_structured": true,
  "stage3_content": true,
  "stage4_pptx": true,
  "stage5_images": true,
  "stage6_audio": false,
  "stage7_video": false
}
```

### GET /checkpoint/<stage>/<filename>

Returns raw checkpoint JSON for a specific stage.

Errors:

- `CHECKPOINT_NOT_FOUND` (404)

### GET /checkpoints/<path:filepath>

Serves files inside the checkpoints directory (images, audio preview files, etc).

## Stage 2

### POST /pipeline/structure/<filename>

Run Stage 2 section structuring.

Request body (optional):

```json
{
  "split_options": {
    "mode": "custom",
    "min_slides": 10,
    "max_slides": 15
  }
}
```

Success response:

```json
{
  "success": true,
  "result": {
    "toc": ["Section A", "Section B"],
    "groups": [
      {
        "section_title": "Section A",
        "slide_nums": [1, 2, 3],
        "insert_divider_before": false,
        "slide_type_summary": {
          "text_heavy": 2,
          "likely_diagram": 1,
          "image_only": 0
        }
      }
    ],
    "split_config": {
      "mode": "custom",
      "min_slides": 10,
      "max_slides": 15
    },
    "split_plan": {
      "parts": [
        {
          "part": 1,
          "section_titles": ["Section A"],
          "slide_count": 3,
          "start_slide": 1,
          "end_slide": 3,
          "slide_nums": [1, 2, 3]
        }
      ],
      "part_count": 1,
      "min_slides": 10,
      "max_slides": 15,
      "unsplittable_sections": [],
      "split_notes": "..."
    },
    "structure_policy_version": 2,
    "source_stage1_parse_policy_version": 3
  }
}
```

Errors:

- `STAGE2_STRUCTURE_FAILED` (500)

## Stage 3

### POST /pipeline/content/<filename>

Run Stage 3 content generation and audit.

Request body: none

Success response:

```json
{
  "success": true,
  "audit": {
    "presentation_title": "...",
    "has_title_slide": true,
    "has_agenda_slide": true,
    "has_conclusion_slide": true,
    "title_slide_num": 1,
    "agenda_slide_num": 2,
    "conclusion_slide_num": 14,
    "content_gaps": [],
    "missing_slides": []
  },
  "missing_slides": [
    {
      "slide_type": "title",
      "title": "...",
      "bullets": [],
      "speaker_notes": "..."
    }
  ],
  "speaker_notes_count": 12,
  "speaker_notes_total_words": 1040,
  "narration_policy": {
    "words_per_minute": 130,
    "min_duration_minutes": 7,
    "max_duration_minutes": 10,
    "target_total_words": [910, 1300],
    "target_per_slide_words": [60, 100],
    "estimated_duration_minutes": 8.2
  },
  "stage3_runtime": {
    "vision_required_count": 2,
    "vision_attempted_count": 2,
    "vision_rate_limited": false,
    "vision_enriched_slides": 1,
    "empty_body_slides": [],
    "empty_body_count": 0,
    "provider_order": ["gemini", "groq"],
    "total_seconds": 22.4
  },
  "original_slides": 11,
  "final_slides": 14
}
```

Errors:

- `STAGE3_CONTENT_FAILED` (500)

## Stage 4

### POST /pipeline/build/<filename>

Build themed PPTX from Stage 3 manifest.

Success response:

```json
{
  "success": true,
  "result": {
    "filename": "notheme",
    "output_path": ".../checkpoints/stage4_pptx/notheme.pptx",
    "total_slides": 14,
    "narrator_words": 1085,
    "estimated_minutes": 8.35,
    "build_policy_version": 11,
    "source_stage3_blueprint_version": 7,
    "source_stage3_notes_policy_version": 7,
    "slide_manifest": []
  }
}
```

Errors:

- `STAGE4_NOT_AVAILABLE` (501)
- `STAGE4_BUILD_FAILED` (500)

### GET /download-pptx/<filename>

Download generated PPTX (`<filename>_ai_generated.pptx`).

Errors:

- `PPTX_NOT_FOUND` (404)

## Revised PPTX Upload

### POST /upload-revised-pptx

Upload a manually edited PPTX to override Stage 4 output for Stages 5-7.

Request:

- `multipart/form-data`
- field `file`: `.pptx`
- field `filename`: checkpoint key

Success response:

```json
{
  "success": true,
  "saved_path": ".../checkpoints/stage5_input/notheme.pptx",
  "slides": 14,
  "narrator_words": 1192,
  "message": "Revised PPTX saved. 14 slides, ~1192 narrator words."
}
```

Errors:

- `REVISED_UPLOAD_MISSING_FILE` (400)
- `REVISED_UPLOAD_MISSING_FILENAME` (400)
- `REVISED_UPLOAD_INVALID_FILE` (400)

## Stage 5

### POST /pipeline/images/<filename>

Export slide images from PPTX.

Success response:

```json
{
  "success": true,
  "result": {
    "filename": "notheme",
    "slide_count": 14,
    "output_dir": ".../checkpoints/stage5_images/notheme",
    "method": "PowerPoint COM",
    "images": [".../slide_01.png"]
  }
}
```

Errors:

- `STAGE5_NOT_AVAILABLE` (501)
- `STAGE5_IMAGES_FAILED` (500)

## Voices and Stage 6

### GET /voices

Return available voices.

```json
{
  "success": true,
  "default_voice_id": "kokoro:af_heart",
  "default_language_code": "en-IN",
  "default_provider": "auto",
  "default_transcript_format": "both",
  "default_transcript_language_mode": "always_english",
  "default_keyword_policy": "keep_english",
  "default_protected_keywords": ["API", "SDK", "REST"],
  "languages": {
    "en-IN": "English (India)",
    "hi-IN": "Hindi"
  },
  "voices": {
    "kokoro:af_heart": {
      "name": "Heart (Female)",
      "gender": "Female",
      "provider": "kokoro",
      "voice": "af_heart",
      "languages": ["en-IN"]
    },
    "sarvam:shubh": {
      "name": "Shubh (Sarvam)",
      "gender": "Unknown",
      "provider": "sarvam",
      "voice": "shubh",
      "languages": ["en-IN", "hi-IN", "ta-IN"]
    }
  }
}
```

Errors:

- `VOICE_LIST_FAILED` (500)

### POST /voices/preview

Generate voice preview clip.

Request body:

```json
{
  "voice_id": "sarvam:shubh",
  "language_code": "hi-IN",
  "provider": "sarvam",
  "keyword_policy": "keep_english",
  "protected_keywords": ["API", "SDK"]
}
```

Notes:

- Preview text is fixed English copy by product policy.
- For non-English preview requests with unsupported providers, routing auto-selects Sarvam.

Success response:

```json
{
  "success": true,
  "result": {
    "voice_id": "sarvam:shubh",
    "voice_name": "Shubh (Sarvam)",
    "provider_requested": "sarvam",
    "provider_used": "sarvam",
    "provider_chain": ["sarvam", "kokoro", "piper", "gtts"],
    "provider_routing": {
      "requested_provider": "sarvam",
      "requested_language": "hi-IN",
      "auto_routed": false,
      "reason": "",
      "effective_primary": "sarvam"
    },
    "language_code": "hi-IN",
    "engine": "sarvam",
    "preview_text_policy": "english_only_fixed",
    "keyword_policy": "keep_english",
    "protected_keywords": ["API", "SDK"],
    "matched_keywords": ["API"],
    "preview_url": "/checkpoints/stage6_audio/previews/preview_sarvam_shubh_hi-IN.mp3"
  }
}
```

Failure response:

- `VOICE_PREVIEW_FAILED` (400)
- `VOICE_PREVIEW_EXCEPTION` (500)

### POST /pipeline/audio/<filename>

Generate one MP3 per slide from PPTX notes.

Request body (optional):

```json
{
  "voice": "sarvam:shubh",
  "language_code": "ta-IN",
  "provider": "sarvam",
  "transcript_format": "both",
  "transcript_language_mode": "always_english",
  "keyword_policy": "keep_english",
  "protected_keywords": ["API", "SDK", "REST"]
}
```

Success response:

```json
{
  "success": true,
  "result": {
    "filename": "notheme",
    "voice_id": "sarvam:shubh",
    "voice": "Shubh (Sarvam)",
    "language_code": "ta-IN",
    "provider_requested": "sarvam",
    "provider_used": "sarvam",
    "provider_chain": ["sarvam", "kokoro", "piper", "gtts"],
    "provider_routing": {
      "requested_provider": "sarvam",
      "requested_language": "ta-IN",
      "auto_routed": false,
      "reason": "",
      "effective_primary": "sarvam"
    },
    "slide_count": 14,
    "total_words": 1102,
    "output_dir": ".../checkpoints/stage6_audio/notheme",
    "audio_files": [],
    "transcript_formats": ["json", "srt"],
    "transcript_language_mode": "always_english",
    "transcripts": {
      "language_mode": "always_english",
      "transcript_language_code": "en-IN",
      "narration_language_code": "ta-IN",
      "formats": ["json", "srt"],
      "files": {
        "json": ".../checkpoints/stage6_audio/notheme/transcript_en.json",
        "srt": ".../checkpoints/stage6_audio/notheme/transcript_en.srt"
      },
      "segments": 14,
      "protected_keywords": ["API", "SDK", "REST"],
      "matched_keywords": ["API"]
    },
    "keyword_policy": "keep_english",
    "protected_keywords": ["API", "SDK", "REST"],
    "matched_keywords": ["API"],
    "engine_summary": {
      "sarvam": 12,
      "gtts": 1,
      "silence": 1
    },
    "primary_engine": "sarvam",
    "fallback_used": true
  }
}
```

Errors:

- `STAGE6_NOT_AVAILABLE` (501)
- `STAGE6_AUDIO_FAILED` (500)

## Stage 7

### POST /pipeline/video/<filename>

Render final MP4 from Stage 5 PNGs and Stage 6 MP3s.

Request body (optional):

```json
{
  "force": false
}
```

Use `force=true` to rebuild even if a valid stage7 checkpoint exists.

Success response:

```json
{
  "success": true,
  "result": {
    "filename": "notheme",
    "output_path": ".../checkpoints/stage7_video/notheme.mp4",
    "total_slides": 14,
    "total_duration_seconds": 484.1,
    "total_duration_minutes": 8.07,
    "fps": 24,
    "codec": "libx264"
  }
}
```

Errors:

- `STAGE7_NOT_AVAILABLE` (501)
- `STAGE7_VIDEO_FAILED` (500)

### GET /download-video/<filename>

Download final video (`<filename>_video.mp4`).

Errors:

- `VIDEO_NOT_FOUND` (404)
