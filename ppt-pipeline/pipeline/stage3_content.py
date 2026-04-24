import re
import json
import os
import base64
import time
import requests
from dotenv import load_dotenv
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled
from pipeline.stage2_structurer import parse_llm_json, call_ai
from pipeline.config import get_ai_config

load_dotenv()
checkpoint_mgr = CheckpointManager()
AI_CFG = get_ai_config()

WORDS_PER_MINUTE = 130
MIN_DURATION_MINUTES = 7
MAX_DURATION_MINUTES = 10
MIN_TOTAL_WORDS = MIN_DURATION_MINUTES * WORDS_PER_MINUTE
MAX_TOTAL_WORDS = MAX_DURATION_MINUTES * WORDS_PER_MINUTE
MIN_NOTE_WORDS_FLOOR = 45
NOTES_POLICY_VERSION = 7
TYPED_BLUEPRINT_VERSION = 7
BULLET_CHARS = '●•·○◦▪▸►▶–—*-'
STAGE3_ENABLE_VISION_ENRICHMENT = str(os.getenv('STAGE3_ENABLE_VISION_ENRICHMENT', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}
STAGE3_MAX_VISION_SLIDES = max(0, int(os.getenv('STAGE3_MAX_VISION_SLIDES', '1') or 1))
STAGE3_MAX_REBALANCE_ATTEMPTS = max(0, int(os.getenv('STAGE3_MAX_REBALANCE_ATTEMPTS', '1') or 1))
STAGE3_REBALANCE_MARGIN_WORDS = max(0, int(os.getenv('STAGE3_REBALANCE_MARGIN_WORDS', '80') or 80))
_stage3_provider_default = ','.join(AI_CFG.get('provider_order', ['groq', 'gemini', 'openrouter']))
STAGE3_PROVIDER_ORDER = [
    p.strip().lower() for p in os.getenv('STAGE3_PROVIDER_ORDER', _stage3_provider_default).split(',') if p.strip()
]
STAGE3_AI_TIMEOUT_SECONDS = max(
    5,
    int(os.getenv('STAGE3_AI_TIMEOUT_SECONDS', str(AI_CFG.get('http', {}).get('timeout_seconds', 45))) or 45),
)
STAGE3_AI_MAX_RETRIES = max(
    0,
    int(os.getenv('STAGE3_AI_MAX_RETRIES', str(AI_CFG.get('http', {}).get('max_retries', 2))) or 2),
)
STAGE3_AI_RETRY_BACKOFF_SECONDS = max(
    0.0,
    float(
        os.getenv(
            'STAGE3_AI_RETRY_BACKOFF_SECONDS',
            str(AI_CFG.get('http', {}).get('retry_backoff_seconds', 1.5)),
        ) or 1.5
    ),
)
STAGE3_SKIP_AUDIT_AI = str(os.getenv('STAGE3_SKIP_AUDIT_AI', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}
STAGE3_PRECISION_MODE = str(os.getenv('STAGE3_PRECISION_MODE', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}
STAGE3_VISION_IMAGE_ONLY = str(os.getenv('STAGE3_VISION_IMAGE_ONLY', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}
STAGE3_RATE_LIMIT_SAFE_MODE = str(os.getenv('STAGE3_RATE_LIMIT_SAFE_MODE', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}

DIAGRAM_HINTS = {'image_only', 'likely_diagram'}
CONTINUATION_CONNECTORS = ('and', 'or', 'to', 'of', 'for', 'with', 'in', 'on', 'by', 'via')

GENERIC_TITLE_KEYS = {
    'agenda',
    'table of contents',
    'toc',
    'conclusion',
    'summary',
    'overview',
    'title',
    'introduction',
}

STRUCTURAL_SOURCE_TITLE_KEYS = {
    'agenda',
    'table of contents',
    'toc',
    'conclusion',
    'summary',
    'closing',
    'takeaways',
}

SUPPLEMENTAL_NOTE_SENTENCES = [
    'This also clarifies practical impact, trade-offs, and where this concept is applied in real workflows.',
    'Notice how this idea connects to the previous topic and prepares the transition to the next section.',
    'From an implementation perspective, this point helps guide design choices and execution priorities.',
    'Keep this takeaway in mind, because it influences both architecture decisions and measurable outcomes.',
]


def _call_ai_stage3(prompt, label='stage3_call'):
    start = time.perf_counter()
    print(f'  Stage3 AI start: {label}')
    result = call_ai(
        prompt,
        provider_order=STAGE3_PROVIDER_ORDER,
        timeout_seconds=STAGE3_AI_TIMEOUT_SECONDS,
        max_retries=STAGE3_AI_MAX_RETRIES,
        retry_backoff_seconds=STAGE3_AI_RETRY_BACKOFF_SECONDS,
    )
    elapsed = time.perf_counter() - start
    print(f'  Stage3 AI done: {label} ({elapsed:.1f}s)')
    return result, elapsed


def _is_rate_limit_error(error):
    msg = str(error or '').lower()
    return (
        '429' in msg
        or 'rate limit' in msg
        or 'resource_exhausted' in msg
        or 'quota' in msg
        or 'too many requests' in msg
    )


def _heuristic_audit(slides, toc):
    titles = [_clean_text(s.get('title', '')) for s in slides]
    titles_l = [t.lower() for t in titles]

    def _find_title_index(patterns):
        for idx, title in enumerate(titles_l, start=1):
            if any(p in title for p in patterns):
                return idx
        return None

    title_slide_num = 1 if slides and _clean_text(slides[0].get('title', '')) else None
    agenda_slide_num = _find_title_index(['agenda', 'table of contents', 'toc'])
    conclusion_slide_num = _find_title_index(['conclusion', 'summary', 'closing', 'takeaways'])

    missing = []
    if not title_slide_num:
        missing.append('title')
    if not agenda_slide_num:
        missing.append('agenda')
    if not conclusion_slide_num:
        missing.append('conclusion')

    return {
        'presentation_title': titles[0] if titles and titles[0] else (toc[0] if toc else 'AI Presentation'),
        'has_title_slide': bool(title_slide_num),
        'has_agenda_slide': bool(agenda_slide_num),
        'has_conclusion_slide': bool(conclusion_slide_num),
        'title_slide_num': title_slide_num,
        'agenda_slide_num': agenda_slide_num,
        'conclusion_slide_num': conclusion_slide_num,
        'content_gaps': [],
        'missing_slides': missing,
    }


def _word_count(text):
    if not text:
        return 0
    return len(re.findall(r"\b\w+\b", text))


def _clean_text(text):
    return re.sub(r'\s+', ' ', (text or '')).strip()


def _trim_to_word_limit(text, max_words):
    text = _clean_text(text)
    if not text:
        return ''
    words = text.split()
    if len(words) <= max_words:
        return text

    trimmed = ' '.join(words[:max_words])
    last_stop = max(trimmed.rfind('.'), trimmed.rfind('!'), trimmed.rfind('?'))
    if last_stop > len(trimmed) // 2:
        return trimmed[:last_stop + 1].strip()
    return trimmed.rstrip() + '...'


def _depad_note(text):
    note = _clean_text(text)
    if not note:
        return ''

    fillers = set(SUPPLEMENTAL_NOTE_SENTENCES) | {
        'Keep this takeaway in mind, because it influences both architecture decisions and measurable outcomes.',
        'This also clarifies practical impact, trade-offs, and where this concept is applied in real workflows.',
    }
    for phrase in fillers:
        note = note.replace(phrase, '')
    note = re.sub(r'\s+', ' ', note).strip()
    return note


def _target_note_range(slide_count):
    safe_count = max(1, int(slide_count))
    avg_target = MIN_TOTAL_WORDS // safe_count
    base_floor = 45 if safe_count >= 18 else 60
    min_words = max(base_floor, min(avg_target, 80))
    max_words = max(min_words + 20, min(130, avg_target + 35))
    return min_words, max_words


def _fallback_note(slide, min_words, max_words):
    title = _clean_text(slide.get('title')) or f"Slide {slide.get('slide_num', '')}"
    bullets = _normalize_bullet_lines(slide.get('bullets') or _extract_candidate_bullets(slide), max_items=4)

    if bullets:
        joined = '; '.join(bullets[:3])
        note = (
            f'This slide covers {title}. '
            f'The key points are: {joined}. '
            'Focus on how these points connect to implementation choices and practical outcomes.'
        )
    else:
        raw_lines = [ln.strip() for ln in (slide.get('raw_text') or '').split('\n') if ln.strip()]
        raw_preview = ' '.join(raw_lines[:2])
        if raw_preview:
            note = f'This slide introduces {title}. {raw_preview}. We use this as context for the next section.'
        else:
            note = f'This slide presents {title}. It is primarily visual, so the narration explains the main idea and why it matters.'

    if _word_count(note) < min_words:
        note += f' In practice, {title} helps guide architecture, execution sequencing, and validation priorities.'

    return _trim_to_word_limit(note, max_words)


def _normalize_bullet_lines(bullets, max_items=6):
    cleaned = []
    for raw in bullets or []:
        line = _clean_text(raw)
        if not line:
            continue
        if line.strip(BULLET_CHARS + ' ').isdigit():
            continue
        cleaned.append(line.lstrip(BULLET_CHARS + ' ').strip())

    if not cleaned:
        return []

    merged = [cleaned[0]]
    for line in cleaned[1:]:
        prev = merged[-1]
        starts_lower = line[:1].islower()
        prev_open = prev.endswith((':', ',', ';', '(', '/')) or prev.split()[-1].lower() in CONTINUATION_CONNECTORS

        if starts_lower or prev_open:
            merged[-1] = f'{prev} {line}'.strip()
        else:
            merged.append(line)

    deduped = []
    seen = set()
    for line in merged:
        key = _title_key(line)
        if key and key in seen:
            continue
        if key:
            seen.add(key)
        deduped.append(line)
        if len(deduped) >= max_items:
            break
    return deduped


def _extract_candidate_bullets(slide, max_items=6):
    title = _clean_text(slide.get('title', ''))
    lines = [ln.strip() for ln in (slide.get('raw_text') or '').split('\n') if ln.strip()]
    if lines and title and lines[0].lower() == title.lower():
        lines = lines[1:]

    bullets = []
    seen = set()
    for ln in lines:
        cleaned = ln.lstrip(BULLET_CHARS + ' ').strip()
        if len(cleaned) < 4 or cleaned.isdigit():
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        bullets.append(cleaned)
        if len(bullets) >= max_items:
            break
    return bullets


def _fallback_bullets(slide, max_items=6):
    extracted = _normalize_bullet_lines(_extract_candidate_bullets(slide, max_items=max_items), max_items=max_items)
    if extracted:
        return extracted

    hint = str(slide.get('content_type_hint', '')).strip().lower()
    if hint in DIAGRAM_HINTS:
        return []

    lines = [ln.strip() for ln in (slide.get('raw_text') or '').split('\n') if ln.strip()]
    title = _clean_text(slide.get('title', ''))
    if lines and title and _clean_text(lines[0]).lower() == title.lower():
        lines = lines[1:]
    return _normalize_bullet_lines(lines, max_items=max_items)


def _coerce_slide_generation_map(parsed, slides, min_words_per_slide, max_words_per_slide):
    items = parsed.get('slides') if isinstance(parsed, dict) else None
    if not isinstance(items, list):
        return None

    by_num = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        slide_num = item.get('slide_num')
        try:
            slide_num = int(slide_num)
        except (TypeError, ValueError):
            continue
        by_num[slide_num] = item

    bullets_map = {}
    notes_map = {}
    missing_slide_nums = []
    for slide in slides:
        num = slide['slide_num']
        generated = by_num.get(num)
        if not generated:
            bullets_map[str(num)] = _fallback_bullets(slide)
            notes_map[str(num)] = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
            missing_slide_nums.append(num)
            continue

        raw_bullets = generated.get('bullets', [])
        bullets = []
        if isinstance(raw_bullets, list):
            for b in raw_bullets:
                cleaned = _clean_text(b)
                if cleaned:
                    bullets.append(cleaned)
        if not bullets:
            bullets = _fallback_bullets(slide)
        bullets_map[str(num)] = bullets[:6]

        note = _clean_text(generated.get('speaker_notes', ''))
        if not note:
            note = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
        note = _trim_to_word_limit(note, max_words_per_slide)
        if _word_count(note) < min_words_per_slide:
            note = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
        notes_map[str(num)] = note

    return bullets_map, notes_map, missing_slide_nums


def _generate_single_slide_content(slide, total, min_words_per_slide, max_words_per_slide):
    slide_num = slide['slide_num']
    per_prompt = f"""You are an expert presentation writer.

Slide {slide_num} of {total}
Title: {slide.get('title', '')}
Content: {_clean_text(slide.get('raw_text', ''))[:700]}

Generate both:
1) 3-6 in-slide bullets for the body
2) speaker notes in 4-6 sentences, between {min_words_per_slide} and {max_words_per_slide} words

Return ONLY JSON:
{{
  "slide_num": {slide_num},
  "bullets": ["...", "..."],
  "speaker_notes": "..."
}}
Start with {{ and end with }}"""

    try:
        raw, _ = _call_ai_stage3(per_prompt, label=f'single_slide_{slide_num}')
        parsed = parse_llm_json(raw)
        raw_bullets = parsed.get('bullets', []) if isinstance(parsed, dict) else []
        bullets = []
        if isinstance(raw_bullets, list):
            bullets = [_clean_text(b) for b in raw_bullets if _clean_text(b)]
        if not bullets:
            bullets = _fallback_bullets(slide)

        note = _clean_text(parsed.get('speaker_notes', '')) if isinstance(parsed, dict) else ''
        if not note:
            note = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
        note = _trim_to_word_limit(note, max_words_per_slide)
        if _word_count(note) < min_words_per_slide:
            note = _fallback_note(slide, min_words_per_slide, max_words_per_slide)

        return bullets[:6], note
    except Exception as ex:
        print(f'  Slide {slide_num}: targeted regeneration failed: {ex}')
        return _fallback_bullets(slide), _fallback_note(slide, min_words_per_slide, max_words_per_slide)


def _generate_slide_content_batch(slides, toc, min_words_per_slide, max_words_per_slide):
    payload = [
        {
            'slide_num': s['slide_num'],
            'title': s.get('title', ''),
            'raw_text': _clean_text(s.get('raw_text', ''))[:700],
            'word_count': int(s.get('word_count', 0) or 0),
            'has_image': bool(s.get('image_path')),
        }
        for s in slides
    ]

    prompt = f"""You are an expert presentation writer.

For each slide, produce:
1) in-slide bullet content suitable for the slide body
2) narrator speaker notes that explain the slide clearly

STRICT RULES:
1. Return one object for every input slide number.
2. `bullets` should be 3-6 concise, meaningful bullets.
3. Do not leave bullets empty even if raw text is sparse.
4. Speaker notes must explain the slide in 4-6 sentences.
5. Each speaker note must be between {min_words_per_slide} and {max_words_per_slide} words.
6. Keep facts grounded in provided title/raw text and TOC context.

TOC:
{json.dumps(toc, ensure_ascii=False)}

Slides:
{json.dumps(payload, ensure_ascii=False, indent=2)}

Return ONLY JSON in this shape:
{{
  "slides": [
    {{
      "slide_num": 1,
      "bullets": ["...", "..."],
      "speaker_notes": "..."
    }}
  ]
}}
Start with {{ and end with }}"""

    for attempt in range(1, 3):
        try:
            raw, _ = _call_ai_stage3(prompt, label=f'batch_slide_content_attempt_{attempt}')
            parsed = parse_llm_json(raw)
            mapped = _coerce_slide_generation_map(parsed, slides, min_words_per_slide, max_words_per_slide)
            if mapped:
                bullets_map, notes_map, missing_slide_nums = mapped
                print(f'  Batch slide-content generation succeeded (attempt {attempt})')
                # Improve only a small number of missing slides to avoid long tail latency.
                if missing_slide_nums:
                    total = len(slides)
                    limited = missing_slide_nums[:2]
                    for num in limited:
                        slide = next((s for s in slides if s['slide_num'] == num), None)
                        if not slide:
                            continue
                        bullets, note = _generate_single_slide_content(slide, total, min_words_per_slide, max_words_per_slide)
                        bullets_map[str(num)] = bullets
                        notes_map[str(num)] = note
                    print(f'  Filled {len(limited)} missing batch slides with targeted regeneration')
                return bullets_map, notes_map
        except Exception as ex:
            print(f'  WARNING: Batch slide-content generation attempt {attempt} failed: {ex}')

    print('  Falling back to deterministic content generation (no extra AI fan-out).')
    bullets_map = {}
    notes_map = {}
    for slide in slides:
        key = str(slide['slide_num'])
        bullets_map[key] = _fallback_bullets(slide)
        notes_map[key] = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
    return bullets_map, notes_map


def _enhance_sparse_bullets_with_vision(filename, slides, toc, bullets_map, skip_slide_nums=None):
    if not STAGE3_ENABLE_VISION_ENRICHMENT or STAGE3_MAX_VISION_SLIDES <= 0:
        return 0

    skip = {int(n) for n in (skip_slide_nums or [])}

    enhanced = 0
    sparse_sorted = sorted(slides, key=lambda s: int(s.get('word_count', 0) or 0))
    for slide in sparse_sorted:
        slide_num = int(slide.get('slide_num', 0) or 0)
        if slide_num in skip:
            continue
        if STAGE3_VISION_IMAGE_ONLY and not _is_image_dominant_slide(slide):
            continue
        if not _is_sparse_slide(slide):
            continue
        if enhanced >= STAGE3_MAX_VISION_SLIDES:
            break

        image_path = _infer_stage1_image_path(filename, slide)
        if not image_path:
            continue

        key = str(slide_num)
        prompt = f"""You are extracting grounded visible information from a presentation slide image.

Task: return 3-6 concise body bullets using ONLY content visible in the image and title context.
Do not invent facts.

Slide number: {slide_num}
Title: {slide.get('title', '')}
TOC context: {json.dumps(toc, ensure_ascii=False)}
Current bullets: {json.dumps(bullets_map.get(key, []), ensure_ascii=False)}

Return ONLY JSON:
{{
  "slide_num": {slide_num},
  "bullets": ["...", "..."]
}}
Start with {{ and end with }}"""

        parsed = _call_gemini_vision_json(prompt, image_path)
        if isinstance(parsed, dict) and parsed.get('_rate_limited'):
            if STAGE3_RATE_LIMIT_SAFE_MODE:
                break
            continue
        if not isinstance(parsed, dict):
            continue

        raw_bullets = parsed.get('bullets', [])
        bullets = []
        if isinstance(raw_bullets, list):
            bullets = [_clean_text(b) for b in raw_bullets if _clean_text(b)]
        if bullets:
            bullets_map[key] = bullets[:6]
            enhanced += 1

    return enhanced


def _run_mandatory_vision_for_diagrams(filename, slides, toc, bullets_map):
    required = [
        s for s in slides
        if str(s.get('content_type_hint', '')).strip().lower() in DIAGRAM_HINTS
        or bool(s.get('title_from_image'))
    ]
    required_nums = [int(s['slide_num']) for s in required]

    if not required:
        return {}, [], required_nums, 0, False
    if not STAGE3_ENABLE_VISION_ENRICHMENT or not _has_usable_gemini_key():
        return {}, [], required_nums, 0, False

    inferred_titles = {}
    inferred_nums = []
    attempted = 0
    rate_limited = False

    for slide in required:
        image_path = _infer_stage1_image_path(filename, slide)
        if not image_path:
            continue

        slide_num = int(slide['slide_num'])
        key = str(slide_num)
        attempted += 1
        prompt = f"""You are analyzing a presentation slide image.

Extract grounded visible text only.
Return a concise slide title and up to 4 bullets only if clearly readable in the image.
If readable bullets are not available, return an empty bullets array.

Slide number: {slide_num}
Current title: {slide.get('title', '')}
TOC context: {json.dumps(toc, ensure_ascii=False)}

Return ONLY JSON:
{{
  "title": "...",
  "bullets": ["...", "..."]
}}
Start with {{ and end with }}"""

        parsed = _call_gemini_vision_json(prompt, image_path)
        if isinstance(parsed, dict) and parsed.get('_rate_limited'):
            rate_limited = True
            if STAGE3_RATE_LIMIT_SAFE_MODE:
                break
            continue
        if not isinstance(parsed, dict):
            continue

        title = _clean_text(parsed.get('title', ''))
        if title:
            inferred_titles[slide_num] = title
            inferred_nums.append(slide_num)

        raw_bullets = parsed.get('bullets', [])
        if isinstance(raw_bullets, list):
            norm = _normalize_bullet_lines(raw_bullets, max_items=6)
            if norm:
                bullets_map[key] = norm

    return inferred_titles, sorted(set(inferred_nums)), required_nums, attempted, rate_limited


def _format_contextual_title(base_title, as_diagram=False):
    base = _clean_text(base_title)
    if not base:
        return ''
    if as_diagram:
        if 'diagram' in base.lower():
            return base
        return f'{base} Diagram'
    if 'overview' in base.lower():
        return base
    return f'{base} Overview'


def _infer_missing_titles_without_vision(slides, existing_title_map=None):
    existing_title_map = existing_title_map or {}
    inferred = {}

    def _meaningful(text):
        cleaned = _clean_text(text)
        return bool(cleaned) and not _is_generic_title(cleaned)

    for idx, slide in enumerate(slides):
        slide_num = int(slide.get('slide_num', 0) or 0)
        if slide_num in existing_title_map:
            continue

        raw_title = _clean_text(slide.get('title', ''))
        # Preserve explicitly structural slides (e.g., Agenda) to avoid leaking them into body flow.
        if _is_structural_source_title(raw_title) and not bool(slide.get('title_from_image')):
            continue

        needs_inference = bool(slide.get('title_from_image')) or not _meaningful(raw_title)
        if not needs_inference:
            continue

        prev_title = ''
        for j in range(idx - 1, -1, -1):
            candidate = _clean_text(slides[j].get('title', ''))
            if _meaningful(candidate):
                prev_title = candidate
                break

        next_title = ''
        for j in range(idx + 1, len(slides)):
            candidate = _clean_text(slides[j].get('title', ''))
            if _meaningful(candidate):
                next_title = candidate
                break

        hint = str(slide.get('content_type_hint', '')).strip().lower()
        is_diagram_like = hint in DIAGRAM_HINTS or bool(slide.get('image_path'))

        candidate_title = ''
        if prev_title:
            candidate_title = _format_contextual_title(prev_title, as_diagram=is_diagram_like)
        elif next_title:
            candidate_title = _format_contextual_title(next_title, as_diagram=is_diagram_like)

        if candidate_title:
            inferred[slide_num] = candidate_title

    return inferred, sorted(inferred.keys())


def _apply_inferred_titles_to_slides(slides, title_map):
    for slide in slides:
        slide_num = int(slide.get('slide_num', 0) or 0)
        inferred = _clean_text(title_map.get(slide_num, ''))
        if not inferred:
            continue
        current = _clean_text(slide.get('title', ''))
        if not current or _is_generic_title(current):
            slide['title'] = inferred
            slide['title_inferred'] = True


def _generate_notes_from_evidence_batch(slides, bullets_map, min_words_per_slide, max_words_per_slide):
    payload = [
        {
            'slide_num': s['slide_num'],
            'title': s.get('title', ''),
            'bullets': bullets_map.get(str(s['slide_num']), []),
            'raw_excerpt': _clean_text(s.get('raw_text', ''))[:300],
        }
        for s in slides
    ]

    prompt = f"""You are an expert presentation narrator.

Write speaker notes for each slide using only the provided title, bullets, and raw excerpt.

STRICT RULES:
1. One note per slide number.
2. 4-6 sentences per slide.
3. {min_words_per_slide}-{max_words_per_slide} words per slide.
4. Explain clearly and connect ideas; do not invent facts not present in provided evidence.

Slides:
{json.dumps(payload, ensure_ascii=False, indent=2)}

Return ONLY JSON in this shape:
{{
  "speaker_notes": {{
    "1": "...",
    "2": "..."
  }}
}}
Start with {{ and end with }}"""

    try:
        raw, _ = _call_ai_stage3(prompt, label='notes_from_evidence_batch')
        parsed = parse_llm_json(raw)
        notes_map = _coerce_notes_map(parsed, slides)
        if notes_map:
            cleaned = {}
            for slide in slides:
                key = str(slide['slide_num'])
                note = _trim_to_word_limit(_depad_note(notes_map.get(key, '')), max_words_per_slide)
                if _word_count(note) < min_words_per_slide:
                    note = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
                cleaned[key] = note
            return cleaned, 'ai', False
    except Exception as ex:
        print(f'  WARNING: Evidence-based notes batch failed: {ex}')
        if _is_rate_limit_error(ex):
            fallback = {}
            for slide in slides:
                fallback[str(slide['slide_num'])] = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
            return fallback, 'fallback_rate_limited', True

    fallback = {}
    for slide in slides:
        fallback[str(slide['slide_num'])] = _fallback_note(slide, min_words_per_slide, max_words_per_slide)
    return fallback, 'fallback', False


def _infer_stage1_image_path(filename, slide):
    existing = slide.get('image_path')
    if existing and os.path.exists(existing):
        return existing
    guessed = os.path.join(
        checkpoint_mgr.base_dir,
        'stage1_parsed',
        filename,
        f"slide_{int(slide['slide_num']):02d}.png",
    )
    return guessed if os.path.exists(guessed) else ''


def _is_image_dominant_slide(slide):
    hint = str(slide.get('content_type_hint', '')).strip().lower()
    if hint in DIAGRAM_HINTS:
        return True
    if bool(slide.get('title_from_image')):
        return True
    return False


def _has_usable_gemini_key():
    api_key = (os.getenv('GEMINI_API_KEY') or '').strip()
    if not api_key:
        return False
    lowered = api_key.lower()
    if 'your_gemini_api_key_here' in lowered:
        return False
    return True


def _is_sparse_slide(slide):
    wc = int(slide.get('word_count', 0) or 0)
    raw = _clean_text(slide.get('raw_text', ''))
    return wc <= 20 or len(raw) <= 90


def _call_gemini_vision_json(prompt, image_path):
    api_key = (os.getenv('GEMINI_API_KEY') or '').strip()
    if not _has_usable_gemini_key() or not image_path or not os.path.exists(image_path):
        return None

    ai_cfg = get_ai_config()
    model = ai_cfg['models']['gemini']
    timeout_seconds = STAGE3_AI_TIMEOUT_SECONDS

    with open(image_path, 'rb') as f:
        img_b64 = base64.b64encode(f.read()).decode('utf-8')

    url = f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}'
    payload = {
        'contents': [{
            'parts': [
                {'text': prompt},
                {'inline_data': {'mime_type': 'image/png', 'data': img_b64}},
            ]
        }]
    }

    try:
        response = requests.post(url, json=payload, timeout=timeout_seconds)
        response.raise_for_status()
        data = response.json()
        text = data['candidates'][0]['content']['parts'][0]['text']
        return parse_llm_json(text)
    except Exception as ex:
        if _is_rate_limit_error(ex):
            print(f'  Vision enrichment rate-limited: {ex}')
            return {'_rate_limited': True}
        print(f'  Vision enrichment failed: {ex}')
        return None


def _enhance_sparse_slides_with_vision(filename, slides, toc, bullets_map, notes_map, min_words_per_slide, max_words_per_slide):
    if not STAGE3_ENABLE_VISION_ENRICHMENT or STAGE3_MAX_VISION_SLIDES <= 0:
        return 0

    enhanced = 0
    sparse_sorted = sorted(slides, key=lambda s: int(s.get('word_count', 0) or 0))
    for slide in sparse_sorted:
        if STAGE3_VISION_IMAGE_ONLY and not _is_image_dominant_slide(slide):
            continue
        if not _is_sparse_slide(slide):
            continue
        if enhanced >= STAGE3_MAX_VISION_SLIDES:
            break
        image_path = _infer_stage1_image_path(filename, slide)
        if not image_path:
            continue

        slide_num = slide['slide_num']
        key = str(slide_num)
        prompt = f"""You are an expert presentation writer.

Use the slide image and heading context to improve this sparse slide's body bullets and narrator notes.

Slide number: {slide_num}
Title: {slide.get('title', '')}
TOC context: {json.dumps(toc, ensure_ascii=False)}
Current bullets: {json.dumps(bullets_map.get(key, []), ensure_ascii=False)}
Current speaker notes: {notes_map.get(key, '')}

Return ONLY JSON:
{{
  "slide_num": {slide_num},
  "bullets": ["...", "..."],
  "speaker_notes": "4-6 sentences, between {min_words_per_slide} and {max_words_per_slide} words"
}}
Start with {{ and end with }}"""

        parsed = _call_gemini_vision_json(prompt, image_path)
        if isinstance(parsed, dict) and parsed.get('_rate_limited'):
            if STAGE3_RATE_LIMIT_SAFE_MODE:
                break
            continue
        if not isinstance(parsed, dict):
            continue

        raw_bullets = parsed.get('bullets', [])
        bullets = []
        if isinstance(raw_bullets, list):
            bullets = [_clean_text(b) for b in raw_bullets if _clean_text(b)]
        if bullets:
            bullets_map[key] = bullets[:6]

        note = _clean_text(parsed.get('speaker_notes', ''))
        if note:
            note = _trim_to_word_limit(note, max_words_per_slide)
            if _word_count(note) >= min_words_per_slide:
                notes_map[key] = note
                enhanced += 1

    return enhanced


def _coerce_notes_map(parsed, slides):
    notes = parsed.get('speaker_notes') if isinstance(parsed, dict) else None
    if not isinstance(notes, dict):
        return None

    result = {}
    for slide in slides:
        key = str(slide['slide_num'])
        raw = notes.get(key)
        if raw is None:
            raw = notes.get(slide['slide_num'])
        cleaned = _depad_note(raw)
        if not cleaned:
            return None
        result[key] = cleaned
    return result


def _rebalance_notes_to_duration(slides, speaker_notes, min_words_per_slide, max_words_per_slide):
    current = {str(k): _depad_note(v) for k, v in speaker_notes.items()}

    for attempt in range(1, STAGE3_MAX_REBALANCE_ATTEMPTS + 1):
        total_words = sum(_word_count(current.get(str(s['slide_num']), '')) for s in slides)
        if MIN_TOTAL_WORDS <= total_words <= MAX_TOTAL_WORDS:
            break

        direction = 'expand' if total_words < MIN_TOTAL_WORDS else 'compress'
        notes_payload = [
            {
                'slide_num': s['slide_num'],
                'title': s.get('title', ''),
                'content': _clean_text(s.get('raw_text', ''))[:320],
                'current_notes': current.get(str(s['slide_num']), ''),
            }
            for s in slides
        ]

        prompt = f"""You are an expert presenter and instructional writer.

Rewrite the speaker notes for each slide to improve clarity, explanation quality, and pacing.

STRICT CONSTRAINTS:
1. Keep one note per slide in the same slide numbers.
2. Explain the slide clearly with context and transitions. Do not invent facts not grounded in slide content.
3. Each slide note MUST be between {min_words_per_slide} and {max_words_per_slide} words.
4. Total words across all slide notes MUST be between {MIN_TOTAL_WORDS} and {MAX_TOTAL_WORDS}.
5. This attempt should {direction} the total narration length.

Slides with current notes:
{json.dumps(notes_payload, ensure_ascii=False, indent=2)}

Return ONLY JSON in this shape:
{{
  "speaker_notes": {{
    "1": "...",
    "2": "..."
  }}
}}
Start with {{ and end with }}"""

        try:
            raw, _ = _call_ai_stage3(prompt, label=f'rebalance_attempt_{attempt}')
            parsed = parse_llm_json(raw)
            candidate = _coerce_notes_map(parsed, slides)
            if candidate:
                current = candidate
                print(f'  Rebalanced narration notes (attempt {attempt})')
        except Exception as ex:
            print(f'  WARNING: Rebalance attempt {attempt} failed: {ex}')

    total_words = sum(_word_count(current.get(str(s['slide_num']), '')) for s in slides)
    if total_words > MAX_TOTAL_WORDS:
        ratio = MAX_TOTAL_WORDS / max(1, total_words)
        compressed = {}
        for s in slides:
            key = str(s['slide_num'])
            src = current.get(key, '')
            curr_words = _word_count(src)
            target = max(min_words_per_slide, min(max_words_per_slide, int(curr_words * ratio)))
            compressed[key] = _trim_to_word_limit(src, target)
        current = compressed
        total_words = sum(_word_count(current.get(str(s['slide_num']), '')) for s in slides)

    if total_words < MIN_TOTAL_WORDS:
        # Deterministic fallback expansion if AI rebalance is still short.
        add_idx = 0
        ordered = sorted(slides, key=lambda x: _word_count(current.get(str(x['slide_num']), '')))
        for s in ordered:
            key = str(s['slide_num'])
            note = current.get(key, '')
            additions_for_slide = 0
            while _word_count(note) < max_words_per_slide and total_words < MIN_TOTAL_WORDS and additions_for_slide < 3:
                sentence = SUPPLEMENTAL_NOTE_SENTENCES[add_idx % len(SUPPLEMENTAL_NOTE_SENTENCES)]
                if sentence in note:
                    title = _clean_text(s.get('title', '')) or f"Slide {s['slide_num']}"
                    sentence = f'In this context, {title} influences implementation priorities and measurable results.'
                note += ' ' + sentence
                add_idx += 1
                additions_for_slide += 1
                note = _trim_to_word_limit(note, max_words_per_slide)
                current[key] = note
                total_words = sum(_word_count(current.get(str(sl['slide_num']), '')) for sl in slides)
            if total_words >= MIN_TOTAL_WORDS:
                break

    if total_words < MIN_TOTAL_WORDS:
        # Final pass: add one concise, title-grounded sentence per slide until minimum duration is met.
        for s in slides:
            key = str(s['slide_num'])
            note = current.get(key, '')
            if _word_count(note) >= max_words_per_slide:
                continue
            title = _clean_text(s.get('title', '')) or f"Slide {s['slide_num']}"
            sentence = f'This also clarifies why {title} matters for decisions, execution, and expected outcomes in practice.'
            candidate = _trim_to_word_limit((note + ' ' + sentence).strip(), max_words_per_slide)
            current[key] = candidate
            total_words = sum(_word_count(current.get(str(sl['slide_num']), '')) for sl in slides)
            if total_words >= MIN_TOTAL_WORDS:
                break

    return current


def _title_key(text):
    cleaned = _clean_text(text).lower()
    cleaned = re.sub(r'[^a-z0-9\s]+', '', cleaned)
    return re.sub(r'\s+', ' ', cleaned).strip()


def _is_generic_title(text):
    key = _title_key(text)
    return (not key) or (key in GENERIC_TITLE_KEYS) or bool(re.match(r'^slide\s+\d+$', key, flags=re.IGNORECASE))


def _humanize_filename(filename):
    base = os.path.splitext(os.path.basename(filename or ''))[0]
    cleaned = _clean_text(re.sub(r'[_\-]+', ' ', base))
    if not cleaned:
        return ''
    return ' '.join(word.upper() if len(word) <= 3 and word.isalpha() else word.capitalize() for word in cleaned.split())


def _reserve_exact_title(title, seen_titles):
    key = _title_key(title)
    if key:
        seen_titles.add(key)
    return title


def _is_structural_source_title(text):
    key = _title_key(text)
    if not key:
        return False
    if key in STRUCTURAL_SOURCE_TITLE_KEYS:
        return True
    return any(key.startswith(f'{prefix} ') for prefix in ('agenda', 'conclusion', 'summary'))


def _to_overview_title(title):
    cleaned = _clean_text(title)
    if not cleaned:
        return 'Presentation Overview'
    lowered = cleaned.lower()
    if 'overview' in lowered or 'presentation' in lowered:
        return cleaned
    return f'{cleaned} Overview'


def _ensure_unique_title(title, seen_titles):
    base = _clean_text(title) or 'Untitled Slide'
    key = _title_key(base)
    if key and key not in seen_titles:
        seen_titles.add(key)
        return base

    suffix = 2
    while True:
        candidate = f'{base} ({suffix})'
        ckey = _title_key(candidate)
        if ckey not in seen_titles:
            seen_titles.add(ckey)
            return candidate
        suffix += 1


def _infer_presentation_title(slides, toc, filename=''):
    for slide in slides or []:
        candidate = _clean_text(slide.get('title', ''))
        if not _is_generic_title(candidate):
            return candidate

    for entry in toc or []:
        candidate = _clean_text(entry)
        if not _is_generic_title(candidate):
            return f'{candidate} Overview'

    from_name = _humanize_filename(filename)
    if from_name and not _is_generic_title(from_name):
        return from_name

    return 'Presentation Overview'


def _classify_body_archetype(slide, bullets):
    hint = str(slide.get('content_type_hint', '')).strip().lower()
    raw = _clean_text(slide.get('raw_text', ''))
    wc = int(slide.get('word_count', 0) or 0)
    has_image = bool(slide.get('image_path'))

    if hint == 'image_only':
        return 'diagram'

    if hint == 'likely_diagram':
        if bullets and len(bullets) >= 2:
            return 'image_content'
        return 'diagram'

    if has_image and (wc <= 24 or len(raw) <= 100) and bullets:
        return 'image_content'
    return 'content'


def _extract_source_agenda_bullets(slides):
    for slide in slides or []:
        title = _clean_text(slide.get('title', ''))
        if _title_key(title) != 'agenda':
            continue

        from_stage1 = _normalize_bullet_lines(slide.get('bullets') or [], max_items=9)
        if from_stage1:
            return from_stage1

        lines = [ln.strip() for ln in (slide.get('raw_text') or '').split('\n') if ln.strip()]
        if lines and _clean_text(lines[0]).lower() == title.lower():
            lines = lines[1:]
        return _normalize_bullet_lines(lines, max_items=9)
    return []


def _build_agenda_bullets(source_agenda_bullets, toc, body_entries):
    bullets = []
    seen = set()

    # Keep agenda synchronized with what is actually in the generated body.
    for entry in body_entries or []:
        cleaned = _clean_text(entry.get('title', ''))
        if not cleaned or _is_generic_title(cleaned):
            continue
        key = _title_key(cleaned)
        if key in seen:
            continue
        seen.add(key)
        bullets.append(cleaned)
        if len(bullets) >= 7:
            return bullets

    for item in source_agenda_bullets or []:
        cleaned = _clean_text(item)
        if not cleaned or _is_generic_title(cleaned):
            continue
        key = _title_key(cleaned)
        if key in seen:
            continue
        seen.add(key)
        bullets.append(cleaned)
        if len(bullets) >= 7:
            return bullets

    for t in toc or []:
        cleaned = _clean_text(t)
        if not cleaned or _is_generic_title(cleaned):
            continue
        key = _title_key(cleaned)
        if key in seen:
            continue
        seen.add(key)
        bullets.append(cleaned)
        if len(bullets) >= 7:
            break

    if not bullets:
        for entry in body_entries:
            cleaned = _clean_text(entry.get('title', ''))
            key = _title_key(cleaned)
            if not cleaned or _is_generic_title(cleaned) or key in seen:
                continue
            seen.add(key)
            bullets.append(cleaned)
            if len(bullets) >= 7:
                break

    return bullets


def _is_low_information_bullet(text):
    key = _title_key(text)
    if not key:
        return True
    low_info_prefixes = (
        'visual walkthrough',
        'builds on',
        'leads into',
        'key topic in this section',
        'diagram overview',
    )
    return any(key.startswith(prefix) for prefix in low_info_prefixes)


def _collect_neighbor_evidence(slide_num, slides, slide_bullets=None, max_items=6, max_hops_per_side=3):
    slide_bullets = slide_bullets or {}
    idx = next((i for i, s in enumerate(slides) if int(s.get('slide_num', 0) or 0) == int(slide_num)), -1)
    if idx < 0:
        return []

    evidence = []
    seen = set()

    for step in (-1, 1):
        hops = 0
        j = idx + step
        while 0 <= j < len(slides) and hops < max_hops_per_side and len(evidence) < max_items:
            src = slides[j]
            src_title = _clean_text(src.get('title', ''))
            if _is_structural_source_title(src_title):
                j += step
                continue

            raw = (
                slide_bullets.get(str(src.get('slide_num')))
                or src.get('bullets')
                or _extract_candidate_bullets(src, max_items=4)
            )
            bullets = _normalize_bullet_lines(raw, max_items=4)
            for bullet in bullets:
                if _is_low_information_bullet(bullet):
                    continue
                key = _title_key(bullet)
                if not key or key in seen:
                    continue
                seen.add(key)
                evidence.append(bullet)
                if len(evidence) >= max_items:
                    break

            hops += 1
            j += step

    return evidence


def _build_diagram_context_bullets(slide, slides, toc, slide_bullets=None, max_items=3):
    slide_num = int(slide.get('slide_num', 0) or 0)
    title = _clean_text(slide.get('title', ''))
    idx = next((i for i, s in enumerate(slides) if int(s.get('slide_num', 0) or 0) == slide_num), -1)

    def _neighbor_title(step):
        if idx < 0:
            return ''
        j = idx + step
        while 0 <= j < len(slides):
            candidate = _clean_text(slides[j].get('title', ''))
            if candidate and not _is_generic_title(candidate) and not _is_structural_source_title(candidate):
                return candidate
            j += step
        return ''

    bullets = _collect_neighbor_evidence(slide_num, slides, slide_bullets=slide_bullets, max_items=max_items + 2)[:max_items]

    prev_title = _neighbor_title(-1)
    next_title = _neighbor_title(1)

    if len(bullets) < max_items and title and not _is_generic_title(title):
        bullets.append(f'{title}: key component interactions and flow.')

    if len(bullets) < max_items and toc and idx >= 0:
        toc_entry = _clean_text(toc[min(idx, len(toc) - 1)])
        if toc_entry and not _is_generic_title(toc_entry):
            bullets.append(f'Focus area: {toc_entry}.')

    if len(bullets) < max_items and prev_title and next_title:
        bullets.append(f'Connects {prev_title} with {next_title} in one end-to-end view.')

    if len(bullets) < max_items and prev_title:
        bullets.append(f'Extends the concepts introduced in {prev_title}.')

    if len(bullets) < max_items and next_title:
        bullets.append(f'Sets up the next section on {next_title}.')

    cleaned = []
    seen = set()
    for bullet in _normalize_bullet_lines(bullets, max_items=max_items * 2):
        if _is_low_information_bullet(bullet):
            continue
        key = _title_key(bullet)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(bullet)
        if len(cleaned) >= max_items:
            break
    return cleaned


def _build_conclusion_bullets(body_entries):
    bullets = []
    for entry in body_entries:
        entry_bullets = _normalize_bullet_lines(entry.get('bullets') or [], max_items=2)
        candidate = entry_bullets[0] if entry_bullets else ''
        if not candidate:
            title = _clean_text(entry.get('title', ''))
            if title and not re.match(r'^slide\s+\d+$', title, flags=re.IGNORECASE):
                if entry.get('type') == 'diagram':
                    candidate = f'Use the {title} visual to explain key component relationships and flow.'
                else:
                    candidate = f'Apply {title} in implementation and validation decisions.'
        if not candidate:
            continue
        bullets.append(candidate)
        if len(bullets) >= 5:
            break

    if not bullets:
        bullets = [
            'Recap the core concepts covered in this presentation',
            'Apply these ideas in practical implementation decisions',
            'Use this foundation to plan the next execution steps',
        ]
    return bullets


def _compose_title_notes(deck_title):
    return (
        f'This presentation introduces {deck_title}. '
        'We will walk through the main concepts in a structured flow, '
        'connecting each section to practical implementation choices and outcomes.'
    )


def _compose_agenda_notes(agenda_bullets):
    if not agenda_bullets:
        return (
            'This agenda outlines the key topics we will cover. '
            'Each section builds on the previous one to form a coherent narrative '
            'from concepts to practical execution.'
        )

    listed = ', '.join(agenda_bullets[:4])
    return (
        f'We will cover {listed}. '
        'As we move through each section, keep track of how the ideas connect, '
        'because those links are what make the overall approach actionable.'
    )


def _compose_conclusion_notes(conclusion_bullets):
    if not conclusion_bullets:
        return (
            'In conclusion, this presentation summarized the essential concepts and '
            'their practical implications. The next step is to apply these takeaways '
            'in implementation and validation.'
        )
    joined = '; '.join(conclusion_bullets[:3])
    return (
        f'In conclusion, we covered the following core outcomes: {joined}. '
        'These takeaways should guide design decisions, execution priorities, and '
        'how we validate the final delivery quality.'
    )


def _build_typed_blueprint(
    slides,
    toc,
    slide_bullets,
    speaker_notes,
    min_words_per_slide,
    max_words_per_slide,
    filename='',
    title_inference_map=None,
    inferred_title_nums=None,
):
    seen_titles = set()
    all_body_entries_seed = []
    title_inference_map = title_inference_map or {}
    inferred_title_nums = set(inferred_title_nums or [])

    for slide in slides:
        key = str(slide['slide_num'])
        bullets = _normalize_bullet_lines(
            slide_bullets.get(key) or slide.get('bullets') or _fallback_bullets(slide),
            max_items=6,
        )
        archetype = _classify_body_archetype(slide, bullets)
        if archetype == 'diagram' and bullets:
            bullets = bullets[:3]

        raw_title = _clean_text(slide.get('title', ''))
        inferred_title = _clean_text(title_inference_map.get(int(slide['slide_num']), ''))
        title_inferred = bool(slide.get('title_inferred', False))
        if (not raw_title or _is_generic_title(raw_title)) and inferred_title:
            raw_title = inferred_title
            title_inferred = True
        elif not raw_title:
            raw_title = f"Slide {slide['slide_num']}"
            title_inferred = int(slide['slide_num']) in inferred_title_nums

        notes = _clean_text(speaker_notes.get(key, ''))
        if not notes:
            notes = _fallback_note(slide, min_words_per_slide, max_words_per_slide)

        all_body_entries_seed.append({
            'type': archetype,
            'source_slide_num': slide['slide_num'],
            'raw_title': raw_title,
            'title_inferred': bool(title_inferred),
            'image_path': _infer_stage1_image_path(filename, slide),
            'bullets': bullets[:6],
            'speaker_notes': _trim_to_word_limit(notes, max_words_per_slide),
        })

    body_entries_seed = [
        seed for seed in all_body_entries_seed
        if not _is_structural_source_title(seed.get('raw_title', ''))
    ]
    if not body_entries_seed:
        body_entries_seed = all_body_entries_seed

    # Reserve canonical synthetic slide names first; duplicates will be renamed on body slides.
    agenda_title = _reserve_exact_title('Agenda', seen_titles)
    conclusion_title = _reserve_exact_title('Conclusion', seen_titles)

    deck_candidate = _infer_presentation_title(slides, toc, filename=filename)
    first_body_title = next((
        _clean_text(seed.get('raw_title', ''))
        for seed in body_entries_seed
        if _clean_text(seed.get('raw_title', '')) and not _is_generic_title(seed.get('raw_title', ''))
    ), '')
    if first_body_title:
        deck_candidate = _to_overview_title(first_body_title)
    if any(_title_key(seed.get('raw_title', '')) == _title_key(deck_candidate) for seed in body_entries_seed):
        deck_candidate = _to_overview_title(deck_candidate)
    if _is_generic_title(deck_candidate):
        first_body = next((
            _clean_text(seed.get('raw_title', ''))
            for seed in body_entries_seed
            if not _is_generic_title(seed.get('raw_title', ''))
        ), '')
        deck_candidate = _to_overview_title(first_body) if first_body else 'Presentation Overview'
    deck_title = _ensure_unique_title(deck_candidate, seen_titles)

    body_entries = []
    for seed in body_entries_seed:
        raw_title = _clean_text(seed.get('raw_title', ''))
        if not raw_title:
            raw_title = f"Slide {seed['source_slide_num']}"
        deduped = _ensure_unique_title(raw_title, seen_titles)

        normalized_bullets = _normalize_bullet_lines(seed.get('bullets', []), max_items=6)
        if seed.get('type') == 'diagram' and normalized_bullets:
            normalized_bullets = normalized_bullets[:3]

        body_entries.append({
            'type': seed['type'],
            'source_slide_num': seed['source_slide_num'],
            'title': deduped,
            'title_inferred': bool(seed.get('title_inferred', False)),
            'image_path': seed.get('image_path') or '',
            'embed_source_image': bool(seed.get('image_path')) and seed.get('type') in {'image_content', 'diagram'},
            'bullets': normalized_bullets,
            'speaker_notes': seed['speaker_notes'],
        })

    source_agenda_bullets = _extract_source_agenda_bullets(slides)
    agenda_bullets = _build_agenda_bullets(source_agenda_bullets, toc, body_entries)
    conclusion_bullets = _build_conclusion_bullets(body_entries)

    typed_blueprint = [
        {
            'type': 'title',
            'title': deck_title,
            'bullets': [],
            'speaker_notes': _compose_title_notes(deck_title),
        },
        {
            'type': 'agenda',
            'title': agenda_title,
            'bullets': agenda_bullets,
            'speaker_notes': _compose_agenda_notes(agenda_bullets),
        },
    ]

    typed_blueprint.extend(body_entries)
    typed_blueprint.append({
        'type': 'conclusion',
        'title': conclusion_title,
        'bullets': conclusion_bullets,
        'speaker_notes': _compose_conclusion_notes(conclusion_bullets),
    })

    return typed_blueprint


def _build_audit_from_blueprint(typed_blueprint, content_gaps=None):
    title_idx = next((i + 1 for i, s in enumerate(typed_blueprint) if s.get('type') == 'title'), None)
    agenda_idx = next((i + 1 for i, s in enumerate(typed_blueprint) if s.get('type') == 'agenda'), None)
    conclusion_idx = next((i + 1 for i, s in enumerate(typed_blueprint) if s.get('type') == 'conclusion'), None)

    return {
        'presentation_title': typed_blueprint[0].get('title', 'AI Presentation') if typed_blueprint else 'AI Presentation',
        'has_title_slide': bool(title_idx),
        'has_agenda_slide': bool(agenda_idx),
        'has_conclusion_slide': bool(conclusion_idx),
        'title_slide_num': title_idx,
        'agenda_slide_num': agenda_idx,
        'conclusion_slide_num': conclusion_idx,
        'content_gaps': content_gaps or [],
        'missing_slides': [],
    }


def generate_content(filename):
    """Stage 3: Deterministic blueprint contract + evidence-grounded notes."""

    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage3_content', filename):
        cached = checkpoint_mgr.load('stage3_content', filename)
        if (
            cached is not None
            and 'error' not in cached
            and cached.get('notes_policy_version') == NOTES_POLICY_VERSION
            and cached.get('typed_blueprint_version') == TYPED_BLUEPRINT_VERSION
        ):
            print(f'Valid stage 3 checkpoint found for {filename}')
            return cached
        print(f'Invalid or outdated stage 3 checkpoint for {filename}, re-running...')

    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    structured = checkpoint_mgr.load('stage2_structured', filename)
    if not parsed:
        raise Exception('Stage 1 checkpoint not found. Run Stage 1 first.')
    if not structured:
        raise Exception('Stage 2 checkpoint not found. Run Stage 2 first.')

    slides = parsed['slides']
    toc = structured.get('toc', [])
    total = len(slides)

    stage3_timings = {}
    stage3_start = time.perf_counter()

    estimated_total_slides = total + 3
    min_words_per_slide, max_words_per_slide = _target_note_range(estimated_total_slides)
    print(
        f'Narration target: {MIN_TOTAL_WORDS}-{MAX_TOTAL_WORDS} words total, '
        f'{min_words_per_slide}-{max_words_per_slide} words per slide'
    )

    print('Stage 3 Step 1: Building evidence-first bullets...')
    slide_bullets = {}
    for slide in slides:
        key = str(slide['slide_num'])
        from_stage1 = _normalize_bullet_lines(slide.get('bullets') or [], max_items=6)
        slide_bullets[key] = from_stage1 if from_stage1 else _fallback_bullets(slide)

    title_inference_map, inferred_title_nums, vision_required_nums, vision_attempted, vision_rate_limited = _run_mandatory_vision_for_diagrams(
        filename,
        slides,
        toc,
        slide_bullets,
    )
    contextual_title_map, contextual_inferred_nums = _infer_missing_titles_without_vision(
        slides,
        existing_title_map=title_inference_map,
    )
    for num, title in contextual_title_map.items():
        if num not in title_inference_map:
            title_inference_map[num] = title

    all_inferred_title_nums = sorted(set(inferred_title_nums + contextual_inferred_nums))
    _apply_inferred_titles_to_slides(slides, title_inference_map)

    stage3_timings['vision_required_slides'] = vision_required_nums
    stage3_timings['vision_required_count'] = len(vision_required_nums)
    stage3_timings['vision_attempted_count'] = vision_attempted
    stage3_timings['vision_rate_limited'] = bool(vision_rate_limited)
    stage3_timings['vision_title_inferred_slides'] = inferred_title_nums
    stage3_timings['vision_title_inferred_count'] = len(inferred_title_nums)
    stage3_timings['context_title_inferred_slides'] = contextual_inferred_nums
    stage3_timings['context_title_inferred_count'] = len(contextual_inferred_nums)
    stage3_timings['all_title_inferred_slides'] = all_inferred_title_nums
    stage3_timings['all_title_inferred_count'] = len(all_inferred_title_nums)
    if vision_required_nums and vision_attempted == 0:
        stage3_timings['vision_unavailable'] = True

    if STAGE3_RATE_LIMIT_SAFE_MODE:
        enhanced_count = 0
        stage3_timings['optional_vision_skipped'] = True
    else:
        enhanced_count = _enhance_sparse_bullets_with_vision(
            filename,
            slides,
            toc,
            slide_bullets,
            skip_slide_nums=vision_required_nums,
        )
    stage3_timings['vision_enriched_slides'] = enhanced_count
    if enhanced_count:
        print(f'  Vision-enriched sparse slide bullets: {enhanced_count}')

    print('Stage 3 Step 2: Generating speaker notes from slide evidence...')
    speaker_notes, notes_generation_mode, notes_rate_limited = _generate_notes_from_evidence_batch(
        slides,
        slide_bullets,
        min_words_per_slide=min_words_per_slide,
        max_words_per_slide=max_words_per_slide,
    )
    stage3_timings['notes_generation_mode'] = notes_generation_mode
    stage3_timings['notes_rate_limited'] = bool(notes_rate_limited)

    diagram_context_fallback_slides = []
    for slide in slides:
        key = str(slide['slide_num'])
        normalized_bullets = _normalize_bullet_lines(slide_bullets.get(key, []), max_items=6)
        archetype = _classify_body_archetype(slide, normalized_bullets)
        if archetype == 'diagram':
            if normalized_bullets:
                slide_bullets[key] = normalized_bullets[:3]
            else:
                fallback = []
                if bool(vision_rate_limited) or stage3_timings.get('vision_unavailable'):
                    fallback = _build_diagram_context_bullets(
                        slide,
                        slides,
                        toc,
                        slide_bullets=slide_bullets,
                        max_items=3,
                    )
                slide_bullets[key] = fallback
                if fallback:
                    diagram_context_fallback_slides.append(int(slide['slide_num']))
        else:
            slide_bullets[key] = normalized_bullets if normalized_bullets else _fallback_bullets(slide)
        if key not in speaker_notes or not _clean_text(speaker_notes[key]):
            speaker_notes[key] = _fallback_note(slide, min_words_per_slide, max_words_per_slide)

    stage3_timings['diagram_context_fallback_slides'] = diagram_context_fallback_slides
    stage3_timings['diagram_context_fallback_count'] = len(diagram_context_fallback_slides)

    pre_rebalance_words = sum(_word_count(text) for text in speaker_notes.values())
    if (
        pre_rebalance_words < (MIN_TOTAL_WORDS - STAGE3_REBALANCE_MARGIN_WORDS)
        or pre_rebalance_words > (MAX_TOTAL_WORDS + STAGE3_REBALANCE_MARGIN_WORDS)
    ) and not (STAGE3_RATE_LIMIT_SAFE_MODE and (notes_rate_limited or vision_rate_limited)):
        rebalance_start = time.perf_counter()
        speaker_notes = _rebalance_notes_to_duration(
            slides,
            speaker_notes,
            min_words_per_slide=min_words_per_slide,
            max_words_per_slide=max_words_per_slide,
        )
        stage3_timings['rebalance_seconds'] = round(time.perf_counter() - rebalance_start, 2)
    elif STAGE3_RATE_LIMIT_SAFE_MODE and (notes_rate_limited or vision_rate_limited):
        stage3_timings['rebalance_skipped_due_rate_limit'] = True

    print('Stage 3 Step 3: Building strict typed blueprint...')
    typed_blueprint = _build_typed_blueprint(
        slides,
        toc,
        slide_bullets,
        speaker_notes,
        min_words_per_slide,
        max_words_per_slide,
        filename=filename,
        title_inference_map=title_inference_map,
        inferred_title_nums=all_inferred_title_nums,
    )
    content_gaps = []
    for i, item in enumerate(typed_blueprint, start=1):
        if item.get('type') in {'title', 'agenda', 'conclusion'}:
            continue
        bullets = _normalize_bullet_lines(item.get('bullets', []), max_items=6)
        if bullets:
            continue
        content_gaps.append({
            'slide_num': i,
            'source_slide_num': item.get('source_slide_num'),
            'type': item.get('type'),
            'title': item.get('title'),
        })
    stage3_timings['empty_body_slides'] = [g['slide_num'] for g in content_gaps]
    stage3_timings['empty_body_count'] = len(content_gaps)

    audit = _build_audit_from_blueprint(typed_blueprint, content_gaps=content_gaps)

    missing_slides_content = []
    for item in typed_blueprint:
        if item.get('type') in {'title', 'agenda', 'conclusion'}:
            missing_slides_content.append({
                'slide_type': item['type'],
                'title': item['title'],
                'bullets': item.get('bullets', []),
                'speaker_notes': item.get('speaker_notes', ''),
            })

    insert_order = [
        {'position': 'first', 'content': missing_slides_content[0]},
        {'position': 'second', 'content': missing_slides_content[1]},
        {'position': 'last', 'content': missing_slides_content[2]},
    ] if len(missing_slides_content) == 3 else []

    blueprint_note_words = sum(_word_count(_clean_text(s.get('speaker_notes', ''))) for s in typed_blueprint)
    estimated_minutes = round(blueprint_note_words / WORDS_PER_MINUTE, 2)
    print(
        f'Speaker notes pacing: {blueprint_note_words} words '
        f'(~{estimated_minutes} min at {WORDS_PER_MINUTE} wpm)'
    )

    result = {
        'filename': filename,
        'generation_mode': 'deterministic_contract',
        'audit': audit,
        'missing_slides_content': missing_slides_content,
        'insert_order': insert_order,
        'speaker_notes': speaker_notes,
        'typed_blueprint_version': TYPED_BLUEPRINT_VERSION,
        'typed_blueprint': typed_blueprint,
        'title_inference': {
            'required_slides': stage3_timings.get('vision_required_slides', []),
            'vision_inferred_slides': stage3_timings.get('vision_title_inferred_slides', []),
            'context_inferred_slides': stage3_timings.get('context_title_inferred_slides', []),
            'inferred_slides': stage3_timings.get('all_title_inferred_slides', []),
        },
        'speaker_notes_total_words': blueprint_note_words,
        'notes_policy_version': NOTES_POLICY_VERSION,
        'narration_policy': {
            'words_per_minute': WORDS_PER_MINUTE,
            'min_duration_minutes': MIN_DURATION_MINUTES,
            'max_duration_minutes': MAX_DURATION_MINUTES,
            'target_total_words': [MIN_TOTAL_WORDS, MAX_TOTAL_WORDS],
            'target_per_slide_words': [min_words_per_slide, max_words_per_slide],
            'estimated_duration_minutes': estimated_minutes,
        },
        'stage3_runtime': {
            **stage3_timings,
            'rate_limit_safe_mode': STAGE3_RATE_LIMIT_SAFE_MODE,
            'rate_limited': bool(notes_rate_limited or vision_rate_limited),
            'provider_order': STAGE3_PROVIDER_ORDER,
            'ai_timeout_seconds': STAGE3_AI_TIMEOUT_SECONDS,
            'ai_max_retries': STAGE3_AI_MAX_RETRIES,
            'precision_mode': True,
            'total_seconds': round(time.perf_counter() - stage3_start, 2),
        },
        'original_slide_count': total,
        'final_slide_count': len(typed_blueprint),
    }

    checkpoint_mgr.save('stage3_content', filename, result)
    print(f'Stage 3 complete! Original: {total} slides -> Final: {result["final_slide_count"]} slides')
    return result
