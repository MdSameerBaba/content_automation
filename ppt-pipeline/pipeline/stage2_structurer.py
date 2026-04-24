
import re
import json
import os
import time
import requests
from dotenv import load_dotenv
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled
from pipeline.config import get_ai_config

load_dotenv()

checkpoint_mgr = CheckpointManager()
STRUCTURE_POLICY_VERSION = 2
CONTENT_TYPES = ('text_heavy', 'likely_diagram', 'image_only')


class AIProviderError(Exception):
    """Raised when all configured LLM providers fail."""


def parse_llm_json(raw_text):
    raw_text = raw_text.strip()
    # Method 1: extract from ```json ... ``` code blocks
    match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', raw_text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass
    # Method 2: direct parse
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        pass
    # Method 3: find outermost { }
    start = raw_text.find('{')
    end = raw_text.rfind('}')
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(raw_text[start:end+1])
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not parse JSON: {raw_text[:300]}")


def _post_json(url, headers, payload, timeout_seconds):
    response = requests.post(url, headers=headers, json=payload, timeout=timeout_seconds)
    response.raise_for_status()
    return response.json()


def _request_with_retries(request_fn, retries, backoff_seconds):
    last_err = None
    for attempt in range(retries + 1):
        try:
            return request_fn()
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError, requests.exceptions.HTTPError) as err:
            last_err = err
            if attempt == retries:
                break
            sleep_seconds = backoff_seconds * (attempt + 1)
            print(f'  Request failed (attempt {attempt + 1}/{retries + 1}), retrying in {sleep_seconds:.1f}s...')
            time.sleep(sleep_seconds)
    raise last_err


def _call_provider(provider, prompt, system_msg, models, timeout_seconds, retries, backoff_seconds):
    if provider == 'groq':
        api_key = os.getenv('GROQ_API_KEY')
        if not api_key:
            raise ValueError('GROQ_API_KEY is not configured')

        url = 'https://api.groq.com/openai/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        payload = {
            'model': models['groq'],
            'messages': [
                {'role': 'system', 'content': system_msg},
                {'role': 'user', 'content': prompt}
            ]
        }
        data = _request_with_retries(
            lambda: _post_json(url, headers, payload, timeout_seconds),
            retries,
            backoff_seconds,
        )
        return data['choices'][0]['message']['content']

    if provider == 'gemini':
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise ValueError('GEMINI_API_KEY is not configured')

        url = f'https://generativelanguage.googleapis.com/v1beta/models/{models["gemini"]}:generateContent?key={api_key}'
        payload = {'contents': [{'parts': [{'text': prompt}]}]}
        data = _request_with_retries(
            lambda: _post_json(url, None, payload, timeout_seconds),
            retries,
            backoff_seconds,
        )
        return data['candidates'][0]['content']['parts'][0]['text']

    if provider == 'openrouter':
        api_key = os.getenv('OPENROUTER_API_KEY')
        if not api_key:
            raise ValueError('OPENROUTER_API_KEY is not configured')

        url = 'https://openrouter.ai/api/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {api_key}',
            'HTTP-Referer': 'https://openrouter.ai/',
            'X-Title': 'ppt-pipeline'
        }
        payload = {
            'model': models['openrouter'],
            'messages': [
                {'role': 'system', 'content': system_msg},
                {'role': 'user', 'content': prompt}
            ]
        }
        data = _request_with_retries(
            lambda: _post_json(url, headers, payload, timeout_seconds),
            retries,
            backoff_seconds,
        )
        return data['choices'][0]['message']['content']

    raise ValueError(f'Unsupported AI provider: {provider}')


def call_ai(
    prompt,
    provider_order=None,
    timeout_seconds=None,
    max_retries=None,
    retry_backoff_seconds=None,
):
    ai_config = get_ai_config()
    providers = provider_order or ai_config['provider_order']
    models = ai_config['models']
    resolved_timeout_seconds = (
        ai_config['http']['timeout_seconds'] if timeout_seconds is None else timeout_seconds
    )
    retries = ai_config['http']['max_retries'] if max_retries is None else max_retries
    backoff_seconds = (
        ai_config['http']['retry_backoff_seconds'] if retry_backoff_seconds is None else retry_backoff_seconds
    )

    system_msg = (
        'You are a JSON-only responder. '
        'Respond with ONLY a valid JSON object. '
        'No explanation, no markdown, no code blocks. '
        'Start your response with { and end with }.'
    )

    provider_errors = []
    for provider in providers:
        try:
            print(f'  Trying AI provider: {provider}')
            return _call_provider(
                provider=provider,
                prompt=prompt,
                system_msg=system_msg,
                models=models,
                timeout_seconds=resolved_timeout_seconds,
                retries=retries,
                backoff_seconds=backoff_seconds,
            )
        except Exception as ex:
            msg = f'{provider} failed: {ex}'
            provider_errors.append(msg)
            print(f'  {msg}')

    raise AIProviderError(
        'All configured AI providers failed. '
        'Check API keys, model names, and network connectivity. '
        f'Providers attempted: {providers}. '
        f'Errors: {" | ".join(provider_errors)}'
    )


def _normalize_split_options(split_options):
    defaults = {
        'mode': 'auto',
        'min_slides': 10,
        'max_slides': 15,
    }

    if not isinstance(split_options, dict):
        return defaults

    mode = str(split_options.get('mode', 'custom' if split_options else 'auto')).lower()
    if mode not in ('auto', 'custom'):
        mode = 'custom'

    try:
        min_slides = int(split_options.get('min_slides', defaults['min_slides']))
    except (TypeError, ValueError):
        min_slides = defaults['min_slides']

    try:
        max_slides = int(split_options.get('max_slides', defaults['max_slides']))
    except (TypeError, ValueError):
        max_slides = defaults['max_slides']

    min_slides = max(1, min_slides)
    max_slides = max(min_slides, max_slides)

    return {
        'mode': mode,
        'min_slides': min_slides,
        'max_slides': max_slides,
    }


def _normalize_groups(groups):
    normalized = []
    for idx, group in enumerate(groups or [], 1):
        nums = []
        for raw in group.get('slide_nums', []):
            if isinstance(raw, int):
                nums.append(raw)
            elif isinstance(raw, str) and raw.strip().isdigit():
                nums.append(int(raw.strip()))
        nums = sorted(set(nums))
        if not nums:
            continue
        title = str(group.get('section_title') or f'Section {idx}').strip() or f'Section {idx}'
        normalized.append({'section_title': title, 'slide_nums': nums})

    normalized.sort(key=lambda g: min(g['slide_nums']))
    return normalized


def _infer_content_type_hint(slide):
    hint = str(slide.get('content_type_hint', '')).strip().lower()
    if hint in CONTENT_TYPES:
        return hint

    wc = int(slide.get('word_count', 0) or 0)
    if wc == 0:
        return 'image_only'
    if wc < 10:
        return 'likely_diagram'
    return 'text_heavy'


def _build_slide_lookup(slides):
    return {
        int(s.get('slide_num')): s
        for s in slides or []
        if isinstance(s, dict) and str(s.get('slide_num', '')).strip().isdigit()
    }


def _group_slide_type_summary(slide_nums, slide_lookup):
    summary = {k: 0 for k in CONTENT_TYPES}
    for num in slide_nums:
        slide = slide_lookup.get(num)
        if not slide:
            continue
        hint = _infer_content_type_hint(slide)
        summary[hint] += 1
    return summary


def _enrich_groups(groups, slide_lookup):
    enriched = []
    for idx, g in enumerate(_normalize_groups(groups)):
        enriched.append({
            **g,
            'insert_divider_before': idx > 0,
            'slide_type_summary': _group_slide_type_summary(g['slide_nums'], slide_lookup),
        })
    return enriched


def _vision_title_required_slides(slides):
    required = []
    for s in slides or []:
        try:
            num = int(s.get('slide_num'))
        except (TypeError, ValueError):
            continue
        if bool(s.get('title_from_image')):
            required.append(num)
    return sorted(set(required))


def _finalize_part(part_index, section_chunks):
    all_slides = []
    for chunk in section_chunks:
        all_slides.extend(chunk['slide_nums'])
    all_slides = sorted(set(all_slides))

    return {
        'part': part_index,
        'section_titles': [chunk['section_title'] for chunk in section_chunks],
        'slide_count': len(all_slides),
        'start_slide': min(all_slides),
        'end_slide': max(all_slides),
        'slide_nums': all_slides,
    }


def _build_split_plan(groups, min_slides=10, max_slides=15):
    chunks = _normalize_groups(groups)
    total_slides = len(sorted({n for g in chunks for n in g['slide_nums']}))

    if not chunks:
        return {
            'parts': [],
            'part_count': 0,
            'min_slides': min_slides,
            'max_slides': max_slides,
            'unsplittable_sections': [],
            'split_notes': 'No sections available to split.',
            'notes': ['No sections available to split.'],
        }

    if total_slides < min_slides:
        parts = [_finalize_part(1, chunks)]
        split_msg = (
            f'Total slides ({total_slides}) below min_slides threshold ({min_slides}). '
            'Kept as single part.'
        )
        return {
            'parts': parts,
            'part_count': len(parts),
            'min_slides': min_slides,
            'max_slides': max_slides,
            'unsplittable_sections': [],
            'split_notes': split_msg,
            'notes': [
                'Sections stay intact; splits happen only at section boundaries.',
                split_msg,
            ],
        }

    parts_raw = []
    current = []
    current_size = 0

    for chunk in chunks:
        size = len(chunk['slide_nums'])
        if not current:
            current = [chunk]
            current_size = size
            continue

        # Never split a section across parts; split only on section boundaries.
        if current_size >= min_slides and (current_size + size) > max_slides:
            parts_raw.append(current)
            current = [chunk]
            current_size = size
        else:
            current.append(chunk)
            current_size += size

    if current:
        parts_raw.append(current)

    if len(parts_raw) > 1:
        last_size = sum(len(section['slide_nums']) for section in parts_raw[-1])
        if last_size < min_slides:
            parts_raw[-2].extend(parts_raw[-1])
            parts_raw.pop()

    parts = [_finalize_part(i + 1, section_chunks) for i, section_chunks in enumerate(parts_raw)]

    unsplittable = []
    for chunk in chunks:
        chunk_size = len(chunk['slide_nums'])
        if chunk_size > max_slides:
            unsplittable.append({
                'section_title': chunk['section_title'],
                'slide_count': chunk_size,
            })

    notes = ['Sections stay intact; splits happen only at section boundaries.']
    if unsplittable:
        notes.append('Some sections exceed max slides and were kept intact to preserve narrative continuity.')

    split_msg = (
        f'Split into {len(parts)} part(s) using section boundaries '
        f'with min={min_slides} and max={max_slides} slides.'
    )

    return {
        'parts': parts,
        'part_count': len(parts),
        'min_slides': min_slides,
        'max_slides': max_slides,
        'unsplittable_sections': unsplittable,
        'split_notes': split_msg,
        'notes': notes,
    }


def _deterministic_fallback_structure(total):
    return {
        'toc': ['Main Content'],
        'groups': [
            {
                'section_title': 'Main Content',
                'slide_nums': list(range(1, total + 1)),
            }
        ],
        'fallback_reason': 'AI providers unavailable; applied deterministic single-section fallback.',
    }


def structure_slides(filename, split_options=None):
    split_cfg = _normalize_split_options(split_options)

    # Load stage 1 output first so cache validation can account for parse version.
    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    if not parsed:
        raise Exception('Stage 1 checkpoint not found. Run Stage 1 first.')
    slides = parsed['slides']
    total = len(slides)
    slide_lookup = _build_slide_lookup(slides)
    source_parse_policy_version = int(parsed.get('parse_policy_version', 0) or 0)

    # Load checkpoint but skip if it's a bad/error checkpoint
    base_struct = None
    cached_fallback = None
    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage2_structured', filename):
        cached = checkpoint_mgr.load('stage2_structured', filename)
        if cached is not None and 'error' not in cached:
            groups = cached.get('groups', [])
            has_empty = any(len(g.get('slide_nums', [])) == 0 for g in groups)
            unassigned = cached.get('unassigned_slides', [])
            if not has_empty and not unassigned and len(groups) > 0:
                cached_fallback = cached
            policy_ok = int(cached.get('structure_policy_version', 0) or 0) == STRUCTURE_POLICY_VERSION
            parse_policy_ok = int(cached.get('source_stage1_parse_policy_version', 0) or 0) == source_parse_policy_version
            if not has_empty and not unassigned and len(groups) > 0 and policy_ok and parse_policy_ok:
                print(f'Valid checkpoint found for {filename}, reusing logical sections.')
                base_struct = cached
        if base_struct is None:
            print(f'Invalid checkpoint for {filename}, re-running stage 2...')

    if base_struct is None:
        prompt = f"""
You are an expert presentation designer. Analyze the following slides and group them into logical sections.

STRICT RULES:
1. Every slide MUST be assigned to exactly one group. No slide can be skipped.
2. Every group MUST have at least one slide. No empty groups allowed.
3. The toc list must exactly match the section_title values in groups.
4. slide_nums must contain integers only. All slides from 1 to {total} must appear exactly once across all groups.
5. Keep narrative continuity: do not split a continuing topic into multiple section titles.
6. Prefer section boundaries that can later be bundled into ~10-15 slide parts.

Return ONLY a JSON object with this exact structure:
{{
  "toc": ["Section Title 1", "Section Title 2"],
  "groups": [
    {{"section_title": "Section Title 1", "slide_nums": [1, 2, 3]}},
    {{"section_title": "Section Title 2", "slide_nums": [4, 5, 6]}}
  ]
}}

Slides to group:
{json.dumps([
    {
        'slide_num': s['slide_num'],
        'title': s.get('title', ''),
        'title_from_image': bool(s.get('title_from_image', False)),
        'content_type_hint': s.get('content_type_hint', ''),
        'bullets': (s.get('bullets') or [])[:8],
        'raw_text': (s.get('raw_text') or '')[:300],
    }
    for s in slides
], ensure_ascii=False, indent=2)}

IMPORTANT: All {total} slides must be assigned. Start response with {{ and end with }}
"""

        try:
            raw = call_ai(prompt)
        except Exception as ex:
            if cached_fallback is not None:
                print(f'WARNING: Stage 2 AI unavailable ({ex}). Reusing prior logical groups.')
                base_struct = cached_fallback
            else:
                print(f'WARNING: Stage 2 AI unavailable ({ex}). Using deterministic fallback structure.')
                base_struct = _deterministic_fallback_structure(total)
            raw = None

        if base_struct is not None:
            normalized_groups = _normalize_groups(base_struct.get('groups', []))
            base_struct['groups'] = _enrich_groups(normalized_groups, slide_lookup)
            base_struct['toc'] = [g['section_title'] for g in base_struct['groups']]
            split_plan = _build_split_plan(
                base_struct.get('groups', []),
                min_slides=split_cfg['min_slides'],
                max_slides=split_cfg['max_slides'],
            )

            result = {
                **base_struct,
                'groups': base_struct['groups'],
                'split_config': split_cfg,
                'split_plan': split_plan,
                'vision_title_required_slides': _vision_title_required_slides(slides),
                'structure_policy_version': STRUCTURE_POLICY_VERSION,
                'source_stage1_parse_policy_version': source_parse_policy_version,
            }

            checkpoint_mgr.save('stage2_structured', filename, result)
            return result

        try:
            struct = parse_llm_json(raw)
        except Exception as ex:
            struct = {'error': 'Could not parse AI response', 'raw': raw, 'exception': str(ex)}
            checkpoint_mgr.save('stage2_structured', filename, struct)
            return struct

        normalized_groups = _normalize_groups(struct.get('groups', []))

        # Validate all slides assigned
        all_assigned = []
        for group in normalized_groups:
            all_assigned.extend(group.get('slide_nums', []))

        missing = [i for i in range(1, total + 1) if i not in all_assigned]
        if missing:
            print(f'WARNING: Slides {missing} were not assigned to any group!')
            struct['unassigned_slides'] = missing

        # Remove any empty groups and enrich with Stage 1-derived metadata.
        struct['groups'] = _enrich_groups(normalized_groups, slide_lookup)
        struct['toc'] = [g['section_title'] for g in struct['groups']]
        base_struct = struct

    enriched_groups = _enrich_groups(base_struct.get('groups', []), slide_lookup)

    split_plan = _build_split_plan(
        enriched_groups,
        min_slides=split_cfg['min_slides'],
        max_slides=split_cfg['max_slides'],
    )

    result = {
        **base_struct,
        'groups': enriched_groups,
        'split_config': split_cfg,
        'split_plan': split_plan,
        'vision_title_required_slides': _vision_title_required_slides(slides),
        'structure_policy_version': STRUCTURE_POLICY_VERSION,
        'source_stage1_parse_policy_version': source_parse_policy_version,
    }

    checkpoint_mgr.save('stage2_structured', filename, result)
    return result