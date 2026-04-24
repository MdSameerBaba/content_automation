"""
Stage 4 (v4): Build themed PPTX using reference.pptx as base.

Fixes:
- XML-level paragraph replacement (no double-content from reference)
- clean_body_lines(): strips bullet chars, filters orphans, joins continuations
- trim_notes(): caps notes at 100 words for ≤10min video target
- Formatting preserved via pPr/rPr template extraction from first para
"""

import os
import re
import copy
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled

checkpoint_mgr = CheckpointManager()

_PIPELINE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
REFERENCE_PATH = os.path.join(_PIPELINE_ROOT, 'theme', 'reference.pptx')
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage4_pptx')

BULLET_CHARS = '●•·○◦▪▸►▶–—*-'
BUILD_POLICY_VERSION = 11


# ── Text cleaning ─────────────────────────────────────────────────────

def clean_body_lines(lines):
    """
    Clean raw parsed lines for slide body content:
    1. Strip leading bullet symbols and whitespace
    2. Filter orphan lines (< 3 chars or pure numbers)
    3. Join continuation lines (lowercase-start short fragments) to previous
    Returns at most 8 lines.
    """
    cleaned = []
    for line in lines:
        # Strip leading bullet chars + spaces
        stripped = line.lstrip(BULLET_CHARS + ' ').strip()
        # Filter empty / too-short / pure-number orphans
        if len(stripped) < 3 or stripped.isdigit():
            continue
        cleaned.append(stripped)

    # Join continuation lines to the previous one
    if cleaned:
        joined = [cleaned[0]]
        for line in cleaned[1:]:
            prev = joined[-1]
            is_continuation = (
                line[0].islower()                            # starts lowercase
                or (len(line.split()) <= 3                  # very short fragment
                    and not prev.endswith(('.', '!', '?', ':')))
            )
            if is_continuation:
                joined[-1] = prev + ' ' + line
            else:
                joined.append(line)
        cleaned = joined

    return cleaned[:8]


def trim_notes(notes, max_words=100):
    """Cap speaker notes at max_words for ≤10-min video pacing."""
    if not notes:
        return notes
    words = notes.split()
    if len(words) <= max_words:
        return notes
    # Trim to last complete sentence within limit
    trimmed = ' '.join(words[:max_words])
    last_stop = max(trimmed.rfind('.'), trimmed.rfind('!'), trimmed.rfind('?'))
    if last_stop > len(trimmed) // 2:
        return trimmed[:last_stop + 1]
    return trimmed + '...'


def _clean_text(text):
    return re.sub(r'\s+', ' ', (text or '')).strip()


def _trim_title(title, max_words=14):
    title = _clean_text(title)
    if not title:
        return 'Untitled Slide'
    words = title.split()
    if len(words) <= max_words:
        return title
    return ' '.join(words[:max_words]).rstrip() + '...'


def _trim_bullets(bullets, max_items=7, max_words_per_line=18):
    cleaned = []
    for raw in bullets or []:
        line = _clean_text(raw)
        if not line:
            continue
        words = line.split()
        if len(words) > max_words_per_line:
            line = ' '.join(words[:max_words_per_line]).rstrip() + '...'
        cleaned.append(line)
        if len(cleaned) >= max_items:
            break
    return cleaned


# ── XML helpers ───────────────────────────────────────────────────────

def _get_format_templates(txBody):
    """Extract pPr and rPr from the first non-empty paragraph as formatting templates."""
    paras = txBody.findall(qn('a:p'))
    for first_p in paras:
        runs = first_p.findall(qn('a:r'))
        if not runs:
            continue
        pPr = first_p.find(qn('a:pPr'))
        rPr = runs[0].find(qn('a:rPr'))
        return (copy.deepcopy(pPr) if pPr is not None else None,
                copy.deepcopy(rPr) if rPr is not None else None)
    return None, None


def _xml_replace_text(txBody, text_lines):
    """
    Remove ALL existing <a:p> elements and replace with one per text line.
    Preserves run/paragraph formatting from the original first paragraph.
    """
    pPr_tmpl, rPr_tmpl = _get_format_templates(txBody)

    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)

    for line in (text_lines if text_lines else ['']):
        p_elem = etree.SubElement(txBody, qn('a:p'))
        if pPr_tmpl is not None:
            p_elem.insert(0, copy.deepcopy(pPr_tmpl))
        r_elem = etree.SubElement(p_elem, qn('a:r'))
        if rPr_tmpl is not None:
            r_elem.insert(0, copy.deepcopy(rPr_tmpl))
        t_elem = etree.SubElement(r_elem, qn('a:t'))
        t_elem.text = line


def _set_ph_text(slide, idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            _xml_replace_text(ph.text_frame._txBody, [text])
            return True
    return False


def _set_ph_bullets(slide, idx, lines):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            _xml_replace_text(ph.text_frame._txBody, lines if lines else [''])
            return True
    return False


def _set_notes(slide, text):
    if not text:
        return
    try:
        tf = slide.notes_slide.notes_text_frame
        _xml_replace_text(tf._txBody, [text])
    except Exception as e:
        print(f'    warn: notes failed: {e}')


def _find_layout(prs, preferred_names):
    for name in preferred_names:
        for layout in prs.slide_layouts:
            if name.lower() in layout.name.lower():
                return layout
    return prs.slide_layouts[min(1, len(prs.slide_layouts) - 1)]


def _resolve_body_placeholder_idx(slide, slide_type):
    if slide_type == 'title':
        return None

    left_tolerance = 91440 * 3  # ~0.3 inch in EMU
    candidates = []

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            continue
        if not hasattr(ph, 'text_frame'):
            continue

        left = int(getattr(ph, 'left', 0) or 0)
        width = int(getattr(ph, 'width', 0) or 0)
        height = int(getattr(ph, 'height', 0) or 0)
        area = width * height
        has_text = bool(_clean_text(getattr(ph, 'text', '')))
        ph_type = int(ph.placeholder_format.type)

        candidates.append({
            'idx': idx,
            'left': left,
            'area': area,
            'has_text': has_text,
            'is_body': ph_type == 2,
        })

    if not candidates:
        return None

    min_left = min(c['left'] for c in candidates)
    left_band = [c for c in candidates if c['left'] <= min_left + left_tolerance]
    pool = left_band if left_band else candidates

    pool.sort(
        key=lambda c: (
            1 if c['has_text'] else 0,
            c['area'],
            1 if c['is_body'] else 0,
            -c['left'],
        ),
        reverse=True,
    )
    return pool[0]['idx']


def _remove_slide(prs, slide_index):
    """Remove slide by index using python-pptx internals for dynamic final deck length."""
    sld_id_lst = prs.slides._sldIdLst
    sld_id = sld_id_lst[slide_index]
    rel_id = sld_id.rId
    prs.part.drop_rel(rel_id)
    del sld_id_lst[slide_index]


def _normalize_blueprint_entry(item, index):
    slide_type = _clean_text(item.get('type', 'content')).lower() or 'content'
    if slide_type not in {'title', 'agenda', 'content', 'image_content', 'diagram', 'conclusion'}:
        slide_type = 'content'

    title = _trim_title(item.get('title') or f'Slide {index}')
    bullets = _trim_bullets(item.get('bullets', []))
    notes = trim_notes(_clean_text(item.get('speaker_notes') or item.get('notes', '')))

    return {
        'type': slide_type,
        'title': title,
        'bullets': bullets,
        'notes': notes,
        'source_slide_num': item.get('source_slide_num'),
        'title_inferred': bool(item.get('title_inferred', False)),
        'embed_source_image': bool(item.get('embed_source_image', False)),
        'image_path': item.get('image_path') or '',
    }


def _ordered_from_typed_blueprint(stage3):
    typed = stage3.get('typed_blueprint')
    if not isinstance(typed, list) or not typed:
        return None

    ordered = []
    for idx, item in enumerate(typed, start=1):
        if not isinstance(item, dict):
            continue
        ordered.append(_normalize_blueprint_entry(item, idx))
    return ordered if ordered else None


def _ordered_from_legacy(parsed, stage3):
    slides_data = parsed['slides']
    insert_order = stage3.get('insert_order', [])
    slide_bullets = stage3.get('slide_bullets', {})
    speaker_notes = stage3.get('speaker_notes', {})
    missing_by_type = {m['content']['slide_type']: m['content'] for m in insert_order if isinstance(m, dict) and m.get('content')}

    ordered = []
    if 'title' in missing_by_type:
        tc = missing_by_type['title']
        ordered.append(_normalize_blueprint_entry({
            'type': 'title',
            'title': tc.get('title'),
            'bullets': tc.get('bullets', []),
            'speaker_notes': tc.get('speaker_notes', ''),
        }, len(ordered) + 1))

    if 'agenda' in missing_by_type:
        ac = missing_by_type['agenda']
        ordered.append(_normalize_blueprint_entry({
            'type': 'agenda',
            'title': ac.get('title', 'Agenda'),
            'bullets': ac.get('bullets', []),
            'speaker_notes': ac.get('speaker_notes', ''),
        }, len(ordered) + 1))

    for s in slides_data:
        slide_num = s['slide_num']
        title_text = s.get('title') or f'Slide {slide_num}'
        raw_text = s.get('raw_text', '')
        notes_raw = speaker_notes.get(str(slide_num), '')

        raw_lines = [ln.strip() for ln in raw_text.split('\n') if ln.strip()]
        if raw_lines and _clean_text(raw_lines[0]).lower() == _clean_text(title_text).lower():
            raw_lines = raw_lines[1:]

        ai_body = slide_bullets.get(str(slide_num), [])
        body = ai_body if isinstance(ai_body, list) and ai_body else clean_body_lines(raw_lines)
        archetype = 'image_content' if s.get('image_path') and int(s.get('word_count', 0) or 0) <= 24 else 'content'

        ordered.append(_normalize_blueprint_entry({
            'type': archetype,
            'source_slide_num': slide_num,
            'title': title_text,
            'bullets': body,
            'speaker_notes': notes_raw,
        }, len(ordered) + 1))

    if 'conclusion' in missing_by_type:
        cc = missing_by_type['conclusion']
        ordered.append(_normalize_blueprint_entry({
            'type': 'conclusion',
            'title': cc.get('title', 'Conclusion'),
            'bullets': cc.get('bullets', []),
            'speaker_notes': cc.get('speaker_notes', ''),
        }, len(ordered) + 1))

    return ordered


def _build_ordered_manifest(parsed, stage3):
    typed = _ordered_from_typed_blueprint(stage3)
    if typed:
        return typed
    return _ordered_from_legacy(parsed, stage3)


# ── Main builder ──────────────────────────────────────────────────────

def build_pptx(filename):
    """Stage 4: build themed PPTX from stage3 manifest using reference.pptx."""

    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    stage3 = checkpoint_mgr.load('stage3_content', filename)

    if not parsed:
        raise Exception('Stage 1 checkpoint missing.')
    if not stage3:
        raise Exception('Stage 3 checkpoint missing.')

    source_blueprint_version = int(stage3.get('typed_blueprint_version', 0) or 0)
    source_notes_policy_version = int(stage3.get('notes_policy_version', 0) or 0)

    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage4_pptx', filename):
        cached = checkpoint_mgr.load('stage4_pptx', filename)
        if (
            cached
            and 'error' not in cached
            and cached.get('build_policy_version') == BUILD_POLICY_VERSION
            and int(cached.get('source_stage3_blueprint_version', 0) or 0) == source_blueprint_version
            and int(cached.get('source_stage3_notes_policy_version', 0) or 0) == source_notes_policy_version
        ):
            print(f'Valid stage 4 checkpoint for {filename}')
            return cached

    ordered = _build_ordered_manifest(parsed, stage3)

    print(f'Building {len(ordered)} slides...')

    # ── Open reference and modify in-place ───────────────────────────
    if not os.path.exists(REFERENCE_PATH):
        raise FileNotFoundError(
            f'Reference template not found: {REFERENCE_PATH}. '
            'Expected file at ppt-pipeline/theme/reference.pptx'
        )
    prs = Presentation(REFERENCE_PATH)
    ref_slides = list(prs.slides)
    print(f'Reference: {len(ref_slides)} slides')

    # The reference deck is a style scaffold, not a fixed output length.
    # Trim extra scaffold slides so final output follows generated manifest length.
    if len(ref_slides) > len(ordered):
        for idx in range(len(ref_slides) - 1, len(ordered) - 1, -1):
            _remove_slide(prs, idx)
        ref_slides = list(prs.slides)
        print(f'Trimmed reference scaffold to {len(ref_slides)} slides for dynamic output length')

    slide_manifest = []

    for i, item in enumerate(ordered):
        title = item['title']
        bullets = item.get('bullets', [])
        notes = item.get('notes', '')
        slide_type = item.get('type', 'content')

        if i < len(ref_slides):
            slide = ref_slides[i]
        else:
            if slide_type == 'title':
                layout = _find_layout(prs, ['title', 'cover'])
            elif slide_type == 'diagram':
                layout = _find_layout(prs, ['SECTION_TITLE_AND_DESCRIPTION', 'section', 'TITLE_AND_BODY'])
            else:
                layout = _find_layout(prs, ['TITLE AND CONTENT', 'content', 'TITLE_AND_BODY'])
            slide = prs.slides.add_slide(layout)

        ph_indices = [ph.placeholder_format.idx for ph in slide.placeholders]
        print(f'  [{i+1}] {slide_type} "{title[:50]}" | phs={ph_indices} | {len(bullets)} bullets')

        effective_bullets = _trim_bullets(bullets)

        # Apply content to slide
        _set_ph_text(slide, 0, title)

        # Clear both body placeholders first to avoid reference bleed-through.
        _set_ph_bullets(slide, 1, [''])
        _set_ph_bullets(slide, 2, [''])

        target_body_idx = _resolve_body_placeholder_idx(slide, slide_type)
        if target_body_idx is not None:
            _set_ph_bullets(slide, target_body_idx, effective_bullets if effective_bullets else [''])

        _set_notes(slide, notes)

        slide_manifest.append({
            'slide_num': i + 1,
            'type': slide_type,
            'title': title,
            'bullets': effective_bullets,
            'notes': notes,
            'source_slide_num': item.get('source_slide_num'),
            'title_inferred': bool(item.get('title_inferred', False)),
            'embed_source_image': bool(item.get('embed_source_image', False)),
            'image_path': item.get('image_path') or '',
        })

    # Save the PPTX
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f'{filename}.pptx')
    prs.save(out_path)
    print(f'Saved PPTX to: {out_path}')

    # Calculate totals
    total_words = sum(len(item['notes'].split()) for item in ordered if item.get('notes'))
    est_mins = total_words / 130

    print(f'\nNarrator: ~{total_words} words -> ~{est_mins:.1f} min at 130 wpm')
    final_slide_count = len(prs.slides)
    print(f'Stage 4 complete: {final_slide_count} slides -> {out_path}')

    result = {
        'filename': filename,
        'output_path': out_path,
        'total_slides': final_slide_count,
        'narrator_words': total_words,
        'estimated_minutes': round(est_mins, 2),
        'build_policy_version': BUILD_POLICY_VERSION,
        'source_stage3_blueprint_version': source_blueprint_version,
        'source_stage3_notes_policy_version': source_notes_policy_version,
        'slide_manifest': slide_manifest,
    }
    checkpoint_mgr.save('stage4_pptx', filename, result)
    return result
