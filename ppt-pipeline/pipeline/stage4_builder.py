"""
Stage 4 (v4): Build themed PPTX using reference.pptx as base.

Fixes:
- XML-level paragraph replacement (no double-content from reference)
- clean_body_lines(): strips bullet chars, filters orphans, joins continuations
- Empty body fallback: uses speaker notes when pdfplumber got nothing
- trim_notes(): caps notes at 100 words for ≤10min video target
- Formatting preserved via pPr/rPr template extraction from first para
"""

import os
import re
import copy
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pipeline.checkpoint import CheckpointManager

checkpoint_mgr = CheckpointManager()

REFERENCE_PATH = os.path.join('theme', 'reference.pptx')
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage4_pptx')

BULLET_CHARS = '●•·○◦▪▸►▶–—*-'


# ── Text cleaning ─────────────────────────────────────────────────────

def clean_body_lines(lines, fallback_notes=''):
    """
    Clean raw parsed lines for slide body content:
    1. Strip leading bullet symbols and whitespace
    2. Filter orphan lines (< 3 chars or pure numbers)
    3. Join continuation lines (lowercase-start short fragments) to previous
    4. Fall back to first 3 sentences of speaker notes if result is empty
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

    # Fall back to speaker notes if body is still empty
    if not cleaned and fallback_notes:
        sentences = re.split(r'(?<=[.!?])\s+', fallback_notes.strip())
        # Split long sentences into ≤ 15-word chunks for readability
        for s in sentences[:4]:
            s = s.strip()
            if len(s) > 10:
                words = s.split()
                if len(words) > 15:
                    # chunk into two
                    mid = len(words) // 2
                    cleaned.append(' '.join(words[:mid]))
                    cleaned.append(' '.join(words[mid:]))
                else:
                    cleaned.append(s)
            if len(cleaned) >= 4:
                break

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


# ── Main builder ──────────────────────────────────────────────────────

def build_pptx(filename):
    """Stage 4: build themed PPTX from stage3 manifest using reference.pptx."""

    if checkpoint_mgr.exists('stage4_pptx', filename):
        cached = checkpoint_mgr.load('stage4_pptx', filename)
        if cached and 'error' not in cached:
            print(f'Valid stage 4 checkpoint for {filename}')
            return cached

    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    stage3 = checkpoint_mgr.load('stage3_content', filename)

    if not parsed:
        raise Exception('Stage 1 checkpoint missing.')
    if not stage3:
        raise Exception('Stage 3 checkpoint missing.')

    slides_data     = parsed['slides']
    insert_order    = stage3['insert_order']
    speaker_notes   = stage3['speaker_notes']
    missing_by_type = {m['content']['slide_type']: m['content'] for m in insert_order}

    # ── Build ordered manifest ────────────────────────────────────────
    ordered = []

    if 'title' in missing_by_type:
        tc = missing_by_type['title']
        ordered.append({
            'type': 'title',
            'title': tc['title'],
            'bullets': tc.get('bullets', []),
            'notes': trim_notes(tc.get('speaker_notes', '')),
        })

    for s in slides_data:
        slide_num  = s['slide_num']
        title_text = s.get('title') or f'Slide {slide_num}'
        raw_text   = s.get('raw_text', '')
        notes_raw  = speaker_notes.get(str(slide_num), '')

        # Raw lines from pdfplumber
        raw_lines = [ln.strip() for ln in raw_text.split('\n') if ln.strip()]
        # Drop first line if it duplicates the title
        if raw_lines and raw_lines[0].lower() == title_text.lower():
            raw_lines = raw_lines[1:]

        body = clean_body_lines(raw_lines, fallback_notes=notes_raw)
        notes = trim_notes(notes_raw)

        # Fallback bullets from notes (used only when ref ph1 had text but pdfplumber got nothing)
        notes_sentences = re.split(r'(?<=[.!?])\s+', notes_raw.strip())
        fallback_bullets = [s.strip() for s in notes_sentences[:3] if len(s.strip()) > 10]

        ordered.append({
            'type': 'content',
            'slide_num': slide_num,
            'title': title_text,
            'bullets': body,
            'fallback_bullets': fallback_bullets,
            'notes': notes,
        })

    if 'conclusion' in missing_by_type:
        cc = missing_by_type['conclusion']
        ordered.append({
            'type': 'conclusion',
            'title': cc['title'],
            'bullets': cc.get('bullets', []),
            'notes': trim_notes(cc.get('speaker_notes', '')),
        })

    print(f'Building {len(ordered)} slides...')

    # ── Open reference and modify in-place ───────────────────────────
    prs = Presentation(REFERENCE_PATH)
    ref_slides = list(prs.slides)
    print(f'Reference: {len(ref_slides)} slides')

    slide_manifest = []

    for i, item in enumerate(ordered):
        title   = item['title']
        bullets = item['bullets']
        notes   = item['notes']

        if i < len(ref_slides):
            slide = ref_slides[i]
        else:
            layout = _find_layout(prs, ['TITLE_AND_BODY', 'TITLE AND CONTENT', 'content'])
            slide  = prs.slides.add_slide(layout)

        # ── Read reference state BEFORE modifying this slide ───────────
        ref_ph0_text     = ''
        ref_ph1_has_text = False
        for ph in slide.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 0:
                ref_ph0_text = ph.text_frame.text.strip()
            elif idx == 1:
                ref_ph1_has_text = bool(ph.text_frame.text.strip())

        # ── FIX 1: Title fallback — use reference title if pdfplumber ──
        # returned "Slide N" (i.e. it couldn't extract the real title)
        if re.match(r'^Slide \d+$', title) and ref_ph0_text:
            title = ref_ph0_text
            print(f'    INFO: fallback title from reference: "{title}"')

        ph_indices = [ph.placeholder_format.idx for ph in slide.placeholders]
        print(f'  [{i+1}] {item["type"]} "{title[:50]}" | phs={ph_indices} | {len(bullets)} bullets')

        # ── FIX 3: Agenda slide — replace bullets with actual slide titles
        if title.lower().strip() in ('agenda', 'table of contents', 'toc'):
            seen, agenda_items = set(), []
            for s in slides_data:
                t = s.get('title', '').strip()
                if t and t.lower() not in ('agenda', 'table of contents', 'toc', '') and t.lower() not in seen:
                    seen.add(t.lower())
                    agenda_items.append(t)
            effective_bullets = agenda_items[:7]
            print(f'    INFO: agenda slide — using {len(effective_bullets)} actual slide titles')

        # ── FIX 2: Image vs text slide decision ───────────────────────
        if title.lower().strip() in ('agenda', 'table of contents', 'toc'):
            # Agenda handled above, effective_bullets already set
            pass
        elif ref_ph1_has_text:
            # This slide is expected to have body text
            # Use parsed content; if empty, fall back to notes sentences
            effective_bullets = bullets or item.get('fallback_bullets', [])
            if not bullets:
                print(f'    INFO: text-slide empty body — using notes fallback ({len(effective_bullets)} lines)')
        else:
            # Reference ph1 was empty -> visual/image slide
            effective_bullets = []
            print(f'    INFO: image slide — no body text')

        # Apply content to slide
        _set_ph_text(slide, 0, title)
        _set_ph_bullets(slide, 1, effective_bullets)
        _set_notes(slide, notes)

        slide_manifest.append({
            'slide_num': i + 1,
            'type': item['type'],
            'title': title,
            'bullets': effective_bullets,
            'notes': notes,
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
    print(f'Stage 4 complete: {len(ordered)} slides -> {out_path}')

    result = {
        'filename': filename,
        'output_path': out_path,
        'total_slides': len(ordered),
        'narrator_words': total_words,
        'estimated_minutes': round(est_mins, 2),
        'slide_manifest': slide_manifest,
    }
    checkpoint_mgr.save('stage4_pptx', filename, result)
    return result
