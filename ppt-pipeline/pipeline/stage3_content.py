import re
import json
import os
from dotenv import load_dotenv
from pipeline.checkpoint import CheckpointManager
from pipeline.stage2_structurer import parse_llm_json, call_ai

load_dotenv()
checkpoint_mgr = CheckpointManager()


def generate_content(filename):
    """Stage 3: Audit slides, generate missing ones, write speaker notes."""

    # Check for valid checkpoint
    if checkpoint_mgr.exists('stage3_content', filename):
        cached = checkpoint_mgr.load('stage3_content', filename)
        if cached is not None and 'error' not in cached:
            print(f'Valid stage 3 checkpoint found for {filename}')
            return cached
        print(f'Invalid stage 3 checkpoint for {filename}, re-running...')

    # Load previous stages
    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    structured = checkpoint_mgr.load('stage2_structured', filename)
    if not parsed:
        raise Exception('Stage 1 checkpoint not found. Run Stage 1 first.')
    if not structured:
        raise Exception('Stage 2 checkpoint not found. Run Stage 2 first.')

    slides = parsed['slides']
    groups = structured['groups']
    toc = structured['toc']
    total = len(slides)

    # ── STEP 1: Audit existing slides ────────────────────────────────
    print('Stage 3 Step 1: Auditing slides...')
    audit_prompt = f"""You are an expert presentation designer auditing a slide deck.

Analyze these {total} slides and return a JSON audit report.

Slides:
{json.dumps([{'slide_num': s['slide_num'], 'title': s['title'], 'raw_text': s['raw_text'][:400]} for s in slides], ensure_ascii=False, indent=2)}

Current TOC sections: {json.dumps(toc)}

Return ONLY this JSON structure:
{{
  "presentation_title": "guessed title of the presentation",
  "has_title_slide": true,
  "has_agenda_slide": true,
  "has_conclusion_slide": false,
  "title_slide_num": 1,
  "agenda_slide_num": 2,
  "conclusion_slide_num": null,
  "content_gaps": ["brief description of any topic gaps you notice"],
  "missing_slides": ["conclusion"]
}}

Rules:
- missing_slides should only contain items from: "title", "agenda", "conclusion"
- Only include a type in missing_slides if it is genuinely absent from the deck
- Start with {{ and end with }}"""

    audit_raw = call_ai(audit_prompt)
    try:
        audit = parse_llm_json(audit_raw)
    except Exception as ex:
        raise Exception(f'Stage 3 audit parse failed: {ex}\nRaw: {audit_raw[:300]}')

    print(f'Audit complete: missing={audit.get("missing_slides", [])}')

    # ── STEP 2: Generate missing slides ──────────────────────────────
    print('Stage 3 Step 2: Generating missing slides...')
    missing_slides_content = []

    for slide_type in audit.get('missing_slides', []):
        print(f'  Generating {slide_type} slide...')
        gen_prompt = f"""You are an expert presentation designer.

Generate content for a "{slide_type}" slide for this presentation:
Title: {audit.get('presentation_title', 'AI Presentation')}
TOC sections: {json.dumps(toc)}
Existing slide titles: {json.dumps([s['title'] for s in slides])}

Return ONLY this JSON:
{{
  "slide_type": "{slide_type}",
  "title": "slide title text",
  "bullets": ["bullet 1", "bullet 2", "bullet 3"],
  "speaker_notes": "2-3 SHORT sentences of professional speaker notes, maximum 60 words total. Speak naturally as if presenting."
}}

Rules:
- title slide: title = presentation name, bullets = []
- agenda slide: bullets = list of TOC section names (max 6 items)
- conclusion slide: bullets = 4-5 key takeaways from the content
Start with {{ and end with }}"""

        gen_raw = call_ai(gen_prompt)
        try:
            generated = parse_llm_json(gen_raw)
            missing_slides_content.append(generated)
        except Exception as ex:
            print(f'  WARNING: Could not generate {slide_type} slide: {ex}')

    # ── STEP 3: Generate speaker notes for all existing slides ────────
    print('Stage 3 Step 3: Generating speaker notes...')
    speaker_notes = {}

    for slide in slides:
        slide_num = slide['slide_num']
        existing_notes = slide.get('notes_text', '').strip()

        if existing_notes and len(existing_notes) > 50:
            speaker_notes[str(slide_num)] = existing_notes
            print(f'  Slide {slide_num}: keeping existing notes')
        else:
            notes_prompt = f"""You are an expert presenter writing speaker notes for a slide deck video.

Slide {slide_num} of {total}
Title: {slide['title']}
Content: {slide['raw_text'][:500]}

Target: The narrator should finish this slide in 45-55 seconds.
Write EXACTLY 2-3 SHORT sentences of speaker notes. Maximum 70 words total.
Add context and transitions naturally — do NOT just repeat the bullets.

Return ONLY this JSON:
{{
  "slide_num": {slide_num},
  "speaker_notes": "your 2-3 sentence notes here, maximum 70 words"
}}
Start with {{ and end with }}"""

            notes_raw = call_ai(notes_prompt)
            try:
                notes_parsed = parse_llm_json(notes_raw)
                speaker_notes[str(slide_num)] = notes_parsed.get('speaker_notes', '')
                print(f'  Slide {slide_num}: generated notes')
            except Exception as ex:
                print(f'  Slide {slide_num}: notes generation failed: {ex}')
                speaker_notes[str(slide_num)] = ''

    # ── STEP 4: Build final manifest ──────────────────────────────────
    print('Stage 3 Step 4: Building final manifest...')

    insert_order = []
    for ms in missing_slides_content:
        t = ms.get('slide_type', '')
        if t == 'title':
            insert_order.append({'position': 'first', 'content': ms})
        elif t == 'agenda':
            insert_order.append({'position': 'second', 'content': ms})
        elif t == 'conclusion':
            insert_order.append({'position': 'last', 'content': ms})

    result = {
        'filename': filename,
        'audit': audit,
        'missing_slides_content': missing_slides_content,
        'insert_order': insert_order,
        'speaker_notes': speaker_notes,
        'original_slide_count': total,
        'final_slide_count': total + len(missing_slides_content)
    }

    checkpoint_mgr.save('stage3_content', filename, result)
    print(f'Stage 3 complete! Original: {total} slides -> Final: {result["final_slide_count"]} slides')
    return result
