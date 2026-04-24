import os
import re
import pdfplumber
from .checkpoint import CheckpointManager, is_cache_reuse_enabled

checkpoint_mgr = CheckpointManager()
PARSE_POLICY_VERSION = 3

BULLET_RE = re.compile(r'^[\u2022\u25cf\u25cb\u25e6\u25aa\u25b8\u25ba\u25c0\u25b6\u25b7\u2013\u2014\-*]\s*(.+)$')
NUMBERED_RE = re.compile(r'^\d+[\.)]\s+(.+)$')
MARKER_ONLY_RE = re.compile(r'^[\u2022\u25cf\u25cb\u25e6\u25aa\u25b8\u25ba\u25c0\u25b6\u25b7\u2013\u2014\-*]+$')


def _content_type_hint(word_count):
    if word_count == 0:
        return 'image_only'
    if word_count < 10:
        return 'likely_diagram'
    return 'text_heavy'


def _estimate_page_image_size(page, resolution):
    # PDF page dimensions are in points (72 dpi).
    width_px = int(round(float(page.width) * float(resolution) / 72.0))
    height_px = int(round(float(page.height) * float(resolution) / 72.0))
    return max(1, width_px), max(1, height_px)


def _read_image_dimensions(image_path):
    try:
        from PIL import Image
        with Image.open(image_path) as img:
            return int(img.width), int(img.height)
    except Exception:
        return None, None


def _extract_bullets(lines, title):
    bullets = []
    normalized_title = (title or '').strip().lower()
    candidate_lines = []
    last_plain_line = ''

    for i, line in enumerate(lines):
        cleaned = line.strip()
        if not cleaned:
            continue
        if i == 0 and normalized_title and cleaned.lower() == normalized_title:
            continue

        match = BULLET_RE.match(cleaned) or NUMBERED_RE.match(cleaned)
        if match:
            bullet_text = match.group(1).strip()
            if bullet_text:
                bullets.append(bullet_text)
            continue

        if MARKER_ONLY_RE.match(cleaned):
            if last_plain_line and (not bullets or bullets[-1] != last_plain_line):
                bullets.append(last_plain_line)
            continue

        # Join wrapped lines to the previous bullet when likely continuation.
        if bullets:
            prev = bullets[-1]
            starts_lower = cleaned[0].islower()
            prev_continuation = prev.endswith((':', ',', 'and', 'or', 'to', 'of', 'the'))
            if starts_lower or prev_continuation:
                bullets[-1] = f'{prev} {cleaned}'.strip()
                last_plain_line = cleaned
                continue

        if not cleaned.isdigit():
            candidate_lines.append(cleaned)
            last_plain_line = cleaned

    if bullets:
        return bullets

    # Fallback: treat list-like short/medium lines as bullets when explicit markers are absent.
    fallback = []
    for ln in candidate_lines:
        words = ln.split()
        if 3 <= len(words) <= 20:
            fallback.append(ln)
        if len(fallback) >= 8:
            break

    return fallback


def _export_slide_image(page, out_path, resolution=140):
    """Best-effort PNG export for a PDF page used as visual context later in the pipeline."""
    est_width, est_height = _estimate_page_image_size(page, resolution)
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        img = page.to_image(resolution=resolution)
        img.save(out_path, format='PNG')
        if os.path.exists(out_path):
            actual_w, actual_h = _read_image_dimensions(out_path)
            return out_path, (actual_w or est_width), (actual_h or est_height)
        return '', est_width, est_height
    except Exception:
        return '', est_width, est_height

def parse_pdf(filepath):
    filename = os.path.splitext(os.path.basename(filepath))[0]
    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage1_parsed', filename):
        cached = checkpoint_mgr.load('stage1_parsed', filename)
        if (
            isinstance(cached, dict)
            and cached.get('parse_policy_version') == PARSE_POLICY_VERSION
            and isinstance(cached.get('slides'), list)
        ):
            return cached
        print(f'Stage 1 checkpoint for {filename} is outdated; re-parsing PDF...')

    slides = []
    image_dir = os.path.join(checkpoint_mgr.base_dir, 'stage1_parsed', filename)
    with pdfplumber.open(filepath) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ''
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            title = lines[0] if lines else ''
            word_count = len(text.split())
            bullets = _extract_bullets(lines, title)
            image_path, slide_width_px, slide_height_px = _export_slide_image(
                page,
                os.path.join(image_dir, f'slide_{i:02d}.png')
            )
            title_from_image = bool(image_path and not title.strip())
            slides.append({
                'slide_num': i,
                'title': title,
                'title_from_image': title_from_image,
                'raw_text': text,
                'bullets': bullets,
                'word_count': word_count,
                'content_type_hint': _content_type_hint(word_count),
                'slide_width_px': slide_width_px,
                'slide_height_px': slide_height_px,
                'image_path': image_path,
            })
    result = {
        'filename': filename,
        'page_count': len(slides),
        'slides': slides,
        'parse_policy_version': PARSE_POLICY_VERSION,
    }
    checkpoint_mgr.save('stage1_parsed', filename, result)
    return result
