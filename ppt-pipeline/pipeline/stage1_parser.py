import os
import pdfplumber
from .checkpoint import CheckpointManager

checkpoint_mgr = CheckpointManager()

def parse_pdf(filepath):
    filename = os.path.splitext(os.path.basename(filepath))[0]
    if checkpoint_mgr.exists('stage1_parsed', filename):
        return checkpoint_mgr.load('stage1_parsed', filename)

    slides = []
    with pdfplumber.open(filepath) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ''
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            title = lines[0] if lines else ''
            word_count = len(text.split())
            slides.append({
                'slide_num': i,
                'title': title,
                'raw_text': text,
                'word_count': word_count
            })
    result = {
        'filename': filename,
        'page_count': len(slides),
        'slides': slides
    }
    checkpoint_mgr.save('stage1_parsed', filename, result)
    return result
