"""Stage 5: Export PPTX slides to PNG images using LibreOffice or comtypes (Windows COM)."""

import os
import glob
import shutil
import subprocess
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled

checkpoint_mgr = CheckpointManager()
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage5_images')


def _run_command(cmd, timeout, context):
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    if result.returncode != 0:
        raise RuntimeError(f'{context} failed (exit {result.returncode}): {result.stderr[:500]}')
    return result


def _get_input_pptx(filename):
    """
    Return the PPTX path to use for export (priority: human-revised > AI-generated).
    """
    revised = os.path.join(checkpoint_mgr.base_dir, 'stage5_input', f'{filename}.pptx')
    if os.path.exists(revised):
        print(f'  Using human-revised PPTX: {revised}')
        return revised
    ai_gen = os.path.join(checkpoint_mgr.base_dir, 'stage4_pptx', f'{filename}.pptx')
    if os.path.exists(ai_gen):
        print(f'  Using AI-generated PPTX: {ai_gen}')
        return ai_gen
    raise FileNotFoundError(f'No PPTX found for "{filename}". Run Stage 4 or upload a revised PPTX.')


def _export_via_libreoffice(pptx_path, out_dir):
    """Export each slide as PNG using LibreOffice headless."""
    candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        'soffice',
    ]
    soffice = None
    for c in candidates:
        if os.path.exists(c) or shutil.which(c):
            soffice = c
            break
    if not soffice:
        raise EnvironmentError('LibreOffice not found. Please install from https://www.libreoffice.org/')

    # First convert to PDF (LibreOffice handles multi-page better)
    import tempfile
    temp_dir = tempfile.mkdtemp()
    pdf_path = ''
    
    cmd = [soffice, '--headless', '--convert-to', 'pdf', '--outdir', temp_dir, pptx_path]
    print(f'  Converting to PDF: {" ".join(cmd)}')
    pdf_result = _run_command(cmd, timeout=120, context='LibreOffice PDF conversion')

    generated_pdfs = sorted(glob.glob(os.path.join(temp_dir, '*.pdf')))
    if generated_pdfs:
        pdf_path = generated_pdfs[0]
    
    if not pdf_path or not os.path.exists(pdf_path):
        listing = ', '.join(os.listdir(temp_dir)) if os.path.exists(temp_dir) else '<temp dir missing>'
        raise RuntimeError(
            f'PDF not created by LibreOffice in {temp_dir}. '
            f'Directory contents: [{listing}]. '
            f'stdout: {pdf_result.stdout[:300]} stderr: {pdf_result.stderr[:300]}'
        )
    
    # Now convert PDF to PNG using pdf2image or similar
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=200)
        for i, img in enumerate(images, 1):
            img.save(os.path.join(out_dir, f'slide_{i:02d}.png'), 'PNG')
        print(f'  Exported {len(images)} slides via PDF')
    except ImportError:
        # Fallback: try direct PNG conversion with better options
        cmd = [soffice, '--headless', '--convert-to', 'png', '--outdir', out_dir, pptx_path]
        print(f'  Running: {" ".join(cmd)}')
        _run_command(cmd, timeout=120, context='LibreOffice PNG conversion')
    
    # Clean up temp
    try:
        for tmp_pdf in glob.glob(os.path.join(temp_dir, '*.pdf')):
            if os.path.exists(tmp_pdf):
                os.remove(tmp_pdf)
        if os.path.exists(temp_dir):
            os.rmdir(temp_dir)
    except OSError:
        # Non-fatal cleanup issue.
        pass
    
    return True


def _export_via_com(pptx_path, out_dir):
    """Export each slide as PNG using PowerPoint COM automation (Windows only)."""
    try:
        import comtypes.client
    except ImportError:
        raise EnvironmentError('comtypes not installed. Run: pip install comtypes')

    import comtypes.client
    pptx_abs = os.path.abspath(pptx_path)
    out_abs  = os.path.abspath(out_dir)

    ppt_app = comtypes.client.CreateObject('PowerPoint.Application')
    ppt_app.Visible = 1
    try:
        prs = ppt_app.Presentations.Open(pptx_abs, ReadOnly=True, WithWindow=False)
        slide_count = prs.Slides.Count
        print(f'  COM: {slide_count} slides found')
        for i in range(1, slide_count + 1):
            out_file = os.path.join(out_abs, f'slide_{i:02d}.png')
            prs.Slides(i).Export(out_file, 'PNG', 1920, 1080)
            print(f'    Exported slide {i}/{slide_count}')
        prs.Close()
        return slide_count
    finally:
        ppt_app.Quit()


def export_images(filename):
    """Stage 5: export each PPTX slide to a PNG image."""

    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage5_images', filename):
        cached = checkpoint_mgr.load('stage5_images', filename)
        if cached and 'error' not in cached:
            print(f'Valid Stage 5 checkpoint for {filename}')
            return cached

    pptx_path = _get_input_pptx(filename)
    out_dir   = os.path.join(OUTPUT_DIR, filename)
    os.makedirs(out_dir, exist_ok=True)

    # Clean old PNGs
    for f in glob.glob(os.path.join(out_dir, '*.png')):
        os.remove(f)

    # Try COM first (Windows), fall back to LibreOffice
    slide_count = 0
    method_used = 'unknown'
    export_errors = []
    try:
        print('  Trying PowerPoint COM export...')
        slide_count = _export_via_com(pptx_path, out_dir)
        method_used = 'PowerPoint COM'
    except Exception as com_err:
        export_errors.append(f'COM export failed: {com_err}')
        print(f'  COM failed ({com_err}), trying LibreOffice...')
        try:
            _export_via_libreoffice(pptx_path, out_dir)
            method_used = 'LibreOffice'
        except Exception as lo_err:
            export_errors.append(f'LibreOffice export failed: {lo_err}')
            raise RuntimeError('Stage 5 image export failed. ' + ' | '.join(export_errors)) from lo_err

    # Collect and rename exported PNGs to slide_01.png format
    pngs = sorted(glob.glob(os.path.join(out_dir, '*.png')))

    # LibreOffice names them <stem>1.png, <stem>2.png etc. — rename to slide_XX.png
    renamed = []
    for idx, src in enumerate(pngs, 1):
        dst = os.path.join(out_dir, f'slide_{idx:02d}.png')
        if src != dst:
            os.rename(src, dst)
        renamed.append(dst)
    slide_count = len(renamed)

    if slide_count == 0:
        raise RuntimeError('Stage 5 produced zero PNGs. Check PowerPoint/LibreOffice export dependencies.')

    print(f'  Exported {slide_count} slides via {method_used} -> {out_dir}')

    result = {
        'filename': filename,
        'slide_count': slide_count,
        'output_dir': out_dir,
        'method': method_used,
        'images': renamed,
    }
    checkpoint_mgr.save('stage5_images', filename, result)
    return result
