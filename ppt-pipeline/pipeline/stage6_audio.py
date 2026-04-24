"""Stage 6: Generate narration audio (MP3) from PPTX speaker notes using Kokoro with voice selection."""

import os
import subprocess
import imageio_ffmpeg
from pptx import Presentation
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled
from pipeline.config import get_tts_config

# Available Kokoro voices
KOKORO_VOICES = {
    'af_heart': {'name': 'Heart (Female)', 'gender': 'Female'},
    'af_sarah': {'name': 'Sarah (Female)', 'gender': 'Female'},
    'am_adam': {'name': 'Adam (Male)', 'gender': 'Male'},
    'am_michael': {'name': 'Michael (Male)', 'gender': 'Male'},
    'bf_emma': {'name': 'Emma (British Female)', 'gender': 'Female'},
    'bm_george': {'name': 'George (British Male)', 'gender': 'Male'},
    'af_nova': {'name': 'Nova (Female)', 'gender': 'Female'},
}

TTS_CONFIG = get_tts_config()
DEFAULT_VOICE = TTS_CONFIG['default_voice']
if DEFAULT_VOICE not in KOKORO_VOICES:
    DEFAULT_VOICE = 'af_heart'

# Sample text for voice preview
PREVIEW_TEXT = TTS_CONFIG['preview_text']
DEFAULT_SILENCE_MS = TTS_CONFIG['silence_duration_ms']

checkpoint_mgr = CheckpointManager()
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage6_audio')


def _run_command(command, timeout):
    result = subprocess.run(command, capture_output=True, timeout=timeout, text=True)
    if result.returncode != 0:
        raise RuntimeError(f'Command failed ({result.returncode}): {" ".join(command)}\n{result.stderr[:300]}')
    return result


def _get_input_pptx(filename):
    revised = os.path.join(checkpoint_mgr.base_dir, 'stage5_input', f'{filename}.pptx')
    if os.path.exists(revised):
        return revised
    ai_gen = os.path.join(checkpoint_mgr.base_dir, 'stage4_pptx', f'{filename}.pptx')
    if os.path.exists(ai_gen):
        return ai_gen
    raise FileNotFoundError(f'No PPTX found for "{filename}".')


def _extract_notes_from_pptx(pptx_path):
    """Read speaker notes from each slide of the PPTX."""
    prs = Presentation(pptx_path)
    notes = {}
    for idx, slide in enumerate(prs.slides, 1):
        try:
            text = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            text = ''
        notes[idx] = text
    return notes


def _generate_gtts(text, out_path):
    """Generate audio using gTTS (Google TTS) - free, reliable."""
    try:
        from gtts import gTTS
        tts = gTTS(text=text, lang='en', slow=False)
        tts.save(out_path)
        return os.path.exists(out_path) and os.path.getsize(out_path) > 100
    except Exception as e:
        print(f'    gTTS failed: {e}')
    return False


def _generate_kokoro(text, out_path, voice=DEFAULT_VOICE):
    """Generate audio using Kokoro TTS with specified voice."""
    wav_path = out_path.replace('.mp3', '.wav')
    try:
        from kokoro import KPipeline
        import soundfile as sf
        import numpy as np
        
        pipeline = KPipeline(lang_code='a', repo_id='hexgrad/Kokoro-82M')
        generator = pipeline(text, voice=voice, speed=1, split_pattern=r'\n+')
        
        audio_chunks = []
        for i, (gs, ps, audio) in enumerate(generator):
            audio_chunks.append(audio)
        
        if audio_chunks:
            full_audio = np.concatenate(audio_chunks)
            sf.write(wav_path, full_audio, 24000)
            
            # Use ffmpeg to convert to mp3
            ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
            _run_command([
                ffmpeg_path, '-i', wav_path,
                '-acodec', 'libmp3lame', '-y', out_path
            ], timeout=30)
            
            return os.path.exists(out_path)
    except Exception as e:
        print(f'    Kokoro failed: {e}')
    finally:
        if os.path.exists(wav_path):
            try:
                os.remove(wav_path)
            except OSError:
                pass
    return False


def _generate_piper(text, out_path):
    """Generate audio using Piper TTS as fallback."""
    text_file = None
    wav_path = out_path.replace('.mp3', '.wav')
    try:
        import tempfile
        text_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False)
        text_file.write(text)
        text_file.close()
        
        _run_command([
            'piper', '--text_file', text_file.name,
            '--output_file', wav_path
        ], timeout=30)
        
        ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
        _run_command([
            ffmpeg_path, '-i', wav_path,
            '-acodec', 'libmp3lame', '-y', out_path
        ], timeout=30)
        
        return True
    except Exception as e:
        print(f'    Piper failed: {e}')
    finally:
        if text_file is not None:
            try:
                os.unlink(text_file.name)
            except OSError:
                pass
        if os.path.exists(wav_path):
            try:
                os.remove(wav_path)
            except OSError:
                pass
    return False


def _generate_silence(out_path, duration_ms=DEFAULT_SILENCE_MS):
    """Generate a silent MP3 when no text or TTS fails."""
    temp_wav = out_path.replace('.mp3', '.wav')
    try:
        ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
        _run_command([
            ffmpeg_path,
            '-f', 'lavfi', '-i', f'anullsrc=r=44100:cl=mono',
            '-t', str(duration_ms/1000), '-y', temp_wav
        ], timeout=10)
        _run_command([
            ffmpeg_path, '-i', temp_wav, '-acodec', 'libmp3lame', '-y', out_path
        ], timeout=10)
        if os.path.exists(temp_wav):
            os.remove(temp_wav)
    except Exception:
        with open(out_path, 'wb') as f:
            f.write(b'\xff\xfb\x90\x00' * 200)


def get_voices():
    """Return available voices."""
    return KOKORO_VOICES


def generate_preview(voice_id, preview_text=None):
    """Generate a 6-7 second voice preview."""
    if voice_id not in KOKORO_VOICES:
        return {'error': 'Invalid voice ID'}
    
    text = preview_text or PREVIEW_TEXT
    
    # Create preview directory
    preview_dir = os.path.join(OUTPUT_DIR, 'previews')
    os.makedirs(preview_dir, exist_ok=True)
    
    out_path = os.path.join(preview_dir, f'preview_{voice_id}.mp3')

    # Try preview engines in order. Kokoro supports voice-specific preview.
    engine = None
    warning = None

    success = _generate_kokoro(text, out_path, voice=voice_id)
    if success:
        engine = 'kokoro'
    else:
        success = _generate_piper(text, out_path)
        if success:
            engine = 'piper'
            warning = 'Kokoro is unavailable, preview generated with Piper and may not match the selected Kokoro voice.'
        else:
            success = _generate_gtts(text, out_path)
            if success:
                engine = 'gtts'
                warning = 'Kokoro is unavailable, preview generated with gTTS and may not match the selected Kokoro voice.'

    if success and os.path.exists(out_path):
        result = {
            'voice_id': voice_id,
            'voice_name': KOKORO_VOICES[voice_id]['name'],
            'engine': engine,
            'preview_url': f'/checkpoints/stage6_audio/previews/preview_{voice_id}.mp3',
        }
        if warning:
            result['warning'] = warning
        return result

    return {'error': 'Failed to generate preview. Install Kokoro (preferred) or gTTS/Piper fallback dependencies.'}


def generate_audio(filename, voice=None):
    """Stage 6: generate one MP3 per slide from speaker notes."""
    
    # Use specified voice or default
    selected_voice = voice if voice and voice in KOKORO_VOICES else DEFAULT_VOICE

    if is_cache_reuse_enabled() and checkpoint_mgr.exists('stage6_audio', filename):
        cached = checkpoint_mgr.load('stage6_audio', filename)
        # Check if cached voice matches selected voice
        if cached and 'error' not in cached and cached.get('voice_id') == selected_voice:
            if 'engine_summary' in cached:
                print(f'Valid Stage 6 checkpoint for {filename} (voice: {selected_voice})')
                return cached
            print(f'Stage 6 checkpoint for {filename} is missing engine metadata; rebuilding...')

    pptx_path = _get_input_pptx(filename)
    notes     = _extract_notes_from_pptx(pptx_path)
    out_dir   = os.path.join(OUTPUT_DIR, filename)
    os.makedirs(out_dir, exist_ok=True)

    audio_files = []
    total_words = 0

    print(f'  Using voice: {KOKORO_VOICES[selected_voice]["name"]}')
    
    for slide_idx, text in notes.items():
        out_path = os.path.join(out_dir, f'slide_{slide_idx:02d}.mp3')
        clean_text = text.strip()
        engine_used = 'silence'

        if clean_text:
            print(f'  Slide {slide_idx}: generating audio ({len(clean_text.split())} words)...')
            
            # Try Kokoro with selected voice
            success = _generate_kokoro(clean_text, out_path, voice=selected_voice)
            if success:
                engine_used = 'kokoro'
            
            # Try Piper if Kokoro fails
            if not success:
                print(f'  Slide {slide_idx}: Kokoro failed, trying Piper...')
                success = _generate_piper(clean_text, out_path)
                if success:
                    engine_used = 'piper'
            
            # Try gTTS as last resort
            if not success:
                print(f'  Slide {slide_idx}: Piper failed, trying gTTS...')
                success = _generate_gtts(clean_text, out_path)
                if success:
                    engine_used = 'gtts'
            
            if not success:
                print(f'  Slide {slide_idx}: All TTS failed, generating silence...')
                _generate_silence(out_path)
            else:
                total_words += len(clean_text.split())
        else:
            print(f'  Slide {slide_idx}: no notes -> generating silence...')
            _generate_silence(out_path)

        audio_files.append({
            'slide': slide_idx,
            'path': out_path,
            'words': len(clean_text.split()) if clean_text else 0,
            'has_speech': engine_used != 'silence',
            'engine': engine_used,
        })

    engine_summary = {}
    for item in audio_files:
        engine = item.get('engine', 'unknown')
        engine_summary[engine] = engine_summary.get(engine, 0) + 1

    non_silence = {k: v for k, v in engine_summary.items() if k != 'silence'}
    if non_silence:
        primary_engine = max(non_silence, key=non_silence.get)
    else:
        primary_engine = 'silence'

    result = {
        'filename': filename,
        'voice_id': selected_voice,
        'voice': KOKORO_VOICES[selected_voice]['name'],
        'slide_count': len(audio_files),
        'total_words': total_words,
        'output_dir': out_dir,
        'audio_files': audio_files,
        'engine_summary': engine_summary,
        'primary_engine': primary_engine,
        'fallback_used': any(engine in engine_summary for engine in ('piper', 'gtts', 'silence')),
    }
    checkpoint_mgr.save('stage6_audio', filename, result)
    return result
