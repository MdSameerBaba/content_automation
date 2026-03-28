"""Stage 6: Generate narration audio (MP3) from PPTX speaker notes using Kokoro with voice selection."""

import os
import subprocess
import imageio_ffmpeg
from pptx import Presentation
from pipeline.checkpoint import CheckpointManager

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

DEFAULT_VOICE = 'af_heart'

# Sample text for voice preview
PREVIEW_TEXT = "Hello! This is a preview of my voice. You can hear how I sound and choose the one you like best."

checkpoint_mgr = CheckpointManager()
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage6_audio')


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
            wav_path = out_path.replace('.mp3', '.wav')
            sf.write(wav_path, full_audio, 24000)
            
            # Use ffmpeg to convert to mp3
            ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
            subprocess.run([
                ffmpeg_path, '-i', wav_path,
                '-acodec', 'libmp3lame', '-y', out_path
            ], capture_output=True, timeout=30)
            
            if os.path.exists(wav_path):
                os.remove(wav_path)
            return os.path.exists(out_path)
    except Exception as e:
        print(f'    Kokoro failed: {e}')
    return False


def _generate_piper(text, out_path):
    """Generate audio using Piper TTS as fallback."""
    try:
        import tempfile
        text_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False)
        text_file.write(text)
        text_file.close()
        
        wav_path = out_path.replace('.mp3', '.wav')
        
        result = subprocess.run([
            'piper', '--text_file', text_file.name,
            '--output_file', wav_path
        ], capture_output=True, timeout=30)
        
        if result.returncode != 0:
            os.unlink(text_file.name)
            return False
        
        subprocess.run([
            'ffmpeg', '-i', wav_path,
            '-acodec', 'libmp3lame', '-y', out_path
        ], capture_output=True, timeout=30)
        
        os.unlink(text_file.name)
        if os.path.exists(wav_path):
            os.remove(wav_path)
        return True
    except Exception as e:
        print(f'    Piper failed: {e}')
    return False


def _generate_silence(out_path, duration_ms=1500):
    """Generate a silent MP3 when no text or TTS fails."""
    temp_wav = out_path.replace('.mp3', '.wav')
    try:
        subprocess.run([
            'ffmpeg', '-f', 'lavfi', '-i', f'anullsrc=r=44100:cl=mono',
            '-t', str(duration_ms/1000), '-y', temp_wav
        ], capture_output=True, timeout=10)
        subprocess.run([
            'ffmpeg', '-i', temp_wav, '-acodec', 'libmp3lame', '-y', out_path
        ], capture_output=True, timeout=10)
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
    
    # Generate with specified voice
    success = _generate_kokoro(text, out_path, voice=voice_id)
    
    if success and os.path.exists(out_path):
        return {
            'voice_id': voice_id,
            'voice_name': KOKORO_VOICES[voice_id]['name'],
            'preview_url': f'/checkpoints/stage6_audio/previews/preview_{voice_id}.mp3',
        }
    else:
        return {'error': 'Failed to generate preview'}


def generate_audio(filename, voice=None):
    """Stage 6: generate one MP3 per slide from speaker notes."""
    
    # Use specified voice or default
    selected_voice = voice if voice and voice in KOKORO_VOICES else DEFAULT_VOICE

    if checkpoint_mgr.exists('stage6_audio', filename):
        cached = checkpoint_mgr.load('stage6_audio', filename)
        # Check if cached voice matches selected voice
        if cached and 'error' not in cached and cached.get('voice_id') == selected_voice:
            print(f'Valid Stage 6 checkpoint for {filename} (voice: {selected_voice})')
            return cached

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

        if clean_text:
            print(f'  Slide {slide_idx}: generating audio ({len(clean_text.split())} words)...')
            
            # Try Kokoro with selected voice
            success = _generate_kokoro(clean_text, out_path, voice=selected_voice)
            
            # Try Piper if Kokoro fails
            if not success:
                print(f'  Slide {slide_idx}: Kokoro failed, trying Piper...')
                success = _generate_piper(clean_text, out_path)
            
            # Try gTTS as last resort
            if not success:
                print(f'  Slide {slide_idx}: Piper failed, trying gTTS...')
                success = _generate_gtts(clean_text, out_path)
            
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
            'has_speech': bool(clean_text),
        })

    print(f'  Audio: {len(audio_files)} files, ~{total_words} words -> {out_dir}')

    result = {
        'filename': filename,
        'voice_id': selected_voice,
        'voice': KOKORO_VOICES[selected_voice]['name'],
        'slide_count': len(audio_files),
        'total_words': total_words,
        'output_dir': out_dir,
        'audio_files': audio_files,
    }
    checkpoint_mgr.save('stage6_audio', filename, result)
    return result
