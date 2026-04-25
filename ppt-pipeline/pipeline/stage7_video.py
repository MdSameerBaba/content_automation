"""Stage 7: Combine slide PNGs + MP3 audio into a final MP4 video using ffmpeg directly."""

import os
import glob
import subprocess
import imageio_ffmpeg
from pipeline.checkpoint import CheckpointManager, is_cache_reuse_enabled

# Get ffmpeg path
FFMPEG_PATH = imageio_ffmpeg.get_ffmpeg_exe()

checkpoint_mgr = CheckpointManager()
OUTPUT_DIR = os.path.join(checkpoint_mgr.base_dir, 'stage7_video')


def _run_ffmpeg(cmd, context):
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f'{context} failed (exit {result.returncode}): {result.stderr[:500]}')
    return result


def _checkpoint_path(stage, filename):
    return os.path.join(checkpoint_mgr.base_dir, stage, f'{filename}.json')


def _is_cached_video_valid(filename, cached):
    if not cached or 'error' in cached:
        return False

    out_path = cached.get('output_path')
    if not out_path or not os.path.exists(out_path):
        return False

    stage7_cp = _checkpoint_path('stage7_video', filename)
    stage5_cp = _checkpoint_path('stage5_images', filename)
    stage6_cp = _checkpoint_path('stage6_audio', filename)
    if not os.path.exists(stage7_cp):
        return False

    # If stage5/stage6 changed after stage7, rebuild video.
    stage7_mtime = max(os.path.getmtime(stage7_cp), os.path.getmtime(out_path))
    for dep in (stage5_cp, stage6_cp):
        if os.path.exists(dep) and os.path.getmtime(dep) > stage7_mtime:
            return False

    return True


def _get_audio_duration(audio_path):
    """Get duration of an MP3 file in seconds."""
    try:
        try:
            # MoviePy v2.x
            from moviepy import AudioFileClip
        except ImportError:
            # MoviePy v1.x fallback
            from moviepy.audio.io.AudioFileClip import AudioFileClip
        clip = AudioFileClip(audio_path)
        duration = clip.duration
        clip.close()
        return duration
    except Exception:
        return 3.0


def create_video(filename, force=False):
    """Stage 7: combine PNGs + MP3s into final MP4 using ffmpeg directly."""

    cached = checkpoint_mgr.load('stage7_video', filename) if checkpoint_mgr.exists('stage7_video', filename) else None
    cache_allowed = is_cache_reuse_enabled()
    if cache_allowed and not force and _is_cached_video_valid(filename, cached):
        print(f'Valid Stage 7 checkpoint for {filename}')
        return cached

    if force:
        print(f'Forcing Stage 7 rebuild for {filename}...')
    elif cached:
        print(f'Stage 7 cache is stale for {filename}; rebuilding...')

    # Load stage 5 and 6 checkpoints
    stage5 = checkpoint_mgr.load('stage5_images', filename)
    stage6 = checkpoint_mgr.load('stage6_audio', filename)

    if not stage5:
        raise Exception('Stage 5 checkpoint missing. Run Stage 5 first.')
    if not stage6:
        raise Exception('Stage 6 checkpoint missing. Run Stage 6 first.')

    images_dir = stage5['output_dir']
    audio_dir = stage6['output_dir']
    slide_count = stage5['slide_count']

    # Collect slide PNGs in order
    images = sorted(glob.glob(os.path.join(images_dir, 'slide_*.png')))
    if not images:
        raise FileNotFoundError(f'No PNG images found in {images_dir}')

    print(f'Stage 7: building video from {len(images)} slides...')

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f'{filename}.mp4')

    # Process each slide: create video clip with audio, then concatenate
    slide_clips = []
    
    for idx, img_path in enumerate(images, 1):
        slide_num = idx
        audio_path = os.path.join(audio_dir, f'slide_{slide_num:02d}.mp3')
        
        if os.path.exists(audio_path):
            duration = _get_audio_duration(audio_path)
            # Create slide video with audio using filter_complex
            slide_clip_path = os.path.join(OUTPUT_DIR, f'slide_{slide_num:02d}.mp4')
            
            cmd = [
                FFMPEG_PATH, '-y',
                '-loop', '1', '-i', img_path,
                '-i', audio_path,
                '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
                '-c:a', 'aac', '-b:a', '128k',
                '-shortest',
                '-movflags', '+faststart',
                '-t', str(duration),
                slide_clip_path
            ]
            
            try:
                _run_ffmpeg(cmd, f'slide {slide_num} video+audio render')
            except Exception as render_err:
                print(f'  Warning: {render_err}. Using fallback without audio.')
                # Fallback: video without audio
                cmd = [
                    FFMPEG_PATH, '-y',
                    '-loop', '1', '-i', img_path,
                    '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
                    '-t', '3', '-shortest',
                    slide_clip_path
                ]
                _run_ffmpeg(cmd, f'slide {slide_num} fallback render')
            
            slide_clips.append(slide_clip_path)
        else:
            # No audio, create silent video
            slide_clip_path = os.path.join(OUTPUT_DIR, f'slide_{slide_num:02d}.mp4')
            cmd = [
                FFMPEG_PATH, '-y',
                '-loop', '1', '-i', img_path,
                '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
                '-t', '3', '-shortest',
                slide_clip_path
            ]
            _run_ffmpeg(cmd, f'slide {slide_num} silent render')
            slide_clips.append(slide_clip_path)
    
    # Concatenate all slide videos.
    # Re-encode (not stream-copy) so mixed audio/silent clips are compatible.
    concat_list_path = os.path.join(OUTPUT_DIR, 'concat_list.txt')
    with open(concat_list_path, 'w') as f:
        for clip_path in slide_clips:
            f.write(f"file '{clip_path}'\n")

    cmd = [
        FFMPEG_PATH, '-y',
        '-f', 'concat',
        '-safe', '0',
        '-i', concat_list_path,
        '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
        '-c:a', 'aac', '-b:a', '128k',
        '-movflags', '+faststart',
        out_path
    ]

    print(f'  Concatenating {len(slide_clips)} slides...')
    try:
        _run_ffmpeg(cmd, 'final video concat')
    except Exception as concat_err:
        print(f'FFmpeg concat error: {concat_err}')
        # Clean up temp clips before re-raising so we don't leave disk clutter.
        for _cp in slide_clips:
            if os.path.exists(_cp):
                try:
                    os.remove(_cp)
                except OSError:
                    pass
        if os.path.exists(concat_list_path):
            try:
                os.remove(concat_list_path)
            except OSError:
                pass
        raise RuntimeError(f'Final video concat failed: {concat_err}') from concat_err
    
    # Cleanup temp files
    try:
        for clip_path in slide_clips:
            if os.path.exists(clip_path):
                os.remove(clip_path)
        if os.path.exists(concat_list_path):
            os.remove(concat_list_path)
    except OSError:
        # Non-fatal cleanup issue.
        pass

    # Calculate total duration
    total_duration = sum(
        _get_audio_duration(os.path.join(audio_dir, f'slide_{i:02d}.mp3'))
        for i in range(1, len(images) + 1)
        if os.path.exists(os.path.join(audio_dir, f'slide_{i:02d}.mp3'))
    )

    print(f'Stage 7 complete: {out_path} ({total_duration:.1f}s total)')

    result = {
        'filename': filename,
        'output_path': out_path,
        'total_slides': len(images),
        'total_duration_seconds': round(total_duration, 1),
        'total_duration_minutes': round(total_duration / 60, 2),
        'fps': 24,
        'codec': 'libx264',
    }
    checkpoint_mgr.save('stage7_video', filename, result)
    return result
