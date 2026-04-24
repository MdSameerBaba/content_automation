import json
import os
from functools import lru_cache


CONFIG_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config', 'pipeline_config.json'))


DEFAULT_CONFIG = {
    'ai': {
        'provider_order': ['groq', 'gemini', 'openrouter'],
        'models': {
            'groq': 'llama-3.3-70b-versatile',
            'gemini': 'gemini-2.0-flash',
            'openrouter': 'anthropic/claude-sonnet-4-5',
        },
        'http': {
            'timeout_seconds': 45,
            'max_retries': 2,
            'retry_backoff_seconds': 1.5,
        },
    },
    'tts': {
        'default_voice': 'af_heart',
        'preview_text': 'Hello! This is a preview of my voice. You can hear how I sound and choose the one you like best.',
        'silence_duration_ms': 1500,
    },
}


def _deep_merge(base, override):
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


@lru_cache(maxsize=1)
def get_config():
    config = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            loaded = json.load(f)
            config = _deep_merge(config, loaded)
    return config


def get_ai_config():
    config = get_config().get('ai', {})

    provider_order = config.get('provider_order', ['groq', 'gemini', 'openrouter'])
    env_order = os.getenv('AI_PROVIDER_ORDER')
    if env_order:
        provider_order = [p.strip().lower() for p in env_order.split(',') if p.strip()]

    models = config.get('models', {})
    models = {
        'groq': os.getenv('GROQ_MODEL', models.get('groq', DEFAULT_CONFIG['ai']['models']['groq'])),
        'gemini': os.getenv('GEMINI_MODEL', models.get('gemini', DEFAULT_CONFIG['ai']['models']['gemini'])),
        'openrouter': os.getenv('MODEL', models.get('openrouter', DEFAULT_CONFIG['ai']['models']['openrouter'])),
    }

    http = config.get('http', {})
    timeout_seconds = int(os.getenv('AI_HTTP_TIMEOUT_SECONDS', http.get('timeout_seconds', 45)))
    max_retries = int(os.getenv('AI_HTTP_MAX_RETRIES', http.get('max_retries', 2)))
    retry_backoff_seconds = float(os.getenv('AI_HTTP_RETRY_BACKOFF_SECONDS', http.get('retry_backoff_seconds', 1.5)))

    return {
        'provider_order': provider_order,
        'models': models,
        'http': {
            'timeout_seconds': timeout_seconds,
            'max_retries': max_retries,
            'retry_backoff_seconds': retry_backoff_seconds,
        },
    }


def get_tts_config():
    config = get_config().get('tts', {})
    return {
        'default_voice': os.getenv('DEFAULT_VOICE', config.get('default_voice', DEFAULT_CONFIG['tts']['default_voice'])),
        'preview_text': config.get('preview_text', DEFAULT_CONFIG['tts']['preview_text']),
        'silence_duration_ms': int(os.getenv('SILENCE_DURATION_MS', config.get('silence_duration_ms', 1500))),
    }
