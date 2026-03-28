
import re
import json
import os
import sys
import requests
from dotenv import load_dotenv
from pipeline.checkpoint import CheckpointManager

load_dotenv()

checkpoint_mgr = CheckpointManager()


def parse_llm_json(raw_text):
    raw_text = raw_text.strip()
    # Method 1: extract from ```json ... ``` code blocks
    match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', raw_text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass
    # Method 2: direct parse
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        pass
    # Method 3: find outermost { }
    start = raw_text.find('{')
    end = raw_text.rfind('}')
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(raw_text[start:end+1])
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not parse JSON: {raw_text[:300]}")


def call_ai(prompt):
    groq_api_key = os.getenv('GROQ_API_KEY')
    groq_model = os.getenv('GROQ_MODEL', 'llama-3.3-70b-versatile')
    gemini_api_key = os.getenv('GEMINI_API_KEY')
    gemini_model = os.getenv('GEMINI_MODEL', 'gemini-2.0-flash')
    openrouter_api_key = os.getenv('OPENROUTER_API_KEY')
    openrouter_model = os.getenv('MODEL', 'anthropic/claude-sonnet-4-5')

    system_msg = (
        'You are a JSON-only responder. '
        'Respond with ONLY a valid JSON object. '
        'No explanation, no markdown, no code blocks. '
        'Start your response with { and end with }.'
    )

    if groq_api_key:
        url = 'https://api.groq.com/openai/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {groq_api_key}',
            'Content-Type': 'application/json'
        }
        payload = {
            'model': groq_model,
            'messages': [
                {'role': 'system', 'content': system_msg},
                {'role': 'user', 'content': prompt}
            ]
        }
        try:
            response = requests.post(url, headers=headers, json=payload)
            response.raise_for_status()
            return response.json()['choices'][0]['message']['content']
        except requests.exceptions.HTTPError as e:
            print(f'Groq error {e.response.status_code}: {e.response.text}', file=sys.stderr)
            raise

    elif gemini_api_key:
        url = f'https://generativelanguage.googleapis.com/v1beta/models/{gemini_model}:generateContent?key={gemini_api_key}'
        payload = {'contents': [{'parts': [{'text': prompt}]}]}
        response = requests.post(url, json=payload)
        response.raise_for_status()
        return response.json()['candidates'][0]['content']['parts'][0]['text']

    elif openrouter_api_key:
        url = 'https://openrouter.ai/api/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {openrouter_api_key}',
            'HTTP-Referer': 'https://openrouter.ai/',
            'X-Title': 'ppt-pipeline'
        }
        payload = {
            'model': openrouter_model,
            'messages': [
                {'role': 'system', 'content': system_msg},
                {'role': 'user', 'content': prompt}
            ]
        }
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']

    else:
        raise Exception('No AI provider configured. Set GROQ_API_KEY, GEMINI_API_KEY, or OPENROUTER_API_KEY in .env')


def structure_slides(filename):
    # Load checkpoint but skip if it's a bad/error checkpoint
    if checkpoint_mgr.exists('stage2_structured', filename):
        cached = checkpoint_mgr.load('stage2_structured', filename)
        if cached is not None and 'error' not in cached:
            groups = cached.get('groups', [])
            has_empty = any(len(g.get('slide_nums', [])) == 0 for g in groups)
            unassigned = cached.get('unassigned_slides', [])
            if not has_empty and not unassigned and len(groups) > 0:
                print(f'Valid checkpoint found for {filename}, skipping stage 2.')
                return cached
        print(f'Invalid checkpoint for {filename}, re-running stage 2...')

    # Load stage 1 output
    parsed = checkpoint_mgr.load('stage1_parsed', filename)
    if not parsed:
        raise Exception('Stage 1 checkpoint not found. Run Stage 1 first.')
    slides = parsed['slides']
    total = len(slides)

    prompt = f"""
You are an expert presentation designer. Analyze the following slides and group them into logical sections.

STRICT RULES:
1. Every slide MUST be assigned to exactly one group. No slide can be skipped.
2. Every group MUST have at least one slide. No empty groups allowed.
3. The toc list must exactly match the section_title values in groups.
4. slide_nums must contain integers only. All slides from 1 to {total} must appear exactly once across all groups.

Return ONLY a JSON object with this exact structure:
{{
  "toc": ["Section Title 1", "Section Title 2"],
  "groups": [
    {{"section_title": "Section Title 1", "slide_nums": [1, 2, 3]}},
    {{"section_title": "Section Title 2", "slide_nums": [4, 5, 6]}}
  ]
}}

Slides to group:
{json.dumps([{'slide_num': s['slide_num'], 'title': s['title'], 'raw_text': s['raw_text'][:300]} for s in slides], ensure_ascii=False, indent=2)}

IMPORTANT: All {total} slides must be assigned. Start response with {{ and end with }}
"""

    raw = call_ai(prompt)

    try:
        struct = parse_llm_json(raw)
    except Exception as ex:
        struct = {'error': 'Could not parse AI response', 'raw': raw, 'exception': str(ex)}
        checkpoint_mgr.save('stage2_structured', filename, struct)
        return struct

    # Validate all slides assigned
    all_assigned = []
    for group in struct.get('groups', []):
        all_assigned.extend(group.get('slide_nums', []))

    missing = [i for i in range(1, total + 1) if i not in all_assigned]
    if missing:
        print(f'WARNING: Slides {missing} were not assigned to any group!')
        struct['unassigned_slides'] = missing

    # Remove any empty groups
    struct['groups'] = [g for g in struct.get('groups', []) if g.get('slide_nums')]
    struct['toc'] = [g['section_title'] for g in struct['groups']]

    checkpoint_mgr.save('stage2_structured', filename, struct)
    return struct