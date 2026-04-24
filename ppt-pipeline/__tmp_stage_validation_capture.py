import argparse
import json
import os
import re
import io
import contextlib
from pipeline.stage1_parser import parse_pdf
from pipeline.stage2_structurer import structure_slides
from pipeline.stage3_content import generate_content
from pipeline.stage4_builder import build_pptx

CODES = ("400", "401", "429")
PROVIDER_HINTS = ("provider", "openai", "anthropic", "gemini", "groq", "api")


def detect_provider_http_from_text(text):
    matches = []
    lines = text.splitlines()
    for i, line in enumerate(lines, 1):
        l = line.lower()
        if any(c in line for c in CODES) and any(h in l for h in PROVIDER_HINTS):
            matches.append({"line": i, "text": line[:300]})
    return matches


def find_http_errors(obj, path="root", out=None):
    if out is None:
        out = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}"
            lk = str(k).lower()
            if isinstance(v, int) and v in {400, 401, 429} and (lk in {"status", "status_code", "http_status", "code", "error_code"} or "provider" in p.lower() or "error" in p.lower()):
                out.append({"path": p, "value": v})
            elif isinstance(v, str):
                if re.search(r"\\b(400|401|429)\\b", v) and any(h in (p + " " + v).lower() for h in PROVIDER_HINTS):
                    out.append({"path": p, "value": v[:240]})
            find_http_errors(v, p, out)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            find_http_errors(v, f"{path}[{i}]", out)
    return out

parser = argparse.ArgumentParser()
parser.add_argument("--pdf", required=True)
args = parser.parse_args()

s1 = parse_pdf(args.pdf)
s2 = structure_slides('notheme', split_options={'mode': 'custom', 'min_slides': 10, 'max_slides': 15})

buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    s3 = generate_content('notheme')
    s4 = build_pptx('notheme')
runtime_stdout = buf.getvalue()

bp = s3.get('typed_blueprint') or []
slide6_bp = next((x for x in bp if x.get('source_slide_num') == 6), None)
manifest = s4.get('slide_manifest') or []
slide6_manifest = next((x for x in manifest if x.get('source_slide_num') == 6), None)

obj_matches = find_http_errors({"stage3": s3, "stage4": s4})
stdout_matches = detect_provider_http_from_text(runtime_stdout)
all_matches = obj_matches + [{"path": f"stdout.line_{m['line']}", "value": m['text']} for m in stdout_matches]

summary = {
    "pdf_path": args.pdf,
    "cache_disabled": os.environ.get("PIPELINE_USE_CACHE") == "0",
    "stage1": {"page_count": s1.get('page_count')},
    "stage2": {"toc": s2.get('toc', [])},
    "stage3": {
        "typed_blueprint_version": s3.get('typed_blueprint_version'),
        "stage3_runtime": s3.get('stage3_runtime'),
        "title_inference": s3.get('title_inference'),
        "source_slide_num_6_entry": slide6_bp
    },
    "stage4": {
        "total_slides": s4.get('total_slides'),
        "source_slide_num_6_manifest_entry": slide6_manifest
    },
    "provider_http_errors_400_401_429": {
        "occurred": len(all_matches) > 0,
        "matches": all_matches[:20]
    }
}

print(json.dumps(summary, ensure_ascii=False, indent=2))
