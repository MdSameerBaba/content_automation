import os, json, re, io, traceback
from pathlib import Path
from contextlib import redirect_stdout, redirect_stderr

from pipeline.stage1_parser import parse_pdf
from pipeline.stage2_structurer import structure_slides
from pipeline.stage3_content import generate_content
from pipeline.stage4_builder import build_pptx

cwd = Path.cwd()
candidates = [cwd / 'uploads' / 'notheme.pdf', cwd.parent / 'uploads' / 'notheme.pdf']
pdf_path = None
for c in candidates:
    if c.exists():
        pdf_path = c
        break
if not pdf_path:
    raise FileNotFoundError('notheme.pdf not found in uploads paths')

def run_stage(name, func, *args, **kwargs):
    out = io.StringIO()
    err = io.StringIO()
    result = None
    exc = None
    with redirect_stdout(out), redirect_stderr(err):
        try:
            result = func(*args, **kwargs)
        except Exception:
            exc = traceback.format_exc()
    return {
        'name': name,
        'result': result,
        'stdout': out.getvalue(),
        'stderr': err.getvalue(),
        'exception': exc
    }

s1r = run_stage('stage1', parse_pdf, str(pdf_path))
s2r = run_stage('stage2', structure_slides, 'notheme', split_options={'mode':'custom','min_slides':10,'max_slides':15})
s3r = run_stage('stage3', generate_content, 'notheme')
s4r = run_stage('stage4', build_pptx, 'notheme')

s1 = s1r['result'] or {}
s2 = s2r['result'] or {}
s3 = s3r['result'] or {}
s4 = s4r['result'] or {}

slides = s1.get('slides') or []
slide6 = slides[5] if len(slides) >= 6 else {}
stage1_ok = (
    s1.get('page_count') == 9
    and slide6.get('title_from_image') is True
    and slide6.get('content_type_hint') == 'image_only'
)

groups = s2.get('groups') or []
groups_exist = isinstance(groups, list) and len(groups) > 0
groups_have_fields = groups_exist and all(
    isinstance(g, dict) and ('insert_divider_before' in g) and ('slide_type_summary' in g)
    for g in groups
)
vision_required = s2.get('vision_title_required_slides') or []
stage2_ok = groups_exist and groups_have_fields and (6 in vision_required)

typed_blueprint = s3.get('typed_blueprint') or []
raw_tbv = s3.get('typed_blueprint_version', 0)
try:
    tbv_num = int(raw_tbv)
except Exception:
    m = re.search(r'\d+', str(raw_tbv))
    tbv_num = int(m.group()) if m else 0

stage3_ok_basic = (
    tbv_num >= 6
    and s3.get('final_slide_count') == 11
    and ('slide_bullets' not in s3)
)

bp_by_source = {}
for item in typed_blueprint:
    if isinstance(item, dict) and item.get('source_slide_num') in (4,6,9):
        bp_by_source[item.get('source_slide_num')] = item

def diag_ok(item):
    if not isinstance(item, dict):
        return False
    return (
        item.get('type') == 'diagram'
        and (item.get('bullets') == [] or item.get('bullets') is None)
        and item.get('embed_source_image') is True
    )

stage3_quality = all(diag_ok(bp_by_source.get(n)) for n in (4,6,9))

slide6_bp = bp_by_source.get(6) or {}
stage3_title6 = bool(slide6_bp) and (slide6_bp.get('title') != 'Slide 6')

manifest = s4.get('slide_manifest') or []
slide6_manifest = None
for m in manifest:
    if isinstance(m, dict) and m.get('source_slide_num') == 6:
        slide6_manifest = m
        break
stage4_ok = (
    s4.get('total_slides') == 11
    and isinstance(slide6_manifest, dict)
    and slide6_manifest.get('title') != 'Slide 6'
)

combined_logs = '\n'.join([
    s1r.get('stdout',''), s1r.get('stderr',''),
    s2r.get('stdout',''), s2r.get('stderr',''),
    s3r.get('stdout',''), s3r.get('stderr',''),
    s4r.get('stdout',''), s4r.get('stderr',''),
    s1r.get('exception') or '', s2r.get('exception') or '', s3r.get('exception') or '', s4r.get('exception') or ''
])
provider_error_lines = []
for line in combined_logs.splitlines():
    if re.search(r'\b(429|400|401)\b', line):
        provider_error_lines.append(line.strip())

check_map = {
    'stage1_ok': stage1_ok,
    'stage2_ok': stage2_ok,
    'stage3_ok_basic': stage3_ok_basic,
    'stage3_quality': stage3_quality,
    'stage3_title6': stage3_title6,
    'stage4_ok': stage4_ok,
}
failed_checks = [k for k,v in check_map.items() if not v]

report = {
    'run_config': {
        'cwd': str(cwd),
        'python': os.sys.executable,
        'pdf_path': str(pdf_path),
        'env': {
            'PIPELINE_USE_CACHE': os.getenv('PIPELINE_USE_CACHE'),
            'STAGE3_PROVIDER_ORDER': os.getenv('STAGE3_PROVIDER_ORDER'),
            'STAGE3_ENABLE_VISION_ENRICHMENT': os.getenv('STAGE3_ENABLE_VISION_ENRICHMENT'),
            'STAGE3_VISION_IMAGE_ONLY': os.getenv('STAGE3_VISION_IMAGE_ONLY'),
            'STAGE3_RATE_LIMIT_SAFE_MODE': os.getenv('STAGE3_RATE_LIMIT_SAFE_MODE')
        }
    },
    'checks': check_map,
    'stage3_runtime_flags': s3.get('stage3_runtime'),
    'provider_errors_detected': {
        'any': len(provider_error_lines) > 0,
        'lines': provider_error_lines[:30]
    },
    'failed_checks': failed_checks,
    'stage_exceptions': {
        'stage1': s1r.get('exception'),
        'stage2': s2r.get('exception'),
        'stage3': s3r.get('exception'),
        'stage4': s4r.get('exception')
    },
    'evidence': {
        'stage1': {
            'page_count': s1.get('page_count'),
            'slide6': {
                'title': slide6.get('title'),
                'title_from_image': slide6.get('title_from_image'),
                'content_type_hint': slide6.get('content_type_hint')
            }
        },
        'stage2': {
            'group_count': len(groups) if isinstance(groups, list) else None,
            'group_fields_present_all': groups_have_fields,
            'vision_title_required_slides': vision_required
        },
        'stage3': {
            'typed_blueprint_version': s3.get('typed_blueprint_version'),
            'final_slide_count': s3.get('final_slide_count'),
            'has_top_level_slide_bullets': ('slide_bullets' in s3),
            'source_4': bp_by_source.get(4),
            'source_6': bp_by_source.get(6),
            'source_9': bp_by_source.get(9)
        },
        'stage4': {
            'total_slides': s4.get('total_slides'),
            'source6_manifest_entry': slide6_manifest
        }
    }
}

print(json.dumps(report, ensure_ascii=False, indent=2))
