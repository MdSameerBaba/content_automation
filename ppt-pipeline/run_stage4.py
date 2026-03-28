import os
os.system('del /f checkpoints\\stage4_pptx\\notheme.json 2>nul')
from pipeline.stage4_builder import build_pptx
r = build_pptx('notheme')
print('DONE:', r['total_slides'], 'slides ->', r['output_path'])
