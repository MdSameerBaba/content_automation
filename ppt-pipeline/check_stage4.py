from pipeline.checkpoint import CheckpointManager
cm = CheckpointManager()
r = cm.load('stage4_pptx', 'notheme')
print('Total slides:', r['total_slides'])
print('Narrator words:', r.get('narrator_words', 'N/A'))
print('Estimated minutes:', r.get('estimated_minutes', 'N/A'))
print()
for s in r['slide_manifest']:
    print(f"  Slide {s['idx']:2d} [{s['type']:12s}] {s['bullets']} bullets | {s['notes_words']} words notes | \"{s['title'][:45]}\"")
