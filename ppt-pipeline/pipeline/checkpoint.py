import os
import json

# Always resolve relative to ppt-pipeline/ (parent of this file's directory)
# so checkpoints land in the same place regardless of launch CWD
_PIPELINE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

class CheckpointManager:
    def __init__(self, base_dir='checkpoints'):
        if os.path.isabs(base_dir):
            self.base_dir = base_dir
        else:
            self.base_dir = os.path.join(_PIPELINE_ROOT, base_dir)

    def _get_path(self, stage, filename):
        return os.path.join(self.base_dir, stage, f"{filename}.json")

    def save(self, stage, filename, data):
        path = self._get_path(stage, filename)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load(self, stage, filename):
        path = self._get_path(stage, filename)
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None

    def exists(self, stage, filename):
        return os.path.exists(self._get_path(stage, filename))
