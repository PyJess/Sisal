import json
import os

def load_file(filepath:str):
    with open(filepath, encoding="utf-8") as f:
        return f.read()


def load_json(filepath:str):
    with open(filepath, encoding="utf-8") as f:
        return json.load(f)


def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def get_user_path(user_id: str, subfolder: str = "") -> str:
    """Get the user-specific path."""
    base_dir = os.path.join(os.path.dirname(__file__),"outputs", user_id, subfolder)
    os.makedirs(base_dir, exist_ok=True)
    return base_dir