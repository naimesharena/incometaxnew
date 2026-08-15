# Auto-generated full field mapping loader - loads from JSON to avoid escaping issues
import json, pathlib

def load_full_mapping():
    json_path = pathlib.Path(__file__).parent / "full_field_mapping.json"
    if json_path.exists():
        with open(json_path, encoding="utf-8") as f:
            return json.load(f)
    return {}

FULL_FIELD_MAPPING = load_full_mapping()
# Total entries
TOTAL_ENTRIES = len(FULL_FIELD_MAPPING)
