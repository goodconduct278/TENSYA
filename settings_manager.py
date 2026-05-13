import json
from pathlib import Path

SETTINGS_PATH = Path(__file__).parent / 'settings.json'

DEFAULT_SETTINGS = {
    "src_sheet":     "内訳(3)",
    "src_start_row": 6,
    "src_col_name":  "D",
    "src_col_spec":  "E",
    "src_col_qty":   "F",
    "src_col_unit":  "G",
    "src_col_price": "H",
    "dst_sheet":     "防水工事",
    "dst_start_row": 8,
    "dst_col_name":  "B",
    "dst_col_spec":  "C",
    "dst_col_qty":   "D",
    "dst_col_unit":  "E",
    "dst_col_price": "F",
    "dst_col_amount":"G",
    "dst_layout":    "3行標準",
    "sum_keyword":   "合計",
    "backup_mode":   "常に作成",
}

def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        with open(SETTINGS_PATH, encoding='utf-8') as f:
            data = json.load(f)
        for k, v in DEFAULT_SETTINGS.items():
            if k not in data:
                data[k] = v
        return data
    return DEFAULT_SETTINGS.copy()

def save_settings(settings: dict) -> None:
    with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
