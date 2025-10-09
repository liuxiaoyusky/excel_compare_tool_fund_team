from seg_mapping_config import sec_mapping
from pathlib import Path
import json

bond_code_set = {
    'MR',
    'MF',
    'G',
    'CB',
    'B',
}

stack_code_set = {
    'S',
    'P',
}

_MAPPING_OVERRIDE_PATH = Path(__file__).parent / 'mapping_override.json'


def _normalize_mapping(raw: dict) -> dict:
    """Uppercase keys/values; ensure str; None -> '' for values."""
    norm: dict[str, str] = {}
    if not isinstance(raw, dict):
        return norm
    for k, v in raw.items():
        if k is None:
            continue
        key = str(k).strip().upper()
        if not key:
            continue
        val = '' if v is None else str(v).strip().upper()
        norm[key] = val
    return norm


def _load_mapping_override() -> dict:
    if not _MAPPING_OVERRIDE_PATH.exists():
        return {}
    try:
        data = json.loads(_MAPPING_OVERRIDE_PATH.read_text(encoding='utf-8'))
        return _normalize_mapping(data)
    except Exception:
        # 读取失败时忽略覆盖，使用默认映射
        return {}


def save_mapping_override(mapping: dict) -> None:
    """Persist override mapping to JSON and refresh in-memory merged mapping."""
    override = _normalize_mapping(mapping)
    _MAPPING_OVERRIDE_PATH.write_text(json.dumps(override, ensure_ascii=False, indent=2, sort_keys=True), encoding='utf-8')
    # 刷新内存中的合并映射（默认 sec_mapping + 覆盖）
    merged = dict(_normalize_mapping(sec_mapping))
    merged.update(override)
    missing_isin_or_stack_code_mapping_dict.clear()
    missing_isin_or_stack_code_mapping_dict.update(merged)


# 初始加载：默认 sec_mapping 与本地覆盖合并（覆盖优先）
_base = dict(_normalize_mapping(sec_mapping))
_override = _load_mapping_override()
_base.update(_override)
missing_isin_or_stack_code_mapping_dict = _base

# 绝对容差：当 |hsbc - spectra| <= 此值时视为相等
TOLERANCE_ABS = 0.000

# 历史归档开关与目录
# 当 ENABLE_HISTORY 为 True 时，完成一次对比后将输入与输出快照保存到 HISTORY_DIR/<timestamp>/
ENABLE_HISTORY = True
HISTORY_DIR = Path(__file__).parent / 'history'