"""
i18n.py
--------
軽量なJSONベースの多言語対応ユーティリティ。

- locales/<lang>.json からキー/値を読み込み
- set_locale("auto"|"en"|"ja"|...) で現在ロケールを切替
- tr(key, **kwargs) で翻訳取得（en→キーの順にフォールバック）
"""
from __future__ import annotations

import json
import locale as pylocale
from pathlib import Path
from typing import Dict, Optional

_current: str = "en"
_fallback: str = "en"
_translations: Dict[str, Dict[str, str]] = {}


def _locales_dir() -> Path:
    # このファイルの隣にある locales ディレクトリ
    return Path(__file__).resolve().parent / "locales"


def _load_lang(lang: str) -> Dict[str, str]:
    if lang in _translations:
        return _translations[lang]
    try:
        fp = _locales_dir() / f"{lang}.json"
        if fp.exists():
            with open(fp, "r", encoding="utf-8") as f:
                _translations[lang] = json.load(f)
        else:
            _translations[lang] = {}
    except Exception:
        _translations[lang] = {}
    return _translations[lang]


def available_locales() -> Dict[str, str]:
    """利用可能なロケール一覧を返す（コード→表示名）。"""
    names = {
        "auto": "Auto",
    }
    locales_path = _locales_dir()
    for f in locales_path.iterdir():
        if f.suffix == ".json":
            lang_code = f.stem
            try:
                with open(f, "r", encoding="utf-8") as json_file:
                    data = json.load(json_file)
                    # lang_name フィールドがあればそれを使用、なければデフォルトで設定
                    lang_name = data.get("lang_name", lang_code)
                    names[lang_code] = lang_name
            except Exception:
                names[lang_code] = lang_code  # エラー時はファイル名をそのまま使用
    return dict(sorted(names.items())) # ソートして返す


def detect_system_lang() -> str:
    try:
        lang, _ = pylocale.getdefaultlocale()  # type: ignore[arg-type]
        if not lang:
            return _fallback
        lang = lang.lower()
        if lang.startswith("ja"):
            return "ja"
        return "en"
    except Exception:
        return _fallback


def set_locale(lang: Optional[str]) -> str:
    global _current
    if not lang or lang == "auto":
        lang = detect_system_lang()
    # 事前にロード
    _load_lang(lang)
    _load_lang(_fallback)
    _current = lang
    return _current


def current_locale() -> str:
    return _current


def tr(key: str, **kwargs) -> str:
    """キーに対応する翻訳を返す。未定義は en→キーの順にフォールバック。"""
    val = _load_lang(_current).get(key)
    if val is None:
        val = _load_lang(_fallback).get(key, key)
    try:
        return val.format(**kwargs)
    except Exception:
        return val

