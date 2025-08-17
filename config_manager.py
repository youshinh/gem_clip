# config_manager.py
import json
# import shutil
from pathlib import Path
from typing import Dict, Optional
from pydantic import BaseModel
import tkinter.messagebox as messagebox

from common_models import Prompt, AppConfig
from constants import CONFIG_FILE, APP_NAME
# Import paths module. When this module is executed as a script (no package
# context), relative imports may fail. Fallback to loading the module
# directly from its file path.
try:
    from . import paths  # type: ignore
except ImportError:
    import importlib.util
    import os
    _paths_spec = importlib.util.spec_from_file_location(
        "paths", os.path.join(os.path.dirname(__file__), "paths.py")
    )
    assert _paths_spec and _paths_spec.loader
    paths = importlib.util.module_from_spec(_paths_spec)
    _paths_spec.loader.exec_module(paths)  # type: ignore

def _read_json(path: Path) -> Optional[dict]:
    """Read a JSON file and return the parsed data or None on error."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"ERROR: Failed to read JSON from {path}: {e}")
        return None

def _write_json(path: Path, data: dict) -> bool:
    """Write the given data as JSON to the specified path.

    Returns True on success, False on error.
    """
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"ERROR: Failed to write JSON to {path}: {e}")
        return False

def _migrate_v1_to_v2(data: dict) -> dict:
    """Migrate a v1 configuration dictionary to the v2 format.

    The v1 format lacks a version field and stores prompts, max_history_size, hotkey,
    and optionally api_key. This function adds a version field (2) and returns
    the updated dictionary. Additional structural changes can be made here in
    future migrations.
    """
    data = data.copy()
    data["version"] = 2
    # ensure required top-level keys exist
    data.setdefault("prompts", {})
    data.setdefault("max_history_size", 20)
    data.setdefault("hotkey", "ctrl+shift+c")
    return data

def _migrate_v2_to_v3(data: dict) -> dict:
    """Migrate a v2 configuration dictionary to the v3 format.

    v3 introduces granular hotkeys: hotkey_prompt_list, hotkey_refine, hotkey_matrix
    and deprecates the single 'hotkey' field.
    """
    data = data.copy()
    # carry over existing single hotkey to prompt_list if present
    legacy = data.get("hotkey")
    data.setdefault("hotkey_prompt_list", legacy or "ctrl+shift+c")
    data.setdefault("hotkey_refine", "ctrl+shift+r")
    data.setdefault("hotkey_matrix", None)
    # bump version
    data["version"] = 3
    # Remove legacy key to avoid confusion
    if "hotkey" in data:
        try:
            del data["hotkey"]
        except Exception:
            pass
    return data



def _migrate_v3_to_v4(data: dict) -> dict:
    """Migrate v3 to v4 by adding matrix summary prompt placeholders."""
    data = data.copy()
    data.setdefault("matrix_row_summary_prompt", None)
    data.setdefault("matrix_col_summary_prompt", None)
    data.setdefault("matrix_matrix_summary_prompt", None)
    data["version"] = 4
    return data

def _migrate_v4_to_v5(data: dict) -> dict:
    """Migrate v4 to v5 by adding flow settings."""
    data = data.copy()
    data.setdefault("max_flow_steps", 5)
    data["version"] = 5
    return data

def _migrate_v5_to_v6(data: dict) -> dict:
    """Migrate v5 to v6 by adding language setting."""
    data = data.copy()
    data.setdefault("language", "auto")
    data["version"] = 6
    return data

def _migrate_v6_to_v7(data: dict) -> dict:
    """Migrate v6 to v7 by adding theme mode setting."""
    data = data.copy()
    data.setdefault("theme_mode", "system")
    data["version"] = 7
    return data

def load_config() -> Optional[AppConfig]:
    """Load the application configuration with automatic migration support.

    This function attempts to read the configuration from the user-specific
    configuration directory. If the configuration file does not exist in the
    new location, but an old ``config.json`` exists in the current working
    directory, the old file is migrated to the new location and upgraded to
    the latest version.

    Returns:
        An instance of ``AppConfig`` if loading and validation succeed,
        otherwise ``None``. Errors are displayed to the user via a
        message box.
    """
    # Determine paths
    new_config_path: Path = paths.get_config_file_path()
    old_config_path: Path = Path(CONFIG_FILE)

    # If new config does not exist, attempt to migrate from legacy app-name locations
    # used before the project was renamed. This preserves user settings across rename.
    if not new_config_path.exists():
        try:
            import sys, os
            legacy_names = ["Geminiクリップボード", "Gemini Clipboard"]
            legacy_base: Optional[Path] = None
            if sys.platform.startswith("win"):
                appdata = os.environ.get("APPDATA")
                if appdata:
                    for nm in legacy_names:
                        cand = Path(appdata) / nm
                        if (cand / CONFIG_FILE).exists():
                            legacy_base = cand
                            break
            elif sys.platform == "darwin":
                for nm in legacy_names:
                    cand = Path.home() / "Library" / "Application Support" / nm
                    if (cand / CONFIG_FILE).exists():
                        legacy_base = cand
                        break
            else:
                xdg_config_home = os.environ.get("XDG_CONFIG_HOME")
                if xdg_config_home:
                    for nm in legacy_names:
                        cand = Path(xdg_config_home) / nm
                        if (cand / CONFIG_FILE).exists():
                            legacy_base = cand
                            break
                if legacy_base is None:
                    for nm in legacy_names:
                        cand = Path.home() / ".config" / nm
                        if (cand / CONFIG_FILE).exists():
                            legacy_base = cand
                            break
            if legacy_base is not None:
                data = _read_json(legacy_base / CONFIG_FILE)
                if data is not None:
                    # Write to new location with minimal changes (structural migration happens later)
                    if _write_json(new_config_path, data):
                        messagebox.showinfo(APP_NAME, f"旧アプリ名の設定を移行しました: {legacy_base / CONFIG_FILE} → {new_config_path}")
        except Exception:
            # Non-fatal: continue with other migration paths
            pass
    # If new config does not exist but an old config file exists in the working directory,
    # migrate it to the new location.
    if not new_config_path.exists() and old_config_path.exists():
        data = _read_json(old_config_path)
        if data is not None:
            # Detect version. If missing or < 2, migrate structure.
            if data.get("version", 1) < 2:
                data = _migrate_v1_to_v2(data)
            # Write migrated data to new location
            if _write_json(new_config_path, data):
                # Optionally keep a backup of the old file. We leave it in place for now.
                messagebox.showinfo(APP_NAME, f"旧設定ファイル {old_config_path} を {new_config_path} に移行しました。")
        else:
            # If reading fails, create a default config in new location
            create_default_config()
    # If new config still does not exist, create a default one
    if not new_config_path.exists():
        create_default_config()
        messagebox.showinfo(APP_NAME, f"{new_config_path.name} を作成しました。プロンプトやホットキーを編集してください。")
    # Read the configuration
    try:
        data = _read_json(new_config_path)
        if data is None:
            raise ValueError("設定ファイルが空、または読み込みに失敗しました。")
        # Handle migration if needed
        ver = data.get("version", 1)
        if ver < 2:
            data = _migrate_v1_to_v2(data)
            ver = 2
        if ver < 3:
            data = _migrate_v2_to_v3(data)
            ver = 3
        if ver < 4:
            data = _migrate_v3_to_v4(data)
            ver = 4
        if ver < 5:
            data = _migrate_v4_to_v5(data)
            ver = 5
        if ver < 6:
            data = _migrate_v5_to_v6(data)
            ver = 6
        if ver < 7:
            data = _migrate_v6_to_v7(data)
            ver = 7
        if data.get("version") != ver:
            data["version"] = ver
        _write_json(new_config_path, data)
        return AppConfig(**data)
    except Exception as e:
        messagebox.showerror("設定エラー", f"設定ファイルの読み込みに失敗しました: {e}")
        return None

def save_config(config: AppConfig):
    """Save the current configuration to the user-specific configuration file."""
    # Persist explicit None values (e.g., to keep hotkeys disabled on restart)
    data = config.model_dump(by_alias=True, exclude_none=False)
    # Ensure version is set to latest
    data["version"] = config.version if hasattr(config, "version") else 4
    config_path: Path = paths.get_config_file_path()
    try:
        _write_json(config_path, data)
    except Exception as e:
        print(f"ERROR: Failed to save config: {e}")
        # Keep this silent; message boxes should be handled at higher level

def create_default_config():
    """Create a default configuration file in the user-specific configuration directory."""
    default_config = {
        "version": 7,
        "prompts": {
            "check": {
                "name": "誤字脱字を修正",
                "model": "gemini-2.5-flash",
                "system_prompt": (
                    "あなたはプロの編集者です。以下のテキストに含まれる誤字、脱字、文法的な誤りを修正し、"
                    "自然で読みやすい文章にしてください。内容は変更せず、修正後のテキストのみを返してください。"
                ),
                "parameters": {"temperature": 0.2},
            },
            "summarize": {
                "name": "複雑な内容を箇条書きで要約",
                "model": "gemini-2.5-flash",
                "system_prompt": (
                    "以下の専門的なテキストの要点を抽出し、簡潔に最大5つまでの箇条書きで要約してください。"
                ),
                "parameters": {"temperature": 0.5},
            },
            "ocr": {
                "name": "画像からテキストを抽出",
                "model": "gemini-2.5-flash",
                "system_prompt": (
                    "この画像に含まれるテキストを正確に読み取り、そのまま出力してください。"
                    "テキストの構造やレイアウトも可能な限り保持してください。"
                ),
                "parameters": {"temperature": 0.1},
            },
        },
        "max_history_size": 20,
        "hotkey_prompt_list": "ctrl+shift+c",
        "hotkey_refine": "ctrl+shift+r",
        "hotkey_matrix": None,
        "matrix_row_summary_prompt": None,
        "matrix_col_summary_prompt": None,
        "matrix_matrix_summary_prompt": None,
        "max_flow_steps": 5,
        "language": "auto",
        "theme_mode": "system",
    }
    config_path: Path = paths.get_config_file_path()
    _write_json(config_path, default_config)
