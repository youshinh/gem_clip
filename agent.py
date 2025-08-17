# agent.py
import asyncio
import base64
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
from io import BytesIO
import hashlib
import json
from pathlib import Path
from typing import Dict, Literal, Optional, List, Any, Callable
import traceback # 追加

import keyring
import pyperclip
# winsound is Windows-only; import lazily/optionally
try:
    import winsound  # type: ignore
except Exception:  # pragma: no cover - non-Windows
    winsound = None  # type: ignore
import keyboard # 追加
import ctypes
import sys
from PIL import Image, ImageGrab
from google.api_core import exceptions
from pystray import Icon, Menu, MenuItem
from google.generativeai import types
import google.generativeai as genai

from common_models import BaseAgent, LlmAgent, Prompt, PromptParameters, create_image_part # create_image_partをインポート
from config_manager import load_config, save_config
from constants import API_SERVICE_ID, APP_NAME, COMPLETION_SOUND_FILE, ICON_FILE
from matrix_batch_processor import MatrixBatchProcessorWindow
from ui_components import ActionSelectorWindow, NotificationPopup, SettingsWindow, ResizableInputDialog
from i18n import tr

class ClipboardToolAgent(BaseAgent):
    def __init__(self, name: str = "ClipboardToolAgent", description: str = "クリップボード操作とLLM処理を行うエージェント"):
        super().__init__(name, description)
        self.config = load_config()
        if not self.config:
            sys.exit(1)
        
        # API価格情報を読み込む
        self.api_price_info = self._load_api_price_info()

        self.api_key = self._get_api_key()
        # genai.configure を使用してAPIキーを設定
        if self.api_key:
            genai.configure(api_key=self.api_key)

        self.task_queue = queue.Queue()
        self.loop = None
        self.worker_thread = threading.Thread(target=self._async_worker, daemon=True)
        self._worker_running = True
        self._loop_ready_event = threading.Event()
        self.worker_thread.start()
        self._loop_ready_event.wait(timeout=5)
        if not self._loop_ready_event.is_set():
            sys.exit(1)

        self.app: Optional[ctk.CTk] = None
        self.matrix_batch_processor_window: Optional[MatrixBatchProcessorWindow] = None
        self._current_notification_popup_window: Optional[NotificationPopup] = None
        self._current_action_selector_window: Optional[ActionSelectorWindow] = None
        self._settings_window: Optional[SettingsWindow] = None

        self.clipboard_history = []
        self.max_history_size = self.config.max_history_size
        self._clipboard_monitor_thread: Optional[threading.Thread] = None
        self._clipboard_monitor_running = False
        self._on_history_updated_callback: Optional[Callable[[List[str]], None]] = None

        # Initialize hotkey-related attributes before registering
        self._hotkey_thread: Optional[threading.Thread] = None
        self._hotkey_user32 = None
        self._hotkey_id_map: Dict[int, Callable] = {}
        # list to store Windows hotkey specifications for the original implementation. Not used by the new Windows hotkey handler.
        self._hotkey_registrations: List[tuple] = []

        # New Windows-hotkey-related attributes for a more robust implementation. These
        # will be used by _register_hotkey_windows2() and friends to avoid conflicts
        # with the original implementation. Each hotkey is registered in a dedicated
        # message-loop thread so that WM_HOTKEY messages are delivered to the correct
        # message queue.
        self._win_hotkey_thread: Optional[threading.Thread] = None
        self._win_hotkey_user32 = None
        self._win_hotkey_id_map: Dict[int, Callable] = {}
        self._win_hotkey_registrations: List[tuple] = []
        self._register_hotkey()

    def _load_api_price_info(self) -> Dict:
        """api_price.jsonファイルを読み込む"""
        try:
            price_file_path = Path("api_price.json")
            if price_file_path.exists():
                with open(price_file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                print(f"WARNING: {price_file_path} が見つかりません。デフォルトの価格情報を使用します。")
                return {}
        except Exception as e:
            print(f"ERROR: API価格情報の読み込みに失敗しました: {e}")
            return {}

    def _get_model_pricing(self, model_name: str, input_token_count: int = 0) -> tuple:
        """モデル名と入力トークン数に基づいて価格情報を取得する"""
        if not self.api_price_info:
            return 0.0, 0.0
            
        # モデル名の完全一致を試す
        model_info = self.api_price_info.get(model_name)
        if model_info:
            # 階層情報がある場合
            if "tiers" in model_info:
                # トークン数に応じた価格を取得
                for tier in sorted(model_info["tiers"], key=lambda x: x.get("threshold_tokens", 0), reverse=True):
                    if input_token_count <= tier.get("threshold_tokens", 0) or tier.get("threshold_tokens", 0) == -1:
                        return (
                            tier.get("input_cost_per_thousand_tokens", 0.0),
                            tier.get("output_cost_per_thousand_tokens", 0.0)
                        )
                # デフォルト価格を使用
                default = model_info.get("default", {})
                return (
                    default.get("input_cost_per_thousand_tokens", 0.0),
                    default.get("output_cost_per_thousand_tokens", 0.0)
                )
            # 階層情報がない場合（フラットな構造）
            elif "input_cost_per_thousand_tokens" in model_info:
                return (
                    model_info.get("input_cost_per_thousand_tokens", 0.0),
                    model_info.get("output_cost_per_thousand_tokens", 0.0)
                )
        
        # 部分一致で検索
        for key, model_info in self.api_price_info.items():
            if key in model_name:
                if "tiers" in model_info:
                    for tier in sorted(model_info["tiers"], key=lambda x: x.get("threshold_tokens", 0), reverse=True):
                        if input_token_count <= tier.get("threshold_tokens", 0) or tier.get("threshold_tokens", 0) == -1:
                            return (
                                tier.get("input_cost_per_thousand_tokens", 0.0),
                                tier.get("output_cost_per_thousand_tokens", 0.0)
                            )
                    default = model_info.get("default", {})
                    return (
                        default.get("input_cost_per_thousand_tokens", 0.0),
                        default.get("output_cost_per_thousand_tokens", 0.0)
                    )
                elif "input_cost_per_thousand_tokens" in model_info:
                    return (
                        model_info.get("input_cost_per_thousand_tokens", 0.0),
                        model_info.get("output_cost_per_thousand_tokens", 0.0)
                    )
        
        # 見つからない場合はデフォルト値
        return 0.0, 0.0

    def _register_hotkey(self):
        """
        Register global hotkeys. On Windows, use the system API (RegisterHotKey) for higher reliability.
        On other platforms, fall back to the keyboard library.
        """
        # Always clean up existing hotkeys first. This resets both the legacy
        # RegisterHotKey implementation and the newer implementation below.
        self._unregister_hotkeys_windows()
        # Clean up any hotkeys registered by the new Windows implementation
        try:
            self._unregister_hotkeys_windows2()
        except AttributeError:
            # Method may not exist yet during initialization
            pass
        # Windows-specific registration
        if sys.platform.startswith("win"):
            try:
                # Use the new Windows hotkey registration for improved reliability
                self._register_hotkey_windows2()
                return
            except Exception as e:
                # If Windows registration fails, fall back to keyboard library
                print(f"WARNING: Windows用ホットキーの登録に失敗しました: {e}. keyboardライブラリを使用します。")
        # Fallback: keyboard library
        # Remove any existing hotkeys registered by keyboard
        try:
            keyboard.remove_all_hotkeys()
        except Exception:
            pass
        # Register configured hotkeys if present
        if getattr(self.config, 'hotkey_prompt_list', None):
            try:
                keyboard.add_hotkey(self.config.hotkey_prompt_list, self._show_action_selector_gui)
                print(f"INFO: ホットキー '" + str(self.config.hotkey_prompt_list) + "'（リスト表示）を登録しました (keyboard fallback)。")
            except Exception as e:
                print(f"ERROR: キーボードライブラリでリスト表示ホットキーの登録に失敗しました: {e}")
        if getattr(self.config, 'hotkey_refine', None):
            try:
                keyboard.add_hotkey(self.config.hotkey_refine, self.handle_refine)
                print(f"INFO: ホットキー '" + str(self.config.hotkey_refine) + "'（追加指示）を登録しました (keyboard fallback)。")
            except Exception as e:
                print(f"ERROR: キーボードライブラリで追加指示ホットキーの登録に失敗しました: {e}")
        if getattr(self.config, 'hotkey_matrix', None):
            try:
                keyboard.add_hotkey(self.config.hotkey_matrix, self.show_matrix_batch_processor_window)
                print(f"INFO: ホットキー '" + str(self.config.hotkey_matrix) + "'（マトリクス）を登録しました (keyboard fallback)。")
            except Exception as e:
                print(f"ERROR: キーボードライブラリでマトリクスホットキーの登録に失敗しました: {e}")

    def update_hotkey(self, target: str, new_hotkey: Optional[str]):
        """Update a specific hotkey and re-register all.

        target: 'prompt_list' | 'refine' | 'matrix'
        new_hotkey: hotkey string like 'ctrl+shift+g', or None/"" to disable
        """
        try:
            # Normalize empty to None
            new_val = new_hotkey if new_hotkey else None
            if target == 'prompt_list':
                setattr(self.config, 'hotkey_prompt_list', new_val)
                # For backward compatibility, also mirror to deprecated field
                self.config.hotkey = None
            elif target == 'refine':
                setattr(self.config, 'hotkey_refine', new_val)
            elif target == 'matrix':
                setattr(self.config, 'hotkey_matrix', new_val)
            else:
                raise ValueError(f"Unknown hotkey target: {target}")
            # Re-register
            self._register_hotkey()
            save_config(self.config)
            print(f"INFO: {target} ホットキーを '{new_val}' に更新しました。")
            return True
        except Exception as e:
            print(f"ERROR: ホットキーの更新に失敗しました: {e}")
            return False

    def _parse_hotkey_to_win(self, hotkey_str: str) -> Optional[tuple]:
        """
        Parse a human-readable hotkey string (e.g., 'ctrl+shift+g') into a tuple of
        (modifier_flags, virtual_key_code) suitable for the Windows RegisterHotKey API.

        Returns None if parsing fails.
        """
        if not hotkey_str:
            return None
        mods = 0
        parts = [p.strip() for p in hotkey_str.lower().split('+')]
        key_part = parts[-1]
        for mod in parts[:-1]:
            if mod == 'ctrl':
                mods |= 0x0002  # MOD_CONTROL
            elif mod == 'shift':
                mods |= 0x0004  # MOD_SHIFT
            elif mod == 'alt':
                mods |= 0x0001  # MOD_ALT
            elif mod == 'win':
                mods |= 0x0008  # MOD_WIN
        # Determine the virtual-key code
        if len(key_part) == 1:
            vk = ord(key_part.upper())
        elif key_part.startswith('f') and key_part[1:].isdigit():
            fn = int(key_part[1:])
            vk = 0x70 + (fn - 1)  # F1 -> 0x70
        else:
            return None
        return mods, vk

    def _register_hotkey_windows(self):
        """
        Use the Windows RegisterHotKey API via ctypes to register the main and refine hotkeys.
        Each hotkey gets a unique ID and is stored in _hotkey_id_map with its callback.
        A dedicated message loop thread is started to listen for WM_HOTKEY events.
        """
        """
        Prepare registration data and start the message loop thread. The actual
        RegisterHotKey calls happen in the message loop thread to ensure WM_HOTKEY
        messages are delivered to the correct thread.
        """
        # Parse hotkeys to a list of (mod, vk, callback)
        hotkey_specs: List[tuple] = []
        main_hotkey = self._parse_hotkey_to_win(self.config.hotkey)
        refine_hotkey = self._parse_hotkey_to_win('ctrl+shift+r')
        if main_hotkey:
            hotkey_specs.append((main_hotkey[0], main_hotkey[1], lambda: self._show_action_selector_gui()))
        if refine_hotkey:
            hotkey_specs.append((refine_hotkey[0], refine_hotkey[1], lambda: self.handle_refine()))
        if not hotkey_specs:
            raise RuntimeError('No valid hotkeys specified for Windows registration.')
        # Store registrations for the thread
        self._hotkey_registrations = hotkey_specs
        # Start the message loop thread if not already running
        if not self._hotkey_thread or not self._hotkey_thread.is_alive():
            self._hotkey_thread = threading.Thread(target=self._hotkey_message_loop, daemon=True)
            self._hotkey_thread.start()
        print('INFO: Windows API によるグローバルホットキーを登録しました。')

    def _unregister_hotkeys_windows(self):
        """
        Unregister all hotkeys that were registered via the Windows API and stop the
        message loop thread by posting a quit message. Safe to call on non-Windows platforms.
        """
        # Unregister each hotkey
        if self._hotkey_user32 and self._hotkey_id_map:
            for hid in list(self._hotkey_id_map.keys()):
                try:
                    self._hotkey_user32.UnregisterHotKey(None, hid)
                except Exception:
                    pass
            self._hotkey_id_map.clear()
        # Signal the message loop thread to exit by posting WM_QUIT (0x0012)
        if self._hotkey_user32 and self._hotkey_thread and self._hotkey_thread.is_alive():
            try:
                # Post WM_QUIT directly to the thread ID; threading.Thread.ident
                # corresponds to the Win32 thread ID on Windows.
                self._hotkey_user32.PostThreadMessageW(self._hotkey_thread.ident, 0x0012, 0, 0)
            except Exception:
                pass
        self._hotkey_thread = None

    def _hotkey_message_loop(self):
        """
        Dedicated message loop to listen for WM_HOTKEY messages. When a hotkey is
        triggered, dispatch the associated callback via the Tkinter main loop to ensure
        thread safety.
        """
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        msg = wintypes.MSG()
        while True:
            result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if result == 0:  # WM_QUIT received
                break
            if msg.message == 0x0312:  # WM_HOTKEY
                callback = self._hotkey_id_map.get(msg.wParam)
                if callback:
                    # Dispatch callback in the GUI thread
                    if self.app:
                        self.app.after(0, callback)
                    else:
                        callback()
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    # ----------------------------------------------------------------------
    # New Windows hotkey implementation using its own message loop thread
    # ----------------------------------------------------------------------
    def _register_hotkey_windows2(self):
        """
        Improved Windows global hotkey registration. This method parses the
        configured hotkey and the refine hotkey into (modifier, virtual key)
        tuples and stores them with their callbacks in `_win_hotkey_registrations`.
        It then unregisters any existing Windows hotkeys and starts a new
        message loop thread which performs the actual `RegisterHotKey` calls.
        """
        # Unregister any existing Windows hotkeys registered via this new implementation first
        self._unregister_hotkeys_windows2()
        # Build the registration list
        self._win_hotkey_registrations = []
        # parse prompt list hotkey
        main_tuple = self._parse_hotkey_to_win(getattr(self.config, 'hotkey_prompt_list', None))
        if main_tuple:
            self._win_hotkey_registrations.append((main_tuple[0], main_tuple[1], self._show_action_selector_gui))
        # parse refine hotkey
        refine_tuple = self._parse_hotkey_to_win(getattr(self.config, 'hotkey_refine', None))
        if refine_tuple:
            self._win_hotkey_registrations.append((refine_tuple[0], refine_tuple[1], self.handle_refine))
        # parse matrix hotkey
        matrix_tuple = self._parse_hotkey_to_win(getattr(self.config, 'hotkey_matrix', None))
        if matrix_tuple:
            self._win_hotkey_registrations.append((matrix_tuple[0], matrix_tuple[1], self.show_matrix_batch_processor_window))
        if not self._win_hotkey_registrations:
            raise RuntimeError('No valid hotkeys specified for Windows registration.')
        planned = len(self._win_hotkey_registrations)
        # Start the message loop thread
        self._win_hotkey_thread = threading.Thread(target=self._hotkey_message_loop2, daemon=True)
        self._win_hotkey_thread.start()
        try:
            print(f"INFO: Windowsホットキー用スレッドを起動しました（{planned} 件を登録予定）。")
        except Exception:
            pass

    def _unregister_hotkeys_windows2(self):
        """
        Request the new Windows hotkey message loop thread to exit. A WM_QUIT
        message is posted to the thread's message queue and the thread is
        joined. State variables are cleared afterwards. Safe to call even if
        no thread is running.
        """
        if sys.platform.startswith("win"):
            # If the dedicated hotkey thread is running, ask it to exit cleanly
            if self._win_hotkey_thread and self._win_hotkey_thread.is_alive():
                try:
                    import ctypes
                    user32 = ctypes.windll.user32
                    # Post WM_QUIT directly to the thread ID; threading.Thread.ident is a Win32 thread ID
                    user32.PostThreadMessageW(self._win_hotkey_thread.ident, 0x0012, 0, 0)
                except Exception:
                    pass
                # Wait briefly for the thread to process WM_QUIT and unregister its hotkeys
                try:
                    self._win_hotkey_thread.join(timeout=1.5)
                except Exception:
                    pass
                # As a fallback (e.g., if the thread did not exit in time), attempt to
                # unregister any known IDs from this thread as best-effort to avoid duplicates.
                if self._win_hotkey_thread and self._win_hotkey_thread.is_alive():
                    try:
                        user32 = ctypes.windll.user32
                        for hid in list(self._win_hotkey_id_map.keys()):
                            try:
                                user32.UnregisterHotKey(None, hid)
                            except Exception:
                                pass
                    except Exception:
                        pass
            # Reset state after we've given the thread a chance to clean up
            self._win_hotkey_thread = None
            self._win_hotkey_user32 = None
            self._win_hotkey_id_map.clear()
            self._win_hotkey_registrations = []

    def _hotkey_message_loop2(self):
        """
        Message loop thread for the improved Windows hotkey implementation. This
        function runs in its own thread. It registers all hotkeys listed in
        `_win_hotkey_registrations` via `RegisterHotKey` and then enters
        a standard Win32 message loop. Upon receiving a WM_HOTKEY message,
        the associated callback is invoked (via Tkinter's `after` if a GUI
        exists). When a WM_QUIT is received, all hotkeys are unregistered
        and the thread exits.
        """
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        # store user32 pointer
        self._win_hotkey_user32 = user32
        # register hotkeys in this thread
        self._win_hotkey_id_map = {}
        next_id = 1
        regs = list(self._win_hotkey_registrations)
        # Clear the shared list
        self._win_hotkey_registrations = []
        for mod, vk, cb in regs:
            try:
                if not user32.RegisterHotKey(None, next_id, mod, vk):
                    raise ctypes.WinError()
                self._win_hotkey_id_map[next_id] = cb
                try:
                    print(f"INFO: Windowsホットキー登録成功 id={next_id} (mod=0x{mod:02X}, vk=0x{vk:02X})")
                except Exception:
                    pass
                next_id += 1
            except Exception as e:
                print(f"ERROR: RegisterHotKey failed: {e}")
        # message loop
        msg = wintypes.MSG()
        while True:
            result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if result == 0:
                break
            if msg.message == 0x0312:
                # WM_HOTKEY
                try:
                    print(f"DEBUG: WM_HOTKEY 受信 id={msg.wParam}")
                except Exception:
                    pass
                callback = self._win_hotkey_id_map.get(msg.wParam)
                if callback:
                    try:
                        if self.app:
                            self.app.after(0, callback)
                        else:
                            callback()
                    except Exception:
                        pass
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        # unregister on exit
        for hid in list(self._win_hotkey_id_map.keys()):
            try:
                user32.UnregisterHotKey(None, hid)
            except Exception:
                pass
        self._win_hotkey_id_map.clear()
        self._win_hotkey_user32 = None

    def set_ui_elements(self, app: ctk.CTk, on_history_updated_callback: Optional[Callable[[List[str]], None]] = None):
        self.app = app
        # 監視は常に開始する（履歴ボタン等で履歴を利用するため）
        if on_history_updated_callback:
            self._on_history_updated_callback = on_history_updated_callback
        self._start_clipboard_monitor()

        # 追加指示用の直近結果・設定の初期化
        if not hasattr(self, 'last_result_text'):
            self.last_result_text = None
        if not hasattr(self, 'last_prompt_config'):
            self.last_prompt_config = None
        if not hasattr(self, 'last_generation_params'):
            self.last_generation_params = {}

    def _update_notification_message(self, chunk: str):
        if self._current_notification_popup_window and self._current_notification_popup_window.winfo_exists():
            self._current_notification_popup_window.update_message(chunk)

    def _show_notification_ui(self, title: str, message: str, level: Literal["info", "warning", "error", "success"] = "info", duration_ms: Optional[int] = 3000):
        if self.app:
            if self._current_notification_popup_window and self._current_notification_popup_window.winfo_exists():
                # 既存のウィンドウを再設定
                self._current_notification_popup_window.reconfigure(title, message, level, duration_ms)
            else:
                # 新しいウィンドウを作成
                self._current_notification_popup_window = NotificationPopup(
                    title=title,
                    message=message,
                    parent_app=self.app,
                    level=level,
                    on_destroy_callback=lambda: setattr(self, '_current_notification_popup_window', None)
                )
                self._current_notification_popup_window.show_at_cursor(title, message, level, duration_ms)

    def show_matrix_batch_processor_window(self, icon=None, item=None):
        if self.app:
            self.app.after(0, self._show_matrix_batch_processor_gui)

    def _show_matrix_batch_processor_gui(self):
        if self.matrix_batch_processor_window and self.matrix_batch_processor_window.winfo_exists():
            self.matrix_batch_processor_window.destroy()
            self.matrix_batch_processor_window = None

        # マトリクスに含めるフラグが立っているプロンプトのみをデフォルトで表示する
        # すべてのプロンプトをコピーして使用すると行列UIでの編集が設定ファイルに影響しない
        filtered_prompts = {pid: prompt for pid, prompt in self.config.prompts.items() if getattr(prompt, "include_in_matrix", False)}
        self.matrix_batch_processor_window = MatrixBatchProcessorWindow(
            prompts=filtered_prompts,
            on_processing_completed=self._on_batch_processing_completed,
            llm_agent_factory=self._create_llm_agent_for_matrix,
            notification_callback=self._show_notification_ui,
            worker_loop=self.loop,
            parent_app=self.app,
            agent=self
        )
        self.matrix_batch_processor_window.deiconify()
        self.matrix_batch_processor_window.lift()
        try:
            self.matrix_batch_processor_window.grab_set()
            self.app.wait_window(self.matrix_batch_processor_window)
        finally:
            if self.matrix_batch_processor_window and self.matrix_batch_processor_window.winfo_exists():
                try:
                    self.matrix_batch_processor_window.grab_release()
                except tk.TclError as e:
                    print(f"WARNING: _show_matrix_batch_processor_gui - grab_release中にTclErrorが発生しました: {e}")

    def _create_llm_agent_for_matrix(self, name: str, prompt_config: Prompt) -> LlmAgent:
        return LlmAgent(name=name, prompt_config=prompt_config)

    def _on_batch_processing_completed(self, result: str):
        pass

    def notify_prompts_changed(self):
        """Notify open Matrix window to refresh its prompt set from current config."""
        try:
            if self.app and self.matrix_batch_processor_window and self.matrix_batch_processor_window.winfo_exists():
                # Reflect latest prompts into matrix window on UI thread
                self.app.after(0, lambda: self.matrix_batch_processor_window.on_prompts_updated(self.config.prompts))
        except Exception:
            pass

    def _async_worker(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        self.loop.set_debug(False) # 内部ログを抑制するためデバッグモードを無効化
        self._loop_ready_event.set()

        def _run_task_from_queue():
            try:
                task_info = self.task_queue.get_nowait()
                if task_info is None:
                    self.loop.call_soon_threadsafe(self.loop.stop)
                    return

                async def _process_task_internal():
                    try:
                        await self.run_async(**task_info)
                    except Exception as e:
                        error_message = tr("notify.agent_run_error", details=str(e))
                        if self.app:
                            self.app.after(0, lambda msg=error_message: self._show_notification_ui(tr("common.error"), msg, "error"))
                    finally:
                        self.task_queue.task_done()
                
                try:
                    self.loop.create_task(_process_task_internal())
                except Exception as e:
                    error_message = tr("notify.task_create_unexpected", details=str(e))
                    if self.app:
                        self.app.after(0, lambda msg=error_message: self._show_notification_ui(tr("common.error"), msg, "error"))
                    self.task_queue.task_done()
            except queue.Empty:
                pass
            except Exception as e:
                if self.app:
                    self.app.after(0, lambda e=e: self._show_notification_ui(tr("common.error"), tr("notify.worker_queue_unexpected", details=str(e)), "error"))
            finally:
                if self._worker_running:
                    self.loop.call_later(0.1, _run_task_from_queue)

        self.loop.call_soon_threadsafe(_run_task_from_queue)
        try:
            self.loop.run_forever()
        finally:
            if self.loop and not self.loop.is_closed():
                self.loop.close()

    def _copy_to_clipboard_and_notify(self, processed_text: str, prompt_config: Prompt, cost_message: str = ""):
        try:
            pyperclip.copy(processed_text)
            time.sleep(0.05)
            pasted_text = pyperclip.paste()
            if pasted_text != processed_text:
                 self.app.clipboard_clear()
                 self.app.clipboard_append(processed_text)
                 self.app.update()
        except Exception as e:
            print(f"ERROR: Clipboard operation failed: {e}")

        if not self.app:
            return

        threading.Thread(target=self._play_completion_sound, daemon=True).start()
        self._show_notification_ui(tr("notify.done_title"), tr("notify.copied_fmt", name=prompt_config.name, cost=cost_message), level="success")

    def _get_api_key(self) -> Optional[str]:
        return keyring.get_password(API_SERVICE_ID, "api_key")

    async def _process_clipboard_content(self, file_paths: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        contents = []
        if file_paths:
            # genai.Client を使用せず、genai.upload_file を直接使用
            for file_path in file_paths:
                try:
                    import mimetypes
                    mime_type, _ = mimetypes.guess_type(file_path)
                    if not mime_type:
                        mime_type = "application/octet-stream"
                    uploaded_file = await asyncio.to_thread(genai.upload_file, path=file_path, mime_type=mime_type) # genai.upload_file を使用
                    contents.append({"type": "file", "file_ref": uploaded_file})
                except Exception as e:
                    error_message = tr("notify.file_upload_failed", details=str(e))
                    self._show_notification_ui(tr("notify.file_upload_error"), error_message, "error")
                    raise RuntimeError(error_message)
            return contents
        else:
            max_retries = 3
            for i in range(max_retries):
                try:
                    image = ImageGrab.grabclipboard()
                    if isinstance(image, Image.Image):
                        if image.mode != 'RGB':
                            image = image.convert('RGB')
                        buffered = BytesIO()
                        image.save(buffered, format="PNG")
                        image_data = base64.b64encode(buffered.getvalue()).decode("utf-8")
                        contents.append({"type": "image", "data": image_data})
                        return contents
                    else:
                        original_text = pyperclip.paste()
                        if not original_text:
                            raise ValueError(tr("notify.clipboard_empty"))
                        contents.append({"type": "text", "data": original_text})
                        return contents
                except Exception as e:
                    await asyncio.sleep(0.1 * (i + 1))
            self._show_notification_ui(tr("notify.clipboard_error"), tr("notify.clipboard_get_failed"), "error")
            raise RuntimeError(tr("notify.clipboard_get_failed"))

    async def run_async(self, prompt_id: Optional[str] = None, file_paths: Optional[List[str]] = None, system_prompt: Optional[str] = None, model: Optional[str] = None, temperature: Optional[float] = None, top_p: Optional[float] = None, top_k: Optional[int] = None, max_output_tokens: Optional[int] = None, stop_sequences: Optional[List[str]] = None, refine_instruction: Optional[str] = None) -> str:
        if not self.app:
            raise RuntimeError("UI application not initialized.")

        if not self.api_key:
            error_message = tr("notify.api_key_missing_message")
            self._show_notification_ui(tr("notify.api_key_missing_title"), error_message, level="error")
            raise RuntimeError(error_message)

        try:
            final_prompt_name = tr("free_input.manual_prompt")
            final_system_prompt = system_prompt
            final_model_name = model
            final_temperature = temperature

            if refine_instruction:
                # 追加指示: prompt_id が指定されていればその設定を、無ければ直近設定を使用
                if prompt_id:
                    prompt_config = self.config.prompts.get(prompt_id)
                    if not prompt_config:
                        raise ValueError(tr("notify.prompt_missing_fmt", id=prompt_id))
                else:
                    if not getattr(self, 'last_result_text', None) or not getattr(self, 'last_prompt_config', None):
                        raise ValueError(tr("notify.no_last_result"))
                    prompt_config = self.last_prompt_config
                final_prompt_name = f"{prompt_config.name}{tr('refine.suffix')}"
                final_system_prompt = prompt_config.system_prompt
                final_model_name = prompt_config.model
                final_temperature = prompt_config.parameters.temperature
            elif prompt_id:
                prompt_config = self.config.prompts.get(prompt_id)
                if not prompt_config:
                    raise ValueError(tr("notify.prompt_missing_fmt", id=prompt_id))
                final_prompt_name = prompt_config.name
                final_system_prompt = prompt_config.system_prompt
                final_model_name = prompt_config.model
                final_temperature = prompt_config.parameters.temperature
            elif not (final_system_prompt and final_model_name and final_temperature is not None):
                raise ValueError(tr("notify.no_content"))

            self._show_notification_ui(tr("notify.running_fmt", name=final_prompt_name), tr("notify.sending"), duration_ms=None)

            contents_to_send = []
            if refine_instruction:
                contents_to_send = [
                    f"{tr('refine.prev_output_label')}\n{self.last_result_text}",
                    f"{tr('refine.additional_input_label')}\n{refine_instruction}",
                    tr("refine.requirements_text"),
                ]
            else:
                # まず一時入力オーバーライドがあればそれを使用
                temp_override = getattr(self, '_temp_input_for_processing', None)
                if temp_override is not None:
                    processed_contents = [temp_override]
                    try:
                        delattr(self, '_temp_input_for_processing')
                    except Exception:
                        pass
                else:
                    processed_contents = await self._process_clipboard_content(file_paths)
                for content_info in processed_contents:
                    if content_info["type"] == "image":
                        image_part = create_image_part(content_info["data"]) # 共通関数を呼び出し
                        contents_to_send.append(image_part)
                    elif content_info["type"] == "file":
                        contents_to_send.append(content_info["file_ref"])
                    else: # テキストの場合
                        text_content = content_info['data']
                        contents_to_send.append(text_content)

            # Web検索ツールの有効化判定（SDK差異に備えつつ試す）
            tools_list = None
            try:
                enable_web_flag = False
                cfg = None
                if refine_instruction:
                    cfg = self.last_prompt_config
                elif prompt_id:
                    cfg = self.config.prompts.get(prompt_id)
                if cfg:
                    enable_web_flag = bool(getattr(cfg, 'enable_web', False))
                has_text_url = any(isinstance(c, str) and c.strip().startswith(("http://", "https://")) for c in contents_to_send)
                if enable_web_flag or has_text_url:
                    tools_list = [{"google_search": {}}]
            except Exception:
                tools_list = None

            model_instance = genai.GenerativeModel(final_model_name, system_instruction=final_system_prompt)

            # GenerationConfig を構築（安全設定はデフォルト、ツールは有効なら付与）
            try:
                generate_content_config = types.GenerationConfig(
                    temperature=final_temperature,
                    top_p=top_p if prompt_id is None else (prompt_config.parameters.top_p if prompt_config and prompt_config.parameters else None),
                    top_k=top_k if prompt_id is None else (prompt_config.parameters.top_k if prompt_config and prompt_config.parameters else None),
                    max_output_tokens=max_output_tokens if prompt_id is None else (prompt_config.parameters.max_output_tokens if prompt_config and prompt_config.parameters else None),
                    stop_sequences=stop_sequences if prompt_id is None else (prompt_config.parameters.stop_sequences if prompt_config and prompt_config.parameters else None),
                    tools=tools_list,
                )
            except TypeError:
                # toolsフィールドが未対応のSDKの場合はツールなしで再構築
                generate_content_config = types.GenerationConfig(
                    temperature=final_temperature,
                    top_p=top_p if prompt_id is None else (prompt_config.parameters.top_p if prompt_config and prompt_config.parameters else None),
                    top_k=top_k if prompt_id is None else (prompt_config.parameters.top_k if prompt_config and prompt_config.parameters else None),
                    max_output_tokens=max_output_tokens if prompt_id is None else (prompt_config.parameters.max_output_tokens if prompt_config and prompt_config.parameters else None),
                    stop_sequences=stop_sequences if prompt_id is None else (prompt_config.parameters.stop_sequences if prompt_config and prompt_config.parameters else None),
                )
            
            input_token_count = 0
            try:
                # model_instance.count_tokens を使用
                count_tokens_response = await asyncio.to_thread(
                    model_instance.count_tokens,
                    contents=contents_to_send,
                )
                input_token_count = count_tokens_response.total_tokens
                print(f"DEBUG: Input token count: {input_token_count}")
            except Exception as e:
                print(f"WARNING: Failed to count input tokens: {e}")

            full_response_text = ""
            # 生成（Web検索ツールがエラーならフォールバック）
            def _gen(stream_flag: bool, config):
                return model_instance.generate_content(
                    contents=contents_to_send,
                    stream=stream_flag,
                    generation_config=config,
                )

            try:
                responses = _gen(True, generate_content_config)
            except Exception as e:
                # google_search キーで失敗した可能性。fallback: google_search_retrieval
                if tools_list:
                    try:
                        alt_tools = [{"google_search_retrieval": {}}]
                        alt_config = types.GenerationConfig(
                            temperature=generate_content_config.temperature,
                            top_p=generate_content_config.top_p,
                            top_k=generate_content_config.top_k,
                            max_output_tokens=generate_content_config.max_output_tokens,
                            stop_sequences=generate_content_config.stop_sequences,
                            tools=alt_tools,
                        )
                        responses = _gen(True, alt_config)
                    except Exception:
                        # 最終フォールバック: ツールなし
                        responses = _gen(True, types.GenerationConfig(
                            temperature=generate_content_config.temperature,
                            top_p=generate_content_config.top_p,
                            top_k=generate_content_config.top_k,
                            max_output_tokens=generate_content_config.max_output_tokens,
                            stop_sequences=generate_content_config.stop_sequences,
                        ))
                else:
                    raise

            for chunk in responses:
                # chunk.text を参照する際に ValueError を吐くことがあるため安全に取得する
                text = None
                try:
                    text = chunk.text
                except Exception:
                    # `chunk.text` が取得できない場合は候補が安全性によりブロックされているとみなす
                    pass

                if text:
                    full_response_text += text
                    # クロージャ内で chunk.text を再度評価しないよう text を閉じ込める
                    self.app.after(0, lambda c=text: self._update_notification_message(c))
                elif chunk.prompt_feedback and chunk.prompt_feedback.block_reason:
                    # Prompt was blocked due to safety settings
                    full_response_text = tr("safety.request_blocked_message")
                    self.app.after(0, lambda: self._show_notification_ui(tr("safety.request_blocked_title"), full_response_text, level="error"))
                    break  # Stop processing further chunks
                elif chunk.candidates and (not chunk.candidates[0].content.parts or chunk.candidates[0].finish_reason):
                    # Candidate was blocked or finished due to safety settings or other reasons
                    full_response_text = tr("safety.response_blocked_message")
                    self.app.after(0, lambda: self._show_notification_ui(tr("safety.response_blocked_title"), full_response_text, level="error"))
                    break  # Stop processing further chunks

            output_token_count = (await asyncio.to_thread(
                model_instance.count_tokens,
                contents=[full_response_text],
            )).total_tokens
            print(f"DEBUG: Output token count: {output_token_count}")

            # 価格情報を取得（推定コストの表示用）
            input_cost_per_thousand_tokens, output_cost_per_thousand_tokens = self._get_model_pricing(final_model_name, input_token_count)

            estimated_cost = (input_token_count / 1000) * input_cost_per_thousand_tokens + \
                             (output_token_count / 1000) * output_cost_per_thousand_tokens

            cost_message_suffix = ""
            if input_cost_per_thousand_tokens == 0.0 and output_cost_per_thousand_tokens == 0.0:
                cost_message_suffix = tr("pricing.unavailable_suffix")

            cost_message = (f"{tr('pricing.estimated_cost_prefix')}{estimated_cost:.6f}{cost_message_suffix}"
                            if (input_token_count or output_token_count) else cost_message_suffix)

            if full_response_text:
                final_prompt_config = Prompt(
                    name=final_prompt_name,
                    model=final_model_name,
                    system_prompt=final_system_prompt,
                    parameters=PromptParameters(
                        temperature=final_temperature,
                        top_p=generate_content_config.top_p,
                        top_k=generate_content_config.top_k,
                        max_output_tokens=generate_content_config.max_output_tokens,
                        stop_sequences=generate_content_config.stop_sequences,
                    ),
                    enable_web=bool(tools_list)
                )
                self._copy_to_clipboard_and_notify(full_response_text, final_prompt_config, cost_message)
                # 直近結果を保持（追加指示用）
                self.last_result_text = full_response_text
                self.last_prompt_config = final_prompt_config
                self.last_generation_params = {
                    "temperature": final_temperature,
                    "top_p": generate_content_config.top_p,
                    "top_k": generate_content_config.top_k,
                    "max_output_tokens": generate_content_config.max_output_tokens,
                    "stop_sequences": generate_content_config.stop_sequences,
                }
            else:
                if self._current_notification_popup_window and self._current_notification_popup_window.winfo_exists():
                    self.app.after(0, self._current_notification_popup_window.destroy)

            return full_response_text

        except exceptions.GoogleAPICallError as e:
            error_message = tr("notify.api_error_message", code=e.code, message=e.message)
            self._show_notification_ui(tr("notify.api_error_title"), error_message, level="error")
            raise RuntimeError(error_message)
        except Exception as e:
            error_message = tr("notify.unexpected_error", details=str(e))
            if "finish_reason: SAFETY" in str(e) or (hasattr(e, '__cause__') and e.__cause__ and "finish_reason: SAFETY" in str(e.__cause__)):
                error_message = tr("safety.request_blocked_message")
                self._show_notification_ui(tr("safety.request_blocked_title"), error_message, level="error")
            else:
                self._show_notification_ui(tr("common.error"), error_message, level="error")
            traceback.print_exc() # スタックトレースを出力
            raise

    def _play_completion_sound(self):
        """Play a short completion sound, if supported on this platform.

        - On Windows, use winsound to play the bundled file or a system alias.
        - On other platforms, silently skip (no external deps introduced).
        """
        sound_file = Path(COMPLETION_SOUND_FILE)
        try:
            if winsound is not None:
                if sound_file.exists():
                    winsound.PlaySound(str(sound_file.resolve()), winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception as e:
            print(f"ERROR: Sound playback failed: {e}")

    def add_prompt(self, prompt_id: str, prompt: Prompt):
        if prompt_id in self.config.prompts:
            raise ValueError(tr("prompt.id_exists", id=prompt_id))
        self.config.prompts[prompt_id] = prompt

    def update_prompt(self, prompt_id: str, updated_prompt: Prompt):
        if prompt_id not in self.config.prompts:
            raise ValueError(tr("prompt.id_missing", id=prompt_id))
        self.config.prompts[prompt_id] = updated_prompt

    def delete_prompt(self, prompt_id: str):
        if prompt_id not in self.config.prompts:
            raise ValueError(tr("prompt.id_missing", id=prompt_id))
        del self.config.prompts[prompt_id]

    def _on_prompt_selected(self, prompt_id: str, file_paths: Optional[List[str]] = None):
        self._run_process_in_thread(prompt_id=prompt_id, file_paths=file_paths)

    def _run_process_in_thread(self, **kwargs):
        try:
            self.task_queue.put(kwargs)
        except Exception as e:
            self._show_notification_ui(tr("common.error"), tr("notify.task_enqueue_failed", details=str(e)), level="error")

    def _show_action_selector_gui(self, *args, **kwargs):
        """
        ActionSelectorWindowを表示する。
        pystrayからの呼び出しと、内部からの呼び出し(file_paths付き)の両方に対応する。
        """
        try:
            print("DEBUG: _show_action_selector_gui invoked.")
        except Exception:
            pass
        file_paths = kwargs.get('file_paths')

        # 既存のウィンドウがあれば、安全に破棄する
        if self._current_action_selector_window and self._current_action_selector_window.winfo_exists():
            self._current_action_selector_window.destroy()
            self._current_action_selector_window = None

        # カーソル位置を取得
        cursor_x = self.app.winfo_pointerx()
        cursor_y = self.app.winfo_pointery()

        # 新しいウィンドウを作成して表示する
        self.app.after(50, lambda: self._create_and_show_action_selector(file_paths=file_paths, cursor_pos=(cursor_x, cursor_y)))

    def _create_and_show_action_selector(self, file_paths: Optional[List[str]] = None, cursor_pos: Optional[tuple] = None):
        """ActionSelectorWindowを作成して表示するヘルパーメソッド"""
        # 既に別のインスタンスが存在していたら何もしない（念のため）
        if self._current_action_selector_window and self._current_action_selector_window.winfo_exists():
            return

        if self.app:
            self._current_action_selector_window = ActionSelectorWindow(
                prompts=self.config.prompts,
                on_prompt_selected_callback=self._on_prompt_selected,
                agent=self,
                file_paths=file_paths,
                on_destroy_callback=lambda: setattr(self, '_current_action_selector_window', None)
            )
            self._current_action_selector_window.show_at_cursor(cursor_pos=cursor_pos)

    def _show_main_window(self, icon=None, item=None):
        if self.app:
            try:
                # Release any existing Tk grab globally to avoid blocked interactions
                try:
                    cur = self.app.tk.call('grab', 'current')
                    if cur:
                        try:
                            self.app.nametowidget(cur).grab_release()
                        except Exception:
                            # Fallback: release grab at root level if possible
                            self.app.grab_release()
                except Exception:
                    pass
                self.app.deiconify()
                # Temporarily set topmost to reliably lift above other toplevels (e.g., matrix window)
                try:
                    self.app.attributes("-topmost", True)
                except Exception:
                    pass
                self.app.lift()
                try:
                    self.app.focus_force()
                except Exception:
                    pass
                # Drop topmost shortly after so normal z-order behavior resumes
                self.app.after(250, lambda: self._unset_topmost_safe())
            except Exception:
                pass

    def _unset_topmost_safe(self):
        try:
            if self.app and self.app.winfo_exists():
                self.app.attributes("-topmost", False)
        except Exception:
            pass

    def show_settings_window(self, icon=None, item=None):
        print("DEBUG: show_settings_window called.")
        if self.app:
            if self._settings_window and self._settings_window.winfo_exists():
                print("DEBUG: show_settings_window - Existing settings window found, destroying it.")
                self._settings_window.destroy()
                self._settings_window = None
            self._settings_window = SettingsWindow(parent_app=self.app, agent=self)
            self._settings_window.deiconify()
            self._settings_window.lift()
            try:
                print(f"DEBUG: show_settings_window - Calling grab_set. Current grab: {self.app.grab_current()}")
                self._settings_window.grab_set()
                print(f"DEBUG: show_settings_window - grab_set called. New grab: {self.app.grab_current()}")
                self.app.wait_window(self._settings_window)
                print("DEBUG: show_settings_window - Settings window closed.")
            finally:
                if self._settings_window and self._settings_window.winfo_exists():
                    try:
                        print(f"DEBUG: show_settings_window - Calling grab_release. Current grab: {self.app.grab_current()}")
                        self._settings_window.grab_release()
                        print("DEBUG: show_settings_window - grab_release called successfully.")
                    except tk.TclError as e:
                        print(f"WARNING: show_settings_window - grab_release中にTclErrorが発生しました: {e}")
                self._settings_window = None
        else:
            print("ERROR: show_settings_window - self.app is not initialized.")

    def handle_free_input(self):
        # ActionSelectorWindowが存在し、grab_setされている場合、grab_releaseを呼び出す
        if self._current_action_selector_window and self._current_action_selector_window.winfo_exists():
            if self._current_action_selector_window.grab_current() == str(self._current_action_selector_window):
                try:
                    self._current_action_selector_window.grab_release()
                    print("DEBUG: ActionSelectorWindow grab_release called before opening ResizableInputDialog.")
                except tk.TclError as e:
                    print(f"WARNING: handle_free_input - ActionSelectorWindow grab_release中にTclErrorが発生しました: {e}")

        dialog = ResizableInputDialog(parent_app=self.app, text=tr("free_input.prompt_label"), title=tr("free_input.title"))
        dialog.show() # ここでshow()メソッドを呼び出す
        prompt_text = dialog.get_input()
        if prompt_text:
            # Check if there's a temporary file path from a previous file attachment
            file_paths_to_process = getattr(self, '_temp_file_paths_for_processing', None)
            self._run_process_in_thread(system_prompt=prompt_text, model="gemini-2.5-flash-lite", temperature=1.0, file_paths=file_paths_to_process)
            # Clear the temporary file path after use
            if hasattr(self, '_temp_file_paths_for_processing'):
                del self._temp_file_paths_for_processing

    def _on_prompt_selected(self, prompt_id: str, file_paths: Optional[List[str]] = None):
        # もし一時ファイルパスが設定されていればそれを使用し、そうでなければNone
        final_file_paths = file_paths if file_paths else getattr(self, '_temp_file_paths_for_processing', None)
        
        self._run_process_in_thread(prompt_id=prompt_id, file_paths=final_file_paths)
        
        # 処理後、一時ファイルパスをクリア
        if hasattr(self, '_temp_file_paths_for_processing'):
            del self._temp_file_paths_for_processing

    def handle_refine(self, icon=None, item=None):
        if not getattr(self, 'last_result_text', None):
            self._show_notification_ui(tr("refine.title"), tr("notify.no_last_result"), level="warning")
            return
        dialog = ResizableInputDialog(parent_app=self.app, text=tr("refine.prompt_label"), title=tr("refine.title"))
        dialog.show()
        instruction = dialog.get_input()
        if instruction:
            self._run_process_in_thread(refine_instruction=instruction)

    def handle_file_attach(self):
        file_paths = filedialog.askopenfilenames() # 複数ファイル選択を許可
        if file_paths:
            # 絶対パスに正規化して保存
            self._temp_file_paths_for_processing = [str(Path(p).resolve()) for p in file_paths]
            
            # 選択されたファイル名を通知として表示
            # file_names = [Path(p).name for p in file_paths]
            # notification_message = "以下のファイルを添付しました:\n" + "\n".join(file_names)
            # self._show_notification_ui("ファイル添付", notification_message, level="info", duration_ms=5000)

            # ファイル選択後、プロンプト選択画面を再表示
            self._show_action_selector_gui(file_paths=list(file_paths))
        # else:
        #     self._show_notification_ui("ファイル添付", "ファイルが選択されませんでした。", level="info", duration_ms=2000)

    def quit_app(self, icon=None, item=None):
        # Unregister Windows hotkeys (if any) and stop the listener thread
        try:
            self._unregister_hotkeys_windows()
        except Exception:
            pass
        # Remove all keyboard hotkeys to clean up fallback registrations
        try:
            keyboard.unhook_all()
        except Exception:
            pass
        self._worker_running = False
        if self._clipboard_monitor_thread and self._clipboard_monitor_thread.is_alive():
            self._clipboard_monitor_running = False
            self._clipboard_monitor_thread.join(timeout=1.0)
        if self.task_queue:
            self.task_queue.put(None)
        if self.worker_thread and self.worker_thread.is_alive():
            self.worker_thread.join(timeout=1.0)
        
        if self._current_action_selector_window and self._current_action_selector_window.winfo_exists():
            self._current_action_selector_window.destroy()

        if self.matrix_batch_processor_window and self.matrix_batch_processor_window.winfo_exists():
            self.matrix_batch_processor_window.destroy()

        if self._settings_window and self._settings_window.winfo_exists():
            self._settings_window.destroy()

        if self.app:
            self.app.quit()
            self.app.destroy()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.stop()
        sys.exit(0)

    def _start_clipboard_monitor(self):
        if not self._clipboard_monitor_running:
            self._clipboard_monitor_running = True
            self._clipboard_monitor_thread = threading.Thread(target=self._clipboard_monitor, daemon=True)
            self._clipboard_monitor_thread.start()

    def stop_clipboard_monitor(self):
        """Stop the clipboard monitoring thread gracefully."""
        if self._clipboard_monitor_running:
            self._clipboard_monitor_running = False
            if self._clipboard_monitor_thread and self._clipboard_monitor_thread.is_alive():
                try:
                    self._clipboard_monitor_thread.join(timeout=1.0)
                except Exception:
                    pass

    def _clipboard_monitor(self):
        """テキストだけでなく、画像やファイルの履歴も収集する。"""
        last_signature: Optional[str] = None

        while self._clipboard_monitor_running:
            try:
                items_to_add = None
                signature = None

                # 1) 画像/ファイルのクリップボードを優先チェック
                try:
                    clip_obj = ImageGrab.grabclipboard()
                except Exception:
                    clip_obj = None

                if isinstance(clip_obj, Image.Image):
                    image = clip_obj
                    if image.mode != 'RGB':
                        image = image.convert('RGB')
                    # Convert image to bytes and compress to reduce memory footprint. Use zlib
                    # to compress the raw PNG bytes before Base64 encoding. This makes
                    # history storage more compact for large images.
                    with BytesIO() as buffer:
                        image.save(buffer, format='PNG')
                        image_bytes = buffer.getvalue()
                    try:
                        import zlib
                        compressed = zlib.compress(image_bytes)
                        encoded = base64.b64encode(compressed).decode('utf-8')
                        items_to_add = [{"type": "image_compressed", "data": encoded}]
                    except Exception:
                        # Fallback to storing uncompressed Base64 if compression fails
                        encoded = base64.b64encode(image_bytes).decode('utf-8')
                        items_to_add = [{"type": "image", "data": encoded}]
                    signature = "img:" + hashlib.sha1(image_bytes).hexdigest()
                elif isinstance(clip_obj, list):
                    file_paths = [p for p in clip_obj if isinstance(p, str)]
                    if file_paths:
                        items_to_add = [{"type": "file", "data": p} for p in file_paths]
                        signature = "files:" + "|".join(file_paths)

                # 2) テキストのチェック（上で何も取得できなかった場合）
                if items_to_add is None:
                    try:
                        text_content = pyperclip.paste()
                    except Exception:
                        text_content = ""
                    if text_content:
                        items_to_add = [{"type": "text", "data": text_content}]
                        signature = "text:" + hashlib.sha1(text_content.encode('utf-8')).hexdigest()

                # 3) 新規内容のみ履歴に追加
                if items_to_add and signature != last_signature:
                    for it in items_to_add:
                        self._add_to_history(it)
                    last_signature = signature

                time.sleep(0.5)
            except Exception:
                time.sleep(1)

    def _add_to_history(self, content: Any):
        """履歴にテキスト/画像/ファイル項目を追加。既存重複は先頭へ移動。"""
        # 正規化と空チェック
        if isinstance(content, str):
            normalized = content.strip()
            if not normalized:
                return
        elif isinstance(content, dict):
            # type/data がなければ無視
            if "type" not in content or "data" not in content:
                return
            normalized = content
        else:
            return

        # 既存重複を削除（等価比較）
        existing_index = None
        for idx, item in enumerate(self.clipboard_history):
            try:
                if item == normalized:
                    existing_index = idx
                    break
            except Exception:
                continue
        if existing_index is not None:
            self.clipboard_history.pop(existing_index)

        self.clipboard_history.insert(0, normalized)
        if len(self.clipboard_history) > self.max_history_size:
            self.clipboard_history = self.clipboard_history[:self.max_history_size]
        if self._on_history_updated_callback:
            self.app.after(0, lambda: self._on_history_updated_callback(self.clipboard_history))

    def create_tray_icon(self):
        image = Image.open(ICON_FILE)
        menu = (
            MenuItem(tr('tray.list'), self._show_action_selector_gui, default=True),
            MenuItem(tr('tray.matrix'), self.show_matrix_batch_processor_window),
            # MenuItem('追加指示…', self.handle_refine),
            MenuItem(tr('tray.manager'), self._show_main_window),
            MenuItem(tr('tray.settings'), self.show_settings_window),
            MenuItem(tr('tray.quit'), self.quit_app)
        )
        self.tray_icon = Icon(APP_NAME, image, APP_NAME, menu)
        return self.tray_icon

    def run(self):
        self.icon = self.create_tray_icon()
        self.icon.run()
