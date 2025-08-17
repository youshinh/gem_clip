import mimetypes
import customtkinter as ctk
import tkinter as tk
import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
from tkinter import messagebox, filedialog
from typing import List, Dict, Any, Optional, Callable
import asyncio
import threading
import time
import pyperclip
from common_models import LlmAgent, Prompt
from PIL import Image
from io import BytesIO
import base64
from google.generativeai import types
import google.generativeai as genai
# from google.api_core import exceptions
import styles
from i18n import tr
from pathlib import Path
from constants import DELETE_ICON_FILE
import traceback
from common_models import create_image_part # create_image_partをインポート
from google.generativeai.generative_models import GenerativeModel # GenerativeModelをインポート
from history_dialogs import HistoryEditDialog
from CTkMessagebox import CTkMessagebox

class SizerGrip(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.bind("<B1-Motion>", self._on_motion)
        self.bind("<ButtonPress-1>", self._on_press)
        self.configure(cursor="sizing")

    def _on_press(self, event):
        self._start_x = event.x
        self._start_y = event.y
        self._toplevel = self.winfo_toplevel()

    def _on_motion(self, event):
        x, y = (event.x, event.y)
        w, h = (self.master.winfo_width() + x, self.master.winfo_height() + y)
        self.master.config(width=w, height=h)


class MatrixBatchProcessorWindow(ctk.CTkToplevel):
    def __init__(self, prompts: Dict[str, Prompt], on_processing_completed: Callable, llm_agent_factory: Callable[[str, Prompt], LlmAgent], notification_callback: Callable[[str, str, str], None], worker_loop: asyncio.AbstractEventLoop, parent_app: ctk.CTk, agent: Any):
        super().__init__(parent_app)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.prompts = prompts
        logging.debug(f"DEBUG: __init__ - Initial self.prompts keys: {list(self.prompts.keys())}")
        try:
            self._initial_prompts = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
            logging.debug(f"DEBUG: __init__ - _initial_prompts keys: {list(self._initial_prompts.keys())}")
        except Exception:
            self._initial_prompts = dict(self.prompts)
            logging.debug(f"DEBUG: __init__ - _initial_prompts (exception) keys: {list(self._initial_prompts.keys())}")
        self.on_processing_completed = on_processing_completed
        self.llm_agent_factory = llm_agent_factory
        self.notification_callback = notification_callback
        self.worker_loop = worker_loop
        self.parent_app = parent_app
        self.agent = agent
        self.geometry(styles.MATRIX_WINDOW_GEOMETRY)
        self.title(tr("matrix.window_title")) # ウィンドウタイトルを設定
        self.resizable(True, True)
        # NOTE: Do not mark as transient to keep normal window controls and allow other windows to lift above it
        self._is_closing = False

        self._cursor_update_job = None
        self._start_cursor_monitoring()

        self.processing_tasks: List[asyncio.Task] = []
        self.semaphore = asyncio.Semaphore(5)
        self.total_tasks = 0
        self.completed_tasks = 0
        self.progress_lock = threading.Lock()
        # 入力セルのフレーム参照を保持して、部分更新で再描画を最小化
        self._input_row_frames: List[ctk.CTkFrame] = []

        # 進捗エリアの背景はMATRIX_TOP_BG_COLOR（キャンバスより少し暗め）
        self.progress_frame = ctk.CTkFrame(self, fg_color=styles.MATRIX_TOP_BG_COLOR)
        self.progress_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.progress_label = ctk.CTkLabel(self.progress_frame, text=tr("matrix.progress_fmt", done=0, total=0), font=styles.MATRIX_FONT_BOLD)
        self.progress_label.pack(fill="x")


        self.input_data: List[Dict[str, Any]] = [{"type": "text", "data": ""}]
        self.checkbox_states: List[List[ctk.BooleanVar]] = []
        self.results: List[List[ctk.StringVar]] = []
        self._full_results: List[List[str]] = []
        self._row_summaries: List[ctk.StringVar] = []
        self._col_summaries: List[ctk.StringVar] = []
        self._history_popup: Optional['ClipboardHistorySelectorPopup'] = None

        self.summarize_row_button: Optional[ctk.CTkButton] = None
        self.summarize_col_button: Optional[ctk.CTkButton] = None

        # --- Drag-and-drop (column reorder) state ---
        self._col_header_frames: List[ctk.CTkFrame] = []
        self._col_drag_data: Dict[str, Any] = {}
        self._col_drag_active_frame: Optional[ctk.CTkFrame] = None
        self._col_drop_line_id: Optional[int] = None  # legacy (canvas)
        self._col_drop_indicator_widget: Optional[tk.Frame] = None

        # --- Flow run state ---
        try:
            self.max_flow_steps: int = int(getattr(self.agent.config, 'max_flow_steps', 5))
        except Exception:
            self.max_flow_steps: int = 5
        self._result_textboxes: List[List[Optional[ctk.CTkTextbox]]] = []
        self._cell_style: List[List[str]] = []  # "normal" or "flow"
        self._flow_cancel_requested: bool = False
        self._flow_tasks: List[asyncio.Task] = []

        # --- UIリサイズ用プロパティ ---
        # 各列の幅と各行の高さを保持するリスト。0番目は固定列/ヘッダ行に対応。
        self._column_widths: List[int] = []
        self._row_heights: List[int] = []
        # リサイズ中の状態管理
        self._current_column_resizing: Optional[int] = None
        self._col_resize_start_x: int = 0
        self._col_resize_initial_width: int = 0
        self._current_row_resizing: Optional[int] = None
        self._row_resize_start_y: int = 0
        self._row_resize_initial_height: int = 0

        self._create_toolbar()
        self._init_tabs()
        self._create_main_grid_frame()
        self.after(100, self._update_ui) # 遅延させてUIを更新
        self.state('zoomed') # ウィンドウを最大化

    def on_closing(self):
        """ウィンドウが閉じられる際の処理"""
        if messagebox.askokcancel(tr("confirm.exit_title"), tr("confirm.exit_message")):
            # セッション保存の確認（Yes: 保存して終了 / No: 保存せず終了 / Cancel: 中止）
            save_choice = messagebox.askyesnocancel(tr("confirm.session_save_title"), tr("confirm.session_save_message"))
            if save_choice is None:
                return
            try:
                # Persist current active tab prompts/state before closing
                if hasattr(self, '_tabs') and self._tabs:
                    try:
                        self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                        self._tabs[self._active_tab_index]['state'] = self._snapshot_state()
                    except Exception:
                        pass
                    if save_choice:
                        self._save_session()
            except Exception:
                pass
            self._is_closing = True
            if hasattr(self, '_cursor_update_job') and self._cursor_update_job:
                self.after_cancel(self._cursor_update_job)
                self._cursor_update_job = None
            
            # 実行中の非同期タスクをキャンセル
            if hasattr(self, 'processing_tasks'):
                for task in self.processing_tasks:
                    if not task.done():
                        task.cancel()
            
            self.destroy()

    def on_prompts_updated(self, updated_prompts: Dict[str, Prompt]):
        """外部（プロンプト管理）での変更を即時反映する。
        - 設定画面の「マトリクス」チェックはデフォルトタブに表示するプロンプト。
        - デフォルト以外がアクティブでも、デフォルトタブの内容のみ更新する。
        - アクティブがデフォルトの場合は表示も即時更新。
        """
        try:
            filtered = {pid: p for pid, p in updated_prompts.items() if getattr(p, 'include_in_matrix', False)}
            # デフォルトタブを探す
            default_idx = next((i for i, t in enumerate(self._tabs) if str(t.get('name')) == tr('matrix.tab.default')), None)
            if default_idx is None:
                # なければ作る
                self._tabs.insert(0, {'name': tr('matrix.tab.default'), 'prompts_obj': {}, 'state': None})
                default_idx = 0
            # デフォルトタブのプロンプトを更新
            self._tabs[default_idx]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in filtered.items()}
            logging.debug(f"DEBUG: on_prompts_updated - default tab prompts updated: {list(filtered.keys())}")
            # アクティブがデフォルトなら表示も更新
            if self._active_tab_index == default_idx:
                self.prompts = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in filtered.items()}
                self.checkbox_states = []
                self.results = []
                self._full_results = []
                self._update_ui()
        except Exception:
            try:
                self._update_ui()
            except Exception:
                pass

    def _create_toolbar(self):
        """ツールバーを作成し、ボタンを配置する"""
        toolbar_frame = ctk.CTkFrame(self, fg_color=styles.MATRIX_TOP_BG_COLOR)
        toolbar_frame.pack(fill="x", padx=10, pady=10)
        # Columns: 0..6 action buttons, 7 settings
        toolbar_frame.grid_columnconfigure(0, weight=1)
        toolbar_frame.grid_columnconfigure(1, weight=1)
        toolbar_frame.grid_columnconfigure(2, weight=1)
        toolbar_frame.grid_columnconfigure(3, weight=1)
        toolbar_frame.grid_columnconfigure(4, weight=1)
        toolbar_frame.grid_columnconfigure(5, weight=1)
        toolbar_frame.grid_columnconfigure(6, weight=1)
        toolbar_frame.grid_columnconfigure(7, weight=0)

        ctk.CTkButton(toolbar_frame, text=tr("matrix.add_input"), command=self._add_input_row, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(toolbar_frame, text=tr("matrix.add_prompt"), command=self._add_prompt_column, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(toolbar_frame, text=tr("matrix.add_set"), command=self._add_prompt_set_tab, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(toolbar_frame, text=tr("matrix.clear"), command=self._clear_active_set, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(toolbar_frame, text=tr("matrix.set_manager"), command=self._open_set_manager, fg_color=styles.MATRIX_BUTTON_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=4, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(toolbar_frame, text=tr("matrix.session_manager"), command=self._open_session_manager, fg_color=styles.MATRIX_BUTTON_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=5, padx=5, pady=5, sticky="ew")
        # タブ削除は各タブの「×」で実行
        
        try:
            config_icon_path = Path("config.ico")
            if not config_icon_path.exists():
                raise FileNotFoundError("config.ico not found")

            from PIL import Image
            icon_img = Image.open(config_icon_path)
            size = (24, 24)
            config_icon = ctk.CTkImage(light_image=icon_img, dark_image=icon_img, size=size)

            settings_button = ctk.CTkButton(toolbar_frame, text="", image=config_icon, width=28, height=28, command=self._open_summary_settings)
        except Exception as e:
            print(f"Error loading settings icon: {e}")
            settings_button = ctk.CTkButton(toolbar_frame, text=tr("settings.title"), width=60, command=self._open_summary_settings, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        
        # 設定ボタンは一番最後（右端）
        settings_button.grid(row=0, column=7, padx=5, pady=5, sticky="e")

    # --- Prompt set tabs (max 5) ---
    def _serialize_prompts(self, prompts_dict: Dict[str, Prompt]) -> Dict[str, dict]:
        out: Dict[str, dict] = {}
        for pid, p in prompts_dict.items():
            try:
                out[pid] = p.model_dump(by_alias=True, exclude_none=True)
            except Exception:
                out[pid] = {
                    'name': getattr(p, 'name', ''),
                    'model': getattr(p, 'model', ''),
                    'system_prompt': getattr(p, 'system_prompt', ''),
                    'thinking_level': getattr(p, 'thinking_level', 'Balanced'),
                    'enable_web': getattr(p, 'enable_web', False),
                    'parameters': getattr(getattr(p, 'parameters', None), 'model_dump', lambda **_: {})()
                }
        return out

    def _deserialize_prompts(self, data: Dict[str, dict]) -> Dict[str, Prompt]:
        out: Dict[str, Prompt] = {}
        for pid, pd in data.items():
            try:
                out[pid] = Prompt(**pd)
            except Exception:
                continue
        return out

    def _snapshot_state(self) -> dict:
        chk: list[list[bool]] = []
        for r in self.checkbox_states:
            try:
                chk.append([bool(v.get()) if hasattr(v, 'get') else bool(v) for v in r])
            except Exception:
                chk.append([False for _ in r])
        results_full: list[list[str]] = []
        for r in self._full_results:
            results_full.append([str(c or '') for c in r])
        row_summ = [sv.get() if hasattr(sv, 'get') else str(sv or '') for sv in self._row_summaries]
        col_summ = [sv.get() if hasattr(sv, 'get') else str(sv or '') for sv in self._col_summaries]
        return {'checkbox': chk, 'full_results': results_full, 'row_summaries': row_summ, 'col_summaries': col_summ}

    def _apply_state(self, state: Optional[dict]):
        if not state:
            return
        chk = state.get('checkbox', [])
        self.checkbox_states = []
        for r in range(len(self.input_data)):
            row = []
            src = chk[r] if r < len(chk) else []
            for c in range(len(self.prompts)):
                val = bool(src[c]) if c < len(src) else False
                row.append(ctk.BooleanVar(value=val))
            self.checkbox_states.append(row)
        self._full_results = []
        self.results = []
        fr = state.get('full_results', [])
        for r in range(len(self.input_data)):
            self._full_results.append([])
            self.results.append([])
            src_row = fr[r] if r < len(fr) else []
            for c in range(len(self.prompts)):
                cell_full = str(src_row[c]) if c < len(src_row) else ''
                self._full_results[r].append(cell_full)
                sv = ctk.StringVar(value=self._truncate_result(cell_full))
                self.results[r].append(sv)
        self._row_summaries = []
        rs = state.get('row_summaries', [])
        for r in range(len(self.input_data)):
            self._row_summaries.append(ctk.StringVar(value=str(rs[r]) if r < len(rs) else ''))
        self._col_summaries = []
        cs = state.get('col_summaries', [])
        for c in range(len(self.prompts)):
            self._col_summaries.append(ctk.StringVar(value=str(cs[c]) if c < len(cs) else ''))
    def _prompt_set_dir(self) -> Path:
        try:
            base = Path.cwd()
        except Exception:
            base = Path('.')
        d = base / 'prompt_set'
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _session_file(self) -> Path:
        return self._prompt_set_dir() / 'session.json'

    def _sessions_dir(self) -> Path:
        d = self._prompt_set_dir() / 'sessions'
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _session_file_for(self, name: str) -> Path:
        safe = name.strip().replace('/', '_').replace('\\', '_')
        return self._sessions_dir() / f"{safe}.json"

    def _init_tabs(self):
        # Custom browser-like tab bar container
        # タブバーの背景は進捗エリアと統一（MATRIX_TOP_BG_COLOR）
        self.tabbar_frame = ctk.CTkFrame(self, fg_color=styles.MATRIX_TOP_BG_COLOR)
        self.tabbar_frame.pack(fill='x', padx=10, pady=(0,0))
        # Per-tab storage
        self._tabs: list[dict] = []
        self._active_tab_index: int = 0
        self._tab_slot_width: Optional[int] = None
        # Load from session or default and render
        self._start_with_default_only: bool = True
        self._load_session_or_default()
        self._render_tabbar()
        # Resize: adjust tab widths only (no rebuild) to avoid flicker
        self.tabbar_frame.bind("<Configure>", lambda e: self._adjust_tabbar_widths())

    def _rebuild_tabs(self):
        # Clean and clamp
        self._tabs = [t for t in self._tabs if t.get('name') is not None]
        if not self._tabs:
            self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {}, 'state': None}]
            self._active_tab_index = 0
        if self._active_tab_index >= len(self._tabs):
            self._active_tab_index = max(0, len(self._tabs) - 1)
        # Apply active prompts/state
        active = self._tabs[self._active_tab_index]
        prompts_obj = active.get('prompts_obj') if isinstance(active.get('prompts_obj', {}), dict) else self._deserialize_prompts(active.get('prompts', {}))
        self.prompts = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in prompts_obj.items()} if prompts_obj else {}
        logging.debug(f"DEBUG: _rebuild_tabs - self.prompts keys set to: {list(self.prompts.keys())}")
        self.checkbox_states = []
        self.results = []
        self._full_results = []
        self._row_summaries = []
        self._col_summaries = []
        self._apply_state(active.get('state'))
        self._render_tabbar()

    def _on_tab_clicked(self, idx: int):
        if idx == self._active_tab_index:
            return
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                self._tabs[self._active_tab_index]['state'] = self._snapshot_state()
        except Exception:
            pass
        self._active_tab_index = idx
        try:
            t = self._tabs[self._active_tab_index]
            prompts_obj = t.get('prompts_obj') if isinstance(t.get('prompts_obj', {}), dict) else self._deserialize_prompts(t.get('prompts', {}))
            self.prompts = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in prompts_obj.items()} if prompts_obj else {}
            logging.debug(f"DEBUG: _on_tab_clicked - self.prompts keys set to: {list(self.prompts.keys())}")
        except Exception as e:
            self.prompts = {}
            logging.error(f"ERROR: _on_tab_clicked - Error setting self.prompts: {e}")
        # 状態を適用し新規フレームに描画（確実な切替を優先）
        self._result_textboxes = []
        self._cell_style = []
        self.checkbox_states = []
        self.results = []
        self._full_results = []
        self._row_summaries = []
        self._col_summaries = []
        self._apply_state(self._tabs[self._active_tab_index].get('state'))
        new_frame = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self.scrollable_content_frame = new_frame
        try:
            self.canvas.itemconfigure(self._window_id, window=self.scrollable_content_frame)
        except Exception:
            pass
        self._update_ui()
        self._render_tabbar()

    def _render_tabbar(self):
        # Clear and rebuild tab buttons
        for w in self.tabbar_frame.winfo_children():
            w.destroy()
        max_tabs = 5
        slot_w = self._compute_slot_width()
        self._tab_slot_width = slot_w
        # Determine bottom blend color (match matrix canvas background)
        try:
            cur_fg = self.canvas_frame.cget("fg_color")
            appearance_mode_index = 0 if ctk.get_appearance_mode() == "Light" else 1
            blend_color = cur_fg[appearance_mode_index] if isinstance(cur_fg, tuple) else cur_fg
        except Exception:
            # Fallback to matrix canvas background color constant
            mc = styles.MATRIX_CANVAS_BACKGROUND_COLOR
            appearance_mode_index = 0 if ctk.get_appearance_mode() == "Light" else 1
            blend_color = mc[appearance_mode_index] if isinstance(mc, tuple) else mc
        for i in range(max_tabs):
            is_real = i < len(self._tabs)
            is_active = (i == self._active_tab_index)
            # アクティブタブはCANVAS色、未使用スロットはタブバー背景色、
            # 非アクティブの実タブは既存のグレー
            if not is_real:
                fg = styles.MATRIX_TOP_BG_COLOR
            else:
                fg = blend_color if is_active else ("#D9D9D9", "#1F1F1F")
            # Square-corner tabs (no rounded corners, no border lines)
            outer = ctk.CTkFrame(self.tabbar_frame, corner_radius=0, border_width=0, fg_color=fg, border_color=styles.MATRIX_HEADER_BORDER_COLOR)
            outer.pack(side="left", padx=4, pady=0)
            outer.configure(width=slot_w, height=34)
            outer.pack_propagate(False)
            if is_real:
                name = str(self._tabs[i].get('name') or tr('matrix.tab.auto_name_fmt', n=i+1))
                inner = ctk.CTkFrame(outer, fg_color="transparent")
                inner.pack(fill="both", expand=True, padx=10, pady=(2,4))
                title = ctk.CTkLabel(inner, text=name, font=styles.MATRIX_FONT_BOLD, anchor="w")
                title.pack(side="left", fill="x", expand=True)
                # Double-click to rename
                title.bind("<Double-Button-1>", lambda e, idx=i: self._rename_tab(idx))
                # Use delete icon same as columns/inputs
                try:
                    icon_path = Path(DELETE_ICON_FILE)
                    if icon_path.exists():
                        from PIL import Image
                        _img = Image.open(icon_path)
                        _img.thumbnail((16,16))
                        tab_delete_icon = ctk.CTkImage(light_image=_img, dark_image=_img, size=(16,16))
                    else:
                        tab_delete_icon = None
                except Exception:
                    tab_delete_icon = None
                close_btn = ctk.CTkButton(inner, text="", image=tab_delete_icon, width=24, height=24,
                                          fg_color=styles.MATRIX_DELETE_BUTTON_COLOR,
                                          hover_color=styles.MATRIX_DELETE_BUTTON_HOVER_COLOR,
                                          text_color=styles.DEFAULT_BUTTON_TEXT_COLOR,
                                          command=lambda idx=i: self._delete_tab_index(idx))
                close_btn.pack(side="right")
                # Drag handlers on outer area (not on close button)
                for wdg in (outer, inner, title):
                    wdg.bind("<ButtonPress-1>", lambda e, idx=i: self._on_tab_press(e, idx))
                    wdg.bind("<B1-Motion>", self._on_tab_motion)
                    wdg.bind("<ButtonRelease-1>", self._on_tab_release)
                # bottom edge is masked globally for all tabs (see below)
            else:
                # 未使用スロットはCANVAS背景色のまま（中身なし）
                pass
        # No mask needed when corners are square; ensure full-width adjustment
        # Ensure widths reflect current container width
        self._adjust_tabbar_widths()

    def _compute_slot_width(self) -> int:
        max_tabs = 5
        try:
            total_w = max(int(self.tabbar_frame.winfo_width()) - 40, 600)
        except Exception:
            total_w = 1000
        return max(140, total_w // max_tabs)

    def _adjust_tabbar_widths(self):
        slot_w = self._compute_slot_width()
        if self._tab_slot_width == slot_w:
            return
        self._tab_slot_width = slot_w
        for child in self.tabbar_frame.winfo_children():
            try:
                child.configure(width=slot_w)
            except Exception:
                pass

    # --- Tab drag & drop handlers ---
    def _compute_tab_drop_index(self, x_root: int) -> int:
        try:
            children = [c for c in self.tabbar_frame.winfo_children()][:len(self._tabs)]
            if not children:
                return 0
            centers = [c.winfo_rootx() + (c.winfo_width() // 2) for c in children]
            for i, cx in enumerate(centers):
                if x_root < cx:
                    return i
            return len(children) - 1
        except Exception:
            return 0

    def _on_tab_press(self, event, idx: int):
        self._tab_drag = { 'start_idx': idx, 'current_idx': idx, 'moved': False, 'start_x': event.x_root }

    def _on_tab_motion(self, event):
        if not self._tab_drag or self._tab_drag.get('start_idx') is None:
            return
        try:
            if abs(int(event.x_root) - int(self._tab_drag.get('start_x', event.x_root))) > 3:
                self._tab_drag['moved'] = True
        except Exception:
            self._tab_drag['moved'] = True
        self._tab_drag['current_idx'] = self._compute_tab_drop_index(event.x_root)

    def _on_tab_release(self, event):
        if not self._tab_drag or self._tab_drag.get('start_idx') is None:
            return
        start = int(self._tab_drag.get('start_idx'))
        moved = bool(self._tab_drag.get('moved'))
        drop = self._compute_tab_drop_index(event.x_root)
        # reset drag state
        self._tab_drag = { 'start_idx': None, 'current_idx': None, 'moved': False }
        if not moved:
            # simple click
            self._on_tab_clicked(start)
            return
        try:
            if not (0 <= start < len(self._tabs)):
                return
            drop = max(0, min(drop, len(self._tabs) - 1))
            if drop == start:
                return
            moving = self._tabs.pop(start)
            self._tabs.insert(drop, moving)
            # adjust active index
            if self._active_tab_index == start:
                self._active_tab_index = drop
            else:
                if start < self._active_tab_index <= drop:
                    self._active_tab_index -= 1
                elif drop <= self._active_tab_index < start:
                    self._active_tab_index += 1
            self._rebuild_tabs()
        except Exception:
            pass

    def _delete_tab_index(self, idx: int):
        if not (0 <= idx < len(self._tabs)):
            return
        try:
            tab_name = str(self._tabs[idx].get('name') or '')
        except Exception:
            tab_name = ''
        if not messagebox.askyesno(tr("common.delete_confirm_title"), tr("matrix.tab.delete_confirm", name=tab_name)):
            return
        del self._tabs[idx]
        if not self._tabs:
            self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {}, 'state': None}]
            self._active_tab_index = 0
        else:
            if self._active_tab_index >= len(self._tabs):
                self._active_tab_index = len(self._tabs) - 1
        # 再描画と内容の即時切替
        self._rebuild_tabs()
        # 高速切替: 既存フレームがあれば表示差し替え、なければ構築
        tab_frame = self._tabs[self._active_tab_index].get('frame')
        if tab_frame:
            self.scrollable_content_frame = tab_frame
            try:
                self.canvas.itemconfigure(self._window_id, window=self.scrollable_content_frame)
                self.after(1, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
            except Exception:
                pass
        else:
            self._update_ui()
        try:
            self._save_session()
        except Exception:
            pass

    def _rename_tab(self, idx: int):
        try:
            current = str(self._tabs[idx].get('name') or '')
        except Exception:
            current = ''
        new_name = self._prompt_text_input(tr("matrix.tab.new_name_title"), tr("matrix.tab.new_name_label"), default=current)
        if not new_name:
            return
        self._tabs[idx]['name'] = new_name
        self._render_tabbar()
        try:
            self._save_session()
        except Exception:
            pass

    def _load_session_or_default(self):
        import json
        sf = self._session_file()
        if sf.exists() and not getattr(self, '_start_with_default_only', False):
            try:
                data = json.loads(sf.read_text(encoding='utf-8'))
                tabs_data = data.get('tabs', [])
                self._tabs = []
                for t_data in tabs_data[:5]:
                    # Ensure the tab has a name before adding it
                    name = t_data.get('name')
                    if not name:
                        continue # Skip invalid tab data
                    prompts = t_data.get('prompts') or {}
                    state = t_data.get('state') or None
                    prompts_obj = self._deserialize_prompts(prompts)
                    self._tabs.append({'name': name, 'prompts_obj': prompts_obj, 'state': state})
                    logging.debug(f"DEBUG: _load_session_or_default - Loaded tab '{name}' with prompts_obj keys: {list(prompts_obj.keys())}")
                self._active_tab_index = int(data.get('active', 0)) if self._tabs else 0
                # Fallback: if loaded tabs are empty and we have initial prompts, seed default
                if (not self._tabs) or (len(self._tabs) == 1 and not self._tabs[0].get('prompts_obj') and getattr(self, '_initial_prompts', {})):
                    initial_prompts_filtered = {pid: p for pid, p in getattr(self, '_initial_prompts', {}).items() if getattr(p, 'include_in_matrix', False)}
                    self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in initial_prompts_filtered.items()}, 'state': None}]
                    self._active_tab_index = 0
                    logging.debug(f"DEBUG: _load_session_or_default - Fallback to default tab with filtered _initial_prompts keys: {list(initial_prompts_filtered.keys())}")
                if not self._tabs:
                    raise ValueError('empty')
                # Ensure UI tabs are constructed when loading from session
                self._rebuild_tabs()
                return
            except Exception as e:
                logging.error(f"ERROR: _load_session_or_default - Error loading session: {e}")
                pass
        # Default single tab from current prompts (store Prompt objects)
        base_prompts = getattr(self, '_initial_prompts', self.prompts)
        base_prompts_filtered = {pid: p for pid, p in base_prompts.items() if getattr(p, 'include_in_matrix', False)}
        self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in base_prompts_filtered.items()}, 'state': None}]
        self._active_tab_index = 0
        logging.debug(f"DEBUG: _load_session_or_default - Initializing with default tab from filtered base_prompts keys: {list(base_prompts_filtered.keys())}")
        # Build UI tabs
        self._rebuild_tabs()
        # After first build, disable the flag so future loads can open sessions if needed
        self._start_with_default_only = False

    def _save_session(self):
        """セッション保存: 現在開いている全タブ（最大5）とアクティブタブ、
        各タブのプロンプトセットとマトリクスのチェック/結果状態を `prompt_set/session.json` に保存します。
        次回起動時はこのセッションを復元します（プリセットとは別管理）。"""
        import json
        # Snapshot active before saving and serialize
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                self._tabs[self._active_tab_index]['state'] = self._snapshot_state()
        except Exception:
            pass
        tabs_payload = []
        for t in self._tabs[:5]:
            try:
                tabs_payload.append({
                    'name': t.get('name'),
                    'prompts': self._serialize_prompts(t.get('prompts_obj', {})),
                    'state': t.get('state')
                })
            except Exception:
                continue
        data = { 'tabs': tabs_payload, 'active': self._active_tab_index }
        try:
            self._session_file().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        except Exception:
            pass

    def _save_session_as(self, name: str):
        """セッションを任意の名前で保存（複数保存対応）。"""
        import json
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                self._tabs[self._active_tab_index]['state'] = self._snapshot_state()
        except Exception:
            pass
        tabs_payload = []
        for t in self._tabs[:5]:
            try:
                tabs_payload.append({
                    'name': t.get('name'),
                    'prompts': self._serialize_prompts(t.get('prompts_obj', {})),
                    'state': t.get('state')
                })
            except Exception:
                continue
        data = { 'tabs': tabs_payload, 'active': self._active_tab_index }
        self._session_file_for(name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')

    def _load_session_named(self, name: str):
        import json
        f = self._session_file_for(name)
        data = json.loads(f.read_text(encoding='utf-8'))
        tabs_data = data.get('tabs', [])
        self._tabs = []
        for t_data in tabs_data[:5]:
            name2 = t_data.get('name')
            if not name2:
                continue
            prompts = t_data.get('prompts') or {}
            state = t_data.get('state') or None
            prompts_obj = self._deserialize_prompts(prompts)
            self._tabs.append({'name': name2, 'prompts_obj': prompts_obj, 'state': state})
        self._active_tab_index = int(data.get('active', 0)) if self._tabs else 0
        if not self._tabs:
            self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {}, 'state': None}]
            self._active_tab_index = 0
        self._rebuild_tabs()

    def _add_prompt_set_tab(self):
        # Limit to 5 tabs
        if len(self._tabs) >= 5:
            CTkMessagebox(title=tr("matrix.tab.limit_title"), message=tr("matrix.tab.limit_message", max=5), icon="warning").wait_window()
            return
        # Show preset chooser
        preset = self._choose_preset_dialog()
        if preset is None:
            return
        name, prompts = preset
        # Snapshot current active tab BEFORE switching
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                self._tabs[self._active_tab_index]['state'] = self._snapshot_state()
        except Exception:
            pass
        if not name:
            # Default tab name
            name = tr("matrix.tab.auto_name_fmt", n=len(self._tabs)+1)
        # Ensure display name uniqueness to avoid user confusion
        base_name = name
        suffix = 1
        existing = {t.get('name') for t in self._tabs}
        while name in existing:
            suffix += 1
            name = f"{base_name} ({suffix})"
        # Use prompts from preset (deserialize to Prompt objects). Empty set -> {}
        prompts_obj = self._deserialize_prompts(prompts) if isinstance(prompts, dict) else {}
        self._tabs.append({'name': name, 'prompts_obj': prompts_obj, 'state': None})
        self._active_tab_index = len(self._tabs) - 1
        self._rebuild_tabs()
        self._update_ui()

    def _save_active_prompt_set(self):
        """プリセット保存: 現在のアクティブタブのプロンプトセットのみを
        名前付きJSON（`prompt_set/<name>.json`）として保存します。これは再利用用のテンプレートで、
        セッションのタブ状態・結果は含みません。"""
        # Ask for name
        name = self._prompt_text_input(tr("matrix.preset.save_title"), tr("matrix.preset.save_label"))
        if not name:
            return
        try:
            # Save prompts to file
            import json, time
            serialized = self._serialize_prompts(self.prompts)
            self._tabs[self._active_tab_index]['name'] = name
            self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
            file = self._prompt_set_dir() / f"{name}.json"
            file.write_text(json.dumps({'name': name, 'prompts': serialized}, ensure_ascii=False, indent=2), encoding='utf-8')
            # Update tab title
            self._rebuild_tabs()
            CTkMessagebox(title=tr("common.success"), message=tr("matrix.preset.saved"), icon="info").wait_window()
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=tr("matrix.preset.save_failed", details=str(e)), icon="cancel").wait_window()

    def _delete_active_tab(self):
        """アクティブなタブ（プロンプトセット）を削除する"""
        if not self._tabs:
            return
        # 確認ダイアログ
        try:
            tab_name = self._tabs[self._active_tab_index]['name'] if 0 <= self._active_tab_index < len(self._tabs) else ""
        except Exception:
            tab_name = ""
        if not messagebox.askyesno(tr("common.delete_confirm_title"), tr("matrix.tab.delete_confirm", name=tab_name)):
            return
        try:
            del self._tabs[self._active_tab_index]
        except Exception:
            return
        # 最低1つは保持：空ならデフォルト空タブを作成
        if not self._tabs:
            self._tabs = [{'name': tr('matrix.tab.default'), 'prompts_obj': {}, 'state': None}]
            self._active_tab_index = 0
        else:
            # 削除位置に応じてインデックス調整
            if self._active_tab_index >= len(self._tabs):
                self._active_tab_index = len(self._tabs) - 1
        # UI再構築と保存
        self._rebuild_tabs()
        # 新しいアクティブタブの内容を反映
        try:
            t = self._tabs[self._active_tab_index]
            prompts_obj = t.get('prompts_obj', {})
            self.prompts = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in prompts_obj.items()} if prompts_obj else {}
        except Exception:
            self.prompts = {}
        self.checkbox_states = []
        self.results = []
        self._full_results = []
        self._row_summaries = []
        self._col_summaries = []
        self._update_ui()
        try:
            self._save_session()
        except Exception:
            pass

    def _choose_preset_dialog(self) -> Optional[tuple[str, dict]]:
        # Build list of presets by mtime desc
        presets = []
        for p in self._prompt_set_dir().glob('*.json'):
            try:
                presets.append((p, p.stat().st_mtime))
            except Exception:
                pass
        presets.sort(key=lambda x: x[1], reverse=True)
        names = [pp[0].stem for pp in presets]
        # 固定選択肢: デフォルト, 空のセット
        names.insert(0, tr("matrix.preset.empty"))
        names.insert(0, tr("matrix.tab.default"))

        # Simple chooser using CTkOptionMenu in a small dialog
        dlg = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        dlg.title(tr("matrix.preset.add_from_title"))
        dlg.geometry("360x160")
        dlg.transient(self)
        dlg.grab_set()
        var = ctk.StringVar(value=tr("matrix.tab.default") if names else tr("matrix.tab.default"))
        ctk.CTkLabel(dlg, text=tr("matrix.preset.choose_label"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(padx=12, pady=(16,6))
        menu = ctk.CTkOptionMenu(dlg, values=names, variable=var)
        menu.pack(padx=12, pady=6)
        chosen: list = []
        def _ok():
            chosen.append(var.get())
            dlg.destroy()
        ctk.CTkButton(dlg, text=tr("common.ok"), command=_ok, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(pady=10)
        self.wait_window(dlg)
        if not chosen:
            return None
        sel = chosen[0]
        if sel == tr("matrix.preset.empty"):
            return (tr("matrix.tab.auto_name_fmt", n=len(self._tabs)+1), {})
        if sel == tr("matrix.tab.default"):
            # 初期プロンプト群のうち include_in_matrix=True のもの
            base_prompts = getattr(self, '_initial_prompts', self.prompts)
            default_prompts = {pid: p for pid, p in base_prompts.items() if getattr(p, 'include_in_matrix', False)}
            return (tr("matrix.tab.default"), self._serialize_prompts(default_prompts))
        # Load preset
        try:
            import json
            file = self._prompt_set_dir() / f"{sel}.json"
            data = json.loads(file.read_text(encoding='utf-8'))
            prompts = data.get('prompts', {})
            return (data.get('name', sel), prompts)
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=tr("matrix.preset.save_failed", details=str(e)), icon="cancel").wait_window()
            return None

    def _prompt_text_input(self, title: str, label: str, default: str = "") -> Optional[str]:
        dlg = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        dlg.title(title)
        dlg.geometry("360x140")
        dlg.transient(self)
        dlg.grab_set()
        ctk.CTkLabel(dlg, text=label, text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(padx=12, pady=(16,6))
        entry = ctk.CTkEntry(dlg, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        entry.pack(padx=12, pady=6, fill='x')
        try:
            entry.insert(0, default)
        except Exception:
            pass
        value: list[str] = []
        def _ok():
            value.append(entry.get().strip())
            dlg.destroy()
        ctk.CTkButton(dlg, text=tr("common.save"), command=_ok, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(pady=10)
        self.wait_window(dlg)
        return value[0] if value else None

    def _clear_active_set(self):
        # Clear prompts/results for active tab only; inputs remain shared
        try:
            self._tabs[self._active_tab_index]['prompts_obj'] = {}
            self.prompts = {}
        except Exception:
            self.prompts = {}
        self.checkbox_states = []
        self.results = []
        self._full_results = []
        self._row_summaries = []
        self._col_summaries = []
        self._result_textboxes = []
        self._cell_style = []
        self._update_ui()

    def _open_summary_settings(self):
        # 統一されたプロンプト管理画面を開く。入力グラブ等を解放し、アプリ画面を前面・フォーカスにする
        try:
            # クリップボード履歴ポップアップ等が掴んでいる場合は閉じてグラブを解放
            try:
                if hasattr(self, '_history_popup') and self._history_popup and self._history_popup.winfo_exists():
                    try:
                        if self._history_popup.grab_current() == str(self._history_popup):
                            self._history_popup.grab_release()
                    except Exception:
                        pass
                    self._history_popup.destroy()
                    self._history_popup = None
            except Exception:
                pass

            # どこかでgrabされていれば強制解放
            try:
                cur = self.tk.call('grab', 'current')
                if cur:
                    try:
                        self.nametowidget(cur).grab_release()
                    except Exception:
                        self.grab_release()
            except Exception:
                pass

            # このウィンドウが入力グラブしている場合は解放
            try:
                if self.grab_current() == str(self):
                    self.grab_release()
            except Exception:
                pass

            # 念のため最前面やトランジェントを解除し、このウィンドウを一旦隠す
            try:
                self.transient(None)
            except Exception:
                pass
            try:
                self.attributes("-topmost", False)
            except Exception:
                pass
            # 表示は維持しつつ操作を無効化（可能なら）
            try:
                self.attributes('-disabled', True)
            except Exception:
                pass

            # プロンプト管理を前面に
            if hasattr(self.agent, '_show_main_window'):
                self.agent._show_main_window()
            # マネージャが閉じられたら再度このウィンドウの操作を有効化
            try:
                self._watch_manager_to_reenable()
            except Exception:
                pass
        except Exception:
            pass

    def _open_set_manager(self):
        dlg = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        dlg.title(tr("matrix.set.manager_title"))
        dlg.geometry("520x240")
        dlg.transient(self)
        dlg.grab_set()
        ctk.CTkLabel(dlg, text=tr("matrix.set.save_delete"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(padx=12, pady=(12,6))
        # Save-as row (unified with session manager style)
        row1 = ctk.CTkFrame(dlg, fg_color="transparent")
        row1.pack(fill='x', padx=12, pady=6)
        name_entry = ctk.CTkEntry(row1, placeholder_text=tr("matrix.set.placeholder"), fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        name_entry.pack(side='left', fill='x', expand=True)
        def do_save():
            name = name_entry.get().strip()
            if not name:
                CTkMessagebox(title=tr("common.warning"), message=tr("matrix.set.name_required"), icon="warning").wait_window()
                return
            try:
                import json
                serialized = self._serialize_prompts(self.prompts)
                # 更新: タブ表示名とプロンプトオブジェクト
                self._tabs[self._active_tab_index]['name'] = name
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                file = self._prompt_set_dir() / f"{name}.json"
                file.write_text(json.dumps({'name': name, 'prompts': serialized}, ensure_ascii=False, indent=2), encoding='utf-8')
                self._rebuild_tabs()
                CTkMessagebox(title=tr("common.success"), message=tr("matrix.set.saved"), icon="info").wait_window()
                dlg.destroy()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.set.save_failed", details=str(e)), icon="cancel").wait_window()
        ctk.CTkButton(row1, text=tr("common.save"), command=do_save, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(side='left', padx=(6,0))
        # 削除UI
        from pathlib import Path as _Path
        options = []
        try:
            options = [p.stem for p in self._prompt_set_dir().glob('*.json')]
        except Exception:
            options = []
        var = ctk.StringVar(value=options[0] if options else "")
        row2 = ctk.CTkFrame(dlg, fg_color="transparent")
        row2.pack(fill='x', padx=12, pady=6)
        ctk.CTkLabel(row2, text=tr("matrix.set.delete_target"), width=80, anchor='w', text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(side='left')
        menu = ctk.CTkOptionMenu(row2, values=options or [""], variable=var)
        menu.pack(side='left', padx=(6,6))
        def do_delete():
            name = var.get().strip()
            if not name:
                return
            try:
                f = self._prompt_set_dir() / f"{name}.json"
                if f.exists():
                    f.unlink()
                    CTkMessagebox(title=tr("common.success"), message=tr("matrix.set.deleted"), icon="info").wait_window()
                    dlg.destroy()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.set.delete_failed", details=str(e)), icon="cancel").wait_window()
        ctk.CTkButton(row2, text=tr("common.delete"), command=do_delete, fg_color=styles.DELETE_BUTTON_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR, hover_color=styles.DELETE_BUTTON_HOVER_COLOR).pack(side='left', padx=(6,0))

    def _open_session_manager(self):
        dlg = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        dlg.title(tr("matrix.session.manager_title"))
        dlg.geometry("520x260")
        dlg.transient(self)
        dlg.grab_set()
        ctk.CTkLabel(dlg, text=tr("matrix.session.save_load_delete"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(padx=12, pady=(12,6))
        # Save-as row
        row1 = ctk.CTkFrame(dlg, fg_color="transparent")
        row1.pack(fill='x', padx=12, pady=6)
        name_entry = ctk.CTkEntry(row1, placeholder_text=tr("matrix.session.placeholder"), fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        name_entry.pack(side='left', fill='x', expand=True)
        def do_save_as():
            name = name_entry.get().strip()
            if not name:
                CTkMessagebox(title=tr("common.warning"), message=tr("matrix.session.name_required"), icon="warning").wait_window()
                return
            try:
                self._save_session_as(name)
                CTkMessagebox(title=tr("common.success"), message=tr("matrix.session.saved"), icon="info").wait_window()
                dlg.destroy()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.session.save_failed", details=str(e)), icon="cancel").wait_window()
        ctk.CTkButton(row1, text=tr("common.save"), command=do_save_as, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(side='left', padx=(6,0))

        # Existing sessions
        row2 = ctk.CTkFrame(dlg, fg_color="transparent")
        row2.pack(fill='x', padx=12, pady=6)
        try:
            sess_files = sorted(self._sessions_dir().glob('*.json'), key=lambda p: p.stat().st_mtime, reverse=True)
            sess_names = [p.stem for p in sess_files]
        except Exception:
            sess_names = []
        sess_var = ctk.StringVar(value=(sess_names[0] if sess_names else ""))
        ctk.CTkLabel(row2, text=tr("matrix.session.label"), width=80, anchor='w', text_color=styles.HISTORY_ITEM_TEXT_COLOR).pack(side='left')
        sess_menu = ctk.CTkOptionMenu(row2, values=sess_names or [""], variable=sess_var)
        sess_menu.pack(side='left', padx=(6,6))
        def do_load_named():
            name = sess_var.get().strip()
            if not name:
                return
            try:
                self._load_session_named(name)
                CTkMessagebox(title=tr("common.success"), message=tr("matrix.session.loaded"), icon="info").wait_window()
                dlg.destroy()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.session.load_failed", details=str(e)), icon="cancel").wait_window()
        def do_delete_named():
            name = sess_var.get().strip()
            if not name:
                return
            try:
                f = self._session_file_for(name)
                if f.exists():
                    f.unlink()
                    CTkMessagebox(title=tr("common.success"), message=tr("matrix.session.deleted"), icon="info").wait_window()
                    dlg.destroy()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.session.delete_failed", details=str(e)), icon="cancel").wait_window()
        ctk.CTkButton(row2, text=tr("common.load"), command=do_load_named, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(side='left', padx=(6,0))
        ctk.CTkButton(row2, text=tr("common.delete"), command=do_delete_named, fg_color=styles.DELETE_BUTTON_COLOR, hover_color=styles.DELETE_BUTTON_HOVER_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).pack(side='left', padx=(6,0))

    def _watch_manager_to_reenable(self):
        try:
            # 管理画面（root）が非表示または最小化されたら再有効化
            if not self.parent_app.winfo_viewable():
                try:
                    self.attributes('-disabled', False)
                except Exception:
                    pass
                try:
                    self.lift()
                    self.focus_force()
                except Exception:
                    pass
                return
        except Exception:
            return
        # まだ開いている場合は定期的に監視
        try:
            self.after(300, self._watch_manager_to_reenable)
        except Exception:
            pass

    def _create_main_grid_frame(self):
        """マトリクスグリッドを配置するためのスクロール可能フレームを作成する"""
        self.canvas_frame = ctk.CTkFrame(self, fg_color=styles.MATRIX_CANVAS_BACKGROUND_COLOR)
        self.canvas_frame.pack(fill="both", expand=True, padx=10, pady=(0,5))

        current_fg_color = self.canvas_frame.cget("fg_color")
        appearance_mode_index = 0 if ctk.get_appearance_mode() == "Light" else 1
        canvas_bg_color = current_fg_color[appearance_mode_index] if isinstance(current_fg_color, tuple) else current_fg_color
        self.canvas = ctk.CTkCanvas(self.canvas_frame, highlightthickness=0, bg=canvas_bg_color)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.v_scrollbar = ctk.CTkScrollbar(self.canvas_frame, orientation="vertical", command=self.canvas.yview)
        self.v_scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.h_scrollbar = ctk.CTkScrollbar(self, orientation="horizontal", command=self.canvas.xview)
        self.h_scrollbar.pack(fill="x", side="bottom", padx=10, pady=(0, 10))
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        # タブごとの内容フレーム（キャッシュ）
        self.scrollable_content_frame = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self._window_id = self.canvas.create_window((0, 0), window=self.scrollable_content_frame, anchor="nw")
        # タブにフレームを格納し、切替時はフレームを差し替える
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['frame'] = self.scrollable_content_frame
        except Exception:
            pass

        self.scrollable_content_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.scrollable_content_frame.grid_columnconfigure(0, weight=0)
        self.scrollable_content_frame.grid_rowconfigure(0, weight=0)

        self.run_button_frame = ctk.CTkFrame(self, fg_color=styles.MATRIX_TOP_BG_COLOR)
        self.run_button_frame.pack(fill="x", padx=10, pady=10, side="bottom")
        self.run_button_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        # Order: 実行, フロー実行, 行まとめ, 列まとめ, 行列まとめ, エクセル出力
        ctk.CTkButton(self.run_button_frame, text=tr("matrix.run"), command=self._run_batch_processing, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.flow_run_button = ctk.CTkButton(self.run_button_frame, text=tr("matrix.run_flow"), command=self._run_flow_processing, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.flow_run_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.summarize_row_button = ctk.CTkButton(self.run_button_frame, text=tr("matrix.run_row_summary"), command=self._summarize_rows, state="disabled", fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.summarize_row_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        self.summarize_col_button = ctk.CTkButton(self.run_button_frame, text=tr("matrix.run_col_summary"), command=self._summarize_columns, state="disabled", fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.summarize_col_button.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        self.summarize_matrix_button = ctk.CTkButton(self.run_button_frame, text=tr("matrix.matrix_summary"), command=self._summarize_matrix, state="disabled", fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.summarize_matrix_button.grid(row=0, column=4, padx=5, pady=5, sticky="ew")

        self.export_excel_button = ctk.CTkButton(self.run_button_frame, text="Excel", command=self._export_to_excel, state="disabled", fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.export_excel_button.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.tooltip_window = None

    def _on_canvas_configure(self, event):
        self.canvas.coords(self._window_id, 0, 0)
        self._on_frame_configure(event)

    def _update_input_row_display(self, row_idx: int):
        for widget in self.scrollable_content_frame.grid_slaves(row=row_idx + 1, column=0):
            widget.destroy()
        self._add_input_row_widgets(row_idx, self.input_data[row_idx])

    def _show_tooltip(self, text):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        x = self.winfo_pointerx() + 25
        y = self.winfo_pointery() + 20
        self.tooltip_window = tk.Toplevel(self)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"{x}+{y}")
        label = tk.Label(self.tooltip_window, text=text, justify='left', background="#ffffe0", relief='solid', borderwidth=1, font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def _hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

    def _update_row_summary_column(self):
        """行まとめ列の表示/非表示、および内容の更新を行う"""
        num_prompts = len(self.prompts)
        summary_col_idx = num_prompts + 1

        for widget in self.scrollable_content_frame.grid_slaves(column=summary_col_idx):
            widget.destroy()

        if self._row_summaries:
            self.scrollable_content_frame.grid_columnconfigure(summary_col_idx, weight=0, minsize=styles.MATRIX_CELL_WIDTH)
            summary_header_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_HEADER_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
            summary_header_frame.grid_propagate(False)
            ctk.CTkLabel(summary_header_frame, text=tr("matrix.row_summary_header"), font=styles.MATRIX_FONT_BOLD).pack(fill="x", padx=2, pady=2)
            summary_header_frame.grid(row=0, column=summary_col_idx, padx=5, pady=5, sticky="nsew")

            for r_idx, summary_var in enumerate(self._row_summaries):
                summary_cell_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_CELL_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
                summary_cell_frame.grid_propagate(False)
                summary_cell_frame.grid(row=r_idx + 1, column=summary_col_idx, padx=5, pady=5, sticky="nsew")
                summary_textbox = ctk.CTkTextbox(summary_cell_frame, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                summary_textbox.insert("1.0", summary_var.get())
                summary_textbox.configure(state="disabled")
                summary_textbox.pack(fill="both", expand=True)
                summary_textbox.bind("<Button-1>", lambda e, r=r_idx: self._show_full_row_summary_popup(r))
                summary_textbox.bind("<Enter>", lambda e: e.widget.configure(cursor="hand2"))
                summary_textbox.bind("<Leave>", lambda e: e.widget.configure(cursor=""))
                summary_var.trace_add("write", lambda name, index, mode, sv=summary_var, tb=summary_textbox: self._update_textbox_from_stringvar(sv, tb))
        else:
            self.scrollable_content_frame.grid_columnconfigure(summary_col_idx, weight=0, minsize=0)

    def _update_column_summary_row(self):
        """列まとめ行の表示/非表示、および内容の更新を行う"""
        num_inputs = len(self.input_data)
        summary_row_idx = num_inputs + 1

        for widget in self.scrollable_content_frame.grid_slaves(row=summary_row_idx):
            widget.destroy()

        if self._col_summaries:
            self.scrollable_content_frame.grid_rowconfigure(summary_row_idx, weight=0, minsize=styles.MATRIX_RESULT_CELL_HEIGHT)
            summary_header_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_HEADER_BORDER_COLOR, height=styles.MATRIX_RESULT_CELL_HEIGHT)
            summary_header_frame.grid_propagate(False)
            ctk.CTkLabel(summary_header_frame, text=tr("matrix.col_summary_header"), font=styles.MATRIX_FONT_BOLD).pack(fill="x", padx=2, pady=2)
            summary_header_frame.grid(row=summary_row_idx, column=0, padx=5, pady=5, sticky="nsew")

            for c_idx, summary_var in enumerate(self._col_summaries):
                summary_cell_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_CELL_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
                summary_cell_frame.grid(row=summary_row_idx, column=c_idx + 1, padx=5, pady=5, sticky="nsew")
                summary_cell_frame.grid_propagate(False)
                
                summary_cell_frame.grid_rowconfigure(0, weight=1)
                summary_cell_frame.grid_columnconfigure(0, weight=1)

                summary_textbox = ctk.CTkTextbox(summary_cell_frame, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                summary_textbox.insert("1.0", summary_var.get())
                summary_textbox.configure(state="disabled")
                summary_textbox.grid(row=0, column=0, sticky="nsew")
                summary_textbox.bind("<Button-1>", lambda e, c=c_idx: self._show_full_col_summary_popup(c))
                summary_textbox.bind("<Enter>", lambda e: e.widget.configure(cursor="hand2"))
                summary_textbox.bind("<Leave>", lambda e: e.widget.configure(cursor=""))
                summary_var.trace_add("write", lambda *args, sv=summary_var, tb=summary_textbox: self._update_textbox_from_stringvar(sv, tb))

                sizer = SizerGrip(summary_cell_frame)
                sizer.grid(row=1, column=1, sticky="se")
        else:
            self.scrollable_content_frame.grid_rowconfigure(summary_row_idx, weight=0, minsize=0)

    def _update_ui(self):
        self.scrollable_content_frame.update_idletasks()
        if self._is_closing or not self.winfo_exists():
            return

        for widget in self.scrollable_content_frame.winfo_children():
            widget.destroy()

        num_prompts = len(self.prompts)
        num_inputs = len(self.input_data)
        
        total_cols = 1 + num_prompts + (1 if self._row_summaries else 0)
        for i in range(total_cols):
            self.scrollable_content_frame.grid_columnconfigure(i, weight=0)
        self.scrollable_content_frame.grid_columnconfigure(0, weight=0)
        for i in range(1, num_prompts + 1):
            self.scrollable_content_frame.grid_columnconfigure(i, weight=0, minsize=styles.MATRIX_CELL_WIDTH)
        if self._row_summaries:
            self.scrollable_content_frame.grid_columnconfigure(num_prompts + 1, weight=0, minsize=styles.MATRIX_CELL_WIDTH)

        total_rows = 1 + num_inputs + (1 if self._col_summaries else 0)
        for i in range(total_rows):
            self.scrollable_content_frame.grid_rowconfigure(i, weight=0)
        self.scrollable_content_frame.grid_rowconfigure(0, weight=0, minsize=styles.MATRIX_RESULT_CELL_HEIGHT)
        for i in range(1, num_inputs + 1):
            self.scrollable_content_frame.grid_rowconfigure(i, weight=0, minsize=styles.MATRIX_RESULT_CELL_HEIGHT)
        if self._col_summaries:
            self.scrollable_content_frame.grid_rowconfigure(num_inputs + 1, weight=0, minsize=styles.MATRIX_RESULT_CELL_HEIGHT)

        ctk.CTkLabel(self.scrollable_content_frame, text="").grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        # Reset header frames list before rebuilding
        self._col_header_frames = []

        for col_idx, (prompt_id, prompt_config) in enumerate(self.prompts.items()):
            self._add_prompt_header_widgets(col_idx, prompt_id, prompt_config)

        for row_idx, input_item in enumerate(self.input_data):
            self._add_input_row_widgets(row_idx, input_item)

        self._update_row_summary_column()
        self._update_column_summary_row()
        # アクティブタブに現在の変数参照を保存（高速切替用）
        try:
            if 0 <= self._active_tab_index < len(self._tabs):
                self._tabs[self._active_tab_index]['vars'] = {
                    'checkbox_states': self.checkbox_states,
                    'results': self.results,
                    '_full_results': self._full_results,
                    '_row_summaries': self._row_summaries,
                    '_col_summaries': self._col_summaries,
                    '_result_textboxes': self._result_textboxes,
                    '_cell_style': self._cell_style
                }
                self._tabs[self._active_tab_index]['frame'] = self.scrollable_content_frame
        except Exception:
            pass

    def _add_prompt_header_widgets(self, col_idx: int, prompt_id: str, prompt_config: Prompt):
        header_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_HEADER_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
        header_frame.grid(row=0, column=col_idx + 1, padx=5, pady=5, sticky="nsew")
        header_frame.grid_propagate(False)
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_rowconfigure(1, weight=1)

        # Track for drag-and-drop reordering
        setattr(header_frame, "_prompt_id", prompt_id)
        while len(self._col_header_frames) <= col_idx:
            self._col_header_frames.append(None)  # type: ignore
        self._col_header_frames[col_idx] = header_frame

        # Bind drag events to the header frame
        header_frame.bind("<ButtonPress-1>", self._on_col_press)
        header_frame.bind("<B1-Motion>", self._on_col_motion)
        header_frame.bind("<ButtonRelease-1>", self._on_col_release)

        col_header_inner_frame = ctk.CTkFrame(header_frame)
        col_header_inner_frame.pack(fill="x", padx=2, pady=2)
        col_header_inner_frame.grid_columnconfigure(0, weight=1)
        col_header_inner_frame.grid_columnconfigure(1, weight=0)

        col_letter = chr(ord('A') + col_idx)
        col_num_label = ctk.CTkLabel(col_header_inner_frame, text=f"{col_letter}", font=styles.MATRIX_FONT_BOLD, anchor="center")
        col_num_label.grid(row=0, column=0, sticky="ew")
        # Bind drag events to child widgets as well so drags starting on them work
        col_num_label.bind("<ButtonPress-1>", self._on_col_press)
        col_num_label.bind("<B1-Motion>", self._on_col_motion)
        col_num_label.bind("<ButtonRelease-1>", self._on_col_release)

        delete_col_icon = None
        try:
            icon_path = Path(DELETE_ICON_FILE)
            if icon_path.exists():
                from PIL import Image
                icon_img = Image.open(icon_path)
                size = (16, 16)
                icon_img.thumbnail(size)
                delete_col_icon = ctk.CTkImage(light_image=icon_img, dark_image=icon_img, size=size)
        except Exception:
            delete_col_icon = None

        delete_col_button = ctk.CTkButton(col_header_inner_frame, text="" if delete_col_icon else tr("common.delete"), image=delete_col_icon, width=24, height=24, fg_color=styles.MATRIX_DELETE_BUTTON_COLOR, hover_color=styles.MATRIX_DELETE_BUTTON_HOVER_COLOR, command=lambda c=col_idx: self._delete_column(c))
        delete_col_button.grid(row=0, column=1, padx=(5, 0), sticky="e")
        delete_col_button.bind("<ButtonPress-1>", self._on_col_press)
        delete_col_button.bind("<B1-Motion>", self._on_col_motion)
        delete_col_button.bind("<ButtonRelease-1>", self._on_col_release)
        
        prompt_name_entry = ctk.CTkEntry(header_frame, placeholder_text=tr("prompt.header_name"))
        prompt_name_entry.insert(0, prompt_config.name)
        prompt_name_entry.configure(state="readonly")
        # Open editor only on click release when not dragging
        prompt_name_entry.bind("<ButtonRelease-1>", lambda e, p_id=prompt_id: self._open_editor_if_not_drag(p_id))
        # Also allow drag start from the entry area
        prompt_name_entry.bind("<ButtonPress-1>", self._on_col_press)
        prompt_name_entry.bind("<B1-Motion>", self._on_col_motion)
        prompt_name_entry.bind("<ButtonRelease-1>", self._on_col_release)
        prompt_name_entry.pack(fill="x", padx=2, pady=2)

        system_prompt_textbox = ctk.CTkTextbox(header_frame, height=styles.MATRIX_RESULT_CELL_HEIGHT, wrap="word")
        system_prompt_textbox.insert("1.0", prompt_config.system_prompt)
        try:
            system_prompt_textbox.tag_configure("left", justify="left")
            system_prompt_textbox.tag_add("left", "1.0", "end")
        except Exception:
            pass
        system_prompt_textbox.configure(state="disabled")
        system_prompt_textbox.bind("<ButtonRelease-1>", lambda e, p_id=prompt_id: self._open_editor_if_not_drag(p_id))
        system_prompt_textbox.bind("<ButtonPress-1>", self._on_col_press)
        system_prompt_textbox.bind("<B1-Motion>", self._on_col_motion)
        system_prompt_textbox.bind("<ButtonRelease-1>", self._on_col_release)
        system_prompt_textbox.pack(fill="both", expand=True, padx=2, pady=2)

    # --- Column drag-and-drop handlers (similar to prompt manager row DnD) ---
    def _on_col_press(self, event):
        try:
            idx = self._compute_col_drop_index(event.x_root)
            if not (0 <= idx < len(self._col_header_frames)):
                return
            target_frame = self._col_header_frames[idx]
            if target_frame is None:
                return
            self._col_drag_data = {"frame": target_frame, "index": idx, "current_index": idx, "moved": False, "start_x": event.x_root}
            # Highlight dragged column header
            if self._col_drag_active_frame and self._col_drag_active_frame.winfo_exists():
                try:
                    self._col_drag_active_frame.configure(fg_color=styles.HISTORY_ITEM_FG_COLOR)
                except Exception:
                    pass
            self._col_drag_active_frame = target_frame
            try:
                self._col_drag_active_frame.configure(fg_color=styles.DRAG_ACTIVE_ROW_COLOR)
            except Exception:
                pass
        except Exception:
            return

    def _on_col_motion(self, event):
        if not self._col_drag_data:
            return
        try:
            # Mark as moved when exceeding small threshold
            try:
                if abs(int(event.x_root) - int(self._col_drag_data.get("start_x", event.x_root))) > 3:
                    self._col_drag_data["moved"] = True
            except Exception:
                self._col_drag_data["moved"] = True
            new_index = self._compute_col_drop_index(event.x_root)
            current_index = self._col_drag_data.get("current_index", 0)
            if 0 <= new_index < len(self._col_header_frames):
                # Draw a white boundary indicator line between columns
                self._draw_col_drop_indicator(event.x_root)
            if new_index != current_index and 0 <= new_index < len(self._col_header_frames):
                self._col_drag_data["current_index"] = new_index
        except Exception:
            return

    def _on_col_release(self, event):
        if not self._col_drag_data:
            return
        try:
            # Compute final index and reorder prompt mapping
            drop_index = self._compute_col_drop_index(event.x_root)
            start_index = self._col_drag_data.get("index", 0)
            if 0 <= drop_index < len(self._col_header_frames) and drop_index != start_index:
                # Compute old/new id order
                old_ids = list(self.prompts.keys())
                moving_id = old_ids.pop(start_index)
                new_ids = old_ids.copy()
                new_ids.insert(drop_index, moving_id)

                # Reorder per-column state arrays to match new order
                try:
                    id_to_old_index = {pid: i for i, pid in enumerate(list(self.prompts.keys()))}
                    # checkbox, results, full_results
                    for r in range(len(self.input_data)):
                        if r < len(self.checkbox_states):
                            old_row = self.checkbox_states[r]
                            new_row = []
                            for pid in new_ids:
                                oi = id_to_old_index.get(pid, None)
                                if oi is None or oi >= len(old_row):
                                    new_row.append(ctk.BooleanVar(value=False))
                                else:
                                    new_row.append(old_row[oi])
                            self.checkbox_states[r] = new_row
                        if r < len(self.results):
                            old_row_r = self.results[r]
                            new_row_r = []
                            for pid in new_ids:
                                oi = id_to_old_index.get(pid, None)
                                if oi is None or oi >= len(old_row_r):
                                    new_row_r.append(ctk.StringVar(value=""))
                                else:
                                    new_row_r.append(old_row_r[oi])
                            self.results[r] = new_row_r
                        if r < len(self._full_results):
                            old_row_f = self._full_results[r]
                            new_row_f = []
                            for pid in new_ids:
                                oi = id_to_old_index.get(pid, None)
                                if oi is None or oi >= len(old_row_f):
                                    new_row_f.append("")
                                else:
                                    new_row_f.append(old_row_f[oi])
                            self._full_results[r] = new_row_f
                    # per-column summaries
                    if self._col_summaries:
                        old_cols = self._col_summaries
                        new_cols = []
                        for pid in new_ids:
                            oi = id_to_old_index.get(pid, None)
                            if oi is None or oi >= len(old_cols):
                                new_cols.append(ctk.StringVar(value=""))
                            else:
                                new_cols.append(old_cols[oi])
                        self._col_summaries = new_cols
                    # widths
                    if self._column_widths:
                        old_w = self._column_widths
                        # Ensure length equals number of prompts
                        while len(old_w) < len(new_ids):
                            old_w.append(styles.MATRIX_CELL_WIDTH)
                        new_w = []
                        for pid in new_ids:
                            oi = id_to_old_index.get(pid, 0)
                            new_w.append(old_w[oi])
                        self._column_widths = new_w
                except Exception:
                    pass

                # Rebuild dict in new order
                new_prompts = {pid: self.prompts[pid] for pid in new_ids}
                self.prompts = new_prompts
                # Rebuild UI to reflect new column order
                self._update_ui()
            # Reset highlights
            for fr in self._col_header_frames:
                if fr is None or not fr.winfo_exists():
                    continue
                try:
                    fr.configure(fg_color=styles.HISTORY_ITEM_FG_COLOR)
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self._col_drag_data = {}
            self._col_drag_active_frame = None
            # Remove drop indicator line
            # Remove drop indicator (line/frame)
            try:
                if self._col_drop_line_id is not None:
                    self.canvas.delete(self._col_drop_line_id)
            except Exception:
                pass
            self._col_drop_line_id = None
            try:
                if self._col_drop_indicator_widget is not None and self._col_drop_indicator_widget.winfo_exists():
                    self._col_drop_indicator_widget.destroy()
            except Exception:
                pass
            self._col_drop_indicator_widget = None

    def _compute_col_drop_index(self, x_root: int) -> int:
        frames = [f for f in self._col_header_frames if f is not None]
        if not frames:
            return 0
        mids = []
        for f in frames:
            try:
                left = f.winfo_rootx()
                w = f.winfo_width() or styles.MATRIX_CELL_WIDTH
                mids.append(left + w / 2)
            except Exception:
                mids.append(0)
        if x_root <= mids[0]:
            return 0
        if x_root >= mids[-1]:
            return len(frames) - 1
        best = 0
        best_d = float('inf')
        for i, m in enumerate(mids):
            d = abs(x_root - m)
            if d < best_d:
                best_d = d
                best = i
        return best

    def _draw_col_drop_indicator(self, x_root: int) -> None:
        try:
            # Compute boundary positions between header frames
            frames = [f for f in self._col_header_frames if f is not None and f.winfo_exists()]
            if not frames:
                return
            lefts = []
            rights = []
            for f in frames:
                lx = f.winfo_rootx()
                w = f.winfo_width() or styles.MATRIX_CELL_WIDTH
                lefts.append(lx)
                rights.append(lx + w)
            boundaries: List[int] = []
            boundaries.append(lefts[0])
            for i in range(1, len(frames)):
                boundaries.append(int((rights[i-1] + lefts[i]) / 2))
            boundaries.append(rights[-1])
            # Find nearest boundary to pointer
            bx = min(boundaries, key=lambda b: abs(b - x_root))
            # Convert to local coordinate within the scrollable content frame
            local_x = bx - self.scrollable_content_frame.winfo_rootx()
            # Create/update overlay frame as vertical indicator
            h = max(2, self.scrollable_content_frame.winfo_height())
            if self._col_drop_indicator_widget is None or not self._col_drop_indicator_widget.winfo_exists():
                self._col_drop_indicator_widget = tk.Frame(self.scrollable_content_frame, bg="#FFFFFF", width=2, height=h)
                self._col_drop_indicator_widget.place(x=local_x, y=0, width=2, relheight=1.0)
            else:
                self._col_drop_indicator_widget.place_configure(x=local_x, y=0)
        except Exception:
            pass

    def _open_editor_if_not_drag(self, prompt_id: str):
        # Only open editor when no drag occurred recently
        try:
            if self._col_drag_data and self._col_drag_data.get("moved"):
                return
        except Exception:
            pass
        self._open_prompt_editor(prompt_id)

    def _add_input_row_widgets(self, row_idx: int, input_item: Dict[str, Any]):
        while len(self._full_results) <= row_idx:
            self._full_results.append([])
        input_cell_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_CELL_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
        input_cell_frame.grid(row=row_idx + 1, column=0, padx=5, pady=5, sticky="nsew")
        input_cell_frame.grid_propagate(False)
        input_cell_frame.grid_columnconfigure(2, weight=1)
        while len(self._input_row_frames) <= row_idx:
            self._input_row_frames.append(None)
        self._input_row_frames[row_idx] = input_cell_frame

        row_header_frame = ctk.CTkFrame(input_cell_frame)
        row_header_frame.grid(row=0, column=0, padx=(5, 2), pady=2, sticky="w")
        row_header_frame.grid_columnconfigure(0, weight=1)

        row_num_label = ctk.CTkLabel(row_header_frame, text=f"{row_idx + 1}", font=styles.MATRIX_FONT_BOLD)
        row_num_label.grid(row=0, column=0, sticky="w")

        delete_row_icon = None
        try:
            icon_path = Path(DELETE_ICON_FILE)
            if icon_path.exists():
                from PIL import Image
                icon_img = Image.open(icon_path)
                size = (16, 16)
                icon_img.thumbnail(size)
                delete_row_icon = ctk.CTkImage(light_image=icon_img, dark_image=icon_img, size=size)
        except Exception:
            delete_row_icon = None

        delete_row_button = ctk.CTkButton(row_header_frame, text="" if delete_row_icon else tr("common.delete"), image=delete_row_icon, width=24, height=24, fg_color=styles.MATRIX_DELETE_BUTTON_COLOR, hover_color=styles.MATRIX_DELETE_BUTTON_HOVER_COLOR, command=lambda r=row_idx: self._delete_row(r))
        delete_row_button.grid(row=0, column=1, padx=(5, 0), sticky="e")

        input_label = ctk.CTkLabel(input_cell_frame, text=tr("action.input").rstrip(':'), font=styles.MATRIX_FONT_BOLD, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        input_label.grid(row=0, column=1, padx=(5, 2), pady=2, sticky="w")

        if input_item["type"] == "text":
            input_entry = ctk.CTkEntry(input_cell_frame, placeholder_text=tr("matrix.input_placeholder", n=row_idx + 1), fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
            input_entry.insert(0, input_item["data"])
            input_entry.configure(state="readonly")
            input_entry.bind("<Button-1>", lambda event, r=row_idx: self._open_history_edit_dialog(r))
            input_entry.grid(row=0, column=2, padx=2, pady=2, sticky="ew")
        elif input_item["type"] in ("image", "image_compressed"):
            try:
                raw_bytes = base64.b64decode(input_item["data"])
                if input_item["type"] == "image_compressed":
                    import zlib
                    raw_bytes = zlib.decompress(raw_bytes)
                image = Image.open(BytesIO(raw_bytes))
                image.thumbnail(styles.MATRIX_IMAGE_THUMBNAIL_SIZE)
                ctk_image = ctk.CTkImage(light_image=image, dark_image=image, size=styles.MATRIX_IMAGE_THUMBNAIL_SIZE)
                image_label = ctk.CTkLabel(input_cell_frame, image=ctk_image, text="", text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                image_label.grid(row=0, column=2, padx=2, pady=2, sticky="w")
                image_label.bind("<Button-1>", lambda event=None, r=row_idx: self._show_image_preview(r))
            except Exception as e:
                error_label = ctk.CTkLabel(input_cell_frame, text=tr("matrix.image_error"), text_color=styles.NOTIFICATION_COLORS["error"])
                error_label.grid(row=0, column=2, padx=2, pady=2, sticky="w")
                print(f"DEBUG: 画像表示エラー (行{row_idx}): {str(e)}")
        elif input_item["type"] == "file":
            file_path = Path(input_item["data"])
            file_label = ctk.CTkLabel(input_cell_frame, text=file_path.name, text_color=styles.HISTORY_ITEM_TEXT_COLOR, anchor="w")
            file_label.grid(row=0, column=2, padx=2, pady=2, sticky="ew")
            file_label.bind("<Enter>", lambda event=None, p=str(file_path): self._show_tooltip(p))
            file_label.bind("<Leave>", lambda event=None: self._hide_tooltip())

        attach_button = ctk.CTkButton(input_cell_frame, text=tr("action.attach"), width=50, fg_color=styles.FILE_ATTACH_BUTTON_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR, command=lambda r=row_idx: self._select_input_source(r))
        attach_button.grid(row=0, column=3, padx=2, pady=2, sticky="e")

        history_button = ctk.CTkButton(input_cell_frame, text=tr("history.button"), width=50, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR, command=lambda r=row_idx: self._show_clipboard_history_popup(r))
        history_button.grid(row=0, column=4, padx=2, pady=2, sticky="e")
        input_cell_frame.grid_columnconfigure(4, weight=0)

        for col_idx in range(len(self.prompts)):
            self._create_result_cell(row_idx, col_idx)

    def _create_result_cell(self, row_idx: int, col_idx: int):
        """指定された行と列に結果表示用のセルウィジェットを作成する"""
        while len(self.checkbox_states) <= row_idx:
            self.checkbox_states.append([])
        while len(self.checkbox_states[row_idx]) <= col_idx:
            self.checkbox_states[row_idx].append(ctk.BooleanVar(value=False))
        
        while len(self.results) <= row_idx:
            self.results.append([])
        while len(self.results[row_idx]) <= col_idx:
            self.results[row_idx].append(ctk.StringVar(value=""))
        
        while len(self._full_results) <= row_idx:
            self._full_results.append([])
        while len(self._full_results[row_idx]) <= col_idx:
            self._full_results[row_idx].append("")

        cell_frame = ctk.CTkFrame(self.scrollable_content_frame, border_width=1, border_color=styles.MATRIX_CELL_BORDER_COLOR, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT)
        cell_frame.grid(row=row_idx + 1, column=col_idx + 1, padx=5, pady=5, sticky="nsew")
        cell_frame.grid_propagate(False)
        cell_frame.grid_columnconfigure(1, weight=1)

        checkbox = ctk.CTkCheckBox(cell_frame, text="", variable=self.checkbox_states[row_idx][col_idx], width=15)
        checkbox.grid(row=0, column=0, padx=0, pady=2, sticky="w")

        result_textbox = ctk.CTkTextbox(cell_frame, wrap="word", height=styles.MATRIX_RESULT_CELL_HEIGHT, font=styles.MATRIX_RESULT_FONT, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        result_textbox.grid(row=0, column=1, padx=(0,2), pady=2, sticky="nsew")
        result_textbox.insert("1.0", self.results[row_idx][col_idx].get())
        result_textbox.configure(state="disabled")
        result_textbox.bind("<Button-1>", lambda event, r=row_idx, c=col_idx: self._show_full_result_popup(r, c))
        self.results[row_idx][col_idx].trace_add("write", lambda *args, sv=self.results[row_idx][col_idx], tb=result_textbox: self._update_textbox_from_stringvar(sv, tb))

        # Track textbox reference and style matrix
        while len(self._result_textboxes) <= row_idx:
            self._result_textboxes.append([])
        while len(self._result_textboxes[row_idx]) <= col_idx:
            self._result_textboxes[row_idx].append(None)
        self._result_textboxes[row_idx][col_idx] = result_textbox

        while len(self._cell_style) <= row_idx:
            self._cell_style.append([])
        while len(self._cell_style[row_idx]) <= col_idx:
            self._cell_style[row_idx].append("normal")

    def _add_input_row(self):
        try:
            self.configure(cursor='watch')
            self.update_idletasks()
        except Exception:
            pass

        new_row_idx = len(self.input_data)
        self.input_data.append({"type": "text", "data": ""})
        if self._row_summaries:
            self._row_summaries.append(ctk.StringVar(value=""))

        self._add_input_row_widgets(new_row_idx, self.input_data[new_row_idx])
        
        self.after(10, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        try:
            self.configure(cursor='')
        except Exception:
            pass

    def _add_prompt_column(self):
        try:
            self.configure(cursor='watch')
            self.update_idletasks()
        except Exception:
            pass

        new_col_idx = len(self.prompts)
        new_prompt_id = f"prompt_{new_col_idx + 1}"
        new_prompt_name = tr("prompt.default_name_fmt", n=new_col_idx + 1)
        new_prompt_config = Prompt(name=new_prompt_name, model="gemini-2.5-flash", system_prompt=tr("prompt.new_placeholder"))
        
        self.prompts[new_prompt_id] = new_prompt_config
        try:
            # update active tab prompt objects snapshot
            self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
        except Exception:
            pass
        self._column_widths.append(styles.MATRIX_CELL_WIDTH)

        if self._col_summaries:
            self._col_summaries.append(ctk.StringVar(value=""))

        self._add_prompt_header_widgets(new_col_idx, new_prompt_id, new_prompt_config)
        for row_idx in range(len(self.input_data)):
            self._create_result_cell(row_idx, new_col_idx)

        self.after(10, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        try:
            self.configure(cursor='')
        except Exception:
            pass

    def _clear_all(self):
        if not messagebox.askyesno(tr("matrix.clear_confirm_title"), tr("matrix.clear_confirm_message")):
            return

        self.input_data = [{"type": "text", "data": ""}]
        self.prompts = {}
        self.checkbox_states = []
        self.results = []
        self._full_results = []
        self._row_summaries = []
        self._col_summaries = []
        self._row_heights = []
        self._column_widths = []
        self.total_tasks = 0
        self.completed_tasks = 0

        if self.summarize_row_button:
            self.summarize_row_button.configure(state="disabled")
        if self.summarize_col_button:
            self.summarize_col_button.configure(state="disabled")
        if self.summarize_matrix_button:
            self.summarize_matrix_button.configure(state="disabled")
        if self.export_excel_button:
            self.export_excel_button.configure(state="disabled")

        try:
            self.progress_label.configure(text=tr("matrix.progress_fmt", done=0, total=0))
        except Exception:
            pass

        self._update_ui()

    def _run_batch_processing(self):
        checked_tasks = []
        for r_idx, row_input in enumerate(self.input_data):
            for c_idx, prompt_id in enumerate(self.prompts.keys()):
                if self.checkbox_states[r_idx][c_idx].get():
                    checked_tasks.append((r_idx, c_idx, row_input, prompt_id))
        
        if not checked_tasks:
            messagebox.showinfo(tr("matrix.run_title"), tr("matrix.no_checked_combinations"))
            return

        self.total_tasks = len(checked_tasks)
        self.completed_tasks = 0
        self._update_progress_label()

        num_inputs = len(self.input_data)
        num_prompts = len(self.prompts)

        while len(self.results) < num_inputs:
            self.results.append([])
        for r_idx in range(num_inputs):
            while len(self.results[r_idx]) < num_prompts:
                self.results[r_idx].append(ctk.StringVar(value=""))

        while len(self._full_results) < num_inputs:
            self._full_results.append([])
        for r_idx in range(num_inputs):
            while len(self._full_results[r_idx]) < num_prompts:
                self._full_results[r_idx].append("")

        for r_idx, c_idx, _, _ in checked_tasks:
            self.results[r_idx][c_idx].set(tr("common.processing"))
            # Normal run uses default color
            self._set_cell_style(r_idx, c_idx, "normal")

        asyncio.run_coroutine_threadsafe(self._execute_llm_tasks(checked_tasks), self.worker_loop)

    def _set_cell_style(self, r_idx: int, c_idx: int, style: str):
        try:
            while len(self._cell_style) <= r_idx:
                self._cell_style.append([])
            while len(self._cell_style[r_idx]) <= c_idx:
                self._cell_style[r_idx].append("normal")
            self._cell_style[r_idx][c_idx] = style
            tb = None
            if 0 <= r_idx < len(self._result_textboxes) and 0 <= c_idx < len(self._result_textboxes[r_idx]):
                tb = self._result_textboxes[r_idx][c_idx]
            if tb and tb.winfo_exists():
                tb.configure(state="normal")
                if style == "flow":
                    tb.configure(text_color=styles.FLOW_RESULT_TEXT_COLOR)
                else:
                    tb.configure(text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                tb.configure(state="disabled")
        except Exception:
            pass

    def _update_cell_on_main_thread(self, r_idx: int, c_idx: int, text_content: str, is_final: bool = False):
        if self._is_closing or not self.winfo_exists():
            return
        current_text = ""
        if 0 <= r_idx < len(self.results) and 0 <= c_idx < len(self.results[r_idx]):
            current_text = self.results[r_idx][c_idx].get()
        else:
            return

        new_text = text_content if is_final else current_text + text_content
        
        try:
            self.results[r_idx][c_idx].set(new_text)
        except tk.TclError:
            pass

        if is_final:
            if 0 <= r_idx < len(self._full_results) and 0 <= c_idx < len(self._full_results[r_idx]):
                self._full_results[r_idx][c_idx] = new_text
            else:
                print(f"ERROR: _update_cell_on_main_thread - _full_results のインデックス ({r_idx}, {c_idx}) が範囲外です。最終結果の保存をスキップします。")
            
            with self.progress_lock:
                self.completed_tasks += 1
                self._update_progress_label()

    async def _execute_llm_tasks(self, tasks_to_run: List[tuple]):
        self.processing_tasks = []
        for r_idx, c_idx, row_input, prompt_id in tasks_to_run:
            prompt_config = self.prompts.get(prompt_id)
            if prompt_config:
                task = asyncio.create_task(self._process_single_cell(r_idx, c_idx, row_input, prompt_config))
                self.processing_tasks.append(task)
            else:
                print(f"ERROR: _execute_llm_tasks - prompt_id '{prompt_id}' not found.")
                error_msg = tr("matrix.error_no_prompt_config")
                self.after(0, self._update_cell_on_main_thread, r_idx, c_idx, error_msg, True)

        await asyncio.gather(*self.processing_tasks)
        
        def show_completion_notification():
            if self.summarize_row_button:
                self.summarize_row_button.configure(state="normal")
            if self.summarize_col_button:
                self.summarize_col_button.configure(state="normal")
            if self.export_excel_button:
                self.export_excel_button.configure(state="normal")
            if self.summarize_matrix_button:
                self.summarize_matrix_button.configure(state="normal")
            for r_idx, c_idx, _, _ in tasks_to_run:
                if 0 <= r_idx < len(self.checkbox_states) and 0 <= c_idx < len(self.checkbox_states[r_idx]):
                    self.checkbox_states[r_idx][c_idx].set(False)
        
        self.after(0, show_completion_notification)

    def _confirm_flow(self, plans: Dict[int, List[int]]) -> bool:
        # Build message: steps per row and warnings
        lines = []
        total_steps = 0
        for r_idx, cols in plans.items():
            if not cols:
                continue
            letters = [chr(ord('A') + c) for c in cols]
            flow_str = " → ".join(letters)
            lines.append(f"{tr('action.input').rstrip(':')}{r_idx+1}: {len(cols)} {tr('matrix.flow.steps_label')} ({flow_str})")
            total_steps += len(cols)
        if total_steps == 0:
            messagebox.showinfo(tr("matrix.flow.running_title"), tr("matrix.no_checked_combinations"))
            return False
        # Overwrite notice
        overwrite = False
        for r_idx, cols in plans.items():
            for c_idx in cols:
                if 0 <= r_idx < len(self._full_results) and 0 <= c_idx < len(self._full_results[r_idx]):
                    if self._full_results[r_idx][c_idx]:
                        overwrite = True
                        break
            if overwrite:
                break
        msg = "\n".join(lines)
        if overwrite:
            msg += f"\n\n{tr('matrix.flow.overwrite_note')}"
        msg += f"\n\n{tr('matrix.flow.max_steps_label')}: {self.max_flow_steps}"
        res = messagebox.askokcancel(tr("matrix.flow.confirm_title"), msg)
        return bool(res)

    def _run_flow_processing(self):
        # Build per-row plans of selected columns, limited and sorted by column index (A..)
        plans: Dict[int, List[int]] = {}
        for r_idx, row_input in enumerate(self.input_data):
            sel_cols = [c_idx for c_idx, _ in enumerate(self.prompts.keys()) if self.checkbox_states[r_idx][c_idx].get()]
            sel_cols.sort()
            if sel_cols:
                try:
                    self.max_flow_steps = int(getattr(self.agent.config, 'max_flow_steps', self.max_flow_steps))
                except Exception:
                    pass
                plans[r_idx] = sel_cols[: int(self.max_flow_steps)]

        if not self._confirm_flow(plans):
            return

        # Initialize UI states
        total_steps = sum(len(cols) for cols in plans.values())
        self.total_tasks = total_steps
        self.completed_tasks = 0
        self._update_progress_label()

        # Mark target cells as processing and flow-styled
        for r_idx, cols in plans.items():
            for c_idx in cols:
                self.results[r_idx][c_idx].set(tr("common.processing"))
                self._set_cell_style(r_idx, c_idx, "flow")

        # Launch per-row flows concurrently
        self._flow_cancel_requested = False
        try:
            self.flow_run_button.configure(state="disabled")
        except Exception:
            pass
        self._show_flow_progress_dialog()
        asyncio.run_coroutine_threadsafe(self._execute_flow_tasks(plans), self.worker_loop)

    async def _execute_flow_for_row(self, r_idx: int, cols: List[int]):
        # Conversation history as alternating user/model messages (dicts)
        conv: List[Dict[str, Any]] = []
        # Prepare initial user parts
        input_item = self.input_data[r_idx]
        initial_parts: List[Any] = []
        if input_item["type"] == "text":
            initial_parts = [{"text": input_item["data"]}]
        elif input_item["type"] in ("image", "image_compressed"):
            img_b64 = input_item["data"]
            if input_item["type"] == "image_compressed":
                try:
                    import zlib
                    img_b64 = base64.b64encode(zlib.decompress(base64.b64decode(img_b64))).decode("utf-8")
                except Exception:
                    pass
            initial_parts = [create_image_part(img_b64)]
        elif input_item["type"] == "file":
            file_path = input_item["data"]
            try:
                mime_type, _ = mimetypes.guess_type(file_path)
                if not mime_type:
                    mime_type = "application/octet-stream"
                uploaded_file = await asyncio.to_thread(genai.upload_file, path=file_path, mime_type=mime_type)
                initial_parts = [uploaded_file]
            except Exception as e:
                err = tr("matrix.error_prefix") + tr("notify.file_upload_failed", details=str(e))
                self.after(0, self._update_cell_on_main_thread, r_idx, cols[0], err, True)
                return
        else:
            initial_parts = [{"text": ""}]

        current_parts = initial_parts
        # Sequentially run per selected column
        for step_idx, c_idx in enumerate(cols):
            if self._flow_cancel_requested:
                break
            # Build user message for current step: prepend instruction as prefix (text) or as separate part
            prompt_id = list(self.prompts.keys())[c_idx]
            prompt_config = self.prompts[prompt_id]
            combined_parts: List[Any] = []
            try:
                # If current parts are purely text, combine into one text blob with instruction prefix
                if all(isinstance(p, dict) and 'text' in p for p in current_parts):
                    joined = "\n\n".join(str(p.get('text', '')) for p in current_parts)
                    prefix = str(getattr(prompt_config, 'system_prompt', '') or '')
                    combined_text = f"{prefix}\n\n---\n\n{joined}" if prefix else joined
                    combined_parts = [{"text": combined_text}]
                else:
                    # Non-text (image/file) inputs: include instruction as separate text part before input
                    instr = str(getattr(prompt_config, 'system_prompt', '') or '')
                    if instr:
                        combined_parts.append({"text": instr})
                    combined_parts.extend(list(current_parts))
            except Exception:
                combined_parts = list(current_parts) if current_parts else [{"text": ""}]

            try:
                conv.append({"role": "user", "parts": combined_parts})
            except Exception:
                conv.append({"role": "user", "parts": [{"text": str(combined_parts)}]})
            try:
                async with self.semaphore:
                    gemini_model = GenerativeModel(prompt_config.model, system_instruction=prompt_config.system_prompt)
                    # Detect tools setting
                    has_url_text = any(isinstance(p, dict) and "text" in p and isinstance(p["text"], str) and p["text"].strip().startswith(("http://", "https://")) for p in combined_parts)
                    tools_list = [{"google_search": {}}] if getattr(prompt_config, 'enable_web', False) or has_url_text else None
                    generate_content_config = types.GenerationConfig(
                        temperature=prompt_config.parameters.temperature,
                        top_p=prompt_config.parameters.top_p,
                        top_k=prompt_config.parameters.top_k,
                        max_output_tokens=prompt_config.parameters.max_output_tokens,
                        stop_sequences=prompt_config.parameters.stop_sequences
                    )

                    def _gen_sync(config, contents, tools=None):
                        if tools is not None:
                            return gemini_model.generate_content(contents=contents, generation_config=config, tools=tools)
                        return gemini_model.generate_content(contents=contents, generation_config=config)

                    # Call model with conversation so far; fallback on tool errors
                    try:
                        response = await asyncio.to_thread(_gen_sync, generate_content_config, conv, tools_list)
                    except Exception:
                        if tools_list:
                            try:
                                alt_tools = [{"google_search_retrieval": {}}]
                                response = await asyncio.to_thread(_gen_sync, generate_content_config, conv, alt_tools)
                            except Exception:
                                response = await asyncio.to_thread(_gen_sync, generate_content_config, conv, None)
                        else:
                            raise

                    # Extract text
                    def _extract_text(resp) -> str:
                        try:
                            if getattr(resp, 'candidates', None):
                                cand = resp.candidates[0]
                                content = getattr(cand, 'content', None)
                                parts = getattr(content, 'parts', None) if content else None
                                if parts:
                                    return "".join(p.text for p in parts if hasattr(p, 'text'))
                            return getattr(resp, 'text', '') or ''
                        except Exception:
                            return ''
                    out_text = _extract_text(response)
                    if not out_text:
                        out_text = tr("matrix.response_empty")
            except Exception as e:
                out_text = tr("matrix.error_prefix") + str(e)

            # Update cell with result and style, and uncheck the box
            self.after(0, self._update_cell_on_main_thread, r_idx, c_idx, out_text, True)
            self._set_cell_style(r_idx, c_idx, "flow")
            try:
                self.after(0, lambda rr=r_idx, cc=c_idx: self.checkbox_states[rr][cc].set(False))
            except Exception:
                pass
            # Append model message and set next input as its text
            try:
                conv.append({"role": "model", "parts": [{"text": out_text}]})
            except Exception:
                conv.append({"role": "model", "parts": [{"text": str(out_text)}]})
            current_parts = [{"text": out_text}]

        # After row flow, uncheck relevant checkboxes
        def _clear_checks():
            for c_idx in cols:
                try:
                    self.checkbox_states[r_idx][c_idx].set(False)
                except Exception:
                    pass
        self.after(0, _clear_checks)

        # Enable summary/export buttons after flows (all tasks completion is handled globally too)
    async def _execute_flow_tasks(self, plans: Dict[int, List[int]]):
        self._flow_tasks = []
        for r_idx, cols in plans.items():
            task = asyncio.create_task(self._execute_flow_for_row(r_idx, cols))
            self._flow_tasks.append(task)
        try:
            await asyncio.gather(*self._flow_tasks)
        except asyncio.CancelledError:
            pass
        def _enable_actions():
            try:
                if self.summarize_row_button:
                    self.summarize_row_button.configure(state="normal")
                if self.summarize_col_button:
                    self.summarize_col_button.configure(state="normal")
                if self.summarize_matrix_button:
                    self.summarize_matrix_button.configure(state="normal")
                if self.export_excel_button:
                    self.export_excel_button.configure(state="normal")
                self.flow_run_button.configure(state="normal")
                self._close_flow_progress_dialog()
            except Exception:
                pass
        self.after(0, _enable_actions)

    async def _process_single_cell(self, r_idx: int, c_idx: int, input_item: Dict[str, Any], prompt_config: Prompt):
        full_result = ""
        try:
            async with self.semaphore:
                gemini_model = GenerativeModel(prompt_config.model, system_instruction=prompt_config.system_prompt)
                contents_to_send = []

                if input_item["type"] == "text":
                    contents_to_send.append(input_item["data"])
                elif input_item["type"] in ("image", "image_compressed"):
                    img_b64 = input_item["data"]
                    if input_item["type"] == "image_compressed":
                        try:
                            import zlib
                            decompressed = zlib.decompress(base64.b64decode(img_b64))
                            img_b64 = base64.b64encode(decompressed).decode('utf-8')
                        except Exception:
                            pass
                    image_part = create_image_part(img_b64)
                    contents_to_send.append(image_part)
                elif input_item["type"] == "file":
                    file_path = input_item["data"]
                    try:
                        mime_type, _ = mimetypes.guess_type(file_path)
                        if not mime_type:
                            mime_type = "application/octet-stream"
                        uploaded_file = await asyncio.to_thread(genai.upload_file, path=file_path, mime_type=mime_type)
                        contents_to_send.append(uploaded_file)
                    except Exception as e:
                        raise RuntimeError(tr("notify.file_upload_failed", details=str(e)))
                else:
                    raise ValueError(f"Unsupported input type: {input_item['type']}")

                has_url_text = any(isinstance(p, str) and p.strip().startswith(("http://", "https://")) for p in contents_to_send)
                tools_list = [{"google_search": {}}] if getattr(prompt_config, 'enable_web', False) or has_url_text else None

                try:
                    generate_content_config = types.GenerationConfig(temperature=prompt_config.parameters.temperature, top_p=prompt_config.parameters.top_p, top_k=prompt_config.parameters.top_k, max_output_tokens=prompt_config.parameters.max_output_tokens, stop_sequences=prompt_config.parameters.stop_sequences, tools=tools_list)
                except TypeError:
                    generate_content_config = types.GenerationConfig(temperature=prompt_config.parameters.temperature, top_p=prompt_config.parameters.top_p, top_k=prompt_config.parameters.top_k, max_output_tokens=prompt_config.parameters.max_output_tokens, stop_sequences=prompt_config.parameters.stop_sequences)

                def _gen_sync(config):
                    return gemini_model.generate_content(contents=contents_to_send, generation_config=config)

                try:
                    response = await asyncio.to_thread(_gen_sync, generate_content_config)
                except Exception:
                    if tools_list:
                        try:
                            alt_tools = [{"google_search_retrieval": {}}]
                            alt_config = types.GenerationConfig(temperature=generate_content_config.temperature, top_p=generate_content_config.top_p, top_k=generate_content_config.top_k, max_output_tokens=generate_content_config.max_output_tokens, stop_sequences=generate_content_config.stop_sequences, tools=alt_tools)
                            response = await asyncio.to_thread(_gen_sync, alt_config)
                        except Exception:
                            no_tool_config = types.GenerationConfig(temperature=generate_content_config.temperature, top_p=generate_content_config.top_p, top_k=generate_content_config.top_k, max_output_tokens=generate_content_config.max_output_tokens, stop_sequences=generate_content_config.stop_sequences)
                            response = await asyncio.to_thread(_gen_sync, no_tool_config)
                    else:
                        raise
                
                def _extract_text(resp) -> str:
                    try:
                        if getattr(resp, 'candidates', None):
                            cand = resp.candidates[0]
                            content = getattr(cand, 'content', None)
                            parts = getattr(content, 'parts', None) if content else None
                            if parts:
                                return "".join(p.text for p in parts if hasattr(p, 'text'))
                        return getattr(resp, 'text', '') or ''
                    except Exception:
                        return ''

                if response.prompt_feedback and response.prompt_feedback.block_reason:
                    full_result = tr("safety.request_blocked_message")
                    self.after(0, lambda: self.notification_callback(tr("safety.request_blocked_title"), full_result, level="error"))
                elif not response.candidates:
                    full_result = tr("matrix.no_response")
                    self.after(0, lambda: self.notification_callback(tr("common.info"), full_result, level="error"))
                else:
                    extracted = _extract_text(response)
                    full_result = extracted if extracted else (tr("matrix.response_empty") + f" finish_reason={getattr(response.candidates[0], 'finish_reason', None)}")

        except Exception as e:
            full_result = tr("matrix.error_prefix") + str(e)
            self.after(0, lambda err=e: self.notification_callback(tr("matrix.processing_error_title"), tr("matrix.cell_error_fmt", row=r_idx+1, col=c_idx+1, details=str(err)), "error"))
            traceback.print_exc()
        finally:
            self.after(0, self._update_cell_on_main_thread, r_idx, c_idx, full_result, True)

    async def _summarize_content_with_llm(self, content_list: List[str], summary_type: str, r_idx: Optional[int] = None, c_idx: Optional[int] = None) -> str:
        combined_content = "\n\n".join(content_list)
        summary_prompt_text = f"以下の{summary_type}の情報を要約してください。重要なポイントを簡潔にまとめてください。\n\n{combined_content}"
        
        cfg_prompt: Optional[Prompt] = None
        try:
            if r_idx is not None:
                cfg_prompt = getattr(self.agent.config, 'matrix_row_summary_prompt', None)
            elif c_idx is not None:
                cfg_prompt = getattr(self.agent.config, 'matrix_col_summary_prompt', None)
            else:
                cfg_prompt = getattr(self.agent.config, 'matrix_matrix_summary_prompt', None)
        except Exception:
            cfg_prompt = None
        
        summary_prompt_config = cfg_prompt or Prompt(name=f"{summary_type}要約", model="gemini-2.5-flash-lite", system_prompt="与えられた情報を簡潔に要約してください。" )
        
        full_summary_result = ""
        try:
            generation_config = genai.types.GenerationConfig(temperature=summary_prompt_config.parameters.temperature, top_p=summary_prompt_config.parameters.top_p, top_k=summary_prompt_config.parameters.top_k, max_output_tokens=summary_prompt_config.parameters.max_output_tokens, stop_sequences=summary_prompt_config.parameters.stop_sequences)
            gemini_model = GenerativeModel(summary_prompt_config.model, system_instruction=summary_prompt_config.system_prompt)
            response = await asyncio.to_thread(gemini_model.generate_content, contents=[summary_prompt_text], generation_config=generation_config)
            
            def _extract_text(resp) -> str:
                try:
                    if getattr(resp, 'candidates', None):
                        cand = resp.candidates[0]
                        content = getattr(cand, 'content', None)
                        parts = getattr(content, 'parts', None) if content else None
                        if parts:
                            return "".join(p.text for p in parts if hasattr(p, 'text'))
                    return getattr(resp, 'text', '') or ''
                except Exception:
                    return ''

            if response.prompt_feedback and response.prompt_feedback.block_reason:
                full_summary_result = tr("safety.request_blocked_message")
                self.after(0, lambda: self.notification_callback(tr("safety.request_blocked_title"), full_summary_result, level="error"))
            elif not response.candidates:
                full_summary_result = tr("matrix.final_summary.none")
                self.after(0, lambda: self.notification_callback(tr("common.info"), full_summary_result, level="error"))
            else:
                extracted = _extract_text(response)
                full_summary_result = extracted if extracted else tr("matrix.response_empty")
            
        except Exception as e:
            full_summary_result = tr("matrix.final_summary.error_fmt", details=str(e))
            self.after(0, lambda err=e: self.notification_callback(tr("matrix.final_summary.error_title"), tr("matrix.final_summary.error_fmt", details=str(err)), "error"))
            traceback.print_exc()
        
        return full_summary_result

    async def _summarize_rows_async(self):

        # initialize row summaries
        self._row_summaries = [ctk.StringVar(value=tr("common.processing")) for _ in range(len(self.input_data))]
        self.after(0, self._update_row_summary_column)

        summary_tasks = []
        for r_idx in range(len(self.input_data)):
            row_results = [self._full_results[r_idx][c_idx] for c_idx in range(len(self.prompts))]
            valid_results = [res for res in row_results if res and res != tr("common.processing") and not res.startswith(tr("matrix.error_prefix").strip())]
            
            if valid_results:
                task = asyncio.create_task(self._summarize_content_with_llm(valid_results, f"{tr('matrix.row_summary_header')} {r_idx+1}", r_idx=r_idx))
                summary_tasks.append((r_idx, task))
            else:
                self._row_summaries[r_idx].set(tr("matrix.summary.target_none"))

        results = await asyncio.gather(*[task for _, task in summary_tasks])

        for i, (r_idx, _) in enumerate(summary_tasks):
            self.after(0, lambda r=r_idx, s=results[i]: self._row_summaries[r].set(s))
        
        self.after(0, self._update_row_summary_column)
        try:
            CTkMessagebox(title=tr("matrix.row_summary_header"), message=tr("matrix.summary.row_done"), icon="info").wait_window()
        except Exception:
            pass

    async def _summarize_columns_async(self):

        self._col_summaries = [ctk.StringVar(value=tr("common.processing")) for _ in range(len(self.prompts))]
        self.after(0, self._update_column_summary_row)

        summary_tasks = []
        for c_idx in range(len(self.prompts)):
            col_results = [self._full_results[r_idx][c_idx] for r_idx in range(len(self.input_data))]
            valid_results = [res for res in col_results if res and res != tr("common.processing") and not res.startswith(tr("matrix.error_prefix").strip())]

            if valid_results:
                task = asyncio.create_task(self._summarize_content_with_llm(valid_results, f"{tr('matrix.col_summary_header')} {chr(ord('A') + c_idx)}", c_idx=c_idx))
                summary_tasks.append((c_idx, task))
            else:
                self._col_summaries[c_idx].set(tr("matrix.summary.target_none"))

        results = await asyncio.gather(*[task for _, task in summary_tasks])

        for i, (c_idx, _) in enumerate(summary_tasks):
            self.after(0, lambda c=c_idx, s=results[i]: self._col_summaries[c].set(s))
        
        self.after(0, self._update_column_summary_row)
        try:
            CTkMessagebox(title=tr("matrix.col_summary_header"), message=tr("matrix.summary.col_done"), icon="info").wait_window()
        except Exception:
            pass

    def _summarize_rows(self):
        asyncio.run_coroutine_threadsafe(self._summarize_rows_async(), self.worker_loop)

    def _summarize_columns(self):
        asyncio.run_coroutine_threadsafe(self._summarize_columns_async(), self.worker_loop)

    def _summarize_matrix(self):
        asyncio.run_coroutine_threadsafe(self._summarize_matrix_async(), self.worker_loop)

    async def _summarize_matrix_async(self):
        if not self._row_summaries or any(s.get() in ["", tr("common.processing")] for s in self._row_summaries):
            await self._summarize_rows_async()
            await asyncio.sleep(0.1)

        if not self._col_summaries or any(s.get() in ["", tr("common.processing")] for s in self._col_summaries):
            await self._summarize_columns_async()
            await asyncio.sleep(0.1)

        # まとめテキスト生成（エラーのみ除外。対象なしでもテキストとして許容）
        row_summary_texts = [f"【行 {i+1} のまとめ】\n{s.get()}" for i, s in enumerate(self._row_summaries) if s.get() and "エラー" not in s.get()]
        col_summary_texts = [f"【列 {chr(ord('A') + i)} のまとめ】\n{s.get()}" for i, s in enumerate(self._col_summaries) if s.get() and "エラー" not in s.get()]

        if not row_summary_texts and not col_summary_texts:
            try:
                CTkMessagebox(title=tr("matrix.matrix_summary"), message=tr("matrix.final_summary.none"), icon="warning").wait_window()
            except Exception:
                pass
            self.after(0, lambda: self._update_matrix_summary_cell(""))
            return

        combined_summaries = "\n\n".join(row_summary_texts + col_summary_texts)
        final_summary_prompt = f"以下の各行・各列の要約情報を基に、全体を俯瞰した総合的な結論や洞察を導き出してください。\n\n---\n\n{combined_summaries}"

        final_summary = await self._summarize_content_with_llm([final_summary_prompt], tr("matrix.matrix_summary"))

        if "エラー" not in final_summary:
            pyperclip.copy(final_summary)
            self.after(0, lambda: self._show_final_summary_popup(final_summary))
            self.after(0, lambda: self._update_matrix_summary_cell(final_summary))
            try:
                CTkMessagebox(title=tr("matrix.matrix_summary"), message=tr("matrix.final_summary.copied"), icon="info").wait_window()
            except Exception:
                pass
        else:
            try:
                CTkMessagebox(title=tr("matrix.final_summary.error_title"), message=tr("matrix.final_summary.error_fmt", details=final_summary), icon="cancel").wait_window()
            except Exception:
                pass

    def _update_matrix_summary_cell(self, summary_text: str):
        num_inputs = len(self.input_data)
        num_prompts = len(self.prompts)
        summary_row_idx = num_inputs + 1
        summary_col_idx = num_prompts + 1

        for widget in self.scrollable_content_frame.grid_slaves(row=summary_row_idx, column=summary_col_idx):
            widget.destroy()

        resizable_frame = tk.Frame(self.scrollable_content_frame, borderwidth=1, relief="solid")
        resizable_frame.grid(row=summary_row_idx, column=summary_col_idx, padx=5, pady=5, sticky="nsew")
        resizable_frame.grid_rowconfigure(0, weight=1)
        resizable_frame.grid_columnconfigure(0, weight=1)

        summary_textbox = ctk.CTkTextbox(resizable_frame, width=styles.MATRIX_CELL_WIDTH, height=styles.MATRIX_RESULT_CELL_HEIGHT, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        summary_textbox.grid(row=0, column=0, sticky="nsew")
        summary_textbox.insert("1.0", summary_text)
        summary_textbox.configure(state="disabled")
        summary_textbox.bind("<Button-1>", lambda e: self._show_final_summary_popup(summary_text))
        summary_textbox.bind("<Enter>", lambda e: e.widget.configure(cursor="hand2"))
        summary_textbox.bind("<Leave>", lambda e: e.widget.configure(cursor=""))

        sizer = SizerGrip(resizable_frame)
        sizer.grid(row=1, column=1, sticky="se")

    def _show_final_summary_popup(self, summary_text: str):
        popup = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        popup.title(tr("matrix.final_summary.title"))
        popup.geometry("700x500")
        popup.transient(self)
        popup.grab_set()

        textbox = ctk.CTkTextbox(popup, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        textbox.insert("1.0", summary_text)
        textbox.configure(state="normal")

        button_frame = ctk.CTkFrame(popup, fg_color="transparent")
        button_frame.pack(pady=5)

        ctk.CTkButton(button_frame, text=tr("common.save"), width=100, command=lambda: [self._update_matrix_summary_cell(textbox.get("1.0","end-1c")), popup.destroy()]).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text=tr("common.copy"), width=100, command=lambda: [pyperclip.copy(textbox.get("1.0","end-1c")), CTkMessagebox(title=tr("common.copy"), message=tr("common.copied_to_clipboard"), icon="info").wait_window()]).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text=tr("common.close"), width=100, command=popup.destroy).pack(side="left", padx=5)

        self.wait_window(popup)

    def _truncate_result(self, result: str, max_length: int = 100) -> str:
        if len(result) > max_length:
            return result[:max_length] + "..."
        return result

    def _update_textbox_from_stringvar(self, string_var: ctk.StringVar, textbox: ctk.CTkTextbox):
        if self._is_closing or not self.winfo_exists():
            return
        try:
            textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.insert("1.0", string_var.get())
            # Apply style-based text color if tracked
            try:
                # Find indices of this textbox if possible
                for r_idx, row in enumerate(self._result_textboxes):
                    for c_idx, tb in enumerate(row):
                        if tb is textbox:
                            style = None
                            try:
                                style = self._cell_style[r_idx][c_idx]
                            except Exception:
                                style = None
                            if style == "flow":
                                textbox.configure(text_color=styles.FLOW_RESULT_TEXT_COLOR)
                            else:
                                textbox.configure(text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                            raise StopIteration
            except StopIteration:
                pass
            textbox.configure(state="disabled")
        except tk.TclError:
            pass

    def _show_full_result_popup(self, r_idx: int, c_idx: int):
        full_result = self._full_results[r_idx][c_idx]
        popup = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        popup.title(tr("matrix.result_preview_title_fmt", row=r_idx+1, col=c_idx+1))
        popup.geometry(styles.MATRIX_POPUP_GEOMETRY)
        textbox = ctk.CTkTextbox(popup, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        textbox.insert("1.0", full_result)
        textbox.configure(state="normal")
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=5)
        ctk.CTkButton(btn_frame, text=tr("common.save"), width=100, command=lambda: self._save_full_result_and_close_popup(popup, textbox, r_idx, c_idx)).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.copy"), width=100, command=lambda: [pyperclip.copy(textbox.get("1.0","end-1c")), CTkMessagebox(title=tr("common.copy"), message=tr("common.copied_to_clipboard"), icon="info").wait_window()]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.close"), width=100, command=popup.destroy).pack(side="left", padx=5)
        popup.grab_set()
        self.wait_window(popup)
        self.grab_release()

    def _show_full_row_summary_popup(self, r_idx: int):
        full_summary = self._row_summaries[r_idx].get()
        popup = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        popup.title(tr("matrix.row_result_preview_title_fmt", row=r_idx+1))
        popup.geometry(styles.MATRIX_POPUP_GEOMETRY)
        textbox = ctk.CTkTextbox(popup, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        textbox.insert("1.0", full_summary)
        textbox.configure(state="normal")
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=5)
        ctk.CTkButton(btn_frame, text=tr("common.save"), width=100, command=lambda: self._save_full_row_summary_and_close_popup(popup, textbox, r_idx)).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.copy"), width=100, command=lambda: [pyperclip.copy(textbox.get("1.0","end-1c")), CTkMessagebox(title=tr("common.copy"), message=tr("common.copied_to_clipboard"), icon="info").wait_window()]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.close"), width=100, command=popup.destroy).pack(side="left", padx=5)
        popup.grab_set()
        self.wait_window(popup)
        self.grab_release()

    def _show_full_col_summary_popup(self, c_idx: int):
        full_summary = self._col_summaries[c_idx].get()
        popup = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        popup.title(tr("matrix.col_result_preview_title_fmt", col=chr(ord('A') + c_idx)))
        popup.geometry(styles.MATRIX_POPUP_GEOMETRY)
        textbox = ctk.CTkTextbox(popup, wrap="word", fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        textbox.insert("1.0", full_summary)
        textbox.configure(state="normal")
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=5)
        ctk.CTkButton(btn_frame, text=tr("common.save"), width=100, command=lambda: self._save_full_col_summary_and_close_popup(popup, textbox, c_idx)).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.copy"), width=100, command=lambda: [pyperclip.copy(textbox.get("1.0","end-1c")), CTkMessagebox(title=tr("common.copy"), message=tr("common.copied_to_clipboard"), icon="info").wait_window()]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=tr("common.close"), width=100, command=popup.destroy).pack(side="left", padx=5)
        popup.grab_set()
        self.wait_window(popup)
        self.grab_release()

    def _save_full_result_and_close_popup(self, popup: ctk.CTkToplevel, textbox: ctk.CTkTextbox, r_idx: int, c_idx: int):
        edited_content = textbox.get("1.0", "end-1c")
        self._full_results[r_idx][c_idx] = edited_content
        self.results[r_idx][c_idx].set(self._truncate_result(edited_content))
        popup.destroy()

    def _save_full_row_summary_and_close_popup(self, popup: ctk.CTkToplevel, textbox: ctk.CTkTextbox, r_idx: int):
        edited = textbox.get("1.0", "end-1c")
        try:
            self._row_summaries[r_idx].set(edited)
        except Exception:
            pass
        popup.destroy()

    def _save_full_col_summary_and_close_popup(self, popup: ctk.CTkToplevel, textbox: ctk.CTkTextbox, c_idx: int):
        edited_content = textbox.get("1.0", "end-1c")
        self._col_summaries[c_idx].set(edited_content)
        popup.destroy()

    def _delete_row(self, row_idx: int):
        if not messagebox.askyesno(tr("matrix.delete_row_title"), f"{tr('matrix.delete_row_confirm_fmt', row=row_idx + 1)}\n{tr('common.cannot_undo')}"):
            return
        try:
            for widget in list(self.scrollable_content_frame.grid_slaves(row=row_idx + 1)):
                widget.destroy()
        except Exception:
            pass
        def finalize_delete():
            try:
                if 0 <= row_idx < len(self.input_data):
                    self.input_data.pop(row_idx)
                    self.checkbox_states.pop(row_idx)
                    self.results.pop(row_idx)
                    self._full_results.pop(row_idx)
                    if self._row_summaries and 0 <= row_idx < len(self._row_summaries):
                        self._row_summaries.pop(row_idx)
                    if self._row_heights and 0 <= row_idx + 1 < len(self._row_heights):
                        self._row_heights.pop(row_idx + 1)
            except Exception:
                pass
            if not self.input_data:
                self._clear_all()
            self._update_ui()
        self.after(10, finalize_delete)

    def _open_prompt_editor(self, prompt_id: str):
        from ui_components import PromptEditorDialog
        
        current_prompt = self.prompts.get(prompt_id)
        if not current_prompt:
            return

        dlg = PromptEditorDialog(self, title=tr("prompt.edit_title_fmt", name=current_prompt.name), prompt=current_prompt)
        result_prompt = dlg.get_result()

        if result_prompt:
            self.prompts[prompt_id] = result_prompt
            try:
                self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
            except Exception:
                pass
            self._update_prompt_header_display(prompt_id)

    def _update_prompt_header_display(self, prompt_id: str):
        try:
            col_idx = list(self.prompts.keys()).index(prompt_id)
        except ValueError:
            return

        for widget in self.scrollable_content_frame.grid_slaves(row=0, column=col_idx + 1):
            widget.destroy()
        
        prompt_config = self.prompts[prompt_id]
        self._add_prompt_header_widgets(col_idx, prompt_id, prompt_config)

    def _delete_column(self, col_idx: int):
        col_letter = chr(ord('A') + col_idx)
        if not messagebox.askyesno(tr("matrix.delete_col_title"), f"{tr('matrix.delete_col_confirm_fmt', col=col_letter)}\n{tr('common.cannot_undo')}"):
            return
        try:
            for widget in list(self.scrollable_content_frame.grid_slaves(column=col_idx + 1)):
                widget.destroy()
        except Exception:
            pass
        def finalize_delete():
            try:
                prompt_keys = list(self.prompts.keys())
                if 0 <= col_idx < len(prompt_keys):
                    del self.prompts[prompt_keys[col_idx]]
                for r_idx in range(len(self.input_data)):
                    if r_idx < len(self.checkbox_states) and 0 <= col_idx < len(self.checkbox_states[r_idx]):
                        self.checkbox_states[r_idx].pop(col_idx)
                    if r_idx < len(self.results) and 0 <= col_idx < len(self.results[r_idx]):
                        self.results[r_idx].pop(col_idx)
                    if r_idx < len(self._full_results) and 0 <= col_idx < len(self._full_results[r_idx]):
                        self._full_results[r_idx].pop(col_idx)
                if self._col_summaries and 0 <= col_idx < len(self._col_summaries):
                    self._col_summaries.pop(col_idx)
                if self._column_widths and 0 <= col_idx + 1 < len(self._column_widths):
                    self._column_widths.pop(col_idx + 1)
                try:
                    self._tabs[self._active_tab_index]['prompts_obj'] = {pid: (p.model_copy(deep=True) if hasattr(p, 'model_copy') else Prompt(**p.model_dump())) for pid, p in self.prompts.items()}
                except Exception:
                    pass
            except Exception:
                pass
            if not self.prompts:
                self._clear_all()
            self._update_ui()
        self.after(10, finalize_delete)

    def _select_input_source(self, row_idx: int):
        file_path = filedialog.askopenfilename(title=tr("matrix.select_input_file"), filetypes=[("All Supported", "*.png *.jpg *.jpeg *.gif *.bmp *.webp *.pdf *.txt *.md *.csv *.py *.mp3 *.wav *.xlsx *.doc *.docx"), ("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("PDF", "*.pdf"), ("Text", "*.txt *.md *.csv"), ("All Files", "*.*")])
        if not file_path:
            return
        file_path_obj = Path(file_path)
        file_type = "file"
        data_content = file_path
        if file_path_obj.suffix.lower() in [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]:
            try:
                image = Image.open(file_path)
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                with BytesIO() as buffer:
                    image.save(buffer, format="PNG")
                    data_content = base64.b64encode(buffer.getvalue()).decode('utf-8')
                file_type = "image"
            except Exception as e:
                messagebox.showerror(tr("common.error_title"), tr("matrix.image_preview_failed", details=str(e)))
                return
        self.input_data[row_idx] = {"type": file_type, "data": data_content}
        self._update_input_row_display(row_idx)
    
    def _show_image_preview(self, row_idx: int):
        input_item = self.input_data[row_idx]
        if input_item["type"] in ("image", "image_compressed"):
            try:
                raw = base64.b64decode(input_item["data"])
                if input_item["type"] == "image_compressed":
                    import zlib
                    raw = zlib.decompress(raw)
                image = Image.open(BytesIO(raw))
                popup = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
                popup.title(tr("matrix.image_preview_title_fmt", row=row_idx+1))
                max_width = self.winfo_width() * 0.8
                max_height = self.winfo_height() * 0.8
                image.thumbnail((max_width, max_height), Image.LANCZOS)
                ctk_image = ctk.CTkImage(light_image=image, dark_image=image, size=image.size)
                image_label = ctk.CTkLabel(popup, image=ctk_image, text="", text_color=styles.HISTORY_ITEM_TEXT_COLOR)
                image_label.pack(padx=10, pady=10)
                close_button = ctk.CTkButton(popup, text=tr("common.close"), command=popup.destroy, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
                close_button.pack(pady=5)
                popup.grab_set()
                self.wait_window(popup)
                self.grab_release()
            except Exception as e:
                messagebox.showerror(tr("common.error_title"), tr("matrix.image_preview_failed", details=str(e)))
        else:
            messagebox.showinfo(tr("common.info"), tr("matrix.not_image_row"))

    def _update_progress_label(self):
        if self._is_closing or not self.winfo_exists():
            return
        try:
            self.progress_label.configure(text=tr("matrix.progress_fmt", done=self.completed_tasks, total=self.total_tasks))
        except tk.TclError:
            pass

    def _cancel_flow_processing(self):
        # Signal cancellation and attempt to cancel tasks
        self._flow_cancel_requested = True
        try:
            for t in list(self._flow_tasks):
                try:
                    if not t.done():
                        t.cancel()
                except Exception:
                    pass
            # Update progress dialog state
            try:
                if hasattr(self, '_flow_dialog_label') and self._flow_dialog_label:
                    self._flow_dialog_label.configure(text=tr("matrix.flow.stopping_message"))
            except Exception:
                pass
        except Exception:
            pass

    # --- Flow progress dialog ---
    def _show_flow_progress_dialog(self):
        try:
            if getattr(self, '_flow_dialog', None) and self._flow_dialog.winfo_exists():
                return
        except Exception:
            pass
        dlg = ctk.CTkToplevel(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        dlg.title(tr("matrix.flow.running_title"))
        dlg.geometry("360x140")
        dlg.transient(self)
        dlg.grab_set()
        lbl = ctk.CTkLabel(dlg, text=tr("matrix.flow.running_message"), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        lbl.pack(padx=16, pady=(20, 10))
        btn = ctk.CTkButton(dlg, text=tr("matrix.flow.cancel"), width=100, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR, command=self._cancel_flow_processing)
        btn.pack(pady=10)
        self._flow_dialog = dlg
        self._flow_dialog_label = lbl

    def _close_flow_progress_dialog(self):
        try:
            if getattr(self, '_flow_dialog', None) and self._flow_dialog.winfo_exists():
                self._flow_dialog.grab_release()
                self._flow_dialog.destroy()
        except Exception:
            pass
        self._flow_dialog = None
        self._flow_dialog_label = None

    def _start_cursor_monitoring(self):
        def check_cursor():
            if self._is_closing or not self.winfo_exists():
                return
            try:
                mouse_x = self.winfo_pointerx()
                mouse_y = self.winfo_pointery()
                window_x = self.winfo_rootx()
                window_y = self.winfo_rooty()
                width = self.winfo_width()
                height = self.winfo_height()
                rel_x = mouse_x - window_x
                rel_y = mouse_y - window_y
                if 0 <= rel_x <= width and 0 <= rel_y <= height:
                    self._update_cursor_direct(rel_x, rel_y)
                else:
                    if self.cget("cursor") != "":
                        self.configure(cursor="")
            except Exception:
                try:
                    self.configure(cursor="")
                except:
                    pass
            if self.winfo_exists():
                self._cursor_update_job = self.after(100, check_cursor)
        check_cursor()

    def _update_cursor_direct(self, x, y):
        border_width = 8
        width = self.winfo_width()
        height = self.winfo_height()
        cursor_type = ""
        on_top_border = y < border_width
        on_bottom_border = y > height - border_width
        on_left_border = x < border_width
        on_right_border = x > width - border_width
        if on_top_border and on_left_border:
            cursor_type = "top_left_corner"
        elif on_top_border and on_right_border:
            cursor_type = "top_right_corner"
        elif on_bottom_border and on_left_border:
            cursor_type = "bottom_left_corner"
        elif on_bottom_border and on_right_border:
            cursor_type = "bottom_right_corner"
        elif on_top_border or on_bottom_border:
            cursor_type = "sb_v_double_arrow"
        elif on_left_border or on_right_border:
            cursor_type = "sb_h_double_arrow"
        if self.cget("cursor") != cursor_type:
            try:
                self.configure(cursor=cursor_type)
            except tk.TclError:
                pass

    def _export_to_excel(self):
        try:
            header_parts = [tr("matrix.export.header_input"), tr("matrix.export.header_prompt_name"), tr("matrix.export.header_system_prompt")] + [p.name for p in self.prompts.values()]
            if self._row_summaries:
                header_parts.append(tr("matrix.row_summary_header"))
            tsv_data = ["\t".join(header_parts)]
            for r_idx, input_item in enumerate(self.input_data):
                input_display = input_item["data"] if input_item["type"] == "text" else f"[{input_item['type']}]"
                row_parts = [input_display, "", ""]
                for c_idx in range(len(self.prompts)):
                    cell_result = self._full_results[r_idx][c_idx] if r_idx < len(self._full_results) and c_idx < len(self._full_results[r_idx]) else ""
                    row_parts.append(cell_result.replace('\n', ' ').replace('\r', ''))
                if self._row_summaries:
                    row_summary = self._row_summaries[r_idx].get() if r_idx < len(self._row_summaries) else ""
                    row_parts.append(row_summary.replace('\n', ' ').replace('\r', ''))
                tsv_data.append("\t".join(row_parts))
            if self._col_summaries:
                col_summary_parts = [tr("matrix.col_summary_header"), "", ""] + [self._col_summaries[c_idx].get().replace('\n', ' ').replace('\r', '') if c_idx < len(self._col_summaries) else "" for c_idx in range(len(self.prompts))]
                if self._row_summaries:
                    col_summary_parts.append("")
                tsv_data.append("\t".join(col_summary_parts))
            pyperclip.copy("\n".join(tsv_data))
            messagebox.showinfo(tr("matrix.export.title"), tr("matrix.export.copied"))
        except Exception as e:
            messagebox.showerror(tr("common.error_title"), tr("matrix.export.error", details=str(e)))

    def _show_clipboard_history_popup(self, row_idx: int):
        if not self.agent or not hasattr(self.agent, 'clipboard_history'):
            self.notification_callback(tr("common.error"), tr("history.unavailable"), "error")
            return
        history_for_popup = []
        for item in self.agent.clipboard_history:
            if isinstance(item, str):
                history_for_popup.append({"type": "text", "data": item})
            elif isinstance(item, dict) and "type" in item and "data" in item:
                history_for_popup.append(item)
        def on_select(selected_item: Dict[str, Any]):
            self._set_input_data_from_history(row_idx, selected_item)
        if self._history_popup and self._history_popup.winfo_exists():
            try:
                self._history_popup.destroy()
            except Exception:
                pass
        def _on_popup_destroy():
            self._history_popup = None
        self._history_popup = ClipboardHistorySelectorPopup(parent_app=self, clipboard_history=history_for_popup, on_select_callback=on_select, on_destroy_callback=_on_popup_destroy)
        self._history_popup.show_at_cursor()

    def _open_history_edit_dialog(self, row_idx: int):
        try:
            if not (0 <= row_idx < len(self.input_data)):
                return
            item = self.input_data[row_idx]
            if not isinstance(item, dict) or item.get('type') != 'text':
                return
            dlg = HistoryEditDialog(self.parent_app, initial_value=item.get('data', ''))
            dlg.show()
            new_text = dlg.get_input()
            if new_text is not None:
                self.input_data[row_idx] = {"type": "text", "data": new_text}
                self._set_input_data_from_history(row_idx, self.input_data[row_idx])
        except Exception:
            pass

    def _set_input_data_from_history(self, row_idx: int, selected_item: Dict[str, Any]):
        if not (0 <= row_idx < len(self.input_data)):
            self.notification_callback(tr("common.error"), tr("matrix.invalid_row_index"), "error")
            return
        self.input_data[row_idx] = selected_item
        try:
            self._update_input_row_display(row_idx)
        except Exception as e:
            print(f"ERROR: _set_input_data_from_history - UIの更新に失敗: {e}")
            traceback.print_exc()
            self._update_ui()

## Duplicate MatrixSummarySettingsDialog removed; use main prompt manager window instead

class ClipboardHistorySelectorPopup(ctk.CTkToplevel):
    """クリップボード履歴から項目を選択するためのポップアップウィンドウ。"""
    def __init__(self, parent_app: ctk.CTk, clipboard_history: List[Dict[str, Any]], on_select_callback: Callable[[Dict[str, Any]], None], on_destroy_callback: Optional[Callable] = None):
        super().__init__(parent_app)
        self.transient(parent_app)
        self.grab_set()

        self.withdraw()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        
        self.on_select_callback = on_select_callback
        self._on_destroy_callback = on_destroy_callback
        self._history_items = clipboard_history
        self._buttons: List[ctk.CTkButton] = []
        self._current_selection_index = 0
        self._is_destroying = False

        self.main_frame = ctk.CTkFrame(self, fg_color=styles.POPUP_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.scrollable_frame = ctk.CTkScrollableFrame(self.main_frame)
        self.scrollable_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        for i, item in enumerate(self._history_items):
            try:
                if item.get("type") == "text":
                    text = item.get("data", "")
                    label = text.replace("\n", " ")[:80] + ("..." if len(text) > 80 else "")
                elif item.get("type") in ("image", "image_compressed"):
                    label = tr("history.image")
                elif item.get("type") == "file":
                    label = f"[" + tr("history.file_name_prefix", name=Path(item.get('data', '')).name) + "]"
                else:
                    label = tr("history.unknown")
            except Exception:
                label = tr("history.display_error", error="")

            button = ctk.CTkButton(self.scrollable_frame, text=label, command=lambda i=item: self._on_item_selected(i), anchor="w", fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
            button.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
            self._buttons.append(button)

        cancel_button = ctk.CTkButton(self.main_frame, text=tr("common.cancel"), command=self.destroy, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        cancel_button.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))

        self.bind("<Escape>", lambda e: self.destroy())
        self._update_selection_highlight()

    def _on_item_selected(self, item: Dict[str, Any]):
        try:
            self.master.after(1, self.on_select_callback, item)
        finally:
            self.destroy()

    def _update_selection_highlight(self):
        for i, button in enumerate(self._buttons):
            if i == self._current_selection_index:
                button.configure(border_color=styles.HIGHLIGHT_BORDER_COLOR, border_width=styles.HIGHLIGHT_BORDER_WIDTH)
            else:
                button.configure(border_width=0)

    def show_at_cursor(self):
        self.update_idletasks()
        width = 400
        height = 300
        x = self.winfo_pointerx()
        y = self.winfo_pointery()
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        if x + width > screen_width:
            x = screen_width - width
        if y + height > screen_height:
            y = screen_height - height
            
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.deiconify()
        self.lift()
        self.focus_force()

    def destroy(self):
        self.grab_release()
        self._is_destroying = True
        if self._on_destroy_callback:
            self._on_destroy_callback()
        super().destroy()
