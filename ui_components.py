# ui_components.py
import tkinter as tk
from tkinter import filedialog
from typing import Dict, Optional, Callable, Any, Literal, List, TYPE_CHECKING
from io import BytesIO
from PIL import Image
import base64
import zlib
from pathlib import Path
import sys

import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from history_dialogs import HistoryEditDialog

from common_models import Prompt, PromptParameters
from matrix_batch_processor import MatrixBatchProcessorWindow
import styles
from constants import API_SERVICE_ID, SUPPORTED_MODELS, model_id_to_label, model_label_to_id
from i18n import tr, set_locale, available_locales
import keyring
import google.generativeai as genai
from config_manager import save_config
import keyboard

"""Reusable UI components and dialogs for the application.

This module contains:
- BaseDialog: common toplevel styling and behaviour
- PromptParameterEditorFrame: editor for LLM parameters
- ActionSelectorWindow: quick prompt picker and attachments bar
- NotificationPopup: streaming result popup near cursor
- PromptEditorDialog: prompt CRUD dialog
- SettingsWindow: API key and hotkey settings
- ResizableInputDialog: free input helper
"""

# Agent クラスの循環参照を避けるための前方宣言
if TYPE_CHECKING:
    from agent import Agent

# --- 共通ダイアログクラス ---
class BaseDialog(ctk.CTkToplevel):
    """
    アプリケーション内のポップアップダイアログの共通基底クラス。

    このクラスでは背景色・テキスト色・ウィンドウサイズの中央表示を統一し、
    子クラスで個別のUI要素を追加する際の基本設定を提供します。
    """
    def __init__(self, parent_app: Optional[ctk.CTk] = None, title: str = "", geometry: str = "400x300"):
        super().__init__(parent_app)
        # タイトル
        if title:
            self.title(title)
        # 背景色を統一
        self.configure(fg_color=styles.POPUP_BG_COLOR)
        # ウィンドウサイズを設定し中央に配置
        try:
            width, height = [int(x) for x in geometry.split("x")]
        except Exception:
            width, height = 400, 300
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        # リサイズ許可
        self.resizable(True, True)
        # モーダル表示
        if parent_app is not None:
            self.transient(parent_app)
            self.grab_set()
        # 常に最前面に表示
        self.attributes("-topmost", True)
        # デフォルトの閉じる動作
        self.protocol("WM_DELETE_WINDOW", self.destroy)

class PromptParameterEditorFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.grid_columnconfigure(0, weight=1) # ラベルの列も伸縮可能にする
        self.grid_columnconfigure(1, weight=1) # 入力ウィジェットの列も伸縮可能にする

        ctk.CTkLabel(self, text=tr("params.temperature"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        temp_frame = ctk.CTkFrame(self, fg_color="transparent")
        temp_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        temp_frame.grid_columnconfigure(0, weight=1)
        self.temperature_slider = ctk.CTkSlider(temp_frame, from_=0, to=2, number_of_steps=20, command=self._update_temperature_label)
        self.temperature_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.temperature_value_label = ctk.CTkLabel(temp_frame, width=4, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.temperature_value_label.grid(row=0, column=1, padx=5)

        ctk.CTkLabel(self, text=tr("params.top_p"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        top_p_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_p_frame.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        top_p_frame.grid_columnconfigure(0, weight=1)
        self.top_p_slider = ctk.CTkSlider(top_p_frame, from_=0, to=1, number_of_steps=20, command=self._update_top_p_label)
        self.top_p_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.top_p_value_label = ctk.CTkLabel(top_p_frame, width=4, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.top_p_value_label.grid(row=0, column=1, padx=5)

        ctk.CTkLabel(self, text=tr("params.top_k"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        top_k_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_k_frame.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        top_k_frame.grid_columnconfigure(0, weight=1)
        self.top_k_slider = ctk.CTkSlider(top_k_frame, from_=1, to=100, number_of_steps=99, command=self._update_top_k_label)
        self.top_k_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.top_k_value_label = ctk.CTkLabel(top_k_frame, width=4, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.top_k_value_label.grid(row=0, column=1, padx=5)

        ctk.CTkLabel(self, text=tr("params.max_output_tokens"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.max_output_tokens_entry = ctk.CTkEntry(self, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.max_output_tokens_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self, text=tr("params.stop_sequences"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.stop_sequences_entry = ctk.CTkEntry(self, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.stop_sequences_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

    def _update_temperature_label(self, value):
        self.temperature_value_label.configure(text=f"{float(value):.1f}")

    def _update_top_p_label(self, value):
        self.top_p_value_label.configure(text=f"{float(value):.2f}")

    def _update_top_k_label(self, value):
        self.top_k_value_label.configure(text=f"{int(value)}")

    def get_parameters(self) -> PromptParameters:
        return PromptParameters(
            temperature=float(self.temperature_slider.get()),
            top_p=float(self.top_p_slider.get()),
            top_k=int(self.top_k_slider.get()),
            max_output_tokens=int(self.max_output_tokens_entry.get()) if self.max_output_tokens_entry.get() else None,
            stop_sequences=[s.strip() for s in self.stop_sequences_entry.get().split(",")] if self.stop_sequences_entry.get() else None
        )

    def set_parameters(self, parameters: PromptParameters):
        self.temperature_slider.set(parameters.temperature)
        self._update_temperature_label(parameters.temperature)
        if parameters.top_p is not None:
            self.top_p_slider.set(parameters.top_p)
            self._update_top_p_label(parameters.top_p)
        if parameters.top_k is not None:
            self.top_k_slider.set(parameters.top_k)
            self._update_top_k_label(parameters.top_k)
        if parameters.max_output_tokens is not None:
            self.max_output_tokens_entry.insert(0, str(parameters.max_output_tokens))
        if parameters.stop_sequences is not None:
            self.stop_sequences_entry.insert(0, ",".join(parameters.stop_sequences))

# 以降は上部のインポートを再利用（重複インポートを削除）

class ActionSelectorWindow(ctk.CTkToplevel):
    def __init__(self, prompts: Dict[str, Prompt], on_prompt_selected_callback: Callable, agent: "Agent", file_paths: Optional[List[str]] = None, on_destroy_callback: Optional[Callable] = None):
        super().__init__()
        self.withdraw() # ウィンドウを非表示で初期化
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.geometry(styles.ACTION_SELECTOR_GEOMETRY)  # 固定サイズを設定
        self.prompts = prompts
        self.on_prompt_selected_callback = on_prompt_selected_callback
        self.agent = agent
        self._on_destroy_callback = on_destroy_callback
        self.buttons: List[ctk.CTkButton] = []  # non-prompt buttons
        self.prompt_buttons: List[ctk.CTkButton] = []  # prompt list buttons
        self.current_selection_index = 0  # index within prompt_buttons
        self._is_destroying = False
        self.attached_file_paths = file_paths
        self._history_label_to_item: Dict[str, Dict[str, Any]] = {}
        self._selected_history_item: Optional[Dict[str, Any]] = None
        # フォーカス外れ時の自動クローズ用
        self._pending_close_id: Optional[str] = None

        self.grid_columnconfigure(0, weight=1)
        self.button_height = styles.ACTION_SELECTOR_BUTTON_HEIGHT
        self.spacing = styles.ACTION_SELECTOR_SPACING
        self.margin = styles.ACTION_SELECTOR_MARGIN

        # クリップボード履歴選択エリア（最上部）
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=self.margin, pady=(self.margin, 0))
        header_frame.grid_columnconfigure(0, weight=0, minsize=35)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=0)

        clipboard_label = ctk.CTkLabel(header_frame, text=tr("action.input"), text_color=styles.HISTORY_ITEM_TEXT_COLOR, anchor="w")
        clipboard_label.grid(row=0, column=0, sticky="w", padx=(0,8))

        # 履歴の表示用ラベルを整形
        labels: List[str] = []
        self._history_label_to_item.clear()
        for it in getattr(self.agent, 'clipboard_history', [])[:50]:
            try:
                if isinstance(it, str):
                    lbl = it.replace('\n', ' ')[:60] + ("…" if len(it) > 60 else "")
                    self._history_label_to_item[lbl] = {"type": "text", "data": it}
                elif isinstance(it, dict) and 'type' in it and 'data' in it:
                    t = it.get('type')
                    if t == 'text':
                        data = it.get('data', '')
                        lbl = data.replace('\n', ' ')[:60] + ("…" if len(data) > 60 else "")
                    elif t in ('image', 'image_compressed'):
                        lbl = tr("history.image")
                    elif t == 'file':
                        try:
                            from pathlib import Path
                            lbl = f"[" + tr("history.file_name_prefix", name=Path(it.get('data','')).name) + "]"
                        except Exception:
                            lbl = tr("history.file")
                    else:
                        lbl = tr("history.unknown")
                    # ラベルの重複を避けるため少しユニーク化
                    uniq_lbl = lbl
                    c = 1
                    while uniq_lbl in self._history_label_to_item:
                        c += 1
                        uniq_lbl = f"{lbl} ({c})"
                    self._history_label_to_item[uniq_lbl] = it
                else:
                    continue
            except Exception:
                continue
        labels = list(self._history_label_to_item.keys()) or [tr("history.empty")]

        # 固定幅コンテナ内に OptionMenu を配置し、リスト長に依存して横幅が変わらないようにする
        om_container = ctk.CTkFrame(header_frame, fg_color="transparent", width=345, height=28)
        om_container.grid(row=0, column=1, sticky="w")
        try:
            om_container.grid_propagate(False)
        except Exception:
            pass
        self.history_variable = ctk.StringVar(value=labels[0])
        self.history_menu = ctk.CTkOptionMenu(om_container, values=labels, variable=self.history_variable, width=312, height=28,command=lambda _v: self._on_history_changed())
        self.history_menu.pack(fill="x")

        self.edit_history_button = ctk.CTkButton(header_frame, text=tr("common.edit"), width=60, command=self._on_edit_history, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.edit_history_button.grid(row=0, column=2, padx=(5,0))

        # 設定アイコンはアクションボタン行の右端へ移動

        self._on_history_changed()  # 初期状態反映

        # 入力の下に「添付ファイル表示 + ファイル添付ボタン」を配置
        attach_row_frame = ctk.CTkFrame(self, fg_color="transparent")
        attach_row_frame.pack(fill="x", padx=self.margin, pady=(4, 0))
        attach_row_frame.grid_columnconfigure(0, weight=1)
        attach_row_frame.grid_columnconfigure(1, weight=0)

        # 添付ファイル名表示（左側、伸縮）
        self.attached_files_label = ctk.CTkLabel(attach_row_frame, text="", text_color=styles.HISTORY_ITEM_TEXT_COLOR, anchor="w", justify="left")
        self.attached_files_label.grid(row=0, column=0, padx=(0,6), sticky="ew")

        # 添付ボタン（右側）: 編集ボタンと幅を合わせる
        try:
            attach_btn_width = int(self.edit_history_button.cget("width"))
            if attach_btn_width <= 0:
                attach_btn_width = 60
        except Exception:
            attach_btn_width = 60
        file_attach_button = ctk.CTkButton(
            attach_row_frame,
            text=tr("action.attach"),
            width=attach_btn_width,
            command=self._on_file_attach,
            height=self.button_height,
            fg_color=styles.DEFAULT_BUTTON_FG_COLOR,
            text_color=styles.DEFAULT_BUTTON_TEXT_COLOR
        )
        file_attach_button.grid(row=0, column=1, padx=2, sticky="ew")
        # 添付ボタンはプロンプトナビ対象外
        try:
            # Enterキーで添付が発火しないよう、Returnを抑止（フォーカスが当たっても誤発火を防ぐ）
            file_attach_button.bind("<Return>", lambda e: "break")
            file_attach_button.bind("<KeyPress-Return>", lambda e: "break")
        except Exception:
            pass

        self.scrollable_frame = ctk.CTkScrollableFrame(self)
        self.scrollable_frame.pack(fill="both", expand=True, padx=self.margin, pady=self.margin)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        prompt_items = list(self.prompts.items())
        for i, (prompt_id, prompt) in enumerate(prompt_items):
            button = ctk.CTkButton(self.scrollable_frame, text=prompt.name,
                                   command=lambda p_id=prompt_id, p_conf=prompt: self._on_prompt_selected(p_id, p_conf),
                                   height=self.button_height,
                                   anchor="w",
                                   fg_color=styles.DEFAULT_BUTTON_FG_COLOR,
                                   text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
            button.grid(row=i, column=0, padx=5, pady=self.spacing, sticky="ew")
            self.prompt_buttons.append(button)

        # 機能ボタンを追加（自由入力・マトリクス・設定 を横一列）
        action_buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_buttons_frame.pack(fill="x", padx=self.margin, pady=self.spacing)
        # 列: 0=自由入力, 1=マトリクス, 2=設定
        action_buttons_frame.grid_columnconfigure(0, weight=1)
        action_buttons_frame.grid_columnconfigure(1, weight=1)
        action_buttons_frame.grid_columnconfigure(2, weight=0)

        free_input_button = ctk.CTkButton(action_buttons_frame, text=tr("action.free_input"), command=self._on_free_input, height=self.button_height, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        free_input_button.grid(row=0, column=0, padx=2, sticky="ew")
        self.buttons.append(free_input_button)

        matrix_button = ctk.CTkButton(action_buttons_frame, text=tr("action.matrix"), command=self._on_matrix, height=self.button_height, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        matrix_button.grid(row=0, column=1, padx=2, sticky="ew")
        self.buttons.append(matrix_button)

        # 設定アイコン（右端）
        try:
            candidates = []
            try:
                candidates.append(Path.cwd() / "config.ico")
            except Exception:
                pass
            try:
                candidates.append(Path(__file__).resolve().parent / "config.ico")
                candidates.append(Path(__file__).resolve().parent.parent / "config.ico")
            except Exception:
                pass
            try:
                meipass = getattr(sys, "_MEIPASS", None)
                if meipass:
                    candidates.append(Path(meipass) / "config.ico")
            except Exception:
                pass
            cfg_path = next((p for p in candidates if p.exists()), None)
            if cfg_path is None:
                raise FileNotFoundError("config.ico not found in candidates")
            cfg_img_pil = Image.open(cfg_path).convert("RGBA")
            self._cfg_img = ctk.CTkImage(light_image=cfg_img_pil, dark_image=cfg_img_pil, size=(20, 20))
            self.open_manager_button = ctk.CTkButton(action_buttons_frame, text="", image=self._cfg_img, width=28, height=28, command=self._on_open_prompt_manager)
        except Exception:
            self.open_manager_button = ctk.CTkButton(action_buttons_frame, text=tr("settings.title"), width=60, command=self._on_open_prompt_manager, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.open_manager_button.grid(row=0, column=2, padx=(6,0), sticky="e")

        # 追指示チェックは設けない（回答はクリップボードから処理可能）

        cancel_button = ctk.CTkButton(self, text=tr("common.cancel"), command=self.destroy,
                                     height=self.button_height, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        cancel_button.pack(fill="x", padx=self.margin, pady=(0, self.margin))
        self.buttons.append(cancel_button)

        self._update_selection_highlight()

        self.bind("<Up>", self._on_key_up)
        self.bind("<Down>", self._on_key_down)
        # Tab/Shift-Tab でもプロンプト移動できるように
        self.bind("<Tab>", self._on_key_down)
        self.bind("<ISO_Left_Tab>", self._on_key_up)  # Shift+Tab on some platforms
        self.bind("<Shift-Tab>", self._on_key_up)
        self.bind("<Return>", self._on_key_enter)
        self.bind("<Escape>", self._on_key_escape)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<FocusIn>", self._on_focus_in)
        # ユーザー操作があれば自動クローズをキャンセル
        for seq in ("<Motion>", "<ButtonPress>", "<KeyPress>"):
            self.bind_all(seq, self._on_user_activity, add=True)
        
        # 添付ファイルがあれば表示を更新
        if self.attached_file_paths:
            self.update_attached_files_display(self.attached_file_paths)

    def _on_open_prompt_manager(self):
        # Open main prompt manager and close this window safely (deferred)
        if getattr(self, '_is_destroying', False):
            return
        try:
            if hasattr(self.agent, '_show_main_window'):
                self.agent._show_main_window()
        except Exception:
            pass
        finally:
            self._close_safely()

    def _close_safely(self):
        # Guard multiple closes and defer actual destroy to avoid Tk race
        if getattr(self, '_is_destroying', False):
            return
        self._is_destroying = True
        try:
            self.withdraw()
        except Exception:
            pass
        try:
            self.after(20, lambda: super(ActionSelectorWindow, self).destroy())
        except Exception:
            pass

    def _on_prompt_selected(self, prompt_id: str, prompt_config: Prompt):
        if not self._is_destroying:
            # 選択された履歴があれば、エージェントに一時入力として渡す
            if self._selected_history_item is not None:
                try:
                    item = self._selected_history_item
                    # 画像が圧縮形式ならここで非圧縮の base64 PNG へ変換しておく
                    if isinstance(item, dict) and item.get('type') == 'image_compressed':
                        try:
                            raw = base64.b64decode(item.get('data', ''))
                            raw = zlib.decompress(raw)
                            # raw は PNG バイト列のはず。再エンコードして 'image' とする
                            b64_png = base64.b64encode(raw).decode('utf-8')
                            item = {"type": "image", "data": b64_png}
                        except Exception:
                            # 失敗時はそのまま渡し、下流でのフォールバックに任せる
                            pass
                    # エージェントに一時的な入力オーバーライドをセット
                    setattr(self.agent, '_temp_input_for_processing', item)
                except Exception:
                    pass
            self.on_prompt_selected_callback(prompt_id=prompt_id, file_paths=self.attached_file_paths)
            self.destroy()

    def _on_key_up(self, event):
        if not self.prompt_buttons:
            return "break"
        self.current_selection_index = (self.current_selection_index - 1) % len(self.prompt_buttons)
        self._update_selection_highlight()
        target = self.prompt_buttons[self.current_selection_index]
        if target and target.winfo_exists():
            target.focus_set()
            self._scroll_prompt_into_view(target)
        return "break"

    def _on_key_down(self, event):
        if not self.prompt_buttons:
            return "break"
        self.current_selection_index = (self.current_selection_index + 1) % len(self.prompt_buttons)
        self._update_selection_highlight()
        target = self.prompt_buttons[self.current_selection_index]
        if target and target.winfo_exists():
            target.focus_set()
            self._scroll_prompt_into_view(target)
        return "break"

    def _on_key_enter(self, event):
        if self.prompt_buttons and self.prompt_buttons[self.current_selection_index].winfo_exists():
            self.prompt_buttons[self.current_selection_index].invoke()
        return "break"

    def _on_key_escape(self, event):
        self.destroy()
        return "break"

    def _on_free_input(self):
        if not self._is_destroying:
            dlg = ResizableInputDialog(self, title=tr("free_input.title"), text=tr("free_input.prompt_label"), agent=self.agent, enable_history=True)
            dlg.show()
            # ユーザーが入力したテキストを取得
            prompt_text = dlg.get_input()
            if prompt_text:
                # モデル名は選択肢の表示から実際のIDを抽出する
                model_name = "gemini-2.5-flash-lite"
                try:
                    if hasattr(dlg, 'model_variable'):
                        display_val = dlg.model_variable.get()
                        # 表示文字列の先頭のスペース区切り部分をモデル名とする
                        # 例: "gemini-2.5-flash-lite (高速、低精度)" -> "gemini-2.5-flash-lite"
                        model_name = display_val.split(" ")[0] if display_val else model_name
                except Exception:
                    pass
                # 温度スライダーから温度を取得
                temperature_val = 1.0
                try:
                    if hasattr(dlg, 'temperature_slider'):
                        temperature_val = float(dlg.temperature_slider.get())
                except Exception:
                    pass
                # 選択されたパラメータでLLM処理を実行
                self.agent._run_process_in_thread(
                    system_prompt=prompt_text,
                    model=model_name,
                    temperature=temperature_val,
                    file_paths=self.attached_file_paths
                )
            # 入力ダイアログを閉じた後は自分自身を閉じる
            self.destroy()

    def _on_file_attach(self):
        if self._is_destroying:
            return
        # 一時的に隠してファイルダイアログの邪魔をしない
        try:
            self.attributes("-topmost", False)
        except Exception:
            pass
        try:
            self.withdraw()
        except Exception:
            pass
        # ファイル選択
        paths = filedialog.askopenfilenames()
        # 再表示（最前面に戻す）
        try:
            self.deiconify()
            self.lift()
            self.attributes("-topmost", True)
        except Exception:
            pass
        # 選択結果の反映（ウィンドウは破棄せず、ラベルのみ更新）
        if paths:
            try:
                from pathlib import Path
                abs_paths = [str(Path(p).resolve()) for p in paths]
            except Exception:
                abs_paths = list(paths)
            # エージェントに一時パスをセット（後続の実行で使用）
            try:
                setattr(self.agent, '_temp_file_paths_for_processing', abs_paths)
            except Exception:
                pass
            # 自身の表示も更新
            self.attached_file_paths = abs_paths
            self.update_attached_files_display(abs_paths)
            # フォーカスを現在のプロンプトに戻す
            if self.prompt_buttons:
                try:
                    self.prompt_buttons[self.current_selection_index].focus_set()
                except Exception:
                    pass

    def _on_matrix(self):
        if not self._is_destroying:
            self.agent.show_matrix_batch_processor_window()
            self.destroy()

    def _update_selection_highlight(self):
        highlight_border_color = styles.HIGHLIGHT_BORDER_COLOR
        highlight_border_width = styles.HIGHLIGHT_BORDER_WIDTH
        for i, button in enumerate(self.prompt_buttons):
            if i == self.current_selection_index:
                button.configure(border_color=highlight_border_color, border_width=highlight_border_width)
            else:
                button.configure(border_width=0)

    def _get_scroll_canvas(self):
        try:
            canvas = getattr(self.scrollable_frame, "_parent_canvas", None)
        except Exception:
            canvas = None
        if canvas is not None:
            return canvas
        # フォールバック: 子ウィジェットからCanvasを探索
        try:
            import tkinter as tk
            for ch in self.scrollable_frame.winfo_children():
                if isinstance(ch, tk.Canvas):
                    return ch
        except Exception:
            pass
        return None

    def _scroll_prompt_into_view(self, btn):
        try:
            self.update_idletasks()
            canvas = self._get_scroll_canvas()
            if canvas is None:
                return
            # 可視領域（Canvas座標）
            view_top = canvas.canvasy(0)
            view_height = canvas.winfo_height()
            view_bottom = view_top + view_height
            # ボタン上端をCanvas座標に変換
            btn_top_canvas = canvas.canvasy(btn.winfo_rooty() - canvas.winfo_rooty())
            btn_bottom_canvas = btn_top_canvas + btn.winfo_height()
            # スクロール領域
            bbox = canvas.bbox("all")
            if not bbox:
                return
            y0, y1 = bbox[1], bbox[3]
            total_height = max(1, y1 - y0)
            margin = 10
            if btn_top_canvas < view_top:
                new_top = max(0, btn_top_canvas - margin)
                frac = max(0.0, min(1.0, new_top / max(1, total_height - view_height)))
                canvas.yview_moveto(frac)
            elif btn_bottom_canvas > view_bottom:
                new_top = max(0, btn_bottom_canvas - view_height + margin)
                frac = max(0.0, min(1.0, new_top / max(1, total_height - view_height)))
                canvas.yview_moveto(frac)
        except Exception:
            pass

    def show_at_cursor(self, cursor_pos: Optional[tuple] = None):
        self.update_idletasks() # Ensure geometry is up to date

        # スタイル定義から固定のウィンドウサイズを読み込む
        geometry_parts = styles.ACTION_SELECTOR_GEOMETRY.split('x')
        window_width = int(geometry_parts[0])
        window_height = int(geometry_parts[1])

        # 画面サイズを取得
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # マウスカーソル位置を取得
        if cursor_pos:
            cursor_x, cursor_y = cursor_pos
        else:
            cursor_x, cursor_y = self.winfo_pointerx(), self.winfo_pointery()
        
        # ウィンドウをマウスカーソルの右下に配置
        offset_x = 20
        offset_y = 20

        x = cursor_x + offset_x
        y = cursor_y + offset_y

        # 画面の右端からはみ出さないように調整
        if x + window_width > screen_width - 10:
            x = screen_width - window_width - 10

        # 画面の下端からはみ出さないように調整
        if y + window_height > screen_height - 50:
            y = screen_height - window_height - 50

        x = max(10, x)
        y = max(10, y)
        
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 位置設定後にウィンドウを表示
        self.deiconify()
        self.lift()
        self.focus_force()
        if self.prompt_buttons and self.prompt_buttons[0].winfo_exists():
            self.current_selection_index = 0
            self.after(1, lambda: self.prompt_buttons[0].focus_set())

    def destroy(self):
        print("DEBUG: ActionSelectorWindow destroy called.")
        self._is_destroying = True
        # ウィンドウが存在し、まだgrab_setされている場合のみgrab_releaseを呼び出す
        if self.winfo_exists() and self.grab_current() == str(self):
            try:
                self.grab_release()
                print("DEBUG: grab_release called successfully.")
            except tk.TclError as e:
                print(f"WARNING: destroy - grab_release中にTclErrorが発生しました: {e}")
        if self._on_destroy_callback:
            self._on_destroy_callback()
        super().destroy()

    def _on_history_changed(self):
        try:
            label = self.history_variable.get()
            item = self._history_label_to_item.get(label)
            self._selected_history_item = item
            # 画像の場合は編集不可、それ以外は可
            is_image = isinstance(item, dict) and item and item.get('type') in ('image', 'image_compressed')
            state = 'disabled' if is_image else 'normal'
            try:
                self.edit_history_button.configure(state=state)
            except Exception:
                pass
        except Exception:
            self._selected_history_item = None
            try:
                self.edit_history_button.configure(state='disabled')
            except Exception:
                pass

    def _on_edit_history(self):
        """選択中の履歴を編集。テキストはテキスト編集、ファイルはファイル再選択。画像は無効化済み。"""
        item = self._selected_history_item
        if not isinstance(item, dict) or 'type' not in item:
            return
        t = item.get('type')
        if t == 'text':
            # テキスト編集ダイアログ（テキスト専用UI）
            dlg = HistoryEditDialog(self, initial_value=item.get('data',''))
            dlg.show()
            new_text = dlg.get_input()
            if new_text is not None:
                self._selected_history_item = {"type": "text", "data": new_text}
        elif t == 'file':
            # ファイル再選択
            from tkinter import filedialog
            new_path = filedialog.askopenfilename()
            if new_path:
                self._selected_history_item = {"type": "file", "data": new_path}
        else:
            # 画像などはここに来ない（ボタン無効化済み）
            pass

    def update_attached_files_display(self, file_paths: Optional[List[str]]):
        if not self.attached_files_label.winfo_exists(): return
        if file_paths:
            file_names = [Path(p).name for p in file_paths]
            body = ", ".join(file_names)
            # 表示を簡潔に（長すぎる場合は省略）
            prefix = tr("history.file_label_prefix")
            text = prefix + body
            max_len = 60
            if len(text) > max_len:
                text = text[: max_len - 1] + "…"
            self.attached_files_label.configure(text=text)
        else:
            self.attached_files_label.configure(text="")

    def _on_focus_out(self, event):
        # フォーカスが外れたら、少し待ってから本当に外部へ移ったままかを確認し、閉じる
        self._schedule_close_after_delay(800)

    def _on_focus_in(self, event):
        # フォーカスが戻ったら、クローズ予定をキャンセル
        self._cancel_scheduled_close()

    def _on_user_activity(self, event):
        # 何らかのユーザー操作があれば、クローズ予定をキャンセル
        self._cancel_scheduled_close()

    def _schedule_close_after_delay(self, delay_ms: int):
        self._cancel_scheduled_close()
        try:
            self._pending_close_id = self.after(delay_ms, self._close_if_still_inactive)
        except Exception:
            self._pending_close_id = None

    def _cancel_scheduled_close(self):
        if self._pending_close_id is not None:
            try:
                self.after_cancel(self._pending_close_id)
            except Exception:
                pass
            self._pending_close_id = None

    def _close_if_still_inactive(self):
        self._pending_close_id = None
        try:
            # 1) まだこのウィンドウ内にフォーカスがあるなら閉じない
            w = self.focus_get()
            if w is not None and self._is_child_of_self(w):
                return
            # 2) マウスカーソルがウィンドウ内にあるなら閉じない（メニュー操作等の誤判定を緩和）
            px, py = self.winfo_pointerx(), self.winfo_pointery()
            wx, wy = self.winfo_rootx(), self.winfo_rooty()
            ww, wh = self.winfo_width(), self.winfo_height()
            if wx <= px <= wx + ww and wy <= py <= wy + wh:
                return
        except Exception:
            # 例外時は安全側（閉じる）
            pass
        # ここまで来たら非アクティブ状態が継続していると判断して閉じる
        try:
            self.destroy()
        except Exception:
            pass

    def _is_child_of_self(self, widget):
        """指定されたウィジェットがこのウィンドウの子孫であるかを確認する"""
        parent = widget
        while parent is not None:
            if parent == self:
                return True
            parent = parent.master
        return False

class NotificationPopup(ctk.CTkToplevel):
    def __init__(self, title: str, message: str, parent_app, level: Literal["info", "warning", "error", "success"] = "info", on_destroy_callback: Optional[Callable] = None):
        super().__init__(parent_app)
        # self.transient(parent_app) # 親ウィンドウとの連携を強化 (一時的に削除)
        self.title(title)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        # self.attributes("-topmost", True) # 常に最前面表示 (一時的に削除)
        self.title_text = title
        self._on_destroy_callback = on_destroy_callback
        self.level = level
        self._close_after_id = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.content_frame = ctk.CTkFrame(self, corner_radius=8)
        self.content_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.message_label = ctk.CTkTextbox(self.content_frame, font=ctk.CTkFont(size=12), wrap="word", height=150, border_width=1, border_color=styles.HIGHLIGHT_BORDER_COLOR)
        self.message_label.grid(row=0, column=0, padx=10, pady=(10, 10), sticky="nsew")
        self.message_label.insert("0.0", message)
        self.message_label.configure(state="disabled")

        self.attributes("-alpha", styles.NOTIFICATION_POPUP_ALPHA)

        self._set_colors_by_level(self.level)
        self.initial_width = styles.NOTIFICATION_POPUP_INITIAL_WIDTH

    def destroy(self):
        if self._close_after_id:
            self.after_cancel(self._close_after_id)
            self._close_after_id = None
        if self.winfo_exists() and self.grab_current() == str(self):
            self.grab_release()
        if self._on_destroy_callback:
            self._on_destroy_callback()
        super().destroy()

    def reconfigure(self, title: str, message: str, level: Literal["info", "warning", "error", "success"] = "info", duration_ms: Optional[int] = 3000):
        if self._close_after_id:
            self.after_cancel(self._close_after_id)
            self._close_after_id = None

        self.title(title) # OSのタイトルバーにタイトルを設定
        self.message_label.configure(state="normal")
        self.message_label.delete("0.0", "end")
        self.message_label.insert("0.0", message)
        self.message_label.see("end")
        self.message_label.configure(state="disabled")

        self._set_colors_by_level(level)
        self.update_idletasks()
        self._adjust_window_size()

        self.state('normal')
        self.lift()
        self.wm_attributes("-topmost", True) # 最前面表示をここで設定

        if duration_ms is not None:
            self._close_after_id = self.after(duration_ms, self._on_timeout_destroy)

    def update_message(self, new_chunk: str):
        self.message_label.configure(state="normal")
        self.message_label.insert("end", new_chunk)
        self.message_label.see("end")
        self.message_label.configure(state="disabled")
        self.update_idletasks()
        self._adjust_window_size()

    def _set_colors_by_level(self, level: Literal["info", "warning", "error", "success"]):
        self.content_frame.configure(fg_color=styles.NOTIFICATION_COLORS.get(level, styles.NOTIFICATION_COLORS["info"]))

    def _adjust_window_size(self):
        self.update_idletasks()
        # タイトルバーの高さはOSが管理するため、ここではメッセージエリアの高さのみを考慮
        message_height = self.message_label.winfo_reqheight() + 20 # 上下のパディングを考慮
        min_height = message_height + 20 # 適当な余白

        screen_height = self.winfo_screenheight()
        max_height = screen_height // 2
        new_height = min(max(min_height, 150), max_height)
        self.geometry(f"{self.initial_width}x{new_height}+{self.winfo_x()}+{self.winfo_y()}")

    def show_at_cursor(self, title: str, message: str, level: Literal["info", "warning", "error", "success"] = "info", duration_ms: Optional[int] = 3000):
        self.reconfigure(title, message, level, duration_ms)

        # マウスカーソル位置を取得
        cursor_x, cursor_y = self.winfo_pointerx(), self.winfo_pointery()

        # 画面サイズを取得
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # ウィンドウの現在のサイズを取得 (reconfigureで調整済み)
        self.update_idletasks() # 最新のサイズを反映させる
        window_width = self.winfo_width()
        window_height = self.winfo_height()

        # マウスカーソルから少しオフセットを持たせる
        offset_x = 20
        offset_y = 20

        x = cursor_x + offset_x
        y = cursor_y + offset_y

        # ウィンドウが画面の右端からはみ出さないように調整
        if x + window_width > screen_width - 10:  # 右端から10pxの余裕
            x = screen_width - window_width - 10

        # ウィンドウが画面の下端からはみ出さないように調整 (タスクバーなどを考慮して少し余裕を持たせる)
        if y + window_height > screen_height - 50:  # 下端から50pxの余裕
            y = screen_height - window_height - 50

        # 負の値を防ぐ (画面の左上端より内側に表示されるように)
        x = max(10, x)
        y = max(10, y)

        # ウィンドウを配置
        self.geometry(f"{window_width}x{window_height}+{x}+{y}") # サイズと位置を同時に設定
        self.state('normal')  # 最小化状態から戻す
        self.lift()
        self.wm_attributes("-topmost", True) # 最前面表示をここで設定

    def _on_timeout_destroy(self):
        self.destroy()

class PromptEditorDialog(BaseDialog):
    def __init__(self, parent_app, title: str, prompt: Optional[Prompt] = None):
        # Initialize using the common BaseDialog to unify styling and geometry
        super().__init__(parent_app, title=title, geometry=styles.PROMPT_EDITOR_GEOMETRY)
        # Override default close handler to call on_cancel
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        self.result: Optional[Prompt] = None

        self.grid_columnconfigure(1, weight=1)

        # Build UI model options from centralized constants
        self.available_models = [label for _, label in SUPPORTED_MODELS]
        self.model_variable = ctk.StringVar()

        ctk.CTkLabel(self, text=tr("prompt.name"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.name_entry = ctk.CTkEntry(self, width=300, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.name_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self, text=tr("common.model"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.model_optionmenu = ctk.CTkOptionMenu(self, values=self.available_models, variable=self.model_variable, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.model_optionmenu.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.parameter_editor = PromptParameterEditorFrame(self)
        self.parameter_editor.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self, text=tr("prompt.thinking_level"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.thinking_level_optionmenu = ctk.CTkOptionMenu(self, values=["Fast", "Balanced", "High Quality", "Unlimited"], fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.thinking_level_optionmenu.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        # Web 検索の有効/無効
        self.enable_web_var = ctk.BooleanVar(value=False)
        ctk.CTkLabel(self, text=tr("prompt.enable_web"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=4, column=0, padx=10, pady=(0, 10), sticky="w")
        self.enable_web_switch = ctk.CTkSwitch(self, text="", variable=self.enable_web_var)
        self.enable_web_switch.grid(row=4, column=1, padx=10, pady=(0, 10), sticky="w")

        ctk.CTkLabel(self, text=tr("prompt.system_prompt"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.system_prompt_textbox = ctk.CTkTextbox(self, width=480, height=400, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR, border_width=1, border_color=styles.HIGHLIGHT_BORDER_COLOR)
        self.system_prompt_textbox.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_rowconfigure(6, weight=1)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=7, column=0, columnspan=2, pady=10)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(button_frame, text=tr("common.save"), command=self.on_save, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=0, padx=10, pady=5)
        ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self.on_cancel, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR).grid(row=0, column=1, padx=10, pady=5)

        if prompt:
            self.name_entry.insert(0, prompt.name)
            # セントラル定義から表示名へ
            self.model_variable.set(model_id_to_label(prompt.model))
            self.parameter_editor.set_parameters(prompt.parameters)
            self.thinking_level_optionmenu.set(prompt.thinking_level)
            self.enable_web_var.set(getattr(prompt, 'enable_web', False))
            self.system_prompt_textbox.insert("0.0", prompt.system_prompt)
        else:
            self.parameter_editor.set_parameters(PromptParameters())
            self.model_variable.set(self.available_models[1])  # Default to flash
            self.thinking_level_optionmenu.set("Balanced")
            self.enable_web_var.set(False)

    def on_save(self):
        try:
            name = self.name_entry.get()
            model = model_label_to_id(self.model_variable.get())
            parameters = self.parameter_editor.get_parameters()
            thinking_level = self.thinking_level_optionmenu.get()
            enable_web = bool(self.enable_web_var.get())
            system_prompt = self.system_prompt_textbox.get("0.0", "end-1c")

            if not name or not system_prompt or not model:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.validation_missing"), icon="warning")
                return

            self.result = Prompt(
                name=name,
                model=model,
                system_prompt=system_prompt,
                thinking_level=thinking_level,
                enable_web=enable_web,
                parameters=parameters
            )
            self.destroy()
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=f"Unexpected error: {e}", icon="cancel")

    def on_cancel(self):
        # print("DEBUG: PromptEditorDialog on_cancel called. Destroying dialog.")
        self.result = None
        self.destroy()

    def get_result(self) -> Optional[Prompt]:
        # print("DEBUG: PromptEditorDialog get_result called. Waiting for window.")
        self.master.wait_window(self)
        # print("DEBUG: PromptEditorDialog get_result - Window closed. Returning result.")
        return self.result

    def destroy(self):
        # print("DEBUG: PromptEditorDialog destroy called.")
        if self.grab_current() == str(self):
            try:
                self.grab_release()
                # print("DEBUG: PromptEditorDialog grab_release called successfully.")
            except tk.TclError as e:
                print(f"WARNING: PromptEditorDialog destroy - grab_release中にTclErrorが発生しました: {e}")
        super().destroy()



class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent_app, agent):
        super().__init__(parent_app, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        self.title(tr("settings.title"))
        # スタイルからジオメトリを読み込み、サイズを維持して中央に配置
        geometry_parts = styles.SETTINGS_WINDOW_GEOMETRY.split('x')
        width = int(geometry_parts[0])
        height = int(geometry_parts[1])
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.transient(parent_app)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.agent = agent
        self.parent_app = parent_app

        self.grid_columnconfigure(1, weight=1)

        api_key_frame = ctk.CTkFrame(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        api_key_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        api_key_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(api_key_frame, text=tr("settings.api_key"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.api_key_entry = ctk.CTkEntry(api_key_frame, placeholder_text=tr("api.placeholder"), fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.api_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        current_api_key = keyring.get_password(API_SERVICE_ID, "api_key")
        if current_api_key:
            self.api_key_entry.insert(0, "*" * (len(current_api_key) - 4) + current_api_key[-4:])
            self.api_key_entry.configure(state="readonly")

        ctk.CTkButton(api_key_frame, text=tr("common.save"), command=self._save_api_key, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(api_key_frame, text=tr("common.delete"), command=self._delete_api_key, fg_color=styles.DELETE_BUTTON_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=1, column=2, padx=5, pady=5)

        hotkey_frame = ctk.CTkFrame(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        hotkey_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        hotkey_frame.grid_columnconfigure(1, weight=1)
        hotkey_frame.grid_columnconfigure(4, weight=0)

        # Hotkey: Prompt List
        ctk.CTkLabel(hotkey_frame, text=tr("settings.hotkey.list"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.hotkey_prompt_label = ctk.CTkLabel(hotkey_frame, text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_prompt_list', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.hotkey_prompt_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.btn_set_prompt = ctk.CTkButton(hotkey_frame, text=tr("common.edit"), width=90, command=lambda: self._set_hotkey('prompt_list'), fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.btn_set_prompt.grid(row=0, column=2, padx=5, pady=5)
        self.btn_clear_prompt = ctk.CTkButton(hotkey_frame, text=tr("common.disable"), width=90, command=lambda: self._clear_hotkey('prompt_list'), fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        self.btn_clear_prompt.grid(row=0, column=3, padx=5, pady=5)

        # Hotkey: Refine
        ctk.CTkLabel(hotkey_frame, text=tr("settings.hotkey.refine"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.hotkey_refine_label = ctk.CTkLabel(hotkey_frame, text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_refine', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.hotkey_refine_label.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.btn_set_refine = ctk.CTkButton(hotkey_frame, text=tr("common.edit"), width=90, command=lambda: self._set_hotkey('refine'), fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.btn_set_refine.grid(row=1, column=2, padx=5, pady=5)
        self.btn_clear_refine = ctk.CTkButton(hotkey_frame, text=tr("common.disable"), width=90, command=lambda: self._clear_hotkey('refine'), fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        self.btn_clear_refine.grid(row=1, column=3, padx=5, pady=5)

        # Hotkey: Matrix
        ctk.CTkLabel(hotkey_frame, text=tr("settings.hotkey.matrix"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.hotkey_matrix_label = ctk.CTkLabel(hotkey_frame, text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_matrix', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.hotkey_matrix_label.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.btn_set_matrix = ctk.CTkButton(hotkey_frame, text=tr("common.edit"), width=90, command=lambda: self._set_hotkey('matrix'), fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR)
        self.btn_set_matrix.grid(row=2, column=2, padx=5, pady=5)
        self.btn_clear_matrix = ctk.CTkButton(hotkey_frame, text=tr("common.disable"), width=90, command=lambda: self._clear_hotkey('matrix'), fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        self.btn_clear_matrix.grid(row=2, column=3, padx=5, pady=5)

        # Language selector (below hotkeys)
        lang_frame = ctk.CTkFrame(self, fg_color=styles.HISTORY_ITEM_FG_COLOR)
        lang_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        lang_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(lang_frame, text=tr("settings.language"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        try:
            names = available_locales()
        except Exception:
            names = {"auto": "Auto", "en": "English", "ja": "日本語"}
        self._lang_codes = list(names.keys())
        values = list(names.values())
        self._lang_var = ctk.StringVar()
        self._lang_menu = ctk.CTkOptionMenu(lang_frame, values=values, variable=self._lang_var)
        cur_lang = getattr(self.agent.config, 'language', 'auto') or 'auto'
        if cur_lang not in self._lang_codes:
            cur_lang = 'auto'
        try:
            self._lang_var.set(values[self._lang_codes.index(cur_lang)])
        except Exception:
            self._lang_var.set(values[0])
        self._lang_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(self, text=tr("settings.history.max"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.max_history_entry = ctk.CTkEntry(self, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.max_history_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.max_history_entry.insert(0, str(self.agent.config.max_history_size))

        # Flow settings
        ctk.CTkLabel(self, text=tr("settings.flow.max_steps"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.max_flow_steps_entry = ctk.CTkEntry(self, fg_color=styles.HISTORY_ITEM_FG_COLOR, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.max_flow_steps_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
        try:
            self.max_flow_steps_entry.insert(0, str(getattr(self.agent.config, 'max_flow_steps', 5)))
        except Exception:
            self.max_flow_steps_entry.insert(0, "5")

        ctk.CTkButton(self, text=tr("settings.save_settings"), command=self._save_settings, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=5, column=0, columnspan=2, padx=10, pady=20)

    def _save_api_key(self):
        new_api_key = self.api_key_entry.get()
        if new_api_key.startswith("****"):
            CTkMessagebox(title=tr("api.title"), message=tr("settings.save_done_message"), icon="info").wait_window()
            return
        if new_api_key:
            try:
                keyring.set_password(API_SERVICE_ID, "api_key", new_api_key)
                self.agent.api_key = new_api_key
                genai.configure(api_key=self.agent.api_key)
                CTkMessagebox(title=tr("common.success"), message=tr("settings.save_done_message"), icon="info").wait_window()
                self.api_key_entry.delete(0, ctk.END)
                self.api_key_entry.insert(0, "*" * (len(new_api_key) - 4) + new_api_key[-4:])
                self.api_key_entry.configure(state="readonly")
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("api.save_failed", details=str(e)), icon="cancel").wait_window()
        else:
            CTkMessagebox(title=tr("common.error"), message=tr("api.enter_key"), icon="warning").wait_window()

    def _delete_api_key(self):
        confirm = CTkMessagebox(title=tr("api.confirm_delete_title"), message=tr("api.confirm_delete_message"), icon="question", option_1=tr("common.cancel"), option_2=tr("common.delete")).get()
        if confirm == tr("common.delete"):
            try:
                keyring.delete_password(API_SERVICE_ID, "api_key")
                self.agent.api_key = None
                CTkMessagebox(title=tr("common.success"), message=tr("api.deleted"), icon="info").wait_window()
                self.api_key_entry.configure(state="normal")
                self.api_key_entry.delete(0, ctk.END)
                self.api_key_entry.insert(0, "")
                self.api_key_entry.focus_set()
            except keyring.errors.NoKeyringError:
                CTkMessagebox(title=tr("common.error"), message=tr("api.keyring_missing"), icon="error").wait_window()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("api.delete_failed", details=str(e)), icon="cancel").wait_window()

    def _fmt_hotkey(self, value: Optional[str]) -> str:
        return value if value else tr("common.unspecified")

    def _set_hotkey(self, target: str):
        # Update target label to prompt user
        if target == 'prompt_list':
            self.hotkey_prompt_label.configure(text=tr("hotkey.press_new"))
        elif target == 'refine':
            self.hotkey_refine_label.configure(text=tr("hotkey.press_new"))
        elif target == 'matrix':
            self.hotkey_matrix_label.configure(text=tr("hotkey.press_new"))
        self.update_idletasks()
        # Disable buttons during capture
        self._set_hotkey_buttons_state('disabled')
        import threading
        def _capture():
            try:
                recorded_hotkey = keyboard.read_hotkey(suppress=False)
            except Exception as e:
                self.after(0, lambda: self._on_hotkey_capture_error(e))
                return
            self.after(0, lambda: self._on_hotkey_captured(target, recorded_hotkey))
        threading.Thread(target=_capture, daemon=True).start()

    def _on_hotkey_capture_error(self, e: Exception):
        CTkMessagebox(title=tr("common.error"), message=tr("hotkey.read_failed", details=str(e)), icon="cancel").wait_window()
        self._refresh_hotkey_labels()
        self._set_hotkey_buttons_state('normal')

    def _on_hotkey_captured(self, target: str, recorded_hotkey: str):
        # Preview selection on proper label
        if target == 'prompt_list':
            self.hotkey_prompt_label.configure(text=recorded_hotkey)
        elif target == 'refine':
            self.hotkey_refine_label.configure(text=recorded_hotkey)
        elif target == 'matrix':
            self.hotkey_matrix_label.configure(text=recorded_hotkey)
        self.update_idletasks()

        msg_box = CTkMessagebox(
            title=tr("hotkey.confirm_title"),
            message=tr("hotkey.confirm_message", hotkey=recorded_hotkey),
            icon="question",
            option_1=tr("common.cancel"),
            option_2=tr("common.yes")
        )
        response = msg_box.get()
        if response == tr("common.yes"):
            if self.agent.update_hotkey(target, recorded_hotkey):
                CTkMessagebox(title=tr("common.success"), message=tr("hotkey.updated"), icon="info").wait_window()
            else:
                CTkMessagebox(title=tr("common.error"), message=tr("hotkey.update_failed"), icon="cancel").wait_window()
                self._refresh_hotkey_labels()
        else:
            self._refresh_hotkey_labels()
        self._set_hotkey_buttons_state('normal')

    def _clear_hotkey(self, target: str):
        try:
            if self.agent.update_hotkey(target, None):
                CTkMessagebox(title=tr("common.success"), message=tr("hotkey.disabled"), icon="info").wait_window()
            else:
                CTkMessagebox(title=tr("common.error"), message=tr("hotkey.disable_failed"), icon="cancel").wait_window()
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=tr("hotkey.disable_error", details=str(e)), icon="cancel").wait_window()
        self._refresh_hotkey_labels()

    def _set_hotkey_buttons_state(self, state: str):
        for b in (
            self.btn_set_prompt, self.btn_clear_prompt,
            self.btn_set_refine, self.btn_clear_refine,
            self.btn_set_matrix, self.btn_clear_matrix,
        ):
            try:
                b.configure(state=state)
            except Exception:
                pass

    def _refresh_hotkey_labels(self):
        self.hotkey_prompt_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_prompt_list', None)))
        self.hotkey_refine_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_refine', None)))
        self.hotkey_matrix_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_matrix', None)))

    def _save_settings(self):
        try:
            new_max_history_size = int(self.max_history_entry.get())
            if new_max_history_size <= 0:
                raise ValueError(tr("settings.history.max_invalid"))
            self.agent.config.max_history_size = new_max_history_size
            self.agent.max_history_size = new_max_history_size
            try:
                new_max_flow_steps = int(self.max_flow_steps_entry.get())
                if new_max_flow_steps <= 0:
                    raise ValueError(tr("settings.flow.max_steps_invalid"))
                self.agent.config.max_flow_steps = new_max_flow_steps
            except Exception as e:
                raise ValueError(e)
            # 言語を保存
            try:
                # CTkOptionMenu内部の値配列からインデックス取得
                idx = self._lang_menu._values.index(self._lang_var.get())  # type: ignore[attr-defined]
                self.agent.config.language = getattr(self, '_lang_codes', ['auto'])[idx]
            except Exception:
                self.agent.config.language = 'auto'
            # 反映
            try:
                set_locale(self.agent.config.language)
                try:
                    self.parent_app.title(tr("app.title"))
                except Exception:
                    pass
            except Exception:
                pass
            save_config(self.agent.config)
            if len(self.agent.clipboard_history) > self.agent.max_history_size:
                self.agent.clipboard_history = self.agent.clipboard_history[-self.agent.max_history_size:]
            if self.agent._on_history_updated_callback:
                self.agent._on_history_updated_callback(self.agent.clipboard_history)
            CTkMessagebox(title=tr("settings.save_done_title"), message=tr("settings.save_done_message"), icon="info").wait_window()
        except ValueError as e:
            CTkMessagebox(title=tr("common.error"), message=f"Invalid value: {e}", icon="warning").wait_window()
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=f"Save failed: {e}", icon="cancel").wait_window()

    def on_close(self):
        self.grab_release()
        self.destroy()

class ResizableInputDialog(BaseDialog):
    def __init__(self, parent_app, title: str, text: str, initial_value: str = "", agent: Optional[Any] = None, enable_history: bool = False):
        # Use BaseDialog for unified styling; set a default geometry with minimum size
        super().__init__(parent_app, title=title, geometry="500x300")
        # Override close protocol to call our own handler
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.result: Optional[str] = None
        self._agent = agent
        self._enable_history = enable_history

        # 最小サイズを設定
        initial_width = 500
        initial_height = 300
        self.minsize(initial_width, initial_height)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # テキストボックスが拡大するように設定

        # ラベル
        self.label = ctk.CTkLabel(self, text=text, wraplength=initial_width - 40, text_color=styles.POPUP_TEXT_COLOR)
        self.label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # テキスト入力エリア (CTkTextboxを使用)
        self.textbox = ctk.CTkTextbox(self, height=150, width=initial_width - 40, border_width=1, border_color=styles.HIGHLIGHT_BORDER_COLOR)  # 初期高さを設定
        self.textbox.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        self.textbox.insert("0.0", initial_value)
        self.textbox.focus_set()

        # パラメータフレーム (モデル選択と温度設定)
        parameter_frame = ctk.CTkFrame(self, fg_color="transparent")
        parameter_frame.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        parameter_frame.grid_columnconfigure(0, weight=1)
        parameter_frame.grid_columnconfigure(1, weight=1)

        # モデル選択
        ctk.CTkLabel(parameter_frame, text=tr("common.model"), text_color=styles.POPUP_TEXT_COLOR).grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
        # 利用可能なモデル（表示用とAPI用を対応付ける）
        self.available_models = [
            "gemini-2.5-flash-lite (高速、低精度)",
            "gemini-2.5-flash (普通)",
            "gemini-2.5-pro (低速、高精度)"
        ]
        # 初期値として最初のモデルを選択
        self.model_variable = ctk.StringVar(value=self.available_models[0])
        self.model_optionmenu = ctk.CTkOptionMenu(parameter_frame, values=self.available_models, variable=self.model_variable)
        self.model_optionmenu.grid(row=0, column=1, pady=5, sticky="ew")

        # 温度設定
        ctk.CTkLabel(parameter_frame, text=tr("params.temperature"), text_color=styles.POPUP_TEXT_COLOR).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="w")
        temp_control_frame = ctk.CTkFrame(parameter_frame, fg_color="transparent")
        temp_control_frame.grid(row=1, column=1, pady=5, sticky="ew")
        temp_control_frame.grid_columnconfigure(0, weight=1)
        # スライダーと表示ラベル
        self.temperature_slider = ctk.CTkSlider(temp_control_frame, from_=0, to=2, number_of_steps=20)
        self.temperature_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.temperature_slider.set(1.0)
        # 表示ラベル
        self.temperature_value_label = ctk.CTkLabel(temp_control_frame, width=4, text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.temperature_value_label.grid(row=0, column=1, padx=5)
        # スライダー値変更時にラベルを更新する関数
        def _update_temp_label(value):
            try:
                self.temperature_value_label.configure(text=f"{float(value):.1f}")
            except Exception:
                pass
        self.temperature_slider.configure(command=_update_temp_label)
        # 初期値ラベル更新
        _update_temp_label(self.temperature_slider.get())

        # ボタンフレーム
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        # このフレームは縦方向に伸縮するが、横方向にはその子ウィジェットに応じてサイズが決定される
        button_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="ew")
        # OK とキャンセルボタンが均等にスペースを使用するように、各列のweightを1に設定
        button_frame.grid_columnconfigure((0, 1), weight=1)

        # 履歴から挿入（任意）
        # 履歴選択用のウィジェットと挿入ボタンは独立したフレームに配置し、固定幅にする
        if self._enable_history and self._agent is not None:
            history_items: list[str] = []
            try:
                # 直近50件の履歴を収集（テキストのみ）
                for it in getattr(self._agent, 'clipboard_history', [])[:50]:
                    if isinstance(it, str):
                        history_items.append(it)
                    elif isinstance(it, dict) and it.get('type') == 'text':
                        history_items.append(it.get('data', ''))
            except Exception:
                history_items = []

            # ラベルを整形し、無い場合でもプレースホルダーを用意する
            def _labelize(s: str) -> str:
                s1 = s.replace('\n', ' ')
                return s1[:60] + ('…' if len(s1) > 60 else '')
            if history_items:
                labels = [_labelize(s) for s in history_items]
                # 実値をマッピング
                self._history_map = {lbl: val for lbl, val in zip(labels, history_items)}
                no_history = False
            else:
                # 履歴がない場合でもレイアウトを保つためプレースホルダーを設定
                labels = [tr("history.empty")]
                self._history_map = {labels[0]: ""}
                no_history = True

            # 履歴メニューと挿入ボタンを含むサブフレーム
            # このフレームはボタンフレームの横幅いっぱいに広がり、その内部でオプションメニューが余白分を埋める
            history_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
            history_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
            # オプションメニューが残りスペースを埋めるように列のweightを設定
            history_frame.grid_columnconfigure(0, weight=1)
            history_frame.grid_columnconfigure(1, weight=0)

            # 履歴選択用ドロップダウン
            # widthを指定して最小幅を確保しつつ、sticky="ew"で横方向に拡大させる
            self.history_menu = ctk.CTkOptionMenu(history_frame, values=labels, width=240)
            self.history_menu.grid(row=0, column=0, padx=(0,10), sticky="ew")

            # 挿入ボタンは右端に配置し、固定幅とする
            insert_btn = ctk.CTkButton(history_frame, text=tr("history.insert"), width=100, command=self._insert_selected_history)
            insert_btn.grid(row=0, column=1, sticky="e")

            # 履歴が無い場合はメニューとボタンを無効化
            if no_history:
                try:
                    self.history_menu.configure(state="disabled")
                    insert_btn.configure(state="disabled")
                except Exception:
                    pass

        # OKボタンとキャンセルボタンは下段に並べる
        self.ok_button = ctk.CTkButton(button_frame, text=tr("common.ok"), command=self._on_ok)
        self.ok_button.grid(row=1, column=0, padx=(0, 10), sticky="ew")

        self.cancel_button = ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self._on_cancel)
        self.cancel_button.grid(row=1, column=1, padx=(10, 0), sticky="ew")

    def show(self):
        self.update_idletasks() # 最新のサイズを反映させる

        # ウィンドウの幅と高さを取得
        window_width = self.winfo_width()
        window_height = self.winfo_height()

        # 画面サイズを取得
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # マウスカーソル位置を取得
        cursor_x, cursor_y = self.winfo_pointerx(), self.winfo_pointery()

        # ウィンドウをマウスカーソルの右下に配置
        # マウスカーソルから少しオフセットを持たせる
        offset_x = 20
        offset_y = 20

        x = cursor_x + offset_x
        y = cursor_y + offset_y

        # 画面の右端からはみ出さないように調整
        if x + window_width > screen_width - 10:
            x = screen_width - window_width - 10

        # 画面の下端からはみ出さないように調整 (タスクバーなどを考慮して少し余裕を持たせる)
        if y + window_height > screen_height - 50:
            y = screen_height - window_height - 50

        # 負の値を防ぐ (画面の左上端より内側に表示されるように)
        x = max(10, x)
        y = max(10, y)

        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.deiconify() # ウィンドウを表示する
        self.lift()
        self.focus_force()
        self.grab_set() # モーダルにするためにgrab_setを呼び出す
        self.wait_window(self) # ウィンドウが閉じるまで待機

    def _on_ok(self):
        self.result = self.textbox.get("0.0", "end-1c")
        self.destroy()

    def _insert_selected_history(self):
        try:
            if hasattr(self, 'history_menu') and hasattr(self, '_history_map'):
                lbl = self.history_menu.get()
                val = self._history_map.get(lbl, '')
                if val:
                    self.textbox.insert('insert', val)
        except Exception:
            pass

    def _on_cancel(self):
        self.result = None
        self.destroy()

    def _on_closing(self):
        self.result = None
        self.destroy()

    def get_input(self) -> Optional[str]:
        return self.result

    def destroy(self):
        if self.grab_current() == str(self):
            try:
                self.grab_release()
            except tk.TclError as e:
                print(f"WARNING: ResizableInputDialog destroy - grab_release中にTclErrorが発生しました: {e}")
        self.attributes("-topmost", False) # 最前面表示を解除
        self.attributes("-topmost", False) # 最前面表示を解除
        super().destroy()


class MatrixSummarySettingsDialog(BaseDialog):
    """行/列/行列まとめプロンプトの編集用ダイアログ。"""
    def __init__(self, parent_app, agent):
        super().__init__(parent_app, title="マトリクス設定", geometry="520x420")
        self.agent = agent
        self.grid_columnconfigure(1, weight=1)

        def row(prompt):
            frame = ctk.CTkFrame(self, fg_color="transparent")
            frame.grid_columnconfigure(0, weight=0)
            frame.grid_columnconfigure(1, weight=1)
            frame.grid_columnconfigure(2, weight=0)
            return frame

        # 行まとめ
        rframe = row(self)
        rframe.grid(row=0, column=0, columnspan=2, padx=10, pady=(10,5), sticky="ew")
        ctk.CTkLabel(rframe, text=tr("matrix.row_summary"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=(0,8), sticky="w")
        self.row_label = ctk.CTkLabel(rframe, text=self._prompt_title(getattr(agent.config, 'matrix_row_summary_prompt', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.row_label.grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(rframe, text=tr("common.edit"), width=60, command=self._edit_row).grid(row=0, column=2, padx=(8,0))

        # 列まとめ
        cframe = row(self)
        cframe.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(cframe, text=tr("matrix.col_summary"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=(0,8), sticky="w")
        self.col_label = ctk.CTkLabel(cframe, text=self._prompt_title(getattr(agent.config, 'matrix_col_summary_prompt', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.col_label.grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(cframe, text=tr("common.edit"), width=60, command=self._edit_col).grid(row=0, column=2, padx=(8,0))

        # 行列まとめ
        mframe = row(self)
        mframe.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(mframe, text=tr("matrix.matrix_summary"), text_color=styles.HISTORY_ITEM_TEXT_COLOR).grid(row=0, column=0, padx=(0,8), sticky="w")
        self.matrix_label = ctk.CTkLabel(mframe, text=self._prompt_title(getattr(agent.config, 'matrix_matrix_summary_prompt', None)), text_color=styles.HISTORY_ITEM_TEXT_COLOR)
        self.matrix_label.grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(mframe, text=tr("common.edit"), width=60, command=self._edit_matrix).grid(row=0, column=2, padx=(8,0))

        # フッター
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=(10,10), sticky="ew")
        btn_frame.grid_columnconfigure((0,1), weight=1)
        ctk.CTkButton(btn_frame, text=tr("common.close"), command=self.on_close, fg_color=styles.DEFAULT_BUTTON_FG_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR).grid(row=0, column=1, padx=5, sticky="e")

    def _prompt_title(self, p: Optional[Prompt]) -> str:
        return p.name if isinstance(p, Prompt) else tr("common.unspecified")

    def _edit_row(self):
        self._edit_target('matrix_row_summary_prompt', tr("matrix.row_edit_title"))

    def _edit_col(self):
        self._edit_target('matrix_col_summary_prompt', tr("matrix.col_edit_title"))

    def _edit_matrix(self):
        self._edit_target('matrix_matrix_summary_prompt', tr("matrix.matrix_edit_title"))

    def _edit_target(self, attr: str, title: str):
        current = getattr(self.agent.config, attr, None)
        dlg = PromptEditorDialog(self, title=title, prompt=current)
        result = dlg.get_result()
        if result:
            setattr(self.agent.config, attr, result)
            save_config(self.agent.config)
            # ラベル更新
            if attr.endswith('row_summary_prompt'):
                self.row_label.configure(text=self._prompt_title(result))
            elif attr.endswith('col_summary_prompt'):
                self.col_label.configure(text=self._prompt_title(result))
            elif attr.endswith('matrix_summary_prompt'):
                self.matrix_label.configure(text=self._prompt_title(result))

    def on_close(self):
        self.destroy()

    def get_input(self) -> Optional[str]:
        self.master.wait_window(self)
        return self.result
