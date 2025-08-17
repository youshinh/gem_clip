# app.py
import customtkinter as ctk
import tkinter as tk
from CTkMessagebox import CTkMessagebox
# import pyperclip
from typing import List, Optional

from agent import ClipboardToolAgent
from i18n import tr, set_locale
from constants import APP_NAME
import styles
from ui_components import PromptEditorDialog
from config_manager import save_config, load_config
from common_models import Prompt

class ClipboardToolApp:
    def __init__(self):
        ctk.set_appearance_mode("System")
        self.app = ctk.CTk()
        self.app.withdraw()

        try:
            _cfg = load_config()
            if _cfg: set_locale(getattr(_cfg, 'language', 'auto'))
        except Exception:
            set_locale('auto')

        self.app.title(tr("app.title"))
        self.app.configure(fg_color=styles.APP_BG_COLOR)
        self._set_window_icon()

        self.app.update_idletasks()
        width, height = [int(x) for x in styles.MAIN_WINDOW_GEOMETRY.split('x')]
        x = (self.app.winfo_screenwidth() // 2) - (width // 2)
        y = (self.app.winfo_screenheight() // 2) - (height // 2)
        self.app.geometry(f"{width}x{height}+{x}+{y}")

        self.app.grid_columnconfigure(0, weight=1)
        self.app.grid_rowconfigure(1, weight=1)

        self.agent = ClipboardToolAgent()
        self.agent.set_ui_elements(self.app)

        self._setup_ui()

        self.app.protocol("WM_DELETE_WINDOW", self.app.withdraw)
        self._row_frames: List[ctk.CTkFrame] = []
        self._drag_data = {}
        self._drag_active_frame: Optional[ctk.CTkFrame] = None
        self._row_drop_indicator_widget: Optional[tk.Frame] = None

        self._create_prompt_list_frame()

    def _setup_ui(self):
        # --- Matrix Summary Settings Section ---
        self.matrix_section = ctk.CTkFrame(self.app, fg_color=styles.FRAME_BG_COLOR, corner_radius=8)
        self.matrix_section.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="ew")
        self.matrix_section.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.matrix_section, text=tr("matrix.section.title"), font=styles.FONT_BOLD).grid(row=0, column=0, columnspan=3, padx=10, pady=(5,10), sticky="w")

        def create_matrix_row(label_text, prompt_attr, edit_cmd, row):
            ctk.CTkLabel(self.matrix_section, text=label_text).grid(row=row, column=0, padx=(10, 8), pady=2, sticky="w")
            label = ctk.CTkLabel(self.matrix_section, text=self._prompt_title(getattr(self.agent.config, prompt_attr, None)), text_color=styles.TEXT_SECONDARY_COLOR, anchor="w")
            label.grid(row=row, column=1, padx=0, pady=2, sticky="ew")
            ctk.CTkButton(self.matrix_section, text=tr("common.edit"), width=60, command=lambda: edit_cmd(prompt_attr, label)).grid(row=row, column=2, padx=(8, 10), pady=2)
            return label

        self.row_summary_label = create_matrix_row(tr("matrix.row_summary"), 'matrix_row_summary_prompt', self._edit_matrix_prompt, 1)
        self.col_summary_label = create_matrix_row(tr("matrix.col_summary"), 'matrix_col_summary_prompt', self._edit_matrix_prompt, 2)
        self.matrix_summary_label = create_matrix_row(tr("matrix.matrix_summary"), 'matrix_matrix_summary_prompt', self._edit_matrix_prompt, 3)

        # --- Prompt List Frame ---
        self.prompt_list_frame = ctk.CTkFrame(self.app, fg_color=styles.FRAME_BG_COLOR, corner_radius=8)
        self.prompt_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.prompt_list_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.prompt_list_frame, text=tr("prompt.list.title"), font=styles.FONT_BOLD).grid(row=0, column=0, padx=10, pady=(5,10), sticky="w")

        self.prompt_scroll_frame = ctk.CTkScrollableFrame(self.prompt_list_frame, fg_color="transparent")
        self.prompt_scroll_frame.grid(row=1, column=0, sticky="nsew")
        self.prompt_scroll_frame.grid_columnconfigure(0, weight=1)
        self.prompt_list_frame.grid_rowconfigure(1, weight=1)

        self.prompt_scroll_frame.bind("<B1-Motion>", self._on_row_motion)
        self.prompt_scroll_frame.bind("<ButtonRelease-1>", self._on_row_release)

        # --- Button Frame ---
        button_frame = ctk.CTkFrame(self.app, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(button_frame, text=tr("app.prompts.add"), command=self._add_prompt).grid(row=0, column=0, padx=(0,5), pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text=tr("app.prompts.save"), command=self._save_settings).grid(row=0, column=1, padx=(5,0), pady=5, sticky="ew")

    def _set_window_icon(self) -> None:
        """Set window/taskbar icon to icon.ico where supported.

        - Windows: use iconbitmap with .ico directly.
        - Others: convert .ico to PNG and set via iconphoto.
        """
        try:
            import sys, os
            if sys.platform.startswith("win"):
                # Windows: _config.ico を使用
                icon_path = os.path.join(os.path.dirname(__file__), "_config.ico")
                try:
                    self.app.iconbitmap(icon_path)
                except Exception:
                    pass
        except Exception:
            pass

    def _edit_matrix_prompt(self, attr: str, label_widget: ctk.CTkLabel):
        current = getattr(self.agent.config, attr, None)
        dialog = PromptEditorDialog(self.app, title=tr("prompt.edit_dialog.title"), prompt=current)
        updated = dialog.get_result()
        if updated and self.agent.config:
            try:
                setattr(self.agent.config, attr, updated)
                save_config(self.agent.config)
                # refresh label
                label_widget.configure(text=updated.name if isinstance(updated, Prompt) else tr("common.unspecified"))
                self.agent.notify_prompts_changed()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("matrix.summary_save_failed", details=str(e)), icon="cancel")

    def _create_prompt_list_frame(self):
        for widget in self.prompt_scroll_frame.winfo_children():
            widget.destroy()
        self._row_frames = []

        if not (self.agent.config and self.agent.config.prompts):
            ctk.CTkLabel(self.prompt_scroll_frame, text=tr("prompt.none"), text_color=styles.TEXT_SECONDARY_COLOR).grid(row=0, column=0, padx=10, pady=10)
            return

        for row_num, (prompt_id, prompt_config) in enumerate(self.agent.config.prompts.items()):
            row_frame = ctk.CTkFrame(self.prompt_scroll_frame, fg_color="transparent")
            setattr(row_frame, "_prompt_id", prompt_id)
            row_frame.grid(row=row_num, column=0, sticky="ew", padx=5, pady=3)
            row_frame.grid_columnconfigure(1, weight=1)
            self._row_frames.append(row_frame)

            drag_handle = ctk.CTkLabel(row_frame, text="≡", width=18, anchor="center", text_color=styles.TEXT_SECONDARY_COLOR)
            drag_handle.grid(row=0, column=0, padx=(6, 4), pady=5)

            name_label = ctk.CTkLabel(row_frame, text=prompt_config.name, anchor="w", text_color=styles.TEXT_COLOR)
            name_label.grid(row=0, column=1, padx=6, pady=5, sticky="ew")

            var = ctk.BooleanVar(value=getattr(prompt_config, 'include_in_matrix', False))
            checkbox = ctk.CTkCheckBox(row_frame, text=tr("prompt.include_in_matrix"), variable=var, command=lambda p_id=prompt_id, v=var: self._toggle_prompt_matrix(p_id, v.get()), text_color=styles.TEXT_SECONDARY_COLOR)
            checkbox.grid(row=0, column=2, padx=5, pady=5, sticky="w")

            edit_button = ctk.CTkButton(row_frame, text=tr("common.edit"), width=60, command=lambda p_id=prompt_id: self._edit_prompt(p_id))
            edit_button.grid(row=0, column=3, padx=5, pady=5, sticky="e")

            delete_button = ctk.CTkButton(row_frame, text=tr("common.delete"), width=60, fg_color=styles.DESTRUCTIVE_COLOR, hover_color=styles.DESTRUCTIVE_HOVER_COLOR, command=lambda p_id=prompt_id: self._delete_prompt(p_id))
            delete_button.grid(row=0, column=4, padx=5, pady=5, sticky="e")

            for widget in (drag_handle, name_label, checkbox, edit_button, delete_button):
                widget.bind("<ButtonPress-1>", lambda event, frame=row_frame: self._on_row_press(event, frame))
                widget.bind("<B1-Motion>", self._on_row_motion)
                widget.bind("<ButtonRelease-1>", self._on_row_release)

    def _on_row_press(self, event):
        """Record the frame and y-coordinate when a row press starts for drag-and-drop reordering."""
        try:
            # 既存の境界線が残っていたら念のため除去
            try:
                if self._row_drop_indicator_widget is not None and self._row_drop_indicator_widget.winfo_exists():
                    self._row_drop_indicator_widget.destroy()
                self._row_drop_indicator_widget = None
            except Exception:
                self._row_drop_indicator_widget = None

            # Y座標から直接インデックスを計算
            idx = self._compute_drop_index(event.y_root)

            # 計算されたインデックスが有効範囲内かチェック
            if not (0 <= idx < len(self._row_frames)):
                return

            target_frame = self._row_frames[idx]
            self._drag_data = {"frame": target_frame, "index": idx, "current_index": idx}

            # ドラッグ中の行をハイライト
            if self._drag_active_frame:
                self._drag_active_frame.configure(fg_color="transparent")
            self._drag_active_frame = target_frame
            self._drag_active_frame.configure(fg_color=styles.DRAG_ACTIVE_ROW_COLOR)
            
            return "break" # イベントの伝播を停止させない

        except Exception as e:
            return None # エラー時もNoneを返すことで、イベント処理を継続させない

    def _on_row_release(self, event):
        """Handle drop event: compute target index and reorder prompts."""
        # まずは境界線を即座に消す（体感を軽くする）
        try:
            canvas = self._get_prompt_list_canvas()
            if canvas is not None and self._row_drop_line_id is not None:
                canvas.delete(self._row_drop_line_id)
        except Exception:
            pass
        self._row_drop_line_id = None
        try:
            if self._row_drop_indicator_widget is not None and self._row_drop_indicator_widget.winfo_exists():
                self._row_drop_indicator_widget.destroy()
        except Exception:
            pass
        self._row_drop_indicator_widget = None

        if not self._drag_data:
            return

        # Finalize visual order to config order using frame-id mapping
        try:
            if self.agent.config and self._row_frames:
                # Build current visual order of prompt IDs
                visual_ids = [getattr(f, "_prompt_id", None) for f in self._row_frames]
                # Filter out any Nones just in case
                visual_ids = [pid for pid in visual_ids if pid is not None]
                # Current config order
                current_ids = list(self.agent.config.prompts.keys())
                if visual_ids != current_ids:
                    # Reconstruct prompts dict preserving new order
                    new_prompts = {pid: self.agent.config.prompts[pid] for pid in visual_ids if pid in self.agent.config.prompts}
                    # In case some prompts exist in config but no frame (shouldn't happen), append them
                    for pid in current_ids:
                        if pid not in new_prompts:
                            new_prompts[pid] = self.agent.config.prompts[pid]
                    self.agent.config.prompts = new_prompts
                    # Persist once on drop to avoid flicker and extra IO
                    save_config(self.agent.config)
                    self.agent.notify_prompts_changed()
        except Exception:
            pass

        # Reset drag state and remove highlights
        for row_f in self._row_frames:
            try:
                row_f.configure(fg_color="transparent")
            except Exception:
                pass
        self._drag_data = {}
        self._drag_active_frame = None
        # ここまでに境界線は除去済み
        # Do not rebuild the list to avoid flicker; frames are already repositioned during drag
        
    def _on_row_motion(self, event):
        """
        Handle drag motion: update potential drop target highlighting while dragging.
        """
        if not self._drag_data:
            return

        dragged_frame = self._drag_data["frame"]
        start_index = self._drag_data["index"]
        current_index = self._drag_data["current_index"]
        
        # Compute tentative new index based on current mouse y position
        new_index = self._compute_drop_index(event.y_root)

        # If the target index has changed, update the internal list and reposition frames
        if current_index != new_index:
            # Update the internal list of frames
            # This logic needs to move the actual frame objects within _row_frames
            # to reflect the real-time visual reordering.
            
            # Remove the dragged frame from its current position
            popped_frame = self._row_frames.pop(current_index)
            # Insert it at the new potential position
            self._row_frames.insert(new_index, popped_frame)
            
            # Update the current index in drag_data
            self._drag_data["current_index"] = new_index
            
            # Reposition all frames visually
            self._reposition_row_frames()

            # Draw boundary indicator between rows (white horizontal line)
            self._draw_row_drop_indicator(event.y_root)

            # Keep dragged frame highlighted
            try:
                for row_f in self._row_frames:
                    if row_f != dragged_frame:
                        row_f.configure(fg_color="transparent")
                dragged_frame.configure(fg_color=styles.DRAG_ACTIVE_ROW_COLOR)
            except Exception:
                pass

    def _reposition_row_frames(self):
        """Re-grid the row frames according to their order in self._row_frames."""
        for idx, frame in enumerate(self._row_frames):
            # try:
            frame.grid_configure(row=idx)
            # except Exception:
            #     pass

    def _compute_drop_index(self, y_root: int) -> int:
        """Compute target row index based on pointer absolute y position."""
        frames = self._row_frames
        if not frames:
            return 0
        mids = []
        for f in frames:
            try:
                top = f.winfo_rooty()
                h = f.winfo_height() or 1
                mids.append(top + h / 2)
            except Exception as e:
                mids.append(0) # Fallback in case of error

        # clamp to bounds
        if not mids: # Handle case where mids is empty (no frames)
            return 0

        if y_root <= mids[0]:
            return 0
        if y_root >= mids[-1]:
            return len(frames) - 1
        # find nearest
        best = 0
        best_d = float('inf')
        for i, m in enumerate(mids):
            d = abs(y_root - m)
            if d < best_d:
                best_d = d
                best = i
        return best

    def _get_prompt_list_canvas(self) -> Optional[tk.Canvas]:
        try:
            canvas = getattr(self.prompt_list_frame, "_parent_canvas", None)
            if isinstance(canvas, tk.Canvas):
                return canvas
            # Fallback: search children
            for ch in self.prompt_list_frame.winfo_children():
                if isinstance(ch, tk.Canvas):
                    return ch
        except Exception:
            pass
        return None

    def _draw_row_drop_indicator(self, y_root: int) -> None:
        parent = self.prompt_list_frame
        if parent is None or not self._row_frames:
            return
        try:
            # Compute boundaries between frames
            tops = []
            bottoms = []
            for f in self._row_frames:
                ty = f.winfo_rooty()
                h = f.winfo_height() or 1
                tops.append(ty)
                bottoms.append(ty + h)
            boundaries = []
            boundaries.append(tops[0])  # before first
            for i in range(1, len(self._row_frames)):
                boundaries.append(int((bottoms[i-1] + tops[i]) / 2))
            boundaries.append(bottoms[-1])  # after last
            by = min(boundaries, key=lambda b: abs(b - y_root))
            local_y = by - parent.winfo_rooty()
            w = max(2, parent.winfo_width())
            if self._row_drop_indicator_widget is None or not self._row_drop_indicator_widget.winfo_exists():
                self._row_drop_indicator_widget = tk.Frame(parent, bg="#FFFFFF", height=2)
                self._row_drop_indicator_widget.place(x=0, y=local_y, relwidth=1.0, height=2)
            else:
                self._row_drop_indicator_widget.place_configure(y=local_y)
        except Exception:
            pass

    def _add_prompt(self):
        dialog = PromptEditorDialog(self.app, tr("prompt.add_title"))
        new_prompt = dialog.get_result()
        if new_prompt and self.agent.config:
            prompt_id = new_prompt.name.lower().replace(" ", "_")
            try:
                self.agent.add_prompt(prompt_id, new_prompt)
                self._create_prompt_list_frame()
                save_config(self.agent.config)
                self.agent.notify_prompts_changed()
                CTkMessagebox(title=tr("common.success"), message=tr("prompt.add_success", name=new_prompt.name), icon="info")
            except ValueError as e:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.add_failed", details=str(e)), icon="cancel")
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.add_unexpected", details=str(e)), icon="cancel")

    def _edit_prompt(self, prompt_id: str):
        if not self.agent.config:
            return
        current_prompt = self.agent.config.prompts.get(prompt_id)
        if not current_prompt:
            CTkMessagebox(title=tr("common.error"), message=tr("prompt.not_found"), icon="cancel")
            return

        dialog = PromptEditorDialog(self.app, tr("prompt.edit_title_fmt", name=current_prompt.name), prompt=current_prompt)
        updated_prompt = dialog.get_result()
        if updated_prompt:
            try:
                new_prompt_id = updated_prompt.name.lower().replace(" ", "_")
                if new_prompt_id != prompt_id:
                    # 名前が変更され、IDが変わる場合
                    # 既に同じIDが存在するかチェック
                    if new_prompt_id in self.agent.config.prompts and new_prompt_id != prompt_id:
                        raise ValueError(tr("prompt.id_exists", id=new_prompt_id))
                    # 再注文を保持しつつIDと内容を更新する
                    items = list(self.agent.config.prompts.items())
                    # 現在の位置を取得
                    old_index = next((i for i, (pid, _) in enumerate(items) if pid == prompt_id), None)
                    if old_index is None:
                        raise ValueError(tr("prompt.id_missing", id=prompt_id))
                    # 削除し、新しいIDで挿入
                    items.pop(old_index)
                    items.insert(old_index, (new_prompt_id, updated_prompt))
                    # 更新
                    self.agent.config.prompts = {k: v for k, v in items}
                    CTkMessagebox(title=tr("common.success"), message=tr("prompt.rename_success_fmt", old=current_prompt.name, new=updated_prompt.name), icon="info")
                else:
                    # IDが変わらない場合、単純に更新
                    self.agent.update_prompt(prompt_id, updated_prompt)
                    CTkMessagebox(title=tr("common.success"), message=tr("prompt.update_success_fmt", name=updated_prompt.name), icon="info")
                # UIを再生成し、設定を保存
                save_config(self.agent.config)
                # 設定を再読み込みして最新状態を反映
                new_config = load_config()
                if new_config:
                    # 保持している他の設定値（履歴サイズやAPIキー）を維持しつつプロンプトだけ更新
                    self.agent.config.prompts = new_config.prompts
                self._create_prompt_list_frame()
                self.agent.notify_prompts_changed()
            except ValueError as e:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.update_failed", details=str(e)), icon="cancel")
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.update_unexpected", details=str(e)), icon="cancel")

    def _delete_prompt(self, prompt_id: str):
        if not self.agent.config:
            return
        if not self.agent.config:
            return
        prompt = self.agent.config.prompts.get(prompt_id)
        prompt_name = prompt.name if prompt else prompt_id
        msg_box = CTkMessagebox(title=tr("prompt.delete_title"), message=tr("prompt.delete_confirm", name=prompt_name), icon="question", option_1=tr("common.cancel"), option_2=tr("common.delete"))
        response = msg_box.get()
        if response == tr("common.delete"):
            try:
                self.agent.delete_prompt(prompt_id)
                self._create_prompt_list_frame()
                save_config(self.agent.config)
                self.agent.notify_prompts_changed()
                CTkMessagebox(title=tr("common.success"), message=tr("prompt.delete_success", name=prompt_name), icon="info")
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("prompt.delete_unexpected", details=str(e)), icon="cancel")

    def _save_settings(self):
        if not self.agent.config:
            return
        try:
            save_config(self.agent.config)
            CTkMessagebox(title=tr("settings.save_done_title"), message=tr("settings.save_done_message"), icon="info")
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=tr("settings.save_failed_unexpected", details=str(e)), icon="cancel")

    def _toggle_prompt_matrix(self, prompt_id: str, include: bool) -> None:
        """
        指定されたプロンプトのマトリクスフラグを更新し、設定を保存します。

        Parameters
        ----------
        prompt_id : str
            対象のプロンプトのID。
        include : bool
            マトリクスに含めるかどうかのフラグ。
        """
        # プロンプトが存在する場合のみ処理
        if not self.agent.config:
            return
        prompt = self.agent.config.prompts.get(prompt_id)
        if prompt:
            # フラグを更新
            prompt.include_in_matrix = include
            # 設定を保存
            try:
                save_config(self.agent.config)
                self.agent.notify_prompts_changed()
            except Exception as e:
                # 保存に失敗した場合はユーザに通知する
                CTkMessagebox(title=tr("common.error"), message=tr("settings.save_failed", details=str(e)), icon="cancel")

    def run(self):
        import threading
        tray_thread = threading.Thread(target=self.agent.run, daemon=True)
        tray_thread.start()

        if not self.agent.api_key:
            self.app.after(1000, self.agent.show_settings_window)

        self.app.mainloop()
