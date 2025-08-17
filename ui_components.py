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

# --- BaseDialog: A Modern Foundation for All Dialogs ---
class BaseDialog(ctk.CTkToplevel):
    """
    A modern, unified base class for all dialog windows in the application.

    This class provides a consistent look and feel by handling:
    - Unified background color from the central style guide.
    - Centered positioning on the screen.
    - Modality (grabbing focus).
    - Safe window destruction.
    - Standardized padding and layout configurations.
    """
    def __init__(self, parent_app: Optional[ctk.CTk] = None, title: str = "", geometry: str = "400x300"):
        super().__init__(parent_app)

        self.title(title)
        self.configure(fg_color=styles.POPUP_BG_COLOR)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Center the window
        self.update_idletasks()
        try:
            width, height = [int(x) for x in geometry.split('x')]
        except (ValueError, TypeError):
            width, height = 400, 300

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

        self.resizable(True, True)
        self.attributes("-topmost", True)

        # Set modality
        if parent_app:
            self.transient(parent_app)
            self.grab_set()

        # To be populated by subclasses
        self.result: Any = None

    def _on_closing(self):
        """Handle the window close event."""
        self.result = None
        self.destroy()

    def destroy(self):
        """Safely destroy the window, releasing the grab if set."""
        if self.winfo_exists() and self.grab_current() == str(self):
            try:
                self.grab_release()
            except tk.TclError as e:
                # This can happen if the window is destroyed by other means
                print(f"INFO: TclError during grab_release in BaseDialog: {e}")
        if self.winfo_exists():
            super().destroy()

    def show(self) -> Any:
        """
        Show the dialog, wait for it to close, and return the result.
        This makes the dialog behave in a blocking (modal) way.
        """
        self.master.wait_window(self)
        return self.result

class PromptParameterEditorFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(1, weight=1)

        # Temperature
        ctk.CTkLabel(self, text=tr("params.temperature"), text_color=styles.TEXT_COLOR).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        temp_frame = ctk.CTkFrame(self, fg_color="transparent")
        temp_frame.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        temp_frame.grid_columnconfigure(0, weight=1)
        self.temperature_slider = ctk.CTkSlider(temp_frame, from_=0, to=2, number_of_steps=20, command=self._update_temperature_label)
        self.temperature_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.temperature_value_label = ctk.CTkLabel(temp_frame, width=4, text_color=styles.TEXT_SECONDARY_COLOR)
        self.temperature_value_label.grid(row=0, column=1, padx=5)

        # Top P
        ctk.CTkLabel(self, text=tr("params.top_p"), text_color=styles.TEXT_COLOR).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        top_p_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_p_frame.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        top_p_frame.grid_columnconfigure(0, weight=1)
        self.top_p_slider = ctk.CTkSlider(top_p_frame, from_=0, to=1, number_of_steps=20, command=self._update_top_p_label)
        self.top_p_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.top_p_value_label = ctk.CTkLabel(top_p_frame, width=4, text_color=styles.TEXT_SECONDARY_COLOR)
        self.top_p_value_label.grid(row=0, column=1, padx=5)

        # Top K
        ctk.CTkLabel(self, text=tr("params.top_k"), text_color=styles.TEXT_COLOR).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        top_k_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_k_frame.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        top_k_frame.grid_columnconfigure(0, weight=1)
        self.top_k_slider = ctk.CTkSlider(top_k_frame, from_=1, to=100, number_of_steps=99, command=self._update_top_k_label)
        self.top_k_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.top_k_value_label = ctk.CTkLabel(top_k_frame, width=4, text_color=styles.TEXT_SECONDARY_COLOR)
        self.top_k_value_label.grid(row=0, column=1, padx=5)

        # Max Output Tokens
        ctk.CTkLabel(self, text=tr("params.max_output_tokens"), text_color=styles.TEXT_COLOR).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.max_output_tokens_entry = ctk.CTkEntry(self, fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, border_color=styles.BORDER_COLOR, border_width=1)
        self.max_output_tokens_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        # Stop Sequences
        ctk.CTkLabel(self, text=tr("params.stop_sequences"), text_color=styles.TEXT_COLOR).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.stop_sequences_entry = ctk.CTkEntry(self, fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, border_color=styles.BORDER_COLOR, border_width=1)
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
        self.withdraw()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.geometry(styles.ACTION_SELECTOR_GEOMETRY)
        self.configure(fg_color=styles.FRAME_BG_COLOR) # Use FRAME_BG_COLOR
        self.prompts = prompts
        self.on_prompt_selected_callback = on_prompt_selected_callback
        self.agent = agent
        self._on_destroy_callback = on_destroy_callback
        self.buttons: List[ctk.CTkButton] = []
        self.prompt_buttons: List[ctk.CTkButton] = []
        self.current_selection_index = 0
        self._is_destroying = False
        self.attached_file_paths = file_paths
        self._history_label_to_item: Dict[str, Dict[str, Any]] = {}
        self._selected_history_item: Optional[Dict[str, Any]] = None
        self._pending_close_id: Optional[str] = None

        self.grid_columnconfigure(0, weight=1)
        self.margin = styles.ACTION_SELECTOR_MARGIN

        # ... (rest of the __init__ method with updated styles)
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=self.margin, pady=(self.margin, 5))
        header_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header_frame, text=tr("action.input"), text_color=styles.TEXT_COLOR, anchor="w").grid(row=0, column=0, sticky="w", padx=(0,8))

        labels = self._get_history_labels()
        self.history_variable = ctk.StringVar(value=labels[0])
        self.history_menu = ctk.CTkOptionMenu(header_frame, values=labels, variable=self.history_variable, command=lambda v: self._on_history_changed())
        self.history_menu.grid(row=0, column=1, sticky="ew")

        self.edit_history_button = ctk.CTkButton(header_frame, text=tr("common.edit"), width=60, command=self._on_edit_history)
        self.edit_history_button.grid(row=0, column=2, padx=(5,0))

        self._on_history_changed()

        attach_row_frame = ctk.CTkFrame(self, fg_color="transparent")
        attach_row_frame.pack(fill="x", padx=self.margin, pady=5)
        attach_row_frame.grid_columnconfigure(0, weight=1)

        self.attached_files_label = ctk.CTkLabel(attach_row_frame, text="", text_color=styles.TEXT_SECONDARY_COLOR, anchor="w", justify="left")
        self.attached_files_label.grid(row=0, column=0, padx=(0,6), sticky="ew")

        ctk.CTkButton(attach_row_frame, text=tr("action.attach"), width=60, command=self._on_file_attach).grid(row=0, column=1, padx=2, sticky="e")

        self.scrollable_frame = ctk.CTkScrollableFrame(self, fg_color=styles.APP_BG_COLOR)
        self.scrollable_frame.pack(fill="both", expand=True, padx=self.margin, pady=5)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        for i, (prompt_id, prompt) in enumerate(self.prompts.items()):
            button = ctk.CTkButton(self.scrollable_frame, text=prompt.name, command=lambda p_id=prompt_id, p_conf=prompt: self._on_prompt_selected(p_id, p_conf), anchor="w")
            button.grid(row=i, column=0, padx=5, pady=styles.ACTION_SELECTOR_SPACING, sticky="ew")
            self.prompt_buttons.append(button)

        action_buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_buttons_frame.pack(fill="x", padx=self.margin, pady=5)
        action_buttons_frame.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(action_buttons_frame, text=tr("action.free_input"), command=self._on_free_input).grid(row=0, column=0, padx=(0,2), sticky="ew")
        ctk.CTkButton(action_buttons_frame, text=tr("action.matrix"), command=self._on_matrix).grid(row=0, column=1, padx=(2,0), sticky="ew")

        # Settings Icon
        try:
            # ... (icon loading logic remains the same)
            cfg_path = next((p for p in [Path.cwd()/"config.ico", Path(__file__).resolve().parent/"config.ico"] if p.exists()), None)
            cfg_img_pil = Image.open(cfg_path).convert("RGBA")
            self._cfg_img = ctk.CTkImage(light_image=cfg_img_pil, dark_image=cfg_img_pil, size=(20, 20))
            open_manager_button = ctk.CTkButton(action_buttons_frame, text="", image=self._cfg_img, width=28, height=28, command=self._on_open_prompt_manager, fg_color="transparent")
        except Exception:
            open_manager_button = ctk.CTkButton(action_buttons_frame, text=tr("settings.title"), width=60, command=self._on_open_prompt_manager)
        open_manager_button.grid(row=0, column=2, padx=(8,0))

        cancel_button = ctk.CTkButton(self, text=tr("common.cancel"), command=self.destroy, fg_color=styles.CANCEL_BUTTON_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        cancel_button.pack(fill="x", padx=self.margin, pady=(5, self.margin))

        self._update_selection_highlight()
        self._bind_events()
        if self.attached_file_paths: self.update_attached_files_display(self.attached_file_paths)

    def _get_history_labels(self) -> List[str]:
        self._history_label_to_item.clear()
        labels = []
        for item in getattr(self.agent, 'clipboard_history', [])[:50]:
            label = ""
            if isinstance(item, str):
                label = item
                self._history_label_to_item[label] = {"type": "text", "data": item}
            elif isinstance(item, dict) and 'type' in item:
                if item['type'] == 'text': label = item.get('data', '')
                elif item['type'] in ('image', 'image_compressed'): label = tr("history.image")
                elif item['type'] == 'file': label = f"[{tr('history.file_name_prefix', name=Path(item.get('data','')).name)}]"
                else: label = tr("history.unknown")

                unique_label = label.replace('\n', ' ')[:60] + ("…" if len(label) > 60 else "")
                c = 1
                while unique_label in self._history_label_to_item:
                    c += 1
                    unique_label = f"{unique_label.split('(')[0].strip()} ({c})"
                self._history_label_to_item[unique_label] = item
                labels.append(unique_label)

        return labels if labels else [tr("history.empty")]

    def _bind_events(self):
        self.bind("<Up>", self._on_key_up)
        self.bind("<Down>", self._on_key_down)
        self.bind("<Tab>", self._on_key_down)
        self.bind("<Shift-Tab>", self._on_key_up)
        self.bind("<Return>", self._on_key_enter)
        self.bind("<Escape>", self._on_key_escape)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<FocusIn>", self._on_focus_in)
        for seq in ("<Motion>", "<ButtonPress>", "<KeyPress>"):
            self.bind_all(seq, self._on_user_activity, add=True)

    # ... (rest of the methods for ActionSelectorWindow, largely unchanged logic but check for style calls)
    def _on_open_prompt_manager(self):
        if getattr(self, '_is_destroying', False): return
        if hasattr(self.agent, '_show_main_window'): self.agent._show_main_window()
        self._close_safely()

    def _close_safely(self):
        if getattr(self, '_is_destroying', False): return
        self._is_destroying = True
        self.withdraw()
        self.after(20, lambda: super(ActionSelectorWindow, self).destroy())

    def _on_prompt_selected(self, prompt_id: str, prompt_config: Prompt):
        if self._is_destroying: return
        if self._selected_history_item:
            item = self._selected_history_item
            if isinstance(item, dict) and item.get('type') == 'image_compressed':
                try:
                    raw = zlib.decompress(base64.b64decode(item.get('data', '')))
                    item = {"type": "image", "data": base64.b64encode(raw).decode('utf-8')}
                except Exception: pass
            setattr(self.agent, '_temp_input_for_processing', item)
        self.on_prompt_selected_callback(prompt_id=prompt_id, file_paths=self.attached_file_paths)
        self.destroy()

    def _on_key_up(self, event):
        if not self.prompt_buttons: return "break"
        self.current_selection_index = (self.current_selection_index - 1 + len(self.prompt_buttons)) % len(self.prompt_buttons)
        self._update_selection_highlight()
        self.prompt_buttons[self.current_selection_index].focus_set()
        self._scroll_prompt_into_view(self.prompt_buttons[self.current_selection_index])
        return "break"

    def _on_key_down(self, event):
        if not self.prompt_buttons: return "break"
        self.current_selection_index = (self.current_selection_index + 1) % len(self.prompt_buttons)
        self._update_selection_highlight()
        self.prompt_buttons[self.current_selection_index].focus_set()
        self._scroll_prompt_into_view(self.prompt_buttons[self.current_selection_index])
        return "break"

    def _on_key_enter(self, event):
        if self.prompt_buttons: self.prompt_buttons[self.current_selection_index].invoke()
        return "break"

    def _on_key_escape(self, event):
        self.destroy()
        return "break"

    def _on_free_input(self):
        if self._is_destroying: return
        dlg = ResizableInputDialog(self, title=tr("free_input.title"), text=tr("free_input.prompt_label"), agent=self.agent, enable_history=True)
        prompt_text = dlg.show()
        if prompt_text:
            model_name = "gemini-2.5-flash-lite"
            if hasattr(dlg, 'model_variable'): model_name = dlg.model_variable.get().split(" ")[0]
            temperature_val = 1.0
            if hasattr(dlg, 'temperature_slider'): temperature_val = float(dlg.temperature_slider.get())
            self.agent._run_process_in_thread(system_prompt=prompt_text, model=model_name, temperature=temperature_val, file_paths=self.attached_file_paths)
        self.destroy()

    def _on_file_attach(self):
        if self._is_destroying: return
        self.attributes("-topmost", False)
        paths = filedialog.askopenfilenames()
        self.attributes("-topmost", True)
        if paths:
            abs_paths = [str(Path(p).resolve()) for p in paths]
            setattr(self.agent, '_temp_file_paths_for_processing', abs_paths)
            self.attached_file_paths = abs_paths
            self.update_attached_files_display(abs_paths)
            if self.prompt_buttons: self.prompt_buttons[self.current_selection_index].focus_set()

    def _on_matrix(self):
        if self._is_destroying: return
        self.agent.show_matrix_batch_processor_window()
        self.destroy()

    def _update_selection_highlight(self):
        for i, button in enumerate(self.prompt_buttons):
            is_selected = (i == self.current_selection_index)
            button.configure(border_color=styles.HIGHLIGHT_BORDER_COLOR if is_selected else self.cget("fg_color"), border_width=styles.HIGHLIGHT_BORDER_WIDTH if is_selected else 1)

    def _get_scroll_canvas(self):
        return getattr(self.scrollable_frame, "_parent_canvas", None)

    def _scroll_prompt_into_view(self, btn):
        self.update_idletasks()
        canvas = self._get_scroll_canvas()
        if not canvas: return
        view_top, view_height = canvas.canvasy(0), canvas.winfo_height()
        btn_top, btn_height = canvas.canvasy(btn.winfo_rooty() - canvas.winfo_rooty()), btn.winfo_height()
        if btn_top < view_top:
            canvas.yview_moveto(btn_top / max(1, canvas.bbox("all")[3]))
        elif btn_top + btn_height > view_top + view_height:
            canvas.yview_moveto((btn_top + btn_height - view_height) / max(1, canvas.bbox("all")[3]))

    def show_at_cursor(self, cursor_pos: Optional[tuple] = None):
        self.update_idletasks()
        w, h = [int(d) for d in styles.ACTION_SELECTOR_GEOMETRY.split('x')]
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        cx, cy = cursor_pos or (self.winfo_pointerx(), self.winfo_pointery())
        x, y = max(10, min(cx + 20, sw - w - 10)), max(10, min(cy + 20, sh - h - 50))
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
        self.lift()
        self.focus_force()
        if self.prompt_buttons: self.prompt_buttons[0].focus_set()
        self._update_selection_highlight()

    def destroy(self):
        if self._is_destroying: return
        self._is_destroying = True
        if self.winfo_exists() and self.grab_current() == str(self):
            self.grab_release()
        if self._on_destroy_callback: self._on_destroy_callback()
        super().destroy()

    def _on_history_changed(self):
        item = self._history_label_to_item.get(self.history_variable.get())
        self._selected_history_item = item
        is_image = isinstance(item, dict) and item.get('type') in ('image', 'image_compressed')
        self.edit_history_button.configure(state='disabled' if is_image else 'normal')

    def _on_edit_history(self):
        item = self._selected_history_item
        if not isinstance(item, dict) or 'type' not in item: return
        if item.get('type') == 'text':
            dlg = HistoryEditDialog(self, initial_value=item.get('data',''))
            new_text = dlg.show()
            if new_text is not None: self._selected_history_item = {"type": "text", "data": new_text}
        elif item.get('type') == 'file':
            if new_path := filedialog.askopenfilename(): self._selected_history_item = {"type": "file", "data": new_path}

    def update_attached_files_display(self, file_paths: Optional[List[str]]):
        if not self.attached_files_label.winfo_exists(): return
        if file_paths:
            names = ", ".join([Path(p).name for p in file_paths])
            text = tr("history.file_label_prefix") + names
            if len(text) > 60: text = text[:59] + "…"
            self.attached_files_label.configure(text=text)
        else:
            self.attached_files_label.configure(text="")

    def _schedule_close_after_delay(self, delay_ms: int):
        self._cancel_scheduled_close()
        self._pending_close_id = self.after(delay_ms, self._close_if_still_inactive)

    def _cancel_scheduled_close(self):
        if self._pending_close_id: self.after_cancel(self._pending_close_id)
        self._pending_close_id = None

    def _on_focus_out(self, event): self._schedule_close_after_delay(800)
    def _on_focus_in(self, event): self._cancel_scheduled_close()
    def _on_user_activity(self, event): self._cancel_scheduled_close()

    def _close_if_still_inactive(self):
        self._pending_close_id = None
        try:
            if self.focus_get() and self._is_child_of_self(self.focus_get()): return
            px, py = self.winfo_pointerx(), self.winfo_pointery()
            wx, wy, ww, wh = self.winfo_rootx(), self.winfo_rooty(), self.winfo_width(), self.winfo_height()
            if wx <= px <= wx + ww and wy <= py <= wy + wh: return
        except Exception: pass
        self.destroy()

    def _is_child_of_self(self, widget):
        parent = widget
        while parent:
            if parent == self: return True
            parent = parent.master
        return False

class NotificationPopup(ctk.CTkToplevel):
    def __init__(self, title: str, message: str, parent_app, level: Literal["info", "warning", "error", "success"] = "info", on_destroy_callback: Optional[Callable] = None):
        super().__init__(parent_app)
        self.title(title)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._on_destroy_callback = on_destroy_callback
        self._close_after_id = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.content_frame = ctk.CTkFrame(self, corner_radius=8, border_width=1)
        self.content_frame.grid(row=0, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.message_label = ctk.CTkTextbox(self.content_frame, wrap="word", height=150, border_width=0, font=styles.FONT_NORMAL)
        self.message_label.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.message_label.insert("0.0", message)
        self.message_label.configure(state="disabled")

        self.attributes("-alpha", styles.NOTIFICATION_POPUP_ALPHA)
        self._set_colors_by_level(level)
        self.initial_width = styles.NOTIFICATION_POPUP_INITIAL_WIDTH

    def destroy(self):
        if self._close_after_id: self.after_cancel(self._close_after_id)
        if self._on_destroy_callback: self._on_destroy_callback()
        super().destroy()

    def reconfigure(self, title: str, message: str, level: Literal["info", "warning", "error", "success"], duration_ms: Optional[int]):
        if self._close_after_id: self.after_cancel(self._close_after_id)

        self.title(title)
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
        self.attributes("-topmost", True)

        if duration_ms: self._close_after_id = self.after(duration_ms, self.destroy)

    def update_message(self, new_chunk: str):
        self.message_label.configure(state="normal")
        self.message_label.insert("end", new_chunk)
        self.message_label.see("end")
        self.message_label.configure(state="disabled")
        self.update_idletasks()
        self._adjust_window_size()

    def _set_colors_by_level(self, level: str):
        bg_color, border_color = styles.NOTIFICATION_COLORS.get(level, styles.NOTIFICATION_COLORS["info"])
        self.content_frame.configure(fg_color=bg_color, border_color=border_color)
        self.message_label.configure(fg_color=bg_color)

    def _adjust_window_size(self):
        self.update_idletasks()
        req_height = self.message_label.winfo_reqheight() + 40
        max_height = self.winfo_screenheight() // 2
        new_height = min(max(req_height, 150), max_height)
        self.geometry(f"{self.initial_width}x{new_height}")

    def show_at_cursor(self, title: str, message: str, level: Literal["info", "warning", "error", "success"], duration_ms: Optional[int]):
        self.reconfigure(title, message, level, duration_ms)
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        cx, cy = self.winfo_pointerx(), self.winfo_pointery()
        x, y = max(10, min(cx + 20, sw - w - 10)), max(10, min(cy + 20, sh - h - 50))
        self.geometry(f"+{x}+{y}")
        self.state('normal')
        self.lift()
        self.attributes("-topmost", True)

class PromptEditorDialog(BaseDialog):
    def __init__(self, parent_app, title: str, prompt: Optional[Prompt] = None):
        super().__init__(parent_app, title=title, geometry=styles.PROMPT_EDITOR_GEOMETRY)
        self.result: Optional[Prompt] = None

        self.grid_columnconfigure(1, weight=1)

        self.available_models = [label for _, label in SUPPORTED_MODELS]
        self.model_variable = ctk.StringVar()

        ctk.CTkLabel(self, text=tr("prompt.name"), text_color=styles.TEXT_COLOR).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.name_entry = ctk.CTkEntry(self, width=300, fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, border_color=styles.BORDER_COLOR, border_width=1)
        self.name_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self, text=tr("common.model"), text_color=styles.TEXT_COLOR).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.model_optionmenu = ctk.CTkOptionMenu(self, values=self.available_models, variable=self.model_variable, fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, button_color=styles.ACCENT_COLOR, button_hover_color=styles.ACCENT_HOVER_COLOR)
        self.model_optionmenu.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.parameter_editor = PromptParameterEditorFrame(self, fg_color="transparent")
        self.parameter_editor.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self, text=tr("prompt.thinking_level"), text_color=styles.TEXT_COLOR).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.thinking_level_optionmenu = ctk.CTkOptionMenu(self, values=["Fast", "Balanced", "High Quality", "Unlimited"], fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, button_color=styles.ACCENT_COLOR, button_hover_color=styles.ACCENT_HOVER_COLOR)
        self.thinking_level_optionmenu.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        self.enable_web_var = ctk.BooleanVar(value=False)
        ctk.CTkLabel(self, text=tr("prompt.enable_web"), text_color=styles.TEXT_COLOR).grid(row=4, column=0, padx=10, pady=(0, 10), sticky="w")
        self.enable_web_switch = ctk.CTkSwitch(self, text="", variable=self.enable_web_var, progress_color=styles.ACCENT_COLOR)
        self.enable_web_switch.grid(row=4, column=1, padx=10, pady=(0, 10), sticky="w")

        ctk.CTkLabel(self, text=tr("prompt.system_prompt"), text_color=styles.TEXT_COLOR).grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.system_prompt_textbox = ctk.CTkTextbox(self, width=480, height=400, fg_color=styles.COMPONENT_BG_COLOR, text_color=styles.TEXT_COLOR, border_color=styles.BORDER_COLOR, border_width=1)
        self.system_prompt_textbox.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_rowconfigure(6, weight=1)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=7, column=0, columnspan=2, pady=20, padx=10, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(button_frame, text=tr("common.save"), command=self.on_save, fg_color=styles.ACCENT_COLOR, text_color=styles.DEFAULT_BUTTON_TEXT_COLOR, hover_color=styles.ACCENT_HOVER_COLOR).grid(row=0, column=0, padx=(0,5), pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self._on_closing, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR).grid(row=0, column=1, padx=(5,0), pady=5, sticky="ew")

        if prompt:
            self.name_entry.insert(0, prompt.name)
            self.model_variable.set(model_id_to_label(prompt.model))
            self.parameter_editor.set_parameters(prompt.parameters)
            self.thinking_level_optionmenu.set(prompt.thinking_level)
            self.enable_web_var.set(getattr(prompt, 'enable_web', False))
            self.system_prompt_textbox.insert("0.0", prompt.system_prompt)
        else:
            self.parameter_editor.set_parameters(PromptParameters())
            self.model_variable.set(self.available_models[1])
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
                name=name, model=model, system_prompt=system_prompt,
                thinking_level=thinking_level, enable_web=enable_web, parameters=parameters
            )
            self.destroy()
        except Exception as e:
            CTkMessagebox(title=tr("common.error"), message=f"Unexpected error: {e}", icon="cancel")

    def _on_closing(self):
        self.result = None
        self.destroy()


class SettingsWindow(BaseDialog):
    def __init__(self, parent_app, agent):
        super().__init__(parent_app, title=tr("settings.title"), geometry=styles.SETTINGS_WINDOW_GEOMETRY)
        self.agent = agent
        self.parent_app = parent_app

        self.configure(fg_color=styles.APP_BG_COLOR)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        def create_section_frame(parent):
            frame = ctk.CTkFrame(parent, fg_color=styles.FRAME_BG_COLOR, corner_radius=8)
            frame.grid(column=0, padx=10, pady=10, sticky="ew")
            frame.grid_columnconfigure(1, weight=1)
            return frame

        api_key_frame = create_section_frame(main_frame)
        ctk.CTkLabel(api_key_frame, text=tr("settings.api_key"), font=styles.FONT_BOLD).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.api_key_entry = ctk.CTkEntry(api_key_frame, placeholder_text=tr("api.placeholder"), border_color=styles.BORDER_COLOR, border_width=1)
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        current_api_key = keyring.get_password(API_SERVICE_ID, "api_key")
        if current_api_key:
            self.api_key_entry.insert(0, "*" * (len(current_api_key) - 4) + current_api_key[-4:])
            self.api_key_entry.configure(state="readonly")

        api_button_frame = ctk.CTkFrame(api_key_frame, fg_color="transparent")
        api_button_frame.grid(row=0, column=2, padx=10, pady=5)
        ctk.CTkButton(api_button_frame, text=tr("common.save"), command=self._save_api_key).pack(side="top", fill="x", expand=True)
        ctk.CTkButton(api_button_frame, text=tr("common.delete"), command=self._delete_api_key, fg_color=styles.DESTRUCTIVE_COLOR, hover_color=styles.DESTRUCTIVE_HOVER_COLOR).pack(side="top", fill="x", expand=True, pady=(5,0))

        hotkey_frame = create_section_frame(main_frame)
        def create_hotkey_row(parent, label, key_attr, row):
            ctk.CTkLabel(parent, text=label).grid(row=row, column=0, padx=10, pady=8, sticky="w")
            label_widget = ctk.CTkLabel(parent, text=self._fmt_hotkey(getattr(self.agent.config, key_attr, None)), text_color=styles.TEXT_SECONDARY_COLOR)
            label_widget.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
            ctk.CTkButton(parent, text=tr("common.edit"), width=80, command=lambda: self._set_hotkey(key_attr, label_widget)).grid(row=row, column=2, padx=5, pady=5)
            ctk.CTkButton(parent, text=tr("common.disable"), width=80, command=lambda: self._clear_hotkey(key_attr, label_widget), fg_color=styles.CANCEL_BUTTON_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR).grid(row=row, column=3, padx=5, pady=5)
            return label_widget

        self.hotkey_prompt_label = create_hotkey_row(hotkey_frame, tr("settings.hotkey.list"), 'hotkey_prompt_list', 0)
        self.hotkey_refine_label = create_hotkey_row(hotkey_frame, tr("settings.hotkey.refine"), 'hotkey_refine', 1)
        self.hotkey_matrix_label = create_hotkey_row(hotkey_frame, tr("settings.hotkey.matrix"), 'hotkey_matrix', 2)

        other_settings_frame = create_section_frame(main_frame)
        ctk.CTkLabel(other_settings_frame, text=tr("settings.language")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        try:
            names = available_locales()
        except Exception:
            names = {"auto": "Auto", "en": "English", "ja": "日本語"}
        self._lang_codes = list(names.keys())
        self._lang_var = ctk.StringVar(value=names.get(getattr(self.agent.config, 'language', 'auto'), "Auto"))
        self._lang_menu = ctk.CTkOptionMenu(other_settings_frame, values=list(names.values()), variable=self._lang_var)
        self._lang_menu.grid(row=0, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(other_settings_frame, text=tr("settings.history.max")).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.max_history_entry = ctk.CTkEntry(other_settings_frame, border_color=styles.BORDER_COLOR, border_width=1)
        self.max_history_entry.grid(row=1, column=1, columnspan=3, padx=10, pady=10, sticky="ew")
        self.max_history_entry.insert(0, str(self.agent.config.max_history_size))

        ctk.CTkLabel(other_settings_frame, text=tr("settings.flow.max_steps")).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.max_flow_steps_entry = ctk.CTkEntry(other_settings_frame, border_color=styles.BORDER_COLOR, border_width=1)
        self.max_flow_steps_entry.grid(row=2, column=1, columnspan=3, padx=10, pady=10, sticky="ew")
        self.max_flow_steps_entry.insert(0, str(getattr(self.agent.config, 'max_flow_steps', 5)))

        save_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        save_button_frame.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
        save_button_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(save_button_frame, text=tr("settings.save_and_close"), command=self._save_settings).grid(row=0, column=0, sticky="ew")

    def _save_api_key(self):
        new_api_key = self.api_key_entry.get()
        if new_api_key.startswith("****"): return
        if new_api_key:
            try:
                keyring.set_password(API_SERVICE_ID, "api_key", new_api_key)
                self.agent.api_key = new_api_key
                genai.configure(api_key=self.agent.api_key)
                CTkMessagebox(title=tr("common.success"), message=tr("settings.save_done_message"), icon="check")
                self.api_key_entry.delete(0, ctk.END)
                self.api_key_entry.insert(0, "*" * (len(new_api_key) - 4) + new_api_key[-4:])
                self.api_key_entry.configure(state="readonly")
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("api.save_failed", details=str(e)), icon="cancel")
        else:
            CTkMessagebox(title=tr("common.error"), message=tr("api.enter_key"), icon="warning")

    def _delete_api_key(self):
        if CTkMessagebox(title=tr("api.confirm_delete_title"), message=tr("api.confirm_delete_message"), icon="question", option_1=tr("common.cancel"), option_2=tr("common.delete")).get() == tr("common.delete"):
            try:
                keyring.delete_password(API_SERVICE_ID, "api_key")
                self.agent.api_key = None
                CTkMessagebox(title=tr("common.success"), message=tr("api.deleted"), icon="check")
                self.api_key_entry.configure(state="normal")
                self.api_key_entry.delete(0, ctk.END)
                self.api_key_entry.focus_set()
            except Exception as e:
                CTkMessagebox(title=tr("common.error"), message=tr("api.delete_failed", details=str(e)), icon="cancel")

    def _fmt_hotkey(self, value: Optional[str]) -> str:
        return value or tr("common.unspecified")

    def _set_hotkey(self, target_attr: str, label: ctk.CTkLabel):
        label.configure(text=tr("hotkey.press_new"))
        self.update_idletasks()
        self._set_all_hotkey_buttons_state('disabled')
        threading.Thread(target=lambda: self.after(0, self._capture_hotkey, target_attr, label), daemon=True).start()

    def _capture_hotkey(self, target_attr: str, label: ctk.CTkLabel):
        try:
            hotkey = keyboard.read_hotkey(suppress=False)
            self.after(0, self._on_hotkey_captured, target_attr, label, hotkey)
        except Exception as e:
            self.after(0, lambda: CTkMessagebox(title=tr("common.error"), message=tr("hotkey.read_failed", details=str(e)), icon="cancel"))
            self.after(0, self._refresh_all_hotkey_labels)
            self.after(0, lambda: self._set_all_hotkey_buttons_state('normal'))

    def _on_hotkey_captured(self, target_attr: str, label: ctk.CTkLabel, hotkey: str):
        label.configure(text=hotkey)
        if CTkMessagebox(title=tr("hotkey.confirm_title"), message=tr("hotkey.confirm_message", hotkey=hotkey), icon="question", option_1=tr("common.cancel"), option_2=tr("common.yes")).get() == tr("common.yes"):
            if not self.agent.update_hotkey(target_attr, hotkey):
                CTkMessagebox(title=tr("common.error"), message=tr("hotkey.update_failed"), icon="cancel")
        self._refresh_all_hotkey_labels()
        self._set_all_hotkey_buttons_state('normal')

    def _clear_hotkey(self, target_attr: str, label: ctk.CTkLabel):
        self.agent.update_hotkey(target_attr, None)
        self._refresh_all_hotkey_labels()

    def _set_all_hotkey_buttons_state(self, state: str):
        for frame in self.winfo_children():
            if isinstance(frame, ctk.CTkFrame):
                for widget in frame.winfo_children():
                    if isinstance(widget, ctk.CTkButton):
                        widget.configure(state=state)

    def _refresh_all_hotkey_labels(self):
        self.hotkey_prompt_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_prompt_list', None)))
        self.hotkey_refine_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_refine', None)))
        self.hotkey_matrix_label.configure(text=self._fmt_hotkey(getattr(self.agent.config, 'hotkey_matrix', None)))

    def _save_settings(self):
        try:
            self.agent.config.max_history_size = int(self.max_history_entry.get())
            self.agent.config.max_flow_steps = int(self.max_flow_steps_entry.get())
            lang_name = self._lang_var.get()
            lang_code = next((code for code, name in available_locales().items() if name == lang_name), 'auto')
            self.agent.config.language = lang_code
            set_locale(lang_code)
            if self.parent_app: self.parent_app.title(tr("app.title"))
            save_config(self.agent.config)
            self.agent.apply_config_changes()
            CTkMessagebox(title=tr("settings.save_done_title"), message=tr("settings.save_done_message"), icon="check")
            self.destroy()
        except (ValueError, Exception) as e:
            CTkMessagebox(title=tr("common.error"), message=f"Save failed: {e}", icon="cancel")

class ResizableInputDialog(BaseDialog):
    def __init__(self, parent_app, title: str, text: str, initial_value: str = "", agent: Optional[Any] = None, enable_history: bool = False):
        super().__init__(parent_app, title=title, geometry="500x350")
        self._agent = agent
        self._enable_history = enable_history
        self.minsize(400, 300)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.label = ctk.CTkLabel(self, text=text, wraplength=460, text_color=styles.TEXT_COLOR)
        self.label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        self.textbox = ctk.CTkTextbox(self, border_width=1, border_color=styles.BORDER_COLOR)
        self.textbox.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        if initial_value: self.textbox.insert("0.0", initial_value)
        self.textbox.focus_set()

        parameter_frame = ctk.CTkFrame(self, fg_color="transparent")
        parameter_frame.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        parameter_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(parameter_frame, text=tr("common.model"), text_color=styles.TEXT_COLOR).grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
        self.available_models = ["gemini-2.5-flash-lite (高速、低精度)", "gemini-2.5-flash (普通)", "gemini-2.5-pro (低速、高精度)"]
        self.model_variable = ctk.StringVar(value=self.available_models[0])
        self.model_optionmenu = ctk.CTkOptionMenu(parameter_frame, values=self.available_models, variable=self.model_variable)
        self.model_optionmenu.grid(row=0, column=1, pady=5, sticky="ew")

        ctk.CTkLabel(parameter_frame, text=tr("params.temperature"), text_color=styles.TEXT_COLOR).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="w")
        temp_control_frame = ctk.CTkFrame(parameter_frame, fg_color="transparent")
        temp_control_frame.grid(row=1, column=1, pady=5, sticky="ew")
        temp_control_frame.grid_columnconfigure(0, weight=1)
        self.temperature_slider = ctk.CTkSlider(temp_control_frame, from_=0, to=2, number_of_steps=20, command=lambda v: self.temperature_value_label.configure(text=f"{v:.1f}"))
        self.temperature_slider.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.temperature_slider.set(1.0)
        self.temperature_value_label = ctk.CTkLabel(temp_control_frame, width=4, text="1.0", text_color=styles.TEXT_SECONDARY_COLOR)
        self.temperature_value_label.grid(row=0, column=1, padx=5)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        if self._enable_history and self._agent:
            history_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
            history_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
            history_frame.grid_columnconfigure(0, weight=1)
            history_items = [item.get('data', '') for item in getattr(self._agent, 'clipboard_history', []) if isinstance(item, dict) and item.get('type') == 'text']
            labels = [s.replace('\n', ' ')[:60] + ('…' if len(s) > 60 else '') for s in history_items]
            if labels:
                self._history_map = dict(zip(labels, history_items))
                self.history_menu = ctk.CTkOptionMenu(history_frame, values=labels)
                self.history_menu.grid(row=0, column=0, padx=(0,10), sticky="ew")
                ctk.CTkButton(history_frame, text=tr("history.insert"), width=100, command=self._insert_selected_history).grid(row=0, column=1, sticky="e")

        ctk.CTkButton(button_frame, text=tr("common.ok"), command=self._on_ok).grid(row=1, column=0, padx=(0, 5), sticky="ew")
        ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self._on_closing, fg_color=styles.CANCEL_BUTTON_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR).grid(row=1, column=1, padx=(5, 0), sticky="ew")

    def _on_ok(self):
        self.result = self.textbox.get("0.0", "end-1c")
        self.destroy()

    def _insert_selected_history(self):
        if hasattr(self, 'history_menu') and hasattr(self, '_history_map'):
            val = self._history_map.get(self.history_menu.get(), '')
            if val: self.textbox.insert('insert', val)


class MatrixSummarySettingsDialog(BaseDialog):
    def __init__(self, parent_app, agent):
        super().__init__(parent_app, title=tr("matrix.settings.title"), geometry="520x250")
        self.agent = agent
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        main_frame.grid_columnconfigure(1, weight=1)

        def create_row(parent, label_text, row_index, prompt_attr, edit_cmd):
            ctk.CTkLabel(parent, text=label_text, text_color=styles.TEXT_COLOR).grid(row=row_index, column=0, padx=(0,8), pady=10, sticky="w")
            label = ctk.CTkLabel(parent, text=self._prompt_title(getattr(agent.config, prompt_attr, None)), text_color=styles.TEXT_SECONDARY_COLOR, anchor="w")
            label.grid(row=row_index, column=1, sticky="ew")
            ctk.CTkButton(parent, text=tr("common.edit"), width=70, command=edit_cmd).grid(row=row_index, column=2, padx=(8,0))
            return label

        self.row_label = create_row(main_frame, tr("matrix.row_summary"), 0, 'matrix_row_summary_prompt', self._edit_row)
        self.col_label = create_row(main_frame, tr("matrix.col_summary"), 1, 'matrix_col_summary_prompt', self._edit_col)
        self.matrix_label = create_row(main_frame, tr("matrix.matrix_summary"), 2, 'matrix_matrix_summary_prompt', self._edit_matrix)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        ctk.CTkButton(btn_frame, text=tr("common.close"), command=self.destroy, fg_color=styles.CANCEL_BUTTON_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR).pack()

    def _prompt_title(self, p: Optional[Prompt]) -> str:
        return p.name if isinstance(p, Prompt) else tr("common.unspecified")

    def _edit_target(self, attr: str, title: str, label_widget: ctk.CTkLabel):
        current = getattr(self.agent.config, attr, None)
        editor = PromptEditorDialog(self, title=title, prompt=current)
        result = editor.show()
        if result:
            setattr(self.agent.config, attr, result)
            save_config(self.agent.config)
            label_widget.configure(text=self._prompt_title(result))
            if hasattr(self.agent, "notify_prompts_changed"):
                self.agent.notify_prompts_changed()

    def _edit_row(self): self._edit_target('matrix_row_summary_prompt', tr("matrix.row_edit_title"), self.row_label)
    def _edit_col(self): self._edit_target('matrix_col_summary_prompt', tr("matrix.col_edit_title"), self.col_label)
    def _edit_matrix(self): self._edit_target('matrix_matrix_summary_prompt', tr("matrix.matrix_edit_title"), self.matrix_label)
