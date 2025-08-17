import customtkinter as ctk
from typing import Optional

import styles
from i18n import tr
from ui_components import BaseDialog

class HistoryEditDialog(BaseDialog):
    """A dialog for editing text history items."""
    def __init__(self, parent_app: Optional[ctk.CTk] = None, title: str = "", initial_value: str = ""):
        super().__init__(parent_app, title=(title or tr("history.edit_title")), geometry="600x360")
        self.result: Optional[str] = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.label = ctk.CTkLabel(self, text=tr("history.edit_label"), text_color=styles.TEXT_COLOR)
        self.label.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="w")

        self.textbox = ctk.CTkTextbox(self, border_width=1, border_color=styles.BORDER_COLOR)
        self.textbox.grid(row=1, column=0, padx=16, pady=8, sticky="nsew")
        self.textbox.insert("0.0", initial_value)
        self.textbox.focus_set()

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=16, pady=(8, 16), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        self.ok_button = ctk.CTkButton(button_frame, text=tr("common.ok"), command=self._on_ok)
        self.ok_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.cancel_button = ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self._on_closing, fg_color=styles.CANCEL_BUTTON_COLOR, hover_color=styles.CANCEL_BUTTON_HOVER_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        self.cancel_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def _on_ok(self):
        self.result = self.textbox.get("0.0", "end-1c")
        self.destroy()
