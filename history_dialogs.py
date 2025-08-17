import customtkinter as ctk
import tkinter as tk
from typing import Optional
import styles
from i18n import tr

class HistoryEditDialog(ctk.CTkToplevel):
    """
    テキスト履歴の編集専用ダイアログ。
    モデル選択や温度設定は含まず、テキストエリアとOK/キャンセルのみ。
    """
    def __init__(self, parent_app: Optional[ctk.CTk] = None, title: str = "", initial_value: str = ""):
        super().__init__(parent_app)
        self.title(title or tr("history.edit_title"))
        self.configure(fg_color=styles.POPUP_BG_COLOR)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # 中央表示
        self.update_idletasks()
        width, height = 600, 360
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

        # モーダル化
        if parent_app is not None:
            self.transient(parent_app)
            try:
                self.grab_set()
            except Exception:
                pass

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.label = ctk.CTkLabel(self, text=tr("history.edit_label"), text_color=styles.POPUP_TEXT_COLOR)
        self.label.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="w")

        self.textbox = ctk.CTkTextbox(self, height=220, border_width=1, border_color=styles.HIGHLIGHT_BORDER_COLOR)
        self.textbox.grid(row=1, column=0, padx=16, pady=8, sticky="nsew")
        self.textbox.insert("0.0", initial_value)
        self.textbox.focus_set()

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=16, pady=(8, 16), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        self.ok_button = ctk.CTkButton(button_frame, text=tr("common.ok"), command=self._on_ok)
        self.ok_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.cancel_button = ctk.CTkButton(button_frame, text=tr("common.cancel"), command=self._on_cancel, fg_color=styles.CANCEL_BUTTON_COLOR, text_color=styles.CANCEL_BUTTON_TEXT_COLOR)
        self.cancel_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")

        self._result: Optional[str] = None

    def show(self):
        self.lift()
        self.focus_force()
        self.wait_window(self)

    def get_input(self) -> Optional[str]:
        return self._result

    def _on_ok(self):
        self._result = self.textbox.get("0.0", "end-1c")
        self._safe_destroy()

    def _on_cancel(self):
        self._result = None
        self._safe_destroy()

    def _safe_destroy(self):
        try:
            if self.grab_current() == str(self):
                self.grab_release()
        except Exception:
            pass
        self.destroy()
