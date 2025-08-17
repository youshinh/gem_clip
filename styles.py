# styles.py

# --- 色定義 (モノクローム調) ---
# 全体的に彩度を抑え、目に優しいグレーを基調とした配色

# ボタンの基本色
DEFAULT_BUTTON_FG_COLOR = ("#555555", "#333333")  # ダークグレー
DEFAULT_BUTTON_TEXT_COLOR = ("#FFFFFF", "#F0F0F0")  # ホワイト

# 強調表示用の境界線色
# 強調表示用の境界線色（赤系に変更）
HIGHLIGHT_BORDER_COLOR = ("#007ACC", "#007ACC") # 青系の強調色

# 各種ボタンの色
FILE_ATTACH_BUTTON_COLOR = ("#6c757d", "#5a6268") # 標準ボタンと少し変えたグレー
MATRIX_BUTTON_COLOR = ("#6c757d", "#5a6268")      # 標準ボタンと少し変えたグレー
CANCEL_BUTTON_COLOR = ("#B0B0B0", "#4D4D4D")      # 明るめのグレー
CANCEL_BUTTON_TEXT_COLOR = ("#000000", "#FFFFFF")  # 黒と白
DELETE_BUTTON_COLOR = "#C44141"                   # 彩度を抑えた赤
DELETE_BUTTON_HOVER_COLOR = "#A43737"             # 少し濃い赤

# 通知ポップアップの色（彩度を抑えて目に優しく）
NOTIFICATION_COLORS = {
    "info":    ("#A9CCE3", "#546E7A"), # ソフトな青
    "warning": ("#F7DC6F", "#B7950B"), # 落ち着いた黄色
    "error":   ("#E6B0AA", "#922B21"), # 落ち着いた赤
    "success": ("#A9DFBF", "#239B56")  # 落ち着いた緑
}

# 履歴アイテムの色
HISTORY_ITEM_FG_COLOR = ("#F0F0F0", "#2C2C2C")   # 明るいグレー / 暗いグレー
HISTORY_ITEM_TEXT_COLOR = ("#1C1C1C", "#EAEAEA") # オフブラック / オフホワイト
FLOW_RESULT_TEXT_COLOR = ("#D94B8C", "#FF8FB7")  # 明るいピンク系（ライト/ダーク）

# マトリクス上部（進捗/タブバー）の背景色（キャンバスより少し暗め）
MATRIX_TOP_BG_COLOR = ("#F0F0F0", "#242424")

# --- フォント定義 (変更なし) ---
FONT_BOLD = ("bold", 14)
FONT_NORMAL = (None, 12)
FONT_LARGE_BOLD = (None, 24, "bold")

# --- サイズ・ジオメトリ定義 (変更なし) ---
ACTION_SELECTOR_BUTTON_HEIGHT = 30
ACTION_SELECTOR_SPACING = 5
ACTION_SELECTOR_MARGIN = 10
ACTION_SELECTOR_GEOMETRY = "350x500"
HIGHLIGHT_BORDER_WIDTH = 1

NOTIFICATION_POPUP_INITIAL_WIDTH = 500
NOTIFICATION_POPUP_ALPHA = 0.9

PROMPT_EDITOR_GEOMETRY = "500x800"
PROMPT_SELECTION_GEOMETRY = "400x800"
SETTINGS_WINDOW_GEOMETRY = "520x460"
MAIN_WINDOW_GEOMETRY = "480x600"

HISTORY_FRAME_HEIGHT = 200

# --- Matrix Batch Processor ---
# マトリックス関連の色もモノクロームに統一
MATRIX_WINDOW_GEOMETRY = "1400x700"
MATRIX_HEADER_BORDER_COLOR = "#A0A0A0"  # 明るいグレーの境界線
MATRIX_CELL_BORDER_COLOR = "#404040"    # よりコントラストの高いグレーの境界線
MATRIX_DELETE_BUTTON_COLOR = "#404040"  # よりコントラストの高いグレーの境界線
# MATRIX_DELETE_BUTTON_COLOR = "#C44141"  # 彩度を抑えた赤
MATRIX_DELETE_BUTTON_HOVER_COLOR = "#A43737" # 少し濃い赤
MATRIX_FONT_BOLD = (None, 12, "bold")
MATRIX_RESULT_FONT = (None, 10)
MATRIX_IMAGE_THUMBNAIL_SIZE = (50, 50)
MATRIX_POPUP_GEOMETRY = "400x400"
MATRIX_RESULT_CELL_HEIGHT = 60
MATRIX_CELL_WIDTH = 250 # セルの固定幅を追加
# 固定（左端）列の推奨最小幅（行番号・削除ボタン・入力・添付/履歴ボタンが収まる程度）
MATRIX_FIXED_COL_WIDTH = 240
MATRIX_SUMMARY_FG_COLOR = ("#FAFAFA", "#3A3A3A")  # 行・列・最終まとめ用に少し明るい背景

HISTORY_SELECTOR_POPUP_WIDTH = 350
HISTORY_SELECTOR_POPUP_MAX_HEIGHT = 400

# メインキャンバスの背景色
MATRIX_CANVAS_BACKGROUND_COLOR = ("#E0E0E0", "#2C2C2C") # 明るいグレー / 暗いグレー（ダークは#2C2C2Cに統一）

# --- Drag & drop highlight colors ---
# 行ドラッグ時に現在掴んでいる行を示す背景色。
# 左がライトモード用、右がダークモード用の色です。
DRAG_ACTIVE_ROW_COLOR = ("#DDE6F7", "#3A4A6F")
# ドロップターゲットを示す背景色。アクティブ行とは異なる色で区別します。
DRAG_TARGET_ROW_COLOR = ("#C8DDF0", "#2F3F61")

# --- Popup スタイル ---
# ポップアップウィンドウの背景色とテキスト色を統一します。
POPUP_BG_COLOR = ("#F5F5F5", "#2C2C2C")    # 明るめグレー / ダークグレー
POPUP_TEXT_COLOR = ("#1C1C1C", "#EAEAEA")   # テキスト色（オフブラック / オフホワイト）

# --- 添付ファイル表示エリア ---
# 左側の添付ファイル名表示パネルの枠と背景
ATTACH_AREA_BORDER_COLOR = ("#F7DC6F", "#B7950B")  # 落ち着いた黄色系
ATTACH_AREA_BG_COLOR = ("#FFF9E6", "#3A3A2A")      # 薄いクリーム/ダーク黄土
