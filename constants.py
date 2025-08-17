APP_NAME = "Geminiクリップボード"
CONFIG_FILE = "config.json"
API_SERVICE_ID = "gemini_clip"
ICON_FILE = "icon.ico"
COMPLETION_SOUND_FILE = "completion.mp3"
DELETE_ICON_FILE = "delete_icon.png"  # 任意の16-24px程度のPNGを配置すると使用されます

# Token prices in USD per 1000 tokens (example values, replace with actual prices)
TOKEN_PRICES = {
    "gemini-2.5-flash-lite": {"input": 0.000000, "output": 0.000000}, # Placeholder, replace with actual prices
    "gemini-2.5-flash": {"input": 0.000000, "output": 0.000000},   # Placeholder, replace with actual prices
    "gemini-2.5-pro": {"input": 0.000000, "output": 0.000000},     # Placeholder, replace with actual prices
}

# Supported model options for UI selection (id, label)
SUPPORTED_MODELS: list[tuple[str, str]] = [
    ("gemini-2.5-flash-lite", "gemini-2.5-flash-lite (高速、低精度)"),
    ("gemini-2.5-flash", "gemini-2.5-flash (普通)"),
    ("gemini-2.5-pro", "gemini-2.5-pro (低速、高精度)"),
]

def model_id_to_label(model_id: str) -> str:
    for mid, label in SUPPORTED_MODELS:
        if mid == model_id:
            return label
    return SUPPORTED_MODELS[0][1]

def model_label_to_id(label: str) -> str:
    # Option menu stores full label; split-safe mapping
    for mid, lbl in SUPPORTED_MODELS:
        if lbl == label:
            return mid
    # Fallback: if given an id-like string (e.g., from legacy config), accept it
    return label.split(" ")[0]
