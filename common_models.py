from typing import List, Dict, Any, Optional, Literal, AsyncIterator
from pydantic import BaseModel, Field, ConfigDict
import google.generativeai as genai
from google.generativeai import types
from google.api_core import exceptions
import base64
from io import BytesIO
from PIL import Image

def create_image_part(image_data_base64: str | bytes) -> Dict[str, Any]:
    """
    Base64エンコードされた画像データ（またはバイト列）から、
    Gemini API 用の画像パートを生成する。

    - 入力が base64 文字列の場合: デコードしてバイト列へ。
    - zlib 圧縮が施されている場合: 可能なら解凍。
    - PIL で一度読み込み、PNG として正規化してから `inline_data` に格納。
    """
    # 1) bytes へ正規化
    raw: bytes
    if isinstance(image_data_base64, bytes):
        raw = image_data_base64
    else:
        raw = base64.b64decode(image_data_base64)

    # 2) 必要なら zlib 解凍（履歴で圧縮されている場合のフォールバック）
    try:
        # 典型的な zlib ヘッダ（0x78, 0x9C/0xDA 等）を簡易判定
        if len(raw) > 2 and raw[0] == 0x78 and raw[1] in (0x01, 0x5E, 0x9C, 0xDA):
            import zlib
            raw = zlib.decompress(raw)
    except Exception:
        # 解凍に失敗した場合は元のバイト列をそのまま使用
        pass

    # 3) PIL で読み込んで PNG 化（壊れたデータや非PNGでも正規化）
    try:
        with Image.open(BytesIO(raw)) as im:
            if im.mode != 'RGB':
                im = im.convert('RGB')
            buf = BytesIO()
            # optimize=True でサイズを抑えつつ互換性維持
            im.save(buf, format='PNG', optimize=True)
            png_bytes = buf.getvalue()
    except Exception:
        # 画像として読み込めない場合は、最終手段としてそのまま送る
        # （API 側で拒否される可能性はあるが、以前の挙動に近い）
        png_bytes = raw

    return {"inline_data": {"mime_type": "image/png", "data": png_bytes}}

class Event(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)
    content: Optional[types.GenerateContentResponse] = None
    is_final: bool = False

    def is_final_response(self) -> bool:
        return self.is_final

# ADKのモッククラス (ADKがインストールされていない場合のために一時的に定義)
class BaseAgent:
    def __init__(self, name: str, description: str = ""):
        self.name = name
        self.description = description
        self.parent_agent: Optional[BaseAgent] = None
        self.sub_agents: List[BaseAgent] = []

    async def run_async(self, *args, **kwargs) -> Any:
        raise NotImplementedError

    async def run_live(self, *args, **kwargs) -> Any:
        raise NotImplementedError

class LlmAgent(BaseAgent):
    def __init__(self, name: str, description: str = "", prompt_config: Optional['Prompt'] = None):
        super().__init__(name, description)
        self.prompt_config = prompt_config
        if prompt_config:
            self.model = prompt_config.model
            self.instruction = prompt_config.system_prompt
            self.temperature = prompt_config.parameters.temperature
        else:
            # デフォルト値またはエラーハンドリング
            self.model = "gemini-2.5-flash-lite"
            self.instruction = ""
            self.temperature = 1.0

    async def run_async(self, content: str) -> AsyncIterator[Event]:
        """LLM処理を実行する (非同期ジェネレータとしてストリーミングをサポート)"""
        if not self.prompt_config:
            raise ValueError("LlmAgentがprompt_configで初期化されていません。")

        model = genai.GenerativeModel(self.model)
        generation_config = genai.types.GenerationConfig(
            temperature=self.temperature
        )

        # LlmAgentはテキストコンテンツのみを処理すると仮定
        contents = f"{self.instruction}\n\n---\n\n{content}"
        
        try:
            print(f"DEBUG: LlmAgent '{self.name}' - LLMストリーミングリクエスト送信中。モデル: {self.model}, 温度: {self.temperature}, コンテンツ先頭: {contents[:100]}...")
            response_stream = await model.generate_content_async(contents, generation_config=generation_config, stream=True)
            
            async for chunk in response_stream:
                if chunk.parts: # chunk.textではなくchunk.partsを使用
                    print(f"DEBUG: LlmAgent '{self.name}' - LLMストリーミングチャンク受信。")
                    yield Event(content=chunk)
            yield Event(is_final=True) # ストリームの最後にis_final=TrueのEventを送信
        except exceptions.GoogleAPICallError as e:
            print(f"ERROR: LlmAgent '{self.name}' - APIエラーが発生しました (コード: {e.code}): {e.message}")
            raise RuntimeError(f"APIエラーが発生しました (コード: {e.code}): {e.message}")
        except Exception as e:
            print(f"ERROR: LlmAgent '{self.name}' - LLM呼び出し中に予期せぬエラーが発生しました: {e}")
            raise RuntimeError(f"LLM呼び出し中に予期せぬエラーが発生しました: {e}")

    async def run_live(self, content: Any) -> Any:
        # リアルタイム処理ロジック
        pass

class PromptParameters(BaseModel):
    temperature: float = 1.0
    top_p: Optional[float] = None
    top_k: Optional[int] = None
    max_output_tokens: Optional[int] = None
    stop_sequences: Optional[List[str]] = None

class Prompt(BaseModel):
    name: str
    model: Literal["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.5-flash-lite"] = "gemini-2.5-flash-lite"
    system_prompt: str
    thinking_level: Literal["Fast", "Balanced", "High Quality", "Unlimited"] = "Balanced"
    enable_web: bool = False
    parameters: PromptParameters = Field(default_factory=PromptParameters)

    # マトリクスプロンプトにデフォルトで含めるかどうかを示すフラグ
    include_in_matrix: bool = False

class AppConfig(BaseModel):
    """Top-level configuration model for the application.

    Attributes:
        version: Configuration schema version. Increment this when
            introducing breaking changes. Older versions will be migrated
            automatically by ``config_manager``.
        prompts: Mapping of prompt identifiers to prompt definitions.
        max_history_size: Maximum number of history items to retain in
            clipboard history.
        api_key: Optional API key for Gemini (or other LLM) services.
        hotkey_prompt_list: Optional hotkey for opening the prompt list.
        hotkey_refine: Optional hotkey for the refine dialog.
        hotkey_matrix: Optional hotkey for opening the matrix processor.
        hotkey: Deprecated single global hotkey (v2 and earlier). Kept for migration.
    """
    version: int = 6
    prompts: Dict[str, Prompt]
    max_history_size: int = 20
    api_key: Optional[str] = None
    # New granular hotkeys (v3)
    hotkey_prompt_list: Optional[str] = "ctrl+shift+c"
    hotkey_refine: Optional[str] = "ctrl+shift+r"
    hotkey_matrix: Optional[str] = None
    # Deprecated legacy field (v2). Keep for migration/back-compat reading.
    hotkey: Optional[str] = None
    # Matrix summary prompts (v4)
    matrix_row_summary_prompt: Optional[Prompt] = None
    matrix_col_summary_prompt: Optional[Prompt] = None
    matrix_matrix_summary_prompt: Optional[Prompt] = None
    # Flow execution settings (v5)
    max_flow_steps: int = 5
    # Language (v6)
    language: Optional[str] = "auto"
