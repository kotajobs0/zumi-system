import os
import tempfile
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-1.5-pro")

MINUTES_PROMPT = """
あなたは優秀な議事録作成アシスタントです。
以下の音声を文字起こしし、議事録を日本語で作成してください。

【出力フォーマット】

■ 議事録

日時：（音声内に言及があれば記載、なければ「不明」）
参加者：（音声から読み取れる人物名、不明な場合は「不明」）

━━━━━━━━━━━━━━━━━━━━━━
【議題・話し合われた内容】
（箇条書きで簡潔に）

【決定事項】
（決定されたことを箇条書きで）

【TODO・次のアクション】
（誰が何をするかを箇条書きで）

【その他・補足】
（重要な補足があれば記載）
━━━━━━━━━━━━━━━━━━━━━━

【文字起こし全文】
（音声の全文テキスト）
"""


def generate_minutes(audio_bytes: bytes, mime_type: str = "audio/m4a") -> str:
    """
    音声バイトデータからGemini APIを使って議事録を生成する。
    """
    model = genai.GenerativeModel(MODEL_NAME)

    with tempfile.NamedTemporaryFile(suffix=".m4a", delete=False) as tmp:
        tmp.write(audio_bytes)
        tmp_path = tmp.name

    try:
        uploaded_file = genai.upload_file(path=tmp_path, mime_type=mime_type)

        response = model.generate_content(
            [MINUTES_PROMPT, uploaded_file],
            generation_config=genai.GenerationConfig(
                temperature=0.2,
                max_output_tokens=4096,
            ),
        )

        genai.delete_file(uploaded_file.name)
        return response.text

    finally:
        os.unlink(tmp_path)
