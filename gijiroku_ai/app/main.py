import os
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, Header, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse
from dotenv import load_dotenv

from linebot.v3 import WebhookParser
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    ApiClient,
    Configuration,
    MessagingApi,
    MessagingApiBlob,
    ReplyMessageRequest,
    PushMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import MessageEvent, AudioMessageContent, FileMessageContent

from app.gemini_handler import generate_minutes

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

LINE_CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]

parser = WebhookParser(LINE_CHANNEL_SECRET)
line_config = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("議事録Bot 起動完了")
    yield
    logger.info("議事録Bot 停止")


app = FastAPI(title="議事録Bot", lifespan=lifespan)


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/webhook")
async def webhook(
    request: Request,
    background_tasks: BackgroundTasks,
    x_line_signature: str = Header(alias="X-Line-Signature"),
):
    body = await request.body()
    body_text = body.decode("utf-8")

    try:
        events = parser.parse(body_text, x_line_signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")

    for event in events:
        if isinstance(event, MessageEvent) and isinstance(event.message, (AudioMessageContent, FileMessageContent)):
            # すぐに処理中メッセージを送信
            _reply_text(event.reply_token, "音声を受信しました。議事録を作成中です...\n（1〜2分かかる場合があります）")
            # 音声処理はバックグラウンドで実行
            background_tasks.add_task(process_audio, event.message.id, event.source.user_id)

    # LINEへは即座に200 OKを返す
    return JSONResponse(content={"status": "ok"})


def process_audio(message_id: str, user_id: str):
    try:
        logger.info(f"音声処理開始: message_id={message_id}")
        audio_bytes = _download_audio(message_id)
        logger.info(f"音声ダウンロード完了: {len(audio_bytes)} bytes")
        minutes = generate_minutes(audio_bytes, mime_type="audio/m4a")
        logger.info("議事録生成完了")
        _push_long_text(user_id, minutes)
    except Exception as e:
        logger.error(f"議事録生成エラー: {e}")
        _push_text(user_id, f"エラーが発生しました。\n{str(e)}")


def _download_audio(message_id: str) -> bytes:
    with ApiClient(line_config) as api_client:
        blob_api = MessagingApiBlob(api_client)
        return blob_api.get_message_content(message_id)


def _reply_text(reply_token: str, text: str):
    with ApiClient(line_config) as api_client:
        messaging_api = MessagingApi(api_client)
        messaging_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=text)],
            )
        )


def _push_text(user_id: str, text: str):
    with ApiClient(line_config) as api_client:
        messaging_api = MessagingApi(api_client)
        messaging_api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=text)],
            )
        )


def _push_long_text(user_id: str, text: str, chunk_size: int = 4900):
    chunks = [text[i : i + chunk_size] for i in range(0, len(text), chunk_size)]
    for chunk in chunks:
        _push_text(user_id, chunk)
