"""Fractionate Edge — FastAPI Backend

Orchestrates local AI models, manages conversations, handles document processing,
and provides streaming chat via SSE. All services bind to 127.0.0.1 only.
"""

import asyncio
import hashlib
import json
import logging
import os
import platform
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Optional

import httpx
import yaml
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from sse_starlette.sse import EventSourceResponse

from database import Database
from models import Florence2Manager, LlamaServerManager

# ── Logging ──────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
)
logger = logging.getLogger("fractionate")

# ── Configuration ────────────────────────────────────────────────

BASE_DIR = Path(os.environ.get("FRACTIONATE_HOME", Path.home() / ".fractionate"))
CONFIG_PATH = BASE_DIR / "config.yaml"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH) as f:
            return yaml.safe_load(f) or {}
    return {}


config = load_config()

DB_PATH = str(
    Path(os.path.expanduser(
        config.get("database", {}).get("path", str(BASE_DIR / "data.db"))
    ))
)

models_cfg = config.get("models", {})
falcon_cfg = models_cfg.get("falcon3_7b", {})
florence_cfg = models_cfg.get("florence2", {})

FALCON_MODEL_PATH = os.path.expanduser(
    falcon_cfg.get("path", str(BASE_DIR / "models" / "falcon3-7b-1.58bit" / "model.gguf"))
)
FALCON_CONTEXT_SIZE = falcon_cfg.get("context_size", 4096)
FALCON_AUTO_START = falcon_cfg.get("auto_start", False)

# Determine llama-server binary path based on platform
if platform.system() == "Windows":
    LLAMA_BINARY = str(BASE_DIR / "bitnet" / "build" / "bin" / "Release" / "llama-server.exe")
else:
    LLAMA_BINARY = str(BASE_DIR / "bitnet" / "build" / "bin" / "llama-server")

FLORENCE_MODEL_PATH = os.path.expanduser(
    florence_cfg.get("path", str(BASE_DIR / "models" / "florence-2-base"))
)

# Thread configuration
falcon_threads = falcon_cfg.get("threads", "auto")
FALCON_THREADS = None if falcon_threads == "auto" else int(falcon_threads)

# ── Globals ──────────────────────────────────────────────────────

db = Database(DB_PATH)
llama = LlamaServerManager(LLAMA_BINARY, FALCON_MODEL_PATH)
florence = Florence2Manager(FLORENCE_MODEL_PATH)


# ── App Lifecycle ────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Starting Fractionate Edge backend")
    logger.info("Database: %s", DB_PATH)
    logger.info("Falcon model: %s", FALCON_MODEL_PATH)
    logger.info("Florence model: %s", FLORENCE_MODEL_PATH)

    await db.connect()
    await db.log_model_event("system", "start", "Backend started")

    if FALCON_AUTO_START:
        try:
            await llama.start(threads=FALCON_THREADS, context_size=FALCON_CONTEXT_SIZE)
            await db.log_model_event("falcon3-7b", "start", "Auto-started")
        except Exception as e:
            logger.error("Failed to auto-start Falcon: %s", e)
            await db.log_model_event("falcon3-7b", "error", str(e))

    yield

    logger.info("Shutting down Fractionate Edge backend")
    if llama.is_running:
        await llama.stop()
        await db.log_model_event("falcon3-7b", "stop", "Backend shutdown")
    if florence.is_loaded:
        florence.unload()
        await db.log_model_event("florence2", "unloaded", "Backend shutdown")
    await db.close()


app = FastAPI(title="Fractionate Edge", version="0.1.0", lifespan=lifespan)


# ── Request / Response Models ────────────────────────────────────

class ChatRequest(BaseModel):
    conversation_id: str
    message: str
    system_prompt: Optional[str] = None


class CreateConversationRequest(BaseModel):
    title: str = "New Conversation"


class UpdateConversationRequest(BaseModel):
    title: str


# ── Health ───────────────────────────────────────────────────────

def _system_info() -> dict:
    import shutil

    try:
        import psutil
        total_ram = psutil.virtual_memory().total // (1024 * 1024)
        available_ram = psutil.virtual_memory().available // (1024 * 1024)
    except ImportError:
        total_ram = 0
        available_ram = 0

    return {
        "total_ram_mb": total_ram,
        "available_ram_mb": available_ram,
        "cpu_cores": os.cpu_count() or 0,
        "platform": platform.system().lower(),
    }


@app.get("/api/health")
async def health():
    conv_count = await db.get_conversations_count() if db.is_connected else 0
    doc_count = await db.get_cached_documents_count() if db.is_connected else 0

    return {
        "status": "ok",
        "models": {
            "falcon3_7b": llama.get_status(),
            "florence2": florence.get_status(),
        },
        "system": _system_info(),
        "database": {
            "connected": db.is_connected,
            "conversations_count": conv_count,
            "cached_documents": doc_count,
        },
    }


# ── Model Management ────────────────────────────────────────────

@app.post("/api/models/{model_name}/start")
async def start_model(model_name: str):
    if model_name == "falcon3-7b":
        if llama.is_running:
            return {"status": "already_running", "pid": llama.pid}
        try:
            await llama.start(threads=FALCON_THREADS, context_size=FALCON_CONTEXT_SIZE)
            await db.log_model_event("falcon3-7b", "start", f"pid={llama.pid}")
            return {"status": "started", "pid": llama.pid}
        except Exception as e:
            await db.log_model_event("falcon3-7b", "error", str(e))
            raise HTTPException(status_code=500, detail=str(e))

    elif model_name == "florence2":
        if florence.is_loaded:
            return {"status": "already_loaded"}
        try:
            florence.load()
            await db.log_model_event("florence2", "loaded")
            return {"status": "loaded"}
        except Exception as e:
            await db.log_model_event("florence2", "error", str(e))
            raise HTTPException(status_code=500, detail=str(e))

    else:
        raise HTTPException(status_code=404, detail=f"Unknown model: {model_name}")


@app.post("/api/models/{model_name}/stop")
async def stop_model(model_name: str):
    if model_name == "falcon3-7b":
        if not llama.is_running:
            return {"status": "not_running"}
        await llama.stop()
        await db.log_model_event("falcon3-7b", "stop")
        return {"status": "stopped"}

    elif model_name == "florence2":
        if not florence.is_loaded:
            return {"status": "not_loaded"}
        florence.unload()
        await db.log_model_event("florence2", "unloaded")
        return {"status": "unloaded"}

    else:
        raise HTTPException(status_code=404, detail=f"Unknown model: {model_name}")


# ── Conversations ────────────────────────────────────────────────

@app.get("/api/conversations")
async def list_conversations():
    conversations = await db.list_conversations()
    return {
        "conversations": [
            {
                "id": c.id,
                "title": c.title,
                "updated_at": c.updated_at,
                "message_count": c.message_count,
                "has_documents": c.has_documents,
            }
            for c in conversations
        ]
    }


@app.post("/api/conversations")
async def create_conversation(req: CreateConversationRequest):
    conv = await db.create_conversation(req.title)
    return {
        "id": conv.id,
        "title": conv.title,
        "created_at": conv.created_at,
        "updated_at": conv.updated_at,
    }


@app.get("/api/conversations/{conv_id}")
async def get_conversation(conv_id: str):
    conv = await db.get_conversation(conv_id)
    if not conv:
        raise HTTPException(status_code=404, detail="Conversation not found")

    messages = await db.get_messages(conv_id)
    return {
        "id": conv.id,
        "title": conv.title,
        "created_at": conv.created_at,
        "updated_at": conv.updated_at,
        "message_count": conv.message_count,
        "has_documents": conv.has_documents,
        "messages": [
            {
                "id": m.id,
                "role": m.role,
                "content": m.content,
                "has_document": m.has_document,
                "document_name": m.document_name,
                "document_type": m.document_type,
                "created_at": m.created_at,
            }
            for m in messages
        ],
    }


@app.put("/api/conversations/{conv_id}")
async def update_conversation(conv_id: str, req: UpdateConversationRequest):
    conv = await db.update_conversation(conv_id, req.title)
    if not conv:
        raise HTTPException(status_code=404, detail="Conversation not found")
    return {"id": conv.id, "title": conv.title, "updated_at": conv.updated_at}


@app.delete("/api/conversations/{conv_id}")
async def delete_conversation(conv_id: str):
    deleted = await db.delete_conversation(conv_id)
    if not deleted:
        raise HTTPException(status_code=404, detail="Conversation not found")
    return {"status": "deleted"}


# ── Chat (Streaming SSE) ────────────────────────────────────────

async def _auto_title(conversation_id: str, first_message: str):
    """Generate a short title from the first message using the LLM."""
    if not llama.is_running:
        # Fallback: use first few words
        words = first_message.split()[:5]
        title = " ".join(words)
        if len(words) == 5:
            title += "..."
        await db.update_conversation(conversation_id, title)
        return

    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            r = await client.post(
                f"http://127.0.0.1:{llama.port}/v1/chat/completions",
                json={
                    "model": "falcon3-7b",
                    "messages": [
                        {
                            "role": "user",
                            "content": f"Generate a short 3-5 word title for this conversation. Reply with ONLY the title, nothing else.\n\nUser message: {first_message[:500]}",
                        }
                    ],
                    "stream": False,
                    "temperature": 0.3,
                    "max_tokens": 20,
                },
            )
            if r.status_code == 200:
                data = r.json()
                title = data["choices"][0]["message"]["content"].strip().strip('"\'')
                if title:
                    await db.update_conversation(conversation_id, title[:100])
    except Exception as e:
        logger.warning("Auto-title failed: %s", e)


@app.post("/api/chat")
async def chat(request: ChatRequest):
    # Verify conversation exists
    conv = await db.get_conversation(request.conversation_id)
    if not conv:
        raise HTTPException(status_code=404, detail="Conversation not found")

    if not llama.is_running:
        raise HTTPException(status_code=503, detail="Falcon3-7B is not running. Start it first.")

    # Save user message
    user_msg = await db.save_message(request.conversation_id, "user", request.message)

    # Build message history
    history = await db.get_messages(request.conversation_id, limit=20)
    messages = []

    # Add system prompt if provided
    if request.system_prompt:
        messages.append({"role": "system", "content": request.system_prompt})
    else:
        messages.append({
            "role": "system",
            "content": "You are a helpful, concise AI assistant running locally on the user's machine via Fractionate Edge. You help with questions, document analysis, and general tasks.",
        })

    for m in history:
        messages.append({"role": m.role, "content": m.content})

    # Check if this is the first exchange (for auto-titling)
    is_first = conv.message_count <= 1

    msg_id = str(uuid.uuid4())

    async def stream_response():
        full_response = ""
        token_count = 0
        start_time = time.time()

        yield {"data": json.dumps({"type": "start", "message_id": msg_id})}

        try:
            async with httpx.AsyncClient(timeout=300.0) as client:
                async with client.stream(
                    "POST",
                    f"http://127.0.0.1:{llama.port}/v1/chat/completions",
                    json={
                        "model": "falcon3-7b",
                        "messages": messages,
                        "stream": True,
                        "temperature": 0.7,
                        "max_tokens": 2048,
                    },
                ) as response:
                    async for line in response.aiter_lines():
                        if line.startswith("data: "):
                            data = line[6:]
                            if data.strip() == "[DONE]":
                                break
                            try:
                                chunk = json.loads(data)
                                delta = chunk.get("choices", [{}])[0].get("delta", {})
                                token = delta.get("content", "")
                                if token:
                                    full_response += token
                                    token_count += 1
                                    yield {
                                        "data": json.dumps({
                                            "type": "token",
                                            "content": token,
                                        })
                                    }
                            except json.JSONDecodeError:
                                continue
        except httpx.ConnectError:
            yield {
                "data": json.dumps({
                    "type": "error",
                    "content": "Lost connection to Falcon3-7B. The model may have crashed.",
                })
            }
            return
        except Exception as e:
            logger.error("Streaming error: %s", e)
            yield {
                "data": json.dumps({
                    "type": "error",
                    "content": f"Error during generation: {str(e)}",
                })
            }
            return

        # Save assistant response
        elapsed = time.time() - start_time
        tps = token_count / elapsed if elapsed > 0 else 0

        if full_response:
            await db.save_message(request.conversation_id, "assistant", full_response)

        yield {
            "data": json.dumps({
                "type": "done",
                "message_id": msg_id,
                "total_tokens": token_count,
                "tokens_per_second": round(tps, 1),
            })
        }

        # Auto-title after first exchange
        if is_first and full_response:
            await _auto_title(request.conversation_id, request.message)

    return EventSourceResponse(stream_response())


# ── Document Processing ──────────────────────────────────────────

def _compute_file_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _extract_text_pdf(file_data: bytes) -> tuple[str, dict]:
    """Extract text from a text-based PDF using PyMuPDF."""
    import fitz  # PyMuPDF

    doc = fitz.open(stream=file_data, filetype="pdf")
    pages_text = []
    for page in doc:
        text = page.get_text()
        if text.strip():
            pages_text.append(text)
    doc.close()

    combined = "\n\n---\n\n".join(pages_text)
    metadata = {"pages": len(pages_text), "method": "text_pdf"}
    return combined, metadata


def _pdf_to_images(file_data: bytes) -> list:
    """Convert PDF pages to PIL Images for vision processing."""
    import fitz
    from PIL import Image
    import io

    doc = fitz.open(stream=file_data, filetype="pdf")
    images = []
    for page in doc:
        # Render at 150 DPI for balance of quality and speed
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data)).convert("RGB")
        images.append(img)
    doc.close()
    return images


def _is_scanned_pdf(file_data: bytes) -> bool:
    """Heuristic: if a PDF has very little text per page, it's likely scanned."""
    import fitz

    doc = fitz.open(stream=file_data, filetype="pdf")
    total_chars = 0
    for page in doc:
        total_chars += len(page.get_text().strip())
    doc.close()

    chars_per_page = total_chars / max(len(doc), 1)
    return chars_per_page < 50  # Less than 50 chars per page = likely scanned


@app.post("/api/documents")
async def upload_document(
    file: UploadFile = File(...),
    conversation_id: Optional[str] = Form(None),
):
    start_time = time.time()
    file_data = await file.read()
    file_hash = _compute_file_hash(file_data)
    filename = file.filename or "unknown"
    file_ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    # Check cache
    cached = await db.get_cached_document(file_hash)
    if cached:
        return {
            "document_id": cached.id,
            "original_name": cached.original_name,
            "extracted_text": cached.extracted_text,
            "extraction_method": cached.extraction_method,
            "metadata": cached.metadata,
            "cached": True,
            "processing_time_seconds": round(time.time() - start_time, 2),
        }

    extracted_text = ""
    extraction_method = ""
    metadata = {}

    try:
        if file_ext == "pdf":
            if _is_scanned_pdf(file_data):
                # Scanned PDF — use Florence-2 vision
                images = _pdf_to_images(file_data)
                page_texts = []
                for i, img in enumerate(images):
                    result = florence.process_image(img)
                    caption = _extract_caption_text(result.get("caption", {}))
                    ocr = _extract_ocr_text(result.get("ocr", {}))
                    page_text = f"[Page {i + 1}]\n"
                    if caption:
                        page_text += f"Description: {caption}\n"
                    if ocr:
                        page_text += f"Text: {ocr}\n"
                    page_texts.append(page_text)
                extracted_text = "\n\n".join(page_texts)
                extraction_method = "florence2"
                metadata = {"pages": len(images)}
            else:
                # Text PDF — fast extraction
                extracted_text, metadata = _extract_text_pdf(file_data)
                extraction_method = "text_pdf"

        elif file_ext in ("jpg", "jpeg", "png", "tiff", "tif", "bmp", "webp"):
            from PIL import Image
            import io

            img = Image.open(io.BytesIO(file_data)).convert("RGB")
            result = florence.process_image(img)
            caption = _extract_caption_text(result.get("caption", {}))
            ocr = _extract_ocr_text(result.get("ocr", {}))
            parts = []
            if caption:
                parts.append(f"Description: {caption}")
            if ocr:
                parts.append(f"Text: {ocr}")
            extracted_text = "\n".join(parts) if parts else "No text could be extracted from this image."
            extraction_method = "florence2"
            metadata = {"width": img.width, "height": img.height}

        elif file_ext == "docx":
            import docx
            import io

            doc = docx.Document(io.BytesIO(file_data))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            extracted_text = "\n\n".join(paragraphs)
            extraction_method = "docx"
            metadata = {"paragraphs": len(paragraphs)}

        else:
            # Try to read as plain text
            try:
                extracted_text = file_data.decode("utf-8")
                extraction_method = "plaintext"
            except UnicodeDecodeError:
                raise HTTPException(
                    status_code=400,
                    detail=f"Unsupported file type: .{file_ext}",
                )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Document processing error: %s", e)
        raise HTTPException(status_code=500, detail=f"Failed to process document: {str(e)}")
    finally:
        # Unload Florence-2 after processing to free RAM
        if florence.is_loaded:
            florence.unload()
            await db.log_model_event("florence2", "unloaded", "Auto-unload after document processing")

    # Cache the result
    doc_record = await db.cache_document(
        file_hash=file_hash,
        original_name=filename,
        extracted_text=extracted_text,
        extraction_method=extraction_method,
        metadata=metadata,
    )

    elapsed = round(time.time() - start_time, 2)

    return {
        "document_id": doc_record.id,
        "original_name": filename,
        "pages": metadata.get("pages", 1),
        "extracted_text": extracted_text,
        "extraction_method": extraction_method,
        "metadata": metadata,
        "cached": False,
        "processing_time_seconds": elapsed,
    }


def _extract_caption_text(caption_result) -> str:
    """Extract plain text from Florence-2 caption result."""
    if isinstance(caption_result, str):
        return caption_result
    if isinstance(caption_result, dict):
        return caption_result.get("<MORE_DETAILED_CAPTION>", str(caption_result))
    return str(caption_result)


def _extract_ocr_text(ocr_result) -> str:
    """Extract plain text from Florence-2 OCR result."""
    if isinstance(ocr_result, str):
        return ocr_result
    if isinstance(ocr_result, dict):
        # OCR_WITH_REGION returns {<OCR_WITH_REGION>: {'quad_boxes': [...], 'labels': [...]}}
        region_data = ocr_result.get("<OCR_WITH_REGION>", {})
        if isinstance(region_data, dict) and "labels" in region_data:
            return " ".join(region_data["labels"])
        return str(ocr_result)
    return str(ocr_result)


# ── Run ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8081,
        log_level="info",
    )
