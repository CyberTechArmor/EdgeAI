"""Fractionate Edge — SQLite Database Manager

Handles all database operations: schema creation, CRUD for conversations,
messages, document cache, model logs, and configuration.
"""

import aiosqlite
import uuid
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Conversation:
    id: str
    title: str
    created_at: str
    updated_at: str
    message_count: int = 0
    has_documents: bool = False


@dataclass
class Message:
    id: str
    conversation_id: str
    role: str
    content: str
    has_document: bool = False
    document_name: Optional[str] = None
    document_type: Optional[str] = None
    created_at: str = ""


@dataclass
class CachedDocument:
    id: str
    file_hash: str
    original_name: str
    extracted_text: str
    extraction_method: str
    metadata: dict = field(default_factory=dict)
    created_at: str = ""


SCHEMA = """
CREATE TABLE IF NOT EXISTS conversations (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL DEFAULT 'New Conversation',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS messages (
    id TEXT PRIMARY KEY,
    conversation_id TEXT NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    has_document BOOLEAN DEFAULT FALSE,
    document_name TEXT,
    document_type TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (conversation_id) REFERENCES conversations(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS document_cache (
    id TEXT PRIMARY KEY,
    file_hash TEXT UNIQUE NOT NULL,
    original_name TEXT NOT NULL,
    extracted_text TEXT NOT NULL,
    extraction_method TEXT,
    metadata TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS model_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model_name TEXT NOT NULL,
    event TEXT NOT NULL,
    details TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS config (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);
"""


class Database:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._db: Optional[aiosqlite.Connection] = None

    async def connect(self):
        """Open database connection and initialize schema."""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self._db = await aiosqlite.connect(self.db_path)
        self._db.row_factory = aiosqlite.Row
        await self._db.execute("PRAGMA journal_mode=WAL")
        await self._db.execute("PRAGMA foreign_keys=ON")
        await self._db.executescript(SCHEMA)
        await self._db.commit()

    async def close(self):
        """Close database connection."""
        if self._db:
            await self._db.close()
            self._db = None

    @property
    def is_connected(self) -> bool:
        return self._db is not None

    # ── Conversations ────────────────────────────────────────────

    async def create_conversation(self, title: str = "New Conversation") -> Conversation:
        conv_id = str(uuid.uuid4())
        now = datetime.now(timezone.utc).isoformat()
        await self._db.execute(
            "INSERT INTO conversations (id, title, created_at, updated_at) VALUES (?, ?, ?, ?)",
            (conv_id, title, now, now),
        )
        await self._db.commit()
        return Conversation(id=conv_id, title=title, created_at=now, updated_at=now)

    async def get_conversation(self, conv_id: str) -> Optional[Conversation]:
        cursor = await self._db.execute(
            "SELECT * FROM conversations WHERE id = ?", (conv_id,)
        )
        row = await cursor.fetchone()
        if not row:
            return None
        msg_count = await self._count_messages(conv_id)
        has_docs = await self._has_documents(conv_id)
        return Conversation(
            id=row["id"],
            title=row["title"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
            message_count=msg_count,
            has_documents=has_docs,
        )

    async def list_conversations(self) -> list[Conversation]:
        cursor = await self._db.execute(
            "SELECT * FROM conversations ORDER BY updated_at DESC"
        )
        rows = await cursor.fetchall()
        conversations = []
        for row in rows:
            msg_count = await self._count_messages(row["id"])
            has_docs = await self._has_documents(row["id"])
            conversations.append(
                Conversation(
                    id=row["id"],
                    title=row["title"],
                    created_at=row["created_at"],
                    updated_at=row["updated_at"],
                    message_count=msg_count,
                    has_documents=has_docs,
                )
            )
        return conversations

    async def update_conversation(self, conv_id: str, title: str) -> Optional[Conversation]:
        now = datetime.now(timezone.utc).isoformat()
        cursor = await self._db.execute(
            "UPDATE conversations SET title = ?, updated_at = ? WHERE id = ?",
            (title, now, conv_id),
        )
        await self._db.commit()
        if cursor.rowcount == 0:
            return None
        return await self.get_conversation(conv_id)

    async def delete_conversation(self, conv_id: str) -> bool:
        cursor = await self._db.execute(
            "DELETE FROM conversations WHERE id = ?", (conv_id,)
        )
        await self._db.commit()
        return cursor.rowcount > 0

    async def touch_conversation(self, conv_id: str):
        """Update the updated_at timestamp."""
        now = datetime.now(timezone.utc).isoformat()
        await self._db.execute(
            "UPDATE conversations SET updated_at = ? WHERE id = ?", (now, conv_id)
        )
        await self._db.commit()

    # ── Messages ─────────────────────────────────────────────────

    async def save_message(
        self,
        conversation_id: str,
        role: str,
        content: str,
        has_document: bool = False,
        document_name: Optional[str] = None,
        document_type: Optional[str] = None,
    ) -> Message:
        msg_id = str(uuid.uuid4())
        now = datetime.now(timezone.utc).isoformat()
        await self._db.execute(
            """INSERT INTO messages
               (id, conversation_id, role, content, has_document, document_name, document_type, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (msg_id, conversation_id, role, content, has_document, document_name, document_type, now),
        )
        await self._db.commit()
        await self.touch_conversation(conversation_id)
        return Message(
            id=msg_id,
            conversation_id=conversation_id,
            role=role,
            content=content,
            has_document=has_document,
            document_name=document_name,
            document_type=document_type,
            created_at=now,
        )

    async def get_messages(self, conversation_id: str, limit: int = 50) -> list[Message]:
        cursor = await self._db.execute(
            """SELECT * FROM messages
               WHERE conversation_id = ?
               ORDER BY created_at ASC
               LIMIT ?""",
            (conversation_id, limit),
        )
        rows = await cursor.fetchall()
        return [
            Message(
                id=row["id"],
                conversation_id=row["conversation_id"],
                role=row["role"],
                content=row["content"],
                has_document=bool(row["has_document"]),
                document_name=row["document_name"],
                document_type=row["document_type"],
                created_at=row["created_at"],
            )
            for row in rows
        ]

    # ── Document Cache ───────────────────────────────────────────

    async def get_cached_document(self, file_hash: str) -> Optional[CachedDocument]:
        cursor = await self._db.execute(
            "SELECT * FROM document_cache WHERE file_hash = ?", (file_hash,)
        )
        row = await cursor.fetchone()
        if not row:
            return None
        metadata = json.loads(row["metadata"]) if row["metadata"] else {}
        return CachedDocument(
            id=row["id"],
            file_hash=row["file_hash"],
            original_name=row["original_name"],
            extracted_text=row["extracted_text"],
            extraction_method=row["extraction_method"],
            metadata=metadata,
            created_at=row["created_at"],
        )

    async def cache_document(
        self,
        file_hash: str,
        original_name: str,
        extracted_text: str,
        extraction_method: str,
        metadata: Optional[dict] = None,
    ) -> CachedDocument:
        doc_id = str(uuid.uuid4())
        now = datetime.now(timezone.utc).isoformat()
        meta_json = json.dumps(metadata) if metadata else None
        await self._db.execute(
            """INSERT INTO document_cache
               (id, file_hash, original_name, extracted_text, extraction_method, metadata, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (doc_id, file_hash, original_name, extracted_text, extraction_method, meta_json, now),
        )
        await self._db.commit()
        return CachedDocument(
            id=doc_id,
            file_hash=file_hash,
            original_name=original_name,
            extracted_text=extracted_text,
            extraction_method=extraction_method,
            metadata=metadata or {},
            created_at=now,
        )

    async def get_cached_documents_count(self) -> int:
        cursor = await self._db.execute("SELECT COUNT(*) as cnt FROM document_cache")
        row = await cursor.fetchone()
        return row["cnt"]

    # ── Model Log ────────────────────────────────────────────────

    async def log_model_event(self, model_name: str, event: str, details: Optional[str] = None):
        await self._db.execute(
            "INSERT INTO model_log (model_name, event, details) VALUES (?, ?, ?)",
            (model_name, event, details),
        )
        await self._db.commit()

    async def get_model_log(self, model_name: str, limit: int = 20) -> list[dict]:
        cursor = await self._db.execute(
            """SELECT * FROM model_log
               WHERE model_name = ?
               ORDER BY created_at DESC
               LIMIT ?""",
            (model_name, limit),
        )
        rows = await cursor.fetchall()
        return [
            {
                "id": row["id"],
                "model_name": row["model_name"],
                "event": row["event"],
                "details": row["details"],
                "created_at": row["created_at"],
            }
            for row in rows
        ]

    # ── Config ───────────────────────────────────────────────────

    async def get_config(self, key: str) -> Optional[str]:
        cursor = await self._db.execute(
            "SELECT value FROM config WHERE key = ?", (key,)
        )
        row = await cursor.fetchone()
        return row["value"] if row else None

    async def set_config(self, key: str, value: str):
        now = datetime.now(timezone.utc).isoformat()
        await self._db.execute(
            """INSERT INTO config (key, value, updated_at) VALUES (?, ?, ?)
               ON CONFLICT(key) DO UPDATE SET value = ?, updated_at = ?""",
            (key, value, now, value, now),
        )
        await self._db.commit()

    # ── Stats ────────────────────────────────────────────────────

    async def get_conversations_count(self) -> int:
        cursor = await self._db.execute("SELECT COUNT(*) as cnt FROM conversations")
        row = await cursor.fetchone()
        return row["cnt"]

    # ── Helpers ──────────────────────────────────────────────────

    async def _count_messages(self, conversation_id: str) -> int:
        cursor = await self._db.execute(
            "SELECT COUNT(*) as cnt FROM messages WHERE conversation_id = ?",
            (conversation_id,),
        )
        row = await cursor.fetchone()
        return row["cnt"]

    async def _has_documents(self, conversation_id: str) -> bool:
        cursor = await self._db.execute(
            "SELECT COUNT(*) as cnt FROM messages WHERE conversation_id = ? AND has_document = TRUE",
            (conversation_id,),
        )
        row = await cursor.fetchone()
        return row["cnt"] > 0
