import os
import uuid
from dataclasses import dataclass
from pathlib import Path

from ai import OllamaClient
from document import page_count, render_page

try:
    import chromadb
    from chromadb.api.types import Documents, EmbeddingFunction, Embeddings
    HAS_CHROMADB = True
except ImportError:
    HAS_CHROMADB = False

try:
    import pytesseract  # type: ignore
except ImportError:
    pytesseract = None


@dataclass
class DocumentChunk:
    id: str
    text: str
    source: str
    page: int



class OllamaEmbeddingFunction:
    def __init__(self, client: OllamaClient):
        self.client = client
    def __call__(self, input):
        return self.client.embed(list(input))
    def name(self) -> str:
        return "ollama_embedding_function"


if not HAS_CHROMADB:
    class RAGDatabase:
        """Stub when chromadb is not installed."""
        def __init__(self, db_path, ollama_client):
            pass
        def ingest_document(self, file_path):
            return 0
        def search(self, query, n_results=5):
            return []
        def clear(self):
            pass
else:
    class RAGDatabase:
        def __init__(self, db_path: Path, ollama_client: OllamaClient):
            self.db_path = db_path
            self.client = chromadb.PersistentClient(path=str(db_path))
            self.embed_func = OllamaEmbeddingFunction(ollama_client)
            self.collection = self.client.get_or_create_collection(
                name="notebook_sources", embedding_function=self.embed_func
            )

        def ingest_document(self, file_path: Path) -> int:
            text_pages = _extract_text_by_page(file_path)
            chunks = []
            for page_num, text in enumerate(text_pages):
                if not text.strip():
                    continue
                page_chunks = _chunk_text(text, chunk_size=1000, overlap=200)
                for chunk_text in page_chunks:
                    chunks.append(
                        DocumentChunk(
                            id=f"{file_path.name}_{page_num}_{uuid.uuid4().hex[:8]}",
                            text=chunk_text,
                            source=file_path.name,
                            page=page_num,
                        )
                    )
            if chunks:
                self.collection.add(
                    ids=[c.id for c in chunks],
                    documents=[c.text for c in chunks],
                    metadatas=[{"source": c.source, "page": c.page} for c in chunks],
                )
            return len(chunks)

        def search(self, query: str, n_results: int = 5) -> list[DocumentChunk]:
            results = self.collection.query(query_texts=[query], n_results=n_results)
            chunks = []
            if not results.get("documents") or not results["documents"][0]:
                return chunks
            docs = results["documents"][0]
            metas = results["metadatas"][0]
            ids = results["ids"][0]
            for i in range(len(docs)):
                chunks.append(
                    DocumentChunk(id=ids[i], text=docs[i],
                                  source=metas[i].get("source", "Unknown"),
                                  page=metas[i].get("page", 0))
                )
            return chunks

        def clear(self) -> None:
            try:
                self.client.delete_collection("notebook_sources")
                self.collection = self.client.create_collection(
                    name="notebook_sources", embedding_function=self.embed_func
                )
            except Exception:
                pass


def _extract_text_by_page(file_path: Path) -> list[str]:
    suffix = file_path.suffix.lower()
    if suffix in {".txt", ".md", ".csv"}:
        return [file_path.read_text(encoding="utf-8", errors="replace")]
    pages = page_count(file_path)
    text_pages = []
    for i in range(pages):
        try:
            if suffix == ".pdf":
                import pymupdf as fitz
                doc = fitz.open(file_path)
                page_text = doc[i].get_text()
                doc.close()
                if page_text.strip():
                    text_pages.append(page_text)
                    continue
        except Exception:
            pass
        try:
            img = render_page(file_path, i, zoom=1.5)
            if pytesseract:
                text_pages.append(pytesseract.image_to_string(img))
            else:
                text_pages.append("")
        except Exception:
            text_pages.append("")
    return text_pages


def _chunk_text(text: str, chunk_size: int, overlap: int) -> list[str]:
    words = text.split()
    chunks = []
    i = 0
    while i < len(words):
        chunks.append(" ".join(words[i : i + chunk_size]))
        i += chunk_size - overlap
    return chunks

