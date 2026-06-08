from __future__ import annotations

import json
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any

from PIL import Image

from document import image_to_base64_png, prepare_for_vision
from templates import ExtractionTemplate


@dataclass
class OllamaClient:
    host: str = "http://localhost:11434"
    vision_model: str = "qwen2.5vl:7b"
    chat_model: str = "gemma3:12b"
    embed_model: str = "nomic-embed-text"
    language_hint: str = "Tamil and English"
    timeout: int = 180
    max_retries: int = 1
    _connection_cache: dict[str, Any] = field(default_factory=dict, repr=False)

    def check_connection(self) -> tuple[bool, str]:
        request = urllib.request.Request(f"{self.host.rstrip('/')}/api/tags", method="GET")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                data = json.loads(response.read().decode("utf-8"))
                models = [m.get("name", "") for m in data.get("models", [])]
                self._connection_cache["models"] = models
                self._connection_cache["connected"] = True
                return True, ""
        except urllib.error.URLError as exc:
            self._connection_cache["connected"] = False
            return False, f"Could not reach Ollama at {self.host}: {exc}"

    def list_models(self) -> list[str]:
        """Return list of available model names from Ollama."""
        if "models" in self._connection_cache:
            return self._connection_cache["models"]
        request = urllib.request.Request(f"{self.host.rstrip('/')}/api/tags", method="GET")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                data = json.loads(response.read().decode("utf-8"))
                models = [m.get("name", "") for m in data.get("models", [])]
                self._connection_cache["models"] = models
                return models
        except Exception:
            return []

    def is_model_available(self, model: str) -> bool:
        """Check if a specific model is available in Ollama."""
        models = self.list_models()
        return any(model in m for m in models)

    def generate(
        self,
        prompt: str,
        *,
        model: str | None = None,
        images: list[Image.Image] | None = None,
        json_mode: bool = False,
        system: str | None = None,
    ) -> str:
        user_message: dict[str, Any] = {"role": "user", "content": prompt}
        if images:
            user_message["images"] = [image_to_base64_png(prepare_for_vision(image)) for image in images]
        messages: list[dict[str, Any]] = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append(user_message)
        payload: dict[str, Any] = {
            "model": model or self.chat_model,
            "messages": messages,
            "stream": False,
            "options": {"temperature": 0},
        }
        if json_mode:
            payload["format"] = "json"

        last_error: Exception | None = None
        for attempt in range(self.max_retries + 1):
            try:
                return self._send_request(payload)
            except RuntimeError as exc:
                last_error = exc
                error_str = str(exc).lower()
                # Don't retry on model-not-found or invalid request errors
                if "not found" in error_str or "invalid" in error_str:
                    raise
                if attempt < self.max_retries:
                    wait = 2 ** attempt
                    time.sleep(wait)
                    continue
                raise
        raise last_error or RuntimeError("Unexpected retry exhaustion")

    def _send_request(self, payload: dict[str, Any]) -> str:
        request = urllib.request.Request(
            f"{self.host.rstrip('/')}/api/chat",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            try:
                detail = json.loads(detail).get("error", detail)
            except json.JSONDecodeError:
                pass
            hint = _error_hint(detail, exc.code)
            raise RuntimeError(f"Ollama error: {detail}\n💡 {hint}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(
                f"Could not reach Ollama at {self.host}: {exc}\n"
                "💡 Make sure Ollama is running: `ollama serve`"
            ) from exc
        message = data.get("message", {})
        if isinstance(message, dict):
            return str(message.get("content", "")).strip()
        return str(data.get("response", "")).strip()

    def extract_crop(self, field_name: str, crop: Image.Image) -> str:
        prompt = (
            "Extract only the value for this field from the cropped document image.\n"
            f"Field: {field_name}\n"
            f"Expected languages: {self.language_hint}.\n"
            "Tamil OCR is important: preserve Tamil script exactly as Unicode Tamil. Do not translate Tamil names, "
            "addresses, scheme names, or table text into English. Return only the value, with no label. "
            "If the crop is blank or unreadable, return an empty response."
        )
        system = "You are a careful OCR and document-table extraction model for Tamil and English legal/finance documents."
        return self.generate(prompt, model=self.vision_model, images=[crop], system=system)

    def extract_table_rows(self, template: ExtractionTemplate, image: Image.Image) -> list[dict[str, str]]:
        fields_json = json.dumps(template.field_names(), ensure_ascii=False)
        prompt = (
            "Extract data only from the visible table in this document image. Ignore all paragraphs, signatures, seals, "
            "headings, and notes outside the table.\n"
            f"Expected languages: {self.language_hint}. Tamil recognition is important. Preserve Tamil script exactly "
            "as Unicode Tamil; do not translate or transliterate names, addresses, scheme names, or table text.\n"
            "Return one JSON object for each actual table data row. Do not return the header row. "
            "If a person/name cell spans multiple rows, repeat that person/name value in each related row. "
            "Preserve Tamil and English text exactly as much as possible. Preserve multi-line cell content using \\n.\n"
            f"Use exactly these JSON keys, in this order: {fields_json}\n"
            'For missing values, use an empty string. Return only a JSON object shaped like {"rows":[...]}, '
            "with no markdown and no explanation."
        )
        system = "You are a precise Tamil-English OCR engine. Your output must be valid JSON only."
        response = self.generate(prompt, model=self.vision_model, images=[image], json_mode=True, system=system)
        return parse_json_rows(response, template.field_names())

    def chat(self, prompt: str) -> str:
        system = "You are a local AI assistant. Answer clearly in Tamil or English based on the user's language."
        return self.generate(prompt, model=self.chat_model, system=system)

    def embed(self, texts: list[str]) -> list[list[float]]:
        payload: dict[str, Any] = {
            "model": self.embed_model,
            "input": texts,
        }
        request = urllib.request.Request(
            f"{self.host.rstrip('/')}/api/embed",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
                return data.get("embeddings", [])
        except Exception as exc:
            raise RuntimeError(f"Ollama embedding failed: {exc}") from exc


def _error_hint(detail: str, code: int) -> str:
    """Provide actionable hints based on common Ollama errors."""
    lower = detail.lower()
    if "not found" in lower or code == 404:
        return "The model may not be downloaded yet. Run: `ollama pull <model-name>`"
    if "out of memory" in lower or "oom" in lower:
        return "The model is too large for your GPU. Try a smaller model or increase VRAM."
    if "missing data required for image input" in lower or "does not support images" in lower:
        return "The selected model does not support vision (images). Try switching to a vision model like `qwen2.5vl` or `llama3.2-vision`."
    if code == 500:
        return "Internal Ollama error. Try restarting Ollama: `ollama serve` or check if the model supports vision inputs."
    return "Check Ollama logs for more details."


def ocr_crop(crop: Image.Image) -> str:
    try:
        import pytesseract  # type: ignore
    except ImportError:
        return ""
    try:
        text = pytesseract.image_to_string(crop, lang="eng+tam", config="--psm 6")
    except Exception:
        try:
            text = pytesseract.image_to_string(crop, config="--psm 6")
        except Exception:
            return ""
    lines = [" ".join(line.split()) for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


def parse_json_rows(text: str, fields: list[str]) -> list[dict[str, str]]:
    payload = _extract_json_payload(text)
    if payload is None:
        return []
    try:
        data = json.loads(payload)
    except json.JSONDecodeError:
        return []
    if isinstance(data, dict):
        for key in ("rows", "data", "table"):
            if isinstance(data.get(key), list):
                data = data[key]
                break
        else:
            data = [data]
    if not isinstance(data, list):
        return []

    normalized: list[dict[str, str]] = []
    for row in data:
        if not isinstance(row, dict):
            continue
        key_lookup = {_normalize_key(str(key)): key for key in row.keys()}
        values: dict[str, str] = {}
        for field_item in fields:
            value = row.get(field_item, "")
            if value == "":
                matched_key = key_lookup.get(_normalize_key(field_item))
                value = row.get(matched_key, "") if matched_key is not None else ""
            if value is None:
                value = ""
            values[field_item] = str(value).strip()
        if any(values.values()):
            normalized.append(values)
    return normalized


def _extract_json_payload(text: str) -> str | None:
    stripped = text.strip()
    if not stripped:
        return None
    if stripped.startswith("```"):
        lines = stripped.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        stripped = "\n".join(lines).strip()
    if stripped.startswith("[") and stripped.endswith("]"):
        return stripped
    if stripped.startswith("{") and stripped.endswith("}"):
        return stripped
    start_array = stripped.find("[")
    end_array = stripped.rfind("]")
    if start_array >= 0 and end_array > start_array:
        return stripped[start_array : end_array + 1]
    start_obj = stripped.find("{")
    end_obj = stripped.rfind("}")
    if start_obj >= 0 and end_obj > start_obj:
        return stripped[start_obj : end_obj + 1]
    return None


def _normalize_key(value: str) -> str:
    return "".join(char.lower() for char in value if char.isalnum())


def tesseract_status() -> tuple[bool, str]:
    try:
        import pytesseract  # type: ignore
    except ImportError:
        return False, "Python package pytesseract is not installed."
    try:
        pytesseract.get_tesseract_version()
    except Exception as exc:
        return False, f"Tesseract OCR is not available on PATH: {exc}"
    return True, ""


def build_chat_prompt(table_context: str, history: list[tuple[str, str]], question: str) -> str:
    recent_history = history[-8:]
    lines = [
        "You are a local AI assistant for a spreadsheet created from document extraction.",
        "Answer using the table context when relevant. You can answer in Tamil or English depending on the user's language.",
        "",
        "Table context:",
        table_context or "(no rows yet)",
        "",
        "Recent conversation:",
    ]
    for role, content in recent_history:
        lines.append(f"{role}: {content}")
    lines.extend(["", f"User: {question}", "Assistant:"])
    return "\n".join(lines)


def extraction_summary(template: ExtractionTemplate, values: dict[str, str]) -> str:
    parts = [f"{name}: {values.get(name, '')}" for name in template.field_names()]
    return "; ".join(parts)
