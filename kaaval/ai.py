from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any

from PIL import Image

from bamini import convert_rows_to_bamini, unicode_to_bamini
from document import image_to_base64_png, prepare_for_vision
from templates import ExtractionTemplate

DEFAULT_GEMINI_MODEL = "gemini-3.5-flash"

try:
    import pytesseract  # type: ignore
    import os
    _tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(_tess_path):
        pytesseract.pytesseract.tesseract_cmd = _tess_path
except ImportError:
    pass


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
        normalized = _normalize_key(model)
        return any(model == m or normalized in _normalize_key(m) for m in models)

    def is_vision_model(self, model: str | None = None) -> bool:
        """Heuristic to check if a model name suggests vision capabilities."""
        m = (model or self.vision_model).lower()
        return any(keyword in m for keyword in ["vl", "vision", "llava", "pixtral", "moondream", "minicpm"])

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
        result = self.generate(prompt, model=self.vision_model, images=[crop], system=system)
        return unicode_to_bamini(result)

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
        rows = parse_json_rows(response, template.field_names())
        return convert_rows_to_bamini(rows)

    def chat(self, prompt: str) -> str:
        system = "You are a local AI assistant. Answer clearly in Tamil or English based on the user's language."
        return self.generate(prompt, model=self.chat_model, system=system)

    def extract_flexible(self, headers: list[str], *, images: list[Image.Image] | None = None, text: str | None = None, instruction: str | None = None, fewshot_template=None) -> list[dict[str, str]]:
        """Extract data into a JSON list based on headers, using either vision (images) or OCR text."""
        headers_json = json.dumps({h: "" for h in headers}, ensure_ascii=False)
        default_inst = (
            "Extract structured data from this document. Understand the context and find values for each field. "
            "The document may contain Tamil and English. Preserve Tamil script exactly as Unicode Tamil.\n"
            "Return a JSON array of objects, where each object represents one entry/entity found in the document."
        )
        prompt = (
            f"{instruction or default_inst}\n\n"
            f"Use exactly these JSON keys: {headers_json}\n"
            "If a value is not found, use \"Not Specified\".\n"
            "Return ONLY a raw JSON array, with no markdown formatting if possible."
        )
        
        all_images = []
        if fewshot_template and hasattr(fewshot_template, 'examples') and fewshot_template.examples:
            prompt += "\n\n=== TRAINING EXAMPLES ==="
            for i, ex in enumerate(fewshot_template.examples, 1):
                if getattr(ex, 'image_path', ''):
                    try:
                        img = Image.open(ex.image_path).convert("RGB")
                        all_images.append(img)
                        prompt += f"\nExample {i} Input Document: [See attached Image {len(all_images)}]\n"
                    except Exception:
                        pass
                prompt += f"Example {i} Expected Output:\n{ex.output_json}\n"
            prompt += "\n=== END OF TRAINING EXAMPLES ===\nNow process the new input document:"

        if images:
            for img in images:
                all_images.append(img)
                prompt += f"\n[See attached Image {len(all_images)} for New Input Document]\n"
        
        if text:
            prompt += f"\n\nDocument Text:\n{text}"

        system = "You are a precise data extraction engine. Output must be valid JSON array only."
        model = self.vision_model if images else self.chat_model
        response = self.generate(prompt, model=model, images=all_images if all_images else None, json_mode=True, system=system)
        rows = parse_json_rows(response, headers)
        return convert_rows_to_bamini(rows)

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


@dataclass
class GeminiClient:
    api_key: str | None = None
    model: str = DEFAULT_GEMINI_MODEL
    language_hint: str = "Tamil and English"
    timeout: int = 300
    endpoint: str = "https://generativelanguage.googleapis.com/v1beta"

    def _api_key(self) -> str:
        key = (self.api_key or os.environ.get("GEMINI_API_KEY") or "").strip()
        if not key:
            raise RuntimeError("Gemini API key is missing. Enter it in the Gemini key field or set GEMINI_API_KEY.")
        return key

    def generate(
        self,
        prompt: str,
        *,
        images: list[Image.Image] | None = None,
        system: str | None = None,
        schema: dict[str, Any] | None = None,
    ) -> str:
        parts: list[dict[str, Any]] = [{"text": prompt}]
        for image in images or []:
            parts.append(
                {
                    "inline_data": {
                        "mime_type": "image/png",
                        "data": image_to_base64_png(prepare_for_vision(image)),
                    }
                }
            )
        payload: dict[str, Any] = {
            "contents": [{"role": "user", "parts": parts}],
            "generationConfig": {"temperature": 0},
        }
        if system:
            payload["systemInstruction"] = {"parts": [{"text": system}]}
        if schema:
            payload["generationConfig"]["response_mime_type"] = "application/json"
            payload["generationConfig"]["response_schema"] = _gemini_rest_schema(schema)
        return self._send_request(payload)

    def _send_request(self, payload: dict[str, Any]) -> str:
        url = f"{self.endpoint.rstrip('/')}/models/{self.model}:generateContent"
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "x-goog-api-key": self._api_key(),
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            try:
                parsed = json.loads(detail)
                detail = parsed.get("error", {}).get("message", detail)
            except json.JSONDecodeError:
                pass
            raise RuntimeError(f"Gemini API error: {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Could not reach Gemini API: {exc}") from exc

        texts: list[str] = []
        for candidate in data.get("candidates", []):
            content = candidate.get("content", {})
            for part in content.get("parts", []):
                if isinstance(part, dict) and part.get("text"):
                    texts.append(str(part.get("text")))
        if texts:
            return "\n".join(texts).strip()
        if data.get("promptFeedback"):
            raise RuntimeError(f"Gemini returned no text. Feedback: {data.get('promptFeedback')}")
        return ""

    def extract_flexible(
        self,
        headers: list[str],
        *,
        images: list[Image.Image] | None = None,
        text: str | None = None,
        instruction: str | None = None,
        fewshot_template=None
    ) -> list[dict[str, str]]:
        headers_json = json.dumps(headers, ensure_ascii=False)
        prompt = (
            f"{instruction or 'Extract structured data from this document.'}\n\n"
            f"Use exactly these JSON keys in every row, in this order: {headers_json}\n"
            "Extract one row per distinct petitioner/entity. Preserve Tamil script exactly as Unicode Tamil. "
            "For missing values, return an empty string. Do not invent values."
        )
        
        all_images = []
        if fewshot_template and hasattr(fewshot_template, 'examples') and fewshot_template.examples:
            prompt += "\n\n=== TRAINING EXAMPLES ==="
            for i, ex in enumerate(fewshot_template.examples, 1):
                if getattr(ex, 'image_path', ''):
                    try:
                        img = Image.open(ex.image_path).convert("RGB")
                        all_images.append(img)
                        prompt += f"\nExample {i} Input Document: [See attached Image {len(all_images)}]\n"
                    except Exception:
                        pass
                prompt += f"Example {i} Expected Output:\n{ex.output_json}\n"
            prompt += "\n=== END OF TRAINING EXAMPLES ===\nNow process the new input document:"

        if images:
            for img in images:
                all_images.append(img)
                prompt += f"\n[See attached Image {len(all_images)} for New Input Document]\n"
                
        if text:
            prompt += f"\n\nDocument Text:\n{text}"
        schema = {
            "type": "object",
            "properties": {
                "rows": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {header: {"type": "string"} for header in headers},
                        "required": headers,
                    },
                }
            },
            "required": ["rows"],
        }
        system = "You are a precise Tamil-English legal and financial document extraction engine. Output valid JSON only."
        response = self.generate(prompt, images=all_images if all_images else None, system=system, schema=schema)
        rows = parse_json_rows(response, headers)
        return convert_rows_to_bamini(rows)


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


def _gemini_rest_schema(schema: dict[str, Any]) -> dict[str, Any]:
    converted: dict[str, Any] = {}
    for key, value in schema.items():
        if key == "type" and isinstance(value, str):
            converted[key] = value.upper()
        elif key in {"properties", "$defs", "defs"} and isinstance(value, dict):
            converted[key] = {
                str(prop_name): _gemini_rest_schema(prop_schema)
                for prop_name, prop_schema in value.items()
                if isinstance(prop_schema, dict)
            }
        elif key == "items" and isinstance(value, dict):
            converted[key] = _gemini_rest_schema(value)
        elif isinstance(value, dict):
            converted[key] = _gemini_rest_schema(value)
        elif isinstance(value, list):
            converted[key] = [_gemini_rest_schema(item) if isinstance(item, dict) else item for item in value]
        else:
            converted[key] = value
    return converted


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
    return unicode_to_bamini("\n".join(lines))


def parse_json_rows(text: str, fields: list[str]) -> list[dict[str, str]]:
    if not text or not text.strip():
        return []
    payload = _extract_json_payload(text)
    if payload is None:
        return []
    data = _loads_json_loose(payload)
    if data is None:
        return []
    data = _coerce_json_rows(data)
    if data is None:
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
                matched_key = _match_json_key(field_item, key_lookup)
                value = row.get(matched_key, "") if matched_key is not None else ""
            if value is None:
                value = ""
            values[field_item] = _stringify_cell_value(value)
        if any(_is_meaningful_cell_value(value) for value in values.values()):
            normalized.append(values)
    return normalized


def _extract_json_payload(text: str) -> str | None:
    if not text:
        return None
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
    decoder = json.JSONDecoder()
    for start, char in enumerate(stripped):
        if char not in "[{":
            continue
        try:
            _data, end = decoder.raw_decode(stripped[start:])
        except json.JSONDecodeError:
            continue
        return stripped[start : start + end]
    start_array = stripped.find("[")
    end_array = stripped.rfind("]")
    if start_array >= 0 and end_array > start_array:
        return stripped[start_array : end_array + 1]
    start_obj = stripped.find("{")
    end_obj = stripped.rfind("}")
    if start_obj >= 0 and end_obj > start_obj:
        return stripped[start_obj : end_obj + 1]
    return None


def _loads_json_loose(payload: str) -> Any | None:
    candidates = [payload, _remove_trailing_json_commas(payload)]
    for candidate in candidates:
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue
    return None


def _remove_trailing_json_commas(payload: str) -> str:
    return re.sub(r",(\s*[}\]])", r"\1", payload)


def _coerce_json_rows(data: Any) -> list[Any] | None:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("rows", "data", "table", "records", "items", "entries", "documents", "petitions", "results", "result"):
            if key in data:
                rows = _coerce_json_rows(data[key])
                if rows is not None:
                    return rows
        return [data]
    return None


def _match_json_key(field: str, key_lookup: dict[str, str]) -> str | None:
    normalized = _normalize_key(field)
    if normalized in key_lookup:
        return key_lookup[normalized]
    for alias in _field_aliases(normalized):
        if alias in key_lookup:
            return key_lookup[alias]
    fuzzy_matches = [
        original
        for key, original in key_lookup.items()
        if key and normalized and (key in normalized or normalized in key or any(alias in key for alias in _field_aliases(normalized)))
    ]
    if len(fuzzy_matches) == 1:
        return fuzzy_matches[0]
    return None


def _field_aliases(normalized_field: str) -> list[str]:
    aliases: list[str] = []
    if any(marker in normalized_field for marker in ("mobile", "phone", "contact")):
        aliases.extend(["mobile", "mobileno", "phonenumber", "phone", "contact", "contactnumber"])
    if "sender" in normalized_field:
        aliases.extend(["sender", "sendername", "senderaddress", "fromaddress"])
    if any(marker in normalized_field for marker in ("petitioner", "complainant", "applicant")):
        aliases.extend(["petitioner", "petitionername", "complainant", "applicant", "nameaddress"])
    if any(marker in normalized_field for marker in ("crime", "fir")):
        aliases.extend(["crimeno", "crimenumber", "firno", "firnumber"])
    if "id" in normalized_field:
        aliases.extend(["id", "idno", "idnumber", "aadhar", "aadhaar", "pan"])
    if any(marker in normalized_field for marker in ("plan", "scheme")):
        aliases.extend(["plan", "planname", "scheme", "schemename"])
    if "agent" in normalized_field:
        aliases.append("agent")
    if "statement" in normalized_field:
        aliases.extend(["statement", "161statement"])
    return aliases


def _stringify_cell_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        parts = [_stringify_cell_value(item) for item in value]
        return "\n".join(part for part in parts if part)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value).strip()


def _is_meaningful_cell_value(value: str) -> bool:
    text = str(value).strip()
    if not text:
        return False
    return text.lower() not in {"not specified", "not found", "n/a", "na", "none", "null", "-", "--"}


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
