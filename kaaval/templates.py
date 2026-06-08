from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path

APP_DIR = Path.home() / ".kaaval"
TEMPLATE_FILE = APP_DIR / "templates.json"
SETTINGS_FILE = APP_DIR / "settings.json"


# ── Settings persistence ──────────────────────────────────────────────

def load_settings() -> dict:
    """Load user settings from disk. Returns empty dict on any failure."""
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_settings(settings: dict) -> None:
    """Merge *settings* into the on-disk file (never overwrites unrelated keys)."""
    current = load_settings()
    current.update(settings)
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(current, indent=2, ensure_ascii=False), encoding="utf-8")


# ── Extraction Templates ──────────────────────────────────────────────

@dataclass
class TemplateField:
    name: str
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    page: int = 0

    @property
    def has_box(self) -> bool:
        return self.width > 0 and self.height > 0


@dataclass
class ExtractionTemplate:
    name: str
    fields: list[TemplateField] = field(default_factory=list)
    template_id: str = field(default_factory=lambda: uuid.uuid4().hex)

    def field_names(self) -> list[str]:
        return [f.name for f in self.fields]


class TemplateStore:
    def __init__(self, path: Path = TEMPLATE_FILE) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.templates: list[ExtractionTemplate] = []
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            self.templates = []
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.templates = []
            return
        self.templates = []
        for item in payload.get("templates", []):
            fields = [TemplateField(**f) for f in item.get("fields", [])]
            self.templates.append(
                ExtractionTemplate(
                    name=item.get("name", "Untitled"),
                    fields=fields,
                    template_id=item.get("template_id") or uuid.uuid4().hex,
                )
            )

    def save(self) -> None:
        payload = {"templates": [asdict(t) for t in self.templates]}
        self.path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def all(self) -> list[ExtractionTemplate]:
        return list(self.templates)

    def get(self, template_id: str) -> ExtractionTemplate | None:
        for t in self.templates:
            if t.template_id == template_id:
                return t
        return None

    def upsert(self, template: ExtractionTemplate) -> None:
        for i, existing in enumerate(self.templates):
            if existing.template_id == template.template_id:
                self.templates[i] = template
                self.save()
                return
        self.templates.append(template)
        self.save()

    def delete(self, template_id: str) -> None:
        self.templates = [t for t in self.templates if t.template_id != template_id]
        self.save()


# ── Few-Shot Training Templates ───────────────────────────────────────

FEWSHOT_FILE = APP_DIR / "fewshot.json"

@dataclass
class FewShotExample:
    image_path: str = ""
    output_json: str = ""

@dataclass
class FewShotTemplate:
    name: str
    examples: list[FewShotExample] = field(default_factory=list)
    template_id: str = field(default_factory=lambda: uuid.uuid4().hex)

class FewShotStore:
    def __init__(self, path: Path = FEWSHOT_FILE) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.templates: list[FewShotTemplate] = []
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            self.templates = []
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.templates = []
            return
        self.templates = []
        for item in payload.get("templates", []):
            examples = [FewShotExample(**f) for f in item.get("examples", [])]
            self.templates.append(
                FewShotTemplate(
                    name=item.get("name", "Untitled"),
                    examples=examples,
                    template_id=item.get("template_id") or uuid.uuid4().hex,
                )
            )

    def save(self) -> None:
        payload = {"templates": [asdict(t) for t in self.templates]}
        self.path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def all(self) -> list[FewShotTemplate]:
        return list(self.templates)

    def get(self, template_id: str) -> FewShotTemplate | None:
        for t in self.templates:
            if t.template_id == template_id:
                return t
        return None

    def upsert(self, template: FewShotTemplate) -> None:
        for i, existing in enumerate(self.templates):
            if existing.template_id == template.template_id:
                self.templates[i] = template
                self.save()
                return
        self.templates.append(template)
        self.save()

    def delete(self, template_id: str) -> None:
        self.templates = [t for t in self.templates if t.template_id != template_id]
        self.save()
