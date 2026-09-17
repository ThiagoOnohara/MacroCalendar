"""Persistent, non-secret configuration for MacroCalendar."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


DEFAULT_COUNTRIES = [
    "brazil",
    "united states",
    "euro zone",
    "canada",
    "united kingdom",
    "china",
    "sweden",
    "norway",
    "south korea",
    "mexico",
    "chile",
    "colombia",
    "south africa",
    "czech republic",
    "hungary",
    "poland",
    "japan",
    "switzerland",
    "france",
    "spain",
    "italy",
]

DEFAULT_CATEGORIES = ["inflation", "employment", "central_banks", "economic_activity"]


def default_config_path() -> Path:
    configured = os.environ.get("MACROCALENDAR_CONFIG")
    if configured:
        return Path(configured).expanduser()

    app_data = os.environ.get("APPDATA")
    base = Path(app_data) if app_data else Path.home() / "AppData" / "Roaming"
    return base / "MacroCalendar" / "config.json"


@dataclass
class AppConfig:
    """Settings that affect collection and Outlook publication."""

    calendar_name: Optional[str] = None
    days: int = 10
    timezone: str = "America/Sao_Paulo"
    verify_tls: bool = False
    reminder_minutes: int = 15
    duration_minutes: int = 30
    countries: List[str] = field(default_factory=lambda: list(DEFAULT_COUNTRIES))
    categories: List[str] = field(default_factory=lambda: list(DEFAULT_CATEGORIES))
    include_holidays: bool = True

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "AppConfig":
        config_path = path or default_config_path()
        if not config_path.exists():
            return cls()

        try:
            raw = json.loads(config_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(f"Não foi possível ler a configuração: {config_path}\n{exc}") from exc

        if not isinstance(raw, dict):
            raise ValueError(f"A configuração precisa ser um objeto JSON: {config_path}")

        allowed = {field_name for field_name in cls.__dataclass_fields__}
        values: Dict[str, Any] = {key: value for key, value in raw.items() if key in allowed}
        config = cls(**values)
        config.validate()
        return config

    def save(self, path: Optional[Path] = None) -> Path:
        config_path = path or default_config_path()
        config_path.parent.mkdir(parents=True, exist_ok=True)
        config_path.write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return config_path

    def validate(self) -> None:
        if not isinstance(self.verify_tls, bool):
            raise ValueError("verify_tls deve ser true ou false")
        if self.days < 0:
            raise ValueError("days deve ser maior ou igual a zero")
        if self.reminder_minutes < 0:
            raise ValueError("reminder_minutes deve ser maior ou igual a zero")
        if self.duration_minutes <= 0:
            raise ValueError("duration_minutes deve ser maior que zero")
        if not self.countries:
            raise ValueError("countries não pode ser vazio")
        if not self.categories:
            raise ValueError("categories não pode ser vazio")
