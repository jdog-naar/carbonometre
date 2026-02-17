from __future__ import annotations

import os
import re
import tomllib
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
LABS_DIR = REPO_ROOT / "labs"
DEFAULT_LAB_ID = "cerege"

DEFAULT_SCIENCE_TEAMS = {
    "CLIMAT": "Climat",
    "ENV_DURABLE": "Environnement Durable",
    "TERRE_PLANETES": "Terre et Planetes",
    "RHC": "Ressources Hydrosystemes et Carbonates",
    "SAR": "SAR",
}

DEFAULT_THEME = {
    "dark": {
        "app_bg_start": "#05070d",
        "app_bg_end": "#0b1220",
        "text": "#dbe7ff",
        "sidebar_bg_start": "#02040a",
        "sidebar_bg_end": "#0a1325",
        "metric_bg": "#0e1a31",
        "metric_border": "#1f355f",
        "button_bg": "#112a57",
        "button_border": "#294b8a",
        "button_text": "#e8f0ff",
        "button_bg_hover": "#173872",
        "button_border_hover": "#3a67b8",
        "input_bg": "#0d1a33",
        "input_text": "#e8f0ff",
        "input_border": "#2b4574",
    },
    "light": {
        "app_bg_start": "#ffffff",
        "app_bg_end": "#f7fff4",
        "text": "#102214",
        "sidebar_bg_start": "#f2ffe9",
        "sidebar_bg_end": "#e6f9d8",
        "metric_bg": "#ffffff",
        "metric_border": "#9bd27e",
        "button_bg": "#6dbf3f",
        "button_border": "#57a631",
        "button_text": "#ffffff",
        "button_bg_hover": "#58a533",
        "button_border_hover": "#4b8f2d",
        "input_bg": "#ffffff",
        "input_text": "#102214",
        "input_border": "#9bd27e",
        "file_drop_bg": "#ffffff",
        "file_drop_border": "#7fbe5d",
        "file_drop_text": "#102214",
    },
}


def _slug(value: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9_-]+", "_", value.strip().lower())
    s = re.sub(r"_+", "_", s).strip("_")
    return s or DEFAULT_LAB_ID


def _read_lab_raw(lab_id: str) -> dict:
    wanted = _slug(lab_id)
    target = LABS_DIR / f"{wanted}.toml"
    fallback = LABS_DIR / f"{DEFAULT_LAB_ID}.toml"
    path = target if target.exists() else fallback
    if not path.exists():
        return {}
    try:
        return tomllib.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _theme_value(raw: dict, mode: str, key: str) -> str:
    mode_raw = raw.get("theme", {}).get(mode, {}) if isinstance(raw.get("theme", {}), dict) else {}
    fallback = DEFAULT_THEME[mode][key]
    value = mode_raw.get(key, fallback) if isinstance(mode_raw, dict) else fallback
    return str(value)


def _load_lab_config() -> dict:
    requested_lab_id = os.getenv("LAB_ID", DEFAULT_LAB_ID)
    raw = _read_lab_raw(requested_lab_id)
    lab_raw = raw.get("lab", {}) if isinstance(raw.get("lab", {}), dict) else {}
    teams_raw = raw.get("teams", {}) if isinstance(raw.get("teams", {}), dict) else {}

    effective_lab_id = _slug(str(lab_raw.get("id", requested_lab_id)))
    if effective_lab_id == "":
        effective_lab_id = DEFAULT_LAB_ID

    teams: dict[str, str] = {}
    for code, label in teams_raw.items():
        k = str(code).strip().upper()
        if not k:
            continue
        teams[k] = str(label).strip() or k
    if not teams:
        teams = dict(DEFAULT_SCIENCE_TEAMS)

    return {
        "lab_id": effective_lab_id,
        "lab_name": str(lab_raw.get("name", effective_lab_id.upper())).strip() or effective_lab_id.upper(),
        "logo_path": str(lab_raw.get("logo_path", "")).strip(),
        "teams": teams,
        "theme": {
            "dark": {k: _theme_value(raw, "dark", k) for k in DEFAULT_THEME["dark"]},
            "light": {k: _theme_value(raw, "light", k) for k in DEFAULT_THEME["light"]},
        },
    }


LAB_CONFIG = _load_lab_config()
LAB_ID = LAB_CONFIG["lab_id"]
LAB_NAME = LAB_CONFIG["lab_name"]
LAB_LOGO_PATH = LAB_CONFIG["logo_path"]
LAB_THEME = LAB_CONFIG["theme"]

TEAM_OPTIONS = {
    "": "Non renseignee",
    "ADMIN_GESTION": "ADMIN/GESTION",
    **LAB_CONFIG["teams"],
}

POSTES = ["achats", "domicile_travail", "campagnes_terrain", "missions", "heures_calcul"]

DEFAULT_FACTORS = {
    "achats_kgco2e_per_eur": 0.30,
    "achats_category_factors": {
        "Ordinateur portable": 0.35,
        "Ordinateur fixe": 0.40,
        "Serveur": 0.55,
        "Instrument scientifique": 0.45,
        "Consommables labo": 0.25,
        "Mobilier": 0.20,
        "Autre": 0.30,
    },
    "domicile_mode_factors": {
        "car": 0.20,
        "train": 0.01,
        "bus": 0.10,
        "metro": 0.005,
        "tram": 0.005,
        "bike": 0.0,
        "walk": 0.0,
        "other": 0.15,
    },
    "campagnes_mode_factors": {
        "plane": 0.19,
        "boat": 0.25,
        "car": 0.20,
        "train": 0.01,
        "other": 0.15,
    },
    "missions_mode_factors": {
        "plane": 0.19,
        "train": 0.01,
        "car": 0.20,
    },
    "heures_calcul_kgco2e_per_kwh": 0.05,
}

SCHEMA_VERSION = "1.0"
