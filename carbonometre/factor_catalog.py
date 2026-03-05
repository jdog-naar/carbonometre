from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pandas as pd

from carbonometre import constants

CATALOG_PATH = constants.REPO_ROOT / "labs" / "factors_catalog.xlsx"
PLATFORM_CATALOG_PATH = constants.REPO_ROOT / "labs" / "factors_plateforme.xlsx"

_COLUMNS = ["factor_group", "factor_key", "facteur_co2", "incertitude_pct", "source", "actif"]

_MAIN_SHEETS: dict[str, list[str]] = {
    "achats": ["achats_kgco2e_per_eur", "achats_category_factors"],
    "domicile_travail": ["domicile_mode_factors"],
    "campagnes_terrain": ["campagnes_mode_factors"],
    "missions": ["missions_mode_factors"],
    "heures_calcul": ["heures_calcul_kgco2e_per_kwh"],
}

_PLATFORM_SHEETS: dict[str, list[str]] = {
    "plateforme": [
        "plateforme_material_factors",
        "plateforme_usage_kgco2e_per_hour",
        "plateforme_maintenance_kgco2e_per_eur",
        "plateforme_invoice_kgco2e_per_eur",
    ]
}

_loaded_once = False
_BASE_DEFAULT_FACTORS = deepcopy(constants.DEFAULT_FACTORS)
_BASE_DEFAULT_FACTOR_REFERENCES = deepcopy(constants.DEFAULT_FACTOR_REFERENCES)
_BASE_DEFAULT_FACTOR_UNCERTAINTY_PCT = deepcopy(constants.DEFAULT_FACTOR_UNCERTAINTY_PCT)


def _ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _is_active(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    s = str(value or "").strip().lower()
    if s in {"", "nan", "none", "null"}:
        return True
    return s not in {"false", "0", "no", "non", "off"}


def _factor_rows_for_group(group: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    values = constants.DEFAULT_FACTORS.get(group)
    refs = constants.DEFAULT_FACTOR_REFERENCES.get(group, "")
    uncs = constants.DEFAULT_FACTOR_UNCERTAINTY_PCT.get(group, 10.0)
    if isinstance(values, dict):
        for key, value in values.items():
            rows.append(
                {
                    "factor_group": group,
                    "factor_key": str(key),
                    "facteur_co2": float(value),
                    "incertitude_pct": float(uncs.get(key, 10.0)) if isinstance(uncs, dict) else float(uncs),
                    "source": str(refs.get(key, "")) if isinstance(refs, dict) else str(refs or ""),
                    "actif": True,
                }
            )
        return rows
    rows.append(
        {
            "factor_group": group,
            "factor_key": "__default__",
            "facteur_co2": float(values if values is not None else 0.0),
            "incertitude_pct": float(uncs if uncs is not None else 10.0),
            "source": str(refs or ""),
            "actif": True,
        }
    )
    return rows


def _write_default_catalog(path: Path, sheet_map: dict[str, list[str]]) -> None:
    _ensure_parent_dir(path)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet, groups in sheet_map.items():
            rows: list[dict[str, Any]] = []
            for group in groups:
                rows.extend(_factor_rows_for_group(group))
            pd.DataFrame(rows, columns=_COLUMNS).to_excel(writer, sheet_name=sheet, index=False)


def _ensure_catalog_files_exist() -> None:
    if not CATALOG_PATH.exists():
        _write_default_catalog(CATALOG_PATH, _MAIN_SHEETS)
    if not PLATFORM_CATALOG_PATH.exists():
        _write_default_catalog(PLATFORM_CATALOG_PATH, _PLATFORM_SHEETS)


def _set_factor(group: str, key: str, value: float, uncertainty_pct: float, source: str) -> None:
    target = constants.DEFAULT_FACTORS.get(group)
    if isinstance(target, dict):
        target[key] = float(value)
    else:
        constants.DEFAULT_FACTORS[group] = float(value)

    target_unc = constants.DEFAULT_FACTOR_UNCERTAINTY_PCT.get(group)
    if isinstance(target_unc, dict):
        target_unc[key] = float(uncertainty_pct)
    else:
        constants.DEFAULT_FACTOR_UNCERTAINTY_PCT[group] = float(uncertainty_pct)

    target_ref = constants.DEFAULT_FACTOR_REFERENCES.get(group)
    if isinstance(target_ref, dict):
        target_ref[key] = str(source or "")
    else:
        constants.DEFAULT_FACTOR_REFERENCES[group] = str(source or "")


def _disable_factor(group: str, key: str) -> None:
    target = constants.DEFAULT_FACTORS.get(group)
    if isinstance(target, dict):
        target.pop(key, None)
    else:
        constants.DEFAULT_FACTORS[group] = 0.0

    target_unc = constants.DEFAULT_FACTOR_UNCERTAINTY_PCT.get(group)
    if isinstance(target_unc, dict):
        target_unc.pop(key, None)
    else:
        constants.DEFAULT_FACTOR_UNCERTAINTY_PCT[group] = 10.0

    target_ref = constants.DEFAULT_FACTOR_REFERENCES.get(group)
    if isinstance(target_ref, dict):
        target_ref.pop(key, None)
    else:
        constants.DEFAULT_FACTOR_REFERENCES[group] = ""


def _apply_catalog(path: Path, sheet_map: dict[str, list[str]]) -> None:
    if not path.exists():
        return
    try:
        workbook = pd.read_excel(path, sheet_name=None)
    except Exception:
        return

    allowed_groups = {g for groups in sheet_map.values() for g in groups}
    for sheet_name, frame in workbook.items():
        if sheet_name not in sheet_map:
            continue
        if frame is None or frame.empty:
            continue
        cols = {str(c).strip().lower(): c for c in frame.columns}
        required = {"factor_group", "factor_key", "facteur_co2"}
        if not required.issubset(set(cols.keys())):
            continue

        for _, row in frame.iterrows():
            group = str(row.get(cols["factor_group"], "")).strip()
            key = str(row.get(cols["factor_key"], "")).strip()
            if not group or group not in allowed_groups:
                continue
            if not key:
                key = "__default__"
            if not isinstance(constants.DEFAULT_FACTORS.get(group), dict):
                key = "__default__"

            if not _is_active(row.get(cols.get("actif", ""), True)):
                _disable_factor(group, key)
                continue

            try:
                factor_value = float(row.get(cols["facteur_co2"], 0.0) or 0.0)
            except Exception:
                continue

            unc_col = cols.get("incertitude_pct")
            src_col = cols.get("source")
            try:
                uncertainty_pct = float(row.get(unc_col, 10.0) if unc_col else 10.0)
            except Exception:
                uncertainty_pct = 10.0
            source_raw = row.get(src_col, "") if src_col else ""
            source = "" if pd.isna(source_raw) else str(source_raw)
            if uncertainty_pct < 0:
                uncertainty_pct = 10.0

            _set_factor(group, key, factor_value, uncertainty_pct, source)


def load_factor_catalogs_once() -> dict[str, str]:
    global _loaded_once
    if _loaded_once:
        return {"main": str(CATALOG_PATH), "plateforme": str(PLATFORM_CATALOG_PATH)}

    # Keep a clean fallback in memory for robustness before applying external catalogs.
    constants.DEFAULT_FACTORS.clear()
    constants.DEFAULT_FACTORS.update(deepcopy(_BASE_DEFAULT_FACTORS))
    constants.DEFAULT_FACTOR_REFERENCES.clear()
    constants.DEFAULT_FACTOR_REFERENCES.update(deepcopy(_BASE_DEFAULT_FACTOR_REFERENCES))
    constants.DEFAULT_FACTOR_UNCERTAINTY_PCT.clear()
    constants.DEFAULT_FACTOR_UNCERTAINTY_PCT.update(deepcopy(_BASE_DEFAULT_FACTOR_UNCERTAINTY_PCT))

    _ensure_catalog_files_exist()
    _apply_catalog(CATALOG_PATH, _MAIN_SHEETS)
    _apply_catalog(PLATFORM_CATALOG_PATH, _PLATFORM_SHEETS)
    _loaded_once = True
    return {"main": str(CATALOG_PATH), "plateforme": str(PLATFORM_CATALOG_PATH)}
