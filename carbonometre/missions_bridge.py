from __future__ import annotations

import os
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from carbonometre.constants import TEAM_OPTIONS

COUNTRY_LABEL_BY_CODE = {
    "FR": "France",
    "BE": "Belgique",
    "CH": "Suisse",
    "DE": "Allemagne",
    "AT": "Autriche",
    "ES": "Espagne",
    "IT": "Italie",
    "GB": "Royaume-Uni",
    "PT": "Portugal",
    "NL": "Pays-Bas",
    "US": "Etats-Unis",
    "CA": "Canada",
    "CN": "Chine",
    "JP": "Japon",
    "AU": "Australie",
    "IL": "Israël",
    "GR": "Grece",
    "NO": "Norvege",
    "TN": "Tunisie",
    "MA": "Maroc",
    "PL": "Pologne",
    "BR": "Bresil",
    "TW": "Taiwan",
    "HR": "Croatie",
}


def _load_moulinette_symbols():
    repo_root = Path(__file__).resolve().parents[1]
    moulinette_root = repo_root / "Moulinette_missions"
    geocache_path = moulinette_root / "Data" / "Config" / "geocache.json"

    import sys

    if str(moulinette_root) not in sys.path:
        sys.path.insert(0, str(moulinette_root))

    old_cwd = Path.cwd()
    try:
        os.chdir(moulinette_root)
        from Libs.EasyEnums import Enm
        from Libs.EmissionsCalculator import compute_emissions_df, compute_emissions_one_row, dist_calculator
        from Libs.MissionsLoader import load_data
    finally:
        os.chdir(old_cwd)

    geocache_path.parent.mkdir(parents=True, exist_ok=True)
    if not geocache_path.exists():
        geocache_path.write_text("{}", encoding="utf-8")
    dist_calculator.cache_path = str(geocache_path)

    return Enm, compute_emissions_df, compute_emissions_one_row, dist_calculator, load_data


def _sanitize_geocache_city_country(city: str, country_code: str) -> None:
    """Remove inconsistent cached entries like 'Tel-Aviv (IL)' pointing to US."""
    city_clean = str(city or "").strip()
    cc = str(country_code or "").strip().upper()
    if not city_clean or len(cc) != 2:
        return

    repo_root = Path(__file__).resolve().parents[1]
    geocache_path = repo_root / "Moulinette_missions" / "Data" / "Config" / "geocache.json"
    if not geocache_path.exists():
        return

    try:
        import json

        raw = json.loads(geocache_path.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return
        key = f"{city_clean} ({cc})"
        payload = raw.get(key)
        if isinstance(payload, dict):
            payload_cc = str(payload.get("countryCode", "")).strip().upper()
            if payload_cc and payload_cc != cc:
                del raw[key]
                geocache_path.write_text(json.dumps(raw, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    except Exception:
        # Keep mission computation working even if cache cleanup fails.
        return


def _country_for_geocoding(country_value: str) -> str:
    raw = str(country_value or "").strip()
    code = raw.upper()
    if len(code) == 2 and code in COUNTRY_LABEL_BY_CODE:
        return COUNTRY_LABEL_BY_CODE[code]
    return raw


def _distance_correction_from_moulinette(
    one_way_km: float,
    total_geodesic_km: float,
    is_round_trip: bool,
    emission_transport_detailed: str,
) -> tuple[float, float]:
    """Return (corrected_total_km, correction_factor_vs_geodesic)."""
    detail = str(emission_transport_detailed or "").strip().lower()
    if total_geodesic_km <= 0:
        return 0.0, 1.0

    if detail.startswith("train"):
        corrected_total = float(total_geodesic_km) * 1.2
    elif detail == "car":
        corrected_total = float(total_geodesic_km) * 1.3
    elif detail.startswith("plane"):
        corrected_one_way = float(one_way_km) + 95.0
        corrected_total = corrected_one_way * (2.0 if is_round_trip else 1.0)
    else:
        corrected_total = float(total_geodesic_km)

    correction_factor = corrected_total / float(total_geodesic_km) if total_geodesic_km > 0 else 1.0
    return float(corrected_total), float(correction_factor)


def compute_single_mission_with_moulinette(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    mission_id: str,
    departure_city: str,
    departure_country: str,
    arrival_city: str,
    arrival_country: str,
    transport_mode: str,
    round_trip: bool,
) -> dict[str, Any]:
    """Compute one mission directly from city/country inputs using Moulinette logic."""
    _sanitize_geocache_city_country(departure_city, departure_country)
    _sanitize_geocache_city_country(arrival_city, arrival_country)
    Enm, _, compute_emissions_one_row, dist_calculator, _, = _load_moulinette_symbols()

    transport_map = {
        "plane": Enm.MAIN_TRANSPORT_PLANE,
        "train": Enm.MAIN_TRANSPORT_TRAIN,
        "car": Enm.MAIN_TRANSPORT_CAR,
    }

    row = pd.Series(
        {
            Enm.COL_MISSION_ID: mission_id.strip(),
            Enm.COL_DEPARTURE_CITY: departure_city.strip(),
            Enm.COL_DEPARTURE_COUNTRY: _country_for_geocoding(departure_country),
            Enm.COL_ARRIVAL_CITY: arrival_city.strip(),
            Enm.COL_ARRIVAL_COUNTRY: _country_for_geocoding(arrival_country),
            Enm.COL_ROUND_TRIP: Enm.ROUNDTRIP_YES if round_trip else Enm.ROUNDTRIP_NO,
            Enm.COL_MAIN_TRANSPORT: transport_map.get(transport_mode, Enm.MAIN_TRANSPORT_PLANE),
            Enm.COL_TRANSPORT_TYPE: transport_mode,
        }
    )

    out = compute_emissions_one_row(row)
    if out is None:
        raise ValueError("Impossible de calculer cette mission (lieu non reconnu).")
    dist_calculator.save_cache(force=True)

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    one_way_dist_km = float(out.get("one_way_dist_km", 0) or 0)
    total_geodesic_km = float(out.get("final_dist_km", 0) or 0)
    emission_transport = str(out.get("transport_for_emissions_detailed", "") or "")
    corrected_total_km, correction_factor = _distance_correction_from_moulinette(
        one_way_km=one_way_dist_km,
        total_geodesic_km=total_geodesic_km,
        is_round_trip=bool(round_trip),
        emission_transport_detailed=emission_transport,
    )
    emissions_kgco2e = float(out.get("co2e_emissions_kg", 0) or 0)
    effective_factor = emissions_kgco2e / corrected_total_km if corrected_total_km > 0 else 0.0

    return {
        "record_id": str(uuid.uuid4()),
        "dossier_id": dossier_id,
        "dossier_type": dossier_type,
        "is_anonymous": bool(is_anonymous),
        "team_code": team_code,
        "team_label": TEAM_OPTIONS.get(team_code, ""),
        "person_label": "" if is_anonymous else person_label.strip(),
        "created_at": now,
        "updated_at": now,
        "poste": "missions",
        "mission_id": mission_id.strip(),
        "departure_city": departure_city.strip(),
        "departure_country": departure_country.strip(),
        "arrival_city": arrival_city.strip(),
        "arrival_country": arrival_country.strip(),
        "departure_lat": float(out.get("departure_loc").latitude) if out.get("departure_loc") is not None else None,
        "departure_lon": float(out.get("departure_loc").longitude) if out.get("departure_loc") is not None else None,
        "arrival_lat": float(out.get("arrival_loc").latitude) if out.get("arrival_loc") is not None else None,
        "arrival_lon": float(out.get("arrival_loc").longitude) if out.get("arrival_loc") is not None else None,
        "t_type": transport_mode,
        "round_trip": bool(round_trip),
        "distance_one_way_km": one_way_dist_km,
        "distance_total_km": total_geodesic_km,
        "distance_corrected_total_km": corrected_total_km,
        "distance_correction_factor": correction_factor,
        "emission_transport": emission_transport,
        "mission_effective_factor_kgco2e_per_km": float(effective_factor),
        "emissions_kgco2e": emissions_kgco2e,
        "uncertainty_kgco2e": float(out.get("emission_uncertainty", 0) or 0),
    }


def run_moulinette_from_uploads(missions_bytes: bytes, missions_filename: str, config_bytes: bytes, config_filename: str) -> tuple[pd.DataFrame, bytes]:
    """Run Moulinette_missions on uploaded files and return computed dataframe + map image bytes."""
    _, compute_emissions_df, _, _, load_data = _load_moulinette_symbols()

    with tempfile.TemporaryDirectory(prefix="carbonometre_missions_") as tmpdir:
        tmp = Path(tmpdir)
        data_file = tmp / missions_filename
        conf_file = tmp / config_filename
        map_file = tmp / "missions_map.png"

        data_file.write_bytes(missions_bytes)
        conf_file.write_bytes(config_bytes)

        df_data = load_data(data_file, conf_file)
        df_emissions = compute_emissions_df(df_data)
        try:
            from Libs.VisualMap import generate_visual_map

            generate_visual_map(df_emissions, map_file)
        except Exception:
            # Keep missions calculations available even if map dependencies are missing.
            pass

        map_bytes = map_file.read_bytes() if map_file.exists() else b""

    return df_emissions, map_bytes


def missions_df_to_entries(
    df_emissions: pd.DataFrame,
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
) -> list[dict[str, Any]]:
    """Convert Moulinette rows to Carbonometre entries."""
    entries: list[dict[str, Any]] = []
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    for _, row in df_emissions.iterrows():
        is_round_trip = str(row.get("round_trip", "")).strip().lower() == "oui"
        one_way_dist_km = float(row.get("one_way_dist_km", 0) or 0)
        total_geodesic_km = float(row.get("final_dist_km", 0) or 0)
        emission_transport = str(row.get("transport_for_emissions_detailed", "") or "")
        corrected_total_km, correction_factor = _distance_correction_from_moulinette(
            one_way_km=one_way_dist_km,
            total_geodesic_km=total_geodesic_km,
            is_round_trip=is_round_trip,
            emission_transport_detailed=emission_transport,
        )
        emissions_kgco2e = float(row.get("co2e_emissions_kg", 0) or 0)
        effective_factor = emissions_kgco2e / corrected_total_km if corrected_total_km > 0 else 0.0
        entries.append(
            {
                "record_id": str(uuid.uuid4()),
                "dossier_id": dossier_id,
                "dossier_type": dossier_type,
                "is_anonymous": bool(is_anonymous),
                "team_code": team_code,
                "team_label": TEAM_OPTIONS.get(team_code, ""),
                "person_label": "" if is_anonymous else person_label.strip(),
                "created_at": now,
                "updated_at": now,
                "poste": "missions",
                "mission_id": row.get("mission_id", ""),
                "departure_city": row.get("departure_city", ""),
                "departure_country": row.get("departure_country", ""),
                "arrival_city": row.get("arrival_city", ""),
                "arrival_country": row.get("arrival_country", ""),
                "departure_lat": float(row.get("departure_loc").latitude) if row.get("departure_loc") is not None else None,
                "departure_lon": float(row.get("departure_loc").longitude) if row.get("departure_loc") is not None else None,
                "arrival_lat": float(row.get("arrival_loc").latitude) if row.get("arrival_loc") is not None else None,
                "arrival_lon": float(row.get("arrival_loc").longitude) if row.get("arrival_loc") is not None else None,
                "t_type": row.get("t_type", ""),
                "round_trip": is_round_trip,
                "distance_one_way_km": one_way_dist_km,
                "distance_total_km": total_geodesic_km,
                "distance_corrected_total_km": corrected_total_km,
                "distance_correction_factor": correction_factor,
                "emission_transport": emission_transport,
                "mission_effective_factor_kgco2e_per_km": float(effective_factor),
                "emissions_kgco2e": emissions_kgco2e,
                "uncertainty_kgco2e": float(row.get("emission_uncertainty", 0) or 0),
            }
        )

    return entries
