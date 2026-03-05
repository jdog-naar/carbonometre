from __future__ import annotations

import uuid
from datetime import datetime, timezone
from typing import Any

import pandas as pd

from carbonometre.constants import DEFAULT_FACTORS, DEFAULT_FACTOR_REFERENCES, POSTES, TEAM_OPTIONS


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _base_entry(dossier_id: str, dossier_type: str, is_anonymous: bool, team_code: str, person_label: str) -> dict[str, Any]:
    if team_code and team_code not in TEAM_OPTIONS and team_code != "PLATEFORMES":
        raise ValueError("Equipe invalide")
    return {
        "record_id": str(uuid.uuid4()),
        "dossier_id": dossier_id,
        "dossier_type": dossier_type,
        "is_anonymous": bool(is_anonymous),
        "team_code": team_code,
        "team_label": TEAM_OPTIONS.get(team_code, ""),
        "person_label": "" if is_anonymous else person_label.strip(),
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }


def add_achat(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    item_label: str,
    amount_eur: float,
    factor_kgco2e_per_eur: float,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    emissions = float(amount_eur) * float(factor_kgco2e_per_eur)
    row.update(
        {
            "poste": "achats",
            "item_label": item_label,
            "amount_eur": float(amount_eur),
            "factor_kgco2e_per_eur": float(factor_kgco2e_per_eur),
            "emissions_kgco2e": emissions,
        }
    )
    return row


def add_domicile(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    transport_mode: str,
    distance_one_way_km: float,
    days_per_year: float,
    round_trip: bool,
    factor_kgco2e_per_km: float | None,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    factor = (
        float(factor_kgco2e_per_km)
        if factor_kgco2e_per_km is not None
        else DEFAULT_FACTORS["domicile_mode_factors"].get(transport_mode, DEFAULT_FACTORS["domicile_mode_factors"]["other"])
    )
    annual_km = float(distance_one_way_km) * (2 if round_trip else 1) * float(days_per_year)
    emissions = annual_km * factor
    row.update(
        {
            "poste": "domicile_travail",
            "transport_mode": transport_mode,
            "distance_one_way_km": float(distance_one_way_km),
            "days_per_year": float(days_per_year),
            "round_trip": bool(round_trip),
            "factor_kgco2e_per_km": factor,
            "annual_km": annual_km,
            "emissions_kgco2e": emissions,
        }
    )
    return row


def add_campagne(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    campaign_label: str,
    segment_mode: str,
    distance_km: float,
    passengers_count: float,
    round_trip: bool,
    factor_kgco2e_per_km_per_person: float | None,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    factor = (
        float(factor_kgco2e_per_km_per_person)
        if factor_kgco2e_per_km_per_person is not None
        else DEFAULT_FACTORS["campagnes_mode_factors"].get(segment_mode, DEFAULT_FACTORS["campagnes_mode_factors"]["other"])
    )
    effective_km = float(distance_km) * (2 if round_trip else 1)
    emissions = effective_km * float(passengers_count) * factor
    row.update(
        {
            "poste": "campagnes_terrain",
            "campaign_label": campaign_label,
            "segment_mode": segment_mode,
            "distance_km": float(distance_km),
            "passengers_count": float(passengers_count),
            "round_trip": bool(round_trip),
            "factor_kgco2e_per_km_per_person": factor,
            "effective_km": effective_km,
            "emissions_kgco2e": emissions,
        }
    )
    return row


def add_mission(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    mission_id: str,
    mission_mode: str,
    distance_one_way_km: float,
    round_trip: bool,
    factor_kgco2e_per_km: float | None,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    factor = (
        float(factor_kgco2e_per_km)
        if factor_kgco2e_per_km is not None
        else DEFAULT_FACTORS["missions_mode_factors"].get(mission_mode, DEFAULT_FACTORS["missions_mode_factors"]["plane"])
    )
    distance_total_km = float(distance_one_way_km) * (2 if round_trip else 1)
    emissions = distance_total_km * factor
    uncertainty = emissions * 0.3
    row.update(
        {
            "poste": "missions",
            "mission_id": mission_id,
            "mission_mode": mission_mode,
            "distance_one_way_km": float(distance_one_way_km),
            "distance_total_km": distance_total_km,
            "round_trip": bool(round_trip),
            "factor_kgco2e_per_km": factor,
            "emissions_kgco2e": emissions,
            "uncertainty_kgco2e": uncertainty,
        }
    )
    return row


def add_heures_calcul(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    compute_label: str,
    compute_type: str,
    hours: float,
    power_kw: float,
    kwh: float,
    factor_kgco2e_per_kwh: float,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    energy_kwh = float(kwh) if float(kwh) > 0 else float(hours) * float(power_kw)
    emissions = energy_kwh * float(factor_kgco2e_per_kwh)
    row.update(
        {
            "poste": "heures_calcul",
            "compute_label": compute_label,
            "compute_type": compute_type,
            "hours": float(hours),
            "power_kw": float(power_kw),
            "kwh": float(energy_kwh),
            "factor_kgco2e_per_kwh": float(factor_kgco2e_per_kwh),
            "emissions_kgco2e": emissions,
        }
    )
    return row


def add_plateforme(
    dossier_id: str,
    dossier_type: str,
    is_anonymous: bool,
    team_code: str,
    person_label: str,
    platform_name: str,
    user_role: str,
    usage_hours: float,
    usage_dates_label: str,
    usage_description: str,
    material_type: str,
    material_purchase_eur: float,
    maintenance_costs_eur: float,
    invoice_eur: float,
    usage_factor_kgco2e_per_hour: float,
    material_factor_kgco2e_per_eur: float,
    maintenance_factor_kgco2e_per_eur: float,
    invoice_factor_kgco2e_per_eur: float,
) -> dict[str, Any]:
    row = _base_entry(dossier_id, dossier_type, is_anonymous, team_code, person_label)
    role = str(user_role).strip().lower()
    is_manager = role == "responsable"
    usage_emissions = float(usage_hours) * float(usage_factor_kgco2e_per_hour)
    material_emissions = float(material_purchase_eur) * float(material_factor_kgco2e_per_eur) if is_manager else 0.0
    maintenance_emissions = float(maintenance_costs_eur) * float(maintenance_factor_kgco2e_per_eur) if is_manager else 0.0
    invoice_emissions = float(invoice_eur) * float(invoice_factor_kgco2e_per_eur) if is_manager else 0.0
    emissions = usage_emissions + material_emissions + maintenance_emissions + invoice_emissions
    row.update(
        {
            "poste": "plateforme",
            "platform_name": str(platform_name).strip(),
            "platform_user_role": role,
            "usage_hours": float(usage_hours),
            "usage_dates_label": str(usage_dates_label).strip(),
            "usage_description": str(usage_description).strip(),
            "factor_usage_kgco2e_per_hour": float(usage_factor_kgco2e_per_hour),
            "usage_emissions_kgco2e": usage_emissions,
            "material_type": str(material_type).strip(),
            "material_purchase_eur": float(material_purchase_eur if is_manager else 0.0),
            "maintenance_costs_eur": float(maintenance_costs_eur if is_manager else 0.0),
            "invoice_eur": float(invoice_eur if is_manager else 0.0),
            "factor_material_kgco2e_per_eur": float(material_factor_kgco2e_per_eur),
            "factor_maintenance_kgco2e_per_eur": float(maintenance_factor_kgco2e_per_eur),
            "factor_invoice_kgco2e_per_eur": float(invoice_factor_kgco2e_per_eur),
            "material_emissions_kgco2e": material_emissions,
            "maintenance_emissions_kgco2e": maintenance_emissions,
            "invoice_emissions_kgco2e": invoice_emissions,
            "emissions_kgco2e": emissions,
        }
    )
    return row


def entries_to_df(entries: list[dict[str, Any]]) -> pd.DataFrame:
    if not entries:
        return pd.DataFrame(columns=["record_id", "dossier_id", "dossier_type", "poste", "emissions_kgco2e"])
    df = pd.DataFrame(entries)
    if "emissions_kgco2e" in df.columns:
        df["emissions_kgco2e"] = pd.to_numeric(df["emissions_kgco2e"], errors="coerce").fillna(0.0)
    return df


def build_synthese(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["scope_type", "scope_value", "poste", "emissions_kgco2e", "forms_count", "entries_count", "updated_at"])

    now = now_iso()
    rows: list[dict[str, Any]] = []

    def _append(scope_type: str, scope_value: str, poste: str, part: pd.DataFrame) -> None:
        rows.append(
            {
                "scope_type": scope_type,
                "scope_value": scope_value,
                "poste": poste,
                "emissions_kgco2e": float(part["emissions_kgco2e"].sum()),
                "forms_count": int(part["dossier_id"].nunique()),
                "entries_count": int(len(part)),
                "updated_at": now,
            }
        )

    _append("global", "ALL", "ALL", df)
    for poste in POSTES:
        p = df[df["poste"] == poste]
        if not p.empty:
            _append("global", "ALL", poste, p)

    for team_code, part in df.groupby("team_code", dropna=False):
        team_value = team_code if isinstance(team_code, str) and team_code else "UNSPECIFIED"
        _append("team", team_value, "ALL", part)
        for poste in POSTES:
            pp = part[part["poste"] == poste]
            if not pp.empty:
                _append("team", team_value, poste, pp)

    return pd.DataFrame(rows)


def build_factors_df() -> pd.DataFrame:
    def _factor_reference(group: str, key: str = "") -> str:
        ref_group = DEFAULT_FACTOR_REFERENCES.get(group, "")
        if isinstance(ref_group, dict):
            return str(ref_group.get(key, "")).strip()
        return str(ref_group).strip()

    rows = [
        {
            "factor_id": "achats_default",
            "factor_domain": "achats",
            "factor_name": "achats_kgco2e_per_eur",
            "factor_value": DEFAULT_FACTORS["achats_kgco2e_per_eur"],
            "factor_reference": _factor_reference("achats_kgco2e_per_eur"),
            "factor_unit": "kgCO2e/EUR",
            "source": "CEREGE default v1",
            "source_version": "1.0",
            "valid_from": "",
            "valid_to": "",
            "is_default": True,
        },
        {
            "factor_id": "heures_default",
            "factor_domain": "heures_calcul",
            "factor_name": "heures_calcul_kgco2e_per_kwh",
            "factor_value": DEFAULT_FACTORS["heures_calcul_kgco2e_per_kwh"],
            "factor_reference": _factor_reference("heures_calcul_kgco2e_per_kwh"),
            "factor_unit": "kgCO2e/kWh",
            "source": "CEREGE default v1",
            "source_version": "1.0",
            "valid_from": "",
            "valid_to": "",
            "is_default": True,
        },
    ]
    for category, val in DEFAULT_FACTORS["achats_category_factors"].items():
        rows.append(
            {
                "factor_id": f"achats_{category.lower().replace(' ', '_')}",
                "factor_domain": "achats",
                "factor_name": category,
                "factor_value": val,
                "factor_reference": _factor_reference("achats_category_factors", category),
                "factor_unit": "kgCO2e/EUR",
                "source": "CEREGE default v1",
                "source_version": "1.0",
                "valid_from": "",
                "valid_to": "",
                "is_default": True,
            }
        )
    for mode, val in DEFAULT_FACTORS["domicile_mode_factors"].items():
        rows.append(
            {
                "factor_id": f"dom_{mode}",
                "factor_domain": "domicile_travail",
                "factor_name": mode,
                "factor_value": val,
                "factor_reference": _factor_reference("domicile_mode_factors", mode),
                "factor_unit": "kgCO2e/km",
                "source": "CEREGE default v1",
                "source_version": "1.0",
                "valid_from": "",
                "valid_to": "",
                "is_default": True,
            }
        )
    for mode, val in DEFAULT_FACTORS["campagnes_mode_factors"].items():
        rows.append(
            {
                "factor_id": f"camp_{mode}",
                "factor_domain": "campagnes_terrain",
                "factor_name": mode,
                "factor_value": val,
                "factor_reference": _factor_reference("campagnes_mode_factors", mode),
                "factor_unit": "kgCO2e/km/person",
                "source": "CEREGE default v1",
                "source_version": "1.0",
                "valid_from": "",
                "valid_to": "",
                "is_default": True,
            }
        )
    for mode, val in DEFAULT_FACTORS["missions_mode_factors"].items():
        rows.append(
            {
                "factor_id": f"mission_{mode}",
                "factor_domain": "missions",
                "factor_name": mode,
                "factor_value": val,
                "factor_reference": _factor_reference("missions_mode_factors", mode),
                "factor_unit": "kgCO2e/km",
                "source": "CEREGE default v1",
                "source_version": "1.0",
                "valid_from": "",
                "valid_to": "",
                "is_default": True,
            }
        )
    for material, val in DEFAULT_FACTORS["plateforme_material_factors"].items():
        rows.append(
            {
                "factor_id": f"plateforme_material_{material.lower().replace(' ', '_')}",
                "factor_domain": "plateforme",
                "factor_name": material,
                "factor_value": val,
                "factor_reference": _factor_reference("plateforme_material_factors", material),
                "factor_unit": "kgCO2e/EUR",
                "source": "CEREGE default v1",
                "source_version": "1.0",
                "valid_from": "",
                "valid_to": "",
                "is_default": True,
            }
        )
    rows.append(
        {
            "factor_id": "plateforme_usage_default",
            "factor_domain": "plateforme",
            "factor_name": "usage_hours",
            "factor_value": DEFAULT_FACTORS["plateforme_usage_kgco2e_per_hour"],
            "factor_reference": _factor_reference("plateforme_usage_kgco2e_per_hour"),
            "factor_unit": "kgCO2e/h",
            "source": "CEREGE default v1",
            "source_version": "1.0",
            "valid_from": "",
            "valid_to": "",
            "is_default": True,
        }
    )
    rows.append(
        {
            "factor_id": "plateforme_maintenance_default",
            "factor_domain": "plateforme",
            "factor_name": "maintenance",
            "factor_value": DEFAULT_FACTORS["plateforme_maintenance_kgco2e_per_eur"],
            "factor_reference": _factor_reference("plateforme_maintenance_kgco2e_per_eur"),
            "factor_unit": "kgCO2e/EUR",
            "source": "CEREGE default v1",
            "source_version": "1.0",
            "valid_from": "",
            "valid_to": "",
            "is_default": True,
        }
    )
    rows.append(
        {
            "factor_id": "plateforme_invoice_default",
            "factor_domain": "plateforme",
            "factor_name": "invoice",
            "factor_value": DEFAULT_FACTORS["plateforme_invoice_kgco2e_per_eur"],
            "factor_reference": _factor_reference("plateforme_invoice_kgco2e_per_eur"),
            "factor_unit": "kgCO2e/EUR",
            "source": "CEREGE default v1",
            "source_version": "1.0",
            "valid_from": "",
            "valid_to": "",
            "is_default": True,
        }
    )

    return pd.DataFrame(rows)
