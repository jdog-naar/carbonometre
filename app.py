from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import json
import re
import time
import unicodedata

import altair as alt
import pandas as pd
import pydeck as pdk
import streamlit as st
import streamlit.components.v1 as components

from carbonometre.calculations import (
    add_achat,
    add_campagne,
    add_domicile,
    add_heures_calcul,
    add_plateforme,
    build_synthese,
    entries_to_df,
)
from carbonometre.constants import DEFAULT_FACTORS, DEFAULT_FACTOR_REFERENCES, DEFAULT_FACTOR_UNCERTAINTY_PCT, LAB_ID, LAB_LOGO_PATH, LAB_NAME, LAB_THEME, TEAM_OPTIONS
from carbonometre.excel_io import export_excel_bytes, import_excel_entries
from carbonometre.factor_catalog import CATALOG_PATH, PLATFORM_CATALOG_PATH, load_factor_catalogs_once
from carbonometre.missions_bridge import COUNTRY_LABEL_BY_CODE, compute_single_mission_with_moulinette

load_factor_catalogs_once()

if "ui_lang" not in st.session_state:
    st.session_state.ui_lang = "FR"
if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "darkmode"

# Keep language/theme as single source of truth from toggle state,
# regardless of where the toggle is rendered (home header or sidebar).
if "theme_toggle" in st.session_state:
    st.session_state.ui_theme = "darkmode" if bool(st.session_state.get("theme_toggle", False)) else "lightmode"
if "lang_toggle" in st.session_state:
    st.session_state.ui_lang = "EN" if bool(st.session_state.get("lang_toggle", False)) else "FR"


def tf(fr: str, en: str) -> str:
    return en if st.session_state.get("ui_lang", "FR") == "EN" else fr


root = Path(__file__).resolve().parent
configured_logo = None
if LAB_LOGO_PATH:
    cfg = Path(LAB_LOGO_PATH)
    configured_logo = cfg if cfg.is_absolute() else (root / cfg)
favicon_candidates = [
    configured_logo,
    root / "cerege_logo.png",
    root / "cerege-logo.png",
    root / "logo.png",
    root / "logo.jpg",
    root / "logo.jpeg",
    root / "logo.svg",
]
favicon_path = next((p for p in favicon_candidates if p is not None and p.exists()), None)

st.set_page_config(
    page_title=f"Carbonometre {LAB_NAME}",
    page_icon=str(favicon_path) if favicon_path else "🌍",
    layout="wide",
)
header_left, header_right = st.columns([8, 2])
with header_left:
    st.markdown(f'<h1 class="app-main-title">Carbonometre {LAB_NAME}</h1>', unsafe_allow_html=True)
    st.caption(tf("MVP local: calcul CO2e par poste, sauvegarde dossier, export/reimport Excel", "Local MVP: CO2e computation by category, form save, export/reimport Excel"))
with header_right:
    candidates = [
        configured_logo,
        root / "cerege_logo.png",
        root / "cerege-logo.png",
        root / "logo.png",
        root / "logo.jpg",
        root / "logo.jpeg",
        root / "logo.svg",
    ]
    logo_path = next((p for p in candidates if p is not None and p.exists()), None)
    if logo_path is None:
        dynamic = sorted([p for p in root.iterdir() if p.is_file() and "logo" in p.name.lower() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".svg"}])
        logo_path = dynamic[0] if dynamic else None
    if logo_path:
        st.image(str(logo_path), use_container_width=True)
    if st.session_state.get("app_stage", "home") == "home":
        row_theme = st.columns([1.2, 1.0, 1.2], vertical_alignment="center")
        row_theme[0].markdown("<div style='text-align:right;'>lightmode ☀️</div>", unsafe_allow_html=True)
        with row_theme[1]:
            dark_on = st.toggle(
                "theme",
                value=st.session_state.ui_theme == "darkmode",
                key="theme_toggle",
                label_visibility="collapsed",
            )
        row_theme[2].markdown("<div style='text-align:left;'>🌙 darkmode</div>", unsafe_allow_html=True)
        st.session_state.ui_theme = "darkmode" if dark_on else "lightmode"

        row_lang = st.columns([1.2, 1.0, 1.2], vertical_alignment="center")
        row_lang[0].markdown("<div style='text-align:right;'>FR 🇫🇷</div>", unsafe_allow_html=True)
        with row_lang[1]:
            en_on = st.toggle(
                "lang",
                value=st.session_state.ui_lang == "EN",
                key="lang_toggle",
                label_visibility="collapsed",
            )
        row_lang[2].markdown("<div style='text-align:left;'>EN 🇬🇧</div>", unsafe_allow_html=True)
        st.session_state.ui_lang = "EN" if en_on else "FR"

st.markdown(
    """
<style>
.app-main-title {
  font-size: 56px !important;
  line-height: 1.05 !important;
  margin: 0 0 0.2rem 0 !important;
  font-weight: 700 !important;
}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<style>
/* Keep a CSS baseline, JS below applies inline style robustly */
div[data-testid="stToggle"] [data-baseweb="checkbox"] > label > div {
  transition: background-color 120ms ease, border-color 120ms ease;
}
</style>
""",
    unsafe_allow_html=True,
)

components.html(
    """
<script>
(function() {
  const OFF_BG = "#173872";
  const OFF_BORDER = "#1f355f";
  const ON_BG = "#f59e0b";
  const ON_BORDER = "#b45309";

  function styleToggle(toggleRoot) {
    const input = toggleRoot.querySelector('input[type="checkbox"]');
    if (!input) return;
    const checked = !!input.checked;
    const targets = toggleRoot.querySelectorAll(
      '[data-baseweb="checkbox"] > label > div, [data-baseweb="checkbox"] div[aria-hidden="true"], input[type="checkbox"] + div'
    );
    targets.forEach((el) => {
      el.style.backgroundColor = checked ? ON_BG : OFF_BG;
      el.style.borderColor = checked ? ON_BORDER : OFF_BORDER;
    });
  }

  function restyleAll() {
    document.querySelectorAll('div[data-testid="stToggle"]').forEach(styleToggle);
  }

  restyleAll();
  const obs = new MutationObserver(restyleAll);
  obs.observe(document.body, { subtree: true, childList: true, attributes: true });
})();
</script>
""",
    height=0,
)

if "entries" not in st.session_state:
    st.session_state.entries = []
if "dossier_id" not in st.session_state:
    st.session_state.dossier_id = ""
if "missions_map_bytes" not in st.session_state:
    st.session_state.missions_map_bytes = b""
if "loaded_form_path" not in st.session_state:
    st.session_state.loaded_form_path = ""
if "loaded_owner_label" not in st.session_state:
    st.session_state.loaded_owner_label = ""
if "person_label_input" not in st.session_state:
    st.session_state.person_label_input = ""
if "poste_type_input" not in st.session_state:
    st.session_state.poste_type_input = ""
if "mode_input" not in st.session_state:
    st.session_state.mode_input = "item_unique"
if "mode_input_ui" not in st.session_state:
    st.session_state.mode_input_ui = "item_unique"
if "pending_mode_change" not in st.session_state:
    st.session_state.pending_mode_change = None
if "mode_ui_reset_to" not in st.session_state:
    st.session_state.mode_ui_reset_to = ""
if "project_name_input" not in st.session_state:
    st.session_state.project_name_input = "Mon projet"
if "platform_name_input" not in st.session_state:
    st.session_state.platform_name_input = "Microscope Confocal"
if "team_code_input" not in st.session_state:
    st.session_state.team_code_input = ""
if "year_input" not in st.session_state:
    current_year = str(datetime.now(timezone.utc).year)
    st.session_state.year_input = current_year if current_year in {"2026", "2025", "2024", "2023", "2022", "2021", "2020", "1998"} else "2026"
if "reload_year_input" not in st.session_state:
    current_year = str(datetime.now(timezone.utc).year)
    st.session_state.reload_year_input = current_year if current_year in {"2026", "2025", "2024", "2023", "2022", "2021", "2020", "1998"} else "2026"
if "confirm_overwrite_pending" not in st.session_state:
    st.session_state.confirm_overwrite_pending = False
if "pending_loaded_meta" not in st.session_state:
    st.session_state.pending_loaded_meta = None
if "pending_loaded_entries" not in st.session_state:
    st.session_state.pending_loaded_entries = None
if "pending_loaded_path" not in st.session_state:
    st.session_state.pending_loaded_path = ""
if "loaded_notice" not in st.session_state:
    st.session_state.loaded_notice = ""
if "uncertainty_notice" not in st.session_state:
    st.session_state.uncertainty_notice = ""
if "editing_row_idx" not in st.session_state:
    st.session_state.editing_row_idx = None
if "mission_steps_input" not in st.session_state:
    st.session_state.mission_steps_input = [
        {
            "departure_city": "Aix-en-Provence",
            "departure_country": "FR",
            "arrival_city": "Paris",
            "arrival_country": "FR",
            "transport_mode": "train",
            "round_trip": False,
        }
    ]
if "mission_preview_rows" not in st.session_state:
    st.session_state.mission_preview_rows = []
if "mission_preview_key" not in st.session_state:
    st.session_state.mission_preview_key = ""
if "platform_users_input" not in st.session_state:
    st.session_state.platform_users_input = [
        {
            "user_label": "",
            "user_role": "responsable",
            "usage_hours": 0.0,
            "usage_dates_label": "",
            "usage_description": "",
            "material_type": "Pompe a vide",
            "material_purchase_eur": 0.0,
            "maintenance_costs_eur": 0.0,
            "invoice_eur": 0.0,
        }
    ]
if "platform_active_form" not in st.session_state:
    st.session_state.platform_active_form = ""
if "platform_preview_row" not in st.session_state:
    st.session_state.platform_preview_row = None
if "platform_involved_teams_input" not in st.session_state:
    st.session_state.platform_involved_teams_input = []
if "app_stage" not in st.session_state:
    st.session_state.app_stage = "home"
if "home_new_mode_input" not in st.session_state:
    st.session_state.home_new_mode_input = "bilan_personnel"
if "home_reload_year_input" not in st.session_state:
    st.session_state.home_reload_year_input = "2026"
if "home_reload_team_input" not in st.session_state:
    st.session_state.home_reload_team_input = ""
if "home_transition_target_view" not in st.session_state:
    st.session_state.home_transition_target_view = ""
if "form_baseline_signature" not in st.session_state:
    st.session_state.form_baseline_signature = "[]"
if "active_form_identity" not in st.session_state:
    st.session_state.active_form_identity = ""
if "pending_form_switch" not in st.session_state:
    st.session_state.pending_form_switch = None
if "snoozed_form_identity" not in st.session_state:
    st.session_state.snoozed_form_identity = ""

COUNTRY_OPTIONS = [
    "FR",
    "BE",
    "CH",
    "DE",
    "ES",
    "IT",
    "GB",
    "PT",
    "NL",
    "US",
    "CA",
    "CN",
    "JP",
    "AU",
    "IL",
    "Autre (saisie manuelle)",
]
SAVE_MODE_SUFFIX = {"bilan_personnel": "bilan_annuel", "bilan_projet": "projet", "plateforme": "plateforme"}
STORAGE_ROOT = Path(__file__).resolve().parent / LAB_ID
CURRENT_YEAR = datetime.now(timezone.utc).year
BASE_YEAR_OPTIONS = [str(y) for y in range(2026, 2019, -1)] + ["1998"]
YEAR_OPTIONS = BASE_YEAR_OPTIONS
YEAR_OPTIONS_ASC = sorted(YEAR_OPTIONS, key=int)
DOSSIER_TYPE_BY_MODE = {
    "item_unique": "item",
    "bilan_personnel": "personnel",
    "bilan_projet": "projet",
    "plateforme": "plateforme",
}
POSTE_TYPE_OPTIONS = [
    "",
    "Chercheur.euses (CR/DR)",
    "Enseignant.es-Chercheur.euses",
    "PAR (Permanent)",
    "PAR (Contractuel)",
    "Doctorant.es",
    "Post-doctorantes",
    "Stagiaire",
]
PLATFORM_POSTE_TYPE_OPTIONS = ["responsable", "utilisateur"]
ADMIN_GESTION_POSTE_TYPE_OPTIONS = ["", "Contractuel", "Permanent"]
EMPLOYMENT_BY_POSTE_TYPE = {
    "Chercheur.euses (CR/DR)": "Permanent",
    "Enseignant.es-Chercheur.euses": "Permanent",
    "PAR (Permanent)": "Permanent",
    "PAR (Contractuel)": "Contractuel",
    "Doctorant.es": "Contractuel",
    "Post-doctorantes": "Contractuel",
    "Stagiaire": "Contractuel",
}
POSTE_TYPE_LEGACY_MAP = {
    "CR": "Chercheur.euses (CR/DR)",
    "DR": "Chercheur.euses (CR/DR)",
    "Chercheur.euses": "Chercheur.euses (CR/DR)",
    "Doctorant": "Doctorant.es",
    "ITA": "PAR (Permanent)",
    "ITA Contractuel": "PAR (Contractuel)",
    "Post-Doctorant": "Post-doctorantes",
    "Post-Doctorante": "Post-doctorantes",
    "Researcher": "Chercheur.euses (CR/DR)",
    "Teacher-Researcher": "Enseignant.es-Chercheur.euses",
    "PhD": "Doctorant.es",
    "Post-Doctoral Fellow": "Post-doctorantes",
}
MODE_OPTIONS = ["item_unique", "bilan_personnel", "bilan_projet", "plateforme"]
POSTE_OPTIONS = ["achats", "domicile_travail", "campagnes_terrain", "missions", "heures_calcul"]
PLATFORM_OPTIONS = [
    "Microscope Confocal",
    "Spectrometre de Masse",
    "Laser Femto",
    "Cryo-MEB",
    "Diffractometre RX",
]
PLATFORM_INVOLVED_TEAM_CODES = [code for code in TEAM_OPTIONS.keys() if code and code != "ADMIN_GESTION"]

GEOCACHE_PATH = Path(__file__).resolve().parent / "Moulinette_missions" / "Data" / "Config" / "geocache.json"


def _normalize_city_key(value: str) -> str:
    txt = unicodedata.normalize("NFKD", str(value or ""))
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = txt.lower().strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def _build_city_country_index() -> dict[str, dict[str, int]]:
    out: dict[str, dict[str, int]] = {}
    if not GEOCACHE_PATH.exists():
        return out
    try:
        raw = json.loads(GEOCACHE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return out
    if not isinstance(raw, dict):
        return out
    for key, payload in raw.items():
        if not isinstance(payload, dict):
            continue
        cc = str(payload.get("countryCode", "")).strip().upper()
        if not cc:
            continue
        key_s = str(key or "").strip()
        variants: list[tuple[str, int]] = [(key_s, 3)]
        if " (" in key_s:
            base_part, suffix_part = key_s.split(" (", 1)
            base_city = base_part.strip()
            suffix_clean = suffix_part.strip().rstrip(")").strip().upper()
            # If suffix is an ISO-like 2-letter code, only trust it when it matches countryCode.
            if re.fullmatch(r"[A-Z]{2}", suffix_clean):
                if suffix_clean == cc:
                    variants.append((base_city, 4))
            else:
                variants.append((base_city, 4))
        address = str(payload.get("address", "")).strip()
        if address:
            variants.append((address.split(",", 1)[0].strip(), 1))
        for v, weight in variants:
            nv = _normalize_city_key(v)
            if not nv:
                continue
            out.setdefault(nv, {})
            out[nv][cc] = out[nv].get(cc, 0) + int(weight)
    return out


CITY_COUNTRY_INDEX = _build_city_country_index()


def _country_option_label(country_value: str) -> str:
    raw = str(country_value or "").strip()
    code = raw.upper()
    if len(code) == 2 and code in COUNTRY_LABEL_BY_CODE:
        return COUNTRY_LABEL_BY_CODE[code]
    return raw


def _country_option_code(country_value: str) -> str:
    raw = str(country_value or "").strip()
    if not raw:
        return ""
    code = raw.upper()
    if len(code) == 2:
        return code
    normalized_raw = _normalize_city_key(raw)
    for cc, label in COUNTRY_LABEL_BY_CODE.items():
        if _normalize_city_key(label) == normalized_raw:
            return cc
    return raw


def _year_sort_key(y: str) -> tuple[int, str]:
    return (int(y), y) if str(y).isdigit() else (-1, str(y))


def _factor_reference(group: str, key: str = "") -> str:
    ref_group = DEFAULT_FACTOR_REFERENCES.get(group, "")
    if isinstance(ref_group, dict):
        return str(ref_group.get(key, "")).strip()
    return str(ref_group).strip()


def _factor_uncertainty_pct(group: str, key: str = "") -> float:
    unc_group = DEFAULT_FACTOR_UNCERTAINTY_PCT.get(group, 10.0)
    if isinstance(unc_group, dict):
        value = unc_group.get(key, 10.0)
    else:
        value = unc_group
    try:
        out = float(value)
    except Exception:
        out = 10.0
    return out if out >= 0 else 10.0


def _format_factor_value_pm(value: float, uncertainty_pct: float) -> str:
    base = float(value)
    unc = abs(base) * float(uncertainty_pct if uncertainty_pct >= 0 else 10.0) / 100.0
    return f"{base:.6g} ± {unc:.6g}"


def _render_factor_source(reference: str) -> None:
    st.markdown(f"**{tf('Source', 'Source')}**")
    ref = str(reference or "").strip()
    if not ref:
        st.caption(tf("Reference a completer", "Reference to fill"))
        return
    if ref.startswith(("http://", "https://")):
        st.markdown(f"[{tf('Ouvrir la source', 'Open source')}]({ref})")
        return
    st.caption(ref)


def _render_default_factor_with_source(
    label: str,
    value: float,
    uncertainty_pct: float,
    reference: str,
    *,
    ignored: bool = False,
    key: str | None = None,
) -> None:
    col_factor, col_source = st.columns([3, 2])
    with col_factor:
        kwargs: dict = {
            "value": _format_factor_value_pm(value, uncertainty_pct),
            "disabled": True,
        }
        if key:
            kwargs["key"] = key
        st.text_input(label, **kwargs)
        if ignored:
            st.caption(tf("Facteur affiche mais non pris en compte.", "Factor displayed but not used."))
    with col_source:
        _render_factor_source(reference)


def _external_emissions_input(
    checkbox_label: str,
    key_prefix: str,
    *,
    default_checked: bool = False,
    default_value: float = 0.0,
) -> tuple[bool, float | None, str]:
    c1, c2, c3 = st.columns([1.8, 1.2, 1.2])
    use_external = c1.checkbox(checkbox_label, value=default_checked, key=f"{key_prefix}_use_external")
    external_total = None
    external_uncertainty_raw = ""
    if use_external:
        external_total = c2.number_input(
            tf("Total emissions (kgCO2e)", "Total emissions (kgCO2e)"),
            min_value=0.0,
            value=float(default_value),
            step=0.1,
            key=f"{key_prefix}_external_total",
        )
        external_uncertainty_raw = c3.text_input(
            tf("Incertitude (kgCO2e)", "Uncertainty (kgCO2e)"),
            value=str(st.session_state.get(f"{key_prefix}_external_uncertainty", "")),
            key=f"{key_prefix}_external_uncertainty",
            help=tf("Laisser vide: 10% des emissions.", "Leave empty: 10% of emissions."),
        )
    return bool(use_external), float(external_total) if external_total is not None else None, external_uncertainty_raw


def _apply_external_emissions(
    row: dict,
    use_external: bool,
    external_total_kgco2e: float | None,
) -> None:
    row["uses_external_emissions"] = bool(use_external)
    row["external_emissions_kgco2e"] = float(external_total_kgco2e) if use_external and external_total_kgco2e is not None else None
    if use_external and external_total_kgco2e is not None:
        row["emissions_kgco2e"] = float(external_total_kgco2e)


def _apply_factor_uncertainty(
    row: dict,
    uncertainty_pct: float,
) -> None:
    emissions = float(row.get("emissions_kgco2e", 0.0) or 0.0)
    pct_value = float(uncertainty_pct if uncertainty_pct >= 0 else 10.0)
    row["uncertainty_factor_pct"] = pct_value
    row["uncertainty_kgco2e"] = emissions * pct_value / 100.0
    row["uncertainty_missing_defaulted"] = False


def _apply_external_uncertainty(
    row: dict,
    use_external: bool,
    uncertainty_kgco2e_raw: str,
    *,
    default_pct_if_missing: float = 10.0,
) -> bool:
    if not use_external:
        return False
    emissions = float(row.get("emissions_kgco2e", 0.0) or 0.0)
    raw = str(uncertainty_kgco2e_raw or "").strip().replace(",", ".")
    missing = False
    if raw:
        try:
            uncertainty_kg = float(raw)
            if uncertainty_kg < 0:
                missing = True
                uncertainty_kg = emissions * float(default_pct_if_missing) / 100.0
        except Exception:
            missing = True
            uncertainty_kg = emissions * float(default_pct_if_missing) / 100.0
    else:
        missing = True
        uncertainty_kg = emissions * float(default_pct_if_missing) / 100.0
    emissions = float(row.get("emissions_kgco2e", 0.0) or 0.0)
    pct_value = (float(uncertainty_kg) / emissions * 100.0) if emissions > 0 else float(default_pct_if_missing)
    row["uncertainty_factor_pct"] = float(pct_value)
    row["uncertainty_kgco2e"] = float(uncertainty_kg)
    row["uncertainty_missing_defaulted"] = bool(missing)
    return bool(missing)


def _set_uncertainty_notice_if_missing(missing: bool) -> None:
    if missing:
        st.session_state.uncertainty_notice = tf(
            "🔺 Incertitude non fournie: 10% applique par defaut.",
            "🔺 Uncertainty not provided: default 10% applied.",
        )


def _mission_default_step(arrival_city: str = "", arrival_country: str = "FR") -> dict:
    return {
        "departure_city": "",
        "departure_country": "FR",
        "arrival_city": arrival_city,
        "arrival_country": arrival_country or "FR",
        "transport_mode": "plane",
        "round_trip": False,
        "use_external": False,
        "external_total_kgco2e": None,
        "external_uncertainty_kgco2e": "",
    }


def _add_mission_step_with_split_destination() -> None:
    steps = st.session_state.get("mission_steps_input", [])
    if not isinstance(steps, list) or not steps:
        st.session_state.mission_steps_input = [_mission_default_step()]
        return
    updated = [dict(s) for s in steps]
    last = updated[-1]
    final_city = str(last.get("arrival_city", "") or "")
    final_country = str(last.get("arrival_country", "FR") or "FR")
    last["arrival_city"] = ""
    last["arrival_country"] = final_country
    updated.append(_mission_default_step(arrival_city=final_city, arrival_country=final_country))
    st.session_state.mission_steps_input = updated


def _mission_calc_input_key(
    mission_id: str,
    mission_ref: str,
    steps: list[dict],
    use_external: bool,
    external_total: float | None,
    external_unc_raw: str,
) -> str:
    steps_key = tuple(
        (
            str(s.get("departure_city", "")),
            str(s.get("departure_country", "")),
            str(s.get("arrival_city", "")),
            str(s.get("arrival_country", "")),
            str(s.get("transport_mode", "")),
            bool(s.get("round_trip", False)),
            bool(s.get("use_external", False)),
            None if s.get("external_total_kgco2e") is None else float(s.get("external_total_kgco2e")),
            str(s.get("external_uncertainty_kgco2e", "")).strip(),
        )
        for s in steps
    )
    return str(
        (
            _clean(mission_id),
            _clean(mission_ref),
            steps_key,
            bool(use_external),
            None if external_total is None else float(external_total),
            str(external_unc_raw or "").strip(),
        )
    )


def _format_total_with_uncertainty(total: float, uncertainty: float) -> str:
    if abs(float(total)) > 100:
        return f"{total:.0f} ± {uncertainty:.0f} kgCO2e"
    return f"{total:.1f} ± {uncertainty:.1f} kgCO2e"


def _ensure_uncertainty_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "emissions_kgco2e" not in out.columns:
        out["emissions_kgco2e"] = 0.0
    out["emissions_kgco2e"] = pd.to_numeric(out["emissions_kgco2e"], errors="coerce").fillna(0.0)

    if "uncertainty_factor_pct" in out.columns:
        out["uncertainty_factor_pct"] = pd.to_numeric(out["uncertainty_factor_pct"], errors="coerce")
    else:
        out["uncertainty_factor_pct"] = pd.NA

    if "uncertainty_kgco2e" in out.columns:
        out["uncertainty_kgco2e"] = pd.to_numeric(out["uncertainty_kgco2e"], errors="coerce")
    else:
        out["uncertainty_kgco2e"] = pd.NA

    if "uncertainty_missing_defaulted" not in out.columns:
        out["uncertainty_missing_defaulted"] = False
    out["uncertainty_missing_defaulted"] = out["uncertainty_missing_defaulted"].fillna(False).astype(bool)

    missing_all = out["uncertainty_factor_pct"].isna() & out["uncertainty_kgco2e"].isna()
    out.loc[missing_all, "uncertainty_factor_pct"] = 10.0
    out.loc[missing_all, "uncertainty_kgco2e"] = out.loc[missing_all, "emissions_kgco2e"] * 0.10
    out.loc[missing_all, "uncertainty_missing_defaulted"] = True

    missing_pct = out["uncertainty_factor_pct"].isna() & out["uncertainty_kgco2e"].notna()
    out.loc[missing_pct, "uncertainty_factor_pct"] = (
        out.loc[missing_pct, "uncertainty_kgco2e"] / out.loc[missing_pct, "emissions_kgco2e"].replace(0, pd.NA) * 100.0
    ).fillna(10.0)

    missing_kg = out["uncertainty_kgco2e"].isna() & out["uncertainty_factor_pct"].notna()
    out.loc[missing_kg, "uncertainty_kgco2e"] = out.loc[missing_kg, "emissions_kgco2e"] * out.loc[missing_kg, "uncertainty_factor_pct"] / 100.0

    out["emissions_low_kgco2e"] = (out["emissions_kgco2e"] - out["uncertainty_kgco2e"]).clip(lower=0.0)
    out["emissions_high_kgco2e"] = out["emissions_kgco2e"] + out["uncertainty_kgco2e"]
    return out


def _styled_dataframe(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    is_light = st.session_state.get("ui_theme", "darkmode") == "lightmode"
    if is_light:
        bg = "#ffffff"
        fg = "#102214"
        header_bg = "#f2f8ee"
        border = "#d8e6d0"
    else:
        bg = "#0d1a33"
        fg = "#e8f0ff"
        header_bg = "#102248"
        border = "#2b4574"
    return (
        df.style.set_properties(**{"background-color": bg, "color": fg, "border-color": border})
        .set_table_styles(
            [
                {"selector": "th", "props": [("background-color", header_bg), ("color", fg), ("border-color", border)]},
                {"selector": "td", "props": [("border-color", border)]},
                {"selector": "table", "props": [("border-collapse", "collapse"), ("width", "100%")]},
            ]
        )
    )


def _render_themed_dataframe(
    df: pd.DataFrame,
    *,
    use_container_width: bool = True,
    hide_index: bool = False,
    light_use_table: bool = True,
) -> None:
    is_light = st.session_state.get("ui_theme", "darkmode") == "lightmode"
    styled = _styled_dataframe(df)
    if hide_index:
        styled = styled.hide(axis="index")
    if is_light and light_use_table:
        st.table(styled)
        return
    st.dataframe(styled, use_container_width=use_container_width, hide_index=hide_index)


def _mode_label(m: str) -> str:
    return {
        "item_unique": "Quickcheck",
        "bilan_personnel": tf("Bilan personnel", "Personal assessment"),
        "bilan_projet": tf("Bilan projet", "Project assessment"),
        "plateforme": tf("Plateforme", "Platform"),
    }.get(m, m)


def _poste_label(p: str) -> str:
    return {
        "achats": tf("Achats", "Purchases"),
        "domicile_travail": tf("Domicile-travail", "Commute"),
        "campagnes_terrain": tf("Campagnes terrain", "Field campaigns"),
        "missions": tf("Missions", "Missions"),
        "heures_calcul": tf("Heures de calcul", "Compute hours"),
        "plateforme": tf("Plateforme", "Platform"),
    }.get(p, p)


def _normalize_mode_value(v: str) -> str:
    mapping = {
        "item_unique": "item_unique",
        "quickcheck": "item_unique",
        "Quickcheck": "item_unique",
        "bilan_personnel": "bilan_personnel",
        "bilan_projet": "bilan_projet",
        "plateforme": "plateforme",
        "Item unique": "item_unique",
        "Single item": "item_unique",
        "Bilan personnel": "bilan_personnel",
        "Personal assessment": "bilan_personnel",
        "Bilan projet": "bilan_projet",
        "Project assessment": "bilan_projet",
        "Plateforme": "plateforme",
        "Platform": "plateforme",
    }
    return mapping.get(str(v), "item_unique")


def _normalize_poste_type_for_overview(team_code: str, poste_type: str) -> str:
    p = str(poste_type or "").strip()
    if team_code == "ADMIN_GESTION":
        if p.lower() == "contractuel":
            return "Admin C"
        if p.lower() == "permanent":
            return "Admin"
    return POSTE_TYPE_LEGACY_MAP.get(p, p)


def _employment_type(team_code: str, poste_type: str) -> str:
    p = str(poste_type or "").strip()
    if team_code == "ADMIN_GESTION":
        if p.lower() == "contractuel":
            return "Contractuel"
        if p.lower() == "permanent":
            return "Permanent"
        return ""
    normalized = POSTE_TYPE_LEGACY_MAP.get(p, p)
    return EMPLOYMENT_BY_POSTE_TYPE.get(normalized, "")


def _poste_type_label(value: str) -> str:
    mapping = {
        "": "",
        "Chercheur.euses (CR/DR)": tf("Chercheur.euses (CR/DR)", "Researcher"),
        "Enseignant.es-Chercheur.euses": tf("Enseignant.es-Chercheur.euses", "Teacher-Researcher"),
        "PAR (Permanent)": tf("PAR (Permanent)", "PAR (Permanent)"),
        "PAR (Contractuel)": tf("PAR (Contractuel)", "PAR (Contractual)"),
        "Doctorant.es": tf("Doctorant.es", "PhD"),
        "Post-doctorantes": tf("Post-doctorantes", "Post-Doctoral Fellow"),
        "Stagiaire": tf("Stagiaire", "Intern"),
        "Contractuel": tf("Contractuel", "Contractual"),
        "Permanent": tf("Permanent", "Permanent"),
    }
    return mapping.get(str(value), str(value))


def _clean(v: str) -> str:
    return v.strip()


def _domicile_days_per_year_from_row(row: dict) -> float:
    try:
        if row.get("days_per_year") is not None:
            return float(row.get("days_per_year", 0.0) or 0.0)
        days = float(row.get("days_per_week", 0.0) or 0.0)
        weeks = float(row.get("weeks_per_year", 0.0) or 0.0)
        return days * weeks
    except Exception:
        return 0.0


DOMICILE_DAYS_PRESETS = [
    ("p1", "4x/sem • 35/an", 140.0),
    ("p2", "3x/sem • 45/an", 135.0),
    ("p3", "5x/sem • 45/an", 225.0),
]


def _set_domicile_days_from_preset(target_days_key: str, days_value: float) -> None:
    st.session_state[target_days_key] = float(days_value)


def _render_domicile_days_preset_cell(col, target_days_key: str, preset_id: str, label: str, days_value: float, button_key_prefix: str) -> None:
    with col:
        st.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
        st.button(
            label,
            key=f"{button_key_prefix}_{preset_id}",
            on_click=_set_domicile_days_from_preset,
            args=(target_days_key, days_value),
            use_container_width=True,
        )


def _slug(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_-]+", "_", _clean(s))
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "anonyme"


def _team_folder(team_code: str, mode: str = "", is_anonymous: bool = False) -> str:
    if mode == "plateforme":
        return "PLATEFORMES"
    if mode == "bilan_projet" and not team_code:
        return "ANONYME"
    return team_code if team_code else "UNSPECIFIED"


def _year_folder(year_value: str) -> str:
    y = str(year_value or "").strip()
    return y if y in YEAR_OPTIONS else YEAR_OPTIONS[0]


def _year_from_path(path: Path) -> str:
    try:
        rel = path.relative_to(STORAGE_ROOT)
    except ValueError:
        return ""
    parts = rel.parts
    if parts and parts[0] in YEAR_OPTIONS:
        return parts[0]
    return ""


def _team_from_path(path: Path) -> str:
    try:
        rel = path.relative_to(STORAGE_ROOT)
    except ValueError:
        return ""
    parts = rel.parts
    return parts[1] if len(parts) > 1 else ""


def _list_saved_forms() -> list[Path]:
    if not STORAGE_ROOT.exists():
        return []
    return sorted(STORAGE_ROOT.rglob("*.xlsx"), key=lambda p: str(p).lower())


def _default_save_path(mode: str, year_value: str, team_code: str, person_label: str, project_name: str = "") -> Path:
    suffix = SAVE_MODE_SUFFIX[mode]
    owner = _slug(person_label) if _clean(person_label) else "anonyme"
    year_folder = _year_folder(year_value)
    folder = _team_folder(team_code=team_code, mode=mode, is_anonymous=owner == "anonyme")
    if mode == "bilan_projet":
        project = _slug(project_name) if _clean(project_name) else "projet"
        return STORAGE_ROOT / year_folder / folder / f"{owner}_{project}_{suffix}.xlsx"
    if mode == "plateforme":
        platform = _slug(st.session_state.get("platform_name_input", "plateforme"))
        return STORAGE_ROOT / year_folder / folder / f"{platform}_{owner}_{suffix}.xlsx"
    return STORAGE_ROOT / year_folder / folder / f"{owner}_{suffix}.xlsx"


def _ensure_storage_tree() -> None:
    STORAGE_ROOT.mkdir(parents=True, exist_ok=True)
    team_folders = [code for code in TEAM_OPTIONS.keys() if code]
    team_folders.extend(["UNSPECIFIED", "ANONYME", "PLATEFORMES"])
    team_folders = sorted(set(team_folders))
    for year in YEAR_OPTIONS:
        for team in team_folders:
            (STORAGE_ROOT / year / team).mkdir(parents=True, exist_ok=True)


def _apply_loaded_metadata(meta: dict, entries: list[dict], source_path: str = "") -> None:
    st.session_state.dossier_id = str(meta.get("dossier_id", st.session_state.dossier_id))
    st.session_state.loaded_owner_label = str(meta.get("owner_label", ""))
    if st.session_state.loaded_owner_label:
        st.session_state.person_label_input = st.session_state.loaded_owner_label

    st.session_state.poste_type_input = str(meta.get("poste_type", st.session_state.get("poste_type_input", "")))
    st.session_state.is_anonymous = bool(meta.get("anonyme_default", st.session_state.get("is_anonymous", True)))

    dossier_type = str(meta.get("dossier_type", "")).strip().lower()
    mode_from_type = {"item": "item_unique", "personnel": "bilan_personnel", "projet": "bilan_projet", "plateforme": "plateforme"}.get(dossier_type, "")
    if mode_from_type:
        st.session_state.mode_input = mode_from_type

    project_name = str(meta.get("dossier_name", "")).strip()
    if project_name:
        st.session_state.project_name_input = project_name
    if mode_from_type == "plateforme":
        platform_name = str(meta.get("platform_name", project_name or "")).strip()
        if platform_name:
            st.session_state.platform_name_input = platform_name

    team_code = str(meta.get("team_code", "")).strip()
    if not team_code and entries:
        team_values = sorted({str(e.get("team_code", "")).strip() for e in entries if str(e.get("team_code", "")).strip()})
        if len(team_values) == 1:
            team_code = team_values[0]
    st.session_state.team_code_input = team_code if (team_code in TEAM_OPTIONS or team_code == "PLATEFORMES") else ""
    meta_year = str(meta.get("year", "")).strip()
    path_year = _year_from_path(Path(source_path)) if source_path else ""
    year_value = meta_year if meta_year in YEAR_OPTIONS else path_year
    if year_value in YEAR_OPTIONS:
        st.session_state.year_input = year_value


def _consume_pending_loaded_state() -> None:
    meta = st.session_state.pending_loaded_meta
    entries = st.session_state.pending_loaded_entries
    if meta is None or entries is None:
        return

    st.session_state.entries = entries
    st.session_state.loaded_form_path = st.session_state.pending_loaded_path or ""
    st.session_state.confirm_overwrite_pending = False
    _apply_loaded_metadata(meta, entries, source_path=st.session_state.loaded_form_path)
    source_name = Path(st.session_state.loaded_form_path).name if st.session_state.loaded_form_path else tf("fichier importe", "imported file")
    st.session_state.loaded_notice = tf(
        f"Dossier charge: {source_name} ({len(entries)} ligne(s))",
        f"Form loaded: {source_name} ({len(entries)} row(s))",
    )

    st.session_state.pending_loaded_meta = None
    st.session_state.pending_loaded_entries = None
    st.session_state.pending_loaded_path = ""
    st.session_state.form_baseline_signature = _entries_signature(st.session_state.entries)
    st.session_state.active_form_identity = ""
    st.session_state.snoozed_form_identity = ""


def _entries_signature(entries: list[dict]) -> str:
    try:
        return json.dumps(entries, sort_keys=True, ensure_ascii=False, default=str, separators=(",", ":"))
    except Exception:
        return str(len(entries))


def _is_form_dirty() -> bool:
    baseline = str(st.session_state.get("form_baseline_signature", ""))
    current = _entries_signature(st.session_state.get("entries", []))
    return baseline != current


def _current_owner_for_save(is_anonymous: bool, person_label: str) -> str:
    if is_anonymous:
        return "anonyme"
    owner = _clean(person_label)
    return owner or "anonyme"


def _compute_form_identity(mode: str, year_value: str, team_code: str, owner_for_save: str, project_name: str) -> str:
    if mode not in SAVE_MODE_SUFFIX:
        return ""
    project_part = _slug(project_name) if mode == "bilan_projet" else "-"
    return "|".join([mode, _year_folder(year_value), _team_folder(team_code=team_code, mode=mode), _slug(owner_for_save), project_part])


def _load_local_form_now(path: Path) -> None:
    meta, imported_entries = import_excel_entries(path.read_bytes())
    # Queue load state for next rerun so widget-backed keys are updated safely
    # before widgets are instantiated.
    st.session_state.pending_loaded_meta = meta
    st.session_state.pending_loaded_entries = imported_entries
    st.session_state.pending_loaded_path = str(path)


def _create_new_form_context(target_path: Path, owner_for_save: str) -> None:
    st.session_state.entries = []
    st.session_state.loaded_form_path = str(target_path)
    st.session_state.loaded_owner_label = owner_for_save
    st.session_state.confirm_overwrite_pending = False
    st.session_state.form_baseline_signature = _entries_signature(st.session_state.entries)
    st.session_state.loaded_notice = tf(
        f"Nouveau dossier initialise: {target_path.name}",
        f"New form initialized: {target_path.name}",
    )


def _execute_form_switch(target_path: Path, identity: str, owner_for_save: str) -> None:
    if target_path.exists():
        _load_local_form_now(target_path)
    else:
        _create_new_form_context(target_path, owner_for_save)
    st.session_state.active_form_identity = identity
    st.session_state.pending_form_switch = None
    st.session_state.snoozed_form_identity = ""
    st.rerun()


def _auto_resolve_form_context(mode: str, year_value: str, team_code: str, is_anonymous: bool, person_label: str, project_name: str) -> None:
    if mode not in SAVE_MODE_SUFFIX:
        st.session_state.active_form_identity = ""
        st.session_state.pending_form_switch = None
        st.session_state.snoozed_form_identity = ""
        return

    owner_for_save = _current_owner_for_save(is_anonymous, person_label)
    identity = _compute_form_identity(mode, year_value, team_code, owner_for_save, project_name)
    if not identity:
        return
    if identity == st.session_state.get("active_form_identity", ""):
        st.session_state.pending_form_switch = None
        return
    if identity == st.session_state.get("snoozed_form_identity", ""):
        return

    target_path = _default_save_path(
        mode=mode,
        year_value=year_value,
        team_code=team_code,
        person_label=owner_for_save,
        project_name=project_name,
    )
    current_path = str(st.session_state.get("loaded_form_path", "") or "")
    if current_path and str(target_path) == current_path:
        st.session_state.active_form_identity = identity
        st.session_state.pending_form_switch = None
        return

    if _is_form_dirty():
        st.session_state.pending_form_switch = {
            "identity": identity,
            "target_path": str(target_path),
            "owner_for_save": owner_for_save,
            "exists": bool(target_path.exists()),
        }
        return

    _execute_form_switch(target_path, identity, owner_for_save)


def _save_local_form(
    save_path: Path,
    year_value: str,
    dossier_id: str,
    dossier_type: str,
    project_name: str,
    platform_name: str,
    is_anonymous: bool,
    owner_for_save: str,
    poste_type: str,
    team_code: str,
    entries: list[dict],
) -> None:
    save_path.parent.mkdir(parents=True, exist_ok=True)
    save_meta = {
        "dossier_id": dossier_id,
        "lab_id": LAB_ID,
        "year": _year_folder(year_value),
        "dossier_type": dossier_type,
        "dossier_name": project_name if dossier_type == "projet" else "",
        "platform_name": platform_name if dossier_type == "plateforme" else "",
        "anonyme_default": bool(is_anonymous),
        "owner_label": owner_for_save,
        "poste_type": poste_type,
        "team_code": team_code,
        "schema_version": "1.0",
        "exported_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
    }
    save_bytes = export_excel_bytes(meta=save_meta, entries=entries)
    save_path.write_bytes(save_bytes)


def _missions_arcs_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    wanted = ["poste", "mission_id", "departure_city", "arrival_city", "departure_lat", "departure_lon", "arrival_lat", "arrival_lon", "round_trip"]
    cols = [c for c in wanted if c in df.columns]
    if len(cols) < 8:
        return pd.DataFrame()
    mdf = df[df["poste"] == "missions"][cols].copy()
    mdf = mdf.dropna(subset=["departure_lat", "departure_lon", "arrival_lat", "arrival_lon"])
    return mdf


def _mission_deck(arcs_df: pd.DataFrame, height: int = 320) -> pdk.Deck:
    center_lat = float(pd.concat([arcs_df["departure_lat"], arcs_df["arrival_lat"]]).mean())
    center_lon = float(pd.concat([arcs_df["departure_lon"], arcs_df["arrival_lon"]]).mean())
    line_layer = pdk.Layer(
        "LineLayer",
        data=arcs_df,
        get_source_position=["departure_lon", "departure_lat"],
        get_target_position=["arrival_lon", "arrival_lat"],
        get_color=[40, 116, 240, 200],
        auto_highlight=True,
        pickable=True,
        get_width=3,
    )
    labels_df = arcs_df.copy()
    labels_df["mid_lon"] = (labels_df["departure_lon"] + labels_df["arrival_lon"]) / 2.0
    labels_df["mid_lat"] = (labels_df["departure_lat"] + labels_df["arrival_lat"]) / 2.0
    labels_df["arrow_label"] = labels_df["round_trip"].apply(lambda v: "↔" if bool(v) else "→")
    text_layer = pdk.Layer(
        "TextLayer",
        data=labels_df,
        get_position=["mid_lon", "mid_lat"],
        get_text="arrow_label",
        get_size=18,
        get_color=[10, 10, 10, 220],
        get_angle=0,
        get_text_anchor="'middle'",
        get_alignment_baseline="'center'",
        pickable=False,
    )
    view_state = pdk.ViewState(latitude=center_lat, longitude=center_lon, zoom=3, pitch=0, bearing=0)
    tooltip = {"text": "{mission_id}\n{departure_city} -> {arrival_city}"}
    return pdk.Deck(
        layers=[line_layer, text_layer],
        initial_view_state=view_state,
        tooltip=tooltip,
        map_provider="carto",
        map_style="light",
        height=height,
    )


def _country_input(label: str, key_prefix: str, default_country: str = "FR", city_hint: str = "") -> str:
    other_label = tf("Autre (saisie manuelle)", "Other (manual input)")
    base_options = [c for c in COUNTRY_OPTIONS if c != "Autre (saisie manuelle)"]
    city_key = _normalize_city_key(city_hint)
    suggested_scores = CITY_COUNTRY_INDEX.get(city_key, {})
    suggested = sorted(suggested_scores.keys(), key=lambda c: (-int(suggested_scores.get(c, 0)), c))
    country_options = suggested[:] if suggested else base_options[:]
    if not country_options:
        country_options = ["FR"]
    country_options.append(other_label)

    select_key = f"{key_prefix}_select"
    current_value = st.session_state.get(select_key, "")
    if suggested:
        # When a city match exists, always default to the top-ranked country for this city.
        st.session_state[select_key] = suggested[0]
    elif current_value not in country_options:
        default_value = _country_option_code(default_country)
        st.session_state[select_key] = default_value if default_value in country_options else country_options[0]

    selected = st.selectbox(label, country_options, key=select_key, format_func=_country_option_label)
    if selected == other_label:
        custom = st.text_input(f"{label} {tf('(saisie)', '(manual)')}", value="", key=f"{key_prefix}_custom")
        return _clean(custom)
    return selected


def _propagate_arrival_city_to_next(step_idx: int) -> None:
    steps = st.session_state.get("mission_steps_input", [])
    if not isinstance(steps, list) or step_idx < 0 or step_idx + 1 >= len(steps):
        return
    src_key = f"mis_step_{step_idx}_arr_city"
    dst_key = f"mis_step_{step_idx + 1}_dep_city"
    src = _clean(st.session_state.get(src_key, ""))
    dst = _clean(st.session_state.get(dst_key, ""))
    if src and not dst:
        st.session_state[dst_key] = src


def _propagate_departure_city_to_prev(step_idx: int) -> None:
    steps = st.session_state.get("mission_steps_input", [])
    if not isinstance(steps, list) or step_idx <= 0 or step_idx >= len(steps):
        return
    src_key = f"mis_step_{step_idx}_dep_city"
    dst_key = f"mis_step_{step_idx - 1}_arr_city"
    src = _clean(st.session_state.get(src_key, ""))
    dst = _clean(st.session_state.get(dst_key, ""))
    if src and not dst:
        st.session_state[dst_key] = src


def _load_entries_from_local_forms() -> list[dict]:
    entries: list[dict] = []
    for path in _list_saved_forms():
        try:
            meta, imported_entries = import_excel_entries(path.read_bytes())
            team_code = str(meta.get("team_code", "")).strip()
            poste_type_raw = str(meta.get("poste_type", "")).strip()
            poste_type_norm = _normalize_poste_type_for_overview(team_code, poste_type_raw)
            employment = _employment_type(team_code, poste_type_raw)
            year_value = str(meta.get("year", "")).strip() or _year_from_path(path)
            for row in imported_entries:
                row["meta_team_code"] = team_code
                row["meta_poste_type"] = poste_type_norm
                row["meta_employment_type"] = employment
                row["meta_year"] = year_value
            entries.extend(imported_entries)
        except Exception:
            continue
    return entries


def _render_labo_overview() -> None:
    is_light = st.session_state.get("ui_theme", "darkmode") == "lightmode"
    chart_bg = "#ffffff" if is_light else "#0d1a33"
    axis_label_color = "#102214" if is_light else "#e8f0ff"
    axis_title_color = "#102214" if is_light else "#e8f0ff"
    legend_label_color = "#102214" if is_light else "#e8f0ff"
    legend_title_color = "#102214" if is_light else "#e8f0ff"
    grid_color = "#dfe7dc" if is_light else "#35507a"

    def _style_chart(chart: alt.Chart) -> alt.Chart:
        return (
            chart.configure(background=chart_bg)
            .configure_view(fill=chart_bg, stroke=None)
            .configure_axis(
                labelColor=axis_label_color,
                titleColor=axis_title_color,
                gridColor=grid_color,
            )
            .configure_legend(
                labelColor=legend_label_color,
                titleColor=legend_title_color,
            )
        )

    def _render_share_bar(data: pd.DataFrame, label_col: str, section_title: str) -> None:
        if data.empty:
            return
        d = data.copy()
        total = float(d["emissions_kgco2e"].sum())
        if total <= 0:
            return
        d["pct"] = d["emissions_kgco2e"] / total * 100
        d["pct_label"] = d["pct"].map(lambda x: f"{x:.1f}%")
        d["abs_label"] = d["emissions_kgco2e"].map(lambda x: f"{x:.1f} kg")
        d["t_label"] = d["emissions_kgco2e"].map(lambda x: x / 1000.0)
        d["legend_label"] = d.apply(
            lambda r: f"{r[label_col]} {r['t_label']:.2f} t eq CO2 ({r['pct']:.1f}%)",
            axis=1,
        )
        d["bar_group"] = section_title

        st.caption(
            tf("Total", "Total") + f" {section_title}: {total:.1f} kgCO2e"
        )

        bar = (
            alt.Chart(d)
            .mark_bar(size=48)
            .encode(
                y=alt.Y("bar_group:N", axis=None),
                x=alt.X("emissions_kgco2e:Q", stack="normalize", axis=alt.Axis(format="%")),
                color=alt.Color("legend_label:N", legend=alt.Legend(title=section_title)),
                tooltip=[
                    alt.Tooltip("legend_label:N", title=section_title),
                    alt.Tooltip("emissions_kgco2e:Q", title="kgCO2e", format=".2f"),
                    alt.Tooltip("pct:Q", title="%", format=".1f"),
                ],
            )
            .properties(height=130)
        )
        st.altair_chart(_style_chart(bar), use_container_width=True)

    st.write("## " + tf("Vue d'ensemble du labo", "Lab overview"))
    storage_label = f"`{STORAGE_ROOT.name}/`"
    st.caption(
        tf(
            f"Consolidation des formulaires locaux sauvegardes dans le dossier {storage_label}",
            f"Consolidation of local forms saved in folder {storage_label}.",
        )
    )
    st.caption(
        tf(
            "Note: les bilans personnels et projets peuvent contenir des recouvrements. Les plateformes sont superposees pour l'instant.",
            "Note: personal and project assessments may overlap. Platform forms are currently superposed.",
        )
    )

    local_entries = _load_entries_from_local_forms()

    if not local_entries:
        st.info(
            tf(
                f"Aucune donnee disponible. Sauvegarde d'abord des formulaires dans `{STORAGE_ROOT.name}/`.",
                f"No data available. Save forms first in `{STORAGE_ROOT.name}/`.",
            )
        )
        return

    ag_df = _ensure_uncertainty_columns(entries_to_df(local_entries))
    if "meta_team_code" not in ag_df.columns:
        ag_df["meta_team_code"] = ""
    if "meta_poste_type" not in ag_df.columns:
        ag_df["meta_poste_type"] = ""
    if "meta_employment_type" not in ag_df.columns:
        ag_df["meta_employment_type"] = ""
    if "meta_year" not in ag_df.columns:
        ag_df["meta_year"] = ""

    yearly_df = ag_df.copy()
    yearly_df["meta_year"] = yearly_df["meta_year"].astype(str).str.strip()
    yearly_df = yearly_df[yearly_df["meta_year"] != ""]

    st.write(f"### {tf('Evolution interannuelle des emissions', 'Interannual emissions trend')}")
    c1, c2, c3 = st.columns(3)
    color_by_team = c1.checkbox(tf("Equipe", "Team"), key="year_stack_team")
    color_by_poste = c2.checkbox(tf("Poste", "Category"), key="year_stack_poste")
    color_by_status = c3.checkbox(tf("Statut", "Status"), key="year_stack_status")

    if yearly_df.empty:
        st.caption(tf("Aucune annee renseignee dans les formulaires.", "No year is defined in saved forms."))
    else:
        yearly_total = (
            yearly_df.groupby("meta_year", as_index=False)["emissions_kgco2e"]
            .sum()
            .sort_values("meta_year")
        )
        yearly_total_chart = (
            alt.Chart(yearly_total)
            .mark_bar()
            .encode(
                x=alt.X("meta_year:N", title=tf("Annee", "Year"), sort=YEAR_OPTIONS_ASC),
                y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"),
                tooltip=[alt.Tooltip("meta_year:N", title=tf("Annee", "Year")), alt.Tooltip("emissions_kgco2e:Q", title="kgCO2e", format=".2f")],
            )
            .properties(height=280)
        )
        st.altair_chart(_style_chart(yearly_total_chart), use_container_width=True)

        def _render_yearly_stacked(df: pd.DataFrame, color_col: str, chart_title: str) -> None:
            d = df.copy()
            d = d[d[color_col].astype(str).str.strip() != ""]
            if d.empty:
                return
            if color_col == "team_group":
                d["color_label"] = d["team_group"].map(lambda x: TEAM_OPTIONS.get(x, x))
            elif color_col == "poste":
                d["color_label"] = d["poste"].map(_poste_label)
            else:
                d["color_label"] = d[color_col]
            grouped = (
                d.groupby(["meta_year", "color_label"], as_index=False)["emissions_kgco2e"]
                .sum()
                .sort_values(["meta_year", "emissions_kgco2e"], ascending=[True, False])
            )
            chart = (
                alt.Chart(grouped)
                .mark_bar()
                .encode(
                    x=alt.X("meta_year:N", title=tf("Annee", "Year"), sort=YEAR_OPTIONS_ASC),
                    y=alt.Y("emissions_kgco2e:Q", title="kgCO2e", stack=True),
                    color=alt.Color("color_label:N", title=chart_title),
                    tooltip=[
                        alt.Tooltip("meta_year:N", title=tf("Annee", "Year")),
                        alt.Tooltip("color_label:N", title=chart_title),
                        alt.Tooltip("emissions_kgco2e:Q", title="kgCO2e", format=".2f"),
                    ],
                )
                .properties(height=300)
            )
            st.altair_chart(_style_chart(chart), use_container_width=True)

        yearly_df["team_group"] = yearly_df["meta_team_code"].replace("", "UNSPECIFIED")
        if color_by_team:
            st.caption(tf("Coloration par equipe", "Color split by team"))
            _render_yearly_stacked(yearly_df, "team_group", tf("Equipe", "Team"))
        if color_by_poste:
            st.caption(tf("Coloration par poste", "Color split by category"))
            _render_yearly_stacked(yearly_df, "poste", tf("Poste", "Category"))
        if color_by_status:
            st.caption(tf("Coloration par statut", "Color split by status"))
            _render_yearly_stacked(yearly_df, "meta_employment_type", tf("Statut", "Status"))

    available_years = sorted(yearly_df["meta_year"].dropna().astype(str).unique(), key=int)
    if not available_years:
        st.info(tf("Aucune annee disponible pour l'analyse detaillee.", "No year available for detailed analysis."))
        return
    default_year_idx = available_years.index("1998") if "1998" in available_years else 0
    selected_year = st.selectbox(
        tf("Annee etudiee", "Study year"),
        options=available_years,
        index=default_year_idx,
        key="ag_year",
    )

    ag_df = ag_df[ag_df["meta_year"].astype(str) == str(selected_year)].copy()

    poste_filter = st.selectbox(tf("Filtre poste", "Category filter"), options=["ALL", "achats", "domicile_travail", "campagnes_terrain", "missions", "heures_calcul", "plateforme"], key="ag_poste")
    team_filter = st.selectbox(tf("Filtre equipe", "Team filter"), options=["ALL", *TEAM_OPTIONS.keys()], format_func=lambda x: tf("Toutes", "All") if x == "ALL" else TEAM_OPTIONS[x], key="ag_team")
    poste_type_values = sorted([v for v in ag_df["meta_poste_type"].dropna().astype(str).unique() if v])
    poste_type_filter = st.selectbox(
        tf("Filtre type de poste", "Position type filter"),
        options=["ALL", *poste_type_values],
        key="ag_poste_type",
    )
    employment_values = ["Permanent", "Contractuel"]
    employment_filter = st.selectbox(
        tf("Filtre statut", "Employment filter"),
        options=["ALL", *employment_values],
        key="ag_employment",
    )

    filtered = ag_df.copy()
    if poste_filter != "ALL":
        filtered = filtered[filtered["poste"] == poste_filter]
    if team_filter != "ALL":
        filtered = filtered[filtered["team_code"] == team_filter]
    if poste_type_filter != "ALL":
        filtered = filtered[filtered["meta_poste_type"] == poste_type_filter]
    if employment_filter != "ALL":
        filtered = filtered[filtered["meta_employment_type"] == employment_filter]

    total_emissions = float(filtered["emissions_kgco2e"].sum()) if not filtered.empty else 0.0
    total_uncertainty = float(filtered["uncertainty_kgco2e"].sum()) if not filtered.empty else 0.0
    uncertainty_pct = (total_uncertainty / total_emissions * 100.0) if total_emissions > 0 else 0.0
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(tf("Emissions", "Emissions"), f"{total_emissions:.2f} kgCO2e")
    m2.metric(tf("Incertitude", "Uncertainty"), f"±{total_uncertainty:.2f} ({uncertainty_pct:.1f}%)")
    m3.metric(tf("Formulaires", "Forms"), int(filtered["dossier_id"].nunique()) if not filtered.empty else 0)
    m4.metric(tf("Lignes", "Rows"), len(filtered))

    if filtered.empty:
        st.warning(tf("Aucune ligne apres filtrage.", "No rows after filtering."))
        return

    by_team = (
        filtered.assign(team=filtered["team_code"].replace("", "UNSPECIFIED"))
        .groupby("team", as_index=False)["emissions_kgco2e"]
        .sum()
        .sort_values("emissions_kgco2e", ascending=False)
    )
    by_poste = filtered.groupby("poste", as_index=False)["emissions_kgco2e"].sum().sort_values("emissions_kgco2e", ascending=False)
    by_poste_type = (
        filtered[filtered["meta_poste_type"].astype(str) != ""]
        .groupby("meta_poste_type", as_index=False)["emissions_kgco2e"]
        .sum()
        .sort_values("emissions_kgco2e", ascending=False)
    )
    by_employment = (
        filtered[filtered["meta_employment_type"].astype(str) != ""]
        .groupby("meta_employment_type", as_index=False)["emissions_kgco2e"]
        .sum()
        .sort_values("emissions_kgco2e", ascending=False)
    )

    st.write(f"### {tf('Emissions par equipe', 'Emissions by team')}")
    team_chart = alt.Chart(by_team).mark_bar().encode(x=alt.X("team:N", title=tf("Equipe", "Team")), y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"), tooltip=["team", "emissions_kgco2e"]).properties(height=320)
    st.altair_chart(_style_chart(team_chart), use_container_width=True)
    _render_share_bar(by_team, "team", tf("Equipe", "Team"))

    st.write(f"### {tf('Emissions par poste', 'Emissions by category')}")
    poste_chart = alt.Chart(by_poste).mark_bar().encode(x=alt.X("poste:N", title=tf("Poste", "Category")), y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"), tooltip=["poste", "emissions_kgco2e"]).properties(height=320)
    st.altair_chart(_style_chart(poste_chart), use_container_width=True)
    _render_share_bar(by_poste, "poste", tf("Poste", "Category"))

    st.write(f"### {tf('Incertitudes', 'Uncertainties')}")
    unc_by_poste = (
        filtered.groupby("poste", as_index=False)[["emissions_kgco2e", "uncertainty_kgco2e"]]
        .sum()
        .sort_values("emissions_kgco2e", ascending=False)
    )
    if not unc_by_poste.empty:
        unc_by_poste["emissions_low_kgco2e"] = (unc_by_poste["emissions_kgco2e"] - unc_by_poste["uncertainty_kgco2e"]).clip(lower=0.0)
        unc_by_poste["emissions_high_kgco2e"] = unc_by_poste["emissions_kgco2e"] + unc_by_poste["uncertainty_kgco2e"]
        error = (
            alt.Chart(unc_by_poste)
            .mark_errorbar(ticks=True)
            .encode(
                x=alt.X("poste:N", title=tf("Poste", "Category")),
                y=alt.Y("emissions_low_kgco2e:Q", title="kgCO2e"),
                y2="emissions_high_kgco2e:Q",
                tooltip=[
                    alt.Tooltip("poste:N", title=tf("Poste", "Category")),
                    alt.Tooltip("emissions_kgco2e:Q", title=tf("Emission centrale", "Central emission"), format=".2f"),
                    alt.Tooltip("uncertainty_kgco2e:Q", title=tf("Incertitude", "Uncertainty"), format=".2f"),
                ],
            )
        )
        points = alt.Chart(unc_by_poste).mark_circle(size=80).encode(
            x=alt.X("poste:N", title=tf("Poste", "Category")),
            y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"),
        )
        st.altair_chart(_style_chart((error + points).properties(height=300)), use_container_width=True)

    unc_year = (
        filtered.groupby("meta_year", as_index=False)[["emissions_kgco2e", "uncertainty_kgco2e"]]
        .sum()
        .sort_values("meta_year")
    )
    if not unc_year.empty:
        unc_year["low"] = (unc_year["emissions_kgco2e"] - unc_year["uncertainty_kgco2e"]).clip(lower=0.0)
        unc_year["high"] = unc_year["emissions_kgco2e"] + unc_year["uncertainty_kgco2e"]
        band = alt.Chart(unc_year).mark_area(opacity=0.25).encode(
            x=alt.X("meta_year:N", title=tf("Annee", "Year"), sort=YEAR_OPTIONS_ASC),
            y=alt.Y("low:Q", title="kgCO2e"),
            y2="high:Q",
        )
        line = alt.Chart(unc_year).mark_line(point=True).encode(
            x=alt.X("meta_year:N", title=tf("Annee", "Year"), sort=YEAR_OPTIONS_ASC),
            y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"),
        )
        st.altair_chart(_style_chart((band + line).properties(height=280)), use_container_width=True)

    if not by_poste_type.empty:
        st.write(f"### {tf('Emissions par type de poste', 'Emissions by position type')}")
        poste_type_chart = (
            alt.Chart(by_poste_type)
            .mark_bar()
            .encode(
                x=alt.X("meta_poste_type:N", title=tf("Type de poste", "Position type")),
                y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"),
                tooltip=["meta_poste_type", "emissions_kgco2e"],
            )
            .properties(height=320)
        )
        st.altair_chart(_style_chart(poste_type_chart), use_container_width=True)
        _render_share_bar(by_poste_type, "meta_poste_type", tf("Type de poste", "Position type"))

    if not by_employment.empty:
        st.write(f"### {tf('Emissions par statut (Permanent / Contractuel)', 'Emissions by employment status (Permanent / Contractual)')}")
        employment_chart = (
            alt.Chart(by_employment)
            .mark_bar()
            .encode(
                x=alt.X("meta_employment_type:N", title=tf("Statut", "Status")),
                y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"),
                tooltip=["meta_employment_type", "emissions_kgco2e"],
            )
            .properties(height=320)
        )
        st.altair_chart(_style_chart(employment_chart), use_container_width=True)
        _render_share_bar(by_employment, "meta_employment_type", tf("Statut", "Status"))

    st.write(f"### {tf('Detail', 'Detail')}")
    _render_themed_dataframe(filtered, use_container_width=True, light_use_table=False)

    st.write(f"### {tf('Synthese', 'Summary')}")
    synth_df = build_synthese(filtered)
    _render_themed_dataframe(synth_df, use_container_width=True)


def _render_factors_catalog_overview() -> None:
    st.write(f"### {tf('Catalogue des facteurs', 'Factors catalog')}")
    st.caption(
        tf(
            "Visualisation en lecture seule des fichiers Excel de facteurs.",
            "Read-only preview of factor Excel files.",
        )
    )

    catalogs = [
        (tf("Catalogue general", "Main catalog"), CATALOG_PATH),
        (tf("Catalogue plateforme", "Platform catalog"), PLATFORM_CATALOG_PATH),
    ]
    for title, path in catalogs:
        st.write(f"#### {title}")
        st.caption(str(path))
        if not path.exists():
            st.warning(tf("Fichier introuvable.", "File not found."))
            continue
        try:
            workbook = pd.read_excel(path, sheet_name=None)
        except Exception as exc:
            st.error(tf(f"Lecture impossible: {exc}", f"Could not read file: {exc}"))
            continue

        if not workbook:
            st.info(tf("Aucune feuille dans ce fichier.", "No sheet in this file."))
            continue

        sheet_names = list(workbook.keys())
        tabs = st.tabs(sheet_names)
        for tab, sheet_name in zip(tabs, sheet_names):
            with tab:
                sheet_df = workbook.get(sheet_name, pd.DataFrame())
                if sheet_df.empty:
                    st.info(tf("Feuille vide.", "Empty sheet."))
                else:
                    _render_themed_dataframe(sheet_df, use_container_width=True, light_use_table=False)


def _render_home_page() -> None:
    def _queue_home_transition(target_view: str) -> None:
        st.session_state.home_transition_target_view = target_view
        st.rerun()

    pending_target = str(st.session_state.get("home_transition_target_view", "")).strip()
    if pending_target in {"formulaire", "labo_overview"}:
        st.markdown(
            """
<style>
.home-exit-wrap {
  animation: home-slide-left 320ms ease-in forwards;
  transform-origin: center center;
}
@keyframes home-slide-left {
  from { opacity: 1; transform: translateX(0px) scale(1); filter: blur(0px); }
  to   { opacity: 0; transform: translateX(-72px) scale(0.985); filter: blur(1px); }
}
</style>
<div class="home-exit-wrap">
  <h3 style="margin-top:0;">Transition...</h3>
</div>
            """,
            unsafe_allow_html=True,
        )
        time.sleep(0.32)
        st.session_state.top_view_sidebar = pending_target
        st.session_state.app_stage = "form"
        st.session_state.home_transition_target_view = ""
        st.rerun()

    st.markdown(f"## {tf('Bienvenue dans le Carbonometre', 'Welcome to Carbonometre')}")
    st.caption(
        tf(
            "Choisissez votre mode de demarrage.",
            "Choose how you want to start.",
        )
    )
    st.markdown(f"### {tf('Qui etes vous ?', 'Who are you?')}")
    home_mode_for_identity = st.session_state.get("home_new_mode_input", "bilan_personnel")
    is_anonymous = st.session_state.get("is_anonymous", True)
    st.text_input(tf("Nom / identifiant", "Name / identifier"), key="person_label_input", disabled=is_anonymous)
    st.checkbox(tf("Anonyme", "Anonymous"), value=is_anonymous, key="is_anonymous", disabled=home_mode_for_identity == "plateforme")

    row_top = st.columns(2)
    with row_top[0]:
        st.markdown("### Quickcheck")
        st.caption(tf("Saisie rapide sans ouvrir un dossier sauvegardable.", "Quick input without opening a savable form."))
        if st.button("Quickcheck", key="home_quickcheck"):
            st.session_state.mode_input = "item_unique"
            _queue_home_transition("formulaire")
    with row_top[1]:
        st.markdown(f"### {tf('Visualiser les donnees du labo', 'View lab data')}")
        st.caption(tf("Acces direct a la vue consolidee du labo.", "Direct access to consolidated lab overview."))
        if st.button(tf("Visualiser les donnees du labo", "View lab data"), key="home_open_labo"):
            _queue_home_transition("labo_overview")

    col_new, col_continue = st.columns(2)
    with col_new:
        st.markdown(f"### {tf('Start new form', 'Start new form')}")
        savable_modes = ["bilan_personnel", "bilan_projet", "plateforme"]
        if st.session_state.get("home_new_mode_input", "") not in savable_modes:
            st.session_state.home_new_mode_input = savable_modes[0]
        new_mode = st.selectbox(tf("Mode", "Mode"), savable_modes, key="home_new_mode_input", format_func=_mode_label)
        if st.session_state.get("year_input", "") not in YEAR_OPTIONS:
            st.session_state.year_input = YEAR_OPTIONS[0]
        st.selectbox(tf("Annee", "Year"), options=YEAR_OPTIONS, key="year_input")
        if new_mode == "bilan_projet":
            st.text_input(tf("Nom du projet", "Project name"), key="project_name_input")
        if new_mode == "plateforme":
            st.session_state.is_anonymous = False
            st.session_state.team_code_input = "PLATEFORMES"
            if st.session_state.get("platform_name_input", "") not in PLATFORM_OPTIONS:
                st.session_state.platform_name_input = PLATFORM_OPTIONS[0]
            st.selectbox(tf("Plateforme", "Platform"), options=PLATFORM_OPTIONS, key="platform_name_input")
            st.caption(tf("Equipe: PLATEFORMES (imposee pour ce mode)", "Team: PLATEFORMES (enforced for this mode)"))
        else:
            st.selectbox(tf("Equipe (optionnel)", "Team (optional)"), options=list(TEAM_OPTIONS.keys()), format_func=lambda k: TEAM_OPTIONS[k], key="team_code_input")
        if st.button(tf("Start new form", "Start new form"), key="home_start_new_form"):
            st.session_state.mode_input = new_mode
            _queue_home_transition("formulaire")

    with col_continue:
        st.markdown(f"### {tf('Continue existing form', 'Continue existing form')}")
        saved_forms = _list_saved_forms()
        selected_path = None
        if saved_forms:
            years_with_forms = sorted(
                {
                    rel.parts[0]
                    for path in saved_forms
                    for rel in [path.relative_to(STORAGE_ROOT)]
                    if rel.parts
                },
                key=_year_sort_key,
                reverse=True,
            )
            reload_year_options = sorted(set(YEAR_OPTIONS).union(years_with_forms), key=_year_sort_key, reverse=True)
            current_year = str(datetime.now(timezone.utc).year)
            if st.session_state.get("home_reload_year_input", "") not in reload_year_options:
                st.session_state.home_reload_year_input = current_year if current_year in reload_year_options else reload_year_options[0]
            selected_reload_year = st.selectbox(
                tf("Annee des dossiers locaux", "Local forms year"),
                options=reload_year_options,
                key="home_reload_year_input",
            )
            year_forms = [p for p in saved_forms if _year_from_path(p) == selected_reload_year]
            team_reload_options = sorted({_team_from_path(p) for p in year_forms if _team_from_path(p)})
            if team_reload_options:
                if st.session_state.get("home_reload_team_input", "") not in team_reload_options:
                    st.session_state.home_reload_team_input = team_reload_options[0]
                selected_reload_team = st.selectbox(
                    tf("Equipe des dossiers locaux", "Local forms team"),
                    options=team_reload_options,
                    key="home_reload_team_input",
                    format_func=lambda t: TEAM_OPTIONS.get(t, t),
                )
                team_forms = [p for p in year_forms if _team_from_path(p) == selected_reload_team]
                if team_forms:
                    selected_path = st.selectbox(
                        tf("Dossiers locaux", "Local forms"),
                        options=team_forms,
                        key="home_reload_form_input",
                        format_func=lambda p: str(p.relative_to(STORAGE_ROOT / selected_reload_year / selected_reload_team)),
                    )
        else:
            st.caption(tf("Aucun dossier local sauvegarde pour l'instant.", "No local saved form yet."))
        if st.button(tf("Continue existing form", "Continue existing form"), key="home_continue_existing", disabled=selected_path is None):
            if selected_path is not None:
                file_bytes = selected_path.read_bytes()
                meta, imported_entries = import_excel_entries(file_bytes)
                st.session_state.pending_loaded_meta = meta
                st.session_state.pending_loaded_entries = imported_entries
                st.session_state.pending_loaded_path = str(selected_path)
                _queue_home_transition("formulaire")



def _apply_theme_css() -> None:
    dark_theme = LAB_THEME.get("dark", {})
    light_theme = LAB_THEME.get("light", {})

    if st.session_state.ui_theme == "darkmode":
        css = """
<style>
.stApp{
  background: linear-gradient(180deg, @@APP_BG_START@@ 0%, @@APP_BG_END@@ 100%);
  color: @@TEXT@@;
}
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stApp p, .stApp label, .stApp div, .stApp span {
  color: @@TEXT@@;
}
section[data-testid="stSidebar"]{
  background: linear-gradient(180deg, @@SIDEBAR_BG_START@@ 0%, @@SIDEBAR_BG_END@@ 100%);
}
div[data-testid="stMetric"]{
  background: @@METRIC_BG@@;
  border: 1px solid @@METRIC_BORDER@@;
  border-radius: 10px;
  padding: 8px 10px;
}
.stButton button{
  background: @@BUTTON_BG@@;
  border: 1px solid @@BUTTON_BORDER@@;
  color: @@BUTTON_TEXT@@;
}
.stButton button:hover{
  background: @@BUTTON_BG_HOVER@@;
  border-color: @@BUTTON_BORDER_HOVER@@;
}
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stTextArea"] textarea{
  background: @@INPUT_BG@@ !important;
  color: @@INPUT_TEXT@@ !important;
  border-color: @@INPUT_BORDER@@ !important;
}
</style>
"""
        css = (
            css.replace("@@APP_BG_START@@", dark_theme.get("app_bg_start", "#05070d"))
            .replace("@@APP_BG_END@@", dark_theme.get("app_bg_end", "#0b1220"))
            .replace("@@TEXT@@", dark_theme.get("text", "#dbe7ff"))
            .replace("@@SIDEBAR_BG_START@@", dark_theme.get("sidebar_bg_start", "#02040a"))
            .replace("@@SIDEBAR_BG_END@@", dark_theme.get("sidebar_bg_end", "#0a1325"))
            .replace("@@METRIC_BG@@", dark_theme.get("metric_bg", "#0e1a31"))
            .replace("@@METRIC_BORDER@@", dark_theme.get("metric_border", "#1f355f"))
            .replace("@@BUTTON_BG@@", dark_theme.get("button_bg", "#112a57"))
            .replace("@@BUTTON_BORDER@@", dark_theme.get("button_border", "#294b8a"))
            .replace("@@BUTTON_TEXT@@", dark_theme.get("button_text", "#e8f0ff"))
            .replace("@@BUTTON_BG_HOVER@@", dark_theme.get("button_bg_hover", "#173872"))
            .replace("@@BUTTON_BORDER_HOVER@@", dark_theme.get("button_border_hover", "#3a67b8"))
            .replace("@@INPUT_BG@@", dark_theme.get("input_bg", "#0d1a33"))
            .replace("@@INPUT_TEXT@@", dark_theme.get("input_text", "#e8f0ff"))
            .replace("@@INPUT_BORDER@@", dark_theme.get("input_border", "#2b4574"))
        )
        st.markdown(css, unsafe_allow_html=True)
        return

    css = """
<style>
.stApp{
  background: linear-gradient(180deg, @@APP_BG_START@@ 0%, @@APP_BG_END@@ 100%);
  color: @@TEXT@@;
}
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stApp p, .stApp label, .stApp div, .stApp span {
  color: @@TEXT@@;
}
section[data-testid="stSidebar"]{
  background: linear-gradient(180deg, @@SIDEBAR_BG_START@@ 0%, @@SIDEBAR_BG_END@@ 100%);
}
div[data-testid="stMetric"]{
  background: @@METRIC_BG@@;
  border: 1px solid @@METRIC_BORDER@@;
  border-radius: 10px;
  padding: 8px 10px;
}
.stButton button{
  background: @@BUTTON_BG@@;
  border: 1px solid @@BUTTON_BORDER@@;
  color: @@BUTTON_TEXT@@;
}
.stButton button:hover{
  background: @@BUTTON_BG_HOVER@@;
  border-color: @@BUTTON_BORDER_HOVER@@;
}
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stTextArea"] textarea{
  background: @@INPUT_BG@@ !important;
  color: @@INPUT_TEXT@@ !important;
  border-color: @@INPUT_BORDER@@ !important;
}
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"]{
  background: #ffffff !important;
  color: #102214 !important;
  border: 1px solid #d8e6d0 !important;
}
div[data-testid="stDataFrame"] [role="grid"],
div[data-testid="stDataEditor"] [role="grid"]{
  background: #ffffff !important;
  color: #102214 !important;
}
div[data-testid="stDataFrame"] [role="columnheader"],
div[data-testid="stDataEditor"] [role="columnheader"]{
  background: #f2f8ee !important;
  color: #102214 !important;
}
div[data-testid="stDataFrame"] [role="rowheader"],
div[data-testid="stDataEditor"] [role="rowheader"]{
  background: #f2f8ee !important;
  color: #102214 !important;
}
div[data-testid="stDataFrame"] [role="gridcell"],
div[data-testid="stDataEditor"] [role="gridcell"]{
  background: #ffffff !important;
  color: #102214 !important;
}
div[data-testid="stDataFrame"] canvas,
div[data-testid="stDataEditor"] canvas{
  background: #ffffff !important;
}
div[data-testid="stDataFrame"] div[data-testid="stMarkdownContainer"],
div[data-testid="stDataEditor"] div[data-testid="stMarkdownContainer"]{
  color: #102214 !important;
}
div[data-testid="stNumberInput"] button{
  background: #f1f5f9 !important;
  color: #102214 !important;
  border-color: #cbd5e1 !important;
}
div[data-testid="stNumberInput"] input{
  background: #ffffff !important;
  color: #102214 !important;
  border-color: #9bd27e !important;
}
/* Select dropdown menu (BaseWeb portal) */
div[role="listbox"]{
  background: #ffffff !important;
  color: #102214 !important;
  border: 1px solid #cbd5e1 !important;
}
li[role="option"]{
  background: #ffffff !important;
  color: #102214 !important;
}
li[role="option"][aria-selected="true"]{
  background: #eef6ff !important;
  color: #0f2b46 !important;
}
div[data-testid="stFileUploaderDropzone"]{
  background: @@FILE_DROP_BG@@ !important;
  border: 1px dashed @@FILE_DROP_BORDER@@ !important;
}
div[data-testid="stFileUploaderDropzone"] *{
  color: @@FILE_DROP_TEXT@@ !important;
}
/* Vega/Altair action buttons visible on hover */
.vega-actions,
.vega-actions a{
  background: #ffffff !important;
  color: #102214 !important;
  border-color: #cbd5e1 !important;
}
div[data-testid="stVegaLiteChart"] button,
div[data-testid="stVegaLiteChart"] summary{
  background: #ffffff !important;
  color: #102214 !important;
  border: 1px solid #cbd5e1 !important;
}
</style>
"""
    css = (
        css.replace("@@APP_BG_START@@", light_theme.get("app_bg_start", "#ffffff"))
        .replace("@@APP_BG_END@@", light_theme.get("app_bg_end", "#f7fff4"))
        .replace("@@TEXT@@", light_theme.get("text", "#102214"))
        .replace("@@SIDEBAR_BG_START@@", light_theme.get("sidebar_bg_start", "#f2ffe9"))
        .replace("@@SIDEBAR_BG_END@@", light_theme.get("sidebar_bg_end", "#e6f9d8"))
        .replace("@@METRIC_BG@@", light_theme.get("metric_bg", "#ffffff"))
        .replace("@@METRIC_BORDER@@", light_theme.get("metric_border", "#9bd27e"))
        .replace("@@BUTTON_BG@@", light_theme.get("button_bg", "#6dbf3f"))
        .replace("@@BUTTON_BORDER@@", light_theme.get("button_border", "#57a631"))
        .replace("@@BUTTON_TEXT@@", light_theme.get("button_text", "#ffffff"))
        .replace("@@BUTTON_BG_HOVER@@", light_theme.get("button_bg_hover", "#58a533"))
        .replace("@@BUTTON_BORDER_HOVER@@", light_theme.get("button_border_hover", "#4b8f2d"))
        .replace("@@INPUT_BG@@", light_theme.get("input_bg", "#ffffff"))
        .replace("@@INPUT_TEXT@@", light_theme.get("input_text", "#102214"))
        .replace("@@INPUT_BORDER@@", light_theme.get("input_border", "#9bd27e"))
        .replace("@@FILE_DROP_BG@@", light_theme.get("file_drop_bg", "#ffffff"))
        .replace("@@FILE_DROP_BORDER@@", light_theme.get("file_drop_border", "#7fbe5d"))
        .replace("@@FILE_DROP_TEXT@@", light_theme.get("file_drop_text", "#102214"))
    )
    st.markdown(css, unsafe_allow_html=True)


_ensure_storage_tree()
_consume_pending_loaded_state()
_apply_theme_css()

mode = _normalize_mode_value(st.session_state.get("mode_input", "item_unique"))
st.session_state.mode_input = mode
pending_mode_ui_reset = _normalize_mode_value(str(st.session_state.get("mode_ui_reset_to", "")))
raw_mode_ui_reset = str(st.session_state.get("mode_ui_reset_to", "")).strip()
if raw_mode_ui_reset and pending_mode_ui_reset in MODE_OPTIONS:
    st.session_state.mode_input_ui = pending_mode_ui_reset
    st.session_state.mode_ui_reset_to = ""
if st.session_state.get("mode_input_ui", "") not in MODE_OPTIONS:
    st.session_state.mode_input_ui = mode
if mode == "plateforme":
    st.session_state.is_anonymous = False
    st.session_state.team_code_input = "PLATEFORMES"
dossier_type = DOSSIER_TYPE_BY_MODE.get(mode, "item")
dossier_id = st.session_state.get("dossier_id", "")
project_name = st.session_state.get("project_name_input", "")
is_anonymous = st.session_state.get("is_anonymous", True)
person_label = st.session_state.get("person_label_input", "")
team_code = st.session_state.get("team_code_input", "")
year_value = st.session_state.get("year_input", "2026")
poste_type = st.session_state.get("poste_type_input", "")

if st.session_state.get("app_stage", "home") == "home":
    _render_home_page()
    st.stop()

st.markdown(f"### {tf('Mode actif', 'Active mode')}: {_mode_label(mode)}")

with st.sidebar:
    st.header(tf("Navigation", "Navigation"))
    top_view = st.radio(
        tf("Vue", "View"),
        ["formulaire", "labo_overview", "factors_catalog"],
        key="top_view_sidebar",
        format_func=lambda v: (
            tf("Formulaire", "Form")
            if v == "formulaire"
            else (
                tf("Vue d'ensemble du labo", "Lab overview")
                if v == "labo_overview"
                else tf("Catalogue facteurs", "Factors catalog")
            )
        ),
    )
    if st.session_state.loaded_notice:
        st.success(st.session_state.loaded_notice)
        st.session_state.loaded_notice = ""
    if st.session_state.uncertainty_notice:
        st.warning(st.session_state.uncertainty_notice)
        st.session_state.uncertainty_notice = ""
    row_theme = st.columns([1.2, 1.0, 1.2], vertical_alignment="center")
    row_theme[0].markdown("<div style='text-align:right;'>lightmode ☀️</div>", unsafe_allow_html=True)
    with row_theme[1]:
        dark_on = st.toggle(
            "theme",
            value=st.session_state.ui_theme == "darkmode",
            key="theme_toggle",
            label_visibility="collapsed",
        )
    row_theme[2].markdown("<div style='text-align:left;'>🌙 darkmode</div>", unsafe_allow_html=True)
    st.session_state.ui_theme = "darkmode" if dark_on else "lightmode"
    row_lang = st.columns([1.2, 1.0, 1.2], vertical_alignment="center")
    row_lang[0].markdown("<div style='text-align:right;'>FR 🇫🇷</div>", unsafe_allow_html=True)
    with row_lang[1]:
        en_on = st.toggle(
            "lang",
            value=st.session_state.ui_lang == "EN",
            key="lang_toggle",
            label_visibility="collapsed",
        )
    row_lang[2].markdown("<div style='text-align:left;'>EN 🇬🇧</div>", unsafe_allow_html=True)
    st.session_state.ui_lang = "EN" if en_on else "FR"
    st.markdown(
        """
<style>
section[data-testid="stSidebar"] div[role="radiogroup"] label{
  border: 1px solid #64748b !important;
  border-radius: 8px !important;
  padding: 4px 8px !important;
  background: transparent !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"]{
  border: 2px solid #22c55e !important;
  background: rgba(34, 197, 94, 0.15) !important;
}
</style>
""",
        unsafe_allow_html=True,
    )

    if top_view == "formulaire":
        st.divider()
        st.header(tf("Dossier", "Form"))
        if st.session_state.get("mode_input_ui", "") not in MODE_OPTIONS:
            st.session_state.mode_input_ui = mode
        selected_mode_ui = st.selectbox(tf("Mode", "Mode"), MODE_OPTIONS, key="mode_input_ui", format_func=_mode_label)
        confirmed_mode = _normalize_mode_value(st.session_state.get("mode_input", "item_unique"))
        pending_mode_change = st.session_state.get("pending_mode_change")
        if selected_mode_ui != confirmed_mode:
            expected = {"from": confirmed_mode, "to": selected_mode_ui}
            if not isinstance(pending_mode_change, dict) or pending_mode_change != expected:
                st.session_state.pending_mode_change = expected
                pending_mode_change = expected
        if isinstance(pending_mode_change, dict):
            st.warning(
                tf(
                    f"Etes vous sur(e) de vouloir changer le mode vers {_mode_label(str(pending_mode_change.get('to', confirmed_mode)))} ?",
                    f"Are you sure you want to switch mode to {_mode_label(str(pending_mode_change.get('to', confirmed_mode)))}?",
                )
            )
            c_yes, c_no = st.columns(2)
            if c_yes.button(tf("oui", "yes"), key="confirm_mode_change_yes"):
                new_mode = _normalize_mode_value(str(pending_mode_change.get("to", confirmed_mode)))
                st.session_state.mode_input = new_mode
                st.session_state.mode_ui_reset_to = new_mode
                st.session_state.pending_mode_change = None
                st.rerun()
            if c_no.button(tf("non", "no"), key="confirm_mode_change_no"):
                st.session_state.mode_ui_reset_to = confirmed_mode
                st.session_state.pending_mode_change = None
                st.rerun()
            mode = confirmed_mode
        else:
            mode = confirmed_mode
        dossier_type = DOSSIER_TYPE_BY_MODE[mode]
        project_name = st.session_state.get("project_name_input", "")
        if mode in SAVE_MODE_SUFFIX:
            if st.session_state.get("year_input", "") not in YEAR_OPTIONS:
                st.session_state.year_input = YEAR_OPTIONS[0]
            year_value = st.selectbox(tf("Annee", "Year"), options=YEAR_OPTIONS, key="year_input")
        else:
            year_value = st.session_state.get("year_input", "2026")
        if mode == "bilan_projet":
            project_name = st.text_input(tf("Nom du projet", "Project name"), key="project_name_input")
        if mode == "plateforme":
            if st.session_state.get("platform_name_input", "") not in PLATFORM_OPTIONS:
                st.session_state.platform_name_input = PLATFORM_OPTIONS[0]
            st.selectbox(
                tf("Plateforme", "Platform"),
                options=PLATFORM_OPTIONS,
                key="platform_name_input",
            )
        if not st.session_state.dossier_id:
            default_id = datetime.now(timezone.utc).strftime("DOSSIER-%Y%m%d-%H%M%S")
            st.session_state.dossier_id = default_id
        dossier_id = st.session_state.dossier_id

        st.divider()
        st.subheader(tf("Identite", "Identity"))
        if mode == "plateforme":
            st.session_state.is_anonymous = False
        is_anonymous = st.session_state.get("is_anonymous", True)
        person_label = st.text_input(tf("Nom / identifiant", "Name / identifier"), key="person_label_input", disabled=is_anonymous)
        is_anonymous = st.checkbox(tf("Anonyme", "Anonymous"), value=is_anonymous, key="is_anonymous", disabled=mode == "plateforme")
        if mode == "plateforme":
            team_code = "PLATEFORMES"
            st.session_state.team_code_input = team_code
            st.caption(tf("Equipe: PLATEFORMES (imposee pour ce mode)", "Team: PLATEFORMES (enforced for this mode)"))
        else:
            team_code = st.selectbox(tf("Equipe (optionnel)", "Team (optional)"), options=list(TEAM_OPTIONS.keys()), format_func=lambda k: TEAM_OPTIONS[k], key="team_code_input")
        if mode == "plateforme":
            poste_options = PLATFORM_POSTE_TYPE_OPTIONS
        else:
            poste_options = ADMIN_GESTION_POSTE_TYPE_OPTIONS if team_code == "ADMIN_GESTION" else POSTE_TYPE_OPTIONS
        normalized_poste_type = _normalize_poste_type_for_overview(team_code, st.session_state.get("poste_type_input", ""))
        st.session_state.poste_type_input = normalized_poste_type
        if st.session_state.get("poste_type_input", "") not in poste_options:
            st.session_state.poste_type_input = poste_options[0] if poste_options else ""
        poste_type = st.selectbox(
            tf("Type de poste", "Position type"),
            options=poste_options,
            key="poste_type_input",
            format_func=_poste_type_label,
        )
        status_info = _employment_type(team_code, poste_type)
        if status_info:
            st.caption(tf(f"Statut: {status_info}", f"Status: {status_info}"))

        _auto_resolve_form_context(
            mode=mode,
            year_value=year_value,
            team_code=team_code,
            is_anonymous=is_anonymous,
            person_label=person_label,
            project_name=project_name,
        )

        pending_switch = st.session_state.get("pending_form_switch")
        if isinstance(pending_switch, dict):
            target_path = Path(str(pending_switch.get("target_path", "")))
            action_label = (
                tf("ouvrir le dossier existant", "open existing form")
                if bool(pending_switch.get("exists", False))
                else tf("creer un nouveau dossier", "create a new form")
            )
            st.warning(
                tf(
                    f"Changement de contexte detecte ({year_value} / {person_label or 'anonyme'}). Des modifications non sauvegardees existent. Voulez-vous {action_label} ?",
                    f"Context change detected ({year_value} / {person_label or 'anonymous'}). Unsaved changes exist. Do you want to {action_label}?",
                )
            )
            c_open, c_stay = st.columns(2)
            if c_open.button(tf("Continuer", "Continue"), key="confirm_form_switch_yes"):
                _execute_form_switch(
                    target_path=target_path,
                    identity=str(pending_switch.get("identity", "")),
                    owner_for_save=str(pending_switch.get("owner_for_save", "anonyme")),
                )
            if c_stay.button(tf("Rester ici", "Stay here"), key="confirm_form_switch_no"):
                st.session_state.snoozed_form_identity = str(pending_switch.get("identity", ""))
                st.session_state.pending_form_switch = None

        st.divider()
        st.subheader(tf("Sauvegarde locale", "Local save"))
        if mode in SAVE_MODE_SUFFIX:
            if st.session_state.loaded_form_path:
                st.caption(f"{tf('Fichier courant', 'Current file')}: {Path(st.session_state.loaded_form_path).name}")
            if st.button(tf("Sauvegarder le dossier localement", "Save form locally")):
                if mode == "plateforme" and is_anonymous:
                    st.error(tf("Mode plateforme: sauvegarde anonyme interdite.", "Platform mode: anonymous save is forbidden."))
                else:
                    owner_for_save = st.session_state.loaded_owner_label or person_label or "anonyme"
                    save_path = (
                        Path(st.session_state.loaded_form_path)
                        if st.session_state.loaded_form_path
                        else _default_save_path(mode=mode, year_value=year_value, team_code=team_code, person_label=owner_for_save, project_name=project_name)
                    )
                    # If this form was reopened from local storage, ask for explicit confirmation before overwrite.
                    if st.session_state.loaded_form_path:
                        st.session_state.confirm_overwrite_pending = True
                    else:
                        _save_local_form(
                            save_path=save_path,
                            year_value=year_value,
                            dossier_id=dossier_id,
                            dossier_type=dossier_type,
                            project_name=project_name,
                            platform_name=st.session_state.get("platform_name_input", ""),
                            is_anonymous=is_anonymous,
                            owner_for_save=owner_for_save,
                            poste_type=poste_type,
                            team_code=team_code,
                            entries=st.session_state.entries,
                        )
                        st.session_state.loaded_form_path = str(save_path)
                        st.session_state.loaded_owner_label = owner_for_save
                        st.session_state.form_baseline_signature = _entries_signature(st.session_state.entries)
                        st.session_state.active_form_identity = _compute_form_identity(
                            mode=mode,
                            year_value=year_value,
                            team_code=team_code,
                            owner_for_save=owner_for_save,
                            project_name=project_name,
                        )
                        st.session_state.snoozed_form_identity = ""
                        st.success(f"{tf('Sauvegarde ok', 'Saved')}: {save_path.relative_to(Path(__file__).resolve().parent)}")

            if st.session_state.confirm_overwrite_pending and st.session_state.loaded_form_path:
                owner_for_msg = st.session_state.loaded_owner_label or person_label or "anonyme"
                st.warning(
                    tf(
                        f"Etes vous vraiment {owner_for_msg} et voulez vous vraiment modifier ce dossier ?",
                        f"Are you really {owner_for_msg} and do you really want to modify this form?",
                    )
                )
                c_yes, c_no = st.columns(2)
                if c_yes.button(tf("oui", "yes"), key="confirm_overwrite_yes"):
                    owner_for_save = st.session_state.loaded_owner_label or person_label or "anonyme"
                    save_path = Path(st.session_state.loaded_form_path)
                    _save_local_form(
                        save_path=save_path,
                        year_value=year_value,
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        project_name=project_name,
                        platform_name=st.session_state.get("platform_name_input", ""),
                        is_anonymous=is_anonymous,
                        owner_for_save=owner_for_save,
                        poste_type=poste_type,
                        team_code=team_code,
                        entries=st.session_state.entries,
                    )
                    st.session_state.loaded_owner_label = owner_for_save
                    st.session_state.confirm_overwrite_pending = False
                    st.session_state.form_baseline_signature = _entries_signature(st.session_state.entries)
                    st.session_state.active_form_identity = _compute_form_identity(
                        mode=mode,
                        year_value=year_value,
                        team_code=team_code,
                        owner_for_save=owner_for_save,
                        project_name=project_name,
                    )
                    st.session_state.snoozed_form_identity = ""
                    st.success(f"{tf('Sauvegarde ok', 'Saved')}: {save_path.relative_to(Path(__file__).resolve().parent)}")
                if c_no.button(tf("non", "no"), key="confirm_overwrite_no"):
                    st.session_state.confirm_overwrite_pending = False
        else:
            st.caption(tf("Sauvegarde locale disponible uniquement pour les modes sauvegardables.", "Local save is available only for savable modes."))

        st.divider()
        st.subheader(tf("Reprendre un dossier", "Reload a form"))
        saved_forms = _list_saved_forms()
        if saved_forms:
            years_with_forms = sorted(
                {
                    rel.parts[0]
                    for path in saved_forms
                    for rel in [path.relative_to(STORAGE_ROOT)]
                    if rel.parts
                },
                key=_year_sort_key,
                reverse=True,
            )
            reload_year_options = sorted(set(YEAR_OPTIONS).union(years_with_forms), key=_year_sort_key, reverse=True)
            current_year = str(datetime.now(timezone.utc).year)
            if st.session_state.get("reload_year_input", "") not in reload_year_options:
                st.session_state.reload_year_input = current_year if current_year in reload_year_options else reload_year_options[0]

            selected_reload_year = st.selectbox(
                tf("Annee des dossiers locaux", "Local forms year"),
                options=reload_year_options,
                key="reload_year_input",
            )
            year_forms = [p for p in saved_forms if _year_from_path(p) == selected_reload_year]
            if year_forms:
                team_reload_options = sorted({_team_from_path(p) for p in year_forms if _team_from_path(p)})
                if team_reload_options:
                    if st.session_state.get("reload_team_input", "") not in team_reload_options:
                        st.session_state.reload_team_input = team_reload_options[0]
                    selected_reload_team = st.selectbox(
                        tf("Equipe des dossiers locaux", "Local forms team"),
                        options=team_reload_options,
                        key="reload_team_input",
                        format_func=lambda t: TEAM_OPTIONS.get(t, t),
                    )
                    team_forms = [p for p in year_forms if _team_from_path(p) == selected_reload_team]
                    if team_forms:
                        selected_path = st.selectbox(
                            tf("Dossiers locaux", "Local forms"),
                            options=team_forms,
                            format_func=lambda p: str(p.relative_to(STORAGE_ROOT / selected_reload_year / selected_reload_team)),
                        )
                        if st.button(tf("Ouvrir dossier local", "Open local form")):
                            file_bytes = selected_path.read_bytes()
                            meta, imported_entries = import_excel_entries(file_bytes)
                            st.session_state.pending_loaded_meta = meta
                            st.session_state.pending_loaded_entries = imported_entries
                            st.session_state.pending_loaded_path = str(selected_path)
                            st.rerun()
                    else:
                        st.caption(tf("Aucun dossier local pour cette equipe.", "No local saved form for this team."))
                else:
                    st.caption(tf("Aucune equipe disponible pour cette annee.", "No team found for this year."))
            else:
                st.caption(tf("Aucun dossier local pour cette annee.", "No local saved form for this year."))
        else:
            st.caption(tf("Aucun dossier local sauvegarde pour l'instant.", "No local saved form yet."))

        st.divider()
        imported = st.file_uploader(tf("Reprendre un dossier (Excel exporte)", "Load a form (exported Excel)"), type=["xlsx"])
        if imported is not None:
            meta, imported_entries = import_excel_entries(imported.read())
            st.session_state.pending_loaded_meta = meta
            st.session_state.pending_loaded_entries = imported_entries
            st.session_state.pending_loaded_path = ""
            st.rerun()

    st.divider()
    if st.button(tf("Visualiser catalogues facteurs", "View factors catalogs"), key="sidebar_open_factors"):
        st.session_state.top_view_sidebar = "factors_catalog"
        st.rerun()
    st.divider()
    if st.button(tf("Retour a la page d'accueil", "Back to home page"), key="sidebar_back_home"):
        st.session_state.app_stage = "home"
        st.rerun()

if top_view == "labo_overview":
    _render_labo_overview()
    st.stop()
if top_view == "factors_catalog":
    _render_factors_catalog_overview()
    st.stop()

if mode == "plateforme":
    poste = "plateforme"
    st.caption(tf("Mode plateforme: formulaire specialise.", "Platform mode: dedicated form."))
else:
    poste = st.selectbox(
        tf("Poste", "Category"),
        POSTE_OPTIONS,
        format_func=_poste_label,
        key="poste_input",
    )

col1, col2 = st.columns([2, 1])
with col1:
    st.write(f"### {tf('Saisie', 'Input')}")
    new_row = None
    platform_rows_added = False
    entries = st.session_state.entries
    editing_idx = st.session_state.get("editing_row_idx")
    editing_row = None
    if isinstance(editing_idx, int) and 0 <= editing_idx < len(entries):
        editing_row = entries[editing_idx]

    if editing_row is not None:
        st.info(tf("Edition de ligne en cours", "Row editing in progress"))
        old = editing_row
        if old.get("poste") == "achats":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", old.get("item_label", ""))), key=f"edit_ref_achat_{editing_idx}")
            achat_categories = DEFAULT_FACTORS["achats_category_factors"]
            cur_cat = old.get("item_label", "")
            default_cat = cur_cat if cur_cat in achat_categories else list(achat_categories.keys())[0]
            e_cat = st.selectbox(
                tf("Type d'achat", "Purchase type"),
                options=list(achat_categories.keys()),
                index=list(achat_categories.keys()).index(default_cat),
                key=f"edit_achat_cat_{editing_idx}",
            )
            e_amount = st.number_input(
                tf("Montant (€)", "Amount (€)"),
                min_value=0.0,
                value=float(old.get("amount_eur", 0.0) or 0.0),
                step=100.0,
                key=f"edit_achat_amount_{editing_idx}",
            )
            e_default_factor = float(achat_categories.get(e_cat, achat_categories.get(default_cat, 0.3)))
            e_unc_pct = _factor_uncertainty_pct("achats_category_factors", e_cat)
            e_use_external_state = bool(st.session_state.get(f"edit_achat_{editing_idx}_use_external", bool(old.get("uses_external_emissions", False))))
            _render_default_factor_with_source(
                tf("Facteur (kgCO2e / €)", "Factor (kgCO2e / €)"),
                e_default_factor,
                e_unc_pct,
                _factor_reference("achats_category_factors", e_cat),
                ignored=e_use_external_state,
                key=f"edit_achat_factor_display_{editing_idx}",
            )
            e_use_external, e_external_total, e_external_unc_raw = _external_emissions_input(
                tf("Chiffre fabricant", "Manufacturer figure"),
                f"edit_achat_{editing_idx}",
                default_checked=bool(old.get("uses_external_emissions", False)),
                default_value=float(old.get("external_emissions_kgco2e", old.get("emissions_kgco2e", 0.0)) or 0.0),
            )
            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_achat_{editing_idx}"):
                updated = add_achat(
                    old.get("dossier_id", dossier_id),
                    old.get("dossier_type", dossier_type),
                    bool(old.get("is_anonymous", is_anonymous)),
                    old.get("team_code", team_code),
                    old.get("person_label", person_label),
                    _clean(e_cat),
                    e_amount,
                    e_default_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_external_emissions(updated, e_use_external, e_external_total)
                _apply_factor_uncertainty(updated, e_unc_pct)
                missing_unc = _apply_external_uncertainty(updated, e_use_external, e_external_unc_raw)
                _set_uncertainty_notice_if_missing(missing_unc)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_achat_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        elif old.get("poste") == "domicile_travail":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", "")), key=f"edit_ref_dom_{editing_idx}")
            e_mode = st.selectbox(
                tf("Mode de transport", "Transport mode"),
                list(DEFAULT_FACTORS["domicile_mode_factors"].keys()),
                index=list(DEFAULT_FACTORS["domicile_mode_factors"].keys()).index(old.get("transport_mode", "car"))
                if old.get("transport_mode", "car") in DEFAULT_FACTORS["domicile_mode_factors"]
                else 0,
                key=f"edit_dom_mode_{editing_idx}",
            )
            e_dist = st.number_input(tf("Distance aller simple (km)", "One-way distance (km)"), min_value=0.0, value=float(old.get("distance_one_way_km", 0.0) or 0.0), step=1.0, key=f"edit_dom_dist_{editing_idx}")
            edit_days_key = f"edit_dom_days_year_{editing_idx}"
            row = st.columns([1.8, 1.0, 1.0, 1.0], vertical_alignment="center")
            with row[0]:
                e_days_year = st.number_input(
                    tf("Jours/an", "Days/year"),
                    min_value=0.0,
                    max_value=366.0,
                    value=float(_domicile_days_per_year_from_row(old)),
                    step=1.0,
                    key=edit_days_key,
                )
            for col, (pid, label, days_value) in zip(row[1:], DOMICILE_DAYS_PRESETS):
                _render_domicile_days_preset_cell(
                    col=col,
                    target_days_key=edit_days_key,
                    preset_id=pid,
                    label=label,
                    days_value=days_value,
                    button_key_prefix=f"edit_dom_days_preset_{editing_idx}",
                )
            e_rt = st.checkbox(tf("Aller-retour", "Round trip"), value=bool(old.get("round_trip", True)), key=f"edit_dom_rt_{editing_idx}")
            e_default_factor = float(DEFAULT_FACTORS["domicile_mode_factors"].get(e_mode, DEFAULT_FACTORS["domicile_mode_factors"]["other"]))
            e_unc_pct = _factor_uncertainty_pct("domicile_mode_factors", e_mode)
            e_use_external_state = bool(st.session_state.get(f"edit_dom_{editing_idx}_use_external", bool(old.get("uses_external_emissions", False))))
            _render_default_factor_with_source(
                tf("Facteur (kgCO2e / km)", "Factor (kgCO2e / km)"),
                e_default_factor,
                e_unc_pct,
                _factor_reference("domicile_mode_factors", e_mode),
                ignored=e_use_external_state,
                key=f"edit_dom_factor_display_{editing_idx}",
            )
            e_use_external, e_external_total, e_external_unc_raw = _external_emissions_input(
                tf("Declaration compagnie de transport", "Transport company declaration"),
                f"edit_dom_{editing_idx}",
                default_checked=bool(old.get("uses_external_emissions", False)),
                default_value=float(old.get("external_emissions_kgco2e", old.get("emissions_kgco2e", 0.0)) or 0.0),
            )
            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_dom_{editing_idx}"):
                updated = add_domicile(
                    old.get("dossier_id", dossier_id),
                    old.get("dossier_type", dossier_type),
                    bool(old.get("is_anonymous", is_anonymous)),
                    old.get("team_code", team_code),
                    old.get("person_label", person_label),
                    e_mode,
                    e_dist,
                    e_days_year,
                    e_rt,
                    e_default_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_external_emissions(updated, e_use_external, e_external_total)
                _apply_factor_uncertainty(updated, e_unc_pct)
                missing_unc = _apply_external_uncertainty(updated, e_use_external, e_external_unc_raw)
                _set_uncertainty_notice_if_missing(missing_unc)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_dom_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        elif old.get("poste") == "campagnes_terrain":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", old.get("campaign_label", ""))), key=f"edit_ref_camp_{editing_idx}")
            e_label = st.text_input(tf("Nom campagne", "Campaign name"), value=str(old.get("campaign_label", "")), key=f"edit_camp_label_{editing_idx}")
            e_mode = st.selectbox(
                tf("Mode de transport", "Transport mode"),
                list(DEFAULT_FACTORS["campagnes_mode_factors"].keys()),
                index=list(DEFAULT_FACTORS["campagnes_mode_factors"].keys()).index(old.get("segment_mode", "plane"))
                if old.get("segment_mode", "plane") in DEFAULT_FACTORS["campagnes_mode_factors"]
                else 0,
                key=f"edit_camp_mode_{editing_idx}",
            )
            e_dist = st.number_input(tf("Distance segment (km)", "Segment distance (km)"), min_value=0.0, value=float(old.get("distance_km", 0.0) or 0.0), step=10.0, key=f"edit_camp_dist_{editing_idx}")
            e_pass = st.number_input(tf("Nombre de personnes", "Number of people"), min_value=1.0, value=float(old.get("passengers_count", 1.0) or 1.0), step=1.0, key=f"edit_camp_pass_{editing_idx}")
            e_rt = st.checkbox(tf("Aller-retour", "Round trip"), value=bool(old.get("round_trip", True)), key=f"edit_camp_rt_{editing_idx}")
            e_default_factor = float(DEFAULT_FACTORS["campagnes_mode_factors"].get(e_mode, DEFAULT_FACTORS["campagnes_mode_factors"]["other"]))
            e_unc_pct = _factor_uncertainty_pct("campagnes_mode_factors", e_mode)
            e_use_external_state = bool(st.session_state.get(f"edit_camp_{editing_idx}_use_external", bool(old.get("uses_external_emissions", False))))
            _render_default_factor_with_source(
                tf("Facteur (kgCO2e / km / personne)", "Factor (kgCO2e / km / person)"),
                e_default_factor,
                e_unc_pct,
                _factor_reference("campagnes_mode_factors", e_mode),
                ignored=e_use_external_state,
                key=f"edit_camp_factor_display_{editing_idx}",
            )
            e_use_external, e_external_total, e_external_unc_raw = _external_emissions_input(
                tf("Declaration prestataire transport", "Transport provider declaration"),
                f"edit_camp_{editing_idx}",
                default_checked=bool(old.get("uses_external_emissions", False)),
                default_value=float(old.get("external_emissions_kgco2e", old.get("emissions_kgco2e", 0.0)) or 0.0),
            )
            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_camp_{editing_idx}"):
                updated = add_campagne(
                    old.get("dossier_id", dossier_id),
                    old.get("dossier_type", dossier_type),
                    bool(old.get("is_anonymous", is_anonymous)),
                    old.get("team_code", team_code),
                    old.get("person_label", person_label),
                    _clean(e_label),
                    e_mode,
                    e_dist,
                    e_pass,
                    e_rt,
                    e_default_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_external_emissions(updated, e_use_external, e_external_total)
                _apply_factor_uncertainty(updated, e_unc_pct)
                missing_unc = _apply_external_uncertainty(updated, e_use_external, e_external_unc_raw)
                _set_uncertainty_notice_if_missing(missing_unc)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_camp_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        elif old.get("poste") == "missions":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", old.get("mission_id", ""))), key=f"edit_ref_mis_{editing_idx}")
            e_mid = st.text_input("Mission ID", value=str(old.get("mission_id", "")), key=f"edit_mis_id_{editing_idx}")
            col_dep_e, col_arr_e = st.columns(2)
            with col_dep_e:
                e_dep_city = st.text_input(tf("Ville depart", "Departure city"), value=str(old.get("departure_city", "")), key=f"edit_mis_dep_city_{editing_idx}")
                e_dep_country = _country_input(
                    tf("Pays depart", "Departure country"),
                    f"edit_dep_country_{editing_idx}",
                    default_country=str(old.get("departure_country", "FR") or "FR"),
                    city_hint=e_dep_city,
                )
            with col_arr_e:
                e_arr_city = st.text_input(tf("Ville arrivee", "Arrival city"), value=str(old.get("arrival_city", "")), key=f"edit_mis_arr_city_{editing_idx}")
                e_arr_country = _country_input(
                    tf("Pays arrivee", "Arrival country"),
                    f"edit_arr_country_{editing_idx}",
                    default_country=str(old.get("arrival_country", "FR") or "FR"),
                    city_hint=e_arr_city,
                )
            cur_mode = str(old.get("t_type", "plane"))
            if cur_mode not in {"plane", "train", "car"}:
                cur_mode = "plane"
            e_mode = st.selectbox(tf("Transport principal", "Main transport"), ["plane", "train", "car"], index=["plane", "train", "car"].index(cur_mode), key=f"edit_mis_mode_{editing_idx}")
            e_rt = st.checkbox(tf("Aller-retour", "Round trip"), value=bool(old.get("round_trip", True)), key=f"edit_mis_rt_{editing_idx}")
            e_unc_pct = _factor_uncertainty_pct("missions_mode_factors", e_mode)
            e_use_external, e_external_total, e_external_unc_raw = _external_emissions_input(
                tf("Declaration compagnie de transport", "Transport company declaration"),
                f"edit_mis_{editing_idx}",
                default_checked=bool(old.get("uses_external_emissions", False)),
                default_value=float(old.get("external_emissions_kgco2e", old.get("emissions_kgco2e", 0.0)) or 0.0),
            )
            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_mis_{editing_idx}"):
                updated = compute_single_mission_with_moulinette(
                    dossier_id=old.get("dossier_id", dossier_id),
                    dossier_type=old.get("dossier_type", dossier_type),
                    is_anonymous=bool(old.get("is_anonymous", is_anonymous)),
                    team_code=old.get("team_code", team_code),
                    person_label=old.get("person_label", person_label),
                    mission_id=_clean(e_mid),
                    departure_city=_clean(e_dep_city),
                    departure_country=_clean(e_dep_country),
                    arrival_city=_clean(e_arr_city),
                    arrival_country=_clean(e_arr_country),
                    transport_mode=e_mode,
                    round_trip=e_rt,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_external_emissions(updated, e_use_external, e_external_total)
                _apply_factor_uncertainty(updated, e_unc_pct)
                missing_unc = _apply_external_uncertainty(updated, e_use_external, e_external_unc_raw)
                _set_uncertainty_notice_if_missing(missing_unc)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_mis_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        elif old.get("poste") == "heures_calcul":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", old.get("compute_label", ""))), key=f"edit_ref_comp_{editing_idx}")
            e_label = st.text_input(tf("Nom job / machine", "Job / machine name"), value=str(old.get("compute_label", "")), key=f"edit_comp_label_{editing_idx}")
            cur_type = str(old.get("compute_type", "cpu"))
            if cur_type not in {"cpu", "gpu", "mixed"}:
                cur_type = "cpu"
            e_type = st.selectbox(tf("Type", "Type"), ["cpu", "gpu", "mixed"], index=["cpu", "gpu", "mixed"].index(cur_type), key=f"edit_comp_type_{editing_idx}")
            e_hours = st.number_input(tf("Heures", "Hours"), min_value=0.0, value=float(old.get("hours", 0.0) or 0.0), step=1.0, key=f"edit_comp_hours_{editing_idx}")
            e_power = st.number_input(tf("Puissance moyenne (kW)", "Average power (kW)"), min_value=0.0, value=float(old.get("power_kw", 0.0) or 0.0), step=0.05, key=f"edit_comp_power_{editing_idx}")
            e_kwh = st.number_input(tf("Energie (kWh, optionnel)", "Energy (kWh, optional)"), min_value=0.0, value=float(old.get("kwh", 0.0) or 0.0), step=1.0, key=f"edit_comp_kwh_{editing_idx}")
            e_default_factor = float(DEFAULT_FACTORS["heures_calcul_kgco2e_per_kwh"])
            e_unc_pct = _factor_uncertainty_pct("heures_calcul_kgco2e_per_kwh")
            e_use_external_state = bool(st.session_state.get(f"edit_comp_{editing_idx}_use_external", bool(old.get("uses_external_emissions", False))))
            _render_default_factor_with_source(
                tf("Facteur (kgCO2e / kWh)", "Factor (kgCO2e / kWh)"),
                e_default_factor,
                e_unc_pct,
                _factor_reference("heures_calcul_kgco2e_per_kwh"),
                ignored=e_use_external_state,
                key=f"edit_comp_factor_display_{editing_idx}",
            )
            e_use_external, e_external_total, e_external_unc_raw = _external_emissions_input(
                tf("Declaration fournisseur", "Supplier declaration"),
                f"edit_comp_{editing_idx}",
                default_checked=bool(old.get("uses_external_emissions", False)),
                default_value=float(old.get("external_emissions_kgco2e", old.get("emissions_kgco2e", 0.0)) or 0.0),
            )
            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_comp_{editing_idx}"):
                updated = add_heures_calcul(
                    old.get("dossier_id", dossier_id),
                    old.get("dossier_type", dossier_type),
                    bool(old.get("is_anonymous", is_anonymous)),
                    old.get("team_code", team_code),
                    old.get("person_label", person_label),
                    _clean(e_label),
                    e_type,
                    e_hours,
                    e_power,
                    e_kwh,
                    e_default_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_external_emissions(updated, e_use_external, e_external_total)
                _apply_factor_uncertainty(updated, e_unc_pct)
                missing_unc = _apply_external_uncertainty(updated, e_use_external, e_external_unc_raw)
                _set_uncertainty_notice_if_missing(missing_unc)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_comp_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        elif old.get("poste") == "plateforme":
            e_ref = st.text_input(tf("Reference", "Reference"), value=str(old.get("ref_label", old.get("platform_name", ""))), key=f"edit_ref_platform_{editing_idx}")
            current_platform = str(old.get("platform_name", PLATFORM_OPTIONS[0]))
            if current_platform not in PLATFORM_OPTIONS:
                current_platform = PLATFORM_OPTIONS[0]
            e_platform = st.selectbox(tf("Plateforme", "Platform"), PLATFORM_OPTIONS, index=PLATFORM_OPTIONS.index(current_platform), key=f"edit_platform_name_{editing_idx}")
            role_opts = ["responsable", "utilisateur"]
            cur_role = str(old.get("platform_user_role", "utilisateur")).lower()
            if cur_role not in role_opts:
                cur_role = "utilisateur"
            e_role = st.selectbox(tf("Role", "Role"), role_opts, index=role_opts.index(cur_role), key=f"edit_platform_role_{editing_idx}")
            e_user_label = st.text_input(tf("Utilisateur", "User"), value=str(old.get("person_label", "")), key=f"edit_platform_user_{editing_idx}")
            e_usage_hours = st.number_input(tf("Heures d'utilisation", "Usage hours"), min_value=0.0, value=float(old.get("usage_hours", 0.0) or 0.0), step=1.0, key=f"edit_platform_usage_hours_{editing_idx}")
            e_usage_dates = st.text_input(tf("Dates approximatives", "Approximate dates"), value=str(old.get("usage_dates_label", "")), key=f"edit_platform_usage_dates_{editing_idx}")
            e_usage_desc = st.text_area(tf("Description d'usage", "Usage description"), value=str(old.get("usage_description", "")), key=f"edit_platform_usage_desc_{editing_idx}")
            material_options = list(DEFAULT_FACTORS["plateforme_material_factors"].keys())
            current_material = str(old.get("material_type", material_options[0]))
            if current_material not in material_options:
                current_material = material_options[0]
            e_material = st.selectbox(tf("Type de materiel achete", "Purchased material type"), material_options, index=material_options.index(current_material), key=f"edit_platform_material_{editing_idx}")
            e_purchase = st.number_input(tf("Achat de materiel (€)", "Equipment purchase (€)"), min_value=0.0, value=float(old.get("material_purchase_eur", 0.0) or 0.0), step=100.0, key=f"edit_platform_purchase_{editing_idx}")
            e_maintenance = st.number_input(tf("Frais de maintenance (€)", "Maintenance costs (€)"), min_value=0.0, value=float(old.get("maintenance_costs_eur", 0.0) or 0.0), step=100.0, key=f"edit_platform_maintenance_{editing_idx}")
            e_invoice = st.number_input(tf("Facture (€)", "Invoice (€)"), min_value=0.0, value=float(old.get("invoice_eur", 0.0) or 0.0), step=100.0, key=f"edit_platform_invoice_{editing_idx}")

            f_usage = float(DEFAULT_FACTORS["plateforme_usage_kgco2e_per_hour"])
            u_usage = _factor_uncertainty_pct("plateforme_usage_kgco2e_per_hour")
            _render_default_factor_with_source(
                tf("Facteur usage (kgCO2e / h)", "Usage factor (kgCO2e / h)"),
                f_usage,
                u_usage,
                _factor_reference("plateforme_usage_kgco2e_per_hour"),
                key=f"edit_platform_factor_usage_{editing_idx}",
            )
            f_material = float(DEFAULT_FACTORS["plateforme_material_factors"][e_material])
            u_material = _factor_uncertainty_pct("plateforme_material_factors", e_material)
            _render_default_factor_with_source(
                tf("Facteur materiel (kgCO2e / €)", "Material factor (kgCO2e / €)"),
                f_material,
                u_material,
                _factor_reference("plateforme_material_factors", e_material),
                key=f"edit_platform_factor_material_{editing_idx}",
            )
            f_maintenance = float(DEFAULT_FACTORS["plateforme_maintenance_kgco2e_per_eur"])
            u_maintenance = _factor_uncertainty_pct("plateforme_maintenance_kgco2e_per_eur")
            _render_default_factor_with_source(
                tf("Facteur maintenance (kgCO2e / €)", "Maintenance factor (kgCO2e / €)"),
                f_maintenance,
                u_maintenance,
                _factor_reference("plateforme_maintenance_kgco2e_per_eur"),
                key=f"edit_platform_factor_maintenance_{editing_idx}",
            )
            f_invoice = float(DEFAULT_FACTORS["plateforme_invoice_kgco2e_per_eur"])
            u_invoice = _factor_uncertainty_pct("plateforme_invoice_kgco2e_per_eur")
            _render_default_factor_with_source(
                tf("Facteur facture (kgCO2e / €)", "Invoice factor (kgCO2e / €)"),
                f_invoice,
                u_invoice,
                _factor_reference("plateforme_invoice_kgco2e_per_eur"),
                key=f"edit_platform_factor_invoice_{editing_idx}",
            )

            csave, ccancel = st.columns(2)
            if csave.button(tf("Enregistrer modifications", "Save changes"), key=f"save_edit_platform_{editing_idx}"):
                updated = add_plateforme(
                    old.get("dossier_id", dossier_id),
                    old.get("dossier_type", dossier_type),
                    False,
                    "PLATEFORMES",
                    e_user_label,
                    e_platform,
                    e_role,
                    e_usage_hours,
                    e_usage_dates,
                    e_usage_desc,
                    e_material,
                    e_purchase,
                    e_maintenance,
                    e_invoice,
                    f_usage,
                    f_material,
                    f_maintenance,
                    f_invoice,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                _apply_factor_uncertainty(updated, max(u_usage, u_material, u_maintenance, u_invoice))
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_platform_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        st.divider()

    if editing_row is None and poste == "achats":
        achat_ref = st.text_input(tf("Reference", "Reference"), value=tf("Achat 1", "Purchase 1"))
        achat_categories = DEFAULT_FACTORS["achats_category_factors"]
        achat_category = st.selectbox(tf("Type d'achat", "Purchase type"), options=list(achat_categories.keys()))
        amount_eur = st.number_input(tf("Montant (€)", "Amount (€)"), min_value=0.0, value=1500.0, step=100.0)
        default_factor = float(achat_categories[achat_category])
        unc_pct = _factor_uncertainty_pct("achats_category_factors", achat_category)
        use_external_state = bool(st.session_state.get("add_achat_use_external", False))
        _render_default_factor_with_source(
            tf("Facteur (kgCO2e / €)", "Factor (kgCO2e / €)"),
            default_factor,
            unc_pct,
            _factor_reference("achats_category_factors", achat_category),
            ignored=use_external_state,
            key=f"achat_factor_display_{achat_category}",
        )
        use_external, external_total, external_unc_raw = _external_emissions_input(
            tf("Chiffre fabricant", "Manufacturer figure"),
            "add_achat",
        )
        if st.button(tf("Calculer / Ajouter", "Compute / Add"), key="add_achats"):
            new_row = add_achat(
                dossier_id,
                dossier_type,
                is_anonymous,
                team_code,
                person_label,
                _clean(achat_category),
                amount_eur,
                default_factor,
            )
            new_row["ref_label"] = _clean(achat_ref)
            _apply_external_emissions(new_row, use_external, external_total)
            _apply_factor_uncertainty(new_row, unc_pct)
            missing_unc = _apply_external_uncertainty(new_row, use_external, external_unc_raw)
            _set_uncertainty_notice_if_missing(missing_unc)

    elif editing_row is None and poste == "domicile_travail":
        domicile_ref = st.text_input(tf("Reference", "Reference"), value=tf("Trajet domicile", "Commute trip"))
        mode_t = st.selectbox(tf("Mode de transport", "Transport mode"), list(DEFAULT_FACTORS["domicile_mode_factors"].keys()))
        dist = st.number_input(tf("Distance aller simple (km)", "One-way distance (km)"), min_value=0.0, value=12.0, step=1.0)
        add_days_key = "add_dom_days_year"
        row = st.columns([1.8, 1.0, 1.0, 1.0], vertical_alignment="center")
        with row[0]:
            days_year = st.number_input(tf("Jours/an", "Days/year"), min_value=0.0, max_value=366.0, value=180.0, step=1.0, key=add_days_key)
        for col, (pid, label, days_value) in zip(row[1:], DOMICILE_DAYS_PRESETS):
            _render_domicile_days_preset_cell(
                col=col,
                target_days_key=add_days_key,
                preset_id=pid,
                label=label,
                days_value=days_value,
                button_key_prefix="add_dom_days_preset",
            )
        round_trip = st.checkbox(tf("Aller-retour", "Round trip"), value=True)
        default_factor = float(DEFAULT_FACTORS["domicile_mode_factors"][mode_t])
        unc_pct = _factor_uncertainty_pct("domicile_mode_factors", mode_t)
        use_external_state = bool(st.session_state.get("add_dom_use_external", False))
        _render_default_factor_with_source(
            tf("Facteur (kgCO2e / km)", "Factor (kgCO2e / km)"),
            default_factor,
            unc_pct,
            _factor_reference("domicile_mode_factors", mode_t),
            ignored=use_external_state,
            key=f"dom_factor_display_{mode_t}",
        )
        use_external, external_total, external_unc_raw = _external_emissions_input(
            tf("Declaration compagnie de transport", "Transport company declaration"),
            "add_dom",
        )
        if st.button(tf("Calculer / Ajouter", "Compute / Add"), key="add_domicile"):
            new_row = add_domicile(
                dossier_id,
                dossier_type,
                is_anonymous,
                team_code,
                person_label,
                mode_t,
                dist,
                days_year,
                round_trip,
                default_factor,
            )
            new_row["ref_label"] = _clean(domicile_ref)
            _apply_external_emissions(new_row, use_external, external_total)
            _apply_factor_uncertainty(new_row, unc_pct)
            missing_unc = _apply_external_uncertainty(new_row, use_external, external_unc_raw)
            _set_uncertainty_notice_if_missing(missing_unc)

    elif editing_row is None and poste == "campagnes_terrain":
        campaign_ref = st.text_input(tf("Reference", "Reference"), value=tf("Campagne 1", "Campaign 1"))
        campaign_label = st.text_input(tf("Nom campagne", "Campaign name"), value=tf("Campagne terrain 1", "Field campaign 1"))
        segment_mode = st.selectbox(tf("Mode de transport", "Transport mode"), list(DEFAULT_FACTORS["campagnes_mode_factors"].keys()))
        dist = st.number_input(tf("Distance segment (km)", "Segment distance (km)"), min_value=0.0, value=800.0, step=10.0)
        passengers = st.number_input(tf("Nombre de personnes", "Number of people"), min_value=1.0, value=1.0, step=1.0)
        round_trip = st.checkbox(tf("Aller-retour", "Round trip"), value=True, key="camp_rt")
        default_factor = float(DEFAULT_FACTORS["campagnes_mode_factors"][segment_mode])
        unc_pct = _factor_uncertainty_pct("campagnes_mode_factors", segment_mode)
        use_external_state = bool(st.session_state.get("add_camp_use_external", False))
        _render_default_factor_with_source(
            tf("Facteur (kgCO2e / km / personne)", "Factor (kgCO2e / km / person)"),
            default_factor,
            unc_pct,
            _factor_reference("campagnes_mode_factors", segment_mode),
            ignored=use_external_state,
            key=f"camp_factor_display_{segment_mode}",
        )
        use_external, external_total, external_unc_raw = _external_emissions_input(
            tf("Declaration prestataire transport", "Transport provider declaration"),
            "add_camp",
        )
        if st.button(tf("Calculer / Ajouter", "Compute / Add"), key="add_campagne"):
            new_row = add_campagne(
                dossier_id,
                dossier_type,
                is_anonymous,
                team_code,
                person_label,
                _clean(campaign_label),
                segment_mode,
                dist,
                passengers,
                round_trip,
                default_factor,
            )
            new_row["ref_label"] = _clean(campaign_ref)
            _apply_external_emissions(new_row, use_external, external_total)
            _apply_factor_uncertainty(new_row, unc_pct)
            missing_unc = _apply_external_uncertainty(new_row, use_external, external_unc_raw)
            _set_uncertainty_notice_if_missing(missing_unc)

    elif editing_row is None and poste == "missions":
        mission_ref = st.text_input(tf("Reference", "Reference"), value="MISSION-001")
        st.caption(tf("Calcul mission en direct via Moulinette_missions (geocodage + distance + facteurs)", "Direct mission computation via Moulinette_missions (geocoding + distance + factors)"))
        mission_id = st.text_input("Mission ID", value="MISSION-001")
        steps = st.session_state.get("mission_steps_input", [])
        if not isinstance(steps, list) or not steps:
            steps = [_mission_default_step(arrival_city="Paris", arrival_country="FR")]
            st.session_state.mission_steps_input = steps

        collected_steps: list[dict] = []
        for idx, step in enumerate(steps):
            st.markdown(f"**{tf('Etape', 'Step')} {idx + 1}**")
            c_dep, c_arr, c_mode = st.columns([2.1, 2.1, 1.4])
            with c_dep:
                dep_city = st.text_input(
                    tf("Ville depart", "Departure city"),
                    value=str(step.get("departure_city", "")),
                    key=f"mis_step_{idx}_dep_city",
                    on_change=_propagate_departure_city_to_prev,
                    args=(idx,),
                )
                dep_country = _country_input(
                    tf("Pays depart", "Departure country"),
                    f"mis_step_{idx}_dep_country",
                    default_country=str(step.get("departure_country", "FR") or "FR"),
                    city_hint=dep_city,
                )
            with c_arr:
                arr_city = st.text_input(
                    tf("Ville arrivee", "Arrival city"),
                    value=str(step.get("arrival_city", "")),
                    key=f"mis_step_{idx}_arr_city",
                    on_change=_propagate_arrival_city_to_next,
                    args=(idx,),
                )
                arr_country = _country_input(
                    tf("Pays arrivee", "Arrival country"),
                    f"mis_step_{idx}_arr_country",
                    default_country=str(step.get("arrival_country", "FR") or "FR"),
                    city_hint=arr_city,
                )
            with c_mode:
                mode_opts = ["plane", "train", "car"]
                cur_mode = str(step.get("transport_mode", "plane"))
                if cur_mode not in mode_opts:
                    cur_mode = "plane"
                mode_val = st.selectbox(
                    tf("Transport", "Transport"),
                    mode_opts,
                    index=mode_opts.index(cur_mode),
                    key=f"mis_step_{idx}_mode",
                )
                step_rt = st.checkbox(
                    tf("Aller-retour", "Round trip"),
                    value=bool(step.get("round_trip", False)),
                    key=f"mis_step_{idx}_rt",
                )

            use_external_step, external_total_step, external_unc_step = _external_emissions_input(
                tf("Declaration compagnie de transport", "Transport company declaration"),
                f"mis_step_{idx}",
                default_checked=bool(step.get("use_external", False)),
                default_value=float(step.get("external_total_kgco2e", 0.0) or 0.0),
            )
            collected_steps.append(
                {
                    "departure_city": _clean(dep_city),
                    "departure_country": _clean(dep_country),
                    "arrival_city": _clean(arr_city),
                    "arrival_country": _clean(arr_country),
                    "transport_mode": mode_val,
                    "round_trip": bool(step_rt),
                    "use_external": bool(use_external_step),
                    "external_total_kgco2e": None if external_total_step is None else float(external_total_step),
                    "external_uncertainty_kgco2e": str(external_unc_step or "").strip(),
                }
            )

        b_add, b_del = st.columns(2)
        if b_add.button(tf("Ajouter etape", "Add step"), key="mission_add_step"):
            _add_mission_step_with_split_destination()
            st.rerun()
        if b_del.button(tf("Supprimer derniere etape", "Remove last step"), key="mission_remove_step", disabled=len(steps) <= 1):
            st.session_state.mission_steps_input = steps[:-1] if len(steps) > 1 else steps
            st.rerun()

        st.session_state.mission_steps_input = collected_steps
        mission_input_key = _mission_calc_input_key(mission_id, mission_ref, collected_steps, False, None, "")
        preview_rows = st.session_state.get("mission_preview_rows", [])
        preview_valid = bool(preview_rows) and st.session_state.get("mission_preview_key", "") == mission_input_key

        b_calc, b_commit = st.columns(2)
        calc_clicked = b_calc.button(tf("Calculer", "Compute"), key="calc_mission_steps")

        if calc_clicked:
            try:
                for idx, s in enumerate(collected_steps):
                    if not s["departure_city"] or not s["arrival_city"]:
                        raise ValueError(
                            tf(
                                f"Etape {idx + 1}: renseigne ville depart et ville arrivee.",
                                f"Step {idx + 1}: fill departure and arrival cities.",
                            )
                        )

                mission_clean = _clean(mission_id)
                new_rows: list[dict] = []
                for idx, s in enumerate(collected_steps):
                    row = compute_single_mission_with_moulinette(
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        is_anonymous=is_anonymous,
                        team_code=team_code,
                        person_label=person_label,
                        mission_id=mission_clean,
                        departure_city=s["departure_city"],
                        departure_country=s["departure_country"],
                        arrival_city=s["arrival_city"],
                        arrival_country=s["arrival_country"],
                        transport_mode=s["transport_mode"],
                        round_trip=s["round_trip"],
                    )
                    row["ref_label"] = _clean(mission_ref)
                    row["mission_step_index"] = idx + 1
                    row["mission_step_count"] = len(collected_steps)
                    row["mission_segment_id"] = f"{mission_clean}-S{idx + 1:02d}"
                    emissions_val = float(row.get("emissions_kgco2e", 0.0) or 0.0)
                    uncertainty_val = float(row.get("uncertainty_kgco2e", 0.0) or 0.0)
                    unc_pct = (uncertainty_val / emissions_val * 100.0) if emissions_val > 0 else 10.0
                    row["mission_factor_default"] = float(row.get("mission_effective_factor_kgco2e_per_km", 0.0) or 0.0)
                    row["mission_factor_uncertainty_pct"] = float(unc_pct)
                    row["mission_factor_source"] = str(row.get("emission_transport", "") or "")
                    _apply_factor_uncertainty(row, unc_pct)
                    new_rows.append(row)

                any_missing_unc = False
                for row, step_data in zip(new_rows, collected_steps):
                    step_use_external = bool(step_data.get("use_external", False))
                    step_external_total = step_data.get("external_total_kgco2e", None)
                    step_external_unc_raw = str(step_data.get("external_uncertainty_kgco2e", "")).strip()
                    _apply_external_emissions(
                        row,
                        step_use_external,
                        float(step_external_total) if step_external_total is not None else None,
                    )
                    missing_unc = _apply_external_uncertainty(row, step_use_external, step_external_unc_raw)
                    any_missing_unc = any_missing_unc or missing_unc
                _set_uncertainty_notice_if_missing(any_missing_unc)

                st.session_state.mission_preview_rows = new_rows
                st.session_state.mission_preview_key = mission_input_key
            except Exception as exc:  # pragma: no cover - UI path
                st.error(tf(f"Erreur moulinette: {exc}", f"Moulinette error: {exc}"))

        preview_rows = st.session_state.get("mission_preview_rows", [])
        preview_valid = bool(preview_rows) and st.session_state.get("mission_preview_key", "") == mission_input_key
        add_clicked = b_commit.button(tf("Ajouter", "Add"), key="add_mission_direct", disabled=not preview_valid)

        if add_clicked:
            rows_to_add = st.session_state.get("mission_preview_rows", [])
            st.session_state.entries.extend(rows_to_add)
            total_dist = sum(float(r.get("distance_total_km", 0.0) or 0.0) for r in rows_to_add)
            total_em = sum(float(r.get("emissions_kgco2e", 0.0) or 0.0) for r in rows_to_add)
            st.success(
                tf(
                    f"Mission ajoutee: {len(rows_to_add)} etape(s) | Distance totale: {total_dist:.1f} km | Emissions: {total_em:.2f} kgCO2e",
                    f"Mission added: {len(rows_to_add)} step(s) | Total distance: {total_dist:.1f} km | Emissions: {total_em:.2f} kgCO2e",
                )
            )
            st.session_state.mission_preview_rows = []
            st.session_state.mission_preview_key = ""

        preview_rows = st.session_state.get("mission_preview_rows", [])
        preview_valid = bool(preview_rows) and st.session_state.get("mission_preview_key", "") == mission_input_key

        if preview_valid:
            st.write(f"#### {tf('Apercu calcul mission', 'Mission calculation preview')}")
            preview_df = pd.DataFrame(
                [
                    {
                        tf("Etape", "Step"): int(r.get("mission_step_index", 0)),
                        tf("Trajet", "Route"): f"{r.get('departure_city', '')} -> {r.get('arrival_city', '')}",
                        tf("Mode", "Mode"): str(r.get("t_type", "")),
                        tf("Distance geo (km)", "Geodesic distance (km)"): float(r.get("distance_total_km", 0.0) or 0.0),
                        tf("Distance corrigee (km)", "Corrected distance (km)"): float(r.get("distance_corrected_total_km", 0.0) or 0.0),
                        tf("Coeff correction distance", "Distance correction factor"): f"x{float(r.get('distance_correction_factor', 1.0) or 1.0):.2f}",
                        tf("Facteur", "Factor"): _format_factor_value_pm(
                            float(r.get("mission_factor_default", 0.0) or 0.0),
                            float(r.get("mission_factor_uncertainty_pct", 10.0) or 10.0),
                        ),
                        tf("Type facteur Moulinette", "Moulinette factor type"): str(r.get("emission_transport", "")),
                        "kgCO2e": float(r.get("emissions_kgco2e", 0.0) or 0.0),
                    }
                    for r in preview_rows
                ]
            )
            _render_themed_dataframe(preview_df, use_container_width=True, hide_index=True)

        if st.session_state.missions_map_bytes:
            st.write(tf("Carte des emissions (Moulinette)", "Emissions map (Moulinette)"))
            st.image(st.session_state.missions_map_bytes)

    elif editing_row is None and poste == "heures_calcul":
        compute_ref = st.text_input(tf("Reference", "Reference"), value=tf("Calcul 1", "Compute 1"))
        compute_label = st.text_input(tf("Nom job / machine", "Job / machine name"), value="Cluster A")
        compute_type = st.selectbox(tf("Type", "Type"), ["cpu", "gpu", "mixed"])
        hours = st.number_input(tf("Heures", "Hours"), min_value=0.0, value=100.0, step=1.0)
        power_kw = st.number_input(tf("Puissance moyenne (kW)", "Average power (kW)"), min_value=0.0, value=0.3, step=0.05)
        kwh = st.number_input(tf("Energie (kWh, optionnel)", "Energy (kWh, optional)"), min_value=0.0, value=0.0, step=1.0)
        default_factor = float(DEFAULT_FACTORS["heures_calcul_kgco2e_per_kwh"])
        unc_pct = _factor_uncertainty_pct("heures_calcul_kgco2e_per_kwh")
        use_external_state = bool(st.session_state.get("add_comp_use_external", False))
        _render_default_factor_with_source(
            tf("Facteur (kgCO2e / kWh)", "Factor (kgCO2e / kWh)"),
            default_factor,
            unc_pct,
            _factor_reference("heures_calcul_kgco2e_per_kwh"),
            ignored=use_external_state,
            key="comp_factor_display",
        )
        use_external, external_total, external_unc_raw = _external_emissions_input(
            tf("Declaration fournisseur", "Supplier declaration"),
            "add_comp",
        )
        if st.button(tf("Calculer / Ajouter", "Compute / Add"), key="add_calcul"):
            new_row = add_heures_calcul(
                dossier_id,
                dossier_type,
                is_anonymous,
                team_code,
                person_label,
                _clean(compute_label),
                compute_type,
                hours,
                power_kw,
                kwh,
                default_factor,
            )
            new_row["ref_label"] = _clean(compute_ref)
            _apply_external_emissions(new_row, use_external, external_total)
            _apply_factor_uncertainty(new_row, unc_pct)
            missing_unc = _apply_external_uncertainty(new_row, use_external, external_unc_raw)
            _set_uncertainty_notice_if_missing(missing_unc)

    elif editing_row is None and poste == "plateforme":
        platform_ref = st.text_input(tf("Reference", "Reference"), value=tf("Plateforme 1", "Platform 1"))
        platform_name = st.session_state.get("platform_name_input", PLATFORM_OPTIONS[0])
        st.caption(f"{tf('Plateforme selectionnee', 'Selected platform')}: {platform_name}")
        st.markdown(f"**{tf('Equipes impliquees', 'Involved teams')}**")
        selected_involved_teams: list[str] = []
        team_checkbox_cols = st.columns(3) if len(PLATFORM_INVOLVED_TEAM_CODES) >= 3 else st.columns(max(1, len(PLATFORM_INVOLVED_TEAM_CODES)))
        for idx, code in enumerate(PLATFORM_INVOLVED_TEAM_CODES):
            default_checked = code in st.session_state.get("platform_involved_teams_input", [])
            checked = team_checkbox_cols[idx % len(team_checkbox_cols)].checkbox(
                TEAM_OPTIONS.get(code, code),
                value=default_checked,
                key=f"platform_involved_team_{code}",
            )
            if checked:
                selected_involved_teams.append(code)
        st.session_state.platform_involved_teams_input = selected_involved_teams
        role_opts = PLATFORM_POSTE_TYPE_OPTIONS
        current_role = poste_type if poste_type in role_opts else "utilisateur"
        identity_missing = not _clean(person_label) or (poste_type not in role_opts)
        st.caption(
            tf(
                f"Utilisateur courant: {person_label or 'anonyme'} ({current_role})",
                f"Current user: {person_label or 'anonymous'} ({current_role})",
            )
        )
        b_achats, b_usage, b_frais = st.columns(3)
        if b_achats.button(
            tf("Ajouter achats", "Add purchases"),
            key="platform_pick_achats",
            disabled=current_role != "responsable",
        ):
            st.session_state.platform_active_form = "achats"
            st.session_state.platform_preview_row = None
        if b_usage.button(tf("Ajouter utilisation", "Add usage"), key="platform_pick_usage"):
            st.session_state.platform_active_form = "utilisation"
            st.session_state.platform_preview_row = None
        if b_frais.button(
            tf("Ajouter frais", "Add expenses"),
            key="platform_pick_frais",
            disabled=current_role != "responsable",
        ):
            st.session_state.platform_active_form = "frais"
            st.session_state.platform_preview_row = None

        active_form = str(st.session_state.get("platform_active_form", "")).strip()
        material_options = list(DEFAULT_FACTORS["plateforme_material_factors"].keys())
        factor_usage = float(DEFAULT_FACTORS["plateforme_usage_kgco2e_per_hour"])
        unc_usage = _factor_uncertainty_pct("plateforme_usage_kgco2e_per_hour")
        factor_maintenance = float(DEFAULT_FACTORS["plateforme_maintenance_kgco2e_per_eur"])
        unc_maintenance = _factor_uncertainty_pct("plateforme_maintenance_kgco2e_per_eur")
        factor_invoice = float(DEFAULT_FACTORS["plateforme_invoice_kgco2e_per_eur"])
        unc_invoice = _factor_uncertainty_pct("plateforme_invoice_kgco2e_per_eur")
        if not active_form:
            st.info(tf("Choisis un type d'ajout pour afficher les champs.", "Choose an add type to show fields."))
        else:
            if active_form == "achats":
                st.markdown(f"**{tf('Ajouts achats', 'Purchase entry')}**")
                purchase_date = st.date_input(tf("Date d'achat", "Purchase date"), key="platform_purchase_date")
                material_type = st.selectbox(tf("Type de materiel achete", "Purchased material type"), material_options, key="platform_material_type_single")
                material_purchase = st.number_input(tf("Montant achat (€)", "Purchase amount (€)"), min_value=0.0, value=0.0, step=100.0, key="platform_material_purchase_single")
                factor_material = float(DEFAULT_FACTORS["plateforme_material_factors"][material_type])
                unc_material = _factor_uncertainty_pct("plateforme_material_factors", material_type)
                _render_default_factor_with_source(
                    tf("Facteur materiel (kgCO2e / €)", "Material factor (kgCO2e / €)"),
                    factor_material,
                    unc_material,
                    _factor_reference("plateforme_material_factors", material_type),
                    key="platform_factor_material_single",
                )
                c_compute, c_add = st.columns(2)
                calc_clicked = c_compute.button(tf("Calculer", "Compute"), key="platform_calc_achats")
                if calc_clicked:
                    preview_row = add_plateforme(
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        is_anonymous=False,
                        team_code="PLATEFORMES",
                        person_label=person_label,
                        platform_name=platform_name,
                        user_role=current_role,
                        usage_hours=0.0,
                        usage_dates_label=str(purchase_date),
                        usage_description=tf("Achat plateforme", "Platform purchase"),
                        material_type=material_type,
                        material_purchase_eur=material_purchase,
                        maintenance_costs_eur=0.0,
                        invoice_eur=0.0,
                        usage_factor_kgco2e_per_hour=factor_usage,
                        material_factor_kgco2e_per_eur=factor_material,
                        maintenance_factor_kgco2e_per_eur=factor_maintenance,
                        invoice_factor_kgco2e_per_eur=factor_invoice,
                    )
                    preview_row["platform_entry_type"] = "achats"
                    preview_row["platform_involved_team_codes"] = ";".join(selected_involved_teams)
                    preview_row["platform_involved_teams"] = ", ".join(TEAM_OPTIONS.get(c, c) for c in selected_involved_teams)
                    preview_row["ref_label"] = _clean(platform_ref)
                    _apply_factor_uncertainty(preview_row, unc_material)
                    st.session_state.platform_preview_row = preview_row
                add_disabled = not (
                    isinstance(st.session_state.get("platform_preview_row"), dict)
                    and st.session_state.platform_preview_row.get("platform_entry_type") == "achats"
                )
                if c_add.button(tf("Ajouter", "Add"), key="platform_add_achats", disabled=add_disabled):
                    if identity_missing:
                        st.error(tf("Complete d'abord l'identite a gauche (nom et type de poste).", "Please complete identity on the left first (name and position type)."))
                    else:
                        row_to_add = st.session_state.get("platform_preview_row")
                        if isinstance(row_to_add, dict):
                            st.session_state.entries.append(row_to_add)
                            st.session_state.platform_preview_row = None
                            platform_rows_added = True
                            st.success(tf("Ligne plateforme ajoutee.", "Platform row added."))
            elif active_form == "utilisation":
                st.markdown(f"**{tf('Ajouts utilisation', 'Usage entry')}**")
                usage_hours = st.number_input(tf("Heures d'utilisation", "Usage hours"), min_value=0.0, value=0.0, step=1.0, key="platform_usage_hours_single")
                c_date1, c_date2 = st.columns(2)
                usage_date_from = c_date1.date_input(tf("Date debut", "Start date"), key="platform_usage_date_from")
                usage_date_to = c_date2.date_input(tf("Date fin", "End date"), key="platform_usage_date_to")
                usage_desc = st.text_area(tf("Description", "Description"), key="platform_usage_desc_single")
                _render_default_factor_with_source(
                    tf("Facteur usage (kgCO2e / h)", "Usage factor (kgCO2e / h)"),
                    factor_usage,
                    unc_usage,
                    _factor_reference("plateforme_usage_kgco2e_per_hour"),
                    key="platform_factor_usage_single",
                )
                c_compute, c_add = st.columns(2)
                calc_clicked = c_compute.button(tf("Calculer", "Compute"), key="platform_calc_usage")
                if calc_clicked:
                    preview_row = add_plateforme(
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        is_anonymous=False,
                        team_code="PLATEFORMES",
                        person_label=person_label,
                        platform_name=platform_name,
                        user_role=current_role,
                        usage_hours=usage_hours,
                        usage_dates_label=f"{usage_date_from} -> {usage_date_to}",
                        usage_description=_clean(usage_desc),
                        material_type=material_options[0],
                        material_purchase_eur=0.0,
                        maintenance_costs_eur=0.0,
                        invoice_eur=0.0,
                        usage_factor_kgco2e_per_hour=factor_usage,
                        material_factor_kgco2e_per_eur=float(DEFAULT_FACTORS["plateforme_material_factors"][material_options[0]]),
                        maintenance_factor_kgco2e_per_eur=factor_maintenance,
                        invoice_factor_kgco2e_per_eur=factor_invoice,
                    )
                    preview_row["platform_entry_type"] = "utilisation"
                    preview_row["platform_involved_team_codes"] = ";".join(selected_involved_teams)
                    preview_row["platform_involved_teams"] = ", ".join(TEAM_OPTIONS.get(c, c) for c in selected_involved_teams)
                    preview_row["ref_label"] = _clean(platform_ref)
                    _apply_factor_uncertainty(preview_row, unc_usage)
                    st.session_state.platform_preview_row = preview_row
                add_disabled = not (
                    isinstance(st.session_state.get("platform_preview_row"), dict)
                    and st.session_state.platform_preview_row.get("platform_entry_type") == "utilisation"
                )
                if c_add.button(tf("Ajouter", "Add"), key="platform_add_usage", disabled=add_disabled):
                    if identity_missing:
                        st.error(tf("Complete d'abord l'identite a gauche (nom et type de poste).", "Please complete identity on the left first (name and position type)."))
                    else:
                        row_to_add = st.session_state.get("platform_preview_row")
                        if isinstance(row_to_add, dict):
                            st.session_state.entries.append(row_to_add)
                            st.session_state.platform_preview_row = None
                            platform_rows_added = True
                            st.success(tf("Ligne plateforme ajoutee.", "Platform row added."))
            else:
                st.markdown(f"**{tf('Ajouts frais', 'Expense entry')}**")
                expense_date = st.date_input(tf("Date frais", "Expense date"), key="platform_expense_date")
                maintenance_costs = st.number_input(tf("Montant frais (€)", "Expense amount (€)"), min_value=0.0, value=0.0, step=100.0, key="platform_maintenance_single")
                expense_desc = st.text_area(tf("Description", "Description"), key="platform_expense_desc_single")
                _render_default_factor_with_source(
                    tf("Facteur frais (kgCO2e / €)", "Expense factor (kgCO2e / €)"),
                    factor_maintenance,
                    unc_maintenance,
                    _factor_reference("plateforme_maintenance_kgco2e_per_eur"),
                    key="platform_factor_frais_single",
                )
                c_compute, c_add = st.columns(2)
                calc_clicked = c_compute.button(tf("Calculer", "Compute"), key="platform_calc_frais")
                if calc_clicked:
                    preview_row = add_plateforme(
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        is_anonymous=False,
                        team_code="PLATEFORMES",
                        person_label=person_label,
                        platform_name=platform_name,
                        user_role=current_role,
                        usage_hours=0.0,
                        usage_dates_label=str(expense_date),
                        usage_description=_clean(expense_desc) or tf("Frais plateforme", "Platform expense"),
                        material_type=material_options[0],
                        material_purchase_eur=0.0,
                        maintenance_costs_eur=maintenance_costs,
                        invoice_eur=0.0,
                        usage_factor_kgco2e_per_hour=factor_usage,
                        material_factor_kgco2e_per_eur=float(DEFAULT_FACTORS["plateforme_material_factors"][material_options[0]]),
                        maintenance_factor_kgco2e_per_eur=factor_maintenance,
                        invoice_factor_kgco2e_per_eur=factor_invoice,
                    )
                    preview_row["platform_entry_type"] = "frais"
                    preview_row["platform_involved_team_codes"] = ";".join(selected_involved_teams)
                    preview_row["platform_involved_teams"] = ", ".join(TEAM_OPTIONS.get(c, c) for c in selected_involved_teams)
                    preview_row["ref_label"] = _clean(platform_ref)
                    _apply_factor_uncertainty(preview_row, unc_maintenance)
                    st.session_state.platform_preview_row = preview_row
                add_disabled = not (
                    isinstance(st.session_state.get("platform_preview_row"), dict)
                    and st.session_state.platform_preview_row.get("platform_entry_type") == "frais"
                )
                if c_add.button(tf("Ajouter", "Add"), key="platform_add_frais", disabled=add_disabled):
                    if identity_missing:
                        st.error(tf("Complete d'abord l'identite a gauche (nom et type de poste).", "Please complete identity on the left first (name and position type)."))
                    else:
                        row_to_add = st.session_state.get("platform_preview_row")
                        if isinstance(row_to_add, dict):
                            st.session_state.entries.append(row_to_add)
                            st.session_state.platform_preview_row = None
                            platform_rows_added = True
                            st.success(tf("Ligne plateforme ajoutee.", "Platform row added."))

            preview_data = st.session_state.get("platform_preview_row")
            if isinstance(preview_data, dict):
                st.caption(tf(f"Calcul termine: {preview_data['emissions_kgco2e']:.2f} kgCO2e", f"Computed: {preview_data['emissions_kgco2e']:.2f} kgCO2e"))

    if new_row is not None and not platform_rows_added:
        st.session_state.entries.append(new_row)
        st.success(tf(f"Ligne ajoutee: {new_row['poste']} -> {new_row['emissions_kgco2e']:.2f} kgCO2e", f"Row added: {new_row['poste']} -> {new_row['emissions_kgco2e']:.2f} kgCO2e"))

with col2:
    st.write(f"### {tf('Resultat dossier', 'Form result')}")
    df = _ensure_uncertainty_columns(entries_to_df(st.session_state.entries))
    total = float(df["emissions_kgco2e"].sum()) if not df.empty else 0.0
    total_unc = float(df["uncertainty_kgco2e"].sum()) if not df.empty else 0.0
    st.metric(tf("Total", "Total"), _format_total_with_uncertainty(total, total_unc))
    st.metric(tf("Nombre de lignes", "Number of rows"), len(df))
    if mode == "plateforme" and not df.empty and "platform_entry_type" in df.columns:
        donut_df = (
            df[df["poste"] == "plateforme"]
            .assign(platform_entry_type=lambda d: d["platform_entry_type"].fillna("").replace("", tf("non precise", "unspecified")))
            .groupby("platform_entry_type", as_index=False)["emissions_kgco2e"]
            .sum()
        )
        if not donut_df.empty:
            st.write(tf("Repartition plateforme", "Platform breakdown"))
            donut = (
                alt.Chart(donut_df)
                .mark_arc(innerRadius=58)
                .encode(
                    theta=alt.Theta("emissions_kgco2e:Q", title="kgCO2e"),
                    color=alt.Color("platform_entry_type:N", title=tf("Poste", "Category")),
                    tooltip=[
                        alt.Tooltip("platform_entry_type:N", title=tf("Poste", "Category")),
                        alt.Tooltip("emissions_kgco2e:Q", title="kgCO2e", format=".2f"),
                    ],
                )
                .properties(height=260)
            )
            st.altair_chart(donut, use_container_width=True)

    if st.button(tf("Supprimer toutes les lignes", "Delete all rows")):
        st.session_state.entries = []
        st.session_state.missions_map_bytes = b""
        st.rerun()

st.markdown(
    """
<style>
/* Keep responsive typography only in detail area tables */
div[data-testid="stDataFrame"] *{
  font-size: clamp(11px, 0.95vw, 13px) !important;
}
/* Ensure action buttons (e.g. Modifier/Supprimer) remain readable on narrow screens */
div[data-testid="stButton"] button{
  font-size: clamp(10px, 0.95vw, 13px) !important;
  white-space: nowrap;
}
</style>
""",
    unsafe_allow_html=True,
)
st.write(f"### {tf('Detail', 'Detail')}")
df = _ensure_uncertainty_columns(entries_to_df(st.session_state.entries))
if df.empty:
    st.info(tf("Aucune ligne pour l'instant", "No row yet."))
else:
    if mode in {"bilan_personnel", "bilan_projet"}:
        view_df = df.reset_index(names="row_idx")
        h1, h2, h3, h4, h5, h6, h7, h8, h9 = st.columns([1.2, 1.4, 1.2, 2.0, 1.3, 1.0, 1.0, 0.8, 0.8])
        h1.write(tf("Poste", "Category"))
        h2.write("Ref")
        h3.write(tf("Equipe", "Team"))
        h4.write(tf("Personne", "Person"))
        h5.write(tf("Mis a jour le", "Updated at"))
        h6.write("kgCO2e")
        h7.write(tf("± kgCO2e", "± kgCO2e"))
        h8.write("")
        h9.write("")

        for _, row in view_df.iterrows():
            ref = row.get("ref_label") or row.get("mission_id") or row.get("item_label") or row.get("campaign_label") or row.get("compute_label") or row.get("record_id")
            c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns([1.2, 1.4, 1.2, 2.0, 1.3, 1.0, 1.0, 0.8, 0.8])
            c1.write(str(row.get("poste", "")))
            c2.write(str(ref))
            c3.write(str(row.get("team_code", "")))
            c4.write(str(row.get("person_label", "")))
            c5.write(str(row.get("updated_at", "")))
            c6.write(f"{float(row.get('emissions_kgco2e', 0)):.2f}")
            unc_text = f"{float(row.get('uncertainty_kgco2e', 0)):.2f}"
            if bool(row.get("uncertainty_missing_defaulted", False)):
                unc_text = f"🔺 {unc_text}"
            c7.write(unc_text)
            if c8.button(tf("Modifier", "Edit"), key=f"edit_row_{int(row['row_idx'])}_{row.get('record_id', '')}"):
                st.session_state.editing_row_idx = int(row["row_idx"])
                st.rerun()
            if c9.button(tf("Supprimer", "Delete"), key=f"delete_row_{int(row['row_idx'])}_{row.get('record_id', '')}_del"):
                del st.session_state.entries[int(row["row_idx"])]
                if st.session_state.get("editing_row_idx") == int(row["row_idx"]):
                    st.session_state.editing_row_idx = None
                st.rerun()
    else:
        show_cols = [
            c
            for c in [
                "poste",
                "team_code",
                "person_label",
                "item_label",
                "mission_id",
                "emission_transport",
                "distance_total_km",
                "distance_corrected_total_km",
                "distance_correction_factor",
                "mission_effective_factor_kgco2e_per_km",
                "campaign_label",
                "compute_label",
                "platform_name",
                "material_type",
                "emissions_kgco2e",
                "uncertainty_kgco2e",
                "uncertainty_factor_pct",
                "updated_at",
            ]
            if c in df.columns
        ]
        display_cols = {
            "distance_corrected_total_km": tf("Dist. corrigee (km)", "Corr. dist. (km)"),
            "distance_correction_factor": tf("Coeff. corr.", "Corr. coeff."),
            "mission_effective_factor_kgco2e_per_km": tf("Facteur CO2 reel", "Actual CO2 factor"),
        }
        _render_themed_dataframe(df[show_cols].rename(columns=display_cols), use_container_width=True)

    if mode in {"bilan_personnel", "bilan_projet"}:
        st.write(f"### {tf('Repartition des postes (%)', 'Category split (%)')}")
        donut_df = (
            df.groupby("poste", as_index=False)["emissions_kgco2e"].sum().rename(columns={"emissions_kgco2e": "emissions"})
        )
        total_emissions = float(donut_df["emissions"].sum())
        if total_emissions > 0:
            donut_df["pct"] = (donut_df["emissions"] / total_emissions) * 100
            donut_df["poste_display"] = donut_df["poste"].apply(_poste_label)
            donut_df["poste_legend"] = donut_df.apply(lambda r: f"{r['poste_display']} ({r['pct']:.1f}%)", axis=1)
            is_light = st.session_state.get("ui_theme", "darkmode") == "lightmode"
            legend_text_color = "#111111" if is_light else "#e8f0ff"
            chart_bg = "#ffffff" if is_light else None
            chart = (
                alt.Chart(donut_df)
                .mark_arc(innerRadius=70)
                .encode(
                theta=alt.Theta(field="emissions", type="quantitative"),
                color=alt.Color(
                    field="poste_legend",
                    type="nominal",
                    legend=alt.Legend(
                        title=tf("Poste", "Category"),
                        orient="right",
                        offset=0,
                        padding=0,
                        labelColor=legend_text_color,
                        titleColor=legend_text_color,
                    ),
                ),
                tooltip=[
                    alt.Tooltip("poste_display:N", title=tf("Poste", "Category")),
                    alt.Tooltip("emissions:Q", title="kgCO2e", format=".2f"),
                    alt.Tooltip("pct:Q", title="%", format=".1f"),
                ],
                )
                .properties(height=320)
            )
            if chart_bg:
                chart = chart.configure(background=chart_bg).configure_view(fill=chart_bg, stroke=None)
            st.altair_chart(chart, use_container_width=True)

        missions_arcs = _missions_arcs_df(df)
        if not missions_arcs.empty:
            st.write(f"### {tf('Cartes des trajets missions', 'Mission route maps')}")
            cards = missions_arcs.to_dict(orient="records")
            ncols = 2
            cols = st.columns(ncols)
            for idx, mission in enumerate(cards):
                one = pd.DataFrame([mission])
                with cols[idx % ncols]:
                    st.caption(f"{mission.get('mission_id', '')}: {mission.get('departure_city', '')} -> {mission.get('arrival_city', '')}")
                    st.pydeck_chart(_mission_deck(one, height=260), use_container_width=True)

    if mode == "item_unique":
        missions_arcs = _missions_arcs_df(df)
        if not missions_arcs.empty:
            st.write(f"### {tf('Carte du trajet mission', 'Mission route map')}")
            st.pydeck_chart(_mission_deck(missions_arcs, height=360), use_container_width=True)

st.write(f"### {tf('Export', 'Export')}")
owner_label_for_meta = st.session_state.loaded_owner_label or person_label or ""
meta = {
    "dossier_id": dossier_id,
    "lab_id": LAB_ID,
    "year": _year_folder(year_value),
    "dossier_type": dossier_type,
    "dossier_name": project_name if mode == "bilan_projet" else "",
    "platform_name": st.session_state.get("platform_name_input", "") if mode == "plateforme" else "",
    "anonyme_default": bool(is_anonymous),
    "owner_label": owner_label_for_meta,
    "poste_type": st.session_state.get("poste_type_input", ""),
    "team_code": team_code,
    "schema_version": "1.0",
    "exported_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
}
excel_bytes = export_excel_bytes(meta=meta, entries=st.session_state.entries)
st.download_button(
    tf("Telecharger le fichier Excel", "Download Excel file"),
    data=excel_bytes,
    file_name=f"carbonometre_{dossier_id}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
