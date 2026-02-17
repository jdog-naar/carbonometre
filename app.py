from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import re

import altair as alt
import pandas as pd
import pydeck as pdk
import streamlit as st

from carbonometre.calculations import (
    add_achat,
    add_campagne,
    add_domicile,
    add_heures_calcul,
    build_synthese,
    entries_to_df,
)
from carbonometre.constants import DEFAULT_FACTORS, LAB_ID, LAB_LOGO_PATH, LAB_NAME, LAB_THEME, TEAM_OPTIONS
from carbonometre.excel_io import export_excel_bytes, import_excel_entries
from carbonometre.missions_bridge import compute_single_mission_with_moulinette

if "ui_lang" not in st.session_state:
    st.session_state.ui_lang = "FR"


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
if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "darkmode"
if "poste_type_input" not in st.session_state:
    st.session_state.poste_type_input = ""
if "mode_input" not in st.session_state:
    st.session_state.mode_input = "item_unique"
if "project_name_input" not in st.session_state:
    st.session_state.project_name_input = "Mon projet"
if "team_code_input" not in st.session_state:
    st.session_state.team_code_input = ""
if "year_input" not in st.session_state:
    current_year = str(datetime.now(timezone.utc).year)
    st.session_state.year_input = current_year if current_year in {"2026", "2025", "2024", "2023", "2022", "2021", "2020", "1998"} else "2026"
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
if "editing_row_idx" not in st.session_state:
    st.session_state.editing_row_idx = None

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
    "Autre (saisie manuelle)",
]
SAVE_MODE_SUFFIX = {"bilan_personnel": "bilan_annuel", "bilan_projet": "projet"}
STORAGE_ROOT = Path(__file__).resolve().parent / LAB_ID
CURRENT_YEAR = datetime.now(timezone.utc).year
BASE_YEAR_OPTIONS = [str(y) for y in range(2026, 2019, -1)] + ["1998"]
YEAR_OPTIONS = BASE_YEAR_OPTIONS
YEAR_OPTIONS_ASC = sorted(YEAR_OPTIONS, key=int)
DOSSIER_TYPE_BY_MODE = {
    "item_unique": "item",
    "bilan_personnel": "personnel",
    "bilan_projet": "projet",
}
POSTE_TYPE_OPTIONS = [
    "",
    "CR",
    "DR",
    "Doctorant",
    "ITA",
    "ITA Contractuel",
    "Post-Doctorant",
    "Stagiaire",
]
ADMIN_GESTION_POSTE_TYPE_OPTIONS = ["", "Contractuel", "Permanent"]
EMPLOYMENT_BY_POSTE_TYPE = {
    "CR": "Permanent",
    "DR": "Permanent",
    "ITA": "Permanent",
    "Doctorant": "Contractuel",
    "ITA Contractuel": "Contractuel",
    "Post-Doctorant": "Contractuel",
    "Stagiaire": "Contractuel",
}
MODE_OPTIONS = ["item_unique", "bilan_personnel", "bilan_projet"]
POSTE_OPTIONS = ["achats", "domicile_travail", "campagnes_terrain", "missions", "heures_calcul"]


def _mode_label(m: str) -> str:
    return {
        "item_unique": tf("Item unique", "Single item"),
        "bilan_personnel": tf("Bilan personnel", "Personal assessment"),
        "bilan_projet": tf("Bilan projet", "Project assessment"),
    }.get(m, m)


def _poste_label(p: str) -> str:
    return {
        "achats": tf("Achats", "Purchases"),
        "domicile_travail": tf("Domicile-travail", "Commute"),
        "campagnes_terrain": tf("Campagnes terrain", "Field campaigns"),
        "missions": tf("Missions", "Missions"),
        "heures_calcul": tf("Heures de calcul", "Compute hours"),
    }.get(p, p)


def _normalize_mode_value(v: str) -> str:
    mapping = {
        "item_unique": "item_unique",
        "bilan_personnel": "bilan_personnel",
        "bilan_projet": "bilan_projet",
        "Item unique": "item_unique",
        "Single item": "item_unique",
        "Bilan personnel": "bilan_personnel",
        "Personal assessment": "bilan_personnel",
        "Bilan projet": "bilan_projet",
        "Project assessment": "bilan_projet",
    }
    return mapping.get(str(v), "item_unique")


def _normalize_poste_type_for_overview(team_code: str, poste_type: str) -> str:
    p = str(poste_type or "").strip()
    if team_code == "ADMIN_GESTION":
        if p.lower() == "contractuel":
            return "Admin C"
        if p.lower() == "permanent":
            return "Admin"
    return p


def _employment_type(team_code: str, poste_type: str) -> str:
    p = str(poste_type or "").strip()
    if team_code == "ADMIN_GESTION":
        if p.lower() == "contractuel":
            return "Contractuel"
        if p.lower() == "permanent":
            return "Permanent"
        return ""
    return EMPLOYMENT_BY_POSTE_TYPE.get(p, "")


def _clean(v: str) -> str:
    return v.strip()


def _slug(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_-]+", "_", _clean(s))
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "anonyme"


def _team_folder(team_code: str, mode: str = "", is_anonymous: bool = False) -> str:
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
    return STORAGE_ROOT / year_folder / folder / f"{owner}_{suffix}.xlsx"


def _ensure_storage_tree() -> None:
    STORAGE_ROOT.mkdir(parents=True, exist_ok=True)
    team_folders = [code for code in TEAM_OPTIONS.keys() if code]
    team_folders.extend(["UNSPECIFIED", "ANONYME"])
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
    mode_from_type = {"item": "item_unique", "personnel": "bilan_personnel", "projet": "bilan_projet"}.get(dossier_type, "")
    if mode_from_type:
        st.session_state.mode_input = mode_from_type

    project_name = str(meta.get("dossier_name", "")).strip()
    if project_name:
        st.session_state.project_name_input = project_name

    team_code = str(meta.get("team_code", "")).strip()
    if not team_code and entries:
        team_values = sorted({str(e.get("team_code", "")).strip() for e in entries if str(e.get("team_code", "")).strip()})
        if len(team_values) == 1:
            team_code = team_values[0]
    st.session_state.team_code_input = team_code if team_code in TEAM_OPTIONS else ""
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


def _save_local_form(
    save_path: Path,
    year_value: str,
    dossier_id: str,
    dossier_type: str,
    project_name: str,
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
    wanted = ["poste", "mission_id", "departure_city", "arrival_city", "departure_lat", "departure_lon", "arrival_lat", "arrival_lon"]
    cols = [c for c in wanted if c in df.columns]
    if len(cols) < 8:
        return pd.DataFrame()
    mdf = df[df["poste"] == "missions"][cols].copy()
    mdf = mdf.dropna(subset=["departure_lat", "departure_lon", "arrival_lat", "arrival_lon"])
    return mdf


def _mission_deck(arcs_df: pd.DataFrame, height: int = 320) -> pdk.Deck:
    center_lat = float(pd.concat([arcs_df["departure_lat"], arcs_df["arrival_lat"]]).mean())
    center_lon = float(pd.concat([arcs_df["departure_lon"], arcs_df["arrival_lon"]]).mean())
    layer = pdk.Layer(
        "ArcLayer",
        data=arcs_df,
        get_source_position=["departure_lon", "departure_lat"],
        get_target_position=["arrival_lon", "arrival_lat"],
        get_source_color=[30, 144, 255, 180],
        get_target_color=[220, 20, 60, 180],
        auto_highlight=True,
        pickable=True,
        get_width=3,
    )
    view_state = pdk.ViewState(latitude=center_lat, longitude=center_lon, zoom=3, pitch=30)
    tooltip = {"text": "{mission_id}\n{departure_city} -> {arrival_city}"}
    return pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        tooltip=tooltip,
        map_provider="carto",
        map_style="light",
        height=height,
    )


def _country_input(label: str, key_prefix: str, default_country: str = "FR") -> str:
    other_label = tf("Autre (saisie manuelle)", "Other (manual input)")
    country_options = [c for c in COUNTRY_OPTIONS if c != "Autre (saisie manuelle)"] + [other_label]
    default_idx = country_options.index(default_country) if default_country in country_options else 0
    selected = st.selectbox(label, country_options, index=default_idx, key=f"{key_prefix}_select")
    if selected == other_label:
        custom = st.text_input(f"{label} {tf('(saisie)', '(manual)')}", value="", key=f"{key_prefix}_custom")
        return _clean(custom)
    return selected


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
        st.altair_chart(bar, use_container_width=True)

    st.write("## " + tf("Vue d'ensemble du labo", "Lab overview"))
    storage_label = f"`{STORAGE_ROOT.name}/`"
    st.caption(
        tf(
            f"Consolidation des formulaires locaux sauvegardes dans le dossier {storage_label}",
            f"Consolidation of local forms saved in folder {storage_label}.",
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

    ag_df = entries_to_df(local_entries)
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
        st.altair_chart(yearly_total_chart, use_container_width=True)

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
            st.altair_chart(chart, use_container_width=True)

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

    poste_filter = st.selectbox(tf("Filtre poste", "Category filter"), options=["ALL", "achats", "domicile_travail", "campagnes_terrain", "missions", "heures_calcul"], key="ag_poste")
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

    m1, m2, m3 = st.columns(3)
    m1.metric(tf("Emissions", "Emissions"), f"{filtered['emissions_kgco2e'].sum():.2f} kgCO2e")
    m2.metric(tf("Formulaires", "Forms"), int(filtered["dossier_id"].nunique()) if not filtered.empty else 0)
    m3.metric(tf("Lignes", "Rows"), len(filtered))

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
    st.altair_chart(team_chart, use_container_width=True)
    _render_share_bar(by_team, "team", tf("Equipe", "Team"))

    st.write(f"### {tf('Emissions par poste', 'Emissions by category')}")
    poste_chart = alt.Chart(by_poste).mark_bar().encode(x=alt.X("poste:N", title=tf("Poste", "Category")), y=alt.Y("emissions_kgco2e:Q", title="kgCO2e"), tooltip=["poste", "emissions_kgco2e"]).properties(height=320)
    st.altair_chart(poste_chart, use_container_width=True)
    _render_share_bar(by_poste, "poste", tf("Poste", "Category"))

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
        st.altair_chart(poste_type_chart, use_container_width=True)
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
        st.altair_chart(employment_chart, use_container_width=True)
        _render_share_bar(by_employment, "meta_employment_type", tf("Statut", "Status"))

    st.write(f"### {tf('Detail', 'Detail')}")
    st.dataframe(filtered, use_container_width=True)

    st.write(f"### {tf('Synthese', 'Summary')}")
    st.dataframe(build_synthese(filtered), use_container_width=True)


_ensure_storage_tree()
_consume_pending_loaded_state()

mode = _normalize_mode_value(st.session_state.get("mode_input", "item_unique"))
st.session_state.mode_input = mode
dossier_type = DOSSIER_TYPE_BY_MODE.get(mode, "item")
dossier_id = st.session_state.get("dossier_id", "")
project_name = st.session_state.get("project_name_input", "")
is_anonymous = st.session_state.get("is_anonymous", True)
person_label = st.session_state.get("person_label_input", "")
team_code = st.session_state.get("team_code_input", "")
year_value = st.session_state.get("year_input", "2026")
poste_type = st.session_state.get("poste_type_input", "")

with st.sidebar:
    st.header(tf("Navigation", "Navigation"))
    top_view = st.radio(
        tf("Vue", "View"),
        ["formulaire", "labo_overview"],
        key="top_view_sidebar",
        format_func=lambda v: tf("Formulaire", "Form") if v == "formulaire" else tf("Vue d'ensemble du labo", "Lab overview"),
    )
    if st.session_state.loaded_notice:
        st.success(st.session_state.loaded_notice)
        st.session_state.loaded_notice = ""
    lcol, mcol, rcol = st.columns([1.2, 0.9, 1.2])
    lcol.markdown("lightmode ☀️")
    with mcol:
        dark_on = st.toggle(
            "theme",
            value=st.session_state.ui_theme == "darkmode",
            key="theme_toggle",
            label_visibility="collapsed",
        )
    rcol.markdown("🌙 darkmode")
    st.session_state.ui_theme = "darkmode" if dark_on else "lightmode"
    l2, m2, r2 = st.columns([1.0, 0.9, 1.0])
    l2.markdown("FR 🇫🇷")
    with m2:
        en_on = st.toggle(
            "lang",
            value=st.session_state.ui_lang == "EN",
            key="lang_toggle",
            label_visibility="collapsed",
        )
    r2.markdown("EN 🇬🇧")
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
        mode = st.selectbox(tf("Mode", "Mode"), MODE_OPTIONS, key="mode_input", format_func=_mode_label)
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
        if not st.session_state.dossier_id:
            default_id = datetime.now(timezone.utc).strftime("DOSSIER-%Y%m%d-%H%M%S")
            st.session_state.dossier_id = default_id
        dossier_id = st.session_state.dossier_id

        st.divider()
        st.subheader(tf("Identite", "Identity"))
        is_anonymous = st.session_state.get("is_anonymous", True)
        person_label = st.text_input(tf("Nom / identifiant", "Name / identifier"), key="person_label_input", disabled=is_anonymous)
        is_anonymous = st.checkbox(tf("Anonyme", "Anonymous"), value=is_anonymous, key="is_anonymous")
        team_code = st.selectbox(tf("Equipe (optionnel)", "Team (optional)"), options=list(TEAM_OPTIONS.keys()), format_func=lambda k: TEAM_OPTIONS[k], key="team_code_input")
        poste_options = ADMIN_GESTION_POSTE_TYPE_OPTIONS if team_code == "ADMIN_GESTION" else POSTE_TYPE_OPTIONS
        if st.session_state.get("poste_type_input", "") not in poste_options:
            st.session_state.poste_type_input = ""
        poste_type = st.selectbox(tf("Type de poste", "Position type"), options=poste_options, key="poste_type_input")
        status_info = _employment_type(team_code, poste_type)
        if status_info:
            st.caption(tf(f"Statut: {status_info}", f"Status: {status_info}"))

        st.divider()
        st.subheader(tf("Reprendre un dossier", "Reload a form"))
        saved_forms = _list_saved_forms()
        if saved_forms:
            selected_path = st.selectbox(
                tf("Dossiers locaux", "Local forms"),
                options=saved_forms,
                format_func=lambda p: str(p.relative_to(STORAGE_ROOT)),
            )
            if st.button(tf("Ouvrir dossier local", "Open local form")):
                file_bytes = selected_path.read_bytes()
                meta, imported_entries = import_excel_entries(file_bytes)
                st.session_state.pending_loaded_meta = meta
                st.session_state.pending_loaded_entries = imported_entries
                st.session_state.pending_loaded_path = str(selected_path)
                st.rerun()
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
        st.subheader(tf("Sauvegarde locale", "Local save"))
        if mode in SAVE_MODE_SUFFIX:
            if st.session_state.loaded_form_path:
                st.caption(f"{tf('Fichier courant', 'Current file')}: {Path(st.session_state.loaded_form_path).name}")
            if st.button(tf("Sauvegarder le dossier localement", "Save form locally")):
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
                        is_anonymous=is_anonymous,
                        owner_for_save=owner_for_save,
                        poste_type=poste_type,
                        team_code=team_code,
                        entries=st.session_state.entries,
                    )
                    st.session_state.loaded_form_path = str(save_path)
                    st.session_state.loaded_owner_label = owner_for_save
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
                        is_anonymous=is_anonymous,
                        owner_for_save=owner_for_save,
                        poste_type=poste_type,
                        team_code=team_code,
                        entries=st.session_state.entries,
                    )
                    st.session_state.loaded_owner_label = owner_for_save
                    st.session_state.confirm_overwrite_pending = False
                    st.success(f"{tf('Sauvegarde ok', 'Saved')}: {save_path.relative_to(Path(__file__).resolve().parent)}")
                if c_no.button(tf("non", "no"), key="confirm_overwrite_no"):
                    st.session_state.confirm_overwrite_pending = False
        else:
            st.caption(tf("Sauvegarde locale disponible uniquement pour bilan_personnel et bilan_projet.", "Local save is available only for bilan_personnel and bilan_projet."))

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
    st.markdown(
        css,
        unsafe_allow_html=True,
    )
else:
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
div[data-testid="stFileUploaderDropzone"]{
  background: @@FILE_DROP_BG@@ !important;
  border: 1px dashed @@FILE_DROP_BORDER@@ !important;
}
div[data-testid="stFileUploaderDropzone"] *{
  color: @@FILE_DROP_TEXT@@ !important;
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
    st.markdown(
        css,
        unsafe_allow_html=True,
    )

if top_view == "labo_overview":
    _render_labo_overview()
    st.stop()

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
            e_factor = st.number_input(
                tf("Facteur (kgCO2e / €)", "Factor (kgCO2e / €)"),
                min_value=0.0,
                value=float(old.get("factor_kgco2e_per_eur", achat_categories.get(default_cat, 0.3)) or 0.0),
                step=0.01,
                key=f"edit_achat_factor_{editing_idx}",
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
                    e_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
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
            e_days = st.number_input(tf("Jours/semaine", "Days/week"), min_value=0.0, max_value=7.0, value=float(old.get("days_per_week", 0.0) or 0.0), step=1.0, key=f"edit_dom_days_{editing_idx}")
            e_weeks = st.number_input(tf("Semaines/an", "Weeks/year"), min_value=0.0, max_value=52.0, value=float(old.get("weeks_per_year", 0.0) or 0.0), step=1.0, key=f"edit_dom_weeks_{editing_idx}")
            e_rt = st.checkbox(tf("Aller-retour", "Round trip"), value=bool(old.get("round_trip", True)), key=f"edit_dom_rt_{editing_idx}")
            e_factor = st.number_input(tf("Facteur (kgCO2e / km)", "Factor (kgCO2e / km)"), min_value=0.0, value=float(old.get("factor_kgco2e_per_km", 0.0) or 0.0), step=0.01, key=f"edit_dom_factor_{editing_idx}")
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
                    e_days,
                    e_weeks,
                    e_rt,
                    e_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
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
            e_factor = st.number_input(tf("Facteur (kgCO2e / km / personne)", "Factor (kgCO2e / km / person)"), min_value=0.0, value=float(old.get("factor_kgco2e_per_km_per_person", 0.0) or 0.0), step=0.01, key=f"edit_camp_factor_{editing_idx}")
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
                    e_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
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
                e_dep_country = _country_input(tf("Pays depart", "Departure country"), f"edit_dep_country_{editing_idx}", default_country=str(old.get("departure_country", "FR") or "FR"))
            with col_arr_e:
                e_arr_city = st.text_input(tf("Ville arrivee", "Arrival city"), value=str(old.get("arrival_city", "")), key=f"edit_mis_arr_city_{editing_idx}")
                e_arr_country = _country_input(tf("Pays arrivee", "Arrival country"), f"edit_arr_country_{editing_idx}", default_country=str(old.get("arrival_country", "FR") or "FR"))
            cur_mode = str(old.get("t_type", "plane"))
            if cur_mode not in {"plane", "train", "car"}:
                cur_mode = "plane"
            e_mode = st.selectbox(tf("Transport principal", "Main transport"), ["plane", "train", "car"], index=["plane", "train", "car"].index(cur_mode), key=f"edit_mis_mode_{editing_idx}")
            e_rt = st.checkbox(tf("Aller-retour", "Round trip"), value=bool(old.get("round_trip", True)), key=f"edit_mis_rt_{editing_idx}")
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
            e_factor = st.number_input(tf("Facteur (kgCO2e / kWh)", "Factor (kgCO2e / kWh)"), min_value=0.0, value=float(old.get("factor_kgco2e_per_kwh", 0.0) or 0.0), step=0.01, key=f"edit_comp_factor_{editing_idx}")
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
                    e_factor,
                )
                updated["record_id"] = old.get("record_id", updated.get("record_id"))
                updated["created_at"] = old.get("created_at", updated.get("created_at"))
                updated["ref_label"] = _clean(e_ref)
                st.session_state.entries[editing_idx] = updated
                st.session_state.editing_row_idx = None
                st.rerun()
            if ccancel.button(tf("Annuler", "Cancel"), key=f"cancel_edit_comp_{editing_idx}"):
                st.session_state.editing_row_idx = None
                st.rerun()

        st.divider()

    if editing_row is None and poste == "achats":
        achat_ref = st.text_input(tf("Reference", "Reference"), value=tf("Achat 1", "Purchase 1"))
        achat_categories = DEFAULT_FACTORS["achats_category_factors"]
        achat_category = st.selectbox(tf("Type d'achat", "Purchase type"), options=list(achat_categories.keys()))
        amount_eur = st.number_input(tf("Montant (€)", "Amount (€)"), min_value=0.0, value=1500.0, step=100.0)
        factor = st.number_input(
            tf("Facteur (kgCO2e / €)", "Factor (kgCO2e / €)"),
            min_value=0.0,
            value=float(achat_categories[achat_category]),
            step=0.01,
            key=f"achat_factor_{achat_category}",
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
                factor,
            )
            new_row["ref_label"] = _clean(achat_ref)

    elif editing_row is None and poste == "domicile_travail":
        domicile_ref = st.text_input(tf("Reference", "Reference"), value=tf("Trajet domicile", "Commute trip"))
        mode_t = st.selectbox(tf("Mode de transport", "Transport mode"), list(DEFAULT_FACTORS["domicile_mode_factors"].keys()))
        dist = st.number_input(tf("Distance aller simple (km)", "One-way distance (km)"), min_value=0.0, value=12.0, step=1.0)
        days = st.number_input(tf("Jours/semaine", "Days/week"), min_value=0.0, max_value=7.0, value=4.0, step=1.0)
        weeks = st.number_input(tf("Semaines/an", "Weeks/year"), min_value=0.0, max_value=52.0, value=45.0, step=1.0)
        round_trip = st.checkbox(tf("Aller-retour", "Round trip"), value=True)
        factor = st.number_input(
            tf("Facteur (kgCO2e / km)", "Factor (kgCO2e / km)"),
            min_value=0.0,
            value=float(DEFAULT_FACTORS["domicile_mode_factors"][mode_t]),
            step=0.01,
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
                days,
                weeks,
                round_trip,
                factor,
            )
            new_row["ref_label"] = _clean(domicile_ref)

    elif editing_row is None and poste == "campagnes_terrain":
        campaign_ref = st.text_input(tf("Reference", "Reference"), value=tf("Campagne 1", "Campaign 1"))
        campaign_label = st.text_input(tf("Nom campagne", "Campaign name"), value=tf("Campagne terrain 1", "Field campaign 1"))
        segment_mode = st.selectbox(tf("Mode de transport", "Transport mode"), list(DEFAULT_FACTORS["campagnes_mode_factors"].keys()))
        dist = st.number_input(tf("Distance segment (km)", "Segment distance (km)"), min_value=0.0, value=800.0, step=10.0)
        passengers = st.number_input(tf("Nombre de personnes", "Number of people"), min_value=1.0, value=1.0, step=1.0)
        round_trip = st.checkbox(tf("Aller-retour", "Round trip"), value=True, key="camp_rt")
        factor = st.number_input(
            tf("Facteur (kgCO2e / km / personne)", "Factor (kgCO2e / km / person)"),
            min_value=0.0,
            value=float(DEFAULT_FACTORS["campagnes_mode_factors"][segment_mode]),
            step=0.01,
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
                factor,
            )
            new_row["ref_label"] = _clean(campaign_ref)

    elif editing_row is None and poste == "missions":
        mission_ref = st.text_input(tf("Reference", "Reference"), value="MISSION-001")
        st.caption(tf("Calcul mission en direct via Moulinette_missions (geocodage + distance + facteurs)", "Direct mission computation via Moulinette_missions (geocoding + distance + factors)"))
        mission_id = st.text_input("Mission ID", value="MISSION-001")
        col_dep, col_arr = st.columns(2)
        with col_dep:
            departure_city = st.text_input(tf("Ville depart", "Departure city"), value="Aix-en-Provence")
            departure_country = _country_input(tf("Pays depart", "Departure country"), "dep_country", default_country="FR")
        with col_arr:
            arrival_city = st.text_input(tf("Ville arrivee", "Arrival city"), value="Paris")
            arrival_country = _country_input(tf("Pays arrivee", "Arrival country"), "arr_country", default_country="FR")

        mission_mode = st.selectbox(tf("Transport principal", "Main transport"), ["plane", "train", "car"])
        round_trip = st.checkbox(tf("Aller-retour", "Round trip"), value=True, key="mis_rt")

        if st.button(tf("Calculer / Ajouter", "Compute / Add"), key="add_mission_direct"):
            try:
                new_row = compute_single_mission_with_moulinette(
                    dossier_id=dossier_id,
                    dossier_type=dossier_type,
                    is_anonymous=is_anonymous,
                    team_code=team_code,
                    person_label=person_label,
                    mission_id=_clean(mission_id),
                    departure_city=_clean(departure_city),
                    departure_country=_clean(departure_country),
                    arrival_city=_clean(arrival_city),
                    arrival_country=_clean(arrival_country),
                    transport_mode=mission_mode,
                    round_trip=round_trip,
                )
                new_row["ref_label"] = _clean(mission_ref)
                st.success(
                    tf(
                        f"Distance: {new_row['distance_total_km']:.1f} km | Emissions: {new_row['emissions_kgco2e']:.2f} kgCO2e",
                        f"Distance: {new_row['distance_total_km']:.1f} km | Emissions: {new_row['emissions_kgco2e']:.2f} kgCO2e",
                    )
                )

                # For short/medium missions, compare train vs plane when user selected one of those modes.
                if mission_mode in {"plane", "train"} and float(new_row["distance_one_way_km"]) < 4000:
                    alt_mode = "train" if mission_mode == "plane" else "plane"
                    alt_row = compute_single_mission_with_moulinette(
                        dossier_id=dossier_id,
                        dossier_type=dossier_type,
                        is_anonymous=is_anonymous,
                        team_code=team_code,
                        person_label=person_label,
                        mission_id=_clean(mission_id),
                        departure_city=_clean(departure_city),
                        departure_country=_clean(departure_country),
                        arrival_city=_clean(arrival_city),
                        arrival_country=_clean(arrival_country),
                        transport_mode=alt_mode,
                        round_trip=round_trip,
                    )
                    delta_kg = float(new_row["emissions_kgco2e"]) - float(alt_row["emissions_kgco2e"])
                    pct = (delta_kg / float(alt_row["emissions_kgco2e"]) * 100) if float(alt_row["emissions_kgco2e"]) > 0 else 0.0
                    color = "red" if delta_kg > 0 else "green"
                    sign = "+" if delta_kg > 0 else ""
                    st.info(
                        tf(
                            f"Mode alternatif ({alt_mode}): {alt_row['emissions_kgco2e']:.2f} kgCO2e",
                            f"Alternative mode ({alt_mode}): {alt_row['emissions_kgco2e']:.2f} kgCO2e",
                        )
                    )
                    st.markdown(
                        tf(
                            f"""<span style="color:{color};font-weight:600;">Différentiel ({mission_mode} - {alt_mode}) : {sign}{delta_kg:.2f} kgCO2e ({sign}{pct:.1f}%)</span>""",
                            f"""<span style="color:{color};font-weight:600;">Difference ({mission_mode} - {alt_mode}): {sign}{delta_kg:.2f} kgCO2e ({sign}{pct:.1f}%)</span>""",
                        ),
                        unsafe_allow_html=True,
                    )
            except Exception as exc:  # pragma: no cover - UI path
                st.error(tf(f"Erreur moulinette: {exc}", f"Moulinette error: {exc}"))

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
        factor = st.number_input(
            tf("Facteur (kgCO2e / kWh)", "Factor (kgCO2e / kWh)"),
            min_value=0.0,
            value=float(DEFAULT_FACTORS["heures_calcul_kgco2e_per_kwh"]),
            step=0.01,
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
                factor,
            )
            new_row["ref_label"] = _clean(compute_ref)

    if new_row is not None:
        st.session_state.entries.append(new_row)
        st.success(tf(f"Ligne ajoutee: {new_row['poste']} -> {new_row['emissions_kgco2e']:.2f} kgCO2e", f"Row added: {new_row['poste']} -> {new_row['emissions_kgco2e']:.2f} kgCO2e"))

with col2:
    st.write(f"### {tf('Resultat dossier', 'Form result')}")
    df = entries_to_df(st.session_state.entries)
    total = float(df["emissions_kgco2e"].sum()) if not df.empty else 0.0
    st.metric(tf("Total", "Total"), f"{total:.2f} kgCO2e")
    st.metric(tf("Nombre de lignes", "Number of rows"), len(df))

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
df = entries_to_df(st.session_state.entries)
if df.empty:
    st.info(tf("Aucune ligne pour l'instant", "No row yet."))
else:
    if mode in {"bilan_personnel", "bilan_projet"}:
        view_df = df.reset_index(names="row_idx")
        h1, h2, h3, h4, h5, h6, h7, h8 = st.columns([1.2, 1.4, 1.2, 2.0, 1.4, 1.0, 0.8, 0.8])
        h1.write(tf("Poste", "Category"))
        h2.write("Ref")
        h3.write(tf("Equipe", "Team"))
        h4.write(tf("Personne", "Person"))
        h5.write(tf("Mis a jour le", "Updated at"))
        h6.write("kgCO2e")
        h7.write("")
        h8.write("")

        for _, row in view_df.iterrows():
            ref = row.get("ref_label") or row.get("mission_id") or row.get("item_label") or row.get("campaign_label") or row.get("compute_label") or row.get("record_id")
            c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([1.2, 1.4, 1.2, 2.0, 1.4, 1.0, 0.8, 0.8])
            c1.write(str(row.get("poste", "")))
            c2.write(str(ref))
            c3.write(str(row.get("team_code", "")))
            c4.write(str(row.get("person_label", "")))
            c5.write(str(row.get("updated_at", "")))
            c6.write(f"{float(row.get('emissions_kgco2e', 0)):.2f}")
            if c7.button(tf("Modifier", "Edit"), key=f"edit_row_{int(row['row_idx'])}_{row.get('record_id', '')}"):
                st.session_state.editing_row_idx = int(row["row_idx"])
                st.rerun()
            if c8.button(tf("Supprimer", "Delete"), key=f"delete_row_{int(row['row_idx'])}_{row.get('record_id', '')}_del"):
                del st.session_state.entries[int(row["row_idx"])]
                if st.session_state.get("editing_row_idx") == int(row["row_idx"]):
                    st.session_state.editing_row_idx = None
                st.rerun()
    else:
        show_cols = [
            c
            for c in [
                "record_id",
                "poste",
                "team_code",
                "person_label",
                "item_label",
                "mission_id",
                "campaign_label",
                "compute_label",
                "emissions_kgco2e",
                "updated_at",
            ]
            if c in df.columns
        ]
        st.dataframe(df[show_cols], use_container_width=True)

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
