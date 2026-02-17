from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd

from carbonometre.calculations import build_factors_df, build_synthese, entries_to_df
from carbonometre.constants import POSTES, SCHEMA_VERSION


def export_excel_bytes(meta: dict[str, Any], entries: list[dict[str, Any]]) -> bytes:
    df = entries_to_df(entries)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([meta]).to_excel(writer, index=False, sheet_name="meta")

        for poste in POSTES:
            sheet_df = df[df["poste"] == poste].copy() if not df.empty else pd.DataFrame()
            sheet_df.to_excel(writer, index=False, sheet_name=poste)

        build_factors_df().to_excel(writer, index=False, sheet_name="factors")
        build_synthese(df).to_excel(writer, index=False, sheet_name="synthese")

    return output.getvalue()


def import_excel_entries(file_bytes: bytes) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    meta = {}
    entries: list[dict[str, Any]] = []

    if "meta" in xls.sheet_names:
        meta_df = pd.read_excel(xls, "meta")
        if not meta_df.empty:
            meta = meta_df.iloc[0].to_dict()

    for poste in POSTES:
        if poste not in xls.sheet_names:
            continue
        df = pd.read_excel(xls, poste)
        if df.empty:
            continue
        rows = df.to_dict(orient="records")
        for row in rows:
            row["poste"] = poste
            entries.append(row)

    if "schema_version" not in meta:
        meta["schema_version"] = SCHEMA_VERSION

    return meta, entries
