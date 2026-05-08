import io
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Grundeinstellungen
# ---------------------------------------------------------------------------

DAY_NAMES = {
    1: "Montag",
    2: "Dienstag",
    3: "Mittwoch",
    4: "Donnerstag",
    5: "Freitag",
    6: "Samstag",
}

# Fallback-Positionen, falls Spaltenüberschriften nicht erkannt werden.
# Index 0 = Spalte A. In der Tourenplanung ist normalerweise:
# A CSB, B SAP, C Name, D Strasse, E Plz, F Ort, G Mo, H Die, I Mitt, J Don, K Fr, L Sam.
DAY_COLUMNS_TOUR = {
    1: 6,
    2: 7,
    3: 8,
    4: 9,
    5: 10,
    6: 11,
}

SAP_COL_INDEX = 0
SAP_DAY_COL_INDEX = 6
TOUR_SAP_COL_INDEX = 1

# Es werden nur diese vier Blätter der Tourenplanung geprüft.
TOUR_SHEET_CANDIDATES = [
    "DIREKT",
    "MK",
    "HUPA_NMS",
    "HUPA_MALCHOW",
]

DAY_COLUMN_CANDIDATES = {
    1: ["mo", "montag"],
    2: ["die", "di", "dienstag"],
    3: ["mitt", "mit", "mi", "mittwoch"],
    4: ["don", "do", "donnerstag"],
    5: ["fr", "frei", "freitag"],
    6: ["sam", "sa", "samstag"],
}

# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------


def value_to_clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_sap_series(series: pd.Series) -> pd.Series:
    """Normalisiert SAP-Nummern. Aus 12345.0 wird 12345."""
    if series.empty:
        return series.astype(str)

    result = series.copy()
    numeric = pd.to_numeric(result, errors="coerce")
    is_int = numeric.notna() & (numeric == numeric.round())

    out = result.astype(str)
    out = out.where(~is_int, numeric.where(is_int).astype("Int64").astype(str))
    out = out.str.strip()
    out = out.replace({"nan": "", "<NA>": "", "None": ""})
    return out


def normalize_header_name(value) -> str:
    """Vereinfacht Spaltenüberschriften für robuste Erkennung."""
    text = "" if value is None or pd.isna(value) else str(value)
    text = text.strip().lower()
    text = (
        text.replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )
    return "".join(ch for ch in text if ch.isalnum())


def normalized_candidates(values: List[str]) -> List[str]:
    return [normalize_header_name(value) for value in values]


def pick_first_matching_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    candidate_set = set(candidates)
    for column in columns:
        if normalize_header_name(column) in candidate_set:
            return column
    return None


def pick_column_by_name_or_position(
    columns: List[str],
    candidates: List[str],
    fallback_index: Optional[int] = None,
) -> Optional[str]:
    found = pick_first_matching_column(columns, normalized_candidates(candidates))
    if found is not None:
        return found
    if fallback_index is not None and len(columns) > fallback_index:
        return columns[fallback_index]
    return None


def make_unique_columns(raw_columns: List[object]) -> List[str]:
    """Erzeugt eindeutige Spaltennamen, auch wenn Excel leere oder gleiche Überschriften enthält."""
    result: List[str] = []
    seen: Dict[str, int] = {}
    for index, value in enumerate(raw_columns, start=1):
        name = value_to_clean_text(value)
        if not name:
            name = f"Spalte_{index}"
        count = seen.get(name, 0) + 1
        seen[name] = count
        if count > 1:
            name = f"{name}_{count}"
        result.append(name)
    return result


def read_excel_with_detected_header(excel: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """Liest ein Blatt und erkennt eine Kopfzeile mit SAP und Tages-Spalten selbst."""
    raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=object)
    if raw.empty:
        return pd.DataFrame()

    header_row: Optional[int] = None
    max_scan_rows = min(len(raw), 25)
    day_names_flat = {
        normalize_header_name(candidate)
        for values in DAY_COLUMN_CANDIDATES.values()
        for candidate in values
    }

    for row_index in range(max_scan_rows):
        values = [normalize_header_name(value) for value in raw.iloc[row_index].tolist()]
        value_set = set(values)
        has_sap = "sap" in value_set or "sapnummer" in value_set or "sapnr" in value_set
        day_hits = sum(1 for value in values if value in day_names_flat)
        if has_sap and day_hits >= 2:
            header_row = row_index
            break

    if header_row is None:
        return pd.read_excel(excel, sheet_name=sheet_name, header=0, dtype=object)

    df = raw.iloc[header_row + 1:].copy()
    df.columns = make_unique_columns(raw.iloc[header_row].tolist())
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def select_tour_sheet_names(excel: pd.ExcelFile) -> Tuple[List[str], List[str]]:
    """Wählt die vier echten Blätter der Tourenplanung aus."""
    available_by_normalized = {normalize_header_name(name): name for name in excel.sheet_names}
    selected: List[str] = []
    missing: List[str] = []

    for expected in TOUR_SHEET_CANDIDATES:
        real_name = available_by_normalized.get(normalize_header_name(expected))
        if real_name and real_name not in selected:
            selected.append(real_name)
        else:
            missing.append(expected)

    if selected:
        return selected, missing

    # Fallback: falls die Datei anders benannt wurde, werden die ersten vier Blätter geprüft.
    return excel.sheet_names[:4], TOUR_SHEET_CANDIDATES


def day_value_is_set(value) -> bool:
    """Ein Tag gilt als vorhanden, wenn in Mo bis Sam ein echter Wert steht."""
    if value is None or pd.isna(value):
        return False
    text = str(value).strip()
    if not text:
        return False
    if text.lower() in {"nan", "none", "<na>", "-", "--"}:
        return False
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(number) and float(number) == 0:
        return False
    return True


def normalize_day_code_series(series: pd.Series) -> pd.Series:
    """Normalisiert Liefertage aus SAP: 1 bis 6 oder Text wie Montag/Mo."""
    numeric = pd.to_numeric(series, errors="coerce")
    text = series.astype(str).map(normalize_header_name)
    text_map = {
        "1": 1, "mo": 1, "montag": 1,
        "2": 2, "di": 2, "die": 2, "dienstag": 2,
        "3": 3, "mi": 3, "mitt": 3, "mit": 3, "mittwoch": 3,
        "4": 4, "do": 4, "don": 4, "donnerstag": 4,
        "5": 5, "fr": 5, "frei": 5, "freitag": 5,
        "6": 6, "sa": 6, "sam": 6, "samstag": 6,
    }
    mapped = text.map(text_map)
    return numeric.where(numeric.notna(), mapped)


def merge_customer_info(base: Dict[str, Dict[str, str]], sap: str, info: Dict[str, str]) -> None:
    target = base.setdefault(sap, {"name": "", "strasse": "", "ort": ""})
    for key in ["name", "strasse", "ort"]:
        if not target.get(key) and info.get(key):
            target[key] = info[key]


# ---------------------------------------------------------------------------
# Dateien lesen
# ---------------------------------------------------------------------------


def read_sap_file(uploaded_file) -> Tuple[Dict[str, Set[int]], str, int]:
    """Liest die SAP-Datei. Fallback: Spalte A = SAP, Spalte G = Liefertag."""
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = read_excel_with_detected_header(excel, sheet_name)

    if df.empty:
        return {}, sheet_name, 0

    columns = list(df.columns)
    sap_column = pick_column_by_name_or_position(
        columns,
        ["SAP", "SAP Nummer", "SAP-Nr", "SAP Nr", "Kundennummer", "Kunden Nummer"],
        SAP_COL_INDEX,
    )
    day_column = pick_column_by_name_or_position(
        columns,
        ["Liefertag", "Liefer Tag", "LT", "Tag", "Liefertag Code", "Liefertagcode"],
        SAP_DAY_COL_INDEX,
    )

    if sap_column is None or day_column is None:
        return {}, sheet_name, 0

    work = df[[sap_column, day_column]].copy()
    work.columns = ["sap", "tag"]
    work["sap"] = normalize_sap_series(work["sap"])
    work["tag_num"] = normalize_day_code_series(work["tag"])

    mask = (
        work["sap"].ne("")
        & work["tag_num"].between(1, 6, inclusive="both")
        & work["tag_num"].notna()
    )
    filtered = work.loc[mask, ["sap", "tag_num"]].copy()
    filtered["tag_int"] = filtered["tag_num"].astype(int)

    days_by_sap: Dict[str, Set[int]] = (
        filtered.groupby("sap")["tag_int"].agg(set).to_dict()
    )

    return days_by_sap, sheet_name, len(filtered)


def read_tourenplanung(uploaded_file) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, Dict[str, str]]]:
    """Liest nur die vier Blätter DIREKT, MK, HUPA_NMS und HUPA_MALCHOW."""
    excel = pd.ExcelFile(uploaded_file)
    sheet_names, missing_sheet_names = select_tour_sheet_names(excel)

    frames: List[pd.DataFrame] = []
    customer_info: Dict[str, Dict[str, str]] = {}

    for sheet_name in sheet_names:
        df = read_excel_with_detected_header(excel, sheet_name)
        if df.empty:
            continue

        columns = list(df.columns)
        sap_column = pick_column_by_name_or_position(
            columns,
            ["SAP", "SAP Nummer", "SAP-Nr", "SAP Nr", "Kundennummer", "Kunden Nummer"],
            TOUR_SAP_COL_INDEX,
        )
        if sap_column is None:
            continue

        csb_column = pick_column_by_name_or_position(columns, ["CSB", "CSB Nummer", "CSB-Nr", "CSB Nr"], 0)
        name_column = pick_column_by_name_or_position(
            columns,
            ["Name", "Kundenname", "Marktname", "Kunde", "Bezeichnung", "Filialname"],
            2,
        )
        strasse_column = pick_column_by_name_or_position(
            columns,
            ["Strasse", "Straße", "Str", "Anschrift", "Adresse", "Strassenname", "Straßenname", "Strasse Hausnummer", "Straße Hausnummer"],
            3,
        )
        plz_column = pick_column_by_name_or_position(columns, ["Plz", "PLZ", "Postleitzahl"], 4)
        ort_column = pick_column_by_name_or_position(columns, ["Ort", "Stadt", "Plz Ort", "PLZ Ort", "Ortname"], 5)

        rename_map = {sap_column: "sap"}
        if csb_column and csb_column != sap_column:
            rename_map[csb_column] = "csb"

        for day_num, col_index in DAY_COLUMNS_TOUR.items():
            day_column = pick_column_by_name_or_position(
                columns,
                DAY_COLUMN_CANDIDATES[day_num],
                col_index,
            )
            if day_column and day_column != sap_column:
                rename_map[day_column] = f"tag_{day_num}"

        if name_column and name_column != sap_column:
            rename_map[name_column] = "name"
        if strasse_column and strasse_column != sap_column:
            rename_map[strasse_column] = "strasse"
        if ort_column and ort_column != sap_column:
            rename_map[ort_column] = "ort"
        if plz_column and plz_column != sap_column:
            rename_map[plz_column] = "plz"

        work = df.rename(columns=rename_map).copy()
        work["sap"] = normalize_sap_series(work["sap"])
        work = work[work["sap"].ne("")].copy()
        if work.empty:
            continue

        for column in ["name", "strasse", "ort", "plz"]:
            if column not in work.columns:
                work[column] = ""

        info_df = work[["sap", "name", "strasse", "ort", "plz"]].copy()
        info_df["name"] = info_df["name"].map(value_to_clean_text)
        info_df["strasse"] = info_df["strasse"].map(value_to_clean_text)
        info_df["ort"] = info_df["ort"].map(value_to_clean_text)
        info_df["plz"] = info_df["plz"].map(value_to_clean_text)
        info_df["ort_kombi"] = info_df.apply(
            lambda row: " ".join(v for v in [row["plz"], row["ort"]] if v).strip(),
            axis=1,
        )

        for _, row in info_df.iterrows():
            merge_customer_info(
                customer_info,
                row["sap"],
                {
                    "name": row["name"],
                    "strasse": row["strasse"],
                    "ort": row["ort_kombi"] or row["ort"],
                },
            )

        day_value_columns = list(dict.fromkeys(
            f"tag_{d}" for d in DAY_COLUMNS_TOUR.keys() if f"tag_{d}" in work.columns
        ))
        if not day_value_columns:
            continue

        work["blatt"] = sheet_name
        long = work.melt(
            id_vars=["sap", "blatt"],
            value_vars=day_value_columns,
            var_name="tag_col",
            value_name="wert",
        )
        long["tag_num"] = long["tag_col"].str.replace("tag_", "", regex=False).astype(int)
        long["wert_gesetzt"] = long["wert"].map(day_value_is_set)

        long = long[long["sap"].ne("") & long["wert_gesetzt"]]
        frames.append(long[["sap", "blatt", "tag_num", "wert"]])

    if not frames:
        empty = pd.DataFrame(columns=["sap", "blatt", "tag_num", "wert"])
        return empty, sheet_names, missing_sheet_names, customer_info

    return pd.concat(frames, ignore_index=True), sheet_names, missing_sheet_names, customer_info


# ---------------------------------------------------------------------------
# Vergleich
# ---------------------------------------------------------------------------


def days_to_text(days: Set[int] | List[int]) -> str:
    return ", ".join(f"{d} {DAY_NAMES[d]}" for d in sorted(days))


def _export_columns() -> List[str]:
    return [
        "Blatt Tourenplanung",
        "SAP Nummer",
        "Name",
        "Straße",
        "Ort",
        "Fehlende LT",
        "LT SAP",
        "LT Tourenplanung",
    ]


def _empty_result_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_export_columns())


def build_missing_in_sap(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """Tage stehen in der Tourenplanung, fehlen aber in SAP."""
    if tour_df.empty:
        return _empty_result_df()

    known = tour_df[tour_df["sap"].ne("")].copy()
    if known.empty:
        return _empty_result_df()

    known["fehlt"] = known.apply(
        lambda row: row["tag_num"] not in days_by_sap.get(row["sap"], set()),
        axis=1,
    )
    missing = known[known["fehlt"]]
    if missing.empty:
        return _empty_result_df()

    days_in_tour: Dict[str, Set[int]] = tour_df.groupby("sap")["tag_num"].agg(set).to_dict()
    sheets_by_sap: Dict[str, str] = tour_df.groupby("sap")["blatt"].agg(
        lambda x: ", ".join(sorted(set(map(str, x))))
    ).to_dict()

    agg = missing.groupby("sap", as_index=False).agg(
        tage=("tag_num", lambda x: sorted(set(x))),
    )

    agg["Blatt Tourenplanung"] = agg["sap"].map(lambda s: sheets_by_sap.get(s, ""))
    agg["Name"] = agg["sap"].map(lambda s: customer_info.get(s, {}).get("name", ""))
    agg["Straße"] = agg["sap"].map(lambda s: customer_info.get(s, {}).get("strasse", ""))
    agg["Ort"] = agg["sap"].map(lambda s: customer_info.get(s, {}).get("ort", ""))
    agg["Fehlende LT"] = agg["tage"].map(days_to_text)
    agg["LT SAP"] = agg["sap"].map(lambda s: days_to_text(days_by_sap.get(s, set())) or "(keine hinterlegt)")
    agg["LT Tourenplanung"] = agg["sap"].map(lambda s: days_to_text(days_in_tour.get(s, set())))
    agg["_SapSort"] = pd.to_numeric(agg["sap"], errors="coerce").fillna(9_999_999_999)

    agg = agg.rename(columns={"sap": "SAP Nummer"}).sort_values(
        ["Blatt Tourenplanung", "_SapSort"]
    ).reset_index(drop=True)

    return agg[_export_columns()]


def build_missing_in_tour(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """Tage stehen in SAP, fehlen aber in der Tourenplanung."""
    days_in_tour: Dict[str, Set[int]] = {}
    sheets_by_sap: Dict[str, str] = {}

    if not tour_df.empty:
        days_in_tour = tour_df.groupby("sap")["tag_num"].agg(set).to_dict()
        sheets_by_sap = tour_df.groupby("sap")["blatt"].agg(
            lambda x: ", ".join(sorted(set(map(str, x))))
        ).to_dict()

    rows: List[dict] = []
    for sap, sap_days in days_by_sap.items():
        tour_days = days_in_tour.get(sap, set())
        fehlend = sorted(sap_days - tour_days)
        if not fehlend:
            continue

        info = customer_info.get(sap, {})
        rows.append({
            "Blatt Tourenplanung": sheets_by_sap.get(sap, "(nicht in Tourenplanung vorhanden)"),
            "SAP Nummer": sap,
            "Name": info.get("name", ""),
            "Straße": info.get("strasse", ""),
            "Ort": info.get("ort", ""),
            "Fehlende LT": days_to_text(fehlend),
            "LT SAP": days_to_text(sap_days),
            "LT Tourenplanung": days_to_text(tour_days) or "(nicht in Tourenplanung vorhanden)",
            "_SapSort": int(sap) if sap.isdigit() else 9_999_999_999,
        })

    if not rows:
        return _empty_result_df()

    df = pd.DataFrame(rows)
    df = df.sort_values(["Blatt Tourenplanung", "_SapSort"]).reset_index(drop=True)
    return df[_export_columns()]


def _add_count_column(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return df

    out = df.copy()
    if "Fehlende LT" in out.columns:
        out["Anzahl LT"] = out["Fehlende LT"].fillna("").map(
            lambda s: len([t for t in str(s).split(",") if t.strip()])
        )
    else:
        out["Anzahl LT"] = 0

    cols = list(out.columns)
    cols.remove("Anzahl LT")
    if "SAP Nummer" in cols:
        idx = cols.index("SAP Nummer") + 1
        cols.insert(idx, "Anzahl LT")
    else:
        cols.insert(0, "Anzahl LT")
    return out[cols]


def _filter_dataframe(df: pd.DataFrame, suche: str, blatt: str = "Alle") -> pd.DataFrame:
    if df is None or df.empty:
        return df

    work = df
    if blatt and blatt != "Alle" and "Blatt Tourenplanung" in work.columns:
        work = work[work["Blatt Tourenplanung"].astype(str).str.contains(blatt, na=False, regex=False)]

    if suche:
        such = suche.strip().lower()
        if such:
            spalten = [c for c in ["Blatt Tourenplanung", "SAP Nummer", "Name", "Straße", "Ort"] if c in work.columns]
            mask = pd.Series(False, index=work.index)
            for c in spalten:
                mask = mask | work[c].astype(str).str.lower().str.contains(such, na=False)
            work = work[mask]

    return work


def build_overview(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for sheet_name in TOUR_SHEET_CANDIDATES:
        rows.append({
            "Blatt": sheet_name,
            "Fehlt in SAP": int(missing_sap["Blatt Tourenplanung"].astype(str).str.contains(sheet_name, na=False, regex=False).sum()) if not missing_sap.empty else 0,
            "Fehlt in Tour": int(missing_tour["Blatt Tourenplanung"].astype(str).str.contains(sheet_name, na=False, regex=False).sum()) if not missing_tour.empty else 0,
        })

    rows.append({
        "Blatt": "Nicht in Tourenplanung vorhanden",
        "Fehlt in SAP": 0,
        "Fehlt in Tour": int(missing_tour["Blatt Tourenplanung"].astype(str).str.contains("nicht in Tourenplanung", case=False, na=False, regex=False).sum()) if not missing_tour.empty else 0,
    })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Excel-Erzeugung
# ---------------------------------------------------------------------------


def build_excel(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame) -> bytes:
    """Schreibt eine Excel mit allen Unterschieden und beiden Einzelrichtungen."""
    output = io.BytesIO()

    missing_sap_out = _add_count_column(missing_sap)
    missing_tour_out = _add_count_column(missing_tour)

    all_parts = []
    if missing_sap_out is not None and not missing_sap_out.empty:
        part = missing_sap_out.copy()
        part.insert(0, "Prüfung", "Fehlt in SAP")
        all_parts.append(part)
    if missing_tour_out is not None and not missing_tour_out.empty:
        part = missing_tour_out.copy()
        part.insert(0, "Prüfung", "Fehlt in Tour")
        all_parts.append(part)

    if all_parts:
        all_out = pd.concat(all_parts, ignore_index=True)
    else:
        all_out = pd.DataFrame(columns=["Prüfung"] + list(_add_count_column(_empty_result_df()).columns))

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_out.to_excel(writer, index=False, sheet_name="Alle Unterschiede", na_rep="")
        _format_sheet(writer, "Alle Unterschiede", all_out)

        missing_sap_out.to_excel(writer, index=False, sheet_name="Fehlt in SAP", na_rep="")
        _format_sheet(writer, "Fehlt in SAP", missing_sap_out)

        missing_tour_out.to_excel(writer, index=False, sheet_name="Fehlt in Tour", na_rep="")
        _format_sheet(writer, "Fehlt in Tour", missing_tour_out)

        wb = writer.book
        for ws in wb.worksheets:
            ws.sheet_state = "visible"

    return output.getvalue()


_RIGHT_ALIGN_COLS = {"SAP Nummer", "Anzahl LT"}

_COL_WIDTH_HINTS = {
    "Prüfung": (14, 18),
    "Blatt Tourenplanung": (18, 34),
    "SAP Nummer": (10, 12),
    "Anzahl LT": (10, 11),
    "Name": (24, 42),
    "Straße": (20, 32),
    "Ort": (22, 36),
    "Fehlende LT": (24, 48),
    "LT SAP": (24, 48),
    "LT Tourenplanung": (24, 48),
}


def _format_sheet(writer, sheet_name: str, df: pd.DataFrame) -> None:
    if df is None:
        return

    ws = writer.sheets[sheet_name]
    n_rows = len(df)
    n_cols = len(df.columns)

    header_fill = PatternFill(start_color="FF2F3A4A", end_color="FF2F3A4A", fill_type="solid")
    zebra_fill = PatternFill(start_color="FFF4F6F8", end_color="FFF4F6F8", fill_type="solid")

    thin = Side(style="thin", color="FFD7DEE8")
    medium = Side(style="medium", color="FF8A94A6")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    body_font = Font(name="Calibri", size=11)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=False)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_left
        cell.border = border_thin
    ws.row_dimensions[1].height = 24

    columns = list(df.columns)
    group_column = "Prüfung" if "Prüfung" in columns else "Blatt Tourenplanung"
    group_idx = columns.index(group_column) if group_column in columns else None
    group_values = df[group_column].astype(str).tolist() if group_idx is not None and n_rows > 0 else []

    for row_offset in range(n_rows):
        excel_row = row_offset + 2
        is_zebra = (row_offset % 2) == 1
        new_group = False
        if group_idx is not None and row_offset > 0:
            new_group = group_values[row_offset] != group_values[row_offset - 1]

        ws.row_dimensions[excel_row].height = 20

        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.font = body_font

            if col_name == "Anzahl LT":
                cell.alignment = align_center
            elif col_name in _RIGHT_ALIGN_COLS:
                cell.alignment = align_right
            else:
                cell.alignment = align_left

            if is_zebra:
                cell.fill = zebra_fill

            top_side = medium if new_group else thin
            cell.border = Border(left=thin, right=thin, top=top_side, bottom=thin)

    for col_idx, col_name in enumerate(columns, start=1):
        sample = df[col_name].astype(str).head(300).tolist() if col_name in df.columns else []
        max_len = max([len(str(col_name))] + [len(v) for v in sample] + [8])
        min_w, max_w = _COL_WIDTH_HINTS.get(col_name, (12, 50))
        width = min(max(max_len + 3, min_w), max_w)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    if n_rows > 0 and n_cols > 0:
        last_col = get_column_letter(n_cols)
        ws.auto_filter.ref = f"A1:{last_col}{n_rows + 1}"

    if "Anzahl LT" in columns and n_rows > 0:
        from openpyxl.formatting.rule import ColorScaleRule
        col_letter = get_column_letter(columns.index("Anzahl LT") + 1)
        rng = f"{col_letter}2:{col_letter}{n_rows + 1}"
        rule = ColorScaleRule(
            start_type="num", start_value=1, start_color="FFE2F0D9",
            mid_type="num", mid_value=3, mid_color="FFFFF2CC",
            end_type="num", end_value=6, end_color="FFF4B084",
        )
        ws.conditional_formatting.add(rng, rule)

    try:
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.print_options.gridLines = False
        if n_rows > 0:
            ws.print_title_rows = "1:1"
    except Exception:
        pass

    ws.sheet_state = "visible"


# ---------------------------------------------------------------------------
# Streamlit-Oberfläche
# ---------------------------------------------------------------------------


st.set_page_config(page_title="SAP ↔ Tourenplanung", layout="wide")

st.markdown(
    """
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1280px; }
        [data-testid="stMetric"] { background: #f7f8fa; border: 1px solid #e3e7ee; border-radius: 14px; padding: 14px 16px; }
        [data-testid="stFileUploader"] section { border: 1px dashed #b9c0cc; border-radius: 14px; background: #fafbfc; }
        div.stButton > button { border-radius: 12px; height: 3rem; font-weight: 700; }
        .small-note { color: #586174; font-size: 0.92rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SAP ↔ Tourenplanung")
st.markdown(
    """
    <div class="small-note">
    Prüft immer beide Richtungen: <b>Tourenplanung → SAP</b> und <b>SAP → Tourenplanung</b>.
    In der Tourenplanung werden nur die vier Blätter <b>DIREKT</b>, <b>MK</b>, <b>HUPA_NMS</b> und <b>HUPA_MALCHOW</b> gelesen.
    Es gibt keine fest hinterlegte Kundenliste mehr.
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()

upload_left, upload_right = st.columns(2)
with upload_left:
    sap_datei = st.file_uploader(
        "SAP-Datei",
        help="Erwartet: SAP Nummer und Liefertag 1 bis 6. Fallback: Spalte A = SAP Nummer, Spalte G = Liefertag.",
        type=["xlsx", "xlsm", "xls"],
        key="sap_datei",
    )

with upload_right:
    tourenplanung_datei = st.file_uploader(
        "Tourenplanung",
        help="Geprüft werden: DIREKT, MK, HUPA_NMS und HUPA_MALCHOW. Erwartet: SAP, Name, Straße, Ort, Mo bis Sam.",
        type=["xlsx", "xlsm", "xls"],
        key="tourenplanung_datei",
    )

run = st.button("Unterschiede prüfen und Excel erzeugen", type="primary", use_container_width=True)

if run:
    if not sap_datei or not tourenplanung_datei:
        st.error("Bitte SAP-Datei und Tourenplanung hochladen.")
        st.stop()

    try:
        days_by_sap, sap_sheet, sap_rows = read_sap_file(sap_datei)
        tour_df, tour_sheets, missing_tour_sheets, customer_info = read_tourenplanung(tourenplanung_datei)

        if sap_rows == 0:
            st.warning("In der SAP-Datei wurden keine gültigen Liefertage erkannt.")
        if tour_df.empty:
            st.warning(
                "In der Tourenplanung wurden keine gesetzten Liefertage erkannt. "
                f"Geprüfte Blätter: {', '.join(tour_sheets)}."
            )
        if missing_tour_sheets:
            st.warning("Nicht alle erwarteten Blätter wurden gefunden: " + ", ".join(missing_tour_sheets))

        missing_sap = build_missing_in_sap(tour_df, days_by_sap, customer_info)
        missing_tour = build_missing_in_tour(tour_df, days_by_sap, customer_info)
        excel_bytes = build_excel(missing_sap, missing_tour)

        all_count = len(missing_sap) + len(missing_tour)

        st.session_state["result"] = {
            "missing_sap": missing_sap,
            "missing_tour": missing_tour,
            "excel_bytes": excel_bytes,
            "sap_sheet": sap_sheet,
            "sap_rows": sap_rows,
            "tour_sheets": tour_sheets,
            "tour_rows": len(tour_df),
            "all_count": all_count,
        }
    except Exception as exc:
        import traceback
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
        with st.expander("Technische Details", expanded=False):
            st.code(traceback.format_exc(), language="python")
        st.session_state.pop("result", None)


# ---------------------------------------------------------------------------
# Ergebnisanzeige
# ---------------------------------------------------------------------------


result = st.session_state.get("result")
if result:
    missing_sap = result["missing_sap"]
    missing_tour = result["missing_tour"]

    st.divider()

    head_left, head_right = st.columns([3, 1])
    with head_left:
        st.subheader("Ergebnis")
        st.caption(
            f"SAP: Blatt {result['sap_sheet']}, {result['sap_rows']} Liefertage übernommen · "
            f"Tourenplanung: {', '.join(result['tour_sheets'])}, {result.get('tour_rows', 0)} gesetzte Liefertage erkannt"
        )
    with head_right:
        st.download_button(
            label="Excel herunterladen",
            data=result["excel_bytes"],
            file_name="sap_tourenplanung_alle_unterschiede.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    m1, m2, m3 = st.columns(3)
    m1.metric("Alle Unterschiede", result["all_count"])
    m2.metric("Fehlt in SAP", len(missing_sap))
    m3.metric("Fehlt in Tour", len(missing_tour))

    tabs = st.tabs([
        "Übersicht",
        f"Alle Unterschiede ({result['all_count']})",
        f"Fehlt in SAP ({len(missing_sap)})",
        f"Fehlt in Tour ({len(missing_tour)})",
    ])

    with tabs[0]:
        overview = build_overview(missing_sap, missing_tour)
        st.dataframe(overview, use_container_width=True, hide_index=True)

        if result["all_count"] == 0:
            st.success("Keine Unterschiede gefunden.")

    combined_parts = []
    if not missing_sap.empty:
        part = _add_count_column(missing_sap).copy()
        part.insert(0, "Prüfung", "Fehlt in SAP")
        combined_parts.append(part)
    if not missing_tour.empty:
        part = _add_count_column(missing_tour).copy()
        part.insert(0, "Prüfung", "Fehlt in Tour")
        combined_parts.append(part)
    combined = pd.concat(combined_parts, ignore_index=True) if combined_parts else pd.DataFrame()

    with tabs[1]:
        if combined.empty:
            st.info("Keine Treffer.")
        else:
            suche = st.text_input("Suchen", key="filter_all_suche", placeholder="SAP Nummer, Name, Straße, Ort oder Blatt")
            filtered = _filter_dataframe(combined, suche)
            st.caption(f"{len(filtered)} von {len(combined)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[2]:
        if missing_sap.empty:
            st.info("Keine Treffer: Es gibt keine Tage in der Tourenplanung, die in SAP fehlen.")
        else:
            f1, f2 = st.columns([1, 2])
            blatt = f1.selectbox("Blatt", ["Alle"] + TOUR_SHEET_CANDIDATES, key="filter_sap_blatt")
            suche = f2.text_input("Suchen", key="filter_sap_suche", placeholder="SAP Nummer, Name, Straße oder Ort")
            filtered = _add_count_column(_filter_dataframe(missing_sap, suche, blatt))
            st.caption(f"{len(filtered)} von {len(missing_sap)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[3]:
        if missing_tour.empty:
            st.info("Keine Treffer: Es gibt keine Tage in SAP, die in der Tourenplanung fehlen.")
        else:
            f1, f2 = st.columns([1, 2])
            blatt = f1.selectbox("Blatt", ["Alle"] + TOUR_SHEET_CANDIDATES + ["nicht in Tourenplanung"], key="filter_tour_blatt")
            suche = f2.text_input("Suchen", key="filter_tour_suche", placeholder="SAP Nummer, Name, Straße oder Ort")
            filtered = _add_count_column(_filter_dataframe(missing_tour, suche, blatt))
            st.caption(f"{len(filtered)} von {len(missing_tour)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)
