import io
import re
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

# Es werden nur diese Bereiche geprüft:
# - NMS komplett
# - Malchow komplett
# - Direkt nur mit den genannten Tournummern
DIRECT_TOURS: Set[str] = {"1058", "2058", "3058", "4058", "5058", "6030"}

# Die Direkt-Tournummern sind an den Liefertag gekoppelt.
# Dadurch wird bei "Fehlt in Tour" nicht nur die SAP-Nummer,
# sondern auch die erwartete Tournummer für genau diesen Liefertag ausgegeben.
DIRECT_TOUR_BY_DAY: Dict[int, str] = {
    1: "1058",
    2: "2058",
    3: "3058",
    4: "4058",
    5: "5058",
    6: "6030",
}

TOUR_SHEET_RULES = [
    {
        "bereich": "NMS",
        "candidates": ["NMS", "HUPA_NMS", "HUPA NMS", "HUPA-NMS", "HuPa NMS"],
        "direct_filter": False,
    },
    {
        "bereich": "Malchow",
        "candidates": ["Malchow", "HUPA_MALCHOW", "HUPA MALCHOW", "HUPA-MALCHOW", "HuPa Malchow"],
        "direct_filter": False,
    },
    {
        "bereich": "Direkt",
        "candidates": ["DIREKT", "Direkt"],
        "direct_filter": True,
    },
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


def select_tour_sheet_rules(excel: pd.ExcelFile) -> Tuple[List[dict], List[str]]:
    """Findet die Blätter NMS, Malchow und Direkt anhand robuster Namensvarianten."""
    available_by_normalized = {normalize_header_name(name): name for name in excel.sheet_names}
    selected: List[dict] = []
    missing: List[str] = []

    for rule in TOUR_SHEET_RULES:
        real_name: Optional[str] = None
        for candidate in rule["candidates"]:
            normalized = normalize_header_name(candidate)
            if normalized in available_by_normalized:
                real_name = available_by_normalized[normalized]
                break

        # Fallback: enthält der Blattname den Suchbegriff?
        if real_name is None:
            for normalized_name, original_name in available_by_normalized.items():
                if rule["bereich"] == "NMS" and "nms" in normalized_name:
                    real_name = original_name
                    break
                if rule["bereich"] == "Malchow" and "malchow" in normalized_name:
                    real_name = original_name
                    break
                if rule["bereich"] == "Direkt" and normalized_name == "direkt":
                    real_name = original_name
                    break

        if real_name:
            selected.append({
                "sheet_name": real_name,
                "bereich": rule["bereich"],
                "direct_filter": bool(rule["direct_filter"]),
            })
        else:
            missing.append(rule["bereich"])

    return selected, missing


def day_value_is_set(value) -> bool:
    """Ein Tag gilt als vorhanden, wenn in Montag bis Samstag ein echter Wert steht."""
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


def normalize_tour_display(value) -> str:
    """Schreibt Tournummern ohne .0, lässt Textwerte wie 22+3 aber stehen."""
    if value is None or pd.isna(value):
        return ""
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(number) and float(number).is_integer():
        return str(int(number))
    return str(value).strip()


def extract_tour_numbers(value) -> Set[str]:
    """Extrahiert Tournummern aus einer Zelle. Wichtig für den Direkt-Filter."""
    if value is None or pd.isna(value):
        return set()

    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(number) and float(number).is_integer():
        return {str(int(number))}

    text = str(value).strip()
    found = set(re.findall(r"\d+", text))
    return found


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


def days_to_text(days: Set[int] | List[int]) -> str:
    return ", ".join(f"{d} {DAY_NAMES[d]}" for d in sorted(days))


def day_to_text(day: int) -> str:
    return f"{day} {DAY_NAMES[day]}"


def join_unique(values) -> str:
    cleaned = [value_to_clean_text(v) for v in values if value_to_clean_text(v)]
    return ", ".join(sorted(set(cleaned)))


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
    """Liest NMS und Malchow komplett, aus Direkt nur die definierten Tournummern."""
    excel = pd.ExcelFile(uploaded_file)
    selected_rules, missing_sheet_names = select_tour_sheet_rules(excel)

    frames: List[pd.DataFrame] = []
    customer_info: Dict[str, Dict[str, str]] = {}

    for sheet_rule in selected_rules:
        sheet_name = sheet_rule["sheet_name"]
        bereich = sheet_rule["bereich"]
        direct_filter = sheet_rule["direct_filter"]

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

        work["Bereich"] = bereich
        work["Blatt Tourenplanung"] = sheet_name
        long = work.melt(
            id_vars=["sap", "Bereich", "Blatt Tourenplanung"],
            value_vars=day_value_columns,
            var_name="tag_col",
            value_name="wert",
        )
        long["tag_num"] = long["tag_col"].str.replace("tag_", "", regex=False).astype(int)
        long["wert_gesetzt"] = long["wert"].map(day_value_is_set)
        long["Tournummer"] = long["wert"].map(normalize_tour_display)

        long = long[long["sap"].ne("") & long["wert_gesetzt"]].copy()

        if direct_filter:
            long["tournummer_set"] = long["wert"].map(extract_tour_numbers)
            long = long[long["tournummer_set"].map(lambda values: bool(values & DIRECT_TOURS))].copy()
            long = long.drop(columns=["tournummer_set"], errors="ignore")

        if not long.empty:
            frames.append(long[["sap", "Bereich", "Blatt Tourenplanung", "tag_num", "Tournummer"]])

    checked_sheets = [rule["sheet_name"] for rule in selected_rules]

    if not frames:
        empty = pd.DataFrame(columns=["sap", "Bereich", "Blatt Tourenplanung", "tag_num", "Tournummer"])
        return empty, checked_sheets, missing_sheet_names, customer_info

    return pd.concat(frames, ignore_index=True), checked_sheets, missing_sheet_names, customer_info


# ---------------------------------------------------------------------------
# Vergleich
# ---------------------------------------------------------------------------


def _export_columns() -> List[str]:
    return [
        "Bereich",
        "Blatt Tourenplanung",
        "Tournummer",
        "SAP Nummer",
        "Name",
        "Straße",
        "Ort",
        "Liefertag",
        "LT SAP",
        "LT Tourenplanung",
        "Hinweis",
    ]


def _empty_result_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_export_columns())


def make_scope_key(row) -> Tuple[str, str, str]:
    return (
        value_to_clean_text(row.get("sap", "")),
        value_to_clean_text(row.get("Bereich", "")),
        value_to_clean_text(row.get("Blatt Tourenplanung", "")),
    )


def build_tour_day_overview(tour_df: pd.DataFrame) -> pd.DataFrame:
    """Eine Zeile je SAP, Bereich, Blatt und Liefertag aus dem geprüften Tourbereich.

    Wichtig: Direkt ist vorher bereits auf die erlaubten Tournummern gefiltert.
    Dadurch werden Unterschiede nicht mehr nur pro SAP-Nummer, sondern genau im
    geprüften Bereich und mit der konkreten Tages-Tournummer bewertet.
    """
    if tour_df.empty:
        return pd.DataFrame(columns=["sap", "Bereich", "Blatt Tourenplanung", "tag_num", "Tournummer"])

    return tour_df.groupby(["sap", "Bereich", "Blatt Tourenplanung", "tag_num"], as_index=False).agg(
        Tournummer=("Tournummer", join_unique),
    )


def build_days_in_scope(tour_df: pd.DataFrame) -> Dict[Tuple[str, str, str], Set[int]]:
    """Liefertage je SAP + Bereich + Blatt.

    Das verhindert, dass ein Tag aus NMS versehentlich einen fehlenden Tag in Direkt
    ausgleicht. Genau hier werden die Tournummern beziehungsweise Bereiche sauber berücksichtigt.
    """
    days: Dict[Tuple[str, str, str], Set[int]] = {}
    if tour_df.empty:
        return days

    for row in tour_df[["sap", "Bereich", "Blatt Tourenplanung", "tag_num"]].drop_duplicates().to_dict("records"):
        key = (row["sap"], row["Bereich"], row["Blatt Tourenplanung"])
        days.setdefault(key, set()).add(int(row["tag_num"]))
    return days


def scope_days_text(row, days_in_scope: Dict[Tuple[str, str, str], Set[int]]) -> str:
    key = make_scope_key(row)
    return days_to_text(days_in_scope.get(key, set())) or "(nicht gesetzt)"


def direct_expected_tour_for_day(day: int) -> str:
    """Erwartete Direkt-Tournummer je Liefertag."""
    return DIRECT_TOUR_BY_DAY.get(int(day), "(Direkt-Tour unbekannt)")


def add_common_result_columns(
    df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    days_in_scope: Dict[Tuple[str, str, str], Set[int]],
    customer_info: Dict[str, Dict[str, str]],
    hinweis: str,
) -> pd.DataFrame:
    if df.empty:
        return _empty_result_df()

    out = df.copy()
    if "Tournummer" not in out.columns:
        out["Tournummer"] = ""

    out["SAP Nummer"] = out["sap"]
    out["Name"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("name", ""))
    out["Straße"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("strasse", ""))
    out["Ort"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("ort", ""))
    out["Liefertag"] = out["tag_num"].map(day_to_text)
    out["LT SAP"] = out["sap"].map(lambda s: days_to_text(days_by_sap.get(s, set())) or "(keine hinterlegt)")
    out["LT Tourenplanung"] = out.apply(lambda row: scope_days_text(row, days_in_scope), axis=1)
    out["Hinweis"] = hinweis
    out["_BereichSort"] = out["Bereich"].map({"NMS": 1, "Malchow": 2, "Direkt": 3}).fillna(99)
    out["_SapSort"] = pd.to_numeric(out["sap"], errors="coerce").fillna(9_999_999_999)
    out = out.sort_values(["_BereichSort", "Blatt Tourenplanung", "_SapSort", "tag_num", "Tournummer"]).reset_index(drop=True)
    return out[_export_columns()]


def build_missing_in_sap(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """Tage stehen in der geprüften Tourenplanung, fehlen aber in SAP.

    Direkt wird hier ausschließlich über die erlaubten Tageszellen bewertet:
    1058, 2058, 3058, 4058, 5058 und 6030.
    """
    if tour_df.empty:
        return _empty_result_df()

    tour_days = build_tour_day_overview(tour_df)
    days_in_scope = build_days_in_scope(tour_df)

    missing = tour_days[
        tour_days.apply(lambda row: int(row["tag_num"]) not in days_by_sap.get(row["sap"], set()), axis=1)
    ].copy()

    return add_common_result_columns(
        missing,
        days_by_sap,
        days_in_scope,
        customer_info,
        "In der geprüften Tourenplanung vorhanden, in SAP nicht als Liefertag hinterlegt.",
    )


def build_missing_in_tour(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """Tage stehen in SAP, fehlen aber im geprüften Tourbereich.

    Die Prüfung läuft je SAP + Bereich + Blatt. Dadurch zählt ein Tag aus NMS
    nicht versehentlich als Treffer für Direkt oder Malchow.

    Für Direkt wird die erwartete Tournummer aus dem Liefertag abgeleitet:
    Montag 1058, Dienstag 2058, Mittwoch 3058, Donnerstag 4058,
    Freitag 5058, Samstag 6030.
    """
    if tour_df.empty:
        return _empty_result_df()

    days_in_scope = build_days_in_scope(tour_df)
    scope_rows = (
        tour_df[["sap", "Bereich", "Blatt Tourenplanung"]]
        .drop_duplicates()
        .sort_values(["Bereich", "Blatt Tourenplanung", "sap"])
        .to_dict("records")
    )

    rows: List[dict] = []
    for scope in scope_rows:
        sap = scope["sap"]
        bereich = scope["Bereich"]
        sheet_name = scope["Blatt Tourenplanung"]
        key = (sap, bereich, sheet_name)

        sap_days = days_by_sap.get(sap, set())
        tour_days = days_in_scope.get(key, set())
        missing_days = sorted(sap_days - tour_days)
        if not missing_days:
            continue

        for day in missing_days:
            if bereich == "Direkt":
                tournummer = direct_expected_tour_for_day(day)
                # Sicherheit: Nur die ausdrücklich gewünschten Direkt-Touren ausgeben.
                if tournummer not in DIRECT_TOURS:
                    continue
            else:
                tournummer = "(fehlt in Tour)"

            rows.append({
                "sap": sap,
                "tag_num": day,
                "Bereich": bereich,
                "Blatt Tourenplanung": sheet_name,
                "Tournummer": tournummer,
            })

    if not rows:
        return _empty_result_df()

    missing = pd.DataFrame(rows)
    return add_common_result_columns(
        missing,
        days_by_sap,
        days_in_scope,
        customer_info,
        "In SAP vorhanden, im geprüften Tourbereich aber nicht gesetzt.",
    )


def _filter_dataframe(df: pd.DataFrame, suche: str, bereich: str = "Alle") -> pd.DataFrame:
    if df is None or df.empty:
        return df

    work = df
    if bereich and bereich != "Alle" and "Bereich" in work.columns:
        work = work[work["Bereich"].astype(str).str.contains(bereich, na=False, regex=False)]

    if suche:
        such = suche.strip().lower()
        if such:
            spalten = [
                c for c in [
                    "Bereich", "Blatt Tourenplanung", "Tournummer", "SAP Nummer",
                    "Name", "Straße", "Ort", "Liefertag", "LT SAP", "LT Tourenplanung", "Hinweis"
                ]
                if c in work.columns
            ]
            mask = pd.Series(False, index=work.index)
            for c in spalten:
                mask = mask | work[c].astype(str).str.lower().str.contains(such, na=False)
            work = work[mask]

    return work


def build_overview(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame, tour_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for bereich in ["NMS", "Malchow", "Direkt"]:
        selected_rows = int((tour_df["Bereich"] == bereich).sum()) if not tour_df.empty else 0
        selected_customers = int(tour_df.loc[tour_df["Bereich"] == bereich, "sap"].nunique()) if not tour_df.empty else 0
        selected_tours = ""
        if bereich == "Direkt":
            selected_tours = ", ".join(sorted(DIRECT_TOURS))

        rows.append({
            "Bereich": bereich,
            "Direkt-Tourfilter": selected_tours,
            "Geprüfte Tour-Zeilen": selected_rows,
            "Geprüfte SAP-Nummern": selected_customers,
            "Fehlt in SAP": int((missing_sap["Bereich"].astype(str).str.contains(bereich, na=False, regex=False)).sum()) if not missing_sap.empty else 0,
            "Fehlt in Tour": int((missing_tour["Bereich"].astype(str).str.contains(bereich, na=False, regex=False)).sum()) if not missing_tour.empty else 0,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Excel-Erzeugung
# ---------------------------------------------------------------------------


def build_excel(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame) -> bytes:
    """Schreibt eine Excel mit allen Unterschieden und beiden Einzelrichtungen."""
    output = io.BytesIO()

    all_parts = []
    if missing_sap is not None and not missing_sap.empty:
        part = missing_sap.copy()
        part.insert(0, "Prüfung", "Fehlt in SAP / zu viel in Tour")
        all_parts.append(part)
    if missing_tour is not None and not missing_tour.empty:
        part = missing_tour.copy()
        part.insert(0, "Prüfung", "Fehlt in Tour / zu viel in SAP")
        all_parts.append(part)

    if all_parts:
        all_out = pd.concat(all_parts, ignore_index=True)
    else:
        all_out = pd.DataFrame(columns=["Prüfung"] + _export_columns())

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_out.to_excel(writer, index=False, sheet_name="Alle Unterschiede", na_rep="")
        _format_sheet(writer, "Alle Unterschiede", all_out)

        missing_sap.to_excel(writer, index=False, sheet_name="Fehlt in SAP", na_rep="")
        _format_sheet(writer, "Fehlt in SAP", missing_sap)

        missing_tour.to_excel(writer, index=False, sheet_name="Fehlt in Tour", na_rep="")
        _format_sheet(writer, "Fehlt in Tour", missing_tour)

        wb = writer.book
        for ws in wb.worksheets:
            ws.sheet_state = "visible"

    return output.getvalue()


_RIGHT_ALIGN_COLS = {"SAP Nummer"}

_COL_WIDTH_HINTS = {
    "Prüfung": (22, 32),
    "Bereich": (10, 16),
    "Blatt Tourenplanung": (18, 34),
    "Tournummer": (12, 22),
    "SAP Nummer": (10, 12),
    "Name": (24, 42),
    "Straße": (20, 32),
    "Ort": (22, 36),
    "Liefertag": (16, 22),
    "LT SAP": (24, 48),
    "LT Tourenplanung": (24, 48),
    "Hinweis": (36, 65),
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

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_left
        cell.border = border_thin
    ws.row_dimensions[1].height = 24

    columns = list(df.columns)
    group_column = "Prüfung" if "Prüfung" in columns else "Bereich"
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
            cell.alignment = align_right if col_name in _RIGHT_ALIGN_COLS else align_left
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
        .small-note { color: #586174; font-size: 0.92rem; line-height: 1.45; }
        .scope-box { background:#fafbfc; border:1px solid #e3e7ee; border-radius:14px; padding:14px 16px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SAP ↔ Tourenplanung")
st.markdown(
    f"""
    <div class="scope-box">
    <b>Prüfumfang</b><br>
    Geprüft werden nur <b>NMS</b> und <b>Malchow</b> komplett sowie aus <b>Direkt</b> nur Tageszellen mit den Tournummern
    <b>{', '.join(sorted(DIRECT_TOURS))}</b>.<br>
    Die Auswertung zeigt immer beide Richtungen: <b>Fehlt in SAP / zu viel in Tour</b> und
    <b>Fehlt in Tour / zu viel in SAP</b>. Die <b>Tournummer</b> wird in jeder Ergebniszeile ausgegeben.
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
        help="Geprüft werden: NMS, Malchow und Direkt mit den Tournummern 1058, 2058, 3058, 4058, 5058, 6030.",
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
                "In der Tourenplanung wurden keine passenden Liefertage erkannt. "
                "Bitte prüfen, ob die Blätter NMS, Malchow und Direkt vorhanden sind und ob Direkt die gewünschten Tournummern enthält."
            )
        if missing_tour_sheets:
            st.warning("Nicht alle erwarteten Bereiche wurden gefunden: " + ", ".join(missing_tour_sheets))

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
            "tour_df": tour_df,
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
    tour_df = result["tour_df"]

    st.divider()

    head_left, head_right = st.columns([3, 1])
    with head_left:
        st.subheader("Ergebnis")
        st.caption(
            f"SAP: Blatt {result['sap_sheet']}, {result['sap_rows']} Liefertage übernommen · "
            f"Tourenplanung: {', '.join(result['tour_sheets'])}, {result.get('tour_rows', 0)} passende Tour-Liefertage erkannt"
        )
        st.caption(
            "Hinweis: Für 'Fehlt in Tour' werden nur SAP-Nummern geprüft, die im ausgewählten Tourbereich vorkommen. "
            "So werden keine fremden Bereiche aus der SAP-Datei als Fehler ausgegeben."
        )
    with head_right:
        st.download_button(
            label="Excel herunterladen",
            data=result["excel_bytes"],
            file_name="sap_tourenplanung_nms_malchow_direkt_1058_unterschiede.xlsx",
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
        overview = build_overview(missing_sap, missing_tour, tour_df)
        st.dataframe(overview, use_container_width=True, hide_index=True)

        if result["all_count"] == 0:
            st.success("Keine Unterschiede gefunden.")

    combined_parts = []
    if not missing_sap.empty:
        part = missing_sap.copy()
        part.insert(0, "Prüfung", "Fehlt in SAP / zu viel in Tour")
        combined_parts.append(part)
    if not missing_tour.empty:
        part = missing_tour.copy()
        part.insert(0, "Prüfung", "Fehlt in Tour / zu viel in SAP")
        combined_parts.append(part)
    combined = pd.concat(combined_parts, ignore_index=True) if combined_parts else pd.DataFrame()

    with tabs[1]:
        if combined.empty:
            st.info("Keine Treffer.")
        else:
            f1, f2 = st.columns([1, 2])
            bereich = f1.selectbox("Bereich", ["Alle", "NMS", "Malchow", "Direkt"], key="filter_all_bereich")
            suche = f2.text_input("Suchen", key="filter_all_suche", placeholder="Tournummer, SAP Nummer, Name, Straße, Ort oder Liefertag")
            filtered = _filter_dataframe(combined, suche, bereich)
            st.caption(f"{len(filtered)} von {len(combined)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[2]:
        if missing_sap.empty:
            st.info("Keine Treffer: Es gibt keine Tage in der Tourenplanung, die in SAP fehlen.")
        else:
            f1, f2 = st.columns([1, 2])
            bereich = f1.selectbox("Bereich", ["Alle", "NMS", "Malchow", "Direkt"], key="filter_sap_bereich")
            suche = f2.text_input("Suchen", key="filter_sap_suche", placeholder="Tournummer, SAP Nummer, Name, Straße, Ort oder Liefertag")
            filtered = _filter_dataframe(missing_sap, suche, bereich)
            st.caption(f"{len(filtered)} von {len(missing_sap)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[3]:
        if missing_tour.empty:
            st.info("Keine Treffer: Es gibt keine Tage in SAP, die im ausgewählten Tourbereich fehlen.")
        else:
            f1, f2 = st.columns([1, 2])
            bereich = f1.selectbox("Bereich", ["Alle", "NMS", "Malchow", "Direkt"], key="filter_tour_bereich")
            suche = f2.text_input("Suchen", key="filter_tour_suche", placeholder="Tournummer, SAP Nummer, Name, Straße, Ort oder Liefertag")
            filtered = _filter_dataframe(missing_tour, suche, bereich)
            st.caption(f"{len(filtered)} von {len(missing_tour)} Zeilen")
            st.dataframe(filtered, use_container_width=True, hide_index=True)
