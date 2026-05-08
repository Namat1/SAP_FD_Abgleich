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

# Excel-Positionen, wenn keine Überschriften erkannt werden.
# Index 0 = Spalte A.
# Tourenplanung: A CSB, B SAP, C Name, D Strasse, E Plz, F Ort, G Mo, H Die, I Mitt, J Don, K Fr, L Sam.
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

# Direkt wird hart nur über diese Tournummern geprüft.
DIRECT_TOURS: Set[str] = {"1058", "2058", "3058", "4058", "5058", "6030"}

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
        "candidates": [
            "NMS", "HUPA_NMS", "HUPA NMS", "HUPA-NMS", "HuPa NMS", "Neumünster", "Neumuenster",
        ],
        "direct_filter": False,
    },
    {
        "bereich": "Malchow",
        "candidates": [
            "Malchow", "HUPA_MALCHOW", "HUPA MALCHOW", "HUPA-MALCHOW", "HuPa Malchow",
        ],
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

SAP_TOUR_COLUMN_CANDIDATES = [
    "Tournummer", "Tour Nummer", "Tour-Nr", "Tour Nr", "Tour",
    "Transportgruppe", "Transport Gruppe", "Transport-Gruppen", "Transportgruppen",
    "TG", "T-Gruppe", "Tourengruppe", "Tour Gruppe", "Rahmentour", "Rahmen Tour",
    "CSB Tournummer", "CSB Tour", "Route", "Routennummer",
]

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

    numeric = pd.to_numeric(series, errors="coerce")
    is_int = numeric.notna() & (numeric == numeric.round())

    out = series.astype(str)
    out = out.where(~is_int, numeric.where(is_int).astype("Int64").astype(str))
    out = out.str.strip()
    out = out.replace({"nan": "", "<NA>": "", "None": ""})
    return out


def normalize_header_name(value) -> str:
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
    """Findet nur die Blätter NMS, Malchow und Direkt."""
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

        if real_name is None:
            for normalized_name, original_name in available_by_normalized.items():
                if rule["bereich"] == "NMS" and "nms" in normalized_name:
                    real_name = original_name
                    break
                if rule["bereich"] == "Malchow" and "malchow" in normalized_name:
                    real_name = original_name
                    break
                if rule["bereich"] == "Direkt" and "direkt" in normalized_name:
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


def normalize_tour_text(value) -> str:
    """Normiert eine einzelne Tournummer für Anzeige und Vergleich."""
    if value is None or pd.isna(value):
        return ""

    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "<na>", "-", "--"}:
        return ""

    # 6.030 oder 6 030 wird zu 6030. Tausendertrennzeichen werden entfernt.
    compact = re.sub(r"(?<=\d)[\.\s](?=\d{3}\b)", "", text)
    compact = compact.replace("\u00a0", " ").strip()

    numeric = pd.to_numeric(pd.Series([compact]), errors="coerce").iloc[0]
    if pd.notna(numeric) and float(numeric).is_integer():
        return str(int(numeric))

    # 6030.0 als Text
    if re.fullmatch(r"\d+\.0+", compact):
        return compact.split(".")[0]

    return compact


def extract_tour_numbers(value) -> List[str]:
    """Extrahiert einzelne Tournummern aus einer Zelle.

    Wichtig: Eine Zelle wie "6029 / 6030" wird in zwei Tournummern gesplittet.
    Für Direkt wird danach hart nur auf 1058, 2058, 3058, 4058, 5058 und 6030 gefiltert.
    """
    if not day_value_is_set(value):
        return []

    normalized = normalize_tour_text(value)
    if not normalized:
        return []

    # Reine Zahl
    if re.fullmatch(r"\d+", normalized):
        return [normalized]

    # Zunächst Tausendertrennzeichen entfernen, damit 6.030 als 6030 erkannt wird.
    text = re.sub(r"(?<=\d)[\.\s](?=\d{3}\b)", "", normalized)
    text = text.replace("\u00a0", " ")

    found: List[str] = []
    for raw in re.findall(r"\d+(?:\.0+)?", text):
        cleaned = raw.split(".")[0] if raw.endswith(".0") or re.fullmatch(r"\d+\.0+", raw) else raw
        if cleaned and cleaned not in found:
            found.append(cleaned)

    if found:
        return found

    # Fallback für Textwerte ohne Ziffern. Für NMS/Malchow soll der Wert nicht verloren gehen.
    return [normalized]


def normalize_day_code_series(series: pd.Series) -> pd.Series:
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
    return f"{day} {DAY_NAMES[int(day)]}"


def join_unique(values) -> str:
    cleaned = [value_to_clean_text(v) for v in values if value_to_clean_text(v)]
    return ", ".join(sorted(set(cleaned), key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else x)))


def expected_direct_tour(day: int) -> str:
    return DIRECT_TOUR_BY_DAY.get(int(day), "")

# ---------------------------------------------------------------------------
# Dateien lesen
# ---------------------------------------------------------------------------


def read_sap_file(uploaded_file) -> Tuple[Dict[str, Set[int]], pd.DataFrame, str, int, Optional[str]]:
    """Liest die SAP-Datei.

    Standard: Spalte A = SAP, Spalte G = Liefertag.
    Wenn in SAP zusätzlich eine Spalte für Tournummer / Transportgruppe gefunden wird,
    wird sie ebenfalls gelesen und für einen exakten Abgleich genutzt.
    """
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = read_excel_with_detected_header(excel, sheet_name)

    empty_records = pd.DataFrame(columns=["sap", "tag_num", "Tournummer SAP"])
    if df.empty:
        return {}, empty_records, sheet_name, 0, None

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
    tour_column = pick_first_matching_column(columns, normalized_candidates(SAP_TOUR_COLUMN_CANDIDATES))

    if sap_column is None or day_column is None:
        return {}, empty_records, sheet_name, 0, tour_column

    needed = [sap_column, day_column]
    if tour_column and tour_column not in needed:
        needed.append(tour_column)

    work = df[needed].copy()
    work = work.rename(columns={sap_column: "sap", day_column: "tag"})
    work["sap"] = normalize_sap_series(work["sap"])
    work["tag_num"] = normalize_day_code_series(work["tag"])

    if tour_column:
        work["Tournummer SAP"] = work[tour_column].map(lambda value: ", ".join(extract_tour_numbers(value)))
    else:
        work["Tournummer SAP"] = ""

    mask = (
        work["sap"].ne("")
        & work["tag_num"].between(1, 6, inclusive="both")
        & work["tag_num"].notna()
    )
    filtered = work.loc[mask, ["sap", "tag_num", "Tournummer SAP"]].copy()
    filtered["tag_num"] = filtered["tag_num"].astype(int)

    # Wenn mehrere Tournummern in einer SAP-Zelle stehen, je Tournummer eine Zeile.
    rows: List[dict] = []
    for _, row in filtered.iterrows():
        tours = [tour for tour in str(row["Tournummer SAP"]).split(",") if tour.strip()]
        if not tours:
            rows.append({"sap": row["sap"], "tag_num": int(row["tag_num"]), "Tournummer SAP": ""})
        else:
            for tour in tours:
                rows.append({"sap": row["sap"], "tag_num": int(row["tag_num"]), "Tournummer SAP": tour.strip()})

    records = pd.DataFrame(rows, columns=["sap", "tag_num", "Tournummer SAP"]).drop_duplicates().reset_index(drop=True)
    days_by_sap: Dict[str, Set[int]] = records.groupby("sap")["tag_num"].agg(set).to_dict() if not records.empty else {}

    return days_by_sap, records, sheet_name, len(records), tour_column


def read_tourenplanung(uploaded_file) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, Dict[str, str]]]:
    """Liest NMS und Malchow komplett; aus Direkt nur die exakten gewünschten Tournummern."""
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
            day_column = pick_column_by_name_or_position(columns, DAY_COLUMN_CANDIDATES[day_num], col_index)
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

        day_value_columns = [f"tag_{d}" for d in DAY_COLUMNS_TOUR.keys() if f"tag_{d}" in work.columns]
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
        long = long[long["sap"].ne("") & long["wert"].map(day_value_is_set)].copy()

        rows: List[dict] = []
        for _, row in long.iterrows():
            extracted_tours = extract_tour_numbers(row["wert"])

            if direct_filter:
                # Direkt: exakt nur die erlaubten Tournummern behalten.
                extracted_tours = [tour for tour in extracted_tours if tour in DIRECT_TOURS]
                if not extracted_tours:
                    continue

            for tournummer in extracted_tours:
                rows.append({
                    "sap": row["sap"],
                    "Bereich": row["Bereich"],
                    "Blatt Tourenplanung": row["Blatt Tourenplanung"],
                    "tag_num": int(row["tag_num"]),
                    "Tournummer Tour": tournummer,
                })

        if rows:
            frames.append(pd.DataFrame(rows))

    checked_sheets = [rule["sheet_name"] for rule in selected_rules]

    if not frames:
        empty = pd.DataFrame(columns=["sap", "Bereich", "Blatt Tourenplanung", "tag_num", "Tournummer Tour"])
        return empty, checked_sheets, missing_sheet_names, customer_info

    return pd.concat(frames, ignore_index=True).drop_duplicates().reset_index(drop=True), checked_sheets, missing_sheet_names, customer_info

# ---------------------------------------------------------------------------
# Vergleich
# ---------------------------------------------------------------------------


def _export_columns() -> List[str]:
    return [
        "Richtung",
        "Bereich",
        "Blatt Tourenplanung",
        "SAP Nummer",
        "Name",
        "Straße",
        "Ort",
        "Liefertag",
        "Tournummer Tour",
        "Tournummer SAP",
        "LT SAP",
        "LT Tourenplanung",
        "Hinweis",
    ]


def _empty_result_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_export_columns())


def build_days_by_scope(tour_df: pd.DataFrame) -> Dict[Tuple[str, str, str], Set[int]]:
    result: Dict[Tuple[str, str, str], Set[int]] = {}
    if tour_df.empty:
        return result
    for row in tour_df[["sap", "Bereich", "Blatt Tourenplanung", "tag_num"]].drop_duplicates().to_dict("records"):
        key = (row["sap"], row["Bereich"], row["Blatt Tourenplanung"])
        result.setdefault(key, set()).add(int(row["tag_num"]))
    return result


def build_tours_by_scope_day(tour_df: pd.DataFrame) -> Dict[Tuple[str, str, str, int], Set[str]]:
    result: Dict[Tuple[str, str, str, int], Set[str]] = {}
    if tour_df.empty:
        return result
    for row in tour_df.to_dict("records"):
        key = (row["sap"], row["Bereich"], row["Blatt Tourenplanung"], int(row["tag_num"]))
        result.setdefault(key, set()).add(value_to_clean_text(row["Tournummer Tour"]))
    return result


def build_sap_tours_by_day(sap_records: pd.DataFrame) -> Dict[Tuple[str, int], Set[str]]:
    result: Dict[Tuple[str, int], Set[str]] = {}
    if sap_records.empty:
        return result
    for row in sap_records.to_dict("records"):
        key = (row["sap"], int(row["tag_num"]))
        tour = value_to_clean_text(row.get("Tournummer SAP", ""))
        if tour:
            result.setdefault(key, set()).add(tour)
    return result


def scope_days_text(row, days_by_scope: Dict[Tuple[str, str, str], Set[int]]) -> str:
    key = (row["sap"], row["Bereich"], row["Blatt Tourenplanung"])
    return days_to_text(days_by_scope.get(key, set())) or "(nicht gesetzt)"


def sap_tours_text(sap_tours_by_day: Dict[Tuple[str, int], Set[str]], sap: str, day: int) -> str:
    return join_unique(sap_tours_by_day.get((sap, int(day)), set()))


def add_common_result_columns(
    df: pd.DataFrame,
    richtung: str,
    days_by_sap: Dict[str, Set[int]],
    sap_tours_by_day: Dict[Tuple[str, int], Set[str]],
    days_by_scope: Dict[Tuple[str, str, str], Set[int]],
    customer_info: Dict[str, Dict[str, str]],
    hinweis: str,
) -> pd.DataFrame:
    if df.empty:
        return _empty_result_df()

    out = df.copy()
    if "Tournummer Tour" not in out.columns:
        out["Tournummer Tour"] = ""
    if "Tournummer SAP" not in out.columns:
        out["Tournummer SAP"] = out.apply(lambda row: sap_tours_text(sap_tours_by_day, row["sap"], row["tag_num"]), axis=1)

    out["Richtung"] = richtung
    out["SAP Nummer"] = out["sap"]
    out["Name"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("name", ""))
    out["Straße"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("strasse", ""))
    out["Ort"] = out["sap"].map(lambda s: customer_info.get(s, {}).get("ort", ""))
    out["Liefertag"] = out["tag_num"].map(day_to_text)
    out["LT SAP"] = out["sap"].map(lambda s: days_to_text(days_by_sap.get(s, set())) or "(keine hinterlegt)")
    out["LT Tourenplanung"] = out.apply(lambda row: scope_days_text(row, days_by_scope), axis=1)
    out["Hinweis"] = hinweis

    out["_BereichSort"] = out["Bereich"].map({"NMS": 1, "Malchow": 2, "Direkt": 3}).fillna(99)
    out["_SapSort"] = pd.to_numeric(out["sap"], errors="coerce").fillna(9_999_999_999)
    out["_TourSort"] = pd.to_numeric(out["Tournummer Tour"].replace("", pd.NA), errors="coerce").fillna(9_999_999_999)
    out = out.sort_values(["_BereichSort", "Blatt Tourenplanung", "_SapSort", "tag_num", "_TourSort"]).reset_index(drop=True)
    return out[_export_columns()]


def build_missing_in_sap(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    sap_records: pd.DataFrame,
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """Tourenplanung hat den Tag / die Tour, SAP nicht.

    Wenn SAP eine Tournummer-Spalte hat, wird exakt SAP Nummer + Liefertag + Tournummer geprüft.
    Ohne SAP-Tournummer-Spalte wird auf SAP Nummer + Liefertag zurückgefallen.
    """
    if tour_df.empty:
        return _empty_result_df()

    sap_tours_by_day = build_sap_tours_by_day(sap_records)
    days_by_scope = build_days_by_scope(tour_df)

    rows: List[dict] = []
    for row in tour_df.to_dict("records"):
        sap = row["sap"]
        day = int(row["tag_num"])
        tour_tour = value_to_clean_text(row["Tournummer Tour"])
        sap_tours_for_day = sap_tours_by_day.get((sap, day), set())

        if sap_tours_for_day:
            fehlt = tour_tour not in sap_tours_for_day
        else:
            fehlt = day not in days_by_sap.get(sap, set())

        if fehlt:
            rows.append({
                "sap": sap,
                "Bereich": row["Bereich"],
                "Blatt Tourenplanung": row["Blatt Tourenplanung"],
                "tag_num": day,
                "Tournummer Tour": tour_tour,
                "Tournummer SAP": join_unique(sap_tours_for_day),
            })

    missing = pd.DataFrame(rows)
    return add_common_result_columns(
        missing,
        "Fehlt in SAP / zu viel in Tour",
        days_by_sap,
        sap_tours_by_day,
        days_by_scope,
        customer_info,
        "Dieser Liefertag mit dieser Tournummer steht im geprüften Tourbereich, ist so aber nicht in SAP vorhanden.",
    )


def build_missing_in_tour(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    sap_records: pd.DataFrame,
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """SAP hat den Tag / die Tour, die geprüfte Tourenplanung nicht.

    Wichtig: Die Rückrichtung wird auf den tatsächlich geprüften Tourbereich begrenzt.
    Es werden also nicht mehr alle SAP-Tage eines Kunden gegen den Direkt-Bereich geworfen.

    - Wenn SAP eine Tournummer enthält, wird exakt SAP Nummer + Liefertag + Tournummer verglichen.
    - Für Direkt werden dabei ausschließlich 1058, 2058, 3058, 4058, 5058 und 6030 bewertet.
    - Wenn SAP keine Tournummer enthält, wird Direkt in der Rückrichtung nicht pauschal über alle
      SAP-Tage geprüft, weil dann nicht sicher erkennbar ist, ob der SAP-Tag zu einer der
      gewünschten Direkt-Touren gehört. Dadurch erscheinen keine fremden Direkt-Kunden mehr.
    """
    if tour_df.empty:
        return _empty_result_df()

    sap_tours_by_day = build_sap_tours_by_day(sap_records)
    days_by_scope = build_days_by_scope(tour_df)
    tours_by_scope_day = build_tours_by_scope_day(tour_df)

    has_sap_tour_numbers = False
    if sap_records is not None and not sap_records.empty and "Tournummer SAP" in sap_records.columns:
        has_sap_tour_numbers = sap_records["Tournummer SAP"].fillna("").astype(str).str.strip().ne("").any()

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
        scope_key = (sap, bereich, sheet_name)

        sap_days = days_by_sap.get(sap, set())
        tour_days = days_by_scope.get(scope_key, set())

        for day in sorted(sap_days):
            sap_tours_for_day = sap_tours_by_day.get((sap, int(day)), set())
            tour_tours_for_day = tours_by_scope_day.get((sap, bereich, sheet_name, int(day)), set())

            if sap_tours_for_day:
                # Exakter Tournummernvergleich, wenn SAP-Tournummern vorhanden sind.
                if bereich == "Direkt":
                    # Direkt: nur die gewünschten Tournummern aus SAP bewerten.
                    sap_tours_for_day = {tour for tour in sap_tours_for_day if tour in DIRECT_TOURS}
                    if not sap_tours_for_day:
                        continue

                missing_tours = sorted(
                    sap_tours_for_day - tour_tours_for_day,
                    key=lambda x: int(x) if str(x).isdigit() else str(x),
                )
                for sap_tour in missing_tours:
                    rows.append({
                        "sap": sap,
                        "Bereich": bereich,
                        "Blatt Tourenplanung": sheet_name,
                        "tag_num": int(day),
                        "Tournummer Tour": join_unique(tour_tours_for_day) or "(fehlt in Tour)",
                        "Tournummer SAP": sap_tour,
                    })
            else:
                # Ohne SAP-Tournummer kann Direkt nicht zuverlässig gegen einzelne Tournummern geprüft werden.
                # Sonst würde jeder weitere SAP-Liefertag des Kunden als Fehler im Direktbereich auftauchen.
                if bereich == "Direkt" and not has_sap_tour_numbers:
                    continue

                # Für NMS und Malchow bleibt der reine Liefertagsvergleich erhalten.
                if int(day) not in tour_days:
                    rows.append({
                        "sap": sap,
                        "Bereich": bereich,
                        "Blatt Tourenplanung": sheet_name,
                        "tag_num": int(day),
                        "Tournummer Tour": "(fehlt in Tour)",
                        "Tournummer SAP": "",
                    })

    missing = pd.DataFrame(rows)
    return add_common_result_columns(
        missing,
        "Fehlt in Tour / zu viel in SAP",
        days_by_sap,
        sap_tours_by_day,
        days_by_scope,
        customer_info,
        "Dieser Liefertag beziehungsweise diese SAP-Tour ist in SAP vorhanden, fehlt aber im geprüften Tourbereich.",
    )


def build_all_differences(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame) -> pd.DataFrame:
    frames = [df for df in [missing_sap, missing_tour] if df is not None and not df.empty]
    if not frames:
        return _empty_result_df()
    return pd.concat(frames, ignore_index=True)


def build_overview(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame, tour_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for bereich in ["NMS", "Malchow", "Direkt"]:
        selected_rows = int((tour_df["Bereich"] == bereich).sum()) if not tour_df.empty else 0
        selected_customers = int(tour_df.loc[tour_df["Bereich"] == bereich, "sap"].nunique()) if not tour_df.empty else 0
        rows.append({
            "Bereich": bereich,
            "Geprüfte Tour-Zeilen": selected_rows,
            "Geprüfte SAP-Nummern": selected_customers,
            "Fehlt in SAP / zu viel in Tour": int((missing_sap["Bereich"] == bereich).sum()) if not missing_sap.empty else 0,
            "Fehlt in Tour / zu viel in SAP": int((missing_tour["Bereich"] == bereich).sum()) if not missing_tour.empty else 0,
        })
    return pd.DataFrame(rows)


def _filter_dataframe(df: pd.DataFrame, suche: str, bereich: str = "Alle") -> pd.DataFrame:
    if df is None or df.empty:
        return df

    work = df
    if bereich and bereich != "Alle" and "Bereich" in work.columns:
        work = work[work["Bereich"].astype(str).eq(bereich)]

    if suche:
        such = suche.strip().lower()
        if such:
            spalten = [c for c in _export_columns() if c in work.columns]
            mask = pd.Series(False, index=work.index)
            for c in spalten:
                mask = mask | work[c].astype(str).str.lower().str.contains(such, na=False, regex=False)
            work = work[mask]

    return work

# ---------------------------------------------------------------------------
# Excel-Ausgabe
# ---------------------------------------------------------------------------


_RIGHT_ALIGN_COLS = {"SAP Nummer", "Tournummer Tour", "Tournummer SAP"}

_COL_WIDTH_HINTS = {
    "Richtung": (24, 34),
    "Bereich": (12, 14),
    "Blatt Tourenplanung": (18, 26),
    "SAP Nummer": (10, 12),
    "Name": (24, 42),
    "Straße": (20, 32),
    "Ort": (22, 36),
    "Liefertag": (16, 18),
    "Tournummer Tour": (14, 18),
    "Tournummer SAP": (14, 18),
    "LT SAP": (24, 48),
    "LT Tourenplanung": (24, 48),
    "Hinweis": (40, 70),
}


def _format_sheet(writer, sheet_name: str, df: pd.DataFrame) -> None:
    if df is None:
        return
    ws = writer.sheets[sheet_name]
    n_rows = len(df)
    n_cols = len(df.columns)

    header_fill = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
    zebra_fill = PatternFill(start_color="FFE8EFF7", end_color="FFE8EFF7", fill_type="solid")

    thin = Side(style="thin", color="FFCBD5E0")
    medium = Side(style="medium", color="FF305496")
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
    bereich_idx = columns.index("Bereich") if "Bereich" in columns else None
    bereich_values = df["Bereich"].tolist() if bereich_idx is not None else []

    for row_offset in range(n_rows):
        excel_row = row_offset + 2
        is_zebra = (row_offset % 2) == 1

        new_group = False
        if bereich_idx is not None and row_offset > 0:
            if bereich_values[row_offset] != bereich_values[row_offset - 1]:
                new_group = True

        ws.row_dimensions[excel_row].height = 20

        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.font = body_font
            if col_name in _RIGHT_ALIGN_COLS:
                cell.alignment = align_right
            elif col_name in {"Bereich", "Liefertag"}:
                cell.alignment = align_center
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
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 3, min_w), max_w)

    ws.freeze_panes = "A2"
    if n_cols > 0:
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


def build_excel(missing_sap: pd.DataFrame, missing_tour: pd.DataFrame, all_diff: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_diff.to_excel(writer, index=False, sheet_name="Alle Unterschiede", na_rep="")
        _format_sheet(writer, "Alle Unterschiede", all_diff)

        missing_sap.to_excel(writer, index=False, sheet_name="Fehlt in SAP", na_rep="")
        _format_sheet(writer, "Fehlt in SAP", missing_sap)

        missing_tour.to_excel(writer, index=False, sheet_name="Fehlt in Tour", na_rep="")
        _format_sheet(writer, "Fehlt in Tour", missing_tour)

        wb = writer.book
        for ws in wb.worksheets:
            ws.sheet_state = "visible"

    return output.getvalue()

# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------


st.set_page_config(page_title="SAP gegen Tourenplanung", layout="wide")

st.title("SAP gegen Tourenplanung")
st.caption("Prüfung nur für NMS, Malchow und die ausgewählten Direkt-Touren. Direkt wird hart auf Kunden aus diesen Touren begrenzt.")

st.markdown(
    """
<div style="padding:14px 16px;border:1px solid #d9e2ec;border-radius:12px;background:#f8fafc;margin-bottom:14px;">
<b>Geprüft wird:</b><br>
NMS komplett · Malchow komplett · Direkt nur mit den Tournummern
<b>1058, 2058, 3058, 4058, 5058 und 6030</b>.<br><br>
<b>Die Tournummern werden jetzt aus jeder Tageszelle einzeln gelesen.</b>
Wenn in einer Zelle zum Beispiel <b>6029 / 6030</b> steht, wird für Direkt nur <b>6030</b> bewertet.
Wenn die SAP-Datei eine Tournummer- oder Transportgruppen-Spalte enthält, wird exakt nach
<b>SAP Nummer + Liefertag + Tournummer</b> verglichen. Ohne SAP-Tournummer wird Direkt in der Rückrichtung nicht pauschal über alle SAP-Tage geprüft.
</div>
""",
    unsafe_allow_html=True,
)

col_upload_1, col_upload_2 = st.columns(2)
with col_upload_1:
    sap_datei = st.file_uploader(
        "SAP hochladen",
        type=["xlsx", "xlsm", "xls"],
        key="sap_datei",
        help="Standard: SAP Nummer in Spalte A, Liefertag in Spalte G. Eine Tournummer-Spalte wird automatisch erkannt, falls vorhanden.",
    )
with col_upload_2:
    tourenplanung_datei = st.file_uploader(
        "Tourenplanung hochladen",
        type=["xlsx", "xlsm", "xls"],
        key="tourenplanung_datei",
        help="Geprüft werden nur NMS, Malchow und Direkt. Direkt wird auf die Tournummern 1058, 2058, 3058, 4058, 5058 und 6030 gefiltert.",
    )

run = st.button("Unterschiede prüfen", type="primary", use_container_width=True)

if run:
    if not sap_datei or not tourenplanung_datei:
        st.error("Bitte beide Excel-Dateien hochladen.")
        st.stop()

    try:
        days_by_sap, sap_records, sap_sheet, sap_rows, sap_tour_column = read_sap_file(sap_datei)
        tour_df, checked_sheets, missing_sheets, customer_info = read_tourenplanung(tourenplanung_datei)

        if sap_rows == 0:
            st.warning("In der SAP-Datei wurden keine gültigen Liefertage erkannt.")
        if tour_df.empty:
            st.warning("In der Tourenplanung wurden im geprüften Bereich keine gültigen Tournummern erkannt.")
        if missing_sheets:
            st.warning("Diese Blätter wurden nicht gefunden: " + ", ".join(missing_sheets))

        missing_sap = build_missing_in_sap(tour_df, days_by_sap, sap_records, customer_info)
        missing_tour = build_missing_in_tour(tour_df, days_by_sap, sap_records, customer_info)
        all_diff = build_all_differences(missing_sap, missing_tour)
        excel_bytes = build_excel(missing_sap, missing_tour, all_diff)

        st.session_state["result"] = {
            "days_by_sap": days_by_sap,
            "sap_records": sap_records,
            "sap_sheet": sap_sheet,
            "sap_rows": sap_rows,
            "sap_tour_column": sap_tour_column,
            "tour_df": tour_df,
            "checked_sheets": checked_sheets,
            "missing_sheets": missing_sheets,
            "missing_sap": missing_sap,
            "missing_tour": missing_tour,
            "all_diff": all_diff,
            "excel_bytes": excel_bytes,
        }
    except Exception as exc:
        import traceback
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
        with st.expander("Technische Details", expanded=False):
            st.code(traceback.format_exc(), language="python")
        st.session_state.pop("result", None)

result = st.session_state.get("result")
if result:
    missing_sap = result["missing_sap"]
    missing_tour = result["missing_tour"]
    all_diff = result["all_diff"]
    tour_df = result["tour_df"]

    st.divider()
    head_left, head_right = st.columns([3, 1])
    with head_left:
        st.subheader("Ergebnis")
        tour_col_text = result["sap_tour_column"] or "keine Tournummer-Spalte erkannt"
        st.caption(
            f"SAP: Blatt {result['sap_sheet']}, {result['sap_rows']} gültige SAP-Zeilen · "
            f"SAP-Tournummer-Spalte: {tour_col_text} · "
            f"Tourenplanung: {', '.join(result['checked_sheets'])}"
        )
    with head_right:
        st.download_button(
            label="Excel herunterladen",
            data=result["excel_bytes"],
            file_name="sap_tourenplanung_nms_malchow_direkt_tournummern_exakt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Alle Unterschiede", len(all_diff))
    m2.metric("Fehlt in SAP", len(missing_sap))
    m3.metric("Fehlt in Tour", len(missing_tour))
    m4.metric("Geprüfte Tour-Zeilen", len(tour_df))

    overview = build_overview(missing_sap, missing_tour, tour_df)

    tab_overview, tab_all, tab_sap, tab_tour, tab_scope = st.tabs([
        "Übersicht",
        f"Alle Unterschiede ({len(all_diff)})",
        f"Fehlt in SAP ({len(missing_sap)})",
        f"Fehlt in Tour ({len(missing_tour)})",
        f"Geprüfter Tourbereich ({len(tour_df)})",
    ])

    with tab_overview:
        st.dataframe(overview, use_container_width=True, hide_index=True)
        if all_diff.empty:
            st.success("Keine Unterschiede gefunden.")

    def render_result_table(df: pd.DataFrame, key_prefix: str) -> None:
        if df is None or df.empty:
            st.info("Keine Treffer.")
            return
        filter_col_1, filter_col_2 = st.columns([1, 2])
        bereiche = ["Alle"] + sorted(df["Bereich"].dropna().unique().tolist())
        bereich = filter_col_1.selectbox("Bereich", bereiche, key=f"{key_prefix}_bereich")
        suche = filter_col_2.text_input("Suchen", key=f"{key_prefix}_suche")
        filtered = _filter_dataframe(df, suche, bereich)
        st.caption(f"{len(filtered)} von {len(df)} Zeilen")
        st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tab_all:
        render_result_table(all_diff, "all")

    with tab_sap:
        render_result_table(missing_sap, "sap")

    with tab_tour:
        render_result_table(missing_tour, "tour")

    with tab_scope:
        if tour_df.empty:
            st.info("Keine geprüften Tourzeilen.")
        else:
            scope_view = tour_df.rename(columns={
                "sap": "SAP Nummer",
                "tag_num": "Tag Nummer",
            }).copy()
            scope_view["Liefertag"] = scope_view["Tag Nummer"].map(day_to_text)
            scope_view = scope_view[[
                "Bereich", "Blatt Tourenplanung", "SAP Nummer", "Liefertag", "Tournummer Tour"
            ]].sort_values(["Bereich", "Blatt Tourenplanung", "SAP Nummer", "Liefertag", "Tournummer Tour"])
            suche = st.text_input("Suchen im geprüften Bereich", key="scope_suche")
            if suche:
                mask = pd.Series(False, index=scope_view.index)
                for col in scope_view.columns:
                    mask = mask | scope_view[col].astype(str).str.lower().str.contains(suche.lower(), na=False, regex=False)
                scope_view = scope_view[mask]
            st.dataframe(scope_view, use_container_width=True, hide_index=True)
