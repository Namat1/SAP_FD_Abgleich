import base64
import html
import io
import re
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Feste Prüfregeln
# ---------------------------------------------------------------------------

DAY_NAMES = {
    1: "Montag",
    2: "Dienstag",
    3: "Mittwoch",
    4: "Donnerstag",
    5: "Freitag",
    6: "Samstag",
}

DAY_SHORT = {
    1: "Mo",
    2: "Di",
    3: "Mi",
    4: "Do",
    5: "Fr",
    6: "Sa",
}

DAY_COLUMNS_TOUR = {
    1: 6,   # G = Mo
    2: 7,   # H = Di
    3: 8,   # I = Mi
    4: 9,   # J = Do
    5: 10,  # K = Fr
    6: 11,  # L = Sa
}

DAY_COLUMN_CANDIDATES = {
    1: ["mo", "montag"],
    2: ["die", "di", "dienstag"],
    3: ["mitt", "mit", "mi", "mittwoch"],
    4: ["don", "do", "donnerstag"],
    5: ["fr", "frei", "freitag"],
    6: ["sam", "sa", "samstag"],
}

# Aus DIREKT werden ausschließlich diese Touren berücksichtigt.
DIRECT_ALLOWED_TOURS: Set[str] = {"1058", "2058", "3058", "4058", "5058", "6030"}

SAP_COL_INDEX = 0       # SAP-Datei Fallback: A = SAP Nummer
SAP_DAY_COL_INDEX = 6   # SAP-Datei Fallback: G = Liefertag
TOUR_CSB_COL_INDEX = 0  # Tourenplanung Fallback: A = CSB
TOUR_SAP_COL_INDEX = 1  # Tourenplanung Fallback: B = SAP


# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------

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


def value_to_clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_sap_series(series: pd.Series) -> pd.Series:
    if series.empty:
        return series.astype(str)
    numeric = pd.to_numeric(series, errors="coerce")
    is_int = numeric.notna() & (numeric == numeric.round())
    out = series.astype(str)
    out = out.where(~is_int, numeric.where(is_int).astype("Int64").astype(str))
    out = out.str.strip()
    return out.replace({"nan": "", "<NA>": "", "None": "", "NaT": ""})


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


def pick_first_matching_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    candidate_set = {normalize_header_name(c) for c in candidates}
    for column in columns:
        if normalize_header_name(column) in candidate_set:
            return column
    return None


def pick_column_by_name_or_position(
    columns: List[str],
    candidates: List[str],
    fallback_index: Optional[int],
) -> Optional[str]:
    found = pick_first_matching_column(columns, candidates)
    if found is not None:
        return found
    if fallback_index is not None and len(columns) > fallback_index:
        return columns[fallback_index]
    return None


def make_unique_columns(raw_columns: List[object]) -> List[str]:
    result: List[str] = []
    seen: Dict[str, int] = {}
    for index, value in enumerate(raw_columns, start=1):
        name = value_to_clean_text(value) or f"Spalte_{index}"
        count = seen.get(name, 0) + 1
        seen[name] = count
        if count > 1:
            name = f"{name}_{count}"
        result.append(name)
    return result


def read_excel_with_detected_header(excel: pd.ExcelFile, sheet_name: str, kind: str) -> pd.DataFrame:
    raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=object)
    if raw.empty:
        return pd.DataFrame()

    day_names_flat = {
        normalize_header_name(candidate)
        for values in DAY_COLUMN_CANDIDATES.values()
        for candidate in values
    }

    header_row: Optional[int] = None
    for row_index in range(min(len(raw), 30)):
        values = [normalize_header_name(value) for value in raw.iloc[row_index].tolist()]
        value_set = set(values)
        has_sap = bool(value_set & {"sap", "sapnummer", "sapnr", "sapnum", "kundennummer", "kundennr"})

        if kind == "tour":
            day_hits = sum(1 for value in values if value in day_names_flat)
            if has_sap and day_hits >= 2:
                header_row = row_index
                break
        else:
            has_day = bool(value_set & {"liefertag", "liefertagcode", "liefercode", "lt", "tag"})
            if has_sap and has_day:
                header_row = row_index
                break

    if header_row is None:
        return pd.read_excel(excel, sheet_name=sheet_name, header=0, dtype=object)

    df = raw.iloc[header_row + 1:].copy()
    df.columns = make_unique_columns(raw.iloc[header_row].tolist())
    return df.dropna(how="all").reset_index(drop=True)


def day_value_is_set(value) -> bool:
    if value is None or pd.isna(value):
        return False
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "<na>", "-", "--"}:
        return False
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(number) and float(number) == 0:
        return False
    return True


def extract_tour_numbers(value) -> List[str]:
    if value is None or pd.isna(value):
        return []
    return re.findall(r"\d+", value_to_clean_text(value))


def select_relevant_sheets(excel: pd.ExcelFile) -> Dict[str, str]:
    """Prüfbereich wie in fd_Vergleich.py: NMS, Malchow und ausgewählte DIREKT-Touren. MK bleibt außen vor."""
    result: Dict[str, str] = {}
    for sheet_name in excel.sheet_names:
        norm = normalize_header_name(sheet_name)
        if "direkt" in norm:
            result.setdefault("Direkt", sheet_name)
        elif "nms" in norm or "neumuenster" in norm or "neumunster" in norm:
            result.setdefault("NMS", sheet_name)
        elif "malchow" in norm:
            result.setdefault("Malchow", sheet_name)
    return result


def get_tour_columns(df: pd.DataFrame) -> Dict[int, str]:
    columns = list(df.columns)
    result: Dict[int, str] = {}
    for day_num, fallback_index in DAY_COLUMNS_TOUR.items():
        col = pick_column_by_name_or_position(columns, DAY_COLUMN_CANDIDATES[day_num], fallback_index)
        if col is not None:
            result[day_num] = col
    return result


def days_text(days: Set[int]) -> str:
    if not days:
        return "–"
    return ", ".join(DAY_SHORT[d] for d in sorted(days))


def result_columns() -> List[str]:
    return [
        "Bereich",
        "CSB",
        "SAP Nummer",
        "Name",
        "Ort",
        "Liefertage Tour",
        "Fehlt in SAP",
    ]


# ---------------------------------------------------------------------------
# SAP lesen
# ---------------------------------------------------------------------------

def read_sap_file(uploaded_file) -> Tuple[Dict[str, Set[int]], Set[str], str, int]:
    """Liest SAP Nummer + Liefertag.

    sap_customers enthält ALLE gefundenen SAP-Nummern, auch wenn bei einem Kunden
    kein gültiger Liefertag hinterlegt ist. So wird 'Kunde fehlt' sauber erkannt.
    """
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = read_excel_with_detected_header(excel, sheet_name, kind="sap")
    if df.empty:
        return {}, set(), sheet_name, 0

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

    if sap_column is None:
        return {}, set(), sheet_name, 0

    sap_series = normalize_sap_series(df[sap_column])
    sap_customers = set(sap_series[sap_series.ne("")].astype(str))

    if day_column is None:
        return {}, sap_customers, sheet_name, len(sap_customers)

    work = pd.DataFrame({"sap": sap_series, "tag": df[day_column]})
    work["tag_num"] = normalize_day_code_series(work["tag"])
    mask = work["sap"].ne("") & work["tag_num"].notna() & work["tag_num"].between(1, 6, inclusive="both")
    filtered = work.loc[mask, ["sap", "tag_num"]].copy()
    filtered["tag_int"] = filtered["tag_num"].astype(int)

    days_by_sap: Dict[str, Set[int]] = filtered.groupby("sap")["tag_int"].agg(set).to_dict()
    return days_by_sap, sap_customers, sheet_name, len(sap_customers)


# ---------------------------------------------------------------------------
# Tourendatei lesen
# ---------------------------------------------------------------------------

def read_tour_customers(uploaded_file) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Erzeugt genau einen Datensatz pro relevantem Tour-Kunden.

    NMS: alle Kunden im Blatt
    Malchow: alle Kunden im Blatt
    Direkt: nur Kunden, die mindestens auf einer der Touren
            1058/2058/3058/4058/5058/6030 stehen
    """
    excel = pd.ExcelFile(uploaded_file)
    selected_sheets = select_relevant_sheets(excel)
    rows: List[dict] = []

    for bereich, sheet_name in selected_sheets.items():
        df = read_excel_with_detected_header(excel, sheet_name, kind="tour")
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

        csb_column = pick_column_by_name_or_position(columns, ["CSB", "CSB Nummer", "CSB-Nr", "CSB Nr"], TOUR_CSB_COL_INDEX)
        name_column = pick_column_by_name_or_position(columns, ["Name", "Kundenname", "Marktname", "Kunde", "Bezeichnung", "Filialname"], 2)
        strasse_column = pick_column_by_name_or_position(columns, ["Strasse", "Straße", "Str", "Anschrift", "Adresse", "Strassenname", "Straßenname", "Strasse Hausnummer", "Straße Hausnummer"], 3)
        plz_column = pick_column_by_name_or_position(columns, ["Plz", "PLZ", "Postleitzahl"], 4)
        ort_column = pick_column_by_name_or_position(columns, ["Ort", "Stadt", "Plz Ort", "PLZ Ort", "Ortname"], 5)
        day_columns = get_tour_columns(df)

        work = df.copy()
        work["_sap"] = normalize_sap_series(work[sap_column])

        for _, row in work.iterrows():
            sap = value_to_clean_text(row.get("_sap", ""))
            if not sap:
                continue

            tour_days: Set[int] = set()
            has_allowed_direct = False

            for day_num, day_column in day_columns.items():
                cell_value = row.get(day_column, "")
                if not day_value_is_set(cell_value):
                    continue

                if bereich == "Direkt":
                    numbers = set(extract_tour_numbers(cell_value))
                    if numbers & DIRECT_ALLOWED_TOURS:
                        has_allowed_direct = True
                        tour_days.add(day_num)
                else:
                    tour_days.add(day_num)

            if bereich == "Direkt" and not has_allowed_direct:
                continue

            rows.append({
                "Bereich": bereich,
                "Blatt Tourenplanung": sheet_name,
                "CSB": value_to_clean_text(row.get(csb_column, "")) if csb_column else "",
                "SAP Nummer": sap,
                "Name": value_to_clean_text(row.get(name_column, "")) if name_column else "",
                "Straße": value_to_clean_text(row.get(strasse_column, "")) if strasse_column else "",
                "PLZ": value_to_clean_text(row.get(plz_column, "")) if plz_column else "",
                "Ort": value_to_clean_text(row.get(ort_column, "")) if ort_column else "",
                "_tour_days": tour_days,
            })

    if not rows:
        empty_cols = ["Bereich", "Blatt Tourenplanung", "CSB", "SAP Nummer", "Name", "Straße", "PLZ", "Ort", "_tour_days"]
        return pd.DataFrame(columns=empty_cols), selected_sheets

    # Falls ein Kunde mehrfach vorkommt: Stammdaten vom ersten Treffer, Liefertage vereinigen.
    grouped_rows: List[dict] = []
    temp = pd.DataFrame(rows)
    for sap, group in temp.groupby("SAP Nummer", sort=False):
        first = group.iloc[0]
        combined_days: Set[int] = set()
        for value in group["_tour_days"]:
            combined_days |= set(value)

        grouped_rows.append({
            "Bereich": ", ".join(sorted(set(group["Bereich"].astype(str)))),
            "Blatt Tourenplanung": ", ".join(sorted(set(group["Blatt Tourenplanung"].astype(str)))),
            "CSB": value_to_clean_text(first.get("CSB", "")),
            "SAP Nummer": str(sap),
            "Name": value_to_clean_text(first.get("Name", "")),
            "Straße": value_to_clean_text(first.get("Straße", "")),
            "PLZ": value_to_clean_text(first.get("PLZ", "")),
            "Ort": value_to_clean_text(first.get("Ort", "")),
            "_tour_days": combined_days,
        })

    out = pd.DataFrame(grouped_rows)
    out["_sap_sort"] = pd.to_numeric(out["SAP Nummer"], errors="coerce").fillna(9_999_999_999)
    out = out.sort_values(["Bereich", "_sap_sort"]).drop(columns=["_sap_sort"]).reset_index(drop=True)
    return out, selected_sheets


# ---------------------------------------------------------------------------
# Vergleich: ausschließlich Tourendatei -> SAP
# ---------------------------------------------------------------------------

def build_differences(
    tour_customers: pd.DataFrame,
    sap_days_by_customer: Dict[str, Set[int]],
    sap_customers: Set[str],
) -> pd.DataFrame:
    """Prüft ausschließlich: Sind alle Liefertage aus der Tourendatei in SAP vorhanden?

    Zusätzliche Liefertage in SAP werden bewusst ignoriert.
    Fehlt ein Kunde komplett in SAP, gelten seine Tour-Liefertage als fehlend.
    """
    if tour_customers.empty:
        return pd.DataFrame(columns=result_columns())

    rows: List[dict] = []

    for _, row in tour_customers.iterrows():
        sap = str(row["SAP Nummer"])
        tour_days = set(row["_tour_days"])
        sap_days = set(sap_days_by_customer.get(sap, set())) if sap in sap_customers else set()

        missing_in_sap = tour_days - sap_days
        if not missing_in_sap:
            continue

        rows.append({
            "Bereich": row.get("Bereich", ""),
            "CSB": row.get("CSB", ""),
            "SAP Nummer": sap,
            "Name": row.get("Name", ""),
            "Ort": row.get("Ort", ""),
            "Liefertage Tour": days_text(tour_days),
            "Fehlt in SAP": days_text(missing_in_sap),
        })

    if not rows:
        return pd.DataFrame(columns=result_columns())

    out = pd.DataFrame(rows, columns=result_columns())
    order = {"NMS": 1, "Malchow": 2, "Direkt": 3}
    out["_bereich_sort"] = out["Bereich"].map(order).fillna(99)
    out["_sap_sort"] = pd.to_numeric(out["SAP Nummer"], errors="coerce").fillna(9_999_999_999)
    out = out.sort_values(["_bereich_sort", "_sap_sort"]).drop(columns=["_bereich_sort", "_sap_sort"])
    return out.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Excel-Ausgabe: nur die relevante Abweichungsliste
# ---------------------------------------------------------------------------

def build_excel(differences: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        differences.to_excel(writer, index=False, sheet_name="Fehlende Liefertage", na_rep="")
        ws = writer.sheets["Fehlende Liefertage"]

        header_fill = PatternFill(start_color="FF374151", end_color="FF374151", fill_type="solid")
        missing_fill = PatternFill(start_color="FFFEE2E2", end_color="FFFEE2E2", fill_type="solid")
        diff_fill = PatternFill(start_color="FFFFF7D6", end_color="FFFFF7D6", fill_type="solid")
        thin = Side(style="thin", color="FFD1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
        body_font = Font(name="Calibri", size=10)

        columns = list(differences.columns)
        for col_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border
        ws.row_dimensions[1].height = 24

        for row_idx in range(2, len(differences) + 2):
            for col_idx in range(1, len(columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = body_font
                cell.border = border
                cell.fill = diff_fill
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            ws.row_dimensions[row_idx].height = 20

        width_hints = {
            "Bereich": 14,
            "CSB": 11,
            "SAP Nummer": 13,
            "Name": 34,
            "Ort": 25,
            "Liefertage Tour": 22,
            "Fehlt in SAP": 22,
        }
        for col_idx, col_name in enumerate(columns, start=1):
            width = width_hints.get(col_name, 20)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        ws.freeze_panes = "A2"
        if len(columns) > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(differences) + 1}"
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"

    return output.getvalue()



# ---------------------------------------------------------------------------
# HTML-Ausgabe: eigenständiger Bericht inkl. eingebettetem Excel-Download
# ---------------------------------------------------------------------------

def build_html_report(
    differences: pd.DataFrame,
    tour_customer_count: int,
    sap_sheet: str,
    selected_sheets: Dict[str, str],
    excel_bytes: bytes,
) -> bytes:
    """Erzeugt eine portable HTML-Auswertung.

    Die Excel-Auswertung wird als Base64 direkt in die HTML-Datei eingebettet.
    Dadurch kann die HTML-Datei lokal geöffnet und die Excel-Datei ohne
    weitere Serververbindung über den Download-Button gespeichert werden.
    """
    excel_b64 = base64.b64encode(excel_bytes).decode("ascii")
    diff_count = len(differences)
    ok_count = max(tour_customer_count - diff_count, 0)

    status_class = "success" if diff_count == 0 else "warning"
    status_title = "Keine Abweichungen gefunden" if diff_count == 0 else f"{diff_count} Kunde(n) mit Abweichung"
    status_text = (
        "Alle Liefertage aus der Tourendatei sind in SAP vorhanden."
        if diff_count == 0
        else "Bei diesen Kunden fehlt mindestens ein Liefertag aus der Tourendatei in SAP."
    )

    sheet_parts = []
    for bereich in ["NMS", "Malchow", "Direkt"]:
        if bereich in selected_sheets:
            sheet_parts.append(f"<span class='chip'><b>{html.escape(bereich)}</b>: {html.escape(str(selected_sheets[bereich]))}</span>")
    sheets_html = "".join(sheet_parts) or "<span class='chip'>Keine relevanten Tourenblätter erkannt</span>"

    if differences.empty:
        table_html = """
        <div class="empty-state">
            <div class="empty-icon">✓</div>
            <h3>Alles in Ordnung</h3>
            <p>Es wurden keine fehlenden Liefertage in SAP gefunden.</p>
        </div>
        """
        search_html = ""
    else:
        rows = []
        for _, row in differences.iterrows():
            rows.append(
                "<tr>"
                f"<td><span class='area'>{html.escape(str(row.get('Bereich', '')))}</span></td>"
                f"<td>{html.escape(str(row.get('CSB', '')))}</td>"
                f"<td class='mono'>{html.escape(str(row.get('SAP Nummer', '')))}</td>"
                f"<td class='name'>{html.escape(str(row.get('Name', '')))}</td>"
                f"<td>{html.escape(str(row.get('Ort', '')))}</td>"
                f"<td>{html.escape(str(row.get('Liefertage Tour', '')))}</td>"
                f"<td><span class='missing'>{html.escape(str(row.get('Fehlt in SAP', '')))}</span></td>"
                "</tr>"
            )
        search_html = """
        <div class="toolbar">
            <input id="searchInput" type="search" placeholder="Kunde, SAP, CSB, Ort oder Liefertag suchen …" oninput="filterTable()">
            <span id="resultCount"></span>
        </div>
        """
        table_html = f"""
        <div class="table-wrap">
            <table id="resultTable">
                <thead>
                    <tr>
                        <th>Bereich</th>
                        <th>CSB</th>
                        <th>SAP Nummer</th>
                        <th>Name</th>
                        <th>Ort</th>
                        <th>Liefertage Tour</th>
                        <th>Fehlt in SAP</th>
                    </tr>
                </thead>
                <tbody>{''.join(rows)}</tbody>
            </table>
        </div>
        """

    report = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Liefertage Tour → SAP</title>
<style>
:root {{
    --bg: #f5f6f8;
    --card: #ffffff;
    --text: #172033;
    --muted: #687386;
    --line: #e4e7ec;
    --accent: #6d55c7;
    --accent-dark: #5842ad;
    --good: #157347;
    --good-bg: #eaf7f0;
    --warn: #9a5a00;
    --warn-bg: #fff6df;
    --bad: #b42318;
    --bad-bg: #fff0ee;
}}
* {{ box-sizing: border-box; }}
body {{ margin: 0; background: var(--bg); color: var(--text); font-family: Inter, Segoe UI, Arial, sans-serif; }}
.page {{ max-width: 1380px; margin: 0 auto; padding: 34px 24px 48px; }}
.header {{ display: flex; justify-content: space-between; gap: 24px; align-items: flex-start; margin-bottom: 24px; }}
.eyebrow {{ color: var(--accent); font-weight: 800; font-size: 12px; letter-spacing: .08em; text-transform: uppercase; margin-bottom: 8px; }}
h1 {{ margin: 0 0 8px; font-size: clamp(28px, 4vw, 42px); letter-spacing: -.03em; }}
.subtitle {{ color: var(--muted); font-size: 16px; max-width: 850px; line-height: 1.5; }}
.download {{ display: inline-flex; align-items: center; justify-content: center; min-height: 48px; padding: 0 18px; border-radius: 12px; background: var(--accent); color: white; font-weight: 800; text-decoration: none; box-shadow: 0 8px 20px rgba(70, 50, 140, .16); white-space: nowrap; }}
.download:hover {{ background: var(--accent-dark); }}
.grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; margin: 20px 0; }}
.card {{ background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 20px; box-shadow: 0 3px 12px rgba(20, 30, 50, .04); }}
.metric-label {{ color: var(--muted); font-size: 13px; font-weight: 700; margin-bottom: 8px; }}
.metric-value {{ font-size: 34px; font-weight: 850; letter-spacing: -.04em; }}
.metric-note {{ color: var(--muted); font-size: 13px; margin-top: 6px; }}
.status {{ border-radius: 14px; padding: 16px 18px; margin: 18px 0; border: 1px solid; }}
.status.success {{ background: var(--good-bg); border-color: #b7e2c9; color: var(--good); }}
.status.warning {{ background: var(--warn-bg); border-color: #f0d391; color: var(--warn); }}
.status strong {{ display: block; font-size: 16px; margin-bottom: 3px; }}
.section-title {{ margin: 28px 0 12px; font-size: 20px; }}
.info {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }}
.chip {{ background: #f1effa; color: #443589; border-radius: 999px; padding: 7px 11px; font-size: 12px; }}
.rules {{ color: var(--muted); font-size: 14px; line-height: 1.55; margin-top: 10px; }}
.toolbar {{ display: flex; justify-content: space-between; align-items: center; gap: 14px; margin: 14px 0 10px; }}
#searchInput {{ width: min(520px, 100%); border: 1px solid var(--line); background: white; border-radius: 12px; padding: 12px 14px; font-size: 14px; outline: none; }}
#searchInput:focus {{ border-color: #a99be0; box-shadow: 0 0 0 3px #eeeafc; }}
#resultCount {{ color: var(--muted); font-size: 13px; font-weight: 700; }}
.table-wrap {{ overflow-x: auto; background: var(--card); border: 1px solid var(--line); border-radius: 16px; box-shadow: 0 3px 12px rgba(20, 30, 50, .04); }}
table {{ width: 100%; border-collapse: collapse; min-width: 980px; }}
th {{ position: sticky; top: 0; background: #252a36; color: #fff; text-align: left; font-size: 12px; letter-spacing: .02em; padding: 13px 14px; }}
td {{ padding: 12px 14px; border-top: 1px solid var(--line); font-size: 13px; vertical-align: middle; }}
tbody tr:hover {{ background: #faf9fe; }}
.name {{ font-weight: 700; }}
.mono {{ font-variant-numeric: tabular-nums; font-family: ui-monospace, SFMono-Regular, Consolas, monospace; }}
.area {{ display: inline-block; background: #f0f2f5; border-radius: 999px; padding: 5px 9px; font-weight: 750; }}
.missing {{ display: inline-block; background: var(--bad-bg); color: var(--bad); border: 1px solid #fac7c2; border-radius: 8px; padding: 5px 8px; font-weight: 800; }}
.empty-state {{ background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 48px 24px; text-align: center; }}
.empty-icon {{ width: 52px; height: 52px; margin: 0 auto 12px; display: grid; place-items: center; border-radius: 50%; background: var(--good-bg); color: var(--good); font-size: 28px; font-weight: 900; }}
.empty-state h3 {{ margin: 0 0 6px; }}
.empty-state p {{ margin: 0; color: var(--muted); }}
.footer {{ color: var(--muted); font-size: 12px; margin-top: 26px; text-align: center; }}
@media (max-width: 820px) {{
    .page {{ padding: 22px 14px 36px; }}
    .header {{ flex-direction: column; }}
    .download {{ width: 100%; }}
    .grid {{ grid-template-columns: 1fr; }}
    .toolbar {{ flex-direction: column; align-items: stretch; }}
    #searchInput {{ width: 100%; }}
}}
</style>
</head>
<body>
<div class="page">
    <div class="header">
        <div>
            <div class="eyebrow">Tourenplanung · SAP-Abgleich</div>
            <h1>Liefertage Tour → SAP</h1>
            <div class="subtitle">Die Tourendatei ist die Vorgabe. Geprüft wird ausschließlich, ob jeder dort eingetragene Liefertag auch in SAP vorhanden ist.</div>
        </div>
        <a class="download" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="Fehlende_Liefertage_in_SAP.xlsx">Excel herunterladen</a>
    </div>

    <div class="grid">
        <div class="card">
            <div class="metric-label">Geprüfte Tour-Kunden</div>
            <div class="metric-value">{tour_customer_count}</div>
            <div class="metric-note">relevante Kunden aus der Tourendatei</div>
        </div>
        <div class="card">
            <div class="metric-label">Ohne Abweichung</div>
            <div class="metric-value">{ok_count}</div>
            <div class="metric-note">Liefertage vollständig in SAP vorhanden</div>
        </div>
        <div class="card">
            <div class="metric-label">Mit fehlendem Liefertag</div>
            <div class="metric-value">{diff_count}</div>
            <div class="metric-note">Kunden in der Fehlerliste</div>
        </div>
    </div>

    <div class="status {status_class}">
        <strong>{html.escape(status_title)}</strong>
        {html.escape(status_text)}
    </div>

    <div class="card">
        <div class="metric-label">Geprüfte Datenbasis</div>
        <div><b>SAP-Blatt:</b> {html.escape(str(sap_sheet))}</div>
        <div class="info">{sheets_html}</div>
        <div class="rules">NMS komplett · Malchow komplett · Direkt nur Touren 1058, 2058, 3058, 4058, 5058 und 6030. Zusätzliche Liefertage in SAP werden bewusst ignoriert.</div>
    </div>

    <h2 class="section-title">Fehlende Liefertage in SAP</h2>
    {search_html}
    {table_html}

    <div class="footer">Erstellt mit der Auswertung „Liefertage Tour → SAP“</div>
</div>
<script>
function filterTable() {{
    const input = document.getElementById('searchInput');
    const table = document.getElementById('resultTable');
    const count = document.getElementById('resultCount');
    if (!input || !table) return;
    const term = input.value.toLowerCase().trim();
    let visible = 0;
    for (const row of table.tBodies[0].rows) {{
        const show = row.innerText.toLowerCase().includes(term);
        row.style.display = show ? '' : 'none';
        if (show) visible++;
    }}
    if (count) count.textContent = visible + ' Treffer';
}}
filterTable();
</script>
</body>
</html>"""
    return report.encode("utf-8")


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Tourendatei gegen SAP", layout="wide")

st.markdown(
    """
    <style>
        .block-container { padding-top: 1.7rem; padding-bottom: 2rem; max-width: 1350px; }
        [data-testid="stMetric"] { border: 1px solid #e5e7eb; border-radius: 12px; padding: 12px 14px; }
        [data-testid="stFileUploader"] section { border-radius: 12px; }
        div.stButton > button { height: 3rem; border-radius: 10px; font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Liefertage Tour → SAP")
st.caption(
    "Die Tourendatei ist die Vorgabe. Geprüft wird ausschließlich, ob jeder dort eingetragene Liefertag auch in SAP vorhanden ist."
)

with st.container(border=True):
    st.markdown("**Prüfbereich**")
    st.write(
        "NMS komplett · Malchow komplett · Direkt nur Touren "
        "1058, 2058, 3058, 4058, 5058 und 6030. "
        "Zusätzliche Liefertage in SAP werden ignoriert."
    )

col1, col2 = st.columns(2)
with col1:
    sap_datei = st.file_uploader(
        "SAP-Datei",
        type=["xlsx", "xlsm", "xls"],
        key="sap_datei",
        help="Benötigt werden SAP Nummer und Liefertag.",
    )
with col2:
    tour_datei = st.file_uploader(
        "Tourendatei",
        type=["xlsx", "xlsm", "xls"],
        key="tour_datei",
        help="NMS und Malchow komplett; Direkt nur die sechs festgelegten Touren.",
    )

run = st.button("Liefertage prüfen", type="primary", use_container_width=True)

if run:
    if not sap_datei or not tour_datei:
        st.error("Bitte SAP-Datei und Tourendatei hochladen.")
        st.stop()

    try:
        sap_days, sap_customers, sap_sheet, sap_customer_count = read_sap_file(sap_datei)
        tour_customers, selected_sheets = read_tour_customers(tour_datei)
        differences = build_differences(tour_customers, sap_days, sap_customers)
        excel_bytes = build_excel(differences)
        html_bytes = build_html_report(
            differences=differences,
            tour_customer_count=len(tour_customers),
            sap_sheet=sap_sheet,
            selected_sheets=selected_sheets,
            excel_bytes=excel_bytes,
        )

        st.session_state["tour_sap_result"] = {
            "sap_sheet": sap_sheet,
            "selected_sheets": selected_sheets,
            "tour_customer_count": len(tour_customers),
            "differences": differences,
            "customers_with_missing_days": len(differences),
            "excel_bytes": excel_bytes,
            "html_bytes": html_bytes,
        }
    except Exception as exc:
        import traceback
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
        with st.expander("Technische Details", expanded=False):
            st.code(traceback.format_exc(), language="python")
        st.session_state.pop("tour_sap_result", None)

result = st.session_state.get("tour_sap_result")
if result:
    st.divider()

    top_left, top_mid, top_right = st.columns([3, 1, 1])
    with top_left:
        st.subheader("Ergebnis")
        st.caption(
            f"{result['tour_customer_count']} relevante Tour-Kunden geprüft · "
            f"SAP-Blatt: {result['sap_sheet']}"
        )
    with top_mid:
        st.download_button(
            "HTML-Auswertung",
            data=result["html_bytes"],
            file_name="Tour_SAP_Auswertung.html",
            mime="text/html",
            use_container_width=True,
            type="primary",
        )
    with top_right:
        st.download_button(
            "Excel herunterladen",
            data=result["excel_bytes"],
            file_name="Fehlende_Liefertage_in_SAP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    c1, c2 = st.columns(2)
    c1.metric("Geprüfte Tour-Kunden", result["tour_customer_count"])
    c2.metric("Kunden mit fehlendem Liefertag in SAP", result["customers_with_missing_days"])

    differences = result["differences"]
    if differences.empty:
        st.success("Alle Liefertage aus der Tourendatei sind in SAP vorhanden.")
    else:
        st.markdown("### Fehlende Liefertage in SAP")
        st.dataframe(
            differences,
            use_container_width=True,
            hide_index=True,
            column_config={
                "SAP Nummer": st.column_config.TextColumn("SAP Nummer"),
                "CSB": st.column_config.TextColumn("CSB"),
            },
        )
