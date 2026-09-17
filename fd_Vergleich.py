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
        "Straße",
        "PLZ",
        "Ort",
        "Tourendatei Soll",
        "SAP Ist",
        "Abweichung",
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
    """Tourendatei ist Soll. SAP wird ausschließlich dagegen geprüft."""
    if tour_customers.empty:
        return pd.DataFrame(columns=result_columns())

    rows: List[dict] = []

    for _, row in tour_customers.iterrows():
        sap = str(row["SAP Nummer"])
        tour_days = set(row["_tour_days"])

        if sap not in sap_customers:
            sap_days: Set[int] = set()
            abweichung = "Kunde fehlt in SAP"
        else:
            sap_days = set(sap_days_by_customer.get(sap, set()))
            missing_in_sap = tour_days - sap_days
            extra_in_sap = sap_days - tour_days

            if not missing_in_sap and not extra_in_sap:
                continue

            parts: List[str] = []
            if missing_in_sap:
                parts.append(f"Fehlt in SAP: {days_text(missing_in_sap)}")
            if extra_in_sap:
                parts.append(f"Zusätzlich in SAP: {days_text(extra_in_sap)}")
            if not parts:
                parts.append("Liefertage weichen ab")
            abweichung = " · ".join(parts)

        rows.append({
            "Bereich": row.get("Bereich", ""),
            "CSB": row.get("CSB", ""),
            "SAP Nummer": sap,
            "Name": row.get("Name", ""),
            "Straße": row.get("Straße", ""),
            "PLZ": row.get("PLZ", ""),
            "Ort": row.get("Ort", ""),
            "Tourendatei Soll": days_text(tour_days),
            "SAP Ist": days_text(sap_days) if sap in sap_customers else "–",
            "Abweichung": abweichung,
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
        differences.to_excel(writer, index=False, sheet_name="Abweichungen", na_rep="")
        ws = writer.sheets["Abweichungen"]

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
            abweichung = str(ws.cell(row=row_idx, column=columns.index("Abweichung") + 1).value or "")
            row_fill = missing_fill if "Kunde fehlt in SAP" in abweichung else diff_fill
            for col_idx in range(1, len(columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = body_font
                cell.border = border
                cell.fill = row_fill
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            ws.row_dimensions[row_idx].height = 20

        width_hints = {
            "Bereich": 14,
            "CSB": 11,
            "SAP Nummer": 13,
            "Name": 34,
            "Straße": 30,
            "PLZ": 9,
            "Ort": 25,
            "Tourendatei Soll": 22,
            "SAP Ist": 22,
            "Abweichung": 46,
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

st.title("Tourendatei → SAP")
st.caption(
    "Die Tourendatei ist der Soll-Stand. Geprüft wird nur, ob diese Kunden in SAP vorhanden sind "
    "und ob ihre Liefertage dort genauso hinterlegt sind."
)

with st.container(border=True):
    st.markdown("**Prüfbereich**")
    st.write(
        "NMS komplett · Malchow komplett · Direkt nur Touren "
        "1058, 2058, 3058, 4058, 5058 und 6030. "
        "Kunden, die nur in SAP stehen, werden nicht geprüft."
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

run = st.button("Abgleich starten", type="primary", use_container_width=True)

if run:
    if not sap_datei or not tour_datei:
        st.error("Bitte SAP-Datei und Tourendatei hochladen.")
        st.stop()

    try:
        sap_days, sap_customers, sap_sheet, sap_customer_count = read_sap_file(sap_datei)
        tour_customers, selected_sheets = read_tour_customers(tour_datei)
        differences = build_differences(tour_customers, sap_days, sap_customers)
        excel_bytes = build_excel(differences)

        missing_customer_count = int(differences["Abweichung"].eq("Kunde fehlt in SAP").sum()) if not differences.empty else 0
        day_diff_count = len(differences) - missing_customer_count

        st.session_state["tour_sap_result"] = {
            "sap_sheet": sap_sheet,
            "sap_customer_count": sap_customer_count,
            "selected_sheets": selected_sheets,
            "tour_customer_count": len(tour_customers),
            "differences": differences,
            "missing_customer_count": missing_customer_count,
            "day_diff_count": day_diff_count,
            "excel_bytes": excel_bytes,
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

    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.subheader("Ergebnis")
        st.caption(
            f"{result['tour_customer_count']} relevante Tour-Kunden geprüft · "
            f"SAP-Blatt: {result['sap_sheet']}"
        )
    with top_right:
        st.download_button(
            "Excel herunterladen",
            data=result["excel_bytes"],
            file_name="Tourendatei_gegen_SAP_Abweichungen.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("Geprüfte Tour-Kunden", result["tour_customer_count"])
    c2.metric("Kunden fehlen in SAP", result["missing_customer_count"])
    c3.metric("Liefertage abweichend", result["day_diff_count"])

    differences = result["differences"]
    if differences.empty:
        st.success("Alle geprüften Kunden aus der Tourendatei sind in SAP vorhanden und die Liefertage stimmen überein.")
    else:
        st.markdown("### Abweichungen")
        st.dataframe(
            differences,
            use_container_width=True,
            hide_index=True,
            column_config={
                "SAP Nummer": st.column_config.TextColumn("SAP Nummer"),
                "CSB": st.column_config.TextColumn("CSB"),
                "PLZ": st.column_config.TextColumn("PLZ"),
            },
        )
