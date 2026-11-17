from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document


TEMPLATE = {
    "Service figures": (2, [("People We Support", 0, 7), ("Referrals", 9, 15)]),
    "Pack_change": (0, [("Pack change", 0, 7)]),
    "Absents": (2, [("Absent Register", 0, 9), ("Grievances", 10, 18)]),
    "Vacancy & Agency": (2, [("Vacancy Info", 0, 6), ("Agency Usage", 7, 12), ("Leavers", 13, 18)]),
    "Safeguarding & WB": (2, [("Safeguarding", 0, 6), ("Whistleblowing", 7, 14), ("Notifications to other agencies", 15, 22)]),
    "Quality, Safety & Compliance": (2, [("Incidents & Accidents", 0, 10), ("Incident detail", 11, 16)]),
    "Complaint & Compliment": (0, [("Complaints & Compliments", 0, 5)]),
    "AOB": (2, [("Key Achievement", 0, 3), ("Any Other Business", 4, 8)]),
    "Dev & Risks": (2, [("Development / Support", 0, 9), ("Risk", 10, 19)]),
    "Audit_Train": (2, [("Audits", 0, 8), ("Training & Supervision", 9, 14), ("Systems & Governance Checks", 15, 19)]),
}
REQUIRED_SHEETS = ["Start_Here", *TEMPLATE.keys(), "Report_Data"]
SOURCE_COLUMNS = {"Source file", "Source sheet", "Source section", "Source row", "_Cover Area", "_Cover Month"}
BOARD_SECTIONS = [
    "Executive summary", "Occupancy and referrals", "Quarter comparison",
    "Improvements, development, risks and issues", "Compliance and monitoring",
    "Complaints and compliments", "Staffing and absence",
]
METRIC_COLUMNS = ["People supported", "Admissions", "Discharges", "Voids", "Referrals listed",
                  "Absence hours", "Agency hours", "Incidents", "Complaints / compliments",
                  "Open risks", "Development items", "Audit records"]


def normalize(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m")
    if isinstance(value, (int, float)) and 20000 <= value <= 80000:
        return (pd.Timestamp("1899-12-30") + pd.to_timedelta(value, unit="D")).strftime("%Y-%m")
    return str(value).strip().casefold()


def excel_column(index: int) -> str:
    name = ""
    while index >= 0:
        index, rem = divmod(index, 26)
        name = chr(65 + rem) + name
        index -= 1
    return name


def clean_table(raw: pd.DataFrame, filename: str, sheet: str, section: str,
                header_row: int, start_col: int, end_col: int) -> pd.DataFrame:
    part = raw.iloc[header_row:, start_col:end_col].copy()
    if part.empty:
        return pd.DataFrame()
    headers = []
    for offset, value in enumerate(part.iloc[0].tolist()):
        label = str(value).strip() if pd.notna(value) else ""
        headers.append(label or f"Column {excel_column(start_col + offset)}")
    part = part.iloc[1:].copy()
    part.columns = headers
    part = part.dropna(how="all")
    part.insert(0, "Source row", part.index + 1)
    part.insert(0, "Source section", section)
    part.insert(0, "Source sheet", sheet)
    part.insert(0, "Source file", filename)
    return part.reset_index(drop=True)


def inspect_workbook(source: Any, filename: str) -> tuple[dict[str, pd.DataFrame], dict[str, Any], list[str]]:
    workbook = pd.ExcelFile(source)
    raw_sheets = {sheet: pd.read_excel(workbook, sheet_name=sheet, header=None, dtype=object)
                  for sheet in workbook.sheet_names}
    cover: dict[str, Any] = {}
    start = raw_sheets.get("Start_Here")
    if start is not None:
        for row in range(min(21, len(start))):
            if start.shape[1] >= 2 and pd.notna(start.iat[row, 0]):
                cover[str(start.iat[row, 0]).strip()] = start.iat[row, 1]
    if isinstance(cover.get("Reporting Month"), (int, float)) and 20000 <= cover["Reporting Month"] <= 80000:
        cover["Reporting Month"] = pd.Timestamp("1899-12-30") + pd.to_timedelta(cover["Reporting Month"], unit="D")

    tables: dict[str, pd.DataFrame] = {}
    for sheet, (header_row, blocks) in TEMPLATE.items():
        raw = raw_sheets.get(sheet)
        if raw is None:
            continue
        for section, start_col, end_col in blocks:
            table = clean_table(raw, filename, sheet, section, header_row, start_col, end_col)
            tables[f"{sheet} — {section}"] = table

    issues = []
    missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in workbook.sheet_names]
    if missing:
        issues.append("Missing expected tabs: " + ", ".join(missing))
    required_cover = ["Reporting Month", "Area", "Completed By", "Role"]
    absent = [key for key in required_cover if not normalize(cover.get(key))]
    if absent:
        issues.append("Missing cover values: " + ", ".join(absent))
    for section_name, table in tables.items():
        if table.empty:
            continue
        fields = [column for column in table.columns if column not in SOURCE_COLUMNS | {"Report Month", "Area"}]
        populated = table[fields].notna().any(axis=1) if fields else pd.Series(False, index=table.index)
        if "Area" in table.columns:
            vals = table.loc[populated, "Area"].map(normalize)
            if (vals == "").any():
                issues.append(f"{section_name}: {int((vals == '').sum())} record(s) missing Area")
            wrong = (vals != "") & (vals != normalize(cover.get("Area")))
            if wrong.any():
                issues.append(f"{section_name}: {int(wrong.sum())} record(s) Area differs from cover")
        if "Report Month" in table.columns:
            vals = table.loc[populated, "Report Month"].map(normalize)
            wrong = (vals != "") & (vals != normalize(cover.get("Reporting Month")))
            if wrong.any():
                issues.append(f"{section_name}: {int(wrong.sum())} record(s) month differs from cover")
    for sheet, raw in raw_sheets.items():
        if raw.astype(str).eq("#REF!").any().any():
            issues.append(f"{sheet}: contains a #REF! formula error")
    for table in tables.values():
        table["_Cover Area"] = cover.get("Area", "")
        table["_Cover Month"] = cover.get("Reporting Month", "")
    return tables, cover, issues


def combine_tables(loaded: list[dict[str, Any]]) -> dict[str, pd.DataFrame]:
    by_section: dict[str, list[pd.DataFrame]] = {}
    for item in loaded:
        for key, frame in item["tables"].items():
            if not frame.empty:
                by_section.setdefault(key, []).append(frame)
    return {key: pd.concat(frames, ignore_index=True, sort=False) for key, frames in by_section.items()}


def section_tables(tables: dict[str, pd.DataFrame], prefix: str) -> pd.DataFrame:
    frames = [frame for name, frame in tables.items() if name.startswith(prefix)]
    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()


def for_area(frame: pd.DataFrame, area: str) -> pd.DataFrame:
    if frame.empty or area == "All areas":
        return frame.copy()
    values = frame.get("Area", pd.Series("", index=frame.index)).fillna("").map(normalize)
    cover_values = frame.get("_Cover Area", pd.Series("", index=frame.index)).fillna("").map(normalize)
    values = values.where(values.ne(""), cover_values)
    return frame[values == normalize(area)].copy()


def for_period(frame: pd.DataFrame, period: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    values = frame.get("Report Month", pd.Series("", index=frame.index)).fillna("").map(normalize)
    cover_values = frame.get("_Cover Month", pd.Series("", index=frame.index)).fillna("").map(normalize)
    values = values.where(values.ne(""), cover_values)
    return frame[values == period].copy()


def num_sum(frame: pd.DataFrame, column: str) -> float:
    if frame.empty or column not in frame:
        return 0
    return float(pd.to_numeric(frame[column], errors="coerce").fillna(0).sum())


def nonempty_count(frame: pd.DataFrame, column: str) -> int:
    if frame.empty or column not in frame:
        return 0
    return int(frame[column].fillna("").astype(str).str.strip().ne("").sum())


def text_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "none"} else text


def checkbox_selected(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return normalize(value) in {"true", "yes", "1", "x", "checked", "✓", "✔"}


def compact_list(values: list[str], limit: int = 3, max_chars: int = 220) -> str:
    clean = [text_value(value) for value in values if text_value(value)]
    text = "; ".join(clean[:limit])
    if len(clean) > limit:
        text += f" (+{len(clean) - limit} more)"
    if len(text) > max_chars:
        text = text[: max_chars - 1].rstrip() + "…"
    return text or "Not recorded"


def metrics_for_area(tables: dict[str, pd.DataFrame], area: str) -> dict[str, float]:
    people = for_area(section_tables(tables, "Service figures — People"), area)
    refs = for_area(section_tables(tables, "Service figures — Referrals"), area)
    absence = for_area(section_tables(tables, "Absents — Absent"), area)
    agency = for_area(section_tables(tables, "Vacancy & Agency — Agency"), area)
    incidents = for_area(section_tables(tables, "Quality, Safety & Compliance — Incidents"), area)
    complaints = for_area(section_tables(tables, "Complaint & Compliment"), area)
    risks = for_area(section_tables(tables, "Dev & Risks — Risk"), area)
    development = for_area(section_tables(tables, "Dev & Risks — Development"), area)
    audits = for_area(section_tables(tables, "Audit_Train — Audits"), area)
    incident_total = num_sum(incidents, "Incident") + num_sum(incidents, "Accident")
    return {
        "People supported": num_sum(people, "Number of PS"),
        "Admissions": num_sum(people, "Admissions"),
        "Discharges": num_sum(people, "Discharge"),
        "Voids": num_sum(people, "Voids"),
        "Referrals listed": nonempty_count(refs, "Referrer"),
        "Absence hours": num_sum(absence, "Hours"),
        "Agency hours": num_sum(agency, "Agency Hours"),
        "Incidents": incident_total,
        "Complaints / compliments": nonempty_count(complaints, "Complaint/ Compliment"),
        "Open risks": int(risks.get("Status", pd.Series(dtype=object)).map(normalize).eq("open").sum()) if not risks.empty else 0,
        "Development items": nonempty_count(development, "Development item"),
        "Audit records": nonempty_count(audits, "Audit Type"),
    }


def build_area_summary(tables: dict[str, pd.DataFrame], areas: list[str]) -> pd.DataFrame:
    rows = []
    for area in areas:
        row = {"Area": area}
        row.update(metrics_for_area(tables, area))
        rows.append(row)
    return pd.DataFrame(rows)


def format_number(value: float) -> str:
    return f"{int(value):,}" if float(value).is_integer() else f"{value:,.1f}"


def report_facts(tables: dict[str, pd.DataFrame], area: str) -> dict[str, Any]:
    metrics = metrics_for_area(tables, area)
    referrals = for_area(section_tables(tables, "Service figures — Referrals"), area)
    risk = for_area(section_tables(tables, "Dev & Risks — Risk"), area)
    dev = for_area(section_tables(tables, "Dev & Risks — Development"), area)
    complaint = for_area(section_tables(tables, "Complaint & Compliment"), area)
    change = for_area(section_tables(tables, "Pack_change"), area)
    aob = for_area(section_tables(tables, "AOB"), area)
    audit = for_area(section_tables(tables, "Audit_Train — Audits"), area)
    monitoring = for_area(section_tables(tables, "Audit_Train — Systems & Governance Checks"), area)
    vacancies = for_area(section_tables(tables, "Vacancy & Agency — Vacancy"), area)
    details = {}
    supporting_records = []
    for key, frame, column in [
        ("risks", risk, "Risk Description"), ("development", dev, "Development item"),
        ("complaints", complaint, "Brief Detail"), ("changes", change, "Change Type"),
        ("audits", audit, "Audit Type"),
    ]:
        if not frame.empty and column in frame:
            details[key] = [str(value).strip() for value in frame[column].dropna()
                            if str(value).strip() and str(value).strip().casefold() not in {"nan", "none"}]
        else:
            details[key] = []
    if not aob.empty:
        details["achievements"] = [str(v).strip() for v in aob.get("Key Achievement", pd.Series(dtype=object)).dropna()
                                    if str(v).strip() and str(v).strip().casefold() != "nan"]
        details["issues"] = [str(v).strip() for v in aob.get("Describe", pd.Series(dtype=object)).dropna()
                             if str(v).strip() and str(v).strip().casefold() != "nan"]
        details["actions"] = [str(v).strip() for v in aob.get("Action Needed", pd.Series(dtype=object)).dropna()
                              if str(v).strip() and str(v).strip().casefold() != "nan"]
    else:
        details["achievements"] = []
        details["issues"] = []
        details["actions"] = []
    details["referral_updates"] = []
    if not referrals.empty:
        for _, row in referrals.iterrows():
            service = text_value(row.get("Service"))
            stage = text_value(row.get("Stage"))
            update = text_value(row.get("Update/ Next action"))
            parts = [value for value in [service, stage, update] if value]
            if parts:
                details["referral_updates"].append(" · ".join(parts))
    details["monitoring"] = []
    if not monitoring.empty and "Check" in monitoring:
        for _, row in monitoring.iterrows():
            check = text_value(row.get("Check"))
            if not check:
                continue
            if checkbox_selected(row.get("Yes")):
                status = "Yes"
            elif checkbox_selected(row.get("No")):
                status = "No"
            else:
                status = "Not recorded"
            details["monitoring"].append({"check": check, "status": status})
    details["change_records"] = []
    if not change.empty:
        for _, row in change.iterrows():
            change_type = text_value(row.get("Change Type"))
            if not change_type:
                continue
            details["change_records"].append({
                "change": change_type,
                "date": text_value(row.get("Effective Date")),
                "impact": text_value(row.get("Impact")),
                "support": text_value(row.get("Mitigation Steps/ Support Needed")),
            })
    details["complaint_counts"] = {
        "Complaints": int(complaint.get("Complaint/ Compliment", pd.Series(dtype=object)).map(normalize).str.startswith("complaint").sum()) if not complaint.empty else 0,
        "Compliments": int(complaint.get("Complaint/ Compliment", pd.Series(dtype=object)).map(normalize).str.startswith("compliment").sum()) if not complaint.empty else 0,
    }
    details["complaint_details"] = [text_value(value) for value in complaint.get("Brief Detail", pd.Series(dtype=object)).dropna() if text_value(value)] if not complaint.empty else []
    details["vacancy_records"] = nonempty_count(vacancies, "Vacancy")
    for category, frame, column in [
        ("Risk", risk, "Risk Description"), ("Development", dev, "Development item"),
        ("Complaint / compliment", complaint, "Brief Detail"), ("Pack change", change, "Change Type"),
    ]:
        if frame.empty or column not in frame:
            continue
        for _, row in frame.iterrows():
            value = row.get(column)
            if pd.isna(value) or not str(value).strip():
                continue
            source = f"{row.get('Source file', '')} · {row.get('Source sheet', '')}, row {row.get('Source row', '')}"
            row_area = row.get("Area")
            if not normalize(row_area):
                row_area = row.get("_Cover Area", "")
            service = row.get("Service", "")
            if pd.isna(service):
                service = ""
            supporting_records.append({
                "Area": row_area,
                "Section": category,
                "Service": service,
                "Summary": str(value).strip(),
                "Source": source,
            })
    return {"metrics": metrics, "details": details, "supporting_records": supporting_records}


def make_report_text(area: str, month: str, facts: dict[str, Any], history_available: bool,
                     hr_summary: dict[str, float] | None, focus: str) -> dict[str, str]:
    m, d = facts["metrics"], facts["details"]
    executive = (
        f"For {area} in {month}, the submitted reports record {format_number(m['People supported'])} people supported, "
        f"{format_number(m['Admissions'])} admissions, {format_number(m['Discharges'])} discharges, and "
        f"{format_number(m['Referrals listed'])} referral records. The submissions list "
        f"{format_number(m['Incidents'])} incidents/accidents and {format_number(m['Open risks'])} risks marked Open."
    )
    if hr_summary:
        executive += f" The mapped HR file adds {format_number(hr_summary.get('Headcount', 0))} people in scope."
    sections = {
        "Executive summary": executive,
        "Occupancy and referrals": (
            f"The source worksheets report {format_number(m['People supported'])} people supported, "
            f"{format_number(m['Admissions'])} admissions, {format_number(m['Discharges'])} discharges, and "
            f"{format_number(m['Voids'])} voids. {format_number(m['Referrals listed'])} referral records are listed. "
            "An occupancy rate is not calculated because the template does not define a capacity denominator."
        ),
        "Quarter comparison": (
            "Historical monthly submissions are available, but the quarter comparison method still needs to be agreed before a trend is calculated."
            if history_available else
            "A quarter comparison is not available from the loaded submissions; additional months are needed."
        ),
        "Improvements, development, risks and issues": (
            f"The source records contain {len(d['development'])} development items, {len(d['risks'])} risk descriptions, "
            f"and {len(d['changes'])} pack changes. Listed risks: " + ("; ".join(d["risks"][:5]) or "None recorded.")
        ),
        "Compliance and monitoring": (
            f"{format_number(m['Incidents'])} incidents/accidents are reported and {len(d['audits'])} audit types are listed. "
            "No compliance percentage is calculated without an agreed denominator and completion rule."
        ),
        "Complaints and compliments": (
            f"{format_number(m['Complaints / compliments'])} complaint/compliment records are listed. "
            "The source details are available in the selected Area's submission."
        ),
        "Staffing and absence": (
            f"The Area Manager template records {format_number(m['Absence hours'])} absence hours across all recorded absence types and "
            f"{format_number(m['Agency hours'])} agency hours. " +
            (f"The mapped HR file records {format_number(hr_summary.get('Absence hours', 0))} absence hours and "
             f"{format_number(hr_summary.get('Headcount', 0))} people in scope." if hr_summary else
             "No HR file is mapped yet; this section uses only the Area Manager template.")
        ),
    }
    if focus.strip():
        sections["Management focus"] = focus.strip()
    return sections


def monitoring_status(facts: dict[str, Any], phrase: str) -> str:
    for item in facts["details"].get("monitoring", []):
        if phrase in normalize(item.get("check")):
            return item.get("status", "Not recorded")
    return "Not recorded"


def add_report_table(doc: Document, heading: str, headers: list[str], rows: list[list[str]]) -> None:
    doc.add_heading(heading, level=1)
    if not rows:
        doc.add_paragraph("No records provided.")
        return
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Light Shading Accent 1"
    for cell, label in zip(table.rows[0].cells, headers):
        cell.text = label
    for values in rows:
        cells = table.add_row().cells
        for cell, value in zip(cells, values):
            cell.text = text_value(value) or "Not recorded"


def build_docx(title: str, month: str, metadata: dict[str, str], executive_summary: str,
               reports: list[dict[str, Any]], focus: str = "") -> bytes:
    doc = Document()
    doc.add_heading(title, 0)

    cover = doc.add_table(rows=2, cols=4)
    cover.style = "Table Grid"
    for cell, label in zip(cover.rows[0].cells, ["Month", "Area", "Completed by", "Role"]):
        cell.text = label
    for cell, key in zip(cover.rows[1].cells, ["Month", "Area", "Completed by", "Role"]):
        cell.text = metadata.get(key, "Not recorded")

    doc.add_heading("Executive summary", level=1)
    doc.add_paragraph(executive_summary)
    if focus.strip():
        doc.add_paragraph("Management focus: " + compact_list([focus], limit=1, max_chars=300))

    occupancy_rows = []
    improvement_rows = []
    compliance_rows = []
    complaint_rows = []
    change_rows = []
    staffing_rows = []
    absence_rows = []
    for report in reports:
        area_name = report["area"]
        facts = report["facts"]
        metrics = facts["metrics"]
        details = facts["details"]
        occupancy_rows.append([
            area_name,
            format_number(metrics.get("People supported", 0)),
            format_number(metrics.get("Admissions", 0)),
            format_number(metrics.get("Discharges", 0)),
            format_number(metrics.get("Voids", 0)),
            format_number(metrics.get("Referrals listed", 0)),
            compact_list(details.get("referral_updates", []), limit=2, max_chars=150),
        ])
        improvement_rows.append([
            area_name,
            compact_list(details.get("achievements", []) + details.get("development", []), limit=2, max_chars=150),
            compact_list(details.get("risks", []), limit=2, max_chars=150),
            compact_list(details.get("issues", []) + details.get("actions", []), limit=2, max_chars=150),
            "Not captured in current template",
        ])
        other_checks = [
            f"Right to work: {item['status']}" for item in details.get("monitoring", [])
            if "right to work" in normalize(item.get("check")) and item.get("status") != "Not recorded"
        ] + [
            f"Nourish devices: {item['status']}" for item in details.get("monitoring", [])
            if "nourish" in normalize(item.get("check")) and item.get("status") != "Not recorded"
        ]
        compliance_rows.append([
            area_name,
            monitoring_status(facts, "monitoring up to date"),
            monitoring_status(facts, "incidents/accidents recorded"),
            monitoring_status(facts, "risk register up to date"),
            monitoring_status(facts, "agency authorisation form submitted"),
            monitoring_status(facts, "cqc capacity tracker updated"),
            compact_list(other_checks, limit=2, max_chars=85),
        ])
        complaint_counts = details.get("complaint_counts", {})
        complaint_rows.append([
            area_name,
            format_number(complaint_counts.get("Compliments", 0)),
            format_number(complaint_counts.get("Complaints", 0)),
            compact_list(details.get("complaint_details", []), limit=2, max_chars=150),
        ])
        for change in details.get("change_records", [])[:2]:
            impact_support = " / ".join(value for value in [change.get("impact", ""), change.get("support", "")] if value)
            change_rows.append([
                area_name,
                change.get("change", ""),
                change.get("date", ""),
                impact_support,
            ])
        metrics_extra = report.get("additional_metrics", report.get("metrics", {}))
        staffing_rows.append([
            area_name,
            metrics_extra.get("HR headcount", "Not provided"),
            format_number(details.get("vacancy_records", 0)),
            format_number(metrics.get("Agency hours", 0)),
        ])
        hr_absence = metrics_extra.get("HR absence hours", "Not mapped")
        absence_rows.append([
            area_name,
            format_number(metrics.get("Absence hours", 0)),
            format_number(hr_absence) if isinstance(hr_absence, (int, float)) else str(hr_absence),
        ])

    add_report_table(doc, "Occupancy and referrals",
                     ["Area", "Supported", "Admitted", "Discharged", "Voids", "Referrals", "Referral detail"],
                     occupancy_rows)
    doc.add_paragraph("Service comments are not collected in the current Area Manager template.")

    doc.add_heading("Quarter comparison", level=1)
    doc.add_paragraph("Not calculated. Prior-quarter data and the comparison method need confirmation.")

    add_report_table(doc, "Service improvement and development, risks and issues",
                     ["Area", "Key achievement / improvement", "Risks", "Issues / actions", "Issues last report"],
                     improvement_rows)
    add_report_table(doc, "Compliance and monitoring",
                     ["Area", "Monitoring current", "Incidents recorded", "Risk register current", "Agency form", "CQC tracker", "Other checks"],
                     compliance_rows)
    add_report_table(doc, "Complaints and compliments",
                     ["Area", "Compliments", "Complaints", "Details"], complaint_rows)
    add_report_table(doc, "Package change information",
                     ["Area", "Change", "Effective date", "Impact / support"], change_rows)
    add_report_table(doc, "Staffing",
                     ["Area", "HR headcount", "Vacancy records", "Agency hours"], staffing_rows)
    add_report_table(doc, "Absence",
                     ["Area", "Area Manager hours (all types)", "HR absence hours"], absence_rows)
    doc.add_paragraph("Area Manager and HR absence totals are shown separately to avoid accidental double counting.")

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


st.set_page_config(page_title="SSP Governance | Report Builder", page_icon="📘", layout="wide")
st.markdown("""
<style>
:root { --blue:#1769aa; --navy:#16324f; --line:#dce5ef; }
.stApp { background:#f5f8fc; color:#172b4d; }
[data-testid="stHeader"] { background:rgba(255,255,255,.92); }
[data-testid="stSidebar"] { background:#ffffff; border-right:1px solid #e4ebf3; }
[data-testid="stSidebar"] * { color:#203b59; }
.block-container { max-width:1280px; padding-top:2rem; padding-bottom:3rem; }
h1,h2,h3 { color:#173b5e; letter-spacing:-.02em; }
[data-testid="stMetric"] { background:#ffffff; border:1px solid #e0e8f1; border-radius:12px; padding:16px 18px; box-shadow:0 2px 8px rgba(28,62,96,.04); }
[data-testid="stMetricLabel"] { color:#59718c; }
[data-testid="stMetricValue"] { color:#163f67; }
div.stButton > button[kind="primary"] { background:#1769aa; border:1px solid #1769aa; border-radius:8px; color:white; font-weight:600; }
div.stButton > button { border-radius:8px; }
.report-card { background:white; border:1px solid #e0e8f1; border-radius:12px; padding:22px 24px; margin:10px 0; box-shadow:0 2px 8px rgba(28,62,96,.035); }
.eyebrow { color:#52708d; font-size:.78rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; }
.app-subtitle { color:#5d7187; margin-top:-.6rem; }
.stTabs [data-baseweb="tab-list"] { gap:8px; }
.stTabs [data-baseweb="tab"] { background:white; border-radius:8px 8px 0 0; padding:9px 16px; }
</style>
""", unsafe_allow_html=True)

if "dashboard_ready" not in st.session_state:
    st.session_state.dashboard_ready = False


def reset_dashboard() -> None:
    st.session_state.dashboard_ready = False

with st.sidebar:
    st.markdown("### SSP Governance")
    page = st.radio("Workspace", ["Board report", "Data sources", "About this draft"], label_visibility="collapsed")
    st.divider()
    st.caption("Prototype · Local session")

st.markdown('<div class="eyebrow">Governance reporting</div>', unsafe_allow_html=True)
st.title("Area report builder")
st.markdown('<div class="app-subtitle">Turn monthly submissions into a clear, reviewable report for your team.</div>', unsafe_allow_html=True)

with st.container(border=True):
    st.subheader("Choose your Area Governance folder")
    st.write("Select the folder that contains the monthly Area Manager Excel reports. The files are uploaded to this local app session so it can build the dashboard.")
    folder_files = st.file_uploader(
        "Select a folder",
        type=["xlsx", "xlsm"],
        accept_multiple_files="directory",
        key="area_folder_upload",
        on_change=reset_dashboard,
        help="Choose the Area Governance Report folder. Excel files in its subfolders are included too.",
    )
    create_dashboard = st.button(
        "Create dashboard",
        type="primary",
        disabled=not folder_files,
    )
    if create_dashboard:
        st.session_state.dashboard_ready = True
    if folder_files:
        eligible_files = [file for file in folder_files
                          if Path(file.name).suffix.lower() in {".xlsx", ".xlsm"}
                          and not Path(file.name).name.startswith("~$")]
        st.caption(f"{len(eligible_files)} Excel report file(s) selected.")
    else:
        eligible_files = []

sources: list[tuple[str, Any]] = [(file.name, file) for file in eligible_files]
if not st.session_state.dashboard_ready:
    st.info("Choose the reports folder, then select **Create dashboard**.")
    st.stop()

loaded: list[dict[str, Any]] = []
for filename, source in sources:
    try:
        tables, cover, issues = inspect_workbook(source, filename)
        loaded.append({"filename": filename, "cover": cover, "tables": tables, "issues": issues})
    except Exception as exc:
        loaded.append({"filename": filename, "cover": {}, "tables": {}, "issues": [f"Could not read workbook: {exc}"]})

hr_upload = None
hr_frame = pd.DataFrame()
hr_summary_by_area: dict[str, dict[str, float]] = {}
hr_columns: list[str] = []
hr_area_col = hr_id_col = hr_absence_col = hr_month_col = None
if page == "Data sources":
    st.subheader("Optional HR data")
    st.write("Add an HR workbook or CSV to bring agreed staffing measures into the report. HR data is aggregated by Area; individual staff records are not shown in the report preview.")
    hr_upload = st.file_uploader("HR workbook or CSV", type=["xlsx", "xlsm", "csv"], key="hr_upload")
else:
    hr_upload = st.session_state.get("hr_upload")
if hr_upload is not None:
    try:
        hr_upload.seek(0)
        hr_frame = pd.read_csv(hr_upload) if hr_upload.name.lower().endswith(".csv") else pd.read_excel(hr_upload, dtype=object)
        hr_columns = [str(c) for c in hr_frame.columns]
        if page == "Data sources":
            st.success(f"Read {len(hr_frame):,} HR rows and {len(hr_columns)} columns. Choose only the fields to aggregate.")
            st.caption("The raw HR rows are not displayed or added to the report. Use approved, minimum-necessary fields.")
            st.selectbox("HR Area column", ["— choose —", *hr_columns], key="hr_area_col")
            st.selectbox("Employee identifier (for distinct headcount)", ["— none —", *hr_columns], key="hr_id_col")
            st.selectbox("Absence hours column (optional)", ["— none —", *hr_columns], key="hr_absence_col")
            st.selectbox("HR reporting month column (recommended)", ["— none —", *hr_columns], key="hr_month_col")
            st.checkbox("This HR file contains only the selected reporting month", key="hr_month_snapshot")
        hr_area_col = st.session_state.get("hr_area_col", "— choose —")
        hr_id_col = st.session_state.get("hr_id_col", "— none —")
        hr_absence_col = st.session_state.get("hr_absence_col", "— none —")
        hr_month_col = st.session_state.get("hr_month_col", "— none —")
        if page == "Data sources" and hr_area_col in hr_frame.columns:
            st.info("Field selections are kept for this local session. The report uses only Area-level aggregates.")
    except Exception as exc:
        st.error(f"Could not read the HR file: {exc}")

if not loaded:
    st.info("Choose the Area Governance folder above and create the dashboard. You can add HR data on the Data sources page after the reports are loaded.")
    st.stop()

issues = [(item["filename"], note) for item in loaded for note in item["issues"]]
covers = [item["cover"] for item in loaded if item["cover"]]
tables = combine_tables(loaded)
areas = sorted({str(cover.get("Area")).strip() for cover in covers if normalize(cover.get("Area"))})
periods = sorted({normalize(cover.get("Reporting Month")) for cover in covers if normalize(cover.get("Reporting Month"))})


def period_label(period_key: str) -> str:
    try:
        return pd.Timestamp(f"{period_key}-01").strftime("%B %Y")
    except Exception:
        return period_key

if page == "About this draft":
    st.subheader("How this draft works")
    st.markdown("""
    - It reads each Area Manager workbook and groups submissions by Area and reporting month.
    - It calculates simple totals directly from named template fields and lists supporting themes from risks, development, change, and complaints.
    - It does not infer occupancy rates, compliance percentages, quarter trends, or HR measures until their definitions and source columns are confirmed.
    - It generates a Word report draft. It does not create a consolidated source workbook.
    - The current prototype is local. Team-wide use will need organisation sign-in and an approved shared deployment.
    """)
    if issues:
        st.subheader("Items needing review")
        for filename, note in issues:
            st.warning(f"{filename}: {note}")
    st.stop()

if page == "Data sources":
    st.subheader("Area Manager submissions")
    source_manifest = pd.DataFrame([{
        "Workbook": item["filename"], "Area": item["cover"].get("Area", ""),
        "Reporting month": item["cover"].get("Reporting Month", ""),
        "Completed by": item["cover"].get("Completed By", ""),
        "Validation items": len(item["issues"]),
    } for item in loaded])
    st.dataframe(source_manifest, use_container_width=True, hide_index=True)
    st.subheader("Validation items")
    if issues:
        for filename, note in issues:
            st.warning(f"{filename}: {note}")
    else:
        st.success("No template or cover issues found in the loaded files.")
    st.stop()

if not areas:
    st.error("No Area values were found in the loaded workbook cover sheets.")
    st.stop()

filter_col1, filter_col2, filter_col3 = st.columns([1.1, 1, 2])
area = filter_col1.selectbox("Report view", ["All areas", *areas])
period_key = filter_col2.selectbox("Reporting month", periods or [""], format_func=period_label)
month = period_label(period_key)
report_focus = filter_col3.text_input("Report focus (optional)", placeholder="e.g. highlight changes, key pressures, and actions")
st.caption("Report focus is copied into the draft as context; it does not change calculated measures.")

report_area = area
period_tables = {name: for_period(frame, period_key) for name, frame in tables.items()}
facts = report_facts(period_tables, report_area)

if not hr_frame.empty and hr_area_col in hr_frame.columns:
    hr_period_frame = hr_frame.copy()
    if hr_month_col in hr_period_frame.columns:
        hr_period_frame = hr_period_frame[hr_period_frame[hr_month_col].map(normalize) == period_key]
    elif not st.session_state.get("hr_month_snapshot", False):
        if page == "Board report":
            st.warning("HR file is not included because no reporting month is mapped. On Data sources, map a month column or confirm it is a single-month snapshot.")
        hr_period_frame = hr_period_frame.iloc[0:0]
    for area_name, subset in hr_period_frame.groupby(hr_area_col, dropna=True):
        summary = {"Headcount": float(subset[hr_id_col].dropna().astype(str).nunique()) if hr_id_col in hr_period_frame.columns else 0}
        summary["Absence hours"] = float(pd.to_numeric(subset[hr_absence_col], errors="coerce").fillna(0).sum()) if hr_absence_col in hr_period_frame.columns else 0
        hr_summary_by_area[normalize(area_name)] = summary
hr_summary = None
if hr_summary_by_area:
    if area == "All areas":
        hr_summary = {key: sum(value.get(key, 0) for value in hr_summary_by_area.values()) for key in ["Headcount", "Absence hours"]}
    else:
        hr_summary = hr_summary_by_area.get(normalize(area), {"Headcount": 0, "Absence hours": 0})

selected_date = pd.Timestamp(f"{period_key}-01") if period_key else pd.Timestamp.today().normalize()
quarter_start_month = ((selected_date.month - 1) // 3) * 3 + 1
quarter_months = {f"{selected_date.year}-{month_num:02d}" for month_num in range(quarter_start_month, quarter_start_month + 3)}
historical_months = {normalize(cover.get("Reporting Month")) for cover in covers}
history_available = quarter_months.issubset(historical_months)
sections = make_report_text(area, month, facts, history_available, hr_summary, report_focus)
area_reports: list[dict[str, Any]] = []
if area == "All areas":
    for area_name in areas:
        area_facts = report_facts(period_tables, area_name)
        area_hr = hr_summary_by_area.get(normalize(area_name))
        area_metrics = dict(area_facts["metrics"])
        if area_hr:
            area_metrics["HR headcount"] = area_hr.get("Headcount", 0)
            if hr_absence_col in hr_frame.columns:
                area_metrics["HR absence hours"] = area_hr.get("Absence hours", 0)
        area_reports.append({
            "area": area_name,
            "facts": area_facts,
            "metrics": area_metrics,
            "sections": make_report_text(area_name, month, area_facts, history_available, area_hr, report_focus),
        })

st.markdown(f"<div class='eyebrow'>{month} · {area}</div>", unsafe_allow_html=True)
st.header("Executive overview")
metrics = facts["metrics"]
kpi_names = ["People supported", "Admissions", "Referrals listed", "Absence hours", "Agency hours"]
kpi_cols = st.columns(len(kpi_names))
for col, name in zip(kpi_cols, kpi_names):
    col.metric(name, format_number(metrics[name]))

if hr_summary:
    hr1, hr2 = st.columns(2)
    hr1.metric("HR headcount", format_number(hr_summary.get("Headcount", 0)))
    hr2.metric("HR absence hours", format_number(hr_summary.get("Absence hours", 0)))

if area == "All areas":
    summary = build_area_summary(period_tables, areas)
    with st.expander("Compare Areas", expanded=True):
        st.dataframe(summary, use_container_width=True, hide_index=True)
        if not summary.empty and "People supported" in summary:
            st.bar_chart(summary.set_index("Area")[["People supported"]], color="#1769aa")
    st.subheader("Area-by-area summaries")
    area_tabs = st.tabs([report["area"] for report in area_reports])
    for area_tab, report in zip(area_tabs, area_reports):
        with area_tab:
            st.markdown(f"#### {report['area']} · {month}")
            measure_names = ["People supported", "Admissions", "Referrals listed", "Absence hours", "Agency hours"]
            area_kpis = st.columns(len(measure_names))
            for col, name in zip(area_kpis, measure_names):
                col.metric(name, format_number(report["metrics"].get(name, 0)))
            if report["metrics"].get("HR headcount") is not None:
                hr_col1, hr_col2 = st.columns(2)
                hr_col1.metric("HR headcount", format_number(report["metrics"].get("HR headcount", 0)))
                hr_col2.metric("HR absence hours", format_number(report["metrics"].get("HR absence hours", 0)))
            st.markdown(f"**Executive summary**  \n{report['sections']['Executive summary']}")
            for section_name in BOARD_SECTIONS[1:]:
                with st.expander(section_name):
                    st.write(report["sections"][section_name])

st.subheader("Board report preview")
st.markdown(f"<div class='report-card'><div class='eyebrow'>Executive summary</div><p>{sections['Executive summary']}</p></div>", unsafe_allow_html=True)
tabs = st.tabs(BOARD_SECTIONS)
for tab, section in zip(tabs, BOARD_SECTIONS):
    with tab:
        st.markdown(f"### {section}")
        st.write(sections[section])
        if section == "Improvements, development, risks and issues":
            if facts["details"]["development"]:
                st.markdown("**Development recorded**")
                for item in facts["details"]["development"][:8]:
                    st.markdown(f"- {item}")
            if facts["details"]["achievements"]:
                st.markdown("**Achievements recorded**")
                for item in facts["details"]["achievements"][:8]:
                    st.markdown(f"- {item}")
            if facts["details"]["risks"]:
                st.markdown("**Risk descriptions**")
                for item in facts["details"]["risks"][:8]:
                    st.markdown(f"- {item}")
        if section == "Complaints and compliments" and facts["details"]["complaints"]:
            for item in facts["details"]["complaints"][:8]:
                st.markdown(f"- {item}")
if report_focus.strip():
    st.subheader("Management focus")
    st.info(report_focus.strip())

selected_covers = [cover for cover in covers
                   if normalize(cover.get("Reporting Month")) == period_key
                   and (area == "All areas" or normalize(cover.get("Area")) == normalize(area))]


def joined_cover_values(key: str) -> str:
    values = []
    seen = set()
    for cover in selected_covers:
        value = text_value(cover.get(key))
        if value and normalize(value) not in seen:
            seen.add(normalize(value))
            values.append(value)
    return "; ".join(values) or "Not recorded"


metadata = {
    "Month": month,
    "Area": area,
    "Completed by": joined_cover_values("Completed By"),
    "Role": joined_cover_values("Role"),
}
single_area_metrics = {"HR headcount": hr_summary.get("Headcount", "Not provided")} if hr_summary else {}
if hr_summary and hr_absence_col in hr_frame.columns:
    single_area_metrics["HR absence hours"] = hr_summary.get("Absence hours", 0)
docx_reports = area_reports if area == "All areas" else [{
    "area": area,
    "facts": facts,
    "additional_metrics": single_area_metrics,
}]
docx = build_docx("SSP Governance Board Report", month, metadata,
                  sections["Executive summary"], docx_reports, report_focus)
st.download_button("Download Board report (Word)", docx,
                   file_name=f"ssp_governance_report_{month.replace(' ', '_')}_{area.replace(' ', '_')}.docx",
                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                   type="primary")
if issues:
    st.caption(f"{len(issues)} data-quality item(s) need review in Data sources before circulation.")
else:
    st.caption("Figures are derived from the submitted template fields. Review the draft before circulation.")
if facts["supporting_records"]:
    with st.expander("View supporting Area and service details"):
        st.dataframe(pd.DataFrame(facts["supporting_records"]), use_container_width=True, hide_index=True)
