"""Adapter: submission JSON (from the web form) -> filled .xlsx bytes.

Reuses the validated Phase-1 engine (model + filler + rows). Fills entirely
in-memory via BytesIO -- no temp files, no LibreOffice in production (Excel
recalculates the formulas when the user opens the downloaded file).
"""
from __future__ import annotations

import io
import os
import re
from datetime import date, datetime

from .model import (
    JobSubmission, OfficeSubmission, PerDiemType,
    MileageRow, CommuteOffset, ExpenseLine, MiscExpense,
    OfficeExpense, CreditCardCharge,
)
from .filler import fill_job_report, fill_office_report

_TEMPLATES = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "_templates")
JOB_TPL = os.path.join(_TEMPLATES, "2026_DomTravelER.xlsx")
OFF_TPL = os.path.join(_TEMPLATES, "2026_OER.xlsx")


def _d(s):
    if not s:
        return None
    if isinstance(s, (date, datetime)):
        return s
    return date.fromisoformat(str(s)[:10])


def _f(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _i(v):
    try:
        return int(v) if v not in (None, "") else 0
    except (TypeError, ValueError):
        return 0


def _safe(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", (name or "report").strip()) or "report"


def _per_diem(v) -> PerDiemType:
    if not v:
        return PerDiemType.NONE
    try:
        return PerDiemType(v)            # match by exact dropdown text
    except ValueError:
        try:
            return PerDiemType[v]        # or by enum member name
        except KeyError:
            return PerDiemType.NONE


def build_job(data: dict) -> JobSubmission:
    return JobSubmission(
        submitter_name=data.get("submitter_name", ""),
        submitter_email=data.get("submitter_email", ""),
        job_number=data.get("job_number", ""),
        customer_name_state=data.get("customer_name_state", ""),
        customer_billing_address=data.get("customer_billing_address", ""),
        customer_po=data.get("customer_po", ""),
        customer_contact=data.get("customer_contact", ""),
        overnight_travel=data.get("overnight_travel", ""),
        days_traveled=_i(data.get("days_traveled")) or None,
        project_description=data.get("project_description", ""),
        mileage=[MileageRow(_d(m.get("date")), m.get("origin", ""), m.get("destination", ""), _f(m.get("miles")))
                 for m in data.get("mileage", []) if any(m.values())],
        commute_offset=CommuteOffset(_f((data.get("commute") or {}).get("rt_commute_miles")),
                                     _i((data.get("commute") or {}).get("num_days"))),
        expenses=[ExpenseLine(
            _d(e.get("date")), _per_diem(e.get("per_diem")),
            airfare=_f(e.get("airfare")), ground_transport=_f(e.get("ground_transport")),
            car_rental_gas=_f(e.get("car_rental_gas")), lodging=_f(e.get("lodging")),
            telephone_fax=_f(e.get("telephone_fax")), parking=_f(e.get("parking")),
            misc=_f(e.get("misc")),
        ) for e in data.get("expenses", []) if any(str(v) for v in e.values() if v not in (None, "", "No Per Diem"))],
        misc_expenses=[MiscExpense(_d(x.get("date")), x.get("description", ""), _f(x.get("amount")))
                       for x in data.get("misc_expenses", []) if any(x.values())],
    )


def build_office(data: dict) -> OfficeSubmission:
    return OfficeSubmission(
        submitter_name=data.get("submitter_name", ""),
        submitter_email=data.get("submitter_email", ""),
        report_date=_d(data.get("report_date")),
        mileage=[MileageRow(_d(m.get("date")), m.get("origin", ""), m.get("destination", ""), _f(m.get("miles")))
                 for m in data.get("mileage", []) if any(m.values())],
        office_expenses=[OfficeExpense(_d(o.get("date")), o.get("description", ""), _f(o.get("cost")))
                         for o in data.get("office_expenses", []) if any(o.values())],
        credit_card_charges=[CreditCardCharge(_d(c.get("date")), c.get("vendor", ""), c.get("description", ""), _f(c.get("amount")))
                             for c in data.get("credit_card_charges", []) if any(c.values())],
    )


def generate_xlsx(data: dict) -> tuple[str, bytes]:
    """Return (filename, xlsx_bytes) for the given submission dict."""
    rtype = (data.get("report_type") or "job").lower()
    buf = io.BytesIO()
    if rtype == "office":
        sub = build_office(data)
        fill_office_report(sub, OFF_TPL, buf, datetime.now())
        fname = f"Office_Expense_{_safe(sub.submitter_name)}.xlsx"
    else:
        sub = build_job(data)
        fill_job_report(sub, JOB_TPL, buf, datetime.now())
        fname = f"Travel_Expense_{_safe(sub.submitter_name)}_{_safe(sub.job_number)}.xlsx"
    return fname, buf.getvalue()
