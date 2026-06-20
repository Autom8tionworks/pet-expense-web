"""Data model for an expense-report submission.

Pure stdlib (dataclasses). No AI, no I/O. This is the contract the rest of the
pipeline fills in. In later phases, Phase-2 receipt scanning produces ExpenseLine
objects and Phase-3 UI produces the per-diem/mileage selections; the shape here
does not need to change.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import Optional


# --------------------------------------------------------------------------- #
# Enums / constants
# --------------------------------------------------------------------------- #
class ReportType(str, Enum):
    JOB = "job"        # 2026_DomTravelER.xlsx
    OFFICE = "office"  # 2026_OER.xlsx


class PerDiemType(str, Enum):
    """The 7 selectable options, matching the job template's dropdown exactly.

    NB: the template's lookup table also carries a leading blank row (rate 0);
    that blank is not user-selectable here. 'No Per Diem' is the explicit zero.
    The string values MUST match 'Per Diem Rates'!B6:B12 character-for-character
    because the in-sheet INDEX/MATCH looks them up by text.
    """
    NONE = "No Per Diem"
    FULL = "Full Day"
    PARTIAL = "Partial Day"
    FULL_D1 = "Full Day, Deduct 1 Meal"
    FULL_D2 = "Full Day, Deduct 2 Meals"
    PARTIAL_D1 = "Partial Day, Deduct 1 Meal"
    PARTIAL_D2 = "Partial Day, Deduct 2 Meals"


# Mirror of 'Per Diem Rates'!C6:C12. Used only for client-side preview/validation;
# the authoritative number is computed by the spreadsheet's own INDEX/MATCH so the
# generated file recalculates correctly even if this constant drifts.
PER_DIEM_RATE: dict[PerDiemType, int] = {
    PerDiemType.NONE: 0,
    PerDiemType.FULL: 68,
    PerDiemType.PARTIAL: 68,
    PerDiemType.FULL_D1: 46,      # 68 - 22
    PerDiemType.FULL_D2: 24,      # 68 - 22 - 22
    PerDiemType.PARTIAL_D1: 46,   # 68 - 22
    PerDiemType.PARTIAL_D2: 24,   # 68 - 22 - 22
}

# Mileage rate is HARDCODED in the templates (job B14, office B10) at $0.68/mile.
# GOTCHA / FUTURE: this should become a configurable value -- it changes
# periodically (IRS standard rate). For Phase 1 we read it back from the template
# rather than re-asserting it, so we never silently disagree with the sheet.
DEFAULT_MILEAGE_RATE = 0.68


# --------------------------------------------------------------------------- #
# Job report — line item types
# --------------------------------------------------------------------------- #
@dataclass
class MileageRow:
    """Section 1 personal-car trip mileage (job rows 10-11 / office rows 6-9)."""
    travel_date: Optional[date]
    origin: str
    destination: str
    miles: float


@dataclass
class CommuteOffset:
    """The 'RT daily commute mileage x #days' deduction subtracted from total
    reimbursable mileage (job row 12). Captured separately from trip mileage so
    the offset can be computed correctly. Office template has no commute offset.
    """
    rt_commute_miles: float = 0.0
    num_days: int = 0


@dataclass
class ExpenseLine:
    """Section 2 (job) per-day travel-expense row, rows 17-26.

    `meals` is intentionally absent as an input: it is auto-calculated by the
    sheet from `per_diem`. Each category below maps to one column. The combined
    'Car rental / Gasoline' column accepts both rental and gas amounts.
    """
    line_date: Optional[date]
    per_diem: PerDiemType = PerDiemType.NONE
    airfare: float = 0.0
    ground_transport: float = 0.0
    car_rental_gas: float = 0.0      # combined rental + gasoline column
    lodging: float = 0.0
    telephone_fax: float = 0.0
    parking: float = 0.0
    misc: float = 0.0                # the in-table 'Misc. Expenses' column (J)


@dataclass
class MiscExpense:
    """Section 3 (job) bottom 'MISCELLANEOUS EXPENSES' block, rows 38-39.

    Distinct from the in-table Misc column above; this is the separate
    Date/Description/Amount section now totalled into 'Total Section 3'.
    """
    expense_date: Optional[date]
    description: str
    amount: float


# --------------------------------------------------------------------------- #
# Office report — line item types
# --------------------------------------------------------------------------- #
@dataclass
class OfficeExpense:
    """Office Section 2 general office expense, rows 13-23."""
    expense_date: Optional[date]
    description: str
    cost: float


@dataclass
class CreditCardCharge:
    """Office Section 3 credit-card / open-account charge, rows 27-29."""
    charge_date: Optional[date]
    vendor: str
    description: str
    amount: float


# --------------------------------------------------------------------------- #
# Submissions
# --------------------------------------------------------------------------- #
@dataclass
class JobSubmission:
    # --- identity / required ---
    submitter_name: str
    job_number: str               # required for job reports
    # --- optional ---
    submitter_email: str = ""     # optional now (no email routing in the web app)
    customer_name_state: str = ""
    customer_billing_address: str = ""
    customer_po: str = ""
    customer_contact: str = ""
    overnight_travel: str = ""
    days_traveled: Optional[int] = None
    project_description: str = ""
    # --- sections ---
    mileage: list[MileageRow] = field(default_factory=list)
    commute_offset: CommuteOffset = field(default_factory=CommuteOffset)
    expenses: list[ExpenseLine] = field(default_factory=list)        # Section 2
    misc_expenses: list[MiscExpense] = field(default_factory=list)   # Section 3
    report_type: ReportType = ReportType.JOB

    def validate(self) -> list[str]:
        errs: list[str] = []
        if not self.submitter_name.strip():
            errs.append("submitter_name is required")
        if self.submitter_email and "@" not in self.submitter_email:
            errs.append("submitter_email must be a valid email")
        if not self.job_number.strip():
            errs.append("job_number is required for a Job report")
        for i, e in enumerate(self.expenses):
            if not isinstance(e.per_diem, PerDiemType):
                errs.append(f"expense row {i}: per_diem must be a PerDiemType")
        return errs


@dataclass
class OfficeSubmission:
    submitter_name: str
    submitter_email: str = ""
    report_date: Optional[date] = None
    mileage: list[MileageRow] = field(default_factory=list)
    office_expenses: list[OfficeExpense] = field(default_factory=list)
    credit_card_charges: list[CreditCardCharge] = field(default_factory=list)
    report_type: ReportType = ReportType.OFFICE

    def validate(self) -> list[str]:
        errs: list[str] = []
        if not self.submitter_name.strip():
            errs.append("submitter_name is required")
        if self.submitter_email and "@" not in self.submitter_email:
            errs.append("submitter_email must be a valid email")
        return errs
