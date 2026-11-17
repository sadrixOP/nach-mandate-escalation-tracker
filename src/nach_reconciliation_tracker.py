"""Automated NACH Mandate Failure Reconciliation & Escalation Tracker.

Reconciles NPCI NACH clearing batch responses against master mandates,
classifies NPCI return codes into actionable resolution buckets, detects
recurring mandate failures & partner-bank spikes, computes escalation SLA
turnaround buckets, exports a formula-driven 4-sheet Excel MIS workbook,
and generates a high-resolution summary dashboard PNG.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

import matplotlib.gridspec as gridspec
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
SCREENSHOT_DIR = BASE_DIR / "screenshots"

PARTNER_BANKS = [
    ("HDFC Bank", 0.23, 0.08),
    ("ICICI Bank", 0.21, 0.09),
    ("SBI", 0.22, 0.13),
    ("Axis Bank", 0.14, 0.10),
    ("Kotak Mahindra", 0.11, 0.09),
    ("Yes Bank", 0.09, 0.17),  # Slight anomaly / higher timeout bounce rate
]

NPCI_RETURN_CODES = {
    "AM04": {
        "description": "Insufficient Funds",
        "category": "Customer Actionable (Soft Bounce)",
        "weight": 0.42,
        "recommended_action": "Trigger Automated Payment Link & SMS Reminder (T+1 Retry)",
    },
    "MD01": {
        "description": "Invalid Mandate / UMRN Not Registered",
        "category": "Mandate Re-registration Required (Hard Bounce)",
        "weight": 0.16,
        "recommended_action": "Initiate Fresh e-NACH / Physical Mandate Onboarding",
    },
    "AC01": {
        "description": "Incorrect / Closed Account Number",
        "category": "Mandate Re-registration Required (Hard Bounce)",
        "weight": 0.12,
        "recommended_action": "Collect Alternate Bank Details & Re-register UMRN",
    },
    "SL01": {
        "description": "Signature Mismatch",
        "category": "Mandate Re-registration Required (Hard Bounce)",
        "weight": 0.10,
        "recommended_action": "Route Customer to Aadhaar eSign / NetBanking e-Mandate",
    },
    "MS03": {
        "description": "Partner Bank Server Timeout / Gateway Error",
        "category": "Partner Bank Technical Issue (Retry Eligible)",
        "weight": 0.13,
        "recommended_action": "Log Bank Tech Escalation & Schedule Same-Day Batch Re-presentment",
    },
    "RR01": {
        "description": "Account Frozen / Regulatory Block",
        "category": "Customer Actionable (Soft Bounce)",
        "weight": 0.07,
        "recommended_action": "Outbound Voice Outreach for Alternate Repayment Mode",
    },
}


def ensure_directories() -> None:
    """Create project directories if missing."""
    for folder in (DATA_DIR, OUTPUT_DIR, SCREENSHOT_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def generate_nach_datasets(num_txns: int = 1500) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Generate master mandates, 1,500 NACH batch clearing txns, and escalation tickets."""
    rng = np.random.default_rng(77)
    bank_names = [b[0] for b in PARTNER_BANKS]
    bank_weights = [b[1] for b in PARTNER_BANKS]
    bank_bounce_rate = {b[0]: b[2] for b in PARTNER_BANKS}

    # Build 420 unique master mandates (UMRN) so recurring failures cluster realistically
    num_mandates = 420
    mandates: list[dict[str, object]] = []
    for m_idx in range(1, num_mandates + 1):
        bank = str(rng.choice(bank_names, p=bank_weights))
        emi_amount = int(rng.choice([2500, 3800, 5200, 7500, 11000, 15500, 22000]))
        mandates.append(
            {
                "umrn": f"UMRN{202600000 + m_idx}",
                "customer_id": f"CUST-{3000 + m_idx}",
                "customer_name": f"Borrower_{m_idx:03d}",
                "partner_bank": bank,
                "max_mandate_limit_inr": emi_amount + 2000,
                "scheduled_emi_inr": emi_amount,
                "mandate_status": "ACTIVE",
            }
        )
    master_df = pd.DataFrame(mandates)

    # Mark ~28 mandates as chronic repeat-bouncers for realistic recurring-failure detection
    chronic_umrns = set(master_df.sample(n=28, random_state=77)["umrn"].tolist())

    code_list = list(NPCI_RETURN_CODES.keys())
    code_weights = [NPCI_RETURN_CODES[c]["weight"] for c in code_list]

    start_dt = datetime(2026, 8, 15, 7, 30, 0)
    txns: list[dict[str, object]] = []
    escalations: list[dict[str, object]] = []
    esc_seq = 1

    for idx in range(1, num_txns + 1):
        # Concentrate chronic mandates slightly higher in batch cycles
        if idx % 9 == 0:
            m_row = master_df[master_df["umrn"].isin(chronic_umrns)].sample(1, random_state=idx).iloc[0]
        else:
            m_row = master_df.iloc[int(rng.integers(0, num_mandates))]

        cycle_day = int(rng.integers(0, 75))
        presentment_dt = start_dt + timedelta(days=cycle_day, minutes=int(rng.integers(0, 480)))
        bank = str(m_row["partner_bank"])
        umrn = str(m_row["umrn"])

        # Determine pass vs. bounce
        bounce_prob = bank_bounce_rate[bank]
        if umrn in chronic_umrns:
            bounce_prob = 0.78
        is_failed = rng.random() < bounce_prob

        if is_failed:
            # Yes Bank & SBI have slightly higher MS03 technical timeouts
            if bank in ("Yes Bank", "SBI") and rng.random() < 0.28:
                ret_code = "MS03"
            elif umrn in chronic_umrns and rng.random() < 0.68:
                ret_code = "AM04"
            else:
                ret_code = str(rng.choice(code_list, p=code_weights))
            clearing_status = "FAILED"
        else:
            ret_code = "0000"
            clearing_status = "CLEARED"

        txn_id = f"NACH-{20260000 + idx}"
        txns.append(
            {
                "txn_id": txn_id,
                "batch_id": f"BATCH-{presentment_dt.strftime('%Y%m%d')}",
                "presentment_date": presentment_dt.strftime("%Y-%m-%d"),
                "umrn": umrn,
                "customer_id": m_row["customer_id"],
                "partner_bank": bank,
                "debit_amount_inr": int(m_row["scheduled_emi_inr"]),
                "clearing_status": clearing_status,
                "npci_return_code": ret_code,
            }
        )

        # Create escalation tickets for hard bounces, technical timeouts, and repeat soft bounces
        if is_failed and (ret_code in ("MD01", "AC01", "SL01", "MS03") or umrn in chronic_umrns):
            created_ts = presentment_dt + timedelta(hours=float(rng.uniform(1.0, 4.0)))
            is_open = cycle_day >= 72 and rng.random() < 0.45
            if ret_code == "MS03":
                sla_target_hrs = 24.0
                turnaround_hrs = float(np.clip(rng.normal(19.0, 11.0), 3.0, 62.0))
            else:
                sla_target_hrs = 48.0
                turnaround_hrs = float(np.clip(rng.normal(28.0, 14.0), 4.0, 74.0))

            if is_open:
                esc_status = "Pending Partner Bank" if ret_code == "MS03" else "Awaiting Customer Action"
                resolved_ts_str = ""
                elapsed_hrs = round(float(rng.uniform(14.0, 56.0)), 1)
            else:
                esc_status = "Resolved"
                resolved_ts = created_ts + timedelta(hours=turnaround_hrs)
                resolved_ts_str = resolved_ts.strftime("%Y-%m-%d %H:%M")
                elapsed_hrs = round(turnaround_hrs, 1)

            escalations.append(
                {
                    "escalation_id": f"ESC-{5000 + esc_seq}",
                    "txn_id": txn_id,
                    "umrn": umrn,
                    "customer_id": m_row["customer_id"],
                    "partner_bank": bank,
                    "npci_return_code": ret_code,
                    "failure_category": NPCI_RETURN_CODES[ret_code]["category"],
                    "assigned_queue": "Bank Tech Desk" if ret_code == "MS03" else "Customer Mandate Ops",
                    "created_at": created_ts.strftime("%Y-%m-%d %H:%M"),
                    "resolved_at": resolved_ts_str,
                    "ticket_status": esc_status,
                    "sla_target_hours": sla_target_hrs,
                    "turnaround_hours": elapsed_hrs,
                }
            )
            esc_seq += 1

    txns_df = pd.DataFrame(txns)
    esc_df = pd.DataFrame(escalations)

    master_df.to_csv(DATA_DIR / "nach_master_mandates.csv", index=False, encoding="utf-8")
    txns_df.to_csv(DATA_DIR / "nach_batch_transactions.csv", index=False, encoding="utf-8")
    esc_df.to_csv(DATA_DIR / "nach_escalations_log.csv", index=False, encoding="utf-8")
    return master_df, txns_df, esc_df


def reconcile_and_analyze(
    master_df: pd.DataFrame, txns_df: pd.DataFrame, esc_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Reconcile transactions against master mandates and flag recurring failures + SLA aging."""
    merged = txns_df.merge(
        master_df[["umrn", "customer_name", "max_mandate_limit_inr", "mandate_status"]],
        on="umrn",
        how="left",
    )

    merged["return_description"] = merged["npci_return_code"].map(
        lambda c: "Cleared Successfully" if c == "0000" else NPCI_RETURN_CODES[c]["description"]
    )
    merged["failure_bucket"] = merged["npci_return_code"].map(
        lambda c: "Cleared" if c == "0000" else NPCI_RETURN_CODES[c]["category"]
    )
    merged["recommended_ops_action"] = merged["npci_return_code"].map(
        lambda c: "No Action Required" if c == "0000" else NPCI_RETURN_CODES[c]["recommended_action"]
    )
    merged["limit_check"] = np.where(
        merged["debit_amount_inr"] <= merged["max_mandate_limit_inr"],
        "Within Sanctioned Limit",
        "Exceeds Limit",
    )

    # Recurring failure watchlist (>= 3 failed presentments on same UMRN)
    failed_only = merged[merged["clearing_status"] == "FAILED"].copy()
    watchlist = (
        failed_only.groupby(["umrn", "customer_id", "customer_name", "partner_bank"], as_index=False)
        .agg(
            failed_attempts=("txn_id", "count"),
            total_failed_exposure_inr=("debit_amount_inr", "sum"),
            top_return_code=("npci_return_code", lambda s: s.mode().iloc[0]),
            last_failure_date=("presentment_date", "max"),
        )
    )
    watchlist = watchlist[watchlist["failed_attempts"] >= 3].sort_values(
        ["failed_attempts", "total_failed_exposure_inr"], ascending=False
    )
    watchlist["risk_tier"] = np.where(
        watchlist["failed_attempts"] >= 5,
        "CRITICAL - Immediate Outreach",
        "HIGH - Watchlist",
    )

    # SLA aging bucket on escalations
    esc_enriched = esc_df.copy()
    esc_enriched["sla_aging_bucket"] = np.select(
        [
            esc_enriched["turnaround_hours"] <= 24.0,
            esc_enriched["turnaround_hours"] <= 48.0,
        ],
        [
            "Within SLA (<24h)",
            "Warning (24-48h)",
        ],
        default="SLA Breached (>48h)",
    )
    esc_enriched["sla_breached"] = np.where(
        esc_enriched["turnaround_hours"] > esc_enriched["sla_target_hours"], "YES", "NO"
    )

    return merged, watchlist, esc_enriched


def _style_header(ws: openpyxl.worksheet.worksheet.Worksheet, row_idx: int, num_cols: int, fill_hex: str = "0F172A") -> None:
    """Apply consistent bold dark-slate header styling."""
    fill = PatternFill("solid", fgColor=fill_hex)
    font = Font(color="FFFFFF", bold=True, size=10.5)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autosize_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, freeze_cell: str = "A2") -> None:
    """Auto-fit column widths and apply freeze panes."""
    ws.freeze_panes = freeze_cell
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 13), 44)


def export_excel_workbook(
    reconciled_df: pd.DataFrame, watchlist_df: pd.DataFrame, esc_df: pd.DataFrame
) -> Path:
    """Build a 4-sheet Advanced Excel MIS workbook with live Excel formulas & conditional styling."""
    wb = openpyxl.Workbook()

    # ------------------------------------------------------------------
    # Sheet 1: MIS_Executive_Dashboard (live formulas referencing Sheets 2 & 4)
    # ------------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "MIS_Executive_Dashboard"
    ws_dash["A1"] = "NACH MANDATE FAILURE RECONCILIATION & ESCALATION MIS TRACKER"
    ws_dash["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws_dash["A2"] = "Formula-Driven Summary Linked to Batch_Reconciliation_Log & Partner_Bank_SLA_Tracker"
    ws_dash["A2"].font = Font(italic=True, size=10, color="475569")

    n_batch = len(reconciled_df) + 1
    n_esc = len(esc_df) + 1

    # Table A: Overall Batch KPIs (Live Excel formulas)
    ws_dash["A4"] = "Metric"
    ws_dash["B4"] = "Value"
    _style_header(ws_dash, 4, 2)
    kpi_rows = [
        ("Total Batch Presentments", f"=COUNTA(Batch_Reconciliation_Log!A2:A{n_batch})"),
        ("Cleared Transactions", f'=COUNTIF(Batch_Reconciliation_Log!H2:H{n_batch},"CLEARED")'),
        ("Failed / Bounced Transactions", f'=COUNTIF(Batch_Reconciliation_Log!H2:H{n_batch},"FAILED")'),
        ("Batch Clearance Rate (%)", "=ROUND(B6/B5*100,2)"),
        ("Total Failed Exposure (INR)", f'=SUMIF(Batch_Reconciliation_Log!H2:H{n_batch},"FAILED",Batch_Reconciliation_Log!G2:G{n_batch})'),
        ("Recurring Watchlist Mandates (>=3 Fails)", len(watchlist_df)),
        ("Total Escalation Tickets", f"=COUNTA(Partner_Bank_SLA_Tracker!A2:A{n_esc})"),
        ("Escalation SLA Compliance (%)", f'=ROUND(COUNTIF(Partner_Bank_SLA_Tracker!O2:O{n_esc},"NO")/B11*100,2)'),
    ]
    for idx, (label, formula_val) in enumerate(kpi_rows, start=5):
        ws_dash.cell(row=idx, column=1, value=label)
        ws_dash.cell(row=idx, column=2, value=formula_val).font = Font(bold=True)

    # Table B: Failure Reasons Breakdown (Live COUNTIF / SUMIF)
    ws_dash["D4"] = "NPCI Code"
    ws_dash["E4"] = "Return Description"
    ws_dash["F4"] = "Failure Category"
    ws_dash["G4"] = "Failure Count"
    ws_dash["H4"] = "Failed Amount (INR)"
    _style_header(ws_dash, 4, 8, fill_hex="1E3A8A")
    # Re-apply cols 1-2 header and col 3 blank spacer cleanly
    _style_header(ws_dash, 4, 2)
    ws_dash["C4"].fill = PatternFill(fill_type=None)

    for r_i, (code, meta) in enumerate(NPCI_RETURN_CODES.items(), start=5):
        ws_dash.cell(row=r_i, column=4, value=code).font = Font(bold=True)
        ws_dash.cell(row=r_i, column=5, value=meta["description"])
        ws_dash.cell(row=r_i, column=6, value=meta["category"])
        ws_dash.cell(row=r_i, column=7, value=f'=COUNTIF(Batch_Reconciliation_Log!I2:I{n_batch},D{r_i})')
        ws_dash.cell(row=r_i, column=8, value=f'=SUMIF(Batch_Reconciliation_Log!I2:I{n_batch},D{r_i},Batch_Reconciliation_Log!G2:G{n_batch})')

    # Table C: Partner Bank Bounce Matrix
    ws_dash["A15"] = "Partner Bank"
    ws_dash["B15"] = "Total Txns"
    ws_dash["C15"] = "Failed Txns"
    ws_dash["D15"] = "Bounce Rate (%)"
    ws_dash["E15"] = "Bank Timeout (MS03)"
    _style_header(ws_dash, 15, 5)
    for b_i, (bank_name, _, _) in enumerate(PARTNER_BANKS, start=16):
        ws_dash.cell(row=b_i, column=1, value=bank_name).font = Font(bold=True)
        ws_dash.cell(row=b_i, column=2, value=f'=COUNTIF(Batch_Reconciliation_Log!F2:F{n_batch},A{b_i})')
        ws_dash.cell(row=b_i, column=3, value=f'=COUNTIFS(Batch_Reconciliation_Log!F2:F{n_batch},A{b_i},Batch_Reconciliation_Log!H2:H{n_batch},"FAILED")')
        ws_dash.cell(row=b_i, column=4, value=f"=ROUND(C{b_i}/B{b_i}*100,2)")
        ws_dash.cell(row=b_i, column=5, value=f'=COUNTIFS(Batch_Reconciliation_Log!F2:F{n_batch},A{b_i},Batch_Reconciliation_Log!I2:I{n_batch},"MS03")')

    # ------------------------------------------------------------------
    # Sheet 2: Batch_Reconciliation_Log (with live Excel IF/COUNTIF formulas)
    # ------------------------------------------------------------------
    ws_batch = wb.create_sheet("Batch_Reconciliation_Log")
    batch_cols = [
        "txn_id", "batch_id", "presentment_date", "umrn", "customer_id",
        "partner_bank", "debit_amount_inr", "clearing_status", "npci_return_code",
        "return_description", "failure_bucket", "limit_check",
        "recommended_ops_action", "excel_repeat_fail_count", "excel_escalation_flag",
    ]
    ws_batch.append(batch_cols)
    _style_header(ws_batch, 1, len(batch_cols))

    green_fill = PatternFill("solid", fgColor="DCFCE7")
    red_fill = PatternFill("solid", fgColor="FEE2E2")
    amber_fill = PatternFill("solid", fgColor="FEF3C7")

    for r_i, row in enumerate(reconciled_df[batch_cols[:-2]].itertuples(index=False), start=2):
        repeat_formula = f'=COUNTIFS($D$2:$D${n_batch},D{r_i},$H$2:$H${n_batch},"FAILED")'
        esc_formula = f'=IF(H{r_i}="CLEARED","OK",IF(OR(I{r_i}="MD01",I{r_i}="AC01",I{r_i}="SL01",I{r_i}="MS03",N{r_i}>=3),"ESCALATE","AUTO-RETRY"))'
        ws_batch.append(list(row) + [repeat_formula, esc_formula])

        status_cell = ws_batch.cell(row=r_i, column=8)
        if row.clearing_status == "CLEARED":
            status_cell.fill = green_fill
        else:
            status_cell.fill = red_fill if row.npci_return_code in ("MD01", "AC01", "SL01", "MS03") else amber_fill
            status_cell.font = Font(bold=True)

    # ------------------------------------------------------------------
    # Sheet 3: Recurring_Failure_Watchlist
    # ------------------------------------------------------------------
    ws_watch = wb.create_sheet("Recurring_Failure_Watchlist")
    watch_cols = list(watchlist_df.columns)
    ws_watch.append(watch_cols)
    _style_header(ws_watch, 1, len(watch_cols))
    for r_i, row in enumerate(watchlist_df.itertuples(index=False), start=2):
        ws_watch.append(list(row))
        risk_cell = ws_watch.cell(row=r_i, column=len(watch_cols))
        risk_cell.fill = red_fill if "CRITICAL" in row.risk_tier else amber_fill
        risk_cell.font = Font(bold=True)

    # ------------------------------------------------------------------
    # Sheet 4: Partner_Bank_SLA_Tracker (with live Excel SLA aging formula)
    # ------------------------------------------------------------------
    ws_sla = wb.create_sheet("Partner_Bank_SLA_Tracker")
    sla_cols = list(esc_df.columns)  # 15 columns; col 14=sla_aging_bucket, col 15=sla_breached
    ws_sla.append(sla_cols)
    _style_header(ws_sla, 1, len(sla_cols))
    for r_i, row in enumerate(esc_df.iloc[:, :13].itertuples(index=False), start=2):
        aging_formula = f'=IF(M{r_i}<=24,"Within SLA (<24h)",IF(M{r_i}<=48,"Warning (24-48h)","SLA Breached (>48h)"))'
        breach_formula = f'=IF(M{r_i}>L{r_i},"YES","NO")'
        ws_sla.append(list(row) + [aging_formula, breach_formula])

        aging_val = esc_df.iloc[r_i - 2]["sla_aging_bucket"]
        aging_cell = ws_sla.cell(row=r_i, column=14)
        if "Within" in aging_val:
            aging_cell.fill = green_fill
        elif "Warning" in aging_val:
            aging_cell.fill = amber_fill
        else:
            aging_cell.fill = red_fill
            aging_cell.font = Font(bold=True)

    for sheet, freeze in [
        (ws_dash, "A5"),
        (ws_batch, "A2"),
        (ws_watch, "A2"),
        (ws_sla, "A2"),
    ]:
        _autosize_sheet(sheet, freeze_cell=freeze)

    out_path = OUTPUT_DIR / "NACH_Mandate_Escalation_Tracker.xlsx"
    wb.save(out_path)
    return out_path


def render_summary_png(
    reconciled_df: pd.DataFrame, watchlist_df: pd.DataFrame, esc_df: pd.DataFrame
) -> Path:
    """Create a 300 DPI executive summary report image for NACH reconciliation & SLA tracking."""
    plt.rcParams.update(
        {
            "figure.facecolor": "#0B1120",
            "axes.facecolor": "#111827",
            "axes.edgecolor": "#1F2937",
            "axes.labelcolor": "#E5E7EB",
            "text.color": "#F8FAFC",
            "xtick.color": "#9CA3AF",
            "ytick.color": "#D1D5DB",
            "grid.color": "#1F2937",
            "font.family": "DejaVu Sans",
        }
    )

    total_txns = len(reconciled_df)
    failed_df = reconciled_df[reconciled_df["clearing_status"] == "FAILED"]
    clear_rate = 100.0 * (1.0 - len(failed_df) / total_txns)
    sla_ok_pct = 100.0 * (esc_df["sla_breached"] == "NO").mean()
    exposure_lakhs = failed_df["debit_amount_inr"].sum() / 100000.0

    fig = plt.figure(figsize=(15.5, 11.0), dpi=300)
    gs = gridspec.GridSpec(3, 2, height_ratios=[0.36, 1.0, 1.0], hspace=0.40, wspace=0.25)

    ax_top = fig.add_subplot(gs[0, :])
    ax_top.axis("off")
    ax_top.text(0.0, 0.92, "AUTOMATED NACH MANDATE FAILURE & ESCALATION TRACKER", fontsize=16.5, fontweight="bold", transform=ax_top.transAxes)
    ax_top.text(
        0.0,
        0.74,
        "Python Reconciliation + Advanced Excel MIS Workbook (COUNTIF/IF Formulas)  |  NPCI Return Code Analytics",
        fontsize=10.5,
        color="#94A3B8",
        transform=ax_top.transAxes,
    )

    kpis = [
        ("BATCH PRESENTMENTS", f"{total_txns:,}", f"Clearance Rate: {clear_rate:.1f}%", "#38BDF8"),
        ("FAILED MANDATES RECONCILED", f"{len(failed_df)}", f"Exposure: INR {exposure_lakhs:.2f}L", "#EF4444"),
        ("RECURRING WATCHLIST", f"{len(watchlist_df)} UMRNs", "Flagged >=3 Consecutive Fails", "#F59E0B"),
        ("ESCALATION SLA COMPLIANCE", f"{sla_ok_pct:.1f}%", f"{len(esc_df)} Bank/Customer Tickets", "#10B981"),
    ]
    for i, (lbl, val, sub, color) in enumerate(kpis):
        x0 = i * 0.253
        ax_top.add_patch(
            plt.Rectangle((x0, 0.04), 0.235, 0.58, transform=ax_top.transAxes, facecolor="#111827", edgecolor=color, lw=1.6, clip_on=False)
        )
        ax_top.text(x0 + 0.014, 0.48, lbl, fontsize=8.5, fontweight="bold", color="#94A3B8", transform=ax_top.transAxes)
        ax_top.text(x0 + 0.014, 0.25, val, fontsize=15.5, fontweight="bold", color=color, transform=ax_top.transAxes)
        ax_top.text(x0 + 0.014, 0.09, sub, fontsize=8.0, color="#CBD5E1", transform=ax_top.transAxes)

    # Panel 1: Failure Root-Cause Distribution (NPCI Return Codes)
    ax1 = fig.add_subplot(gs[1, 0])
    rc_counts = (
        failed_df.groupby(["npci_return_code", "return_description"], as_index=False)
        .size()
        .sort_values("size", ascending=True)
    )
    rc_labels = [f"{r['npci_return_code']} - {r['return_description'][:22]}" for _, r in rc_counts.iterrows()]
    bars1 = ax1.barh(rc_labels, rc_counts["size"], color="#38BDF8", height=0.58)
    ax1.set_title("1. NACH Failure Root-Cause by NPCI Return Code", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax1.set_xlabel("Failed Mandates")
    ax1.grid(axis="x", linestyle="--", alpha=0.45)
    ax1.set_xlim(0, rc_counts["size"].max() * 1.2)
    for b in bars1:
        ax1.text(b.get_width() + 1.5, b.get_y() + b.get_height() / 2, f"{int(b.get_width())}", va="center", fontsize=9)

    # Panel 2: Partner Bank Bounce Rate (%)
    ax2 = fig.add_subplot(gs[1, 1])
    bank_stats = (
        reconciled_df.groupby("partner_bank", as_index=False)
        .agg(
            total=("txn_id", "count"),
            failed=("clearing_status", lambda s: (s == "FAILED").sum()),
        )
    )
    bank_stats["bounce_pct"] = 100.0 * bank_stats["failed"] / bank_stats["total"]
    bank_stats = bank_stats.sort_values("bounce_pct", ascending=True)
    bcols = ["#EF4444" if p > 18 else "#F59E0B" if p > 13 else "#10B981" for p in bank_stats["bounce_pct"]]
    bars2 = ax2.barh(bank_stats["partner_bank"], bank_stats["bounce_pct"], color=bcols, height=0.55)
    ax2.set_title("2. Partner Bank Bounce Rate (%) & Anomaly Detection", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax2.set_xlabel("Bounce Rate (%)")
    ax2.grid(axis="x", linestyle="--", alpha=0.45)
    ax2.set_xlim(0, bank_stats["bounce_pct"].max() * 1.25)
    for b, (_, row) in zip(bars2, bank_stats.iterrows()):
        ax2.text(
            b.get_width() + 0.4,
            b.get_y() + b.get_height() / 2,
            f"{row['bounce_pct']:.1f}% ({int(row['failed'])}/{int(row['total'])})",
            va="center",
            fontsize=8.8,
        )

    # Panel 3: Recurring Failure Cohort
    ax3 = fig.add_subplot(gs[2, 0])
    top_watch = watchlist_df.head(8).sort_values("failed_attempts", ascending=True)
    bars3 = ax3.barh(top_watch["umrn"], top_watch["failed_attempts"], color="#F59E0B", height=0.58)
    ax3.set_title("3. Top Recurring Failure Mandates (Watchlist Cohort)", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax3.set_xlabel("Consecutive / Repeat Failures")
    ax3.grid(axis="x", linestyle="--", alpha=0.45)
    ax3.set_xlim(0, top_watch["failed_attempts"].max() * 1.35)
    for b, (_, row) in zip(bars3, top_watch.iterrows()):
        ax3.text(
            b.get_width() + 0.15,
            b.get_y() + b.get_height() / 2,
            f"{int(row['failed_attempts'])}x | INR {int(row['total_failed_exposure_inr']):,}",
            va="center",
            fontsize=8.5,
        )

    # Panel 4: Escalation SLA Aging Matrix
    ax4 = fig.add_subplot(gs[2, 1])
    aging_order = ["Within SLA (<24h)", "Warning (24-48h)", "SLA Breached (>48h)"]
    aging_counts = esc_df["sla_aging_bucket"].value_counts().reindex(aging_order, fill_value=0)
    acols = ["#10B981", "#F59E0B", "#EF4444"]
    bars4 = ax4.bar(aging_counts.index, aging_counts.values, color=acols, width=0.52)
    ax4.set_title("4. Partner Bank & Customer Escalation SLA Aging", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax4.set_ylabel("Escalation Tickets")
    ax4.grid(axis="y", linestyle="--", alpha=0.45)
    ax4.set_ylim(0, aging_counts.max() * 1.2)
    for b in bars4:
        ax4.text(b.get_x() + b.get_width() / 2, b.get_height() + 1.5, f"{int(b.get_height())}", ha="center", fontsize=9.5, fontweight="bold")

    out_png = SCREENSHOT_DIR / "nach_escalation_tracker_summary.png"
    fig.savefig(out_png, dpi=300, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out_png


def main() -> None:
    """Run NACH mandate reconciliation, generate Excel MIS workbook, and export PNG."""
    ensure_directories()
    master_df, txns_df, esc_df = generate_nach_datasets(num_txns=1500)
    reconciled_df, watchlist_df, esc_enriched = reconcile_and_analyze(master_df, txns_df, esc_df)

    reconciled_df.to_csv(OUTPUT_DIR / "reconciled_batch_report.csv", index=False, encoding="utf-8")
    watchlist_df.to_csv(OUTPUT_DIR / "recurring_failure_watchlist.csv", index=False, encoding="utf-8")
    esc_enriched.to_csv(OUTPUT_DIR / "escalation_sla_tracker.csv", index=False, encoding="utf-8")

    xlsx_path = export_excel_workbook(reconciled_df, watchlist_df, esc_enriched)
    png_path = render_summary_png(reconciled_df, watchlist_df, esc_enriched)

    print(f"[Project 3] Reconciled {len(reconciled_df)} NACH batch txns | Watchlist UMRNs: {len(watchlist_df)}")
    print(f"[Project 3] Formula-driven Excel MIS workbook -> {xlsx_path}")
    print(f"[Project 3] Summary screenshot -> {png_path}")


if __name__ == "__main__":
    main()
