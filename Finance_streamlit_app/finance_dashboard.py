"""
BUSINSESS PAYOFF DASHBOARD
Business Investment Payoff Dashboard
A self-contained Streamlit app that tracks how an entrepreneur is paying off
their initial investment by allocating a percentage of each job's revenue
toward debt, and the rest toward savings.

Storage: local SQLite database (data/dashboard.db).
"""

import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime, date
import plotly.graph_objects as go
import io

# ---------- Page config ----------
st.set_page_config(
    page_title="Dashboard Para tus finazas de Emprendedora Estética ✨",
    page_icon="💅",
    layout="wide",
)

# ---------- Database ----------
DB_DIR = Path("data")
DB_DIR.mkdir(exist_ok=True)
DB_PATH = DB_DIR / "dashboard.db"


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                description TEXT,
                revenue REAL NOT NULL,
                debt_pct REAL NOT NULL,
                to_debt REAL NOT NULL,
                to_savings REAL NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def get_setting(key: str, default=None):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT value FROM settings WHERE key = ?", (key,)
        ).fetchone()
        return row["value"] if row else default


def set_setting(key: str, value):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, str(value)),
        )
        conn.commit()


def add_job(d: date, description: str, revenue: float, debt_pct: float):
    to_debt = revenue * (debt_pct / 100)
    to_savings = revenue - to_debt
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO jobs (date, description, revenue, debt_pct, to_debt, to_savings, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                d.isoformat(),
                description,
                revenue,
                debt_pct,
                to_debt,
                to_savings,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()


def delete_job(job_id: int):
    with get_conn() as conn:
        conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
        conn.commit()


def get_jobs_df() -> pd.DataFrame:
    with get_conn() as conn:
        df = pd.read_sql_query(
            "SELECT * FROM jobs ORDER BY date ASC, id ASC", conn
        )
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"])
    return df


def reset_all_data():
    with get_conn() as conn:
        conn.execute("DELETE FROM jobs")
        conn.execute("DELETE FROM settings")
        conn.commit()


# ---------- Helpers ----------
def money(x: float) -> str:
    return f"${x:,.2f}"


def get_initial_investment() -> float:
    val = get_setting("initial_investment", "0")
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def get_default_debt_pct() -> float:
    val = get_setting("default_debt_pct", "50")
    try:
        return float(val)
    except (TypeError, ValueError):
        return 50.0


# ---------- Charts ----------
def progress_ring(pct: float):
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=pct,
            number={"suffix": "%", "font": {"size": 42}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 1},
                "bar": {"color": "#10b981"},
                "steps": [
                    {"range": [0, 33], "color": "#fee2e2"},
                    {"range": [33, 66], "color": "#fef3c7"},
                    {"range": [66, 100], "color": "#d1fae5"},
                ],
                "threshold": {
                    "line": {"color": "#059669", "width": 4},
                    "thickness": 0.75,
                    "value": pct,
                },
            },
            domain={"x": [0, 1], "y": [0, 1]},
        )
    )
    fig.update_layout(height=280, margin=dict(l=20, r=20, t=20, b=20))
    return fig


def progress_chart(df: pd.DataFrame, initial_investment: float):
    df = df.copy()
    df["cum_debt"] = df["to_debt"].cumsum()
    df["cum_savings"] = df["to_savings"].cumsum()

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df["date"],
            y=df["cum_debt"],
            mode="lines+markers",
            name="Toward debt (cumulative)",
            line=dict(color="#10b981", width=3),
            fill="tozeroy",
            fillcolor="rgba(16, 185, 129, 0.15)",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=df["date"],
            y=df["cum_savings"],
            mode="lines+markers",
            name="Saved (cumulative)",
            line=dict(color="#3b82f6", width=3),
        )
    )
    fig.add_hline(
        y=initial_investment,
        line_dash="dash",
        line_color="#ef4444",
        annotation_text=f"Goal: {money(initial_investment)}",
        annotation_position="top left",
    )
    fig.update_layout(
        height=400,
        margin=dict(l=20, r=20, t=40, b=20),
        xaxis_title="Date",
        yaxis_title="Amount ($)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


# ---------- Exports ----------
def build_excel_report(df: pd.DataFrame, initial_investment: float, month_label: str) -> bytes:
    """Build a multi-sheet xlsx for the given month's data."""
    output = io.BytesIO()

    total_revenue = df["revenue"].sum() if not df.empty else 0
    total_debt = df["to_debt"].sum() if not df.empty else 0
    total_savings = df["to_savings"].sum() if not df.empty else 0

    summary = pd.DataFrame(
        {
            "Metric": [
                "Report period",
                "Initial investment",
                "Jobs logged this month",
                "Total revenue this month",
                "Applied to debt this month",
                "Added to savings this month",
                "Total paid toward debt (all-time)",
                "Total saved (all-time)",
                "Remaining debt",
                "% paid off",
            ],
            "Value": [
                month_label,
                money(initial_investment),
                len(df),
                money(total_revenue),
                money(total_debt),
                money(total_savings),
                "",  # filled by caller
                "",
                "",
                "",
            ],
        }
    )

    detail = df.copy()
    if not detail.empty:
        detail["date"] = detail["date"].dt.strftime("%Y-%m-%d")
        detail = detail[
            ["date", "description", "revenue", "debt_pct", "to_debt", "to_savings"]
        ].rename(
            columns={
                "date": "Date",
                "description": "Job",
                "revenue": "Revenue",
                "debt_pct": "Debt %",
                "to_debt": "To Debt",
                "to_savings": "To Savings",
            }
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        if not detail.empty:
            detail.to_excel(writer, sheet_name="Jobs", index=False)
        else:
            pd.DataFrame({"Note": ["No jobs logged in this period."]}).to_excel(
                writer, sheet_name="Jobs", index=False
            )

    return output.getvalue()


def build_full_excel_report(all_df: pd.DataFrame, month_df: pd.DataFrame,
                             initial_investment: float, month_label: str,
                             total_debt_paid: float, total_saved: float,
                             remaining: float, pct_paid: float) -> bytes:
    output = io.BytesIO()

    summary = pd.DataFrame(
        {
            "Metric": [
                "Report period",
                "Initial investment",
                "Jobs this month",
                "Revenue this month",
                "Applied to debt this month",
                "Added to savings this month",
                "—",
                "All-time paid toward debt",
                "All-time saved",
                "Remaining debt",
                "% paid off",
            ],
            "Value": [
                month_label,
                money(initial_investment),
                len(month_df),
                money(month_df["revenue"].sum()) if not month_df.empty else money(0),
                money(month_df["to_debt"].sum()) if not month_df.empty else money(0),
                money(month_df["to_savings"].sum()) if not month_df.empty else money(0),
                "",
                money(total_debt_paid),
                money(total_saved),
                money(remaining),
                f"{pct_paid:.1f}%",
            ],
        }
    )

    def fmt_jobs(d: pd.DataFrame) -> pd.DataFrame:
        if d.empty:
            return pd.DataFrame({"Note": ["No jobs in this period."]})
        out = d.copy()
        out["date"] = out["date"].dt.strftime("%Y-%m-%d")
        return out[["date", "description", "revenue", "debt_pct", "to_debt", "to_savings"]].rename(
            columns={
                "date": "Date",
                "description": "Job",
                "revenue": "Revenue",
                "debt_pct": "Debt %",
                "to_debt": "To Debt",
                "to_savings": "To Savings",
            }
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        fmt_jobs(month_df).to_excel(writer, sheet_name="This Month", index=False)
        fmt_jobs(all_df).to_excel(writer, sheet_name="All Jobs", index=False)

    return output.getvalue()


def build_pdf_report(month_df: pd.DataFrame, initial_investment: float, month_label: str,
                      total_debt_paid: float, total_saved: float,
                      remaining: float, pct_paid: float) -> bytes:
    """Build a simple monthly PDF summary."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    )

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=letter,
                             leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                             topMargin=0.6 * inch, bottomMargin=0.6 * inch)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Heading1"], fontSize=20, textColor=colors.HexColor("#0f172a")
    )
    h2 = ParagraphStyle(
        "h2", parent=styles["Heading2"], fontSize=13, textColor=colors.HexColor("#334155")
    )

    elements = []
    elements.append(Paragraph("Business Payoff — Monthly Report", title_style))
    elements.append(Paragraph(f"Period: <b>{month_label}</b>", styles["Normal"]))
    elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    elements.append(Spacer(1, 0.25 * inch))

    elements.append(Paragraph("Overall progress", h2))
    overall_data = [
        ["Initial investment", money(initial_investment)],
        ["All-time paid toward debt", money(total_debt_paid)],
        ["All-time saved", money(total_saved)],
        ["Remaining debt", money(remaining)],
        ["% paid off", f"{pct_paid:.1f}%"],
    ]
    overall_tbl = Table(overall_data, colWidths=[2.8 * inch, 2.5 * inch])
    overall_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f1f5f9")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#0f172a")),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(overall_tbl)
    elements.append(Spacer(1, 0.25 * inch))

    elements.append(Paragraph(f"This month ({month_label})", h2))
    if month_df.empty:
        elements.append(Paragraph("No jobs logged this month.", styles["Normal"]))
    else:
        month_revenue = month_df["revenue"].sum()
        month_debt = month_df["to_debt"].sum()
        month_savings = month_df["to_savings"].sum()
        month_data = [
            ["Jobs logged", str(len(month_df))],
            ["Revenue", money(month_revenue)],
            ["Applied to debt", money(month_debt)],
            ["Added to savings", money(month_savings)],
        ]
        month_tbl = Table(month_data, colWidths=[2.8 * inch, 2.5 * inch])
        month_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f1f5f9")),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
        ]))
        elements.append(month_tbl)
        elements.append(Spacer(1, 0.25 * inch))

        elements.append(Paragraph("Job detail", h2))
        detail_rows = [["Date", "Job", "Revenue", "Debt %", "To Debt", "To Savings"]]
        for _, r in month_df.iterrows():
            detail_rows.append([
                r["date"].strftime("%Y-%m-%d"),
                str(r["description"])[:40] if r["description"] else "—",
                money(r["revenue"]),
                f"{r['debt_pct']:.0f}%",
                money(r["to_debt"]),
                money(r["to_savings"]),
            ])
        detail_tbl = Table(detail_rows, colWidths=[0.9 * inch, 2.0 * inch, 0.9 * inch, 0.7 * inch, 0.9 * inch, 0.9 * inch])
        detail_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(detail_tbl)

    doc.build(elements)
    return output.getvalue()


# ---------- App ----------
init_db()
initial_investment = get_initial_investment()
default_debt_pct = get_default_debt_pct()
jobs_df = get_jobs_df()

# ---------- Sidebar ----------
with st.sidebar:
    st.title("⚙️ Setup")

    st.markdown("##### 1. Initial investment")
    st.caption("How much money you put in to start the business.")
    new_investment = st.number_input(
        "Initial investment ($)",
        min_value=0.0,
        value=float(initial_investment),
        step=100.0,
        format="%.2f",
        key="ii_input",
    )
    if st.button("💾 Save initial investment", use_container_width=True):
        set_setting("initial_investment", new_investment)
        st.success("Saved!")
        st.rerun()

    st.markdown("---")
    st.markdown("##### 2. Default debt allocation %")
    st.caption("Pre-fills the slider when logging a new job.")
    new_default_pct = st.slider(
        "Default % of revenue toward debt",
        min_value=0,
        max_value=100,
        value=int(default_debt_pct),
        step=5,
        key="default_pct_slider",
    )
    if st.button("💾 Save default %", use_container_width=True):
        set_setting("default_debt_pct", new_default_pct)
        st.success("Saved!")
        st.rerun()

    st.markdown("---")
    with st.expander("⚠️ Danger zone"):
        st.caption("Deletes ALL jobs and resets settings. Cannot be undone.")
        confirm = st.checkbox("I understand, reset everything")
        if st.button("🗑️ Reset all data", disabled=not confirm, use_container_width=True):
            reset_all_data()
            st.success("All data cleared.")
            st.rerun()

# ---------- Onboarding when no investment is set ----------
if initial_investment <= 0:
    st.title("💰 Business Payoff Dashboard")
    st.markdown("#### Welcome! Let's get you set up in 3 quick steps.")

    with st.container(border=True):
        st.markdown("### 👈 Step 1 — Enter your initial investment")
        st.markdown(
            "Open the sidebar (top-left) and enter the **total amount you invested** "
            "to start your business. Click **Save initial investment**."
        )

    with st.container(border=True):
        st.markdown("### 🎯 Step 2 — Set your default debt allocation")
        st.markdown(
            "Choose what % of each job's revenue should go toward paying back the debt. "
            "Example: **50%** means half of every job pays the debt, half goes to savings. "
            "You can override this per-job later."
        )

    with st.container(border=True):
        st.markdown("### 💼 Step 3 — Log jobs as money comes in")
        st.markdown(
            "Each time you complete a job, log the revenue here. The dashboard will "
            "automatically split it between **debt** and **savings**, and show you progress."
        )

    st.info("Once you save your initial investment in the sidebar, the dashboard will appear here.")
    st.stop()

# ---------- Main dashboard ----------
total_debt_paid = float(jobs_df["to_debt"].sum()) if not jobs_df.empty else 0.0
total_saved = float(jobs_df["to_savings"].sum()) if not jobs_df.empty else 0.0
total_revenue = float(jobs_df["revenue"].sum()) if not jobs_df.empty else 0.0
remaining = max(initial_investment - total_debt_paid, 0.0)
pct_paid = min((total_debt_paid / initial_investment * 100) if initial_investment > 0 else 0.0, 100.0)

st.title("💰 Business Payoff Dashboard")
st.caption(f"Initial investment: **{money(initial_investment)}** · Default debt allocation: **{int(default_debt_pct)}%**")

# Top KPI row
k1, k2, k3, k4 = st.columns(4)
k1.metric("% Paid off", f"{pct_paid:.1f}%")
k2.metric("Paid toward debt", money(total_debt_paid))
k3.metric("Remaining debt", money(remaining))
k4.metric("Savings balance", money(total_saved))

st.markdown("---")

# ---------- Tabs ----------
tab_log, tab_progress, tab_history, tab_export = st.tabs(
    ["➕ Log a job", "📊 Progress", "📋 History", "📥 Monthly report"]
)

# ===== Tab 1: Log a job =====
with tab_log:
    st.subheader("Log a new job")
    st.caption(
        "Each time you finish a job and get paid, log it here. The amount will be "
        "split automatically between debt repayment and your savings balance."
    )

    with st.form("add_job_form", clear_on_submit=True):
        c1, c2 = st.columns([1, 2])
        with c1:
            job_date = st.date_input("Date", value=date.today())
        with c2:
            description = st.text_input(
                "Job description",
                placeholder="e.g. Logo design for ACME Corp",
            )

        c3, c4 = st.columns(2)
        with c3:
            revenue = st.number_input(
                "Revenue from this job ($)",
                min_value=0.0,
                value=0.0,
                step=50.0,
                format="%.2f",
            )
        with c4:
            debt_pct = st.slider(
                "% of this revenue toward debt",
                min_value=0,
                max_value=100,
                value=int(default_debt_pct),
                step=5,
                help="The rest goes to your savings balance.",
            )

        # Live preview
        if revenue > 0:
            preview_debt = revenue * (debt_pct / 100)
            preview_savings = revenue - preview_debt
            st.info(
                f"📊 Preview: **{money(preview_debt)}** to debt · "
                f"**{money(preview_savings)}** to savings"
            )

        submitted = st.form_submit_button("➕ Add job", type="primary", use_container_width=True)
        if submitted:
            if revenue <= 0:
                st.error("Revenue must be greater than zero.")
            else:
                add_job(job_date, description.strip(), revenue, float(debt_pct))
                st.success(
                    f"Logged {money(revenue)} — {money(revenue * debt_pct / 100)} to debt, "
                    f"{money(revenue * (1 - debt_pct / 100))} to savings."
                )
                st.rerun()

# ===== Tab 2: Progress =====
with tab_progress:
    st.subheader("Where you stand")

    if jobs_df.empty:
        st.info("No jobs logged yet. Head to the **Log a job** tab to add your first one.")
    else:
        left, right = st.columns([1, 1])
        with left:
            st.plotly_chart(progress_ring(pct_paid), use_container_width=True)
            st.progress(pct_paid / 100)
            st.caption(
                f"{money(total_debt_paid)} of {money(initial_investment)} paid"
            )

        with right:
            st.markdown("##### Quick stats")
            n_jobs = len(jobs_df)
            avg_rev = total_revenue / n_jobs if n_jobs else 0
            st.metric("Total jobs logged", f"{n_jobs}")
            st.metric("Total revenue earned", money(total_revenue))
            st.metric("Average per job", money(avg_rev))

            if remaining > 0 and avg_rev > 0:
                avg_to_debt_per_job = total_debt_paid / n_jobs
                if avg_to_debt_per_job > 0:
                    jobs_needed = remaining / avg_to_debt_per_job
                    st.caption(
                        f"💡 At your current pace, **~{jobs_needed:.0f} more jobs** "
                        f"like your average will pay off the remaining debt."
                    )
            elif remaining == 0:
                st.success("🎉 Debt fully paid off!")

        st.markdown("---")
        st.subheader("📈 Cumulative progress over time")
        st.plotly_chart(progress_chart(jobs_df, initial_investment), use_container_width=True)

# ===== Tab 3: History =====
with tab_history:
    st.subheader("All logged jobs")

    if jobs_df.empty:
        st.info("No jobs logged yet.")
    else:
        st.caption(f"{len(jobs_df)} jobs logged · click 🗑️ next to a row to delete it.")

        view_df = jobs_df.copy()
        view_df = view_df.sort_values("date", ascending=False).reset_index(drop=True)
        view_df["date"] = view_df["date"].dt.strftime("%Y-%m-%d")

        # Header
        h = st.columns([1.2, 2.5, 1.2, 0.8, 1.2, 1.2, 0.6])
        h[0].markdown("**Date**")
        h[1].markdown("**Job**")
        h[2].markdown("**Revenue**")
        h[3].markdown("**Debt %**")
        h[4].markdown("**To Debt**")
        h[5].markdown("**To Savings**")
        h[6].markdown("**🗑️**")

        for _, row in view_df.iterrows():
            cols = st.columns([1.2, 2.5, 1.2, 0.8, 1.2, 1.2, 0.6])
            cols[0].write(row["date"])
            cols[1].write(row["description"] or "—")
            cols[2].write(money(row["revenue"]))
            cols[3].write(f"{int(row['debt_pct'])}%")
            cols[4].write(money(row["to_debt"]))
            cols[5].write(money(row["to_savings"]))
            if cols[6].button("🗑️", key=f"del_{row['id']}", help="Delete this job"):
                delete_job(int(row["id"]))
                st.rerun()

# ===== Tab 4: Export =====
with tab_export:
    st.subheader("Monthly progress report")
    st.caption("Generate an Excel or PDF file showing the month's activity and overall progress.")

    if jobs_df.empty:
        st.info("Log at least one job before generating a report.")
    else:
        # Pick a month
        jobs_df["year_month"] = jobs_df["date"].dt.to_period("M")
        available_months = sorted(jobs_df["year_month"].unique(), reverse=True)
        month_options = [str(m) for m in available_months]
        current_month = str(pd.Timestamp.now().to_period("M"))
        if current_month not in month_options:
            month_options.insert(0, current_month)

        selected_month = st.selectbox(
            "Select a month",
            options=month_options,
            index=0,
        )

        month_period = pd.Period(selected_month, freq="M")
        month_df = jobs_df[jobs_df["year_month"] == month_period].copy()
        month_label = month_period.strftime("%B %Y")

        # Preview
        with st.container(border=True):
            st.markdown(f"##### Preview — {month_label}")
            if month_df.empty:
                st.warning("No jobs in this month.")
            else:
                p1, p2, p3, p4 = st.columns(4)
                p1.metric("Jobs", len(month_df))
                p2.metric("Revenue", money(month_df["revenue"].sum()))
                p3.metric("To debt", money(month_df["to_debt"].sum()))
                p4.metric("To savings", money(month_df["to_savings"].sum()))

        # Download buttons
        st.markdown("##### Download")
        c1, c2 = st.columns(2)

        with c1:
            try:
                excel_bytes = build_full_excel_report(
                    jobs_df.drop(columns=["year_month"]),
                    month_df.drop(columns=["year_month"]) if not month_df.empty else month_df,
                    initial_investment,
                    month_label,
                    total_debt_paid,
                    total_saved,
                    remaining,
                    pct_paid,
                )
                st.download_button(
                    "📊 Download Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"payoff_report_{selected_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Excel export error: {e}")

        with c2:
            try:
                pdf_bytes = build_pdf_report(
                    month_df.drop(columns=["year_month"]) if not month_df.empty else month_df,
                    initial_investment,
                    month_label,
                    total_debt_paid,
                    total_saved,
                    remaining,
                    pct_paid,
                )
                st.download_button(
                    "📄 Download PDF",
                    data=pdf_bytes,
                    file_name=f"payoff_report_{selected_month}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PDF export error: {e}")

st.markdown("---")
st.caption(f"💾 Data saved locally in `{DB_PATH}` · Last action: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")