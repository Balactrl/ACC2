import streamlit as st
import pandas as pd
import os
import io
import hmac
from datetime import datetime

st.set_page_config(
    page_title="License Allocation Studio",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

MASTER_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Master_License_Database.xlsx")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc; }
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 1.8rem 2.2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
    }
    .hero-banner h1 { color: #ffffff; font-size: 1.9rem; font-weight: 700; margin: 0 0 0.3rem 0; }
    .hero-banner p { color: #cbd5e1; font-size: 0.98rem; margin: 0; }
    .kpi-card {
        background: white;
        padding: 1.2rem 1.4rem;
        border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
    }
    .kpi-label { font-size: 0.8rem; font-weight: 600; text-transform: uppercase; color: #64748b; }
    .kpi-value { font-size: 1.65rem; font-weight: 800; color: #0f172a; margin: 0.25rem 0; }
    .kpi-sub { font-size: 0.78rem; color: #10b981; font-weight: 500; }
    .login-brand { text-align: center; padding: 1.5rem 0 1rem; }
    .login-mark {
        display: inline-flex; align-items: center; justify-content: center;
        width: 64px; height: 64px; border-radius: 20px;
        background: linear-gradient(135deg, #06b6d4, #4f46e5);
        font-size: 2rem; box-shadow: 0 10px 24px rgba(37, 99, 235, 0.25);
    }
    .login-brand h1 { color: #f8fafc; font-size: 1.8rem; margin: 1rem 0 0.25rem; }
    .login-brand p { color: #a5b4fc; margin: 0; }
    div.stDownloadButton > button, div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: white;
        font-weight: 600;
        border: none;
        padding: 0.65rem 1.5rem;
        border-radius: 10px;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

def parse_period(period_str):
    start_str, end_str = str(period_str).upper().split("TO")
    start_dt = pd.to_datetime(start_str.strip(), dayfirst=True)
    end_dt = pd.to_datetime(end_str.strip(), dayfirst=True)
    return pd.Series((start_dt, end_dt))

def normalize_header(col):
    if isinstance(col, (datetime, pd.Timestamp)):
        return col.strftime("%b-%y")
    s = str(col).strip()
    try:
        dt = datetime.strptime(s, "%b-%y")
        return dt.strftime("%b-%y")
    except ValueError:
        pass
    try:
        dt = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%b-%y")
    except ValueError:
        return s

def is_month_col(col_name):
    try:
        datetime.strptime(str(col_name).strip(), "%b-%y")
        return True
    except ValueError:
        return False

def sort_columns(df):
    df = df.copy()
    df.columns = [normalize_header(c) for c in df.columns]
    first_col = next(iter(df.columns))
    df = df[df[first_col].astype(str).str.upper() != "GRAND TOTAL"].copy()

    month_cols = sorted(
        [c for c in df.columns if is_month_col(c)],
        key=lambda x: datetime.strptime(x, "%b-%y")
    )
    base_cols = [c for c in df.columns if c not in month_cols]
    for m in month_cols:
        df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0.0).round(2)
    if "LIESENCE AMOUNT" in df.columns:
        df["LIESENCE AMOUNT"] = pd.to_numeric(df["LIESENCE AMOUNT"], errors="coerce").fillna(0.0).round(2)
    if "PER DAY" in df.columns:
        df["PER DAY"] = pd.to_numeric(df["PER DAY"], errors="coerce").fillna(0.0).round(6)
    return df[base_cols + month_cols], base_cols, month_cols

def add_total_row(dataframe, numeric_cols):
    total_row = {col: "" for col in dataframe.columns}
    first_col = next(iter(dataframe.columns))
    total_row[first_col] = "GRAND TOTAL"
    for col in numeric_cols:
        if col in dataframe.columns:
            total_row[col] = round(pd.to_numeric(dataframe[col], errors="coerce").sum(), 2)
    return pd.concat([dataframe, pd.DataFrame([total_row])], ignore_index=True)

def to_excel_bytes(dataframe, sheet_name):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def find_column(columns, target):
    target_clean = "".join(char for char in str(target).upper() if char.isalnum())
    for col in columns:
        cleaned = "".join(char for char in str(col).upper() if char.isalnum())
        if cleaned == target_clean:
            return col
    return None

def find_site_column(columns):
    for name in ("SITEID", "SITECODE", "ID"):
        column = find_column(columns, name)
        if column is not None:
            return column
    return next(iter(columns))

def text_column_matches(dataframe, column, search_text):
    values = dataframe[column].astype("string").fillna("").str.replace(
        r"\.0$", "", regex=True
    )
    return values.str.contains(search_text, case=False, regex=False, na=False)

def show_login_page():
    st.markdown("""
    <style>
        .stApp {
            background:
                radial-gradient(circle at 18% 12%, rgba(37, 99, 235, 0.28), transparent 34%),
                linear-gradient(135deg, #020617 0%, #0b1120 52%, #111827 100%);
        }
        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stForm"] {
            padding: 1.5rem;
            border: 1px solid #334155;
            border-radius: 18px;
            background: rgba(15, 23, 42, 0.92);
            box-shadow: 0 24px 60px rgba(0, 0, 0, 0.35);
        }
        [data-testid="stForm"] label { color: #e2e8f0; }
        [data-testid="stForm"] input {
            color: #f8fafc;
            background-color: #111827;
            border-color: #475569;
        }
        [data-testid="stForm"] input::placeholder { color: #94a3b8; }
        [data-testid="stFormSubmitButton"] button {
            background: linear-gradient(135deg, #06b6d4, #2563eb) !important;
            color: #ffffff !important;
            border: 0;
        }
    </style>
    """, unsafe_allow_html=True)
    _, login_column, _ = st.columns([1, 1.15, 1])
    with login_column:
        st.markdown("""
        <div class="login-brand">
            <div class="login-mark">💎</div>
            <h1>Welcome back</h1>
            <p>Sign in to License Period Allocation</p>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input(
                "Password",
                placeholder="Enter your password",
                type="password",
            )
            submitted = st.form_submit_button(
                "Sign in",
                type="primary",
                use_container_width=True,
            )

        if submitted:
            login_config = st.secrets.get("login", {})
            expected_username = login_config.get("username", "")
            expected_password = login_config.get("password", "")
            if (
                expected_username
                and expected_password
                and hmac.compare_digest(username, expected_username)
                and hmac.compare_digest(password, expected_password)
            ):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect username or password.")

if not st.session_state.get("authenticated", False):
    show_login_page()
    st.stop()

with st.sidebar:
    st.markdown("### ⚙️ Control Panel")
    if st.button("🚪 Log out", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()
    st.info("**Auto-Save Master File:** Every uploaded batch is automatically merged into `Master_License_Database.xlsx`.")
    st.markdown("#### 📄 Sample Template")
    sample_df = pd.DataFrame(
        [
            ("10001", "BNG", "KA", "192012-RF004320122190", "FA20-21","01-04-2024 TO 01-04-2026", 3000),
            ("10002", "KAR", "TN","192-RB04320122190","FA20-2F", "01-04-2024 TO 01-04-2027", 5000),
        ],
        columns=("SITEID", "SITENAME", "REGION", "DLNO", "FSSAI", "PERIOD DATE", "LIESENCE AMOUNT")
    )
    sample_buf = io.BytesIO()
    with pd.ExcelWriter(sample_buf, engine="openpyxl") as writer:
        sample_df.to_excel(writer, index=False)
    st.download_button("⬇️ Download Sample Excel", sample_buf.getvalue(), "Sample_License_Input.xlsx")

    if os.path.exists(MASTER_FILE):
        st.markdown("---")
        st.markdown("#### 🗄️ Master Storage")
        st.caption(f"Saved at:\n`{MASTER_FILE}`")
        if st.button("🗑️ Clear Master Database"):
            os.remove(MASTER_FILE)
            st.session_state.pop("last_uploaded_sig", None)
            st.session_state.pop("current_batch_df", None)
            st.session_state.pop("current_batch_name", None)
            st.rerun()

st.markdown("""
<div class="hero-banner">
    <h1>💎 License Period & Month-Wise Allocator</h1>
    <p>View your current uploaded file, cumulative Master database, and month-wise paid site reports in one place.</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload New Batch (Excel or CSV)", type=("xlsx", "xls", "csv"))

if uploaded_file is not None:
    file_sig = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("last_uploaded_sig") != file_sig:
        new_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
        new_df.columns = [str(c).strip() for c in new_df.columns]

        if "PERIOD DATE" not in new_df.columns or "LIESENCE AMOUNT" not in new_df.columns:
            st.error("Missing required columns: `PERIOD DATE` and `LIESENCE AMOUNT`.")
            st.stop()

        new_df[["_start", "_end"]] = new_df["PERIOD DATE"].apply(parse_period)
        new_df["DAYS"] = (new_df["_end"] - new_df["_start"]).dt.days + 1
        new_df["PER DAY"] = (new_df["LIESENCE AMOUNT"] / new_df["DAYS"]).round(6)
        new_df["UPLOADED AT"] = datetime.now().strftime("%d-%m-%Y %I:%M %p")

        min_m = new_df["_start"].min().to_period("M")
        max_m = new_df["_end"].max().to_period("M")
        batch_months = pd.period_range(start=min_m, end=max_m, freq="M")

        for period in batch_months:
            col_name = period.strftime("%b-%y")
            m_start = period.to_timestamp()
            m_end = period.to_timestamp(how="end").normalize()

            a_start = new_df["_start"].clip(lower=m_start)
            a_end = new_df["_end"].clip(upper=m_end)
            d_count = ((a_end - a_start).dt.days + 1).clip(lower=0)
            new_df[col_name] = (d_count * (new_df["LIESENCE AMOUNT"] / new_df["DAYS"])).round(2)

        new_clean = new_df.drop(columns=["_start", "_end"])
        new_clean, _, _ = sort_columns(new_clean)

        st.session_state["current_batch_df"] = new_clean.copy()
        st.session_state["current_batch_name"] = uploaded_file.name

        if os.path.exists(MASTER_FILE):
            existing_master = pd.read_excel(MASTER_FILE)
            existing_master, _, _ = sort_columns(existing_master)
            combined_master = pd.concat([existing_master, new_clean], ignore_index=True)
        else:
            combined_master = new_clean.copy()

        combined_master, _, _ = sort_columns(combined_master)
        combined_master.to_excel(MASTER_FILE, index=False, sheet_name="Master_All_Data")

        st.session_state["last_uploaded_sig"] = file_sig
        st.success(f"✅ Processed **{len(new_clean)}** rows from `{uploaded_file.name}` and saved to `Master_License_Database.xlsx`!")

has_master = os.path.exists(MASTER_FILE)
has_batch = "current_batch_df" in st.session_state

if not has_master and not has_batch:
    st.markdown("""
    <div style="text-align:center; padding: 3rem; background: white; border-radius: 14px; border: 1px dashed #cbd5e1; color: #64748b;">
        <h4>No Data Available Yet</h4>
        <p>Upload an Excel or CSV file above to view your current upload, cumulative Master file, and month-wise reports.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

master_df = pd.DataFrame()
m_base_cols = []
m_month_cols = []

if has_master:
    raw_master = pd.read_excel(MASTER_FILE)
    master_df, m_base_cols, m_month_cols = sort_columns(raw_master)
    first_m = next(iter(m_month_cols), "N/A")
    last_m = next(reversed(m_month_cols), "N/A")

    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-label">Cumulative Master Sites</div><div class="kpi-value">{len(master_df):,}</div><div class="kpi-sub">✓ Saved in Master Excel</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-label">Cumulative License Amount</div><div class="kpi-value">{pd.to_numeric(master_df["LIESENCE AMOUNT"], errors="coerce").sum():,.2f}</div><div class="kpi-sub">100% Allocated</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-label">Master Timeline</div><div class="kpi-value" style="font-size:1.35rem;padding-top:0.3rem;">{first_m} → {last_m}</div><div class="kpi-sub">{len(m_month_cols)} Calendar Months</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

# 3 OPTIONS NAVIGATION BAR (Uses st.radio horizontal instead of st.tabs for 100% rendering reliability)
selected_view = st.radio(
    "Select View",
    (
        "📤 1. Current Uploaded File",
        "📋 2. Cumulative Master Data (All Sites)",
        "🎯 3. Month-Wise Site Report"
    ),
    horizontal=True,
    label_visibility="collapsed"
)
st.markdown("---")

if selected_view == "📤 1. Current Uploaded File":
    if has_batch:
        batch_df = st.session_state["current_batch_df"]
        st.subheader(f"Current Upload: {st.session_state['current_batch_name']}")
        st.caption(f"{len(batch_df):,} sites in this uploaded batch")
        st.dataframe(batch_df, width="stretch", hide_index=True)
        st.download_button(
            "⬇️ Download Current Batch",
            to_excel_bytes(batch_df, "Current_Batch"),
            file_name="Current_License_Batch.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Upload a file above to view the current batch. The cumulative master data is still available in option 2.")

elif selected_view == "📋 2. Cumulative Master Data (All Sites)":
    if has_master and not master_df.empty:
        st.subheader("Cumulative Master Data (All Sites)")
        st.caption(f"{len(master_df):,} records saved in the master database")
        st.dataframe(master_df, width="stretch", hide_index=True)
        st.download_button(
            "⬇️ Download Cumulative Master Data",
            to_excel_bytes(master_df, "Master_All_Data"),
            file_name="Cumulative_Master_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No cumulative master records are available yet. Upload a file above to build the master database.")

elif selected_view == "🎯 3. Month-Wise Site Report":
    if has_master and not master_df.empty and m_month_cols:
        st.subheader("Month-Wise Site Report")
        st.caption("Choose a month, date range, Site ID, DLNO, or FSSAI. Any one filter can be used by itself.")
        selected_month = st.selectbox(
            "Select month",
            m_month_cols,
            index=None,
            placeholder="Select a month",
            key="month_report_month",
        )
        date_col1, date_col2 = st.columns(2)
        with date_col1:
            from_date = st.date_input(
                "From Date",
                value=None,
                format="DD/MM/YYYY",
                key="month_report_from_date_empty",
            )
        with date_col2:
            to_date = st.date_input(
                "To Date",
                value=None,
                format="DD/MM/YYYY",
                key="month_report_to_date_empty",
            )
        search_col1, search_col2, search_col3 = st.columns(3)
        with search_col1:
            site_search = st.text_input(
                "Search by Site ID",
                placeholder="Enter a Site ID or part of one",
                key="month_report_site_search",
            ).strip()
        with search_col2:
            dlno_search = st.text_input(
                "Search by DLNO",
                placeholder="Enter a DLNO or part of one",
                key="month_report_dlno_search",
            ).strip()
        with search_col3:
            fssai_search = st.text_input(
                "Search by FSSAI",
                placeholder="Enter an FSSAI number or part of one",
                key="month_report_fssai_search",
            ).strip()

        has_date_filter = from_date is not None or to_date is not None
        has_search_filter = bool(
            selected_month or has_date_filter or site_search or dlno_search or fssai_search
        )
        requested_searches = {
            "Site ID": site_search,
            "DLNO": dlno_search,
            "FSSAI": fssai_search,
        }
        search_columns = {
            "Site ID": find_site_column(master_df.columns),
            "DLNO": find_column(master_df.columns, "DLNO"),
            "FSSAI": find_column(master_df.columns, "FSSAI"),
        }
        missing_search_columns = [
            label
            for label, search_text in requested_searches.items()
            if search_text and search_columns[label] is None
        ]

        if has_search_filter and from_date and to_date and from_date > to_date:
            st.error("From Date must be on or before To Date.")
        elif missing_search_columns:
            st.error(
                "Cannot search by "
                + ", ".join(missing_search_columns)
                + ": the master database does not contain the required column."
            )
        elif has_search_filter:
            period_bounds = master_df["PERIOD DATE"].apply(parse_period)
            license_starts = pd.to_datetime(period_bounds[0])
            license_ends = pd.to_datetime(period_bounds[1])
            date_matches = pd.Series(True, index=master_df.index)
            if from_date:
                date_matches &= license_ends.dt.date >= from_date
            if to_date:
                date_matches &= license_starts.dt.date <= to_date

            if site_search:
                date_matches &= text_column_matches(
                    master_df, search_columns["Site ID"], site_search
                )
            if dlno_search:
                date_matches &= text_column_matches(
                    master_df, search_columns["DLNO"], dlno_search
                )
            if fssai_search:
                date_matches &= text_column_matches(
                    master_df, search_columns["FSSAI"], fssai_search
                )

            if selected_month:
                report_months = [selected_month]
                month_amounts = pd.to_numeric(master_df[selected_month], errors="coerce").fillna(0)
                date_matches &= month_amounts > 0
            else:
                report_months = list(m_month_cols)
                if has_date_filter:
                    report_months = []
                    for month in m_month_cols:
                        month_start = datetime.strptime(month, "%b-%y").date()
                        month_end = (pd.Timestamp(month_start) + pd.offsets.MonthEnd(0)).date()
                        if (
                            (from_date is None or month_end >= from_date)
                            and (to_date is None or month_start <= to_date)
                        ):
                            report_months.append(month)
                    if report_months:
                        range_amounts = master_df[report_months].apply(
                            pd.to_numeric, errors="coerce"
                        ).fillna(0).sum(axis=1)
                        date_matches &= range_amounts > 0
                    else:
                        date_matches &= False

            report_data = master_df.loc[
                date_matches,
                m_base_cols + report_months,
            ].copy()
            if selected_month and not report_data.empty:
                report_data = report_data.sort_values(selected_month, ascending=False)

            if not report_data.empty:
                result_bounds = report_data["PERIOD DATE"].apply(parse_period)
                report_data["FROM DATE"] = result_bounds[0].dt.strftime("%d-%m-%Y")
                report_data["TO DATE"] = result_bounds[1].dt.strftime("%d-%m-%Y")
                report_columns = []
                for column in m_base_cols:
                    if column == "PERIOD DATE":
                        report_columns.extend(["FROM DATE", "TO DATE"])
                    else:
                        report_columns.append(column)
                report_columns.extend(report_months)
                report_data = report_data[report_columns]

            metric_col1, metric_col2 = st.columns(2)
            metric_col1.metric("Sites in report", f"{len(report_data):,}")
            allocation_total = (
                report_data[report_months]
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0)
                .sum()
                .sum()
                if not report_data.empty else 0
            )
            metric_col2.metric(
                "Total allocated",
                f"{allocation_total:,.2f}",
            )

            if report_data.empty:
                st.info("No sites matched the selected filters.")
            else:
                st.dataframe(report_data, width="stretch", hide_index=True)
                report_label = selected_month or "Filtered"
                st.download_button(
                    f"⬇️ Download {report_label} Site Report",
                    to_excel_bytes(report_data, "Site_Report"),
                    file_name=f"Site_Report_{report_label.replace('-', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("Select a month, choose a date, or enter a Site ID, DLNO, or FSSAI to display a report.")
    else:
        st.info("No monthly allocation data is available yet. Upload a file with valid PERIOD DATE and LIESENCE AMOUNT values.")

