# ==========================================================
# PUBLIC SURVEY RESPONSE ANALYSER
# VERSION 7.0 - SMART INSIGHTS & PROFESSIONAL REPORTS
# ==========================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import datetime
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from openpyxl.styles import PatternFill, Font


# ----------------------------------------------------------
# 1. PAGE CONFIGURATION
# ----------------------------------------------------------

st.set_page_config(
    page_title="Public Survey Response Analyser",
    page_icon="📊",
    layout="wide"
)
# ----------------------------------------------------------
# MODERN UI DESIGN + ANIMATIONS
# ----------------------------------------------------------

st.markdown("""
<style>

/* ---------- IMPORT MODERN FONT ---------- */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* ---------- GLOBAL DESIGN ---------- */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(ellipse at 10% 5%, rgba(91, 63, 180, 0.18), transparent 35%),
        radial-gradient(ellipse at 90% 15%, rgba(25, 130, 190, 0.13), transparent 35%),
        #0b1020;
    color: #edf2ff;
}

/* Main content spacing */
.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}

/* ---------- ANIMATIONS ---------- */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(18px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

@keyframes gradientMove {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

@keyframes floating {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-5px); }
}

@keyframes glow {
    0%, 100% { box-shadow: 0 0 12px rgba(139, 92, 246, 0.12); }
    50% { box-shadow: 0 0 25px rgba(139, 92, 246, 0.30); }
}

@keyframes slideIn {
    from {
        opacity: 0;
        transform: translateX(-15px);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

.main .block-container {
    animation: fadeIn 0.7s ease-out;
}

/* ---------- HERO HEADER ---------- */
.hero-container {
    position: relative;
    overflow: hidden;
    padding: 38px 35px;
    margin-bottom: 28px;
    border-radius: 22px;
    background: linear-gradient(
        120deg,
        #312e81,
        #5b21b6,
        #1d4ed8,
        #312e81
    );
    background-size: 300% 300%;
    animation: gradientMove 12s ease infinite, glow 5s infinite;
    border: 1px solid rgba(255,255,255,0.16);
}

.hero-container h1 {
    color: white !important;
    font-size: 2.35rem !important;
    font-weight: 800 !important;
    margin-bottom: 10px;
    letter-spacing: -1px;
}

.hero-container p {
    color: #e0e7ff !important;
    font-size: 1rem;
    margin-bottom: 0;
}

.hero-label {
    display: inline-block;
    padding: 6px 13px;
    margin-bottom: 15px;
    border-radius: 30px;
    background: rgba(255,255,255,0.16);
    color: #ffffff;
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 1px;
}

/* ---------- HEADINGS ---------- */
h1, h2, h3 {
    color: #f1f5ff !important;
    font-weight: 700 !important;
}

h2, h3 {
    letter-spacing: -0.4px;
}

p, label, li {
    color: #cbd5e1;
}

/* ---------- SIDEBAR ---------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #11182b, #15132d);
    border-right: 1px solid rgba(148,163,184,0.15);
}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #e9eaff !important;
}

[data-testid="stSidebar"] .stMarkdown {
    color: #cbd5e1;
}

/* ---------- KPI METRIC CARDS ---------- */
[data-testid="stMetric"] {
    background: linear-gradient(
        145deg,
        rgba(30, 41, 72, 0.95),
        rgba(24, 32, 58, 0.95)
    );
    border: 1px solid rgba(139, 92, 246, 0.25);
    border-radius: 17px;
    padding: 19px 20px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.18);
    transition: all 0.3s ease;
    animation: fadeIn 0.65s ease-out;
}

[data-testid="stMetric"]:hover {
    transform: translateY(-5px);
    border-color: rgba(167,139,250,0.75);
    box-shadow: 0 12px 30px rgba(91,63,180,0.22);
}

[data-testid="stMetricLabel"] {
    color: #a5b4fc !important;
    font-weight: 600;
}

[data-testid="stMetricValue"] {
    color: #ffffff !important;
    font-weight: 800;
}

/* ---------- TABS ---------- */
button[data-baseweb="tab"] {
    color: #aebbd3 !important;
    font-weight: 600 !important;
    border-radius: 10px 10px 0 0;
    transition: all 0.25s ease;
}

button[data-baseweb="tab"]:hover {
    color: #ffffff !important;
    background: rgba(139,92,246,0.12);
}

button[data-baseweb="tab"][aria-selected="true"] {
    color: #c4b5fd !important;
    border-bottom: 3px solid #8b5cf6 !important;
}

/* ---------- BUTTONS ---------- */
.stButton > button,
.stDownloadButton > button {
    border: 1px solid rgba(167,139,250,0.45);
    border-radius: 11px;
    background: linear-gradient(110deg, #6d28d9, #4f46e5);
    color: white !important;
    font-weight: 700;
    padding: 10px 18px;
    transition: all 0.25s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    transform: translateY(-3px);
    background: linear-gradient(110deg, #7c3aed, #2563eb);
    box-shadow: 0 8px 22px rgba(99,102,241,0.28);
    border-color: #c4b5fd;
}

/* ---------- INPUTS ---------- */
.stTextInput input,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"],
.stMultiSelect div[data-baseweb="select"],
.stTextArea textarea {
    background: #151e33 !important;
    color: #f8fafc !important;
    border-radius: 10px !important;
    border-color: rgba(148,163,184,0.25) !important;
}

/* ---------- DATA TABLES ---------- */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(148,163,184,0.18);
    border-radius: 13px;
    overflow: hidden;
}

/* ---------- CHART CONTAINERS ---------- */
[data-testid="stPlotlyChart"] {
    background: rgba(21, 30, 51, 0.65);
    border: 1px solid rgba(148,163,184,0.15);
    border-radius: 16px;
    padding: 10px;
    transition: all 0.3s ease;
    animation: fadeIn 0.7s ease-out;
}

[data-testid="stPlotlyChart"]:hover {
    border-color: rgba(139,92,246,0.45);
    box-shadow: 0 8px 25px rgba(0,0,0,0.16);
}

/* ---------- ALERTS ---------- */
[data-testid="stAlert"] {
    border-radius: 12px;
    border: 1px solid rgba(148,163,184,0.20);
}

/* ---------- FILE UPLOADER ---------- */
[data-testid="stFileUploader"] {
    background: rgba(30,41,72,0.55);
    border: 1px dashed rgba(167,139,250,0.55);
    border-radius: 15px;
    padding: 12px;
    transition: all 0.3s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: #a78bfa;
    background: rgba(91,63,180,0.12);
}

/* ---------- DIVIDERS ---------- */
hr {
    border-color: rgba(148,163,184,0.17) !important;
}

/* ---------- CHECKBOXES AND RADIO ---------- */
.stCheckbox label,
.stRadio label {
    color: #dbeafe !important;
}

/* ---------- CAPTIONS ---------- */
.stCaption,
[data-testid="stCaptionContainer"] {
    color: #94a3b8 !important;
}

/* ---------- REDUCE MOTION FOR ACCESSIBILITY ---------- */
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
        animation-duration: 0.01ms !important;
        transition-duration: 0.01ms !important;
    }
}

/* ---------- MOBILE RESPONSIVENESS ---------- */
@media (max-width: 768px) {
    .block-container {
        padding: 1rem;
    }

    .hero-container {
        padding: 25px 20px;
    }

    .hero-container h1 {
        font-size: 1.65rem !important;
    }
}

</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------
# PLOTLY PROFESSIONAL THEME HELPER
# ----------------------------------------------------------

def apply_chart_theme(fig):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            family="Inter, sans-serif",
            color="#e2e8f0",
            size=12
        ),
        title=dict(
            font=dict(
                size=19,
                color="#c4b5fd"
            ),
            x=0.02
        ),
        margin=dict(
            l=35,
            r=25,
            t=65,
            b=45
        ),
        legend=dict(
            bgcolor="rgba(0,0,0,0)"
        ),
        hoverlabel=dict(
            bgcolor="#1e293b",
            font_color="white"
        )
    )

    fig.update_xaxes(
        gridcolor="rgba(148,163,184,0.12)",
        zerolinecolor="rgba(148,163,184,0.2)"
    )

    fig.update_yaxes(
        gridcolor="rgba(148,163,184,0.12)",
        zerolinecolor="rgba(148,163,184,0.2)"
    )

    return fig



# ----------------------------------------------------------
# PREMIUM ANIMATED HERO SECTION
# ----------------------------------------------------------

st.markdown("""
<div class="hero-container">
    <div class="hero-label">✦ SMART DATA ANALYTICS PLATFORM</div>
    <h1>📊 Public Survey Response Analyser</h1>
    <p>
        Transform raw survey responses into meaningful insights.
        Clean your data, discover patterns, explore interactive
        visualisations and generate professional reports.
    </p>
</div>
""", unsafe_allow_html=True)


# ----------------------------------------------------------
# 2. DATASET UPLOAD
# ----------------------------------------------------------

st.sidebar.title("Dataset Settings")
st.sidebar.header("Upload Your Dataset")

uploaded_file = st.sidebar.file_uploader(
    "Choose CSV or Excel file",
    type=["csv", "xlsx"]
)


# ----------------------------------------------------------
# 3. READ DATASET
# ----------------------------------------------------------

if uploaded_file is not None:

    try:
        file_bytes = uploaded_file.getvalue()

        if uploaded_file.name.lower().endswith(".csv"):
            df = None

            for encoding in ["utf-8-sig", "utf-8", "latin1"]:
                try:
                    df = pd.read_csv(
                        io.BytesIO(file_bytes),
                        encoding=encoding
                    )
                    break
                except UnicodeDecodeError:
                    continue

            if df is None:
                st.error("Unable to read CSV file.")
                st.stop()

        else:
            excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
            selected_sheet = st.sidebar.selectbox(
                "Select Excel Sheet",
                excel_file.sheet_names
            )
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=selected_sheet
            )

        if df.empty:
            st.error("The uploaded dataset is empty.")
            st.stop()

        df.columns = df.columns.astype(str).str.strip()

        if df.columns.duplicated().any():
            st.error(
                "Duplicate column names detected. "
                "Please rename them before uploading."
            )
            st.stop()

        original_df = df.copy(deep=True)

        # --------------------------------------------------
        # 4. MISSING VALUE DETECTION AND CLASSIFICATION
        # --------------------------------------------------

        def detect_blank_values(data):
            result = data.copy(deep=True)
            text_columns = result.select_dtypes(
                include=["object", "string"]
            ).columns

            for column in text_columns:
                result[column] = result[column].replace(
                    r"^\s*$", pd.NA, regex=True
                )
            return result

        original_analysis_df = detect_blank_values(original_df)

        id_keywords = [
            "id", "identifier", "serial", "record_no", "sr_no"
        ]
        date_keywords = ["date", "time", "timestamp"]

        def classify_column(column):
            name = str(column).lower()
            words = re.split(r"[^a-z0-9]+", name)

            if any(word in id_keywords for word in words):
                return "Identifier"

            if any(word in date_keywords for word in words):
                return "Date / Time"

            series = (
                original_analysis_df[column]
                if column in original_analysis_df.columns
                else clean_df[column]
            )

            if pd.api.types.is_datetime64_any_dtype(series):
                return "Date / Time"

            if pd.api.types.is_numeric_dtype(series):
                rating_keywords = [
                    "rating", "score", "satisfaction", "marks"
                ]
                if any(word in name for word in rating_keywords):
                    return "Rating / Numerical"
                return "Numerical"

            non_empty = series.dropna().astype(str)

            if len(non_empty) == 0:
                return "Empty Column"

            average_length = non_empty.str.len().mean()
            unique_ratio = non_empty.nunique() / len(non_empty)

            if average_length >= 40 or (
                average_length >= 20 and unique_ratio > 0.5
            ):
                return "Text Response"

            return "Categorical"

        def detect_id_columns(columns):
            candidates = []
            known_id_names = [
                "id", "respondent_id", "response_id",
                "participant_id", "survey_id",
                "customer_id", "user_id", "record_id"
            ]

            for column in columns:
                normalized_name = re.sub(
                    r"[^a-z0-9]+", "_", str(column).lower()
                ).strip("_")

                if (
                    normalized_name in known_id_names
                    or normalized_name.endswith("_id")
                ):
                    candidates.append(column)

            return candidates

        id_columns = detect_id_columns(original_df.columns)

        # --------------------------------------------------
        # 5. SIDEBAR CLEANING SETTINGS
        # --------------------------------------------------

        st.sidebar.divider()
        st.sidebar.header("🧹 Advanced Data Cleaning")

        remove_duplicates = st.sidebar.checkbox(
            "Remove duplicate rows", value=False
        )
        remove_empty_columns = st.sidebar.checkbox(
            "Remove completely empty columns", value=False
        )
        remove_empty_rows = st.sidebar.checkbox(
            "Remove completely empty rows", value=False
        )
        trim_text = st.sidebar.checkbox(
            "Remove unnecessary text spaces", value=False
        )

        selected_id_column = None
        remove_duplicate_ids = False

        if id_columns:
            st.sidebar.divider()
            st.sidebar.subheader("🆔 Duplicate ID Detection")
            selected_id_column = st.sidebar.selectbox(
                "Select ID column", id_columns
            )
            remove_duplicate_ids = st.sidebar.checkbox(
                "Remove duplicate IDs (keep first)", value=False
            )
        else:
            st.sidebar.info("No ID column automatically detected.")

        missing_strategy = st.sidebar.selectbox(
            "Missing Value Treatment",
            [
                "Keep missing values",
                "Drop rows with missing values",
                "Fill numerical with median and text with Not provided"
            ]
        )

        # --------------------------------------------------
        # 6. ORIGINAL AND CLEANED DATA
        # --------------------------------------------------

        clean_df = original_analysis_df.copy(deep=True)

        def calculate_quality_score(data):
            if data.empty or len(data.columns) == 0:
                return 0.0

            total_cells = data.shape[0] * data.shape[1]
            if total_cells == 0:
                return 0.0

            filled_cells = data.notna().sum().sum()
            completeness = filled_cells / total_cells * 100
            duplicate_percentage = (
                data.duplicated().sum() / len(data)
            ) * 100
            uniqueness = 100 - duplicate_percentage

            return round(
                completeness * 0.70 + uniqueness * 0.30, 2
            )

        original_missing = int(original_analysis_df.isnull().sum().sum())
        original_duplicates = int(original_analysis_df.duplicated().sum())
        original_duplicate_ids = 0

        if selected_id_column is not None:
            valid_ids = original_analysis_df[
                selected_id_column
            ].dropna()
            original_duplicate_ids = int(valid_ids.duplicated().sum())

        original_quality = calculate_quality_score(
            original_analysis_df
        )

        # Apply cleaning choices
        if remove_empty_columns:
            clean_df = clean_df.dropna(axis=1, how="all")

        if remove_empty_rows:
            clean_df = clean_df.dropna(axis=0, how="all")

        if remove_duplicates:
            clean_df = clean_df.drop_duplicates()

        if trim_text:
            text_columns = clean_df.select_dtypes(
                include=["object", "string"]
            ).columns
            for column in text_columns:
                clean_df[column] = clean_df[column].apply(
                    lambda value: (
                        value.strip() if isinstance(value, str) else value
                    )
                )

        if missing_strategy == "Drop rows with missing values":
            clean_df = clean_df.dropna(axis=0)

        elif missing_strategy == (
            "Fill numerical with median and text with Not provided"
        ):
            for column in clean_df.columns:
                if pd.api.types.is_numeric_dtype(clean_df[column]):
                    median_value = clean_df[column].median()
                    if pd.notna(median_value):
                        clean_df[column] = clean_df[column].fillna(
                            median_value
                        )
                else:
                    clean_df[column] = clean_df[column].fillna(
                        "Not provided"
                    )

        if (
            remove_duplicate_ids
            and selected_id_column in clean_df.columns
        ):
            valid_id_mask = clean_df[selected_id_column].notna()
            duplicate_id_mask = (
                clean_df.loc[
                    valid_id_mask, selected_id_column
                ].duplicated(keep="first")
            )
            clean_df = clean_df.drop(
                index=duplicate_id_mask[duplicate_id_mask].index
            )

        cleaned_missing = int(clean_df.isnull().sum().sum())
        cleaned_duplicates = int(clean_df.duplicated().sum())
        cleaned_duplicate_ids = 0

        if (
            selected_id_column is not None
            and selected_id_column in clean_df.columns
        ):
            valid_ids = clean_df[selected_id_column].dropna()
            cleaned_duplicate_ids = int(valid_ids.duplicated().sum())

        cleaned_quality = calculate_quality_score(clean_df)

        # --------------------------------------------------
        # 7. DASHBOARD OVERVIEW AND TABS
        # --------------------------------------------------

        st.success("Your dataset is ready for analysis!")
        st.markdown("## 📌 Dataset Overview")
        st.caption("A quick snapshot of your uploaded survey dataset.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Responses", len(clean_df))
        col2.metric("Total Columns", len(clean_df.columns))
        col3.metric("Missing Values", cleaned_missing)
        col4.metric("Duplicate Rows", cleaned_duplicates)

        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8,tab9 = st.tabs([
            "📄 Data Preview",
            "🧹 Data Quality",
            "📊 Numerical Analysis",
            "📈 Response Analysis",
            "✨ Advanced Cleaning",
            "👥 Group Analysis",
            "🧠 Smart Insights",
            "📑 Professional Reports",
            "🚀 Interactive Dashboard",

        ])

        # --------------------------------------------------
        # TAB 1: DATA PREVIEW
        # --------------------------------------------------

        with tab1:
            st.subheader("Original Survey Dataset")
            st.dataframe(
                original_df, use_container_width=True, height=300
            )
            st.write("Original Dimensions:", original_df.shape)

            st.divider()
            st.subheader("Cleaned Survey Dataset")
            st.dataframe(
                clean_df, use_container_width=True, height=300
            )
            st.write("Cleaned Dimensions:", clean_df.shape)

        # --------------------------------------------------
        # TAB 2: DATA QUALITY
        # --------------------------------------------------

        with tab2:
            st.subheader("Data Quality Report")

            missing_report = pd.DataFrame({
                "Column": clean_df.columns,
                "Missing Values": [
                    clean_df[col].isnull().sum()
                    for col in clean_df.columns
                ],
                "Missing Percentage": [
                    round(clean_df[col].isnull().mean() * 100, 2)
                    for col in clean_df.columns
                ]
            })

            st.write("Missing Value Analysis")
            st.dataframe(
                missing_report,
                use_container_width=True,
                hide_index=True
            )

            st.divider()
            st.subheader("Duplicate Record Analysis")
            col1, col2 = st.columns(2)
            col1.metric("Duplicate Rows", cleaned_duplicates)
            col2.metric("Duplicate IDs", cleaned_duplicate_ids)

            if selected_id_column:
                st.caption(f"ID column used: {selected_id_column}")
            else:
                st.info("No suitable ID column detected.")

        # --------------------------------------------------
        # TAB 3: ADVANCED STATISTICAL ANALYSIS
        # --------------------------------------------------

        with tab3:
            st.header("📊 Advanced Statistical Analysis")
            st.write(
                "Perform detailed statistical analysis "
                "of your survey dataset."
            )

            numeric_columns = clean_df.select_dtypes(
                include="number"
            ).columns.tolist()

            numeric_columns = [
                col for col in numeric_columns
                if classify_column(col) != "Identifier"
            ]

            if numeric_columns:
                selected_numeric = st.selectbox(
                    "Select a numerical column",
                    numeric_columns,
                    key="advanced_numeric_column"
                )
                numeric_data = clean_df[selected_numeric].dropna()

                if len(numeric_data) > 0:
                    mean_value = numeric_data.mean()
                    median_value = numeric_data.median()
                    mode_values = numeric_data.mode().tolist()
                    std_value = numeric_data.std()
                    variance_value = numeric_data.var()
                    minimum_value = numeric_data.min()
                    maximum_value = numeric_data.max()
                    range_value = maximum_value - minimum_value

                    st.subheader("Statistical Summary")
                    cols = st.columns(5)
                    cols[0].metric("Mean", f"{mean_value:.2f}")
                    cols[1].metric("Median", f"{median_value:.2f}")
                    cols[2].metric(
                        "Mode",
                        ", ".join(map(str, mode_values[:3]))
                    )
                    cols[3].metric(
                        "Std Deviation", f"{std_value:.2f}"
                    )
                    cols[4].metric(
                        "Variance", f"{variance_value:.2f}"
                    )

                    statistics = pd.DataFrame({
                        "Statistical Measure": [
                            "Count", "Mean", "Median", "Mode",
                            "Standard Deviation", "Variance",
                            "Minimum", "Maximum", "Range"
                        ],
                        "Value": [
                            len(numeric_data),
                            round(mean_value, 2),
                            round(median_value, 2),
                            ", ".join(map(str, mode_values)),
                            round(std_value, 2),
                            round(variance_value, 2),
                            minimum_value,
                            maximum_value,
                            range_value
                        ]
                    })

                    st.divider()
                    st.subheader("Detailed Statistical Report")
                    st.dataframe(
                        statistics,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.divider()
                    st.subheader("📈 Distribution Analysis")
                    col1, col2 = st.columns(2)

                    with col1:
                        histogram = px.histogram(
                            x=numeric_data,
                            title=f"Distribution of {selected_numeric}",
                            marginal="box",
                            labels={"x": selected_numeric}
                        )
                        st.plotly_chart(
                            histogram, use_container_width=True
                        )

                    with col2:
                        box_plot = px.box(
                            y=numeric_data,
                            title=f"Box Plot of {selected_numeric}",
                            labels={"y": selected_numeric}
                        )
                        st.plotly_chart(
                            box_plot, use_container_width=True
                        )

                    st.divider()
                    st.subheader("🔍 Outlier Detection")
                    q1 = numeric_data.quantile(0.25)
                    q3 = numeric_data.quantile(0.75)
                    iqr = q3 - q1
                    lower_limit = q1 - (1.5 * iqr)
                    upper_limit = q3 + (1.5 * iqr)

                    outliers = numeric_data[
                        (numeric_data < lower_limit)
                        | (numeric_data > upper_limit)
                    ]

                    cols = st.columns(3)
                    cols[0].metric("Total Outliers", len(outliers))
                    cols[1].metric("Lower Limit", f"{lower_limit:.2f}")
                    cols[2].metric("Upper Limit", f"{upper_limit:.2f}")

                    if len(outliers) > 0:
                        st.warning("Unusual values have been detected.")
                        outlier_report = pd.DataFrame({
                            "Original Row Index": outliers.index,
                            "Outlier Value": outliers.values
                        })
                        st.dataframe(
                            outlier_report,
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        st.success(
                            "No outliers detected using the IQR method."
                        )
                else:
                    st.warning(
                        "The selected column contains no valid values."
                    )
            else:
                st.info("No suitable numerical columns detected.")

            st.divider()
            st.subheader("🔗 Correlation Analysis")

            if len(numeric_columns) >= 2:
                correlation_data = clean_df[numeric_columns].corr()
                correlation_fig = px.imshow(
                    correlation_data,
                    text_auto=".2f",
                    aspect="auto",
                    color_continuous_scale="RdBu",
                    zmin=-1,
                    zmax=1,
                    title="Numerical Correlation Heatmap"
                )
                st.plotly_chart(
                    correlation_fig, use_container_width=True
                )
                st.dataframe(
                    correlation_data.round(2),
                    use_container_width=True
                )
                st.caption(
                    "Correlation ranges from -1 to +1. "
                    "Correlation does not prove causation."
                )
            else:
                st.info(
                    "Correlation analysis requires at least "
                    "two suitable numerical columns."
                )

        # --------------------------------------------------
        # TAB 4: RESPONSE ANALYSIS
        # --------------------------------------------------

        with tab4:
            st.subheader("Survey Response Distribution")

            categorical_columns = [
                col for col in clean_df.columns
                if classify_column(col) == "Categorical"
            ]

            if categorical_columns:
                selected_category = st.selectbox(
                    "Select a response column",
                    categorical_columns,
                    key="response_category"
                )

                response_counts = (
                    clean_df[selected_category]
                    .dropna()
                    .astype(str)
                    .value_counts()
                    .head(10)
                    .reset_index()
                )
                response_counts.columns = ["Response", "Count"]

                fig = px.bar(
                    response_counts,
                    x="Response",
                    y="Count",
                    title=f"Response Distribution: {selected_category}",
                    text="Count"
                )
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(
                    response_counts,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No categorical response columns detected.")

        # --------------------------------------------------
        # TAB 5: ADVANCED DATA CLEANING
        # --------------------------------------------------

        with tab5:
            st.header("✨ Advanced Data Cleaning Dashboard")
            st.write(
                "Compare your original dataset with "
                "the cleaned dataset."
            )

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Before Cleaning")
                st.metric("Data Quality Score", f"{original_quality:.2f}%")
                st.metric("Missing Values", original_missing)
                st.metric("Duplicate Rows", original_duplicates)
                st.metric("Duplicate IDs", original_duplicate_ids)
                st.metric("Total Rows", len(original_df))

            with col2:
                st.subheader("After Cleaning")
                quality_delta = round(
                    cleaned_quality - original_quality, 2
                )
                st.metric(
                    "Data Quality Score",
                    f"{cleaned_quality:.2f}%",
                    delta=(
                        f"{quality_delta:+.2f}%"
                        if quality_delta != 0 else None
                    )
                )
                st.metric(
                    "Missing Values",
                    cleaned_missing,
                    delta=(
                        original_missing - cleaned_missing
                        if original_missing != cleaned_missing else None
                    )
                )
                st.metric(
                    "Duplicate Rows",
                    cleaned_duplicates,
                    delta=(
                        original_duplicates - cleaned_duplicates
                        if original_duplicates != cleaned_duplicates else None
                    )
                )
                st.metric(
                    "Duplicate IDs",
                    cleaned_duplicate_ids,
                    delta=(
                        original_duplicate_ids - cleaned_duplicate_ids
                        if original_duplicate_ids != cleaned_duplicate_ids
                        else None
                    )
                )
                st.metric("Total Rows", len(clean_df))

            st.divider()
            st.subheader("Cleaning Operations Report")

            operations = pd.DataFrame({
                "Cleaning Operation": [
                    "Blank Value Detection",
                    "Duplicate Row Removal",
                    "Duplicate ID Removal",
                    "Empty Column Removal",
                    "Empty Row Removal",
                    "Text Space Standardisation",
                    "Missing Value Treatment"
                ],
                "Status": [
                    "Completed",
                    "Applied" if remove_duplicates else "Not Applied",
                    "Applied" if remove_duplicate_ids else "Not Applied",
                    "Applied" if remove_empty_columns else "Not Applied",
                    "Applied" if remove_empty_rows else "Not Applied",
                    "Applied" if trim_text else "Not Applied",
                    missing_strategy
                ]
            })

            st.dataframe(
                operations, use_container_width=True, hide_index=True
            )

            st.divider()
            st.subheader("Data Quality Comparison")

            quality_comparison = pd.DataFrame({
                "Stage": ["Before Cleaning", "After Cleaning"],
                "Quality Score": [original_quality, cleaned_quality]
            })

            quality_fig = px.bar(
                quality_comparison,
                x="Stage",
                y="Quality Score",
                text="Quality Score",
                range_y=[0, 100],
                title="Data Quality Before and After Cleaning"
            )
            st.plotly_chart(quality_fig, use_container_width=True)

            st.caption(
                "Quality score formula: 70% completeness + "
                "30% non-duplicate row percentage. This is an "
                "indicative project metric, not an industry-standard "
                "data quality certification."
            )

            st.divider()
            st.subheader("Cleaning Summary")
            col1, col2 = st.columns(2)

            with col1:
                st.write(f"Original rows: {len(original_df)}")
                st.write(f"Original columns: {len(original_df.columns)}")

            with col2:
                st.write(f"Cleaned rows: {len(clean_df)}")
                st.write(f"Cleaned columns: {len(clean_df.columns)}")

            if cleaned_quality > original_quality:
                st.success(
                    "The calculated data quality score improved."
                )
            elif cleaned_quality == original_quality:
                st.info(
                    "The calculated data quality score remained unchanged."
                )
            else:
                st.warning(
                    "The calculated data quality score decreased. "
                    "Review the selected cleaning operations."
                )

        # --------------------------------------------------
        # TAB 6: GROUP-WISE & SATISFACTION ANALYSIS
        # --------------------------------------------------

        with tab6:
            st.header("👥 Group-wise & Satisfaction Analysis")
            st.write(
                "Compare survey responses across different "
                "demographic groups and analyse their ratings."
            )

            categorical_columns = [
                col for col in clean_df.columns
                if classify_column(col) == "Categorical"
            ]

            analysis_numeric_columns = [
                col for col in clean_df.select_dtypes(
                    include="number"
                ).columns
                if classify_column(col) != "Identifier"
            ]

            if not categorical_columns:
                st.warning(
                    "No suitable categorical or demographic "
                    "columns were detected."
                )
            elif not analysis_numeric_columns:
                st.warning("No suitable numerical columns were detected.")
            else:
                st.subheader("Select Analysis Columns")
                col1, col2 = st.columns(2)

                with col1:
                    selected_group = st.selectbox(
                        "Select demographic / group column",
                        categorical_columns,
                        key="step5_group_column"
                    )

                rating_columns = [
                    col for col in analysis_numeric_columns
                    if classify_column(col) == "Rating / Numerical"
                ]
                default_numeric = (
                    rating_columns[0]
                    if rating_columns
                    else analysis_numeric_columns[0]
                )

                with col2:
                    selected_measure = st.selectbox(
                        "Select numerical rating / measure",
                        analysis_numeric_columns,
                        index=analysis_numeric_columns.index(default_numeric),
                        key="step5_measure_column"
                    )

                analysis_data = clean_df[
                    [selected_group, selected_measure]
                ].dropna().copy()
                analysis_data[selected_group] = (
                    analysis_data[selected_group].astype(str)
                )

                if analysis_data.empty:
                    st.warning(
                        "No valid responses are available for "
                        "the selected columns."
                    )
                else:
                    st.divider()
                    st.subheader("📊 Group-wise Statistical Comparison")

                    group_summary = (
                        analysis_data
                        .groupby(selected_group)[selected_measure]
                        .agg(
                            Responses="count",
                            Average="mean",
                            Median="median",
                            Minimum="min",
                            Maximum="max"
                        )
                        .reset_index()
                    )

                    group_summary[
                        ["Average", "Median", "Minimum", "Maximum"]
                    ] = group_summary[
                        ["Average", "Median", "Minimum", "Maximum"]
                    ].round(2)

                    group_summary = group_summary.sort_values(
                        by="Average", ascending=False
                    )

                    st.dataframe(
                        group_summary,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.divider()
                    st.subheader("📌 Group Comparison Summary")

                    highest_group = group_summary.iloc[0]
                    lowest_group = group_summary.iloc[-1]

                    cols = st.columns(3)
                    cols[0].metric("Total Groups", len(group_summary))
                    cols[1].metric(
                        "Highest Average",
                        f"{highest_group['Average']:.2f}"
                    )
                    cols[2].metric(
                        "Lowest Average",
                        f"{lowest_group['Average']:.2f}"
                    )

                    st.success(
                        f"Highest average: {highest_group[selected_group]} "
                        f"({highest_group['Average']:.2f})"
                    )
                    st.info(
                        f"Lowest average: {lowest_group[selected_group]} "
                        f"({lowest_group['Average']:.2f})"
                    )
                    st.write(
                        f"Total valid responses analysed: {len(analysis_data)}"
                    )

                    st.divider()
                    st.subheader("📈 Average Rating by Group")

                    average_chart = px.bar(
                        group_summary,
                        x=selected_group,
                        y="Average",
                        text="Average",
                        title=f"Average {selected_measure} by {selected_group}",
                        hover_data=["Responses", "Median"]
                    )
                    st.plotly_chart(
                        average_chart, use_container_width=True
                    )

                    st.divider()
                    st.subheader("📊 Number of Responses by Group")

                    response_chart = px.bar(
                        group_summary,
                        x=selected_group,
                        y="Responses",
                        text="Responses",
                        title=f"Response Count by {selected_group}"
                    )
                    st.plotly_chart(
                        response_chart, use_container_width=True
                    )

                    st.divider()
                    st.subheader(
                        "⭐ Threshold-based Positive Response Analysis"
                    )
                    st.write(
                        "Choose a threshold to calculate the percentage "
                        "of responses whose numerical value is equal to "
                        "or greater than that threshold."
                    )

                    minimum_value = float(
                        analysis_data[selected_measure].min()
                    )
                    maximum_value = float(
                        analysis_data[selected_measure].max()
                    )
                    median_value = float(
                        analysis_data[selected_measure].median()
                    )

                    if minimum_value == maximum_value:
                        threshold = minimum_value
                        st.info(
                            f"All values are {minimum_value}. "
                            "The threshold is fixed at this value."
                        )
                    else:
                        threshold = st.number_input(
                            f"Select threshold for {selected_measure}",
                            min_value=minimum_value,
                            max_value=maximum_value,
                            value=median_value,
                            key="step5_threshold"
                        )

                    analysis_data["Positive Response"] = (
                        analysis_data[selected_measure] >= threshold
                    )

                    satisfaction_summary = (
                        analysis_data
                        .groupby(selected_group)
                        .agg(
                            Total_Responses=("Positive Response", "count"),
                            Positive_Responses=("Positive Response", "sum")
                        )
                        .reset_index()
                    )

                    satisfaction_summary[
                        "Positive Response Percentage"
                    ] = (
                        satisfaction_summary["Positive_Responses"]
                        / satisfaction_summary["Total_Responses"]
                        * 100
                    ).round(2)

                    satisfaction_summary = (
                        satisfaction_summary.sort_values(
                            by="Positive Response Percentage",
                            ascending=False
                        )
                    )

                    overall_positive = int(
                        analysis_data["Positive Response"].sum()
                    )
                    overall_percentage = (
                        overall_positive / len(analysis_data) * 100
                    )

                    col1, col2 = st.columns(2)
                    col1.metric(
                        "Overall Positive Responses", overall_positive
                    )
                    col2.metric(
                        "Overall Positive Response Rate",
                        f"{overall_percentage:.2f}%"
                    )

                    st.dataframe(
                        satisfaction_summary,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.divider()
                    st.subheader(
                        "📊 Positive Response Percentage by Group"
                    )

                    satisfaction_chart = px.bar(
                        satisfaction_summary,
                        x=selected_group,
                        y="Positive Response Percentage",
                        text="Positive Response Percentage",
                        title=f"Positive Response Rate by {selected_group}",
                        range_y=[0, 100]
                    )
                    st.plotly_chart(
                        satisfaction_chart, use_container_width=True
                    )

                    st.divider()
                    st.subheader("💡 Automatically Generated Insights")

                    highest_positive = satisfaction_summary.iloc[0]
                    lowest_positive = satisfaction_summary.iloc[-1]

                    st.success(
                        f"{highest_positive[selected_group]} has the "
                        f"highest positive response rate of "
                        f"{highest_positive['Positive Response Percentage']:.2f}%."
                    )
                    st.warning(
                        f"{lowest_positive[selected_group]} has the "
                        f"lowest positive response rate of "
                        f"{lowest_positive['Positive Response Percentage']:.2f}%."
                    )
                    st.caption(
                        "Positive response means a value greater than "
                        "or equal to the selected threshold. It is not "
                        "automatically a validated measure of satisfaction."
                    )

                    st.divider()
                    st.subheader("📥 Download Group Analysis Report")

                    report_buffer = io.BytesIO()
                    with pd.ExcelWriter(
                        report_buffer, engine="openpyxl"
                    ) as writer:
                        group_summary.to_excel(
                            writer,
                            index=False,
                            sheet_name="Group Comparison"
                        )
                        satisfaction_summary.to_excel(
                            writer,
                            index=False,
                            sheet_name="Positive Responses"
                        )

                    st.download_button(
                        label="Download Group Analysis Excel",
                        data=report_buffer.getvalue(),
                        file_name="survey_group_analysis.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key="step5_download_report"
                    )

        # --------------------------------------------------
        # TAB 7: SMART INSIGHTS & EXECUTIVE SUMMARY
        # --------------------------------------------------

        smart_insights = []
        smart_recommendations = []
        executive_report = []

        with tab7:
            st.header("🧠 Smart Survey Insights")
            st.write(
                "Automatically understand your survey data through "
                "important findings, response patterns and recommendations."
            )
            st.divider()

            smart_numeric_columns = [
                col for col in clean_df.select_dtypes(
                    include="number"
                ).columns
                if classify_column(col) != "Identifier"
            ]

            smart_categorical_columns = [
                col for col in clean_df.columns
                if classify_column(col) == "Categorical"
            ]

            st.subheader("Choose Data for Smart Analysis")
            smart_measure = None
            smart_category = None
            col1, col2 = st.columns(2)

            with col1:
                if smart_numeric_columns:
                    smart_measure = st.selectbox(
                        "Select numerical measure",
                        smart_numeric_columns,
                        key="step6_measure"
                    )
                else:
                    st.info("No suitable numerical column found.")

            with col2:
                if smart_categorical_columns:
                    smart_category = st.selectbox(
                        "Select response category",
                        smart_categorical_columns,
                        key="step6_category"
                    )
                else:
                    st.info("No suitable categorical column found.")

            executive_report.extend([
                "PUBLIC SURVEY RESPONSE ANALYSER",
                "SMART INSIGHTS - EXECUTIVE SUMMARY",
                "=" * 55,
                f"Total cleaned responses: {len(clean_df)}",
                f"Total columns: {len(clean_df.columns)}",
                f"Missing values: {cleaned_missing}",
                f"Duplicate rows: {cleaned_duplicates}",
                f"Data quality score: {cleaned_quality}%"
            ])

            st.divider()
            st.subheader("📋 Overall Survey Overview")
            cols = st.columns(4)
            cols[0].metric("Total Responses", len(clean_df))
            cols[1].metric("Missing Values", cleaned_missing)
            cols[2].metric("Duplicate Rows", cleaned_duplicates)
            cols[3].metric("Data Quality Score", f"{cleaned_quality:.2f}%")

            if cleaned_missing == 0:
                quality_message = (
                    "The cleaned dataset contains no missing values."
                )
                smart_recommendations.append(
                    "Continue checking incoming survey data to maintain completeness."
                )
            else:
                quality_message = (
                    f"The cleaned dataset contains {cleaned_missing} missing values."
                )
                smart_recommendations.append(
                    "Review columns with missing values and decide whether "
                    "to collect, fill or exclude those responses."
                )

            smart_insights.append(quality_message)
            executive_report.extend(["", "DATA QUALITY", quality_message])

            if cleaned_duplicates > 0:
                duplicate_message = (
                    f"{cleaned_duplicates} duplicate rows remain "
                    "in the cleaned dataset."
                )
                smart_insights.append(duplicate_message)
                smart_recommendations.append(
                    "Review duplicate records to check whether they "
                    "represent repeated submissions."
                )
            else:
                smart_insights.append(
                    "No duplicate rows were found in the cleaned dataset."
                )

            if smart_measure is not None:
                numeric_data = clean_df[smart_measure].dropna()
                st.divider()
                st.subheader("📊 Numerical Findings")

                if not numeric_data.empty:
                    average = float(numeric_data.mean())
                    median = float(numeric_data.median())
                    minimum = float(numeric_data.min())
                    maximum = float(numeric_data.max())
                    modes = numeric_data.mode().tolist()

                    cols = st.columns(4)
                    cols[0].metric("Average", f"{average:.2f}")
                    cols[1].metric("Median", f"{median:.2f}")
                    cols[2].metric("Minimum", f"{minimum:.2f}")
                    cols[3].metric("Maximum", f"{maximum:.2f}")

                    numerical_message = (
                        f"The average {smart_measure} is {average:.2f}, "
                        f"while the median is {median:.2f}."
                    )
                    smart_insights.append(numerical_message)

                    executive_report.extend([
                        "",
                        "NUMERICAL FINDINGS",
                        f"Measure analysed: {smart_measure}",
                        f"Average: {average:.2f}",
                        f"Median: {median:.2f}",
                        f"Minimum: {minimum:.2f}",
                        f"Maximum: {maximum:.2f}",
                        f"Most common value(s): {', '.join(map(str, modes[:5]))}"
                    ])

                    if average > median:
                        smart_insights.append(
                            "The average is higher than the median, which "
                            "may indicate that some higher values are pulling "
                            "the average upward."
                        )
                    elif average < median:
                        smart_insights.append(
                            "The average is lower than the median, which "
                            "may indicate that some lower values are pulling "
                            "the average downward."
                        )
                    else:
                        smart_insights.append(
                            "The average and median are equal."
                        )

                    if numeric_data.nunique() == 1:
                        smart_recommendations.append(
                            f"The {smart_measure} column contains only one "
                            "distinct value. Check whether the question or "
                            "data collection process produced enough variation."
                        )

                    if len(numeric_data) < len(clean_df):
                        smart_recommendations.append(
                            f"Some responses are missing from {smart_measure}. "
                            "Review those records before making final conclusions."
                        )

                    numerical_fig = px.histogram(
                        x=numeric_data,
                        title=f"Distribution of {smart_measure}",
                        labels={"x": smart_measure}
                    )
                    st.plotly_chart(
                        numerical_fig, use_container_width=True
                    )
                else:
                    st.warning(
                        "The selected numerical column has no valid values."
                    )

            if smart_category is not None:
                category_data = (
                    clean_df[smart_category].dropna().astype(str)
                )
                st.divider()
                st.subheader("📈 Most Common Survey Responses")

                if not category_data.empty:
                    category_counts = (
                        category_data.value_counts()
                        .head(10)
                        .reset_index()
                    )
                    category_counts.columns = ["Response", "Count"]
                    category_counts["Percentage"] = (
                        category_counts["Count"] / len(category_data) * 100
                    ).round(2)

                    st.dataframe(
                        category_counts,
                        use_container_width=True,
                        hide_index=True
                    )

                    most_common = category_counts.iloc[0]
                    common_message = (
                        f"The most common response for {smart_category} "
                        f"is '{most_common['Response']}', selected by "
                        f"{int(most_common['Count'])} respondents "
                        f"({most_common['Percentage']:.2f}%)."
                    )
                    smart_insights.append(common_message)

                    executive_report.extend([
                        "",
                        "MOST COMMON SURVEY RESPONSES",
                        f"Category analysed: {smart_category}"
                    ])

                    for _, row in category_counts.iterrows():
                        executive_report.append(
                            f"{row['Response']}: {int(row['Count'])} "
                            f"responses ({row['Percentage']:.2f}%)"
                        )

                    category_fig = px.bar(
                        category_counts,
                        x="Response",
                        y="Count",
                        text="Count",
                        title=f"Most Common Responses: {smart_category}"
                    )
                    st.plotly_chart(
                        category_fig, use_container_width=True
                    )

                    if len(category_counts) == 1:
                        smart_recommendations.append(
                            f"Only one response category was found for "
                            f"{smart_category}. Check whether the dataset "
                            "contains enough response variety."
                        )
                else:
                    st.warning("No valid categorical responses are available.")

            if smart_category is not None and smart_measure is not None:
                st.divider()
                st.subheader("👥 Demographic Group Insights")

                group_data = clean_df[
                    [smart_category, smart_measure]
                ].dropna().copy()

                if not group_data.empty:
                    group_data[smart_category] = (
                        group_data[smart_category].astype(str)
                    )

                    smart_group_summary = (
                        group_data
                        .groupby(smart_category)[smart_measure]
                        .agg(Responses="count", Average="mean")
                        .reset_index()
                    )
                    smart_group_summary["Average"] = (
                        smart_group_summary["Average"].round(2)
                    )
                    smart_group_summary = smart_group_summary.sort_values(
                        by="Average", ascending=False
                    )

                    st.dataframe(
                        smart_group_summary,
                        use_container_width=True,
                        hide_index=True
                    )

                    best_group = smart_group_summary.iloc[0]
                    lowest_group = smart_group_summary.iloc[-1]

                    highest_message = (
                        f"{best_group[smart_category]} has the highest "
                        f"average {smart_measure}: {best_group['Average']:.2f}."
                    )
                    lowest_message = (
                        f"{lowest_group[smart_category]} has the lowest "
                        f"average {smart_measure}: {lowest_group['Average']:.2f}."
                    )
                    smart_insights.extend([
                        highest_message,
                        lowest_message
                    ])
                    executive_report.extend([
                        "",
                        "DEMOGRAPHIC GROUP COMPARISON",
                        highest_message,
                        lowest_message
                    ])

                    group_fig = px.bar(
                        smart_group_summary,
                        x=smart_category,
                        y="Average",
                        text="Average",
                        title=f"Average {smart_measure} by {smart_category}"
                    )
                    st.plotly_chart(
                        group_fig, use_container_width=True
                    )

                    if (
                        best_group[smart_category]
                        != lowest_group[smart_category]
                    ):
                        difference = (
                            best_group["Average"] - lowest_group["Average"]
                        )
                        smart_insights.append(
                            f"The difference between the highest and lowest "
                            f"group averages is {difference:.2f}."
                        )
                        smart_recommendations.append(
                            f"Investigate why the average {smart_measure} "
                            f"differs between {best_group[smart_category]} "
                            f"and {lowest_group[smart_category]}. Collect "
                            "additional feedback before deciding on action."
                        )
                    else:
                        smart_insights.append(
                            "Only one group was available for this comparison."
                        )
                else:
                    st.info(
                        "There are no valid records for the selected "
                        "group and numerical measure."
                    )

            st.divider()
            st.subheader("💡 Automatically Generated Key Findings")

            for index, insight in enumerate(smart_insights, start=1):
                st.write(f"**{index}.** {insight}")
                executive_report.append(f"{index}. {insight}")

            st.divider()
            st.subheader("🎯 Actionable Recommendations")

            if not smart_recommendations:
                smart_recommendations.append(
                    "Continue collecting survey responses and monitor "
                    "changes in the results over time."
                )

            smart_recommendations.append(
                "Compare results across different groups before making "
                "major public or organisational decisions."
            )
            smart_recommendations.append(
                "Use the survey findings to identify areas that need "
                "further investigation or improvement."
            )

            unique_recommendations = list(
                dict.fromkeys(smart_recommendations)
            )

            for index, recommendation in enumerate(
                unique_recommendations, start=1
            ):
                st.write(f"**{index}.** {recommendation}")
                executive_report.append(
                    f"Recommendation {index}: {recommendation}"
                )

            st.divider()
            st.header("📑 Executive Summary")
            st.write(
                "This is an automatically generated overview "
                "of the selected survey dataset."
            )

            if smart_measure is None and smart_category is None:
                st.warning(
                    "Select at least one suitable analysis column "
                    "to generate detailed findings."
                )

            summary_text = (
                f"The survey dataset contains {len(clean_df)} cleaned "
                f"responses and {len(clean_df.columns)} columns. Its "
                f"calculated data quality score is {cleaned_quality:.2f}%. "
            )

            if smart_measure is not None:
                summary_text += (
                    f"The selected numerical measure is {smart_measure}. "
                )
            if smart_category is not None:
                summary_text += (
                    f"The selected response category is {smart_category}. "
                )

            summary_text += (
                "The findings and recommendations above are based on "
                "the available cleaned dataset."
            )

            st.info(summary_text)
            executive_report.extend([
                "",
                "EXECUTIVE SUMMARY",
                summary_text,
                "",
                "IMPORTANT LIMITATIONS",
                "These findings describe the uploaded dataset and do not "
                "establish cause-and-effect relationships.",
                "Automatically generated recommendations should be reviewed.",
                "The calculated data quality score is an indicative "
                "project metric, not an industry-standard certification."
            ])

            st.divider()
            st.subheader("📥 Download Executive Summary")

            st.download_button(
                label="Download Executive Summary (TXT)",
                data="\n".join(executive_report),
                file_name="survey_executive_summary.txt",
                mime="text/plain",
                key="step6_download_summary"
            )

            st.caption(
                "The report is generated using Python and the uploaded "
                "dataset. No external AI API is required."
            )

        # --------------------------------------------------
        # TAB 8: PROFESSIONAL REPORT GENERATION - STEP 7
        # --------------------------------------------------

        with tab8:
            st.header("📑 Professional Survey Reports")
            st.write(
                "Generate professional PDF and Excel reports "
                "from your uploaded survey dataset."
            )
            st.divider()

            report_date = datetime.now().strftime(
                "%d %B %Y, %I:%M %p"
            )
            report_filename = uploaded_file.name

            st.subheader("📋 Report Overview")
            cols = st.columns(4)
            cols[0].metric("Total Responses", len(clean_df))
            cols[1].metric("Total Columns", len(clean_df.columns))
            cols[2].metric("Data Quality", f"{cleaned_quality:.2f}%")
            cols[3].metric("Missing Values", cleaned_missing)

            st.write(f"**Dataset:** {report_filename}")
            st.write(f"**Report Generated:** {report_date}")
            st.divider()

            # Prepare report tables
            report_quality = pd.DataFrame({
                "Column": clean_df.columns,
                "Data Type": [
                    str(clean_df[col].dtype) for col in clean_df.columns
                ],
                "Missing Values": [
                    int(clean_df[col].isna().sum())
                    for col in clean_df.columns
                ],
                "Missing Percentage": [
                    round(clean_df[col].isna().mean() * 100, 2)
                    for col in clean_df.columns
                ],
                "Unique Values": [
                    int(clean_df[col].nunique(dropna=True))
                    for col in clean_df.columns
                ]
            })

            report_numeric_columns = [
                col for col in clean_df.select_dtypes(
                    include="number"
                ).columns
                if classify_column(col) != "Identifier"
            ]

            if report_numeric_columns:
                report_numerical = (
                    clean_df[report_numeric_columns]
                    .describe()
                    .round(2)
                    .reset_index()
                    .rename(columns={"index": "Statistic"})
                )
            else:
                report_numerical = pd.DataFrame({
                    "Information": ["No suitable numerical columns found."]
                })

            response_tables = []
            for column in clean_df.columns:
                if classify_column(column) == "Categorical":
                    counts = (
                        clean_df[column]
                        .dropna()
                        .astype(str)
                        .value_counts()
                        .reset_index()
                    )
                    if not counts.empty:
                        counts.columns = ["Response", "Count"]
                        counts["Column"] = column
                        counts["Percentage"] = (
                            counts["Count"] / counts["Count"].sum() * 100
                        ).round(2)
                        response_tables.append(
                            counts[
                                ["Column", "Response", "Count", "Percentage"]
                            ]
                        )

            if response_tables:
                report_responses = pd.concat(
                    response_tables, ignore_index=True
                )
            else:
                report_responses = pd.DataFrame({
                    "Information": ["No categorical response columns found."]
                })

            # Build findings and recommendations for professional report
            report_findings = [
                f"The cleaned dataset contains {len(clean_df)} responses.",
                f"The dataset contains {len(clean_df.columns)} columns.",
                f"The calculated data quality score is {cleaned_quality:.2f}%.",
                f"The cleaned dataset contains {cleaned_missing} missing values.",
                f"The cleaned dataset contains {cleaned_duplicates} duplicate rows."
            ]

            report_recommendations = []

            if cleaned_missing > 0:
                report_recommendations.append(
                    "Review missing values and decide whether additional "
                    "responses should be collected or missing data should "
                    "be treated."
                )
            else:
                report_recommendations.append(
                    "Continue monitoring incoming survey responses "
                    "to maintain data completeness."
                )

            if cleaned_duplicates > 0:
                report_recommendations.append(
                    "Investigate remaining duplicate records before "
                    "using the dataset for final analysis."
                )
            else:
                report_recommendations.append(
                    "Continue checking duplicate submissions during "
                    "future data collection."
                )

            for column in report_numeric_columns:
                values = clean_df[column].dropna()
                if not values.empty:
                    report_findings.append(
                        f"{column}: average = {values.mean():.2f}, "
                        f"median = {values.median():.2f}, "
                        f"minimum = {values.min()}, maximum = {values.max()}."
                    )

            report_recommendations.extend([
                "Compare survey responses across relevant groups before "
                "making major decisions.",
                "Use the findings to identify areas requiring further "
                "investigation or improvement."
            ])

            # Group analysis for report
            report_group_summary = pd.DataFrame()
            report_categorical_columns = [
                col for col in clean_df.columns
                if classify_column(col) == "Categorical"
            ]

            if report_categorical_columns and report_numeric_columns:
                report_group_column = report_categorical_columns[0]
                report_measure_column = report_numeric_columns[0]

                group_data = clean_df[
                    [report_group_column, report_measure_column]
                ].dropna().copy()

                if not group_data.empty:
                    group_data[report_group_column] = (
                        group_data[report_group_column].astype(str)
                    )
                    report_group_summary = (
                        group_data
                        .groupby(report_group_column)[report_measure_column]
                        .agg(
                            Responses="count",
                            Average="mean",
                            Median="median",
                            Minimum="min",
                            Maximum="max"
                        )
                        .reset_index()
                    )
                    report_group_summary[
                        ["Average", "Median", "Minimum", "Maximum"]
                    ] = report_group_summary[
                        ["Average", "Median", "Minimum", "Maximum"]
                    ].round(2)
                    report_group_summary = report_group_summary.sort_values(
                        by="Average", ascending=False
                    )

            if report_group_summary.empty:
                report_group_summary = pd.DataFrame({
                    "Information": [
                        "No suitable group comparison was available."
                    ]
                })

            report_cleaning = pd.DataFrame({
                "Cleaning Operation": [
                    "Blank Value Detection",
                    "Duplicate Row Removal",
                    "Duplicate ID Removal",
                    "Empty Column Removal",
                    "Empty Row Removal",
                    "Text Space Standardisation",
                    "Missing Value Treatment"
                ],
                "Status": [
                    "Completed",
                    "Applied" if remove_duplicates else "Not Applied",
                    "Applied" if remove_duplicate_ids else "Not Applied",
                    "Applied" if remove_empty_columns else "Not Applied",
                    "Applied" if remove_empty_rows else "Not Applied",
                    "Applied" if trim_text else "Not Applied",
                    missing_strategy
                ]
            })

            # PDF generator
            def generate_pdf_report():
                pdf_buffer = io.BytesIO()
                document = SimpleDocTemplate(
                    pdf_buffer,
                    pagesize=A4,
                    rightMargin=40,
                    leftMargin=40,
                    topMargin=45,
                    bottomMargin=45
                )

                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    "ReportTitle",
                    parent=styles["Title"],
                    fontSize=20,
                    leading=25,
                    alignment=TA_CENTER,
                    textColor=colors.HexColor("#17365D"),
                    spaceAfter=15
                )
                heading_style = ParagraphStyle(
                    "ReportHeading",
                    parent=styles["Heading2"],
                    fontSize=13,
                    leading=17,
                    textColor=colors.HexColor("#17365D"),
                    spaceBefore=12,
                    spaceAfter=8
                )
                normal_style = styles["BodyText"]
                story = []

                story.append(Paragraph(
                    "PUBLIC SURVEY RESPONSE ANALYSER", title_style
                ))
                story.append(Paragraph(
                    "PROFESSIONAL SURVEY ANALYSIS REPORT", heading_style
                ))
                story.append(Spacer(1, 10))

                story.append(Paragraph(
                    "1. DATASET OVERVIEW", heading_style
                ))
                overview_data = [
                    ["Report Information", "Details"],
                    ["Dataset Name", escape(str(report_filename))],
                    ["Report Generated", escape(report_date)],
                    ["Original Responses", str(len(original_df))],
                    ["Cleaned Responses", str(len(clean_df))],
                    ["Total Columns", str(len(clean_df.columns))],
                    ["Missing Values", str(cleaned_missing)],
                    ["Duplicate Rows", str(cleaned_duplicates)],
                    ["Data Quality Score", f"{cleaned_quality:.2f}%"]
                ]
                overview_table = Table(
                    overview_data, colWidths=[180, 300], repeatRows=1
                )
                overview_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0),
                     colors.HexColor("#17365D")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("PADDING", (0, 0), (-1, -1), 7),
                    ("VALIGN", (0, 0), (-1, -1), "TOP")
                ]))
                story.append(overview_table)

                story.append(Paragraph(
                    "2. DATA QUALITY ANALYSIS", heading_style
                ))
                quality_data = [[
                    "Column", "Data Type", "Missing", "Missing %"
                ]]
                for _, row in report_quality.iterrows():
                    quality_data.append([
                        Paragraph(escape(str(row["Column"])), normal_style),
                        str(row["Data Type"]),
                        str(row["Missing Values"]),
                        f"{row['Missing Percentage']:.2f}%"
                    ])

                quality_table = Table(
                    quality_data,
                    colWidths=[170, 120, 75, 75],
                    repeatRows=1
                )
                quality_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0),
                     colors.HexColor("#17365D")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("PADDING", (0, 0), (-1, -1), 6),
                    ("VALIGN", (0, 0), (-1, -1), "TOP")
                ]))
                story.append(quality_table)

                story.append(Paragraph(
                    "3. NUMERICAL ANALYSIS", heading_style
                ))
                if report_numeric_columns:
                    numerical_data = [[
                        "Column", "Mean", "Median", "Minimum", "Maximum"
                    ]]
                    for column in report_numeric_columns:
                        values = clean_df[column].dropna()
                        if not values.empty:
                            numerical_data.append([
                                Paragraph(escape(str(column)), normal_style),
                                f"{values.mean():.2f}",
                                f"{values.median():.2f}",
                                str(values.min()),
                                str(values.max())
                            ])

                    numerical_table = Table(
                        numerical_data,
                        colWidths=[150, 80, 80, 80, 80],
                        repeatRows=1
                    )
                    numerical_table.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0),
                         colors.HexColor("#17365D")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("PADDING", (0, 0), (-1, -1), 6)
                    ]))
                    story.append(numerical_table)
                else:
                    story.append(Paragraph(
                        "No suitable numerical columns found.",
                        normal_style
                    ))

                story.append(Paragraph(
                    "4. RESPONSE DISTRIBUTION", heading_style
                ))
                if (
                    not report_responses.empty
                    and "Column" in report_responses.columns
                ):
                    response_data = [[
                        "Column", "Response", "Count", "Percentage"
                    ]]
                    for _, row in report_responses.head(40).iterrows():
                        response_data.append([
                            Paragraph(
                                escape(str(row["Column"])), normal_style
                            ),
                            Paragraph(
                                escape(str(row["Response"])), normal_style
                            ),
                            str(row["Count"]),
                            f"{row['Percentage']:.2f}%"
                        ])

                    response_table = Table(
                        response_data,
                        colWidths=[120, 200, 70, 80],
                        repeatRows=1
                    )
                    response_table.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0),
                         colors.HexColor("#17365D")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("PADDING", (0, 0), (-1, -1), 6),
                        ("VALIGN", (0, 0), (-1, -1), "TOP")
                    ]))
                    story.append(response_table)

                    if len(report_responses) > 40:
                        story.append(Paragraph(
                            "Note: PDF displays the first 40 response "
                            "categories. Excel contains all categories.",
                            normal_style
                        ))
                else:
                    story.append(Paragraph(
                        "No categorical response data available.",
                        normal_style
                    ))

                story.append(Paragraph(
                    "5. GROUP ANALYSIS", heading_style
                ))
                if (
                    not report_group_summary.empty
                    and "Information" not in report_group_summary.columns
                ):
                    group_data_pdf = [
                        [str(c) for c in report_group_summary.columns]
                    ]
                    for _, row in report_group_summary.head(30).iterrows():
                        group_data_pdf.append([
                            str(value) for value in row.tolist()
                        ])

                    group_table = Table(group_data_pdf, repeatRows=1)
                    group_table.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0),
                         colors.HexColor("#17365D")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("PADDING", (0, 0), (-1, -1), 6),
                        ("FONTSIZE", (0, 0), (-1, -1), 8)
                    ]))
                    story.append(group_table)

                    if len(report_group_summary) > 30:
                        story.append(Paragraph(
                            "Note: PDF displays the first 30 groups. "
                            "Excel contains all groups.",
                            normal_style
                        ))
                else:
                    story.append(Paragraph(
                        "No suitable group analysis was available.",
                        normal_style
                    ))

                story.append(Paragraph(
                    "6. KEY FINDINGS", heading_style
                ))
                for index, finding in enumerate(
                    report_findings, start=1
                ):
                    story.append(Paragraph(
                        f"{index}. {escape(str(finding))}", normal_style
                    ))
                    story.append(Spacer(1, 5))

                story.append(Paragraph(
                    "7. RECOMMENDATIONS", heading_style
                ))
                for index, recommendation in enumerate(
                    report_recommendations, start=1
                ):
                    story.append(Paragraph(
                        f"{index}. {escape(str(recommendation))}",
                        normal_style
                    ))
                    story.append(Spacer(1, 5))

                story.append(Paragraph(
                    "8. REPORT LIMITATIONS", heading_style
                ))
                limitations = [
                    "This report describes the uploaded survey dataset.",
                    "Statistical relationships do not automatically "
                    "establish cause-and-effect relationships.",
                    "Automatically generated recommendations should be "
                    "reviewed before taking action.",
                    "The data quality score is an indicative project "
                    "metric, not an industry-standard certification."
                ]
                for item in limitations:
                    story.append(Paragraph(
                        "• " + escape(item), normal_style
                    ))
                    story.append(Spacer(1, 4))

                story.append(Spacer(1, 20))
                story.append(Paragraph(
                    "Generated by Public Survey Response Analyser",
                    normal_style
                ))

                document.build(story)
                pdf_buffer.seek(0)
                return pdf_buffer.getvalue()

            # Excel generator
            def generate_excel_report():
                excel_buffer = io.BytesIO()

                overview = pd.DataFrame({
                    "Report Information": [
                        "Application", "Dataset Name", "Report Generated",
                        "Original Responses", "Cleaned Responses",
                        "Original Columns", "Cleaned Columns",
                        "Original Missing Values", "Cleaned Missing Values",
                        "Original Duplicate Rows", "Cleaned Duplicate Rows",
                        "Original Quality Score", "Cleaned Quality Score"
                    ],
                    "Value": [
                        "Public Survey Response Analyser",
                        report_filename,
                        report_date,
                        len(original_df),
                        len(clean_df),
                        len(original_df.columns),
                        len(clean_df.columns),
                        original_missing,
                        cleaned_missing,
                        original_duplicates,
                        cleaned_duplicates,
                        f"{original_quality:.2f}%",
                        f"{cleaned_quality:.2f}%"
                    ]
                })

                findings_df = pd.DataFrame({
                    "Key Findings": report_findings
                })
                recommendations_df = pd.DataFrame({
                    "Recommendations": report_recommendations
                })
                smart_insights_df = pd.DataFrame({
                    "Smart Insights": smart_insights
                })
                smart_recommendations_df = pd.DataFrame({
                    "Smart Recommendations": unique_recommendations
                })

                with pd.ExcelWriter(
                    excel_buffer, engine="openpyxl"
                ) as writer:
                    overview.to_excel(
                        writer, index=False, sheet_name="Overview"
                    )
                    original_df.to_excel(
                        writer, index=False, sheet_name="Original Data"
                    )
                    clean_df.to_excel(
                        writer, index=False, sheet_name="Cleaned Data"
                    )
                    report_quality.to_excel(
                        writer, index=False, sheet_name="Data Quality"
                    )
                    report_numerical.to_excel(
                        writer, index=False, sheet_name="Numerical Analysis"
                    )
                    report_responses.to_excel(
                        writer, index=False, sheet_name="Response Distribution"
                    )
                    report_group_summary.to_excel(
                        writer, index=False, sheet_name="Group Analysis"
                    )
                    report_cleaning.to_excel(
                        writer, index=False, sheet_name="Cleaning Report"
                    )
                    findings_df.to_excel(
                        writer, index=False, sheet_name="Key Findings"
                    )
                    recommendations_df.to_excel(
                        writer, index=False, sheet_name="Recommendations"
                    )
                    smart_insights_df.to_excel(
                        writer, index=False, sheet_name="Smart Insights"
                    )
                    smart_recommendations_df.to_excel(
                        writer, index=False,
                        sheet_name="Smart Recommendations"
                    )

                    for worksheet in writer.sheets.values():
                        worksheet.freeze_panes = "A2"
                        worksheet.auto_filter.ref = worksheet.dimensions

                        for cell in worksheet[1]:
                            cell.fill = PatternFill(
                                fill_type="solid",
                                fgColor="17365D"
                            )
                            cell.font = Font(
                                bold=True,
                                color="FFFFFF"
                            )

                        for column_cells in worksheet.columns:
                            max_length = 0
                            column_letter = (
                                column_cells[0].column_letter
                            )
                            for cell in column_cells:
                                try:
                                    max_length = max(
                                        max_length,
                                        len(str(cell.value))
                                    )
                                except Exception:
                                    pass

                            worksheet.column_dimensions[
                                column_letter
                            ].width = min(max_length + 3, 40)

                excel_buffer.seek(0)
                return excel_buffer.getvalue()

            # Report download controls
            st.subheader("📥 Download Professional Reports")
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📄 PDF Report")
                st.write(
                    "Formatted report with overview, quality analysis, "
                    "numerical findings, response distribution, group "
                    "analysis and recommendations."
                )

                if st.button(
                    "Generate PDF Report",
                    key="step7_generate_pdf"
                ):
                    try:
                        pdf_data = generate_pdf_report()
                        st.download_button(
                            label="📥 Download Professional PDF",
                            data=pdf_data,
                            file_name="professional_survey_report.pdf",
                            mime="application/pdf",
                            key="step7_download_pdf"
                        )
                        st.success("PDF report generated successfully!")
                    except Exception as error:
                        st.error(f"PDF generation failed: {error}")

            with col2:
                st.markdown("### 📊 Excel Report")
                st.write(
                    "Detailed workbook with separate worksheets "
                    "for original data, cleaned data, quality, "
                    "analysis, findings and recommendations."
                )

                if st.button(
                    "Generate Excel Report",
                    key="step7_generate_excel"
                ):
                    try:
                        excel_data = generate_excel_report()
                        st.download_button(
                            label="📥 Download Professional Excel",
                            data=excel_data,
                            file_name="professional_survey_report.xlsx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            key="step7_download_excel"
                        )
                        st.success("Excel report generated successfully!")
                    except Exception as error:
                        st.error(f"Excel generation failed: {error}")

            st.divider()
            st.subheader("📌 Report Contents")
            st.markdown("""
            **PDF Report**
            - Dataset overview and metadata
            - Data quality analysis
            - Numerical analysis
            - Response distribution
            - Group analysis
            - Key findings and recommendations
            - Report limitations

            **Excel Workbook**
            - Overview
            - Original Data
            - Cleaned Data
            - Data Quality
            - Numerical Analysis
            - Response Distribution
            - Group Analysis
            - Cleaning Report
            - Key Findings
            - Recommendations
            - Smart Insights
            - Smart Recommendations
            """)

            st.caption(
                "Reports are generated locally using Python. "
                "No external AI API is required."
            )
        # --------------------------------------------------
        # TAB 9: INTERACTIVE DASHBOARD - STEP 8
        # --------------------------------------------------

        with tab9:

            st.header("🚀 Interactive Survey Dashboard")
            st.write(
                "Explore your survey dataset using interactive "
                "filters, KPIs, charts and group comparisons."
            )

            st.divider()

            # Create a separate copy so dashboard filters
            # do not modify the original cleaned dataset.
            dashboard_df = clean_df.copy(deep=True)

            # --------------------------------------------------
            # 1. INTERACTIVE FILTERS
            # --------------------------------------------------

            st.subheader("🔎 Interactive Data Filters")

            dashboard_categorical = [
                col for col in dashboard_df.columns
                if classify_column(col) == "Categorical"
            ]

            dashboard_numeric = [
                col for col in dashboard_df.select_dtypes(
                    include="number"
                ).columns
                if classify_column(col) != "Identifier"
            ]

            filtered_df = dashboard_df.copy(deep=True)

            filter_col1, filter_col2 = st.columns(2)

            with filter_col1:

                if dashboard_categorical:

                    selected_filter_columns = st.multiselect(
                        "Select categorical columns to filter",
                        dashboard_categorical,
                        key="dashboard_filter_columns"
                    )

                    for column in selected_filter_columns:

                        available_values = sorted(
                            dashboard_df[column]
                            .dropna()
                            .astype(str)
                            .unique()
                            .tolist()
                        )

                        selected_values = st.multiselect(
                            f"Select values for {column}",
                            available_values,
                            default=available_values,
                            key=f"dashboard_values_{column}"
                        )

                        include_missing = st.checkbox(
                            f"Include missing values in {column}",
                            value=True,
                            key=f"dashboard_missing_{column}"
                        )

                        value_mask = (
                            dashboard_df[column]
                            .astype(str)
                            .isin(selected_values)
                        )

                        if include_missing:
                            value_mask = (
                                value_mask
                                | dashboard_df[column].isna()
                            )

                        filtered_df = filtered_df[value_mask]

                else:
                    st.info(
                        "No categorical columns available for filtering."
                    )

            with filter_col2:

                if dashboard_numeric:

                    selected_numeric_filters = st.multiselect(
                        "Select numerical columns to filter",
                        dashboard_numeric,
                        key="dashboard_numeric_filters"
                    )

                    for column in selected_numeric_filters:

                        valid_values = dashboard_df[column].dropna()

                        if valid_values.empty:
                            continue

                        minimum = float(valid_values.min())
                        maximum = float(valid_values.max())

                        if minimum == maximum:
                            st.info(
                                f"{column} has only one numerical value: "
                                f"{minimum}"
                            )
                            continue

                        selected_range = st.slider(
                            f"Select range for {column}",
                            min_value=minimum,
                            max_value=maximum,
                            value=(minimum, maximum),
                            key=f"dashboard_range_{column}"
                        )

                        include_missing_numeric = st.checkbox(
                            f"Include missing values in {column}",
                            value=True,
                            key=f"dashboard_numeric_missing_{column}"
                        )

                        numeric_mask = (
                            dashboard_df[column].between(
                                selected_range[0],
                                selected_range[1]
                            )
                        )

                        if include_missing_numeric:
                            numeric_mask = (
                                numeric_mask
                                | dashboard_df[column].isna()
                            )

                        filtered_df = filtered_df[numeric_mask]

                else:
                    st.info(
                        "No numerical columns available for filtering."
                    )

            st.divider()

            # --------------------------------------------------
            # 2. DASHBOARD KPI CARDS
            # --------------------------------------------------

            st.subheader("📌 Live Dashboard Overview")

            total_responses = len(filtered_df)
            total_columns = len(filtered_df.columns)
            total_missing = int(
                filtered_df.isnull().sum().sum()
            )
            total_duplicates = int(
                filtered_df.duplicated().sum()
            )

            kpi1, kpi2, kpi3, kpi4 = st.columns(4)

            kpi1.metric(
                "Total Responses",
                total_responses
            )

            kpi2.metric(
                "Total Columns",
                total_columns
            )

            kpi3.metric(
                "Missing Values",
                total_missing
            )

            kpi4.metric(
                "Duplicate Rows",
                total_duplicates
            )

            st.caption(
                "All dashboard figures and charts update according "
                "to the filters selected above."
            )

            st.divider()

            if filtered_df.empty:

                st.warning(
                    "No records match the selected filters. "
                    "Change your filters to view dashboard results."
                )

            else:

                # --------------------------------------------------
                # 3. NUMERICAL KPI ANALYSIS
                # --------------------------------------------------

                st.subheader("📊 Numerical KPI Analysis")

                if dashboard_numeric:

                    selected_kpi_column = st.selectbox(
                        "Select numerical column",
                        dashboard_numeric,
                        key="dashboard_kpi_column"
                    )

                    selected_kpi_measure = st.selectbox(
                        "Select KPI measure",
                        ["Mean", "Median", "Minimum", "Maximum"],
                        key="dashboard_kpi_measure"
                    )

                    kpi_data = filtered_df[
                        selected_kpi_column
                    ].dropna()

                    if not kpi_data.empty:

                        if selected_kpi_measure == "Mean":
                            kpi_value = kpi_data.mean()

                        elif selected_kpi_measure == "Median":
                            kpi_value = kpi_data.median()

                        elif selected_kpi_measure == "Minimum":
                            kpi_value = kpi_data.min()

                        else:
                            kpi_value = kpi_data.max()

                        st.metric(
                            f"{selected_kpi_measure} - {selected_kpi_column}",
                            f"{kpi_value:.2f}"
                        )

                    else:
                        st.info(
                            "No valid numerical values are available."
                        )

                else:
                    st.info(
                        "No numerical columns are available for KPI analysis."
                    )

                st.divider()

                # --------------------------------------------------
                # 4. INTERACTIVE RESPONSE DISTRIBUTION
                # --------------------------------------------------

                st.subheader("📈 Response Distribution")

                if dashboard_categorical:

                    selected_chart_category = st.selectbox(
                        "Select categorical column",
                        dashboard_categorical,
                        key="dashboard_chart_category"
                    )

                    response_data = (
                        filtered_df[selected_chart_category]
                        .dropna()
                        .astype(str)
                        .value_counts()
                        .head(15)
                        .reset_index()
                    )

                    response_data.columns = [
                        "Response",
                        "Count"
                    ]

                    if not response_data.empty:

                        response_fig = px.bar(
                            response_data,
                            x="Response",
                            y="Count",
                            text="Count",
                            title=(
                                f"Response Distribution: "
                                f"{selected_chart_category}"
                            ),
                            color="Response"
                        )

                        st.plotly_chart(
                            response_fig,
                            use_container_width=True
                        )

                        st.dataframe(
                            response_data,
                            use_container_width=True,
                            hide_index=True
                        )

                    else:
                        st.info(
                            "No categorical responses available."
                        )

                else:
                    st.info(
                        "No categorical columns available for charts."
                    )

                st.divider()

                # --------------------------------------------------
                # 5. NUMERICAL DISTRIBUTION CHARTS
                # --------------------------------------------------

                st.subheader("📉 Numerical Data Visualization")

                if dashboard_numeric:

                    selected_chart_numeric = st.selectbox(
                        "Select numerical column for visualization",
                        dashboard_numeric,
                        key="dashboard_chart_numeric"
                    )

                    chart_data = filtered_df[
                        selected_chart_numeric
                    ].dropna()

                    if not chart_data.empty:

                        chart_col1, chart_col2 = st.columns(2)

                        with chart_col1:

                            histogram_fig = px.histogram(
                                x=chart_data,
                                title=(
                                    f"Histogram: "
                                    f"{selected_chart_numeric}"
                                ),
                                labels={
                                    "x": selected_chart_numeric,
                                    "y": "Frequency"
                                }
                            )

                            st.plotly_chart(
                                histogram_fig,
                                use_container_width=True
                            )

                        with chart_col2:

                            box_fig = px.box(
                                y=chart_data,
                                title=(
                                    f"Box Plot: "
                                    f"{selected_chart_numeric}"
                                ),
                                labels={
                                    "y": selected_chart_numeric
                                }
                            )

                            st.plotly_chart(
                                box_fig,
                                use_container_width=True
                            )

                    else:
                        st.info(
                            "No valid numerical values available."
                        )

                else:
                    st.info(
                        "No numerical columns available for visualization."
                    )

                st.divider()

                # --------------------------------------------------
                # 6. GROUP-WISE COMPARISON
                # --------------------------------------------------

                st.subheader("👥 Interactive Group Comparison")

                if dashboard_categorical and dashboard_numeric:

                    group_col1, group_col2 = st.columns(2)

                    with group_col1:

                        dashboard_group = st.selectbox(
                            "Select group column",
                            dashboard_categorical,
                            key="dashboard_group_column"
                        )

                    with group_col2:

                        dashboard_measure = st.selectbox(
                            "Select numerical measure",
                            dashboard_numeric,
                            key="dashboard_group_measure"
                        )

                    group_data = filtered_df[
                        [dashboard_group, dashboard_measure]
                    ].dropna().copy()

                    if not group_data.empty:

                        group_data[dashboard_group] = (
                            group_data[dashboard_group].astype(str)
                        )

                        dashboard_group_summary = (
                            group_data
                            .groupby(dashboard_group)[dashboard_measure]
                            .agg(
                                Responses="count",
                                Average="mean",
                                Median="median"
                            )
                            .reset_index()
                        )

                        dashboard_group_summary[
                            ["Average", "Median"]
                        ] = dashboard_group_summary[
                            ["Average", "Median"]
                        ].round(2)

                        st.dataframe(
                            dashboard_group_summary,
                            use_container_width=True,
                            hide_index=True
                        )

                        group_fig = px.bar(
                            dashboard_group_summary,
                            x=dashboard_group,
                            y="Average",
                            text="Average",
                            title=(
                                f"Average {dashboard_measure} "
                                f"by {dashboard_group}"
                            ),
                            hover_data=["Responses", "Median"]
                        )

                        st.plotly_chart(
                            group_fig,
                            use_container_width=True
                        )

                    else:
                        st.info(
                            "No valid records for group comparison."
                        )

                else:
                    st.info(
                        "Group comparison requires categorical "
                        "and numerical columns."
                    )

                st.divider()

                # --------------------------------------------------
                # 7. CORRELATION HEATMAP
                # --------------------------------------------------

                st.subheader("🔗 Interactive Correlation Heatmap")

                if len(dashboard_numeric) >= 2:

                    correlation_data = (
                        filtered_df[dashboard_numeric].corr()
                    )

                    correlation_fig = px.imshow(
                        correlation_data,
                        text_auto=".2f",
                        aspect="auto",
                        color_continuous_scale="RdBu",
                        zmin=-1,
                        zmax=1,
                        title="Numerical Correlation Heatmap"
                    )

                    st.plotly_chart(
                        correlation_fig,
                        use_container_width=True
                    )

                    st.caption(
                        "Correlation ranges from -1 to +1. "
                        "Correlation does not prove causation."
                    )

                else:
                    st.info(
                        "Correlation analysis requires at least "
                        "two numerical columns."
                    )

                st.divider()

                # --------------------------------------------------
                # 8. FILTERED DATA PREVIEW AND DOWNLOAD
                # --------------------------------------------------

                st.subheader("📋 Filtered Dataset Preview")

                st.write(
                    f"Showing {len(filtered_df)} filtered responses."
                )

                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    height=300
                )

                dashboard_csv = filtered_df.to_csv(
                    index=False
                ).encode("utf-8-sig")

                st.download_button(
                    label="📥 Download Filtered Dashboard Data (CSV)",
                    data=dashboard_csv,
                    file_name="filtered_survey_dashboard.csv",
                    mime="text/csv",
                    key="dashboard_filtered_download"
                )

            st.divider()

            st.caption(
                "Interactive Dashboard - Step 8 | "
                "Public Survey Response Analyser"
            )
        # --------------------------------------------------
        # 8. DOWNLOAD CLEANED DATASET
        # --------------------------------------------------

        st.divider()
        st.header("📥 Download Cleaned Dataset")
        col1, col2 = st.columns(2)

        csv_data = clean_df.to_csv(
            index=False
        ).encode("utf-8-sig")

        with col1:
            st.download_button(
                label="Download Cleaned CSV",
                data=csv_data,
                file_name="cleaned_survey_data.csv",
                mime="text/csv"
            )

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(
            excel_buffer, engine="openpyxl"
        ) as writer:
            clean_df.to_excel(
                writer, index=False, sheet_name="Cleaned Data"
            )
            original_df.to_excel(
                writer, index=False, sheet_name="Original Data"
            )

        with col2:
            st.download_button(
                label="Download Cleaned Excel",
                data=excel_buffer.getvalue(),
                file_name="cleaned_survey_data.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )


    except Exception as error:
        st.error(f"An error occurred: {error}")

# ----------------------------------------------------------
# EMPTY DASHBOARD STATE
# ----------------------------------------------------------

else:
    st.info(
        "Upload a CSV or Excel survey dataset "
        "using the sidebar to begin analysis."
    )

    st.subheader("What this application can analyse")
    col1, col2, col3 = st.columns(3)
    col1.write("📋 Dataset Overview")
    col2.write("🧹 Data Quality")
    col3.write("📊 Statistical Analysis")
    st.caption("Supported formats: CSV and XLSX")
