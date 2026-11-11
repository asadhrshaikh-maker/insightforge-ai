import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="Data Copilot Pro",
    page_icon="📊",
    layout="wide"
)

# ==================================================
# DARK THEME
# ==================================================
st.markdown("""
<style>

.stApp{
    background-color:#0f172a;
    color:white;
}

[data-testid="stSidebar"]{
    background:#111827;
}

div[data-testid="stMetric"]{
    background:#1e293b;
    border:1px solid #334155;
    border-radius:15px;
    padding:15px;
}

h1,h2,h3,h4{
    color:white;
}

</style>
""", unsafe_allow_html=True)

# ==================================================
# LOAD DATA
# ==================================================
@st.cache_data
def load_data(file):

    if file.name.endswith(".csv"):
        return pd.read_csv(file)

    return pd.read_excel(file)

# ==================================================
# DETECT DATASET
# ==================================================
def detect_dataset(df):

    cols = [str(c).lower() for c in df.columns]

    if any(x in " ".join(cols) for x in
           ["status", "project", "ra name", "interviewer"]):
        return "Survey Dataset"

    if any(x in " ".join(cols) for x in
           ["sales", "revenue", "price", "totalprice"]):
        return "Sales Dataset"

    return "General Dataset"

# ==================================================
# EXECUTIVE SUMMARY
# ==================================================
def executive_summary(df):

    result = []

    result.append(
        f"Dataset contains {len(df):,} rows and {len(df.columns)} columns."
    )

    numeric = df.select_dtypes(include=np.number)

    if len(numeric.columns):

        biggest = numeric.sum().sort_values(
            ascending=False
        ).index[0]

        result.append(
            f"Highest value metric is {biggest} "
            f"with total {df[biggest].sum():,.2f}"
        )

    cat_cols = df.select_dtypes(include="object")

    for col in cat_cols.columns[:3]:

        try:
            top = df[col].value_counts().idxmax()

            cnt = df[col].value_counts().max()

            result.append(
                f"Top {col}: {top} ({cnt:,} rows)"
            )

        except:
            pass

    return result

# ==================================================
# SIDEBAR
# ==================================================
st.sidebar.title("🤖 Data Copilot Pro")

file = st.sidebar.file_uploader(
    "Upload File",
    type=["csv", "xlsx", "xls"]
)

if not file:

    st.title("📊 Data Copilot Pro")

    st.info(
        "Upload any Excel or CSV file to start analysis."
    )

    st.stop()

# ==================================================
# DATA
# ==================================================
df = load_data(file)

# ==================================================
# FILTERS
# ==================================================
st.sidebar.markdown("---")
st.sidebar.header("🎯 Filters")

filtered = df.copy()

cat_cols = df.select_dtypes(
    include="object"
).columns.tolist()

for col in cat_cols[:5]:

    vals = sorted(
        filtered[col]
        .dropna()
        .astype(str)
        .unique()
    )

    selected = st.sidebar.multiselect(
        col,
        vals,
        default=vals
    )

    if selected:

        filtered = filtered[
            filtered[col]
            .astype(str)
            .isin(selected)
        ]

df = filtered

# ==================================================
# HEADER
# ==================================================
st.title("📊 Data Copilot Pro")

dtype = detect_dataset(df)

st.success(
    f"Detected Dataset Type: {dtype}"
)

# ==================================================
# KPI SECTION
# ==================================================
c1,c2,c3,c4 = st.columns(4)

c1.metric(
    "Rows",
    f"{len(df):,}"
)

c2.metric(
    "Columns",
    len(df.columns)
)

c3.metric(
    "Missing Values",
    int(df.isna().sum().sum())
)

c4.metric(
    "Duplicate Rows",
    int(df.duplicated().sum())
)

# ==================================================
# SUMMARY
# ==================================================
st.subheader("📋 Executive Summary")

for line in executive_summary(df):

    st.write("✅", line)

# ==================================================
# CHARTS
# ==================================================
tab1,tab2,tab3,tab4 = st.tabs([
    "📈 Dashboard",
    "🥧 Charts",
    "🔍 Profiling",
    "💾 Export"
])

# ==================================================
# DASHBOARD
# ==================================================
with tab1:

    st.subheader("Dataset Preview")

    st.dataframe(
        df.head(100),
        use_container_width=True
    )

    numeric_cols = df.select_dtypes(
        include=np.number
    ).columns.tolist()

    if numeric_cols:

        metric = st.selectbox(
            "Select Metric",
            numeric_cols
        )

        if len(cat_cols):

            category = st.selectbox(
                "Group By",
                cat_cols
            )

            top = (
                df.groupby(category)[metric]
                .sum()
                .reset_index()
                .sort_values(
                    metric,
                    ascending=False
                )
                .head(10)
            )

            fig = px.bar(
                top,
                x=category,
                y=metric,
                color=metric,
                title=f"{metric} by {category}"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

# ==================================================
# PIE CHARTS
# ==================================================
with tab2:

    if len(cat_cols):

        pie_col = st.selectbox(
            "Pie Chart Column",
            cat_cols
        )

        pie = (
            df[pie_col]
            .value_counts()
            .head(10)
            .reset_index()
        )

        pie.columns = [pie_col, "Count"]

        fig = px.pie(
            pie,
            names=pie_col,
            values="Count",
            hole=0.4,
            title=f"{pie_col} Distribution"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

# ==================================================
# PROFILING
# ==================================================
with tab3:

    profile = []

    for col in df.columns:

        profile.append({
            "Column": col,
            "Type": str(df[col].dtype),
            "Missing": int(df[col].isna().sum()),
            "Unique": int(df[col].nunique())
        })

    profile_df = pd.DataFrame(profile)

    st.dataframe(
        profile_df,
        use_container_width=True
    )

# ==================================================
# EXPORT
# ==================================================
with tab4:

    st.dataframe(
        df.head(100),
        use_container_width=True
    )

    csv = df.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(
        "📥 Download CSV",
        csv,
        file_name="filtered_data.csv",
        mime="text/csv"
    )