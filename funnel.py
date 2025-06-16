import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Funnel Visualizer")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    pivot_df = df.pivot_table(index='ST_Program', columns='LeadStatus', values='GroupTotal', aggfunc='sum').fillna(0)

    # Funnel stage order
    stages = [
        "PRE_EVAL_inclusive", "MID_EVAL_inclusive", "MQL_inclusive",
        "SQL_inclusive", "APPLICATION_IN_PROGRESS_inclusive",
        "APPLIED_inclusive", "ADMITTED_inclusive",
        "ADMITTED_ACCEPT_inclusive", "REGISTERED_inclusive"
    ]

    # Helper to get value or 0 if column missing
    def get(col): return pivot_df[col] if col in pivot_df.columns else 0

    # Compute inclusive metrics (each calculated independently from raw stages)
    inclusive_df = pd.DataFrame(index=pivot_df.index)

    inclusive_df["PRE_EVAL_inclusive"] = get("PRE-EVAL") + get("MID-EVAL") + get("MQL") + get("SQL") + \
        get("APPLICATION IN-PROGRESS") + get("APPLIED") + get("APPLICATION CANCELLED") + \
        get("APPLICATION WITHDRAWN") + get("ADMITTED") + get("ADMITTED/ACCEPT") + \
        get("ADMITTED/DECLINE") + get("ADMITTED/DEFER") + get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["MID_EVAL_inclusive"] = get("MID-EVAL") + get("MQL") + get("SQL") + \
        get("APPLICATION IN-PROGRESS") + get("APPLIED") + get("APPLICATION CANCELLED") + \
        get("APPLICATION WITHDRAWN") + get("ADMITTED") + get("ADMITTED/ACCEPT") + \
        get("ADMITTED/DECLINE") + get("ADMITTED/DEFER") + get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["MQL_inclusive"] = get("MQL") + get("SQL") + get("APPLICATION IN-PROGRESS") + \
        get("APPLIED") + get("APPLICATION CANCELLED") + get("APPLICATION WITHDRAWN") + \
        get("ADMITTED") + get("ADMITTED/ACCEPT") + get("ADMITTED/DECLINE") + \
        get("ADMITTED/DEFER") + get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["SQL_inclusive"] = get("SQL") + get("APPLICATION IN-PROGRESS") + get("APPLIED") + \
        get("APPLICATION CANCELLED") + get("APPLICATION WITHDRAWN") + get("ADMITTED") + \
        get("ADMITTED/ACCEPT") + get("ADMITTED/DECLINE") + get("ADMITTED/DEFER") + \
        get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["APPLICATION_IN_PROGRESS_inclusive"] = get("APPLICATION IN-PROGRESS") + get("APPLIED") + \
        get("APPLICATION CANCELLED") + get("APPLICATION WITHDRAWN") + get("ADMITTED") + \
        get("ADMITTED/ACCEPT") + get("ADMITTED/DECLINE") + get("ADMITTED/DEFER") + \
        get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["APPLIED_inclusive"] = get("APPLIED") + get("ADMITTED") + get("ADMITTED/ACCEPT") + \
        get("ADMITTED/DECLINE") + get("ADMITTED/DEFER") + get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["ADMITTED_inclusive"] = get("ADMITTED") + get("ADMITTED/ACCEPT") + get("ADMITTED/DECLINE") + \
        get("ADMITTED/DEFER") + get("ADMITTED/WITHDRAW") + get("REGISTERED")

    inclusive_df["ADMITTED_ACCEPT_inclusive"] = get("ADMITTED/ACCEPT") + get("REGISTERED")

    inclusive_df["REGISTERED_inclusive"] = get("REGISTERED")

    st.subheader("Funnel Visualizations")

    for program in inclusive_df.index:
        data = inclusive_df.loc[program, stages]
        retention = (data / data.shift(1)).fillna(1) * 100
        retention_labels = retention.round(1).astype(str) + '%'

        fig, ax = plt.subplots(figsize=(10, 5))
        bars = ax.bar(stages, data)
        for i, (bar, val, pct) in enumerate(zip(bars, data, retention_labels)):
            label = f"{int(val)}"
            if i > 0:
                label += f"\n({pct})"
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                    label, ha='center', fontsize=8)
        ax.set_title(f"Program: {program}")
        ax.set_ylabel("Leads")
        ax.set_xticks(range(len(stages)))
        ax.set_xticklabels(stages, rotation=45, ha='right')
        st.pyplot(fig)

