import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_funnel(stage_names, counts, title="Funnel"):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.axis('off')
    ax.set_title(title)

    n = len(counts)
    max_width = 8
    min_width = 2
    height = 1
    spacing = 0.2

    for i in range(n):
        top_width = max_width - (i * (max_width - min_width) / max(n-1, 1))
        bottom_width = max_width - ((i + 1) * (max_width - min_width) / max(n-1, 1)) if i < n - 1 else min_width

        x_top_left = (10 - top_width) / 2
        x_bottom_left = (10 - bottom_width) / 2
        y = -i * (height + spacing)

        polygon = patches.Polygon([
            (x_top_left, y),
            (x_top_left + top_width, y),
            (x_bottom_left + bottom_width, y - height),
            (x_bottom_left, y - height)
        ], closed=True, facecolor='skyblue', edgecolor='black')

        ax.add_patch(polygon)

        ax.text(5, y - height / 2, f"{stage_names[i]}: {counts[i]:,}", 
                ha='center', va='center', fontsize=10, weight='bold')

        if i > 0 and counts[i-1] > 0:
            conv = counts[i] / counts[i-1] * 100
            ax.text(5, y + spacing / 2, f"{conv:.1f}%", ha='center', va='bottom', fontsize=9, color='gray')

    plt.ylim(-n * (height + spacing), spacing)
    plt.xlim(0, 10)
    plt.tight_layout()

    
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

    # for program in inclusive_df.index:
    #     data = inclusive_df.loc[program, stages]
    #     retention = (data / data.shift(1)).fillna(1) * 100
    #     retention_labels = retention.round(1).astype(str) + '%'

    #     fig, ax = plt.subplots(figsize=(10, 5))
    #     bars = ax.bar(stages, data)
    #     for i, (bar, val, pct) in enumerate(zip(bars, data, retention_labels)):
    #         label = f"{int(val)}"
    #         if i > 0:
    #             label += f"\n({pct})"
    #         ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
    #                 label, ha='center', fontsize=8)
    #     ax.set_title(f"Program: {program}")
    #     ax.set_ylabel("Leads")
    #     ax.set_xticks(range(len(stages)))
    #     ax.set_xticklabels(stages, rotation=45, ha='right')
        
    

# After computing inclusive_df
    
    
    for program in inclusive_df.index:
        counts = [inclusive_df.loc[program, stage] if stage in inclusive_df.columns else 0 for stage in stages]
        draw_funnel(stages, counts, title=program)
    

    draw_funnel(stages, total_counts)
