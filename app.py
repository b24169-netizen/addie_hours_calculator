import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Employee Hours Tool", layout="wide")

st.title("Employee Hours Calculator")

st.write("Upload the Excel sheet. The tool will automatically pick Planned vs Actual duration and calculate employee hours.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    # Read excel with row 3 & 4 as headers
    df = pd.read_excel(uploaded_file, header=[2,3])

    st.subheader("Raw Data Preview")
    st.dataframe(df)

    # Extract required columns
    planned_duration = df[("Planned", "Duration")]
    actual_duration = df[("Actual", "Duration")]
    employee = df[("Actual", "Employee")]

    calc_df = pd.DataFrame({
        "Employee": employee,
        "Planned Duration": planned_duration,
        "Actual Duration": actual_duration
    })

    # Calculate chosen hours
    calc_df["Chosen Hours"] = calc_df[["Planned Duration","Actual Duration"]].min(axis=1)

    # Employee totals
    summary = calc_df.groupby("Employee")["Chosen Hours"].sum().reset_index()
    summary.rename(columns={"Chosen Hours":"Total Hours"}, inplace=True)

    st.subheader("Line Item Calculation")
    st.dataframe(calc_df)

    st.subheader("Employee Total Hours")
    st.dataframe(summary)

    # Excel export
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        calc_df.to_excel(writer, sheet_name="Line Items", index=False)
        summary.to_excel(writer, sheet_name="Employee Totals", index=False)

    output.seek(0)

    st.download_button(
        "Download Result Excel",
        data=output,
        file_name="employee_hours_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )