# -*- coding: utf-8 -*-
"""Excel_Compare.ipynb
"""
import streamlit as st
import pandas as pd
import openpyxl

def compare_excel_files(file1, file2):
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        if df1.shape != df2.shape:
              st.error("Error: Files must have the same dimensions for comparison.")
              return None

        changes = []
        for index, row in df1.iterrows():
            for col in df1.columns:
                val1 = df1.loc[index, col]
                val2 = df2.loc[index, col]
                if pd.isna(val1) and pd.isna(val2):
                    continue
                elif pd.notna(val1) and pd.isna(val2):
                    changes.append([len(changes)+1, f"{col}{index+1}", val1, ""])
                elif pd.isna(val1) and pd.notna(val2):
                    changes.append([len(changes)+1, f"{col}{index+1}", "", val2])
                elif val1 != val2:
                    changes.append([len(changes)+1, f"{col}{index+1}", val1, val2])

        changes_df = pd.DataFrame(changes, columns=["Change Number", "Cell Reference", "Old Value", "New Value"])
        return changes_df
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

st.title("Excel File Comparison")

uploaded_file1 = st.file_uploader("Choose the first Excel file", type="xlsx")
uploaded_file2 = st.file_uploader("Choose the second Excel file", type="xlsx")

if uploaded_file1 is not None and uploaded_file2 is not None:
    changes_df = compare_excel_files(uploaded_file1, uploaded_file2)

    if changes_df is not None:
        st.write(changes_df)

        # Download the changes as an Excel file
        output = pd.ExcelWriter('comparison_report.xlsx', engine='openpyxl')
        changes_df.to_excel(output, sheet_name='Changes', index=False)
        output.save()
        with open("comparison_report.xlsx", "rb") as fp:
            btn = st.download_button(
                label="Download Excel file",
                data=fp,
                file_name="comparison_report.xlsx",
                mime="application/vnd.ms-excel"
            )
