import pandas as pd
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, NamedStyle, PatternFill

# ========================= CONFIG =========================
input_dir = "."  # Streamlit uses current directory

# For Streamlit - we'll use uploaded file instead of local file

def clean_money(x):
    if pd.isna(x) or str(x).strip() in ['', 'nan', 'NaN']:
        return 0.0
    x = str(x).replace('$', '').replace(',', '').strip()
    if x.startswith('(') and x.endswith(')'):
        x = '-' + x[1:-1]
    try:
        return round(float(x))
    except:
        return 0.0

# ========================= STREAMLIT APP =========================
import streamlit as st

st.set_page_config(page_title="BlueBox Processor", layout="wide")
st.title("🟦 BlueBox Sales Processor")
st.markdown("Upload your raw Excel file to generate the formatted report.")

uploaded_file = st.file_uploader("Choose an Excel file (.xlsx or .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Processing your file..."):
        df = pd.read_excel(uploaded_file)

        # Drop unwanted columns
        df = df.drop(df.columns[3], axis=1)
        df = df.drop(df.columns[8:16], axis=1)

        account_col = df.columns[0]
        category_col = df.columns[1]
        subtotal_col = df.columns[2]

        # Grand Totals
        total_rows = df[df[category_col] == 'Total'].copy()
        total_rows['Sales'] = total_rows.iloc[:, 3].apply(clean_money)
        total_rows['Sales PY'] = total_rows.iloc[:, 4].apply(clean_money)

        grand_totals = total_rows.groupby(account_col).agg({
            'Sales': 'sum',
            'Sales PY': 'sum'
        }).reset_index()

        # Detail Pivot
        detail_df = df[(df[subtotal_col] == 'Total') & (df[category_col] != 'Total')].copy()
        detail_df[category_col] = detail_df[category_col].astype(str).str.strip()
        detail_df['Sales'] = detail_df.iloc[:, 3].apply(clean_money)

        pivot_df = detail_df.pivot_table(
            index=account_col,
            columns=category_col,
            values='Sales',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot_df = pivot_df.rename(columns={pivot_df.columns[0]: 'AccountId - AccountName'})

        final_df = pivot_df.merge(grand_totals, on='AccountId - AccountName', how='left')

        final_df['$ Diff YOY'] = final_df['Sales'] - final_df['Sales PY']
        final_df['% Diff YOY'] = final_df['Sales'] / final_df['Sales PY'].replace(0, pd.NA)

        desired_order = [
            'AccountId - AccountName', 'Sales', 'Sales PY', '$ Diff YOY', '% Diff YOY',
            'All Other Preventive', 'All Other Restorative', 'Blocks', 'Class II',
            'Diagnostics', 'Handpiece', 'Impression Materials', 'IP',
            'Pharma', 'Polishing & Treatments', 'Scaling - Equipment',
            'Scaling - Inserts', 'SUC', 'Temporization', 'Treatments'
        ]

        for col in desired_order:
            if col not in final_df.columns:
                final_df[col] = 0

        final_df = final_df[desired_order]

        numeric_cols = [c for c in final_df.columns if c != 'AccountId - AccountName']
        final_df[numeric_cols] = final_df[numeric_cols].round(0).fillna(0)

        # Total Row
        total_row = {'AccountId - AccountName': 'Total'}
        for col in numeric_cols:
            if col != '% Diff YOY':
                total_row[col] = final_df[col].sum()

        total_sales = total_row.get('Sales', 0)
        total_sales_py = total_row.get('Sales PY', 0)
        total_row['% Diff YOY'] = (total_sales / total_sales_py) if total_sales_py != 0 else 0

        final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
        final_df = final_df[final_df['AccountId - AccountName'] != 'Total']
        final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

        # Save to Excel
        output_file = "Processed_Output.xlsx"
        final_df.to_excel(output_file, index=False)

        # ========================= FORMATTING =========================
        wb = load_workbook(output_file)
        ws = wb.active

        light_yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        aqua_marine = PatternFill(start_color="7FFFD4", end_color="7FFFD4", fill_type="solid")

        regular_font = Font(size=8)
        bold_font = Font(size=8, bold=True)
        red_font = Font(size=8, color="FF0000")
        bold_red_font = Font(size=8, color="FF0000", bold=True)

        # Header
        for cell in ws[1]:
            cell.font = bold_font

        # Number Styles
        currency_style = NamedStyle(name="currency")
        currency_style.number_format = '$#,##0;($#,##0)'

        percent_style = NamedStyle(name="percent")
        percent_style.number_format = '0%'

        header_values = [cell.value for cell in ws[1]]

        # Number formatting
        for col_idx in range(2, ws.max_column + 1):
            col_name = header_values[col_idx - 1]
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row_idx, col_idx)
                if isinstance(cell.value, (int, float)) and cell.value is not None:
                    if col_name == '% Diff YOY':
                        cell.style = percent_style
                    else:
                        cell.style = currency_style

        # Fonts + Backgrounds
        for row_idx in range(2, ws.max_row + 1):
            is_total_row = (row_idx == ws.max_row)
            
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row_idx, col_idx)
                
                # Font - Red negatives in Column D
                if col_idx == 4:  # Column D
                    if isinstance(cell.value, (int, float)) and cell.value < 0:
                        cell.font = bold_red_font if is_total_row else red_font
                    else:
                        cell.font = bold_font if is_total_row else regular_font
                else:
                    cell.font = bold_font if is_total_row else regular_font
                
                # B-E: Light Yellow
                if 2 <= col_idx <= 5:
                    cell.fill = light_yellow
                
                # F-T: Zero → Blank + Aqua
                elif 6 <= col_idx <= 20:
                    if isinstance(cell.value, (int, float)) and cell.value == 0:
                        cell.value = None
                        cell.fill = aqua_marine

        # Freeze & Auto width
        ws.freeze_panes = 'A2'

        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[column_letter].width = max_length + 3

        wb.save(output_file)

        st.success("✅ Processing Complete!")

        # Download Button
        with open(output_file, "rb") as file:
            st.download_button(
                label="📥 Download Processed Excel File",
                data=file,
                file_name="Processed_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.dataframe(final_df, use_container_width=True)

else:
    st.info("👆 Please upload your Excel file to begin processing.")
