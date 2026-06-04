import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, NamedStyle
import io

st.set_page_config(page_title="BlueBox Sales Processor", layout="wide")

st.title("🦷 BlueBox Sales Processor")
st.markdown("### Upload your sales data and get a clean, formatted report")

st.info("""
**How to use this tool:**
1. Open SalesHub and click "Account Performance"
2. Right click the first + sign you see and select Expand --> All
3. Click ... in the upper right corner of the box and select Export
4. Select Data with Current Format
5. Open the file and Save As 'input_data' and select Save As Type "CSV(comma delimited)" 
6. Click the button below to upload it
7. Wait a few seconds for processing
8. Download the formatted Excel report

**Note:** Make sure your file has the standard BlueBox columns.
""")

uploaded_file = st.file_uploader("Upload your input_data.csv file", type=["csv"])

if uploaded_file is not None:
    with st.spinner("Processing your file... This may take 10-20 seconds"):
        try:
            df = pd.read_csv(uploaded_file, encoding='cp1252')

            # [All your processing code remains the same - I'm keeping it short here]
            df = df.drop(df.columns[3], axis=1)
            df = df.drop(df.columns[8:16], axis=1)

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

            total_rows = df[df.iloc[:, 1] == 'Total'].copy()
            total_rows['Sales'] = total_rows.iloc[:, 3].apply(clean_money)
            total_rows['Sales PY'] = total_rows.iloc[:, 4].apply(clean_money)

            grand_totals = total_rows.groupby(df.columns[0]).agg({
                'Sales': 'sum', 'Sales PY': 'sum'
            }).reset_index()

            detail_df = df[(df.iloc[:, 2] == 'Total') & (df.iloc[:, 1] != 'Total')].copy()
            detail_df.iloc[:, 1] = detail_df.iloc[:, 1].astype(str).str.strip()
            detail_df['Sales'] = detail_df.iloc[:, 3].apply(clean_money)

            pivot_df = detail_df.pivot_table(
                index=detail_df.columns[0],
                columns=detail_df.columns[1],
                values='Sales',
                aggfunc='sum',
                fill_value=0
            ).reset_index()

            pivot_df = pivot_df.rename(columns={pivot_df.columns[0]: 'AccountId - AccountName'})

            final_df = pivot_df.merge(grand_totals, on='AccountId - AccountName', how='left')

            final_df['$ Diff YOY'] = final_df['Sales'] - final_df['Sales PY']
            final_df['% Diff YOY'] = (final_df['Sales'] / final_df['Sales PY'].replace(0, pd.NA))

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
            final_df[numeric_cols] = final_df[numeric_cols].round(0)

            # Total Row
            total_row = {'AccountId - AccountName': 'Total'}
            for col in numeric_cols:
                if col != '% Diff YOY':
                    total_row[col] = final_df[col].sum()
            total_sales = total_row.get('Sales', 0)
            total_sales_py = total_row.get('Sales PY', 0)
            total_row['% Diff YOY'] = (total_sales / total_sales_py) if total_sales_py != 0 else 0

            final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
            final_df = final_df.sort_values('AccountId - AccountName').reset_index(drop=True)

            # Create formatted Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)

            wb = load_workbook(output)
            ws = wb.active

            # Formatting
            regular_font = Font(size=8)
            bold_font = Font(size=8, bold=True)

            for cell in ws[1]:
                cell.font = bold_font

            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row_idx, 1).font = regular_font
                for col_idx in range(2, ws.max_column + 1):
                    ws.cell(row_idx, col_idx).font = regular_font

            # Bold Total Row
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(ws.max_row, col_idx).font = bold_font

            currency_style = NamedStyle(name="currency")
            currency_style.number_format = '$#,##0;($#,##0)'
            percent_style = NamedStyle(name="percent")
            percent_style.number_format = '0%'

            for col_idx in range(2, ws.max_column + 1):
                col_name = ws.cell(1, col_idx).value
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row_idx, col_idx)
                    if isinstance(cell.value, (int, float)) and cell.value is not None:
                        if col_name == '% Diff YOY':
                            cell.style = percent_style
                        else:
                            cell.style = currency_style
                        if cell.value < 0:
                            cell.font = Font(size=8, color="FF0000", bold=True)

            ws.freeze_panes = 'A2'

            final_output = io.BytesIO()
            wb.save(final_output)
            final_output.seek(0)

            st.success("✅ Processing Complete!")
            st.download_button(
                label="📥 Download Processed_Report.xlsx",
                data=final_output,
                file_name="Processed_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

else:
    st.info("👆 Please upload your `input_data.csv` file to start processing.")