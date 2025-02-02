import streamlit as st
import pandas as pd
from io import BytesIO  # Used for in-memory file handling

# Set the title of the app
st.title("Excel Table Comparison Tool")

# Upload files
st.sidebar.header("Upload Your Files")
file1 = st.sidebar.file_uploader("Upload First Excel File (Older)", type=["xlsx"])
file2 = st.sidebar.file_uploader("Upload Second Excel File (Newer)", type=["xlsx"])

def get_valid_sheet(file, primary_sheet, fallback_sheet, header=2):
    """Checks available sheets and selects the first valid one."""
    xls = pd.ExcelFile(file)
    available_sheets = [sheet.strip().lower() for sheet in xls.sheet_names]  # Trim and lowercase

    # Normalize sheet names for comparison
    primary_sheet = primary_sheet.strip().lower()
    fallback_sheet = fallback_sheet.strip().lower()

    selected_sheet = None
    if primary_sheet in available_sheets:
        selected_sheet = xls.sheet_names[available_sheets.index(primary_sheet)]
    elif fallback_sheet in available_sheets:
        selected_sheet = xls.sheet_names[available_sheets.index(fallback_sheet)]
    
    if selected_sheet:
        return pd.read_excel(xls, sheet_name=selected_sheet, header=header).fillna('')
    else:
        raise ValueError(f"Neither '{primary_sheet}' nor '{fallback_sheet}' found in the uploaded file. Please check the exact sheet names.")

if file1 and file2:
    try:
        # Read the specific sheet or fallback sheet
        df1 = get_valid_sheet(file1, 'Distribution', 'UNOPS Total Distribution')
        df2 = get_valid_sheet(file2, 'Distribution', 'UNOPS Total Distribution')

        # Ignore the last 2 rows in each table
        df1 = df1.iloc[:-2]
        df2 = df2.iloc[:-2]

        # Ensure required columns exist in both files
        if 'Description' in df1.columns and 'Description' in df2.columns and 'Agency' in df1.columns and 'Agency' in df2.columns:
            # Normalize and standardize the columns for comparison
            df1['Description'] = df1['Description'].astype(str).str.strip().str.lower()
            df2['Description'] = df2['Description'].astype(str).str.strip().str.lower()
            df1['Agency'] = df1['Agency'].astype(str).str.strip().str.lower()
            df2['Agency'] = df2['Agency'].astype(str).str.strip().str.lower()

            # Create a combined key that prioritizes 'Description' and falls back to 'Agency'
            df1['comparison_key'] = df1.apply(lambda row: row['Description'] if row['Description'] else row['Agency'], axis=1)
            df2['comparison_key'] = df2.apply(lambda row: row['Description'] if row['Description'] else row['Agency'], axis=1)

            # Find all unique keys in df1 and df2
            set1 = set(df1['comparison_key'])
            set2 = set(df2['comparison_key'])

            # Identify added and removed rows
            added_keys = set2 - set1
            removed_keys = set1 - set2

            added_rows = df2[df2['comparison_key'].isin(added_keys)]
            removed_rows = df1[df1['comparison_key'].isin(removed_keys)]

            # Display added and removed rows
            st.subheader("Rows Added to Newer File")
            st.dataframe(added_rows)

            st.subheader("Rows Removed from Older File")
            st.dataframe(removed_rows)

            # Combine results into a single Excel file with two sheets
            output_combined = BytesIO()
            with pd.ExcelWriter(output_combined, engine="openpyxl") as writer:
                added_rows.drop(columns=['comparison_key'], inplace=False).to_excel(writer, index=False, sheet_name="Added Rows")
                removed_rows.drop(columns=['comparison_key'], inplace=False).to_excel(writer, index=False, sheet_name="Removed Rows")
            combined_data = output_combined.getvalue()

            # Allow users to download the combined Excel file
            st.download_button(
                label="Download Combined Results as Excel",
                data=combined_data,
                file_name="comparison_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("The columns 'Description' and 'Agency' must exist in both files. Please check your Excel files.")
    except Exception as e:
        st.error(f"An error occurred while reading the Excel files: {e}")
else:
    st.warning("Please upload both Excel files to proceed.")
