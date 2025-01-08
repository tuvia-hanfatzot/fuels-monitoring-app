import streamlit as st
import pandas as pd
from io import BytesIO  # Used for in-memory file handling

# Set the title of the app
st.title("Excel Table Comparison Tool")

# Upload files
st.sidebar.header("Upload Your Files")
file1 = st.sidebar.file_uploader("Upload First Excel File (Older)", type=["xlsx"])
file2 = st.sidebar.file_uploader("Upload Second Excel File (Newer)", type=["xlsx"])

if file1 and file2:
    try:
        # Read the specific sheet and skip rows above the actual table
        df1 = pd.read_excel(file1, sheet_name='Distribution', header=2).fillna('')  # Fill NA values with an empty string
        df2 = pd.read_excel(file2, sheet_name='Distribution', header=2).fillna('')  # Fill NA values with an empty string

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

            # Display previews of the dataframes
            st.subheader("Preview of First File (Older - Sheet: Distribution)")
            st.dataframe(df1)

            st.subheader("Preview of Second File (Newer - Sheet: Distribution)")
            st.dataframe(df2)

            # Display added rows
            st.subheader("Rows Added in Newer File")
            if not added_rows.empty:
                st.dataframe(added_rows)

                # Allow users to download added rows
                output_added = BytesIO()
                with pd.ExcelWriter(output_added, engine="openpyxl") as writer:
                    added_rows.drop(columns=['comparison_key'], inplace=False).to_excel(writer, index=False, sheet_name="Added Rows")
                added_data = output_added.getvalue()

                st.download_button(
                    label="Download Added Rows as Excel",
                    data=added_data,
                    file_name="added_rows.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No rows were added in the newer file.")

            # Display removed rows
            st.subheader("Rows Removed in Newer File")
            if not removed_rows.empty:
                st.dataframe(removed_rows)

                # Allow users to download removed rows
                output_removed = BytesIO()
                with pd.ExcelWriter(output_removed, engine="openpyxl") as writer:
                    removed_rows.drop(columns=['comparison_key'], inplace=False).to_excel(writer, index=False, sheet_name="Removed Rows")
                removed_data = output_removed.getvalue()

                st.download_button(
                    label="Download Removed Rows as Excel",
                    data=removed_data,
                    file_name="removed_rows.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No rows were removed in the newer file.")

        else:
            st.error("The columns 'Description' and 'Agency' must exist in both files. Please check your Excel files.")
    except Exception as e:
        st.error(f"An error occurred while reading the Excel files: {e}")
else:
    st.warning("Please upload both Excel files to proceed.")
