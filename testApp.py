import streamlit as st
import pandas as pd

# Function to preprocess a single sheet
def preprocess_sheet(df):
    """Preprocess the DataFrame by removing rows above the 'TOTAL' row."""
    total_row_index = df.apply(lambda row: row.astype(str).str.contains("TOTAL", case=False, na=False)).any(axis=1).idxmax()
    processed_df = df.iloc[total_row_index:].reset_index(drop=True)

    if total_row_index != 0:
        processed_df.columns = processed_df.iloc[0].astype(str)
        processed_df = processed_df[1:].reset_index(drop=True)

    return processed_df

# Function to collect maximum marks for each sheet
def create_max_marks_UI(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    max_marks = {}

    selected_sheets = st.multiselect("Select sheets to process:", xls.sheet_names)

    if selected_sheets:
        st.write("Set maximum marks for the selected sheets:")
        for name in selected_sheets:
            max_marks_input = st.number_input(
                f"Enter maximum marks for sheet {name}:", min_value=1, step=1, format="%d", key=f"max_marks_{name}"
            )
            max_marks[name] = max_marks_input

    return max_marks

# Function to process Excel file
def process_excel(uploaded_file, max_marks, metadata):
    xls = pd.ExcelFile(uploaded_file)
    result_file_path = "Result.xlsx"
    result_file = pd.ExcelWriter(result_file_path, engine="openpyxl")

    for name in max_marks.keys():
        st.write(f"Processing sheet: {name}")
        df = pd.read_excel(xls, sheet_name=name)
        df = preprocess_sheet(df)
        df.columns = df.columns.str.strip().str.upper()

        if "TOTAL" not in df.columns:
            st.error(f"The column 'TOTAL' is missing in sheet '{name}'. Skipping...")
            continue

        # Get the maximum marks for this sheet
        max_mark = max_marks[name]

        # Calculate the thresholds based on 40% and 80% of the maximum marks
        weak_test_threshold = 0.40 * max_mark
        bright_test_threshold = 0.80 * max_mark

        df_weak = df[df["TOTAL"] <= weak_test_threshold]
        df_bright = df[df["TOTAL"] >= bright_test_threshold]

        df_weak.reset_index(drop=True, inplace=True)
        df_bright.reset_index(drop=True, inplace=True)

        # Adding a "Signature" column
        df_weak["SIGNATURE"] = ""
        df_bright["SIGNATURE"] = ""

        # Create metadata DataFrame with the new "Course" field
        metadata_info = pd.DataFrame(
            {
                "Metadata": ["Name of the Faculty:", "Program:", "Year:", "Semester:", "Course:"],
                "Details": [metadata["faculty"], metadata["program"], metadata["year"], metadata["semester"], metadata["course"]],
            }
        )

        # Save metadata and student data to Excel
        metadata_info.to_excel(result_file, sheet_name=name + " (Weak)", index=False, header=False)
        df_weak.to_excel(result_file, sheet_name=name + " (Weak)", startrow=5, index=False)
        metadata_info.to_excel(result_file, sheet_name=name + " (Bright)", index=False, header=False)
        df_bright.to_excel(result_file, sheet_name=name + " (Bright)", startrow=5, index=False)

        st.success(f"Sheet {name} processed successfully!")

    result_file.close()
    return result_file_path

# Streamlit App
st.title("Test Scoring App")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.write("Excel file uploaded successfully!")

    # Include the new "Course" metadata field
    metadata = {
        "faculty": st.text_input("Name of Faculty"),
        "program": st.text_input("Program"),
        
        # Dropdown for Year
        "year": st.selectbox(
            "Select Year",
            ["F.Y. B.TECH", "S.Y. B.TECH", "T.Y. B.TECH", "Final Year B.TECH"]
        ),
        
        # Dropdown for Semester
        "semester": st.selectbox(
            "Select Semester",
            ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
        ),
        
        "course": st.text_input("Course"),  # Added Course field
    }

    max_marks = create_max_marks_UI(uploaded_file)

    if max_marks and st.button("Start Processing"):
        if all(val is not None and val > 0 for val in max_marks.values()):
            result_path = process_excel(uploaded_file, max_marks, metadata)

            # Excel Download
            with open(result_path, "rb") as f:
                st.download_button(
                    label="Download Processed Excel File",
                    data=f,
                    file_name="Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("Please enter valid maximum marks for all sheets.")