import streamlit as st
import pandas as pd

# Function to preprocess a single sheet
def preprocess_sheet(df):
    """Preprocess the DataFrame by removing rows above the 'TOTAL' row."""
    # Convert all values to strings and find the row containing 'TOTAL' (case-insensitive)
    total_row_index = df.apply(lambda row: row.astype(str).str.contains("TOTAL", case=False, na=False)).any(axis=1).idxmax()
    
    # Remove rows above the "TOTAL" row
    processed_df = df.iloc[total_row_index:].reset_index(drop=True)
    
    # Set the first row as the header
    processed_df.columns = processed_df.iloc[0].astype(str)  # Ensure headers are strings
    processed_df = processed_df[1:].reset_index(drop=True)
    
    return processed_df

# Function to collect thresholds for each sheet
def create_thresholdUI(uploaded_file):
    """Create the UI to select sheets and input thresholds."""
    xls = pd.ExcelFile(uploaded_file)
    thresholds = {}

    # Multi-selection dropdown for sheets
    selected_sheets = st.multiselect("Select sheets to process:", xls.sheet_names)

    if selected_sheets:
        st.write("Set thresholds for the selected sheets:")
        # Collect thresholds for each selected sheet
        for name in selected_sheets:
            weak_test_threshold = st.number_input(
                f"Enter threshold for sheet {name} (Weak Student):", min_value=0, step=1, format="%d", key=f"weak_{name}"
            )
            bright_test_threshold = st.number_input(
                f"Enter threshold for sheet {name} (Bright Student):", min_value=0, step=1, format="%d", key=f"bright_{name}"
            )
            thresholds[name] = [weak_test_threshold, bright_test_threshold]

    return thresholds

# Function to process the Excel file based on thresholds
def process_excel(uploaded_file, thresholds):
    """Process the uploaded Excel file based on thresholds."""
    xls = pd.ExcelFile(uploaded_file)
    result_file = pd.ExcelWriter("Result.xlsx", engine='openpyxl')

    for name in thresholds.keys():
        st.write(f"Processing sheet: {name}")

        # Read and preprocess the data
        df = pd.read_excel(xls, sheet_name=name)
        df = preprocess_sheet(df)

        # Normalize column names for consistent access
        df.columns = df.columns.str.strip().str.upper()

        # Check for the "TOTAL" column (case-insensitive)
        if "TOTAL" not in df.columns:
            st.error(f"The column 'TOTAL' is missing in sheet '{name}'. Skipping...")
            continue

        # Extract thresholds
        weak_test_threshold, bright_test_threshold = thresholds[name]

        # Filter data for weak and bright students
        df_weak = df[df["TOTAL"] <= weak_test_threshold]
        df_bright = df[df["TOTAL"] >= bright_test_threshold]

        # Reset indexes
        df_weak.reset_index(drop=True, inplace=True)
        df_bright.reset_index(drop=True, inplace=True)

        # Write processed data to the result Excel file
        df_weak.to_excel(result_file, sheet_name=(name + " (Weak)"), index=False)
        df_bright.to_excel(result_file, sheet_name=(name + " (Bright)"), index=False)

        st.success(f"Sheet {name} processed successfully!")

    # Save the final Excel file
    result_file.close()
    return "Result.xlsx"

# Streamlit App
st.title("Test Scoring App")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.write("Excel file uploaded successfully!")

    # Collect thresholds when user uploads a file
    thresholds = create_thresholdUI(uploaded_file)

    # Ensure processing starts only when the button is clicked
    if bool(thresholds) and st.button("Start Processing"):
        # Ensure thresholds are valid before processing
        if all(val is not None and val != 0 for threshold_list in thresholds.values() for val in threshold_list):
            result_path = process_excel(uploaded_file, thresholds)

            # Provide a download link for the processed result
            with open(result_path, "rb") as f:
                st.download_button(
                    label="Download Processed File",
                    data=f,
                    file_name="Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("Please enter valid thresholds for all sheets.")