import streamlit as st
import pandas as pd

# Function to collect thresholds for each sheet
def create_thresholdUI(uploaded_file):
    # Read the Excel file
    xls  = pd.ExcelFile(uploaded_file)
    thresholds = {}

    # Collect thresholds for each sheet
    for name in xls.sheet_names:
        # Display input fields for weak and bright student thresholds
        weak_test_threshold = st.number_input(
            f"Enter threshold for sheet {name} (Weak Student):", min_value=0, step=1, format="%d", key=f"weak_{name}"
        )
        bright_test_threshold = st.number_input(
            f"Enter threshold for sheet {name} (Bright Student):", min_value=0, step=1, format="%d", key=f"bright_{name}"
        )

        # Save these thresholds in the dictionary
        thresholds[name] = [weak_test_threshold, bright_test_threshold]

    return thresholds

# Function to process the uploaded Excel file based on thresholds
def process_excel(uploaded_file, thresholds):
    # Read the Excel file
    xls = pd.ExcelFile(uploaded_file)
    result_file = pd.ExcelWriter("Result.xlsx", engine='openpyxl')

    # Process each sheet
    for name in xls.sheet_names:
        st.write(f"Processing sheet: {name}")

        # Get the thresholds for the current sheet
        weak_test_threshold, bright_test_threshold = thresholds.get(name, [None, None])

        # Ensure thresholds are valid
        if weak_test_threshold is not None and bright_test_threshold is not None:
            # Read and preprocess the data
            df = pd.read_excel(xls, sheet_name=name)
            df_weak = df[df["TOTAL"] <= weak_test_threshold]
            df_bright = df[df["TOTAL"] >= bright_test_threshold]

            # Reset indexes
            df_weak.reset_index(drop=True, inplace=True)
            df_bright.reset_index(drop=True, inplace=True)

            # Write to the result Excel file
            df_weak.to_excel(result_file, sheet_name=(name + " (Weak)"), index=False)
            df_bright.to_excel(result_file, sheet_name=(name + " (Bright)"), index=False)

            st.success(f"Sheet {name} processed successfully!")
        else:
            st.warning(f"Thresholds for sheet {name} are missing or invalid. Please check the input.")

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
    if st.button("Start Processing"):
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