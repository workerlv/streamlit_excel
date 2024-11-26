from io import BytesIO
import streamlit as st
import pandas as pd


st.set_page_config(layout="wide")


def create_excel(data_frame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        data_frame.to_excel(writer, index=False, sheet_name="Results")
        workbook = writer.book
        worksheet = writer.sheets["Results"]
        format1 = workbook.add_format({"num_format": "0.00"})
        worksheet.set_column("A:A", None, format1)
    return output.getvalue()


def sidebar_options(df_lookup_table, df_value_table):
    selected_lookup_values = st.sidebar.selectbox(
        "Select headings", df_lookup_table.columns, index=0
    )
    select_column_to_look_in = st.sidebar.selectbox(
        "Select column to look in", df_lookup_table.columns, index=1
    )
    select_main_values = st.sidebar.selectbox(
        "Select lookup value column", df_value_table.columns, index=0
    )

    lookup_table = pd.DataFrame(
        {
            "keys": df_lookup_table[selected_lookup_values],
            "values": df_lookup_table[select_column_to_look_in],
        }
    )
    value_table = pd.DataFrame({"keys": df_value_table[select_main_values]})

    return lookup_table, value_table


def display_data_preview(df_lookup_table, df_value_table, labels):
    col_1, col_2 = st.columns(2)
    with col_1:
        st.write(f"{labels[0]}:")
        st.dataframe(df_lookup_table)
    with col_2:
        st.write(f"{labels[1]}:")
        st.dataframe(df_value_table)


def sidebar_settings():
    st.sidebar.title("Settings")
    demo_mode = st.sidebar.toggle("Run demo mode", value=False)

    if not demo_mode:
        st.sidebar.write("For detailed info toggle to demo mode")

    return demo_mode


def run_demo_mode():
    st.header("Demo mode")

    st.sidebar.divider()
    st.sidebar.write("Demo mode you can check the functionality with demo data")
    st.sidebar.write(
        "To use your own data press on toggle button in sidebar 'Run demo mode'"
    )
    st.sidebar.divider()

    path_to_lookup_values = "example_data/lookup_values.xlsx"
    path_to_lookup_table = "example_data/lookup_table.xlsx"

    df_lookup_values = pd.read_excel(path_to_lookup_values)
    df_lookup_table = pd.read_excel(path_to_lookup_table)

    st.divider()
    st.write("Data from excel files")

    display_data_preview(
        df_lookup_table, df_lookup_values, ("Lookup Table", "Lookup Values")
    )

    st.divider()
    st.write("Results for demo mode")

    df_lookup_table_final, df_lookup_values_final = sidebar_options(
        df_lookup_table, df_lookup_values
    )

    merged_df = pd.merge(
        df_lookup_values_final, df_lookup_table_final, on="keys", how="left"
    )

    st.dataframe(merged_df)


def run_main_mode():
    st.header("VLOOKUP in excel")

    st.write("Upload look up table and look up values with header row as first row")

    with st.expander("Preview example files"):
        col_ex_1, col_ex_2 = st.columns(2)

        with col_ex_1:
            st.write("Example look up table")
            st.image("example_images/lookup_table.png")

        with col_ex_2:
            st.write("Example look up values")
            st.image("example_images/lookup_values.png")

    col_1, col_2 = st.columns(2)
    # TODO: add csv option
    with col_1:
        uploaded_lookup_table = st.file_uploader("Upload Lookup Table", type=["xlsx"])
    with col_2:
        uploaded_lookup_values = st.file_uploader("Upload Lookup Values", type=["xlsx"])

    if not (uploaded_lookup_table and uploaded_lookup_values):
        st.warning("Please upload both files to proceed.")
        st.stop()

    # Read uploaded files
    df_lookup_table = pd.read_excel(uploaded_lookup_table)
    df_value_table = pd.read_excel(uploaded_lookup_values)

    # Preview uploaded data
    with st.expander("Preview Uploaded Data"):
        display_data_preview(
            df_lookup_table, df_value_table, ("Lookup Table", "Lookup Values")
        )

    df_lookup_table_final, df_lookup_values_final = sidebar_options(
        df_lookup_table, df_value_table
    )

    st.divider()
    st.write("Results:")
    merged_df = pd.merge(
        df_lookup_values_final, df_lookup_table_final, on="keys", how="left"
    )

    st.dataframe(merged_df)

    st.download_button(
        label="Download result", data=create_excel(merged_df), file_name="result.xlsx"
    )


def main():

    demo_mode = sidebar_settings()

    if demo_mode:
        run_demo_mode()
    else:
        run_main_mode()


if __name__ == "__main__":
    main()
