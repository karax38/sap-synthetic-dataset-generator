from __future__ import annotations

import streamlit as st

from generator import create_excel_file, generate_sap_dataset


st.set_page_config(
    page_title="SAP Synthetic Inventory Dataset Generator",
    page_icon=":material/inventory_2:",
    layout="centered",
)


def render_summary(stats: dict[str, int]) -> None:
    col1, col2, col3 = st.columns(3)
    col1.metric("Total materials", f"{stats['total_materials']:,}")
    col2.metric("Total plants", f"{stats['total_plants']:,}")
    col3.metric("MATDOC rows generated", f"{stats['matdoc_rows']:,}")


def main() -> None:
    st.title("SAP Synthetic Inventory Dataset Generator")
    st.write(
        "This tool creates synthetic SAP-like inventory datasets for safety stock optimization "
        "demos, testing, and inventory analytics validation."
    )

    with st.form("generation_form"):
        num_plants = st.number_input(
            "Number of plants",
            min_value=1,
            max_value=50,
            value=4,
            step=1,
        )
        materials_per_plant = st.number_input(
            "Materials per plant",
            min_value=10,
            max_value=5000,
            value=100,
            step=10,
        )
        years_of_history = st.number_input(
            "Years of history",
            min_value=1,
            max_value=10,
            value=3,
            step=1,
        )
        generate_clicked = st.form_submit_button(
            "Generate Dataset",
            type="primary",
            use_container_width=True,
        )

    if "generated" not in st.session_state:
        st.session_state.generated = None

    if generate_clicked:
        with st.spinner("Generating synthetic SAP dataset..."):
            dataset = generate_sap_dataset(
                num_plants=int(num_plants),
                materials_per_plant=int(materials_per_plant),
                years_of_history=int(years_of_history),
            )
            excel_bytes = create_excel_file(dataset["tables"])
            st.session_state.generated = {
                "dataset": dataset,
                "excel_bytes": excel_bytes,
            }

    generated = st.session_state.generated
    if generated:
        stats = generated["dataset"]["stats"]
        render_summary(stats)

        st.download_button(
            label="Download Excel Dataset",
            data=generated["excel_bytes"],
            file_name="synthetic_sap_dataset.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("Generation details", expanded=False):
            metadata = generated["dataset"]["metadata"]
            st.write(
                f"Generated from {metadata['start_date']} to {metadata['end_date']} "
                f"with seed `{metadata['seed']}`."
            )


if __name__ == "__main__":
    main()
