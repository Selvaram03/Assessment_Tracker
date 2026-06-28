"""
Assessment Automation
Author : Karthi Dass

Main Streamlit Application
"""

from pathlib import Path
import tempfile

import streamlit as st
import pandas as pd

from modules.parser import AssessmentParser
from modules.validator import Validator
from modules.merger import Merger
from modules.ranking import RankingEngine
from modules.excel_engine import ExcelEngine

from config import *


# ==========================================================
# Page Configuration
# ==========================================================

st.set_page_config(
    page_title="Assessment Automation",
    page_icon="📊",
    layout="wide"
)


def rerun_app():
    rerun_fn = getattr(st, "rerun", None)

    if rerun_fn is None:
        rerun_fn = getattr(st, "experimental_rerun", None)

    if rerun_fn is None:
        raise AttributeError(
            "This Streamlit version does not support rerun."
        )

    rerun_fn()


# ==========================================================
# Session State
# ==========================================================

if "advanced_df" not in st.session_state:
    st.session_state.advanced_df = None

if "intermediate_df" not in st.session_state:
    st.session_state.intermediate_df = None

if "summary" not in st.session_state:
    st.session_state.summary = None

if "output_file" not in st.session_state:
    st.session_state.output_file = None

if "combined_df" not in st.session_state:
    st.session_state.combined_df = None

if "selected_master_type" not in st.session_state:
    st.session_state.selected_master_type = "IT"


# ==========================================================
# Title
# ==========================================================

st.title("📊 Assessment Automation")

st.caption(
    "Automates Assessment Sheet Generation "
    "using Master Workbook"
)

st.divider()


# ==========================================================
# Sidebar
# ==========================================================

with st.sidebar:

    st.header("Project Information")

    master_type = st.selectbox(
        "Select Master",
        ["IT", "ECE"],
        index=0
    )

    if st.session_state.selected_master_type != master_type:

        for key in [
            "advanced_df",
            "intermediate_df",
            "combined_df",
            "summary",
            "output_file"
        ]:

            if key in st.session_state:

                del st.session_state[key]

        st.session_state.selected_master_type = master_type

        rerun_app()

    st.info(
        """
        ✅ Upload Master Workbook (optional)

        ✅ Use local master template if no file is uploaded

        ✅ Upload Assessment Files

        ✅ Automatic Ranking

        ✅ Automatic Excel Update

        ✅ Download Updated Workbook
        """
    )

    st.divider()

    st.write(
        "**Supported Files**"
    )

    st.write(
        "- Master Workbook (.xlsx)"
    )

    st.write(
        "- Daily Score Reports (.xlsx)"
    )


# ==========================================================
# Upload Section
# ==========================================================

col1, col2 = st.columns(2)

with col1:

    master_file = st.file_uploader(

        f"Upload {master_type} Master Workbook (optional)",

        type=["xlsx"],

        accept_multiple_files=False

    )

with col2:

    uploaded_files = st.file_uploader(

        "Upload Assessment Files",

        type=["xlsx"],

        accept_multiple_files=True

    )


st.divider()


# ==========================================================
# Preview Uploaded Files
# ==========================================================

if uploaded_files:

    st.subheader("Uploaded Files")

    file_names = pd.DataFrame({

        "File Name":

        [f.name for f in uploaded_files]

    })

    st.dataframe(

        file_names,

        use_container_width=True,

        hide_index=True

    )


# ==========================================================
# Start Processing
# ==========================================================

generate = st.button(

    "🚀 Generate Assessment Workbook",

    use_container_width=True,

    type="primary"

)

# ==========================================================
# Processing
# ==========================================================

if generate:

    if not uploaded_files:

        st.error(
            "Please upload Assessment Files."
        )

        st.stop()

    progress = st.progress(0)

    status = st.empty()

    try:

        temp_dir = Path(tempfile.mkdtemp())

        # ---------------------------------------
        # Resolve Master Workbook
        # ---------------------------------------

        status.info(
            "Resolving Master Workbook..."
        )

        progress.progress(10)

        if master_file is not None:

            master_path = temp_dir / master_file.name

            with open(master_path, "wb") as file:

                file.write(
                    master_file.getbuffer()
                )

        else:

            master_path = MASTER_TEMPLATES[
                master_type
            ]

            if not master_path.exists():

                st.error(
                    f"{master_type} master template not found at {master_path}. "
                    "Upload the selected master workbook or place the template there."
                )

                st.stop()

        status.info(
            "Master Workbook Ready."
        )

        # ---------------------------------------
        # Parse Uploaded Files
        # ---------------------------------------

        status.info(
            "Reading Assessment Files..."
        )

        parser = AssessmentParser()

        parsed_files = parser.load_uploaded_files(
            uploaded_files
        )

        progress.progress(25)

        # ---------------------------------------
        # Validation
        # ---------------------------------------

        status.info(
            "Validating Files..."
        )

        validator = Validator()

        validation = validator.validate(
            parsed_files
        )

        assessment_date = validation[
            "assessment_date"
        ]

        progress.progress(40)

        output_file = OUTPUT_DIR / (

            f"Assessment_"

            f"{assessment_date}.xlsx"

        )

        engine = ExcelEngine(
            master_path,
            master_type=master_type
        )

        ranking = RankingEngine()

        if master_type == "ECE":

            status.info(
                "Combining ECE Files..."
            )

            combined_df = pd.concat(
                [
                    file.dataframe
                    for file in parsed_files
                ],
                ignore_index=True
            )

            progress.progress(55)

            status.info(
                "Calculating Ranking..."
            )

            combined_df = ranking.process(
                combined_df
            )

            progress.progress(70)

            status.info(
                "Updating Workbook..."
            )

            summary = engine.process(

                combined_df,

                pd.DataFrame(),

                assessment_date,

                output_file

            )

            st.session_state.combined_df = combined_df
            st.session_state.advanced_df = None
            st.session_state.intermediate_df = None

        else:

            # ---------------------------------------
            # Merge
            # ---------------------------------------

            status.info(
                "Merging Files..."
            )

            merger = Merger()

            advanced_df, intermediate_df = (

                merger.merge(

                    parsed_files

                )

            )

            progress.progress(55)

            # ---------------------------------------
            # Ranking
            # ---------------------------------------

            status.info(
                "Calculating Ranking..."
            )

            advanced_df = ranking.process(
                advanced_df
            )

            intermediate_df = ranking.process(
                intermediate_df
            )

            progress.progress(70)

            # ---------------------------------------
            # Excel Processing
            # ---------------------------------------

            status.info(
                "Updating Workbook..."
            )

            summary = engine.process(

                advanced_df,

                intermediate_df,

                assessment_date,

                output_file

            )

            st.session_state.advanced_df = (
                advanced_df
            )

            st.session_state.intermediate_df = (
                intermediate_df
            )

            st.session_state.combined_df = None

        progress.progress(100)

        # ---------------------------------------
        # Store Session
        # ---------------------------------------

        st.session_state.summary = summary

        st.session_state.output_file = output_file

        if master_type == "ECE":

            skipped_total = summary.get("ECE", {}).get("skipped", 0)

        else:

            skipped_total = sum(
                summary.get(section, {}).get("skipped", 0)
                for section in ("Advanced", "Intermediate")
            )

        if skipped_total:

            st.warning(
                f"{skipped_total} uploaded register numbers were not found "
                "in the master workbook and were skipped."
            )

        status.success(
            "Workbook Generated Successfully."
        )

    except Exception as error:

        progress.empty()

        status.empty()

        st.exception(error)

# ==========================================================
# Processing Summary
# ==========================================================

if st.session_state.summary:

    st.divider()

    st.subheader("Processing Summary")

    if master_type == "ECE":

        ece_summary = st.session_state.summary["ECE"]

        st.success("ECE")

        st.metric(

            "Students",

            ece_summary["total"]

        )

        st.metric(

            "Updated",

            ece_summary["updated"]

        )

        st.metric(

            "Absent",

            ece_summary["absent"]

        )

    else:

        col1, col2 = st.columns(2)

        with col1:

            st.success("Advanced")

            st.metric(

                "Students",

                st.session_state.summary["Advanced"]["total"]

            )

            st.metric(

                "Updated",

                st.session_state.summary["Advanced"]["updated"]

            )

            st.metric(

                "Absent",

                st.session_state.summary["Advanced"]["absent"]

            )

        with col2:

            st.success("Intermediate")

            st.metric(

                "Students",

                st.session_state.summary["Intermediate"]["total"]

            )

            st.metric(

                "Updated",

                st.session_state.summary["Intermediate"]["updated"]

            )

            st.metric(

                "Absent",

                st.session_state.summary["Intermediate"]["absent"]

            )



# ==========================================================
# Preview Data
# ==========================================================

if master_type == "ECE" and st.session_state.combined_df is not None:

    st.divider()

    with st.expander(
        "ECE Preview",
        expanded=False
    ):

        st.dataframe(

            st.session_state.combined_df,

            use_container_width=True,

            hide_index=True

        )


if master_type == "IT" and st.session_state.advanced_df is not None:

    st.divider()

    with st.expander(
        "Advanced Preview",
        expanded=False
    ):

        st.dataframe(

            st.session_state.advanced_df,

            use_container_width=True,

            hide_index=True

        )


if master_type == "IT" and st.session_state.intermediate_df is not None:

    with st.expander(
        "Intermediate Preview",
        expanded=False
    ):

        st.dataframe(

            st.session_state.intermediate_df,

            use_container_width=True,

            hide_index=True

        )


# ==========================================================
# Download Workbook
# ==========================================================

if st.session_state.output_file:

    st.divider()

    with open(

        st.session_state.output_file,

        "rb"

    ) as workbook:

        st.download_button(

            label="📥 Download Updated Workbook",

            data=workbook,

            file_name=Path(
                st.session_state.output_file
            ).name,

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            use_container_width=True,

            type="primary"

        )


# ==========================================================
# Reset
# ==========================================================

st.divider()

if st.button(

    "🔄 Reset",

    use_container_width=True

):

    for key in [

        "advanced_df",

        "intermediate_df",

        "summary",

        "output_file"

    ]:

        if key in st.session_state:

            del st.session_state[key]

    rerun_app()
