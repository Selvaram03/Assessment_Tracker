"""
Assessment Automation
Author : Karthi Dass

Main Streamlit Application
"""

from pathlib import Path
import tempfile
from datetime import datetime

import streamlit as st
import pandas as pd

from modules.parser import AssessmentParser
from modules.validator import Validator
from modules.ranking import RankingEngine
from modules.excel_engine import ExcelEngine

from config import *

MASTER_DISPLAY_TO_KEY = {
    "IT": "IT",
    "ECE": "ECE",
    "Shanmuga College": "SSCET",
}

COMBINED_MASTER_TYPES = {"ECE", "SSCET"}


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


def parse_assessment_date(value):

    text = str(value).strip()

    try:
        return datetime.strptime(text, "%d-%m-%Y")
    except ValueError:
        return datetime.max


def group_files_by_date(uploaded_files):

    grouped = {}

    for file in uploaded_files:
        grouped.setdefault(file.assessment_date, []).append(file)

    return grouped


def build_output_path(assessment_dates):

    if not assessment_dates:
        return OUTPUT_DIR / OUTPUT_FILENAME

    if len(assessment_dates) == 1:
        return OUTPUT_DIR / f"Assessment_{assessment_dates[0]}.xlsx"

    return OUTPUT_DIR / (
        f"Assessment_{assessment_dates[0]}_to_{assessment_dates[-1]}.xlsx"
    )


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

if "date_summary_df" not in st.session_state:
    st.session_state.date_summary_df = None


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

    master_display_name = st.selectbox(
        "Select Master",
        list(MASTER_DISPLAY_TO_KEY.keys()),
        index=0
    )

    master_type = MASTER_DISPLAY_TO_KEY[master_display_name]

    if st.session_state.selected_master_type != master_type:

        for key in [
            "advanced_df",
            "intermediate_df",
            "combined_df",
            "summary",
            "output_file",
            "date_summary_df"
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

    master_file_types = ["xlsx"]

    master_file = st.file_uploader(

        f"Upload {master_display_name} Master File (optional)",

        type=master_file_types,

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

            master_path = MASTER_TEMPLATES[master_type]

            if not master_path.exists():

                st.error(
                    f"{master_display_name} master template not found at {master_path}. "
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

        assessment_dates = sorted(
            validation["assessment_dates"],
            key=parse_assessment_date
        )

        progress.progress(40)

        output_file = build_output_path(assessment_dates)

        engine = ExcelEngine(
            master_path,
            master_type=master_type
        )

        engine.load_workbook()
        engine.create_backup()
        engine.build_register_maps()

        ranking = RankingEngine()

        grouped_files = group_files_by_date(parsed_files)
        ordered_dates = sorted(
            grouped_files.keys(),
            key=parse_assessment_date
        )

        summary_rows = []
        overall_advanced = {
            "total": 0,
            "updated": 0,
            "absent": 0,
            "skipped": 0
        }
        overall_intermediate = {
            "total": 0,
            "updated": 0,
            "absent": 0,
            "skipped": 0
        }
        overall_combined = {
            "total": 0,
            "updated": 0,
            "absent": 0,
            "skipped": 0
        }
        advanced_preview_frames = []
        intermediate_preview_frames = []
        combined_preview_frames = []

        for index, assessment_date in enumerate(ordered_dates, start=1):

            status.info(
                f"Processing {assessment_date} "
                f"({index}/{len(ordered_dates)})..."
            )

            date_files = grouped_files[assessment_date]
            date_df = pd.concat(
                [
                    file.dataframe
                    for file in date_files
                ],
                ignore_index=True
            )

            progress_value = 40 + int((index / max(len(ordered_dates), 1)) * 50)
            progress.progress(progress_value)

            if master_type in COMBINED_MASTER_TYPES:

                processed_df = ranking.process(
                    date_df
                )

                combined_preview_frames.append(
                    processed_df
                )

                date_summary = engine.update_workbook(
                    processed_df,
                    pd.DataFrame(),
                    assessment_date
                )

                engine.update_top10(
                    assessment_date,
                    processed_df,
                    pd.DataFrame()
                )

                engine.update_last10(
                    assessment_date,
                    processed_df,
                    pd.DataFrame()
                )

                combined_summary = date_summary[master_type]

                for key in overall_combined:
                    overall_combined[key] += combined_summary.get(key, 0)

                summary_rows.append({
                    "Assessment Date": assessment_date,
                    "Students": combined_summary.get("total", 0),
                    "Updated": combined_summary.get("updated", 0),
                    "Absent": combined_summary.get("absent", 0),
                    "Skipped": combined_summary.get("skipped", 0),
                })

            else:

                advanced_df, intermediate_df, unmatched_df = (
                    engine.split_uploaded_dataframe(
                        date_df
                    )
                )

                advanced_df = ranking.process(
                    advanced_df
                )

                intermediate_df = ranking.process(
                    intermediate_df
                )

                advanced_preview_frames.append(
                    advanced_df
                )

                intermediate_preview_frames.append(
                    intermediate_df
                )

                date_summary = engine.update_workbook(
                    advanced_df,
                    intermediate_df,
                    assessment_date
                )

                engine.update_top10(
                    assessment_date,
                    advanced_df,
                    intermediate_df
                )

                engine.update_last10(
                    assessment_date,
                    advanced_df,
                    intermediate_df
                )

                advanced_summary = date_summary["Advanced"]
                intermediate_summary = date_summary["Intermediate"]

                for key in overall_advanced:
                    overall_advanced[key] += advanced_summary.get(key, 0)

                for key in overall_intermediate:
                    overall_intermediate[key] += intermediate_summary.get(key, 0)

                summary_rows.append({
                    "Assessment Date": assessment_date,
                    "Advanced Students": advanced_summary.get("total", 0),
                    "Advanced Updated": advanced_summary.get("updated", 0),
                    "Advanced Absent": advanced_summary.get("absent", 0),
                    "Advanced Skipped": advanced_summary.get("skipped", 0),
                    "Intermediate Students": intermediate_summary.get("total", 0),
                    "Intermediate Updated": intermediate_summary.get("updated", 0),
                    "Intermediate Absent": intermediate_summary.get("absent", 0),
                    "Intermediate Skipped": intermediate_summary.get("skipped", 0),
                    "Unmatched Rows": len(unmatched_df),
                })

            progress.progress(progress_value)

        if master_type in COMBINED_MASTER_TYPES:

            summary = {
                master_type: overall_combined
            }

            if combined_preview_frames:
                st.session_state.combined_df = pd.concat(
                    combined_preview_frames,
                    ignore_index=True
                )
            else:
                st.session_state.combined_df = None

            st.session_state.advanced_df = None
            st.session_state.intermediate_df = None

            skipped_total = overall_combined["skipped"]

        else:

            summary = {
                "Advanced": overall_advanced,
                "Intermediate": overall_intermediate
            }

            if advanced_preview_frames:
                st.session_state.advanced_df = pd.concat(
                    advanced_preview_frames,
                    ignore_index=True
                )
            else:
                st.session_state.advanced_df = None

            if intermediate_preview_frames:
                st.session_state.intermediate_df = pd.concat(
                    intermediate_preview_frames,
                    ignore_index=True
                )
            else:
                st.session_state.intermediate_df = None

            st.session_state.combined_df = None

            skipped_total = (
                overall_advanced["skipped"]
                + overall_intermediate["skipped"]
            )

        engine.save_workbook(
            output_file
        )

        progress.progress(100)

        # ---------------------------------------
        # Store Session
        # ---------------------------------------

        st.session_state.summary = summary
        st.session_state.date_summary_df = pd.DataFrame(summary_rows)

        st.session_state.output_file = output_file

        if skipped_total:

            st.warning(
                f"{skipped_total} uploaded register numbers were not found "
                "in the master workbook and were skipped."
            )

        status.success(
            f"Workbook Generated Successfully for {len(ordered_dates)} date(s)."
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

    if master_type in COMBINED_MASTER_TYPES:

        combined_summary = st.session_state.summary[master_type]

        st.success(master_display_name)

        st.metric(

            "Students",

            combined_summary["total"]

        )

        st.metric(

            "Updated",

            combined_summary["updated"]

        )

        st.metric(

            "Absent",

            combined_summary["absent"]

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

    if st.session_state.date_summary_df is not None and not st.session_state.date_summary_df.empty:

        st.caption("Date-wise summary")

        st.dataframe(

            st.session_state.date_summary_df,

            use_container_width=True,

            hide_index=True

        )



# ==========================================================
# Preview Data
# ==========================================================

if master_type in COMBINED_MASTER_TYPES and st.session_state.combined_df is not None:

    st.divider()

    with st.expander(
        f"{master_display_name} Preview",
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

        "date_summary_df",

        "output_file"

    ]:

        if key in st.session_state:

            del st.session_state[key]

    rerun_app()
