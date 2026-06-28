"""
modules/parser.py

Reads uploaded assessment files, validates filenames,
groups them into Advanced/Intermediate and returns
structured objects.

Author : Karthi Dass
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List
import pandas as pd

from config import (
    REQUIRED_COLUMNS,
    REGISTER_COLUMN
)

from modules.helpers import (
    parse_filename,
    get_batch_type,
    clean_register_number
)
from modules.validator import ValidationError


# ============================================================
# Uploaded Assessment File
# ============================================================

@dataclass
class UploadedAssessmentFile:

    file_name: str
    course: str
    batch: str
    assessment_date: str
    sheet_type: str
    dataframe: pd.DataFrame


# ============================================================
# Assessment Parser
# ============================================================

class AssessmentParser:

    def __init__(self):

        self.files: List[UploadedAssessmentFile] = []

    # --------------------------------------------------------

    def _validate_columns(self, df: pd.DataFrame):

        missing = []

        for col in REQUIRED_COLUMNS:

            if col not in df.columns:
                missing.append(col)

        if missing:
            raise ValidationError(
                f"Missing Columns : {missing}"
            )

    # --------------------------------------------------------

    def _clean_dataframe(self, df: pd.DataFrame):

        df.columns = df.columns.str.strip()

        self._validate_columns(df)

        df[REGISTER_COLUMN] = (

            df[REGISTER_COLUMN]

            .apply(clean_register_number)

            .astype(str)

            .str.strip()

        )

        return df

    # --------------------------------------------------------

    def load_uploaded_files(self, uploaded_files):

        """
        uploaded_files

        Streamlit UploadedFile list
        """

        self.files.clear()

        if not uploaded_files:
            return self.files

        for file in uploaded_files:

            course, batch, assessment_date = parse_filename(
                file.name
            )

            sheet_type = get_batch_type(batch)

            df = pd.read_excel(file)

            df = self._clean_dataframe(df)

            assessment = UploadedAssessmentFile(

                file_name=file.name,

                course=course,

                batch=batch,

                assessment_date=assessment_date,

                sheet_type=sheet_type,

                dataframe=df

            )

            self.files.append(
                assessment
            )

        return self.files

    # --------------------------------------------------------

    def validate_dates(self):

        if not self.files:
            raise ValidationError("No assessment files uploaded.")

        dates = {

            x.assessment_date

            for x in self.files

        }

        if len(dates) != 1:

            raise ValidationError(

                "Uploaded files contain multiple dates."

            )

        return list(dates)[0]

    # --------------------------------------------------------

    def group_files(self):

        grouped = {

            "Advanced": [],

            "Intermediate": []

        }

        for file in self.files:

            grouped[
                file.sheet_type
            ].append(file)

        return grouped

    # --------------------------------------------------------

    def get_course(self):

        if not self.files:
            raise ValidationError("No assessment files uploaded.")

        courses = {

            x.course

            for x in self.files

        }

        return list(courses)[0]

    # --------------------------------------------------------

    def summary(self):

        rows = []

        for file in self.files:

            rows.append({

                "File Name": file.file_name,

                "Course": file.course,

                "Batch": file.batch,

                "Sheet": file.sheet_type,

                "Assessment Date": file.assessment_date,

                "Students": len(file.dataframe)

            })

        return pd.DataFrame(rows)

    # --------------------------------------------------------

    def advanced_files(self):

        return [

            x

            for x in self.files

            if x.sheet_type == "Advanced"

        ]

    # --------------------------------------------------------

    def intermediate_files(self):

        return [

            x

            for x in self.files

            if x.sheet_type == "Intermediate"

        ]

    # --------------------------------------------------------

    def advanced_count(self):

        total = 0

        for file in self.advanced_files():

            total += len(file.dataframe)

        return total

    # --------------------------------------------------------

    def intermediate_count(self):

        total = 0

        for file in self.intermediate_files():

            total += len(file.dataframe)

        return total

    # --------------------------------------------------------

    def reset(self):

        self.files.clear()
