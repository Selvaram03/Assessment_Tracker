"""
modules/validator.py

Author : Karthi Dass

Validation Engine

Responsibilities
----------------
1. Validate uploaded files
2. Validate assessment dates
3. Validate course
4. Validate duplicate batches
5. Validate duplicate register numbers
"""

from collections import Counter

import pandas as pd

from config import (
    REGISTER_COLUMN
)


class ValidationError(Exception):
    pass


class Validator:

    def __init__(self):
        pass

    # =====================================================
    # Empty File Validation
    # =====================================================

    def validate_empty_dataframe(
            self,
            dataframe,
            file_name
    ):

        if dataframe.empty:

            raise ValidationError(

                f"{file_name} contains no student records."

            )

    # =====================================================
    # Duplicate Register Numbers
    # =====================================================

    def validate_duplicate_registers(
            self,
            dataframe,
            file_name
    ):

        duplicate = dataframe[

            dataframe.duplicated(

                subset=[REGISTER_COLUMN],

                keep=False

            )

        ]

        if not duplicate.empty:

            registers = sorted(

                duplicate[
                    REGISTER_COLUMN
                ].astype(str).unique()

            )

            raise ValidationError(

                f"Duplicate Register Numbers found in {file_name}\n\n"

                + "\n".join(registers)

            )

    # =====================================================
    # Duplicate Batch Upload
    # =====================================================

    def validate_duplicate_batches(
            self,
            uploaded_files
    ):

        batches_by_date = {}

        for file in uploaded_files:
            batches_by_date.setdefault(file.assessment_date, []).append(
                file.batch
            )

        duplicate_lines = []

        for assessment_date, batches in batches_by_date.items():
            duplicate_batches = [
                batch
                for batch, count in Counter(batches).items()
                if count > 1
            ]

            if duplicate_batches:
                duplicate_lines.append(
                    f"{assessment_date}: {', '.join(sorted(duplicate_batches))}"
                )

        if duplicate_lines:

            raise ValidationError(

                "Duplicate Batch Upload\n\n"

                + "\n".join(duplicate_lines)

            )

    # =====================================================
    # Assessment Date Validation
    # =====================================================

    def validate_assessment_date(
            self,
            uploaded_files
    ):

        dates = [

            file.assessment_date

            for file in uploaded_files
        ]

        unique_dates = set(dates)

        return sorted(unique_dates)

    # =====================================================
    # Course Validation
    # =====================================================

    def validate_course(
            self,
            uploaded_files
    ):

        courses = {

            file.course

            for file in uploaded_files

        }

        if len(courses) > 1:

            raise ValidationError(

                "Multiple Courses Uploaded."

            )

        return list(courses)[0]

    # =====================================================
    # Validate All Uploaded Files
    # =====================================================

    def validate(
            self,
            uploaded_files
    ):

        if not uploaded_files:
            raise ValidationError(
                "No assessment files uploaded."
            )

        self.validate_duplicate_batches(
            uploaded_files
        )

        assessment_date = self.validate_assessment_date(
            uploaded_files
        )

        course = self.validate_course(
            uploaded_files
        )

        for file in uploaded_files:

            self.validate_empty_dataframe(
                file.dataframe,
                file.file_name
            )

            self.validate_duplicate_registers(
                file.dataframe,
                file.file_name
            )

        return {

            "assessment_date": assessment_date,

            "assessment_dates": assessment_date,

            "course": course

        }
