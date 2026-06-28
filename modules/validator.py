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

        batches = [

            file.batch

            for file in uploaded_files

        ]

        duplicate = [

            batch

            for batch, count in Counter(

                batches

            ).items()

            if count > 1

        ]

        if duplicate:

            raise ValidationError(

                "Duplicate Batch Upload\n\n"

                + ", ".join(duplicate)

            )

    # =====================================================
    # Assessment Date Validation
    # =====================================================

    def validate_assessment_date(
            self,
            uploaded_files
    ):

        dates = {

            file.assessment_date

            for file in uploaded_files

        }

        if len(dates) > 1:

            raise ValidationError(

                "Multiple Assessment Dates Uploaded."

            )

        return list(dates)[0]

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

            "course": course

        }
