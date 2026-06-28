"""
modules/merger.py

Author : Karthi Dass

Responsible for

1. Merge all Advanced files
2. Merge all Intermediate files
3. Detect duplicate Register Numbers
4. Calculate Percentage
5. Calculate Competition Rank
6. Return clean DataFrames
"""

from __future__ import annotations

import pandas as pd

from config import (
    REGISTER_COLUMN,
    SCORE_COLUMN,
    TOTAL_COLUMN,
    TIME_COLUMN,
    PERCENTAGE_COLUMN,
    RANK_COLUMN,
    STATUS_COLUMN,
    ATTENDED,
    ABSENT
)

from modules.helpers import (
    calculate_percentage,
    apply_competition_rank,
    clean_register_number
)


class Merger:

    def __init__(self):

        self.advanced_df = pd.DataFrame()

        self.intermediate_df = pd.DataFrame()

    # ==========================================================
    # Merge Uploaded Files
    # ==========================================================

    def merge(self, uploaded_files):

        """
        uploaded_files

        list[UploadedAssessmentFile]
        """

        advanced = []

        intermediate = []

        for file in uploaded_files:

            df = file.dataframe.copy()

            if file.sheet_type == "Advanced":
                advanced.append(df)

            else:
                intermediate.append(df)

        if advanced:

            self.advanced_df = pd.concat(

                advanced,

                ignore_index=True

            )

        else:

            self.advanced_df = pd.DataFrame()

        if intermediate:

            self.intermediate_df = pd.concat(

                intermediate,

                ignore_index=True

            )

        else:

            self.intermediate_df = pd.DataFrame()

        return (

            self.advanced_df,

            self.intermediate_df

        )

    # ==========================================================
    # Clean Register Numbers
    # ==========================================================

    def clean_register_numbers(self, df):

        if df.empty:
            return df

        df[REGISTER_COLUMN] = (

            df[REGISTER_COLUMN]

            .apply(clean_register_number)

            .astype(str)

            .str.strip()

        )

        return df

    # ==========================================================
    # Duplicate Register Numbers
    # ==========================================================

    def check_duplicate_register_numbers(self, df):

        duplicate = df[

            df.duplicated(

                subset=[REGISTER_COLUMN],

                keep=False

            )

        ]

        if not duplicate.empty:

            duplicate = duplicate.sort_values(

                REGISTER_COLUMN

            )

            raise Exception(

                "\nDuplicate Register Numbers Found\n\n"

                +

                duplicate[[REGISTER_COLUMN]]

                .drop_duplicates()

                .to_string(index=False)

            )

    # ==========================================================
    # Calculate Percentage
    # ==========================================================

    def calculate_percentage(self, df):

        if df.empty:
            return df

        df[PERCENTAGE_COLUMN] = df.apply(

            lambda row:

            calculate_percentage(

                row[SCORE_COLUMN],

                row[TOTAL_COLUMN]

            ),

            axis=1

        )

        return df

    # ==========================================================
    # Attendance Status
    # ==========================================================

    def calculate_status(self, df):

        if df.empty:
            return df

        df[STATUS_COLUMN] = df[TOTAL_COLUMN].apply(

            lambda x:

            ATTENDED

            if float(x) > 0

            else ABSENT

        )

        return df

    # ==========================================================
    # Competition Ranking
    # ==========================================================

    def calculate_rank(self, df):

        if df.empty:
            return df

        return apply_competition_rank(df)

    # ==========================================================
    # Prepare DataFrame
    # ==========================================================

    def prepare_dataframe(self, df):

        if df.empty:
            return df

        df = self.clean_register_numbers(df)

        self.check_duplicate_register_numbers(df)

        df = self.calculate_percentage(df)

        df = self.calculate_status(df)

        df = self.calculate_rank(df)

        df = df.reset_index(drop=True)

        return df

    # ==========================================================
    # Process
    # ==========================================================

    def process(self, uploaded_files):

        self.merge(uploaded_files)

        self.advanced_df = self.prepare_dataframe(

            self.advanced_df

        )

        self.intermediate_df = self.prepare_dataframe(

            self.intermediate_df

        )

        return (

            self.advanced_df,

            self.intermediate_df

        )

    # ==========================================================
    # Summary
    # ==========================================================

    def summary(self):

        return {

            "Advanced Students": len(

                self.advanced_df

            ),

            "Intermediate Students": len(

                self.intermediate_df

            ),

            "Advanced Present":

                len(

                    self.advanced_df[

                        self.advanced_df[STATUS_COLUMN]

                        == ATTENDED

                    ]

                ),

            "Intermediate Present":

                len(

                    self.intermediate_df[

                        self.intermediate_df[STATUS_COLUMN]

                        == ATTENDED

                    ]

                )

        }