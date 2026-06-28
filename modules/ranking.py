"""
modules/ranking.py

Author : Karthi Dass

Purpose
-------
1. Convert numeric columns safely
2. Calculate Percentage
3. Calculate Attendance Status
4. Calculate Competition Ranking

Ranking Rule
------------
Primary   : Score Obtained (Highest First)
Secondary : Time Taken (Seconds) (Lowest First)

Competition Ranking Example

Score  Time   Rank
30     120      1
30     120      1
30     135      3
29      80      4
28      90      5
"""

import pandas as pd

from config import (
    SCORE_COLUMN,
    TOTAL_COLUMN,
    TIME_COLUMN,
    PERCENTAGE_COLUMN,
    RANK_COLUMN,
    STATUS_COLUMN,
    ATTENDED,
    ABSENT
)


class RankingEngine:

    def __init__(self):
        pass

    # =====================================================
    # Safe Numeric Conversion
    # =====================================================

    def prepare_numeric_columns(self, df):

        numeric_columns = [
            SCORE_COLUMN,
            TOTAL_COLUMN,
            TIME_COLUMN
        ]

        for col in numeric_columns:

            if col in df.columns:

                df[col] = pd.to_numeric(
                    df[col],
                    errors="coerce"
                ).fillna(0)

        return df

    # =====================================================
    # Percentage
    # =====================================================

    def calculate_percentage(self, df):

        percentage = []

        for _, row in df.iterrows():

            score = row[SCORE_COLUMN]
            total = row[TOTAL_COLUMN]

            if total == 0:
                percentage.append(0)

            else:
                percentage.append(
                    round((score / total) * 100, 2)
                )

        df[PERCENTAGE_COLUMN] = percentage

        return df

    # =====================================================
    # Attendance Status
    # =====================================================

    def calculate_status(self, df):

        status = []

        for _, row in df.iterrows():

            if row[TOTAL_COLUMN] > 0:
                status.append(ATTENDED)
            else:
                status.append(ABSENT)

        df[STATUS_COLUMN] = status

        return df

    # =====================================================
    # Competition Ranking
    # =====================================================

    def competition_rank(self, df):

        """
        Sort

        Score Desc

        Time Asc
        """

        attended_mask = df[TOTAL_COLUMN] > 0

        attended = df[attended_mask].copy()
        absent = df[~attended_mask].copy()

        attended = attended.sort_values(

            by=[
                SCORE_COLUMN,
                TIME_COLUMN
            ],

            ascending=[
                False,
                True
            ]

        ).reset_index(drop=True)

        ranks = []

        previous_score = None
        previous_time = None
        previous_rank = 0

        for index in range(len(attended)):

            score = attended.at[index, SCORE_COLUMN]
            time = attended.at[index, TIME_COLUMN]

            if previous_score is None:

                rank = 1

            elif score == previous_score and time == previous_time:

                rank = previous_rank

            else:

                rank = index + 1

            ranks.append(rank)

            previous_score = score
            previous_time = time
            previous_rank = rank

        attended[RANK_COLUMN] = ranks
        absent[RANK_COLUMN] = 0

        if absent.empty:
            return attended

        if attended.empty:
            return absent

        return pd.concat(
            [
                attended,
                absent
            ],
            ignore_index=True
        )

    # =====================================================
    # Complete Processing
    # =====================================================

    def process(self, df):

        if df.empty:
            return df

        df = self.prepare_numeric_columns(df)

        df = self.calculate_percentage(df)

        df = self.calculate_status(df)

        df = self.competition_rank(df)

        return df

    # =====================================================
    # Summary
    # =====================================================

    def summary(self, df):

        if df.empty:

            return {

                "Students": 0,

                "Present": 0,

                "Absent": 0,

                "Highest": 0,

                "Lowest": 0,

                "Average": 0

            }

        return {

            "Students": len(df),

            "Present": len(

                df[
                    df[STATUS_COLUMN] == ATTENDED
                ]

            ),

            "Absent": len(

                df[
                    df[STATUS_COLUMN] == ABSENT
                ]

            ),

            "Highest": df[SCORE_COLUMN].max(),

            "Lowest": df[SCORE_COLUMN].min(),

            "Average": round(

                df[SCORE_COLUMN].mean(),

                2

            )

        }

    # =====================================================
    # Get Top Students
    # =====================================================

    def top_students(self, df, count=10):

        if df.empty:
            return df

        return df[
            df[STATUS_COLUMN] == ATTENDED
        ].nsmallest(count, RANK_COLUMN)

    # =====================================================
    # Get Last Students
    # =====================================================

    def last_students(self, df, count=10):

        if df.empty:
            return df

        return df[
            df[STATUS_COLUMN] == ATTENDED
        ].nlargest(count, RANK_COLUMN)
