"""
config.py

Central configuration file for Assessment Automation
Author: Karthi Dass
"""

from pathlib import Path

# ==============================================================================
# PROJECT PATHS
# ==============================================================================

ROOT_DIR = Path(__file__).parent

BACKUP_DIR = ROOT_DIR / "backup"
OUTPUT_DIR = ROOT_DIR / "output"
LOG_DIR = ROOT_DIR / "logs"
MASTER_TEMPLATE_NAME = "RIT_Portal login-Assessment_2026-2027_Master.xlsx"
MASTER_TEMPLATE_PATH = ROOT_DIR / "assets" / MASTER_TEMPLATE_NAME
ECE_MASTER_TEMPLATE_NAME = "ECE - Batch (2023-27) - Portal login-Assessment.xlsx"
ECE_MASTER_TEMPLATE_PATH = ROOT_DIR / "assets" / ECE_MASTER_TEMPLATE_NAME
ECE_MAIN_SHEET_NAME = "Sheet1"

MASTER_TEMPLATES = {
    "IT": MASTER_TEMPLATE_PATH,
    "ECE": ECE_MASTER_TEMPLATE_PATH,
}

MASTER_PROFILES = {
    "IT": {
        "template_path": MASTER_TEMPLATE_PATH,
        "mode": "split",
    },
    "ECE": {
        "template_path": ECE_MASTER_TEMPLATE_PATH,
        "mode": "combined",
        "main_sheet_name": ECE_MAIN_SHEET_NAME,
        "detail_header_row": 2,
        "detail_data_start_row": 3,
        "assessment_header_row": 2,
        "assessment_start_column": 7,
    },
}

BACKUP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# ==============================================================================
# SHEET NAMES (Master Workbook)
# ==============================================================================

ADVANCED_SHEET = "Advanced"
INTERMEDIATE_SHEET = "Intermediate"

TOP10_SHEET = "Top 10 Performers"
LAST10_SHEET = "Last 10 Performers"

# ==============================================================================
# MASTER SHEET STRUCTURE
# ==============================================================================

DATE_ROW = 1          # Row where assessment date is merged
HEADER_ROW = 2        # Row where Score, Total, %, Rank, Status exist
DATA_START_ROW = 3    # Student data starts here

DATE_BLOCK_SIZE = 5

DATE_HEADERS = [
    "Score",
    "Total",
    "%",
    "Rank",
    "Status"
]

# ==============================================================================
# COLUMN NAMES (Uploaded Assessment Report)
# ==============================================================================

REGISTER_COLUMN = "Register Number"
NAME_COLUMN = "Student Name"
DEPARTMENT_COLUMN = "Department"

SCORE_COLUMN = "Score Obtained"
TOTAL_COLUMN = "Total Marks"

TIME_COLUMN = "Time Taken (Seconds)"

PERCENTAGE_COLUMN = "Percentage"
RANK_COLUMN = "College Rank"

STATUS_COLUMN = "Status"

# ==============================================================================
# STATUS VALUES
# ==============================================================================

ATTENDED = "Attended"
ABSENT = "Absent"

# ==============================================================================
# FILE NAME SETTINGS
# ==============================================================================

ADVANCED_PREFIX = "A"
INTERMEDIATE_PREFIX = "B"

# Example:
# Daily_Score_Report_Java_A1_24_06_2026.xlsx

FILENAME_SEPARATOR = "_"

# ==============================================================================
# RANKING SETTINGS
# ==============================================================================

SORT_COLUMNS = [
    SCORE_COLUMN,
    TIME_COLUMN
]

SORT_ASCENDING = [
    False,   # Score Desc
    True     # Time Asc
]

# ==============================================================================
# DEFAULT VALUES
# ==============================================================================

DEFAULT_SCORE = 0
DEFAULT_TOTAL = 0
DEFAULT_PERCENTAGE = 0
DEFAULT_RANK = 0

# ==============================================================================
# OUTPUT FILE NAME
# ==============================================================================

OUTPUT_FILENAME = "Assessment_Report.xlsx"

# ==============================================================================
# EXCEL FORMATTING
# ==============================================================================

COPY_FONT = True
COPY_BORDER = True
COPY_FILL = True
COPY_ALIGNMENT = True
COPY_NUMBER_FORMAT = True
COPY_PROTECTION = True
COPY_COLUMN_WIDTH = True
COPY_MERGED_CELLS = True

# ==============================================================================
# VALIDATION
# ==============================================================================

REQUIRED_COLUMNS = [

    REGISTER_COLUMN,

    NAME_COLUMN,

    DEPARTMENT_COLUMN,

    SCORE_COLUMN,

    TOTAL_COLUMN,

    TIME_COLUMN

]

# Master Workbook Columns
FIRST_ASSESSMENT_COLUMN = 7
REGISTER_COLUMN_INDEX = 2   # B
NAME_COLUMN_INDEX = 3        # C
