import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Database
DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://neondb_owner:npg_w6orgHf8PROc@ep-cold-darkness-b5abg52r-pooler.c-7.us-east-2.aws.neon.tech/neondb?sslmode=require&channel_binding=require"
)
EXCEL_FILE = BASE_DIR / "school_data" / "Blue_Diamond_Fee_Database.xlsx"
EXCEL_PASSWORD = "BlueDiamond2026"
BACKUP_DIR = BASE_DIR / "backups"
DATA_DIR = BASE_DIR / "school_data"

# School
SCHOOL_NAME = "BLUE DIAMOND PUBLIC SCHOOL"
SCHOOL_ADDRESS = "Dadri, Gautam Budh Nagar, Uttar Pradesh"

# Security
SECRET_KEY = os.environ.get("BLUE_DIAMOND_SECRET", "CHANGE_THIS_2026_BLUE_DIAMOND_SECRET")
SESSION_HOURS = 12
MAX_LOGIN_ATTEMPTS = 3

# Fee defaults
DEFAULT_DUE_DAY = 10
DEFAULT_GRACE_DAYS = 0
DEFAULT_LATE_FEE_TYPE = "Per Day"
DEFAULT_LATE_FEE_AMOUNT = 50
DEFAULT_MAX_LATE_FEE = 1000

# Staff
DEFAULT_WORKING_DAYS = 26
DEFAULT_HALF_DAY_FACTOR = 0.5

# Ensure directories
for d in [DATA_DIR, BACKUP_DIR]:
    d.mkdir(exist_ok=True)