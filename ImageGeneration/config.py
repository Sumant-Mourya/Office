import os

# API Settings
GEMINI_API_KEY = ""
MODEL_ID = "models/gemini-2.5-flash-image"

# Excel column names - DOUBLE CHECK THESE IN YOUR EXCEL
COL_NAME = "imagename"
COL_LINK = "imageurl"  # The column with the image URL (e.g., 'Image_URL')
COL_PROMPT = "prompt"    # The column with the prompt text

TOTAL_BATCHES = 0   # Number of batches to submit at once

# Path Settings
WORKBOOK_PATH = "products aditi.xlsx"
SHEET_NAME = "Sheet1"
OUTPUT_DIR = "images"

# Image settings
TARGET_WIDTH = 1000
TARGET_HEIGHT = 1000

