# CSV Trial Extractor - Usage Guide

## Quick Start

### Running the Program

```bash
python csv_trial_extractor.py
```

Or with the virtual environment:
```bash
.venv\Scripts\python.exe csv_trial_extractor.py
```

## Step-by-Step Instructions

### Tab 1: File Selection

1. **Select Catalog File**: Click "Browse..." and select `CMF_Caralogue.xlsx`
2. **Enter Sheet Name**: Enter the sheet name (e.g., "data") and click "Load Sheet"
3. **Select Columns**:
   - **File Name Column**: Choose the column containing CSV filenames
   - **Experiment Type Column**: Choose the column with experiment types (e.g., Column F)
   - **Select Experiment Type**: Choose which experiment type to process

### Tab 2: Trial Configuration

1. **Select Trial Separator**: Choose the state marker that identifies trial starts (e.g., "MagEntry")
2. **View Num_cat Values**: See unique values associated with the separator
3. **Click "Find Trials"**: Verify trial detection with count and line numbers

### Tab 3: Marker Configuration

Configure up to 15 markers to track:
- **State**: Select the state to track from the dropdown
- **Reward?**: Check if this marker has an associated reward
- **Reward State**: If reward checked, select the reward state

Leave unused marker rows empty.

### Execute Processing

Click **"Start Data Crunching"** at the bottom right to begin processing.

## Output Files

### Individual Excel Files
Located in `processed_data/` folder with three sheets:
- **raw**: Original CSV data (8 columns, header skipped)
- **trial**: Extracted trial data with markers
- **header**: First 11 rows from original CSV

### Aggregated Excel File
Format: `[catalog_name]_[sheet_name]_[experiment_type].xlsx`

Contains one row per file with:
- Filename
- Status (processed/not present/error)
- For each marker: Sum of occurrences and average time

## Notes

- CSV files must use `latin-1` encoding (handles French accents)
- Trials containing 'Finish' in Cat column are excluded as incomplete
- Missing CSV files are marked as "not present" in aggregated output
- All times are in milliseconds relative to trial start

## Dependencies

- Python 3.7+
- pandas
- openpyxl
- tkinter (included in Python)

Install dependencies:
```bash
pip install pandas openpyxl
```

## Future Packaging

This program uses minimal dependencies (pandas, openpyxl, tkinter) to facilitate packaging as a standalone .exe with PyInstaller in future versions.
