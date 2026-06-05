# Testing Guide

## Prerequisites

- Python 3.8 or newer
- Optional virtual environment
- Install any required packages for the script you plan to run

## Recommended Setup

```bash
cd /d d:\Exercises\Python\Pipeline_Testing
python -m venv .venv
.\.venv\Scripts\activate
```

## Manual Validation Steps

1. Verify the sample and test data:
   - `sample_data/`
   - `Test_Data/`

2. Run individual validation scripts:
   - Header checkers: `header_checker_v1.py`, `header_checker_v2.py`, `header_checker_v1.1.py`, `header_checker_v1.3.py`
   - Data ETL: `Data_ETL/extract.py`, `Data_ETL/transform.py`, `Data_ETL/upload_database.py`

3. Example command:

```bash
python header_checker_v1.py
```

4. Use the GUI tools if needed:
   - `header_validator_enterprise_v2/main.py`
   - `header_validator_project/main.py`
   - `Pipeline_with_GUI/gui.py`

## Expected Results

- Scripts should complete without unhandled exceptions.
- Output files generated in `output/` or `Output_Folder/` should match the expected structure for the selected workflow.
- Validation report files should contain header validation and pipeline status details.

## Notes

- This repository does not include a formal automated test suite.
- Use sample files from `sample_data/` and `Test_Data/` to verify logic.
- If additional dependencies are required, install them in the virtual environment before running the script.
