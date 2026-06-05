# Pipeline Testing

This repository contains Python scripts and utilities for data pipeline validation, template generation, header checking, and ETL experimentation.

## Contents

- `Data_ETL/` - Extract, transform, upload, and pipeline scripts for data processing.
- `header_validator_enterprise_v2/` - Enterprise header validation utilities and GUI interface.
- `header_validator_project/` - Project-level header validation utilities and GUI interface.
- `Pipeline_with_GUI/` - Pipeline implementation with a GUI and extraction/transform scripts.
- `sample_data/`, `Test_Data/` - Sample files for testing pipeline and validation logic.
- Individual scripts such as `create_zipfile.py`, `manysheets.py`, and `tkinter_consolidation.py`.

## Getting Started

1. Install Python 3.8 or newer.
2. (Optional) Create a virtual environment:

   ```bash
   python -m venv .venv
   .\.venv\Scripts\activate
   ```

3. Install dependencies if any are required for the selected tool.
   The repository does not include a single shared requirements file, so install packages locally as needed.

4. Run the desired script from the repository root.

## Recommended Usage

- Use `Data_ETL/` scripts for batch extraction and transformation workflows.
- Use `header_validator_enterprise_v2/` or `header_validator_project/` for header validation and GUI-driven checks.
- Use `Pipeline_with_GUI/` for a GUI-based pipeline experience.

## Testing Guide

See `TESTING.md` for instructions on how to run tests and confirm expected behavior.
