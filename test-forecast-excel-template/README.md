# Test Design & Execution Forecast Template

This project generates an Excel template for planning and forecasting test design and execution effort.

## What it creates

An `.xlsx` workbook with the following sheets:

- `Instructions`: How to use the template
- `Inputs`: Assumptions for effort and capacity
- `Test_Plan`: Planned test cases and design status
- `Execution_Tracking`: Execution outcomes and defect references
- `Forecast_Summary`: Auto-calculated KPIs and forecast metrics

## Quick start

1. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

2. Generate the workbook:

   ```powershell
   python src/build_template.py
   ```

3. Open output file:

   - `output/Test_Design_Execution_Forecast_Template.xlsx`

## Notes

- The workbook includes formulas and sample rows.
- You can rerun the script at any time to regenerate the template.
