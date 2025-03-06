# Excel VBA Sales Management Tool

## Project Overview

This Excel VBA Sales Management Tool is designed to help users efficiently manage daily sales data for personal or small business use. It supports data entry, dynamic column management, data validation, and automatic generation of visual charts for easy data analysis.

---

## Workbook Structure

The Excel file consists of three worksheets:

1. **Sheet1 - Input Page:** Interface for data entry.
2. **Sheet2 - Charts:** Automatically generates charts based on user-selected variables.
3. **Sheet3 - Database**: Stores all sales data in a structured table format named "SalesTable."

---

## Features

### 1. Add New Columns
- Users can add new columns directly through the "Input Page" by clicking the "Add Column" button and entering the desired column name.
- Newly created columns automatically reflect in the database and the input page after refreshing.

### 2. Dynamic Input Area
- Click the "Refresh Input Area" button to synchronize input fields with the current structure of the "Database." 


### 3. Batch Data Entry
- Enter data across multiple rows (rows 2-20) on the Input Page.
- Clicking the "Batch Add" button validates each entry and adds them into the database.
- Errors in data entry prompt immediate feedback for corrections.
- Successfully added data rows are cleared (excluding special formula cells).

### 4. Dynamic Chart Generation
- On the "Charts" worksheet, select variables for the X-axis, Y-axis, and chart type via drop-down lists.
- Clicking the chart generation button creates or updates the visual chart based on selections.

---

## How to Use

### Step 1: Add New Columns (Optional)
- From the Input Page, click "Add Column," then enter the new column name when prompted.

### Step 2: Refresh Input Area
- After any changes to the database structure, click "Refresh Input Area" to update column headers.

### Step 3: Data Entry and Validation
- Fill in sales data (rows 2-20). Required fields such as dates and numeric fields are validated automatically.
- Correct any errors indicated by prompts before submission.

### Step 4: Batch Submission
- After filling data, click "Batch Add" to transfer entries to the database.

### Step 5: Chart Visualization
- On the "Charts" sheet, select desired columns and chart type, then click the "Generate Chart" button to visualize your data.


