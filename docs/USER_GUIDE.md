# SPC Automation Tool - User Guide

## 1. First Time Setup
*You only need to do this once per computer.*

### A. Install Python
1. Download **Python** (version 3.9 or newer) from [python.org](https://www.python.org/downloads/).
2. Run the installer.
3. **CRITICAL:** Check the box **"Add Python to PATH"** at the bottom of the installer window before clicking "Install".

### B. Install the Tool
1. Open the folder containing these files.
2. Double-click the `install.bat` file.
3. A black window will open. It will:
   - Create a local "sandbox" (virtual environment).
   - Install the necessary libraries.
4. Wait for the green **"SUCCESS"** message, then close the window.

> **Note:** This installation happens entirely inside this folder. It does not require Admin rights or IT approval.

---

## 2. How to Use (Daily Workflow)

### Step A: Create a Data Template
1. Double-click `get-template.bat`.
2. A file named `SPC-DATA_Input_Template.xlsx` will appear in your folder.
3. Rename this file if you wish, but keep the `SPC-DATA_` prefix (e.g., `SPC-DATA_Lot104.xlsx`).

### Step B: Enter Your Data
1. Open the Excel file.
2. **Metadata (Rows 2-5):** Fill in Part Number, Batch, Date, etc.
3. **Columns:**
   - **Row 6 (Header):** Name your feature (e.g., "Outer Diameter").
   - **Row 7 (Nominal):** Target value.
   - **Row 8 (USL):** Upper Spec Limit.
   - **Row 9 (LSL):** Lower Spec Limit.
   - **Row 10 (Subgroup):** Usually 5.
   - **Rows 11+:** Enter your measurement data.
4. Save and close the file.

### Step C: Run the Reports
1. Double-click `run.bat`.
2. An interactive menu will appear:
   - **Select Files:** Use **Spacebar** to check/uncheck files. Press **Enter** to confirm.
   - **Project Name:** Type a name for this run (e.g., `Run_05`). This creates a new folder for your results.
3. The tool will process the files and generate charts.

### Step D: View Results
1. Open the `output` folder.
2. Open your specific project folder (e.g., `output/Run_05/`).
3. Open the `SPC-RESULTS_...xlsx` file.
4. Click the **"SPC_Charts"** tab at the bottom to view the histograms and tolerance tables.

---

## 3. Troubleshooting

**"Python is not found"**
* You installed Python but didn't check "Add to PATH". Re-install Python and check that box.

**"Access Denied" or Permission Errors**
* Make sure you are running `install.bat` to set up the local sandbox. Do not try to run `pip install` manually in the system terminal.

**The script crashes immediately**
* Ensure your Excel files are **closed** before running the tool.
* Ensure your data filenames start with `SPC-DATA_`.