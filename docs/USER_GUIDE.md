# SPC Generator - User Guide

This guide will help you install, configure, and master the SPC Generator tool.

---

## 1. Installation

### Prerequisites
* **Computer:** Windows 10 or 11.
* **Software:** Python 3.9 or newer.
* **Permissions:** You do not need Admin rights to run the tool, but you need to be able to run batch scripts (`.bat`).

### First-Time Setup
1.  **Install Python:** Download from [python.org](https://www.python.org/).
    * ⚠️ **CRITICAL:** On the first screen of the installer, check the box: **"Add Python to PATH"**.
2.  **Initialize Tool:**
    * Open the tool folder.
    * Double-click `install.bat`.
    * Wait for the green "SUCCESS" message.
    * *Note: If the window closes immediately with an error, you likely missed the "Add to PATH" step in Python.*

---

## 2. Preparing Your Data

The tool uses a strict template to understand your data.

### Step A: Generate the Template
Double-click `get-template.bat`. A file named `SPC-DATA_Input_Template.xlsx` will appear.

### Step B: The Rules of the Input File
1.  **Filename:** You **MUST** rename the file to start with `SPC-DATA_`.
    * ✅ `SPC-DATA_Lot505.xlsx`
    * ❌ `Lot505_Data.xlsx`
2.  **Metadata (Rows 2-5):** Fill out the Part Number, Batch, and Date. This information appears on the PDF report header.
3.  **Columns (Features):** Each column represents one dimension (e.g., "Outer Diameter", "Length", "Hardness").

### Step C: Defining Limits (Smart Detection)
The tool is smart enough to handle two types of limits.

**Option 1: Absolute Limits (Standard)**
* **Nominal:** `10.00`
* **USL:** `10.05`
* **LSL:** `9.95`

**Option 2: Tolerances (Shortcut)**
* **Nominal:** `10.00`
* **USL:** `0.05` (Tool calculates this as 10.05)
* **LSL:** `-0.05` (Tool calculates this as 9.95)

> **Pro Tip:** Set "Subgroup Size" to 5 for standard manufacturing analysis.

---

## 3. Running the Analysis

1.  **Close Excel:** Ensure all your data files are closed. The tool cannot read open files.
2.  **Launch:** Double-click `run.bat`.
3.  **Select Files:**
    * Use **Up/Down arrows** to navigate.
    * Press **Spacebar** to toggle files selection.
    * Press **Enter** to confirm.
4.  **Name Project:** Enter a unique name (e.g., `Validation_Run_01`). This will be the name of the folder where your results are saved.

---

## 4. Interpreting Results

Go to the `output/Your_Project_Name/` folder. You will see two files for every input:

### A. The PDF Report (`..._REPORT.pdf`)
This is your "White Paper."
* **Executive Summary:** A quick table showing every feature and its Pass/Fail status.
* **Bell Curves:** Individual pages for every feature showing the histogram overlaid with the normal distribution curve.
* **Statistics:** Look for the **Cpk** value.
    * **Cpk > 1.33:** Process is CAPABLE (Green).
    * **Cpk < 1.33:** Process is NOT CAPABLE (Red).

### B. The Excel Results (`..._SPC-RESULTS.xlsx`)
Open this file to see the raw data and control charts.

**Understanding the Color Codes:**
The tool automatically highlights cells to help you spot issues quickly:
* <span style="background-color: #FFC7CE; color: #9C0006; font-weight: bold; padding: 2px;">RED Cell</span>: **Out of Tolerance.** The measurement is actually bad scrap.
* <span style="background-color: #FFEB9C; color: #9C6500; font-weight: bold; padding: 2px;">YELLOW Cell</span>: **Warning.** The data is statistically unstable or close to the limit, but not yet scrap.
* **Pattern / Status Column:** This tells you *why* a feature failed.
    * *Example:* "WECO Rule 1" means a single point jumped way outside the expected variation.
    * *Example:* "Trend" means 6 points in a row were constantly increasing (tool wear?).

---

## 5. Troubleshooting

| Error Message | Solution |
| :--- | :--- |
| **"Python is not found"** | Re-install Python and ensure "Add to PATH" is checked. |
| **"Access Denied"** | You are trying to run the code from a zipped folder. Unzip the folder first. |
| **"Permission Error"** | You have the Excel file open. Close it and try again. |
| **"No data files found"** | Your files must start with `SPC-DATA_` and be in the same folder as `run.bat`. |

---

## 6. Advanced: "Split" Data
If you have a gap in production or a tool change, you can type the word `SPLIT` into any data cell in your input Excel. The tool will break the control chart line at that point to visually indicate a process change.