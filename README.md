# SPC Generator

**Automated Process Capability & Attribute Analysis Tool**

The SPC Generator is an automation tool designed for quality and manufacturing engineering applications. It processes raw inspection data from Excel files—handling both numerical measurements and pass/fail attributes—and generates formatted PDF reports, control charts, and statistical summaries.

## Features

* **Graphical Interface:** A windowed dashboard for managing files, configuring project directories, and reviewing execution logs.
* **Batch Processing:** Select and process multiple Excel data files simultaneously.
* **Variable & Attribute Data Support:** * *Numeric:* Calculates standard capability metrics (Cp, Cpk, Mean, Standard Deviation).
    * *Attribute:* Detects text strings (e.g., "Pass", "Fail") and calculates yield and failure rates.
* **Anomaly Detection:** Evaluates continuous data against standard WECO rules to identify statistical instability and trends.
* **Standardized Reporting:** Outputs a PDF "White Paper" containing bell curves for numerical data, bar charts for attribute data, and an executive summary table. Modified Excel files are also generated with color-coded results and formatted control charts.

## Installation

1. Install [Python 3.9+](https://www.python.org/downloads/). Ensure that **"Add to PATH"** is selected during the installation process.
2. Run **`install.bat`**. This script will create a local virtual environment (`venv`) and install all necessary dependencies without affecting your global Python configuration. Wait for the success confirmation before proceeding.

## Usage Guide

### 1. Preparing the Data
1. Launch the tool using **`Launch_GUI.bat`**.
2. Click **Get New Template** to generate `SPC-DATA_Input_Template.xlsx` in your current directory.
3. Open the template and populate it with your measurement data. 
    * The file name must retain the `SPC-DATA_` prefix to be recognized by the tool (e.g., `SPC-DATA_Lot123.xlsx`).
    * You can define limits either via absolute values (USL/LSL) or tolerances (Upper/Lower Tolerance applied to a Nominal).
    * Advanced: Typing `SPLIT` into any data cell will break the output control chart line at that index, which is useful for indicating a process or tool change.

### 2. Running the Analysis
1. In the GUI, click **Refresh File List** to load your newly saved Excel files.
2. Under the Project Configuration section, specify a **New Name** to create a fresh output directory, or select an existing project from the dropdown to append data.
3. Select the files you wish to process from the list and click **Run Analysis**.
4. Review the generated PDF and Excel reports in the `output/<Project_Name>/` directory.

## Mathematical Summary

### Variable Data (Numeric)
The tool evaluates the input data as a single continuous population.
* **Standard Deviation ($\sigma$):** Calculated as the Sample Standard Deviation of the entire dataset (overall variation). 
  *Note: Because this utilizes the total standard deviation rather than within-subgroup variation ($\bar{R}/d_2$), the reported "Cpk" is statistically equivalent to Process Performance (Ppk).*
* **Process Capability ($C_{pk}$):**
  * $C_{pu} = \frac{USL - \bar{x}}{3\sigma}$
  * $C_{pl} = \frac{\bar{x} - LSL}{3\sigma}$
  * $C_{pk} = \min(C_{pu}, C_{pl})$

### Attribute Data (Pass/Fail)
Columns containing non-numeric text are processed in Attribute Mode.
* **Data Transformation:** `PASS`, `OK`, `GOOD` evaluate to **1**. `FAIL`, `NOK`, `BAD` evaluate to **0**.
* **Failure Rate:** $\left( \frac{\text{Count of Fails}}{\text{Total Samples}} \right) \times 100$

## Repository Structure

* `Launch_GUI.bat` - Application entry point.
* `SPC_Tool_GUI.pyw` - Frontend graphical interface logic.
* `install.bat` - Environment setup script.
* `get-template.bat` - Command-line alternative for generating a blank Excel template.
* `archive/run.bat` - Legacy CLI runner.
* `src/` - Core Python backend containing the SPC calculation engine and PDF generation logic.
* `output/` - Target directory for generated PDF and Excel reports (created dynamically).

## Troubleshooting

* **Application immediately closes / Window does not appear:** Ensure `install.bat` was run successfully. Check the root directory for a `crash_log.txt` file for detailed traceback information.
* **"Python not found" error:** Python is either not installed or was not added to your system PATH. Re-run the Python installer and check the corresponding box.
* **"Permission Error" / "Access Denied":** Ensure the target Excel files are not currently open in Microsoft Excel or another application while the tool is running. Close the files and re-run.
* **Files not appearing in list:** Input files must be located in the same directory as the executable scripts and must begin with the `SPC-DATA_` prefix.
