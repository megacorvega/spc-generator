# SPC Generator

![Version](https://img.shields.io/badge/version-4.4.0-blue) ![Python](https://img.shields.io/badge/python-3.9%2B-green) ![License](https://img.shields.io/badge/license-MIT-lightgrey)

**Automated Process Capability & Attribute Analysis Tool**

The SPC Generator is a local automation tool for Quality and Manufacturing Engineers. It processes raw inspection data from Excel—whether **Numerical measurements** or **Pass/Fail attributes**—and generates professional PDF reports, control charts, and statistical summaries without manual calculation.

---

## 🚀 Key Features

* **⚡ Zero-Setup Batch Processing:** Drop 1 or 100 Excel files into the folder and process them all at once.
* **📊 Variable & Attribute Support:**
    * **Numeric:** Automatically calculates Cp, Cpk, Mean, and Standard Deviation.
    * **Attribute:** Automatically detects text (e.g., "Pass", "Fail", "OK", "NOK") and calculates Failure Rates.
* **🧠 Smart Header Detection:** Fuzzy matching allows you to name columns "USL", "Upper Limit", or "High" without breaking the tool.
* **🚨 Automatic Anomaly Detection:** Scans data against **Standard WECO Rules** to detect instability and trends.
* **📑 "Audit-Ready" Reports:** Generates PDF "White Papers" with Histograms (for numeric) or Bar Charts (for attributes).

---

## 📐 Mathematical Summary & Methodology

### 1. Variable Data (Numeric)
The tool treats the input data as a single continuous population.
* **Standard Deviation ($\sigma$):** The tool calculates the **Sample Standard Deviation** of the entire dataset (Overall Variation).
    $$\hat{\sigma} = \sqrt{\frac{\sum(x_i - \bar{x})^2}{N-1}}$$
    *Note: Because this uses the total standard deviation rather than within-subgroup variation ($\bar{R}/d_2$), the reported "Cpk" is statistically equivalent to **Ppk** (Process Performance).*

* **Process Capability ($C_{pk}$):**
    $$C_{pu} = \frac{USL - \bar{x}}{3\sigma}$$
    $$C_{pl} = \frac{\bar{x} - LSL}{3\sigma}$$
    $$C_{pk} = \min(C_{pu}, C_{pl})$$

### 2. Attribute Data (Pass/Fail)
If a column contains text (non-numeric data), the tool switches to Attribute Mode.
* **Data Transformation:**
    * `PASS`, `OK`, `GOOD` $\rightarrow$ **1**
    * `FAIL`, `NOK`, `BAD` $\rightarrow$ **0**
* **Failure Rate:**
    $$Rate (\%) = \left( \frac{\text{Count of Fails}}{\text{Total Samples}} \right) \times 100$$

---

## 🛠️ How It Works

1.  **Input:** Fill out the Excel template.
    * **Numeric:** Enter measurements (e.g., `10.05`, `10.06`).
    * **Attribute:** Enter text (e.g., `Pass`, `Fail`).
2.  **Process:** The tool scans each column.
    * It detects if limits are absolute (e.g., `10.05`) or tolerances (e.g., `0.05`) and standardizes them.
    * It determines if the feature is Numeric or Attribute based on content.
3.  **Analyze:**
    * **Numeric:** Runs standard distribution analysis and WECO rules.
    * **Attribute:** Calculates yield and plots a Step Chart (1 vs 0).
4.  **Output:** Results are saved in `output/Project_Name/`.

---

## 🏁 Quick Start

### 1. Installation
1.  Install [Python 3.9+](https://www.python.org/downloads/) (Check **"Add to PATH"**).
2.  Double-click `install.bat` to set up the local environment.

### 2. Usage
1.  **Get Template:** Run `get-template.bat`.
2.  **Enter Data:** Save your file as `SPC-DATA_YourFileName.xlsx`.
3.  **Run:** Double-click `run.bat`.

---

## 📂 Project Structure

* `input/` - Place your `SPC-DATA_*.xlsx` files here.
* `output/` - Reports generated here.
* `src/` - Core Python logic.
* `docs/` - Documentation.

---

## 🛡️ Note on Statistics
This tool uses **Overall Standard Deviation** (Long-Term).
* **Use this for:** Ppk reporting, long-term lot validation, and attribute failure rates.
* **Do not use this for:** Short-run Control Charts requiring $\bar{R}$ (Range) or Subgrouping logic.