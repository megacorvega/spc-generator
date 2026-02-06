# SPC Generator

![Version](https://img.shields.io/badge/version-4.3.0-blue) ![Python](https://img.shields.io/badge/python-3.9%2B-green) ![License](https://img.shields.io/badge/license-MIT-lightgrey)

**Turn messy Excel inspection data into professional Engineering capability reports in seconds.**

The SPC Generator is a local automation tool designed for Quality and Manufacturing Engineers. It processes raw inspection data from Excel, performs statistical analysis (Cp/Cpk), and generates PDF "White Paper" style reports and visual control charts without requiring manual calculations.

---

## 🚀 Key Features

* **⚡ Zero-Setup Batch Processing:** Drop 1 or 100 Excel files into the folder and process them all at once.
* **📊 Advanced Statistics:** Automatically calculates **Cp, Cpk, Mean, and Sigma** using industry-standard within-subgroup variation (Unbiased Standard Deviation / $c_4$ constants).
* **🚨 Automatic Anomaly Detection:** Scans data against **8 standard WECO (Western Electric) Rules** to detect instability, trends, and stratification.
* **📑 "Audit-Ready" PDF Reports:** Generates a clean Executive Summary and individual Bell Curve histograms for every feature.
* **📈 Visual Excel Integration:** Outputs a new Excel file with embedded control charts, color-coded pass/fail cells, and histograms.

---

## 🛠️ How It Works

1.  **Input:** You fill out a simple Excel template with your measurements.
2.  **Process:** The tool scans for tolerance limits. It intelligently detects if you entered "Absolute Limits" (e.g., `10.05`) or "Tolerances" (e.g., `0.05`) and standardizes them.
3.  **Analyze:** It runs a full SPC scan, checking for outliers and statistical control.
4.  **Output:** You get a folder containing your reports, organized by Project Name.

---

## 🏁 Quick Start

### 1. Installation (Windows)
No command line knowledge required.
1.  Install [Python 3.9+](https://www.python.org/downloads/) (Make sure to check **"Add to PATH"** during install).
2.  Download this folder.
3.  Double-click `install.bat`. 
    * *This creates a secure, local environment for the tool.*

### 2. Usage
1.  **Get a Template:** Double-click `get-template.bat` to generate a blank input file.
2.  **Enter Data:** Fill in your measurements and save the file with the prefix `SPC-DATA_` (e.g., `SPC-DATA_Batch101.xlsx`).
3.  **Run:** Double-click `run.bat`.
4.  **Select & Go:** Follow the on-screen prompts to select your files and name your project.

---

## 📂 Project Structure

* `input/` - (Root) Place your `SPC-DATA_*.xlsx` files here.
* `output/` - Results are automatically sorted here by Project Name.
* `src/` - The core Python application logic.
* `docs/` - Detailed documentation.

---

## 🛡️ Methodology

The tool uses **within-subgroup standard deviation** (σ_within) to estimate process capability (Cpk), which is the standard for short-run manufacturing analysis.

**σ = R̄ / d₂** or   **s̄ / c₄**

*Note: This tool uses the c₄ method for higher accuracy with varying subgroup sizes.*

---

## 🤝 Contributing

See [docs/TESTING.md](docs/TESTING.md) for information on running the test suite (`pytest`).
