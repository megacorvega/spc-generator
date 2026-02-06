# SPC Generator

Statistical Process Control (SPC) automation tool for generating histograms and tolerance tables from inspection data.

## Features

* **Batch Processing:** Automatically processes multiple Excel inspection files.
* **Statistical Analysis:** Calculates ±3σ tolerance limits using within-subgroup variation (c4 constants).
* **Project Organization:** Automatically routes results into organized project folders (e.g., `output/Batch_A/`).
* **Visual Interface:** Interactive command-line menu for file selection using `rich` and `questionary`.
* **Excel Integration:** Outputs results with embedded charts and formatted tables.

---

## Quick Start

### Installation

**Option A: One-Click (Windows)**
1. Ensure Python 3.9+ is installed.
2. Double-click `install.bat`.

**Option B: Developer Mode**
Open a terminal in the project root and run:
```bash
python -m pip install -e .
```

### Usage

**1. Generate an Input Template**
Open a terminal and run:
```bash
spc-template
```
*This creates `SPC-DATA_Input_Template.xlsx`.*

**2. Fill Data**
Fill the template and save your data files with the prefix `SPC-DATA_` in your working directory (e.g., `SPC-DATA_Lot104.xlsx`).

**3. Run the Analysis**
Double-click `run.bat` (or run `spc-gen` in a terminal).

* **Select Files:** Use `Spacebar` to select files, `Enter` to confirm.
* **Name Project:** Enter a name (e.g., "Run_05").
* **View Results:** Open the `output/Run_05/` folder to see your reports.

---

## Project Structure

```text
spc-generator/
├── output/                  # Ignored by Git (Results go here)
├── src/spc_generator/       # Source Code (Package)
│   ├── generator.py         # Main logic
│   └── template.py          # Template generator
├── docs/                    # Documentation
│   ├── USER_GUIDE.md        # Instructions for End Users
│   └── TESTING.md           # Developer Guide for Tests
├── install.bat              # One-click Installer
├── run.bat                  # One-click Launcher
└── pyproject.toml           # Configuration & Dependencies
```

---

## Development

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov=spc_generator --cov-report=html
```

## Documentation

* [User Guide](docs/USER_GUIDE.md) - Instructions for non-technical users.
* [Testing Guide](docs/TESTING.md) - Deep dive into the test fixtures.