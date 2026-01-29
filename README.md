# FuzzWord - Business Partner Name Matching Tool

A Python tool for matching customer names between Oracle and SAP systems using fuzzy string matching. It finds the best matching names even when they're spelled differently or have variations.

---

## Table of Contents

- [Features](#features)
- [How It Works](#how-it-works)
- [System Requirements](#system-requirements)
- [Installation](#installation)
  - [Step 1: Install Python](#step-1-install-python)
  - [Step 2: Download the Project](#step-2-download-the-project)
  - [Step 3: Create Virtual Environment](#step-3-create-virtual-environment)
  - [Step 4: Activate Virtual Environment](#step-4-activate-virtual-environment)
  - [Step 5: Install Dependencies](#step-5-install-dependencies)
- [Preparing Your Data](#preparing-your-data)
- [Running the Program](#running-the-program)
- [Understanding the Output](#understanding-the-output)
- [Score Interpretation Guide](#score-interpretation-guide)
- [Troubleshooting](#troubleshooting)
- [Deactivating Virtual Environment](#deactivating-virtual-environment)
- [Technical Details](#technical-details)

---

## Features

- **Fuzzy Matching** - Matches names even with spelling variations
- **Top 5 Results** - Returns the 5 most similar matches for each record
- **Similarity Score** - Provides confidence scores (0-100) for each match
- **Smart Text Cleaning** - Removes common business prefixes/suffixes before comparison
- **Progress Tracking** - Shows real-time progress with time estimation
- **Excel Support** - Reads from and writes to Excel files

---

## How It Works

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   Oracle Data   │────▶│  Fuzzy Matching  │────▶│  Match Results  │
│   (Excel)       │     │    Algorithm     │     │    (Excel)      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                        │
        │                        │
        ▼                        ▼
┌─────────────────┐     ┌──────────────────┐
│    SAP Data     │     │   Text Cleaning  │
│    (Excel)      │     │   & Comparison   │
└─────────────────┘     └──────────────────┘
```

1. **Load Data** - Reads Oracle and SAP customer names from Excel
2. **Clean Text** - Removes punctuation, business terms, and normalizes text
3. **Match Names** - Uses token sort ratio algorithm to find similar names
4. **Generate Output** - Creates Excel file with top 5 matches and scores

---

## System Requirements

- **Python** 3.7 or higher
- **Operating System**: Windows 10/11, macOS, or Linux
- **Excel file** with your source data (`data_migration.xlsx`)

---

## Installation

### Step 1: Install Python

#### Check if Python is already installed

**Windows (Command Prompt or PowerShell):**
```cmd
python --version
```

**macOS/Linux (Terminal):**
```bash
python3 --version
```

If you see a version number like `Python 3.10.0`, Python is installed. If not, follow the instructions below.

#### Installing Python

**Windows:**
1. Go to [python.org/downloads](https://www.python.org/downloads/)
2. Click "Download Python 3.x.x" (latest version)
3. Run the installer
4. **IMPORTANT**: Check the box "Add Python to PATH" at the bottom
5. Click "Install Now"
6. After installation, restart your Command Prompt

**macOS:**
```bash
# Option 1: Using Homebrew (recommended)
brew install python3

# Option 2: Download from python.org
# Go to https://www.python.org/downloads/macos/
# Download and run the installer
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

---

### Step 2: Download the Project

**Option A: If you have the project folder:**
- Simply copy the folder to your desired location

**Option B: If using Git:**

**Windows:**
```cmd
cd C:\Users\YourUsername\Desktop
git clone <repository-url> fuzzword
cd fuzzword
```

**macOS/Linux:**
```bash
cd ~/Desktop
git clone <repository-url> fuzzword
cd fuzzword
```

---

### Step 3: Create Virtual Environment

A virtual environment keeps project dependencies isolated from other Python projects.

**Windows:**
```cmd
cd C:\Users\YourUsername\Desktop\fuzzword
python -m venv .venv
```

**macOS/Linux:**
```bash
cd ~/Desktop/fuzzword
python3 -m venv .venv
```

> **Note**: `.venv` is a hidden folder that will be created in your project directory.

---

### Step 4: Activate Virtual Environment

**Windows (Command Prompt):**
```cmd
.venv\Scripts\activate
```

**Windows (PowerShell):**
```powershell
.venv\Scripts\Activate.ps1
```

> **PowerShell Execution Policy Error?** Run this first:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

**macOS/Linux:**
```bash
source .venv/bin/activate
```

**How to know it's working:**
When activated, you'll see `(.venv)` at the beginning of your command line:
```
(.venv) C:\Users\YourUsername\Desktop\fuzzword>     # Windows
(.venv) user@computer:~/Desktop/fuzzword$           # macOS/Linux
```

---

### Step 5: Install Dependencies

With the virtual environment activated, install all required packages:

**Windows:**
```cmd
pip install -r requirements.txt
```

**macOS/Linux:**
```bash
pip install -r requirements.txt
```

**Alternative: Install packages manually:**
```bash
pip install pandas thefuzz openpyxl python-Levenshtein
```

---

## Preparing Your Data

Create an Excel file named `data_migration.xlsx` in the same folder as the script.

The file must have **2 sheets**:

### Sheet 1: "Oracle"

| Column | Description | Example |
|--------|-------------|---------|
| `ID` | Customer ID in Oracle | `ORC001` |
| `Name1` | Name part 1 | `ABC Company` |
| `Name2` | Name part 2 (optional) | `Limited` |

### Sheet 2: "SAP"

| Column | Description | Example |
|--------|-------------|---------|
| `BP_Number` | Business Partner ID in SAP | `SAP10001` |
| `Name1` | Name part 1 | `ABC Co.` |
| `Name2` | Name part 2 (optional) | `Ltd` |

**Example Excel structure:**

```
data_migration.xlsx
├── Sheet: Oracle
│   ├── ID        │ Name1           │ Name2
│   ├── ORC001    │ ABC Company     │ Limited
│   ├── ORC002    │ XYZ Trading     │ Co., Ltd.
│   └── ...
│
└── Sheet: SAP
    ├── BP_Number │ Name1           │ Name2
    ├── SAP10001  │ ABC Co.         │ Ltd
    ├── SAP10002  │ XYZ Trade       │ Company
    └── ...
```

---

## Running the Program

### Step-by-Step Guide

**1. Open Terminal/Command Prompt**

**Windows:**
- Press `Win + R`, type `cmd`, press Enter
- Or search for "Command Prompt" in Start menu

**macOS:**
- Press `Cmd + Space`, type "Terminal", press Enter
- Or go to Applications > Utilities > Terminal

**2. Navigate to the project folder**

**Windows:**
```cmd
cd C:\Users\YourUsername\Desktop\fuzzword
```

**macOS/Linux:**
```bash
cd ~/Desktop/fuzzword
```

**3. Activate virtual environment (if not already active)**

**Windows:**
```cmd
.venv\Scripts\activate
```

**macOS/Linux:**
```bash
source .venv/bin/activate
```

**4. Run the program**

**Windows:**
```cmd
python migrate_script.py
```

**macOS/Linux:**
```bash
python3 migrate_script.py
```

### Expected Output

```
Loading data from Excel...
Cleaning data and preparing Search Keys...
Starting search for top 5 matches (Total: 1000 records)...
Completed 50/1000 | Elapsed: 0.5 min | Est. remaining: 9.5 min
Completed 100/1000 | Elapsed: 1.0 min | Est. remaining: 9.0 min
...
Saving results to file...
Complete! Total time: 10.25 minutes
Results saved to: Match_Result_Final.xlsx
```

---

## Understanding the Output

The program creates `Match_Result_Final.xlsx` with the following columns:

| Column | Description |
|--------|-------------|
| `Oracle_ID` | Original customer ID from Oracle |
| `Oracle_Name` | Full name from Oracle (Name1 + Name2) |
| `Match_1_BP_Number` | Best match - SAP Business Partner ID |
| `Match_1_SAP_Name` | Best match - SAP customer name |
| `Match_1_Score` | Best match - Similarity score (0-100) |
| `Match_2_BP_Number` | 2nd best match - SAP BP ID |
| `Match_2_SAP_Name` | 2nd best match - SAP name |
| `Match_2_Score` | 2nd best match - Score |
| `Match_3_*` | 3rd best match |
| `Match_4_*` | 4th best match |
| `Match_5_*` | 5th best match |

---

## Score Interpretation Guide

| Score Range | Meaning | Recommendation |
|-------------|---------|----------------|
| **90-100** | Excellent match | High confidence - recommended to use |
| **80-89** | Good match | Should verify before using |
| **70-79** | Moderate match | Needs careful review |
| **Below 70** | Poor match | Likely not the same entity |

---

## Troubleshooting

### Error: "File not found: data_migration.xlsx"

**Cause:** The input file is missing or in the wrong location.

**Solution:**
1. Make sure `data_migration.xlsx` is in the same folder as `migrate_script.py`
2. Check the file name is spelled correctly (case-sensitive on macOS/Linux)

```bash
# List files in current directory
# Windows:
dir

# macOS/Linux:
ls -la
```

---

### Error: "Sheet 'Oracle' not found" or "Sheet 'SAP' not found"

**Cause:** Sheet names in the Excel file don't match expected names.

**Solution:**
1. Open `data_migration.xlsx` in Excel
2. Check that sheets are named exactly "Oracle" and "SAP" (case-sensitive)
3. Rename sheets if needed

---

### Error: "KeyError: 'ID'" or "KeyError: 'Name1'"

**Cause:** Required columns are missing in your Excel file.

**Solution:**
Verify your Excel file has these exact column names:

- **Oracle sheet:** `ID`, `Name1`, `Name2`
- **SAP sheet:** `BP_Number`, `Name1`, `Name2`

---

### Error: "python is not recognized as a command"

**Cause:** Python is not installed or not added to PATH.

**Solution (Windows):**
1. Reinstall Python from [python.org](https://www.python.org)
2. During installation, check "Add Python to PATH"
3. Restart Command Prompt

**Solution (macOS/Linux):**
Use `python3` instead of `python`:
```bash
python3 migrate_script.py
```

---

### Error: "No module named 'pandas'" or similar

**Cause:** Dependencies are not installed or virtual environment is not activated.

**Solution:**
1. Make sure virtual environment is activated (you should see `(.venv)` in your prompt)
2. Install dependencies again:
```bash
pip install -r requirements.txt
```

---

### PowerShell: "cannot be loaded because running scripts is disabled"

**Cause:** PowerShell execution policy blocks scripts.

**Solution:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Program is very slow

**Cause:** Large datasets take longer to process.

**Tips:**
- The program shows progress every 50 records
- Processing time depends on dataset size
- Consider splitting large datasets into smaller batches

---

## Deactivating Virtual Environment

When you're done, deactivate the virtual environment:

**Windows/macOS/Linux:**
```bash
deactivate
```

The `(.venv)` prefix will disappear from your prompt.

---

## Technical Details

### Algorithm Used

The program uses **Token Sort Ratio** from the FuzzyWuzzy library:
- Splits strings into tokens (words)
- Sorts tokens alphabetically
- Compares sorted strings
- This handles word order differences (e.g., "ABC Company" matches "Company ABC")

### Text Cleaning

Before comparison, text is cleaned by:
1. Converting to lowercase
2. Removing periods (`.`) and spaces
3. Removing common Thai business terms:
   - Company prefixes: `บริษัท`, `บจก`, `หจก`, `บมจ`, `หสน`
   - Suffixes: `จำกัด`
   - Titles: `คุณ`, `นาง`, `นาย`
   - Other: `ร้าน`

### Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pandas | 2.3.3 | Data manipulation |
| thefuzz | 0.22.1 | Fuzzy matching |
| python-Levenshtein | 0.27.3 | Fast string distance |
| openpyxl | 3.1.5 | Excel file handling |
| RapidFuzz | 3.14.3 | High-performance matching |
| numpy | 2.4.1 | Numerical computations |

---

## Quick Reference Card

### Windows Commands
```cmd
# Navigate to folder
cd C:\Users\YourUsername\Desktop\fuzzword

# Create virtual environment
python -m venv .venv

# Activate virtual environment
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run program
python migrate_script.py

# Deactivate when done
deactivate
```

### macOS/Linux Commands
```bash
# Navigate to folder
cd ~/Desktop/fuzzword

# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run program
python3 migrate_script.py

# Deactivate when done
deactivate
```

---

## License

This project is provided as-is for internal business use.

---

## Support

If you encounter issues not covered in this guide, please check:
1. Python version (`python --version` or `python3 --version`)
2. All dependencies are installed (`pip list`)
3. Input file format matches the required structure
