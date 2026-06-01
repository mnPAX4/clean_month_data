# R&D Clock-In Report — Setup & Run Instructions

This guide explains how to run `clean_month_data.py` on **any computer** to turn a raw Badge Access Log Excel file into an R&D Employee Clock-In Times report.

---

## What you need

1. **Python 3.10 or newer**  
   Download: https://www.python.org/downloads/  
   During install on Windows, check **“Add python to PATH”**.

2. **Two Python packages** (pandas and openpyxl)

3. **The script file**  
   `clean_month_data.py` (keep this file on the computer)

4. **Raw badge data**  
   An `.xlsx` export from your badge system with columns:
   - `Time` — e.g. `May 11, 2026, 06:50:35`
   - `Actor User` — employee full name  

   If row 1 is junk (`Column1`, `Column2`, …), the script skips it automatically.

---

## One-time setup (each computer)

Open a terminal:

- **Windows:** Press `Win + R`, type `cmd` or `powershell`, press Enter  
- **Mac:** Terminal app  
- **Linux:** Terminal

### Step 1 — Go to the folder that contains the script

Replace the path with wherever you saved `clean_month_data.py`:

**Windows (PowerShell or Command Prompt):**
```powershell
cd C:\Users\YourName\work
```

**Mac / Linux:**
```bash
cd /Users/YourName/work
```

### Step 2 — Install dependencies (once per computer)

```bash
pip install -r requirements-clean_month_data.txt
```

If `pip` is not found, try:

```bash
python -m pip install -r requirements-clean_month_data.txt
```

---

## Run the script

### Basic usage (one raw file)

Put the **full path** to your raw badge Excel file in the command below.

**Windows:**
```powershell
python clean_month_data.py "C:\Users\YourName\Downloads\Badge Access Logs - 5-11 to 5-22.xlsx"
```

**Mac:**
```bash
python3 clean_month_data.py "/Users/YourName/Downloads/Badge Access Logs - 5-11 to 5-22.xlsx"
```

**Tip (Windows):** You can type `python clean_month_data.py ` then drag the `.xlsx` file into the terminal to paste its path. Keep the quotes if the path has spaces.

### Choose where the output is saved

By default, the report is saved in the **same folder as the raw file**. To pick a different folder:

**Windows:**
```powershell
python clean_month_data.py "C:\Users\YourName\Downloads\Badge Access Logs - 5-11 to 5-22.xlsx" --output-dir "C:\Users\YourName\Downloads"
```

**Mac:**
```bash
python3 clean_month_data.py "/Users/YourName/Downloads/badge_may.xlsx" --output-dir "/Users/YourName/Downloads"
```

### Multiple raw files for one period

If a month was exported as more than one file, list them all:

```powershell
python clean_month_data.py "C:\path\to\log_part1.xlsx" "C:\path\to\log_part2.xlsx" --output-dir "C:\path\to\output"
```

---

## Output file

The script creates a file named from the **dates in the raw data**:

```
R&D Employee Clock-In Times - MM-DD-YYYY to MM-DD-YYYY.xlsx
```

Example:
```
R&D Employee Clock-In Times - 05-11-2026 to 05-22-2026.xlsx
```

### Sheets inside the workbook

| Sheet | Contents |
|--------|-----------|
| **Daily Detail** | First badge scan per employee per day, plus signed `MinutesLate` (negative = early) |
| **Averages** | Employee, Average Clock-In Time, Average Minutes Late, Scheduled Start Time |
| **Missing Days** | Weekdays in the date range with no badge scan |

---

## Monthly workflow

1. Export badge access logs for the month (`.xlsx`).
2. Open a terminal and `cd` to the folder with `clean_month_data.py`.
3. Run:
   ```powershell
   python clean_month_data.py "FULL\PATH\TO\your_raw_badge_log.xlsx"
   ```
4. Open the generated `R&D Employee Clock-In Times - ... .xlsx` file.

Repeat each month with that month’s raw export.

---

## Troubleshooting

| Problem | What to do |
|--------|------------|
| `'python' is not recognized` | Install Python and enable “Add to PATH”, or use `py clean_month_data.py ...` on Windows |
| `No module named 'pandas'` | Run `pip install -r requirements-clean_month_data.txt` again |
| `Badge log not found` | Check the file path; use quotes around paths with spaces |
| `missing columns` | Raw file must have `Time` and `Actor User` columns |
| `No R&D badge records found` | Names in the log must match the R&D roster in the script, or add employees to `SCHEDULED_START_TIMES` in `clean_month_data.py` |
| Output averages differ from an older report | Reports only cover the dates **in that raw file**. A May-only export will not match a March–May report |

---

## Updating employee schedules

Edit `SCHEDULED_START_TIMES` near the top of `clean_month_data.py` if an employee’s scheduled start time changes or someone new joins R&D.

---

## Quick copy-paste template

```powershell
cd PATH\TO\FOLDER\WITH\SCRIPT
pip install -r requirements-clean_month_data.txt
python clean_month_data.py "PATH\TO\RAW\Badge Access Logs.xlsx"
```

Replace both paths, run in terminal, then open the new `R&D Employee Clock-In Times - ... .xlsx` file.
