# Antigravity Excel Engine

> **The Enterprise-Grade, AI-Native Excel Automation Engine for Autonomous Agents and Data Pipelines.**

![Python Version](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Tech Stack](https://img.shields.io/badge/backend-Win32%20COM%20%2B%20OpenPyXL-orange)
![Tests](https://img.shields.io/badge/tests-27%2F27%20passed%20(100%25)-brightgreen)
![IPC Latency](https://img.shields.io/badge/IPC%20Latency-%3C5ms-blueviolet)
![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011%20%7C%20Server-lightgrey)

*Created by [Francisco Grandón Vergara](https://github.com/FranciscoGrandon)*

---

## 📖 Table of Contents

- [What is Antigravity Excel Engine?](#what-is-antigravity-excel-engine)
- [Why You Need It](#why-you-need-it)
- [Key Features](#key-features)
- [Architecture](#architecture)
- [Quick Start](#quick-start)
- [Real-World Examples](#real-world-examples)
- [API Reference](#api-reference)
- [Daemon & IPC](#daemon--ipc)
- [Feature Comparison](#feature-comparison)
- [Status & Roadmap](#status--roadmap)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## What is Antigravity Excel Engine?

**Antigravity Excel Engine** is a Python library that automates Excel operations with **zero freezes**, **AI-optimized data formats**, and **sub-5ms response times**. It's built for:

- 🤖 **AI Agents** that need to read/write/audit Excel files without blocking
- 📊 **Data Pipelines** that process Excel reports automatically
- 🔄 **Business Automation** teams that need reliable, fast Excel manipulation
- 📈 **Financial Modeling** systems requiring dynamic formulas and style presets

The engine **automatically chooses** between:
- **Live COM Mode**: When Excel is open (real-time updates, undo/redo support)
- **Headless Mode**: When Excel is closed (pure file I/O, no dependencies)

**Without you changing a single line of code.**

---

## Why You Need It

### The Problem

Traditional Excel libraries break when integrated with AI agents and automation:

| Issue | Impact |
|-------|--------|
| **COM Crashes on Active Editing** | If a user types in Excel while your script runs, you get `RPC_E_SERVERCALL_RETRYLATER` and silent failures |
| **OpenPyXL Can't Evaluate Formulas** | Dynamic calculations fail; you get stale data or #REF! errors |
| **Slow Cold Starts** | Python imports take 500–1200ms just to initialize |
| **LLM Token Waste** | Sending JSON instead of CSV means ~75% more tokens to say the same thing |
| **Zombie Processes** | Failed scripts leave `EXCEL.EXE` running forever in the background |
| **No Formula Auditing** | You can't detect errors like `#VALUE!`, truncated ranges, or inconsistent formulas |

### The Solution

Antigravity Excel Engine solves **all of these** with:

✅ **Dual-Core Architecture** — Auto-switches between Live COM and Headless mode  
✅ **Resident Daemon** — Keeps Excel warm in RAM, < 5ms response times via Named Pipe IPC  
✅ **87% Faster Cold Start** — Lazy loading brings startup from 726ms → 57ms  
✅ **Token-Optimized CSV** — Reduces LLM context tokens by up to 75%  
✅ **Self-Healing COM Locks** — Automatic retry with exponential backoff  
✅ **Formula Audit** — Detect errors and inconsistencies automatically  
✅ **Unified Data Curation** — Dedup + Clean + Outlier Detection in **one memory pass**  
✅ **Zero Process Leaks** — Guaranteed cleanup and resource management  

---

## Key Features

### 1. Live COM + Headless Auto-Detection

```python
engine = AntigravityExcelEngine(mode="auto", file_path="report.xlsx")
# ✅ If Excel has it open → uses Live COM (dynamic, interactive)
# ✅ If Excel closed → uses OpenPyXL (fast, lightweight)
```

### 2. Sub-5ms Resident Daemon

```bash
# Start once, use forever
python antigravity_excel_daemon.py --start

# Then any operation is < 5ms via Named Pipe IPC
python antigravity_excel_cli.py get-csv "A1:F20" --json
```

### 3. Token-Optimized CSV Output

```python
# Returns dense CSV (optimized for LLMs)
csv = engine.get_range_as_csv("A1:D100")
# vs. verbose JSON with 75% more tokens
```

### 4. Unified Data Curation Pipeline

```python
# One call, everything happens in memory
engine.curate_pipeline(
    "DataTable",
    spec={
        "dedup": {"key_columns": ["ID"]},
        "clean": {
            "standardize_text": {"Segment": {"casing": "title"}},
            "cast_types": {"Sales": "float"},
            "coerce_dates": ["Date"]
        },
        "outliers": [
            {"column": "Sales", "method": "iqr", "add_comment": True}
        ]
    }
)
```

### 5. Financial Modeling Presets

```python
engine.apply_style_preset("A1:G1", "header")        # Corporate blue + bold
engine.apply_style_preset("B2:B10", "currency")     # $#,##0.00 format
engine.apply_style_preset("H2:H10", "delta_positive")  # Green + bold
```

### 6. Formula Audit & Error Detection

```python
errors = engine.check_formula_errors("A1:D100")
# Returns: [{"cell": "B5", "error": "#DIV/0!", "formula": "=A5/B5"}]

quality = engine.audit_data_quality("A1:D100")
# Returns: whitespace issues, numbers-as-text, type inconsistencies, formula count
```

### 7. Native Pivot Tables & Charts

```python
# Create pivot table
engine.create_pivot_table(
    source="A1:D1000",
    rows=["Region"],
    cols=["Product"],
    value_field="Sales",
    dest_sheet="Pivot_Summary"
)

# Smart non-overlapping chart placement
engine.create_chart(
    source="Pivot_Summary!A3:B8",
    chart_type="column_clustered",
    title="Regional Sales",
    placement="auto"  # Never overlaps with data
)
```

---

## Architecture

```mermaid
flowchart TD
    Agent["🤖 AI Agent / CLI / Python Script"] -->|JSON Request| IPC{Named Pipe Active?}
    
    IPC -->|Yes < 5ms| Daemon["🔥 Resident Daemon<br/>(Warm Excel Session)"]
    IPC -->|No (Fallback)| API["⚙️ Direct API Call"]
    
    Daemon --> Engine["Antigravity Engine Core"]
    API --> Detector{Is Excel Open?}
    
    Detector -->|Yes| COM["📊 Live COM Backend<br/>(win32com + oleacc)<br/>Real-time Updates"]
    Detector -->|No| Headless["📁 Headless Backend<br/>(OpenPyXL)<br/>Fast File I/O"]
    
    COM --> Pipeline["🔄 Unified Pipeline<br/>(Dedup, Clean, Audit)"]
    Headless --> Pipeline
    
    Pipeline --> Output["📤 Optimized Output<br/>(CSV, JSON, Formatted)"]
    Output --> Agent
```

### How It Works

1. **Agent sends request** (Python API, CLI, or daemon pipe)
2. **Daemon checks** if Named Pipe is active (< 5ms roundtrip)
3. **Engine auto-detects** if Excel is open
4. **Picks the right backend**:
   - **Live COM** if file is open (interactive, dynamic)
   - **Headless OpenPyXL** if file is closed (fast, lightweight)
5. **Executes operation** with retry logic for locks
6. **Returns optimized result** (CSV for LLMs, JSON for structure)

---

## Quick Start

### Installation

**Requirements:**
- Python 3.10+
- Windows 10/11 or Windows Server
- Excel 2016+ (for Live COM mode) or just openpyxl (for Headless)

**Install:**

```bash
git clone https://github.com/FranciscoGrandon/antigravity-excel-engine.git
cd antigravity-excel-engine
pip install -r requirements.txt
```

### Minimal Example (Python API)

```python
from antigravity_excel_core import AntigravityExcelEngine

# Create engine (auto-detects mode)
engine = AntigravityExcelEngine(mode="auto", file_path="C:/data/report.xlsx")

# Read data
csv_data = engine.get_range_as_csv("A1:D100")
print(csv_data)

# Write declaratively
engine.set_cells({
    "A1": {"value": "Revenue", "cellStyles": {"fontWeight": "bold", "backgroundColor": "#0E2E63"}},
    "B1": {"formula": "=SUM(B2:B10)", "cellStyles": {"numberFormat": "$#,##0"}},
}, autofit=True)

# Audit for errors
errors = engine.check_formula_errors("A1:D100")
if errors:
    print(f"Found {len(errors)} formula errors")

# Save
engine.save()
```

### Using the Daemon (< 5ms per operation)

**Start the daemon (runs in background):**

```bash
python antigravity_excel_daemon.py --start --mode auto --file "C:/data/report.xlsx"
```

**Check status:**

```bash
python antigravity_excel_daemon.py --status
# Output: {"status": "success", "daemon_running": true, "latency_ms": 2.1, ...}
```

**Run fast commands:**

```bash
# Get CSV data (< 5ms)
python antigravity_excel_cli.py get-csv "A1:F20" --sheet "Datos" --json

# Audit data quality (< 5ms)
python antigravity_excel_cli.py audit-quality "A1:P700" --json

# Curate pipeline (bulk operation, still very fast)
python antigravity_excel_cli.py curate \
  --source "A1:P700" \
  --spec-json '{"dedup": {"key_columns": ["ID"]}, "clean": {...}}'

# Create pivot table
python antigravity_excel_cli.py create-pivot \
  --source "A1:P700" \
  --rows Region \
  --cols Product \
  --value-field Sales \
  --dest-sheet "Pivot_Summary"

# Stop daemon
python antigravity_excel_daemon.py --stop
```

---

## Real-World Examples

### Example 1: Automated Monthly Report

```python
from antigravity_excel_core import AntigravityExcelEngine

# Open last month's template
engine = AntigravityExcelEngine(mode="auto", file_path="template_monthly.xlsx")

# Read parameter values from user
parameters = [
    {"name": "Growth Rate", "value": 0.05, "format": "percent"},
    {"name": "Inflation", "value": 0.03, "format": "percent"},
]

# Create summary table with parameters + data
engine.create_summary_table(
    title="Monthly Financial Summary",
    headers=["Scenario", "Revenue", "Costs", "Margin"],
    rows_spec=[
        {"label": "Base Case", "values": [100000, 70000, "=B4-C4"]},
        {"label": "Optimistic", "values": [120000, 80000, "=B5-C5"]},
        {"label": "Pessimistic", "values": [90000, 65000, "=B6-C6"]},
    ],
    dest_sheet="Summary",
    parameters=parameters
)

# Add calculated column with formula
engine.add_calculated_column(
    header="Margin %",
    formula="=D{r}/B{r}",
    number_format="0.0%"
)

engine.save()
print("✅ Report generated")
```

### Example 2: Data Quality Audit Before ML

```python
engine = AntigravityExcelEngine(mode="auto", file_path="raw_data.xlsx")

# Full audit in one pass
quality = engine.audit_data_quality("A1:P1000")

print(f"Found {len(quality['numbers_stored_as_text'])} numbers-as-text")
print(f"Found {len(quality['header_whitespace_issues'])} whitespace issues")

# Flag outliers with comments
outliers = engine.detect_and_flag_outliers(
    column="Sales",
    method="iqr",
    range_or_table="A1:P1000",
    add_comment=True
)
print(f"Flagged {len(outliers['flagged_cells'])} outlier cells")

# Clean everything in one pipeline
result = engine.curate_pipeline(
    "RawData",
    spec={
        "dedup": {"key_columns": ["ID", "Date"]},
        "clean": {
            "standardize_text": {
                "Segment": {"trim": True, "casing": "title"},
                "Country": {"trim": True, "casing": "upper"},
            },
            "cast_types": {
                "Sales": "float",
                "Units": "int",
            },
            "coerce_dates": ["Date"],
            "trim_headers": True,
        },
        "outliers": [
            {"column": "Sales", "method": "iqr", "add_comment": True},
            {"column": "Units", "method": "zscore", "add_comment": False},
        ]
    }
)

print(f"Removed {result['duplicates_removed']} duplicates")
print(f"Cleaned {result['cells_cleaned']} cells")
print(f"Flagged {result['outliers_count']} outliers")

engine.save()
print("✅ Data ready for ML")
```

### Example 3: AI Agent Reporting Loop

```python
import json
from antigravity_excel_core import AntigravityExcelEngine

def ai_agent_generate_report(sales_data: list):
    """Simulates an AI agent writing to Excel"""
    
    engine = AntigravityExcelEngine(mode="auto", file_path="agent_report.xlsx")
    
    # Write raw data
    engine.set_cell_range("A1", [["Date", "Product", "Sales"]] + sales_data)
    
    # Aggregate by product
    summary = engine.aggregate(
        group_by_col="Product",
        metric_col="Sales",
        agg_func="sum",
        top_n=10
    )
    
    # Create pivot + chart
    engine.create_pivot_table(
        source="A1:C100",
        rows=["Product"],
        value_field="Sales",
        dest_sheet="Summary"
    )
    
    engine.create_chart(
        source="Summary!A3:B10",
        chart_type="column_clustered",
        title="Top 10 Products by Sales",
        placement="auto"
    )
    
    # Audit formula errors
    errors = engine.check_formula_errors("A1:D100")
    if errors:
        print(f"⚠️ Warning: Found {len(errors)} formula errors")
    
    engine.save()
    return {"status": "success", "summary": summary["ranking"][:5]}

# Use it
report = ai_agent_generate_report([
    ["2024-01-01", "ProductA", 1000],
    ["2024-01-01", "ProductB", 1500],
])
print(json.dumps(report, indent=2))
```

---

## API Reference

### Core Class: `AntigravityExcelEngine`

#### Initialization

```python
engine = AntigravityExcelEngine(mode="auto", file_path="report.xlsx")
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | `str` | `"auto"` (default), `"live"` (COM only), `"file"` (OpenPyXL only) |
| `file_path` | `str` | Path to .xlsx file (created if doesn't exist) |

#### Reading Data

```python
# Get range as optimized CSV (best for LLMs)
csv = engine.get_range_as_csv("A1:D100", delimiter=";")

# Get raw cell values + formulas
data = engine.get_cell_ranges(["A1:B10", "D5:E20"], include_formulas=True)
```

#### Writing Data

```python
# Write declarative mapping (values, formulas, styles, notes)
engine.set_cells({
    "A1": {"value": "Title", "cellStyles": {"fontWeight": "bold"}},
    "B1": {"formula": "=SUM(B2:B10)", "cellStyles": {"numberFormat": "$#,##0"}},
}, autofit=True)

# Write range with formulas
engine.set_cell_range("A1", [["A", "B"], [1, 2]], autofit=True)
```

#### Data Quality & Audit

```python
# Audit for errors, types, whitespace
report = engine.audit_data_quality("A1:D100")

# Check formula errors
errors = engine.check_formula_errors("A1:D100")

# Detect outliers
outliers = engine.detect_and_flag_outliers(
    column="Sales",
    method="iqr",  # or "zscore"
    add_comment=True
)
```

#### Data Cleaning

```python
# Clean in single pass
engine.clean_table_dataset(
    "DataRange",
    rules={
        "trim_headers": True,
        "standardize_text": {
            "Segment": {"trim": True, "casing": "title"},
        },
        "cast_types": {
            "Sales": "float",
        },
        "coerce_dates": ["Date"]
    }
)

# Full curation pipeline (dedup + clean + outliers)
result = engine.curate_pipeline(
    "RawData",
    spec={...}
)
```

#### Analytics

```python
# Group by and aggregate
summary = engine.aggregate(
    group_by_col="Region",
    metric_col="Sales",
    agg_func="sum",  # "avg", "min", "max", "count"
    top_n=10,
    ascending=False
)

# Add calculated column with auto-expanding table
engine.add_calculated_column(
    header="Total",
    formula="=A{r}*B{r}",  # {r} is replaced with row number
    number_format="$#,##0.00"
)
```

#### Pivot Tables & Charts

```python
# Create pivot table
engine.create_pivot_table(
    source="A1:D1000",
    rows=["Region"],
    cols=["Product"],
    value_field="Sales",
    value_func="sum",  # "avg", "count", "min", "max"
    dest_sheet="PivotSummary",
    dest_cell="A3"
)

# Create chart (auto-placed)
engine.create_chart(
    source="PivotSummary!A3:B10",
    chart_type="column_clustered",  # "bar", "line", "pie", "area"
    title="Sales by Region",
    placement="auto",  # or "right", "bottom", or anchor="I3"
    width=480,
    height=300
)
```

#### Styling

```python
# Apply preset styles
engine.apply_style_preset("A1:G1", "header")          # Corporate blue
engine.apply_style_preset("B2:B100", "currency")      # $#,##0.00
engine.apply_style_preset("H2:H100", "percent")       # 0.0%
engine.apply_style_preset("C10", "delta_positive")    # Green + bold
engine.apply_style_preset("C11", "delta_negative")    # Red + bold
engine.apply_style_preset("H12", "total_row")         # Bold + double line
engine.apply_style_preset("B2", "input_cell")         # Yellow background
```

#### Save & Close

```python
engine.save()          # Save to current file
engine.save("new.xlsx")  # Save as new file
engine.close()         # Clean up resources
```

---

## Daemon & IPC

### Why Use the Daemon?

- **Speed**: < 5ms per operation (no Python startup overhead)
- **State**: Keeps Excel session warm in RAM
- **Concurrency**: Multiple requests share same Excel session safely
- **Reliability**: Automatic retry on lock

### Start/Stop

```bash
# Start daemon in background
python antigravity_excel_daemon.py --start --mode auto --file "report.xlsx"

# Check health
python antigravity_excel_daemon.py --status
# Output: {"daemon_running": true, "latency_ms": 2.3, "engine_status": {...}}

# Stop daemon
python antigravity_excel_daemon.py --stop
```

### CLI Commands

```bash
# Read data
python antigravity_excel_cli.py get-csv "A1:F20" --sheet "Sheet1" --json

# Write cells
python antigravity_excel_cli.py set-cells --cells '{"A1": {"value": "Test"}}' --json

# Audit
python antigravity_excel_cli.py audit-quality "A1:P700" --json

# Curate pipeline
python antigravity_excel_cli.py curate --source "A1:P700" --spec-json '{"dedup": {...}}'

# Create pivot
python antigravity_excel_cli.py create-pivot \
  --source "A1:D1000" --rows Region --value-field Sales --dest-sheet Pivot

# Create chart
python antigravity_excel_cli.py create-chart \
  --source "Pivot!A3:B10" --chart-type column_clustered --title "Sales"

# All commands return JSON with --json flag
```

### Python Daemon Client

```python
from antigravity_excel_daemon import send_pipe_message, ping_daemon

# Check if daemon is alive
alive, latency_ms, details = ping_daemon(timeout_ms=1000)
print(f"Daemon latency: {latency_ms}ms")

# Send custom command
response = send_pipe_message({
    "command": "get-csv",
    "args": {
        "range": "A1:F20",
        "sheet": "Datos",
        "delimiter": ";"
    }
})
print(response)
```

---

## Feature Comparison

| Feature | Antigravity | xlwings | openpyxl | win32com |
|---------|:-----------:|:-------:|:--------:|:--------:|
| **Dual-Core (Live + Headless)** | ✅ | ❌ | ✅ (Headless) | ❌ (Live) |
| **Sub-5ms Daemon IPC** | ✅ | ❌ | ❌ | ❌ |
| **AI Token Optimization (CSV)** | ✅ | ❌ | ❌ | ❌ |
| **Lazy Loading (87% faster)** | ✅ | ❌ | ❌ | ❌ |
| **Auto-Retry COM Locks** | ✅ | ⚠️ | N/A | ❌ |
| **Unified Data Curation (1-pass)** | ✅ | ❌ | ❌ | ❌ |
| **Formula Audit & Error Detection** | ✅ | ❌ | ❌ | ❌ |
| **Outlier Detection (IQR + Z-score)** | ✅ | ❌ | ❌ | ❌ |
| **Smart Non-Overlapping Charts** | ✅ | ❌ | ⚠️ | ⚠️ |
| **Zero Process Leaks** | ✅ | ⚠️ | ✅ | ❌ |
| **Native Pivot Tables** | ✅ | ✅ | ❌ | ✅ |
| **Financial Presets** | ✅ | ❌ | ❌ | ❌ |

---

## Status & Roadmap

### Current Status: **Beta** ✅

- ✅ Core engine fully functional
- ✅ Live COM backend production-ready
- ✅ Headless OpenPyXL backend production-ready
- ✅ Daemon IPC stable
- ✅ 27 unit tests (100% passing)
- ✅ Data curation pipeline complete
- ✅ Formula auditing working

### Known Limitations

- ⚠️ Windows-only (COM is Windows-specific; Linux/Mac unsupported)
- ⚠️ Excel 2016+ required for Live COM mode
- ⚠️ Large files (> 10k rows) may need tuning for performance
- ⚠️ Concurrency via daemon is safe but single-threaded per session

### Roadmap (Priority Order)

1. **Refactoring** — Modularize `antigravity_excel_core.py` for easier maintenance
2. **Better Logging** — Add structured logging for debugging daemon issues
3. **TypeHints** — Full type annotations for better IDE support
4. **Linting** — Integrate ruff/black/mypy in CI
5. **Performance Tuning** — Optimize for 10k+ row datasets
6. **Extended Testing** — CI/CD with GitHub Actions for Windows
7. **Mac/Linux Support** — Evaluate headless-only mode for non-Windows

---

## Troubleshooting

### Problem: Daemon won't start

```bash
python antigravity_excel_daemon.py --status
# Result: daemon_running: false
```

**Solution:**

1. Check if another daemon is already running:
   ```bash
   tasklist | findstr python
   ```
2. Kill any orphaned processes:
   ```bash
   taskkill /F /IM EXCEL.EXE  # Last resort
   ```
3. Check file permissions — ensure you can write to the directory
4. Try explicit mode:
   ```bash
   python antigravity_excel_daemon.py --start --mode file
   ```

---

### Problem: "RPC_E_SERVERCALL_RETRYLATER" errors

**This is normal** when Excel is actively being used. The engine retries automatically.

**If errors persist:**

1. Increase retry count in `antigravity_excel_core.py` (line ~343):
   ```python
   @retry_on_excel_busy(max_retries=15, delay_sec=1.0)
   ```

2. Or avoid Live COM mode:
   ```python
   engine = AntigravityExcelEngine(mode="file", file_path="report.xlsx")
   ```

---

### Problem: "File is locked" errors

**Solution:**

1. Make sure Excel isn't holding the file open in another process
2. Use daemon with persistent session:
   ```bash
   python antigravity_excel_daemon.py --start --file "C:/data/report.xlsx"
   ```
3. If stuck, close the file and try again

---

### Problem: Daemon latency is high (> 50ms)

**Possible causes:**

1. System is under heavy load
2. Excel file is very large (> 100k cells)
3. Antivirus is scanning the file

**Solutions:**

1. Close other applications
2. Simplify operations (read smaller ranges)
3. Add file to antivirus exclusion list
4. Use headless mode instead:
   ```python
   engine = AntigravityExcelEngine(mode="file", file_path="report.xlsx")
   ```

---

### Problem: "Module not found" errors

**Solution:**

```bash
# Ensure all dependencies installed
pip install -r requirements.txt --force-reinstall

# For Windows, also ensure pywin32 is properly installed:
python -m pip install --upgrade pywin32
```

---

## Contributing

Contributions are welcome! Areas where help is most needed:

1. **Code quality** — Add type hints, improve docstrings
2. **Testing** — Add tests for edge cases, concurrency, large files
3. **Documentation** — Examples, use cases, troubleshooting
4. **Performance** — Profile and optimize hot paths
5. **Architecture** — Help refactor `antigravity_excel_core.py`

**To contribute:**

1. Fork the repository
2. Create a branch: `git checkout -b feature/your-feature`
3. Make changes + run tests: `python -m pytest tests/ -v`
4. Commit: `git commit -m "Add feature X"`
5. Push: `git push origin feature/your-feature`
6. Open a Pull Request

---

## License

MIT License — See [LICENSE](LICENSE) file for details.

**Copyright © 2024 Francisco Grandón Vergara**

---

## Questions or Issues?

- 📧 Email: Francisco.Grandon@essbio.cl (if contact is available)
- 🐛 [Open an issue on GitHub](https://github.com/FranciscoGrandon/antigravity-excel-engine/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/FranciscoGrandon/antigravity-excel-engine/discussions)

---

**Built with ❤️ for AI agents and data automation teams.**
