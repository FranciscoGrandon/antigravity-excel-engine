# Antigravity Excel Engine Examples

This document contains practical examples for real-world usage.

---

## 1. Generate a Monthly Financial Report

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="monthly_report.xlsx")

engine.set_cell_range(
    "A1",
    [
        ["Month", "Revenue", "Costs", "Margin"],
        ["Jan", 100000, 70000, "=B2-C2"],
        ["Feb", 120000, 80000, "=B3-C3"],
        ["Mar", 140000, 90000, "=B4-C4"],
    ],
    autofit=True,
)

engine.apply_style_preset("A1:D1", "header")
engine.apply_style_preset("B2:D5", "currency")
engine.add_calculated_column(
    header="Margin %",
    formula="=D{r}/B{r}",
    number_format="0.0%",
)

engine.save()
print("Monthly report generated")
```

---

## 2. Audit Data Quality Before Training an ML Model

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="raw_data.xlsx")

quality_report = engine.audit_data_quality("A1:K500")
print(quality_report)

outliers = engine.detect_and_flag_outliers(
    column="Sales",
    method="iqr",
    range_or_table="A1:K500",
    add_comment=True,
)
print(outliers)

engine.save()
```

---

## 3. Clean and Normalize a Raw Excel Table

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="raw_table.xlsx")

result = engine.clean_table_dataset(
    target_range_or_table="A1:F1000",
    rules={
        "trim_headers": True,
        "standardize_text": {
            "Segment": {"trim": True, "casing": "title"},
            "Country": {"trim": True, "casing": "title"},
        },
        "cast_types": {
            "Sales": "float",
            "Units": "int",
        },
        "coerce_dates": ["Date"],
    },
)

print(result)
engine.save()
```

---

## 4. Build a Summary Dashboard

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="dashboard.xlsx")

engine.create_summary_table(
    title="Executive Summary",
    headers=["Scenario", "Revenue", "Cost", "Margin", "Growth"],
    rows_spec=[
        {"label": "Base", "values": [100000, 70000, "=B3-C3", 0.12], "formats": ["currency", "currency", "currency", "percent"]},
        {"label": "Optimistic", "values": [150000, 80000, "=B4-C4", 0.18], "formats": ["currency", "currency", "currency", "percent"]},
    ],
    dest_sheet="Summary",
    start_cell="B2",
)

engine.save()
```

---

## 5. AI Agent Reporting Workflow

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="agent_output.xlsx")

engine.set_cell_range(
    "A1",
    [
        ["Date", "Product", "Sales"],
        ["2024-01-01", "Product A", 1000],
        ["2024-01-01", "Product B", 1800],
        ["2024-01-02", "Product A", 1200],
    ],
)

summary = engine.aggregate(
    group_by_col="Product",
    metric_col="Sales",
    agg_func="sum",
    top_n=5,
)

print(summary)

engine.create_pivot_table(
    source="A1:C5",
    rows=["Product"],
    value_field="Sales",
    value_func="sum",
    dest_sheet="PivotSummary",
)

engine.create_chart(
    source="PivotSummary!A3:B6",
    chart_type="column_clustered",
    title="Sales by Product",
    placement="auto",
)

engine.save()
```

---

## 6. Use the Daemon for Fast Repeated Reads

```python
from antigravity_excel_daemon import ping_daemon, send_pipe_message

alive, latency, details = ping_daemon(timeout_ms=1000)
print(alive, latency, details)

resp = send_pipe_message({
    "command": "get-csv",
    "args": {
        "range": "A1:E20",
        "sheet": "Datos",
        "delimiter": ";",
    },
})

print(resp)
```

---

## 7. Build a High-Quality Financial Model

```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="auto", file_path="finance_model.xlsx")

engine.set_cells({
    "A1": {"value": "Metric", "cellStyles": {"fontWeight": "bold"}},
    "B1": {"value": "Value", "cellStyles": {"fontWeight": "bold"}},
    "A2": {"value": "Revenue"},
    "B2": {"value": 100000, "cellStyles": {"numberFormat": "$#,##0"}},
    "A3": {"value": "Costs"},
    "B3": {"value": 65000, "cellStyles": {"numberFormat": "$#,##0"}},
    "A4": {"value": "Margin"},
    "B4": {"formula": "=B2-B3", "cellStyles": {"numberFormat": "$#,##0"}},
}, autofit=True)

engine.apply_style_preset("A1:B1", "header")
engine.apply_style_preset("A4:B4", "total_row")
engine.save()
```
