# Antigravity Excel Engine Architecture

This document explains how the project is structured internally and how the different components interact.

---

## Overview

Antigravity Excel Engine is designed around a simple principle:

> Excel automation must be fast, resilient, and compatible with AI-driven workflows.

The project is built around a dual-backend architecture:

- Live COM backend for Excel sessions already open in the desktop app
- Headless backend for file-driven execution without UI

The engine automatically chooses the most appropriate backend based on runtime conditions.

---

## High-Level Components

### 1. Core Engine

The main orchestration layer lives in `antigravity_excel_core.py`.

This module contains:

- `AntigravityExcelEngine`
- workbook-level operations
- formula validation and audit routines
- data cleaning and deduplication logic
- pivot and chart creation
- formatting and style presets
- helper functions for coordinate parsing and formula sanitization

This is the central component of the project.

### 2. Daemon Layer

The background service is implemented in `antigravity_excel_daemon.py`.

Its purpose is to:

- keep a persistent Excel session alive in memory
- expose ultra-fast IPC through Windows Named Pipes
- reduce repeated operation latency
- serve repeated requests without reinitializing the engine every time

The daemon receives JSON payloads and dispatches them to the engine.

### 3. CLI Layer

The CLI is implemented in `antigravity_excel_cli.py`.

It provides a command-line interface for:

- reading ranges as CSV
- running audits
- generating summaries
- creating charts
- running curation pipelines
- interacting with the daemon

The CLI is the easiest path for scripts, automation jobs, and AI agents.

---

## Execution Flow

### Standard API Flow

```text
Python client
   -> AntigravityExcelEngine
      -> mode detection
         -> Live COM backend or Headless backend
      -> operation executes
      -> result returned
```

### Daemon Flow

```text
CLI / Python client
   -> Named Pipe request
   -> antigravity_excel_daemon.py
      -> AntigravityExcelEngine
      -> operation executes
      -> JSON response returned
```

---

## Backend Strategy

### Live COM Backend

Used when the workbook is already open in Excel.

This mode is ideal when:

- the user expects visual updates in Excel
- formulas must recalculate immediately
- undo/redo behavior matters
- a user is actively editing the workbook

This backend relies on `win32com` and Windows-specific Excel automation services.

Benefits:

- real-time UI behavior
- better compatibility with interactive Excel operations
- full workbook recalculation semantics

Trade-offs:

- Windows-only
- more sensitive to locks and COM failures
- requires the Excel app to be available

### Headless Backend

Used when the workbook is not open or a file-only workflow is preferred.

This mode uses `openpyxl` and is designed for:

- automation scripts
- batch processing
- read-only extraction
- AI pre-processing of workbook data
- pipeline tasks without UI interaction

Benefits:

- lightweight and fast
- easy to run in automation pipelines
- no visible Excel session required
- stable for file-driven computations

Trade-offs:

- no live UI updates
- formulas may not behave exactly like a visible Excel session

---

## Why the Daemon Exists

A direct engine call works, but repeated calls create unnecessary overhead:

- Python startup cost
- Excel COM reinitialization cost
- repeated imports and warm-up time
- extra resource churn

The daemon resolves this by keeping the Excel engine alive in memory.

### Daemon Responsibilities

- keep one background engine alive
- receive incoming JSON commands
- route them to the correct workbook state
- return structured JSON results
- keep the Excel session warm for repeated accesses

This reduces latency for repeated operations and makes agent workflows more responsive.

---

## Internal Design Principles

### 1. Mode Detection First

Every operation begins by selecting the correct runtime mode:

- live if workbook is open
- file if workbook is closed
- fallback if neither state is valid

This ensures compatibility with both interactive and non-interactive Excel use cases.

### 2. Optimized Data Output for LLMs

The project intentionally avoids returning verbose nested JSON for many operations.

Instead, it prefers:

- compact CSV output
- dense structured rows
- direct range extraction
- lightweight summaries

This matters because AI agents operate with token constraints and need compact but high-value data.

### 3. Preserve Excel Semantics Where Possible

The engine does not just expose raw cell access; it attempts to preserve Excel semantics:

- formulas remain formulas
- number formats remain consistent
- styles and colors can be applied declaratively
- structured tables can be built programmatically
- outlier detection can annotate cells without destroying data

### 4. Be Safe Around Excel Locks

Windows COM workflows are fragile under contention.

The engine must account for:

- file locks
- user editing the sheet
- active screen updates
- Excel busy states
- retry conditions after transient failures

This is a core requirement of the project.

---

## Data Flow Patterns

### Read Pattern

```text
Request range -> backend reads cells -> normalize values -> return CSV or JSON
```

### Write Pattern

```text
Input cells -> convert to Excel coordinates -> render styles -> save workbook -> return stats
```

### Audit Pattern

```text
Read range -> validate formulas -> detect anomalies -> annotate cells -> return report
```

### Curation Pattern

```text
Read range -> clean strings -> cast types -> normalize dates -> deduplicate -> flag outliers -> save workbook
```

---

## Core Responsibilities by Module

### `antigravity_excel_core.py`

Primary tasks:

- workbook lifecycle management
- backend orchestration
- formula sanitization and compatibility
- format and style logic
- metadata and audit generation
- data transformation and aggregation
- pivot tables and chart support

### `antigravity_excel_daemon.py`

Primary tasks:

- persistent background service
- Named Pipe message handling
- JSON request parsing
- request dispatching
- process lifecycle control
- health and latency monitoring

### `antigravity_excel_cli.py`

Primary tasks:

- user-facing access layer
- command parsing
- command dispatching to daemon or direct engine
- JSON output formatting
- friendly developer UX for ad-hoc tasks

---

## Performance Considerations

This project is intentionally designed around a specific performance profile:

- fast repeated access
- minimal Python startup cost
- low-latency request/response cycles
- optimized data output
- minimal state churn

### Hot Paths

The hottest areas are:

- reading large ranges
- writing large batches of cells
- CSV output generation
- formula checks
- deduplication and curation
- daemon ping latency

These are the parts most likely to matter for real-world AI workflows.

---

## Risks and Failure Modes

### COM Lock Failures

When Excel is busy, COM operations may fail or stall.

Control strategy:

- retry policy
- lock detection
- fallback to headless execution when possible
- structured warnings and errors

### Resource Leaks

Excel automation can leave zombie processes if cleanup is imperfect.

Control strategy:

- explicit cleanup hooks
- close workbook and release handles
- daemon shutdown path
- process health validation

### Formula Incompatibility

Excel versions and formula namespaces differ across environments.

Control strategy:

- sanitize modern formulas
- prefix compatibility functions where necessary
- keep formula logic explicit and auditable

---

## Extension Points

The architecture is intentionally extensible.

Possible future extensions include:

- backend plugins for new spreadsheet engines
- richer validation layers
- API wrappers for REST or MCP use
- notebook integration
- reporting templates
- workbook snapshots and diffing

---

## Recommended Refactor Direction

Currently, the engine combines many responsibilities in one large file. A logical next step is to split by concern.

### Suggested Module Layout

```text
antigravity_excel_engine/
├── backends/
│   ├── __init__.py
│   ├── live_com.py
│   └── headless.py
├── operations/
│   ├── read.py
│   ├── write.py
│   ├── audit.py
│   ├── aggregate.py
│   └── curation.py
├── formatting/
│   ├── presets.py
│   └── styles.py
├── formulas/
│   ├── sanitize.py
│   └── audit.py
├── ipc/
│   ├── daemon.py
│   └── client.py
├── utils/
│   ├── coordinates.py
│   ├── dates.py
│   └── validation.py
├── engine.py
└── cli.py
```

This would improve maintainability and reduce coupling between runtime concerns and business logic.

---

## Summary

The project is built around a strong idea:

- keep Excel automation fast
- keep it safe under COM locks
- optimize data output for LLM consumption
- make the workflow AI-friendly

The architecture reflects this intentionally:

- the engine handles operations
- the daemon keeps sessions warm
- the CLI makes operations accessible
- the backends adapt to Excel state and runtime context

This is a practical, agent-oriented architecture, not just a raw spreadsheet wrapper.
