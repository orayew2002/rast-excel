# rast-excel

A Go library for generating employee attendance Excel reports from an `.xlsx` template.
Fill a template with real employee data, expand attendance columns for the current month, and write per-employee summary formulas — all without Excel installed.

---

## Table of Contents

- [How It Works](#how-it-works)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Template Placeholders](#template-placeholders)
- [Providing Your Own Employees](#providing-your-own-employees)
- [Custom Formula Keys](#custom-formula-keys)
- [Processing API](#processing-api)
- [Package Overview](#package-overview)
- [Project Structure](#project-structure)
- [CLI Usage](#cli-usage)

---

## How It Works

Processing happens in **two steps**, each powered by the same `processor.Processor` engine:

```
┌──────────────────────────────────────────────────────────┐
│  Step 1 — Structural pass  (reads .xlsx from disk)       │
│                                                          │
│  {{days}}          → expands header to current month     │
│  {{working_time}}  → replaces with localized label       │
│  {{start_process}} → inserts one row per employee        │
└──────────────────────────────────────────────────────────┘
                          │ []byte
                          ▼
┌──────────────────────────────────────────────────────────┐
│  Step 2 — Formula pass  (reads from bytes)               │
│                                                          │
│  {{t}}, {{d}}, …   → per-employee SUMPRODUCT formulas    │
│  {{num_sum}}       → sum of numeric attendance values    │
│  {{num_count}}     → count of numeric attendance cells   │
│  {{}}              → centered style only, no formula     │
└──────────────────────────────────────────────────────────┘
                          │ []byte
                          ▼
                     result.xlsx
```

Every placeholder cell is **cleared** after processing — the output contains only values, formulas, and styles.

---

## Installation

```bash
go get github.com/orayew2002/rast-excel
```

Requires Go 1.22+.

---

## Quick Start

```go
package main

import (
    "fmt"
    "os"

    "github.com/orayew2002/rast-excel/domain"
    "github.com/orayew2002/rast-excel/processor"
    "github.com/orayew2002/rast-excel/template"
)

func main() {
    employees := domain.GenerateEmployees(25) // or supply your own

    data, err := step1("table.xlsx", employees)
    if err != nil {
        fmt.Fprintln(os.Stderr, err)
        os.Exit(1)
    }

    data, err = step2(data, len(employees))
    if err != nil {
        fmt.Fprintln(os.Stderr, err)
        os.Exit(1)
    }

    if err := os.WriteFile("result.xlsx", data, 0644); err != nil {
        fmt.Fprintln(os.Stderr, err)
        os.Exit(1)
    }
}

// step1 injects {{days}}, {{working_time}}, and employee rows.
func step1(input string, employees []domain.Employee) ([]byte, error) {
    registry := template.New()
    template.RegisterDefaults(registry)
    template.RegisterEmployeeHandler(registry, employees)
    return processor.New(registry).ProcessFile(input)
}

// step2 writes per-employee formulas for every {{key}} summary cell.
func step2(data []byte, employeeCount int) ([]byte, error) {
    registry := template.New()
    attStart := template.AttendanceStartCol(0)
    template.RegisterFormulaHandler(registry, employeeCount, attStart, []template.FormulaKey{
        {Key: "{{t}}",         FormulaFn: template.CountIFFormula("T", 1)},
        {Key: "{{d}}",         FormulaFn: template.CountIFFormula("D", 1)},
        {Key: "{{w}}",         FormulaFn: template.CountIFFormula("W", 1)},
        {Key: "{{l}}",         FormulaFn: template.CountIFFormula("L", 1)},
        {Key: "{{a}}",         FormulaFn: template.CountIFFormula("A", 1)},
        {Key: "{{p}}",         FormulaFn: template.CountIFFormula("P", 1)},
        {Key: "{{num_sum}}",   FormulaFn: template.SumNumFormula()},
        {Key: "{{num_count}}", FormulaFn: template.CountNumFormula()},
        {Key: "{{}}",          FormulaFn: nil},
    })
    return processor.New(registry).ProcessBytes(data)
}
```

---

## Template Placeholders

Place these keys inside cells of your `.xlsx` template file.

### Step 1 — Structural Keys

| Key | Description |
|-----|-------------|
| `{{days}}` | Expands the attendance header to cover every day of the current month. Merges header rows and sets column widths automatically. |
| `{{working_time}}` | Replaced with the localized working-time label defined in `domain.KeyMap`. |
| `{{start_process}}` | Marks the row where employee data is inserted. Writes one row per employee: fixed columns (ID, full name, table ID, job position) followed by daily attendance values. |

### Step 2 — Formula Keys

These cells must appear **below** the employee block. A formula is written into the same column for every employee row above, then the placeholder cell is cleared.

| Key | Description |
|-----|-------------|
| `{{t}}` | Count of `"T"` entries in the employee's attendance range |
| `{{d}}` | Count of `"D"` entries |
| `{{w}}` | Count of `"W"` entries |
| `{{l}}` | Count of `"L"` entries |
| `{{a}}` | Count of `"A"` entries |
| `{{p}}` | Count of `"P"` entries |
| `{{num_sum}}` | Sum of numeric attendance values — e.g. `"8", "W", "8"` → `16`. Returns `0` when none. |
| `{{num_count}}` | Count of cells that contain a number — e.g. `"8", "W", "8"` → `2`. Returns `0` when none. |
| `{{}}` | **Style-only.** Applies centered style to each employee cell. No formula written. Useful for visual spacing or separator columns. |

> **Combining keys:** A single cell may hold multiple keys, e.g. `{{d}}{{t}}`.
> Their formulas are summed: `countD + countT`.
>
> **Zero behaviour:** Results equal to `0` are stored as `""` (blank cell) for symbol-count keys.
> `{{num_sum}}` and `{{num_count}}` always display the numeric `0`.

---

## Providing Your Own Employees

Use the `domain.Employee` struct to supply real data:

```go
import "github.com/orayew2002/rast-excel/domain"

employees := []domain.Employee{
    {
        Id:          1,
        FullName:    "Alice Smith",
        TableID:     "001",
        JobPosition: "Engineer",
        // one entry per calendar day of the current month
        Attendance: []string{"W", "8", "W", "L", "8", "W", "W" /*, … */},
    },
    {
        Id:          2,
        FullName:    "Bob Jones",
        TableID:     "002",
        JobPosition: "Manager",
        Attendance:  []string{"8", "8", "W", "8", "8", "D", "W" /*, … */},
    },
}
```

`Attendance` length must equal the number of days in the current month.
Use `domain.GenerateEmployees(n)` to generate random test data.

### Attendance Symbols

| Symbol | Meaning |
|--------|---------|
| `"8"` | Worked 8-hour day (numeric) |
| `"W"` | Weekend |
| `"T"` | Business trip |
| `"D"` | Day off |
| `"L"` | Leave / vacation |
| `"A"` | Absent |
| `"P"` | Public holiday |

---

## Custom Formula Keys

Register additional keys alongside the built-ins:

```go
import (
    "fmt"
    "github.com/orayew2002/rast-excel/template"
)

template.RegisterFormulaHandler(registry, employeeCount, attStart, []template.FormulaKey{
    // built-in
    {Key: "{{t}}", FormulaFn: template.CountIFFormula("T", 1)},

    // count "OT" (overtime) entries
    {Key: "{{ot}}", FormulaFn: func(attRange string) string {
        return fmt.Sprintf(`SUMPRODUCT((%s="OT")*1)`, attRange)
    }},

    // count "T" entries multiplied by 8 hours
    {Key: "{{th}}", FormulaFn: template.CountIFFormula("T", 8)},

    // style-only column — no formula, just styling
    {Key: "{{}}", FormulaFn: nil},
})
```

`FormulaFn` receives the Excel range string for one employee's attendance row
(e.g. `"E5:AF5"`) and must return a valid Excel formula **without** the leading `=`.

Set `FormulaFn` to `nil` for a style-only key.

---

## Processing API

### `processor.New(registry).ProcessFile(path string) ([]byte, error)`

Opens an `.xlsx` file from disk, processes all sheets, and returns the result as bytes.

```go
data, err := processor.New(registry).ProcessFile("table.xlsx")
```

### `processor.New(registry).ProcessBytes(data []byte) ([]byte, error)`

Same as `ProcessFile` but reads from an in-memory byte slice. Use this for step 2, or in HTTP handlers / APIs to avoid touching the filesystem.

```go
// HTTP upload example
func generateReport(w http.ResponseWriter, r *http.Request) {
    templateBytes, _ := io.ReadAll(r.Body)

    data, err := step1bytes(templateBytes, employees)
    // ...
    data, err = step2(data, len(employees))
    // ...
    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Write(data)
}
```

### `template.AttendanceStartCol(employeeCol int) int`

Returns the 0-based column index where attendance data begins. Pass the column where the employee section starts (`0` for column A).

```go
// employee section starts at column A (0) → attendance starts at column E (4)
attStart := template.AttendanceStartCol(0)
```

---

## Package Overview

| Package | Responsibility |
|---------|---------------|
| `domain` | `Employee` struct, `GenerateEmployees`, `KeyMap` for text replacements |
| `template` | Handler registration, `FormulaKey`, formula builders (`CountIFFormula`, `SumNumFormula`, `CountNumFormula`), `StyleManager` |
| `processor` | `Processor` — iterates all cells in all sheets and dispatches to the registry |
| `excel` | `CellName(row, col)`, `IndexToColumn(n)` — coordinate helpers |

### Key Types

```go
// domain
type Employee struct {
    Id          int
    FullName    string
    TableID     string
    JobPosition string
    Attendance  []string // one entry per calendar day
}

// template
type FormulaKey struct {
    Key       string
    FormulaFn func(attRange string) string // nil = style-only
}

// template
type Registry struct { /* … */ }
func (r *Registry) Register(pattern string, handler HandlerFunc)

// processor
type Processor struct { /* … */ }
func (p *Processor) ProcessFile(input string) ([]byte, error)
func (p *Processor) ProcessBytes(data []byte) ([]byte, error)
```

---

## Project Structure

```
rast-excel/
├── main.go                 # CLI entry point
├── domain/
│   ├── domain.go           # Employee struct + GenerateEmployees
│   └── const.go            # KeyMap (text replacements)
├── processor/
│   └── processor.go        # Core engine — open → process sheets → return bytes
├── template/
│   ├── registry.go         # Registry: pattern → HandlerFunc
│   ├── handlers.go         # Built-in handlers + RegisterFormulaHandler
│   └── styles.go           # StyleManager (cached Excel styles)
└── excel/
    └── cell.go             # CellName(), IndexToColumn()
```

---

## CLI Usage

The included `main.go` is a ready-to-run command:

```bash
go run . -input table.xlsx -output result.xlsx
```

| Flag | Default | Description |
|------|---------|-------------|
| `-input` | `table.xlsx` | Path to the template Excel file |
| `-output` | `result.xlsx` | Path for the generated output file |
