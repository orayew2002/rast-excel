# rast-excel

Go package for processing Excel templates with dynamic data injection. Reads an `.xlsx` template with placeholder variables, replaces them with real data, and outputs the result.

## Install

```bash
go get github.com/orayew2002/rast-excel
```

## Project Structure

```
rast-excel/
├── main.go               # CLI entry point (example usage)
├── domain/
│   ├── const.go           # Key-value mappings for text replacements
│   └── domain.go          # Employee struct + fake data generator
├── processor/
│   └── processor.go       # Core engine: open file → process sheets → save/return bytes
├── template/
│   ├── registry.go        # Template handler registry (pattern → handler)
│   ├── handlers.go        # Built-in handlers: {{days}}, {{working_time}}, {{start_process}}
│   └── styles.go          # StyleManager with caching (creates each style once)
└── excel/
    └── cell.go            # Helpers: CellName(), IndexToColumn()
```

## Template Variables

| Variable | Description |
|---|---|
| `{{days}}` | Inserts columns for each day of the current month (1..28/30/31) |
| `{{working_time}}` | Replaces with the working time label from `domain.KeyMap` |
| `{{start_process}}` | Fills employee rows with data you provide |

## Usage

### As a CLI tool

```bash
# default: reads table.xlsx, writes result.xlsx
go run main.go

# custom paths
go run main.go -input template.xlsx -output output.xlsx
```

### As a package in your project

#### Basic — process file from disk

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
    // 1. Create registry and register handlers.
    registry := template.New()
    template.RegisterDefaults(registry)

    // 2. Pass your employee data.
    employees := []domain.Employee{
        {Id: 1, FullName: "Aman Orazow",  TableID: "001", JobPosition: "Engineer",  Attendance: []string{"8", "8", "W", "8"}},
        {Id: 2, FullName: "Aylar Hanowa", TableID: "002", JobPosition: "Designer",  Attendance: []string{"8", "P", "8", "8"}},
        {Id: 3, FullName: "Merdan Oraz",  TableID: "003", JobPosition: "Manager",   Attendance: []string{"8", "8", "8", "A"}},
    }
    template.RegisterEmployeeHandler(registry, employees)

    // 3. Process.
    p := processor.New(registry)
    _, err := p.ProcessFile("template.xlsx", "result.xlsx")
    if err != nil {
        fmt.Fprintf(os.Stderr, "error: %v\n", err)
        os.Exit(1)
    }
}
```

#### Process from bytes (for HTTP handlers / APIs)

```go
func handleUpload(templateBytes []byte) ([]byte, error) {
    registry := template.New()
    template.RegisterDefaults(registry)
    template.RegisterEmployeeHandler(registry, myEmployees)

    p := processor.New(registry)
    return p.ProcessBytes(templateBytes)
}
```

#### Testing with fake data

```go
// domain.GenerateEmployees generates N employees with random names,
// positions, and attendance matching the current month's day count.
employees := domain.GenerateEmployees(25)
template.RegisterEmployeeHandler(registry, employees)
```

### Adding a new template variable

1. Write a handler function:

```go
func handleMyVar(f *excelize.File, sheet string, row, col int, value string) error {
    cell := excel.CellName(row, col)
    replaced := strings.ReplaceAll(value, "{{my_var}}", "actual value")
    return f.SetCellStr(sheet, cell, replaced)
}
```

2. Register it:

```go
registry := template.New()
template.RegisterDefaults(registry)
registry.Register("{{my_var}}", handleMyVar)
```

That's it. The processor will automatically pick it up when scanning cells.

### Adding a new employee column

Edit the `columns` slice in `template/handlers.go`:

```go
var columns = []columnDef{
    {value: func(e domain.Employee) string { return strconv.Itoa(e.Id) }, style: (*StyleManager).Centered},
    {value: func(e domain.Employee) string { return e.FullName },         style: (*StyleManager).Centered},
    {value: func(e domain.Employee) string { return e.TableID },          style: (*StyleManager).Centered},
    {value: func(e domain.Employee) string { return e.JobPosition },      style: (*StyleManager).Centered},
    // add your new column here:
    {value: func(e domain.Employee) string { return e.NewField },         style: (*StyleManager).Left},
}
```

## API Reference

### `template.New() *Registry`
Creates an empty template handler registry.

### `template.RegisterDefaults(r *Registry)`
Registers `{{days}}` and `{{working_time}}` handlers.

### `template.RegisterEmployeeHandler(r *Registry, employees []domain.Employee)`
Registers the `{{start_process}}` handler with your employee data.

### `processor.New(registry) *Processor`
Creates a processor with the given registry.

### `processor.ProcessFile(input, output string) ([]byte, error)`
Opens an Excel file from disk, processes all sheets, saves to disk, and returns the result as bytes.

### `processor.ProcessBytes(data []byte) ([]byte, error)`
Processes an Excel file from raw bytes and returns the result as bytes. No disk I/O.

### `domain.Employee`
```go
type Employee struct {
    Id          int
    FullName    string
    TableID     string
    JobPosition string
    Attendance  []string  // one entry per day, e.g. "8", "W", "P", "A", "L"
}
```

### `domain.GenerateEmployees(n int) []Employee`
Generates N fake employees for testing. Attendance length matches the current month's day count.
