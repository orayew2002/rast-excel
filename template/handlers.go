package template

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/orayew2002/rast-excel/domain"
	"github.com/orayew2002/rast-excel/excel"
	"github.com/xuri/excelize/v2"
)

// RegisterDefaults registers the built-in template handlers (days, working_time).
func RegisterDefaults(r *Registry) {
	r.Register("{{days}}", handleDays)
	r.Register("{{working_time}}", handleWorkingTime)
}

// RegisterEmployeeHandler registers the {{start_process}} handler
// with the given employee list. Pass your own data from your project,
// or use domain.GenerateEmployees() for testing with fake data.
func RegisterEmployeeHandler(r *Registry, employees []domain.Employee) {
	r.Register("{{start_process}}", func(f *excelize.File, sheet string, row, col int, _ string) error {
		return writeEmployees(f, sheet, row, col, employees)
	})
}

// ---------- {{days}} ----------

func handleDays(f *excelize.File, sheet string, row, col int, _ string) error {
	days := currentMonthDays()
	if err := f.InsertCols(sheet, excel.IndexToColumn(col+1), days-1); err != nil {
		return fmt.Errorf("insert cols: %w", err)
	}

	for _, headerRow := range []int{0, 1} {
		topLeft := excel.CellName(headerRow, col)
		bottomRight := excel.CellName(headerRow, col+days-1)
		if err := f.MergeCell(sheet, topLeft, bottomRight); err != nil {
			return fmt.Errorf("merge row %d: %w", headerRow, err)
		}
	}

	for i := range days {
		cell := excel.CellName(row, col+i)
		if err := f.SetCellInt(sheet, cell, int64(i+1)); err != nil {
			return fmt.Errorf("set day %d: %w", i+1, err)
		}
	}

	cell := excel.CellName(row, col)
	styleID, _ := f.GetCellStyle(sheet, cell)

	topLeft := excel.CellName(row, col)
	bottomRight := excel.CellName(row, col+days-1)
	if err := f.SetCellStyle(sheet, topLeft, bottomRight, styleID); err != nil {
		return fmt.Errorf("set style: %w", err)
	}

	colStart := excel.IndexToColumn(col)
	colEnd := excel.IndexToColumn(col + days - 1)
	if err := f.SetColWidth(sheet, colStart, colEnd, 4); err != nil {
		return fmt.Errorf("set col width: %w", err)
	}

	return nil
}

// ---------- {{working_time}} ----------

func handleWorkingTime(f *excelize.File, sheet string, row, col int, value string) error {
	cell := excel.CellName(row, col)

	styleID, _ := f.GetCellStyle(sheet, cell)
	replaced := strings.ReplaceAll(value, "{{working_time}}", domain.KeyMap["{{working_time}}"])

	if err := f.SetCellStr(sheet, cell, replaced); err != nil {
		return fmt.Errorf("set working_time: %w", err)
	}

	if styleID != 0 {
		if err := f.SetCellStyle(sheet, cell, cell, styleID); err != nil {
			return fmt.Errorf("restore style: %w", err)
		}
	}

	return nil
}

// ---------- {{start_process}} ----------

// columnDef describes one employee column: how to extract the value and which style to use.
type columnDef struct {
	value func(emp domain.Employee) string
	style func(sm *StyleManager) (int, error)
}

// columns defines the fixed employee columns in order.
// To add a new column: append one entry here — that's it.
var columns = []columnDef{
	{value: func(e domain.Employee) string { return strconv.Itoa(e.Id) }, style: (*StyleManager).Centered},
	{value: func(e domain.Employee) string { return e.FullName }, style: (*StyleManager).Centered},
	{value: func(e domain.Employee) string { return e.TableID }, style: (*StyleManager).Centered},
	{value: func(e domain.Employee) string { return e.JobPosition }, style: (*StyleManager).Centered},
}

func writeEmployees(f *excelize.File, sheet string, row, col int, employees []domain.Employee) error {
	if err := f.RemoveRow(sheet, row+1); err != nil {
		return fmt.Errorf("remove template row: %w", err)
	}

	count := len(employees)
	if err := f.InsertRows(sheet, row+1, count); err != nil {
		return fmt.Errorf("insert rows: %w", err)
	}

	sm := NewStyleManager(f)

	for i, emp := range employees {
		if err := writeEmployeeRow(f, sm, sheet, row+i, col, emp); err != nil {
			return fmt.Errorf("employee %d: %w", emp.Id, err)
		}
	}

	return nil
}

func writeEmployeeRow(f *excelize.File, sm *StyleManager, sheet string, row, col int, emp domain.Employee) error {
	// Write fixed columns.
	for c, def := range columns {
		cell := excel.CellName(row, col+c)
		if err := f.SetCellStr(sheet, cell, def.value(emp)); err != nil {
			return fmt.Errorf("col %d: %w", c, err)
		}

		styleID, err := def.style(sm)
		if err != nil {
			return fmt.Errorf("style col %d: %w", c, err)
		}
		if err := f.SetCellStyle(sheet, cell, cell, styleID); err != nil {
			return fmt.Errorf("set style col %d: %w", c, err)
		}
	}

	// Write attendance columns after fixed columns.
	attStart := col + len(columns)
	centeredStyle, err := sm.Centered()
	if err != nil {
		return fmt.Errorf("attendance style: %w", err)
	}

	for i, att := range emp.Attendance {
		cell := excel.CellName(row, attStart+i)
		if err := f.SetCellStr(sheet, cell, att); err != nil {
			return fmt.Errorf("attendance %d: %w", i, err)
		}

		if err := f.SetCellStyle(sheet, cell, cell, centeredStyle); err != nil {
			return fmt.Errorf("attendance style %d: %w", i, err)
		}
	}

	return nil
}

// ---------- helpers ----------

func currentMonthDays() int {
	now := time.Now().Local()
	year, month, _ := now.Date()
	first := time.Date(year, month+1, 1, 0, 0, 0, 0, now.Location())
	return first.AddDate(0, 0, -1).Day()
}
