package template

import (
	"fmt"
	"strings"
	"time"

	"github.com/orayew2002/rast-excel/domain"
	"github.com/orayew2002/rast-excel/excel"
	"github.com/xuri/excelize/v2"
)

// RegisterDefaults registers all built-in template handlers.
func RegisterDefaults(r *Registry) {
	r.Register("{{days}}", handleDays)
	r.Register("{{working_time}}", handleWorkingTime)
}

func handleDays(f *excelize.File, sheet string, row, col int, _ string) error {
	days := currentMonthDays()
	fmt.Println("current month days =", days)

	if err := f.InsertCols(sheet, excel.IndexToColumn(col+1), days-1); err != nil {
		return fmt.Errorf("insert cols: %w", err)
	}

	// Merge header rows above the day numbers.
	for _, headerRow := range []int{0, 1} {
		topLeft := excel.CellName(headerRow, col)
		bottomRight := excel.CellName(headerRow, col+days-1)
		if err := f.MergeCell(sheet, topLeft, bottomRight); err != nil {
			return fmt.Errorf("merge row %d: %w", headerRow, err)
		}
	}

	// Fill day numbers (1..N).
	for i := 0; i < days; i++ {
		cell := excel.CellName(row, col+i)
		if err := f.SetCellInt(sheet, cell, int64(i+1)); err != nil {
			return fmt.Errorf("set day %d: %w", i+1, err)
		}
	}

	// Copy style across all day cells.
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

func currentMonthDays() int {
	now := time.Now().Local()
	year, month, _ := now.Date()
	first := time.Date(year, month+1, 1, 0, 0, 0, 0, now.Location())
	return first.AddDate(0, 0, -1).Day()
}
