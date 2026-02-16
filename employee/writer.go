package employee

import (
	"fmt"
	"strconv"

	"github.com/orayew2002/rast-excel/excel"
	excelize "github.com/xuri/excelize/v2"
)

// headers defines the column layout for the employee table.
var headers = []string{"№", "F.I.O.", "Wezipesi", "Bölümi"}

// WriteToFile creates a new Excel file with employee data and saves it to path.
func WriteToFile(employees []Employee, path string) error {
	f := excelize.NewFile()
	sheet := "Sheet1"

	if err := writeHeaders(f, sheet); err != nil {
		return fmt.Errorf("write headers: %w", err)
	}

	if err := writeRows(f, sheet, employees); err != nil {
		return fmt.Errorf("write rows: %w", err)
	}

	if err := autoFitColumns(f, sheet); err != nil {
		return fmt.Errorf("auto fit columns: %w", err)
	}

	if err := f.SaveAs(path); err != nil {
		return fmt.Errorf("save %s: %w", path, err)
	}

	return nil
}

// WriteToBytes creates a new Excel file with employee data and returns it as bytes.
func WriteToBytes(employees []Employee) ([]byte, error) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	if err := writeHeaders(f, sheet); err != nil {
		return nil, fmt.Errorf("write headers: %w", err)
	}

	if err := writeRows(f, sheet, employees); err != nil {
		return nil, fmt.Errorf("write rows: %w", err)
	}

	if err := autoFitColumns(f, sheet); err != nil {
		return nil, fmt.Errorf("auto fit columns: %w", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("write to buffer: %w", err)
	}

	return buf.Bytes(), nil
}

func writeHeaders(f *excelize.File, sheet string) error {
	style, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	if err != nil {
		return err
	}

	for col, header := range headers {
		cell := excel.CellName(0, col)
		if err := f.SetCellStr(sheet, cell, header); err != nil {
			return err
		}
		if err := f.SetCellStyle(sheet, cell, cell, style); err != nil {
			return err
		}
	}

	return nil
}

func writeRows(f *excelize.File, sheet string, employees []Employee) error {
	for i, emp := range employees {
		row := i + 1 // row 0 is headers
		values := []string{
			strconv.Itoa(emp.ID),
			emp.FullName,
			emp.Position,
			emp.Department,
		}
		for col, val := range values {
			cell := excel.CellName(row, col)
			if err := f.SetCellStr(sheet, cell, val); err != nil {
				return fmt.Errorf("employee %d, col %d: %w", emp.ID, col, err)
			}
		}
	}

	return nil
}

func autoFitColumns(f *excelize.File, sheet string) error {
	widths := []float64{5, 30, 25, 20}
	for col, w := range widths {
		colName := excel.IndexToColumn(col)
		if err := f.SetColWidth(sheet, colName, colName, w); err != nil {
			return err
		}
	}
	return nil
}
