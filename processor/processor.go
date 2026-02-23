package processor

import (
	"bytes"
	"fmt"

	"github.com/orayew2002/rast-excel/excel"
	"github.com/orayew2002/rast-excel/template"
	"github.com/xuri/excelize/v2"
)

// Processor applies registered template handlers to Excel files.
type Processor struct {
	registry *template.Registry
}

// New creates a Processor with the given template registry.
func New(registry *template.Registry) *Processor {
	return &Processor{registry: registry}
}

// ProcessFile opens an Excel file from disk, processes all sheets,
// and returns the result as bytes. It does NOT save to disk.
func (p *Processor) ProcessFile(input string) ([]byte, error) {
	f, err := excelize.OpenFile(input)
	if err != nil {
		return nil, fmt.Errorf("open %s: %w", input, err)
	}
	defer f.Close()

	return p.process(f)
}

// ProcessBytes opens an Excel file from raw bytes, processes all sheets,
// and returns the result as bytes.
func (p *Processor) ProcessBytes(data []byte) ([]byte, error) {
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, fmt.Errorf("open from bytes: %w", err)
	}
	defer f.Close()

	return p.process(f)
}

// process runs all sheet handlers and serializes the result to bytes.
func (p *Processor) process(f *excelize.File) ([]byte, error) {
	for _, sheet := range f.GetSheetList() {
		if err := p.processSheet(f, sheet); err != nil {
			return nil, fmt.Errorf("sheet %q: %w", sheet, err)
		}
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("write to buffer: %w", err)
	}

	return buf.Bytes(), nil
}

func (p *Processor) processSheet(f *excelize.File, sheet string) error {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("get rows: %w", err)
	}

	for row := range rows {
		for col := range rows[row] {
			value := rows[row][col]
			if value == "" {
				continue
			}

			if _, err := p.registry.Process(f, sheet, row, col, value); err != nil {
				return fmt.Errorf("cell %s: %w", excel.CellName(row, col), err)
			}
		}
	}

	return nil
}
