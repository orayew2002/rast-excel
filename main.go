package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/orayew2002/rast-excel/domain"
	"github.com/orayew2002/rast-excel/processor"
	"github.com/orayew2002/rast-excel/template"
)

const employeeCount = 25

func main() {
	input := flag.String("input", "table.xlsx", "path to the input Excel file")
	output := flag.String("output", "result.xlsx", "path to the output Excel file")
	flag.Parse()

	// Step 1: inject days, working_time, and employee attendance rows.
	data, err := step1(*input)
	if err != nil {
		fmt.Fprintf(os.Stderr, "step1: %v\n", err)
		os.Exit(1)
	}

	// Step 2: write per-employee formulas for any {{key}} cells below the employee block.
	data, err = step2(data)
	if err != nil {
		fmt.Fprintf(os.Stderr, "step2: %v\n", err)
		os.Exit(1)
	}

	// Step 3: apply [rowSpan:colSpan] merge codes embedded in cell values.
	data, err = step3(data)
	if err != nil {
		fmt.Fprintf(os.Stderr, "step3: %v\n", err)
		os.Exit(1)
	}

	// Step 4: apply borders to &1…&1 ranges.
	data, err = step4(data)
	if err != nil {
		fmt.Fprintf(os.Stderr, "step4: %v\n", err)
		os.Exit(1)
	}

	if err := os.WriteFile(*output, data, 0644); err != nil {
		fmt.Fprintf(os.Stderr, "save: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("done:", *output)
}

var marks = []domain.Mark{
	{Name: "Dynç alyş we baýramçylyk günler", Key: "B"},
	{Name: "Kanuna laýyk işe gelmezlik", Key: "C"},
	{Name: "Gulluk iş saparlary", Key: "W"},
	{Name: "Nobatdaky we goşmaça rugsatlar", Key: "O"},
	{Name: "Işe ýarawsyzlyk (kesel, karantin we ş.m.)", Key: "Y"},
	{Name: "Gowrelilik sebäpli rugsat", Key: "O"},
	{Name: "Emdiryän eneleriň ýeňillikli sagatlary", Key: "I"},
	{Name: "Saglyga zyýanly önümçilikde işleýän işleriň ýeňillikli sagatlary", Key: "ÝS"},
	{Name: "Iş wagtyndan daşary edilen işiň sagatlary", Key: "IWI"},
	{Name: "Bütin smena boýunça işsiz durmaklyk", Key: "ID"},
	{Name: "Smeniň içindäki işsiz durmaklyk", Key: "IID"},
	{Name: "Sebäpsiz işden galmak", Key: "S"},
	{Name: "Işe gijä galmak we işden wagtyndan öň gitmek", Key: "SIG"},
	{Name: "Kärhanañ çäginden daşary gulluk tabşyryklaryny ýerine ýetirmek", Key: "ÇDG"},
	{Name: "Administrasiýañ rugsady boýunça işe gelmezlik", Key: "AR"},
}

func step1(input string) ([]byte, error) {
	registry := template.New()
	template.RegisterDefaults(registry)

	employees := domain.GenerateEmployees(employeeCount)
	template.RegisterEmployeeHandler(registry, employees)

	template.RegisterMarksHandler(registry, marks)

	return processor.New(registry).ProcessFile(input)
}

func step3(data []byte) ([]byte, error) {
	registry := template.New()
	template.RegisterMergeHandler(registry)
	return processor.New(registry).ProcessBytes(data)
}

func step4(data []byte) ([]byte, error) {
	registry := template.New()
	template.RegisterBorderHandler(registry)
	return processor.New(registry).ProcessBytes(data)
}

func step2(data []byte) ([]byte, error) {
	registry := template.New()

	attStart := template.AttendanceStartCol(0)
	template.RegisterFormulaHandler(registry, employeeCount, attStart, []template.FormulaKey{
		{Key: "{{t}}", FormulaFn: template.CountIFFormula("T", 1)},
		{Key: "{{d}}", FormulaFn: template.CountIFFormula("D", 1)},
		{Key: "{{w}}", FormulaFn: template.CountIFFormula("W", 1)},
		{Key: "{{l}}", FormulaFn: template.CountIFFormula("L", 1)},
		{Key: "{{a}}", FormulaFn: template.CountIFFormula("A", 1)},
		{Key: "{{p}}", FormulaFn: template.CountIFFormula("P", 1)},
		{Key: "{{num_sum}}", FormulaFn: template.SumNumFormula()},
		{Key: "{{num_count}}", FormulaFn: template.CountNumFormula()},
		{Key: "{{}}", FormulaFn: nil},
	})

	return processor.New(registry).ProcessBytes(data)
}
