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

	if err := os.WriteFile(*output, data, 0644); err != nil {
		fmt.Fprintf(os.Stderr, "save: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("done:", *output)
}

func step1(input string) ([]byte, error) {
	registry := template.New()
	template.RegisterDefaults(registry)

	employees := domain.GenerateEmployees(employeeCount)
	template.RegisterEmployeeHandler(registry, employees)

	return processor.New(registry).ProcessFile(input)
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
