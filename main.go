package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/orayew2002/rast-excel/domain"
	"github.com/orayew2002/rast-excel/processor"
	"github.com/orayew2002/rast-excel/template"
)

func main() {
	input := flag.String("input", "table.xlsx", "path to the input Excel file")
	output := flag.String("output", "result.xlsx", "path to the output Excel file")
	flag.Parse()

	registry := template.New()
	template.RegisterDefaults(registry)

	// For testing: use fake employees. In your real project, pass your own []domain.Employee.
	employees := domain.GenerateEmployees(25)
	template.RegisterEmployeeHandler(registry, employees)

	p := processor.New(registry)
	if _, err := p.ProcessFile(*input, *output); err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("done:", *output)
}
