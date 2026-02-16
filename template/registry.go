package template

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// HandlerFunc processes a matched template variable in an Excel cell.
// It receives the file, sheet name, 0-based row/col indices, and the raw cell value.
type HandlerFunc func(f *excelize.File, sheet string, row, col int, value string) error

// Registry holds template pattern → handler mappings.
type Registry struct {
	handlers []entry
}

type entry struct {
	pattern string
	handler HandlerFunc
}

// New creates an empty Registry.
func New() *Registry {
	return &Registry{}
}

// Register adds a handler for the given pattern (e.g. "{{days}}").
// Handlers are checked in registration order; the first match wins.
func (r *Registry) Register(pattern string, handler HandlerFunc) {
	r.handlers = append(r.handlers, entry{pattern: pattern, handler: handler})
}

// Process checks the cell value against all registered patterns.
// If a match is found, the corresponding handler is called.
// Returns true if a handler was executed.
func (r *Registry) Process(f *excelize.File, sheet string, row, col int, value string) (bool, error) {
	for _, e := range r.handlers {
		if strings.Contains(value, e.pattern) {
			if err := e.handler(f, sheet, row, col, value); err != nil {
				return false, err
			}

			return true, nil
		}
	}

	return false, nil
}
