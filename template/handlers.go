package template

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/orayew2002/rast-excel/domain"
	"github.com/orayew2002/rast-excel/excel"
	"github.com/xuri/excelize/v2"
)

// RegisterDefaults registers the built-in template handlers (days, working_time).
func RegisterDefaults(r *Registry) {
	r.Register("{{days}}", handleDays)
	r.Register("{{working_time}}", handleWorkingTime)
}

// ---------- Employee columns ----------

// columnDef describes one fixed employee column: value extractor + style.
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

// AttendanceStartCol returns the 0-based column index where attendance data begins
// for an employee section that starts at employeeCol.
func AttendanceStartCol(employeeCol int) int {
	return employeeCol + len(columns)
}

// ---------- RegisterEmployeeHandler ----------

// RegisterEmployeeHandler registers the {{start_process}} handler.
// It writes employee rows (fixed columns + attendance) into the sheet,
// replacing the template row. No formulas are written here — use
// RegisterFormulaHandler in a second pass for that.
func RegisterEmployeeHandler(r *Registry, employees []domain.Employee) {
	r.Register("{{start_process}}", func(f *excelize.File, sheet string, row, col int, _ string) error {
		return writeEmployees(f, sheet, row, col, employees)
	})
}

func writeEmployees(f *excelize.File, sheet string, row, col int, employees []domain.Employee) error {
	if err := f.RemoveRow(sheet, row+1); err != nil {
		return fmt.Errorf("remove template row: %w", err)
	}

	if err := f.InsertRows(sheet, row+1, len(employees)); err != nil {
		return fmt.Errorf("insert rows: %w", err)
	}

	sm := NewStyleManager(f)

	for i, emp := range employees {
		empRow := row + i
		if err := writeEmployeeRow(f, sm, sheet, empRow, col, emp); err != nil {
			return fmt.Errorf("employee %d: %w", emp.Id, err)
		}
	}

	return nil
}

func writeEmployeeRow(f *excelize.File, sm *StyleManager, sheet string, row, col int, emp domain.Employee) error {
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

// ---------- RegisterFormulaHandler ----------

// FormulaKey pairs a template placeholder with an optional formula generator.
//
// Key is the placeholder in the Excel template (e.g. "{{d}}").
// FormulaFn receives the attendance cell range and returns the formula string.
// Set FormulaFn to nil for a style-only key (e.g. "{{}}") that applies
// the centered style to each employee cell without writing a formula.
type FormulaKey struct {
	Key       string
	FormulaFn func(attRange string) string
}

// CountIFFormula returns a FormulaFn that counts occurrences of symbol across
// an attendance range, multiplied by value.
//
//	symbol "W", value 1 → SUMPRODUCT((range="W")*1)  — counts each "W"
func CountIFFormula(symbol string, value int) func(string) string {
	return func(cellRange string) string {
		return fmt.Sprintf(`SUMPRODUCT((%s="%s")*%d)`, cellRange, symbol, value)
	}
}

// SumNumFormula returns a FormulaFn that sums all numeric values in the
// attendance range, ignoring non-numeric cells.
//
//	"8", "W", "8" → 8 + 0 + 8 = 16
func SumNumFormula() func(string) string {
	return func(cellRange string) string {
		return fmt.Sprintf(`IFERROR(SUMPRODUCT(IFERROR(VALUE(%s),0)),0)`, cellRange)
	}
}

// CountNumFormula returns a FormulaFn that counts how many cells in the
// attendance range contain a number, ignoring non-numeric cells.
//
//	"8", "W", "8" → 1 + 0 + 1 = 2
func CountNumFormula() func(string) string {
	return func(cellRange string) string {
		return fmt.Sprintf(`IFERROR(SUMPRODUCT(IFERROR(VALUE(%s)*0+1,0)),0)`, cellRange)
	}
}

// combFormulaHandler is shared across all formula key registrations.
// When any registered key is found in a cell, it combines the formulas
// of ALL keys present in that cell and writes one formula per employee row.
type combFormulaHandler struct {
	employeeCount int
	attStart      int // 0-based column where attendance data begins
	keys          []FormulaKey
	sm            *StyleManager        // lazily initialized on first handle call
	removedRows   map[string]struct{}  // tracks formula rows already removed
}

func (h *combFormulaHandler) handle(f *excelize.File, sheet string, row, col int, value string) error {
	if h.sm == nil {
		h.sm = NewStyleManager(f)
	}
	if h.removedRows == nil {
		h.removedRows = make(map[string]struct{})
	}

	centeredStyle, err := h.sm.Centered()
	if err != nil {
		return fmt.Errorf("formula cell style: %w", err)
	}

	attEnd := h.attStart + currentMonthDays() - 1
	firstEmpRow := row - h.employeeCount

	for empRow := firstEmpRow; empRow < row; empRow++ {
		attRange := excel.CellName(empRow, h.attStart) + ":" + excel.CellName(empRow, attEnd)

		formula, matched := h.buildFormula(value, attRange)
		if !matched {
			continue
		}

		cell := excel.CellName(empRow, col)
		if formula != "" {
			if err := f.SetCellFormula(sheet, cell, formula); err != nil {
				return fmt.Errorf("set formula at %s: %w", cell, err)
			}
		}
		if err := f.SetCellStyle(sheet, cell, cell, centeredStyle); err != nil {
			return fmt.Errorf("set style at %s: %w", cell, err)
		}
	}

	// Remove the template formula row exactly once — multiple keys can live in
	// the same row (e.g. {{t}}, {{d}}, {{w}}), so the handler is called once per
	// cell. Tracking ensures RemoveRow is called only on the first hit.
	key := fmt.Sprintf("%s:%d", sheet, row)
	if _, done := h.removedRows[key]; !done {
		h.removedRows[key] = struct{}{}
		if err := f.RemoveRow(sheet, row+1); err != nil {
			return fmt.Errorf("remove formula row: %w", err)
		}
	}

	return nil
}

// buildFormula collects formulas from all keys found in value, returning the
// combined wrapped formula and whether any key matched. If all matching keys
// have nil FormulaFn (style-only), formula is empty but matched is true.
func (h *combFormulaHandler) buildFormula(value, attRange string) (formula string, matched bool) {
	var parts []string
	for _, k := range h.keys {
		if !strings.Contains(value, k.Key) {
			continue
		}
		matched = true
		if k.FormulaFn != nil {
			parts = append(parts, k.FormulaFn(attRange))
		}
	}
	if len(parts) == 0 {
		return "", matched
	}
	combined := strings.Join(parts, "+")
	return fmt.Sprintf(`IF(%s=0,"",(%s))`, combined, combined), true
}

// RegisterFormulaHandler registers per-employee formula handlers for each key.
//
// When the processor finds a template cell containing any of the registered keys,
// it writes the appropriate Excel formula into each of the employeeCount rows
// directly above that cell (one formula per employee). A cell may contain
// multiple keys (e.g. "{{d}}{{t}}") — the resulting formulas are combined with "+".
//
// attStart is the 0-based column index where employee attendance data begins
// (use AttendanceStartCol to compute it).
//
// Example:
//
//	template.RegisterFormulaHandler(registry, 25, template.AttendanceStartCol(0), []template.FormulaKey{
//	    {Key: "{{d}}", FormulaFn: template.CountIFFormula("8", 8)},
//	    {Key: "{{w}}", FormulaFn: template.CountIFFormula("W", 1)},
//	    {Key: "{{t}}", FormulaFn: func(r string) string {
//	        return fmt.Sprintf(`SUMPRODUCT(IFERROR(VALUE(%s),(%s<>"")*1))`, r, r)
//	    }},
//	})
func RegisterFormulaHandler(r *Registry, employeeCount, attStart int, keys []FormulaKey) {
	h := &combFormulaHandler{
		employeeCount: employeeCount,
		attStart:      attStart,
		keys:          keys,
	}
	for _, k := range keys {
		r.Register(k.Key, h.handle)
	}
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

		headerStyleID, _ := f.GetCellStyle(sheet, topLeft)

		if err := f.MergeCell(sheet, topLeft, bottomRight); err != nil {
			return fmt.Errorf("merge row %d: %w", headerRow, err)
		}

		if headerStyleID != 0 {
			if err := f.SetCellStyle(sheet, topLeft, bottomRight, headerStyleID); err != nil {
				return fmt.Errorf("header style row %d: %w", headerRow, err)
			}
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

// ---------- ReplaceHandler ----------

// ReplaceHandler accumulates key→value pairs and registers a single shared
// handler for all of them. Because the registry stops at the first matched
// handler per cell, sharing one handler ensures ALL pairs are replaced in one
// pass — even when a cell contains several keys at once (e.g. "{{year}} {{month}}").
//
// Usage:
//
//	rh := template.NewReplaceHandler()
//	rh.Add("{{start_year}}", "2026")
//	rh.Add("{{month_tk}}", "Февраль")
//	rh.Register(registry)
type ReplaceHandler struct {
	pairs []replacePair
}

type replacePair struct{ key, val string }

// NewReplaceHandler creates an empty ReplaceHandler.
func NewReplaceHandler() *ReplaceHandler {
	return &ReplaceHandler{}
}

// Add appends a key→val pair. Returns h so calls can be chained.
func (h *ReplaceHandler) Add(key, val string) *ReplaceHandler {
	h.pairs = append(h.pairs, replacePair{key, val})
	return h
}

// Register registers h into r for every key added via Add.
// All keys share the same underlying handler, so whichever key triggers first
// causes all pairs to be replaced in the cell.
func (h *ReplaceHandler) Register(r *Registry) {
	for _, p := range h.pairs {
		r.Register(p.key, h.apply)
	}
}

func (h *ReplaceHandler) apply(f *excelize.File, sheet string, row, col int, value string) error {
	cell := excel.CellName(row, col)

	styleID, _ := f.GetCellStyle(sheet, cell)

	replaced := value
	for _, p := range h.pairs {
		replaced = strings.ReplaceAll(replaced, p.key, p.val)
	}

	if err := f.SetCellStr(sheet, cell, replaced); err != nil {
		return fmt.Errorf("replace handler: %w", err)
	}

	if styleID != 0 {
		if err := f.SetCellStyle(sheet, cell, cell, styleID); err != nil {
			return fmt.Errorf("restore style: %w", err)
		}
	}

	return nil
}

// RegisterReplaceHandler is a convenience wrapper for a single key→val pair.
// For cells that contain multiple keys, use NewReplaceHandler instead.
func RegisterReplaceHandler(r *Registry, key, val string) {
	NewReplaceHandler().Add(key, val).Register(r)
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

// ---------- RegisterMarksHandler ----------

// RegisterMarksHandler registers a handler for {{marks_list}}.
//
// The template cell containing {{marks_list}} must be a merged cell that
// spans the full width of the marks list (e.g. A1:E1). The handler:
//  1. Detects the merge range of the placeholder cell automatically.
//  2. Copies the cell style from the placeholder.
//  3. Removes the template row and inserts one row per mark.
//  4. For each mark: re-applies the same merge, writes the mark content
//     into the merged cell with the copied style.
//
// The underscore separator between Name and Key is computed dynamically so
// that every row has the same total rune width — the Key always appears at
// the same right-edge position regardless of how long the Name is.
//
// Example:
//
//	template.RegisterMarksHandler(registry, []domain.Mark{
//	    {Name: "Dynç alyş we baýramçylyk günler", Key: "B"},
//	    {Name: "Kanuna laýyk işe gelmezlik",      Key: "C"},
//	})
func RegisterMarksHandler(r *Registry, marks []domain.Mark) {
	r.Register("{{marks_list}}", func(f *excelize.File, sheet string, row, col int, _ string) error {
		return writeMarks(f, sheet, row, col, marks)
	})
}

// numericKey reports whether every rune in key is a digit (e.g. "8", "12").
func numericKey(key string) bool {
	if key == "" {
		return false
	}
	for _, r := range key {
		if !unicode.IsDigit(r) {
			return false
		}
	}
	return true
}

func writeMarks(f *excelize.File, sheet string, row, col int, marks []domain.Mark) error {
	// Skip marks whose Key is a plain number (e.g. "8" for worked hours).
	filtered := marks[:0:0]
	for _, m := range marks {
		if !numericKey(m.Key) {
			filtered = append(filtered, m)
		}
	}
	marks = filtered

	placeholder := excel.CellName(row, col)

	// Capture style before the row is removed.
	styleID, _ := f.GetCellStyle(sheet, placeholder)

	// Detect the merge range of the placeholder cell.
	// mergeEndCol stays at col when the cell is not merged.
	mergeEndCol := col
	if merges, err := f.GetMergeCells(sheet); err == nil {
		for _, mc := range merges {
			if mc.GetStartAxis() == placeholder {
				if endCol, _, err := excelize.CellNameToCoordinates(mc.GetEndAxis()); err == nil {
					mergeEndCol = endCol - 1 // excelize returns 1-based; convert to 0-based
				}
				break
			}
		}
	}

	if err := f.RemoveRow(sheet, row+1); err != nil {
		return fmt.Errorf("marks: remove template row: %w", err)
	}

	// Excel templates sometimes contain phantom row elements near R=1048576
	// (an artifact of normal editing). InsertRows fails with ErrMaxRows when
	// any such row's index + n would exceed the limit. Sweeping from the bottom
	// with RemoveRow (offset=-1) is always safe: newRow = R-1 never overflows.
	if err := removePhantomRows(f, sheet, len(marks)); err != nil {
		return fmt.Errorf("marks: clean phantom rows: %w", err)
	}

	if err := f.InsertRows(sheet, row+1, len(marks)); err != nil {
		return fmt.Errorf("marks: insert rows: %w", err)
	}

	// Compute target rune-width so every row's Name+pad+Key has identical length.
	// This keeps the Key abbreviation right-aligned at the same position for all marks.
	const minPad = 4
	targetWidth := 0
	for _, m := range marks {
		w := len([]rune(m.Name)) + len([]rune(m.Key))
		if w > targetWidth {
			targetWidth = w
		}
	}
	targetWidth += minPad

	for i, m := range marks {
		r := row + i
		startCell := excel.CellName(r, col)
		endCell := excel.CellName(r, mergeEndCol)

		// Re-apply the merge for this row.
		if mergeEndCol > col {
			if err := f.MergeCell(sheet, startCell, endCell); err != nil {
				return fmt.Errorf("marks[%d] merge: %w", i, err)
			}
		}

		padLen := targetWidth - len([]rune(m.Name)) - len([]rune(m.Key))
		if padLen < 1 {
			padLen = 1
		}
		content := m.Name + strings.Repeat("_", padLen) + m.Key
		if err := f.SetCellStr(sheet, startCell, content); err != nil {
			return fmt.Errorf("marks[%d] value: %w", i, err)
		}

		if err := f.SetCellStyle(sheet, startCell, endCell, styleID); err != nil {
			return fmt.Errorf("marks[%d] style: %w", i, err)
		}
	}

	return nil
}

// ---------- RegisterMergeHandler ----------

var mergeCodePat = regexp.MustCompile(`\[(\d+):(\d+)\]`)

// RegisterMergeHandler registers a handler that detects [extraRows:extraCols] codes
// embedded in cell values, strips the code, and merges the cell with its neighbours.
//
//	[1:0] → merge with 1 row below, no extra cols
//	[1:1] → merge with 1 row below and 1 col to the right
//	[0:2] → merge 2 cols to the right (horizontal only)
//	[0:0] → strip code only, no merge
//
// Run this in a separate pass (after all row/col insertions are done) so the
// row indices are stable.
func RegisterMergeHandler(r *Registry) {
	r.Register("[", handleMergeCode)
}

func handleMergeCode(f *excelize.File, sheet string, row, col int, value string) error {
	m := mergeCodePat.FindStringSubmatch(value)
	if m == nil {
		return nil // "[" present but not a merge code — skip
	}

	extraRows, _ := strconv.Atoi(m[1])
	extraCols, _ := strconv.Atoi(m[2])

	cleaned := mergeCodePat.ReplaceAllString(value, "")

	cell := excel.CellName(row, col)
	styleID, _ := f.GetCellStyle(sheet, cell)

	if err := f.SetCellStr(sheet, cell, cleaned); err != nil {
		return fmt.Errorf("merge handler: set value: %w", err)
	}

	if extraRows == 0 && extraCols == 0 {
		return nil // nothing to merge
	}

	bottomRight := excel.CellName(row+extraRows, col+extraCols)
	if err := f.MergeCell(sheet, cell, bottomRight); err != nil {
		return fmt.Errorf("merge handler: merge: %w", err)
	}

	if styleID != 0 {
		if err := f.SetCellStyle(sheet, cell, bottomRight, styleID); err != nil {
			return fmt.Errorf("merge handler: style: %w", err)
		}
	}

	return nil
}

// ---------- helpers ----------

// removePhantomRows sweeps the last n rows of the sheet using RemoveRow.
// RemoveRow uses offset=-1, so newRow = R-1 which can never exceed TotalRows —
// it is always safe. This clears any phantom row elements that Excel leaves near
// R=1048576 as editing artifacts, which would otherwise cause InsertRows to
// return ErrMaxRows when n rows are inserted anywhere in the sheet.
func removePhantomRows(f *excelize.File, sheet string, n int) error {
	const totalRows = 1048576
	for r := totalRows; r > totalRows-n; r-- {
		if err := f.RemoveRow(sheet, r); err != nil {
			return err
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
