package excel

import "fmt"

// CellName converts 0-based row and column indices to an Excel cell reference (e.g. 0,0 → "A1").
func CellName(row, col int) string {
	return fmt.Sprintf("%s%d", IndexToColumn(col), row+1)
}

// IndexToColumn converts a 0-based column index to Excel column letters (0→A, 25→Z, 26→AA).
func IndexToColumn(n int) string {
	result := ""
	for n >= 0 {
		result = string(rune('A'+(n%26))) + result
		n = n/26 - 1
	}
	return result
}
