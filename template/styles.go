package template

import "github.com/xuri/excelize/v2"

// StyleManager caches Excel styles so each style is created only once per file.
type StyleManager struct {
	file  *excelize.File
	cache map[string]int
}

// NewStyleManager creates a style manager bound to the given file.
func NewStyleManager(f *excelize.File) *StyleManager {
	return &StyleManager{file: f, cache: make(map[string]int)}
}

// Centered returns a center-aligned bordered style (cached).
func (sm *StyleManager) Centered() (int, error) {
	return sm.getOrCreate("centered", &excelize.Style{
		Font:      defaultFont(),
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border:    defaultBorder(),
	})
}

// Left returns a left-aligned bordered style (cached).
func (sm *StyleManager) Left() (int, error) {
	return sm.getOrCreate("left", &excelize.Style{
		Font:      defaultFont(),
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Border:    defaultBorder(),
	})
}

func (sm *StyleManager) getOrCreate(key string, style *excelize.Style) (int, error) {
	if id, ok := sm.cache[key]; ok {
		return id, nil
	}

	id, err := sm.file.NewStyle(style)
	if err != nil {
		return 0, err
	}

	sm.cache[key] = id
	return id, nil
}

func defaultFont() *excelize.Font {
	return &excelize.Font{Family: "Times New Roman", Size: 11}
}

func defaultBorder() []excelize.Border {
	return []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
	}
}
