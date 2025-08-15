package excel

import "test/internal/models"

type ExcelManager struct {
	Creator *Creator
}

type Creator interface {
	CreateFromTemplate(oldFilename, newFilename, sheet string) error
	CreateFromMap(rtspCamera map[string][]models.Camera, sheet, filename string) error
}
