package excel

import "test/internal/models"

type ExcelManager struct {
	Creator *Creator
	Writer  *Writer
}

type Creator interface {
	CreateFromTemplate(oldFilename, newFilename, sheet string) error
	CreateFromMap(rtspCamera map[string][]models.Camera, sheet, filename string) error
}

type Writer interface {
	WriteByMap(rows []map[string]interface{}, sheetName string) error
	WriteColumn(rows map[string]interface{}, sheetName string) error
}

func NewExcelManager(writer Writer, creator Creator) *ExcelManager {
	return &ExcelManager{
		Creator: &creator,
		Writer:  &writer,
	}
}
