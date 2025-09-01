package writers

import (
	"encoding/json"
	"fmt"
	"log/slog"
	"os"

	"github.com/xuri/excelize/v2"
)

const (
// mainSheet = "Основной файл"
// infoSheet = "Информация"
)

type ExcelWriter struct {
	FilePath    string
	mappingData *MappingData
}

func NewExcelWriter(mappingPath string, filePath string) (*ExcelWriter, error) {
	var excelWriter *ExcelWriter
	excelWriter.FilePath = filePath
	if err := excelWriter.ChangeTemplate(mappingPath); err != nil {
		slog.Error("failed to change template", "error", err)
		return nil, err
	}
	return excelWriter, nil
}

type SheetInfo struct {
	Name string         `json:"name"`
	Map  map[string]int `json:"map"`
}

type MappingData struct {
	FilePath  string       `json:"file_path"`
	SheetInfo []*SheetInfo `json:"sheet_info"`
}

func (w *ExcelWriter) ChangeTemplate(mappingPath string) error {
	if _, err := os.Stat(mappingPath); os.IsNotExist(err) {
		return fmt.Errorf("template file not found: %s", mappingPath)
	}
	data, err := os.ReadFile(mappingPath)
	if err != nil {
		return fmt.Errorf("failed to read template file: %w", err)
	}

	var mappingData MappingData
	if err := json.Unmarshal(data, &mappingData); err != nil {
		return fmt.Errorf("failed to unmarshal template data: %w", err)
	}
	w.mappingData = &mappingData
	return nil
}
func (w *ExcelWriter) GetMapping() *MappingData {
	return w.mappingData
}

func (w *ExcelWriter) WriteByMap(rows []map[string]interface{}, sheetName string) error {
	// Открыть файл
	f, err := excelize.OpenFile(w.FilePath)
	if err != nil {
		return fmt.Errorf("failed to open excel file: %w", err)
	}
	defer f.Close()
	var sheetInfo *SheetInfo
	for _, s := range w.mappingData.SheetInfo {
		if s.Name == sheetName {
			sheetInfo = s
		}
	}
	if sheetInfo == nil {
		return fmt.Errorf("sheet '%s' not found in mapping", sheetName)
	}
	// Проверка, что все поля в данных есть в маппинге
	for _, row := range rows {
		for field := range row {
			if _, ok := sheetInfo.Map[field]; !ok {
				return fmt.Errorf("field '%s' not found in mapping", field)
			}
		}
	}

	// Запись данных
	for i, row := range rows {
		excelRow := i + 2 // +2 потому что строка 1 — это заголовок
		for field, value := range row {
			col := sheetInfo.Map[field]
			cell, err := excelize.CoordinatesToCellName(col, excelRow)
			if err != nil {
				return fmt.Errorf("failed to convert coordinates to cell name: %w", err)
			}
			if err := f.SetCellValue(sheetInfo.Name, cell, value); err != nil {
				return fmt.Errorf("failed to set cell %s: %w", cell, err)
			}
		}
	}

	// Сохранение
	if err := f.SaveAs(w.FilePath); err != nil {
		return fmt.Errorf("failed to save excel file: %w", err)
	}
	return nil
}
