package writers

import (
	"encoding/json"
	"fmt"
	"log/slog"
	"os"
	"test/internal/models"
	"test/internal/views/excel"
	"time"

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
	Name string            `json:"name"`
	Map  map[string]string `json:"map"`
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

func (w *ExcelWriter) WriteData(rows []map[string]interface{}, sheetName string) error {
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
			cell := fmt.Sprintf("%s%d", col, excelRow)
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

type Info struct {
	result  bool
	comment string
}

func InfoToExcel(results map[string]bool, comments map[string]string, rtspCamera map[string][]models.Camera, infoSheet string, infoFileName string) error {
	RtspInfo := make(map[string]Info)
	if !excel.FileExists(infoFileName) {
		err := CreateExcelFile(rtspCamera, infoSheet, infoFileName)
		if err != nil {
			return fmt.Errorf("can't create file %s, with error %s", infoFileName, err.Error())
		}
	}
	time.Sleep(1 * time.Second)
	infoFile, err := excelize.OpenFile(infoFileName)
	if err != nil {
		return fmt.Errorf("can't open file %s, with error %s", infoFileName, err.Error())
	}
	for ip, cameras := range rtspCamera {
		for _, camera := range cameras {
			RtspInfo[camera.Rtsp] = Info{result: results[ip], comment: comments[ip]}
		}
	}
	col, err := excel.FindNextEmptyColumnIndex(infoFile, infoSheet, 1)
	if err != nil {
		return fmt.Errorf("can't find next empty column index, with error %s", err.Error())
	}
	timeCell1, err := excelize.CoordinatesToCellName(col, 1)
	if err != nil {
		return fmt.Errorf("can't convert coordinates to cell name, with error %s", err.Error())
	}
	timeCell2, err := excelize.CoordinatesToCellName(col+1, 1)
	if err != nil {
		return fmt.Errorf("can't convert coordinates to cell name, with error %s", err.Error())
	}
	slog.Info("Insert into column", "column", col)

	if err := infoFile.MergeCell(infoSheet, timeCell1, timeCell2); err != nil {
		return fmt.Errorf("can't merge cells, with error %s", err.Error())
	}
	if err := infoFile.SetCellValue(infoSheet, timeCell1, time.Now().Format("15:04:05")); err != nil {
		return fmt.Errorf("can't set cell value, with error %s", err.Error())
	}
	for row := 2; row <= 5000; row++ {
		name, err := infoFile.GetCellValue(infoSheet, fmt.Sprintf("B%d", row))
		if err != nil {
			return fmt.Errorf("can't get cell value, with error %s", err.Error())
		}
		if name == "" {
			break
		}
		rtsp, err := infoFile.GetCellValue(infoSheet, fmt.Sprintf("D%d", row))
		if err != nil {
			return fmt.Errorf("can't get cell value, with error %s", err.Error())
		}
		slog.Debug("Processing camera", "name", name)
		cellResult, err := excelize.CoordinatesToCellName(col+1, row)
		if err != nil {
			return fmt.Errorf("can't convert coordinates to cell name, with error %s", err.Error())
		}
		cellComment, err := excelize.CoordinatesToCellName(col, row)
		if err != nil {
			return fmt.Errorf("can't convert coordinates to cell name, with error %s", err.Error())
		}
		info := RtspInfo[rtsp]
		slog.Debug("Ip result", "result", info.result)
		if err := infoFile.SetCellValue(infoSheet, cellResult, info.result); err != nil {
			return fmt.Errorf("can't set cell value, with error %s", err.Error())
		}
		slog.Debug("Ip comment", "comment", info.comment)
		if err := infoFile.SetCellValue(infoSheet, cellComment, info.comment); err != nil {
			return fmt.Errorf("can't set cell value, with error %s", err.Error())
		}
	}
	infoFile.Save()
	return nil
}
