package reader

import (
	"encoding/json"
	"fmt"
	"os"
	"test/internal/models"

	"github.com/xuri/excelize/v2"
)

type ExcelReader struct {
	FilePath    string
	mappingData *MappingData
}

type SheetInfo struct {
	Name string         `json:"name"`
	Map  map[string]int `json:"map"`
}

type MappingData struct {
	FilePath  string       `json:"file_path"`
	SheetInfo []*SheetInfo `json:"sheet_info"`
}

func (r *ExcelReader) ChangeTemplate(mappingPath string) error {
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
	r.mappingData = &mappingData
	return nil
}
func (r *ExcelReader) GetMapping() *MappingData {
	return r.mappingData
}

func (r *ExcelReader) ReadByMap(sheetName string) ([]*models.Camera, error) {
	var m map[string]int
	cameras := make([]*models.Camera, 0)
	if r.mappingData == nil {
		return nil, fmt.Errorf("mapping data not initialized")
	}
	for _, sheet := range r.mappingData.SheetInfo {
		if sheet.Name == sheetName {
			m = sheet.Map
		}
	}
	if m == nil {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	f, err := excelize.OpenFile(r.FilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}
	for _, row := range rows {
		camera := &models.Camera{}
		camera.Rtsp = row[m["Rtsp"]]
		camera.Name = row[m["Name"]]
		camera.Location = row[m["Location"]]
		cameras = append(cameras, camera)
	}
	return cameras, nil
}
