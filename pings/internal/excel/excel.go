package excel

import (
	"fmt"
	"log/slog"
	"os"
	"test/internal/models"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	mainSheet = "Основной файл"
	infoSheet = "Информация"
)

type Info struct {
	result  bool
	comment string
}

func InfoToExcel(results map[string]bool, comments map[string]string, rtspCamera map[string][]models.Camera) error {
	RtspInfo := make(map[string]Info)
	infoFileName := infoFileName()
	if !FileExists(infoFileName) {
		err := CreateExcelFile(rtspCamera)
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
	col, err := findNextEmptyColumnIndex(infoFile, infoSheet, 1)
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
func FileExists(path string) bool {
	_, err := os.Stat(path)
	if os.IsNotExist(err) {
		return false
	}
	return err == nil
}

func findNextEmptyColumnIndex(f *excelize.File, sheet string, row int) (int, error) {
	for col := 1; col <= 100; col++ {
		cell, _ := excelize.CoordinatesToCellName(col, row)
		val, _ := f.GetCellValue(sheet, cell)
		if val == "" {
			return col, nil
		}
	}
	return 0, fmt.Errorf("нет пустых колонок")
}
