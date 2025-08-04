package main

import (
	"fmt"
	"log/slog"
	"os"
	"strings"
	"test/pinger"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	mainSheet = "Основной файл"
	infoSheet = "Информация"
)

type Camera struct {
	Rtsp        string
	Name        string
	Ip          string
	Location    string
	DateAdded   string
	Coordinates string
}

func main() {
	logFile, err := os.OpenFile("./archive/logs/app.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		panic("не удалось открыть файл логов: " + err.Error())
	}

	handler := slog.NewTextHandler(logFile, &slog.HandlerOptions{
		Level: slog.LevelInfo,
	})

	logger := slog.New(handler)

	slog.SetDefault(logger)

	err = Ping()
	if err != nil {
		slog.Error("Ping error", "error", err.Error())
	}
	ticker := time.NewTicker(30 * time.Minute)
	for {
		select {
		case <-ticker.C:
			err := Ping()
			if err != nil {
				slog.Error("Ping error", "error", err.Error())
			}
		}
	}

}

func Ping() error {
	slog.Info("Start ping")
	cameras := make([]Camera, 0)
	dubles := make(map[string]int)
	ips := make([]string, 0)
	ipCamera := make(map[string][]Camera)
	ipFileName := generateExcelFilename(int(time.Now().Month()))
	if !fileExists(ipFileName) {
		err := CreateMainExcelFile(ipFileName)
		if err != nil {
			fmt.Println("Error creating Excel file:", err)
			return err
		}
	}
	f, err := excelize.OpenFile(ipFileName)
	if err != nil {
		return fmt.Errorf("can't open file %s, with error %s", ipFileName, err.Error())
	}
	rows, err := f.GetRows(mainSheet)
	if err != nil {
		return fmt.Errorf("can't open read sheet %s, with error %s", mainSheet, err)
	}
	for i, row := range rows {
		if i == 0 || i == 1 {
			continue
		}

		cameras = append(cameras, Camera{Name: row[1], Ip: strings.TrimSpace(row[3]), Location: row[2]})
		if dubles[row[1]] > 0 {
			fmt.Println(row[1])
		} else {
			dubles[row[1]]++
		}
		ipCamera[strings.TrimSpace(row[3])] = append(ipCamera[row[3]], Camera{Name: row[1], Ip: strings.TrimSpace(row[3]), Location: row[2]})
	}
	for ip, _ := range ipCamera {
		ips = append(ips, ip)
	}

	pinger, err := pinger.NewPingManager(ips)
	if err != nil {
		return fmt.Errorf("can't create ping manager with error %s", err.Error())
	}
	results, comments := pinger.Start()
	for ip, result := range results {
		slog.Debug(fmt.Sprintf("%s: %t\n", ip, result))
	}
	for ip, comment := range comments {
		slog.Debug(fmt.Sprintf("%s: %s\n", ip, comment))
	}
	if err := infoToExcel(results, comments, ipCamera); err != nil {
		return fmt.Errorf("can't write info to excel with error %s", err.Error())
	}
	return nil
}

type Info struct {
	result  bool
	comment string
}

func fileExists(path string) bool {
	_, err := os.Stat(path)
	if os.IsNotExist(err) {
		return false
	}
	return err == nil
}

func infoFileName() string {
	now := time.Now()
	return fmt.Sprintf("./archive/архив_%02d.%02d.%d.xlsx", now.Day(), now.Month(), now.Year())
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

func getRussianMonth(m time.Month) string {
	months := map[time.Month]string{
		time.January:   "Январь",
		time.February:  "Февраль",
		time.March:     "Март",
		time.April:     "Апрель",
		time.May:       "Май",
		time.June:      "Июнь",
		time.July:      "Июль",
		time.August:    "Август",
		time.September: "Сентябрь",
		time.October:   "Октябрь",
		time.November:  "Ноябрь",
		time.December:  "Декабрь",
	}
	return months[m]
}

func generateExcelFilename(month int) string {
	now := time.Now()
	monthUpper := getRussianMonth(time.Month(month))
	year := now.Year()
	return fmt.Sprintf("C:/Users/user/Documents/Bitrix24-koordinator@itso.su@itso.bitrix24.ru/Мониторинг и обслуживание/Свод %s %d.xlsx", monthUpper, year)
}

func CreateMainExcelFile(filename string) error {
	readFile, err := excelize.OpenFile(generateExcelFilename(int(time.Now().Month()) - 1))

	if err != nil {
		return err
	}
	cols, err := readFile.GetCols(mainSheet)
	if err != nil {
		return err
	}
	cameras := make([]Camera, len(cols[1]))
	for i := 1; i < len(cols[1]); i++ {
		cameras[i] = Camera{
			Rtsp:        cols[0][i],
			Name:        cols[1][i],
			Location:    cols[2][i],
			Ip:          cols[3][i],
			DateAdded:   cols[4][i],
			Coordinates: cols[5][i],
		}
	}
	newFile := excelize.NewFile()
	if _, err := newFile.NewSheet(mainSheet); err != nil {
		return fmt.Errorf("New sheet error: %v", err)
	}
	if err := newFile.DeleteSheet("Sheet1"); err != nil {
		return fmt.Errorf("Delete sheet error: %v", err)
	}
	slog.Info("Adding cameras to Excel file")

	for i, camera := range cameras {

		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("A%d", i+1), camera.Rtsp); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("B%d", i+1), camera.Name); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("C%d", i+1), camera.Location); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("D%d", i+1), camera.Ip); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("E%d", i+1), camera.DateAdded); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
		if err := newFile.SetCellValue(mainSheet, fmt.Sprintf("F%d", i+1), camera.Coordinates); err != nil {
			return fmt.Errorf("Set cell value error: %v", err)
		}
	}
	if err := newFile.SetCellValue(mainSheet, "A1", "RTSP"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SetCellValue(mainSheet, "B1", "Название"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SetCellValue(mainSheet, "C1", "Объект"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SetCellValue(mainSheet, "D1", "IP"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SetCellValue(mainSheet, "E1", "Дата Добавления"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SetCellValue(mainSheet, "F1", "Координаты"); err != nil {
		return fmt.Errorf("Set cell value error: %v", err)
	}
	if err := newFile.SaveAs(generateExcelFilename(int(time.Now().Month()))); err != nil {
		return err
	}
	return nil
}

func CreateExcelFile(ipCamera map[string][]Camera) error {
	f := excelize.NewFile()
	row := 2
	if _, err := f.NewSheet(infoSheet); err != nil {
		return err
	}
	if err := f.DeleteSheet("Sheet1"); err != nil {
		return err
	}
	if err := f.SetCellValue(infoSheet, "A1", "Имя"); err != nil {
		return err
	}
	if err := f.SetCellValue(infoSheet, "B1", "Объект"); err != nil {
		return err
	}
	if err := f.SetCellValue(infoSheet, "C1", "IP"); err != nil {
		return err
	}
	for _, cameras := range ipCamera {
		for _, camera := range cameras {
			if err := f.SetCellValue(infoSheet, fmt.Sprintf("A%d", row), camera.Name); err != nil {
				return err
			}

			if err := f.SetCellValue(infoSheet, fmt.Sprintf("B%d", row), camera.Location); err != nil {
				return err
			}
			if err := f.SetCellValue(infoSheet, fmt.Sprintf("C%d", row), camera.Ip); err != nil {
				return err
			}
			row++
		}
	}
	if err := f.SaveAs(infoFileName()); err != nil {
		return err
	}
	return nil
}

func infoToExcel(results map[string]bool, comments map[string]string, ipCamera map[string][]Camera) error {
	IpInfo := make(map[string]Info)
	infoFileName := infoFileName()
	if !fileExists(infoFileName) {
		err := CreateExcelFile(ipCamera)
		if err != nil {
			return fmt.Errorf("can't create file %s, with error %s", infoFileName, err.Error())
		}
	}
	time.Sleep(1 * time.Second)
	infoFile, err := excelize.OpenFile(infoFileName)
	if err != nil {
		return fmt.Errorf("can't open file %s, with error %s", infoFileName, err.Error())
	}
	for ip, cameras := range ipCamera {
		for _, camera := range cameras {
			IpInfo[camera.Ip] = Info{result: results[ip], comment: comments[ip]}
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
		ip, err := infoFile.GetCellValue(infoSheet, fmt.Sprintf("C%d", row))
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
		info := IpInfo[ip]
		if ip == "" {
			info = Info{false, "Нет ip у камеры"}
		}
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
