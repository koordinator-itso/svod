package main

import (
	"fmt"
	"log/slog"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	infoSheet = "Информация"
	mainSheet = "Основной файл"
)

type Camera struct {
	Rtsp        string
	Name        string
	Ip          string
	Location    string
	DateAdded   string
	Coordinates string
}

type Info struct {
	result  bool
	comment string
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
func infoFileName() string {
	now := time.Now()
	return fmt.Sprintf("../pings/archive/архив_%02d.%02d.%d.xlsx", now.Day(), now.Month(), now.Year())
}

func fileExists(path string) bool {
	_, err := os.Stat(path)
	if os.IsNotExist(err) {
		return false
	}
	return err == nil
}
func CreateExcelFile(filename string) error {
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
		fmt.Printf("%d:%s\n", i+1, camera)
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

func stringToBool(s string) bool {
	if s == "TRUE" {
		return true
	}
	return false
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

func main() {
	if !fileExists(generateExcelFilename(int(time.Now().Month()))) {
		err := CreateExcelFile(generateExcelFilename(int(time.Now().Month())))
		if err != nil {
			fmt.Println("Error creating Excel file:", err)
			return
		}
	}

	f, err := excelize.OpenFile(generateExcelFilename(int(time.Now().Month())))
	if err != nil {
		slog.Error("Error opening svod file:", "error", err.Error())
	}
	archiveFile, err := excelize.OpenFile(infoFileName())
	if err != nil {
		slog.Error("Error opening info file:", "error", err.Error())
	}

	cameras, err := archiveFile.GetRows(infoSheet)
	if err != nil {
		slog.Error("Error getting rows from info sheet:", "error", err.Error())
	}
	camerasCoordinates := make(map[string]int)
	column, err := findNextEmptyColumnIndex(f, mainSheet, 2)
	if err != nil {
		slog.Error("Error finding next empty column:", "error", err.Error())
	}

	results := make(map[string]int)
	comments := make(map[string]string)
	for i, camera := range cameras {
		if i == 0 {
			continue
		}
		sum := 0
		k := 0
		com := ""
		if len(camera) < 4 {
			results[camera[0]] = 3
			continue
		}
		for j := 4; j < len(camera); j += 2 {
			status := stringToBool(camera[j])
			comment := camera[j-1]
			if status {
				if comment != "" {
					com = comment
				}
				sum += 2

			}
			k++
		}
		if sum >= k {
			results[camera[0]] = 2
			if com != "" {
				results[camera[0]] = 1
				comments[camera[0]] = com
			}
		} else {
			results[camera[0]] = 0
		}
	}
	cams, err := f.GetRows(mainSheet)
	if err != nil {
		slog.Error("Error getting rows:", "error", err.Error())
	}
	for i, cam := range cams {
		if i == 0 {
			continue
		}
		if len(cam) < 2 {
			continue
		}
		camerasCoordinates[cam[1]] = i + 1
	}
	greenStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#C6EFCE"}, // Светло-зелёный
			Pattern: 1,                   // 1 = сплошная заливка
		},
	})
	if err != nil {
		slog.Error("Error creating style:", "error", err.Error())

	}

	redStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFC7CE"}, // Светло-красный
			Pattern: 1,
		},
	})
	if err != nil {
		slog.Error("Error creating style:", "error", err.Error())
	}

	yellowStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFEB9C"}, // Светло-жёлтый
			Pattern: 1,
		},
	})
	if err != nil {
		slog.Error("Error creating style:", "error", err.Error())
	}
	c, err := excelize.CoordinatesToCellName(column, 2)
	if err != nil {
		slog.Error("Error coordinate to cell name:", "error", err.Error())
	}
	f.SetCellValue(mainSheet, c, time.Now().Format("2006-01-02"))
	for camera, status := range results {
		// if status == 1 {
		// 	fmt.Println(camera, comments[camera])
		// }
		// if status == 0 {
		// 	fmt.Println(camera, status)
		// }
		//fmt.Println(camera, status)
		coord, err := excelize.CoordinatesToCellName(column, camerasCoordinates[camera])
		if err != nil {
			slog.Error("Error coordinate to cell name:", "error", err.Error())
		}
		if status == 2 {
			if err := f.SetCellValue(mainSheet, coord, "В сети"); err != nil {
				slog.Error("Error setting cell value:", "error", err.Error())
			}
			if err := f.SetCellStyle(mainSheet, coord, coord, greenStyle); err != nil {
				slog.Error("Error setting cell style:", "error", err.Error())
			}
		}
		if status == 1 {
			if err := f.SetCellValue(mainSheet, coord, fmt.Sprintf("Сомнительная работа. Причина: %s", comments[camera])); err != nil {
				slog.Error("Error setting cell value:", "error", err.Error())
			}
			if err := f.SetCellStyle(mainSheet, coord, coord, yellowStyle); err != nil {
				slog.Error("Error setting cell style:", "error", err.Error())
			}
		}
		if status == 0 {
			previousCoord, err := excelize.CoordinatesToCellName(column-1, camerasCoordinates[camera])
			if err != nil {
				slog.Error("Error coordinate to cell name:", "error", err.Error())
			}
			val, err := f.GetCellValue(mainSheet, previousCoord)

			if err != nil {
				slog.Error("Error getting cell value:", "error", err.Error())
			}
			if strings.HasPrefix(val, "Не в сети не по нашей вине") {
				fmt.Println(previousCoord)
				v, err := f.GetCellValue(mainSheet, previousCoord)
				if err != nil {
					slog.Error("Error getting cell value:", "error", err.Error())
				}
				fmt.Println(v)
				s, err := f.GetCellStyle(mainSheet, previousCoord)
				if err != nil {
					slog.Error("Error getting cell style:", "error", err.Error())
				}
				if err := f.SetCellValue(mainSheet, coord, v); err != nil {
					slog.Error("Error setting cell value:", "error", err.Error())
				}
				if err := f.SetCellStyle(mainSheet, coord, coord, s); err != nil {
					slog.Error("Error setting cell style:", "error", err.Error())
				}

			} else {
				if err := f.SetCellValue(mainSheet, coord, "Не в сети"); err != nil {
					slog.Error("Error setting cell value:", "error", err.Error())
				}
				if err := f.SetCellStyle(mainSheet, coord, coord, redStyle); err != nil {
					slog.Error("Error setting cell style:", "error", err.Error())
				}
			}

		}
		if status == 3 {
			if err := f.SetCellValue(mainSheet, coord, "Недостаточно данных"); err != nil {
				slog.Error("Error setting cell value:", "error", err.Error())
			}
			if err := f.SetCellStyle(mainSheet, coord, coord, yellowStyle); err != nil {
				slog.Error("Error setting cell style:", "error", err.Error())
			}
		}
	}
	if err := f.Save(); err != nil {
		slog.Error("Error saving file:", "error", err.Error())
	}
}
