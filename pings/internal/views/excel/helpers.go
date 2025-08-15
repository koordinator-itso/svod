package excel

import (
	"fmt"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

func FileExists(path string) bool {
	_, err := os.Stat(path)
	if os.IsNotExist(err) {
		return false
	}
	return err == nil
}

func FindNextEmptyColumnIndex(f *excelize.File, sheet string, row int) (int, error) {
	for col := 1; col <= 100; col++ {
		cell, _ := excelize.CoordinatesToCellName(col, row)
		val, _ := f.GetCellValue(sheet, cell)
		if val == "" {
			return col, nil
		}
	}
	return 0, fmt.Errorf("нет пустых колонок")
}

func GenerateExcelFilename(month int) string {
	now := time.Now()
	monthUpper := getRussianMonth(time.Month(month))
	year := now.Year()
	return fmt.Sprintf("C:/Users/user/Documents/Bitrix24-koordinator@itso.su@itso.bitrix24.ru/Мониторинг и обслуживание/Свод %s %d.xlsx", monthUpper, year)
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
