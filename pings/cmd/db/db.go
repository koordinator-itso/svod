package db

import "github.com/xuri/excelize/v2"

const (
	excelFileName = ".xlsx"
)

func main() {
	f, err := excelize.OpenFile(excelFileName)
	if err != nil {
		panic(err)
	}
	defer f.Close()

}
