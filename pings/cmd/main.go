package main

import (
	"fmt"
	"log/slog"
	"os"
	"strings"
	"test/internal/excel"
	"test/internal/models"
	"test/internal/pinger"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	mainSheet = "Основной файл"
	infoSheet = "Информация"
)

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
	cameras := make([]models.Camera, 0)
	dubles := make(map[string]int)
	rtsps := make([]string, 0)
	rtspCamera := make(map[string][]models.Camera)
	ipFileName := excel.GenerateExcelFilename(int(time.Now().Month()))
	if !excel.FileExists(ipFileName) {
		err := excel.CreateMainExcelFile(ipFileName)
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

		cameras = append(cameras, models.Camera{Name: row[1], Ip: strings.TrimSpace(row[3]), Location: row[2], Rtsp: row[0]})
		if dubles[row[1]] > 0 {
			fmt.Println(row[1])
		} else {
			dubles[row[1]]++
		}
		rtspCamera[row[0]] = append(rtspCamera[row[0]], models.Camera{Name: row[1], Ip: strings.TrimSpace(row[3]), Location: row[2], Rtsp: row[0]})
	}
	for rtsp, _ := range rtspCamera {
		rtsps = append(rtsps, rtsp)
	}

	pinger := pinger.NewRTSPManager(rtsps)
	results, comments := pinger.Start()
	for ip, result := range results {
		slog.Debug(fmt.Sprintf("%s: %t\n", ip, result))
	}
	for ip, comment := range comments {
		slog.Debug(fmt.Sprintf("%s: %s\n", ip, comment))
	}
	if err := excel.InfoToExcel(results, comments, rtspCamera); err != nil {
		return fmt.Errorf("can't write info to excel with error %s", err.Error())
	}
	return nil
}
