package pinger

import (
	"context"
	"errors"
	"log/slog"
	"sync"
	"test/internal/db/postgres"
	"time"

	"github.com/bluenviron/gortsplib/v4"
	"github.com/bluenviron/gortsplib/v4/pkg/base"
	"github.com/bluenviron/gortsplib/v4/pkg/format"
	"golang.org/x/sync/errgroup"
)

type RTSPManager struct {
	results  map[string]bool
	comments map[string]string
	mu       sync.Mutex
	repo     *postgres.Repository
}

const (
	concurrencyLimit = 300
)

func NewRTSPManager(repo *postgres.Repository) *RTSPManager {
	return &RTSPManager{
		results:  make(map[string]bool),
		comments: make(map[string]string),
		repo:     repo,
	}
}

func (m *RTSPManager) Start(logger *slog.Logger) error {
	err := m.Ping(logger)
	if err != nil {
		return err
	}
	ticker := time.NewTicker(30 * time.Minute)
	for {
		select {
		case <-ticker.C:
			err := m.Ping(logger)
			if err != nil {
				return err
			}
		}
	}
}

func (m *RTSPManager) Ping(logger *slog.Logger) error {
	var eg errgroup.Group
	sem := make(chan struct{}, concurrencyLimit)

	var streams []string

	cameras, err := m.repo.CameraRepository.GetCameras(context.Background())
	if err != nil {
		logger.Error("GetCameras error", "error", err.Error())
		return err
	}
	for _, camera := range cameras {
		streams = append(streams, camera.Rtsp)
	}

	for _, url := range streams {
		if url == "" {
			continue
		}

		sem <- struct{}{}
		eg.Go(func() error {
			status, comment := checkRTSP(url, logger, sem)
			if err := m.setResult(url, status, comment); err != nil {
				logger.Error("SetResult error", "error", err.Error())
				return err
			}
			return nil
		})

	}
	if err := eg.Wait(); err != nil {
		return err
	}
	return nil
}

func checkRTSP(url string, logger *slog.Logger, sem chan struct{}) (bool, string) {
	defer func() { <-sem }()

	parsedURL, err := base.ParseURL(url)
	if err != nil {
		logger.Error("ParseURL error", "error", err.Error())
		return false, "Некорректный URL: " + err.Error()
	}

	client := gortsplib.Client{
		ReadTimeout:  5 * time.Second,
		WriteTimeout: 5 * time.Second,
	}
	defer client.Close()

	err = client.Start(parsedURL.Scheme, parsedURL.Host)
	if err != nil {
		logger.Error("Start error", "error", err.Error())
		return false, "Ошибка подключения: " + err.Error()
	}

	session, _, err := client.Describe(parsedURL)
	if err != nil {
		logger.Error("Describe error", "error", err.Error())
		return false, "DESCRIBE ошибка: " + err.Error()
	}

	found := false
	for _, media := range session.Medias {
		for _, f := range media.Formats {
			if _, ok := f.(*format.H264); ok {
				found = true
				break
			}
		}
		if found {
			break
		}
	}

	if found {
		return true, ""
	}
	return true, "❌ H264 не найден"
}

func (m *RTSPManager) setResult(url string, ok bool, comment string) error {
	camera, err := m.repo.CameraRepository.GetCameraByRtsp(context.Background(), url)
	if err != nil {

		return err
	}
	if camera == nil {
		return errors.New("camera not found")
	}
	_, err = m.repo.CameraInfoRepository.CreateCameraInfo(string(camera.ID), ok, comment)
	if err != nil {
		return err
	}
	return nil
}
