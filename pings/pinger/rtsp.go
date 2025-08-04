package pinger

import (
	"sync"
	"time"

	"github.com/bluenviron/gortsplib/v4"
	"github.com/bluenviron/gortsplib/v4/pkg/base"
	"github.com/bluenviron/gortsplib/v4/pkg/format"
)

type RTSPManager struct {
	streams  []string
	results  map[string]bool
	comments map[string]string
	mu       sync.Mutex
}

func NewRTSPManager(streams []string) *RTSPManager {
	return &RTSPManager{
		streams:  streams,
		results:  make(map[string]bool),
		comments: make(map[string]string),
	}
}

func (m *RTSPManager) Start() (map[string]bool, map[string]string) {
	var wg sync.WaitGroup
	sem := make(chan struct{}, concurrencyLimit)

	for _, url := range m.streams {
		if url == "" {
			continue
		}

		wg.Add(1)
		sem <- struct{}{}

		go func(url string) {
			defer wg.Done()
			defer func() { <-sem }()

			parsedURL, err := base.ParseURL(url)
			if err != nil {
				m.setResult(url, false, "Некорректный URL: "+err.Error())
				return
			}

			client := gortsplib.Client{
				ReadTimeout:  5 * time.Second,
				WriteTimeout: 5 * time.Second,
			}
			defer client.Close()

			err = client.Start(parsedURL.Scheme, parsedURL.Host)
			if err != nil {
				m.setResult(url, false, "Ошибка подключения: "+err.Error())
				return
			}

			session, _, err := client.Describe(parsedURL)
			if err != nil {
				m.setResult(url, false, "DESCRIBE ошибка: "+err.Error())
				return
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
				m.setResult(url, true, "")
			} else {
				m.setResult(url, true, "❌ H264 не найден")
			}
		}(url)
	}

	wg.Wait()
	return m.results, m.comments
}

func (m *RTSPManager) setResult(url string, ok bool, comment string) {
	m.mu.Lock()
	defer m.mu.Unlock()
	m.results[url] = ok
	m.comments[url] = comment
}
