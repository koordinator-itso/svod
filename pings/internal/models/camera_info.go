package models

import "time"

type CameraInfo struct {
	ID        int
	CameraID  int
	CreatedAt time.Time
	Status    bool
	Comment   string
}
