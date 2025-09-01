package models

import "time"

type Camera struct {
	ID          int
	Rtsp        string
	Name        string
	Ip          string
	Location    string
	DateAdded   time.Time
	Coordinates string
	MacAddress  string
	GroupID     int
}
