package models

type Group struct {
	ID       int
	Comments string
	Cameras  []Camera
}
