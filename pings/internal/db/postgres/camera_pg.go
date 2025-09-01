package postgres

import (
	"context"
	"sync"
	"test/internal/models"

	"github.com/jackc/pgx/v5/pgxpool"
)

type CameraRepository struct {
	db *pgxpool.Pool
	mu *sync.RWMutex
}

const (
	GetCameraByIDQuery        = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras WHERE id = $1"
	GetCamerasByNameQuery     = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras WHERE name = $1"
	GetCamerasQuery           = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras"
	GetCamerasByLocationQuery = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras WHERE location = $1"
	GetCameraByMacQuery       = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras WHERE mac_address = $1"
	UpdateCameraQuery         = "UPDATE cameras SET rtsp = $2, name = $3, ip = $4, location = $5, date_added = $6, coordinates = $7, group_id = $8 WHERE mac_address = $1"
	DeleteCameraQuery         = "DELETE FROM cameras WHERE id = $1"
	CreateCameraQuery         = "INSERT INTO cameras (rtsp, name, ip, location, date_added, coordinates, mac_address, group_id) VALUES ($1, $2, $3, $4, $5, $6, $7, $8) RETURNING id"
	GetCameraByRtspQuery      = "SELECT id, rtsp, name, ip, location, date_added, coordinates, mac_address, group_id FROM cameras WHERE rtsp = $1"
)

func NewCameraRepository(db *pgxpool.Pool) *CameraRepository {
	return &CameraRepository{
		db: db,
		mu: &sync.RWMutex{},
	}
}

func (r *CameraRepository) CreateCamera(ctx context.Context, camera *models.Camera) error {
	r.mu.Lock()
	defer r.mu.Unlock()
	var id int
	err := r.db.QueryRow(ctx, CreateCameraQuery, camera.Rtsp, camera.Name, camera.Ip, camera.Location, camera.DateAdded, camera.Coordinates, camera.MacAddress, camera.GroupID).Scan(&id)
	if err != nil {
		return err
	}
	camera.ID = id
	return nil
}

func (r *CameraRepository) GetCameras(ctx context.Context) ([]*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	rows, err := r.db.Query(ctx, GetCamerasQuery)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var cameras []*models.Camera
	for rows.Next() {
		var camera models.Camera
		if err := rows.Scan(
			&camera.ID,
			&camera.Rtsp,
			&camera.Name,
			&camera.Ip,
			&camera.Location,
			&camera.DateAdded,
			&camera.Coordinates,
			&camera.MacAddress,
			&camera.GroupID,
		); err != nil {
			return nil, err
		}
		cameras = append(cameras, &camera)
	}

	return cameras, nil
}

func (r *CameraRepository) GetCamerasByLocation(ctx context.Context, location string) ([]*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	rows, err := r.db.Query(ctx, GetCamerasByLocationQuery, location)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var cameras []*models.Camera
	for rows.Next() {
		var camera models.Camera
		if err := rows.Scan(
			&camera.ID,
			&camera.Rtsp,
			&camera.Name,
			&camera.Ip,
			&camera.Location,
			&camera.DateAdded,
			&camera.Coordinates,
			&camera.MacAddress,
			&camera.GroupID); err != nil {
			return nil, err
		}
		cameras = append(cameras, &camera)
	}

	return cameras, nil
}
func (r *CameraRepository) GetCamerasByName(ctx context.Context, name string) ([]*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	rows, err := r.db.Query(ctx, GetCamerasByNameQuery, name)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var cameras []*models.Camera
	for rows.Next() {
		var camera models.Camera
		if err := rows.Scan(
			&camera.ID,
			&camera.Rtsp,
			&camera.Name,
			&camera.Ip,
			&camera.Location,
			&camera.DateAdded,
			&camera.Coordinates,
			&camera.MacAddress,
			&camera.GroupID); err != nil {
			return nil, err
		}
		cameras = append(cameras, &camera)
	}

	return cameras, nil
}

func (r *CameraRepository) GetCameraByMac(ctx context.Context, macAddress string) (*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	row := r.db.QueryRow(ctx, GetCameraByMacQuery, macAddress)

	var camera models.Camera
	if err := row.Scan(
		&camera.ID,
		&camera.Rtsp,
		&camera.Name,
		&camera.Ip,
		&camera.Location,
		&camera.DateAdded,
		&camera.Coordinates,
		&camera.MacAddress,
		&camera.GroupID); err != nil {
		return nil, err
	}

	return &camera, nil
}

func (r *CameraRepository) GetCameraByID(ctx context.Context, id int) (*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	row := r.db.QueryRow(ctx, GetCameraByIDQuery, id)

	var camera models.Camera
	if err := row.Scan(
		&camera.ID,
		&camera.Rtsp,
		&camera.Name,
		&camera.Ip,
		&camera.Location,
		&camera.DateAdded,
		&camera.Coordinates,
		&camera.MacAddress,
		&camera.GroupID); err != nil {
		return nil, err
	}

	return &camera, nil
}

func (r *CameraRepository) GetCameraByRtsp(ctx context.Context, rtsp string) (*models.Camera, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	row := r.db.QueryRow(ctx, GetCameraByRtspQuery, rtsp)

	var camera models.Camera
	if err := row.Scan(
		&camera.ID,
		&camera.Rtsp,
		&camera.Name,
		&camera.Ip,
		&camera.Location,
		&camera.DateAdded,
		&camera.Coordinates,
		&camera.MacAddress,
		&camera.GroupID); err != nil {
		return nil, err
	}

	return &camera, nil
}

func (r *CameraRepository) UpdateCamera(ctx context.Context, camera *models.Camera) error {
	r.mu.RLock()
	defer r.mu.RUnlock()

	_, err := r.db.Exec(ctx, UpdateCameraQuery,
		camera.MacAddress,
		camera.Rtsp,
		camera.Name,
		camera.Ip,
		camera.Location,
		camera.DateAdded,
		camera.Coordinates,
		camera.GroupID)

	return err
}

func (r *CameraRepository) DeleteCamera(ctx context.Context, id int) error {
	r.mu.RLock()
	defer r.mu.RUnlock()

	_, err := r.db.Exec(ctx, DeleteCameraQuery, id)

	return err
}
