package postgres

import (
	"context"
	"sync"
	"test/internal/models"
	"time"

	"github.com/jackc/pgx/v5/pgxpool"
)

type CameraInfoRepository struct {
	db *pgxpool.Pool
	mu *sync.RWMutex
}

func NewCameraInfoRepository(pool *pgxpool.Pool) *CameraInfoRepository {
	return &CameraInfoRepository{
		db: pool,
		mu: &sync.RWMutex{},
	}
}

const (
	CreateCameraInfo   = "INSERT INTO camera_info (camera_id,status,comment) VALUES ($1,$2,$3) RETURNING id"
	GetCameraInfo      = "SELECT id,camera_id,status,comment FROM camera_info WHERE camera_id=$1"
	GetAllCameraInfo   = "SELECT id,camera_id,status,comment FROM camera_info"
	GetCameraInfoByDay = "SELECT id,camera_id,status,comment FROM camera_info WHERE created_at>=$1 AND created_at<$2"
)

func (r *CameraInfoRepository) CreateCameraInfo(cameraID string, status bool, comment string) (int, error) {
	r.mu.Lock()
	defer r.mu.Unlock()

	var id int
	err := r.db.QueryRow(context.Background(), CreateCameraInfo, cameraID, status, comment).Scan(&id)
	if err != nil {
		return 0, err
	}

	return id, nil
}

func (r *CameraInfoRepository) GetCameraInfo(cameraID string) (*models.CameraInfo, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	var info models.CameraInfo
	err := r.db.QueryRow(context.Background(), GetCameraInfo, cameraID).Scan(&info.ID, &info.CameraID, &info.Status, &info.Comment)
	if err != nil {
		return nil, err
	}

	return &info, nil
}

func (r *CameraInfoRepository) GetAllCameraInfo() ([]*models.CameraInfo, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	rows, err := r.db.Query(context.Background(), GetAllCameraInfo)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var infos []*models.CameraInfo
	for rows.Next() {
		var info models.CameraInfo
		if err := rows.Scan(&info.ID, &info.CameraID, &info.Status, &info.Comment); err != nil {
			return nil, err
		}
		infos = append(infos, &info)
	}

	return infos, nil
}

func (r *CameraInfoRepository) GetCameraInfoByDay(start, end time.Time) ([]*models.CameraInfo, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	rows, err := r.db.Query(context.Background(), GetCameraInfoByDay, start, end)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var infos []*models.CameraInfo
	for rows.Next() {
		var info models.CameraInfo
		if err := rows.Scan(&info.ID, &info.CameraID, &info.Status, &info.Comment); err != nil {
			return nil, err
		}
		infos = append(infos, &info)
	}

	return infos, nil
}
