package postgres

import (
	"context"

	"github.com/jackc/pgx/v5/pgxpool"
)

type Repository struct {
	CameraRepository     *CameraRepository
	GroupRepository      *GroupRepository
	CameraInfoRepository *CameraInfoRepository
}

func NewRepository(ctx context.Context, connString string) *Repository {
	pool, err := pgxpool.New(ctx, connString)
	if err != nil {
		panic("can't connect to database")
	}
	if err := pool.Ping(ctx); err != nil {
		panic("can't ping database")
	}
	return &Repository{
		CameraRepository:     NewCameraRepository(pool),
		GroupRepository:      NewGroupRepository(pool),
		CameraInfoRepository: NewCameraInfoRepository(pool),
	}
}
