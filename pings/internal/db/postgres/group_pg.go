package postgres

import (
	"context"
	"sync"
	"test/internal/models"

	"github.com/jackc/pgx/v5/pgxpool"
)

type GroupRepository struct {
	db *pgxpool.Pool
	mu *sync.RWMutex
}

const (
	CreateGroupQuery  = `INSERT INTO groups (name) VALUES ($1) RETURNING id`
	GetGroupByIDQuery = `SELECT id, comments FROM groups WHERE id = $1`
)

func NewGroupRepository(db *pgxpool.Pool) *GroupRepository {
	return &GroupRepository{
		db: db,
		mu: &sync.RWMutex{},
	}
}

func (r *GroupRepository) CreateGroup(ctx context.Context, name string) (int, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	var id int
	err := r.db.QueryRow(ctx, CreateGroupQuery, name).Scan(&id)
	if err != nil {
		return 0, err
	}

	return id, nil
}

func (r *GroupRepository) GetGroupByID(ctx context.Context, id int) (*models.Group, error) {
	r.mu.RLock()
	defer r.mu.RUnlock()

	var group models.Group
	err := r.db.QueryRow(ctx, GetGroupByIDQuery, id).Scan(&group.ID, &group.Comments)
	if err != nil {
		return nil, err
	}

	return &group, nil
}
