-- ===============================
-- Создание таблицы групп
-- ===============================
CREATE TABLE groups (
    id SERIAL PRIMARY KEY,
    comments TEXT
);

-- ===============================
-- Создание таблицы камер
-- ===============================
CREATE TABLE cameras (
    id SERIAL PRIMARY KEY,
    rtsp TEXT NOT NULL,
    name TEXT NOT NULL,
    ip TEXT NOT NULL,
    location TEXT,
    date_added TIMESTAMPTZ DEFAULT now(),
    coordinates TEXT,
    mac_address TEXT,
    group_id INT REFERENCES groups(id) ON DELETE SET NULL
);

-- ===============================
-- Таблица информации о камерах
-- ===============================
CREATE TABLE camera_info (
    id SERIAL PRIMARY KEY,
    camera_id INT NOT NULL REFERENCES cameras(id) ON DELETE CASCADE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    status BOOLEAN NOT NULL
);

-- ===============================
-- Индексы для ускорения поиска
-- ===============================

-- быстрый поиск по дате в camera_info
CREATE INDEX idx_camera_info_created_at ON camera_info(created_at);

-- быстрый поиск по group_id в cameras
CREATE INDEX idx_cameras_group_id ON cameras(group_id);
