package config

import (
	"log/slog"

	"github.com/spf13/viper"
)

type Config struct {
	DB struct {
		DSN      string `mapstructure:"dsn"`
		MaxConns int    `mapstructure:"max_conns"`
	} `mapstructure:"db"`

	Excel struct {
		TemplatesPath     string `mapstructure:"templates_path"`
		OutputArchivePath string `mapstructure:"output_archive_path"`
		OutputLogsPath    string `mapstructure:"output_logs_path"`
	} `mapstructure:"excel"`
}

func LoadConfig(path string, log *slog.Logger) *Config {
	viper.SetConfigName("config")
	viper.SetConfigType("yaml")
	viper.AddConfigPath(path)
	viper.AutomaticEnv() // позволяет переопределять env-переменными

	// Маппим env → struct
	viper.SetEnvPrefix("APP") // пример: APP_DB_DSN
	viper.BindEnv("db.dsn")
	viper.BindEnv("db.max_conns")
	viper.BindEnv("excel.templates_path")
	viper.BindEnv("excel.output_archive_path")
	viper.BindEnv("excel.output_logs_path")

	if err := viper.ReadInConfig(); err != nil {
		log.Error("Config file not found: %v (fallback to ENV only)", err)
	}

	var cfg Config
	if err := viper.Unmarshal(&cfg); err != nil {
		log.Error("Config unmarshal error: %v", err)
	}
	return &cfg
}
