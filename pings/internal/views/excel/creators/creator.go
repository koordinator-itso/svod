package excel

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
)

type ExcelCreator struct {
	templatePath string
}

func NewExcelCreator(templatePath string) *ExcelCreator {
	return &ExcelCreator{
		templatePath: templatePath,
	}
}
func (c *ExcelCreator) ChangeTemplate(templatePath string) error {
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		return fmt.Errorf("template file not found: %s", templatePath)
	}
	c.templatePath = templatePath
	return nil
}

func (c *ExcelCreator) CreateFromTemplate(destPath string) error {
	if err := os.MkdirAll(filepath.Dir(destPath), os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directories: %w", err)
	}
	src, err := os.Open(c.templatePath)
	if err != nil {
		return fmt.Errorf("failed to open template: %w", err)
	}
	defer src.Close()

	// Создаём целевой файл
	dst, err := os.Create(destPath)
	if err != nil {
		return fmt.Errorf("failed to create destination file: %w", err)
	}
	defer dst.Close()

	// Копируем содержимое
	if _, err := io.Copy(dst, src); err != nil {
		return fmt.Errorf("failed to copy template: %w", err)
	}

	return nil
}
