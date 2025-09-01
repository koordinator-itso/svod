Acccconst (
	pingCount        = 5
	concurrencyLimit = 300
)onst (
	pingCount        = 5
	concurrencyLimit = 300
)onst (
	pingCount        = 5
	concurrencyLimit = 300
)onst (
	pingCount        = 5
	concurrencyLimit = 300
)const (
	pingCount        = 5
	concurrencyLimit = 300
)PP_NAME = pings
MAIN = ./cmd/main/main.go

# Флаги компиляции
BUILD_FLAGS =
RUN_FLAGS = -config ./config.yaml -mode dev

.PHONY: build run clean

build:
	go build $(BUILD_FLAGS) -o $(APP_NAME) $(MAIN)

run: build
	./$(APP_NAME) $(RUN_FLAGS)

clean:
	rm -f $(APP_NAME)
