@echo off
echo === NTCPL Setup ===

docker --version >nul 2>&1
if errorlevel 1 (
  echo Error: Docker is not installed. Get it from https://docs.docker.com/get-docker/
  pause
  exit /b 1
)

if not exist .env (
  echo Downloading .env (API keys^)...
  curl -fsSL -o .env https://github.com/yashashv-hyphen/ntcpl/releases/download/v1.0/default.env
  echo .env downloaded.
) else (
  echo .env already exists, skipping download.
)

echo Building and starting services...
docker compose up --build -d
if errorlevel 1 (
  docker-compose up --build -d
)

echo.
echo Done! App is running at http://localhost:7860
pause
