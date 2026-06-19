#!/bin/bash
set -e

echo "=== NTCPL Setup ==="

if ! command -v docker &>/dev/null; then
  echo "Error: Docker is not installed. Install it from https://docs.docker.com/get-docker/"
  exit 1
fi

if ! docker compose version &>/dev/null && ! docker-compose version &>/dev/null; then
  echo "Error: Docker Compose is not installed."
  exit 1
fi

if [ ! -f .env ]; then
  echo "Downloading .env (API keys)..."
  curl -fsSL -o .env https://github.com/yashashv-hyphen/ntcpl/releases/download/v1.0/default.env
  echo ".env downloaded."
else
  echo ".env already exists, skipping download."
fi

echo "Building and starting services..."
docker compose up --build -d 2>/dev/null || docker-compose up --build -d

echo ""
echo "Done! App is running at http://localhost:7860"
