# NTCPL - Amazon Product Listing Automation

Automates Amazon product listing creation using AI (Groq, Gemini, OpenAI, FAL, HuggingFace).

## Requirements

- [Docker](https://docs.docker.com/get-docker/) (includes Docker Compose)

That's it. No Python install needed.

## Setup & Run

### Linux / Mac
```bash
git clone https://github.com/yashashv-hyphen/ntcpl.git
cd ntcpl
chmod +x setup.sh && ./setup.sh
```

### Windows
```bat
git clone https://github.com/yashashv-hyphen/ntcpl.git
cd ntcpl
setup.bat
```

The setup script will:
1. Download the `.env` file with all API keys from the [v1.0 release](https://github.com/yashashv-hyphen/ntcpl/releases/tag/v1.0)
2. Build the Docker images
3. Start both services

Then open **http://localhost:7860** in your browser.

## Services

| Service | Port | Description |
|---------|------|-------------|
| Web UI | 7860 | Upload catalog, fill Amazon templates |
| Image service | 7861 | AI image generation (internal) |

## Useful commands

```bash
# View logs
docker compose logs -f

# Stop
docker compose down

# Restart after code changes
docker compose up --build -d
```

## Environment Variables

All API keys come pre-configured in the `.env` downloaded by the setup script.  
See `.env.example` for the full list of variables and what they control.

Key providers:
- **Groq** – text generation (primary)
- **Google Gemini** – vision analysis  
- **FAL** – image generation
- **HuggingFace** – fallback models
- **Cloudflare** – Workers AI
- **OpenAI** – DALL-E image generation
