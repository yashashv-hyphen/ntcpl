# NTCPL — Amazon Product Listing Automation

Automatically fills Amazon listing templates from your product catalog using AI. Upload an Excel catalog, get a filled Amazon template back — with titles, bullet points, descriptions, and AI-generated images.

---

## Prerequisites

Only one thing is required:

- **[Docker Desktop](https://docs.docker.com/get-docker/)** (includes Docker Compose)
  - Windows: [Download Docker Desktop for Windows](https://docs.docker.com/desktop/install/windows-install/)
  - Mac: [Download Docker Desktop for Mac](https://docs.docker.com/desktop/install/mac-install/)
  - Linux: [Install Docker Engine](https://docs.docker.com/engine/install/) + [Docker Compose plugin](https://docs.docker.com/compose/install/)

No Python. No Node. No other installs needed.

---

## Installation & Setup

### Step 1 — Clone the repository

```bash
git clone https://github.com/yashashv-hyphen/ntcpl.git
cd ntcpl
```

### Step 2 — Run the setup script

**Linux / Mac (Terminal):**
```bash
chmod +x setup.sh && ./setup.sh
```

**Windows (Command Prompt or PowerShell — run as Administrator):**
```bat
setup.bat
```

The script will:
1. Check that Docker is installed
2. Download the pre-configured `.env` file with API keys (AI services)
3. Build the Docker images (takes 2–5 minutes on first run)
4. Start both services in the background

### Step 3 — Open the app

Once setup finishes, open your browser and go to:

```
http://localhost:7860
```

You should see the NTCPL web interface.

---

## Using the Tool

### What you need
- Your **product catalog** as an Excel file (`.xlsx`) with columns like SKU, product name, description, image URL, etc.
- An **Amazon listing template** Excel file (the flat-file template downloaded from Amazon Seller Central)

### Workflow

1. **Upload your catalog** — click "Upload Catalog" and select your `.xlsx` file
2. **Upload the Amazon template** — click "Upload Template" and select the Amazon flat-file template
3. **Click "Generate Listings"** — the AI reads your catalog, generates titles, bullets, descriptions, and images
4. **Download the result** — a filled Amazon template ready to upload to Seller Central

### Expected time
- ~30–60 seconds per SKU for text generation
- ~60–90 seconds per SKU for AI image generation (if enabled)

---

## Stopping & Restarting

```bash
# Stop all services
docker compose down

# Start again (no rebuild needed)
docker compose up -d

# Restart after pulling updates
git pull
docker compose up --build -d
```

---

## Troubleshooting

### "Docker is not installed" / setup script fails immediately
Install Docker Desktop from https://docs.docker.com/get-docker/ and make sure it is **running** before re-running the setup script.

### Port 7860 is already in use
Something else on your machine is using port 7860. Either stop that process, or edit `docker-compose.yml` and change `"7860:7860"` to `"7862:7860"` then open `http://localhost:7862`.

### Page loads but "Generate" never completes / spinner runs forever
View logs to see what's happening:
```bash
docker compose logs -f
```
Look for `ERROR` lines. Common causes:
- API rate limit hit — wait a few minutes and retry
- Network issue fetching product images — check that image URLs in your catalog are publicly accessible

### "Cannot connect to Docker daemon"
Docker Desktop is not running. Open Docker Desktop from your Applications/Start Menu and wait for it to start (the whale icon in the taskbar turns solid).

### Build fails with "error" during `docker compose up --build`
Ensure you have a stable internet connection — the build downloads Python packages. Then retry:
```bash
docker compose up --build -d
```

### Windows: `setup.bat` opens and closes immediately
Right-click `setup.bat` → **Run as Administrator**. If it still closes, open Command Prompt as Administrator, `cd` into the project folder, and run `setup.bat` manually to see the error.

### Mac: Permission denied on `setup.sh`
Run `chmod +x setup.sh` first, then `./setup.sh`.

---

## Using Your Own API Keys (Optional)

The setup script downloads a `.env` with shared demo API keys. For production use or higher limits, replace them with your own:

1. Copy the example:
   ```bash
   cp .env.example .env
   ```
2. Edit `.env` and fill in your keys:
   ```
   GROQ_API_KEY=your_groq_key        # free at console.groq.com
   FAL_KEY=your_fal_key              # fal.ai (pay-per-use)
   ```
3. Restart services:
   ```bash
   docker compose up --build -d
   ```

| Provider | What it does | Free tier |
|----------|-------------|-----------|
| [Groq](https://console.groq.com) | Text generation (titles, bullets, descriptions) | Yes — generous free tier |
| [FAL](https://fal.ai) | AI image generation | Pay-per-use (~$0.05/image) |

---

## Services

| Service | Port | Description |
|---------|------|-------------|
| Web UI | 7860 | Main interface — upload catalog, download filled template |
| Image service | 7861 | Internal AI image generation (not exposed to browser) |

---

## Viewing Logs

```bash
# All services
docker compose logs -f

# Web UI only
docker compose logs -f excel-service

# Image service only
docker compose logs -f image-service
```
