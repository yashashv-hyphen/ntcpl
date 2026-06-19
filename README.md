# NTCPL - Amazon Product Listing Automation

Automates Amazon product listing creation using AI (Groq, Gemini, OpenAI, FAL, HuggingFace).

## Quick Setup

### 1. Clone the repo
```bash
git clone https://github.com/yashashv-hyphen/ntcpl.git
cd ntcpl
```

### 2. Download the .env file (contains all API keys)
```bash
curl -L -o .env https://github.com/yashashv-hyphen/ntcpl/releases/download/v1.0/default.env
```

### 3. Install Python dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the web app
```bash
python amazon_template_autofill_web.py
```

Then open `http://localhost:5000` in your browser.

## Docker (alternative)

```bash
curl -L -o .env https://github.com/yashashv-hyphen/ntcpl/releases/download/v1.0/default.env
docker-compose up
```

## Environment Variables

All API keys are in the `.env` file available in the [v1.0 release](https://github.com/yashashv-hyphen/ntcpl/releases/tag/v1.0). Download it as shown above.

Key providers used:
- **Groq** – text generation (primary)
- **Google Gemini** – vision analysis
- **FAL** – image generation
- **HuggingFace** – fallback models
- **Cloudflare** – Workers AI
- **OpenAI** – DALL-E image generation
