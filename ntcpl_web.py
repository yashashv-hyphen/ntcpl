#!/usr/bin/env python3
import base64
import json
import logging
import os
import threading
import uuid
import webbrowser
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory
from huggingface_hub import InferenceClient
from PIL import Image

# =========================================================
# CONFIG & PATHS
# =========================================================

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
log = logging.getLogger(__name__)

APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / "uploaded_product_images"
OUTPUT_DIR = APP_DIR / "generated_images"
INSTRUCTIONS_FILE = APP_DIR / "image_generation_instructions.txt"

IMAGE_MODEL = os.getenv("HF_IMAGE_MODEL", "black-forest-labs/FLUX.1-schnell")
ANALYSIS_MODEL = os.getenv("HF_ANALYSIS_MODEL", "Qwen/Qwen2.5-7B-Instruct")
VISION_MODEL = os.getenv("HF_VISION_MODEL", "Qwen/Qwen2.5-VL-7B-Instruct:fastest")

app = Flask(__name__)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_api_key():
    return os.getenv("HF_TOKEN") or os.getenv("HF_API_KEY") or os.getenv("HUGGINGFACE_API_KEY")

def get_client():
    token = get_api_key()
    if not token:
        raise ValueError("Missing HF_TOKEN/HF_API_KEY")
    return InferenceClient(token=token)

# =========================================================
# CORE LOGIC (Adapted from ntcpl.py)
# =========================================================

def load_instruction_sections():
    if not INSTRUCTIONS_FILE.is_file():
        raise FileNotFoundError(f"Instructions file not found: {INSTRUCTIONS_FILE}")
    text = INSTRUCTIONS_FILE.read_text(encoding="utf-8")
    sections = {}
    current = None
    lines = []
    import re
    for line in text.splitlines():
        header = re.match(r"^\[([A-Z0-9_]+)\]\s*$", line.strip())
        if header:
            if current:
                sections[current] = "\n".join(lines).strip()
            current = header.group(1)
            lines = []
        elif not line.startswith("#"):
            lines.append(line)
    if current:
        sections[current] = "\n".join(lines).strip()
    return sections

def fill_template(template, **values):
    result = template
    for key, value in values.items():
        result = result.replace("{" + key + "}", value or "")
    return " ".join(result.split())

def analyze_product_with_vision(client, image_path, analysis_prompt):
    with open(image_path, "rb") as f:
        data = base64.b64encode(f.read()).decode("utf-8")
    suffix = Path(image_path).suffix.lower()
    mime = "image/jpeg" if suffix in {".jpg", ".jpeg"} else "image/png"
    image_url = f"data:{mime};base64,{data}"
    
    response = client.chat.completions.create(
        model=VISION_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": image_url}},
                    {"type": "text", "text": analysis_prompt},
                ],
            }
        ],
        max_tokens=700,
    )
    return response.choices[0].message.content.strip()

def analyze_product_with_text(client, analysis_prompt):
    response = client.chat.completions.create(
        model=ANALYSIS_MODEL,
        messages=[{"role": "user", "content": analysis_prompt}],
        max_tokens=700,
    )
    return response.choices[0].message.content.strip()

def generate_image(prompt):
    client = get_client()
    image = client.text_to_image(prompt=prompt, model=IMAGE_MODEL)
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()

# =========================================================
# WEB ROUTES
# =========================================================

@app.route("/")
def index():
    hf_available = "true" if get_api_key() else "false"
    
    html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Amazon AI Image Engine</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Inter:wght@400;500&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #ff9900;
            --primary-hover: #ffb84d;
            --bg: #060912;
            --surface: rgba(18, 24, 43, 0.72);
            --surface-hover: rgba(30, 38, 64, 0.7);
            --text: #eaf0ff;
            --text-muted: #9aa6c7;
            --accent: #22d3ee;
            --success: #10b981;
            --error: #ef4444;
            --border: rgba(140, 160, 210, 0.16);
            --gradient: linear-gradient(135deg, #ff9900 0%, #ffb84d 100%);
            --gradient-cool: linear-gradient(135deg, #7c5cff 0%, #22d3ee 100%);
            --glass: rgba(14, 20, 38, 0.7);
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg);
            color: var(--text);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            overflow-x: hidden;
            -webkit-font-smoothing: antialiased;
            background-image:
                radial-gradient(1100px 560px at 8% -8%, rgba(124, 92, 255, 0.20) 0%, transparent 55%),
                radial-gradient(1000px 520px at 95% 10%, rgba(34, 211, 238, 0.16) 0%, transparent 55%),
                radial-gradient(900px 600px at 50% 115%, rgba(255, 153, 0, 0.14) 0%, transparent 55%),
                linear-gradient(180deg, #070b18 0%, #05070f 100%);
            background-attachment: fixed;
        }
        body::before {
            content:""; position: fixed; inset: 0; z-index: 0; pointer-events: none;
            background-image:
                linear-gradient(rgba(140,160,210,0.05) 1px, transparent 1px),
                linear-gradient(90deg, rgba(140,160,210,0.05) 1px, transparent 1px);
            background-size: 46px 46px;
            -webkit-mask-image: radial-gradient(circle at 50% 16%, #000 0%, transparent 78%);
            mask-image: radial-gradient(circle at 50% 16%, #000 0%, transparent 78%);
        }

        header {
            width: 100%;
            padding: 1.6rem 2rem;
            text-align: center;
            background: var(--glass);
            backdrop-filter: blur(16px) saturate(140%);
            border-bottom: 1px solid var(--border);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        header::after {
            content:""; position:absolute; left:0; right:0; bottom:0; height:1px;
            background: linear-gradient(90deg, transparent, rgba(255,153,0,0.6), rgba(124,92,255,0.6), transparent);
        }
        .brandrow {
            display:flex; align-items:center; justify-content:center; gap:14px;
        }
        .brand-mark {
            position: relative; width:48px; height:48px; border-radius:15px;
            background: conic-gradient(from 130deg, #ff9900, #ffb84d, #7c5cff, #22d3ee, #ff9900);
            box-shadow: 0 0 0 1px rgba(255,255,255,0.12), 0 12px 30px rgba(255,153,0,0.35);
            display:grid; place-items:center; flex:none;
        }
        .brand-mark::after { content:""; position:absolute; inset:3px; border-radius:12px; background:#0a0e1b; }
        .brand-spark {
            position: relative; z-index:1; width:21px; height:21px; border-radius:50%;
            background: radial-gradient(circle at 35% 30%, #fff, #ff9900 55%, #b45309);
            animation: sparkPulse 2.8s ease-in-out infinite;
        }
        @keyframes sparkPulse {
            0%,100% { box-shadow: 0 0 14px rgba(255,153,0,0.6); transform: scale(1); }
            50% { box-shadow: 0 0 26px rgba(255,153,0,0.95); transform: scale(1.08); }
        }

        h1 {
            font-family: 'Outfit', sans-serif;
            font-size: 2.3rem;
            font-weight: 700;
            letter-spacing: -0.5px;
            background: linear-gradient(90deg, #ffffff 0%, #ffb84d 50%, #7c5cff 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.25rem;
        }
        .subtitle {
            font-size: 0.9rem; color: var(--text-muted); letter-spacing: 0.3px; margin-top: 0.2rem;
        }

        .container {
            width: 100%;
            max-width: 1200px;
            padding: 2rem;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 2rem;
            position: relative;
            z-index: 1;
        }

        .card {
            background: var(--surface);
            padding: 2rem;
            border-radius: 1.5rem;
            border: 1px solid var(--border);
            box-shadow: 0 24px 60px rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(18px) saturate(140%);
            -webkit-backdrop-filter: blur(18px) saturate(140%);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        .card::after {
            content:""; position:absolute; left:0; top:0; right:0; height:1px;
            background: linear-gradient(90deg, transparent, rgba(255,153,0,0.55), rgba(124,92,255,0.55), transparent);
        }

        .card:hover {
            transform: translateY(-2px);
            box-shadow: 0 30px 70px rgba(0, 0, 0, 0.6);
        }

        .input-group { margin-bottom: 1.5rem; }
        label {
            display: block;
            font-size: 0.875rem;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 0.5rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        input[type="text"], textarea {
            width: 100%;
            padding: 0.75rem 1rem;
            background: rgba(15, 23, 42, 0.5);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 0.75rem;
            color: var(--text);
            font-family: inherit;
            font-size: 1rem;
            transition: border-color 0.2s, box-shadow 0.2s;
        }

        input:focus, textarea:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2);
        }

        #drop-zone {
            width: 100%;
            height: 200px;
            background: rgba(15, 23, 42, 0.3);
            border: 2px dashed rgba(255, 255, 255, 0.1);
            border-radius: 1rem;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
            overflow: hidden;
        }

        #drop-zone:hover {
            border-color: var(--primary);
            background: rgba(99, 102, 241, 0.05);
        }

        #drop-zone img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
            display: none;
        }

        #drop-zone .icon { font-size: 2rem; color: var(--text-muted); margin-bottom: 0.5rem; }

        .btn {
            width: 100%;
            padding: 1rem;
            background: var(--gradient);
            color: #1a1206;
            border: none;
            border-radius: 1rem;
            font-size: 1.05rem;
            font-weight: 700;
            letter-spacing: 0.2px;
            cursor: pointer;
            transition: all 0.3s;
            box-shadow: 0 12px 28px -3px rgba(255, 153, 0, 0.45);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 20px 36px -5px rgba(255, 153, 0, 0.55);
        }

        .btn:active { transform: translateY(0); }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }

        #results {
            display: grid;
            grid-template-columns: 1fr;
            gap: 2rem;
        }

        .result-card {
            background: var(--surface);
            border-radius: 1.5rem;
            overflow: hidden;
            border: 1px solid rgba(255, 255, 255, 0.05);
        }

        .result-img-container {
            width: 100%;
            aspect-ratio: 1;
            background: #000;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
        }

        .result-img-container img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }

        .result-info { padding: 1.5rem; }
        .result-title { font-weight: 700; font-size: 1.25rem; margin-bottom: 0.5rem; color: var(--accent); }

        #status-msg {
            margin-top: 1rem;
            text-align: center;
            font-size: 0.875rem;
            color: var(--text-muted);
        }

        .loader {
            width: 24px;
            height: 24px;
            border: 3px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 1s linear infinite;
            display: none;
        }

        @keyframes spin { to { transform: rotate(360deg); } }

        @media (max-width: 900px) {
            .container { grid-template-columns: 1fr; }
        }

        #api-badge {
            display: inline-flex;
            align-items: center;
            padding: 0.25rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-top: 0.5rem;
        }
        .badge-ok { background: rgba(16, 185, 129, 0.1); color: var(--success); }
        .badge-err { background: rgba(239, 68, 68, 0.1); color: var(--error); }
    </style>
</head>
<body>
    <header>
        <div class="brandrow">
            <span class="brand-mark" aria-hidden="true"><span class="brand-spark"></span></span>
            <h1>Amazon AI Image Engine</h1>
        </div>
        <div class="subtitle">Seller Studio · Lifestyle &amp; Infographic generation for pitch-ready listings</div>
        <div id="api-badge" class="badge-err">HF API KEY MISSING</div>
    </header>

    <div class="container">
        <div class="card">
            <div class="input-group">
                <label>Product Name</label>
                <input type="text" id="product_name" placeholder="E.g. Professional Wireless Earbuds">
            </div>
            
            <div class="input-group">
                <label>Seller Notes / Details</label>
                <textarea id="seller_notes" rows="4" placeholder="Optional: material, color, specific features..."></textarea>
            </div>

            <div class="input-group">
                <label>Main Product Image (Reference)</label>
                <div id="drop-zone" onclick="document.getElementById('file-input').click()">
                    <div class="icon">📸</div>
                    <p id="drop-text">Click or drag image here</p>
                    <img id="preview-img">
                </div>
                <input type="file" id="file-input" style="display:none" accept="image/*">
            </div>

            <button id="generate-btn" class="btn" disabled>
                <span class="loader" id="gen-loader"></span>
                Generate Lifestyle & Infographic
            </button>
            <p id="status-msg"></p>
        </div>

        <div id="results">
            <div class="result-card">
                <div class="result-img-container">
                    <img id="lifestyle-img" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=">
                </div>
                <div class="result-info">
                    <div class="result-title">Lifestyle Image</div>
                    <p class="text-muted">High-quality lifestyle usage visualization.</p>
                </div>
            </div>
            <div class="result-card">
                <div class="result-img-container">
                    <img id="infographic-img" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=">
                </div>
                <div class="result-info">
                    <div class="result-title">Infographic Image</div>
                    <p class="text-muted">Features & benefits breakdown.</p>
                </div>
            </div>
        </div>
    </div>

    <script>
        const hfAvailable = __HF_AVAILABLE__;
        const apiBadge = document.getElementById('api-badge');
        const generateBtn = document.getElementById('generate-btn');
        const fileInput = document.getElementById('file-input');
        const dropZone = document.getElementById('drop-zone');
        const previewImg = document.getElementById('preview-img');
        const dropText = document.getElementById('drop-text');
        
        let selectedFile = null;

        if (hfAvailable) {
            apiBadge.textContent = 'HF API READY';
            apiBadge.className = 'badge-ok';
        }

        fileInput.addEventListener('change', (e) => {
            if (e.target.files && e.target.files[0]) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            selectedFile = file;
            const reader = new FileReader();
            reader.onload = (e) => {
                previewImg.src = e.target.result;
                previewImg.style.display = 'block';
                dropText.style.display = 'none';
                checkReady();
            };
            reader.readAsDataURL(file);
        }

        function checkReady() {
            const name = document.getElementById('product_name').value.trim();
            generateBtn.disabled = !(selectedFile && name && hfAvailable);
        }

        document.getElementById('product_name').addEventListener('input', checkReady);

        generateBtn.addEventListener('click', async () => {
            const name = document.getElementById('product_name').value.trim();
            const notes = document.getElementById('seller_notes').value.trim();
            const statusMsg = document.getElementById('status-msg');
            const loader = document.getElementById('gen-loader');

            const fd = new FormData();
            fd.append('image', selectedFile);
            fd.append('name', name);
            fd.append('notes', notes);

            generateBtn.disabled = true;
            loader.style.display = 'block';
            statusMsg.textContent = 'Analyzing product & generating images (may take 60s)...';

            try {
                const res = await fetch('/api/generate', { method: 'POST', body: fd });
                const data = await res.json();
                
                if (data.ok) {
                    document.getElementById('lifestyle-img').src = data.lifestyle;
                    document.getElementById('infographic-img').src = data.infographic;
                    statusMsg.textContent = 'Generation complete!';
                } else {
                    statusMsg.textContent = 'Error: ' + data.error;
                }
            } catch (err) {
                statusMsg.textContent = 'Error: Connection failed';
            } finally {
                generateBtn.disabled = false;
                loader.style.display = 'none';
            }
        });
    </script>
</body>
</html>
    """
    return html.replace("__HF_AVAILABLE__", hf_available)

@app.route("/api/generate", methods=["POST"])
def api_generate():
    try:
        if not get_api_key():
            return jsonify({"ok": False, "error": "HF API Key not set"}), 400
        
        file = request.files.get("image")
        name = request.form.get("name")
        notes = request.form.get("notes", "")
        
        if not file or not name:
            return jsonify({"ok": False, "error": "Missing image or name"}), 400
        
        # Save temp image
        ext = Path(file.filename).suffix or ".png"
        file_path = UPLOAD_DIR / f"{uuid.uuid4()}{ext}"
        file.save(str(file_path))
        
        # 1. Product Analysis
        log.info(f"Analyzing product: {name}")
        sections = load_instruction_sections()
        analysis_prompt = fill_template(
            sections["PRODUCT_ANALYSIS"],
            product_name=name,
            seller_notes=notes or "None provided.",
            product_details="",
        )
        
        client = get_client()
        product_details = ""
        analysis_method = "vision"
        try:
            product_details = analyze_product_with_vision(client, file_path, analysis_prompt)
        except Exception as e:
            log.warning(f"Vision failed, falling back to text: {e}")
            product_details = analyze_product_with_text(client, analysis_prompt)
            analysis_method = "text"

        # 2. Build Prompts
        lifestyle_prompt = fill_template(
            sections["LIFESTYLE_IMAGE"],
            product_name=name,
            product_details=product_details,
        )
        infographic_prompt = fill_template(
            sections["INFOGRAPHIC_IMAGE"],
            product_name=name,
            product_details=product_details,
        )

        # 3. Generate Images
        log.info("Generating lifestyle image...")
        lifestyle_bytes = generate_image(lifestyle_prompt)
        log.info("Generating infographic image...")
        infographic_bytes = generate_image(infographic_prompt)

        # Save to output bundle
        timestamp = uuid.uuid4().hex[:8]
        bundle_dir = OUTPUT_DIR / f"web_{timestamp}"
        bundle_dir.mkdir(parents=True, exist_ok=True)
        
        (bundle_dir / "lifestyle.png").write_bytes(lifestyle_bytes)
        (bundle_dir / "infographic.png").write_bytes(infographic_bytes)
        (bundle_dir / "original.png").write_bytes(file_path.read_bytes())
        
        ls_b64 = base64.b64encode(lifestyle_bytes).decode("utf-8")
        ig_b64 = base64.b64encode(infographic_bytes).decode("utf-8")
        
        return jsonify({
            "ok": True,
            "lifestyle": f"data:image/png;base64,{ls_b64}",
            "infographic": f"data:image/png;base64,{ig_b64}",
            "bundle": str(bundle_dir.name)
        })

    except Exception as e:
        log.exception("Generation error")
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__ == "__main__":
    port = 7861
    log.info(f"Starting server on port {port}...")
    threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port)
