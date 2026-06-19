#!/usr/bin/env python3
"""Image generation microservice.

Runs on IMAGE_SERVICE_PORT (default 7861).  The excel-service delegates all
AI image generation work here via HTTP so that heavy ML deps (Pillow rendering,
FAL image generation) are isolated in their own container.  Both containers
mount the same Docker volume at /app/generated_images so file paths returned
by this service are directly readable by the excel-service.

Endpoints
---------
POST /api/generate-images
    sku           str   — product SKU (used as filename prefix and sub-directory)
    image_url     str   — reference image URL or local path
    product_name  str   — human-readable name used in AI prompts
    provider      str   — fal | demo
    analysis      dict  — (optional) pre-computed product analysis dict
    bullets       list  — (optional) bullet point strings for PIL infographic
    listing       dict  — (optional) full Groq listing output for infographic details

    Response: {"ok": true, "image_paths": ["/app/generated_images/SKU/..."]}
              {"ok": false, "error": "..."}

GET /healthz
    Returns {"ok": true, "service": "image-service"}
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from flask import Flask, jsonify, request

# All image generation logic lives in the main application module.
# Importing is safe: the Flask dev-server only starts inside __main__.
from amazon_template_autofill_web import (
    APP_DIR,
    log,
    _PRODUCT_IMAGES_AVAILABLE,
    generate_reference_locked_images,
    _download_image_to_dir,
    _reference_conditioning_text,
)

AI_PROVIDER = "groq"

try:
    from amazon_product_images import generateProductImages as _generateProductImages  # type: ignore
except ImportError:
    _generateProductImages = None

app = Flask("image_service")
_PORT = int(os.environ.get("IMAGE_SERVICE_PORT", "7861"))
_DATA_DIR = Path(os.environ.get("AI_GENERATED_IMAGE_DIR", str(APP_DIR / "generated_images")))


@app.post("/api/generate-images")
def api_generate_images():
    payload      = request.get_json(force=True) or {}
    sku          = str(payload.get("sku", "product")).strip() or "product"
    image_url    = str(payload.get("image_url", "")).strip()
    product_name = str(payload.get("product_name", sku)).strip() or sku
    provider     = str(payload.get("provider", AI_PROVIDER)).strip().lower()
    analysis     = payload.get("analysis") or {}
    listing      = payload.get("listing")  or {}
    bullets: list[str] = [
        str(b).strip() for b in (payload.get("bullets") or []) if str(b).strip()
    ]

    if not image_url:
        return jsonify({"ok": False, "error": "image_url is required"}), 400

    out_dir = _DATA_DIR / sku
    out_dir.mkdir(parents=True, exist_ok=True)
    image_paths: list[str] = []

    try:
        ref_path = str(_download_image_to_dir(image_url, out_dir, f"{sku}_ref.jpg"))

        # ── 4-shot pipeline (primary) ─────────────────────────────────────
        _used_4shot = False
        if _PRODUCT_IMAGES_AVAILABLE and _generateProductImages:
            try:
                ref_desc = ""
                try:
                    ref_desc = _reference_conditioning_text(ref_path, product_name)
                except Exception as e:
                    log.debug("Reference conditioning skipped for %s: %s", sku, e)

                _product = dict(analysis)
                _product["sku"] = sku
                shot_results = _generateProductImages(
                    product=_product,
                    reference_image=ref_path,
                    config={
                        "shots": ("main", "lifestyle", "macro", "styled_hero"),
                        "out_dir": out_dir,
                        "provider": provider,
                        "prefix": sku,
                    },
                    reference_description=ref_desc,
                )
                for shot_key in ("lifestyle", "macro", "styled_hero"):
                    if shot_key in shot_results:
                        image_paths.append(str(shot_results[shot_key].resolve()))
                _used_4shot = True
                log.info("4-shot pipeline produced %d images for %s", len(shot_results), sku)
            except Exception as exc:
                log.warning("4-shot pipeline failed for %s: %s — falling back to 2-shot", sku, exc)

        # ── 2-shot fallback (PIL) ─────────────────────────────────────────
        if not _used_4shot:
            try:
                image_paths = generate_reference_locked_images(
                    reference_image_source=ref_path,
                    out_dir=out_dir,
                    prefix=sku,
                    product_name=product_name,
                    bullets=bullets[:5],
                    analysis=analysis,
                    listing=listing,
                )
            except Exception as exc:
                log.warning("PIL generation failed for %s: %s", sku, exc)

    except Exception as exc:
        log.exception("Image generation failed for sku=%s", sku)
        return jsonify({"ok": False, "error": str(exc)}), 500

    return jsonify({"ok": True, "image_paths": image_paths})


@app.get("/healthz")
def healthz():
    return jsonify({"ok": True, "service": "image-service"})


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
    )
    app.run(host="0.0.0.0", port=_PORT, debug=False)
