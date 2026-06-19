# Implementation Plan: Amazon 4-Shot Product Image Generator

## Overview

Add `amazon_product_images.py` — a self-contained module that builds
Amazon-compliant 4-shot image sets (main, lifestyle, macro detail, styled hero)
for any product. It reuses all existing API wrappers
(`_cloudflare_post`, `huggingface_generate_image_bytes`,
`generate_reference_locked_images`) and a new Flask endpoint
`POST /api/generate-product-images` exposes it to the UI and catalog pipeline.

---

## Discovery Notes (read-only findings)

### Image-gen API wrappers available

| Wrapper | File | Supports i2i? | Notes |
|---------|------|--------------|-------|
| `_cloudflare_post(model, body)` | `_web.py` L2268 | No | Returns raw PNG bytes; `CF_IMAGE_MODEL` = `@cf/black-forest-labs/flux-1-schnell` |
| `huggingface_generate_image_bytes(prompt, ref)` | `_web.py` L477 | No | `client.text_to_image`; FLUX.1-schnell; 4 retries |
| `generate_reference_locked_images(...)` | `_web.py` L1506 | N/A | PIL compositor — always available, no API |
| `_reference_conditioning_text(src, name)` | `_web.py` L521 | — | Describes reference image in text for prompt conditioning |

**Neither CF nor HF support true image-to-image.** Reference product fidelity
is achieved via `_reference_conditioning_text()`, which vision-describes the
reference photo and appends the description to every prompt. This is the
existing production approach and is sufficient.

### Product data schema

Every AI provider returns (and every prompt builder consumes) this dict:
```python
{
  "product_type": str,    # "Anarkali kurti with dupatta"
  "category":    str,     # "Women's Ethnic Wear"
  "material":    str,     # "Cotton"
  "colors":      [str],   # ["white", "blue"]
  "features":    [str],   # ["V-neckline", "long sleeves"]
  "usage":       str,     # "Casual wear, festive wear"
  "style":       str,     # "Indo-western"
  "confidence":  float,
}
```

### Active providers (from `.env`)
- **Cloudflare** — credentials present, verified working. **Primary provider.**
- **HuggingFace** — placeholder key (`your_huggingface_api_key_here`), `HF_AVAILABLE=False` after today's fix.
- **PIL fallback** — always available; used when both APIs fail.

### Existing output conventions
- `generated_images/<sku>/` for per-SKU images
- Files: `<prefix>_ai_<N>.png`
- Output dir overridable via `AI_GENERATED_IMAGE_DIR` env var

---

## Architecture Decisions

1. **New file `amazon_product_images.py`** rather than growing `_web.py` past
   6000 lines. The new module is imported into `_web.py` for the Flask endpoint.

2. **No new external dependencies.** All generation calls go through the
   existing wrappers (`_cloudflare_post`, `huggingface_generate_image_bytes`,
   `generate_reference_locked_images`). The new module imports these.

3. **Config via env vars** following the existing `os.getenv("VAR", "default")`
   pattern. Two new vars: `PRODUCT_IMAGE_SIZE` (default `2048`) and
   `PRODUCT_IMAGE_SHOTS` (default `main,lifestyle,macro,styled_hero`).

4. **Provider priority**: CF → HF → PIL fallback, per shot.
   Each shot is independent; one failing doesn't abort the rest.

5. **Reference image for main shot**: Pass `reference_image_path` (local file)
   through `_reference_conditioning_text()` exactly as the existing pipeline
   does. No new API feature needed.

6. **Return format**: `Dict[str, str]` — `{shot_type: absolute_file_path}`.
   The Flask endpoint wraps this as `{ok: true, images: {...}}`.

---

## Dependency Graph

```
product data dict  +  reference_image_path
        │
        ▼
[Task 1] _build_prompt(product, shot_type) → str
        │   uses: product_type, category, material, colors,
        │         features, usage, style  (all from analysis dict)
        │
        ▼
[Task 1] _adapt_by_category(category, shot_type, base_prompt) → str
        │   appends category-specific clauses per spec rules
        │
        ▼
[Task 2] _generate_one_shot(prompt, provider, out_path, ref_src) → Path
        │   → _cloudflare_post()           [existing]
        │   → huggingface_generate_image_bytes()  [existing]
        │   → generate_reference_locked_images()  [existing, PIL]
        │   → _reference_conditioning_text()       [existing]
        │
        ▼
[Task 3] generateProductImages(product, ref_image, config) → Dict[shot, Path]
        │   calls Tasks 1+2 × len(config.shots)
        │   config: shots, size, out_dir, provider
        │
        ▼
[Task 4] POST /api/generate-product-images  (Flask, in _web.py)
        │   accepts {product, reference_image_path, options}
        │   returns {ok, images: {shot: path}}
        │
        ▼
[Task 5] Integration: process_catalog_row() updated to call
         generateProductImages() when generate_images=True
         (replaces current 2-shot pipeline with 4-shot)
```

---

## Task List

### Phase 1: Prompt Library + Adaptive Logic

---

#### Task 1: Prompt templates and category adaptation

**Description:** Create `amazon_product_images.py` with the four prompt
template functions (`_build_main_prompt`, `_build_lifestyle_prompt`,
`_build_macro_prompt`, `_build_styled_hero_prompt`) and the category-adapter
`_adapt_by_category(category, shot_type, prompt)` that applies the
spec's adaptive rules (apparel/flat-lay, glossy reflections, luxury, etc.).
Also define `SHOT_TYPES = ("main", "lifestyle", "macro", "styled_hero")` and
`_build_prompt(product, shot_type)` as the single entry point.

All functions are pure (no I/O) at this stage.

**Acceptance criteria:**
- [ ] `_build_prompt(product, "main")` returns a string containing "white
  background RGB 255,255,255" and "85%" for a generic product dict.
- [ ] `_build_prompt(product, "lifestyle")` returns a string containing
  the product's `usage` field value.
- [ ] `_build_prompt(product, "macro")` contains the first item of
  `product["features"]`.
- [ ] `_build_prompt(product, "styled_hero")` contains the product's
  primary color.
- [ ] For `category="Apparel"`, the main prompt gains "flat lay" and
  "NOT on a model".
- [ ] For `category="Electronics"`, the lifestyle prompt gains "tidy desk".
- [ ] For a product with `style="glossy"` or `material` containing "glass",
  main prompt gains "minimal controlled reflection".
- [ ] Unknown categories pass through without error.

**Verification:**
```python
python3 -c "
from amazon_product_images import _build_prompt
p = {'product_type':'Kurti','category':'Apparel','material':'Cotton',
     'colors':['white'],'features':['V-neckline'],'usage':'festive wear',
     'style':'ethnic','confidence':0.9}
print(_build_prompt(p,'main'))
print('---')
print(_build_prompt(p,'lifestyle'))
"
```

**Dependencies:** None  
**Files:** `amazon_product_images.py` (new)  
**Estimated scope:** S (1 file, ~150 lines)

---

#### Checkpoint 1: Prompts verified
- [ ] All four prompt outputs checked manually for spec compliance.
- [ ] Adaptive clauses appear correctly for Apparel, Electronics, Glossy categories.

---

### Phase 2: Single-Shot Generator

---

#### Task 2: Per-shot generation with provider routing and retry

**Description:** Add `_generate_one_shot(prompt, provider, out_path,
reference_source)` to `amazon_product_images.py`. Routes to the correct
backend:

- `provider="cloudflare"` → `_cloudflare_post(CF_IMAGE_MODEL, {"prompt": conditioned_prompt})`
- `provider="hf"` → `huggingface_generate_image_bytes(conditioned_prompt, reference_source)`
- `provider="pil"` → `generate_reference_locked_images(...)` (PIL compositor; takes product + bullets, not a prompt string)

For CF and HF: condition the prompt by calling `_reference_conditioning_text(reference_source, product_name)` and appending. Wrap each backend call in a retry loop (3 attempts, exponential back-off starting at 5s). On all retries exhausted, fall back to the PIL compositor. Save bytes to `out_path` and return the path.

**Acceptance criteria:**
- [ ] CF provider saves a PNG file at `out_path`.
- [ ] On CF `HTTPError 429`, retries up to 3 times before falling back.
- [ ] On CF permanent failure, falls back to PIL and returns a valid PNG.
- [ ] HF provider path (when `HF_AVAILABLE`) saves a PNG file.
- [ ] PIL provider path produces a valid PIL-composed PNG directly.
- [ ] `out_path` parent directory is created automatically.

**Verification:**
```bash
python3 -c "
import os; os.chdir('/home/venom/amazon_project')
from dotenv import load_dotenv; load_dotenv()
from amazon_product_images import _generate_one_shot
from pathlib import Path
p = Path('/tmp/test_shot_main.png')
result = _generate_one_shot(
    prompt='Professional studio product photograph of a white cotton kurti on pure white background.',
    provider='cloudflare',
    out_path=p,
    reference_source='',
    product={'product_type':'Kurti','category':'Apparel','material':'Cotton',
             'colors':['white'],'features':['V-neckline'],'usage':'festive wear',
             'style':'ethnic','confidence':0.9}
)
print('Saved to:', result, '— size:', result.stat().st_size, 'bytes')
"
```

**Dependencies:** Task 1  
**Files:** `amazon_product_images.py`  
**Estimated scope:** S–M (adds ~80 lines)

---

#### Checkpoint 2: Single shot works
- [ ] CF generates a real PNG; PIL fallback also produces a valid PNG.
- [ ] Retry logic confirmed by temporarily using a bad CF token.

---

### Phase 3: Multi-Shot Orchestrator

---

#### Task 3: `generateProductImages` — 4-shot orchestrator with config

**Description:** Add the public entry point
`generateProductImages(product, reference_image, config)` to
`amazon_product_images.py`.

`config` is a plain dict with defaults:
```python
{
  "shots":    ("main", "lifestyle", "macro", "styled_hero"),
  "size":     int(os.getenv("PRODUCT_IMAGE_SIZE", "2048")),
  "out_dir":  Path(os.getenv("AI_GENERATED_IMAGE_DIR", "generated_amazon_images/ai")),
  "provider": os.getenv("AI_PROVIDER", "cloudflare"),
  "prefix":   product.get("sku") or product.get("product_type", "product"),
}
```

The function:
1. Resolves `reference_source`: if `reference_image` is a URL, downloads to
   a temp file first (reuse `read_image_bytes_any`). If empty string, skips
   reference conditioning.
2. Iterates `config["shots"]`, calls `_build_prompt(product, shot_type)` then
   `_generate_one_shot(...)` for each.
3. Returns `{"main": Path, "lifestyle": Path, "macro": Path, "styled_hero": Path}`
   — only keys for requested shots.
4. Each shot is independent: a failure on one shot does **not** abort others;
   the failed key maps to the PIL fallback path.

**Acceptance criteria:**
- [ ] With `shots=("main", "lifestyle")`, returns exactly 2 keys.
- [ ] Output files are named `<prefix>_main.png`, `<prefix>_lifestyle.png`, etc.
- [ ] When `reference_image` is an S3 URL, the file is downloaded once and
  reused for all shots (not re-downloaded per shot).
- [ ] When `reference_image=""`, reference conditioning is skipped and prompts
  are used as-is.
- [ ] All 4 files exist on disk after a successful run with `shots=all`.
- [ ] Returns in under 120 s for 4 shots via CF (each CF call ≤ 30 s).

**Verification:**
```bash
python3 -c "
import os; os.chdir('/home/venom/amazon_project')
from dotenv import load_dotenv; load_dotenv()
from amazon_product_images import generateProductImages
from pathlib import Path

product = {
    'sku': 'test_kurti',
    'product_type': 'Anarkali kurti',
    'category': 'Apparel',
    'material': 'Cotton',
    'colors': ['white', 'blue'],
    'features': ['V-neckline', 'flared skirt'],
    'usage': 'festive wear',
    'style': 'ethnic',
    'confidence': 0.9,
}
images = generateProductImages(
    product=product,
    reference_image='https://ntcpl-image-catalog-temp.s3.eu-north-1.amazonaws.com/storage/uploads/kurti/kurti_1.jpg',
    config={'out_dir': Path('/tmp/test_4shot'), 'shots': ('main', 'lifestyle', 'macro', 'styled_hero')},
)
for shot, path in images.items():
    print(f'{shot}: {path} ({path.stat().st_size} bytes)')
"
```

**Dependencies:** Tasks 1, 2  
**Files:** `amazon_product_images.py`  
**Estimated scope:** M (adds ~80 lines)

---

#### Checkpoint 3: Full 4-shot pipeline works
- [ ] All 4 PNG files generated and non-empty.
- [ ] Naming convention correct.
- [ ] Reference conditioning text appears in prompts (log at DEBUG level).
- [ ] No regressions in existing image generation paths.

---

### Phase 4: Flask Endpoint

---

#### Task 4: `POST /api/generate-product-images` endpoint

**Description:** Add the Flask route to `amazon_template_autofill_web.py`.

Request body (JSON):
```json
{
  "product": { ...analysis dict fields... },
  "reference_image_path": "/abs/path/or/https://url",
  "options": {
    "shots": ["main", "lifestyle", "macro", "styled_hero"],
    "provider": "cloudflare"
  }
}
```

Response:
```json
{
  "ok": true,
  "images": {
    "main":        "/abs/path/to/prefix_main.png",
    "lifestyle":   "/abs/path/to/prefix_lifestyle.png",
    "macro":       "/abs/path/to/prefix_macro.png",
    "styled_hero": "/abs/path/to/prefix_styled_hero.png"
  }
}
```

Validate that `product` has at minimum `product_type` or `category` (reject
with 400 if both are empty). All other fields are optional.

**Acceptance criteria:**
- [ ] `POST` with a valid product dict and a real CF provider returns 4 paths.
- [ ] Missing `product` field returns `400` with `{"ok": false, "error": "..."}`.
- [ ] Product dict with only `product_type` (no category/material) still
  succeeds (graceful degradation of prompts).
- [ ] Unknown `provider` value returns `400`.

**Verification:**
```bash
TEMPLATE="/home/venom/amazon_project/uploaded_excel_templates/2aa81519e8c04d2ca6e71e5d0786b9a4_LEASH_ANIMAL_COLLAR_1.xlsm"
curl -s -X POST http://localhost:5050/api/generate-product-images \
  -H 'Content-Type: application/json' \
  -d '{
    "product": {
      "sku":"kurti","product_type":"Anarkali kurti","category":"Apparel",
      "material":"Cotton","colors":["white","blue"],
      "features":["V-neckline","flared skirt"],"usage":"festive wear",
      "style":"ethnic","confidence":0.9
    },
    "reference_image_path": "https://ntcpl-image-catalog-temp.s3.eu-north-1.amazonaws.com/storage/uploads/kurti/kurti_1.jpg",
    "options": {"shots": ["main","lifestyle"], "provider": "cloudflare"}
  }' | python3 -m json.tool
```

**Dependencies:** Task 3  
**Files:** `amazon_template_autofill_web.py`  
**Estimated scope:** S (adds ~40 lines)

---

#### Checkpoint 4: Endpoint verified
- [ ] curl test returns `{"ok": true, "images": {...}}`.
- [ ] Error cases return correct 400 responses.
- [ ] No existing endpoints broken (run existing catalog test).

---

### Phase 5: Catalog Pipeline Integration

---

#### Task 5: Replace 2-shot pipeline in `process_catalog_row` with 4-shot

**Description:** Update `process_catalog_row` in `amazon_template_autofill_web.py`
to call `generateProductImages` when `generate_images=True`, replacing the
current `cloudflare_generate_images` / `generate_reference_locked_images` 2-shot
block. Map the 4 returned paths to template columns:

- `main` → `Main Image URL` column (overwrite only if currently a local path,
  not an S3 URL — preserve original URL when it exists)
- `lifestyle` → `Other Image URL` column 1
- `macro` → `Other Image URL` column 2
- `styled_hero` → logged only (Amazon allows up to 9 images; placeholder for
  future upload)

The existing 2-shot fallback path is kept for when `amazon_product_images`
import fails (defensive import).

**Acceptance criteria:**
- [ ] Catalog pipeline with `generate_images=True` now produces 4 image files
  per SKU under `generated_images/<sku>/`.
- [ ] `lifestyle` and `macro` paths appear in the `Other Image URL` cells of
  the returned product row.
- [ ] If `generateProductImages` raises, the pipeline falls back to the old
  2-shot path and logs a warning (no crash).
- [ ] Existing catalog test (kurti catalog, no images) still passes with
  `generate_images=False`.

**Verification:**
```bash
python3 -c "
import os; os.chdir('/home/venom/amazon_project')
from dotenv import load_dotenv; load_dotenv()
import requests, json
resp = requests.post('http://localhost:5050/api/process-catalog', json={
  'template_path': '/home/venom/amazon_project/uploaded_excel_templates/2aa81519e8c04d2ca6e71e5d0786b9a4_LEASH_ANIMAL_COLLAR_1.xlsm',
  'image_catalog_path': '/home/venom/amazon_project/uploaded_catalogs/db65282232434e53984dbdd2935165df_catalog_60.xlsx',
  'provider': 'cloudflare',
  'generate_images': True,
}, timeout=300)
data = resp.json()
print('ok:', data['ok'], 'count:', data.get('count'))
for detail in data.get('details', [])[:1]:
    print('image_paths:', detail['pipeline'].get('image_paths'))
"
```

**Dependencies:** Tasks 3, 4  
**Files:** `amazon_template_autofill_web.py`  
**Estimated scope:** S (modifies ~30 lines in one function)

---

#### Checkpoint 5: Full integration verified
- [ ] Catalog pipeline produces 4 PNGs per SKU.
- [ ] Other Image URL columns populated in saved Excel.
- [ ] No regressions on `generate_images=False` path.
- [ ] Existing `/api/ai-generate-row` and `/api/save` still work.

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| CF `flux-1-schnell` produces poor white-bg compliance for main shot | High | Add explicit negative-prompt clause; validate bg color after generation; flag in logs |
| CF rate limit (429) during 4-shot generation (4× CF calls per product) | Medium | Existing 3-retry + back-off in `_cloudflare_post`; shots are sequential not parallel |
| `_reference_conditioning_text` vision call adds latency per product | Medium | Cache result: compute once, reuse for all 4 shots |
| `amazon_product_images.py` import fails (new module on old deployment) | Low | Wrap import in try/except in `_web.py`; fall back to existing 2-shot pipeline |
| PIL fallback produces visually weaker main image | Low | PIL fallback is always white-bg compliant — acceptable safety net |

## Open Questions

1. Should the 4-shot generation be run **in a background thread** (non-blocking
   response) with a job-polling endpoint? Currently planned as synchronous —
   OK for ≤5 SKUs but will time out in browser for large catalogs.
   → **Ask before Task 4 if batch size > 5 is expected.**

2. Should `styled_hero` path be written anywhere in the template row, or
   only returned in the API response and saved to disk?
   → **Defaulting to disk-only for now; raise before Task 5 if needed.**

3. CF `flux-1-schnell` prompt length is limited. Should prompts be truncated
   at a safe limit (e.g. 500 chars)?
   → **Will add a `_truncate_prompt(p, max_chars=500)` guard in Task 2.**
