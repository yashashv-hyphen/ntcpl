# Task List — Amazon 4-Shot Product Image Generator

## Phase 1: Prompt Library + Adaptive Logic
- [ ] **Task 1** — `_build_prompt()` + `_adapt_by_category()` in `amazon_product_images.py` (new file)
  - 4 shot templates: main, lifestyle, macro, styled_hero
  - Adaptive clauses: apparel flat-lay, glossy reflection guard, luxury moodboard, electronics desk, beauty spa, kitchen action, outdoor daylight
  - Pure functions, no I/O

## Checkpoint 1
- [ ] All 4 prompt outputs pass spec compliance check (white-bg clause, 85% fill, no-text rule)
- [ ] Adaptive clauses verified for Apparel, Electronics, Glossy material

## Phase 2: Single-Shot Generator
- [ ] **Task 2** — `_generate_one_shot(prompt, provider, out_path, reference_source, product)` in `amazon_product_images.py`
  - CF / HF / PIL routing
  - Reference conditioning via `_reference_conditioning_text()` (existing)
  - 3-retry + back-off + PIL fallback
  - Prompt truncation guard at 500 chars for CF

## Checkpoint 2
- [ ] CF generates real PNG for one shot
- [ ] PIL fallback confirmed by forcing CF failure

## Phase 3: Multi-Shot Orchestrator
- [ ] **Task 3** — `generateProductImages(product, reference_image, config)` in `amazon_product_images.py`
  - Iterates shots list, calls Tasks 1+2 per shot
  - Downloads URL reference once, reuses for all shots
  - Returns `{shot_type: Path}` dict
  - Config: shots, size, out_dir, provider, prefix

## Checkpoint 3
- [ ] All 4 PNGs generated from S3 URL reference
- [ ] Shot files named `<prefix>_<shot_type>.png`
- [ ] No regression in existing image generation

## Phase 4: Flask Endpoint
- [ ] **Task 4** — `POST /api/generate-product-images` in `amazon_template_autofill_web.py`
  - Import `generateProductImages` from `amazon_product_images`
  - Validate product dict (needs product_type or category)
  - Return `{ok, images: {shot: abs_path}}`

## Checkpoint 4
- [ ] curl test returns correct JSON
- [ ] Error cases (missing product, unknown provider) return 400
- [ ] Existing catalog endpoint still works

## Phase 5: Catalog Pipeline Integration
- [ ] **Task 5** — Update `process_catalog_row()` in `amazon_template_autofill_web.py`
  - Replace 2-shot block with `generateProductImages()` call
  - Map lifestyle→Other Image URL 1, macro→Other Image URL 2
  - Defensive import with 2-shot fallback

## Checkpoint 5
- [ ] Catalog pipeline produces 4 PNGs per SKU
- [ ] Other Image URL columns populated
- [ ] `generate_images=False` path unaffected
