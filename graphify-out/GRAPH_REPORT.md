# Graph Report - .  (2026-06-01)

## Corpus Check
- 27 files · ~215,475 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 257 nodes · 633 edges · 27 communities (24 shown, 3 thin omitted)
- Extraction: 91% EXTRACTED · 9% INFERRED · 0% AMBIGUOUS · INFERRED: 60 edges (avg confidence: 0.79)
- Token cost: 4,685 input · 2,157 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Template Autofill Core|Template Autofill Core]]
- [[_COMMUNITY_NTCPL Product Analyzer|NTCPL Product Analyzer]]
- [[_COMMUNITY_Web App API Endpoints|Web App API Endpoints]]
- [[_COMMUNITY_AI Copy Generation|AI Copy Generation]]
- [[_COMMUNITY_HF Image Generation|HF Image Generation]]
- [[_COMMUNITY_AI Asset Pipeline|AI Asset Pipeline]]
- [[_COMMUNITY_Instruction Templates|Instruction Templates]]
- [[_COMMUNITY_Leash Infographic|Leash Infographic]]
- [[_COMMUNITY_Listing Quality Validation|Listing Quality Validation]]
- [[_COMMUNITY_Leash Original Product Photo|Leash Original Product Photo]]
- [[_COMMUNITY_AI Product Photo v1|AI Product Photo v1]]
- [[_COMMUNITY_File Upload Endpoints|File Upload Endpoints]]
- [[_COMMUNITY_Leash Lifestyle Image|Leash Lifestyle Image]]
- [[_COMMUNITY_Product Detail Layout|Product Detail Layout]]
- [[_COMMUNITY_Uploaded Product Shots v3|Uploaded Product Shots v3]]
- [[_COMMUNITY_AI Ref Test Image 1|AI Ref Test Image 1]]
- [[_COMMUNITY_AI Ref Test Mock Listing|AI Ref Test Mock Listing]]
- [[_COMMUNITY_Uploaded Product Shots v4|Uploaded Product Shots v4]]
- [[_COMMUNITY_Uploaded Product Shots v2|Uploaded Product Shots v2]]
- [[_COMMUNITY_Uploaded Product Shots v1|Uploaded Product Shots v1]]
- [[_COMMUNITY_Earbuds Lifestyle Image|Earbuds Lifestyle Image]]
- [[_COMMUNITY_HF Vision Model Config|HF Vision Model Config]]
- [[_COMMUNITY_Earbuds Infographic|Earbuds Infographic]]
- [[_COMMUNITY_Earbuds Original Product|Earbuds Original Product]]
- [[_COMMUNITY_NTCPL Web Flask App|NTCPL Web Flask App]]
- [[_COMMUNITY_Earbuds Product Analysis|Earbuds Product Analysis]]
- [[_COMMUNITY_Earbuds Image Prompts|Earbuds Image Prompts]]

## God Nodes (most connected - your core abstractions)
1. `str` - 49 edges
2. `AutoFillGUI` - 28 edges
3. `str` - 20 edges
4. `main()` - 19 edges
5. `Path` - 17 edges
6. `apply_ai_assets_to_row()` - 16 edges
7. `analyze_product_image()` - 14 edges
8. `ColumnMeta` - 13 edges
9. `generate_amazon_listing()` - 13 edges
10. `read_column_metadata()` - 12 edges

## Surprising Connections (you probably didn't know these)
- `load_instruction_sections()` --semantically_similar_to--> `load_instruction_sections()`  [INFERRED] [semantically similar]
  amazon_template_autofill_web.py → ntcpl.py
- `save_generation_bundle()` --references--> `leash product_analysis.txt`  [INFERRED]
  ntcpl.py → generated_images/leash_20260515_125919/product_analysis.txt
- `save_generation_bundle()` --references--> `leash prompts_used.txt`  [INFERRED]
  ntcpl.py → generated_images/leash_20260515_125919/prompts_used.txt
- `load_instruction_sections()` --semantically_similar_to--> `load_instruction_sections()`  [EXTRACTED] [semantically similar]
  ntcpl_web.py → ntcpl.py
- `fill_template()` --semantically_similar_to--> `fill_template()`  [EXTRACTED] [semantically similar]
  ntcpl_web.py → ntcpl.py

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Template Parsing Utilities (shared across CLI, GUI, Web)** — amazon_template_autofill_parse_template_rows, amazon_template_autofill_read_column_metadata, amazon_template_autofill_read_data_definitions, amazon_template_autofill_read_browse_nodes, amazon_template_autofill_read_recommended_node_to_product_type, amazon_template_autofill_main, amazon_template_autofill_gui_autofillgui, amazon_template_autofill_web_load_template_meta [EXTRACTED 1.00]
- **Hugging Face API client initializers across modules** — amazon_template_autofill_hf_client, ntcpl_get_client, ntcpl_web_get_client [INFERRED 0.85]
- **Listing generation pipeline (analyze -> generate -> validate -> refine)** — amazon_template_autofill_web_analyze_product_image, amazon_template_autofill_web_generate_listing, amazon_template_autofill_web_validate_listing, amazon_template_autofill_web_refine_listing, amazon_template_autofill_web_generate_amazon_listing [EXTRACTED 1.00]
- **Modules consuming image_generation_instructions.txt** — ntcpl_load_instruction_sections, ntcpl_web_load_instruction_sections, amazon_template_autofill_web_load_instruction_sections, image_generation_instructions_document [EXTRACTED 1.00]

## Communities (27 total, 3 thin omitted)

### Community 0 - "Template Autofill Core"
Cohesion: 0.13
Nodes (37): all_attrs_by_label(), apply_values(), ask_user_for_required(), build_defaults(), column_letter_from_index(), ColumnMeta, create_marketing_images(), first_attr_by_label() (+29 more)

### Community 1 - "NTCPL Product Analyzer"
Cohesion: 0.13
Nodes (26): Hugging Face Inference Backend, leash product_analysis.txt, leash prompts_used.txt, analyze_product(), analyze_product_with_text(), analyze_product_with_vision(), build_prompts(), display_image() (+18 more)

### Community 2 - "Web App API Endpoints"
Cohesion: 0.17
Nodes (17): all_attrs_by_label(), api_describe_image(), api_meta(), api_save(), build_product_details_for_instructions(), copy_row_style(), describe_image(), _fallback_copy_from_first_image() (+9 more)

### Community 3 - "AI Copy Generation"
Cohesion: 0.20
Nodes (19): analyze_product_image(), build_listing_prompt(), CopyGenerationError, _extract_hf_chat_text(), extract_json(), _fallback_analysis_from_image(), _fallback_listing_from_analysis(), generate_amazon_listing() (+11 more)

### Community 4 - "HF Image Generation"
Cohesion: 0.24
Nodes (15): ai_generate_images_from_prompts(), encode_image_b64(), _fallback_image_bytes(), _hf_image_model(), hf_project_generate_images(), _hf_vision_fallback_models(), _huggingface_api_key(), huggingface_generate_image_bytes() (+7 more)

### Community 5 - "AI Asset Pipeline"
Cohesion: 0.24
Nodes (14): ai_generate_copy_from_image(), analyze_product_hf(), api_ai_generate_assets(), api_ai_generate_row(), apply_ai_assets_to_row(), build_analysis_prompt(), build_instruction_prompts(), fill_template() (+6 more)

### Community 6 - "Instruction Templates"
Cohesion: 0.25
Nodes (12): Image Generation Instructions Format (section-based template), earbuds image_generation_instructions.txt (copy), leash image_generation_instructions.txt (copy), analyze_product_with_text(), analyze_product_with_vision(), api_generate(), fill_template(), generate_image() (+4 more)

### Community 7 - "Leash Infographic"
Cohesion: 0.36
Nodes (9): Customer Features - Feature Section 2, Faduree Features - Comfortable Grip and Durability, Leash Product Infographic, Dog Leash Product, Legro Blemth - Feature Section 1, Metal Clip / Snap Hook, Black Nylon Strap Design, Residritice - Feature Section 4 (+1 more)

### Community 8 - "Listing Quality Validation"
Cohesion: 0.29
Nodes (8): _analysis_terms(), audit_listing_copy(), _extract_visual_keywords(), Return human-readable gaps; empty list means copy aligns with listing rules., validate_listing(), _word_overlap_score(), Listing Copy Validation Pipeline, float

### Community 9 - "Leash Original Product Photo"
Cohesion: 0.38
Nodes (7): Brass Snap Hook Clasp, Brown / Tan Color, Dog Leash - Product Type, Flat Strap Design, Leather Dog Leash - Original Product Photo, Leather Material, White Background Product Presentation

### Community 10 - "AI Product Photo v1"
Cohesion: 0.40
Nodes (6): Light Blue Studio Background with Shadow, Brass Snap Hook Hardware, AI-Generated Product Photography Style, Brown Leather Dog Leash with Brass Snap Hook, Leather Material - Brown, Pet Accessories - Dog Leash

### Community 11 - "File Upload Endpoints"
Cohesion: 0.53
Nodes (6): api_upload_image(), api_upload_template(), ensure_template_upload_dir(), ensure_upload_dir(), generate(), Path

### Community 12 - "Leash Lifestyle Image"
Cohesion: 0.53
Nodes (6): Lifestyle Image - Dog Leash Product, Dog Owner Target Customer, Energetic and Cheerful Mood, Dog Leash with Harness, Outdoor Park Setting with Blue Sky, Happy Dog Wearing Harness on Leash

### Community 13 - "Product Detail Layout"
Cohesion: 0.60
Nodes (5): Amazon Product Listing Content, Dog Leash, Dog Leash Product Details AI Image, Leather Dog Leash with Brass Clip, Product Details Layout with Text and Image Reference

### Community 14 - "Uploaded Product Shots v3"
Cohesion: 0.40
Nodes (5): Brass Snap Hook Clip, Brown Leather Material, Dog Training Equipment, Leather Dog Leash Product Image, Pet Accessory Category

### Community 15 - "AI Ref Test Image 1"
Cohesion: 0.70
Nodes (5): Dog Leash Product - braided rope leash with metal snap clip, Garbled AI-generated Text Overlay - nonsensical bullet point text typical of AI image generation artifacts, ref_test_ai_1.png - AI-generated product reference test image of a dog leash, Outdoor Park Background - autumn trees, blurred background environment, Amazon Product Listing Layout - composite image with product photo, bullet points, and label

### Community 16 - "AI Ref Test Mock Listing"
Cohesion: 0.60
Nodes (5): Grey Dog Leash Product - Coiled Nylon Leash with Metal Clip, 25m, AI-Generated Amazon Product Listing - Grey Dog Leash (ref_test_ai_2), Product Price $219, Product Rating 3.5 Stars, Amazon-Style Product Page UI Layout with CTA Button

### Community 17 - "Uploaded Product Shots v4"
Cohesion: 0.50
Nodes (4): Brass Snap Hook Clasp, Brown Leather Material, Leather Dog Leash Product Image, Pet Accessory

### Community 18 - "Uploaded Product Shots v2"
Cohesion: 0.50
Nodes (4): Pet Supplies - Dog Leash, Brass Snap Hook Clasp, Brown Leather Material, Leather Dog Leash with Brass Snap Hook

### Community 19 - "Uploaded Product Shots v1"
Cohesion: 0.50
Nodes (4): Brass Snap Hook Clasp, Brown Leather Material, Leather Dog Leash Product Image, Pet Accessory

### Community 20 - "Earbuds Lifestyle Image"
Cohesion: 0.67
Nodes (4): Target Audience - Active Consumer, Wireless Bluetooth Earbuds Lifestyle Image, Wireless Bluetooth Earbuds, Lifestyle Scene - Product In Use

### Community 21 - "HF Vision Model Config"
Cohesion: 0.67
Nodes (3): _hf_vision_model(), _hf_vision_model_for_endpoint(), Return an endpoint-safe HF model id for raw HTTP inference calls.     Some SDK-s

### Community 22 - "Earbuds Infographic"
Cohesion: 0.67
Nodes (3): Earbuds Product Features, Wireless Bluetooth Earbuds Infographic Image, Wireless Bluetooth Earbuds Product

### Community 23 - "Earbuds Original Product"
Cohesion: 0.67
Nodes (3): Dark/Black Background, Wireless Bluetooth Earbuds - Original Product Photo, Wireless Bluetooth Earbuds Product Category

## Knowledge Gaps
- **38 isolated node(s):** `Image`, `InferenceClient`, `Image`, `InferenceClient`, `float` (+33 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **3 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `load_instruction_sections()` connect `NTCPL Product Analyzer` to `AI Asset Pipeline`, `Instruction Templates`?**
  _High betweenness centrality (0.085) - this node is a cross-community bridge._
- **Why does `load_instruction_sections()` connect `AI Asset Pipeline` to `NTCPL Product Analyzer`, `Web App API Endpoints`, `HF Image Generation`, `Instruction Templates`?**
  _High betweenness centrality (0.052) - this node is a cross-community bridge._
- **Are the 9 inferred relationships involving `ValueError` (e.g. with `create_marketing_images()` and `.save_excel()`) actually correct?**
  _`ValueError` has 9 INFERRED edges - model-reasoned connections that need verification._
- **Are the 3 inferred relationships involving `AutoFillGUI` (e.g. with `ColumnMeta` and `RequirementMeta`) actually correct?**
  _`AutoFillGUI` has 3 INFERRED edges - model-reasoned connections that need verification._
- **Are the 2 inferred relationships involving `main()` (e.g. with `Amazon Listing Pipeline` and `ValueError`) actually correct?**
  _`main()` has 2 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Image`, `InferenceClient`, `Image` to the rest of the system?**
  _44 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Template Autofill Core` be split into smaller, more focused modules?**
  _Cohesion score 0.125544267053701 - nodes in this community are weakly interconnected._