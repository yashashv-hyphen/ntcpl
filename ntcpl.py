import base64
import os
import re
import shutil
import threading
import tkinter as tk
from datetime import datetime
from io import BytesIO
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from huggingface_hub import InferenceClient
from PIL import Image, ImageTk

# =========================================================
# PATHS & CONFIG
# =========================================================

APP_DIR = Path(__file__).resolve().parent
INSTRUCTIONS_FILE = APP_DIR / "image_generation_instructions.txt"
OUTPUT_DIR = APP_DIR / "generated_images"

DEFAULT_IMAGE_MODEL = "black-forest-labs/FLUX.1-schnell"
ANALYSIS_MODEL = "Qwen/Qwen2.5-7B-Instruct"
VISION_MODEL = "Qwen/Qwen2.5-VL-7B-Instruct:fastest"

PLACEHOLDER_NAME = "Enter Product Name"
PLACEHOLDER_DETAILS = (
    "Optional: dimensions, materials, colors, features, specs, "
    "or anything not visible in the photo"
)

SECTION_PRODUCT_ANALYSIS = "PRODUCT_ANALYSIS"
SECTION_LIFESTYLE = "LIFESTYLE_IMAGE"
SECTION_INFOGRAPHIC = "INFOGRAPHIC_IMAGE"


def get_api_key():
    return os.environ.get("HF_TOKEN") or os.environ.get("HF_API_KEY") or os.environ.get("HUGGINGFACE_API_KEY")


def get_client():
    token = get_api_key()
    if not token:
        raise ValueError(
            "Missing Hugging Face API token.\n\n"
            "Set HF_TOKEN or HF_API_KEY in your environment, for example:\n"
            "  export HF_TOKEN=hf_xxxxxxxx"
        )
    return InferenceClient(token=token)


# =========================================================
# INSTRUCTIONS FILE
# =========================================================

def load_instruction_sections():
    if not INSTRUCTIONS_FILE.is_file():
        raise FileNotFoundError(
            f"Instructions file not found:\n{INSTRUCTIONS_FILE}"
        )

    text = INSTRUCTIONS_FILE.read_text(encoding="utf-8")
    sections = {}
    current = None
    lines = []

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

    required = (
        SECTION_PRODUCT_ANALYSIS,
        SECTION_LIFESTYLE,
        SECTION_INFOGRAPHIC,
    )
    missing = [name for name in required if name not in sections]
    if missing:
        raise ValueError(
            "Instructions file is missing sections: "
            + ", ".join(missing)
        )

    return sections


def fill_template(template, **values):
    result = template
    for key, value in values.items():
        result = result.replace("{" + key + "}", value or "")
    return " ".join(result.split())


def load_instructions_for_generation():
    sections = load_instruction_sections()
    return {
        "analysis_template": sections[SECTION_PRODUCT_ANALYSIS],
        "lifestyle_template": sections[SECTION_LIFESTYLE],
        "infographic_template": sections[SECTION_INFOGRAPHIC],
    }


# =========================================================
# PRODUCT ANALYSIS (AI)
# =========================================================

def encode_image_base64(image_path):
    with open(image_path, "rb") as file:
        data = base64.b64encode(file.read()).decode("utf-8")
    suffix = Path(image_path).suffix.lower()
    mime = "image/jpeg" if suffix in {".jpg", ".jpeg"} else "image/png"
    return f"data:{mime};base64,{data}"


def analyze_product_with_vision(client, image_path, analysis_prompt):
    image_url = encode_image_base64(image_path)
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


def analyze_product(product_name, seller_notes, image_path):
    instructions = load_instructions_for_generation()
    analysis_prompt = fill_template(
        instructions["analysis_template"],
        product_name=product_name,
        seller_notes=seller_notes or "None provided.",
        product_details="",
    )

    client = get_client()

    if image_path:
        try:
            return (
                analyze_product_with_vision(client, image_path, analysis_prompt),
                "vision",
            )
        except Exception:
            analysis_prompt += (
                "\n\nNote: A product photo was uploaded but the vision API "
                "was unavailable. Infer the most likely appearance, materials, "
                "and packaging from the product name and seller notes."
            )

    return analyze_product_with_text(client, analysis_prompt), "text"


# =========================================================
# IMAGE GENERATION
# =========================================================

def generate_image(prompt):
    client = get_client()
    image = client.text_to_image(
        prompt=prompt,
        model=DEFAULT_IMAGE_MODEL,
    )
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def build_prompts(product_name, product_details, instructions):
    lifestyle_prompt = fill_template(
        instructions["lifestyle_template"],
        product_name=product_name,
        product_details=product_details,
    )
    infographic_prompt = fill_template(
        instructions["infographic_template"],
        product_name=product_name,
        product_details=product_details,
    )
    return lifestyle_prompt, infographic_prompt


def slugify(text):
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", text.lower()).strip("_")
    return slug[:60] or "product"


def save_generation_bundle(
    product_name,
    product_details,
    analysis_method,
    lifestyle_prompt,
    infographic_prompt,
    lifestyle_bytes,
    infographic_bytes,
    source_image_path,
):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = OUTPUT_DIR / f"{slugify(product_name)}_{timestamp}"
    folder.mkdir(parents=True, exist_ok=True)

    lifestyle_path = folder / "lifestyle_image.png"
    infographic_path = folder / "infographic_image.png"
    lifestyle_path.write_bytes(lifestyle_bytes)
    infographic_path.write_bytes(infographic_bytes)

    if source_image_path:
        ext = Path(source_image_path).suffix or ".png"
        shutil.copy2(source_image_path, folder / f"original_product{ext}")

    (folder / "product_analysis.txt").write_text(
        product_details,
        encoding="utf-8",
    )
    (folder / "prompts_used.txt").write_text(
        "\n\n".join(
            [
                f"Analysis method: {analysis_method}",
                f"Instructions file: {INSTRUCTIONS_FILE}",
                "",
                "=== LIFESTYLE PROMPT ===",
                lifestyle_prompt,
                "",
                "=== INFOGRAPHIC PROMPT ===",
                infographic_prompt,
            ]
        ),
        encoding="utf-8",
    )
    shutil.copy2(INSTRUCTIONS_FILE, folder / "image_generation_instructions.txt")

    return folder


# =========================================================
# DISPLAY IMAGE
# =========================================================

def display_image(image_bytes, label):
    image = Image.open(BytesIO(image_bytes))
    image.thumbnail((400, 400))
    tk_image = ImageTk.PhotoImage(image)
    label.config(image=tk_image)
    label.image = tk_image


# =========================================================
# UPLOAD IMAGE
# =========================================================

selected_product_image = None


def upload_image():
    global selected_product_image

    file_path = filedialog.askopenfilename(
        filetypes=[("Images", "*.png *.jpg *.jpeg *.webp")]
    )
    if not file_path:
        return

    selected_product_image = file_path
    image = Image.open(file_path)
    image.thumbnail((250, 250))
    tk_image = ImageTk.PhotoImage(image)
    uploaded_image_label.config(image=tk_image)
    uploaded_image_label.image = tk_image


# =========================================================
# FORM HELPERS
# =========================================================

def get_product_name():
    name = product_entry.get().strip()
    if not name or name == PLACEHOLDER_NAME:
        return None
    return name


def get_seller_notes():
    notes = details_text.get("1.0", tk.END).strip()
    if not notes or notes == PLACEHOLDER_DETAILS:
        return ""
    return notes


def clear_name_placeholder(_event=None):
    if product_entry.get() == PLACEHOLDER_NAME:
        product_entry.delete(0, tk.END)
        product_entry.config(fg="black")


def restore_name_placeholder(_event=None):
    if not product_entry.get().strip():
        product_entry.insert(0, PLACEHOLDER_NAME)
        product_entry.config(fg="gray")


def clear_details_placeholder(_event=None):
    if details_text.get("1.0", tk.END).strip() == PLACEHOLDER_DETAILS:
        details_text.delete("1.0", tk.END)
        details_text.config(fg="black")


def restore_details_placeholder(_event=None):
    if not details_text.get("1.0", tk.END).strip():
        details_text.insert("1.0", PLACEHOLDER_DETAILS)
        details_text.config(fg="gray")


# =========================================================
# GENERATION LOGIC
# =========================================================

def set_status(text):
    status_label.config(text=text)


def set_generate_enabled(enabled):
    generate_button.config(state="normal" if enabled else "disabled")


def on_generation_success(result):
    display_image(result["lifestyle_bytes"], output_label_1)
    display_image(result["infographic_bytes"], output_label_2)
    set_status(f"Saved to: {result['folder']}")
    set_generate_enabled(True)
    messagebox.showinfo(
        "Success",
        f"Images generated and saved to:\n{result['folder']}",
    )


def on_generation_error(error_message):
    messagebox.showerror("Generation Failed", error_message)
    set_status("Generation failed.")
    set_generate_enabled(True)


def update_status_from_worker(text):
    root.after(0, lambda: set_status(text))


def generate_amazon_images_worker(product_name, seller_notes, image_path):
    try:
        update_status_from_worker("Analyzing product (reading instructions)...")
        instructions = load_instructions_for_generation()

        product_details, analysis_method = analyze_product(
            product_name,
            seller_notes,
            image_path,
        )

        lifestyle_prompt, infographic_prompt = build_prompts(
            product_name,
            product_details,
            instructions,
        )

        update_status_from_worker("Generating lifestyle image...")
        lifestyle_bytes = generate_image(lifestyle_prompt)

        update_status_from_worker("Generating infographic image...")
        infographic_bytes = generate_image(infographic_prompt)

        folder = save_generation_bundle(
            product_name,
            product_details,
            analysis_method,
            lifestyle_prompt,
            infographic_prompt,
            lifestyle_bytes,
            infographic_bytes,
            image_path,
        )

        root.after(
            0,
            lambda: on_generation_success(
                {
                    "lifestyle_bytes": lifestyle_bytes,
                    "infographic_bytes": infographic_bytes,
                    "folder": folder,
                }
            ),
        )
    except Exception as error:
        root.after(0, lambda: on_generation_error(str(error)))


def generate_amazon_images():
    product_name = get_product_name()
    if not product_name:
        messagebox.showerror("Error", "Please enter a product name.")
        return

    if not selected_product_image:
        messagebox.showerror(
            "Error",
            "Please upload a product image to sell on Amazon.",
        )
        return

    if not get_api_key():
        messagebox.showerror(
            "Missing API Token",
            "Set HF_TOKEN or HF_API_KEY before generating images.",
        )
        return

    try:
        load_instruction_sections()
    except (FileNotFoundError, ValueError) as error:
        messagebox.showerror("Instructions Error", str(error))
        return

    set_status("Starting generation...")
    set_generate_enabled(False)

    thread = threading.Thread(
        target=generate_amazon_images_worker,
        args=(product_name, get_seller_notes(), selected_product_image),
        daemon=True,
    )
    thread.start()


# =========================================================
# GUI
# =========================================================

root = tk.Tk()
root.title("Amazon AI Product Image Generator")
root.geometry("1150x920")
root.configure(bg="#f4f4f4")

title_label = tk.Label(
    root,
    text="Amazon AI Product Image Generator",
    font=("Arial", 24, "bold"),
    bg="#f4f4f4",
)
title_label.pack(pady=15)

subtitle = tk.Label(
    root,
    text=f"Instructions: {INSTRUCTIONS_FILE.name}  |  Output: {OUTPUT_DIR.name}/",
    font=("Arial", 10),
    bg="#f4f4f4",
    fg="#555555",
)
subtitle.pack()

upload_button = tk.Button(
    root,
    text="Upload Product Image (required)",
    command=upload_image,
    font=("Arial", 12),
    padx=20,
    pady=8,
)
upload_button.pack(pady=8)

uploaded_image_label = tk.Label(root, bg="#dddddd", width=250, height=250)
uploaded_image_label.pack(pady=10)

product_entry = tk.Entry(root, width=50, font=("Arial", 14), fg="gray")
product_entry.pack(pady=6)
product_entry.insert(0, PLACEHOLDER_NAME)
product_entry.bind("<FocusIn>", clear_name_placeholder)
product_entry.bind("<FocusOut>", restore_name_placeholder)

details_label = tk.Label(
    root,
    text="Product details (optional — improves AI recognition)",
    font=("Arial", 11),
    bg="#f4f4f4",
)
details_label.pack(pady=(8, 2))

details_text = scrolledtext.ScrolledText(
    root,
    width=70,
    height=4,
    font=("Arial", 11),
    fg="gray",
    wrap=tk.WORD,
)
details_text.pack(pady=4)
details_text.insert("1.0", PLACEHOLDER_DETAILS)
details_text.bind("<FocusIn>", clear_details_placeholder)
details_text.bind("<FocusOut>", restore_details_placeholder)

generate_button = tk.Button(
    root,
    text="Generate Amazon Images",
    command=generate_amazon_images,
    font=("Arial", 14, "bold"),
    bg="#4CAF50",
    fg="white",
    padx=25,
    pady=10,
)
generate_button.pack(pady=14)

status_label = tk.Label(root, text="", font=("Arial", 12), bg="#f4f4f4")
status_label.pack()

output_frame = tk.Frame(root, bg="#f4f4f4")
output_frame.pack(pady=20)

frame1 = tk.Frame(output_frame, bg="#f4f4f4")
frame1.grid(row=0, column=0, padx=20)

tk.Label(
    frame1,
    text="Lifestyle (real-life usage)",
    font=("Arial", 16, "bold"),
    bg="#f4f4f4",
).pack(pady=8)

output_label_1 = tk.Label(frame1, bg="#cccccc", width=400, height=400)
output_label_1.pack()

frame2 = tk.Frame(output_frame, bg="#f4f4f4")
frame2.grid(row=0, column=1, padx=20)

tk.Label(
    frame2,
    text="Infographic (features & specs)",
    font=("Arial", 16, "bold"),
    bg="#f4f4f4",
).pack(pady=8)

output_label_2 = tk.Label(frame2, bg="#cccccc", width=400, height=400)
output_label_2.pack()

if __name__ == "__main__":
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    root.mainloop()
