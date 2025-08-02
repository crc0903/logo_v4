import streamlit as st
from PIL import Image, ImageChops, ImageDraw
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches

PRELOADED_LOGO_DIR = "preloaded_logos"

def load_preloaded_logos():
    logos = []
    if not os.path.exists(PRELOADED_LOGO_DIR):
        os.makedirs(PRELOADED_LOGO_DIR)
    for file in os.listdir(PRELOADED_LOGO_DIR):
        if file.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            name = os.path.splitext(file)[0]
            image = Image.open(os.path.join(PRELOADED_LOGO_DIR, file)).convert("RGBA")
            logos.append((name.lower(), image))
    return logos

def trim_whitespace(image):
    bg = Image.new(image.mode, image.size, (255, 255, 255, 0))  # transparent
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    return image.crop(bbox) if bbox else image

def resize_to_fill_ratio_box(image, target_w, target_h, buffer=0.9):
    img_w, img_h = image.size
    box_ratio = target_w / target_h
    img_ratio = img_w / img_h

    max_w = int(target_w * buffer)
    max_h = int(target_h * buffer)

    if img_ratio > box_ratio:
        new_w = max_w
        new_h = int(max_w / img_ratio)
    else:
        new_h = max_h
        new_w = int(max_h * img_ratio)

    return image.resize((new_w, new_h), Image.LANCZOS)

def add_debug_box(image, box_w, box_h):
    box = Image.new("RGBA", (box_w, box_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(box)
    draw.rectangle([0, 0, box_w - 1, box_h - 1], outline="gray", width=2)
    x = (box_w - image.width) // 2
    y = (box_h - image.height) // 2
    box.paste(image, (x, y), image)
    return box

def create_logo_slide(prs, logos, canvas_width_in, canvas_height_in, logos_per_row):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    canvas_width_px = int(canvas_width_in * 96)
    canvas_height_px = int(canvas_height_in * 96)

    logo_count = len(logos)
    cols = logos_per_row if logos_per_row else max(1, round(math.sqrt(logo_count * canvas_width_in / canvas_height_in)))
    rows = math.ceil(logo_count / cols)

    cell_w = canvas_width_px // cols
    cell_h = canvas_height_px // rows

    left_margin = Inches((10 - canvas_width_in) / 2)
    top_margin = Inches((7.5 - canvas_height_in) / 2)

    for idx, (_, logo_img) in enumerate(logos):
        row, col = divmod(idx, cols)
        trimmed = trim_whitespace(logo_img)
        resized = resize_to_fill_ratio_box(trimmed, cell_w, cell_h)

        # Optional debug: draw the guide box
        final = add_debug_box(resized, cell_w, cell_h)

        img_stream = io.BytesIO()
        final.save(img_stream, format="PNG")
        img_stream.seek(0)

        x_offset = (cell_w - final.width) / 2
        y_offset = (cell_h - final.height) / 2
        left = left_margin + Inches((col * cell_w + x_offset) / 96)
        top = top_margin + Inches((row * cell_h + y_offset) / 96)

        slide.shapes.add_picture(
            img_stream, left, top,
            width=Inches(final.width / 96),
            height=Inches(final.height / 96)
        )

# --- Streamlit UI ---
st.title("Logo Grid PowerPoint Exporter")

uploaded_files = st.file_uploader("Upload logos", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

preloaded = load_preloaded_logos()
selected_preloaded = st.multiselect("Select preloaded logos", options=[name for name, _ in preloaded])

canvas_width_in = st.number_input("Grid width (inches)", min_value=1.0, max_value=20.0, value=10.0)
canvas_height_in = st.number_input("Grid height (inches)", min_value=1.0, max_value=20.0, value=7.5)
logos_per_row = st.number_input("Logos per row (optional)", min_value=0, max_value=50, value=0)

if st.button("Generate PowerPoint"):
    images = []

    if uploaded_files:
        for f in uploaded_files:
            name = os.path.splitext(f.name)[0].lower()
            image = Image.open(f).convert("RGBA")
            images.append((name, image))

    for name, img in preloaded:
        if name in [x.lower() for x in selected_preloaded]:
            images.append((name, img))

    if not images:
        st.warning("Please upload or select logos.")
    else:
        # Sort all logos together alphabetically
        sorted_images = sorted(images, key=lambda x: x[0])

        prs = Presentation()
        create_logo_slide(
            prs, sorted_images,
            canvas_width_in,
            canvas_height_in,
            logos_per_row if logos_per_row > 0 else None
        )

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        st.success("PowerPoint created!")
        st.download_button("Download .pptx", output, file_name="logo_grid.pptx")
