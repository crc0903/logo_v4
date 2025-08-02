import streamlit as st
from PIL import Image, ImageChops
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

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
    bg = Image.new("RGBA", image.size, (255, 255, 255, 0))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    return image.crop(bbox) if bbox else image

def resize_image_proportionally(image, max_width, max_height):
    img_w, img_h = image.size
    ratio = min(max_width / img_w, max_height / img_h)
    return image.resize((int(img_w * ratio), int(img_h * ratio)), Image.LANCZOS)

def create_logo_slide(prs, logos, logos_per_row=5):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    cell_width = slide_width / logos_per_row
    cell_height = cell_width * 2 / 5  # 5:2 box ratio

    rows = math.ceil(len(logos) / logos_per_row)

    for idx, (name, image) in enumerate(logos):
        col = idx % logos_per_row
        row = idx // logos_per_row

        trimmed = trim_whitespace(image)
        resized = resize_image_proportionally(trimmed, cell_width * 0.9, cell_height * 0.9)

        img_stream = io.BytesIO()
        resized.save(img_stream, format="PNG")
        img_stream.seek(0)

        left = cell_width * col + (cell_width - resized.width) / 2
        top = cell_height * row + (cell_height - resized.height) / 2

        slide.shapes.add_picture(
            img_stream,
            left=left,
            top=top,
            width=resized.width,
            height=resized.height
        )

        # Add 5:2 guideline box
        guide = slide.shapes.add_shape(
            autoshape_type_id=1,
            left=cell_width * col,
            top=cell_height * row,
            width=cell_width,
            height=cell_height
        )
        guide.line.color.rgb = RGBColor(100, 100, 100)
        guide.fill.background()

# --- Streamlit UI ---
st.title("Logo Grid PowerPoint Exporter")
st.markdown("Upload logos or select from preloaded options:")

uploaded_files = st.file_uploader("Upload logos", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
preloaded = load_preloaded_logos()
selected_preloaded = st.multiselect("Select preloaded logos", options=sorted([name for name, _ in preloaded]))

logos_per_row = st.number_input("Logos per row", min_value=1, max_value=20, value=5)

if st.button("Generate PowerPoint"):
    logos = []

    # Combine uploaded and selected preloaded logos
    if uploaded_files:
        for f in uploaded_files:
            name = os.path.splitext(f.name)[0].lower()
            image = Image.open(f).convert("RGBA")
            logos.append((name, image))

    for name, image in preloaded:
        if name in [n.lower() for n in selected_preloaded]:
            logos.append((name, image))

    if not logos:
        st.warning("Please upload or select logos.")
    else:
        # Alphabetize
        logos.sort(key=lambda x: x[0])

        prs = Presentation()
        create_logo_slide(prs, logos, logos_per_row)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        st.success("PowerPoint created!")
        st.download_button("Download .pptx", output, file_name="logo_grid.pptx")
