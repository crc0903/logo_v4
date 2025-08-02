import streamlit as st
from PIL import Image, ImageChops
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

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

def resize_to_5x2_box(image, box_width, box_height):
    img_w, img_h = image.size
    target_ratio = box_width / box_height
    img_ratio = img_w / img_h

    if img_ratio > target_ratio:
        scale = box_width / img_w
    else:
        scale = box_height / img_h

    new_w = int(img_w * scale)
    new_h = int(img_h * scale)
    return image.resize((new_w, new_h), Image.LANCZOS)

def create_logo_slide(prs, logos, grid_width_in, grid_height_in):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    grid_width = Inches(grid_width_in)
    grid_height = Inches(grid_height_in)
    box_width = grid_width / 5
    box_height = grid_height / 2

    start_x = (slide_width - grid_width) / 2
    start_y = (slide_height - grid_height) / 2

    # Add guideline box (optional)
    slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=start_x,
        top=start_y,
        width=grid_width,
        height=grid_height
    ).line.color.rgb = RGBColor(180, 180, 180)

    for i, (name, img) in enumerate(logos):
        col = i % 5
        row = i // 5

        x = start_x + col * box_width
        y = start_y + row * box_height

        img = trim_whitespace(img)
        img = resize_to_5x2_box(img, box_width, box_height)

        stream = io.BytesIO()
        img.save(stream, format="PNG")
        stream.seek(0)

        left = x + (box_width - img.width) / 2
        top = y + (box_height - img.height) / 2
        slide.shapes.add_picture(stream, left=left, top=top, width=img.width, height=img.height)

def main():
    st.title("Logo Grid Exporter (5:2 Boxes)")

    st.markdown("Select logos from below and/or upload your own:")
    preloaded = load_preloaded_logos()
    selected = st.multiselect("Preloaded Logos", options=[name for name, _ in preloaded])
    uploaded_files = st.file_uploader("Upload Logos", accept_multiple_files=True, type=["png", "jpg", "jpeg", "webp"])

    grid_width = st.number_input("Grid width (inches)", min_value=1.0, max_value=15.0, value=10.0, step=0.5)
    grid_height = st.number_input("Grid height (inches)", min_value=1.0, max_value=10.0, value=4.0, step=0.5)

    if st.button("Generate PowerPoint"):
        selected_logos = [item for item in preloaded if item[0] in selected]

        for file in uploaded_files:
            name = os.path.splitext(file.name)[0].lower()
            image = Image.open(file).convert("RGBA")
            selected_logos.append((name, image))

        selected_logos.sort(key=lambda x: x[0])  # Alphabetize all

        prs = Presentation()
        create_logo_slide(prs, selected_logos, grid_width, grid_height)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        st.download_button("Download PowerPoint", data=output, file_name="logo_grid.pptx")

if __name__ == "__main__":
    main()
