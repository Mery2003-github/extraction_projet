import fitz  # PyMuPDF
import json
import base64
from PIL import Image
from skimage import io as skio
import io as pyio
import pillow_avif  


def int_color_to_hex(color):
    if isinstance(color, tuple):
        r = int(color[0] * 255)
        g = int(color[1] * 255)
        b = int(color[2] * 255)
        return f'#{r:02X}{g:02X}{b:02X}'
    elif isinstance(color, int):
        r = (color >> 16) & 255
        g = (color >> 8) & 255
        b = color & 255
        return f'#{r:02X}{g:02X}{b:02X}'
    return '#000000'


def process_image_bytes(img_bytes, max_width=150, quality=50, force_white_bg=True):
    pil_img = Image.open(pyio.BytesIO(img_bytes))

    if pil_img.mode != "RGBA":
        pil_img = pil_img.convert("RGBA")

    if force_white_bg:
        bg = Image.new("RGBA", pil_img.size, (255, 255, 255, 255))
        pil_img = Image.alpha_composite(bg, pil_img)

    pil_img = pil_img.convert("RGB")

    w, h = pil_img.size
    aspect_ratio = w / h  
    if w > max_width:
        ratio = max_width / w
        new_h = int(h * ratio)
        pil_img = pil_img.resize((max_width, new_h), Image.LANCZOS)

    buf = pyio.BytesIO()
    pil_img.save(buf, format="AVIF", quality=quality)
    return "data:image/avif;base64," + base64.b64encode(buf.getvalue()).decode("utf-8"), aspect_ratio


def extract_pdf_to_json(pdf_path, max_image_width=150, quality=10, scale_factor=1.0):
    doc = fitz.open(pdf_path)
    pages = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        texts = []

      
        for block in page.get_text("dict")["blocks"]:
            if block['type'] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        color_hex = int_color_to_hex(span.get("color", 0))
                        flags = span.get("flags", 0)
                        texts.append({
    "text": span["text"],
    "top": float(span["bbox"][1]) * scale_factor,
    "left": float(span["bbox"][0]) * scale_factor,
    "width": (float(span["bbox"][2] - span["bbox"][0]) + 1) * scale_factor,  
    "height": float(span["bbox"][3] - span["bbox"][1]) * scale_factor,
    "font_size": float(span["size"]) * scale_factor,
    "font_family": span.get("font", "unknown"),
    "color": color_hex,
    "bold": bool(flags & 2),
    "italic": bool(flags & 1)
})


        rects = []
        for shape in page.get_drawings():
            fill_color_hex = int_color_to_hex(shape.get("fill", 0)) if shape.get("fill") else None
            stroke_color_hex = int_color_to_hex(shape.get("color", 0))
            width_line = (shape.get("width", 1) or 1) * scale_factor
            for item in shape["items"]:
                if item[0] == "re":
                    r = item[1]
                    rects.append({
                        "type": "rectangle",
                        "top": int(r.y0 * scale_factor),
                        "left": int(r.x0 * scale_factor),
                        "width": int(r.width * scale_factor),
                        "height": int(r.height * scale_factor),
                        "stroke_color": stroke_color_hex,
                        "fill_color": fill_color_hex,
                        "width_line": width_line
                    })
                elif item[0] == "l":
                    p0, p1 = item[1], item[2]
                    rects.append({
                        "type": "line",
                        "from": {"x": p0[0] * scale_factor, "y": p0[1] * scale_factor},
                        "to": {"x": p1[0] * scale_factor, "y": p1[1] * scale_factor},
                        "stroke_color": stroke_color_hex,
                        "width_line": width_line
                    })

     
        images = []
        for img in page.get_images(full=True):
            xref = img[0]
            bbox = page.get_image_bbox(img)
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]

            smask_xref = base_image.get("smask")
            if smask_xref:
                mask_image = doc.extract_image(smask_xref)
                mask_bytes = mask_image["image"]

                img_pil = Image.open(pyio.BytesIO(img_bytes)).convert("RGB")
                mask_pil = Image.open(pyio.BytesIO(mask_bytes)).convert("L")
                img_pil.putalpha(mask_pil)

                img_pil = Image.alpha_composite(Image.new("RGBA", img_pil.size, (255, 255, 255, 255)), img_pil)
                img_pil = img_pil.convert("RGB")

                buf = pyio.BytesIO()
                img_pil.save(buf, format="AVIF", quality=quality)
                img_base64 = "data:image/avif;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")
                aspect_ratio = img_pil.width / img_pil.height  
            else:
                img_np = skio.imread(pyio.BytesIO(img_bytes))
                is_icon = img_np.ndim == 2 or (img_np.shape[-1] == 1) or max(img_np.shape[:2]) <= 64
                max_w = 32 if is_icon else max_image_width
                q = 5 if is_icon else quality
                img_base64, aspect_ratio = process_image_bytes(img_bytes, max_width=max_w, quality=q)

            images.append({
                "top": int(bbox.y0 * scale_factor),
                "left": int(bbox.x0 * scale_factor),
                "width": int(bbox.width * scale_factor),
                "height": int(bbox.height * scale_factor),
                "aspect_ratio": aspect_ratio,  
                "base64": img_base64
            })

       
        pages.append({
            "page_index": page_num + 1,
            "page_width": int(page.rect.width * scale_factor),
            "page_height": int(page.rect.height * scale_factor),
            "texts": texts,
            "rectangles": rects,
            "images": images
        })

    return {"pages": pages}

if __name__ == "__main__":
    scale_factor = 1
    pdf_file = r"pdfs\Document sans titre (31).pdf"

  
    data = extract_pdf_to_json(pdf_file, scale_factor=scale_factor)

  
    print("\n--- JSON complet extrait ---\n")
    print(json.dumps(data, indent=2, ensure_ascii=False))

   
    print("\n--- Résumé des dimensions et ratio d'aspect exact ---")
    for page in data["pages"]:
        original_width = page["page_width"] / scale_factor
        original_height = page["page_height"] / scale_factor
        aspect_ratio_original = original_width / original_height

        new_width = page["page_width"]
        new_height = page["page_height"]
        aspect_ratio_new = new_width / new_height

     
        ratio_conserve = aspect_ratio_original == aspect_ratio_new

        print(f"\nPage {page['page_index']}:")
        print(" Dimensions de la feuille PDF :")
        print(f"  - Largeur originale : {original_width}px")
        print(f"  - Hauteur originale : {original_height}px")
        print(f"  - Ratio d'aspect original (W/H) : {aspect_ratio_original}")
        print()
        print(" Dimensions après mise à l'échelle :")
        print(f"  - Largeur cible : {new_width}px")
        print(f"  - Hauteur calculée : {new_height}px")
        print(f"  - Ratio d'aspect nouveau (W/H) : {aspect_ratio_new}")
        print(f"  - Ratio conservé : {ratio_conserve}")
