import fitz  # PyMuPDF
import base64
import json
import io
from PIL import Image

def int_color_to_hex(color):
    if isinstance(color, tuple):
        r = int(color[0]*255)
        g = int(color[1]*255)
        b = int(color[2]*255)
        return f'#{r:02X}{g:02X}{b:02X}'
    elif isinstance(color, int):
        r = (color >> 16) & 255
        g = (color >> 8) & 255
        b = color & 255
        return f'#{r:02X}{g:02X}{b:02X}'
    else:
        return '#000000'

def extract_pdf_to_json(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)

       
        text_blocks = page.get_text("dict")["blocks"]
        texts = []
        for block in text_blocks:
            if block['type'] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        color_hex = int_color_to_hex(span.get("color", 0))
                        flags = span.get("flags", 0)
                        is_bold = bool(flags & 2)
                        is_italic = bool(flags & 1)

                        texts.append({
                            "text": span["text"],
                            "top": int(span["bbox"][1]),
                            "left": int(span["bbox"][0]),
                            "width": int(span["bbox"][2] - span["bbox"][0]),
                            "height": int(span["bbox"][3] - span["bbox"][1]),
                            "font_size": int(span["size"]),
                            "font_family": span.get("font", "unknown"),
                            "color": color_hex,
                            "bold": is_bold,
                            "italic": is_italic
                        })

        
        shapes = page.get_drawings()
        rects = []
        for shape in shapes:
            fill_color_value = shape.get("fill", None)
            fill_color_hex = int_color_to_hex(fill_color_value) if fill_color_value is not None else None
            stroke_color = shape.get("color", 0)
            width_line = shape.get("width", 1) or 1

            for item in shape["items"]:
                if item[0] == "re":
                    r = item[1]
                    rects.append({
                        "type": "rectangle",
                        "top": int(r.y0),
                        "left": int(r.x0),
                        "width": int(r.width),
                        "height": int(r.height),
                        "stroke_color": int_color_to_hex(stroke_color),
                        "fill_color": fill_color_hex,
                        "width_line": width_line
                    })
                elif item[0] == "l":
                    p0, p1 = item[1], item[2]
                    rects.append({
                        "type": "line",
                        "from": {"x": p0[0], "y": p0[1]},
                        "to": {"x": p1[0], "y": p1[1]},
                        "stroke_color": int_color_to_hex(stroke_color),
                        "width_line": width_line
                    })

       
        images = []
        for img in page.get_images(full=True):
            xref = img[0]
            bbox = page.get_image_bbox(img)
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]

            try:
                image_pil = Image.open(io.BytesIO(img_bytes))
                img_io = io.BytesIO()
                # Convertir en WEBP compressé (qualité ajustable 60 ici)
                image_pil.convert("RGBA").save(img_io, format="WEBP", quality=60)
                img_base64 = base64.b64encode(img_io.getvalue()).decode("utf-8")
                mime_type = "image/webp"
            except Exception:
                # Fallback en base64 brut avec extension d'origine
                img_base64 = base64.b64encode(img_bytes).decode("utf-8")
                mime_type = base_image.get("ext", "png")
                mime_type = f"image/{mime_type}"

            images.append({
                "top": int(bbox.y0),
                "left": int(bbox.x0),
                "width": int(bbox.width),
                "height": int(bbox.height),
                "base64": f"data:{mime_type};base64,{img_base64}"
            })

        pages.append({
            "page_index": page_num + 1,
            "page_width": int(page.rect.width),
            "page_height": int(page.rect.height),
            "texts": texts,
            "rectangles": rects,
            "images": images
        })

    return {"pages": pages}


if __name__ == "__main__":
    pdf_file = "pdfs\d.pdf"  # Remplace par le chemin de ton PDF
    data = extract_pdf_to_json(pdf_file)
    print(json.dumps(data, indent=2, ensure_ascii=False))
