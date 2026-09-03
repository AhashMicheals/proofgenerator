import os
import io
import re
import zipfile
from typing import Dict, List, Tuple, Optional, Any
from PIL import Image, ImageDraw, ImageFont

# Disable PIL limit for high-resolution images (> 89 megapixels)
Image.MAX_IMAGE_PIXELS = None

SUPPORTED_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.webp', '.bmp')

def normalize_slug(s: str) -> str:
    """Normalize string by keeping only alphanumeric characters in lowercase."""
    if not s:
        return ""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def _register_single_photo(registry: Dict[str, Any], filename: str, file_bytes: bytes):
    """Helper to index a single photo under multiple fast-lookup keys."""
    if not filename or not file_bytes:
        return
        
    ext = os.path.splitext(filename)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        return
        
    clean_filename = os.path.basename(filename)
    base_name = os.path.splitext(clean_filename)[0]
    norm_base = normalize_slug(base_name)
    norm_full = normalize_slug(clean_filename)
    
    photo_info = {
        "filename": clean_filename,
        "bytes": file_bytes,
        "extension": ext,
        "base_name": base_name,
        "norm_base": norm_base
    }
    
    # Fast O(1) hash lookups
    registry[clean_filename] = photo_info
    registry[clean_filename.lower()] = photo_info
    registry[base_name] = photo_info
    registry[base_name.lower()] = photo_info
    if norm_full:
        registry[norm_full] = photo_info
    if norm_base:
        registry[norm_base] = photo_info

def build_photo_registry(photo_files: List[Any]) -> Dict[str, Any]:
    """
    Build a comprehensive photo lookup registry from:
    1. Uploaded individual image files
    2. Uploaded ZIP files containing 1000s of images (and subfolders)
    3. File system directory paths
    """
    registry = {}
    
    for item in photo_files:
        # 1. Directory path string
        if isinstance(item, str) and os.path.exists(item) and os.path.isdir(item):
            for root, _, files in os.walk(item):
                for f in files:
                    fpath = os.path.join(root, f)
                    try:
                        with open(fpath, "rb") as fp:
                            _register_single_photo(registry, f, fp.read())
                    except Exception:
                        continue
            continue

        # 2. ZIP file on disk
        if isinstance(item, str) and os.path.exists(item) and item.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(item, "r") as zf:
                    for zname in zf.namelist():
                        if zname.startswith("__MACOSX") or zname.endswith("/") or os.path.basename(zname).startswith("."):
                            continue
                        zext = os.path.splitext(zname)[1].lower()
                        if zext in SUPPORTED_EXTENSIONS:
                            try:
                                zbytes = zf.read(zname)
                                _register_single_photo(registry, os.path.basename(zname), zbytes)
                            except Exception:
                                continue
                continue
            except Exception:
                continue

        # 3. Streamlit UploadedFile or file-like object
        if hasattr(item, "name") and (hasattr(item, "read") or hasattr(item, "getvalue")):
            filename = item.name
            if filename.lower().endswith(".zip"):
                try:
                    if hasattr(item, "seek"):
                        item.seek(0)
                    with zipfile.ZipFile(item) as zf:
                        for zname in zf.namelist():
                            if zname.startswith("__MACOSX") or zname.endswith("/") or os.path.basename(zname).startswith("."):
                                continue
                            zext = os.path.splitext(zname)[1].lower()
                            if zext in SUPPORTED_EXTENSIONS:
                                try:
                                    zbytes = zf.read(zname)
                                    _register_single_photo(registry, os.path.basename(zname), zbytes)
                                except Exception:
                                    continue
                    continue
                except Exception:
                    pass

            # Non-zip uploaded image
            file_bytes = item.getvalue() if hasattr(item, "getvalue") else item.read()
            if filename and file_bytes:
                _register_single_photo(registry, filename, file_bytes)
            continue

        # 4. Single file on disk
        if isinstance(item, str) and os.path.exists(item) and os.path.isfile(item):
            filename = os.path.basename(item)
            try:
                with open(item, "rb") as f:
                    file_bytes = f.read()
                _register_single_photo(registry, filename, file_bytes)
            except Exception:
                continue
            continue

        # 5. Tuple (filename, bytes)
        if isinstance(item, tuple) and len(item) == 2:
            filename, file_bytes = item
            if filename.lower().endswith(".zip") and file_bytes:
                try:
                    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                        for zname in zf.namelist():
                            if zname.startswith("__MACOSX") or zname.endswith("/") or os.path.basename(zname).startswith("."):
                                continue
                            zext = os.path.splitext(zname)[1].lower()
                            if zext in SUPPORTED_EXTENSIONS:
                                try:
                                    zbytes = zf.read(zname)
                                    _register_single_photo(registry, os.path.basename(zname), zbytes)
                                except Exception:
                                    continue
                except Exception:
                    continue
            elif filename and file_bytes:
                _register_single_photo(registry, filename, file_bytes)
            
    return registry

def match_photo_for_staff(record: Dict[str, Any], registry: Dict[str, Any], key_cols: Optional[Dict[str, Optional[str]]] = None) -> Optional[Dict[str, Any]]:
    """
    High-performance dynamic photo matching algorithm (O(1) hash map priority):
    1. Check explicit photo_col if detected/configured.
    2. Check any column name containing 'photo', 'img', 'picture'.
    3. Check explicit id_col or title_col.
    4. Check all record field values against registry filenames and normalized slugs.
    """
    if not registry:
        return None
        
    key_cols = key_cols or {}
    photo_col = key_cols.get("photo_col")
    id_col = key_cols.get("id_col")
    title_col = key_cols.get("title_col")

    priority_vals = []
    
    # 1. Explicit Photo Column value
    if photo_col and photo_col in record:
        val = str(record[photo_col]).strip()
        if val:
            priority_vals.append(val)
            
    # 2. Check any column named PHOTO/IMAGE
    for col, val in record.items():
        if col not in ("_meta",) and any(kw in str(col).lower() for kw in ["photo", "image", "pic", "img"]):
            str_val = str(val).strip()
            if str_val and str_val not in priority_vals:
                priority_vals.append(str_val)

    # 3. Explicit ID & Title Column values
    for col_key in (id_col, title_col):
        if col_key and col_key in record:
            val = str(record[col_key]).strip()
            if val and val not in priority_vals:
                priority_vals.append(val)

    # Try fast exact/base/slug O(1) lookups for priority values
    for val in priority_vals:
        base_val = os.path.splitext(os.path.basename(val))[0]
        norm_val = normalize_slug(base_val)
        
        for key in (val, val.lower(), base_val, base_val.lower(), norm_val):
            if key in registry:
                return registry[key]

    # 4. Fallback: Check ALL non-internal field values in record via O(1) lookup
    for col, val in record.items():
        if col.startswith("_"):
            continue
        str_val = str(val).strip()
        if not str_val or str_val == "-":
            continue
            
        base_val = os.path.splitext(os.path.basename(str_val))[0]
        norm_val = normalize_slug(base_val)
        
        for key in (str_val, str_val.lower(), base_val, base_val.lower(), norm_val):
            if key and key in registry:
                return registry[key]

    return None

def process_and_crop_photo(photo_bytes: bytes, target_width: int = 280, target_height: int = 350) -> bytes:
    """
    Reads image bytes (handling huge high-res images), converts mode to RGB safely, crops to portrait ratio, resizes.
    Outputs optimized JPEG byte stream (~30KB) to ensure small Word file size even for 1000s of cards.
    """
    try:
        img = Image.open(io.BytesIO(photo_bytes))
        if img.mode != 'RGB':
            img = img.convert('RGB')
            
        width, height = img.size
        target_ratio = target_width / target_height
        current_ratio = width / height
        
        if current_ratio > target_ratio:
            new_width = int(height * target_ratio)
            left = (width - new_width) // 2
            top = 0
            right = left + new_width
            bottom = height
        else:
            new_height = int(width / target_ratio)
            left = 0
            top = (height - new_height) // 2
            right = width
            bottom = top + new_height
            
        img_cropped = img.crop((left, top, right, bottom))
        img_resized = img_cropped.resize((target_width, target_height), Image.Resampling.LANCZOS)
        
        out_buffer = io.BytesIO()
        img_resized.save(out_buffer, format='JPEG', quality=95)
        return out_buffer.getvalue()
    except Exception:
        return create_placeholder_photo(target_width, target_height)

def create_placeholder_photo(width: int = 280, height: int = 350) -> bytes:
    """
    Generates a clean, neutral 'PHOTO NOT FOUND' image.
    """
    img = Image.new('RGB', (width, height), color=(245, 247, 250))
    draw = ImageDraw.Draw(img)
    
    # Outer border
    draw.rectangle([4, 4, width - 5, height - 5], outline=(180, 185, 195), width=2)
    
    # Avatar silhouette
    center_x = width // 2
    head_y = int(height * 0.38)
    head_r = int(width * 0.18)
    draw.ellipse([center_x - head_r, head_y - head_r, center_x + head_r, head_y + head_r], fill=(205, 210, 220))
    
    shoulder_y = int(height * 0.72)
    shoulder_rx = int(width * 0.32)
    shoulder_ry = int(height * 0.22)
    draw.ellipse([center_x - shoulder_rx, shoulder_y - shoulder_ry, center_x + shoulder_rx, shoulder_y + shoulder_ry], fill=(205, 210, 220))
    
    # Red Banner for missing photo text
    draw.rectangle([8, height - 52, width - 8, height - 16], fill=(220, 53, 69))
    text = "PHOTO NOT FOUND"
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None
        
    if font:
        bbox = draw.textbbox((0, 0), text, font=font)
        t_w = bbox[2] - bbox[0]
        t_h = bbox[3] - bbox[1]
        draw.text((center_x - t_w // 2, height - 34 - t_h // 2), text, fill=(255, 255, 255), font=font)
    else:
        draw.text((center_x - 45, height - 38), text, fill=(255, 255, 255))
        
    out_buffer = io.BytesIO()
    img.save(out_buffer, format='JPEG', quality=90)
    return out_buffer.getvalue()
