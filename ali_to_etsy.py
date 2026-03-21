#!/usr/bin/env python3
"""
AliExpress CSV to Etsy (SKUpid) Bulk Upload Converter
======================================================
Takes the CSV output from the AliExpress scraper and generates
a SKUpid-compatible CSV for bulk upload to Etsy.

Usage:
    python3 ali_to_etsy.py <scraped_csv>
    python3 ali_to_etsy.py  (auto-finds CSV in current folder)

Requirements:
    pip3 install requests Pillow --break-system-packages
"""

import csv
import json
import os
import re
import sys
import base64
import logging
from io import BytesIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import requests as http_requests
except ImportError:
    print("[ERROR] requests not installed. Run: pip3 install requests --break-system-packages")
    sys.exit(1)

try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False
    print("[WARN] Pillow not installed. Images won't be converted. Run: pip3 install Pillow --break-system-packages")

log = logging.getLogger("etsy")
logging.basicConfig(level=logging.INFO, format="%(message)s")

# ============================================================
# CONFIGURATION
# ============================================================
TAXONOMY_ID = 6450                      # Miniatures (Craft Supplies > Doll & Model Supplies)
SHIPPING_TEMPLATE_ID = 302463300039     # Your "Usual" delivery profile
PROCESSING_MIN = 4                      # Min processing days
PROCESSING_MAX = 7                     # Max processing days
QUANTITY = 5                            # Stock per listing
IS_SUPPLY = 1                           # 1 = craft supply (keeps within Etsy policy)
WHO_MADE = "someone_else"
WHEN_MADE = "2020_2026"
IS_CUSTOMIZABLE = 0
MATERIALS = "resin"

# ---- PRICING ----
TARGET_PROFIT_MARGIN = 0.30
ETSY_LISTING_FEE = 0.16                # £0.16 per listing
ETSY_TRANSACTION_FEE = 0.065           # 6.5%
ETSY_PAYMENT_FEE_PCT = 0.04            # 4% + 20p
ETSY_PAYMENT_FEE_FIXED = 0.20          # 20p
USD_TO_GBP = 0.79
ALI_SHIPPING_ESTIMATE = 2.00
MIN_SELL_PRICE = 5.99
MIN_PROFIT = 6.0

# ---- IMAGE REHOSTING ----
IMGBB_API_KEY = "dd9a3b6ab5cabf1a45a24736ffe29e42"
REHOST_WORKERS = 20
MAX_IMAGES_PER_LISTING = 10

# ---- RESIN FILTER ----
RESIN_INCLUDE = [
    "resin", "model kit", "model figure", "figure kit", "garage kit",
    "gk kit", "unpainted", "unassembled", "1/6 scale", "1/8 scale",
    "1/10 scale", "1/12 scale", "1/24 scale", "1/35 scale",
    "statue kit", "bust kit", "diorama", "miniature figure",
    "scale model", "resin cast", "resin figure", "resin statue",
    "moc", "building block", "micro block", "brick set", "brick model",
    "nano block", "diamond block", "mini block", "architecture model",
    "military model", "tank model", "ship model", "airplane model",
    "car model kit", "gundam", "mecha", "robot model",
]

RESIN_EXCLUDE = [
    "phone case", "screen protector", "earphone", "headphone",
    "charger", "cable", "adapter", "usb", "bluetooth",
    "clothing", "shirt", "dress", "pants", "shoe", "sock",
    "food", "snack", "drink", "supplement", "vitamin",
    "cosmetic", "makeup", "skincare", "perfume", "shampoo",
    "pet food", "dog food", "cat food",
    "sticker", "decal only", "poster", "wall art",
    "silicone mold", "silicone mould", "candle mold",
    "jewelry mold", "epoxy mold", "soap mold",
    "resin art supply", "resin pigment", "resin dye",
    "jeep", "ford", "toyota", "bmw", "mercedes", "audi",
    "ferrari", "lamborghini", "porsche", "tesla", "honda",
    "marvel", "disney", "star wars", "pokemon", "transformers",
    "warhammer", "games workshop", "bandai", "kotobukiya", "hasbro",
    "funko", "lego", "nike", "adidas", "supreme",
]


# ============================================================
# IMAGE REHOSTING (AliExpress -> imgbb/freeimage/imgur)
# ============================================================

def _download_and_convert(img_url):
    """Download image from AliExpress, convert to JPEG bytes."""
    resp = http_requests.get(img_url, headers={
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept": "image/*",
        "Referer": "https://www.aliexpress.com/",
    }, timeout=15)
    if resp.status_code != 200:
        return None
    if HAS_PILLOW:
        try:
            img = Image.open(BytesIO(resp.content))
            if img.mode in ('RGBA', 'P', 'LA'):
                bg = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                bg.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
                img = bg
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            jpeg_buffer = BytesIO()
            img.save(jpeg_buffer, format='JPEG', quality=92)
            return jpeg_buffer.getvalue()
        except Exception:
            return resp.content
    return resp.content


def _upload_to_imgbb(jpeg_bytes):
    """Upload JPEG bytes to imgbb. Returns URL or None."""
    try:
        b64 = base64.b64encode(jpeg_bytes).decode("utf-8")
        resp = http_requests.post(
            "https://api.imgbb.com/1/upload",
            data={"key": IMGBB_API_KEY, "image": b64},
            timeout=30,
        )
        if resp.status_code == 200:
            data = resp.json()
            img_data = data.get("data", {})
            url = (img_data.get("image", {}).get("url", "")
                   or img_data.get("display_url", "")
                   or img_data.get("url", ""))
            if url:
                return url
    except Exception:
        pass
    return None


def _upload_to_freeimage(jpeg_bytes):
    """Upload JPEG bytes to freeimage.host. Returns URL or None."""
    try:
        b64 = base64.b64encode(jpeg_bytes).decode("utf-8")
        resp = http_requests.post(
            "https://freeimage.host/api/1/upload",
            data={"key": "6d207e02198a847aa98d0a2a901485a5", "source": b64, "format": "json"},
            timeout=30,
        )
        if resp.status_code == 200:
            data = resp.json()
            url = data.get("image", {}).get("url", "")
            if url:
                return url
    except Exception:
        pass
    return None


def _upload_to_imgur(jpeg_bytes):
    """Upload JPEG bytes to Imgur. Returns URL or None."""
    try:
        b64 = base64.b64encode(jpeg_bytes).decode("utf-8")
        resp = http_requests.post(
            "https://api.imgur.com/3/image",
            headers={"Authorization": "Client-ID 546c25a59c58ad7"},
            data={"image": b64, "type": "base64"},
            timeout=30,
        )
        if resp.status_code == 200:
            data = resp.json()
            link = data.get("data", {}).get("link", "")
            if link:
                return link.replace("http://", "https://")
    except Exception:
        pass
    return None


def rehost_image(img_url):
    """Download image from AliExpress, rehost to imgbb/freeimage/imgur."""
    if not img_url:
        return None
    if img_url.startswith("//"):
        img_url = "https:" + img_url
    # Strip AliExpress resize suffixes for full-size
    img_url = re.sub(r'_\d+x\d+[^.]*\.', '.', img_url)
    img_url = re.sub(r'\.(jpg|png|jpeg)_\d+x\d+[^.]*', r'.\1', img_url, flags=re.IGNORECASE)

    jpeg_bytes = None
    try:
        jpeg_bytes = _download_and_convert(img_url)
    except Exception:
        pass
    if not jpeg_bytes:
        return None

    # Detect promotional banner images (solid color backgrounds like "Sale", "Choice")
    if HAS_PILLOW and _is_promo_image(jpeg_bytes):
        return "PROMO"

    # Try imgbb first
    url = _upload_to_imgbb(jpeg_bytes)
    if url:
        return url
    # Try freeimage
    url = _upload_to_freeimage(jpeg_bytes)
    if url:
        return url
    # Try imgur
    url = _upload_to_imgur(jpeg_bytes)
    if url:
        return url
    return None


def _is_promo_image(jpeg_bytes):
    """Detect AliExpress promotional banner images (Sale, Choice, etc).
    These have large areas of solid color (red, yellow, orange, etc)."""
    try:
        img = Image.open(BytesIO(jpeg_bytes)).convert('RGB')
        img_small = img.resize((50, 50))  # Downsample for speed
        pixels = list(img_small.getdata())
        total = len(pixels)

        # Count pixels that are very saturated solid colors (promo banners)
        promo_count = 0
        for r, g, b in pixels:
            # Solid red (Sale banners)
            if r > 180 and g < 80 and b < 80:
                promo_count += 1
            # Solid yellow/gold (Choice banners)
            elif r > 200 and g > 180 and b < 80:
                promo_count += 1
            # Solid orange (promo banners)
            elif r > 200 and 80 < g < 160 and b < 60:
                promo_count += 1
            # Solid bright green (promo banners)
            elif g > 200 and r < 80 and b < 80:
                promo_count += 1
            # Solid blue (promo banners)
            elif b > 200 and r < 80 and g < 80:
                promo_count += 1

        ratio = promo_count / total
        if ratio > 0.35:  # More than 35% solid promo color
            return True
    except Exception:
        pass
    return False


def rehost_all_images(resin_products):
    """Rehost all images for all products in parallel."""
    tasks = []
    for pi, row in enumerate(resin_products):
        img_str = row.get("product_images", "")
        all_urls = []
        if img_str:
            all_urls = [clean_image_url(u.strip()) for u in img_str.split("|") if u.strip()]
        if not all_urls:
            single = row.get("product_image", "")
            if single:
                all_urls = [clean_image_url(single)]
        for ii, url in enumerate(all_urls[:MAX_IMAGES_PER_LISTING]):
            tasks.append((pi, ii, url))

    print(f"[INFO] Rehosting {len(tasks)} images across {len(resin_products)} products ({REHOST_WORKERS} workers)...")

    results = {}
    done = 0
    failed = 0

    with ThreadPoolExecutor(max_workers=REHOST_WORKERS) as pool:
        future_map = {pool.submit(rehost_image, t[2]): (t[0], t[1], t[2]) for t in tasks}
        for future in as_completed(future_map):
            pi, ii, orig_url = future_map[future]
            done += 1
            try:
                new_url = future.result()
                if new_url and new_url != "PROMO":
                    results[(pi, ii)] = new_url
                elif new_url == "PROMO":
                    pass  # Skip promotional images silently
                else:
                    failed += 1
            except Exception:
                failed += 1
            if done % 50 == 0 or done == len(tasks):
                print(f"[INFO]   {done}/{len(tasks)} images processed ({failed} failed)")

    for pi, row in enumerate(resin_products):
        rehosted = []
        for ii in range(MAX_IMAGES_PER_LISTING):
            if (pi, ii) in results:
                rehosted.append(results[(pi, ii)])
        row["_rehosted_images"] = rehosted

    print(f"[INFO] Rehosted {len(results)} images, {failed} failed")
    return resin_products


def is_resin_model(title):
    t = title.lower()
    for ex in RESIN_EXCLUDE:
        if ex in t:
            return False
    for inc in RESIN_INCLUDE:
        if inc in t:
            return True
    return False


def parse_price(price_str):
    if not price_str or price_str == "N/A":
        return None
    cleaned = re.sub(r'[^\d.,]', '', price_str)
    if ',' in cleaned and '.' not in cleaned:
        cleaned = cleaned.replace(',', '.')
    elif ',' in cleaned and '.' in cleaned:
        cleaned = cleaned.replace(',', '')
    try:
        return float(cleaned)
    except ValueError:
        return None


def ali_to_etsy_price(price_usd):
    """Calculate sell price for 30% margin / £6 min profit after Etsy fees."""
    if not price_usd:
        return 12.99

    cost_gbp = price_usd * USD_TO_GBP
    total_cost = cost_gbp + ALI_SHIPPING_ESTIMATE

    # sell - (sell * transaction_fee) - (sell * payment_pct) - payment_fixed - listing_fee - cost = margin * sell
    # sell * (1 - transaction_fee - payment_pct - margin) = cost + payment_fixed + listing_fee
    denominator = 1 - ETSY_TRANSACTION_FEE - ETSY_PAYMENT_FEE_PCT - TARGET_PROFIT_MARGIN
    if denominator <= 0:
        return 12.99

    sell_price = (total_cost + ETSY_PAYMENT_FEE_FIXED + ETSY_LISTING_FEE) / denominator
    sell_price = round(sell_price, 2)

    # Check minimum £6 profit
    fees = (sell_price * ETSY_TRANSACTION_FEE) + (sell_price * ETSY_PAYMENT_FEE_PCT) + ETSY_PAYMENT_FEE_FIXED + ETSY_LISTING_FEE
    actual_profit = sell_price - fees - total_cost
    if actual_profit < MIN_PROFIT:
        sell_price = (MIN_PROFIT + ETSY_PAYMENT_FEE_FIXED + ETSY_LISTING_FEE + total_cost) / (1 - ETSY_TRANSACTION_FEE - ETSY_PAYMENT_FEE_PCT)
        sell_price = round(sell_price, 2)

    if sell_price < MIN_SELL_PRICE:
        sell_price = MIN_SELL_PRICE

    return sell_price


def clean_title(title):
    """Clean title for Etsy — max 140 chars, no special chars."""
    if not title:
        return "Resin Model Kit"
    title = re.sub(r'[^\w\s\-\.,&\'/()\[\]]', ' ', title)
    # Remove commas — they break SKUpid CSV parsing
    title = title.replace(',', ' -')
    # Remove prohibited phrases
    for phrase in [r'free\s*shipping', r'hot\s*sale', r'wholesale', r'dropship\w*',
                   r'aliexpress', r'cheap', r'from\s*china', r'new\s*arrival']:
        title = re.sub(r'(?i)\b' + phrase + r'\b', '', title)
    title = re.sub(r'\s+', ' ', title).strip()
    if len(title) > 140:
        title = title[:137] + "..."
    if not title or len(title) < 3:
        title = "Resin Model Kit"
    return title


def make_description(title):
    """Generate detailed Etsy product description with SEO keywords."""
    clean = clean_title(title)
    t = title.lower()

    # Detect product attributes from title
    scale = ""
    scale_match = re.search(r'1/(\d+)', title)
    if scale_match:
        scale = "1/" + scale_match.group(1)

    size_mm = ""
    mm_match = re.search(r'(\d+)\s*mm', t)
    if mm_match:
        size_mm = mm_match.group(1) + "mm"

    # Detect theme
    theme = "Fantasy"
    if any(w in t for w in ["wwii", "ww2", "world war", "military", "soldier", "infantry", "tank crew", "tanker"]):
        theme = "Military / WWII"
    elif any(w in t for w in ["ww1 ", "wwi ", "world war i "]):
        theme = "Military / WWI"
    elif any(w in t for w in ["vietnam", "korean war", "modern"]):
        theme = "Modern Military"
    elif any(w in t for w in ["medieval", "knight", "crusad", "viking", "saxon", "roman", "ancient", "historical"]):
        theme = "Historical"
    elif any(w in t for w in ["fantasy", "dragon", "orc", "elf", "demon", "warrior", "monster", "minotaur"]):
        theme = "Fantasy / Tabletop"
    elif any(w in t for w in ["sci-fi", "science fiction", "mecha", "robot", "space", "cyber"]):
        theme = "Science Fiction"
    elif any(w in t for w in ["pirate", "cowboy", "western"]):
        theme = "Adventure"
    elif any(w in t for w in ["girl", "female", "woman", "lady", "maiden"]):
        theme = "Character / Figure"
    elif any(w in t for w in ["civilian", "driver", "mechanic", "worker"]):
        theme = "Civilian / Diorama"

    # Detect type
    product_type = "Figure"
    if "bust" in t:
        product_type = "Bust"
    elif "diorama" in t:
        product_type = "Diorama Set"
    elif any(w in t for w in ["head", "heads"]):
        product_type = "Head Set / Conversion"
    elif any(w in t for w in ["accessori", "stowage", "equipment", "boots", "helmet", "tarp", "tent", "bag"]):
        product_type = "Accessories / Stowage"
    elif any(w in t for w in ["animal", "horse", "cow", "dog"]):
        product_type = "Animal Figure"

    is_unpainted = "unpainted" in t or "unassembled" in t
    is_gk = "gk" in t or "garage kit" in t

    NL = "&#13;&#10;"
    NL2 = NL + NL

    desc = clean + NL2

    # Opening hook
    desc += "Bring your modelling projects to life with this detailed resin " + product_type.lower() + "."
    if scale:
        desc += " " + scale + " scale."
    if size_mm:
        desc += " Approximately " + size_mm + " in height."
    desc += NL2

    # What's included
    desc += "WHAT'S INCLUDED" + NL
    desc += "- Resin " + product_type.lower() + " kit"
    if is_gk:
        desc += " (Garage Kit / GK)"
    desc += NL
    if is_unpainted:
        desc += "- Unassembled and unpainted - ready for your creative touch" + NL
    else:
        desc += "- May require assembly and painting" + NL
    desc += "- Carefully cast in high quality resin" + NL
    if any(w in t for w in ["base", "with base", "including base"]):
        desc += "- Display base included" + NL
    desc += NL

    # Specs
    desc += "SPECIFICATIONS" + NL
    desc += "- Material: Resin" + NL
    if scale:
        desc += "- Scale: " + scale + NL
    if size_mm:
        desc += "- Size: Approximately " + size_mm + NL
    desc += "- Theme: " + theme + NL
    desc += "- Type: " + product_type + NL
    desc += "- Condition: New, unbuilt kit" + NL
    desc += NL

    # Who it's for
    desc += "PERFECT FOR" + NL
    desc += "- Scale model builders and painters" + NL
    desc += "- Wargaming and tabletop RPG enthusiasts" + NL
    desc += "- Military history collectors" + NL
    desc += "- Diorama creators" + NL
    desc += "- Gift for hobby enthusiasts" + NL
    desc += NL

    # Tips
    desc += "MODELLING TIPS" + NL
    desc += "- Wash parts in warm soapy water before priming" + NL
    desc += "- Use superglue (cyanoacrylate) for assembly" + NL
    desc += "- Prime with grey or white primer before painting" + NL
    desc += "- Acrylic or enamel paints recommended" + NL
    desc += NL

    # Shipping
    desc += "SHIPPING AND PROCESSING" + NL
    desc += "- Processing time: 4-7 business days" + NL
    desc += "- Carefully packaged to prevent damage during transit" + NL
    desc += "- Tracking provided on all orders" + NL
    desc += NL

    desc += "Please examine all photos carefully before purchasing. Photos show the kit assembled and painted as a reference - you will receive the unassembled, unpainted kit." + NL
    desc += NL
    desc += "Questions? Send us a message and we will be happy to help!" + NL

    if len(desc) > 5000:
        desc = desc[:4997] + "..."
    return desc


def make_tags(title):
    """Generate up to 13 highly relevant Etsy SEO tags from title."""
    t = title.lower()
    tags = []
    seen = set()

    def add(tag):
        tag = tag.strip()[:20]
        if tag.lower() not in seen and len(tags) < 13 and len(tag) >= 2:
            seen.add(tag.lower())
            tags.append(tag)

    # 1. Extract scale tags
    scale_match = re.search(r'1/(\d+)', title)
    if scale_match:
        num = scale_match.group(1)
        add("1 " + num + " scale model")
        add("1 " + num + " resin figure")
        add("1 " + num + " miniature")

    # 2. Detect and add theme-specific tags
    if any(w in t for w in ["wwii", "ww2", "world war ii"]):
        add("ww2 model kit")
        add("wwii miniature")
        add("ww2 soldier")
    if any(w in t for w in ["ww1 ", "wwi ", "world war i "]):
        add("ww1 model kit")
        add("wwi figure")
    if any(w in t for w in ["military", "soldier", "infantry", "army"]):
        add("military model")
        add("soldier figure")
    if any(w in t for w in ["tank", "tanker", "tank crew", "panzer"]):
        add("tank crew model")
        add("tank model kit")
    if any(w in t for w in ["fantasy", "dragon", "orc", "elf", "demon"]):
        add("fantasy miniature")
        add("tabletop figure")
        add("dnd miniature")
    if any(w in t for w in ["knight", "medieval", "crusad"]):
        add("medieval knight")
        add("historical figure")
    if any(w in t for w in ["viking", "norse"]):
        add("viking model")
        add("norse warrior")
    if any(w in t for w in ["roman", "centurion", "gladiator"]):
        add("roman figure")
        add("ancient warrior")
    if any(w in t for w in ["pirate"]):
        add("pirate model")
        add("pirate figure")
    if any(w in t for w in ["sci-fi", "science fiction", "mecha", "robot", "space"]):
        add("sci-fi model")
        add("mecha figure")
    if any(w in t for w in ["bust"]):
        add("resin bust")
        add("model bust kit")
    if any(w in t for w in ["diorama"]):
        add("diorama supplies")
        add("diorama figure")
    if any(w in t for w in ["head", "heads"]):
        add("conversion bits")
        add("model heads")
    if any(w in t for w in ["pilot", "air force", "airborne"]):
        add("pilot figure")
        add("airborne model")
    if any(w in t for w in ["british", "uk"]):
        add("british soldier")
    if any(w in t for w in ["american", "us ", "u.s", "usaf"]):
        add("us military model")
    if any(w in t for w in ["soviet", "russian", "russia"]):
        add("soviet model")
    if any(w in t for w in ["german", "panzer", "luftwaffe", "dak"]):
        add("german ww2 model")

    # 3. Extract unique meaningful words from title as tags
    words = re.findall(r'\b[a-zA-Z]{4,15}\b', t)
    stopwords = {'the', 'and', 'for', 'with', 'kit', 'model', 'resin', 'figure',
                 'unassembled', 'unpainted', 'theme', 'bust', 'military', 'themes',
                 'soldier', 'that', 'this', 'from', 'have', 'been', 'were', 'will',
                 'would', 'could', 'should', 'about', 'their', 'which', 'there'}
    for word in words:
        if word not in stopwords:
            add(word)

    # 4. Always include core tags if space
    core = ["resin model kit", "unpainted kit", "hobby supplies", "craft supplies",
            "miniature figure", "scale model", "model painting", "resin craft"]
    for tag in core:
        add(tag)

    # Strip invalid characters from tags (Etsy doesn't allow / \ : etc)
    clean_tags = []
    for tag in tags[:13]:
        tag = re.sub(r'[/\\:;!@#$%^&*()=+\[\]{}|<>]', ' ', tag)
        tag = re.sub(r'\s+', ' ', tag).strip()[:20]
        if tag:
            clean_tags.append(tag)
    return ",".join(clean_tags)


def clean_image_url(url):
    """Clean AliExpress image URL."""
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    # Strip size suffixes
    url = re.sub(r'_\d+x\d+[^.]*\.', '.', url)
    url = re.sub(r'\.(jpg|png|jpeg)_\d+x\d+[^.]*', r'.\1', url, flags=re.IGNORECASE)
    # Remove .avif
    avif_match = re.match(r'(https?://.*?\.(jpg|png|jpeg))', url, re.IGNORECASE)
    if avif_match:
        url = avif_match.group(1)
    return url


def main():
    print("=" * 60)
    print("  AliExpress CSV -> Etsy (SKUpid) Bulk Upload")
    print("=" * 60)

    # Find CSV
    if len(sys.argv) >= 2:
        csv_path = sys.argv[1]
    else:
        csvs = [f for f in os.listdir(".") if f.endswith(".csv") and "aliexpress" in f.lower()]
        if not csvs:
            csvs = [f for f in os.listdir(".") if f.endswith(".csv") and "etsy" not in f.lower()]
        if not csvs:
            print("[ERROR] No CSV found. Usage: python3 ali_to_etsy.py <scraped_csv>")
            sys.exit(1)
        if len(csvs) > 1:
            print("Multiple CSVs found:")
            for i, f in enumerate(csvs):
                print("  " + str(i + 1) + ". " + f)
            choice = input("Select: ").strip()
            try:
                csv_path = csvs[int(choice) - 1]
            except (ValueError, IndexError):
                sys.exit(1)
        else:
            csv_path = csvs[0]

    if not os.path.exists(csv_path):
        print("[ERROR] File not found: " + csv_path)
        sys.exit(1)

    # Load CSV
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    print("[INFO] Loaded " + str(len(rows)) + " products")

    # Deduplicate
    seen_ids = set()
    seen_titles = set()
    unique = []
    for row in rows:
        pid = row.get("id", "").strip()
        title = row.get("product_title", "").strip().lower()
        if pid and pid in seen_ids:
            continue
        if title and title in seen_titles:
            continue
        if pid:
            seen_ids.add(pid)
        if title:
            seen_titles.add(title)
        unique.append(row)
    print("[INFO] Deduplicated: " + str(len(unique)) + " unique")

    # Filter resin models
    resin = [r for r in unique if is_resin_model(r.get("product_title", ""))]
    print("[INFO] Resin models: " + str(len(resin)))

    if not resin:
        print("[ERROR] No resin models found.")
        sys.exit(1)

    # Rehost images from AliExpress to imgbb
    resin = rehost_all_images(resin)

    # Build Etsy CSV
    etsy_rows = []
    skipped_no_images = 0
    for row in resin:
        title = clean_title(row.get("product_title", ""))
        price_str = row.get("product_price", "")
        price_usd = parse_price(price_str)
        sell_price = ali_to_etsy_price(price_usd)

        # Use rehosted images
        all_images = row.get("_rehosted_images", [])
        if not all_images:
            skipped_no_images += 1
            continue

        # Limit to 10 images, comma-separated
        images_str = ",".join(all_images[:10])

        etsy_row = {
            "title": title,
            "description": make_description(title),
            "quantity": QUANTITY,
            "price": f"{sell_price + (0.01 if f'{sell_price:.2f}'.endswith('0') else 0):.2f}",
            "is_supply": IS_SUPPLY,
            "who_made": WHO_MADE,
            "is_customizable": IS_CUSTOMIZABLE,
            "when_made": WHEN_MADE,
            "tags": make_tags(title),
            "processing_min": PROCESSING_MIN,
            "processing_max": PROCESSING_MAX,
            "shipping_template_id": SHIPPING_TEMPLATE_ID,
            "readiness_state_id": 1473548284400,
            "shop_section_id": "",
            "images": images_str,
            "variations": "",
            "materials": MATERIALS,
            "taxonomy_id": TAXONOMY_ID,
            "title_de": "",
            "description_de": "",
            "tags_de": "",
            "renew": "",
            "is_personalizable": 0,
            "personalization_is_required": "",
            "personalization_char_count_max": "",
            "personalization_instructions": "",
            "production_partner_ids": "",
        }
        etsy_rows.append(etsy_row)

    # Write output CSV
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = "etsy_upload_" + ts + ".csv"

    fieldnames = [
        "title", "description", "quantity", "price", "is_supply", "who_made",
        "is_customizable", "when_made", "tags", "processing_min", "processing_max",
        "shipping_template_id", "readiness_state_id", "shop_section_id", "images",
        "variations", "materials", "taxonomy_id", "title_de", "description_de",
        "tags_de", "renew", "is_personalizable", "personalization_is_required",
        "personalization_char_count_max", "personalization_instructions",
        "production_partner_ids",
    ]

    with open(output_name, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore", quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        w.writerows(etsy_rows)

    print()
    print("=" * 60)
    print("  DONE! " + str(len(etsy_rows)) + " products -> " + output_name)
    print("=" * 60)
    print()
    print("NEXT STEPS:")
    print("  1. Go to SKUpid > Import")
    print("  2. Upload " + output_name)
    print("  3. Select 'Basic listing importer'")
    print("  4. SKUpid creates listings as DRAFTS")
    print("  5. Review in Etsy Shop Manager > Listings")
    print("  6. Activate the listings you want to publish")
    print()
    print("NOTE: Etsy charges £0.16 per listing when published.")
    print("      " + str(len(etsy_rows)) + " listings = £" + str(round(len(etsy_rows) * 0.16, 2)) + " in listing fees")
    if skipped_no_images:
        print("      " + str(skipped_no_images) + " products skipped (image rehosting failed)")


if __name__ == "__main__":
    main()
