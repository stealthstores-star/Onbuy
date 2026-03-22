#!/usr/bin/env python3
"""
AliExpress CSV -> OnBuy Product Create Template Converter
==========================================================
Takes an AliExpress scrape CSV, assigns GTINs, rehosts images,
generates descriptions, calculates pricing (min £7.50 profit or
35% margin after OnBuy fees, whichever is higher), and outputs
an OnBuy-compatible XLSX ready for upload.

Usage:
    python3 ali_to_onbuy.py
    python3 ali_to_onbuy.py <scraped_csv>

Requirements:
    pip3 install requests Pillow openpyxl --break-system-packages
"""

import csv
import os
import re
import sys
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# openpyxl optional — only needed if you want XLSX output
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# Import image rehosting functions from ali_to_etsy.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ali_to_etsy import (
    rehost_image, clean_image_url, parse_price,
    _download_and_convert, _upload_to_imgbb, _upload_to_freeimage, _upload_to_imgur,
    HAS_PILLOW, _is_promo_image,
)

log = logging.getLogger("onbuy")
logging.basicConfig(level=logging.INFO, format="%(message)s")

# ============================================================
# CONFIGURATION
# ============================================================

# ---- FILES ----
ALIEXPRESS_CSV = "aliexpress_scrape_2026-03-22_1346.csv"
GTIN_CSV = "GTIN13-506112295-EN-0226-B457-4BEF.csv"
ONBUY_TEMPLATE = "OnBuy_Product_Create_Template.xlsx"

# ---- PRICING ----
ONBUY_COMMISSION = 0.09          # 9% category fee (Toys & Hobbies typical)
ALI_SHIPPING_GBP = 3.52          # Flat AliExpress shipping
TARGET_MARGIN = 0.35             # 35% profit margin
MIN_PROFIT = 7.50                # £7.50 minimum profit
MIN_SELL_PRICE = 9.99            # Floor price

# ---- LISTING DEFAULTS ----
CONDITION = "New"
BRAND = "Unbranded"
CATEGORY = "Toys & Games > Action Figures & Collectibles > Model Kits"
STOCK = 5
HANDLING_DAYS = 5
FREE_RETURNS = "No"

# ---- IMAGE REHOSTING ----
REHOST_WORKERS = 20
MAX_IMAGES_PER_LISTING = 10      # 1 default + up to 10 additional on OnBuy

# ---- LIMITS ----
MAX_PRODUCTS = 1000              # Match GTIN count


# ============================================================
# PRICING
# ============================================================

def calculate_onbuy_price(cost_gbp):
    """Calculate sell price: max of (35% margin, £7.50 profit) after OnBuy fees.

    OnBuy takes commission% of sell price.
    Profit = sell_price * (1 - commission) - total_cost
    Margin = profit / sell_price

    For 35% margin on sell price:
        sell * (1 - commission) - cost = 0.35 * sell
        sell = cost / (1 - commission - 0.35)

    For £7.50 min profit:
        sell * (1 - commission) - cost = 7.50
        sell = (cost + 7.50) / (1 - commission)

    We pick whichever gives the HIGHER sell price to satisfy both.
    """
    if not cost_gbp or cost_gbp <= 0:
        return MIN_SELL_PRICE

    total_cost = cost_gbp + ALI_SHIPPING_GBP

    # Price for 35% margin
    denom_margin = 1 - ONBUY_COMMISSION - TARGET_MARGIN
    if denom_margin <= 0:
        price_for_margin = 999.99
    else:
        price_for_margin = total_cost / denom_margin

    # Price for £7.50 min profit
    denom_profit = 1 - ONBUY_COMMISSION
    price_for_profit = (total_cost + MIN_PROFIT) / denom_profit

    # Whichever comes first = whichever gives the higher price
    sell_price = max(price_for_margin, price_for_profit)

    # Round to .99 pricing
    sell_price = round(sell_price)
    if sell_price > 1:
        sell_price = sell_price - 0.01  # e.g. 25 -> 24.99

    if sell_price < MIN_SELL_PRICE:
        sell_price = MIN_SELL_PRICE

    return round(sell_price, 2)


# ============================================================
# DESCRIPTION GENERATOR
# ============================================================

def clean_title_onbuy(title):
    """Clean title for OnBuy — max 150 chars, no AliExpress junk."""
    if not title:
        return "Scale Resin Model Kit"
    # Remove AliExpress promotional words
    for phrase in [r'free\s*shipping', r'hot\s*sale', r'wholesale', r'dropship\w*',
                   r'aliexpress', r'cheap', r'from\s*china', r'new\s*arrival',
                   r'best\s*price', r'top\s*quality', r'high\s*quality',
                   r'in\s*stock', r'fast\s*delivery', r'big\s*sale']:
        title = re.sub(r'(?i)\b' + phrase + r'\b', '', title)
    # Clean special chars but keep useful ones
    title = re.sub(r'[^\w\s\-\.,&\'/()\[\]]', ' ', title)
    title = re.sub(r'\s+', ' ', title).strip()
    # Capitalise nicely
    title = title.title()
    if len(title) > 150:
        title = title[:147] + "..."
    if not title or len(title) < 3:
        title = "Scale Resin Model Kit"
    return title


def _detect_attributes(title):
    """Extract product attributes from the title text."""
    t = title.lower()

    # Scale
    scale = ""
    scale_match = re.search(r'1[/:](\d+)', title)
    if scale_match:
        scale = "1/" + scale_match.group(1)

    # Size
    size_mm = ""
    mm_match = re.search(r'(\d+)\s*mm', t)
    if mm_match:
        size_mm = mm_match.group(1) + "mm"

    # Theme
    theme = "Fantasy"
    if any(w in t for w in ["wwii", "ww2", "world war ii", "world war 2"]):
        theme = "World War II"
    elif any(w in t for w in ["ww1", "wwi", "world war i", "world war 1"]):
        theme = "World War I"
    elif any(w in t for w in ["vietnam", "korean war", "modern military", "modern soldier"]):
        theme = "Modern Military"
    elif any(w in t for w in ["military", "soldier", "infantry", "army", "marine", "paratrooper"]):
        theme = "Military"
    elif any(w in t for w in ["medieval", "knight", "crusad", "viking", "saxon"]):
        theme = "Medieval"
    elif any(w in t for w in ["roman", "ancient", "greek", "egyptian", "spartan"]):
        theme = "Ancient History"
    elif any(w in t for w in ["civil war", "napoleon", "colonial", "historical"]):
        theme = "Historical"
    elif any(w in t for w in ["fantasy", "dragon", "orc", "elf", "demon", "wizard"]):
        theme = "Fantasy"
    elif any(w in t for w in ["sci-fi", "science fiction", "mecha", "robot", "space", "cyber"]):
        theme = "Science Fiction"
    elif any(w in t for w in ["pirate", "cowboy", "western"]):
        theme = "Adventure"
    elif any(w in t for w in ["pilot", "helicopter", "aircrew", "air force"]):
        theme = "Aviation"
    elif any(w in t for w in ["navy", "sailor", "naval", "submarine"]):
        theme = "Naval"
    elif any(w in t for w in ["civilian", "worker", "mechanic"]):
        theme = "Civilian"

    # Type
    product_type = "Figure"
    if "bust" in t:
        product_type = "Bust"
    elif "diorama" in t:
        product_type = "Diorama Set"
    elif any(w in t for w in ["head ", "heads "]):
        product_type = "Head Set"
    elif any(w in t for w in ["accessori", "stowage", "equipment"]):
        product_type = "Accessories"
    elif any(w in t for w in ["animal", "horse", "cow", "dog", "mule"]):
        product_type = "Animal Figure"
    elif any(w in t for w in ["vehicle", "tank", "jeep", "truck", "car "]):
        product_type = "Vehicle Kit"
    elif any(w in t for w in ["building", "ruin", "house", "church"]):
        product_type = "Building / Terrain"

    # Count of figures
    fig_count = ""
    count_match = re.search(r'(\d+)\s*(?:men|man|figures?|soldiers?|people|persons?|pcs)', t)
    if count_match:
        fig_count = count_match.group(1)

    is_unpainted = "unpainted" in t or "unassembled" in t or "unbuilt" in t
    has_base = any(w in t for w in ["base", "with base", "including base", "plinth"])

    return {
        "scale": scale,
        "size_mm": size_mm,
        "theme": theme,
        "product_type": product_type,
        "fig_count": fig_count,
        "is_unpainted": is_unpainted,
        "has_base": has_base,
    }


def make_onbuy_description(title):
    """Generate a clean, readable HTML description for OnBuy from the product title.

    OnBuy accepts basic HTML. We create a structured, informative description
    that reads well and covers what a buyer needs to know.
    """
    clean = clean_title_onbuy(title)
    attr = _detect_attributes(title)

    lines = []

    # Opening paragraph
    opening = f"<p><strong>{clean}</strong></p>\n"
    opening += "<p>"
    opening += f"A highly detailed resin {attr['product_type'].lower()} kit"
    if attr["scale"]:
        opening += f" in {attr['scale']} scale"
    if attr["fig_count"]:
        opening += f", featuring {attr['fig_count']} figure(s)"
    opening += "."
    if attr["size_mm"]:
        opening += f" Approximately {attr['size_mm']} tall."
    if attr["theme"] != "Fantasy":
        opening += f" {attr['theme']} themed."
    if attr["is_unpainted"]:
        opening += " Supplied unassembled and unpainted, ready for your creative touch."
    else:
        opening += " May require assembly and painting."
    opening += "</p>"
    lines.append(opening)

    # What's in the box
    box = "\n<p><strong>What's Included</strong></p>\n<ul>\n"
    box += f"  <li>Resin {attr['product_type'].lower()} kit"
    if attr["fig_count"]:
        box += f" ({attr['fig_count']} figure(s))"
    box += "</li>\n"
    box += "  <li>Cast in high-quality resin for fine detail</li>\n"
    if attr["has_base"]:
        box += "  <li>Display base included</li>\n"
    box += "</ul>"
    lines.append(box)

    # Specifications
    specs = "\n<p><strong>Specifications</strong></p>\n<ul>\n"
    specs += "  <li><strong>Material:</strong> Resin</li>\n"
    if attr["scale"]:
        specs += f"  <li><strong>Scale:</strong> {attr['scale']}</li>\n"
    if attr["size_mm"]:
        specs += f"  <li><strong>Height:</strong> Approx. {attr['size_mm']}</li>\n"
    specs += f"  <li><strong>Theme:</strong> {attr['theme']}</li>\n"
    specs += f"  <li><strong>Type:</strong> {attr['product_type']}</li>\n"
    specs += "  <li><strong>Condition:</strong> Brand new, unbuilt kit</li>\n"
    specs += "</ul>"
    lines.append(specs)

    # Assembly tips
    tips = "\n<p><strong>Assembly &amp; Painting Tips</strong></p>\n<ul>\n"
    tips += "  <li>Wash all parts in warm soapy water before priming to remove mould release</li>\n"
    tips += "  <li>Use super glue (cyanoacrylate) or two-part epoxy for assembly</li>\n"
    tips += "  <li>Prime with a grey or white spray primer before painting</li>\n"
    tips += "  <li>Acrylic or enamel model paints are both suitable</li>\n"
    tips += "</ul>"
    lines.append(tips)

    # Ideal for
    ideal = "\n<p><strong>Ideal For</strong></p>\n<ul>\n"
    ideal += "  <li>Scale model builders and painters</li>\n"
    ideal += "  <li>Diorama creators</li>\n"
    ideal += "  <li>Military history and miniature collectors</li>\n"
    ideal += "  <li>Wargaming and tabletop RPG enthusiasts</li>\n"
    ideal += "  <li>A unique gift for hobby enthusiasts</li>\n"
    ideal += "</ul>"
    lines.append(ideal)

    # Note
    note = "\n<p><em>Please examine all photos carefully. "
    note += "Images show the kit assembled and painted as a reference only &mdash; "
    note += "you will receive the unassembled, unpainted resin kit.</em></p>"
    lines.append(note)

    desc = "\n".join(lines)

    # OnBuy max 50,000 chars
    if len(desc) > 49000:
        desc = desc[:49000]

    return desc


def make_summary_points(title):
    """Generate up to 5 summary bullet points for OnBuy listing."""
    attr = _detect_attributes(title)
    points = []

    desc = f"High-detail resin {attr['product_type'].lower()} kit"
    if attr["scale"]:
        desc += f" in {attr['scale']} scale"
    points.append(desc)

    if attr["theme"]:
        points.append(f"{attr['theme']} themed — perfect for collectors and modellers")

    points.append("Cast in premium resin for crisp, fine detail")

    if attr["is_unpainted"]:
        points.append("Supplied unassembled and unpainted — customise to your vision")
    else:
        points.append("May require assembly and painting for a personalised finish")

    points.append("Ideal for dioramas, display cabinets, and wargaming tables")

    return points[:5]


# ============================================================
# GTIN LOADER
# ============================================================

def load_gtins(gtin_path):
    """Load GTIN-13 numbers from the GS1 CSV export.

    The CSV uses Excel formula format: ="5061122950005"
    We strip the ="..." wrapper to get the raw number.
    """
    gtins = []
    with open(gtin_path, "r", encoding="utf-8-sig") as f:
        # Skip metadata lines until we hit the header
        for line in f:
            if line.strip().startswith("Number"):
                break
        reader = csv.DictReader(f, fieldnames=[
            "Number", "Status", "Description", "Main Brand",
            "Sub Brand", "Product Link", "MPN", "SKU", "Updated"
        ])
        for row in reader:
            raw = row.get("Number", "").strip()
            # Strip Excel formula wrapper: ="5061122950005" -> 5061122950005
            num = re.sub(r'[="\'"]', '', raw).strip()
            if num and num.isdigit() and len(num) == 13:
                gtins.append(num)
    return gtins


# ============================================================
# IMAGE REHOSTING (uses ali_to_etsy functions)
# ============================================================

def rehost_all_images_onbuy(products):
    """Rehost all images for products in parallel using ali_to_etsy's rehost_image."""
    tasks = []
    for pi, row in enumerate(products):
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

    print(f"[INFO] Rehosting {len(tasks)} images across {len(products)} products ({REHOST_WORKERS} workers)...")

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
                    pass
                else:
                    failed += 1
            except Exception:
                failed += 1
            if done % 50 == 0 or done == len(tasks):
                print(f"[INFO]   {done}/{len(tasks)} images processed ({failed} failed)")

    for pi, row in enumerate(products):
        rehosted = []
        for ii in range(MAX_IMAGES_PER_LISTING):
            if (pi, ii) in results:
                rehosted.append(results[(pi, ii)])
        row["_rehosted_images"] = rehosted

    print(f"[INFO] Rehosted {len(results)} images total, {failed} failed")
    return products


# ============================================================
# ONBUY CSV WRITER
# ============================================================

# OnBuy template headers (matching the official template exactly)
ONBUY_HEADERS = [
    "SKU", "Product_Name", "Description", "Default_Image",
    "Brand", "Category", "Condition", "EAN",
    "Price", "Stock", "Handling_Time", "Shipping_Template_Id",
    "Shipping_Weight_(Kg)", "Warranty (Months)", "Free Returns",
    "ASIN", "MPN", "RRP",
    "Parent_Group", "Variant_One_Name", "Variant_One_Value",
    "Variant_Two_Name", "Variant_Two_Value",
    "Clothing Size", "Colour",
    "Summary_Point_One", "Summary_Point_Two", "Summary_Point_Three",
    "Summary_Point_Four", "Summary_Point_Five",
    "Additional_images_One", "Additional_images_Two",
    "Additional_images_Three", "Additional_images_Four",
    "Additional_images_Five", "Additional_images_Six",
    "Additional_images_Seven", "Additional_images_Eight",
    "Additional_images_Nine", "Additional_images_Ten",
]


def write_onbuy_csv(onbuy_rows, output_path):
    """Write products to OnBuy Product Create Template format as CSV."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=ONBUY_HEADERS, extrasaction="ignore",
            quoting=csv.QUOTE_ALL,
        )
        writer.writeheader()
        writer.writerows(onbuy_rows)
    print(f"[INFO] Saved {len(onbuy_rows)} products to {output_path}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("  AliExpress CSV -> OnBuy Product Upload")
    print("=" * 60)

    # ---- Find AliExpress CSV ----
    if len(sys.argv) >= 2:
        csv_path = sys.argv[1]
    else:
        csv_path = ALIEXPRESS_CSV
    if not os.path.exists(csv_path):
        print(f"[ERROR] AliExpress CSV not found: {csv_path}")
        sys.exit(1)

    # ---- Load GTINs ----
    if not os.path.exists(GTIN_CSV):
        print(f"[ERROR] GTIN CSV not found: {GTIN_CSV}")
        sys.exit(1)
    gtins = load_gtins(GTIN_CSV)
    print(f"[INFO] Loaded {len(gtins)} GTINs")
    if len(gtins) < MAX_PRODUCTS:
        print(f"[WARN] Only {len(gtins)} GTINs available, will limit to that")

    # ---- Load AliExpress products ----
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    print(f"[INFO] Loaded {len(rows)} products from {csv_path}")

    # ---- Deduplicate ----
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
    print(f"[INFO] After dedup: {len(unique)} unique products")

    # ---- Filter: must have a parseable price ----
    priced = []
    for row in unique:
        price = parse_price(row.get("product_price", ""))
        if price and price > 0:
            priced.append(row)
    print(f"[INFO] Products with valid price: {len(priced)}")

    # ---- Sort by total_sales descending (best sellers first) ----
    def get_sales(row):
        sales_str = row.get("total_sales", "0")
        nums = re.findall(r'\d+', sales_str)
        return int(nums[0]) if nums else 0
    priced.sort(key=get_sales, reverse=True)

    # ---- Limit to available GTINs ----
    limit = min(len(priced), len(gtins), MAX_PRODUCTS)
    products = priced[:limit]
    print(f"[INFO] Using top {limit} products (by sales)")

    # ---- Rehost images ----
    products = rehost_all_images_onbuy(products)

    # ---- Build OnBuy rows ----
    onbuy_rows = []
    skipped_no_images = 0
    gtin_idx = 0

    for row in products:
        images = row.get("_rehosted_images", [])
        if not images:
            skipped_no_images += 1
            continue

        if gtin_idx >= len(gtins):
            break

        title = row.get("product_title", "")
        clean = clean_title_onbuy(title)
        cost_gbp = parse_price(row.get("product_price", ""))
        sell_price = calculate_onbuy_price(cost_gbp)
        ean = gtins[gtin_idx]
        gtin_idx += 1

        # SKU from AliExpress product ID
        sku = "OB-" + row.get("id", str(gtin_idx)).strip()

        # Description
        description = make_onbuy_description(title)

        # Summary points
        summary = make_summary_points(title)

        # RRP (recommend ~40% above sell price for the "save" badge)
        rrp = round(sell_price * 1.4, 2)

        # Build the row
        onbuy_row = {
            "SKU": sku,
            "Product_Name": clean,
            "Description": description,
            "Default_Image": images[0],
            "Brand": BRAND,
            "Category": CATEGORY,
            "Condition": CONDITION,
            "EAN": ean,
            "Price": f"{sell_price:.2f}",
            "Stock": STOCK,
            "Handling_Time": HANDLING_DAYS,
            "Shipping_Template_Id": "Deliveries",
            "Shipping_Weight_(Kg)": "",
            "Warranty (Months)": "",
            "Free Returns": FREE_RETURNS,
            "ASIN": "",
            "MPN": "",
            "RRP": f"{rrp:.2f}",
            "Parent_Group": "",
            "Variant_One_Name": "",
            "Variant_One_Value": "",
            "Variant_Two_Name": "",
            "Variant_Two_Value": "",
            "Clothing Size": "",
            "Colour": "",
            "Summary_Point_One": summary[0] if len(summary) > 0 else "",
            "Summary_Point_Two": summary[1] if len(summary) > 1 else "",
            "Summary_Point_Three": summary[2] if len(summary) > 2 else "",
            "Summary_Point_Four": summary[3] if len(summary) > 3 else "",
            "Summary_Point_Five": summary[4] if len(summary) > 4 else "",
        }

        # Additional images (up to 10)
        additional_keys = [
            "Additional_images_One", "Additional_images_Two",
            "Additional_images_Three", "Additional_images_Four",
            "Additional_images_Five", "Additional_images_Six",
            "Additional_images_Seven", "Additional_images_Eight",
            "Additional_images_Nine", "Additional_images_Ten",
        ]
        for i, key in enumerate(additional_keys):
            if i + 1 < len(images):
                onbuy_row[key] = images[i + 1]
            else:
                onbuy_row[key] = ""

        onbuy_rows.append(onbuy_row)

    # ---- Write output ----
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"onbuy_upload_{ts}.csv"
    write_onbuy_csv(onbuy_rows, output_name)

    # ---- Summary ----
    print()
    print("=" * 60)
    print(f"  DONE! {len(onbuy_rows)} products -> {output_name}")
    print("=" * 60)
    print()

    # Price summary
    if onbuy_rows:
        prices = [float(r["Price"]) for r in onbuy_rows]
        print(f"  Price range: £{min(prices):.2f} — £{max(prices):.2f}")
        print(f"  Average price: £{sum(prices)/len(prices):.2f}")

    # Profit check on a few examples
    print()
    print("  Sample profit checks (after OnBuy fees):")
    for row, orig in zip(onbuy_rows[:5], products[:5]):
        cost = parse_price(orig.get("product_price", "")) + ALI_SHIPPING_GBP
        sell = float(row["Price"])
        fee = sell * ONBUY_COMMISSION
        profit = sell - fee - cost
        margin = profit / sell * 100 if sell > 0 else 0
        print(f"    £{sell:.2f} sell  |  £{cost:.2f} cost  |  £{fee:.2f} fee  |  £{profit:.2f} profit  |  {margin:.1f}% margin")

    print()
    if skipped_no_images:
        print(f"  {skipped_no_images} products skipped (image rehosting failed)")
    print(f"  {gtin_idx} GTINs used of {len(gtins)} available")
    print()
    print("NEXT STEPS:")
    print(f"  1. Go to OnBuy Seller Portal > Products > Upload")
    print(f"  2. Upload {output_name}")
    print(f"  3. Review and publish listings")
    print()


if __name__ == "__main__":
    main()
