import os
import json
from PIL import Image, ImageDraw, ImageFont

# === CONFIG ===
FOLDER = "C:/Users/Philip Work/Documents/Python"
IMAGE_FOLDER = os.path.join(FOLDER, "affiliate_images")
OUTPUT_FOLDER = os.path.join(FOLDER, "instagram_output")
TALL_OUTPUT_FOLDER = os.path.join(FOLDER, "instagram_output_tall")
FONT_FOLDER = os.path.join(FOLDER, "fonts")
ASSETS_FOLDER = os.path.join(FOLDER, "assets")
JSON_PATH = os.path.join(FOLDER, "todays_deals.json")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TALL_OUTPUT_FOLDER, exist_ok=True)

# === Load product data ===
with open(JSON_PATH, "r", encoding="utf-8") as f:
    products = json.load(f)

# === Load Japanese fonts from font folder (.ttf versions) ===
font_title = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 36)
font_text = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Regular.ttf"), 28)
font_footer = ImageFont.truetype(os.path.join(FONT_FOLDER, "NotoSansJP-Bold.ttf"), 30)

# === Loop through each product ===
for p in products:
    asin = p.get("asin")
    if not asin:
        continue

    image_path = os.path.join(IMAGE_FOLDER, f"{asin}.jpg")
    if not os.path.exists(image_path):
        print(f"❌ Image not found for ASIN: {asin}")
        continue

    try:
        # Load and resize product image
        product_img = Image.open(image_path).convert("RGBA").resize((360, 360))

        # Draw border around the product image
        bordered_img = Image.new("RGBA", (364, 364), (0, 0, 0, 255))
        bordered_img.paste(product_img, (2, 2))

        # === Horizontal layout (wide and short) ===
        bg_width, bg_height = 940, 580
        bg_image = Image.open(os.path.join(ASSETS_FOLDER, "bg_horizontal.png")).convert("RGBA").resize((bg_width, bg_height))
        background = bg_image.copy()
        background.paste(bordered_img, (60, 80), mask=bordered_img)

        draw = ImageDraw.Draw(background)
        text_x = 480
        y_cursor = 80

        title = p.get("title", "")
        lines = [title[i:i+13] for i in range(0, len(title), 13)][:4]
        for line in lines:
            draw.text((text_x, y_cursor), line, font=font_title, fill=(0, 0, 0))
            y_cursor += 40

        stars_line = p.get("stars", "")
        if stars_line:
            raw_rating = stars_line.split("（")[-1].split("）")[0]
            rating_value = float(raw_rating)
            full_stars = int(rating_value)
            half_star = rating_value - full_stars >= 0.5
            stars_only = "★" * full_stars + ("½" if half_star else "")
            review_count = stars_line.split("｜")[-1].replace("）", "").replace("件", "").strip()
            combined = f"{stars_only}（{rating_value:.1f}｜{review_count}件）"
            draw.text((text_x, y_cursor), combined, font=font_text, fill=(255, 153, 0))
            y_cursor += 35

        discount_line = p.get("discount", "")
        if discount_line:
            clean_discount = discount_line.replace("🔻 割引：", "").strip()
            draw.text((text_x, y_cursor), f"割引：{clean_discount}", font=font_text, fill=(192, 0, 0))
            y_cursor += 35

        price_line = p.get("price", "")
        if price_line:
            draw.text((text_x, y_cursor), f"価格：{price_line.replace('💴 価格：', '')}", font=font_text, fill=(0, 0, 0))
            y_cursor += 35

        coupon_line = p.get("coupon", "")
        if coupon_line:
            draw.text((text_x, y_cursor), f"クーポン：{coupon_line.replace('🎟️ 追加クーポン：', '')}", font=font_text, fill=(255, 204, 0))
            y_cursor += 35

        verified_line = p.get("verified", "")
        if verified_line:
            clean_verified = verified_line.replace("🕒 確認済み：", "").replace("🕒", "").replace("確認済み：", "")
            draw.text((text_x, y_cursor), f"確認済み：{clean_verified}", font=font_text, fill=(100, 100, 100))
            y_cursor += 35

        footer_text = "ショップのリンクはプロフィールからご覧ください"
        footer_height = 80
        footer_banner = Image.new("RGBA", (bg_width, footer_height), (180, 0, 0, 255))
        draw_footer = ImageDraw.Draw(footer_banner)
        w = font_footer.getlength(footer_text)
        draw_footer.text(((bg_width - w) // 2, 20), footer_text, font=font_footer, fill=(255, 255, 255))
        background.paste(footer_banner, (0, bg_height - footer_height))

        output_path = os.path.join(OUTPUT_FOLDER, f"{asin}_insta.png")
        background.save(output_path)

        # === Tall layout (portrait style for story) ===
        tall_width, tall_height = 1080, 1920
        tall_bg_image = Image.open(os.path.join(ASSETS_FOLDER, "bg_vertical.png")).convert("RGBA").resize((tall_width, tall_height))
        tall_background = tall_bg_image.copy()
        tall_background.paste(bordered_img.resize((500, 500)), (290, 120), mask=bordered_img.resize((500, 500)))

        draw_tall = ImageDraw.Draw(tall_background)
        y_tall = 670

        title = p.get("title", "")
        lines = [title[i:i+17] for i in range(0, len(title), 17)][:4]
        for line in lines:
            draw_tall.text((100, y_tall), line, font=font_title, fill=(0, 0, 0))
            y_tall += 45

        if stars_line:
            draw_tall.text((100, y_tall), combined, font=font_text, fill=(255, 153, 0))
            y_tall += 35

        if discount_line:
            draw_tall.text((100, y_tall), f"割引：{clean_discount}", font=font_text, fill=(192, 0, 0))
            y_tall += 35

        if price_line:
            draw_tall.text((100, y_tall), f"価格：{price_line.replace('💴 価格：', '')}", font=font_text, fill=(0, 0, 0))
            y_tall += 35

        if coupon_line:
            draw_tall.text((100, y_tall), f"クーポン：{coupon_line.replace('🎟️ 追加クーポン：', '')}", font=font_text, fill=(255, 204, 0))
            y_tall += 35

        if verified_line:
            draw_tall.text((100, y_tall), f"確認済み：{clean_verified}", font=font_text, fill=(100, 100, 100))
            y_tall += 35

        tall_background.paste(footer_banner.resize((tall_width, footer_height)), (0, tall_height - footer_height))

        tall_output_path = os.path.join(TALL_OUTPUT_FOLDER, f"{asin}_story.png")
        tall_background.save(tall_output_path)

    except Exception as e:
        print(f"❌ Error creating image for ASIN {asin}: {e}")

print("✅ All previews generated.")
