import requests
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

# 1. Download the product image from Imgur
imgur_url = "https://i.imgur.com/GrTVXXH.jpeg"
response = requests.get(imgur_url)
product_img = Image.open(BytesIO(response.content)).convert("RGBA")

# 2. Create the background (beige color)
bg_width, bg_height = 768, 1024
background = Image.new("RGBA", (bg_width, bg_height), (245, 235, 220, 255))

# 3. Resize and paste the product image
product_img = product_img.resize((360, 360))
background.paste(product_img, (60, 80))

# 4. Load fonts (system fonts or bundled ones)
font_title = ImageFont.truetype("arialbd.ttf", 36)
font_text = ImageFont.truetype("arial.ttf", 28)
font_footer = ImageFont.truetype("arialbd.ttf", 30)

draw = ImageDraw.Draw(background)

# 5. Define the product data
title = "モンプチ キャットフード\nバッグ ドライ 5種のフィッシュ\nブレンド 成猫用 600g"
stars = "\u2605\u2605\u2605\u2605\u2605  (4.3)"
reviews = "2063件"
discount = "41% OFF"
price = "\u00a5358"
verified = "確認済み：6/19 3:30AM"
footer = "ショップのリンクはプロフィールからご覧ください"

# 6. Add text overlays
text_x = 450

for i, line in enumerate(title.split("\n")):
    draw.text((text_x, 80 + i*40), line, font=font_title, fill=(0, 0, 0))

# Ratings
star_y = 240
draw.text((text_x, star_y), stars, font=font_text, fill=(255, 153, 0))
draw.text((text_x, star_y + 35), reviews, font=font_text, fill=(0, 0, 0))

# Discount and price
draw.text((text_x, star_y + 90), f"割引: {discount}", font=font_text, fill=(192, 0, 0))
draw.text((text_x, star_y + 140), f"価格: {price}", font=font_text, fill=(0, 0, 0))

# Verified
draw.text((text_x, star_y + 200), verified, font=font_text, fill=(80, 80, 80))

# Footer banner
footer_height = 80
footer_bg = Image.new("RGBA", (bg_width, footer_height), (180, 0, 0, 255))
draw_footer = ImageDraw.Draw(footer_bg)
draw_footer.text((bg_width//2 - 250, 20), footer, font=font_footer, fill=(255, 255, 255))

# Paste footer to bottom of the image
background.paste(footer_bg, (0, bg_height - footer_height))

# 7. Save or show
background.save("monpetit_promo_final.png")
background.show()
