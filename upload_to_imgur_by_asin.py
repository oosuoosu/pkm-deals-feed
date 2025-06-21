
import os
import requests
import json

# Configuration
CLIENT_ID = "c296b8f8ec0baa0"
IMAGE_FOLDER = "C:/Users/Philip Work/Documents/Python/affiliate_images"
OUTPUT_JSON = "C:/Users/Philip Work/Documents/Python/imgur_links_by_asin.json"

headers = {"Authorization": f"Client-ID {CLIENT_ID}"}
upload_url = "https://api.imgur.com/3/image"

image_links = {}

for filename in os.listdir(IMAGE_FOLDER):
    if filename.lower().endswith((".jpg", ".jpeg", ".png")):
        asin = os.path.splitext(filename)[0]
        image_path = os.path.join(IMAGE_FOLDER, filename)

        with open(image_path, "rb") as img:
            response = requests.post(upload_url, headers=headers, files={"image": img})

        if response.status_code == 200:
            link = response.json()["data"]["link"]
            image_links[asin] = link
            print(f"✅ Uploaded {filename}: {link}")
        else:
            print(f"❌ Failed to upload {filename}: {response.json()}")

# Save output mapping
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(image_links, f, ensure_ascii=False, indent=2)

print(f"📁 Saved Imgur links to {OUTPUT_JSON}")
