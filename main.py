# pip install playwright pandas openpyxl
# playwright install   # run once in terminal

import pandas as pd
from playwright.sync_api import sync_playwright
import time
import logging
from pathlib import Path
from urllib.parse import urlparse
import os

# ────────────────────────────────────────────────
INPUT_EXCEL     = os.getenv("INPUT_EXCEL", "Cartoonxlsx.xlsx")           # Your input file
OUTPUT_EXCEL    = os.getenv("OUTPUT_EXCEL", "updated_videos_with_links.xlsx")
OUTPUT_M3U      = os.getenv("OUTPUT_M3U", "all_channels_playlist.m3u")
WAIT_AFTER_LOAD = 12                            # seconds — increase if needed
HEADLESS        = True                          # Set False to watch/debug locally (not in CI)
MAX_CONCURRENT  = 3                             # 2–5 recommended
# ────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')

def find_m3u8_urls(page, target_url):
    found_urls = set()

    def on_request(request):
        url_lower = request.url.lower()
        if any(x in url_lower for x in [".m3u8", "master.m3u", "playlist.m3u", "chunklist"]):
            found_urls.add(request.url)

    page.on("request", on_request)

    try:
        logging.info(f"Processing: {target_url}")
        page.goto(target_url, wait_until="networkidle", timeout=60000)
        time.sleep(WAIT_AFTER_LOAD)

        # Optional: try to start playback (adjust selector for your sites)
        # try:
        #     page.click("button:has-text('Play'), [aria-label*='play' i]", timeout=10000)
        #     time.sleep(5)
        # except:
        #     pass

    except Exception as e:
        logging.error(f"Error on {target_url}: {str(e)}")

    return list(found_urls)


def main():
    if not Path(INPUT_EXCEL).is_file():
        logging.error(f"Input file not found: {INPUT_EXCEL}")
        return

    # Read Excel
    df = pd.read_excel(INPUT_EXCEL, dtype=str)
    logging.info(f"Loaded {len(df)} entries from {INPUT_EXCEL}")

    # Make sure required columns exist
    if 'Page URL' not in df.columns:
        logging.error("Excel must contain 'Page URL' column!")
        return

    # Prepare result column if missing
    if 'm3u8_url' not in df.columns:
        df['m3u8_url'] = pd.NA

    found_count = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)

        for i in range(0, len(df), MAX_CONCURRENT):
            batch = df.iloc[i:i + MAX_CONCURRENT]
            context = browser.new_context(
                viewport={'width': 1280, 'height': 720},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128.0.0.0",
                ignore_https_errors=True,
            )

            pages = [context.new_page() for _ in batch['Page URL']]

            for page, (_, row) in zip(pages, batch.iterrows()):
                url = row['Page URL']
                fresh_links = find_m3u8_urls(page, url)

                if fresh_links:
                    # Take the longest / most complete-looking one (usually master)
                    best = max(fresh_links, key=len)
                    df.at[_, 'm3u8_url'] = best
                    found_count += 1
                    logging.info(f"Found → {best}")
                else:
                    logging.warning(f"No m3u8 found for: {url}")

            # Cleanup
            for page in pages:
                try:
                    page.close()
                except:
                    pass
            context.close()

            time.sleep(1)  # small delay between batches

        browser.close()

    # ── Save updated Excel ───────────────────────────────────────
    df.to_excel(OUTPUT_EXCEL, index=False)
    logging.info(f"Updated Excel saved → {OUTPUT_EXCEL} ({found_count} new links)")

    # ── Generate M3U playlist (always, even if empty) ──────────────────────────
    valid = df.dropna(subset=['m3u8_url'])
    with open(OUTPUT_M3U, 'w', encoding='utf-8') as f:
        f.write("#EXTM3U\n")
        f.write("# Generated on " + time.strftime("%Y-%m-%d %H:%M:%S") + "\n\n")

        if len(valid) == 0:
            f.write("# No valid m3u8 links found yet — check logs for errors or add more URLs to input Excel.\n\n")
        else:
            for _, row in valid.iterrows():
                name = f"{row.get('Video Name', 'Unknown')}"
                if 'Season' in row and 'Episode' in row:
                    name = f"S{row['Season']}E{row['Episode']} - {name}"

                f.write(f"#EXTINF:-1 tvg-name=\"{name}\" tvg-language=\"TAM\",{name}\n")
                f.write(f"{row['m3u8_url']}\n\n")

    logging.info(f"M3U playlist created → {OUTPUT_M3U} ({len(valid)} channels)")

if __name__ == "__main__":
    main()

    # Uncomment to run every hour:
    # while True:
    #     main()
    #     print("\nWaiting 1 hour for next refresh...\n")
    #     time.sleep(3600)
