# pip install playwright pandas openpyxl pytz
# playwright install   # run once

import pandas as pd
from playwright.sync_api import sync_playwright
import time
import logging
from pathlib import Path
import os
from datetime import datetime
import pytz

# ────────────────────────────────────────────────
INPUT_EXCEL     = os.getenv("INPUT_EXCEL", "Cartoonxlsx.xlsx")
OUTPUT_EXCEL    = os.getenv("OUTPUT_EXCEL", "updated_videos_with_links.xlsx")
OUTPUT_M3U      = os.getenv("OUTPUT_M3U", "all_channels_playlist.m3u8")
WAIT_AFTER_LOAD = 15
HEADLESS        = True
MAX_CONCURRENT  = 2
# ────────────────────────────────────────────────

# ── IST Timezone ────────────────────────────────
IST = pytz.timezone("Asia/Kolkata")

# ── Logging with IST timestamps ─────────────────
class ISTFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        dt = datetime.fromtimestamp(record.created, IST)
        return dt.strftime(datefmt or "%Y-%m-%d %H:%M:%S")

formatter = ISTFormatter('%(asctime)s | %(levelname)s | %(message)s')
handler = logging.StreamHandler()
handler.setFormatter(formatter)

logging.getLogger().handlers.clear()
logging.getLogger().addHandler(handler)
logging.getLogger().setLevel(logging.INFO)


def find_m3u8_urls(page, target_url):
    found_urls = set()

    def on_request(request):
        url_lower = request.url.lower()
        if any(x in url_lower for x in [".m3u8", "master.m3u", "playlist.m3u", "chunklist"]):
            logging.info(f"Captured m3u8 request: {request.url}")
            found_urls.add(request.url)

    page.on("request", on_request)

    try:
        logging.info(f"Processing: {target_url}")
        page.goto(target_url, wait_until="networkidle", timeout=60000)
        time.sleep(WAIT_AFTER_LOAD)

        try:
            page.click(
                "button:has-text('Play'), [aria-label*='play' i], .play-button, video, div.video-player",
                timeout=10000
            )
            logging.info("Clicked potential play element")
            time.sleep(5)
        except Exception as e:
            logging.warning(f"No play element found or click failed: {str(e)}")

    except Exception as e:
        logging.error(f"Page load/error on {target_url}: {str(e)}")

    if not found_urls:
        logging.warning(f"No m3u8 URLs captured for {target_url}")

    return list(found_urls)


def main():
    logging.info(f"Script started at IST: {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S')}")

    input_path = Path(INPUT_EXCEL)
    if not input_path.is_file():
        logging.error(f"Input file not found: {INPUT_EXCEL}")
        pd.DataFrame(
            {"error": ["No input Excel found — upload Cartoonxlsx.xlsx"]}
        ).to_excel(OUTPUT_EXCEL, index=False)

        with open(OUTPUT_M3U, 'w', encoding='utf-8') as f:
            f.write("#EXTM3U\n")
            f.write("# ERROR: Input Excel not found\n")
        return

    df = pd.read_excel(INPUT_EXCEL, dtype=str)
    logging.info(f"Loaded {len(df)} rows from Excel")

    if 'Page URL' not in df.columns:
        logging.error("Excel must contain 'Page URL' column")
        df.to_excel(OUTPUT_EXCEL, index=False)
        return

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

            for page, (idx, row) in zip(pages, batch.iterrows()):
                fresh_links = find_m3u8_urls(page, row['Page URL'])

                if fresh_links:
                    best = max(fresh_links, key=len)
                    df.at[idx, 'm3u8_url'] = best
                    found_count += 1
                    logging.info(f"Selected m3u8 → {best}")
                else:
                    logging.warning(f"No m3u8 for: {row['Page URL']}")

            for page in pages:
                page.close()
            context.close()

            time.sleep(2 + i % 3)

        browser.close()

    df.to_excel(OUTPUT_EXCEL, index=False)
    logging.info(f"Updated Excel saved → {OUTPUT_EXCEL}")

    valid = df.dropna(subset=['m3u8_url'])

    with open(OUTPUT_M3U, 'w', encoding='utf-8') as f:
        f.write("#EXTM3U\n")
        f.write(
            "# Generated on " +
            datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S") +
            " IST\n\n"
        )

        if len(valid) == 0:
            f.write("# No valid m3u8 links found\n\n")
        else:
            for _, row in valid.iterrows():
                name = row.get('Video Name', 'Unknown')

                if 'Season' in row and 'Episode' in row:
                    name = f"S{int(row['Season']):02}E{int(row['Episode']):02} - {name}"

                f.write(f"#EXTINF:-1 tvg-name=\"{name}\" tvg-language=\"TAM\",{name}\n")
                f.write(f"{row['m3u8_url']}\n\n")

    logging.info(f"M3U playlist created → {OUTPUT_M3U} ({len(valid)} entries)")
    logging.info("Script completed successfully")


if __name__ == "__main__":
    main()
