"""
Dedicated scraper for: IOM
Source: https://www.iom.int/news
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

NOTE (2026-07-21): switched to the global IOM news listing per explicit
request - loop over ALL pages (this listing has ~497 pages via real
<a href="?page=N" rel="next"> links, which find_next_page_url already
handles). This is NOT Lebanon-scoped; it's every IOM news item site-wide.
Since it's ~497 pages, pass --max-pages 500 (or higher) when running, e.g.:
    python IOM.py --max-pages 500
otherwise it'll stop early at common.DEFAULT_MAX_PAGES (300).

(Earlier notes, now superseded: https://www.iom.int/lebanon-crisis was tried
first - that's a static topic/landing page with no pagination and almost no
real article links, so it was swapped for mena.iom.int/authors/iom-lebanon -
but the user then asked to use the global /news listing and page through all
of it instead, so that's what this points at now.)
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'IOM'
URL = 'https://www.iom.int/news'
OUTPUT_FILE = 'IOM.xlsx'
NOTES = ''

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape(ORG, URL, OUTPUT_FILE, args.max_pages, args.max_articles, args.delay, NOTES)
    finally:
        common.shutdown()
