"""
Dedicated scraper for: USAID / U.S. Embassy Lebanon
Source: https://lb.usembassy.gov/category/news/

NOTE: the original config.csv URL (https://lb.usembassy.gov/news/) isn't the
real paginated archive - that's "News Archives" at
https://lb.usembassy.gov/category/news/, a standard WordPress category
listing (30+ pages) with real numbered <a href> pager links
(.../category/news/page/2/, etc.) and each post's link inside an <h2
class="entry-title">.

Getting this working also surfaced a real bug in common.py shared by every
site: WordPress path-based pagination links like ".../category/news/page/2/"
were slipping past PAGINATION_QUERY_PATTERN (which only catches "?page="
query strings) and then wrongly matching DETAIL_PATH_PATTERN too ("news/"
immediately followed by "page" looked like a real article slug), so numbered
pager links were being scraped as if they were articles. That's now fixed
with a PAGINATION_PATH_PATTERN in common.py, so no custom code is needed here
- just the corrected URL.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'USAID / U.S. Embassy Lebanon'
URL = 'https://lb.usembassy.gov/category/news/'
OUTPUT_FILE = 'USAID_US_Embassy_Lebanon.xlsx'
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
