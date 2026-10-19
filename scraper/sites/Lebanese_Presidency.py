"""
Dedicated scraper for: Lebanese Presidency
Source: https://www.presidency.gov.lb/media

Checked live 2026-10-16: the listing is a plain server-rendered page with a
numbered pager (?page=2, ?page=3 ...). Each news item lives at
/media/news/<number>. The old run picked up the section links
(/media/images, /media/videos, /media/category/...) as if they were
articles, and real articles were wrongly rejected as "bot challenge" pages -
both fixed: LINK_PATTERN below only accepts real /media/news/<id> pages, and
common.is_bot_challenge_page() no longer scans <script> text.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Lebanese Presidency'
URL = 'https://www.presidency.gov.lb/media'
OUTPUT_FILE = 'Lebanese_Presidency.xlsx'
NOTES = ''
LINK_PATTERN = r"presidency\.gov\.lb/media/news/\d+/?$"
PAGE_URL_TEMPLATE = 'https://www.presidency.gov.lb/media?page={n}'

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape(ORG, URL, OUTPUT_FILE, args.max_pages, args.max_articles, args.delay, NOTES,
                          link_pattern=LINK_PATTERN, page_url_template=PAGE_URL_TEMPLATE)
    finally:
        common.shutdown()
