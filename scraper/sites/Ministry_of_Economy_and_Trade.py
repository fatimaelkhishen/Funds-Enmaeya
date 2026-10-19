"""
Dedicated scraper for: Ministry of Economy and Trade
Source: https://economy.gov.lb/ar/news

Checked live 2026-10-16: 10 news items per page, numbered pager using
?page=2, ?page=3 ... ; each article is /ar/news/<arabic-slug>.

Previously only page 1 was scraped (pager bug in common.py, now fixed) and
most articles were wrongly thrown away as "bot challenge" pages (also fixed).
Article pages on this server can take 20+ seconds, so the timeout is raised.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Economy and Trade'
URL = 'https://economy.gov.lb/ar/news'
OUTPUT_FILE = 'Ministry_of_Economy_and_Trade.xlsx'
NOTES = ''
LINK_PATTERN = r"economy\.gov\.lb/ar/news/[^/?#]+$"
PAGE_URL_TEMPLATE = 'https://economy.gov.lb/ar/news?page={n}'

common.REQUEST_TIMEOUT = 60   # slow article pages

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
