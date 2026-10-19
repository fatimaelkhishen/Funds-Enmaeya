"""
Dedicated scraper for: ReliefWeb / OCHA
Source: https://reliefweb.int/updates?list=Lebanon%20Updates&advanced-search=%28PC137%29
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

Using the filtered "Lebanon Updates" listing (advanced-search=(PC137), ReliefWeb's
Lebanon country-code filter) instead of the old /country/lbn landing page - it's a
much fuller feed (735 pages at last check) and its pager is a normal Drupal one
with real <a href="...&page=N" rel="next"> links, so common.run_scrape's generic
<a href>-following pagination (find_next_page_url) already handles it with no
special click-pagination driver needed. The listing itself needs Selenium (a
plain request comes back with the page shell before its JS renders the results),
which common.fetch_html already does automatically.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'ReliefWeb / OCHA'
URL = 'https://reliefweb.int/updates?list=Lebanon%20Updates&advanced-search=%28PC137%29'
OUTPUT_FILE = 'ReliefWeb_OCHA.xlsx'
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
