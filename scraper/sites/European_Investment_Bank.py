"""
Dedicated scraper for: European Investment Bank
Source: https://www.eib.org/en/press/all/index

This one is special: EIB's "All releases" page is a JS search widget whose
pagination is a prev/next arrow-click control (a "page N of 503" box), not
real <a href="...page=2"> links, so it can't use the normal fetch+follow-links
approach. It uses common.run_scrape_click_pagination() instead, which keeps a
single Selenium session open and clicks the "Next" arrow repeatedly.

NOTE: this scrapes ALL EIB press releases (503+ pages, global, not just
Lebanon) since no Lebanon-only filter control was confirmed yet. If you find
the CSS selector for a "Lebanon" checkbox in the sidebar's Countries filter,
set FILTER_SELECTOR below to it and re-run - that will filter to Lebanon-only
results before pagination starts, instead of scraping everything.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'European Investment Bank'
URL = 'https://www.eib.org/en/press/all/index'
OUTPUT_FILE = 'European_Investment_Bank.xlsx'
NEXT_SELECTOR = '.nextPagination'
FILTER_SELECTOR = None  # e.g. 'input[value="Lebanon"]' or similar, once known

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape_click_pagination(
            ORG, URL, OUTPUT_FILE, NEXT_SELECTOR,
            filter_selector=FILTER_SELECTOR,
            max_pages=args.max_pages,
            max_articles=args.max_articles,
            delay=args.delay,
        )
    finally:
        common.shutdown()
