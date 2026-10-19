"""
Dedicated scraper for: Ministry of Foreign Affairs and Emigrants
Source: https://mfa.gov.lb/arabic/news

NOTE: the article listing itself is real, static HTML (real links like
/arabic/news/260714 with real headlines), but pagination is a classic
ASP.NET WebForms postback pager - the "next" control and every numbered
page link is a javascript:WebForm_DoPostBackWithOptions(...) handler with
no real href/URL at all:

    <a href="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions
        ("ctl00$BodyPlaceHolder$lbtnNext", "", true, "", "", false, true))">&raquo;</a>

common.find_next_page_url() can't follow this (normalize_url() correctly
rejects javascript: hrefs, so there is nothing to extract a real page-2 URL
from) - the crawl would silently stop after page 1 forever. This uses
common.run_scrape_click_pagination() instead, which keeps one Selenium
session open and clicks the actual rendered "next" control (matched by its
rendered id, since ASP.NET renders "ctl00$BodyPlaceHolder$lbtnNext" as
id="ctl00_BodyPlaceHolder_lbtnNext") rather than trying to read an href.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Foreign Affairs and Emigrants'
URL = 'https://mfa.gov.lb/arabic/news'
OUTPUT_FILE = 'Ministry_of_Foreign_Affairs.xlsx'
NOTES = ''

# Matches the rendered <a id="ctl00_BodyPlaceHolder_lbtnNext" ...> "next" control
# regardless of exact master-page prefix - ASP.NET always renders the control's
# dotted "$"-separated ID with underscores, ending in "_lbtnNext".
NEXT_SELECTOR = "a[id$='lbtnNext']"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape_click_pagination(
            ORG, URL, OUTPUT_FILE, NEXT_SELECTOR,
            max_pages=args.max_pages, max_articles=args.max_articles, delay=args.delay,
        )
    finally:
        common.shutdown()
