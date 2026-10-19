"""
Dedicated scraper for: UNESCO Lebanon
Source: https://www.unesco.org/en/newsroom
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

NOTE (2026-07-21): switched from https://www.unesco.org/en/countries/lb/latestnews
(consistently 500'd / bot-challenged from this network - see git history/
scrape_log.txt for that investigation) to the global newsroom, per request.
This page only renders its first batch of items statically; the rest load via
a "Load more" JS link (<a class="use-ajax" onclick="window.Unesco.clickLoadMore(this)">),
which AJAX-appends more results into the SAME page rather than linking to a
real "page 2" URL - same situation as UNDP Lebanon's "View More" button, so
this uses common.run_scrape_click_pagination (keeps one Selenium session open,
clicks the link repeatedly) instead of the default href-following run_scrape.
NOT Lebanon-scoped - this is UNESCO's worldwide newsroom, same tradeoff as the
IOM global /news switch earlier.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'UNESCO Lebanon'
URL = 'https://www.unesco.org/en/newsroom'
OUTPUT_FILE = 'UNESCO_Lebanon.xlsx'
NEXT_SELECTOR = 'a.use-ajax[onclick*="clickLoadMore"]'

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape_click_pagination(
            ORG, URL, OUTPUT_FILE, NEXT_SELECTOR,
            max_pages=args.max_pages,
            max_articles=args.max_articles,
            delay=args.delay,
        )
    finally:
        common.shutdown()
