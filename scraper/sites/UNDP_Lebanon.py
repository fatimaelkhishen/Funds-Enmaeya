"""
Dedicated scraper for: UNDP Lebanon
Source: https://www.undp.org/lebanon/news-centre
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

This listing page only renders its first ~9 press releases in the initial
HTML; the rest (183 total, confirmed via the page's own
data-totalresult="183" attribute) load only when a "View More" JS button
(<button class="... load-more-custom" data-pageno="0" data-loadedresult="9"
data-totalresult="183">) is clicked - it AJAX-appends more results into the
SAME page rather than linking to a real "page 2" URL. common.run_scrape's
<a href>-following pagination can never find that, so it silently stopped
after page 1. Uses common.run_scrape_click_pagination instead (same approach
as EIB / Ministry of Foreign Affairs), which keeps one Selenium session open
and clicks the button repeatedly until it's gone/disabled.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'UNDP Lebanon'
URL = 'https://www.undp.org/lebanon/news-centre'
OUTPUT_FILE = 'UNDP_Lebanon.xlsx'
NEXT_SELECTOR = 'button.load-more-custom'

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
