"""
Dedicated scraper for: Ministry of Justice
Source: https://www.justice.gov.lb/index.php/news-announcements/2

NOTE: config previously only had the bare homepage, which only shows a
1-item news carousel. The "مشاهدة الكل" (view all) link on the homepage
points at this combined news+announcements listing page (its "#1"/"#2"
fragment just switches between two in-page tabs - stripped harmlessly by
common.normalize_url()'s urldefrag). Detail links look like
".../news-details/1454/2" which already matches common.py's
DETAIL_PATH_PATTERN (news + hyphenated suffix + numeric id). Whether this
page has further real pagination beyond what's shown, or is itself a
JS/AJAX-loaded "load more" list, hasn't been confirmed live.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Justice'
URL = 'https://www.justice.gov.lb/index.php/news-announcements/2'
OUTPUT_FILE = 'Ministry_of_Justice.xlsx'
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
