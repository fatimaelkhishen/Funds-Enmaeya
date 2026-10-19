"""
Dedicated scraper for: Ministry of Tourism
Source: https://mot.gov.lb/en/news/

NOTE: the site does have a real "News" nav item/page (unlike before, when
config only had the bare homepage), but as of this writing the page itself
says "News & Circulars section is coming Soon" - there are no articles
published there yet. Nothing to fix code-wise; this will start picking up
real articles automatically via common.run_scrape() once the ministry
actually publishes something there.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Tourism'
URL = 'https://mot.gov.lb/en/news/'
OUTPUT_FILE = 'Ministry_of_Tourism.xlsx'
NOTES = 'news section exists but has no published articles yet ("coming soon")'

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
