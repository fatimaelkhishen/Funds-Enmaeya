"""
Dedicated scraper for: Ministry of Interior and Municipalities
Source: https://moim.gov.lb/latest-news/

NOTE: config previously only had the bare homepage. The site is a modern
WordPress site with a real "الأخبار" (News) nav item pointing at a real
listing - this should already work fine with common.run_scrape()'s generic
WordPress-style handling (same engine already verified against
USAID_US_Embassy_Lebanon.py's WordPress "News Archives" listing).
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Interior and Municipalities'
URL = 'https://moim.gov.lb/latest-news/'
OUTPUT_FILE = 'Ministry_of_Interior_and_Municipalities.xlsx'
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
