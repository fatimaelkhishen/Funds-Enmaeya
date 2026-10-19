"""
Dedicated scraper for: Ministry of Education and Higher Education
Source: https://www.mehe.gov.lb/ar/media-center/news

NOTE: config previously only had the bare homepage. The real news listing
lives under "المركز الإعلامي" > "الأخبار" - a Drupal 10 site, so this
should already work with common.run_scrape()'s generic Drupal handling
(same engine already verified against United_Nations_in_Lebanon.py and
UNIFIL.py). Live pagination behavior hasn't been confirmed.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Education and Higher Education'
URL = 'https://www.mehe.gov.lb/ar/media-center/news'
OUTPUT_FILE = 'Ministry_of_Education_and_Higher_Education.xlsx'
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
