"""
Dedicated scraper for: Ministry of Industry
Source: http://www.industry.gov.lb/Media/News
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Industry'
URL = 'http://www.industry.gov.lb/Media/News'
OUTPUT_FILE = 'Ministry_of_Industry.xlsx'
# no_selenium: confirmed in scrape_log.txt (2026-07-21 runs) that Selenium simply
# cannot render this ASP.NET/Sitefinity site (industry.gov.lb) - it comes back with
# a near-empty shell (just the <title>, one stray href="#") on EVERY page, in BOTH
# headless mode and force_visible_browser (non-headless) mode, including page 1,
# which a plain static request renders perfectly (10 real article links). Since
# visible mode failed identically to headless, this isn't a headless-fingerprinting
# issue like UNRWA's - Selenium is just broken for this site, period. So: never
# escalate to Selenium at all, and always trust the static fetch.
NOTES = 'no_selenium'

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
