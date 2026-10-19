"""
Dedicated scraper for: European Union Delegation to Lebanon
Source: https://www.eeas.europa.eu/eeas/press-material_en?f%5B0%5D=press_site%3ALebanon&s=203&page=0
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'European Union Delegation to Lebanon'
URL = 'https://www.eeas.europa.eu/eeas/press-material_en?f%5B0%5D=press_site%3ALebanon&s=203&page=0'
OUTPUT_FILE = 'EU_Delegation_Lebanon.xlsx'
NOTES = 'force_selenium - results grid loads via JS/AJAX'

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
