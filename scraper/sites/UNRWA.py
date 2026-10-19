"""
Dedicated scraper for: UNRWA (Lebanon)

unrwa.org itself sits behind a Cloudflare "Just a moment..." challenge - checked
2026-10-16 in a normal, logged-in desktop Chrome too, and it never lets an
automated visit through (the RSS feed is behind the same wall). We do NOT try
to get around that.

Instead this scraper reads UNRWA's own Lebanon publications from ReliefWeb
(the UN OCHA site that republishes every UNRWA report/press release), filtered
to source = UNRWA and country = Lebanon - about 1,100 items, with a normal
paginated listing that works with a plain request.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'UNRWA'
URL = 'https://reliefweb.int/updates?advanced-search=%28S1484%29_%28C137%29'   # S1484 = UNRWA, C137 = Lebanon
OFFICIAL_URL = 'https://www.unrwa.org/tags/lebanon'   # kept for reference (blocked by Cloudflare)
OUTPUT_FILE = 'UNRWA.xlsx'
NOTES = ''
LINK_PATTERN = r"reliefweb\.int/report/[^/]+/[^/?#]+"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape(ORG, URL, OUTPUT_FILE, args.max_pages, args.max_articles, args.delay, NOTES,
                          link_pattern=LINK_PATTERN)
    finally:
        common.shutdown()
