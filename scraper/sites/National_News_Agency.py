"""
Dedicated scraper for: National News Agency - Ministry of Information
Source: https://www.nna-leb.gov.lb/ar/latest-news
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

Confirmed live: pagination is Previous/Next <button> elements with no href at
all (22,229 pages total, per "الصفحة 1 من 22229" in the pager) - a client-side
React/Vue-style app, not real links:

    <div class="flex justify-center gap-2 mt-6">
      <button disabled class="px-3 py-1 rounded bg-gray-200 disabled:opacity-50 ...">السابق</button>
      <span class="px-3 py-1 font-medium">الصفحة 1 من 22229</span>
      <button class="px-3 py-1 rounded bg-gray-200 disabled:opacity-50 ...">التالي</button>
    </div>

Both buttons share identical classes (generic Tailwind utility classes, no
"prev"/"next" class to key off of) and no id - only their position in the DOM
distinguishes them, so NEXT_SELECTOR targets "the last <button> in the pager's
flex row" rather than any class/id. Uses common.run_scrape_click_pagination.

Also worth noting: this site's Tailwind classes include "disabled:opacity-50"
on BOTH buttons always (that's just Tailwind's "apply this style when the
disabled attribute is present" variant syntax) - common.py's disabled-check
used to substring-match "disabled" anywhere in the class list, which would
have matched this literally on every page and stopped pagination immediately.
That's been fixed in common.py to check for "disabled" as an exact class
token / the real disabled attribute instead.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'National News Agency - Ministry of Information'
URL = 'https://www.nna-leb.gov.lb/ar/latest-news'
OUTPUT_FILE = 'National_News_Agency.xlsx'
NEXT_SELECTOR = "div.flex.justify-center.gap-2 button:last-of-type"

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
