"""
Dedicated scraper for: EBRD
Source: https://www.ebrd.com/news?1=1&filterCountry=Lebanon
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

Confirmed live: pagination is a custom JS widget, not real links -

    <ul id="pagination">
      <li class="arrows arrow-left disabled"><span class="ebrd-icon-chevron-left"></span></li>
      <li tabindex="0" data-page="1" class="active">1</li>
      <li tabindex="0" data-page="2" class="">2</li>
      ...
      <li tabindex="0" data-page="695">695</li>
      <li tabindex="0" class="arrows arrow-right"><span class="ebrd-icon-chevron-right"></span></li>
    </ul>

695 pages total. Every page number and the prev/next arrows are plain <li>
elements (no href) with a JS click handler reading data-page - common.run_scrape's
<a href>-following pagination can't see any of that, so it only ever scraped
page 1. Uses common.run_scrape_click_pagination instead, clicking the
right-arrow <li> (class "arrow-right", distinct from the left arrow's
"arrow-left") repeatedly; it picks up a literal "disabled" class once the
real last page is reached, same as the left arrow already shows on page 1.

WORTH DOUBLE-CHECKING: 695 pages seems like a lot for a Lebanon-only filter.
European_Investment_Bank.py hit the exact same situation - its JS "All
releases" widget ignores a similar filter query param and just shows every
country's press releases globally. If EBRD's results here turn out to mostly
be non-Lebanon news too, the "?filterCountry=Lebanon" query param the URL
carries may not actually be applied by this JS widget (it may need the
sidebar filter control itself clicked, not just the URL parameter) - worth
spot-checking the saved articles once a run finishes.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'EBRD'
URL = 'https://www.ebrd.com/news?1=1&filterCountry=Lebanon'
OUTPUT_FILE = 'EBRD.xlsx'
NEXT_SELECTOR = "ul#pagination li.arrow-right"

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
            pager_selector="ul#pagination",
        )
    finally:
        common.shutdown()
