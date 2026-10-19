"""
Dedicated scraper for: UNHCR Lebanon
Source: https://www.unhcr.org/lb/news-and-stories/press-releases
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

This listing's pager is a Drupal Views AJAX "search-pager" widget:

    <li class="search-pager__item search-pager__item--next">
      <div class="search-pager__item-link" role="button" tabindex="0"
           title="Go to next page" rel="next"><span class="page-next">Next</span></div>
    </li>

Every page number, First/Previous/Next/Last control is a <div role="button">
with a "page" attribute, not a real <a href="...page=2">, so it AJAX-replaces
the results in the SAME page instead of navigating anywhere - common.run_scrape's
<a href>-following pagination can't see it at all and only ever scraped page 1.
Uses common.run_scrape_click_pagination instead (same approach as UNDP Lebanon /
EIB), clicking the "Next" control repeatedly.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'UNHCR Lebanon'
URL = 'https://www.unhcr.org/lb/news-and-stories/press-releases'
OUTPUT_FILE = 'UNHCR_Lebanon.xlsx'
NEXT_SELECTOR = 'li.search-pager__item--next .search-pager__item-link'

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
