"""
Dedicated scraper for: Ministry of Energy and Water
Source: https://www.energyandwater.gov.lb/ar/listing/newsroom/news?l=3

Confirmed live: pagination is an Angular controller call, not a real link:

    <div class="page-numbering"><ul>
      <li><a href="javascript:;" onclick="articlesCtrl.loadArticles(0, {"pushState":true})">1</a></li>
      ...
      <li><a href="javascript:;" title="Next" onclick="articlesCtrl.loadArticles(1, {"pushState":true})">
        <i class="fa fa-angle-right"></i></a></li>
      <li><a href="javascript:;" title="Last" onclick="articlesCtrl.loadArticles(11, {"pushState":true})">
        <i class="fa fa-angle-double-right"></i></a></li>
    </ul></div>

href="javascript:;" is a no-op - the real page-N load happens via the
articlesCtrl.loadArticles() JS call in onclick, which AJAX-replaces the
results in the SAME page (no distinct "page 2" URL to follow). common.run_scrape's
<a href>-following pagination can't see that and only ever scraped page 1.
Uses common.run_scrape_click_pagination instead (same approach as UNDP Lebanon /
UNHCR Lebanon), clicking the "Next" link (a[title="Next"]) - it's a real <a>
element so Selenium can click it directly even though its href does nothing.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Energy and Water'
URL = 'https://www.energyandwater.gov.lb/ar/listing/newsroom/news?l=3'
OUTPUT_FILE = 'Ministry_of_Energy_and_Water.xlsx'
NEXT_SELECTOR = '.page-numbering a[title="Next"]'

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
