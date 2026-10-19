"""
Dedicated scraper for: Ministry of Agriculture
Source: http://www.agriculture.gov.lb/Media/News

NOTE (2026-07-22): this used to drive its own dedicated Selenium session
(wait_for_results/fetch_rendered_html below, kept only for reference - no
longer called) on the theory that this listing's news cards are rendered
client-side via JS and need an explicit wait. That theory doesn't hold up:

  - A plain static GET of http://www.agriculture.gov.lb/Media/News (no JS at
    all) already returns the full page - 10 real headline links, each
    properly wrapped in <h2><a href="/Media/News/2026/...">, plus a real
    numbered pager ("عرض النتائج 1-10 (من 2569)" / pages 2-257 via
    "?page=N", with literal ">" / ">|" next/last links) - confirmed by
    fetching it directly. Nothing here needs a browser.
  - Every recent run of the OLD custom-Selenium version (scrape_log.txt,
    2026-07-21/22) shows wait_for_results itself succeeding - it finds the 10
    real "article.news-list-item h2.news-title a[href]" elements fine, and
    even confirms the captured snapshot contains "news-list-item" - but then
    common.extract_candidate_links(), run on that SAME captured HTML, finds
    only 1 total <a href> on the whole page (href="#"). That's the exact
    "Selenium's render of this ASP.NET/Sitefinity site comes back broken"
    failure already diagnosed and fixed (via notes='no_selenium') for
    Ministry_of_Industry.py (industry.gov.lb) and
    Ministry_of_Public_Works_and_Transport.py (mopw.gov.lb) - same CMS
    family, same symptom. The custom driver here forces Selenium
    unconditionally (never even tries a static fetch), so it could never
    benefit from the fix those two scripts got.

So: dropped the custom Selenium wrapper entirely and switched to the same
common.run_scrape() every other site uses, with 'no_selenium' set so it
always trusts the (already confirmed to work great) static fetch.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Agriculture'
URL = 'http://www.agriculture.gov.lb/Media/News'
OUTPUT_FILE = 'Ministry_of_Agriculture.xlsx'
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
