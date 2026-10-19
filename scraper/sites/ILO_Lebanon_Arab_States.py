"""
Dedicated scraper for: ILO Lebanon / Arab States
Source: https://www.ilo.org/all-news-for-country?cf0=946

NOTE: the original config.csv URL (the "Lebanon" country portal page) only shows
the 2 most recent items and has no pager of its own - the real, paginated news
listing for Lebanon is https://www.ilo.org/all-news-for-country?cf0=946 (18 pages).

That pager uses real <a href> links (no Selenium/browser needed), but with markup
common.py's generic pager detector (common.find_next_page_url) can't recognize: the
"next" link has no rel="next", no aria-label, and its visible text is "Go to next
page" (not a bare "next"/"more"/etc.), e.g.

    <a class="ilo--pagination--next-page ilo--pagination--link " href="?cf0=946&page=1"
       title="Go to next page">Go to next page</a>

and a disabled/last-page state is marked with a CSS class
("ilo--pagination--disable") rather than an aria-disabled/disabled attribute. So
this file finds the "next page" link itself; article-link extraction still uses
common.py's normal extract_candidate_links() (it already matches these
/resource/news/... and /resource/article/... article URLs fine), and everything
else reuses common.py's fetch/extract/save pipeline exactly like every other
site's scraper.
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'ILO Lebanon / Arab States'
URL = 'https://www.ilo.org/all-news-for-country?cf0=946'
OUTPUT_FILE = 'ILO_Lebanon_Arab_States.xlsx'
NOTES = ''

# The pager's "next page" link and the CSS class marking it (or first/prev) disabled.
NEXT_LINK_SELECTOR = "a.ilo--pagination--next-page"
DISABLED_CLASS = "ilo--pagination--disable"


def find_ilo_next_page(soup, current_url):
    """ILO-specific pager lookup: common.find_next_page_url() can't see this
    pager (no rel="next", no aria-label, text is "Go to next page" not a bare
    "next"), so match its actual CSS class and disabled-state class instead."""
    next_a = soup.select_one(NEXT_LINK_SELECTOR)
    if not next_a:
        return None
    classes = next_a.get("class") or []
    if DISABLED_CLASS in classes:
        return None
    href = next_a.get("href")
    if not href:
        return None
    return common.normalize_url(current_url, href)


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom pager) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    current_url = URL
    visited_pages = set()
    saved_count = 0
    page_count = 0

    try:
        while current_url and page_count < max_pages:
            if current_url in visited_pages:
                break
            if not common.allowed_by_robots(rp, current_url):
                common.log.warning("[%s] Blocked by robots.txt: %s", ORG, current_url)
                break

            common.log.info("[%s] Listing page %d: %s", ORG, page_count + 1, current_url)
            html, method = common.fetch_html(current_url, session)
            visited_pages.add(current_url)
            page_count += 1
            if not html:
                common.log.warning("[%s] Could not fetch listing page, stopping.", ORG)
                break
            if common.is_bot_challenge_page(html):
                common.log.warning("[%s] Bot/CAPTCHA challenge page, stopping: %s", ORG, current_url)
                break

            soup = BeautifulSoup(html, "html.parser")
            article_links = common.extract_candidate_links(soup, current_url)
            new_links = [l for l in article_links if l not in visited_articles]
            common.log.info("[%s] (%s) Found %d candidate links (%d new).", ORG, method, len(article_links), len(new_links))

            for link in new_links:
                if max_articles and saved_count >= max_articles:
                    break
                time.sleep(delay)
                if not common.allowed_by_robots(rp, link):
                    common.log.warning("[%s] Article blocked by robots.txt: %s", ORG, link)
                    visited_articles.add(link)
                    continue
                art_html, art_method = common.fetch_html(link, session)
                if not art_html:
                    visited_articles.add(link)
                    continue
                if common.is_bot_challenge_page(art_html):
                    common.log.warning("[%s] Bot/CAPTCHA challenge page, skipping (will retry later): %s", ORG, link)
                    continue
                visited_articles.add(link)
                title, date, text = common.extract_article(art_html, link)
                if not text or len(text) < 50:
                    common.log.warning("[%s] Little/no extractable text even via %s: %s", ORG, art_method, link)
                    continue
                common.append_row(ws, title, link, date, text)
                wb.save(out_path)
                saved_count += 1
                common.log.info("[%s] Saved article %d: %s", ORG, saved_count, (title or link)[:80])

            common.save_state(org_key, visited_articles)

            if max_articles and saved_count >= max_articles:
                common.log.info("[%s] Reached --max-articles limit (%d).", ORG, max_articles)
                break

            time.sleep(delay)
            next_url = find_ilo_next_page(soup, current_url)
            if not next_url or next_url in visited_pages:
                common.log.info("[%s] No further pagination found after %d page(s).", ORG, page_count)
                break
            current_url = next_url
    finally:
        wb.save(out_path)

    common.log.info("[%s] Done. %d new article(s) saved to %s.", ORG, saved_count, OUTPUT_FILE)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        run(args.max_pages, args.max_articles, args.delay)
    finally:
        common.shutdown()
