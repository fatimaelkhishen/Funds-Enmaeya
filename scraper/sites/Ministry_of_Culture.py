"""
Dedicated scraper for: Ministry of Culture
Source: http://culture.gov.lb/ar/News

NOTE: the listing page itself never contains real article links in its
static HTML - it's an empty shell with a "Loading..." spinner, and the
actual news cards are injected by an inline script that calls
loadevents(0) on document-ready, which does a jQuery .load() AJAX call to:

    http://culture.gov.lb/ajaxpost/news.php?sid=10&catid=&page=N

Two things follow from that:
  1. The pager itself is 100% javascript:loadevents(N) links - no real
     href/URL for "next page" at all, so common.find_next_page_url() can
     never advance past page 1 no matter what.
  2. BUT the ajaxpost/news.php endpoint turns out to be a perfectly normal,
     directly-fetchable GET endpoint (confirmed by fetching it directly) -
     it's not gated behind any session/token, and returns plain HTML with
     real absolute article links (http://culture.gov.lb//News/<slug> - note
     the CMS's own double slash, harmless) plus the same loadevents(N)
     pager repeated. That means this scraper doesn't need Selenium at all:
     it can page straight through the AJAX endpoint with plain requests,
     which is both simpler and much faster than driving a browser.

The last page number is read off the "»»" (last-page) pager link's
loadevents(N) argument on whatever page is fetched, so the loop knows when
to stop without guessing.
"""
import argparse
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'Ministry of Culture'
LISTING_URL = 'http://culture.gov.lb/ar/News'
AJAX_URL = 'http://culture.gov.lb/ajaxpost/news.php'
OUTPUT_FILE = 'Ministry_of_Culture.xlsx'
NOTES = ''

LAST_PAGE_PATTERN = re.compile(r"loadevents\((\d+)\)")


def fetch_ajax_page(session, page):
    try:
        resp = session.get(
            AJAX_URL, params={"sid": 10, "catid": "", "page": page},
            timeout=common.REQUEST_TIMEOUT, headers={"User-Agent": common.USER_AGENT},
        )
        if resp.status_code == 200:
            return resp.text
        common.log.warning("[%s] AJAX page %d returned HTTP %s", ORG, page, resp.status_code)
    except requests.RequestException as exc:
        common.log.warning("[%s] AJAX page %d request failed: %s", ORG, page, exc)
    return None


def extract_links_and_last_page(html):
    soup = BeautifulSoup(html, "html.parser")
    links, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("javascript:"):
            continue
        norm = common.normalize_url(AJAX_URL, href)
        if norm and norm not in seen:
            seen.add(norm)
            links.append(norm)

    last_page = None
    for a in soup.find_all("a", href=True):
        if a.get_text(strip=True) == "»»":
            m = LAST_PAGE_PATTERN.search(a["href"])
            if m:
                last_page = int(m.group(1))
            break
    return links, last_page


def collect_all_links(session, max_pages, delay):
    all_links, seen = [], set()

    html = fetch_ajax_page(session, 0)
    if not html:
        common.log.error("[%s] Could not fetch the AJAX news endpoint at all.", ORG)
        return all_links

    links, last_page = extract_links_and_last_page(html)
    for l in links:
        if l not in seen:
            seen.add(l)
            all_links.append(l)
    common.log.info("[%s] Page 0: %d link(s). Last page reported as %s.", ORG, len(links), last_page)

    if last_page is None:
        last_page = 0

    page = 1
    while page <= last_page and page < max_pages:
        time.sleep(delay)
        html = fetch_ajax_page(session, page)
        if not html:
            common.log.warning("[%s] AJAX page %d failed to fetch - stopping.", ORG, page)
            break
        links, _ = extract_links_and_last_page(html)
        new_count = 0
        for l in links:
            if l not in seen:
                seen.add(l)
                all_links.append(l)
                new_count += 1
        common.log.info("[%s] Page %d: %d link(s) (%d new).", ORG, page, len(links), new_count)
        page += 1

    return all_links


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom AJAX-endpoint pager driver) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(LISTING_URL)
    wb, ws = common.open_or_create_workbook(out_path)

    saved_count = 0
    try:
        all_links = collect_all_links(session, max_pages, delay)
        common.log.info("[%s] Collected %d unique news link(s) across all pages.", ORG, len(all_links))

        new_links = [l for l in all_links if l not in visited_articles]
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
