"""
Dedicated scraper for: GIZ Lebanon
Source: https://www.giz.de/en/regions/asia/lebanon

This one is special: the "Projects in this country" list on this page is a
JS-rendered widget (not plain Drupal links). Its pager is a row of <button>
elements (aria-label="previous page" / "Go to page 2" / "next page" / "To the
last page") with no real href to follow, and each project card's link
(<a class="list-item__title" href="/en/projects/...">) sits inside a plain
<div>, not a heading tag - so common.py's generic extract_candidate_links()
heuristic (which looks for headline-length text inside <h1>-<h6>, or URLs
that look like "/news/..." or "/press/...") never picks these up.

So instead of common.run_scrape() / common.run_scrape_click_pagination(),
this file drives the pager itself with one Selenium session: it reads the
project links off the current page, clicks "next page" (stopping once that
button is disabled, i.e. the last page), and repeats - collecting links from
every page - then hands the full, de-duplicated link list to common.py's
normal fetch -> extract_article -> save-to-xlsx pipeline, exactly like every
other site's scraper does.
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'GIZ Lebanon'
URL = 'https://www.giz.de/en/regions/asia/lebanon'
OUTPUT_FILE = 'GIZ_Lebanon.xlsx'
NOTES = ''

# CSS selector for each project card's link within the "Projects in this
# country" list (see docstring above for why the generic extractor misses it).
PROJECT_LINK_SELECTOR = "a.list-item__title[href]"
# The pager's "next page" button (disabled with disabled="" on the last page).
NEXT_BUTTON_SELECTOR = "button.pager__next"


def extract_giz_project_links(html, current_url):
    """Page-specific listing extractor: pull every project link out of the
    <li> cards, ignoring the <button> pager controls entirely."""
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.select(PROJECT_LINK_SELECTOR):
        href = common.normalize_url(current_url, a["href"])
        if href:
            links.append(href)
    return links


def collect_all_project_links(max_pages, delay):
    """Drive the button-based pager with Selenium, collecting project links
    from every page (1, 2, 3, ...) until the "next page" button is disabled
    or there isn't one, or max_pages is reached."""
    driver = common._get_driver()
    if driver is None:
        common.log.error("[%s] Selenium/Chrome not available - cannot page through the listing.", ORG)
        return []

    all_links, seen = [], set()
    driver.get(URL)
    time.sleep(common.SELENIUM_WAIT_SECONDS)

    page_num = 1
    while page_num <= max_pages:
        html = driver.page_source
        page_links = extract_giz_project_links(html, driver.current_url)
        new = 0
        for link in page_links:
            if link not in seen:
                seen.add(link)
                all_links.append(link)
                new += 1
        common.log.info("[%s] Pager page %d: found %d link(s) (%d new).", ORG, page_num, len(page_links), new)

        try:
            next_btn = driver.find_element("css selector", NEXT_BUTTON_SELECTOR)
        except Exception:
            common.log.info("[%s] No next-page button found, stopping after %d page(s).", ORG, page_num)
            break

        if next_btn.get_attribute("disabled") is not None:
            common.log.info("[%s] Next-page button is disabled - reached the last page (%d).", ORG, page_num)
            break

        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
            next_btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", next_btn)
            except Exception as exc:
                common.log.warning("[%s] Could not click next-page button, stopping: %s", ORG, exc)
                break

        time.sleep(delay + common.SELENIUM_WAIT_SECONDS)
        page_num += 1

    return all_links


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom pager) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    saved_count = 0
    try:
        all_links = collect_all_project_links(max_pages, delay)
        common.log.info("[%s] Collected %d unique project link(s) across all pager pages.", ORG, len(all_links))

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
