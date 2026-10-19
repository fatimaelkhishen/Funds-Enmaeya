"""
Dedicated scraper for: IMF - Lebanon
Source: https://www.imf.org/en/countries/lbn

NOTE: this country page's "See All Documents" listing (a panel with id/class
"pnlResults") is a JS-driven AJAX panel, not a normal paginated listing:
  - its numbered page links are inert (href="#") - they only work via JS click
  - it has a <button aria-label="Next page" class="pagination-arrow next">
    that re-renders the panel's contents in place instead of navigating to a
    new URL, e.g.:
        <div class="pnlResults"><div class="pages" role="navigation" aria-label="Pagination">
          <div class="pagination"> <span>Page: 1</span>  of 28</div>
          <span class="pnlGroup" role="list">...</span>
          <button type="button" aria-label="Next page" class="pagination-arrow next"></button>
        </div></div>
  - the panel also takes noticeably longer than common.py's default 4-second
    Selenium wait to populate (a plain common.fetch_html() call returned 0
    candidate links on page 1 in an earlier test run), so this file waits
    explicitly for result items / for the page counter to advance instead of
    a fixed sleep.

So instead of common.run_scrape(), this file drives the "Next page" button
itself with one Selenium session, reading result links off each page as:
    <div class="result-item" role="listitem">
      <h6 aria-level="4"><a href="https://www.imf.org/en/publications/...">Title</a></h6>
      <p class="date">February 13, 2026</p>
      <div class="tags">...(href="#" filter buttons, NOT real article links)...</div>
    </div>
via "div.result-item h6 a[href]" specifically, so the href="#" filter-tag
buttons inside the same card are never picked up. It stops once the page
counter ("Page: N of M") reaches M, then hands the full, de-duplicated link
list to common.py's normal fetch -> extract_article -> save-to-xlsx pipeline,
exactly like every other site's scraper.
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

ORG = 'IMF - Lebanon'
URL = 'https://www.imf.org/en/countries/lbn'
OUTPUT_FILE = 'IMF_Lebanon.xlsx'
NOTES = ''

RESULT_LINK_SELECTOR = "div.result-item h6 a[href]"
NEXT_BUTTON_SELECTOR = "button.pagination-arrow.next"
PAGE_COUNTER_SELECTOR = ".pnlResults .pagination"
PAGE_COUNTER_PATTERN = re.compile(r"Page:\s*(\d+)\s*of\s*(\d+)", re.IGNORECASE)

RESULTS_WAIT_SECONDS = 20   # this panel is slower than common.SELENIUM_WAIT_SECONDS
CLICK_SETTLE_SECONDS = 20   # max time to allow the AJAX refresh after clicking "Next"


def extract_imf_links(html, current_url):
    """IMF-specific listing extractor: pull the real document link out of each
    result card's <h6><a>, ignoring the href="#" filter-tag buttons in the
    same card (see docstring above for why the generic extractor is risky here)."""
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.select(RESULT_LINK_SELECTOR):
        href = common.normalize_url(current_url, a["href"])
        if href:
            links.append(href)
    return links


def read_page_counter(driver):
    """Parse the "Page: N of M" counter. Used both to log progress and to
    know when the last page has been reached (no confirmed markup for a
    disabled "Next" button, so this is the reliable stop condition)."""
    try:
        el = driver.find_element("css selector", PAGE_COUNTER_SELECTOR)
    except Exception:
        return None, None
    m = PAGE_COUNTER_PATTERN.search(el.text or "")
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))


def wait_for_results(driver, timeout):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(("css selector", "div.result-item"))
        )
    except Exception:
        pass  # fall through - this round will just find 0 links if it never showed up


def wait_for_page_number(driver, expected_page, timeout):
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, timeout).until(lambda d: read_page_counter(d)[0] == expected_page)
    except Exception:
        pass  # fall through - we'll just re-read whatever's on screen


def collect_all_document_links(max_pages, delay):
    driver = common._get_driver()
    if driver is None:
        common.log.error("[%s] Selenium/Chrome not available - cannot page through the listing.", ORG)
        return []

    all_links, seen = [], set()
    driver.get(URL)
    wait_for_results(driver, RESULTS_WAIT_SECONDS)

    page_num = 1
    while page_num <= max_pages:
        html = driver.page_source
        page_links = extract_imf_links(html, driver.current_url)
        new = 0
        for link in page_links:
            if link not in seen:
                seen.add(link)
                all_links.append(link)
                new += 1
        current, total = read_page_counter(driver)
        common.log.info(
            "[%s] Results page %d%s: found %d link(s) (%d new).",
            ORG, page_num, f" of {total}" if total else "", len(page_links), new,
        )

        if total and current and current >= total:
            common.log.info("[%s] Reached the last page (%d of %d).", ORG, current, total)
            break

        try:
            next_btn = driver.find_element("css selector", NEXT_BUTTON_SELECTOR)
        except Exception:
            common.log.info("[%s] No next-page button found, stopping after %d page(s).", ORG, page_num)
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

        time.sleep(delay)
        wait_for_page_number(driver, page_num + 1, CLICK_SETTLE_SECONDS)
        page_num += 1

    return all_links


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom AJAX pager) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    saved_count = 0
    try:
        all_links = collect_all_document_links(max_pages, delay)
        common.log.info("[%s] Collected %d unique document link(s) across all pages.", ORG, len(all_links))

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
