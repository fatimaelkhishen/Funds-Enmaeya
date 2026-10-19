"""
Dedicated scraper for: World Bank - Lebanon
Source: https://www.worldbank.org/en/news/all?countrycode_exact=LB&apilang=en&lang_exact=English

NOTE: this listing is an Angular app (page-number links carry no href at all,
just "_ngcontent-c0" markers - they're pure JS click handlers), so
common.run_scrape() could read page 1 fine (its generic link extractor already
matches these article cards - confirmed live, 20 candidate links found) but
had no real URL to follow for page 2+, e.g.:

    <ul class="pagination">
      <li class="disabled"><a>« (first)</a></li>
      <li class="disabled"><a>‹ (prev)</a></li>
      <li class="active"><a>1</a></li>
      <li><a>2</a></li> ... <li><a>10</a></li>
      <li><a>› (next)</a></li>
      <li><a>» (last)</a></li>
    </ul>

On page 1 the first two <li> (first/prev) carry class "disabled". By the same
pattern, the last two <li> (next/last) should pick up "disabled" once the real
last page is reached - so this file clicks the "next" <li> (second-to-last -
a single right-chevron, vs. the double-chevron "jump to last" one after it)
repeatedly and stops when THAT <li> is the one marked disabled, rather than
guessing a total page count. It waits for the active page number to actually
change after each click (Angular re-renders via an API call, not a full page
load) before re-reading the list. Article extraction reuses common.py's normal
extract_candidate_links() unchanged, since it already works on this page as-is;
only the pager needed page-specific handling.
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'World Bank - Lebanon'
URL = 'https://www.worldbank.org/en/news/all?countrycode_exact=LB&apilang=en&lang_exact=English'
OUTPUT_FILE = 'World_Bank_Lebanon.xlsx'
NOTES = ''

PAGINATION_LIST_SELECTOR = "ul.pagination"
ACTIVE_PAGE_SELECTOR = "ul.pagination li.active a"
RESULTS_WAIT_SECONDS = 20   # first load: this app is slower than common.SELENIUM_WAIT_SECONDS (4s)
CLICK_SETTLE_SECONDS = 20   # max time to allow the Angular re-render after clicking "next"


def wait_for_results(driver, timeout):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(("css selector", ACTIVE_PAGE_SELECTOR))
        )
    except Exception:
        pass  # fall through - this round will just find 0 links if it never showed up


def read_active_page(driver):
    try:
        el = driver.find_element("css selector", ACTIVE_PAGE_SELECTOR)
        return (el.text or "").strip()
    except Exception:
        return None


def get_pagination_items(driver):
    try:
        ul = driver.find_element("css selector", PAGINATION_LIST_SELECTOR)
        return ul.find_elements("css selector", "li")
    except Exception:
        return []


def click_next_page(driver, delay):
    """Click the "next" <li> (second-to-last in the pager). Returns False if
    it's not found, or it's the one carrying class "disabled" (last page)."""
    items = get_pagination_items(driver)
    if len(items) < 2:
        return False
    next_li = items[-2]
    classes = (next_li.get_attribute("class") or "")
    if "disabled" in classes:
        return False

    before = read_active_page(driver)
    try:
        next_a = next_li.find_element("css selector", "a")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_a)
        next_a.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", next_li.find_element("css selector", "a"))
        except Exception as exc:
            common.log.warning("[%s] Could not click next-page control, stopping: %s", ORG, exc)
            return False

    time.sleep(delay)
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, CLICK_SETTLE_SECONDS).until(lambda d: read_active_page(d) != before)
    except Exception:
        pass  # fall through - we'll just re-read whatever's on screen
    return True


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom Angular pager) ===", ORG)
    driver = common._get_driver()
    if driver is None:
        common.log.error("[%s] Selenium/Chrome not available - cannot page through the listing.", ORG)
        return

    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    saved_count = 0
    page_count = 0

    try:
        if not common.allowed_by_robots(rp, URL):
            common.log.warning("[%s] Blocked by robots.txt: %s", ORG, URL)
            return

        driver.get(URL)
        wait_for_results(driver, RESULTS_WAIT_SECONDS)

        while page_count < max_pages:
            page_count += 1
            html = driver.page_source
            if common.is_bot_challenge_page(html):
                common.log.warning("[%s] Bot/CAPTCHA challenge page, stopping.", ORG)
                break

            soup = BeautifulSoup(html, "html.parser")
            article_links = common.extract_candidate_links(soup, URL)
            new_links = [l for l in article_links if l not in visited_articles]
            active_page = read_active_page(driver)
            common.log.info(
                "[%s] Listing page %s (loop %d): found %d candidate links (%d new).",
                ORG, active_page or "?", page_count, len(article_links), len(new_links),
            )

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

            if not click_next_page(driver, delay):
                common.log.info("[%s] No further pagination found after %d page(s).", ORG, page_count)
                break
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
