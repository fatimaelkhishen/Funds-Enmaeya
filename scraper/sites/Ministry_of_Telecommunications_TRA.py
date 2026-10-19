"""
Dedicated scraper for: Ministry of Telecommunications / TRA
Source: https://www.tra.gov.lb/News

NOTE: config previously pointed at the bare homepage (URL listing unclear).
The real news listing lives at https://www.tra.gov.lb/News - a classic
ASP.NET GridView with real article links:

    <a href="NewsDetails.aspx?pageid=3835">The TRA issues Decision No. 7/2026 ...</a>

but its own titles aren't wrapped in a heading tag and its query param is
"?pageid=" (not one of the generic id-style params common.py's
DETAIL_PATH_PATTERN recognizes), so common.extract_candidate_links()'s
generic heuristics would find 0 links here - this pulls links directly via
the stable "NewsDetails.aspx?pageid=" href pattern instead, the same way
UNICEF_Lebanon.py targets its press-release href prefix.

Pagination is also a GridView postback pager with NO real "next" href and no
single stable "next" control across pages - each page number is its own
javascript:__doPostBack('gvNews','Page$N') handler, and pages beyond the
visible range are reached by clicking a "..." control that reveals the next
block of page-number links:

    1 <a href="javascript:__doPostBack('gvNews','Page$2')">2</a> ...
    <a href="javascript:__doPostBack('gvNews','Page$10')">10</a>
    <a href="javascript:__doPostBack('gvNews','Page$11')">...</a>

So this drives one Selenium session, clicking the link whose visible text is
the next page number (or the "..." control to reveal further page-number
links when the target isn't shown yet) rather than using
common.run_scrape_click_pagination()'s single-fixed-"next"-selector model.
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'Ministry of Telecommunications / TRA'
URL = 'https://www.tra.gov.lb/News'
OUTPUT_FILE = 'Ministry_of_Telecommunications_TRA.xlsx'
NOTES = ''

ARTICLE_LINK_SELECTOR = 'a[href*="NewsDetails.aspx?pageid="]'
INITIAL_WAIT_SECONDS = 15
POSTBACK_SETTLE_SECONDS = 4


def extract_tra_links(html, current_url):
    soup = BeautifulSoup(html, "html.parser")
    links, seen = [], set()
    for a in soup.select(ARTICLE_LINK_SELECTOR):
        href = common.normalize_url(current_url, a["href"])
        if href and href not in seen:
            seen.add(href)
            links.append(href)
    return links


def wait_for_results(driver, timeout):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(("css selector", ARTICLE_LINK_SELECTOR))
        )
    except Exception:
        pass


def find_page_link(driver, page_num):
    from selenium.webdriver.common.by import By
    try:
        return driver.find_element(By.XPATH, f"//a[normalize-space(text())='{page_num}']")
    except Exception:
        return None


def click_ellipsis(driver):
    from selenium.webdriver.common.by import By
    try:
        el = driver.find_element(By.XPATH, "//a[normalize-space(text())='...']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def collect_all_links(max_pages, delay):
    driver = common._get_driver()
    if driver is None:
        common.log.error("[%s] Selenium/Chrome not available - cannot drive the GridView pager.", ORG)
        return []

    driver.get(URL)
    wait_for_results(driver, INITIAL_WAIT_SECONDS)

    links = extract_tra_links(driver.page_source, URL)
    common.log.info("[%s] Page 1: %d link(s).", ORG, len(links))

    page = 1
    while page < max_pages:
        next_page = page + 1
        link_el = find_page_link(driver, next_page)
        if link_el is None:
            if not click_ellipsis(driver):
                common.log.info("[%s] No page-%d control or '...' expander found - stopping after %d page(s).", ORG, next_page, page)
                break
            time.sleep(delay)
            link_el = find_page_link(driver, next_page)
            if link_el is None:
                common.log.info("[%s] Page %d still not reachable after expanding pager - stopping.", ORG, next_page)
                break

        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_el)
            driver.execute_script("arguments[0].click();", link_el)
        except Exception as exc:
            common.log.warning("[%s] Could not click page-%d control, stopping: %s", ORG, next_page, exc)
            break

        time.sleep(delay + POSTBACK_SETTLE_SECONDS)
        wait_for_results(driver, POSTBACK_SETTLE_SECONDS)
        page = next_page

        new_links = extract_tra_links(driver.page_source, URL)
        before = len(links)
        for l in new_links:
            if l not in links:
                links.append(l)
        common.log.info("[%s] Page %d: %d link(s) total (%d new).", ORG, page, len(links), len(links) - before)
        if len(links) == before:
            common.log.info("[%s] Clicking to page %d added no new links - stopping.", ORG, page)
            break

    return links


def run(max_pages, max_articles, delay):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom GridView pager driver) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    saved_count = 0
    try:
        all_links = collect_all_links(max_pages, delay)
        common.log.info("[%s] Collected %d unique news link(s).", ORG, len(all_links))

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
