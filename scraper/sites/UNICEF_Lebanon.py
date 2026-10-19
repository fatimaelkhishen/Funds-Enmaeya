"""
Dedicated scraper for: UNICEF Lebanon
Source: https://www.unicef.org/lebanon/press-center

NOTE: this listing uses a Drupal "load more" AJAX pattern instead of real
pagination - there's no page-2 URL at all. Clicking the button appends more
press-release cards into the SAME page:

    <a href="" class="btn cta--secondary outlined lmlink disable-state"
       id="mosaic-3d85d195-b0c3-48cf-b2ae-8aa36e4ac92f-loadmore-link"
       data-once="drupal-ajax">Load more items</a>

(that "mosaic-<uuid>-loadmore-link" id changes per page load, so this
targets the stable "a.lmlink" class instead.) common.run_scrape() would only
ever see the first batch that loads, exactly like the World Bank Lebanon
listing's "Load More" list - so this file drives one Selenium session,
clicking "Load more items" repeatedly (stopping once the button is gone or a
click stops adding new links) and only then handing everything off to
common.py's normal fetch -> extract_article -> save-to-xlsx pipeline.

Article links are pulled by href prefix ("/lebanon/press-releases/...")
rather than common.py's generic heading-based heuristic, since these teaser
links (data-tracked-component="cta") aren't confirmed to sit inside a real
heading tag - matching on the real content path is more precise here anyway
and avoids ever picking up the "Load more" button itself (its href is empty).
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

import requests
from bs4 import BeautifulSoup

ORG = 'UNICEF Lebanon'
URL = 'https://www.unicef.org/lebanon/press-center'
OUTPUT_FILE = 'UNICEF_Lebanon.xlsx'
NOTES = ''

# FALLBACK (added 2026-10-16): unicef.org loads fine in a normal browser, but
# often serves a bot-check page to scripts, so every article came back
# "Bot/CAPTCHA challenge" and the sheet stayed empty. If the official site
# yields nothing, the same UNICEF Lebanon press releases / reports are read
# from ReliefWeb instead (source = UNICEF, country = Lebanon, ~1,300 items).
FALLBACK_URL = ("https://reliefweb.int/updates?advanced-search=%28C137%29"
                "&search=source.exact%3A%22UN%20Children%27s%20Fund%22")
FALLBACK_LINK_PATTERN = r"reliefweb\.int/report/[^/]+/[^/?#]+"

ARTICLE_LINK_SELECTOR = 'a[href^="/lebanon/press-releases/"]'
LOAD_MORE_SELECTOR = "a.lmlink"
INITIAL_WAIT_SECONDS = 20   # first load: give the listing time to render
CLICK_SETTLE_SECONDS = 15   # time to allow the AJAX "load more" append to finish


def extract_unicef_links(html, current_url):
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
        pass  # fall through - this round will just find 0 links if it never showed up


def click_load_more(driver, delay):
    """Click the "Load more items" button once. Returns False if it's gone
    (all items loaded) or couldn't be clicked."""
    try:
        btn = driver.find_element("css selector", LOAD_MORE_SELECTOR)
    except Exception:
        return False

    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception as exc:
            common.log.warning("[%s] Could not click 'Load more items', stopping: %s", ORG, exc)
            return False

    time.sleep(delay)
    wait_for_results(driver, CLICK_SETTLE_SECONDS)
    return True


def collect_all_links(max_clicks, delay):
    driver = common._get_driver()
    if driver is None:
        common.log.error("[%s] Selenium/Chrome not available - cannot drive 'Load more'.", ORG)
        return []

    driver.get(URL)
    wait_for_results(driver, INITIAL_WAIT_SECONDS)

    links = extract_unicef_links(driver.page_source, URL)
    common.log.info("[%s] Initial batch: %d link(s).", ORG, len(links))

    clicks = 0
    while clicks < max_clicks:
        before = len(links)
        if not click_load_more(driver, delay):
            common.log.info("[%s] 'Load more items' button gone - reached the end after %d click(s).", ORG, clicks)
            break
        clicks += 1
        links = extract_unicef_links(driver.page_source, URL)
        common.log.info("[%s] After click %d: %d link(s) total.", ORG, clicks, len(links))
        if len(links) <= before:
            common.log.info("[%s] Clicking 'Load more' stopped adding new links - stopping.", ORG)
            break

    return links


def run(max_pages, max_articles, delay, manual_captcha=False):
    out_path = os.path.join(common.NEWS_DIR, OUTPUT_FILE)
    org_key = os.path.splitext(OUTPUT_FILE)[0]

    common.log.info("=== %s (custom 'load more' driver) ===", ORG)
    session = requests.Session()
    visited_articles = common.load_state(org_key)
    rp = common.get_robot_parser(URL)
    wb, ws = common.open_or_create_workbook(out_path)

    if manual_captcha:
        common.log.info(
            "[%s] --manual-captcha is on: if an article hits a bot/CAPTCHA "
            "challenge, a visible Chrome window will open and the script will "
            "pause - solve it yourself there, then press Enter in this "
            "terminal to continue.", ORG,
        )

    saved_count = 0
    challenges = 0
    all_links = []
    try:
        all_links = collect_all_links(max_pages, delay)
        common.log.info("[%s] Collected %d unique press-release link(s).", ORG, len(all_links))

        new_links = [l for l in all_links if l not in visited_articles]
        for link in new_links:
            if max_articles and saved_count >= max_articles:
                break
            time.sleep(delay)
            if not common.allowed_by_robots(rp, link):
                common.log.warning("[%s] Article blocked by robots.txt: %s", ORG, link)
                visited_articles.add(link)
                continue
            if manual_captcha:
                art_html, art_method = common.fetch_html(
                    link, session,
                    driver_override=common._get_secondary_driver(),
                    allow_manual_challenge_solve=True,
                )
            else:
                art_html, art_method = common.fetch_html(link, session)
            if not art_html:
                visited_articles.add(link)
                continue
            if common.is_bot_challenge_page(art_html):
                challenges += 1
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
    return saved_count, challenges, len(all_links)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    parser.add_argument(
        "--manual-captcha", action="store_true",
        help=(
            "If an article hits a bot/CAPTCHA challenge, open a visible Chrome "
            "window and pause so you can solve it yourself, then press Enter "
            "to continue. Off by default - only use this for an interactive "
            "run you're watching, never for an unattended/scheduled one."
        ),
    )
    args = parser.parse_args()
    try:
        saved, challenges, n_links = run(args.max_pages, args.max_articles, args.delay,
                                         manual_captcha=args.manual_captcha)
        if saved == 0 and (challenges or n_links == 0):
            common.log.warning("[%s] Official site gave nothing (%d link(s), %d blocked) - "
                               "switching to the ReliefWeb mirror of UNICEF Lebanon.", ORG, n_links, challenges)
            common.run_scrape(ORG, FALLBACK_URL, OUTPUT_FILE, args.max_pages, args.max_articles,
                              args.delay, NOTES, link_pattern=FALLBACK_LINK_PATTERN)
    finally:
        common.shutdown()
