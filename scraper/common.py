"""
common.py
---------
Shared helpers used by every per-organization scraper in scraper/sites/.

Design:
  * Each organization has its own tiny script (scraper/sites/<Org>.py) that
    just declares ORG / URL / OUTPUT_FILE and calls run_scrape(...).
  * All the actual crawling logic lives here so it isn't duplicated 45 times.
  * Fetching tries a plain HTTP request first (fast, no browser needed). If
    the page looks like it needs JavaScript to render its article links or
    its "next page" control (a common issue on modern org/IFI sites and a
    few ministry sites), it automatically retries the SAME url with a
    headless Chrome browser (Selenium) instead. You don't have to know in
    advance which sites are static and which are dynamic - this is detected
    automatically per page.
  * Output is written to one .xlsx file per organization (in the News
    folder, one level above scraper/) with columns:
    Title | URL | Date | Text | Scraped At
"""

import csv
import json
import logging
import os
import re
import sys
import time
import urllib.robotparser
from urllib.parse import urljoin, urlparse, urldefrag
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

try:
    import trafilatura
    HAVE_TRAFILATURA = True
except ImportError:
    HAVE_TRAFILATURA = False

# ---------------------------------------------------------------------------
# Paths / constants
# ---------------------------------------------------------------------------

SCRAPER_DIR = os.path.dirname(os.path.abspath(__file__))
NEWS_DIR = os.path.dirname(SCRAPER_DIR)           # the "News" folder
STATE_DIR = os.path.join(SCRAPER_DIR, "state")    # remembers what's already scraped
LOG_PATH = os.path.join(SCRAPER_DIR, "scrape_log.txt")

# --- Test mode (set by `python run_all.py --test`, or SCRAPER_TEST=1) --------
# Writes every .xlsx into News/_test_output/ and uses a separate, throw-away
# state folder, so a test run never touches your real spreadsheets or the
# "already scraped" memory in scraper/state/.
TEST_MODE = os.environ.get("SCRAPER_TEST", "").strip() not in ("", "0", "false", "False")
if TEST_MODE:
    NEWS_DIR = os.path.join(NEWS_DIR, "_test_output")
    STATE_DIR = os.path.join(SCRAPER_DIR, "state_test")
    LOG_PATH = os.path.join(SCRAPER_DIR, "scrape_log_test.txt")
    os.makedirs(NEWS_DIR, exist_ok=True)

# --- Dry-run / pagination check (python check_pagination.py) ---------------
# Only walks the LISTING pages: logs every page visited and every article link
# found, but never opens an article, never writes an .xlsx and never touches
# scraper/state/. Works for every site script without editing them: in this
# mode load_state() returns a set that says "already scraped" for every link
# (so no script fetches any article) while quietly recording each link it was
# asked about, which is how the link counts in the report are produced.
DRY_RUN = os.environ.get("SCRAPER_DRY_RUN", "").strip() not in ("", "0", "false", "False")
if DRY_RUN:
    NEWS_DIR = os.path.join(NEWS_DIR, "_dry_run_unused")
    STATE_DIR = os.path.join(SCRAPER_DIR, "state_dry")
    LOG_PATH = os.path.join(SCRAPER_DIR, "scrape_log_dry.txt")

# --- Unattended mode (set by run_all.py, GitHub Actions, Task Scheduler) -----
# Never block waiting for keyboard input (the manual-CAPTCHA pause) when
# nobody is at the terminal - that used to hang run_all.py forever on the
# force_visible_browser sites.
UNATTENDED = (
    os.environ.get("SCRAPER_UNATTENDED", "").strip() not in ("", "0", "false", "False")
    or os.environ.get("CI") == "true"
    or not (sys.stdin and sys.stdin.isatty())
)

# --- Incremental mode (run_all.py --incremental) ----------------------------
# Listing pages are newest-first. Once N listing pages in a row contain no
# new article links, everything older was already scraped on a previous run,
# so stop instead of re-walking hundreds of old pages every week. 0 = off.
try:
    STOP_AFTER_KNOWN_PAGES = int(os.environ.get("SCRAPER_STOP_AFTER_KNOWN_PAGES", "0") or 0)
except ValueError:
    STOP_AFTER_KNOWN_PAGES = 0

# Excel files are re-written to disk every SAVE_EVERY saved articles (and
# always at the end of each listing page / run) instead of after every single
# article - re-saving an 8 MB workbook thousands of times was the slowest part
# of a big run.
SAVE_EVERY = 10

# Give up on a site for this run after this many article fetches in a row came
# back as bot-check/CAPTCHA/error pages - the site is blocking us right now, so
# trying the remaining links only burns time. Those links stay un-visited and
# are retried on the next run. 0 = never give up.
MAX_CHALLENGE_STREAK = 5

# NOTE (2026-07-21): this used to end in " EnmaeyaNewsArchiver/1.0" - a custom
# product token appended after an otherwise-real Chrome UA string. That tail
# is exactly the kind of thing a UA-sniffing WAF/edge (Akamai etc., common on
# UN-agency sites) flags as "not a real browser" and soft-blocks by serving a
# generic error page instead of a hard 403 - which is consistent with what
# UNESCO's /en/countries/lb/latestnews did: our session got HTTP 500 three
# times in a row (with retries/backoff) AND Selenium (same UA) got the same
# result, while an unrelated fetch of the identical URL with a plain
# unmodified browser UA succeeded immediately with full content. Removed the
# suffix so every scraper's requests - static and Selenium alike, since
# Selenium's driver options also read this same constant - present as a
# normal browser instead of self-identifying as a scraping tool.
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
REQUEST_TIMEOUT = 20
DEFAULT_MAX_PAGES = 300
MIN_LINK_TEXT_LEN = 12
SELENIUM_WAIT_SECONDS = 4       # time to let JS render before reading page_source

SKIP_DOMAINS = {
    "facebook.com", "www.facebook.com", "m.facebook.com",
    "twitter.com", "x.com", "instagram.com", "linkedin.com",
    "youtube.com", "youtu.be", "t.me",
}
SKIP_EXTENSIONS = (
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".css", ".js", ".zip",
    ".mp4", ".mp3", ".ico", ".woff", ".woff2",
)

NEXT_TEXT_PATTERNS = re.compile(
    r"^\s*(next|more|older|»|›|>>|>|التالي|السابق|المزيد|تحميل المزيد|load more)\s*$",
    re.IGNORECASE,
)

# Some Sitefinity/.NET ministry sites (moe.gov.lb's news listing, confirmed
# 2026-07-22) only hyperlink a short "read more"/"<< للمزيد" label on each
# teaser card - the actual headline text next to it is plain, unlinked text.
# Such a link is always under MIN_LINK_TEXT_LEN and never sits inside a
# heading tag, so the normal looks_like_headline check rejects it outright,
# and its target URL's slug is usually a transliterated/Arabic phrase with no
# English "news"/"press"/etc. keyword, so DETAIL_PATH_PATTERN below doesn't
# catch it either - net result: every real article on the page gets silently
# dropped, 0 links found. This is NOT the same thing as NEXT_TEXT_PATTERNS
# above (a *pagination* "go to next listing page" control) - a "read more"
# link like this lives inside one specific article's own teaser card and
# points straight at that one article's detail page, so it's fine to accept
# unconditionally by text alone (still subject to every other check: same
# domain, not nav, not pagination, etc.) rather than needing the
# heading/length heuristic.
READ_MORE_TEXT_PATTERN = re.compile(
    r"(read\s*more|continue\s*reading|للمزيد|المزيد|اقرأ\s*المزيد|قراءة\s*المزيد|"
    r"تابع\s*القراءة|التفاصيل|تفاصيل)",
    re.IGNORECASE,
)

# A link whose URL contains a pagination query param is a "page N of the
# listing" control, never an individual article - even if its visible text
# happens to look like a headline (some sites reuse a "latest headline"
# teaser widget that links back into the paginated listing).
PAGINATION_QUERY_PATTERN = re.compile(r"[?&](page|paged|pg)=\d+", re.IGNORECASE)

# Same idea, but for WordPress-style *path* pagination ("/page/2/") instead of
# a query string. Without this, a URL like ".../category/news/page/2/" gets
# past PAGINATION_QUERY_PATTERN (no "?page=" to match) and then wrongly
# matches DETAIL_PATH_PATTERN below too - "news/" immediately followed by
# 3+ word characters ("page") looks like a real article slug to that pattern,
# so numbered pager links were being scraped as if they were articles.
# Also covers the colon-separated variant some Sitefinity/.NET ministry
# sites use ("/ar/Media/news/page:2") - same false-positive risk.
PAGINATION_PATH_PATTERN = re.compile(r"/page[:/]\d+/?(?:[?#]|$)", re.IGNORECASE)

# Generic site-navigation sections (About, organizational structure, job
# postings, resource libraries, raw API endpoints, search) that show up in
# the nav/menu on every page of a site and can get mistaken for "long
# headline links" by a text-length heuristic alone.
NAV_EXCLUDE_PATTERN = re.compile(
    r"^/(about|who-we-are|clusters|resources|jobs|careers|contact|search|api|sitemap)(/|\?|$)",
    re.IGNORECASE,
)

# A URL that looks like an actual article/press-release detail page: a real
# slug or id after the news/press keyword, a dated path, an id-style query
# param, or a Drupal-style "/<node-id>-slug" path (e.g. UN sites: articles
# live at "/en/318791-some-headline-slug", with no "news"/"press" segment and
# no heading tag around the link - just a bare numeric node ID followed by a
# hyphen, as opposed to the bare listing root or a pagination link).
#
# The keyword can be plural ("articles/", "stories/") and/or hyphenated
# ("press-releases/", "media-releases/") - UNESCO's "/en/articles/..." and
# UNDP's "/lebanon/press-releases/..." paths were both slipping past this
# check when it only matched the bare singular word immediately before "/".
# "report(s)" covers ReliefWeb, whose individual items live at
# "/report/lebanon/<slug>" - distinct from "news/press/updates" but the same
# kind of individual-document URL, not a listing or nav link.
DETAIL_PATH_PATTERN = re.compile(
    r"(news|press|articles?|stor(?:y|ies)|updates?|media|reports?)(-[\w]+)?/[\w\-%]{3,}"
    r"|[?&](id|newsid|nid|articleid|item)=\d+"
    r"|/\d{4}/\d{1,2}/[\w\-]+"
    r"|/\d{4,}(?:[/?\-]|$)",
    re.IGNORECASE,
)

# Signature phrases from common bot-protection / CAPTCHA interstitial pages
# (Cloudflare, Akamai, etc.) AND from generic server/application error pages
# (IIS/ASP.NET's default "Server Error in '/' Application" screen, plain
# 404/500/502/503 text, etc.). Neither is real article content - if a fetched
# page matches one of these, never save it (a broken/mistyped URL on the
# Ministry of Economy and Trade site once got its literal ASP.NET error page
# saved to the spreadsheet as a real "article" before this list covered
# generic error pages, not just bot challenges).
# FIX (2026-10-16): these used to be searched against the WHOLE raw HTML,
# including <script> blocks, JSON settings and hidden templates. Many real
# article pages (presidency.gov.lb, economy.gov.lb, unicef.org ...) carry
# strings like "Page not found" or "unexpected error" inside their JS
# translations, so every real article got thrown away as a "Bot/CAPTCHA
# challenge page". Now split into two groups and only checked against the
# visible text - see is_bot_challenge_page().
STRONG_CHALLENGE_PATTERNS = re.compile(
    r"(performing security verification|checking your browser before accessing|"
    r"just a moment\.\.\.|attention required.{0,15}cloudflare|verify you are human|"
    r"ddos protection by|enable javascript and cookies to continue|"
    r"please stand by, while we are checking your browser|access denied|"
    r"request unsuccessful\. incapsula|you have been blocked)",
    re.IGNORECASE,
)
ERROR_PAGE_PATTERNS = re.compile(
    r"(server error in '/' application|http error 4\d\d|http error 5\d\d|"
    r"the resource cannot be found|the page cannot be found|page not found|"
    r"internal server error|502 bad gateway|"
    r"503 service (temporarily )?unavailable|504 gateway time-?out|"
    r"the website encountered an unexpected error|"
    r"an unexpected error (has )?occurred.{0,30}try again|out of memory\.)",
    re.IGNORECASE,
)
# Kept for backwards compatibility with any script that still imports it.
BOT_CHALLENGE_PATTERNS = re.compile(
    STRONG_CHALLENGE_PATTERNS.pattern + "|" + ERROR_PAGE_PATTERNS.pattern, re.IGNORECASE
)

# Incapsula/Imperva-style JS "challenge" redirect page (seen on finance.gov.lb,
# 2026-07-22): unlike the interstitials BOT_CHALLENGE_PATTERNS catches above,
# this one has no human-readable "checking your browser"-type text at all - it's
# just a near-empty page (title = the site's short name, e.g. "MOF") whose only
# link is a single long base64url-ish token used as the whole path, e.g.
# href="/GPnSbtkDvEcpbfpvfq-eVTQjBJW9BDZioAzawZXkD24=.html" - a tiny bootstrap
# script computes a cookie/token and redirects to the real page, which a plain
# requests session (no JS) or a fingerprinted/rate-limited Selenium session
# never reaches. Real article/nav links on these ministry sites are always
# hyphenated human words (Arabic or English), never a single 30+ char
# mixed-case/base64url blob - so this is a safe structural signature to flag
# even though it contains none of the phrases above. Capped to short pages only
# (len(html) < 3000) so a legitimately huge real page that happens to link one
# base64-named attachment somewhere isn't misflagged.
INCAPSULA_JS_REDIRECT_PATTERN = re.compile(r'href="/[A-Za-z0-9_-]{30,60}=?\.html"')

EXCEL_HEADERS = ["Title", "URL", "Date", "Text", "Scraped At"]

os.makedirs(STATE_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout), logging.FileHandler(LOG_PATH, encoding="utf-8")],
)
log = logging.getLogger("news_scraper")

# Domains that have already been confirmed to need Selenium during this run,
# so we don't waste time retrying the static fetch for every page on that site.
_dynamic_domains = set()

# Lazily-created Selenium driver (only spun up the first time it's actually needed)
_driver = None

# Second, independent Selenium driver - only used for fetching individual
# article pages during click-pagination mode, so it never disturbs the
# primary driver's position on the paginated listing page. See
# _get_secondary_driver() for why this exists.
_secondary_driver = None


# ---------------------------------------------------------------------------
# Small URL / domain helpers
# ---------------------------------------------------------------------------

def domain_of(url):
    netloc = urlparse(url).netloc.lower()
    # Strip an explicit port (e.g. "www.fao.org:443") so the same host with
    # or without a default port isn't treated as two different domains -
    # some sites' CMS inconsistently bakes ":443" into some article URLs.
    if ":" in netloc:
        netloc = netloc.split(":", 1)[0]
    return netloc


def is_skippable_domain(url):
    d = domain_of(url)
    return any(d == sd or d.endswith("." + sd) for sd in SKIP_DOMAINS)


# Invisible Unicode bidi-control characters (LRM/RLM, LRE/RLE/PDF, LRO/RLO,
# LRI/RLI/FSI/PDI, ALM) - some Arabic-CMS themes embed these in slug/href
# text for RTL display purposes, and they occasionally leak into the actual
# href attribute value. A URL with one of these baked into the middle of it
# is invalid and 500s/404s server-side (seen on the Ministry of Economy and
# Trade site: ".../مدير-‏عام-الاقتصاد..." - note the invisible mark
# right after "مدير-") even though it looks like a normal link when printed.
BIDI_CONTROL_CHARS_PATTERN = re.compile(
    "[" + "".join(chr(c) for c in (
        0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E,
        0x2066, 0x2067, 0x2068, 0x2069, 0x061C,
    )) + "]"
)


def normalize_url(base, href):
    if not href:
        return None
    href = BIDI_CONTROL_CHARS_PATTERN.sub("", href).strip()
    if href.startswith(("mailto:", "tel:", "javascript:", "#")):
        return None
    absolute = urljoin(base, href)
    absolute, _ = urldefrag(absolute)
    return absolute


def get_robot_parser(url):
    parsed = urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = urllib.robotparser.RobotFileParser()
    try:
        # Fetch robots.txt ourselves with a real browser User-Agent instead of
        # letting RobotFileParser.read() do it with urllib's default UA. Some
        # sites (e.g. worldbank.org, behind bot-protection) return HTTP 403 to
        # generic/default user agents on robots.txt specifically, and Python's
        # robotparser treats a 403 there as "disallow everything" - which is
        # wrong when the actual robots.txt (fetched properly) allows the page.
        resp = requests.get(robots_url, timeout=10, headers={"User-Agent": USER_AGENT})
        if resp.status_code == 200:
            rp.parse(resp.text.splitlines())
        elif resp.status_code in (401, 403):
            # Truly restricted robots.txt - be conservative and treat as "no rules found"
            # (RFC 9309 guidance) rather than silently allowing everything.
            log.warning("robots.txt at %s returned HTTP %s - treating as unrestricted.", robots_url, resp.status_code)
            return None
        else:
            return None
    except Exception:
        return None
    return rp


def allowed_by_robots(rp, url):
    if rp is None:
        return True
    try:
        return rp.can_fetch(USER_AGENT, url)
    except Exception:
        return True


# ---------------------------------------------------------------------------
# Fetching: plain requests, with automatic Selenium fallback for JS pages
# ---------------------------------------------------------------------------

def _build_chrome_driver(headless=True):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    try:
        # Preferred: auto-downloads the right chromedriver build for you.
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = None  # fall back to Chrome/Chromedriver already on PATH

    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument(f"user-agent={USER_AGENT}")
    return webdriver.Chrome(service=service, options=options) if service else webdriver.Chrome(options=options)


def _get_driver():
    """Lazily create a single headless Chrome driver, reused for the whole run."""
    global _driver
    if _driver is not None:
        return _driver
    try:
        _driver = _build_chrome_driver()
        return _driver
    except Exception as exc:
        log.error(
            "Could not start a Selenium/Chrome browser (%s). Install Google "
            "Chrome and run 'pip install selenium webdriver-manager' to enable "
            "scraping of JavaScript-rendered pages.", exc,
        )
        return None


def _get_secondary_driver():
    """A second, independent Chrome session (visible, not headless), used
    only to fetch individual article pages during click-pagination mode.

    run_scrape_click_pagination keeps the PRIMARY driver (_get_driver())
    parked on the paginated listing page for the whole run, since that's the
    only place the "N results loaded so far" state lives - the site has no
    real "page 3" URL to revisit, just an in-page button that appends more
    results each click. If fetching an article's full text ever reused that
    same driver (e.g. because fetch_html's static request got a 403 and fell
    back to Selenium), it would navigate the listing tab away to the article
    URL and lose that accumulated state - the next "click Load More" would
    then silently fail (button not found) because the driver is no longer
    even on the listing page. This isolates article fetches onto their own
    tab/session so the listing driver is never touched.

    Not headless (unlike the primary driver) so that, when paired with
    allow_manual_challenge_solve on fetch_html(), a real window is on-screen
    for you to solve a bot-detection challenge in yourself if one comes up.
    """
    global _secondary_driver
    if _secondary_driver is not None:
        return _secondary_driver
    try:
        _secondary_driver = _build_chrome_driver(headless=False)
        return _secondary_driver
    except Exception as exc:
        log.error(
            "Could not start a second Selenium/Chrome browser for article "
            "fetches (%s).", exc,
        )
        return None


def _fetch_dynamic(url, driver_override=None, allow_manual_challenge_solve=False):
    driver = driver_override if driver_override is not None else _get_driver()
    if driver is None:
        return None
    try:
        driver.get(url)
        time.sleep(SELENIUM_WAIT_SECONDS)
        html = driver.page_source
        if allow_manual_challenge_solve and not UNATTENDED and is_bot_challenge_page(html):
            # A real, visible Chrome window (see _get_secondary_driver) is showing
            # this page - pause here so YOU can solve whatever challenge it's showing
            # yourself, in that window. We never solve/click through it ourselves;
            # this just waits for you and then re-reads the page once you say go.
            log.warning(
                "[%s] Bot/CAPTCHA challenge shown - a Chrome window should be open "
                "on your screen. Solve it there yourself, then come back here.", url,
            )
            input(
                f"\n    >>> Bot/CAPTCHA challenge for: {url}\n"
                "    >>> Solve it yourself in the Chrome window, then press Enter here to continue... "
            )
            html = driver.page_source
        return html
    except Exception as exc:
        log.warning("Selenium fetch failed for %s: %s", url, exc)
        return None


def _fetch_static(url, session, retries=2):
    # A 5xx (or 429) is usually a transient hiccup on the ORIGIN server, not
    # something wrong with our request - e.g. UNESCO's site intermittently
    # returns its own "The website encountered an unexpected error. Try again
    # later." page for a few seconds. Previously any non-200 status returned
    # None immediately with no retry at all (the `retries` param only ever
    # covered network-level exceptions), so a fetch that landed exactly during
    # one of these blips would give up on the real page for that whole run and
    # incorrectly fall through to Selenium - which just hits the same
    # still-erroring page a few seconds later and "confirms" 0 real content.
    # Retrying with a short backoff gives a momentary server error time to
    # clear before we conclude the page is genuinely broken/JS-only.
    # 4xx (other than 429) is a real client-side outcome (not found, blocked,
    # etc.) and retrying it wouldn't help, so those still fail fast.
    for attempt in range(retries + 1):
        try:
            resp = session.get(url, timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT})
            if resp.status_code == 200:
                return resp.text
            if resp.status_code == 429 or resp.status_code >= 500:
                log.warning(
                    "HTTP %s for %s (attempt %s/%s)%s", resp.status_code, url,
                    attempt + 1, retries + 1,
                    "" if attempt < retries else " - giving up after retries",
                )
                if attempt < retries:
                    time.sleep(3)
                    continue
                return None
            log.warning("HTTP %s for %s", resp.status_code, url)
            return None
        except requests.RequestException as exc:
            log.warning("Request error (%s/%s) for %s: %s", attempt + 1, retries + 1, url, exc)
            time.sleep(2)
    return None


def _looks_insufficient(html, url):
    """Heuristic: does this HTML look like it needs JS to show real content?

    A couple of nav/footer links surviving a static fetch (e.g. a "Contact us"
    or sidebar link that happens to match the candidate-link heuristic) isn't
    enough to call a page "rendered" - real listing pages have several
    headline links AND a reasonable amount of body text. Sites that load
    their results via a JS/AJAX call after the page loads (seen on a few IFI
    "filtered news" pages) tend to come back from a plain request with just
    the page chrome: a couple of nav links and very little text.
    """
    if not html:
        return True
    soup = BeautifulSoup(html, "html.parser")
    links = extract_candidate_links(soup, url)
    if len(links) == 0:
        # Zero real article links is disqualifying on its own - some faceted
        # search pages (e.g. filter sidebars with hundreds of category/country
        # links) render plenty of *text* statically while the actual results
        # grid loads separately via JS/AJAX, so a text-length check alone
        # would wrongly call this page "sufficient".
        return True
    body_text = soup.get_text(" ", strip=True)
    return len(links) < 3 and len(body_text) < 1500


def fetch_html(url, session, force_dynamic=False, driver_override=None, allow_manual_challenge_solve=False, no_selenium=False):
    """
    Returns rendered HTML for url, using plain requests unless this domain
    was EXPLICITLY forced into Selenium mode (via config.csv's
    "force_selenium"/"force_visible_browser" notes, which populate
    _dynamic_domains up front), or this particular URL's static fetch looks
    insufficient - in which case it transparently retries with headless
    Chrome for just this one fetch.

    Deliberately does NOT cache "this URL needed Selenium" onto the whole
    domain for future calls the way it used to - industry.gov.lb (an ASP.NET
    WebForms/Sitefinity site) has a static listing page 1 that renders
    perfectly (10 real links via a plain request) but one of its ARTICLE
    pages needed Selenium, which used to flip the whole domain into
    Selenium-only mode - and this site's Selenium/headless render of the
    LISTING page turned out to be badly broken (near-empty: just the <title>
    and a stray href="#", confirmed via the _debug_org diagnostic), so every
    subsequent listing page silently produced 0 links instead of just
    reusing the static path that worked fine on page 1. Trying static fresh
    per-URL costs one extra cheap request on genuinely JS-only sites, but
    avoids exactly this kind of silent, domain-wide breakage on sites that
    are only PARTIALLY dynamic.

    driver_override: use this Selenium driver instead of the shared primary
    one for the Selenium fallback. Pass _get_secondary_driver() when fetching
    individual article pages from inside a click-pagination crawl, so the
    primary driver's listing-page position/state is never disturbed.

    allow_manual_challenge_solve: if the Selenium fallback lands on a bot/
    CAPTCHA challenge page, pause and wait for you to solve it yourself in
    the (non-headless) browser window before continuing - see
    _get_secondary_driver(). Off by default since it blocks on terminal
    input; only turn it on for interactive runs (never for run_all.py-style
    unattended batches, which would hang forever waiting for a keypress).

    no_selenium: skip the Selenium fallback entirely and just return whatever
    the static fetch got, even if _looks_insufficient() would otherwise flag
    it as thin. For industry.gov.lb specifically, forcing EITHER a headless
    OR a visible Selenium session (config.csv's "force_visible_browser") was
    tried and confirmed via scrape_log.txt to come back with the same
    near-empty shell on every single page - including page 1, which a plain
    request renders perfectly. That rules out headless-fingerprinting as the
    cause (a visible, non-headless browser hit the exact same wall) and shows
    Selenium simply cannot render this particular ASP.NET/Sitefinity site at
    all, so the only reliable path here is to trust the static fetch and
    never escalate to Selenium, not even per-URL.
    """
    d = domain_of(url)
    if no_selenium:
        return _fetch_static(url, session), "static"
    static_html = None
    if not force_dynamic and d not in _dynamic_domains:
        static_html = _fetch_static(url, session)
        if not _looks_insufficient(static_html, url):
            return static_html, "static"
        log.info("[%s] Page looks JS-rendered / static fetch was thin for this URL - retrying with headless Chrome.", d)

    html = _fetch_dynamic(url, driver_override=driver_override, allow_manual_challenge_solve=allow_manual_challenge_solve)
    # FIX (2026-10-16): if Chrome isn't available / fails, or comes back with a
    # challenge page, don't throw away a perfectly good static copy (article
    # pages often have few "listing-style" links, so they were flagged "thin"
    # and lost whenever Selenium failed).
    if static_html and (not html or (is_bot_challenge_page(html) and not is_bot_challenge_page(static_html))):
        return static_html, "static"
    return html, "selenium"


# ---------------------------------------------------------------------------
# Listing-page parsing: candidate article links + "next page" detection
# ---------------------------------------------------------------------------

def find_next_page_url(soup, current_url, visited=None):
    link_tag = soup.find("link", rel=lambda v: v and "next" in v.lower())
    if link_tag and link_tag.get("href"):
        return normalize_url(current_url, link_tag["href"])

    a_rel = soup.find("a", rel=lambda v: v and "next" in v.lower())
    if a_rel and a_rel.get("href"):
        return normalize_url(current_url, a_rel["href"])

    a_aria = soup.find("a", attrs={"aria-label": re.compile("next", re.IGNORECASE)})
    if a_aria and a_aria.get("href"):
        return normalize_url(current_url, a_aria["href"])

    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        if text and NEXT_TEXT_PATTERNS.match(text):
            candidate = normalize_url(current_url, a["href"])
            if candidate and candidate != current_url:
                return candidate

    visited = visited or set()
    page_links = []
    current_num = None
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        # A bare number in an <a> is usually a numbered pager link ("1 2 3 ...") -
        # but a plain 4-digit year (a site's "News archive: 2019 2020 2021 ..."
        # nav) also passes text.isdigit(), and would otherwise get mistaken for
        # "page 2019 of the listing" here (seen on the Ministry of Finance site).
        # No real listing has that many numbered-link pages, so exclude anything
        # in a plausible calendar-year range.
        if text.isdigit() and not (1900 <= int(text) <= 2100):
            candidate = normalize_url(current_url, a["href"])
            if candidate:
                page_links.append((int(text), candidate))
                cls = " ".join(a.get("class") or [])
                if a.parent is not None:
                    cls += " " + " ".join(a.parent.get("class") or [])
                if a.get("aria-current") or re.search(r"\b(active|current|selected)\b", cls, re.I):
                    current_num = int(text)
    # FIX (2026-10-16): the old code returned the LOWEST numbered link that
    # wasn't the current URL - so on page 2 it returned page 1 (already
    # visited) and the crawl stopped after 2 pages on every numbered-pager
    # site. Now: find which page we're on and go to the next number up.
    if current_num is None:
        for tag in soup.find_all(["span", "strong", "b", "em", "li", "button"],
                                 class_=re.compile(r"(active|current|selected)", re.I)):
            t = tag.get_text(strip=True)
            if t.isdigit() and not (1900 <= int(t) <= 2100):
                current_num = int(t)
                break
    if page_links:
        page_links.sort(key=lambda t: t[0])
        if current_num is not None:
            for num, candidate in page_links:
                if num > current_num and candidate != current_url and candidate not in visited:
                    return candidate
            return None
        for _num, candidate in page_links:
            if candidate != current_url and candidate not in visited:
                return candidate

    return None


def extract_candidate_links(soup, current_url, _debug_org=None, link_pattern=None, exclude_pattern=None):
    """Return article URLs found on a listing page.

    link_pattern / exclude_pattern (optional, per-site): a regex (string or
    compiled). When link_pattern is given, ONLY same-site links whose URL
    matches it are returned - this is how each site script says exactly what
    its article URLs look like, instead of relying on the generic heuristics.
    """
    link_pattern = _compile_pattern(link_pattern)
    exclude_pattern = _compile_pattern(exclude_pattern)
    # Work on a throwaway copy with nav/etc. stripped so sitewide menu links
    # (About, Clusters, Jobs, ...) never get mistaken for article headlines
    # just because a nav item has a long label.
    working = BeautifulSoup(str(soup), "html.parser")
    for tag in working(["script", "style", "nav", "aside", "noscript"]):
        tag.decompose()

    # <form> tags are usually just a search/filter widget and safe to strip
    # so their inputs/labels never get mistaken for content - EXCEPT on
    # classic ASP.NET WebForms/Sitefinity-style sites (industry.gov.lb,
    # mopw.gov.lb, and other "ASP.NET postback pager" ministry sites in this
    # project) where the ENTIRE page - every real headline link and the
    # numbered pager alike - is wrapped in one <form id="aspnetForm"/"form1">.
    # Blanket-decomposing every <form> there deleted the real content right
    # along with it, leaving 0-1 links behind (confirmed via the _debug_org
    # diagnostic: only a single stray href="#" survived on industry.gov.lb,
    # while find_next_page_url - which never strips <form> - still found
    # every real numbered-pager link in that same raw HTML). So: only strip
    # a <form> if the page still has real links OUTSIDE it - if a form is the
    # only place on the page with any links at all, it must be acting as the
    # whole page's content wrapper (not a small widget), so it's left alone
    # rather than decomposed, no matter how many/few links it holds.
    all_links_before_form_strip = len(working.find_all("a", href=True))
    for tag in working.find_all("form"):
        links_in_form = len(tag.find_all("a", href=True))
        links_outside_form = all_links_before_form_strip - links_in_form
        if links_outside_form > 0:
            tag.decompose()

    # "header"/"footer" are only stripped when they're page-level chrome (not
    # nested inside an <article>) - some themes (e.g. UNIFIL/Drupal "uw"
    # theme) wrap each individual teaser's headline in its own
    # <header class="...__header"><h5><a>...</a></h5></header> INSIDE the
    # teaser's <article>, and blanket-stripping every <header> tag was
    # deleting that headline link on every single card, not just the page's
    # real global header.
    for tag in working(["header", "footer"]):
        if tag.find_parent("article") is None:
            tag.decompose()

    # A single <article> is normally an individual-article detail page - safe
    # to scope down to it. But a *listing* page where every teaser card is
    # its own <article> (e.g. UNIFIL's Drupal theme) has many of them, and
    # picking just the first would silently drop every other card - so only
    # take the single-article shortcut when there's exactly one.
    articles = working.find_all("article")
    solo_article = articles[0] if len(articles) == 1 else None
    content_id_match = working.find(id=re.compile("content", re.I))

    # Try each candidate container in priority order, but only accept one if
    # it actually contains at least one link - otherwise fall through to the
    # next. Without this, an unrelated element whose id merely *contains* the
    # substring "content" (e.g. a hidden "my_popup_content" modal shell
    # elsewhere on the page, seen on the Ministry of Public Health site) could
    # get picked over the real listing purely by id-matching first, silently
    # scoping the whole extraction down to an empty container - 0 links found
    # even though the page has hundreds of real article links elsewhere.
    which = "whole page (no main/article/#content found)"
    container = working
    for candidate, label in (
        (working.find("main"), "main"),
        (solo_article, "solo_article"),
        (content_id_match, f"id-match ({content_id_match.get('id')!r} <{content_id_match.name}>)" if content_id_match else None),
    ):
        if candidate is not None and candidate.find_all("a", href=True):
            container, which = candidate, label
            break

    if _debug_org is not None:
        total_links_on_page = len(working.find_all("a", href=True))
        log.info(
            "[%s] extract_candidate_links: container=%s, <article> tags on page=%d, "
            "<a href> in container=%d, <a href> in whole page=%d.",
            _debug_org, which, len(articles),
            len(container.find_all("a", href=True)), total_links_on_page,
        )
        if total_links_on_page <= 2:
            # Zero (or nearly zero) links ANYWHERE on the whole rendered page
            # (not just in the chosen container) usually means Selenium never
            # actually got the real page - an error/"not found" page, a login
            # wall, a longer JS render than SELENIUM_WAIT_SECONDS allows, etc.
            # A lone leftover nav/logo link surviving (total=1) is just as
            # diagnostic as total=0 - both mean the real listing never
            # rendered - so this isn't restricted to the total==0 case only.
            # Log the title, a body-text snippet, and whatever link(s) DID
            # survive, so we can tell which failure mode this is without
            # needing a live browser to look ourselves.
            page_title = working.title.get_text(strip=True) if working.title else "(no <title>)"
            body_text = working.get_text(" ", strip=True)[:300]
            surviving_hrefs = [a.get("href") for a in working.find_all("a", href=True)]
            log.info(
                "[%s] Only %d link(s) anywhere on the page - title=%r, surviving href(s)=%r, body starts: %r",
                _debug_org, total_links_on_page, page_title, surviving_hrefs, body_text,
            )
        elif len(container.find_all("a", href=True)) > 0:
            # The container itself has links, but extract_candidate_links still
            # returned zero candidates - the whole-page link count being >2 meant
            # the branch above never ran, so WHY each of these container links
            # got rejected was previously invisible. Dump per-link diagnostics:
            # href, visible text, and which specific filter(s) it failed/matched,
            # so a new site's rejection reason is visible from the log alone,
            # without needing to reproduce the render in a live browser.
            base_domain_dbg = domain_of(current_url)
            current_path_dbg = urlparse(current_url).path.rstrip("/") or "/"
            for a in container.find_all("a", href=True):
                href_dbg = normalize_url(current_url, a["href"])
                text_dbg = a.get_text(strip=True)
                reasons = []
                if not href_dbg:
                    reasons.append("normalize_url->None (mailto/tel/javascript/#/empty)")
                else:
                    if is_skippable_domain(href_dbg):
                        reasons.append("skippable domain")
                    if href_dbg.lower().endswith(SKIP_EXTENSIONS):
                        reasons.append("skip extension")
                    if domain_of(href_dbg) != base_domain_dbg:
                        reasons.append("different domain")
                    if PAGINATION_QUERY_PATTERN.search(href_dbg):
                        reasons.append("pagination query pattern")
                    if PAGINATION_PATH_PATTERN.search(href_dbg):
                        reasons.append("pagination path pattern")
                    href_path_dbg = urlparse(href_dbg).path.rstrip("/") or "/"
                    if href_path_dbg == current_path_dbg:
                        reasons.append("same path as listing page")
                    if NAV_EXCLUDE_PATTERN.match(href_path_dbg):
                        reasons.append("nav exclude pattern")
                    looks_detail = bool(DETAIL_PATH_PATTERN.search(href_dbg))
                    in_heading = a.find_parent(["h1", "h2", "h3", "h4", "h5", "h6"]) is not None
                    looks_headline = len(text_dbg) >= MIN_LINK_TEXT_LEN and in_heading
                    looks_read_more = bool(READ_MORE_TEXT_PATTERN.search(text_dbg))
                    looks_self_url = text_dbg.startswith(("http://", "https://")) and normalize_url(current_url, text_dbg) == href_dbg
                    if not reasons and not (looks_detail or looks_headline or looks_read_more or looks_self_url):
                        reasons.append(
                            f"no detail-path match AND no headline match AND no "
                            f"read-more-text match AND no self-URL-as-text match "
                            f"(text_len={len(text_dbg)}, in_heading={in_heading}, "
                            f"parent_tags={[p.name for p in a.parents][:4]})"
                        )
                log.info(
                    "[%s]   container link: href=%r text=%r -> %s",
                    _debug_org, href_dbg or a.get("href"), text_dbg[:80],
                    "REJECTED (%s)" % "; ".join(reasons) if reasons else "ACCEPTED",
                )

    links, seen = [], set()
    base_domain = domain_of(current_url)
    current_path = urlparse(current_url).path.rstrip("/") or "/"

    for a in container.find_all("a", href=True):
        href = normalize_url(current_url, a["href"])
        if not href or href in seen:
            continue
        if is_skippable_domain(href):
            continue
        if href.lower().endswith(SKIP_EXTENSIONS):
            continue
        if domain_of(href) != base_domain:
            continue

        parsed = urlparse(href)
        href_path = parsed.path.rstrip("/") or "/"

        if PAGINATION_QUERY_PATTERN.search(href):
            continue  # "page 2 of the listing" control, not an article
        if PAGINATION_PATH_PATTERN.search(href):
            continue  # "/page/2/"-style pager control, not an article
        if href_path == current_path:
            continue  # link back to this same listing page (any query string)
        if NAV_EXCLUDE_PATTERN.match(href_path):
            continue  # generic site nav (About, Clusters, Jobs, Resources, ...)
        if exclude_pattern is not None and exclude_pattern.search(href):
            continue
        if link_pattern is not None:
            if link_pattern.search(href):
                seen.add(href)
                links.append(href)
            continue

        text = a.get_text(strip=True)
        looks_like_detail_path = bool(DETAIL_PATH_PATTERN.search(href))
        # A long link label only counts as a headline if it's actually wrapped
        # in a heading tag, the way teaser/card headlines usually are -
        # ordinary nav links with long labels are not inside a heading. Some
        # sites (e.g. FAO) use h5/h6 for teaser titles, not just h1-h4.
        looks_like_headline = len(text) >= MIN_LINK_TEXT_LEN and a.find_parent(["h1", "h2", "h3", "h4", "h5", "h6"]) is not None
        # Short "read more"/"للمزيد" labels on a teaser card whose headline
        # itself isn't a link (moe.gov.lb and similar) - see
        # READ_MORE_TEXT_PATTERN's definition above for why this is safe to
        # accept without the length/heading requirement the general headline
        # case needs.
        looks_like_read_more = bool(READ_MORE_TEXT_PATTERN.search(text))
        # finance.gov.lb's SharePoint "Latest News" listing (confirmed
        # 2026-07-22): each item is a plain (non-heading) date + headline
        # text, followed by a SEPARATE <a> whose visible text is just the
        # link's own destination URL printed out verbatim (e.g. the link
        # text IS "https://finance.gov.lb/.../Pdf-Preview.aspx?dt=OL1608",
        # identical to its href) rather than any headline words - so
        # looks_like_headline (no heading wrapper, and text is a URL not
        # prose) and looks_like_detail_path (the per-item query param is
        # "?dt=OL1608", not one of the id/newsid/nid/articleid/item names
        # DETAIL_PATH_PATTERN recognizes) both miss every single one of
        # these real article links. A nav/menu link never displays its own
        # raw URL as its clickable text - real site navigation always uses
        # a human label - so a link whose stripped text starts with
        # "http://"/"https://" and matches its own href is a strong,
        # low-false-positive signal of an individual content item's link,
        # safe to accept without the usual length/heading/keyword checks.
        looks_like_self_url_text = text.startswith(("http://", "https://")) and normalize_url(current_url, text) == href

        if looks_like_detail_path or looks_like_headline or looks_like_read_more or looks_like_self_url_text:
            seen.add(href)
            links.append(href)
    return links


def _compile_pattern(p):
    if p is None or hasattr(p, "search"):
        return p
    return re.compile(p, re.IGNORECASE)


def is_bot_challenge_page(html):
    """True if this looks like a Cloudflare/Akamai-style CAPTCHA/JS-challenge
    interstitial rather than real page content - never save these as articles."""
    if not html:
        return False
    # See INCAPSULA_JS_REDIRECT_PATTERN's definition above - this variant has no
    # matching text at all, so it needs its own (length-capped) structural check.
    if len(html) < 3000 and INCAPSULA_JS_REDIRECT_PATTERN.search(html):
        return True
    try:
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        return False
    for tag in soup(["script", "style", "noscript", "template"]):
        tag.decompose()
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    h1 = soup.find("h1")
    head = title + " " + (h1.get_text(" ", strip=True) if h1 else "")
    text = soup.get_text(" ", strip=True)
    # Challenge interstitials are tiny pages; a real article is long.
    if STRONG_CHALLENGE_PATTERNS.search(head):
        return True
    if len(text) < 5000 and STRONG_CHALLENGE_PATTERNS.search(text):
        return True
    # Generic server error pages: only when it's the page's own title/heading,
    # or the whole visible page is short.
    if ERROR_PAGE_PATTERNS.search(head):
        return True
    if len(text) < 1500 and ERROR_PAGE_PATTERNS.search(text):
        return True
    return False


def extract_article(html, url):
    title, date, text = None, None, None

    if HAVE_TRAFILATURA:
        try:
            text = trafilatura.extract(html, include_comments=False, include_tables=False, url=url)
        except Exception:
            text = None
        try:
            meta = trafilatura.extract_metadata(html)
            if meta:
                title = meta.title
                date = meta.date
        except Exception:
            pass

    soup = BeautifulSoup(html, "html.parser")

    if not title:
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
        h1 = soup.find("h1")
        if h1:
            title = h1.get_text(strip=True) or title

    if not date:
        time_tag = soup.find("time")
        if time_tag:
            date = time_tag.get("datetime") or time_tag.get_text(strip=True)
        if not date:
            meta_date = soup.find("meta", attrs={"property": "article:published_time"})
            if meta_date and meta_date.get("content"):
                date = meta_date["content"]

    if not text:
        for tag in soup(["script", "style", "nav", "header", "footer", "form", "noscript"]):
            tag.decompose()
        container = soup.find("article") or soup.find("main") or soup.find(id=re.compile("content", re.I)) or soup
        paragraphs = [p.get_text(" ", strip=True) for p in container.find_all("p")]
        paragraphs = [p for p in paragraphs if len(p) > 30]
        text = "\n\n".join(paragraphs)

    return title, date, text


# ---------------------------------------------------------------------------
# State (already-scraped URLs) so re-runs only fetch new articles
# ---------------------------------------------------------------------------

class _DryRunVisited(set):
    """Dry-run stand-in for the 'already scraped' set - see DRY_RUN above."""
    def __init__(self, org_key):
        super().__init__()
        self.org_key = org_key
        self.links = []
        self._seen = set()

    def __contains__(self, url):
        if url not in self._seen:
            self._seen.add(url)
            self.links.append(url)
        return True          # "already scraped" -> no article is ever fetched

    def add(self, url):
        pass


def _write_dry_report(visited):
    try:
        d = os.path.join(STATE_DIR, "_dry")
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, f"{visited.org_key}.json"), "w", encoding="utf-8") as f:
            json.dump({"links": len(visited.links), "first": visited.links[:3],
                       "last": visited.links[-3:], "all_links": visited.links},
                      f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def load_state(org_key):
    if DRY_RUN:
        return _DryRunVisited(org_key)
    path = os.path.join(STATE_DIR, f"{org_key}.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_state(org_key, visited):
    if DRY_RUN:
        if isinstance(visited, _DryRunVisited):
            _write_dry_report(visited)
        return
    path = os.path.join(STATE_DIR, f"{org_key}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sorted(visited), f, ensure_ascii=False, indent=0)


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

class _NoSaveWorkbook(Workbook):
    def save(self, *args, **kwargs):
        pass


def open_or_create_workbook(path):
    if DRY_RUN:
        wb = _NoSaveWorkbook()
        return wb, wb.active
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Articles"
        ws.append(EXCEL_HEADERS)
    return wb, ws


def append_row(ws, title, url, date, text):
    ws.append([title or "", url or "", date or "", text or "", datetime.now(timezone.utc).isoformat()])


# ---------------------------------------------------------------------------
# Main per-organization crawl, called from each scraper/sites/<Org>.py
# ---------------------------------------------------------------------------

def _write_run_result(org_key, stats):
    """Leave a small JSON summary of this run for run_all.py's report."""
    try:
        d = os.path.join(STATE_DIR, "_last_run")
        os.makedirs(d, exist_ok=True)
        stats = dict(stats, finished_at=datetime.now(timezone.utc).isoformat())
        with open(os.path.join(d, f"{org_key}.json"), "w", encoding="utf-8") as f:
            json.dump(stats, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def run_scrape(org, start_url, output_filename, max_pages=DEFAULT_MAX_PAGES, max_articles=0, delay=1.5, notes="",
               link_pattern=None, exclude_pattern=None, page_url_template=None, page_start=1,
               extra_start_urls=None, extra_article_urls=None):
    """
    Generic crawl used by most site scripts.

    Per-site options (all optional - pass them from scraper/sites/<Org>.py):
      link_pattern      regex: only links matching this are treated as articles.
      exclude_pattern   regex: links matching this are never treated as articles.
      page_url_template e.g. "https://site/news?page={n}" - build listing page
                        URLs directly instead of hunting for a "next" link.
                        start_url is page `page_start`; stops when a page has no
                        article links or repeats the previous page.
      extra_start_urls  more listing pages to crawl the same way (e.g. a
                        ministry's "Announcements" section next to "News").
      extra_article_urls list (or a function returning a list) of article URLs
                        to fetch directly after the listings.

    Returns a stats dict: saved, links_found, challenges, pages, status.
    """
    out_path = os.path.join(NEWS_DIR, output_filename)
    org_key = os.path.splitext(output_filename)[0]
    stats = {"org": org, "saved": 0, "links_found": 0, "attempted": 0, "challenges": 0, "fetch_failures": 0,
             "pages": 0, "status": "ok", "message": ""}

    if "facebook_only" in (notes or "") or is_skippable_domain(start_url):
        log.warning(
            "[%s] Skipping - social-media-only source, cannot be scraped with requests or "
            "Selenium without logging in (against Facebook's terms of use). Use the Facebook "
            "Graph API with a page access token, or export posts manually. See README.md.",
            org,
        )
        stats.update(status="skipped", message="facebook only")
        _write_run_result(org_key, stats)
        return stats

    if "force_selenium" in (notes or ""):
        _dynamic_domains.add(domain_of(start_url))
        log.info("[%s] Forcing headless Chrome (Selenium) for this organization, per config.", org)

    # Some sites render a near-empty shell specifically to HEADLESS Chrome,
    # while an ordinary visible browser sees the full page. This is NOT
    # solving or bypassing any verification challenge - just a visible window.
    force_visible = "force_visible_browser" in (notes or "")
    driver_override = None
    if force_visible:
        _dynamic_domains.add(domain_of(start_url))
        driver_override = _get_secondary_driver()
        log.info("[%s] Forcing a visible (non-headless) Chrome window for this organization, per config.", org)

    # Some sites' Selenium render is simply broken - never escalate to Selenium.
    no_selenium = "no_selenium" in (notes or "")
    if no_selenium:
        log.info(
            "[%s] Selenium fallback disabled for this organization, per config - always using "
            "the plain static fetch, even if it looks thin.", org,
        )

    log.info("=== %s ===%s", org, " (TEST MODE)" if TEST_MODE else "")
    session = requests.Session()
    visited_articles = load_state(org_key)
    rp = get_robot_parser(start_url)
    wb, ws = open_or_create_workbook(out_path)
    unsaved = [0]
    challenge_streak = [0]

    def flush():
        wb.save(out_path)
        save_state(org_key, visited_articles)
        unsaved[0] = 0

    def site_blocked():
        return bool(MAX_CHALLENGE_STREAK and challenge_streak[0] >= MAX_CHALLENGE_STREAK)

    def limit_reached():
        return bool(max_articles and stats["saved"] >= max_articles) or site_blocked()

    def process_links(links):
        for link in links:
            if limit_reached():
                return
            time.sleep(delay)
            if not allowed_by_robots(rp, link):
                log.warning("[%s] Article blocked by robots.txt: %s", org, link)
                visited_articles.add(link)
                continue
            stats["attempted"] += 1
            art_html, art_method = fetch_html(
                link, session, driver_override=driver_override,
                allow_manual_challenge_solve=force_visible, no_selenium=no_selenium,
            )
            if not art_html:
                stats["fetch_failures"] += 1
                visited_articles.add(link)
                continue
            if is_bot_challenge_page(art_html):
                # Don't mark as visited - worth retrying on a future run.
                stats["challenges"] += 1
                challenge_streak[0] += 1
                log.warning("[%s] Bot/CAPTCHA/error page for article, skipping (will retry later): %s", org, link)
                if site_blocked():
                    log.warning("[%s] %d bot-check pages in a row - the site is blocking us right now, "
                                "giving up on it for this run (unscraped links are retried next run).",
                                org, challenge_streak[0])
                    return
                continue
            challenge_streak[0] = 0
            visited_articles.add(link)
            title, date, text = extract_article(art_html, link)
            if not text or len(text) < 50:
                log.warning("[%s] Little/no extractable text even via %s: %s", org, art_method, link)
                continue
            append_row(ws, title, link, date, text)
            stats["saved"] += 1
            unsaved[0] += 1
            if unsaved[0] >= SAVE_EVERY:
                flush()
            log.info("[%s] Saved article %d: %s", org, stats["saved"], (title or link)[:80])

    def crawl_listing(first_url):
        current_url = first_url
        visited_pages = set()
        pages_here = 0
        page_no = page_start
        previous_signature = None
        known_streak = 0
        while current_url and pages_here < max_pages and not limit_reached():
            if current_url in visited_pages:
                break
            if not allowed_by_robots(rp, current_url):
                log.warning("[%s] Blocked by robots.txt: %s", org, current_url)
                break

            log.info("[%s] Listing page %d: %s", org, pages_here + 1, current_url)
            html, method = fetch_html(
                current_url, session, driver_override=driver_override,
                allow_manual_challenge_solve=force_visible, no_selenium=no_selenium,
            )
            visited_pages.add(current_url)
            pages_here += 1
            stats["pages"] += 1
            if not html:
                log.warning("[%s] Could not fetch listing page (even with Selenium fallback), stopping.", org)
                if stats["pages"] == 1:
                    stats.update(status="error", message="listing page could not be fetched")
                break
            if is_bot_challenge_page(html):
                log.warning(
                    "[%s] Listing page returned a bot/CAPTCHA/error screen instead of real "
                    "content: %s - stopping.", org, current_url,
                )
                stats["challenges"] += 1
                if stats["pages"] == 1:
                    stats.update(status="blocked", message="listing page is a bot-challenge/error page")
                break

            soup = BeautifulSoup(html, "html.parser")
            article_links = extract_candidate_links(soup, current_url, link_pattern=link_pattern,
                                                    exclude_pattern=exclude_pattern)
            if not article_links:
                extract_candidate_links(soup, current_url, _debug_org=org)  # log why
            new_links = [l for l in article_links if l not in visited_articles]
            stats["links_found"] += len(article_links)
            log.info("[%s] (%s) Found %d candidate links (%d new).", org, method, len(article_links), len(new_links))

            process_links(new_links)
            flush()

            if site_blocked():
                break
            if limit_reached():
                log.info("[%s] Reached --max-articles limit (%d).", org, max_articles)
                break

            known_streak = known_streak + 1 if (article_links and not new_links) else 0
            if STOP_AFTER_KNOWN_PAGES and known_streak >= STOP_AFTER_KNOWN_PAGES:
                log.info("[%s] %d listing page(s) in a row with nothing new - older pages were "
                         "already scraped, stopping (incremental mode).", org, known_streak)
                break

            time.sleep(delay)
            if page_url_template:
                signature = frozenset(article_links)
                if not article_links or signature == previous_signature:
                    log.info("[%s] No more results after %d page(s).", org, pages_here)
                    break
                previous_signature = signature
                page_no += 1
                next_url = page_url_template.format(n=page_no)
            else:
                next_url = find_next_page_url(soup, current_url, visited=visited_pages)
            if not next_url or next_url in visited_pages:
                log.info("[%s] No further pagination found after %d page(s).", org, pages_here)
                break
            current_url = next_url

    try:
        for listing_url in [start_url] + list(extra_start_urls or []):
            if limit_reached():
                break
            crawl_listing(listing_url)

        if extra_article_urls and not limit_reached():
            extra = extra_article_urls() if callable(extra_article_urls) else extra_article_urls
            extra = [u for u in extra if u not in visited_articles]
            if extra:
                log.info("[%s] Fetching %d extra article URL(s) directly.", org, len(extra))
                stats["links_found"] += len(extra)
                process_links(extra)
    except Exception as exc:
        stats.update(status="error", message=f"{type(exc).__name__}: {exc}")
        log.exception("[%s] Crashed: %s", org, exc)
        raise
    finally:
        flush()
        if DRY_RUN:
            log.info("[%s] DRY RUN: %d listing page(s), %d unique article link(s) found.",
                     org, stats["pages"], len(getattr(visited_articles, "links", [])))
        if (stats["status"] == "ok" and stats["saved"] == 0 and stats["challenges"] >= 2
                and stats["challenges"] >= stats["attempted"]):
            stats.update(status="blocked", message="every article came back as a bot-challenge/error page")
        _write_run_result(org_key, stats)

    log.info("[%s] Done. %d new article(s) saved to %s.", org, stats["saved"], output_filename)
    return stats


def _log_near_selector_matches(org, html, selector):
    """Debug helper for click-pagination mode: when next_selector isn't found
    by Selenium, log any elements elsewhere in the page that share its most
    distinctive class token(s) - e.g. the control got re-rendered as a
    different tag, or picked up extra/different classes after a JS
    hydration pass, or is buried inside a wrapper the selector doesn't
    reach. Cheap way to see what actually happened without a live debugger.
    """
    if not html:
        return
    class_tokens = re.findall(r"\.([\w-]+)", selector)
    if not class_tokens:
        return
    soup = BeautifulSoup(html, "html.parser")
    matches = []
    for token in class_tokens:
        for tag in soup.find_all(attrs={"class": re.compile(rf"(?:^|\s){re.escape(token)}(?:\s|$)")}):
            snippet = str(tag)[:300]
            if snippet not in matches:
                matches.append(snippet)
        if matches:
            break
    if matches:
        log.info(
            "[%s] Selector '%s' not found, but %d element(s) elsewhere match class token(s) %s:",
            org, selector, len(matches), class_tokens,
        )
        for m in matches[:3]:
            log.info("[%s]   %s", org, m)
    else:
        log.info(
            "[%s] Selector '%s' not found and no elements matched its class token(s) %s either "
            "- the control may need a longer render wait, live inside an iframe, or use a "
            "completely different class/tag than expected.",
            org, selector, class_tokens,
        )


def run_scrape_click_pagination(
    org, start_url, output_filename, next_selector,
    filter_selector=None, max_pages=DEFAULT_MAX_PAGES, max_articles=0, delay=1.5,
    pager_selector=None,
):
    """
    For sites whose pagination is a JS click control (prev/next arrow buttons,
    an editable "page N of 503" widget, etc.) instead of real <a href> links -
    e.g. EIB's "All releases" search tool. A normal fetch_html()/find_next_page_url()
    pass can't follow this kind of pagination since there's nothing to click in
    a static HTML parse, so this keeps one Selenium browser session open for the
    whole organization and clicks the "next" control repeatedly instead of
    requesting numbered URLs.

    next_selector: CSS selector for the "next page" element to click.
    filter_selector: optional CSS selector clicked once at the start (e.g. a
        "Lebanon" checkbox in a country filter) before pagination begins.
    pager_selector: optional CSS selector for the whole pager widget (e.g.
        "ul#pagination"). Purely diagnostic - if given, logs which page the
        widget itself thinks is "active" (its own .active/data-page state)
        each iteration, so a stalled-pagination report can tell "the click
        isn't registering with the widget at all" apart from "the widget
        advanced internally but the AJAX content list didn't refresh in
        time" - those need different fixes.
    """
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException

    out_path = os.path.join(NEWS_DIR, output_filename)
    org_key = os.path.splitext(output_filename)[0]
    log.info("=== %s (click-pagination mode) ===", org)

    session = requests.Session()
    visited_articles = load_state(org_key)
    rp = get_robot_parser(start_url)
    wb, ws = open_or_create_workbook(out_path)

    driver = _get_driver()
    if driver is None:
        log.error("[%s] Selenium/Chrome not available - cannot use click-pagination mode.", org)
        _write_run_result(org_key, {"org": org, "saved": 0, "status": "error",
                                    "message": "Google Chrome / Selenium not available"})
        return {"org": org, "saved": 0, "pages": 0}

    saved_count = 0
    page_count = 0
    previous_links_signature = None
    known_streak = 0

    try:
        driver.get(start_url)
        time.sleep(SELENIUM_WAIT_SECONDS)

        if filter_selector:
            try:
                el = driver.find_element(By.CSS_SELECTOR, filter_selector)
                el.click()
                time.sleep(SELENIUM_WAIT_SECONDS)
                log.info("[%s] Applied filter via selector: %s", org, filter_selector)
            except NoSuchElementException:
                log.warning("[%s] Filter selector not found (%s) - continuing unfiltered.", org, filter_selector)
            except Exception as exc:
                log.warning("[%s] Could not apply filter (%s): %s", org, filter_selector, exc)

        while page_count < max_pages:
            page_count += 1
            html = driver.page_source
            current_url = driver.current_url

            if is_bot_challenge_page(html):
                log.warning("[%s] Bot/CAPTCHA challenge encountered, stopping.", org)
                break

            soup = BeautifulSoup(html, "html.parser")

            if pager_selector:
                pager_el = soup.select_one(pager_selector)
                if pager_el is None:
                    log.info("[%s] Page %d: pager_selector %r not found in this page's HTML at all.", org, page_count, pager_selector)
                else:
                    active_el = pager_el.select_one(".active") or pager_el.select_one("[aria-current]")
                    active_desc = (
                        "text=%r data-page=%r class=%r" % (
                            active_el.get_text(strip=True),
                            active_el.get("data-page"),
                            active_el.get("class"),
                        ) if active_el is not None else "no element with class 'active'/[aria-current] found inside pager"
                    )
                    log.info("[%s] Page %d: pager's own active-page indicator -> %s", org, page_count, active_desc)

            article_links = extract_candidate_links(soup, current_url)
            if not article_links:
                extract_candidate_links(soup, current_url, _debug_org=org)  # log why, see below

            # Some "Next" controls stay clickable (no disabled class/attribute
            # we can detect) even past the last page - they just silently
            # re-render the same last page instead of erroring or getting
            # removed. If clicking "next" produced the exact same set of
            # candidate links as the page before it, nothing actually
            # advanced - treat that as "reached the end" rather than looping
            # (bounded by max_pages, but no reason to burn through all of it).
            links_signature = frozenset(article_links)
            if page_count > 1 and links_signature == previous_links_signature:
                log.info(
                    "[%s] Same links as the previous page after clicking next - pagination "
                    "looks stalled/looped, stopping after %d page(s).", org, page_count - 1,
                )
                break
            previous_links_signature = links_signature

            new_links = [l for l in article_links if l not in visited_articles]
            log.info("[%s] Page %d: found %d candidate links (%d new).", org, page_count, len(article_links), len(new_links))

            known_streak = known_streak + 1 if (article_links and not new_links) else 0
            if STOP_AFTER_KNOWN_PAGES and known_streak >= STOP_AFTER_KNOWN_PAGES:
                log.info("[%s] %d page(s) in a row with nothing new - stopping (incremental mode).", org, known_streak)
                break

            for link in new_links:
                if max_articles and saved_count >= max_articles:
                    break
                time.sleep(delay)
                if not allowed_by_robots(rp, link):
                    log.warning("[%s] Article blocked by robots.txt: %s", org, link)
                    visited_articles.add(link)
                    continue
                art_html, art_method = fetch_html(link, session, driver_override=_get_secondary_driver())
                if not art_html:
                    visited_articles.add(link)
                    continue
                if is_bot_challenge_page(art_html):
                    log.warning("[%s] Bot/CAPTCHA challenge page for article, skipping (will retry later): %s", org, link)
                    continue
                visited_articles.add(link)
                title, date, text = extract_article(art_html, link)
                if not text or len(text) < 50:
                    log.warning("[%s] Little/no extractable text even via %s: %s", org, art_method, link)
                    continue
                append_row(ws, title, link, date, text)
                wb.save(out_path)
                saved_count += 1
                log.info("[%s] Saved article %d: %s", org, saved_count, (title or link)[:80])

            save_state(org_key, visited_articles)

            if max_articles and saved_count >= max_articles:
                log.info("[%s] Reached --max-articles limit (%d).", org, max_articles)
                break

            try:
                next_el = driver.find_element(By.CSS_SELECTOR, next_selector)
            except NoSuchElementException:
                log.info("[%s] No next-page control found, stopping after %d page(s).", org, page_count)
                _log_near_selector_matches(org, html, next_selector)
                break

            # Real disabled state: the native boolean "disabled" attribute, or
            # aria-disabled="true". NOT a plain substring check against the
            # class list - Tailwind CSS's own utility naming convention uses
            # variant-prefixed classes like "disabled:opacity-50" (meaning
            # "apply this opacity WHEN disabled"), which contains the literal
            # substring "disabled" on EVERY button using that convention,
            # whether or not it's actually disabled right now. A substring
            # check there means every "Next" button styled this way (seen on
            # the National News Agency site) looks permanently "disabled" and
            # pagination stops after page 1 even though it never was.
            classes = (next_el.get_attribute("class") or "")
            class_tokens = classes.lower().split()
            is_disabled_class = "disabled" in class_tokens
            aria_disabled = (next_el.get_attribute("aria-disabled") or "").lower()
            has_disabled_attr = next_el.get_attribute("disabled") is not None
            if is_disabled_class or aria_disabled == "true" or has_disabled_attr:
                log.info("[%s] Next-page control is disabled - reached the last page (%d).", org, page_count)
                break

            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_el)
                next_el.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", next_el)
            except Exception as exc:
                log.warning("[%s] Could not click the next-page control, stopping: %s", org, exc)
                break

            time.sleep(delay + SELENIUM_WAIT_SECONDS)
    finally:
        wb.save(out_path)

    log.info("[%s] Done. %d new article(s) saved to %s (%d page(s) visited).", org, saved_count, output_filename, page_count)
    _write_run_result(org_key, {"org": org, "saved": saved_count, "pages": page_count, "status": "ok", "message": ""})
    return {"org": org, "saved": saved_count, "pages": page_count}


def shutdown():
    """Call at the very end of a run to close any Selenium browser(s) opened."""
    global _driver, _secondary_driver
    if _driver is not None:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None
    if _secondary_driver is not None:
        try:
            _secondary_driver.quit()
        except Exception:
            pass
        _secondary_driver = None
