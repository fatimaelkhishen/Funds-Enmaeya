"""
Dedicated scraper for: Ministry of Public Works and Transport
Source: http://www.mopw.gov.lb/Media/News

NOTE: config previously only had the bare homepage. This is the same
classic ASP.NET WebForms/Sitefinity-style CMS platform as
Ministry_of_Industry.py (identical "/Media/News?page=N" numeric pagination,
h2-wrapped headline links, dated "/media/news/2025/May/slug" article paths,
and the same "PagerControl"/"UnselectedPage" pager markup).

Both sites were found returning "only 1 link on the whole page" from
common.extract_candidate_links's _debug_org diagnostic. Two separate bugs
in common.py combined to cause this, both fixed generically (no
site-specific code needed here):

1. extract_candidate_links blanket-stripped every <form> tag before
   scanning for links, and these classic ASP.NET WebForms sites wrap their
   ENTIRE page body in one <form id="form1">, so stripping it deleted all
   the real headline links right along with it. Fixed: a <form> is now
   only stripped if doing so still leaves real links elsewhere on the page.

2. The actual reason Selenium was even being used on the listing pages at
   all: industry.gov.lb's static (plain-request) listing page 1 rendered
   perfectly (10 real links, no JS needed) - but one of its ARTICLE pages
   needed Selenium, and fetch_html used to cache "this domain needs
   Selenium" globally once ANY page on it looked JS-thin. That silently
   forced every later LISTING page (2, 3, 4, ...) onto Selenium too, even
   though the listing itself never needed it - and this site's headless
   Selenium render of the listing turned out to be badly broken (just the
   <title> and a stray href="#"). Fixed: fetch_html no longer caches the
   Selenium decision across a whole domain: only the config-level
   force_selenium/force_visible_browser notes forces it up front now, and
   every other URL gets its own fresh static-first attempt.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Public Works and Transport'
URL = 'http://www.mopw.gov.lb/Media/News'
OUTPUT_FILE = 'Ministry_of_Public_Works_and_Transport.xlsx'
# no_selenium: confirmed by fetching the listing page directly - the plain static
# HTML already contains the real headline link (h2-wrapped, "...اعلان-وزير-الاشغال...")
# with no JS needed. But every logged run (scrape_log.txt, 07-15/07-21/07-22) shows
# the static fetch getting flagged "thin" and escalating to Selenium, whose headless
# render of this ASP.NET/Sitefinity site comes back as a near-empty shell (just the
# <title>, one stray href="#") - the exact same failure mode documented for
# Ministry_of_Industry.py (industry.gov.lb). Disabling Selenium for this org and
# always trusting the static fetch fixes the "0 links" result.
NOTES = 'no_selenium'

# Checked live 2026-10-16: the "News" section itself really only contains ONE
# item (May 2025) - that's why the spreadsheet had 1 row; nothing was broken.
# The ministry posts most of its material in two sibling sections instead,
# so those are crawled too:
#   /media/news-(1)  = إعلانات (announcements)   - several pages
#   /media/news-(2)  = قرارات و مذكرات (decisions and memos)
EXTRA_START_URLS = [
    'http://www.mopw.gov.lb/media/news-(1)',
    'http://www.mopw.gov.lb/media/news-(2)',
]
LINK_PATTERN = r"/media/news(-(\(|%28)\d(\)|%29))?/\d{4}/[^/]+/[^/?#]+"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    args = parser.parse_args()
    try:
        common.run_scrape(ORG, URL, OUTPUT_FILE, args.max_pages, args.max_articles, args.delay, NOTES,
                          link_pattern=LINK_PATTERN, extra_start_urls=EXTRA_START_URLS)
    finally:
        common.shutdown()
