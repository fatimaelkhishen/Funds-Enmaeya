"""
Dedicated scraper for: Ministry of Finance
Source: https://www.finance.gov.lb/ar-lb
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.

NOTE (2026-07-22): the homepage's "Latest News" widget (SharePoint-generated)
already has dozens of real, individually-linked items (dated Jan-Jul 2026,
each linking to its own .../Pages/Pdf-Preview.aspx?dt=OL#### page) - fetched
directly and confirmed. So "homepage only" is still accurate (no separate
listing URL, the widget itself isn't paginated - it's just a long list
already), but it was returning 0 links for a different reason: every item's
headline is plain (non-heading) text, and the <a> next to it displays its own
destination URL as its visible text (e.g. link text IS
"https://finance.gov.lb/.../Pdf-Preview.aspx?dt=OL1608") instead of any
headline wording - so neither the heading/length headline check nor
DETAIL_PATH_PATTERN's known query-param names (id/newsid/nid/articleid/item -
this site uses "dt=") matched any of them. Fixed generically in common.py:
added a check that accepts a link whose text is itself a URL matching its own
href (real nav never does this, so it's a safe, low-false-positive signal of
an individual content link).

SEPARATE issue (seen 2026-07-22 on Fatima's actual run - not reproducible from
this project's own dev-side fetches, which got the real page fine every time):
both the static request AND the headless-Selenium fallback instead got an
Incapsula/Imperva-style JS "challenge" redirect page - title just "MOF", one
link to a random-looking "/<base64url token>.html" path, no real content.
common.py's is_bot_challenge_page() didn't recognize this variant before (it
has none of the usual "checking your browser" text) and silently logged "0
candidate links" as if the page were just empty - fixed generically by adding
a structural pattern (INCAPSULA_JS_REDIRECT_PATTERN) so this now gets flagged
and logged as a bot challenge instead of a silent zero. That alone doesn't get
PAST the challenge though - since it showed up on both static AND headless
requests (not just headless), this looks like real IP/session-based bot or
rate-limit detection, unlike the "headless render is simply broken" pattern
seen on industry.gov.lb/mopw.gov.lb/agriculture.gov.lb. Trying
force_visible_browser next (a real, visible Chrome window is less likely to
get flagged than a headless one); if that still doesn't get through, this may
just need a retry later, from a different network, or a manual
allow_manual_challenge_solve run.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Finance'
URL = 'https://www.finance.gov.lb/ar-lb'
OUTPUT_FILE = 'Ministry_of_Finance.xlsx'
# UPDATE 2026-10-16: checked in a normal desktop Chrome on Fatima's own computer -
# finance.gov.lb AND www.finance.gov.lb both stay on the same empty "MOF"
# protection page (0 links, nothing renders), so the site currently blocks all
# automated visits, browser or not. force_visible_browser was removed because
# it made run_all.py stop and wait for a keypress. The script now just tries
# once with a plain request, reports "blocked" in the run summary, and moves on.
# When the site works again it will start saving articles automatically.
NOTES = 'homepage only - no distinct news listing URL found; no_selenium'

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
