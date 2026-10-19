"""
Dedicated scraper for: Ministry of Environment
Source: https://moe.gov.lb/للحصول-على-المعلومات/الاخبار-والنشاطات.aspx
Auto-generated from config.csv - edit URL/OUTPUT_FILE here if the source changes,
or edit config.csv and regenerate.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Environment'
# NOTE (2026-07-22): the old URL, https://www.moe.gov.lb/ ("www"), is a stale
# mirror stuck on 2018-era content, AND its TLS cert doesn't cover the "www"
# hostname at all - requests fails every fetch with
# "CERTIFICATE_VERIFY_FAILED ... Hostname mismatch, certificate is not valid
# for 'www.moe.gov.lb'" (see scrape_log.txt, 2026-07-21/22 runs), which forced
# a Selenium retry that landed on Chrome's own interstitial SSL-warning page
# ("Learn more about this warning" / "Turn on enhanced protection" - not the
# real site at all) and produced 0 links either way.
# The real, current site lives at https://moe.gov.lb/ (no "www", valid cert) -
# confirmed live with 2026 news dated up to June. It DOES have a distinct news
# listing page (contradicting the old NOTES below), linked from the homepage
# as "الأخبار والنشاطات": .../للحصول-على-المعلومات/الاخبار-والنشاطات.aspx.
URL = 'https://moe.gov.lb/للحصول-على-المعلومات/الاخبار-والنشاطات.aspx'
OUTPUT_FILE = 'Ministry_of_Environment.xlsx'
# Each news card here only hyperlinks its short "<< للمزيد" ("read more") text,
# not the full headline - the headline itself sits in plain (unlinked) text.
# That "read more" text is under common.py's MIN_LINK_TEXT_LEN (12 chars) and
# isn't inside a heading tag, and its target URL's Arabic slug doesn't match
# DETAIL_PATH_PATTERN (which only recognizes English keywords like
# news/press/articles) - extract_candidate_links would otherwise reject every
# one of these as a real article link even with the URL/SSL problem fixed.
# Fixed generically in common.py: added READ_MORE_TEXT_PATTERN, which accepts
# a link by "read more"/"للمزيد"-style text alone (no length/heading
# requirement), since such a link always points straight at that one card's
# own article detail page.
NOTES = ''

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
