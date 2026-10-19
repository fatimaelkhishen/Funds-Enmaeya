"""
Dedicated scraper for: Ministry of Labor
Source: https://www.labor.gov.lb/AllNews.aspx?pub=1

Checked live 2026-10-16: AllNews.aspx?pub=1 shows the 10 latest news items,
each linking to LatestNewsDetails.aspx?newsid=<N>. There is NO pager on the
page at all (the old version waited for an "lbtnNext" button that doesn't
exist, so it always stopped with 0 articles). Changing "pub=" switches
category, not page.

The news ids are consecutive (e.g. 25154 ... 25163), so to fetch OLDER news
you can walk the ids backwards:

    python sites/Ministry_of_Labor.py --backfill 300

fetches up to 300 ids below the newest one (only once - after that the
normal weekly run just picks up the latest 10).

This server is slow and sometimes answers with an ASP.NET "Out of memory"
error page; those are detected and retried on the next run.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common

ORG = 'Ministry of Labor'
URL = 'https://www.labor.gov.lb/AllNews.aspx?pub=1'
OUTPUT_FILE = 'Ministry_of_Labor.xlsx'
NOTES = ''
LINK_PATTERN = r"LatestNewsDetails\.aspx\?newsid=\d+"
DETAIL_TEMPLATE = 'https://www.labor.gov.lb/LatestNewsDetails.aspx?newsid={id}'

common.REQUEST_TIMEOUT = 60   # slow server


def backfill_urls(count):
    def _make():
        org_key = os.path.splitext(OUTPUT_FILE)[0]
        ids = [int(m.group(1)) for u in common.load_state(org_key)
               for m in [re.search(r"newsid=(\d+)", u, re.I)] if m]
        if not ids:
            return []
        newest = max(ids)
        return [DETAIL_TEMPLATE.format(id=i) for i in range(newest - 1, max(newest - 1 - count, 0), -1)]
    return _make


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"Scrape news for {ORG}")
    parser.add_argument("--max-pages", type=int, default=common.DEFAULT_MAX_PAGES)
    parser.add_argument("--max-articles", type=int, default=0)
    parser.add_argument("--delay", type=float, default=1.5)
    parser.add_argument("--backfill", type=int, default=0,
                        help="Also fetch this many older news ids below the newest one.")
    args = parser.parse_args()
    try:
        common.run_scrape(ORG, URL, OUTPUT_FILE, args.max_pages, args.max_articles, args.delay, NOTES,
                          link_pattern=LINK_PATTERN,
                          extra_article_urls=backfill_urls(args.backfill) if args.backfill else None)
    finally:
        common.shutdown()
