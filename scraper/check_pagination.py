#!/usr/bin/env python3
"""
check_pagination.py
-------------------
Checks that every scraper (1) finds article links and (2) moves from one
listing page to the next - WITHOUT opening any article and WITHOUT writing
any Excel file or touching scraper/state/. Logs only.

    python check_pagination.py                        # every site, 5 listing pages each
    python check_pagination.py --max-pages 20         # walk further
    python check_pagination.py --only UNRWA Ministry_of_Labor
    python check_pagination.py --only ministry_of     # all ministries

At the end it prints a table (also saved as scraper/pagination_check.csv):
    OK        links were found (PAGES = listing pages visited)
    NO LINKS  the page loaded but no article links were recognised
    BLOCKED   the site served a bot-check / error page
    ERROR     the listing page could not be loaded
    CRASHED   the script raised an error (scroll up / see scrape_log_dry.txt)
Every link each site found is listed in scraper/state_dry/_dry/<site>.json.

Same options as run_all.py (--only, --skip, --max-pages, --delay, --timeout).
"""
import os
import subprocess
import sys

here = os.path.dirname(os.path.abspath(__file__))
sys.exit(subprocess.call([sys.executable, os.path.join(here, "run_all.py"), "--dry-run"] + sys.argv[1:]))
