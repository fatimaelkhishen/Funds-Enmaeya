# Lebanon News Scraper

Scrapes news/press-release pages for the ~45 organizations tracked in
`config.csv` (UN agencies, IFIs, embassies/donors, and Lebanese ministries).
Each organization has:

- its **own dedicated scraper script** in `scraper/sites/<Organization>.py`
  (not one shared file - each is independently runnable), and
- its **own Excel output file** in the `News` folder (one level up from this
  script), e.g. `News/UNDP_Lebanon.xlsx`, with columns
  `Title | URL | Date | Text | Scraped At` - one row per article.

All the actual crawling logic (fetching, pagination detection, article text
extraction, Excel writing) lives once in `scraper/common.py`, which every
site script imports - so there's no duplicated scraping code, but each
organization is still its own file you can run, edit, or debug on its own.

## Quick start: test everything first (updated 2026-10-16)

```
cd scraper
pip install -r requirements.txt
python run_all.py --test
```

`--test` runs every site with 1 listing page and 2 articles each, writes the
spreadsheets to `News/_test_output/` (your real `.xlsx` files and the
`scraper/state/` memory are NOT touched) and ends with a table like:

```
SITE                         STATUS      NEW  TOTAL   MIN  NOTE
Lebanese_Presidency          OK            2      2   0.4
Ministry_of_Finance          BLOCKED       0      0   0.1  listing page is a bot-challenge/error page
...
```

Status meanings: **OK** saved new articles - **NO NEW** ran fine, nothing new -
**BLOCKED** the website served a bot-check / error page - **ERROR** the listing
page could not be loaded - **CRASHED** the script raised an error (see the
log) - **TIMEOUT** stopped after the time limit (articles saved so far are kept).
The table is also saved as `scraper/run_summary.csv` (`run_summary_test.csv`
in test mode).

Test only some sites (names with or without `.py`, any case, or part of a name):

```
python run_all.py --test --only UNRWA Ministry_of_Labor
python run_all.py --test --only ministry_of        # all ministries
python run_all.py --list                           # show all site scripts
```

Check only links + pagination (logs only - no articles opened, nothing saved):

```
python check_pagination.py                    # every site, 5 listing pages each
python check_pagination.py --max-pages 20 --only Lebanese_Presidency
```

Result table -> `pagination_check.csv`; every link each site found ->
`state_dry/_dry/<site>.json`; full log -> `scrape_log_dry.txt`.

Real runs:

```
python run_all.py                  # full crawl of every site
python run_all.py --incremental    # weekly mode - each site stops after 2 listing
                                   # pages in a row with nothing new (used by
                                   # run_weekly.bat and the GitHub workflow)
python run_all.py --skip National_News_Agency --timeout 60
python sites/Ministry_of_Labor.py --backfill 300   # one-off: older Labor news
```

`run_all.py` never stops to wait for a keypress (the manual-CAPTCHA pause is
only used when you run a single site script yourself in a terminal).

### What was fixed on 2026-10-16

- Real articles were being thrown away as "Bot/CAPTCHA challenge" pages
  (Presidency, UNICEF, Economy ...) because words like "Page not found" inside
  the page's JavaScript were matched. Detection now only looks at the visible
  page text.
- Numbered pagers (1 2 3 ...) stopped after page 2 - fixed.
- If Chrome failed, a good plain-request copy of the page was discarded - fixed.
- `run_all.py` could hang forever on UNRWA / Finance waiting for Enter - fixed.
- Excel is now saved every 10 articles instead of after every article (much
  faster on big files like Public Health / NNA).
- Each site script can now declare exactly what its article links look like
  (`LINK_PATTERN`) and how its pages are numbered (`PAGE_URL_TEMPLATE`).
- Lebanese Presidency: new article URL format `/media/news/<id>`.
- UNRWA: unrwa.org blocks all automated visits (Cloudflare) - now reads UNRWA's
  Lebanon publications from ReliefWeb.
- UNICEF Lebanon: falls back to ReliefWeb when unicef.org blocks the script.
- Ministry of Labor: there is no pager, the old script waited for a button that
  doesn't exist - now reads the latest news (+ optional `--backfill`).
- Ministry of Economy and Trade: follows `?page=N`, longer timeout.
- Ministry of Public Works: the News section really has 1 item; the
  Announcements and Decisions sections are now crawled too.
- Ministry of Finance: the site shows an empty protection page even in a
  normal browser right now - reported as BLOCKED instead of hanging.
- Old versions of every changed file: `scraper/_backup_before_fix_2026-10-16/`.

## Handling of dynamic (JavaScript-rendered) pages

Some of these sites render their news list or "next page" control with
JavaScript rather than plain HTML, so a normal HTTP request won't see any
article links. You don't need to figure out in advance which ones those
are: every fetch first tries a plain `requests` call (fast, no browser
needed), and if that page comes back with no article links and very little
text, `common.py` automatically retries the *same* URL using a headless
Chrome browser (Selenium) instead, then keeps using Selenium for the rest
of that site for the remainder of the run. You'll see this in the log as:

```
Page looks JS-rendered / static fetch was thin - retrying with headless Chrome.
```

This needs Google Chrome installed on your machine; `webdriver-manager`
(in requirements.txt) downloads the matching chromedriver automatically the
first time it's needed, so you shouldn't have to install anything else.

## Why you have to run this yourself

This was built and researched in a sandboxed environment with no outbound
internet access (not even to install Python packages), so it could not
actually be executed there. `config.csv` was populated by manually
researching and verifying each organization's current news/press URL, and
all the code has been written and syntax-checked, but it needs to run on a
normal computer with internet access (yours).

## Setup

```
cd scraper
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # Mac/Linux
pip install -r requirements.txt
```

## Run

Run a single organization directly:

```
python sites/UNDP_Lebanon.py
python sites/UNDP_Lebanon.py --max-pages 20 --max-articles 50 --delay 2
```

Or run all 45 in sequence:

```
python run_all.py
python run_all.py --max-pages 20 --delay 2
python run_all.py --only UNDP_Lebanon.py Ministry_of_Culture.py
```

Each script remembers which article URLs it already saved (in
`scraper/state/<org>.json`), so re-running later only appends newly
published articles instead of duplicating rows in the Excel file.

A full run log is written to `scraper/scrape_log.txt`.

## Known limitations (read before assuming a file is "empty" or "wrong")

This uses heuristics (link text length, URL patterns, "next page" link
detection) rather than site-specific rules, so it works well on most sites
but a few will need extra attention:

- **Facebook-only sources** (Ministry of Youth and Sports) - neither
  `requests` nor Selenium is used here, on purpose: scraping Facebook
  without logging in doesn't work, and automating a login would violate
  Facebook's terms of use. Use the Facebook Graph API with a page access
  token, or export posts manually.
- **Sites with no distinct "news" URL** - a few ministries (Environment,
  Finance, Interior, Tourism, Justice, Public Works, Energy and Water,
  Education, TRA, Presidency, Labor) don't have an obviously separate news
  listing page as far as could be found by research alone; `config.csv`
  points these at the homepage. Browse the site once and update the URL in
  `config.csv` (then regenerate/edit the matching file in `scraper/sites/`)
  if you spot a better "News"/"الأخبار" page.
- **PDF-based announcements** - some ministries publish notices as PDFs
  linked from the page rather than as regular HTML news articles; the
  script currently skips non-HTML file types. Ask if you want PDF text
  extraction added.
- **Selenium is slower and heavier** than a plain request, and a small
  number of sites may still resist it (bot-detection, CAPTCHAs, login
  walls) - the log will show a fetch failure in those cases rather than
  silently producing an empty row.

## Editing the source list

`config.csv` (`organization,url,filename,notes`) is the single source of
truth for organization name, URL, and output filename. If you change a URL
or add an organization there, update (or regenerate) the matching file in
`scraper/sites/` to match - each site script just hardcodes those same
three values for clarity and independence.

## A note on scraping etiquette

Every fetch checks `robots.txt` and skips disallowed URLs, uses a real
User-Agent, and waits between requests (`--delay`, default 1.5s) so as not
to hammer these sites. Please keep the scraped content for internal
research/monitoring use, respect each site's terms of use, and attribute/
link back to the original source if you publish anything derived from it.
