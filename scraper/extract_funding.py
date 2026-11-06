#!/usr/bin/env python3
"""
extract_funding.py
------------------
Reads one day's scrape folder (scrapes/YYYY-MM-DD/<Org>.xlsx, written by
run_all.py --daily-dir) and builds ONE Excel file with only the
funding-related articles, in the "Lebanon Ministry Funding" template layout:

    scrapes/YYYY-MM-DD/AI-Gen/Funding_YYYY-MM-DD.xlsx

How it decides what goes in:
  1. Free keyword filter (English / French / Arabic funding words: grant,
     loan, donation, funded, هبة, تمويل, قرض, subvention, ...). Articles with
     none of these words are skipped without any AI call.
  2. Each remaining article's page is opened again, so the model also sees the
     parts the scraper leaves out (dateline, page dates, sidebars, "project
     details" boxes with budgets and start/end dates).
  3. A small OpenAI model reads the scraped text + live page, says whether it is
     about funding in Lebanon and, if so, fills in the template columns. Only
     Lebanon, and only records with a Lebanese ministry from the dashboard's list.

Needs the OPENAI_API_KEY environment variable (a GitHub secret in the daily
workflow). Without it the script writes nothing and exits normally.

Usage:
    python scraper/extract_funding.py                      # today's folder (Beirut date)
    python scraper/extract_funding.py --date 2026-11-02    # a past day
    python scraper/extract_funding.py --date 2026-11-02 --dry-run
                                                           # no AI: just list which articles
                                                           # pass the keyword filter

Cost controls: --model (default gpt-4o-mini, or the OPENAI_MODEL env var),
--max-articles (default 150 per day), --max-chars (default 6000 characters of
article text sent per article).
"""

import argparse
import glob
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRAPER_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(SCRAPER_DIR)
DEFAULT_SCRAPES_DIR = os.path.join(REPO_DIR, "scrapes")

COLUMNS = [
    "Ministry", "Project / Initiative", "Status", "Funder(s)", "Amount / Scale", "Partner(s)",
    "Region", "Funding Date", "Date Basis", "Date Evidence (from article)", "Project Period",
    "In 2025-26?", "Published Date", "Description", "Source / Link", "Enmaeya Reporting Link",
    "Confidence", "Source Feed", "Origin", "Source Checked", "Live-page Evidence", "Funding Type",
    "Sector",  # extra column after the 22 template columns: the dashboard sector for this project
]
COL_WIDTHS = {
    "Ministry": 26, "Project / Initiative": 40, "Status": 12, "Funder(s)": 30, "Amount / Scale": 20,
    "Partner(s)": 30, "Region": 18, "Funding Date": 13, "Date Basis": 20,
    "Date Evidence (from article)": 40, "Project Period": 16, "In 2025-26?": 12, "Published Date": 13,
    "Description": 70, "Source / Link": 14, "Enmaeya Reporting Link": 14, "Confidence": 11,
    "Source Feed": 24, "Origin": 12, "Source Checked": 18, "Live-page Evidence": 40, "Funding Type": 18,
    "Sector": 18,
}

# Words that must appear for an article to be worth an AI call. Kept broad on
# purpose: the AI makes the real decision, this only removes obvious non-funding news.
FUNDING_KEYWORDS = re.compile(
    r"\b(fund(?:s|ed|ing|er|ers)?|grants?|granted|loans?|donat(?:e|ed|es|ion|ions|or|ors)|"
    r"financ(?:e|ed|es|ing)|financial support|contribut(?:e|ed|es|ion|ions)|"
    r"allocat(?:e|ed|es|ion)|disburs(?:e|ed|ement)|budget|aid|assistance|"
    r"handover|handed over|hand-over|sponsor(?:ed|ship)?|"
    r"(?:US\$|USD|EUR|€|\$|£)\s?\d|\d[\d.,]*\s?(?:million|billion|USD|EUR|euros?|dollars?)|"
    r"subventions?|financement|financé|prêts?|bailleurs?)\b"
    r"|هب[ةه]|تمويل|ممول|بتمويل|قرض|منح[ةه]|تبرع|مساعد[ةه]|مساعدات|دعم مالي|"
    r"مليون|مليار|دولار|يورو|موازن[ةه]|اتفاقي[ةه] تمويل",
    re.IGNORECASE,
)

# The Lebanese ministries the funding dashboard knows (same names as MINISTRY_MAP in
# the dashboard repo's build_data.py). The model must pick from this list, so names
# always match the dashboard. Records with no Lebanese ministry are dropped.
LEBANESE_MINISTRIES = [
    "Ministry of Agriculture", "Ministry of Culture", "Ministry of Economy and Trade",
    "Ministry of Education and Higher Education", "Ministry of Energy and Water",
    "Ministry of Environment", "Ministry of Finance", "Ministry of Foreign Affairs and Emigrants",
    "Ministry of Industry", "Ministry of Information", "Ministry of Interior and Municipalities",
    "Ministry of Justice", "Ministry of Labour", "Ministry of Public Health",
    "Ministry of Public Works and Transport", "Ministry of Social Affairs",
    "Ministry of State for Administrative Development",
    "Ministry of State for Administrative Reform (OMSAR)", "Ministry of Telecommunications",
    "Ministry of Tourism", "Ministry of National Defense",
]
NO_MINISTRY = "Not specified"

# The dashboard's six sectors (ids from data.json sectorFilters) and what belongs in each,
# taken from the subtopics on the dashboard's funding-network page. Shared word-for-word with
# classify_sectors.py in the dashboard repo so old and new records are judged the same way.
SECTORS = {
    "Health": "hospitals, primary and community health care, medicines and medical equipment, "
              "public health, disease control, vaccination, mental health, digital health",
    "Education": "schools, universities, access to education, crisis/emergency education, "
                 "teachers, vocational and technical training (TVET), future skills, citizenship",
    "Economy": "economic recovery, jobs and employment, livelihoods, entrepreneurship, SMEs, "
               "private sector, trade, industry, investment, finance, tourism, agricultural value "
               "chains and farmer incomes",
    "Environment": "climate, water and wastewater, waste management, biodiversity, forests and "
                   "land, clean/renewable energy, electricity sector, sustainable agriculture and "
                   "natural resources",
    "Artificial Intelligence": "AI, digital government and e-services, digital transformation of "
                               "public administration, data systems, telecoms, AI skills, "
                               "innovation and tech",
    "Development": "governance, institutions, security forces and justice, social protection and "
                   "social inclusion, humanitarian, food and cash assistance, shelter, refugees, "
                   "youth, gender, municipalities, reconstruction and general recovery",
}
SECTOR_RULES = "\n".join(f"  - {k}: {v}" for k, v in SECTORS.items())

SYSTEM_PROMPT = """You extract funding information about Lebanon from one news article.

ONLY LEBANON: the funding must be for a project, programme, institution or population IN
LEBANON. Set is_lebanon=false for anything about other countries, regions (e.g. Africa, the
EU, small island states), or global/multi-country programmes that only mention Lebanon in
passing. A Lebanon-specific component of a regional programme counts only if the article
describes what happens in Lebanon.

An article counts as FUNDING if it reports money or in-kind resources being granted, loaned,
donated, allocated, disbursed, pledged, signed or launched for a project, programme, institution
or population in Lebanon (e.g. a donor grant, a World Bank loan, an EU-funded project, a handover
of equipment or medicines, a budget allocation by a ministry). General politics, security news,
diplomatic visits with no funding, events, statements, and reports without a funded activity are
NOT funding.

If the article is not about funding, return is_funding=false and an empty records list.
If it is, return one record per distinct funded project/initiative in the article (usually one).

INPUT: you get the SCRAPED ARTICLE TEXT and, when the page could be opened again, the LIVE PAGE
(page metadata dates plus all visible text: header, dateline, sidebars, "project details" or
"related project" boxes with budgets, donors and start/end dates). Use BOTH. Fill every field the
article or page supports: read the whole text, including datelines ("BEIRUT, 18 February 2020",
"18/02/2020"), Arabic dates ("بتاريخ 16/3/2021", "الأربعاء 22 تموز 2026"), and project boxes.
Leave a field empty only when neither source supports it. Never invent facts.

- article_date: the article's real publication date as M/D/YYYY, from the dateline / header /
  page metadata (prefer the date shown on the article over "last modified" metadata). "" if none.

Field rules (write values in English even if the article is Arabic or French):
- ministry: the ONE Lebanese government ministry most involved in or benefiting from the
  funding, chosen from the allowed list. If several are involved, pick the main one. Foreign
  ministries (e.g. a donor country's foreign ministry) are funders, never the ministry.
  Use "Not specified" if no Lebanese ministry is involved. Ministry mapping hints: army /
  Lebanese Armed Forces -> Ministry of National Defense; police / ISF / civil defense /
  municipalities -> Ministry of Interior and Municipalities; hospitals / medicines -> Ministry
  of Public Health; schools / universities -> Ministry of Education and Higher Education;
  electricity / water -> Ministry of Energy and Water; farmers / forests -> Ministry of
  Agriculture; refugees / social protection / cash assistance -> Ministry of Social Affairs.
- project: short name of the project / initiative.
- status: one of Announced, Signed, Planned, Ongoing, Completed.
- funders: donor(s) providing the money, comma-separated, or "Not specified".
- amount: the money for THIS project with currency (e.g. "USD 2.7 million", "EUR 45 million").
  Look everywhere: budget, grant/loan value, "worth"/"valued at" an in-kind donation, Arabic
  figures ("بقيمة 500 ألف دولار"), project boxes on the live page. If no money figure exists, give
  the scale of what was provided ("864 metric tons of rice", "12 ambulances", "3,000 families").
  "Not specified" only if neither exists.
- partners: implementing and co-implementing organisations, comma-separated - include the
  organisation that published the article when it implements the project (e.g. FAO, ILO, UNDP,
  WFP), plus any NGOs, municipalities or other agencies named. "" only if none.
- region: place in Lebanon (e.g. "Beirut", "Bekaa", "South Lebanon"), or "Lebanon" if national.
- funding_date: when the money moved / the agreement was signed / the project was launched or
  handed over, as M/D/YYYY. Work through these in order and use the first that applies:
  1. a date given for the signing, launch, approval, disbursement, announcement or handover
     -> date_basis "In-text: <event>" (Signed, Launched, Approved, Disbursed, Announced, Handover);
  2. the article reports that event as happening now ("today", "on Tuesday", "signed", "launched",
     "handed over", "inaugurated" in a news report) -> use article_date, date_basis
     "Same-day report";
  3. a project start date in the text or a project box -> date_basis "In-text date";
  4. otherwise "" and date_basis "Not stated in article".
  If only month and year are known use the 1st of the month (e.g. "March 2023" -> 3/1/2023).
- date_basis: as above.
- date_evidence: short exact quote (original language) supporting funding_date, or "".
- project_period: start -> end as "YYYY-MM -> YYYY-MM" (or "YYYY -> YYYY"), from stated dates or
  durations ("a three-year project launched in 2024" -> "2024 -> 2027"; "1 March 2023 - 30
  September 2026" -> "2023-03 -> 2026-09"). "" if neither is stated.
- in_2025_26: "Yes" if the funding date, or any part of the project period, falls in 2025 or
  2026; "NO - <year>" if it is clearly entirely in another year (give the funding/start year);
  "Unknown" only if no date at all is available.
- live_page_evidence: if a fact you used (amount, date, period, donor) came from the LIVE PAGE
  rather than the article text, a short exact quote of it; otherwise "".
- description: 1-2 sentence English summary of who funds what, for whom, and how much.
- confidence: High (all key facts explicit), Medium (some inferred), Low (vague/partial).
- funding_type: "External donor", "Ministry budget", "Unknown".
- sector: the ONE dashboard sector that best matches what the PROJECT does (not which ministry
  runs it). E.g. a job-training project with the Ministry of Social Affairs is Economy; a digital
  farmers' registry is Artificial Intelligence; a forest project is Environment. Sectors:
""" + SECTOR_RULES + "\n"

RECORD_FIELDS = ["ministry", "project", "status", "funders", "amount", "partners", "region",
                 "funding_date", "date_basis", "date_evidence", "project_period", "in_2025_26",
                 "description", "confidence", "funding_type", "live_page_evidence", "sector"]

RESPONSE_SCHEMA = {
    "name": "funding_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "required": ["is_funding", "is_lebanon", "article_date", "records"],
        "properties": {
            "is_funding": {"type": "boolean"},
            "is_lebanon": {"type": "boolean"},
            "article_date": {"type": "string"},
            "records": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": RECORD_FIELDS,
                    "properties": {
                        **{f: {"type": "string"} for f in RECORD_FIELDS},
                        "ministry": {"type": "string", "enum": LEBANESE_MINISTRIES + [NO_MINISTRY]},
                        "status": {"type": "string",
                                   "enum": ["Announced", "Signed", "Planned", "Ongoing", "Completed"]},
                        "sector": {"type": "string", "enum": list(SECTORS)},
                        "date_basis": {"type": "string", "enum": [
                            "In-text: Signed", "In-text: Launched", "In-text: Approved",
                            "In-text: Disbursed", "In-text: Announced", "In-text: Handover",
                            "Same-day report", "In-text date", "Not stated in article"]},
                        "confidence": {"type": "string", "enum": ["High", "Medium", "Low"]},
                        "funding_type": {"type": "string",
                                         "enum": ["External donor", "Ministry budget", "Unknown"]},
                    },
                },
            },
        },
    },
}

# Approximate USD per 1M tokens, only used to print a cost estimate in the log.
PRICES = {"gpt-4o-mini": (0.15, 0.60), "gpt-4.1-mini": (0.40, 1.60), "gpt-4.1-nano": (0.10, 0.40),
          "gpt-5-mini": (0.25, 2.00), "gpt-5-nano": (0.05, 0.40)}

SHORT_TEXT = 400  # use the re-opened page's main text when the scraped text is shorter than this


def beirut_today():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Beirut")).strftime("%Y-%m-%d")
    except Exception:
        return (datetime.now(timezone.utc) + timedelta(hours=2)).strftime("%Y-%m-%d")


def us_date(value):
    """'2026-07-07' / datetime -> '7/7/2026' (the template's style). Anything else unchanged."""
    if not value:
        return ""
    if isinstance(value, datetime):
        d = value
    else:
        s = str(value).strip()
        d = None
        for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                d = datetime.strptime(s[:19] if "T" in s else s[:10], fmt)
                break
            except ValueError:
                continue
        if d is None:
            return s
    return f"{d.month}/{d.day}/{d.year}"


def read_day_articles(day_dir):
    articles = []
    for path in sorted(glob.glob(os.path.join(day_dir, "*.xlsx"))):
        name = os.path.basename(path)
        if name.startswith("Funding_") or name.startswith("~$"):
            continue
        feed = os.path.splitext(name)[0].replace("_", " ")
        wb = load_workbook(path, read_only=True)
        rows = list(wb.active.iter_rows(values_only=True))
        wb.close()
        if not rows:
            continue
        header = [str(h or "").strip().lower() for h in rows[0]]
        idx = {k: header.index(k) for k in ("title", "url", "date", "text") if k in header}
        for r in rows[1:]:
            get = lambda k: (r[idx[k]] if k in idx and idx[k] < len(r) else None) or ""
            if not get("url") and not get("title"):
                continue
            articles.append({"feed": feed, "title": str(get("title")), "url": str(get("url")),
                             "date": get("date"), "text": str(get("text"))})
    return articles


def distrust_constant_dates(articles, run_date):
    """Some sites give every article the same date (a site-wide "last modified" or a
    default): e.g. Ministry of Social Affairs pages all scraped as 2023-11-23, Ministry
    of Justice as 2026-01-01. If 3+ articles from one feed share a date that is more
    than 30 days older than the run date, that date is not trusted: it is cleared so
    the AI reads the real date off the article/live page instead."""
    try:
        run = datetime.strptime(run_date, "%Y-%m-%d")
    except ValueError:
        return
    groups = {}
    for a in articles:
        d = str(a["date"] or "")[:10]
        if re.match(r"^\d{4}-\d{2}-\d{2}$", d):
            groups.setdefault((a["feed"], d), []).append(a)
    for (feed, d), group in groups.items():
        if len(group) >= 3 and (run - datetime.strptime(d, "%Y-%m-%d")).days > 30:
            print(f"  {feed}: {len(group)} articles share the date {d} - treated as unknown")
            for a in group:
                a["date"] = ""


def fetch_live_page(url, max_chars):
    """Open the article page again. Returns (main_text, page_text) - page_text is the
    page's metadata dates plus all visible text (header, dateline, sidebars, project boxes),
    which the scraped article text often leaves out. ('', '') on failure."""
    try:
        resp = requests.get(url, timeout=30, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
            "Accept-Language": "en,ar;q=0.8,fr;q=0.6"})
        if resp.status_code != 200 or not resp.text:
            return "", ""
        html = resp.text
    except Exception:
        return "", ""
    main_text = ""
    try:
        import trafilatura
        main_text = trafilatura.extract(html) or ""
    except Exception:
        pass
    page_text = ""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        dates = []
        for m in soup.find_all("meta"):
            key = (m.get("property") or m.get("name") or m.get("itemprop") or "").lower()
            if any(k in key for k in ("published", "date", "modified", "created")) and m.get("content"):
                dates.append(f"{key}: {m['content']}")
        for t in soup.find_all("time"):
            if t.get("datetime"):
                dates.append(f"<time>: {t['datetime']} ({t.get_text(' ', strip=True)[:40]})")
        for js in soup.find_all("script", type="application/ld+json"):
            for k, v in re.findall(r'"(datePublished|dateCreated|dateModified|startDate|endDate)"\s*:\s*"([^"]+)"',
                                   js.string or ""):
                dates.append(f"{k}: {v}")
        for t in soup(["script", "style", "noscript", "svg", "nav", "footer", "form", "iframe"]):
            t.decompose()
        visible = re.sub(r"\s+", " ", soup.get_text(" ")).strip()
        page_text = ("Page metadata dates: " + "; ".join(dict.fromkeys(dates))[:800] + "\n\n"
                     if dates else "") + visible[:max_chars]
    except Exception:
        pass
    return main_text, page_text


def ask_model(api_key, model, article, text, page_text, max_chars, usage):
    user = (f"Source feed: {article['feed']}\n"
            f"Published date (as scraped - may be wrong): {article['date'] or 'unknown'}\n"
            f"Title: {article['title']}\n"
            f"URL: {article['url']}\n\n"
            f"=== SCRAPED ARTICLE TEXT ===\n{text[:max_chars]}\n\n"
            f"=== LIVE PAGE ===\n{page_text or '(the page could not be opened again)'}")
    body = {
        "model": model,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": user}],
        "response_format": {"type": "json_schema", "json_schema": RESPONSE_SCHEMA},
    }
    if not model.startswith("gpt-5"):
        body["temperature"] = 0
    for attempt in range(5):
        try:
            resp = requests.post("https://api.openai.com/v1/chat/completions", json=body, timeout=120,
                                 headers={"Authorization": f"Bearer {api_key}"})
        except requests.RequestException as exc:
            err = str(exc)
        else:
            if resp.status_code == 200:
                data = resp.json()
                u = data.get("usage") or {}
                usage["in"] += u.get("prompt_tokens", 0)
                usage["out"] += u.get("completion_tokens", 0)
                return json.loads(data["choices"][0]["message"]["content"])
            err = f"HTTP {resp.status_code}: {resp.text[:300]}"
            if resp.status_code in (400, 401, 403, 404):
                raise RuntimeError(err)  # bad key / model name - retrying won't help
        wait = 5 * (attempt + 1)
        print(f"    OpenAI error ({err}); retrying in {wait}s")
        time.sleep(wait)
    raise RuntimeError("OpenAI request kept failing")


def write_excel(rows, out_path, run_date, stats):
    wb = Workbook()
    ws = wb.active
    ws.title = "Funding"
    ws.append([f"Lebanon Ministry Funding — Daily extraction {run_date}"])
    ws.append([f"{len(rows)} record(s) from {stats['funding']} funding article(s), out of "
               f"{stats['total']} article(s) scraped on {run_date}. Extracted automatically by AI "
               f"({stats['model']}) from the scraped article text - check important figures at the source."])
    ws.append([])
    ws.append(COLUMNS)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].font = Font(italic=True, color="555555")
    head_fill = PatternFill("solid", fgColor="1F3864")
    for c in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=4, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTHS.get(COLUMNS[c - 1], 18)
    link_col = COLUMNS.index("Source / Link") + 1
    for rec in rows:
        ws.append([rec.get(col, "") for col in COLUMNS])
        r = ws.max_row
        url = rec.get("_url", "")
        if url:
            cell = ws.cell(row=r, column=link_col)
            cell.value = "Open source ▸"
            cell.hyperlink = url
            cell.font = Font(color="0563C1", underline="single")
        for c in range(1, len(COLUMNS) + 1):
            ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B5"
    if rows:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(COLUMNS))}{ws.max_row}"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def main():
    p = argparse.ArgumentParser(description="Build the daily funding Excel from one day's scrape folder.")
    p.add_argument("--date", default=None, help="YYYY-MM-DD (default: today in Beirut)")
    p.add_argument("--scrapes-dir", default=DEFAULT_SCRAPES_DIR)
    p.add_argument("--model", default=os.environ.get("OPENAI_MODEL") or "gpt-4o-mini")
    p.add_argument("--max-articles", type=int, default=150,
                   help="Max articles sent to the AI per run (cost cap). Default 150.")
    p.add_argument("--max-chars", type=int, default=6000,
                   help="Max characters of article text sent per article. Default 6000.")
    p.add_argument("--no-live", action="store_true",
                   help="Don't re-open each article page (cheaper, but fewer filled-in cells).")
    p.add_argument("--dry-run", action="store_true",
                   help="No AI calls: only show which articles pass the keyword filter.")
    args = p.parse_args()

    run_date = args.date or beirut_today()
    day_dir = os.path.join(args.scrapes_dir, run_date)
    out_path = os.path.join(day_dir, "AI-Gen", f"Funding_{run_date}.xlsx")
    if not os.path.isdir(day_dir):
        print(f"No scrape folder for {run_date} ({day_dir}) - nothing to do.")
        return

    articles = read_day_articles(day_dir)
    distrust_constant_dates(articles, run_date)
    candidates = [a for a in articles if FUNDING_KEYWORDS.search(f"{a['title']}\n{a['text']}")]
    print(f"{run_date}: {len(articles)} article(s) scraped, {len(candidates)} mention funding keywords.")

    if args.dry_run:
        for a in candidates:
            print(f"  [{a['feed']}] {a['title'][:100]}")
        return

    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        print("OPENAI_API_KEY is not set - skipping the funding extraction.")
        return

    if len(candidates) > args.max_articles:
        print(f"Only the first {args.max_articles} candidate article(s) are sent to the AI (--max-articles).")
        candidates = candidates[:args.max_articles]

    usage = {"in": 0, "out": 0}
    rows, funding_articles, failures = [], 0, 0
    for i, a in enumerate(candidates, 1):
        text, page_text, checked = a["text"], "", "Scraped text only"
        if a["url"] and not args.no_live:
            main_text, page_text = fetch_live_page(a["url"], args.max_chars)
            checked = "Re-checked at source" if page_text else "Fetch failed (scraped text only)"
            if len(text.strip()) < SHORT_TEXT and len(main_text) > len(text):
                text = main_text
        if not (text.strip() or page_text.strip() or a["title"].strip()):
            continue
        try:
            result = ask_model(api_key, args.model, a, text or a["title"], page_text, args.max_chars, usage)
        except RuntimeError as exc:
            failures += 1
            print(f"  [{i}/{len(candidates)}] AI error, skipped: {a['title'][:70]} ({exc})")
            if "HTTP 401" in str(exc) or "HTTP 404" in str(exc):
                print("Stopping: check the OPENAI_API_KEY secret and the model name.")
                break
            continue
        recs = (result.get("records") or []) if result.get("is_funding") else []
        if recs and not result.get("is_lebanon"):
            verdict, recs = "not Lebanon", []
        else:
            kept = [r for r in recs if r.get("ministry") in LEBANESE_MINISTRIES]
            verdict = "FUNDING" if kept else ("no ministry" if recs else "no")
            recs = kept
        print(f"  [{i}/{len(candidates)}] {verdict:11s} [{a['feed']}] {a['title'][:80]}")
        if recs:
            funding_articles += 1
        # Some sites give the scraper a wrong date (e.g. a "last modified" date shared by every
        # page). Prefer the date the AI read off the article itself when it found one.
        published = us_date(a["date"])
        art_date = (result.get("article_date") or "").strip()
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", art_date):
            published = art_date
        for rec in recs:
            rows.append({
                "Ministry": rec["ministry"], "Project / Initiative": rec["project"],
                "Status": rec["status"], "Funder(s)": rec["funders"], "Amount / Scale": rec["amount"],
                "Partner(s)": rec["partners"], "Region": rec["region"],
                "Funding Date": rec["funding_date"], "Date Basis": rec["date_basis"],
                "Date Evidence (from article)": rec["date_evidence"],
                "Project Period": rec["project_period"], "In 2025-26?": rec["in_2025_26"],
                "Published Date": published, "Description": rec["description"],
                "Enmaeya Reporting Link": "", "Confidence": rec["confidence"],
                "Source Feed": a["feed"], "Origin": "Daily run", "Source Checked": checked,
                "Live-page Evidence": rec["live_page_evidence"], "Funding Type": rec["funding_type"],
                "Sector": rec["sector"],
                "_url": a["url"],
            })

    rows.sort(key=lambda r: (r["Ministry"] == "Not specified", r["Ministry"], r["Source Feed"]))
    write_excel(rows, out_path, run_date, {"total": len(articles), "funding": funding_articles,
                                           "model": args.model})
    pin, pout = PRICES.get(args.model, (0, 0))
    cost = usage["in"] / 1e6 * pin + usage["out"] / 1e6 * pout
    print(f"\nSaved {len(rows)} funding record(s) from {funding_articles} article(s) to {out_path}")
    print(f"AI usage: {usage['in']:,} input + {usage['out']:,} output tokens"
          + (f" (about ${cost:.3f})" if pin else "") + (f"; {failures} article(s) failed" if failures else ""))


if __name__ == "__main__":
    main()
