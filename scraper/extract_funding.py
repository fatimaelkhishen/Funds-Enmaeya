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
  2. Each remaining article goes to a small OpenAI model, which says whether
     it is really about funding and, if so, fills in the template columns.
     If the scraped text is very short, the article link is opened again to
     get the full page before asking the model.

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
]
COL_WIDTHS = {
    "Ministry": 26, "Project / Initiative": 40, "Status": 12, "Funder(s)": 30, "Amount / Scale": 20,
    "Partner(s)": 30, "Region": 18, "Funding Date": 13, "Date Basis": 20,
    "Date Evidence (from article)": 40, "Project Period": 16, "In 2025-26?": 12, "Published Date": 13,
    "Description": 70, "Source / Link": 14, "Enmaeya Reporting Link": 14, "Confidence": 11,
    "Source Feed": 24, "Origin": 12, "Source Checked": 18, "Live-page Evidence": 40, "Funding Type": 18,
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
- amount: amount with currency as written (e.g. "USD 2.7 million", "EUR 45 million"), or a scale
  ("864 metric tons of rice"), or "Not specified".
- partners: implementing partners, comma-separated, or "" if none.
- region: place in Lebanon (e.g. "Beirut", "Bekaa", "South Lebanon"), or "Lebanon" if national.
- funding_date: date the money moved / agreement was signed / project launched, as M/D/YYYY.
  Only if the article states it; otherwise "".
- date_basis: "Not stated in article", "Same-day report", "In-text date", or
  "In-text: <event>" where event is Signed, Launched, Approved, Disbursed, Announced or Handover.
- date_evidence: short exact quote from the article supporting funding_date, or "".
- project_period: e.g. "2025-01 -> 2026-12" or "2026 -> 2029" if stated, else "".
- in_2025_26: "Yes" if the funding happened in 2025 or 2026, "NO - <year>" if clearly another
  year, "Unknown" if unclear.
- description: 1-2 sentence English summary of who funds what, for whom, and how much.
- confidence: High (all key facts explicit), Medium (some inferred), Low (vague/partial).
- funding_type: "External donor", "Ministry budget", "Unknown".
"""

RECORD_FIELDS = ["ministry", "project", "status", "funders", "amount", "partners", "region",
                 "funding_date", "date_basis", "date_evidence", "project_period", "in_2025_26",
                 "description", "confidence", "funding_type"]

RESPONSE_SCHEMA = {
    "name": "funding_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "required": ["is_funding", "is_lebanon", "records"],
        "properties": {
            "is_funding": {"type": "boolean"},
            "is_lebanon": {"type": "boolean"},
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

SHORT_TEXT = 400  # re-open the link when the scraped text is shorter than this


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


def refetch_text(url):
    """Open the article again and pull its main text. Returns '' on failure."""
    try:
        resp = requests.get(url, timeout=30, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
        if resp.status_code != 200 or not resp.text:
            return ""
        html = resp.text
    except Exception:
        return ""
    try:
        import trafilatura
        text = trafilatura.extract(html) or ""
        if text:
            return text
    except Exception:
        pass
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        for t in soup(["script", "style", "nav", "header", "footer"]):
            t.decompose()
        return re.sub(r"\s+", " ", soup.get_text(" ")).strip()
    except Exception:
        return ""


def ask_model(api_key, model, article, text, max_chars, usage):
    user = (f"Source feed: {article['feed']}\n"
            f"Published date (from the site): {article['date'] or 'unknown'}\n"
            f"Title: {article['title']}\n"
            f"URL: {article['url']}\n\n"
            f"Article text:\n{text[:max_chars]}")
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
        text, checked = a["text"], "Scraped text"
        if len(text.strip()) < SHORT_TEXT and a["url"]:
            fresh = refetch_text(a["url"])
            if len(fresh) > len(text):
                text, checked = fresh, "Re-fetched from source"
            else:
                checked = "Re-fetch failed (used scraped text)"
        if not (text.strip() or a["title"].strip()):
            continue
        try:
            result = ask_model(api_key, args.model, a, text or a["title"], args.max_chars, usage)
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
        for rec in recs:
            rows.append({
                "Ministry": rec["ministry"], "Project / Initiative": rec["project"],
                "Status": rec["status"], "Funder(s)": rec["funders"], "Amount / Scale": rec["amount"],
                "Partner(s)": rec["partners"], "Region": rec["region"],
                "Funding Date": rec["funding_date"], "Date Basis": rec["date_basis"],
                "Date Evidence (from article)": rec["date_evidence"],
                "Project Period": rec["project_period"], "In 2025-26?": rec["in_2025_26"],
                "Published Date": us_date(a["date"]), "Description": rec["description"],
                "Enmaeya Reporting Link": "", "Confidence": rec["confidence"],
                "Source Feed": a["feed"], "Origin": "Daily run", "Source Checked": checked,
                "Live-page Evidence": "", "Funding Type": rec["funding_type"], "_url": a["url"],
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
