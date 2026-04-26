"""
Weekly data refresh script.
Fetches recent events from SEC EDGAR, clinicaltrials.gov, and FDA,
then uses Claude to regenerate the editorial sections.

Usage:
  python update_data.py                   # full refresh
  python update_data.py --editorial-only  # regenerate editorial from existing data
  python update_data.py --no-claude       # fetch data only, skip Claude editorial

Requires:
  pip install anthropic requests
  ANTHROPIC_API_KEY env var set
"""

import json
import os
import sys
import time
import argparse
import logging
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import requests
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "requests", "-q"], check=True)
    import requests

try:
    import anthropic
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "anthropic", "-q"], check=True)
    import anthropic

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

BASE_DIR   = Path(__file__).parent.parent
DATA_FILE  = BASE_DIR / "data" / "data.json"

COMPANIES = [
    "AbbVie", "Johnson & Johnson", "AstraZeneca", "Merck", "Pfizer",
    "Eli Lilly", "Sanofi", "Regeneron", "Gilead Sciences", "Roche",
    "Bristol Myers Squibb", "Novartis", "GSK", "Amgen", "UCB",
    "Biogen", "Novo Nordisk", "ViiV Healthcare", "argenx", "Immunovant"
]

COMPANY_TICKERS = {
    "AbbVie": "ABBV", "Johnson & Johnson": "JNJ", "AstraZeneca": "AZN",
    "Merck": "MRK", "Pfizer": "PFE", "Eli Lilly": "LLY",
    "Sanofi": "SNY", "Regeneron": "REGN", "Gilead Sciences": "GILD",
    "Roche": "RHHBY", "Bristol Myers Squibb": "BMY", "Novartis": "NVS",
    "GSK": "GSK", "Amgen": "AMGN", "Biogen": "BIIB", "Novo Nordisk": "NVO",
}


# ── Data fetching ─────────────────────────────────────────────────────────────

def fetch_sec_filings(days_back=14):
    """Fetch recent 8-K / 6-K (press releases) from SEC EDGAR for pharma companies."""
    cutoff = (datetime.utcnow() - timedelta(days=days_back)).strftime("%Y-%m-%d")
    events = []

    for company, ticker in list(COMPANY_TICKERS.items())[:12]:  # limit to avoid rate limits
        url = (
            f"https://efts.sec.gov/LATEST/search-index?q=%22{ticker}%22"
            f"&dateRange=custom&startdt={cutoff}&forms=8-K,6-K&hits.hits._source=period_of_report,file_date,display_names,form_type,entity_name"
        )
        try:
            r = requests.get(url, timeout=10, headers={"User-Agent": "pharma-landscape/1.0 dominiqueadaniel@gmail.com"})
            if r.status_code == 200:
                hits = r.json().get("hits", {}).get("hits", [])
                for h in hits[:3]:
                    src = h.get("_source", {})
                    events.append({
                        "company": company,
                        "ticker": ticker,
                        "form": src.get("form_type", ""),
                        "filed": src.get("file_date", ""),
                        "description": src.get("display_names", ""),
                    })
            time.sleep(0.3)  # EDGAR rate limit
        except Exception as e:
            log.warning(f"SEC EDGAR fetch failed for {ticker}: {e}")

    return events


def fetch_ct_updates(days_back=14):
    """Fetch recent clinical trial status updates from clinicaltrials.gov."""
    cutoff = (datetime.utcnow() - timedelta(days=days_back)).strftime("%Y-%m-%d")
    terms = [
        "immunology autoimmune", "HIV treatment prevention", "RSV vaccine",
        "IL-23 psoriasis IBD", "JAK inhibitor", "bispecific antibody immunology"
    ]
    results = []
    for term in terms[:3]:
        try:
            url = (
                "https://clinicaltrials.gov/api/v2/studies"
                f"?query.term={requests.utils.quote(term)}"
                f"&filter.advanced=AREA[LastUpdatePostDate]RANGE[{cutoff},MAX]"
                "&fields=nctId,briefTitle,overallStatus,phase,studyFirstPostDate,lastUpdatePostDate,sponsorName"
                "&pageSize=10&sort=LastUpdatePostDate:desc"
            )
            r = requests.get(url, timeout=12, headers={"User-Agent": "pharma-landscape/1.0"})
            if r.status_code == 200:
                studies = r.json().get("studies", [])
                for s in studies:
                    proto = s.get("protocolSection", {})
                    ident = proto.get("identificationModule", {})
                    status = proto.get("statusModule", {})
                    design = proto.get("designModule", {})
                    sponsor = proto.get("sponsorCollaboratorsModule", {}).get("leadSponsor", {})
                    results.append({
                        "nct_id": ident.get("nctId", ""),
                        "title": ident.get("briefTitle", ""),
                        "status": status.get("overallStatus", ""),
                        "phase": design.get("phases", []),
                        "sponsor": sponsor.get("name", ""),
                        "last_updated": status.get("lastUpdatePostDateStruct", {}).get("date", ""),
                    })
            time.sleep(0.5)
        except Exception as e:
            log.warning(f"ClinicalTrials.gov fetch failed for '{term}': {e}")
    return results


def fetch_fda_approvals(days_back=30):
    """Fetch recent FDA drug approvals."""
    try:
        url = "https://api.fda.gov/drug/drugsfda.json?search=submissions.submission_status_date:[now-30d+TO+now]+AND+submissions.submission_type:ORIG&limit=20"
        r = requests.get(url, timeout=12, headers={"User-Agent": "pharma-landscape/1.0"})
        if r.status_code == 200:
            results = r.json().get("results", [])
            approvals = []
            for res in results:
                for sub in res.get("submissions", []):
                    if sub.get("submission_status") == "AP" and sub.get("submission_type") == "ORIG":
                        approvals.append({
                            "application_number": res.get("application_number", ""),
                            "brand_name": res.get("openfda", {}).get("brand_name", [""])[0],
                            "generic_name": res.get("openfda", {}).get("generic_name", [""])[0],
                            "sponsor": res.get("sponsor_name", ""),
                            "approval_date": sub.get("submission_status_date", ""),
                        })
            return approvals
    except Exception as e:
        log.warning(f"FDA approvals fetch failed: {e}")
    return []


# ── Claude editorial generation ───────────────────────────────────────────────

def generate_editorial(current_data, sec_filings, ct_updates, fda_approvals):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        log.error("ANTHROPIC_API_KEY not set — skipping editorial generation")
        return current_data.get("editorial", {})

    client = anthropic.Anthropic(api_key=api_key)

    # Build context summary for Claude
    pipeline_summary = []
    for asset in (current_data.get("pipeline") or [])[:40]:
        pipeline_summary.append(
            f"- {asset.get('company')} | {asset.get('asset')} | {asset.get('phase')} | "
            f"{asset.get('area')} | {asset.get('indication','')[:80]} | PoS: {asset.get('pos','')} | "
            f"Expected: {asset.get('expected_data','')}"
        )

    loe_summary = []
    for l in (current_data.get("loe_watchlist") or []):
        loe_summary.append(
            f"- {l.get('product')} ({l.get('company')}): US LOE {l.get('us_loe','')}, "
            f"impact {l.get('revenue_impact','')}"
        )

    sec_summary = "\n".join(
        f"- {e['company']} filed {e['form']} on {e['filed']}"
        for e in sec_filings[:15]
    ) or "(none retrieved)"

    ct_summary = "\n".join(
        f"- {s['title']} | {s['status']} | {s['phase']} | Sponsor: {s['sponsor']} | Updated: {s['last_updated']}"
        for s in ct_updates[:15]
    ) or "(none retrieved)"

    fda_summary = "\n".join(
        f"- {a.get('brand_name','?')} ({a.get('generic_name','?')}) approved {a.get('approval_date','?')} by {a.get('sponsor','?')}"
        for a in fda_approvals[:10]
    ) or "(none retrieved)"

    today = str(date.today())

    prompt = f"""You are a senior biopharma analyst producing a weekly competitive intelligence update for a dashboard covering immunology and infectious disease.

Today's date: {today}

## Pipeline Summary (selected assets)
{chr(10).join(pipeline_summary)}

## LOE / Biosimilar Watchlist
{chr(10).join(loe_summary)}

## Recent SEC Filings (8-K / 6-K) from past 2 weeks
{sec_summary}

## Recent ClinicalTrials.gov Updates
{ct_summary}

## Recent FDA Approvals
{fda_summary}

---

Please produce a JSON object with exactly these three keys:

1. "key_recent_events": array of 5-6 event objects, each with:
   - "date": string like "2026-04" or "2026-W17"
   - "headline": concise headline (max 12 words)
   - "detail": 2-3 sentence explanation of significance for immunology/ID competitive landscape
   - "companies": array of company names mentioned
   - "tags": array of 3-5 short topic tags

2. "implications": array of 5 implication objects, each with:
   - "theme": 6-10 word theme title
   - "body": 4-6 sentence analytical paragraph explaining the strategic/competitive implication
   - "impact": one of "Transformative", "High", "Medium"
   - "companies_affected": array of company names most affected

3. "commentary": array of 5 company commentary objects, each with:
   - "company": company name (choose 5 of the most strategically interesting from the landscape)
   - "headline": 10-15 word strategic assessment headline
   - "body": 4-6 sentence detailed commentary on strategy and execution quality
   - "rating": one of "Strong Execution", "Stable — Monitoring", "Transitional — Monitoring", "Underperforming"
   - "risk": one of "Low", "Low-Medium", "Medium", "High"

Focus on what has changed in the last 2 weeks vs. the prior period. Be analytically precise and specific about mechanisms, indications, revenue figures, and competitive dynamics. Avoid generic statements.

Return only valid JSON, no markdown fences."""

    log.info("Calling Claude API for editorial generation…")
    try:
        message = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
            system="You are a precise, analytically rigorous biopharma analyst. Return only valid JSON.",
        )
        raw = message.content[0].text.strip()
        # Strip markdown fences if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        editorial = json.loads(raw.strip())
        editorial["last_generated"] = today
        log.info("Editorial generated successfully.")
        return editorial
    except Exception as e:
        log.error(f"Claude API call failed: {e}")
        return current_data.get("editorial", {})


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--editorial-only", action="store_true", help="Only regenerate editorial, don't re-fetch external data")
    parser.add_argument("--no-claude", action="store_true", help="Fetch data only, skip Claude API call")
    parser.add_argument("--days-back", type=int, default=14, help="Days back to look for recent events")
    args = parser.parse_args()

    # Load existing data
    if not DATA_FILE.exists():
        log.error(f"data.json not found at {DATA_FILE}. Run generate_data.py first.")
        sys.exit(1)

    with open(DATA_FILE) as f:
        data = json.load(f)

    log.info(f"Loaded data.json ({len(data.get('pipeline',[]))} pipeline assets)")

    if not args.editorial_only:
        log.info("Fetching SEC EDGAR filings…")
        sec_filings = fetch_sec_filings(args.days_back)
        log.info(f"  {len(sec_filings)} filings retrieved")

        log.info("Fetching ClinicalTrials.gov updates…")
        ct_updates = fetch_ct_updates(args.days_back)
        log.info(f"  {len(ct_updates)} trial updates retrieved")

        log.info("Fetching FDA approvals…")
        fda_approvals = fetch_fda_approvals(args.days_back)
        log.info(f"  {len(fda_approvals)} FDA approvals retrieved")
    else:
        sec_filings = []
        ct_updates = []
        fda_approvals = []

    if not args.no_claude:
        editorial = generate_editorial(data, sec_filings, ct_updates, fda_approvals)
        data["editorial"] = editorial

    # Update metadata
    data["meta"]["last_updated"] = str(date.today())

    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    log.info(f"data.json updated → {DATA_FILE}")


if __name__ == "__main__":
    main()
