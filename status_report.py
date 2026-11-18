"""Weekly client status report generator for TPMs.

Reads the project tracker (Excel file with tabs "Summary" and "Risks and Issues",
or a folder with Summary.csv and Risks.csv) and writes one status report per client
(Markdown + email-ready HTML). It also prints a "tracker health check" for the TPM.

Usage:
    python status_report.py --tracker sample/T5_Project_Tracker.xlsx --date 2026-11-23
"""
import argparse
import csv
import datetime as dt
import html
import os
import sys

LEVEL = {"high": 3, "medium": 2, "low": 1}
RAG_COLOR = {"red": "#d93025", "amber": "#f29900", "green": "#188038"}


# ------------------------------------------------------------------ reading
def _norm_key(k):
    return str(k or "").strip().lower()


def _rows_from_table(table, required):
    """table = list of lists, first row = header. Returns list of dicts (lower-case keys).

    A row is kept only if ALL `required` columns have a value. This skips empty rows and
    the explanation notes under the tables (a note fills only the first column).
    """
    if not table:
        return []
    head = [_norm_key(h) for h in table[0]]
    out = []
    for row in table[1:]:
        rec = {head[i]: row[i] for i in range(min(len(head), len(row))) if head[i]}
        if not all(str(rec.get(k) or "").strip() for k in required):
            continue
        out.append(rec)
    return out


SUMMARY_REQUIRED = ("project", "client", "status (rag)")
RISKS_REQUIRED = ("id", "project", "description")


def read_tracker(path):
    if os.path.isdir(path):
        def load(name):
            with open(os.path.join(path, name), newline="", encoding="utf-8-sig") as f:
                return list(csv.reader(f))
        return (_rows_from_table(load("Summary.csv"), SUMMARY_REQUIRED),
                _rows_from_table(load("Risks.csv"), RISKS_REQUIRED))

    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    sheets = {ws.title.strip().lower(): ws for ws in wb.worksheets}
    summary = sheets.get("summary")
    risks = next((ws for name, ws in sheets.items() if name.startswith("risk")), None)
    if summary is None or risks is None:
        sys.exit(f"ERROR: {path} needs the tabs 'Summary' and 'Risks and Issues'. Found: {list(sheets)}")
    as_table = lambda ws: [list(r) for r in ws.iter_rows(values_only=True)]
    return _rows_from_table(as_table(summary), SUMMARY_REQUIRED), _rows_from_table(as_table(risks), RISKS_REQUIRED)


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%a %d-%b-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def score(r):
    # Always compute from Probability x Impact. The Score cell may be an Excel formula
    # with no saved value (openpyxl returns None for it), so we do not trust it.
    return LEVEL.get(_norm_key(r.get("probability")), 0) * LEVEL.get(_norm_key(r.get("impact")), 0)


def fmt_date(d):
    return d.strftime("%a %d %b %Y") if d else "not set"


def txt(v):
    return "" if v is None else str(v).strip()


# ------------------------------------------------------------------ logic
def health_check(summary, risks, today):
    """Problems in the tracker itself that the TPM should fix before sending reports."""
    issues = []
    for r in risks:
        rid = txt(r.get("id")) or "(no ID)"
        if _norm_key(r.get("status")) == "closed":
            continue
        for field in ("owner", "action", "due date"):
            if not txt(r.get(field)):
                issues.append(f"{rid}: missing {field}")
        due = to_date(r.get("due date"))
        if due and due < today:
            issues.append(f"{rid}: OVERDUE since {fmt_date(due)} (owner: {txt(r.get('owner'))})")
        if _norm_key(r.get("probability")) not in LEVEL or _norm_key(r.get("impact")) not in LEVEL:
            issues.append(f"{rid}: probability/impact must be High, Medium or Low")
    for s in summary:
        name = txt(s.get("project"))
        if _norm_key(s.get("status (rag)")) not in RAG_COLOR:
            issues.append(f"{name}: status must be Red, Amber or Green")
        upd = to_date(s.get("last updated"))
        if upd is None or (today - upd).days > 7:
            issues.append(f"{name}: summary not updated in the last 7 days")
    return issues


def build_report(proj, risks, today, internal=False):
    name = txt(proj.get("project"))
    mine = [r for r in risks if _norm_key(r.get("project")) in (_norm_key(name), "both")
            and _norm_key(r.get("status")) != "closed"]
    if not internal:
        mine = [r for r in mine if _norm_key(r.get("share with client")) != "no"]
    mine.sort(key=lambda r: (-score(r), to_date(r.get("due date")) or dt.date.max))
    ms_date = to_date(proj.get("milestone date"))
    days = (ms_date - today).days if ms_date else None
    return {
        "project": name,
        "client": txt(proj.get("client")),
        "contact": txt(proj.get("client contact")),
        "rag": txt(proj.get("status (rag)")).title(),
        "reason": txt(proj.get("status reason")),
        "milestone": txt(proj.get("next milestone")),
        "milestone_date": fmt_date(ms_date),
        "days_left": days,
        "golive": fmt_date(to_date(proj.get("go-live date"))),
        "problem": txt(proj.get("main problem")),
        "done": txt(proj.get("done this week")),
        "next": txt(proj.get("next steps")),
        "tpm": txt(proj.get("tpm")),
        "risks": [{
            "id": txt(r.get("id")), "desc": txt(r.get("description")), "type": txt(r.get("type")),
            "level": "High" if score(r) >= 6 else "Medium" if score(r) >= 3 else "Low",
            "owner": txt(r.get("owner")), "action": txt(r.get("action")),
            "due": fmt_date(to_date(r.get("due date"))),
            "overdue": bool(to_date(r.get("due date")) and to_date(r.get("due date")) < today),
        } for r in mine[:6]],
        "today": fmt_date(today),
    }


def to_markdown(rep):
    days = "" if rep["days_left"] is None else f" ({rep['days_left']} days from today)"
    lines = [
        f"# Weekly status – {rep['project']} ({rep['client']})",
        f"**Date:** {rep['today']}  |  **To:** {rep['contact']}  |  **From:** {rep['tpm']}, PSI",
        "",
        f"## Overall status: {rep['rag'].upper()}",
        rep["reason"],
        "",
        f"- **Next milestone:** {rep['milestone']} – {rep['milestone_date']}{days}",
        f"- **Go-live:** {rep['golive']}",
        f"- **Main problem:** {rep['problem']}",
        "",
        "## Done this week", rep["done"], "",
        "## Next steps", rep["next"], "",
        "## Risks and issues",
        "| Level | Item | Action | Owner | Due |",
        "|---|---|---|---|---|",
    ]
    for r in rep["risks"]:
        due = r["due"] + (" (OVERDUE)" if r["overdue"] else "")
        lines.append(f"| {r['level']} | {r['desc']} | {r['action']} | {r['owner']} | {due} |")
    if not rep["risks"]:
        lines.append("| – | No open risks | | | |")
    lines += ["", "Please reply or call me with any questions."]
    return "\n".join(lines) + "\n"


def to_html(rep):
    e = html.escape
    color = RAG_COLOR.get(rep["rag"].lower(), "#555")
    days = "" if rep["days_left"] is None else f" ({rep['days_left']} days from today)"
    rows = "".join(
        f"<tr><td><b>{e(r['level'])}</b></td><td>{e(r['desc'])}</td><td>{e(r['action'])}</td>"
        f"<td>{e(r['owner'])}</td><td>{e(r['due'])}{' <b style=\"color:#d93025\">OVERDUE</b>' if r['overdue'] else ''}</td></tr>"
        for r in rep["risks"]) or "<tr><td colspan=5>No open risks</td></tr>"
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>{e(rep['project'])} – weekly status</title>
<style>body{{font-family:Arial,sans-serif;max-width:820px;margin:24px auto;padding:0 16px;color:#1f2933;background:#fff}}
h1{{font-size:20px;color:#12305a}} h2{{font-size:15px;color:#12305a;margin-top:20px}}
.rag{{display:inline-block;padding:4px 12px;border-radius:4px;color:#fff;background:{color};font-weight:bold}}
table{{border-collapse:collapse;width:100%;font-size:13px}} td,th{{border:1px solid #c8d1dc;padding:5px;vertical-align:top;text-align:left}}
th{{background:#e8eef6}} .muted{{color:#5f6b7a;font-size:13px}}</style></head><body>
<h1>Weekly status – {e(rep['project'])} ({e(rep['client'])})</h1>
<p class="muted">Date: {e(rep['today'])} | To: {e(rep['contact'])} | From: {e(rep['tpm'])}, PSI</p>
<p><span class="rag">{e(rep['rag'].upper())}</span> {e(rep['reason'])}</p>
<ul><li><b>Next milestone:</b> {e(rep['milestone'])} – {e(rep['milestone_date'])}{e(days)}</li>
<li><b>Go-live:</b> {e(rep['golive'])}</li><li><b>Main problem:</b> {e(rep['problem'])}</li></ul>
<h2>Done this week</h2><p>{e(rep['done'])}</p>
<h2>Next steps</h2><p>{e(rep['next'])}</p>
<h2>Risks and issues</h2><table><tr><th>Level</th><th>Item</th><th>Action</th><th>Owner</th><th>Due</th></tr>{rows}</table>
<p>Please reply or call me with any questions.</p></body></html>"""


# ------------------------------------------------------------------ main
def main(argv=None):
    ap = argparse.ArgumentParser(description="Make weekly client status reports from the project tracker.")
    ap.add_argument("--tracker", required=True, help="Excel tracker (.xlsx) or a folder with Summary.csv and Risks.csv")
    ap.add_argument("--date", help="Report date YYYY-MM-DD (default: today)")
    ap.add_argument("--out", default="reports", help="Output folder (default: reports)")
    ap.add_argument("--project", help="Only this project (e.g. Kavach)")
    ap.add_argument("--internal", action="store_true", help="Include internal items (Share with client = No)")
    a = ap.parse_args(argv)

    today = dt.date.fromisoformat(a.date) if a.date else dt.date.today()
    summary, risks = read_tracker(a.tracker)
    if a.project:
        summary = [s for s in summary if _norm_key(s.get("project")) == _norm_key(a.project)]
        if not summary:
            sys.exit(f"ERROR: project '{a.project}' not found in the Summary tab")

    os.makedirs(a.out, exist_ok=True)
    print(f"Report date: {fmt_date(today)}")
    for proj in summary:
        rep = build_report(proj, risks, today, internal=a.internal)
        base = f"{today.isoformat()}_{rep['project']}_{rep['client'].replace(' ', '_')}" + ("_INTERNAL" if a.internal else "")
        for ext, fn in (("md", to_markdown), ("html", to_html)):
            with open(os.path.join(a.out, f"{base}.{ext}"), "w", encoding="utf-8") as f:
                f.write(fn(rep))
        print(f"  [{rep['rag']:5s}] {rep['project']:8s} -> {a.out}/{base}.md / .html  ({len(rep['risks'])} risks shown)")

    issues = health_check(summary, risks, today)
    print("\nTracker health check:" if issues else "\nTracker health check: OK – nothing to fix.")
    for i in issues:
        print("  -", i)


if __name__ == "__main__":
    main()
