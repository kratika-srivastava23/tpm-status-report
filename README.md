# TPM Weekly Status Report Tool

A small tool for a Technical Program Manager. It reads the project tracker (T5) and writes a **weekly status report for each client**, in Markdown and as email-ready HTML. It also runs a **tracker health check** that lists overdue actions, risks without an owner or action, and summaries older than 7 days.

**Why it saves time:** on Friday, the TPM updates one tracker. One command then writes every client's report. In this scenario, the missed Ledger Friday report (message 5) would not happen again.

## What it does

| Feature | Detail |
|---|---|
| One report per client | Status (Red/Amber/Green) with the reason, the next milestone and days left, go-live date, main problem, work done, next steps, top risks |
| Client vs internal | Items with `Share with client = No` (for example, "Arjun is overloaded") are left out of client reports. `--internal` includes them for the Group CTO. |
| Risk ranking | Score = Probability × Impact (High 3, Medium 2, Low 1). Highest score first, then the earliest due date. |
| Overdue flag | Actions past their due date are marked **OVERDUE** in the report and in the health check. |
| Input | The Excel tracker (`Summary` + `Risks and Issues` tabs) or a folder with `Summary.csv` and `Risks.csv`. |

## How to run

Requirements: Python 3.9 or newer.

```bash
pip install -r requirements.txt

# All clients, with the report date set to the scenario's Monday
python status_report.py --tracker sample/T5_Project_Tracker.xlsx --date 2026-11-23

# One project, internal version (for the Group CTO)
python status_report.py --tracker sample/T5_Project_Tracker.xlsx --project Kavach --internal

# From CSV files instead of Excel
python status_report.py --tracker sample/csv --date 2026-11-23
```

Reports are written to `reports/` (change it with `--out`). Without `--date`, the tool uses today's date.

Example output:

```
Report date: Mon 23 Nov 2026
  [Red  ] Kavach   -> reports/2026-11-23_Kavach_Nirmal_Finance.md / .html  (4 risks shown)
  [Amber] Ledger   -> reports/2026-11-23_Ledger_Sagar_Publications.md / .html  (3 risks shown)

Tracker health check: OK – nothing to fix.
```

Try `--date 2026-11-26` to see the health check list the overdue actions.

## Tracker format

**Summary** tab columns: Project, Client, Client contact, Status (RAG), Status reason, Next milestone, Milestone date, Go-live date, Main problem, Done this week, Next steps, TPM, Last updated.

**Risks and Issues** tab columns: ID, Project (`Kavach`, `Ledger` or `Both`), Type, Description, Probability, Impact, Score, Owner, Action, Due date, Status (`Open`/`Closed`), Share with client (`Yes`/`No`).

Column names are not case-sensitive. Extra columns are ignored.

## How I used AI

I built this tool with **Claude Code** (Anthropic) in VS Code. I gave it the assignment PDF and the tracker from Task 5.

### Prompts I used

1. *"There is one PDF document present in this folder, so just tell me only what has to be done by you and me."* → Claude Code read the assignment and split the work between the AI and me.
2. *"We have to submit the files as written in this image and start building, and name each file and folder as mentioned in this image."* (with a screenshot of the "How to submit" section) → Claude Code built the files with the task-number names, including this tool.
3. Design that Claude Code proposed and I accepted: *"Read the T5 tracker (Excel or CSV). Write one weekly status report per client in Markdown and HTML. Hide internal risks from client reports. Rank risks by probability × impact. Flag overdue actions. Print a health check of the tracker."*

<!-- Add here any later prompts you used to change the tool. -->

### One error the AI made, and how it was corrected

**Error:** In the first version, the tool created a **third, fake project report** named *"RAG rule: Red = a date is at risk…"*. The Summary tab has an explanation note under the table, in column A. Column A is the "Project" column. The code kept every row with a value in "Project", so it read the note as a project. The health check also reported that this fake project had no valid status.

**How I found it:** I ran the tool on the real tracker and read the output. There were 3 reports instead of 2.

**Fix:** a row is now used only if all of its **required columns** have a value (Summary: Project, Client, Status; Risks: ID, Project, Description). A note fills only one column, so it is skipped. See `_rows_from_table()` in `status_report.py`.

**Lesson:** AI code often works on clean test data but fails on real files. I always run it on the real input and check the result.

### Also checked

The tracker's `Score` column is an Excel formula. If the file was never opened in Excel, the Python library returns no value for a formula cell. So the tool does not read the Score cell. It computes the score from Probability × Impact (see `score()`).
