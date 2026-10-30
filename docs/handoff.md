# Handoff: resuming the design interview

The design interview on streamlining IREx's research data collection has two phases. The **Collection Workbook** design is confirmed (round 14), written up as a spec in [GitHub issue #1](https://github.com/zkaiiii04/IREx-Data-Gathering/issues/1), and built: see "Building the Collection Workbook" below. The **Form phase** (Power Apps) is parked. These three files hold everything decided so far:

- [GLOSSARY.md](../GLOSSARY.md): the canonical terms. Write in them.
- [discovery-notes.md](discovery-notes.md): the problem, what the two workbooks contain, decisions settled in rounds 1–14, the confirmed Collection Workbook design, facts already checked, and topics not yet asked.
- `git log` on this branch: one commit per round.

This file holds only what those three don't: how to resume, and the conversational context.

## Resume

1. Load the `grilling` and `domain-modeling` skills with the Skill tool. (`/grill-with-docs` loads both, but only the user can type it.)
2. Read the glossary and the discovery notes in full.
3. If resuming the Form phase, put its first round to the user, built from the "Not yet asked" list. Use the grilling format: numbered ❓ questions, each with a ➡️ recommendation, worded so that "yes" accepts it. Then wait for answers.
4. As answers arrive: update the glossary as soon as a term is settled, move answered questions into a new "Settled (round N)" section, write the next round's frontier as "Open questions (round N+1)", then commit and push. When an answer reverses an earlier decision, say so in the settled entry ("This replaces round X's …").
5. ADRs live in `docs/adr/` (format in the `domain-modeling` skill). 0001 records that the main database stays the hand-edited source of truth. Offer more only when a decision meets the skill's three tests.
6. The interview is finished when the "Not yet asked" list in the notes is empty and the user confirms there's a shared understanding. Build nothing before that.

## Building the Collection Workbook

- **The spec** is [GitHub issue #1](https://github.com/zkaiiii04/IREx-Data-Gathering/issues/1).
- **Built, test-first:** `src/IRExCollection.bas` (the module for the Main Database), `tests/IRExCollectionTests.bas` (the tests), and fixtures with invented data made by `tests/fixtures/make_fixtures.py`. The user guide, including the Power Automate reminder flows, is [collection-workbook.md](collection-workbook.md).
- **Testing:** `python3 tests/run_in_libreoffice.py` runs every test headless here. LibreOffice skips a few Excel-only checks, listed in [tests/README.md](../tests/README.md). The user runs the reference run in desktop Excel on Windows and reports the Results sheet. Until that happens, treat the Excel-only paths as unverified: tables, dropdowns, .xlsx output, replacing a file, and the prompts in `ExportStaffCopy` and `CheckCollectionWorkbook`.
- **The real main database:** the last upload was `/root/.claude/uploads/d01558a1-cf84-46ac-8d31-0622debfd040/8a2acf9a-tbl_IRExDB_Use_This.xlsm`. If it's gone, ask for a re-upload. Never commit it, and never copy names or figures from it into fixtures. A smoke run on it (export 2025 then check) worked. It flags the AHINET label row below `tbl_Education` and one duplicated Media row pair.
- **Where the build departs from the spec** (each needs the user's confirmation):
  - The spec says No. is typed on Publications and Presentations. In fact No. is a row-numbering formula there too, as on Media. The export writes it as values, and matching by No. assumes rows weren't sorted or inserted mid-table in the Main Database.
  - The spec reads rows below a table only when there's no gap. The real rows below `tbl_Education` come after 8 blank rows, so the export reads every filled row below, gaps allowed. This also picks up the bold label row.
  - For a row with no Export ID that matches a record, only the cells it fills in are compared.
  - A key match that finds several records is reported as such, naming the Main Database rows.
  - The highlighted copy is made by copying the file first, so AutoSave can't write marks back into the shared Collection Workbook.

## Context from the conversation

- **The staff workbook is built first.** The user needs to start collecting data now, so the Form comes later (round 10).

- **The user is IREx Admin.** They run the main database with their admin staff and build the KPI Review PowerPoint from its visualisation tabs.
- **The user answers by question number, often by voice.** They once said "Power Automate" when they meant Power Apps. Expect transcription slips: "commuter meetings" meant committee meetings, and "Lian Foundation" meant Lien Foundation. Read for intent, and confirm only real ambiguity.
- **Hard constraints, stated firmly by the user:**
  - The main database keeps every sheet and its purpose.
  - Every record, old and new, stays editable by hand in its own sheet.
  - Nothing may own or overwrite its tables. That rules out SharePoint Lists as the source of truth, Power Query tables that refresh overwrites, and frozen history tables.
  - Staff must be able to report several items in one Submission.
  - Everything goes through one form.
- **Decisions reversed during the interview, so don't propose them again:**
  - SharePoint Lists as the database (rounds 2–3).
  - Importing history into lists.
  - The RIE Namelist as the member list.
  - One activity per response.
  - Fixed reminder dates of 30 Jun and 31 Dec.
  - A fixed number of item slots per activity type in Microsoft Forms (round 5). The user wants an add item button. Forms with a cap is only the fallback if Power Apps isn't available.
  - A "Rehab Staff involved" names box on the form (round 5).
  - Outlook distribution lists for reminders (round 7). IREx Admin keeps the list of addresses.
  - A *Merged* Decision value (round 8). The macro highlights likely duplicates, and IREx Admin fixes them by hand.
  - Locked existing rows, a separate New entries table, or extra tabs in the staff workbook (round 10). Keep it simple: staff add rows to the existing sheets.
  - One review queue mixing staff workbook rows with Form rows (round 10). The staff workbook stands alone.
  - Filling any staff name column automatically from the Submitter (rounds 8–9). On the Form, name columns are questions the Submitter fills in (round 11).
  - A snapshot of the export to compare against (round 13). The check compares against today's main database.
  - An import macro for the staff workbook (round 11). A check macro saves a highlighted copy, and IREx Admin copies rows by hand.
- **"IREx Member" was a misunderstanding.** It meant the IREx Committee. Reporting is open to all Rehab Staff, so no member list is needed for the form. The user offered a committee list for a later session, but nothing currently depends on it.
- **Microsoft 365 enterprise, no Copilot agents.** Forms and Power Automate are confirmed. Power BI is unconfirmed. The user can create Power Apps canvas apps (round 14). Sharing an app and calling a flow from it are untested.

## The workbooks

The two `.xlsm` files were uploads to the earlier session and are **not in the repo**. Keep them out of it: they hold staff names and project finances. Everything the interview needs from them is summarised in the discovery notes. If a later question needs the files again, ask the user to re-upload them. The macros' sheet, table and column names were taken from the upload named under "Building the Collection Workbook".

Reading them:
- `openpyxl` crashes on their pivot caches. Patch the parser before loading:
  ```python
  from openpyxl.reader import workbook as wbr
  class Any(dict):
      def __missing__(self, k): return None
  wbr.WorkbookParser.pivot_caches = property(lambda self: Any())
  ```
- `pip install --target <dir> oletools` lets you read the VBA. The only macros are the Gantt chart's load and save routines.

Keep staff names and finance figures out of anything committed. Describe patterns instead, e.g. "one member's name appears in five spellings".
