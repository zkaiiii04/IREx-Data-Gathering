# Handoff: resuming the design interview

The design interview on streamlining IREx's research data collection is paused at **round 6, unanswered**. These three files hold everything decided so far:

- [GLOSSARY.md](../GLOSSARY.md): the canonical terms. Write in them.
- [discovery-notes.md](discovery-notes.md): the problem, what the two workbooks contain, decisions settled in rounds 1–5, the round-6 questions with recommendations, facts already checked, and topics not yet asked.
- `git log` on this branch: one commit per round.

This file holds only what those three don't: how to resume, and the conversational context.

## Resume

1. Load the `grilling` and `domain-modeling` skills with the Skill tool. (`/grill-with-docs` loads both, but only the user can type it.)
2. Read the glossary and the discovery notes in full.
3. Put round 6 to the user again, worded from the notes. Use the grilling format: numbered ❓ questions, each with a ➡️ recommendation, worded so that "yes" accepts it. Then wait for answers.
4. As answers arrive: update the glossary as soon as a term is settled, move answered questions into a new "Settled (round N)" section, write the next round's frontier as "Open questions (round N+1)", then commit and push. When an answer reverses an earlier decision, say so in the settled entry ("This replaces round X's …").
5. Round 6 Q5 offers an ADR in `docs/adr/` (format in the `domain-modeling` skill). It would record that the main database stays the hand-edited source of truth, with no sync. Write it if the user says yes.
6. The interview is finished when the "Not yet asked" list in the notes is empty and the user confirms there's a shared understanding. Build nothing before that.

## Context from the conversation

- **The user is IREx Admin.** They run the main database with their admin staff and build the KPI Review PowerPoint from its visualisation tabs.
- **The user answers by question number, often by voice.** Expect transcription slips: "commuter meetings" meant committee meetings, and "Lian Foundation" meant Lien Foundation. Read for intent, and confirm only real ambiguity.
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
- **"IREx Member" was a misunderstanding.** It meant the IREx Committee. Reporting is open to all Rehab Staff, so no member list is needed for the form. The user offered a committee list for a later session, but nothing currently depends on it.
- **Microsoft 365 enterprise, no Copilot agents.** Forms and Power Automate are confirmed. Power BI is unconfirmed. Whether TTSH allows Power Apps is unknown, and the user is checking (round 6 Q3).

## The workbooks

The two `.xlsm` files were uploads to the earlier session and are **not in the repo**. Keep them out of it: they hold staff names and project finances. Everything the interview needs from them is summarised in the discovery notes. If a later question needs the files again, ask the user to re-upload them.

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
