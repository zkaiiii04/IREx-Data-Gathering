# Discovery notes: streamlining research data collection

Working record of the design interview. Terms are defined in [GLOSSARY.md](../GLOSSARY.md); this file holds the problem, the decisions made so far and the questions still open.

## The problem

IREx collects its members' research activity (publications, poster presentations, media appearances, research projects) twice a year for department KPI tracking at the KPI Review. Today a shared Excel workbook with one sheet per activity type, already holding earlier data, is emailed to staff, and the collector has to keep reminding people to fill it in. The data ends up in a separate main database Excel file.

Constraints: Microsoft 365 enterprise, no Copilot agents. Microsoft Forms, SharePoint Lists and Power Automate are available. Power BI availability is unknown.

## Settled (round 1)

1. **Purpose.** The data feeds IREx's department KPIs, reviewed at the bi-annual department meeting (the KPI Review).
2. **Ongoing collection.** IREx Members submit activity when it happens rather than in a bi-annual rush. The KPI Review date becomes a cut-off for reporting, not a collection deadline.
3. **Tools.** Forms, SharePoint Lists and Power Automate are available. Power BI is unconfirmed, so don't depend on it.
4. **Research Project is separate.** It's tracked as its own KPI (count per year, funding, status), separate from Research Outputs.
5. **Who submits.** Each IREx Member submits their own activity.
6. **Co-authored outputs.** One record per output, however many IREx Members contributed. Today it is recorded with the "PI" as main contributor and other IREx Members listed as authors (see open Q2: the "PI" naming needs settling).
7. **Existing data.** A mix: some records get updated over time (e.g. status changes), and some are history that's never touched.
8. **Scale.** About 30+ IREx Members.

## Open questions (round 2)

Each has a recommended answer for the user to confirm or change.

1. **Where the main data lives.** The user asked whether existing data must be moved over: yes, once.
   Recommended: SharePoint Lists become the main database (one each for Publications, Poster Presentations, Media Appearances, Research Projects), with existing data imported once. The KPI workbook pulls from the lists via Excel *Get Data → From SharePoint List* and is refreshed before each KPI Review. Alternative: keep the Excel main database and have Power Automate add Forms answers as new rows. That's simpler, but staff can't update their existing records and the file is easy to break.
   To check: does anything else depend on the current main database workbook?
2. **Lead contributor on a publication.** "PI" is reserved for leading a Research Project. Is the publication's main contributor the PI of a related project, or the lead or first-listed IREx author?
   Recommended: call them the *Lead IREx Author*. They submit the publication, and co-authors only check they're listed.
3. **Recording IREx co-authors.** Recommended: pick IREx Members from the directory (a multi-person field), with outside authors as free text. This allows per-person counts and "outputs I'm on" views.
4. **Which half-year an output counts in.** Which dates bound the two reporting periods, and which date of each record decides its period? Do Research Projects count in every period they're active?
   Recommended (guess): Jan–Jun / Jul–Dec. Outputs count by the date they came out. Projects count in each active period, and funding counts when awarded.
5. **Publication statuses.** Recommended: Submitted → Accepted → Published, with only Published counted in the KPIs.
6. **Who sees whose records.** Recommended: every IREx Member can see all records but edits only the ones they submitted or are listed on.
7. **Current columns.** Column headers of each sheet in both workbooks are needed to design the lists and forms.

## Not yet asked

These come after round 2:

- Reminders: what nudges staff, and when (e.g. a monthly "anything new?" digest, a pre-KPI Review deadline reminder).
- Review: does the collector approve submissions before they count?
- Moving the existing data: cleaning duplicates and name variants during import.
- Producing KPI figures and the KPI Review report (Excel refresh vs Power BI).
- Rollout: how staff are introduced to the new process, and running it alongside the old workbook for one cycle.

## Next session

The user will provide:

- the **main database** Excel file, and
- the **staff input** Excel workbook currently emailed out.

Review both (sheets, columns, status values, how co-authors and projects are recorded) before resuming round 2. Many of the open questions may then be answered from the files and only need confirming.
