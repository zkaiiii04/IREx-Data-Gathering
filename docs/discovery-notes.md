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

## What the two workbooks show (reviewed in session 2)

Two files were reviewed: the **staff input workbook** (`IREX_Rehab_Med_Data_Collection_CY2026.xlsm`), which goes out to IREx Members, and the **main database** (`tbl_IRExDB_Use_This.xlsm`), which the collector's team uses.

### How the staff input workbook relates to the main database

- The staff input workbook is a **copy of the main database** with the team-only tabs removed. It contains every existing record, not a blank form, and IREx Members add rows to it. The collector then copies the new rows into the main database by hand.
- The two have drifted apart. The main database has 24 Publications and 4 Poster/Oral Presentations that the staff input copy doesn't have. The staff input copy has 2 Publications and 1 presentation that the main database doesn't have, and one project is only in the main database.
- **Both project tabs ("Research" and "Innovation") hold all 103 projects**, each with all 39 columns, including grant amounts, IO codes, cost centres, budget, utilisation and balance. Every IREx Member who receives the file can see the finances of every project.
- New rows get typed below the Excel table instead of inside it (e.g. Publications row 99), so the table's pivots and lookups miss them.

### What the main database does besides storing KPI data

It is a project management tool as well as the KPI store. Some of its tabs never go to staff:

- **Gantt Chart**, run by macros (the only VBA in either file loads and saves Gantt data), with the task plan in `DB_PrjGanttChart`.
- **RA and AHS/Nursing resource distribution** across projects, with manpower rates by grade.
- **JIF percentile lookup tables** (`Impact Factor-2025Q1/Q2`, `Impact Factor-2026`, about 32,000 journals). The collector adds each Publication's journal ranking from these.
- **Visualisation tabs** (one per team member), built from pivot tables and slicers. They feed the PowerPoint for the KPI Review.
- `tbl_List` contains a table of SharePoint list GUIDs (`tbl_PrjGeneralInfo`, `tbl_ProjectIOCode`, `tbl_ProjectUpdatesConcern`, `tbl_ProjectUpcomingPlan`, `tbl_ProjectUtilisation`), which suggests the project data was connected to SharePoint lists at some point.

This answers round 2's "to check": other work depends on the main database, so it can't simply be replaced by SharePoint Lists.

### KPIs the visualisation tabs report

- Publication count by year of the publication date, split by status (Published / Accepted / Under Review). Count in the top 10% and top 25% of the journal's field by JIF percentile.
- Presentation count by year, split by status (Accepted / Presented).
- Research Project count by year of the LOA (Letter of Award) start date, split into Innovation and Research. The note on the tab says projects count if their LOA starts or ends from 2021.
- Grant totals (total quantum, and the amount awarded to TTSH) by grant type, grant partner and Innovation/Research.
- Counts of Talent Development awards and of R/I/E (Researcher/Innovator/Educator) roles from the namelist.

All grouping is by **calendar year**, sometimes broken down further into quarters or months. Nothing in the files splits the year into halves. The file name says "CY2026".

### Activity types and their columns

| Tab | Rows (main database) | Columns that matter | Notes |
|---|---|---|---|
| Projects (`tbl_PrjGeneralInfo`) | 107 | Grant Ref No, Project Name, Project Alias, Category (Research / Innovation / LF), Grant Partner, Grant Group, Grant Type (Intramural / Extramural / Intramural – Talent Development), PI, Co-I/Collaborator, LOA Start/End, Grant Amount, Amount Awarded to TTSH, Status, Progress Update, DSRB Ref/dates | 41 columns, about 25 of them finance or admin (IO codes, budget, utilisation, manpower, cost centre, RA/admin support). Several columns are copies of other columns joined into one text ("LOA Start & End Date", "Grant Amt & Awarded to TTSH", "Project Name & Grant/Partners"). |
| Publications (`tbl_Publication`) | 116 | Title, 1st Author Name, TTSH Contributor, Status, Date of Publication, PMID, Journal Citation, Associated Project, DOI, Category of Article, Journal Name, JIF percentile (2025, 2026) | The 1st author is often not an IREx Member. "TTSH Contributor" is a free-text list, one name per line, and includes TTSH staff outside IREx. "Full Author Listing" and "Other Remarks" are always empty. DOI is filled in 33 of 116 rows. |
| Poster & Oral Presentation | 118 | Title, Full Name (one presenter), Status, Date of Presentation, Event Start/End, Related Project, Conference Name, Organisation, Country | **Includes oral presentations**, not only posters. Country is filled in 9 of 118 rows. |
| Media | 23 | Full name, Platform (Digital News / Social Media / Websites & Blogs / Newspaper / Radio / Press Release), Date, Title & Description, Source, Link, Associated Project | Not an Excel table. "Full name" is sometimes a programme ("Stroke Rehab Ecosystem") rather than a person. One press event becomes 14 rows, one per outlet. |
| Talent Development | 19 | Name, Family Group, Award Description, Project, LOA Start/End, Grant Amount, Year, Category (Talent Development / Special Award) | Covers personal awards and personal grants. |
| Education | 19 | Date, Year, Type (Workshop/Masterclass, Webinar, Presentation/Sharing), Audience (Professional / Patients, Care Givers & Public), Title, Trainers, Partner, Participants, Local/Overseas, Organisers | Hidden in the staff input workbook. Only in the main database. |
| RIE Namelist (`tbl_RIE_List`) | 36 | Name, Family Group (Rehab Med / Rehab AHS / Rehab Nursing / MSW), R/I/E role, Institution (TTSH / WH), Active/Inactive | The closest thing to a directory of IREx Members. |

### Data quality

- **Names are free text, and each person appears in several forms.** The PI column has 68 distinct spellings, but the namelist has only 36 people. One senior member's name appears in at least five spellings (with and without title, given and family names in either order, in capitals). Per-person counts are unreliable as a result.
- **Statuses aren't controlled.** Values include "In progress" / "In Progress", "Published " with a trailing space, and "Accpeted". One status list covers both projects and outputs (Presented, Published, Applied, Awarded, In Progress, …), and each project status also has an "(Oversight)" variant.
- Some date cells hold Excel serial numbers or text ("20 Aug 25") instead of dates, and others hold "(blank)".
- One Publication is in the main database twice.

## Open questions (round 2, revised after reviewing the files)

The files answered the old Q7 (columns) and the "to check" under Q1. The other questions have been revised to match what the files show; the round is put to the user in the session.

1. **Where the data lives.** The main database can't be replaced, because the Gantt chart, resource distribution and visualisations depend on it.
   Recommended: Research Outputs (and Talent Development / Education if in scope) go to SharePoint Lists filled from Forms. The main database pulls them in with *Get Data → From SharePoint List* in place of its hand-filled tabs, so the existing pivots keep working. Research Projects stay in the main database, maintained by the team, and PIs send new projects and status updates through a form.
2. **Lead contributor on a publication.** The files record a "1st Author Name" (often external) plus a free-text "TTSH Contributor" list. No publication field is called "PI".
3. **Contributors outside IREx.** "TTSH Contributor" includes TTSH staff who aren't IREx Members. Do they count?
4. **Reporting period.** The files group by calendar year. Is the KPI Review cut-off half-yearly?
5. **Publication statuses.** The values in use are Under Review → Accepted → Published, and the KPI pivot reports all three.
6. **Presentations.** The tab holds both poster and oral presentations, but the glossary only has "Poster Presentation".
7. **Talent Development and Education** are tracked but aren't in the glossary or round 1.
8. **Media granularity.** One press event currently becomes one row per outlet.
9. **Family group workbooks.** The file name is "Rehab_Med". Does each family group get its own copy?
10. **"(Oversight)" project statuses**, **project Category "LF"**, and who sees project finances.

## Not yet asked

These come after round 2:

- Reminders: what nudges staff, and when (e.g. a monthly "anything new?" digest, a pre-KPI Review deadline reminder).
- Review: does the collector approve submissions before they count?
- Moving the existing data: cleaning duplicates and name variants during import.
- Producing KPI figures and the KPI Review report (Excel refresh vs Power BI).
- Rollout: how staff are introduced to the new process, and running it alongside the old workbook for one cycle.

## Next session

Resume with the user's answers to the revised round 2.
