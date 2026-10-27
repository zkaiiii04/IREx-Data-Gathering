# Discovery notes: streamlining research data collection

Working record of the design interview. Terms are defined in [GLOSSARY.md](../GLOSSARY.md); this file holds the problem, the decisions made so far and the questions still open.

## The problem

IREx collects its members' research activity (Publications, Presentations, Media Appearances, Talent Development, Education Activities and Projects) twice a year for department KPI tracking at the KPI Review. Today a shared Excel workbook with one sheet per activity type, already holding earlier data, is emailed to staff, and IREx Admin has to keep reminding people to fill it in. The data ends up in a separate main database Excel file.

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

Two files were reviewed: the **staff input workbook** (`IREX_Rehab_Med_Data_Collection_CY2026.xlsm`), which goes out to IREx Members, and the **main database** (`tbl_IRExDB_Use_This.xlsm`), which IREx Admin uses.

### How the staff input workbook relates to the main database

- The staff input workbook is a **copy of the main database** with the team-only tabs removed. It contains every existing record, not a blank form, and IREx Members add rows to it. IREx Admin then copies the new rows into the main database by hand.
- The two have drifted apart. The main database has 24 Publications and 4 Poster/Oral Presentations that the staff input copy doesn't have. The staff input copy has 2 Publications and 1 presentation that the main database doesn't have, and one project is only in the main database.
- **Both project tabs ("Research" and "Innovation") hold all 103 projects**, each with all 39 columns, including grant amounts, IO codes, cost centres, budget, utilisation and balance. Every IREx Member who receives the file can see the finances of every project.
- New rows get typed below the Excel table instead of inside it (e.g. Publications row 99), so the table's pivots and lookups miss them.

### What the main database does besides storing KPI data

It is a project management tool as well as the KPI store. Some of its tabs never go to staff:

- **Gantt Chart**, run by macros (the only VBA in either file loads and saves Gantt data), with the task plan in `DB_PrjGanttChart`.
- **RA and AHS/Nursing resource distribution** across projects, with manpower rates by grade.
- **JIF percentile lookup tables** (`Impact Factor-2025Q1/Q2`, `Impact Factor-2026`, about 32,000 journals). IREx Admin adds each Publication's journal ranking from these.
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

## Settled (round 2)

1. **Where the data lives.** Research Outputs go into SharePoint Lists. The main database pulls them in with *Get Data → From SharePoint List* in place of the tabs filled by hand today, so its pivots and the PowerPoint keep working. Research Projects stay in the main database, maintained by IREx Admin. Staff no longer receive a copy of the main database. (How a new project is reported is open in round 3.)
2. **Which Publications count.** A Publication counts when at least one author is an IREx Member (an *IREx Author*), whether that's the first author or any other author. An external first author doesn't matter as long as an IREx Author is listed.
3. **Reporting period.** Outputs count in the calendar year of their date (publication, presentation or media date). Research Projects count by LOA start year. Each KPI Review shows year-to-date figures: Jan–Jun at mid-year and the full year at year-end.
4. **Statuses.** IREx Admin records and updates statuses, not IREx Members. On submission a member may pick any of the statuses: Under Review / Accepted / Published for a Publication, Accepted / Presented for a Presentation.
5. **Presentation.** "Poster Presentation" becomes *Presentation*, with a type of Poster or Oral and the statuses Accepted → Presented.
6. **Talent Development and Education Activities** are in scope, submitted through the same new process. Talent Development is submitted by the recipient. (Who submits an Education Activity is open in round 3.)
7. **Media Appearance.** One record per outlet, as today, with one or more IREx Members and an optional related project or programme.
8. **One process for everyone.** All IREx Members use the same forms. The forms don't ask for Family Group, because IREx Admin can tell it from the name.
9. **Project terms.** An *Oversight Project* is one IREx oversees but doesn't run. Category "LF" means Lien Foundation.

## Settled (round 3)

1. **No official IREx Member list yet.** The RIE Namelist is *not* the IREx Member list: it includes some people who aren't IREx Members and misses some who are. The user will provide the IREx Member list in a later session. Data collection isn't limited to IREx Members either (who else reports is open in round 4).
2. **Form tool.** Microsoft Forms, with Power Automate copying each response into the matching SharePoint List.
3. **One form for everything, new projects included.** One form covers the Research Outputs, Talent Development, Education Activities and new Projects. The first question picks what is being reported, and the form branches from there. New Projects go to a *Project Notifications* list, and IREx Admin creates the Project in the main database and fills in the finance and admin fields.
4. **Checking submissions.** Each submission arrives as *New*. IREx Admin checks it by hand (duplicates, status, JIF percentile) and marks it *Checked*. Only Checked records count.
5. **Who submits an Education Activity.** Any IREx Member who spoke at the activity or was part of the team that ran it.
6. **Access.** Only IREx Admin can open the lists. Everyone else uses the form only.
7. **Reminders.** Quarterly rather than monthly. Reminders aren't tied to fixed dates: IREx Admin wants to send one about two weeks before committee meetings, and one before 31 Dec, because a separate non-IREx collation exercise happens then.
8. **Project** replaces "Research Project" as the term, with category Research, Innovation or Lien Foundation.

The user questioned old round 3 Q6 (importing existing data). The data flow they have in mind is: form → list → main database. That's correct, so the import is no longer assumed. Round 4 Q3 offers keeping the history in the main database instead.

## Open questions (round 4)

1. **Who reports besides IREx Members**, and which of their records count towards IREx's KPIs as opposed to the separate Dec collation.
   Recommended: the form is open to all TTSH staff. A record counts in IREx KPIs only when it involves at least one IREx Member. Everything is kept so it can also serve the other collation.
2. **Naming the people involved.** Recommended: the form records the submitter automatically from their login. It asks "Which IREx Members were involved?" as a tick list (filled from the user's IREx Member list once provided), plus a free-text "Other contributors" box.
3. **History.** Recommended: no import. Records already in the main database stay there as a frozen history table. A Power Query in the main database appends the list's Checked records to that history, so the pivots show both. Anything submitted after a cut-over date arrives through the form. Old Accepted / Under Review records are updated in the history table by hand.
4. **Where IREx Admin's edits happen.** Data pulled from a list is rewritten on every refresh, so anything typed beside it in Excel gets lost or misaligned.
   Recommended: status changes and *Checked* are set in the list itself. The JIF percentile is filled in automatically by Power Query, which matches the journal name against the Impact Factor tables, and IREx Admin fixes the journal names that don't match.
5. **One item per response.** Microsoft Forms can only jump forward to later sections, so it can't loop back for a second item. Recommended: each response reports one activity, and the thank-you message offers "Submit another response".
6. **Reminder mechanics.** Recommended: a scheduled quarterly reminder flow, plus a reminder flow IREx Admin runs by hand from Power Automate when needed (about two weeks before a committee meeting, and before 31 Dec). Also to confirm: is the *committee meeting* the KPI Review, or a different meeting?
7. **The choices on the first question.** Recommended: Publication / Presentation (poster or oral) / Media Appearance / Talent Development (award or grant) / Education Activity / New Project.

## Not yet asked

These come after round 4:

- Producing KPI figures and the KPI Review report: refreshing the main database's queries and pivots, and whether the PowerPoint can be fed from them more directly.
- Rollout: how staff are introduced to the new process, and running it alongside the old workbook for one cycle.
- The exact questions on each branch of the form.

## Next session

Resume with the user's answers to round 4. The user will also bring the IREx Member list.
