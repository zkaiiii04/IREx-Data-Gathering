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

## Settled (round 4)

1. **"IREx Member" meant the IREx Committee.** Anyone in rehab reports: rehab doctors, nurses and allied health staff (therapists, MSWs and others), together called *Rehab Staff*. Research activity isn't limited to the committee. The form link is shared with the relevant people and needs no access management. It's open to all TTSH staff. (This replaces round 3's "IREx Member list". A member list isn't needed for reporting.)
2. **Several items per Submission.** Someone should be able to report several outputs in one go. This replaces round 3's "first question picks one activity type".
3. **The main database stays exactly as it is.** IREx Admin keeps using the main database day to day for KPI data and project management. Every sheet keeps its current purpose and layout, and every record, old and new, stays editable by hand in its own sheet. So nothing may own its tables: no SharePoint List as the source of truth, no Power Query table that refresh overwrites, no frozen history. This replaces round 2's "Research Outputs go into SharePoint Lists" and round 3's "admin-only lists". IREx Admin's edits (status, Checked, JIF percentile) are made in the main database as now.
4. **KPI Review = the committee meeting.** It happens about twice a year, with dates that move by a few months.
5. **Reminders.** Two Power Automate flows: a scheduled quarterly reminder, and one IREx Admin runs when needed (about two weeks before a KPI Review, and before 31 Dec for the separate non-IREx collation exercise).
6. **The six activity types on the form** are Publication, Presentation (poster or oral), Media Appearance, Talent Development, Education Activity and New Project.

## Settled (round 5)

1. **Submissions workbook.** Form → Power Automate → a separate Submissions workbook (.xlsx, in a SharePoint folder only IREx Admin can open). It has one table per activity type, with the same columns in the same order as the main database's tables, plus *Submitted by* and *Submitted on* columns. Each item in a Submission becomes its own row. IREx Admin moves the rows they accept into the main database, and from then on those rows are ordinary rows in the main database. Power Automate never writes to the main database.

   ```
   Form ──► Power Automate ──► Submissions workbook (.xlsx, admin-only)
                                one table per activity type, one row per item
                                       │  IREx Admin reviews, then moves rows
                                       ▼
                               Main database (.xlsm), unchanged
   ```
2. **Import Submissions macro.** A new VBA module in the main database, next to the Gantt modules, with no sheet changes. It appends each row inside the right Excel table, records in the Submissions workbook that the row was moved, and flags likely duplicate titles or DOIs. Copy and paste by hand stays possible. The user asked how rows that shouldn't be moved are handled, and wants something cleaner than deleting them: round 6 Q1.
3. **Several items per Submission.** The user rejected a fixed number of slots per activity type as messy. They want an *add item* button, so a submitter can add as many items as they need. Microsoft Forms can't do that, so this depends on whether Power Apps is available (round 6 Q3).
4. **Which Publications count.** A Publication counts when at least one author is Rehab Staff (a *Rehab Author*), in any position. A paper whose only TTSH author is a neurologist doesn't count.
5. **Naming people.** The Submitter is recorded automatically from their login. There's no "Rehab Staff involved" box: the form is sent to Rehab Staff, so the Submitter is the Rehab Staff member. (What fills the main database's other name columns is round 6 Q4.)

## Settled (round 6)

1. **Rows that aren't moved.** The Submissions workbook has a *Decision* column (blank, Import, Reject) and a *Reason* column, in place of *Moved*. The macro imports the Import rows and stamps them *Imported* with the date. Reject rows stay as *Rejected*. Each table is filtered to show blank Decisions only. Nothing is deleted, and the Submitter isn't emailed about a rejection.
2. **No New/Checked status.** Moving a row into the main database is the check, and every row in the main database counts. This replaces round 3's New → Checked step and round 4's "Checked" edit.
3. **Add item button: waiting on the user.** The user can confirm they have Power Automate, and is checking whether they have Power Apps. They also asked what the next best option is if Power Apps isn't available and Forms is too limiting: round 7 Q1.
4. **Name columns.** The Submitter's name goes into each item's person column (TTSH Contributor, presenter, Media full name, Talent Development recipient, Education trainer, Project PI). IREx Admin adds other Rehab Staff by hand, as now. When two Rehab Authors both report the same paper, the macro flags it and IREx Admin rejects one. (A Co-I reporting a Project led by someone else is left for the per-activity questions.)
5. **ADR written:** [0001](adr/0001-main-database-stays-hand-edited.md).

## Settled (round 7)

1. **Fallback if Power Apps isn't available.** The user's answer was "built forms and built excel with automated reminders". This is read as *build both a Form and an Excel input, with automated reminders*, but which Excel and how the two relate needs confirming: round 8 Q1. Power Apps still hasn't been checked.
2. **Letting staff check what's already recorded: KIV.** The need is real, because it's why the current file carries previous data, but it's parked. The proposal on hold: a confirmation email after each Submission, plus an *Export Records* macro writing this year's records, with no finance columns, to a read-only file Rehab Staff can open.
3. **Reminder recipients.** IREx Admin keeps a list of addresses in the reminder flows. No distribution lists.
4. **KPI Review out of scope.** The main database, its pivots and visualisation tabs are unchanged, so the KPI Review PowerPoint is produced as now.

## Settled (round 8)

1. **Form and Excel, side by side.** The current Excel file, with automated reminders, stays alongside the Form. Staff can report through either one. This replaces round 2's "staff no longer receive a copy of the main database". How the Excel file is produced, shared and collected is in round 9 Q1–Q3.
2. **Duplicates.** When something is already recorded and its status changes, the Submitter reports it again. When the Import macro finds a likely duplicate (matching title or DOI), it tells IREx Admin which rows they are and highlights them. IREx Admin checks and fixes them by hand. There's no *Merged* Decision value: Decision stays blank, Import or Reject.
3. **Publication questions.** These are as proposed in round 8. The Submissions workbook keeps the main database's current columns, in the same order. **TTSH Contributor isn't filled in automatically.** IREx Admin fills it in, and *Submitted by* shows who the Submitter was. For Publications, this replaces round 6 Q4. The values for Category of Article are still needed from the user.
4. **Presentation questions** are as proposed: Title, Type (Poster/Oral), Status, Date of Presentation, Conference name, Organising body, Country (default Singapore) and Related Project. Full Name = Submitter. Event Start/End aren't asked.

## Settled (round 9)

1. **Export Staff Copy macro.** At the start of each cycle, a macro in the main database writes a fresh staff Excel file. It leaves out the team-only tabs and the project finance and admin columns, except **Grant Amount** and **Amount Awarded to TTSH**, which stay. Projects show those two amounts plus name, PI, category and status. Staff use the file to check what's already recorded, which resolves round 7's KIV. (Talent Development amounts: round 10 Q1.)
2. **Sharing the file.** It's one .xlsx in SharePoint, opened in the browser, with no macros. Existing records are locked, and new items go in a New entries table. The user asked how many sheets that means and whether it doubles them: round 10 Q2.
3. **One review queue.** The user asked what "next to the Form's rows" meant: round 10 Q3 explains it again.
4. **Media Appearance questions** are as proposed. The current Media tab does have a *Full name* column, so it stays a column. Who fills it: round 10 Q4.
5. **Talent Development questions** are as proposed. The recipient's name isn't filled in automatically.
6. **Education Activity questions** are as proposed. Trainers aren't filled in automatically.
7. **New Project questions** are as proposed, including Grant Amount and Amount Awarded to TTSH.

## Settled (round 10)

1. **Talent Development amounts stay in the staff workbook.** Without them, IREx Admin would have to look the amounts up themselves.
2. **Keep the staff workbook simple.** Staff add new rows to the existing sheets, as now. There's no separate New entries table, no locked rows, and no extra tabs. The process has to cope with rows typed outside the Excel table (round 11 Q2). Staff should be able to search for their own name (round 11 Q5). This replaces round 9's locked records with a New entries table.
3. **The staff workbook stands alone and gets built first.** It isn't linked to the Submissions workbook or the Form review queue. The user needs to start collecting data now, so the staff workbook with automated reminders comes first, and the Form, Submissions workbook and Import Submissions macro come later. This replaces round 9's single review queue. How rows get from the staff workbook into the main database: round 11 Q4.
4. **Names on the Form.** The user asked whether this meant names being filled in automatically from the Form. It's explained again in round 11 Q6.
5. **Rollout starts now, without the Form.** The first cycle sends the staff workbook and reminders only. The Form is added in a later cycle.

## Settled (round 11)

1. **One shared staff workbook** sits in a SharePoint folder that Rehab Staff can edit, and its link goes in every reminder.
2. **The staff workbook has no macros.** It's a plain .xlsx. The macros live in the main database. When IREx Admin checks the staff workbook, the macro reads every filled row on each sheet, inside or below the table. Rows with no *No.* are new.
3. **Macros.** The user asked what the edit-detection macro does and how many macros there are. Round 11 Q3 and Q4 merge into a single macro. The staff workbook route needs **two** new macros, both in the main database:
   - **Export Staff Copy** writes the staff workbook for a chosen year.
   - **Check Staff Workbook** compares the staff workbook with the main database (Q4 below).

   The Form route will add a third macro later. The existing Gantt macros are untouched.
4. **No import macro for the staff workbook.** Instead, *Check Staff Workbook* compares it with the main database and saves a highlighted copy, with **new rows in one colour and changed rows in another**. IREx Admin identifies the rows and copies them into the main database by hand. This replaces round 11's proposed *Import Staff Entries* macro.
5. **Visibility in the shared file.** The user asked whether one person's changes can be kept from showing to everyone else: round 12 Q2.
6. **Names on the Form (later).** Each name column (TTSH Contributor, presenter Full Name, Media Full name, Talent Development Name, Education Trainers, PI) is a question on the Form, left blank for the Submitter to fill in. The name columns arrive holding exactly what the Submitter typed. *Submitted by* is still recorded from the sign-in. This replaces round 5's "no Rehab Staff involved box" and the round 9–10 "IREx Admin fills names".
7. **Build order** as proposed: Export Staff Copy → shared staff workbook → reminder flows → Check Staff Workbook. The Form comes later. **Export Staff Copy asks IREx Admin which year to export.** Projects are exported if they start in that year, are still ongoing in it, or were completed in it.

## Settled (round 12)

1. **What the check marks.** Green rows are new and yellow rows are changed, as proposed. The user asked how far "changed" goes, because small spelling or spacing edits would get flagged: round 13 Q1 and Q2.
2. **Shared file visibility.** Sheet View is what the user wants. Everyone can see the content of the shared file, and that's accepted.
3. **Export year rule** as proposed. Research Outputs and Education go by date in Y, and Talent Development by Year = Y. Projects are included if LOA start ≤ 31 Dec Y and LOA end ≥ 1 Jan Y (or there's no end), plus Applied projects that have no LOA yet. Rows with unreadable dates are included and listed for fixing.
4. **When to export.** The export replaces the **staff workbook** (the user calls it the collection workbook), never the main database. It runs once at the start of each year, after a final check. A mid-year re-export is only done straight after a check and copy-over.
5. **The No. column.** The user thinks No. is typed by hand, and some rows may have none. The user asked which sheets have a No. column. That isn't recorded in these notes (the workbooks aren't in the repo), and round 13 Q3 removes the dependency on it.

## Settled (round 13)

1. **No snapshot. The check compares against today's main database.** If IREx Admin edits a row in the main database after the export, that row shows up yellow. That's accepted.
2. **What counts as a change.** The check ignores spaces at the start and end, double spaces and line-break differences, upper or lower case, date formatting (compared as dates), number formatting, and empty against "(blank)". Everything else is flagged, and each yellow cell has a comment showing the old value. Edge cases are handled as proposed in round 13 Q2.
3. **Export ID.** The export writes a hidden Export ID on every row. Matching uses it, not No. The export lists main database rows that have no No.
4. **File name.** The user wants the file named like the current collection file, `IREX_Rehab_Med_Data_Collection_CY2026.xlsm`. See round 14 Q2.

## Settled (round 14)

1. **ID lookup.** The export saves a small lookup file in the admin-only folder, listing each Export ID with its tab, No. and title. The check uses it to find the record in today's main database: by No. if there is one, otherwise by title. Rows it can't find are listed.
2. **File name.** `IREX_Rehab_Med_Data_Collection_CY<year>.xlsx` if Rehab Med's is the only collection file today. Other Family Groups have their own files, so there's one shared file for everyone: `IREX_Rehab_Data_Collection_CY<year>.xlsx`.
3. **The Collection Workbook design below is confirmed** as the shared understanding. Building starts only when the user gives the go-ahead.
4. **Power Apps appears to be available.** The user can create a canvas app. That makes Power Apps, with an *Add another* button, the choice for the Form phase (round 6 Q3), and Microsoft Forms with caps is no longer needed as the fallback. Still to test: sharing an app with a colleague and calling a Power Automate flow from it.
5. **Category of Article.** The user asked where it's stated. It's a column in the main database's Publications table (`tbl_Publication`). Its values are whatever has been entered there: the column's filter dropdown lists them, and so does `tbl_List`, if the column has a dropdown list.

## After round 14: file name and a second look at the main database

- **One Collection Workbook for everyone.** Rehab Med isn't the only collection file today, so one shared file replaces them all: `IREX_Rehab_Data_Collection_CY<year>.xlsx`.
- **Category of Article does exist, but it's hidden.** It's column M of `tbl_Publication` on `DB_Publications`. Columns J (Full Author Listing), M (Category of Article) and N (Other Remarks) are all hidden. Category of Article holds a single value, "Original Research Paper", in 89 of 116 rows. The rest are blank.
- **Which tabs have a No. column:**
  - `DB_Publications`: typed, filled on all 116 rows, no duplicates.
  - `DB_Poster& Oral Presentation`: typed, filled on all 118 rows, no duplicates.
  - `DB_Media`: a **formula**, so it changes when rows move. Media also isn't an Excel table.
  - `DB_Talent Develop`, `DB_Education` and `DB_PrjGeneralInfo` have no No. column.

  This confirms the Export ID decision.
- **The main database also has rows outside its tables.** `DB_Education` has 4 filled rows below `tbl_Education`. The export has to read them too.
- **`tbl_Publication` starts at row 3.** Rows 1–2 hold the Top 10% / Top 25% COUNTIFS inputs.
- **Project statuses come from `tbllist_PrjStatus`** (17 values, including the Oversight variants). The Status column on `DB_PrjGeneralInfo` uses it as a dropdown.

## Collection Workbook design (confirmed in round 14)

**The Collection Workbook**
- One shared .xlsx in a SharePoint folder that Rehab Staff can edit. It has no macros.
- It holds the chosen year's records from the main database: the same tabs and columns as today, minus the team-only tabs and the finance and admin columns. Grant Amount, Amount Awarded to TTSH and Talent Development amounts stay. There's one Projects tab, and a hidden Export ID on every row.
- A note at the top of each sheet tells staff how to find their records: Ctrl+F or the filter arrow, choosing Sheet View if Excel offers it.
- Staff add new rows below the existing ones, anywhere on the sheet, and edit their own existing rows, for example to update a status.

**Reminders**
- Two Power Automate flows, both linking to the Collection Workbook and sending to an address list that IREx Admin keeps:
  - a scheduled quarterly reminder
  - one IREx Admin runs when needed (about two weeks before a KPI Review, and before 31 Dec)

**Main database: one new VBA module, no sheet changes**
- **Export Staff Copy** asks which year to export.
  - Research Outputs and Education are included by date in that year, and Talent Development by Year.
  - Projects are included when LOA start ≤ 31 Dec and LOA end ≥ 1 Jan (or there's no end), plus Applied projects with no LOA.
  - It lists rows with unreadable dates (and includes them) and rows with no No.
  - It writes the Collection Workbook and the ID lookup.
  - It runs once at the start of each year, after a final check. A mid-year re-export only happens straight after a check and copy-over.
- **Check Collection Workbook** reads every filled row on each sheet, inside or below the table, and compares it with today's main database. Formatting differences are ignored. It saves a highlighted copy in the admin-only folder:
  - **green** = new
  - **yellow** = changed, with the changed cells in bold and a comment showing the old value
  - "already recorded as No. X" on unnumbered rows whose title or DOI matches an existing record
  - a summary sheet with counts, deleted rows and extra columns
- IREx Admin copies rows into the main database by hand. Nothing writes into the main database automatically.

**Parked (Form, later):** a Power Apps form (available, per round 14) feeding a Submissions workbook, with an Import Submissions macro. The question lists for each activity type are settled in rounds 8, 9 and 11, and names are left for the Submitter to fill in. Still open: whether Power Apps is available, and the Category of Article values.

## Platform facts checked

These were looked up during the interview, so they don't need to be put to the user again.

- **Microsoft Forms** allows at most 200 questions per form ([Microsoft support](https://support.microsoft.com/topic/ec15323d-92a4-4c33-bf88-3fdb9e5b5fea)), and users report hitting the limit earlier. A Likert question counts once per statement. Forms has no repeating section. Branching goes from a question to a later question or section, or to the end. After submitting, the respondent can be offered "Submit another response".
- **Power Automate Excel Online (Business) connector**: Microsoft's documentation contradicts itself on .xlsm. Older connector docs list only .xlsx and .xlsb. The Office Scripts docs say .xlsm works, chiefly through *Run script*. VBA macros never run from Power Automate. ([connector](https://learn.microsoft.com/connectors/excelonlinebusiness), [Office Scripts](https://learn.microsoft.com/office/dev/scripts/develop/macros-power-automate))
- **Power Apps**: canvas apps that use only standard connectors (SharePoint, Excel Online, Outlook) are covered by Microsoft 365 E1/E3/E5 licences, unless IT has disabled them. Premium connectors (SQL, Dataverse) need a paid licence for every user. A flow that a Power App calls can use its owner's connection rather than the user's, so users don't need access to the target file. These come from community sources ([Microsoft Community](https://techcommunity.microsoft.com/discussions/microsoft-365/power-platform-licenses-and-features-in-office-365-e3-subscription/2553776), [run-only users](https://community.powerapps.com/t5/General-Power-Automate/All-about-quot-Manage-Run-Only-Users-quot/m-p/123305)), so test them in TTSH's tenant. Whether TTSH allows Power Apps is unknown.
- **Excel macros**: Office blocks VBA by default in files that come from the internet or email (Mark of the Web), and VBA doesn't run in Excel for the web. A macro file sent to staff would need them to unblock it and open it in desktop Excel.
- **Sheet Views** in Microsoft 365 Excel (desktop and web) let one person sort or filter a shared file without changing what others see. Sheet Views don't isolate typed cell values.
- **Excel tables on protected sheets** can't be resized or have rows inserted, so a New entries table on a locked sheet needs its blank rows made in advance. Excel for the web supports protected sheets and data validation dropdowns.
- **Microsoft Lists forms**: Microsoft has no official list of the column types they support. A third-party checklist says Person and Choice work ([handsontek](https://lists.handsontek.net/column-type-checklist-microsoft-lists-forms/)). Not needed while SharePoint Lists are out of the design.

## Not yet asked

The Collection Workbook design is complete. What remains is the **Form phase**, which is parked until the user wants to resume it:

- Power Apps form design: screens, the *Add another* behaviour per activity type, and the call to Power Automate. First, test sharing the app with a colleague.
- The Category of Article values, from the Publications table's column.
- Whether the Form phase's Submissions workbook and Import Submissions macro still fit, or whether the user would rather reuse the Collection Workbook's check-and-copy approach.

## Next session

The Collection Workbook design is confirmed and specified in [GitHub issue #1](https://github.com/zkaiiii04/IREx-Data-Gathering/issues/1). Next is either building it (on the user's go-ahead) or resuming the interview for the Form phase. [handoff.md](handoff.md) says how.
