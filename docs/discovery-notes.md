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

## Open questions (round 6)

1. **Rows that aren't moved.** Recommended: replace the *Moved* column with a *Decision* column (blank, Import, Reject) and a *Reason* column. IREx Admin sets Import or Reject. The macro imports the Import rows and stamps them *Imported* with the date; Reject rows stay as *Rejected*. Each table is filtered to show blank Decisions only, so IREx Admin sees just what's waiting. Nothing is deleted, so "what happened to my paper?" always has an answer. No automatic email to the Submitter. The words avoid "Accepted", which is already a Publication and Presentation status.
2. **Drop "Checked".** Round 3 had each record arrive as *New* and be marked *Checked*. With the review now happening in the Submissions workbook, moving a row into the main database *is* the check, and a Checked column would change the main database's layout. Recommended: drop New/Checked. Every row in the main database counts. This would replace round 3 Q4 and round 4's "Checked" edit.
3. **Add item button.** Recommended: if Power Apps is available, build the form as a Power Apps canvas app with an *Add another* button per activity type and a list of the items added so far. On submit, the app calls a Power Automate flow that writes the rows to the Submissions workbook using IREx Admin's connection, so submitters never need access to the workbook. Canvas apps that use only standard connectors (Excel Online, SharePoint) are included in Microsoft 365 E licences unless IT has disabled them. To check: open make.powerapps.com with a TTSH login and try to create a blank canvas app. If Power Apps isn't available, fall back to Microsoft Forms. Branching would make it feel like "Add another?" up to a cap, with the last slot saying "Have more? Submit another response". The form is the only part that changes between the two options; the Submissions workbook and the macro stay the same.

   ```
   Power Apps form ──(Add another … Submit)──► Power Automate ──► Submissions workbook
   ```
4. **Name columns in the main database.** Recommended: the Submitter's name goes into each item's person column (1st author / TTSH Contributor, presenter, Media full name, Talent Development recipient, Education trainer, Project PI). IREx Admin adds other Rehab Staff by hand when they want to, as now. When two Rehab Authors both report the same paper, the macro flags the duplicate and IREx Admin rejects one. Roles that may not be the Submitter's (a Co-I reporting a Project led by someone else) get their own question when the per-activity questions are drafted.
5. **ADR.** Record in `docs/adr/` that the main database stays the hand-edited source of truth, with Submissions staged in a separate workbook and moved by IREx Admin: no sync, no SharePoint Lists. Recommended: yes.

## Platform facts checked

These were looked up during the interview, so they don't need to be put to the user again.

- **Microsoft Forms** allows at most 200 questions per form ([Microsoft support](https://support.microsoft.com/topic/ec15323d-92a4-4c33-bf88-3fdb9e5b5fea)), and users report hitting the limit earlier. A Likert question counts once per statement. Forms has no repeating section. Branching goes from a question to a later question or section, or to the end. After submitting, the respondent can be offered "Submit another response".
- **Power Automate Excel Online (Business) connector**: Microsoft's documentation contradicts itself on .xlsm. Older connector docs list only .xlsx and .xlsb. The Office Scripts docs say .xlsm works, chiefly through *Run script*. VBA macros never run from Power Automate. ([connector](https://learn.microsoft.com/connectors/excelonlinebusiness), [Office Scripts](https://learn.microsoft.com/office/dev/scripts/develop/macros-power-automate))
- **Power Apps**: canvas apps that use only standard connectors (SharePoint, Excel Online, Outlook) are covered by Microsoft 365 E1/E3/E5 licences, unless IT has disabled them. Premium connectors (SQL, Dataverse) need a paid licence for every user. A flow that a Power App calls can use its owner's connection rather than the user's, so users don't need access to the target file. These come from community sources ([Microsoft Community](https://techcommunity.microsoft.com/discussions/microsoft-365/power-platform-licenses-and-features-in-office-365-e3-subscription/2553776), [run-only users](https://community.powerapps.com/t5/General-Power-Automate/All-about-quot-Manage-Run-Only-Users-quot/m-p/123305)), so test them in TTSH's tenant. Whether TTSH allows Power Apps is unknown.
- **Microsoft Lists forms**: Microsoft has no official list of the column types they support. A third-party checklist says Person and Choice work ([handsontek](https://lists.handsontek.net/column-type-checklist-microsoft-lists-forms/)). Not needed while SharePoint Lists are out of the design.

## Not yet asked

These come after round 6:

- Producing KPI figures and the KPI Review report: refreshing the main database's queries and pivots, and whether the PowerPoint can be fed from them more directly.
- Rollout: how staff are introduced to the new process, and running it alongside the old workbook for one cycle.
- The exact questions for each activity type: only what submitters know, mapped to the main database's columns. Derived columns (No., the combined text columns, JIF percentile) stay with IREx Admin.

## Next session

Resume by putting round 6 to the user again; [handoff.md](handoff.md) says how.
