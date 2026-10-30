# Collection Workbook: setting up and running it

How IREx Admin installs and runs the two macros, and sets up the reminder emails. The design is in [GitHub issue #1](https://github.com/zkaiiii04/IREx-Data-Gathering/issues/1); terms are in [GLOSSARY.md](../GLOSSARY.md).

## Install the macros in the Main Database (once)

1. Make a backup copy of the Main Database.
2. Open the Main Database and press **Alt+F11** to open the VBA editor.
3. Choose **File > Import File** and pick `src\IRExCollection.bas` from this repo. A module called `IRExCollection` appears next to the Gantt chart modules.
4. Close the editor and save the Main Database.

Nothing else in the Main Database changes: no sheet, table, column, formula or button. The Gantt chart macros are untouched.

To get buttons for the macros, choose **File > Options > Quick Access Toolbar**, set *Choose commands from* to **Macros**, and add `ExportStaffCopy` and `CheckCollectionWorkbook`. You can also run both from **View > Macros** (Alt+F8).

To update the macros later, remove the `IRExCollection` module (right-click > Remove, don't export) and import the new file.

## Export Staff Copy

Run it at the start of each year, after a final check of last year's Collection Workbook. A mid-year re-export only makes sense straight after a check and copy-over.

1. Run **ExportStaffCopy** and type the year (it suggests the current one).
2. The first time, pick two folders. The macro remembers them for your Windows login:
   - the **SharePoint folder Rehab Staff can edit**, synced to your PC. The Collection Workbook goes here.
   - the **folder only IREx Admin can open**. The ID lookup and the reports go here.
3. If that year's Collection Workbook is already in the shared folder, the macro asks before replacing it. Replacing it loses every row Rehab Staff typed since your last check.
4. **Close any copy of these files first**, including one opened from SharePoint, Teams or Downloads. Excel can't have two workbooks with the same name open at once, so the macro stops before writing anything and names the open file.

It writes three files:

| File | Folder | What it's for |
|---|---|---|
| `IREX_Rehab_Data_Collection_CY2026.xlsx` | shared | The Collection Workbook. It has no macros. |
| `IREX_Rehab_Data_Collection_CY2026 – ID lookup.xlsx` | admin | Each Export ID with its tab, No., title and match keys. The check needs it. Keep it as long as you keep the Collection Workbook. |
| `IREX_Rehab_Data_Collection_CY2026 – export report <date>.xlsx` | admin | Rows exported per tab, rows whose date or Year can't be read, and rows with no No. |

Then share the Collection Workbook's link: in SharePoint, choose **Share > Copy link**, with *People in TTSH with the link can edit*. That link goes in the reminder flows.

### What it exports

| Tab | Rows | Columns |
|---|---|---|
| Publications | Date of Publication in the year | All except the four JIF columns. Hidden columns stay hidden. |
| Poster & Oral Presentation | Date of Presentation in the year | All |
| Media | Media Publishing Date in the year | All, with No. as plain numbers |
| Talent Development | Year is the year | All, including Grant Amount |
| Education | Year is the year, or the date is in the year when Year is empty | All |
| Projects | LOA Start on or before 31 Dec and LOA End on or after 1 Jan (or no LOA End), or a status starting "Applied" with no LOA Start | Grant Ref No, Project Name, Category, Grant / Collaborator Partners, Grant Type, PI, Co-I/Collaborator, LOA Start Date, LOA End Date, Grant Amount, Grant Amt (Awarded to TTSH), Status |

- **Dates it can read:** real dates, Excel date numbers (also when stored as text), and text such as "20 Aug 25", "20 August 2025", "2025-08-20" or "20/8/2025". A row whose date or Year can't be read, such as "(blank)", "Nov 26(KIV)" or "25 & 26 Feb 23", is exported and listed in the report so you can fix it.
- **Rows outside the tables:** every filled row under a table, within the table's columns, is read, even after blank rows. That includes the rows below `tbl_Education`.
- **Dropdowns:** Education type, audience and Local/Overseas, and Project Status are carried over as fixed lists, with room for new rows.
- **Export ID:** each exported row gets one in a hidden last column (`PUB-0001`, `PRES-0001`, `MED-0001`, `TD-0001`, `EDU-0001`, `PRJ-0001`). Ask Rehab Staff not to delete that column.

## Check Collection Workbook

Run it whenever you want to process entries, for example before each KPI Review.

1. Run **CheckCollectionWorkbook**.
2. Pick the Collection Workbook. The synced copy in the shared folder is fine. The macro copies the file before opening it, so the shared file is never opened, changed or saved.
3. Pick its ID lookup. The macro suggests the matching one.

It saves a highlighted copy in the admin folder, named like `IREX_Rehab_Data_Collection_CY2026 – checked 2026-11-02.xlsx`, and opens it. It compares against the Main Database as it is now, and doesn't change it.

| You see | It means |
|---|---|
| Green row, note "New" | No Export ID, and no Main Database record matches it. Copy it over. |
| Yellow row, note "Changed: No. 6, <title>" | Has an Export ID, and some values differ from the Main Database record. Each changed cell is bold, with a comment holding the Main Database value. |
| Note "Already recorded as No. 6, <title>" | No Export ID, but it matches a record: usually a status update typed as a new row, or a row you've already copied. Cells that differ are bold with a comment, and the row is yellow. Only the cells the person filled in are compared. |
| Note "Already recorded more than once (Main Database rows 9, 13)" | It matches several records, so the Main Database probably has a duplicate. |
| Note "Can't find in Main Database" | Has an Export ID, but its record can't be found, for example a record with no No. whose title you changed. When several records match, the note names their rows. |

The **Summary** sheet counts new, changed, already recorded, deleted and can't-find rows, and extra columns, per tab. Below the counts it lists every marked row, every Export ID that's in the lookup but no longer in the Collection Workbook (deleted), every column staff added, and every value typed outside the exported columns.

What the comparison ignores: spaces at the start or end, double spaces, line breaks, upper or lower case, how a date or number is formatted, and an empty cell against "(blank)". Anything else counts as a change, including a one-letter spelling fix. A record you edited in the Main Database after the export also shows as changed.

**Copy the rows you accept into the Main Database by hand.** Nothing is written into the Main Database automatically.

### How a row finds its record

| Tab | Match keys, tried in order |
|---|---|
| Publications | No., then DOI, then Title of Article |
| Poster & Oral Presentation | No., then title |
| Media | Title together with the media source (one story in several outlets is several Media Appearances) |
| Talent Development | Name together with Award Description |
| Education | Title together with Year |
| Projects | Grant Ref No, then Project Name |

The first key that finds exactly one record wins. A key is skipped when its first value is empty. A row with an Export ID uses the values saved in the ID lookup at export time. A row without one uses its own values.

## Things the real Main Database does that you should know

- **No. is a formula on Publications, Presentations and Media.** Each row numbers itself from the row above. Matching by No. therefore assumes rows haven't been inserted mid-table or sorted between the export and the check. Adding rows at the bottom is fine. If you sort or insert, re-export straight after your next check.
- **The bold "AHINET" label row below `tbl_Education`** has no Year, no readable date and no title. It's exported, listed as an unreadable date, and shows as "can't find" in every check. Moving the label out of the table's columns stops that.
- **`DB_Media` has two rows with the same title and outlet.** Checks name both rows until one is removed.

## Choosing the folders again

The folders are remembered for your Windows login. To pick them again, open the VBA editor (Alt+F11), press **Ctrl+G**, type `ResetCollectionFolders` and press Enter.

## Reminder flows (Power Automate)

Two flows, both owned by IREx Admin. Both send the same email.

**Quarterly reminder (scheduled)**
1. In Power Automate, choose **Create > Scheduled cloud flow**.
2. Set it to repeat every **3 months**, starting on the first send date (for example 15 Jan, 9:00), with the time zone set to Singapore.
3. Add the action **Office 365 Outlook > Send an email (V2)**:
   - **To:** the recipient addresses, separated by semicolons. You keep this list here, in the flow.
   - **Subject** and **Body:** the email below.
4. Save.

**Reminder you send when needed (manual)**
1. Choose **Create > Instant cloud flow**, with the trigger **Manually trigger a flow**.
2. Add the same **Send an email (V2)** action.
3. To send it, open the flow and choose **Run**, for example two weeks before a KPI Review, or before 31 Dec.

**Test each flow** by putting only your own address in **To**, running it, and checking the email. Then add the recipients.

To change the wording, edit the action's Subject and Body. The macros don't need to change.

**Email text**

> **Subject:** IREx research activity: please update the Collection Workbook
>
> Dear colleagues,
>
> Please report your research activity in the IREx Collection Workbook: **[link to the Collection Workbook]**
>
> - **New items:** add a new row on the right tab (Publications, Poster & Oral Presentation, Media, Talent Development, Education or Projects).
> - **Changes:** update your existing row if anything has changed, for example a status.
> - **Finding your records:** if Excel offers Sheet View (View > Sheet View > New), choose it first, so that your filter doesn't change what colleagues see. Then press Ctrl+F and search for your name, or use the filter arrow on the name column.
>
> Questions: **[IREx Admin contact]**
>
> Thank you,
> IREx Admin
