# The main database stays the hand-edited source of truth

IREx Admin uses the main database (`tbl_IRExDB_Use_This.xlsm`) every day, for KPI data and for project management: the Gantt chart macros, resource distribution, JIF lookups and the visualisation tabs that feed the KPI Review PowerPoint. Every record in it, old or new, must stay editable by hand in its own sheet. So nothing outside the main database may own or overwrite its tables. Submissions are written by Power Automate into a separate, admin-only Submissions workbook, which has the same columns as the main database's tables. IREx Admin reviews each row there, marks it Import or Reject, and an Import Submissions macro appends the Import rows inside the right tables of the main database. After that they are ordinary rows. There is no sync in either direction.

## Considered Options

- **SharePoint Lists as the source of truth, pulled into the main database with Get Data.** Rejected: a refresh overwrites the query tables, so records could no longer be edited by hand in their own sheets.
- **Importing history into lists, or freezing it in history tables.** Rejected for the same reason: old records must stay editable alongside new ones.
- **Power Automate writing straight into the main database.** Rejected: Microsoft's documentation is inconsistent about the Excel connector and .xlsm files, the file is often open on someone's desktop, and there would be no review step before a row counts.

## Consequences

- Moving a row into the main database is the check. There's no New/Checked status, and every row in the main database counts.
- Rejected Submissions stay in the Submissions workbook with a reason. Nothing is deleted.
- The staff workbook route works the same way. A *Check Staff Workbook* macro compares the shared staff workbook with the main database and saves a highlighted copy (new rows and changed rows in two colours). IREx Admin copies rows into the main database by hand. Nothing writes into the main database automatically.
