# Tests for the Collection Workbook macros

The tests run the two macros in `src/IRExCollection.bas` end to end against the fixture workbooks, then check only what IREx Admin would see: the rows and columns in the Collection Workbook, the ID lookup and the export report, and the fills, bold cells, comments, notes and Summary entries in the highlighted copy.

| File | What it is |
|---|---|
| `IRExCollectionTests.bas` | The test module: the runner and one procedure per test |
| `fixtures/MainDatabase.xlsx` | A fake Main Database with the real sheet, table and column names and invented records |
| `fixtures/check/` | A Collection Workbook exported for 2025 and then edited by "Rehab Staff", with its ID lookup |
| `fixtures/make_fixtures.py` | Writes the fixtures. Each planted case is commented there |
| `run_in_libreoffice.py` | Runs the same tests headless on Linux |
| `find_compile_error.py` | Finds the procedure LibreOffice can't compile |

The fixtures hold invented data only. Never put names or figures from the real workbooks in them.

## Running the tests in Excel (the reference run)

Desktop Excel on Windows is what the macros are for, so this is the run that counts.

1. **Make the test workbook (once).**
   - In Excel, create a blank workbook and save it as **Excel Macro-Enabled Workbook** named `IRExCollectionTests.xlsm`, in this `tests` folder, next to `fixtures`.
   - Press **Alt+F11** to open the VBA editor. Choose **File > Import File**, then import `src\IRExCollection.bas` and `tests\IRExCollectionTests.bas`.
   - Back in Excel, choose **Developer > Insert > Button (Form Control)**, draw a button, assign it the macro `RunAllTests` and label it **Run tests**. If there's no Developer tab, turn it on in **File > Options > Customize Ribbon**.
   - Save.
2. **Run.** Press **Run tests**. The tests take about a minute. They write to `tests\out` and open and close workbooks as they go, so leave Excel alone until a **Results** sheet appears.
3. **Read the results.** Each row is PASS or FAIL, with details for every failure. Send the Results sheet back, as a screenshot or the rows pasted as text.

After a change to either `.bas` file, remove both modules from the test workbook (right-click > Remove, don't export), then import them again.

If Excel blocks the macros, unblock the file: right-click `IRExCollectionTests.xlsm` in File Explorer, choose **Properties**, tick **Unblock**, then reopen it.

## Running the tests in LibreOffice (Linux)

```
python3 tests/run_in_libreoffice.py            # every test
python3 tests/run_in_libreoffice.py Check      # tests whose name contains "Check"
python3 tests/run_in_libreoffice.py --trace    # on an error, name the last line reached
```

This loads both modules into LibreOffice Calc in VBA mode and calls `RunOneHeadless` once per test. LibreOffice's VBA support has gaps, so a few checks only run in Excel and show as skipped here:

- **No Excel tables.** LibreOffice has no ListObjects, so the module finds headers by name instead. That's the same path Excel takes for `DB_Media`.
- **No .xlsx from SaveAs.** It always writes the old .xls format, so file format checks are skipped.
- **Dropdowns.** LibreOffice reads data-validation lists in its own form, so dropdown checks are skipped.
- **Replacing.** `Workbook.Close` leaves documents open, which blocks saving over a file in the same test, so the replace check is skipped. The harness closes every document between tests.

Things LibreOffice Basic does differently, which the module avoids:

- `ReDim` on a member of a Type
- `End(xlUp)`, which LibreOffice works out from whichever window is active
- a module-level `Collection` holding arrays
- `Exit For` in deeply nested loops
- variables named `row`, `shared` or `compare`

If the harness says the modules don't compile, run `python3 tests/find_compile_error.py src/IRExCollection.bas` (or the tests module) to find the procedure.

## Changing the fixtures

Edit `fixtures/make_fixtures.py`, run `python3 tests/fixtures/make_fixtures.py`, and commit the regenerated `.xlsx` files with it. The Collection Workbook fixture copies the layout Export Staff Copy writes: a note in A1, headers on row 3 from column A, records from row 4, and a hidden Export ID as the last exported column. It predates leaving No. out, so it still has No. columns, which is how the tests cover checking an older Collection Workbook.
