Attribute VB_Name = "IRExCollectionTests"
' Tests for the IRExCollection module (Export Staff Copy and Check Collection Workbook).
'
' Each test runs one macro end to end against the fixture workbooks in tests\fixtures
' and asserts only on what IREx Admin sees in the files it writes. Outputs go to
' tests\out, which is cleared test by test.
'
' Excel: put this module and IRExCollection in the test workbook (tests\README.md)
' and press Run tests, which calls RunAllTests.
' LibreOffice: python3 tests/run_in_libreoffice.py calls RunOneHeadless for each test.
Option Explicit

Private Const GREEN As Long = 13561798      ' RGB(198, 239, 206)
Private Const YELLOW As Long = 10284031     ' RGB(255, 235, 156)

Private mFixtures As String
Private mOut As String
Private mInLibreOffice As Boolean
Private mFailures As String
Private mSkips As String
Private mNames() As String
Private mOutcomes() As String
Private mMessages() As String
Private mCount As Long

' ---------------------------------------------------------------------------
' Runners
' ---------------------------------------------------------------------------

Public Sub RunAllTests()
    Dim sep As String, name As Variant
    sep = Application.PathSeparator
    Setup ThisWorkbook.Path & sep & "fixtures" & sep, ThisWorkbook.Path & sep & "out" & sep, False
    For Each name In TestNames()
        RunTest CStr(name)
    Next name
    WriteResultsSheet
End Sub

Public Function TestNamesHeadless() As String
    TestNamesHeadless = Join(TestNames(), "|")
End Function

' Runs one test and returns its result line. The LibreOffice harness calls this
' once per test, so that it can close the documents each test leaves open.
Public Function RunOneHeadless(ByVal fixtures As String, ByVal outDir As String, ByVal name As String) As String
    Setup fixtures, outDir, True
    RunTest name
    RunOneHeadless = mOutcomes(1) & "  " & mNames(1)
    If mMessages(1) <> "" Then RunOneHeadless = RunOneHeadless & vbLf & "      " & Replace(mMessages(1), vbLf, vbLf & "      ")
End Function

Private Function TestNames() As Variant
    TestNames = Array( _
        "Export_writes_the_six_activity_tabs_to_a_workbook_named_for_the_year", _
        "Export_takes_Publications_whose_date_falls_in_the_year", _
        "Export_takes_Presentations_and_Media_Appearances_whose_date_falls_in_the_year", _
        "Export_takes_Talent_Developments_whose_Year_is_the_year", _
        "Export_takes_Education_Activities_by_Year_or_by_date_when_Year_is_empty", _
        "Export_takes_Projects_live_during_the_year_and_Applied_ones_with_no_LOA", _
        "Export_keeps_every_Publications_column_but_the_four_JIF_columns", _
        "Export_keeps_columns_hidden_that_are_hidden_in_the_Main_Database", _
        "Export_keeps_only_the_columns_a_submitter_needs_on_Projects", _
        "Export_keeps_every_column_on_the_other_tabs", _
        "Export_writes_No_as_plain_values", _
        "Export_writes_a_hidden_unique_Export_ID_on_every_row", _
        "Export_keeps_text_that_looks_like_a_number_or_date_as_text", _
        "Export_puts_a_how_to_find_your_records_note_above_each_table", _
        "Export_carries_the_dropdown_lists_over_as_fixed_lists", _
        "Export_puts_the_rows_in_an_Excel_table", _
        "Export_writes_an_ID_lookup_with_each_rows_tab_No_title_and_match_keys", _
        "Export_report_counts_the_rows_exported_per_tab", _
        "Export_report_lists_rows_whose_date_or_Year_cant_be_read", _
        "Export_report_lists_Main_Database_rows_with_no_No", _
        "Export_replaces_an_existing_Collection_Workbook_only_when_told_to", _
        "Export_leaves_the_Main_Database_as_it_was")
End Function

Private Sub Setup(ByVal fixtures As String, ByVal outDir As String, ByVal inLibreOffice As Boolean)
    mFixtures = fixtures
    mOut = outDir
    mInLibreOffice = inLibreOffice
    mCount = 0
    ReDim mNames(1 To 100)
    ReDim mOutcomes(1 To 100)
    ReDim mMessages(1 To 100)
    EnsureFolder mOut
End Sub

Private Sub Dispatch(ByVal name As String)
    Select Case name
        Case "Export_writes_the_six_activity_tabs_to_a_workbook_named_for_the_year": Export_writes_the_six_activity_tabs_to_a_workbook_named_for_the_year
        Case "Export_takes_Publications_whose_date_falls_in_the_year": Export_takes_Publications_whose_date_falls_in_the_year
        Case "Export_takes_Presentations_and_Media_Appearances_whose_date_falls_in_the_year": Export_takes_Presentations_and_Media_Appearances_whose_date_falls_in_the_year
        Case "Export_takes_Talent_Developments_whose_Year_is_the_year": Export_takes_Talent_Developments_whose_Year_is_the_year
        Case "Export_takes_Education_Activities_by_Year_or_by_date_when_Year_is_empty": Export_takes_Education_Activities_by_Year_or_by_date_when_Year_is_empty
        Case "Export_takes_Projects_live_during_the_year_and_Applied_ones_with_no_LOA": Export_takes_Projects_live_during_the_year_and_Applied_ones_with_no_LOA
        Case "Export_keeps_every_Publications_column_but_the_four_JIF_columns": Export_keeps_every_Publications_column_but_the_four_JIF_columns
        Case "Export_keeps_columns_hidden_that_are_hidden_in_the_Main_Database": Export_keeps_columns_hidden_that_are_hidden_in_the_Main_Database
        Case "Export_keeps_only_the_columns_a_submitter_needs_on_Projects": Export_keeps_only_the_columns_a_submitter_needs_on_Projects
        Case "Export_keeps_every_column_on_the_other_tabs": Export_keeps_every_column_on_the_other_tabs
        Case "Export_writes_No_as_plain_values": Export_writes_No_as_plain_values
        Case "Export_writes_a_hidden_unique_Export_ID_on_every_row": Export_writes_a_hidden_unique_Export_ID_on_every_row
        Case "Export_keeps_text_that_looks_like_a_number_or_date_as_text": Export_keeps_text_that_looks_like_a_number_or_date_as_text
        Case "Export_puts_a_how_to_find_your_records_note_above_each_table": Export_puts_a_how_to_find_your_records_note_above_each_table
        Case "Export_carries_the_dropdown_lists_over_as_fixed_lists": Export_carries_the_dropdown_lists_over_as_fixed_lists
        Case "Export_puts_the_rows_in_an_Excel_table": Export_puts_the_rows_in_an_Excel_table
        Case "Export_writes_an_ID_lookup_with_each_rows_tab_No_title_and_match_keys": Export_writes_an_ID_lookup_with_each_rows_tab_No_title_and_match_keys
        Case "Export_report_counts_the_rows_exported_per_tab": Export_report_counts_the_rows_exported_per_tab
        Case "Export_report_lists_rows_whose_date_or_Year_cant_be_read": Export_report_lists_rows_whose_date_or_Year_cant_be_read
        Case "Export_report_lists_Main_Database_rows_with_no_No": Export_report_lists_Main_Database_rows_with_no_No
        Case "Export_replaces_an_existing_Collection_Workbook_only_when_told_to": Export_replaces_an_existing_Collection_Workbook_only_when_told_to
        Case "Export_leaves_the_Main_Database_as_it_was": Export_leaves_the_Main_Database_as_it_was
        Case Else: Fail "no such test"
    End Select
End Sub

Private Sub RunTest(ByVal name As String)
    mFailures = ""
    mSkips = ""
    RunGuarded name
    CloseTestWorkbooks
    mCount = mCount + 1
    If mCount > UBound(mNames) Then
        ReDim Preserve mNames(1 To mCount + 50)
        ReDim Preserve mOutcomes(1 To mCount + 50)
        ReDim Preserve mMessages(1 To mCount + 50)
    End If
    mNames(mCount) = Replace(name, "_", " ")
    If mFailures <> "" Then
        mOutcomes(mCount) = "FAIL"
        mMessages(mCount) = Mid(mFailures, 2)
    Else
        mOutcomes(mCount) = "PASS"
        If mSkips <> "" Then mMessages(mCount) = "skipped in LibreOffice: " & Mid(mSkips, 3)
    End If
End Sub

' Runs a test, recording an error that escapes it as a failure.
Private Sub RunGuarded(ByVal name As String)
    On Error GoTo escaped
    Dispatch name
    Exit Sub
escaped:
    Fail "ERROR " & Err.Number & ": " & Err.Description
End Sub

Private Sub WriteResultsSheet()
    Dim ws As Object, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Results")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Results"
    End If
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("Result", "Test", "Details")
    ws.Range("A1:C1").Font.Bold = True
    For i = 1 To mCount
        ws.Cells(i + 1, 1).Value = mOutcomes(i)
        ws.Cells(i + 1, 2).Value = mNames(i)
        ws.Cells(i + 1, 3).Value = mMessages(i)
        If mOutcomes(i) = "PASS" Then
            ws.Cells(i + 1, 1).Interior.Color = GREEN
        Else
            ws.Cells(i + 1, 1).Interior.Color = RGB(255, 199, 206)
        End If
    Next i
    ws.Columns("A:C").AutoFit
    ws.Activate
End Sub

' ---------------------------------------------------------------------------
' Export Staff Copy
' ---------------------------------------------------------------------------

Private Sub Export_writes_the_six_activity_tabs_to_a_workbook_named_for_the_year()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "IREX_Rehab_Data_Collection_CY2025.xlsx", Dir(SharedFolder() & "*.xlsx"), "file in the shared folder"
    AssertEqual "Publications|Poster & Oral Presentation|Media|Talent Development|Education|Projects", SheetNames(cw), "tabs"
    If mInLibreOffice Then
        Skip "file format and macros"
    Else
        AssertEqual 51, cw.FileFormat, "saved as .xlsx"
        AssertEqual False, cw.HasVBProject, "no macros"
    End If
End Sub

' Dates that can be read (a text date, a date number stored as text) count by
' their year; "(blank)" can't be read, so that row is kept.
Private Sub Export_takes_Publications_whose_date_falls_in_the_year()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "Robotic gait training after stroke: a pilot study|Spasticity outcomes after botulinum toxin|" & _
        "Caregiver burden in spinal cord injury|Upper limb telerehabilitation|Dysphagia screening in acute stroke", _
        ColumnText(cw.Worksheets("Publications"), "Title of Article"), "titles"
End Sub

Private Sub Export_takes_Presentations_and_Media_Appearances_whose_date_falls_in_the_year()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "Early mobilisation in intensive care|Wheelchair skills training|Virtual reality balance training for older adults", _
        ColumnText(cw.Worksheets("Poster & Oral Presentation"), "Title of Presentation/Topic"), "Presentations"
    ' One story carried by three outlets is three Media Appearances; "Nov 26(KIV)" can't be read, so it's kept.
    AssertEqual "The Test Times|The Test Times Print|Test Radio 99|Test Blog", _
        ColumnText(cw.Worksheets("Media"), "Name of Media Published Source (For eg The Straits Times or online eg ChannelnewsAsia.com)"), "Media outlets"
End Sub

' A Talent Development with no Year can't be placed, so it's kept.
Private Sub Export_takes_Talent_Developments_whose_Year_is_the_year()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "Test Fellowship Award|Test Teaching Award|Test Nursing Award", _
        ColumnText(cw.Worksheets("Talent Development"), "Award Description"), "awards"
End Sub

' Rows typed below tbl_Education, after a gap, are read too: the bold label
' row there has no Year and no readable date, so it's kept like any unreadable row.
Private Sub Export_takes_Education_Activities_by_Year_or_by_date_when_Year_is_empty()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "Spasticity management workshop|Pressure injury prevention webinar|Caregiver training sharing|" & _
        "Hand therapy splinting course||Community rehab workshop", _
        ColumnText(cw.Worksheets("Education"), "Title of Course/Educational Materials/Educational Events/Outreach"), "titles"
    AssertEqual "PARTNER PROGRAMME", CellUnder(cw.Worksheets("Education"), "Date/Year/Period", 5), "label row below the table"
End Sub

' Live: LOA Start on or before 31 Dec and LOA End on or after 1 Jan, or no LOA End.
Private Sub Export_takes_Projects_live_during_the_year_and_Applied_ones_with_no_LOA()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "Robot-assisted gait study|Home exercise app|Community stroke registry|Exoskeleton feasibility|" & _
        "Regional falls network|Pressure mapping pilot", _
        ColumnText(cw.Worksheets("Projects"), "Project Name"), "projects"
End Sub

Private Sub Export_keeps_every_Publications_column_but_the_four_JIF_columns()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual Squash("No.|Title of Article|1st Author Name|TTSH Contributor|Status|Date of Publication|pmid|Journal Citation|" & _
        "Associated Research/Innovation Project (If applicable)|Full Author Listing (Optional)|DOI|Category of Article|" & _
        "Other Remarks|Journal Name|Export ID"), HeaderList(cw.Worksheets("Publications")), "Publications headers"
End Sub

Private Sub Export_keeps_columns_hidden_that_are_hidden_in_the_Main_Database()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual Squash("Associated Research/Innovation Project (If applicable)|Full Author Listing (Optional)|" & _
        "Category of Article|Other Remarks|Export ID"), HiddenHeaders(cw.Worksheets("Publications")), "hidden on Publications"
    AssertEqual Squash("Title of Related Research, Innovation Projects /Programmes|Export ID"), _
        HiddenHeaders(cw.Worksheets("Poster & Oral Presentation")), "hidden on Presentations"
    AssertEqual "export id", HiddenHeaders(cw.Worksheets("Media")), "hidden on Media"
End Sub

Private Sub Export_keeps_only_the_columns_a_submitter_needs_on_Projects()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual Squash("Grant Ref No|Project Name|Category|Grant / Collaborator Partners|Grant Type|PI|Co-I/Collaborator|" & _
        "LOA Start Date|LOA End Date|Grant Amount|Grant Amt (Awarded to TTSH)|Status|Export ID"), _
        HeaderList(cw.Worksheets("Projects")), "Projects headers"
    AssertEqual "500000|80000|250000|60000|90000|30000", ColumnText(cw.Worksheets("Projects"), "Grant Amount"), "Grant Amount"
End Sub

Private Sub Export_keeps_every_column_on_the_other_tabs()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual Squash("No.|Title of Presentation/Topic|Full Name|Status|Date of Presentation|Start Date of Event|End Date of Event|" & _
        "Title of Related Research, Innovation Projects /Programmes|Name of Conference/ Symposium/ Seminar|Name of Organisation|" & _
        "Country where the conference/sympoisum/Seminar was held|Other Remarks|Export ID"), _
        HeaderList(cw.Worksheets("Poster & Oral Presentation")), "Presentations"
    AssertEqual Squash("No.|Full name|Platform of Media|Media Publishing Date|" & _
        "Title of Media Published & Description (Briefly provide the description and or purpose of the published media)|" & _
        "Name of Media Published Source (For eg The Straits Times or online eg ChannelnewsAsia.com)|" & _
        "Website link to media/ Source (If applicable)|Associated Project Title (If applicable)|Other Remarks|Export ID"), _
        HeaderList(cw.Worksheets("Media")), "Media"
    AssertEqual Squash("Name|Family Group|Award Description|Project Name|LOA Start Date|LOA End Date|Grant Amount|Year|Category|Export ID"), _
        HeaderList(cw.Worksheets("Talent Development")), "Talent Development"
    AssertEqual "50000|-|-", ColumnText(cw.Worksheets("Talent Development"), "Grant Amount"), "Talent Development Grant Amount"
    AssertEqual Squash("Date/Year/Period|Year|Type of Material/Event|" & _
        "Area(s) of Focus (Professional vs Patients/Care Givers/General Public)|" & _
        "Title of Course/Educational Materials/Educational Events/Outreach|Family Group|Trainer(s)/Speaker(s)|" & _
        "Collaboration Partner(s), if any|Number of Participants, if applicable|Local/Overseas|Country (if overseas)|Organisers|Export ID"), _
        HeaderList(cw.Worksheets("Education")), "Education"
End Sub

' No. is a formula in the Main Database; the export writes the numbers it shows.
Private Sub Export_writes_No_as_plain_values()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "1|2|3|5", ColumnText(cw.Worksheets("Media"), "No."), "Media No."
    AssertEqual "1|3|4|5|6", ColumnText(cw.Worksheets("Publications"), "No."), "Publications No."
    AssertEqual "1|2|4", ColumnText(cw.Worksheets("Poster & Oral Presentation"), "No."), "Presentations No."
    AssertEqual False, HeaderCell(cw.Worksheets("Media"), "No.").Offset(1, 0).HasFormula, "Media No. is not a formula"
End Sub

Private Sub Export_writes_a_hidden_unique_Export_ID_on_every_row()
    Dim cw As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    AssertEqual "PUB-0001|PUB-0002|PUB-0003|PUB-0004|PUB-0005", ColumnText(cw.Worksheets("Publications"), "Export ID"), "Publications"
    AssertEqual "PRES-0001|PRES-0002|PRES-0003", ColumnText(cw.Worksheets("Poster & Oral Presentation"), "Export ID"), "Presentations"
    AssertEqual "MED-0001|MED-0002|MED-0003|MED-0004", ColumnText(cw.Worksheets("Media"), "Export ID"), "Media"
    AssertEqual "TD-0001|TD-0002|TD-0003", ColumnText(cw.Worksheets("Talent Development"), "Export ID"), "Talent Development"
    AssertEqual "EDU-0001|EDU-0002|EDU-0003|EDU-0004|EDU-0005|EDU-0006", ColumnText(cw.Worksheets("Education"), "Export ID"), "Education"
    AssertEqual "PRJ-0001|PRJ-0002|PRJ-0003|PRJ-0004|PRJ-0005|PRJ-0006", ColumnText(cw.Worksheets("Projects"), "Export ID"), "Projects"
End Sub

Private Sub Export_keeps_text_that_looks_like_a_number_or_date_as_text()
    Dim cw As Object, ws As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    Set ws = cw.Worksheets("Publications")
    AssertEqual "20 Aug 25", CellUnder(ws, "Date of Publication", 2), "text date"
    AssertEqual vbString, VarType(CellUnder(ws, "Date of Publication", 2)), "text date stays text"
    AssertEqual "(blank)", CellUnder(ws, "Date of Publication", 3), "(blank)"
    AssertEqual "45700", CellUnder(ws, "Date of Publication", 4), "date number stored as text"
    AssertEqual vbString, VarType(CellUnder(ws, "Date of Publication", 4)), "date number stays text"
    AssertEqual DateSerial(2025, 3, 10), CellUnder(ws, "Date of Publication", 1), "real date"
    AssertEqual "Tan Wei Ling" & vbLf & "Goh Mei Xin", CellUnder(ws, "TTSH Contributor", 1), "line breaks kept"
End Sub

Private Sub Export_puts_a_how_to_find_your_records_note_above_each_table()
    Dim cw As Object, ws As Object, note As String
    Set cw = OpenWorkbook(ExportFixture(2025))
    For Each ws In cw.Worksheets
        note = CStr(ws.Range("A1").Value)
        AssertTrue InStr(note, "Ctrl+F") > 0, ws.Name & ": note mentions Ctrl+F"
        AssertTrue InStr(note, "filter arrow") > 0, ws.Name & ": note mentions the filter arrow"
        AssertTrue InStr(note, "Sheet View") > 0, ws.Name & ": note mentions Sheet View"
        AssertTrue HeaderCell(ws, "Export ID").Row > 1, ws.Name & ": table is below the note"
    Next ws
End Sub

' Education's lists are typed into the Main Database's validation; Project
' Status refers to tbllist_PrjStatus. Rows typed below the export get them too.
Private Sub Export_carries_the_dropdown_lists_over_as_fixed_lists()
    Dim cw As Object, edu As Object
    Set cw = OpenWorkbook(ExportFixture(2025))
    Set edu = cw.Worksheets("Education")
    AssertEqual "", ListOf(edu, "Title of Course/Educational Materials/Educational Events/Outreach", 1), "no list where the Main Database has none"
    If mInLibreOffice Then
        Skip "dropdown lists (LibreOffice reads them in its own form)"
        Exit Sub
    End If
    AssertTrue InStr(ListOf(edu, "Type of Material/Event", 1), "Presentation/Sharing") > 0, "Education type list on the first row"
    AssertTrue InStr(ListOf(edu, "Type of Material/Event", 20), "Webinar") > 0, "Education type list on a new row"
    AssertTrue InStr(ListOf(edu, "Area(s) of Focus (Professional vs Patients/Care Givers/General Public)", 1), "Patients/Care Givers/General Public") > 0, "audience list"
    AssertTrue InStr(ListOf(edu, "Local/Overseas", 1), "Overseas") > 0, "Local/Overseas list"
    AssertEqual "Presented,Published,Pending,Applied,Under Review,Awarded,In Progress,Completed,Incomplete,Dropped," & _
            "Applied (Oversight),Under Review (Oversight),Awarded (Oversight),In Progress (Oversight),Completed (Oversight)," & _
        "Incompleted (Oversight),Dropped (Oversight)", ListOf(cw.Worksheets("Projects"), "Status", 1), "Project Status list"
End Sub

Private Sub Export_puts_the_rows_in_an_Excel_table()
    Dim cw As Object, ws As Object
    If mInLibreOffice Then
        Skip "tables (LibreOffice has no ListObjects)"
        Exit Sub
    End If
    Set cw = OpenWorkbook(ExportFixture(2025))
    For Each ws In cw.Worksheets
        AssertEqual 1, ws.ListObjects.Count, ws.Name & ": one table"
        AssertEqual HeaderCell(ws, "Export ID").Row, ws.ListObjects(1).HeaderRowRange.Row, ws.Name & ": table header row"
        AssertEqual ws.Cells(ws.Rows.Count, HeaderCell(ws, "Export ID").Column).End(-4162).Row, _
            ws.ListObjects(1).Range.Row + ws.ListObjects(1).Range.Rows.Count - 1, ws.Name & ": table's last row"
    Next ws
End Sub

' Lookup rows read "tab|No.|title|Key 1|Key 2". The keys per tab: Publications
' DOI and title; Presentations title; Media title and outlet; Talent
' Development name and award; Education title and Year; Projects Grant Ref No
' and Project Name.
Private Sub Export_writes_an_ID_lookup_with_each_rows_tab_No_title_and_match_keys()
    Dim lk As Object
    ExportFixture 2025
    Set lk = OpenWorkbook(AdminFolder() & "IREX_Rehab_Data_Collection_CY2025 " & ChrW(8211) & " ID lookup.xlsx")
    Dim ws As Object
    Set ws = lk.Worksheets(1)
    AssertEqual "export id|tab|no.|title|key 1|key 2", Squash(Join(Array(ws.Range("A1").Value, ws.Range("B1").Value, _
        ws.Range("C1").Value, ws.Range("D1").Value, ws.Range("E1").Value, ws.Range("F1").Value), "|")), "headers"
    AssertEqual 27, ws.Cells(ws.Rows.Count, 1).End(-4162).Row - 1, "one row per exported record"
    AssertEqual "Publications|1|Robotic gait training after stroke: a pilot study|10.1000/test.0001|Robotic gait training after stroke: a pilot study", _
        LookupRow(ws, "PUB-0001"), "PUB-0001"
    AssertEqual "Publications|3|Spasticity outcomes after botulinum toxin||Spasticity outcomes after botulinum toxin", LookupRow(ws, "PUB-0002"), "PUB-0002"
    AssertEqual "Poster & Oral Presentation|4|Virtual reality balance training for older adults|Virtual reality balance training for older adults|", _
        LookupRow(ws, "PRES-0003"), "PRES-0003"
    AssertEqual "Media|2|New stroke rehab centre opens|New stroke rehab centre opens|The Test Times Print", LookupRow(ws, "MED-0002"), "MED-0002"
    AssertEqual "Talent Development||Test Fellowship Award|Tan Wei Ling|Test Fellowship Award", LookupRow(ws, "TD-0001"), "TD-0001"
    AssertEqual "Education||Spasticity management workshop|Spasticity management workshop|2025", LookupRow(ws, "EDU-0001"), "EDU-0001"
    AssertEqual "Education||Pressure injury prevention webinar|Pressure injury prevention webinar|", LookupRow(ws, "EDU-0002"), "EDU-0002"
    AssertEqual "Projects||Exoskeleton feasibility||Exoskeleton feasibility", LookupRow(ws, "PRJ-0004"), "PRJ-0004"
    AssertEqual "Projects||Robot-assisted gait study|GRT-001|Robot-assisted gait study", LookupRow(ws, "PRJ-0001"), "PRJ-0001"
End Sub

Private Sub Export_report_counts_the_rows_exported_per_tab()
    Dim report As String
    ExportFixture 2025
    report = ReportText()
    AssertContains report, "Exported|Publications|||5", "Publications"
    AssertContains report, "Exported|Poster & Oral Presentation|||3", "Presentations"
    AssertContains report, "Exported|Media|||4", "Media"
    AssertContains report, "Exported|Talent Development|||3", "Talent Development"
    AssertContains report, "Exported|Education|||6", "Education"
    AssertContains report, "Exported|Projects|||6", "Projects"
End Sub

' Lines read "Unreadable date|tab|Main Database row|column|value".
Private Sub Export_report_lists_rows_whose_date_or_Year_cant_be_read()
    Dim report As String
    ExportFixture 2025
    report = ReportText()
    AssertContains report, "Unreadable date|Publications|7|Date of Publication|(blank)", "Publications"
    AssertContains report, "Unreadable date|Media|6|Media Publishing Date|Nov 26(KIV)", "Media"
    AssertContains report, "Unreadable date|Talent Development|5|Year|", "Talent Development"
    AssertContains report, "Unreadable date|Education|7|Date/Year/Period|25 & 26 Feb 25", "Education"
    AssertContains report, "Unreadable date|Education|13|Date/Year/Period|PARTNER PROGRAMME", "Education label row"
    AssertEqual 5, CountOf(report, "Unreadable date|"), "nothing else listed"
End Sub

' Every Main Database row with an empty No., exported or not.
Private Sub Export_report_lists_Main_Database_rows_with_no_No()
    Dim report As String
    ExportFixture 2025
    report = ReportText()
    AssertContains report, "No No.|Poster & Oral Presentation|6|No.|Return to work after stroke", "Presentations"
    AssertEqual 1, CountOf(report, "No No.|"), "nothing else listed"
End Sub

Private Sub Export_replaces_an_existing_Collection_Workbook_only_when_told_to()
    Dim path As String, staffCopy As String
    path = ExportFixture(2025)
    staffCopy = mFixtures & "check" & Application.PathSeparator & "IREX_Rehab_Data_Collection_CY2025.xlsx"
    FileCopy staffCopy, path
    AssertEqual "", RunExportStaffCopy(2025, OpenMainDatabase(), SharedFolder(), AdminFolder(), False), "returns nothing when not replacing"
    AssertEqual FileLen(staffCopy), FileLen(path), "the existing file is kept"
    If mInLibreOffice Then
        Skip "replacing (LibreOffice keeps the first export open, which blocks saving over it)"
        Exit Sub
    End If
    AssertEqual path, RunExportStaffCopy(2025, OpenMainDatabase(), SharedFolder(), AdminFolder(), True), "returns the path when replacing"
    AssertTrue FileLen(staffCopy) <> FileLen(path), "the existing file is replaced"
End Sub

Private Sub Export_leaves_the_Main_Database_as_it_was()
    Dim before As String
    before = Fingerprint(OpenMainDatabase())
    ExportFixture 2025
    AssertEqual before, Fingerprint(OpenMainDatabase()), "sheets, ranges and values"
End Sub

' ---------------------------------------------------------------------------
' Helpers: running the macros
' ---------------------------------------------------------------------------

' Runs Export Staff Copy on the fixture Main Database into a fresh folder and
' returns the Collection Workbook's path.
Private Function ExportFixture(ByVal year As Long) As String
    Dim folder As String
    folder = FreshFolder("export")
    EnsureFolder folder & "shared"
    EnsureFolder folder & "admin"
    ExportFixture = RunExportStaffCopy(year, OpenMainDatabase(), SharedFolder(), AdminFolder(), True)
End Function

Private Function SharedFolder() As String
    SharedFolder = mOut & "export" & Application.PathSeparator & "shared" & Application.PathSeparator
End Function

Private Function AdminFolder() As String
    AdminFolder = mOut & "export" & Application.PathSeparator & "admin" & Application.PathSeparator
End Function

Private Function OpenMainDatabase() As Object
    Set OpenMainDatabase = OpenWorkbook(mFixtures & "MainDatabase.xlsx")
End Function

' ---------------------------------------------------------------------------
' Helpers: reading workbooks
' ---------------------------------------------------------------------------

Private Function OpenWorkbook(ByVal path As String) As Object
    Dim wb As Object
    For Each wb In Workbooks
        If wb.FullName = path Then
            Set OpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set OpenWorkbook = Workbooks.Open(path)
End Function

Private Function SheetNames(ByVal wb As Object) As String
    Dim ws As Object, s As String
    For Each ws In wb.Worksheets
        s = s & "|" & ws.Name
    Next ws
    SheetNames = Mid(s, 2)
End Function

' The header cell matching header (ignoring case, spaces and line breaks) in the
' first 10 rows, or Nothing.
Private Function HeaderCell(ByVal ws As Object, ByVal header As String) As Object
    Dim r As Long, c As Long, lastCol As Long
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For r = 1 To 10
        For c = 1 To lastCol
            If Squash(ws.Cells(r, c).Value) = Squash(header) Then
                Set HeaderCell = ws.Cells(r, c)
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function Squash(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Replace(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squash = LCase(Trim(s))
End Function

' The values under a column's header, "|"-separated, one per row down to the
' column's last filled cell.
Private Function ColumnText(ByVal ws As Object, ByVal header As String) As String
    Dim h As Object, r As Long, lastRow As Long, s As String
    Set h = HeaderCell(ws, header)
    If h Is Nothing Then
        ColumnText = "(no column " & header & ")"
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, h.Column).End(-4162).Row
    For r = h.Row + 1 To lastRow
        s = s & "|" & Show(ws.Cells(r, h.Column).Value)
    Next r
    ColumnText = Mid(s, 2)
End Function

' The headers in the row holding "Export ID", squashed and "|"-separated.
Private Function HeaderList(ByVal ws As Object) As String
    HeaderList = HeadersWhere(ws, False)
End Function

Private Function HiddenHeaders(ByVal ws As Object) As String
    HiddenHeaders = HeadersWhere(ws, True)
End Function

Private Function HeadersWhere(ByVal ws As Object, ByVal hiddenOnly As Boolean) As String
    Dim h As Object, c As Long, s As String
    Set h = HeaderCell(ws, "Export ID")
    If h Is Nothing Then
        HeadersWhere = "(no Export ID column)"
        Exit Function
    End If
    For c = 1 To ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        If Squash(ws.Cells(h.Row, c).Value) <> "" Then
            If Not hiddenOnly Or ws.Columns(c).Hidden Then s = s & "|" & Squash(ws.Cells(h.Row, c).Value)
        End If
    Next c
    HeadersWhere = Mid(s, 2)
End Function

' The dropdown list on the cell n rows under a column's header, "" if none.
Private Function ListOf(ByVal ws As Object, ByVal header As String, ByVal n As Long) As String
    Dim h As Object
    Set h = HeaderCell(ws, header)
    If h Is Nothing Then
        ListOf = "(no column " & header & ")"
        Exit Function
    End If
    On Error Resume Next
    If ws.Cells(h.Row + n, h.Column).Validation.Type = 3 Then ListOf = ws.Cells(h.Row + n, h.Column).Validation.Formula1
    On Error GoTo 0
End Function

Private Function LookupRow(ByVal ws As Object, ByVal exportId As String) As String
    Dim r As Long, c As Long, s As String
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        If ws.Cells(r, 1).Value = exportId Then
            For c = 2 To 6
                s = s & "|" & Show(ws.Cells(r, c).Value)
            Next c
            LookupRow = Mid(s, 2)
            Exit Function
        End If
    Next r
    LookupRow = "(no lookup row " & exportId & ")"
End Function

' Every row of the export report, one line each, cells "|"-separated.
Private Function ReportText() As String
    Dim f As String, wb As Object
    f = Dir(AdminFolder() & "*export report*.xlsx")
    If f = "" Then
        ReportText = "(no export report)"
        Exit Function
    End If
    Set wb = OpenWorkbook(AdminFolder() & f)
    ReportText = SheetText(wb.Worksheets(1))
End Function

Private Function SheetText(ByVal ws As Object) As String
    Dim r As Long, c As Long, lastRow As Long, lastCol As Long, line As String, s As String
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For r = 1 To lastRow
        line = ""
        For c = 1 To lastCol
            line = line & "|" & Show(ws.Cells(r, c).Value)
        Next c
        s = s & vbLf & Mid(line, 2)
    Next r
    SheetText = s & vbLf
End Function

' Sheet names, used ranges and every value, to tell whether a workbook changed.
Private Function Fingerprint(ByVal wb As Object) As String
    Dim ws As Object, s As String
    For Each ws In wb.Worksheets
        s = s & "#" & ws.Name & "@" & ws.UsedRange.Address & SheetText(ws)
    Next ws
    Fingerprint = s
End Function

Private Function CountOf(ByVal text As String, ByVal part As String) As Long
    CountOf = (Len(text) - Len(Replace(text, part, ""))) / Len(part)
End Function

' The value n rows under a column's header.
Private Function CellUnder(ByVal ws As Object, ByVal header As String, ByVal n As Long) As Variant
    Dim h As Object
    Set h = HeaderCell(ws, header)
    If h Is Nothing Then
        CellUnder = "(no column " & header & ")"
    Else
        CellUnder = ws.Cells(h.Row + n, h.Column).Value
    End If
End Function

Private Sub CloseTestWorkbooks()
    Dim wb As Object, toClose As New Collection, i As Long
    On Error Resume Next
    For Each wb In Workbooks
        If InStr(1, wb.FullName, mOut, vbTextCompare) = 1 Or InStr(1, wb.FullName, mFixtures, vbTextCompare) = 1 Then toClose.Add wb
    Next wb
    For i = 1 To toClose.Count
        toClose(i).Close False
    Next i
End Sub

' ---------------------------------------------------------------------------
' Helpers: files
' ---------------------------------------------------------------------------

Private Sub EnsureFolder(ByVal folder As String)
    If Dir(folder, vbDirectory) = "" Then MkDir folder
End Sub

' Returns mOut\<name>\, emptied of files (one level of subfolders too).
Private Function FreshFolder(ByVal name As String) As String
    Dim folder As String, sub_ As Variant
    folder = mOut & name & Application.PathSeparator
    EnsureFolder folder
    DeleteFilesIn folder
    For Each sub_ In Array("shared", "admin")
        If Dir(folder & sub_, vbDirectory) <> "" Then DeleteFilesIn folder & sub_ & Application.PathSeparator
    Next sub_
    FreshFolder = folder
End Function

Private Sub DeleteFilesIn(ByVal folder As String)
    Dim f As String, files As New Collection, i As Long
    f = Dir(folder & "*.*")
    Do While f <> ""
        files.Add folder & f
        f = Dir()
    Loop
    For i = 1 To files.Count
        Kill files(i)
    Next i
End Sub

' ---------------------------------------------------------------------------
' Assertions
' ---------------------------------------------------------------------------

Private Sub Fail(ByVal message As String)
    mFailures = mFailures & vbLf & message
End Sub

Private Sub Skip(ByVal what As String)
    mSkips = mSkips & "; " & what
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal what As String)
    If Not SameValue(expected, actual) Then
        Fail what & ": expected [" & Show(expected) & "] but got [" & Show(actual) & "]"
    End If
End Sub

' Passes when one whole line of text starts with line (text from SheetText).
Private Sub AssertContains(ByVal text As String, ByVal line As String, ByVal what As String)
    If InStr(text, vbLf & line & "|") = 0 And InStr(text, vbLf & line & vbLf) = 0 Then
        Fail what & ": no line [" & line & "] in" & Replace(text, vbLf, vbLf & "        ")
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal what As String)
    If Not condition Then Fail what
End Sub

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then
        SameValue = (Show(a) = Show(b))
    ElseIf VarType(a) = vbDate Or VarType(b) = vbDate Then
        On Error Resume Next
        SameValue = (CDbl(CDate(a)) = CDbl(CDate(b)))
    ElseIf VarType(a) = vbString Or VarType(b) = vbString Then
        SameValue = (CStr(a) = CStr(b))
    Else
        SameValue = (a = b)
    End If
End Function

Private Function Show(ByVal v As Variant) As String
    If IsError(v) Then
        Show = "#error"
    ElseIf IsEmpty(v) Then
        Show = ""
    ElseIf VarType(v) = vbDate Then
        Show = Format(v, "yyyy-mm-dd")
    Else
        Show = CStr(v)
    End If
End Function
