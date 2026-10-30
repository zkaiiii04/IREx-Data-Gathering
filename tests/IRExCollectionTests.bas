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
        "Export_takes_Projects_live_during_the_year_and_Applied_ones_with_no_LOA")
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
        Case Else: Fail "no such test"
    End Select
End Sub

Private Sub RunTest(ByVal name As String)
    mFailures = ""
    mSkips = ""
    Err.Clear
    On Error Resume Next
    Dispatch name
    If Err.Number <> 0 Then Fail "ERROR " & Err.Number & ": " & Err.Description
    On Error GoTo 0
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
