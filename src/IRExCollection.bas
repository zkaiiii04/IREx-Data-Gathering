Attribute VB_Name = "IRExCollection"
' IREx Collection Workbook macros, run from the Main Database.
'
'   ExportStaffCopy           writes the Collection Workbook for one calendar year,
'                             with its ID lookup and an export report
'   CheckCollectionWorkbook   compares a Collection Workbook with the Main Database
'                             and saves a highlighted copy
'
' Neither macro changes the Main Database. Spec: GitHub issue #1.
Option Explicit

Private Const COLLECTION_PREFIX As String = "IREX_Rehab_Data_Collection_CY"
Private Const XL_WORKBOOK As Long = 51          ' xlOpenXMLWorkbook (.xlsx)
Private Const XL_SINGLE_SHEET As Long = -4167   ' xlWBATWorksheet
Private Const XL_UP As Long = -4162
Private Const HEADER_ROW As Long = 3            ' Collection Workbook header row
Private Const MAX_HEADER_SCAN As Long = 30
Private Const MAX_COLUMN_WIDTH As Double = 60
Private Const EXPORT_ID_HEADER As String = "Export ID"
Private Const ROOM_FOR_NEW_ROWS As Long = 500   ' blank rows under the export that get the dropdowns
Private Const XL_VALIDATE_LIST As Long = 3

' How a tab picks the records of a year.
Private Const RULE_DATE As Long = 1             ' its date falls in the year
Private Const RULE_YEAR As Long = 2             ' its Year is the year
Private Const RULE_YEAR_THEN_DATE As Long = 3   ' its Year, or its date when Year is empty
Private Const RULE_PROJECT As Long = 4          ' its LOA overlaps the year, or Applied with no LOA

' Selection outcomes for one row.
Private Const PICK_OUT As Long = 0
Private Const PICK_IN As Long = 1
Private Const PICK_UNREADABLE As Long = 2       ' included, and listed in the export report

' One activity tab. Header patterns are normalised header text (see Norm);
' a trailing * matches any header that starts with the rest.
Private Type TabSpec
    Title As String             ' tab in the Collection Workbook
    Prefix As String            ' Export ID prefix
    SheetName As String         ' Main Database sheet
    TableName As String         ' Main Database table, "" if the sheet has none
    Anchor As String            ' a header that locates the header row when there's no table
    Only As String              ' "|"-separated headers to export in this order; "" = every column
    Drop As String              ' "|"-separated headers never exported
    Rule As Long
    DateHeader As String
    YearHeader As String
    NameColumn As String        ' where Rehab Staff look for their name (for the note)
    NoHeader As String          ' "" if the tab has no No. column
    TitleHeader As String
    Key1Header As String        ' match key fields, kept in the ID lookup
    Key2Header As String
End Type

' A block of records read from a sheet: one header row and the filled rows under it.
Private Type SheetRows
    Found As Boolean
    HeaderRow As Long
    ColCount As Long
    Cols() As Long              ' sheet column of each header
    Headers() As String         ' header text as on the sheet
    Keys() As String            ' normalised header text
    RowCount As Long
    RowNums() As Long           ' sheet row of each record
    Values() As Variant         ' (record, column)
End Type

Private mTabs(1 To 6) As TabSpec
'Rows for the ID lookup and the export report. (Arrays, not Collections: in
' LibreOffice a module-level Collection returns its array items as objects.)
Private mLookup() As Variant
Private mLookupCount As Long
Private mReport() As Variant
Private mReportCount As Long

' ===========================================================================
' Export Staff Copy
' ===========================================================================

' Writes the Collection Workbook for year into collectionFolder and returns its
' path. The ID lookup and the export report go into adminFolder.
' Writes the Collection Workbook for calendarYear into collectionFolder and
' returns its path. The ID lookup and the export report go into adminFolder.
' If the Collection Workbook already exists, it's replaced only when
' replaceExisting is True; otherwise nothing is written and "" is returned.
Public Function RunExportStaffCopy(ByVal calendarYear As Long, ByVal mainDatabase As Object, ByVal collectionFolder As String, ByVal adminFolder As String, Optional ByVal replaceExisting As Boolean = False) As String
    Dim path As String
    path = collectionFolder & CollectionName(calendarYear) & ".xlsx"
    If FileExists(path) And Not replaceExisting Then Exit Function

    LoadTabs
    mLookupCount = 0
    mReportCount = 0
    Dim cw As Object, i As Long, ws As Object
    Set cw = Workbooks.Add(XL_SINGLE_SHEET)
    For i = 1 To 6
        If i = 1 Then
            Set ws = cw.Worksheets(1)
        Else
            Set ws = cw.Worksheets.Add(After:=cw.Worksheets(cw.Worksheets.Count))
        End If
        ws.Name = mTabs(i).Title
        ExportTab mainDatabase, mTabs(i), calendarYear, ws
    Next i
    cw.Worksheets(1).Activate
    SaveNew cw, path
    WriteLookup adminFolder & CollectionName(calendarYear) & " " & ChrW(8211) & " ID lookup.xlsx"
    WriteExportReport adminFolder & CollectionName(calendarYear) & " " & ChrW(8211) & " export report " & Format(Now, "yyyy-mm-dd") & ".xlsx", calendarYear
    RunExportStaffCopy = path
End Function

Private Function CollectionName(ByVal calendarYear As Long) As String
    CollectionName = COLLECTION_PREFIX & calendarYear
End Function

Private Sub ExportTab(ByVal mainDatabase As Object, ByRef spec As TabSpec, ByVal calendarYear As Long, ByVal ws As Object)
    Dim src As SheetRows, srcSheet As Object
    ws.Cells(1, 1).Value = FindYourRecordsNote(spec)
    src = ReadMainSheet(mainDatabase, spec)
    If Not src.Found Then
        ReportLine "Sheet or table not found", spec.Title, "", spec.SheetName, spec.TableName
        Exit Sub
    End If
    Set srcSheet = mainDatabase.Worksheets(spec.SheetName)

    Dim cols() As Long, n As Long, c As Long, r As Long, outRow As Long, idCol As Long
    n = ExportColumns(spec, src, cols)
    idCol = n + 1
    For c = 1 To n
        ws.Cells(HEADER_ROW, c).Value = src.Headers(cols(c))
        ws.Columns(c).ColumnWidth = Application.WorksheetFunction.Min(srcSheet.Columns(src.Cols(cols(c))).ColumnWidth, MAX_COLUMN_WIDTH)
        ws.Columns(c).Hidden = srcSheet.Columns(src.Cols(cols(c))).Hidden
    Next c
    ws.Cells(HEADER_ROW, idCol).Value = EXPORT_ID_HEADER
    ws.Columns(idCol).Hidden = True

    Dim picked As Long, why As Long, id As String
    outRow = HEADER_ROW
    For r = 1 To src.RowCount
        picked = Pick(spec, src, r, calendarYear, why)
        If picked <> PICK_OUT Then
            outRow = outRow + 1
            For c = 1 To n
                CopyNumberFormat srcSheet.Cells(src.RowNums(r), src.Cols(cols(c))), ws.Cells(outRow, c)
                WriteValue ws.Cells(outRow, c), src.Values(r, cols(c))
            Next c
            id = spec.Prefix & "-" & Format(outRow - HEADER_ROW, "0000")
            ws.Cells(outRow, idCol).Value = id
            LookupLine id, spec.Title, ColumnValue(src, r, spec.NoHeader), ColumnValue(src, r, spec.TitleHeader), _
                       ColumnValue(src, r, spec.Key1Header), ColumnValue(src, r, spec.Key2Header)
            If picked = PICK_UNREADABLE Then
                ReportLine "Unreadable date", spec.Title, src.RowNums(r), Tidy(src.Headers(why)), CellText(src.Values(r, why))
            End If
        End If
    Next r
    ReportLine "Exported", spec.Title, "", "", outRow - HEADER_ROW
    If spec.NoHeader <> "" Then
        For r = 1 To src.RowCount
            If IsBlank(ColumnValue(src, r, spec.NoHeader)) Then
                ReportLine "No No.", spec.Title, src.RowNums(r), "No.", CellText(ColumnValue(src, r, spec.TitleHeader))
            End If
        Next r
    End If

    Dim list As String
    For c = 1 To n
        If src.RowCount > 0 Then
            list = DropdownList(srcSheet.Cells(src.RowNums(1), src.Cols(cols(c))))
            If list <> "" Then AddDropdown ws.Range(ws.Cells(HEADER_ROW + 1, c), ws.Cells(outRow + ROOM_FOR_NEW_ROWS, c)), list
        End If
    Next c
    MakeTable ws, spec, Application.WorksheetFunction.Max(outRow, HEADER_ROW + 1), idCol
End Sub

' The cell's dropdown as a fixed list: the list typed into its validation, or
' the values of the range its validation refers to (such as tbllist_PrjStatus).
Private Function DropdownList(ByVal cell As Object) As String
    Dim kind As Long, f As String, target As Object, c As Object, s As String
    On Error Resume Next
    kind = cell.Validation.Type
    If Err.Number <> 0 Or kind <> XL_VALIDATE_LIST Then Exit Function
    f = cell.Validation.Formula1
    On Error GoTo 0
    If Left(f, 1) <> "=" Then
        DropdownList = f
        Exit Function
    End If
    On Error Resume Next
    Set target = cell.Worksheet.Evaluate(Mid(f, 2))
    If target Is Nothing Then Set target = cell.Worksheet.Parent.Names(Mid(f, 2)).RefersToRange
    On Error GoTo 0
    If target Is Nothing Then Exit Function
    For Each c In target.Cells
        If Not IsBlank(c.Value) Then s = s & "," & CellText(c.Value)
    Next c
    DropdownList = Mid(s, 2)
End Function

Private Sub AddDropdown(ByVal target As Object, ByVal list As String)
    On Error Resume Next
    target.Validation.Delete
    target.Validation.Add Type:=XL_VALIDATE_LIST, AlertStyle:=1, Operator:=1, Formula1:=list
    Err.Clear
End Sub

Private Sub MakeTable(ByVal ws As Object, ByRef spec As TabSpec, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim t As Object
    On Error Resume Next        ' LibreOffice has no ListObjects; the sheet works without one
    Set t = ws.ListObjects.Add(1, ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol)), , 1)
    If Not t Is Nothing Then t.Name = "Collection_" & spec.Prefix
    Err.Clear
End Sub

Private Function FindYourRecordsNote(ByRef spec As TabSpec) As String
    FindYourRecordsNote = "How to find your records: if Excel offers Sheet View (View > Sheet View > New), choose it first, " & _
        "so that your sorting and filtering don't change what colleagues see. Then press Ctrl+F and search for your name, " & _
        "or use the filter arrow on the " & spec.NameColumn & " column. " & _
        "To report something new, type it in a new row below the existing ones. " & _
        "If something has changed, such as a status, update your existing row."
End Function

Private Sub CopyNumberFormat(ByVal fromCell As Object, ByVal toCell As Object)
    Dim f As String
    f = fromCell.NumberFormat
    If f <> "" And LCase(f) <> "general" Then toCell.NumberFormat = f
End Sub

' The source columns a tab exports, in order. Returns how many.
Private Function ExportColumns(ByRef spec As TabSpec, ByRef src As SheetRows, ByRef cols() As Long) As Long
    Dim n As Long, c As Long, i As Long, wanted As Variant
    ReDim cols(1 To src.ColCount)
    If spec.Only <> "" Then
        wanted = Split(spec.Only, "|")
        For i = 0 To UBound(wanted)
            c = FindColumn(src, CStr(wanted(i)))
            If c > 0 Then
                n = n + 1
                cols(n) = c
            End If
        Next i
    Else
        For c = 1 To src.ColCount
            If Not MatchesAny(src.Keys(c), spec.Drop) Then
                n = n + 1
                cols(n) = c
            End If
        Next c
    End If
    ExportColumns = n
End Function

' Whether row r of a tab belongs to the year. When the answer is
' PICK_UNREADABLE, why is set to the column that couldn't be read.
Private Function Pick(ByRef spec As TabSpec, ByRef src As SheetRows, ByVal r As Long, ByVal calendarYear As Long, ByRef why As Long) As Long
    Dim dateCol As Long, yearCol As Long, startCol As Long, endCol As Long
    dateCol = FindColumn(src, spec.DateHeader)
    yearCol = FindColumn(src, spec.YearHeader)
    Select Case spec.Rule
        Case RULE_DATE
            why = dateCol
            Pick = PickByDate(CellAt(src, r, dateCol), calendarYear)
        Case RULE_YEAR
            why = yearCol
            Pick = PickByYear(CellAt(src, r, yearCol), calendarYear)
        Case RULE_YEAR_THEN_DATE
            If IsBlank(CellAt(src, r, yearCol)) Then
                why = dateCol
                Pick = PickByDate(CellAt(src, r, dateCol), calendarYear)
            Else
                why = yearCol
                Pick = PickByYear(CellAt(src, r, yearCol), calendarYear)
            End If
        Case RULE_PROJECT
            startCol = FindColumn(src, "loa start date")
            endCol = FindColumn(src, "loa end date")
            Pick = PickProject(CellAt(src, r, startCol), CellAt(src, r, endCol), ColumnValue(src, r, "status"), calendarYear)
            Dim d As Date
            If ReadDate(CellAt(src, r, startCol), d) Or IsBlank(CellAt(src, r, startCol)) Then why = endCol Else why = startCol
    End Select
    If why = 0 Then why = 1
End Function

Private Function PickByYear(ByVal v As Variant, ByVal calendarYear As Long) As Long
    Dim y As Long
    If Not ReadYear(v, y) Then
        PickByYear = PICK_UNREADABLE
    ElseIf y = calendarYear Then
        PickByYear = PICK_IN
    Else
        PickByYear = PICK_OUT
    End If
End Function

Private Function PickProject(ByVal loaStart As Variant, ByVal loaEnd As Variant, ByVal status As Variant, ByVal calendarYear As Long) As Long
    Dim startDate As Date, endDate As Date
    If IsBlank(loaStart) Then
        If Left(Norm(status), 7) = "applied" Then PickProject = PICK_IN Else PickProject = PICK_OUT
    ElseIf Not ReadDate(loaStart, startDate) Then
        PickProject = PICK_UNREADABLE
    ElseIf startDate > DateSerial(calendarYear, 12, 31) Then
        PickProject = PICK_OUT
    ElseIf IsBlank(loaEnd) Then
        PickProject = PICK_IN
    ElseIf Not ReadDate(loaEnd, endDate) Then
        PickProject = PICK_UNREADABLE
    ElseIf endDate >= DateSerial(calendarYear, 1, 1) Then
        PickProject = PICK_IN
    Else
        PickProject = PICK_OUT
    End If
End Function

Private Function PickByDate(ByVal v As Variant, ByVal calendarYear As Long) As Long
    Dim d As Date
    If Not ReadDate(v, d) Then
        PickByDate = PICK_UNREADABLE
    ElseIf Year(d) = calendarYear Then
        PickByDate = PICK_IN
    Else
        PickByDate = PICK_OUT
    End If
End Function

' ===========================================================================
' Tabs
' ===========================================================================

Private Sub LoadTabs()
    With mTabs(1)
        .Title = "Publications": .NameColumn = "TTSH Contributor": .Prefix = "PUB"
        .SheetName = "DB_Publications": .TableName = "tbl_Publication": .Anchor = "title of article"
        .Drop = "journal category & jif percentile listing*|jif percentile*"
        .Rule = RULE_DATE: .DateHeader = "date of publication"
        .NoHeader = "no.": .TitleHeader = "title of article": .Key1Header = "doi": .Key2Header = "title of article"
    End With
    With mTabs(2)
        .Title = "Poster & Oral Presentation": .NameColumn = "Full Name": .Prefix = "PRES"
        .SheetName = "DB_Poster& Oral Presentation": .TableName = "tbl_PosterOralPresentation"
        .Anchor = "title of presentation/topic"
        .Rule = RULE_DATE: .DateHeader = "date of presentation"
        .NoHeader = "no.": .TitleHeader = "title of presentation/topic": .Key1Header = "title of presentation/topic": .Key2Header = ""
    End With
    With mTabs(3)
        .Title = "Media": .NameColumn = "Full name": .Prefix = "MED"
        .SheetName = "DB_Media": .TableName = "": .Anchor = "platform of media"
        .Rule = RULE_DATE: .DateHeader = "media publishing date"
        .NoHeader = "no.": .TitleHeader = "title of media published*": .Key1Header = "title of media published*": .Key2Header = "name of media published source*"
    End With
    With mTabs(4)
        .Title = "Talent Development": .NameColumn = "Name": .Prefix = "TD"
        .SheetName = "DB_Talent Develop": .TableName = "tbl_TalentDevelop": .Anchor = "award description"
        .Rule = RULE_YEAR: .YearHeader = "year"
        .NoHeader = "": .TitleHeader = "award description": .Key1Header = "name": .Key2Header = "award description"
    End With
    With mTabs(5)
        .Title = "Education": .NameColumn = "Trainer(s)/Speaker(s)": .Prefix = "EDU"
        .SheetName = "DB_Education": .TableName = "tbl_Education": .Anchor = "title of course*"
        .Rule = RULE_YEAR_THEN_DATE: .YearHeader = "year": .DateHeader = "date/year/period"
        .NoHeader = "": .TitleHeader = "title of course*": .Key1Header = "title of course*": .Key2Header = "year"
    End With
    With mTabs(6)
        .Title = "Projects": .NameColumn = "PI or Co-I/Collaborator": .Prefix = "PRJ"
        .SheetName = "DB_PrjGeneralInfo": .TableName = "tbl_PrjGeneralInfo": .Anchor = "project name"
        .Only = "grant ref no|project name|category|grant / collaborator partners|grant type|pi|" & _
                "co-i/collaborator|loa start date|loa end date|grant amount|grant amt (awarded to ttsh)|status"
        .Rule = RULE_PROJECT
        .NoHeader = "": .TitleHeader = "project name": .Key1Header = "grant ref no": .Key2Header = "project name"
    End With
End Sub

' ===========================================================================
' Reading sheets
' ===========================================================================

' Reads a tab's records from the Main Database: the header row of its table (or,
' with no table, the row holding its anchor header), then every filled row
' below it within those columns, inside or below the table, gaps allowed.
Private Function ReadMainSheet(ByVal mainDatabase As Object, ByRef spec As TabSpec) As SheetRows
    Dim ws As Object, headerRow As Long, firstCol As Long, lastCol As Long
    On Error Resume Next
    Set ws = mainDatabase.Worksheets(spec.SheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If Not TableHeader(ws, spec.TableName, headerRow, firstCol, lastCol) Then
        If Not ScanHeader(ws, spec.Anchor, headerRow, firstCol, lastCol) Then Exit Function
    End If
    ReadMainSheet = ReadBlock(ws, headerRow, firstCol, lastCol)
End Function

Private Function TableHeader(ByVal ws As Object, ByVal tableName As String, ByRef headerRow As Long, ByRef firstCol As Long, ByRef lastCol As Long) As Boolean
    Dim header As Object
    If tableName = "" Then Exit Function
    On Error Resume Next
    Set header = ws.ListObjects(tableName).HeaderRowRange
    On Error GoTo 0
    If header Is Nothing Then Exit Function
    headerRow = header.Row
    firstCol = header.Column
    lastCol = header.Column + header.Columns.Count - 1
    TableHeader = True
End Function

' Finds the anchor header near the top of the sheet; the header row runs left
' and right of it to the first empty header cell.
Private Function ScanHeader(ByVal ws As Object, ByVal anchor As String, ByRef headerRow As Long, ByRef firstCol As Long, ByRef lastCol As Long) As Boolean
    Dim r As Long, c As Long, maxCol As Long
    maxCol = LastUsedColumn(ws)
    For r = 1 To MAX_HEADER_SCAN
        For c = 1 To maxCol
            If HeaderMatches(Norm(ws.Cells(r, c).Value), anchor) Then
                headerRow = r
                firstCol = c
                Do While firstCol > 1
                    If IsBlank(ws.Cells(r, firstCol - 1).Value) Then Exit Do
                    firstCol = firstCol - 1
                Loop
                lastCol = c
                Do While lastCol < maxCol
                    If IsBlank(ws.Cells(r, lastCol + 1).Value) Then Exit Do
                    lastCol = lastCol + 1
                Loop
                ScanHeader = True
                Exit Function
            End If
        Next c
    Next r
End Function

' (LibreOffice can't ReDim a member of a Type, so arrays are built locally.)
Private Function ReadBlock(ByVal ws As Object, ByVal headerRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As SheetRows
    Dim b As SheetRows, c As Long, r As Long, lastRow As Long, n As Long, i As Long, k As Long
    Dim cols() As Long, headers() As String, keys() As String
    k = lastCol - firstCol + 1
    ReDim cols(1 To k)
    ReDim headers(1 To k)
    ReDim keys(1 To k)
    For i = 1 To k
        c = firstCol + i - 1
        cols(i) = c
        headers(i) = CellText(ws.Cells(headerRow, c).Value)
        keys(i) = Norm(headers(i))
        r = ws.Cells(ws.Rows.Count, c).End(XL_UP).Row
        If r > lastRow Then lastRow = r
    Next i
    b.Found = True
    b.HeaderRow = headerRow
    b.ColCount = k
    b.Cols = cols
    b.Headers = headers
    b.Keys = keys

    Dim filled() As Boolean
    If lastRow <= headerRow Then
        ReadBlock = b
        Exit Function
    End If
    ReDim filled(headerRow + 1 To lastRow)
    For r = headerRow + 1 To lastRow
        For i = 1 To k
            If Not IsBlank(ws.Cells(r, cols(i)).Value) Then
                filled(r) = True
                n = n + 1
                Exit For
            End If
        Next i
    Next r

    b.RowCount = n
    If n > 0 Then
        Dim rowNums() As Long, values() As Variant
        ReDim rowNums(1 To n)
        ReDim values(1 To n, 1 To k)
        n = 0
        For r = headerRow + 1 To lastRow
            If filled(r) Then
                n = n + 1
                rowNums(n) = r
                For i = 1 To k
                    values(n, i) = ws.Cells(r, cols(i)).Value
                Next i
            End If
        Next r
        b.RowNums = rowNums
        b.Values = values
    End If
    ReadBlock = b
End Function

Private Function LastUsedColumn(ByVal ws As Object) As Long
    LastUsedColumn = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
End Function

' The column of a block whose header matches pattern, or 0.
Private Function FindColumn(ByRef b As SheetRows, ByVal pattern As String) As Long
    Dim c As Long
    For c = 1 To b.ColCount
        If HeaderMatches(b.Keys(c), pattern) Then
            FindColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function ColumnValue(ByRef b As SheetRows, ByVal r As Long, ByVal pattern As String) As Variant
    ColumnValue = Empty
    If pattern <> "" Then ColumnValue = CellAt(b, r, FindColumn(b, pattern))
End Function

Private Function CellAt(ByRef b As SheetRows, ByVal r As Long, ByVal c As Long) As Variant
    CellAt = Empty
    If c > 0 Then CellAt = b.Values(r, c)
End Function

Private Function HeaderMatches(ByVal key As String, ByVal pattern As String) As Boolean
    If Right(pattern, 1) = "*" Then
        HeaderMatches = (Left(key, Len(pattern) - 1) = Left(pattern, Len(pattern) - 1))
    Else
        HeaderMatches = (key = pattern)
    End If
End Function

Private Function MatchesAny(ByVal key As String, ByVal patterns As String) As Boolean
    Dim p As Variant
    If patterns = "" Then Exit Function
    For Each p In Split(patterns, "|")
        If HeaderMatches(key, CStr(p)) Then
            MatchesAny = True
            Exit Function
        End If
    Next p
End Function

' ===========================================================================
' Values
' ===========================================================================

' Text with line breaks, tabs and runs of spaces collapsed to one space,
' trimmed and in lower case.
Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    s = CellText(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = LCase(Trim(s))
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#error"
    ElseIf IsEmpty(v) Or IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

' Empty, only spaces or line breaks, or "(blank)".
Private Function IsBlank(ByVal v As Variant) As Boolean
    Dim s As String
    s = Norm(v)
    IsBlank = (s = "" Or s = "(blank)")
End Function

' Reads a real date, an Excel date number (also when stored as text), or a
' text date such as "20 Aug 25", "20 August 2025", "2025-08-20" or "20/8/2025".
Private Function ReadDate(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim n As Double
    If VarType(v) = vbDate Then
        d = v
        ReadDate = True
    ElseIf IsBlank(v) Or IsError(v) Then
        ReadDate = False
    ElseIf ReadNumber(v, n) Then
        If n >= 10000 And n <= 80000 Then           ' 1927 to 2119
            d = DateSerial(1899, 12, 30) + Int(n)
            ReadDate = True
        End If
    Else
        ReadDate = ParseDateText(CStr(v), d)
    End If
End Function

' Reads a whole-number year from 1900 to 2200, as a number or as text.
Private Function ReadYear(ByVal v As Variant, ByRef y As Long) As Boolean
    Dim n As Double
    If VarType(v) = vbDate Then Exit Function
    If Not ReadNumber(v, n) Then Exit Function
    If n <> Int(n) Or n < 1900 Or n > 2200 Then Exit Function
    y = CLng(n)
    ReadYear = True
End Function

Private Function ParseDateText(ByVal s As String, ByRef d As Date) As Boolean
    Dim t As Variant, sep As Variant, a As Long, m As Long, b As Long
    s = Trim(s)
    For Each sep In Array("-", "/", ".", ",")
        s = Replace(s, sep, " ")
    Next sep
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    t = Split(s, " ")
    If UBound(t) <> 2 Then Exit Function
    If IsDigits(t(0)) And MonthNumber(t(1)) > 0 And IsDigits(t(2)) Then           ' 20 Aug 25
        ParseDateText = MakeDate(CLng(t(2)), MonthNumber(t(1)), CLng(t(0)), d)
    ElseIf MonthNumber(t(0)) > 0 And IsDigits(t(1)) And IsDigits(t(2)) Then       ' Aug 20 2025
        ParseDateText = MakeDate(CLng(t(2)), MonthNumber(t(0)), CLng(t(1)), d)
    ElseIf IsDigits(t(0)) And IsDigits(t(1)) And IsDigits(t(2)) Then
        a = CLng(t(0)): m = CLng(t(1)): b = CLng(t(2))
        If Len(t(0)) = 4 Then
            ParseDateText = MakeDate(a, m, b, d)                                   ' 2025-08-20
        Else
            ParseDateText = MakeDate(b, m, a, d)                                   ' 20/08/2025
        End If
    End If
End Function

Private Function MakeDate(ByVal y As Long, ByVal m As Long, ByVal dd As Long, ByRef d As Date) As Boolean
    If y < 100 Then y = 2000 + y
    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    d = DateSerial(y, m, dd)
    MakeDate = (Day(d) = dd)
End Function

Private Function MonthNumber(ByVal token As Variant) As Long
    Dim names As Variant, i As Long, t As String
    t = LCase(CStr(token))
    If Len(t) < 3 Then Exit Function
    names = Array("january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december")
    For i = 0 To 11
        If Left(names(i), Len(t)) = t Or (i = 8 And t = "sept") Then
            MonthNumber = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function IsDigits(ByVal token As Variant) As Boolean
    Dim s As String, i As Long
    s = CStr(token)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then Exit Function
    Next i
    IsDigits = True
End Function

' Reads a number, or text that is a number apart from formatting: spaces,
' thousands commas and a leading $.
Private Function ReadNumber(ByVal v As Variant, ByRef n As Double) As Boolean
    Dim s As String, i As Long, ch As String, digits As Long, dots As Long
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            n = CDbl(v)
            ReadNumber = True
            Exit Function
        Case vbString
        Case Else
            Exit Function
    End Select
    s = Replace(Replace(Trim(CStr(v)), ",", ""), " ", "")
    If Left(s, 2) = "S$" Then s = Mid(s, 3)
    If Left(s, 1) = "$" Then s = Mid(s, 2)
    If Left(s, 1) = "-" Then i = 2 Else i = 1
    For i = i To Len(s)
        ch = Mid(s, i, 1)
        If ch = "." Then
            dots = dots + 1
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Or dots > 1 Then Exit Function
    n = Val(s)
    ReadNumber = True
End Function

' Writes a value as it is. Text that Excel would turn into a number or date
' when typed ("45700", "20 Aug 25", "(5)") is written as text.
Private Sub WriteValue(ByVal cell As Object, ByVal v As Variant)
    If IsEmpty(v) Then
        Exit Sub
    ElseIf IsError(v) Then
        cell.Value = "'" & CellText(v)
    ElseIf VarType(v) = vbString Then
        If TextNeedsQuote(CStr(v)) Then cell.Value = "'" & v Else cell.Value = v
    Else
        cell.Value = v
    End If
End Sub

Private Function TextNeedsQuote(ByVal s As String) As Boolean
    Dim first As String
    If s = "" Then Exit Function
    first = LCase(Left(s, 1))
    If first < "a" Or first > "z" Then
        TextNeedsQuote = True
    ElseIf MonthNumber(Split(Trim(s) & " ", " ")(0)) > 0 Then
        TextNeedsQuote = True
    ElseIf LCase(s) = "true" Or LCase(s) = "false" Then
        TextNeedsQuote = True
    End If
End Function

' ===========================================================================
' Files
' ===========================================================================

Private Sub WriteLookup(ByVal path As String)
    Dim wb As Object, ws As Object, r As Long, c As Long, entry As Variant
    Set wb = Workbooks.Add(XL_SINGLE_SHEET)
    Set ws = wb.Worksheets(1)
    ws.Name = "ID lookup"
    ws.Range("A1:F1").Value = Array(EXPORT_ID_HEADER, "Tab", "No.", "Title", "Key 1", "Key 2")
    ws.Range("A1:F1").Font.Bold = True
    For r = 1 To mLookupCount
        entry = mLookup(r)
        For c = 0 To 5
            WriteValue ws.Cells(r + 1, c + 1), entry(c)
        Next c
    Next r
    SaveNew wb, path
End Sub

Private Sub WriteExportReport(ByVal path As String, ByVal calendarYear As Long)
    Dim wb As Object, ws As Object, r As Long, c As Long, entry As Variant
    Set wb = Workbooks.Add(XL_SINGLE_SHEET)
    Set ws = wb.Worksheets(1)
    ws.Name = "Export report"
    ws.Range("A1").Value = "Export Staff Copy for CY" & calendarYear & ", run " & Format(Now, "d mmm yyyy hh:mm")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:E3").Value = Array("Entry", "Tab", "Main Database row", "Column", "Value")
    ws.Range("A3:E3").Font.Bold = True
    For r = 1 To mReportCount
        entry = mReport(r)
        For c = 0 To 4
            WriteValue ws.Cells(r + 3, c + 1), entry(c)
        Next c
    Next r
    ws.Columns("A:E").ColumnWidth = 24
    SaveNew wb, path
End Sub

Private Sub LookupLine(ByVal id As String, ByVal tabTitle As String, ByVal no As Variant, ByVal title As Variant, ByVal key1 As Variant, ByVal key2 As Variant)
    mLookupCount = mLookupCount + 1
    ReDim Preserve mLookup(1 To mLookupCount)
    mLookup(mLookupCount) = Array(id, tabTitle, no, title, key1, key2)
End Sub

Private Sub ReportLine(ByVal entry As String, ByVal tabTitle As String, ByVal sheetRow As Variant, ByVal column As String, ByVal value As Variant)
    mReportCount = mReportCount + 1
    ReDim Preserve mReport(1 To mReportCount)
    mReport(mReportCount) = Array(entry, tabTitle, sheetRow, column, value)
End Sub

' Header text with line breaks and runs of spaces collapsed, for reports.
Private Function Tidy(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Tidy = Trim(s)
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    Err.Clear
End Function

' Saves a new workbook as .xlsx at path, replacing any file there, and closes it.
Private Sub SaveNew(ByVal wb As Object, ByVal path As String)
    Dim alerts As Boolean
    If FileExists(path) Then Kill path
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=XL_WORKBOOK
    Application.DisplayAlerts = alerts
    wb.Close False
End Sub
