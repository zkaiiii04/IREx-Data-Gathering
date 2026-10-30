Attribute VB_Name = "IRExCollection"
' IREx Collection Workbook macros, run from the Main Database.
'
'   ExportStaffCopy           writes the Collection Workbook for one calendar year,
'                             with its ID lookup and an export report
'   CheckCollectionWorkbook   compares a Collection Workbook with the Main Database
'                             and saves a highlighted copy
'
' Neither macro changes the Main Database. Run them from the Macros dialog
' (Alt+F8) or a Quick Access Toolbar button. Spec: GitHub issue #1.
'
' The first run of each asks for the two folders (the SharePoint folder Rehab
' Staff can edit, and the admin-only folder) and remembers them for this
' Windows user. To choose again, run ResetCollectionFolders from the
' Immediate window (Ctrl+G in the VBA editor).
'
' RunExportStaffCopy and RunCheckCollectionWorkbook do the work without any
' prompts; the tests in tests\IRExCollectionTests.bas call them directly.
Option Explicit

Private Const COLLECTION_PREFIX As String = "IREX_Rehab_Data_Collection_CY"
Private Const SETTINGS_APP As String = "IREx Collection Workbook"
Private Const SETTINGS_SECTION As String = "Folders"
Private Const SHARED_KEY As String = "Collection"
Private Const ADMIN_KEY As String = "Admin"
Private Const PICK_FOLDER As Long = 4           ' msoFileDialogFolderPicker
Private Const PICK_FILE As Long = 3             ' msoFileDialogFilePicker
Private Const XL_WORKBOOK As Long = 51          ' xlOpenXMLWorkbook (.xlsx)
Private Const XL_SINGLE_SHEET As Long = -4167   ' xlWBATWorksheet
Private Const HEADER_ROW As Long = 3            ' Collection Workbook header row
Private Const MAX_HEADER_SCAN As Long = 30
Private Const MAX_COLUMN_WIDTH As Double = 60
Private Const DEFAULT_COLUMN_WIDTH As Double = 14
Private Const MIN_COLUMN_WIDTH As Double = 16    ' wide enough for a date, and for header words not to break
Private Const EXPORT_ID_HEADER As String = "Export ID"
Private Const NO_HEADER As String = "no."       ' the Main Database's row number, never exported
Private Const ROOM_FOR_NEW_ROWS As Long = 500   ' blank rows under the export that get the dropdowns
Private Const XL_VALIDATE_LIST As Long = 3
Private Const MACRO_ERROR As Long = &H80040000   ' vbObjectError, which LibreOffice lacks
Private Const CHECK_NOTE_HEADER As String = "Check note"
Private Const GREEN As Long = 13561798         ' RGB(198, 239, 206): new
Private Const YELLOW As Long = 10284031        ' RGB(255, 235, 156): changed

' Collection Workbook layout.
Private Const XL_CENTER As Long = -4108
Private Const XL_LEFT As Long = -4131
Private Const XL_CONTINUOUS As Long = 1
Private Const WHITE As Long = 16777215
Private Const HEADER_FILL As Long = 7949855     ' RGB(31, 78, 121): dark blue
Private Const NOTE_FILL As Long = 16247773      ' RGB(221, 235, 247): pale blue
Private Const GRID_COLOUR As Long = 12566463    ' RGB(191, 191, 191): light grey
Private Const NOTE_FONT_SIZE As Long = 14
Private Const HEADER_FONT_SIZE As Long = 12
Private Const MIN_ROW_HEIGHT As Double = 24
Private Const ROW_PADDING As Double = 8         ' points added to each row's fitted height
Private Const MAX_ROW_HEIGHT As Double = 409    ' Excel's limit
Private Const TABLE_STYLE As String = "IREx Collection"

' Summary counts, one per tab.
Private Const COUNT_NEW As Long = 1
Private Const COUNT_CHANGED As Long = 2
Private Const COUNT_ALREADY As Long = 3
Private Const COUNT_DELETED As Long = 4
Private Const COUNT_CANT_FIND As Long = 5
Private Const COUNT_EXTRA As Long = 6

' How a tab picks the records of a year.
Private Const RULE_DATE As Long = 1             ' its date falls in the year
Private Const RULE_YEAR As Long = 2             ' its Year is the year
Private Const RULE_YEAR_THEN_DATE As Long = 3   ' its Year, or its date when Year is empty
Private Const RULE_PROJECT As Long = 4          ' its LOA overlaps the year, or Applied with no LOA; never Dropped

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
    Keys As String              ' match keys tried in order, ";"-separated; each a "+"-joined
                                ' list of fields: 0 = No., 1 = Key 1, 2 = Key 2
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
' Rows for the ID lookup and the export report. (Arrays, not Collections: in
' LibreOffice a module-level Collection returns its array items as objects.)
Private mLookup() As Variant
Private mLookupCount As Long
Private mReport() As Variant
Private mReportCount As Long

' The ID lookup, while checking.
Private mIds() As Variant           ' (row, 0 to 5): Export ID, tab, No., title, Key 1, Key 2
Private mIdCount As Long
Private mIdSeen() As Boolean
Private mCounts(1 To 6, 1 To 6) As Long     ' (tab, COUNT_*)

' ===========================================================================
' The macros
' ===========================================================================

Public Sub ExportStaffCopy()
    Const TITLE As String = "Export Staff Copy"
    Dim answer As String, calendarYear As Long, sharedFolder As String, adminFolder As String, path As String
    answer = InputBox("Export the Collection Workbook for which calendar year?", TITLE, Year(Date))
    If answer = "" Then Exit Sub
    If Not ReadYear(answer, calendarYear) Then
        MsgBox "Type a year such as " & Year(Date) & ".", vbExclamation, TITLE
        Exit Sub
    End If
    sharedFolder = RememberedFolder(SHARED_KEY, "Choose the SharePoint folder Rehab Staff can edit (the Collection Workbook goes here)")
    If sharedFolder = "" Then Exit Sub
    adminFolder = RememberedFolder(ADMIN_KEY, "Choose the folder only IREx Admin can open (the ID lookup and reports go here)")
    If adminFolder = "" Then Exit Sub

    path = sharedFolder & CollectionName(calendarYear) & ".xlsx"
    If FileExists(path) Then
        If MsgBox(CollectionName(calendarYear) & ".xlsx is already in " & sharedFolder & vbLf & vbLf & _
                  "Replacing it loses every row Rehab Staff have typed since your last check. " & _
                  "Run Check Collection Workbook first and copy over what you need." & vbLf & vbLf & _
                  "Replace it?", vbYesNo + vbExclamation + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub
    End If

    On Error GoTo failed
    Application.ScreenUpdating = False
    path = RunExportStaffCopy(calendarYear, ThisWorkbook, sharedFolder, adminFolder, True)
    Application.ScreenUpdating = True
    MsgBox "Exported " & mLookupCount & " records to" & vbLf & path & vbLf & vbLf & _
           ReportCount("Unreadable date") & " rows with a date or Year that can't be read, and " & _
           ReportCount("No No.") & " rows with no No., are listed in the export report in" & vbLf & adminFolder, vbInformation, TITLE
    Exit Sub
failed:
    Application.ScreenUpdating = True
    MsgBox "Export Staff Copy stopped: " & Err.Description, vbCritical, TITLE
End Sub

Public Sub CheckCollectionWorkbook()
    Const TITLE As String = "Check Collection Workbook"
    Dim cwPath As String, adminFolder As String, lookupPath As String, copyPath As String
    cwPath = PickFile("Pick the Collection Workbook to check (a synced or downloaded copy)", _
                      GetSetting(SETTINGS_APP, SETTINGS_SECTION, SHARED_KEY, ""), "Collection Workbook")
    If cwPath = "" Then Exit Sub
    adminFolder = RememberedFolder(ADMIN_KEY, "Choose the folder only IREx Admin can open (the ID lookup and reports go here)")
    If adminFolder = "" Then Exit Sub
    lookupPath = PickFile("Pick the ID lookup saved when this Collection Workbook was exported", _
                          adminFolder & LookupNameFor(cwPath), "ID lookup")
    If lookupPath = "" Then Exit Sub

    On Error GoTo failed
    Application.ScreenUpdating = False
    copyPath = RunCheckCollectionWorkbook(cwPath, lookupPath, ThisWorkbook, adminFolder)
    Application.ScreenUpdating = True
    Workbooks.Open copyPath
    MsgBox TotalCount(COUNT_NEW) & " new, " & TotalCount(COUNT_CHANGED) & " changed, " & _
           TotalCount(COUNT_ALREADY) & " already recorded, " & TotalCount(COUNT_DELETED) & " deleted, " & _
           TotalCount(COUNT_CANT_FIND) & " can't be found in the Main Database." & vbLf & vbLf & _
           "The highlighted copy is open, and saved as" & vbLf & copyPath, vbInformation, TITLE
    Exit Sub
failed:
    Application.ScreenUpdating = True
    MsgBox "Check Collection Workbook stopped: " & Err.Description, vbCritical, TITLE
End Sub

' Forgets the two folders, so the next run asks for them again. Run it from
' the Immediate window; the optional argument keeps it out of the Macros dialog.
Public Sub ResetCollectionFolders(Optional ByVal quiet As Boolean = False)
    On Error Resume Next
    DeleteSetting SETTINGS_APP, SETTINGS_SECTION
    On Error GoTo 0
    If Not quiet Then MsgBox "The next run of Export Staff Copy or Check Collection Workbook will ask for the folders again.", vbInformation
End Sub

' A folder chosen once and remembered for this Windows user, with a trailing
' path separator; "" if the picker is cancelled.
Private Function RememberedFolder(ByVal key As String, ByVal prompt As String) As String
    Dim folder As String
    folder = GetSetting(SETTINGS_APP, SETTINGS_SECTION, key, "")
    If folder <> "" Then
        If Dir(folder, vbDirectory) <> "" Then
            RememberedFolder = folder
            Exit Function
        End If
    End If
    With Application.FileDialog(PICK_FOLDER)
        .Title = prompt
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        folder = .SelectedItems(1)
    End With
    If Right(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, key, folder
    RememberedFolder = folder
End Function

Private Function PickFile(ByVal prompt As String, ByVal initial As String, ByVal kind As String) As String
    With Application.FileDialog(PICK_FILE)
        .Title = prompt
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add kind, "*.xlsx"
        If initial <> "" Then .InitialFileName = initial
        If .Show <> -1 Then Exit Function
        PickFile = .SelectedItems(1)
    End With
End Function

' The ID lookup's file name for a Collection Workbook, also when the copy was
' downloaded under a changed name such as "..._CY2026 (1).xlsx".
Private Function LookupNameFor(ByVal collectionPath As String) As String
    Dim name As String, i As Long
    name = BaseName(collectionPath)
    i = InStr(1, name, COLLECTION_PREFIX, vbTextCompare)
    If i > 0 Then name = Mid(name, i, Len(COLLECTION_PREFIX) + 4)
    LookupNameFor = name & " " & ChrW(8211) & " ID lookup.xlsx"
End Function

Private Function ReportCount(ByVal entry As String) As Long
    Dim i As Long
    For i = 1 To mReportCount
        If mReport(i)(0) = entry Then ReportCount = ReportCount + 1
    Next i
End Function

Private Function TotalCount(ByVal what As Long) As Long
    Dim t As Long
    For t = 1 To 6
        TotalCount = TotalCount + mCounts(t, what)
    Next t
End Function

' ===========================================================================
' Export Staff Copy
' ===========================================================================

' Writes the Collection Workbook for calendarYear into collectionFolder and
' returns its path. The ID lookup and the export report go into adminFolder.
' If the Collection Workbook already exists, it's replaced only when
' replaceExisting is True; otherwise nothing is written and "" is returned.
Public Function RunExportStaffCopy(ByVal calendarYear As Long, ByVal mainDatabase As Object, ByVal collectionFolder As String, ByVal adminFolder As String, Optional ByVal replaceExisting As Boolean = False) As String
    Dim path As String, lookupPath As String, reportPath As String
    path = collectionFolder & CollectionName(calendarYear) & ".xlsx"
    If FileExists(path) And Not replaceExisting Then Exit Function
    lookupPath = adminFolder & CollectionName(calendarYear) & " " & ChrW(8211) & " ID lookup.xlsx"
    reportPath = adminFolder & CollectionName(calendarYear) & " " & ChrW(8211) & " export report " & Format(Now, "yyyy-mm-dd") & ".xlsx"
    EnsureNameFree path
    EnsureNameFree lookupPath
    EnsureNameFree reportPath

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
    WriteLookup lookupPath
    WriteExportReport reportPath, calendarYear
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

    Dim cols() As Long, n As Long, c As Long, r As Long, outRow As Long, idCol As Long, width As Double
    n = ExportColumns(spec, src, cols)
    idCol = n + 1
    For c = 1 To n
        ws.Cells(HEADER_ROW, c).Value = src.Headers(cols(c))
        width = srcSheet.Columns(src.Cols(cols(c))).ColumnWidth
        If width <= 0 Then width = DEFAULT_COLUMN_WIDTH        ' a hidden column reads as width 0
        If width < MIN_COLUMN_WIDTH Then width = MIN_COLUMN_WIDTH
        ws.Columns(c).ColumnWidth = Application.WorksheetFunction.Min(width, MAX_COLUMN_WIDTH)
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
            If list <> "" Then
                If Not AddDropdown(ws.Range(ws.Cells(HEADER_ROW + 1, c), ws.Cells(outRow + ROOM_FOR_NEW_ROWS, c)), list) Then
                    ReportLine "Dropdown not carried over", spec.Title, "", Tidy(src.Headers(cols(c))), list
                End If
            End If
        End If
    Next c
    LayOutTable ws, Application.WorksheetFunction.Max(outRow, HEADER_ROW + 1), idCol
    LayOutNote ws, n
    MakeTable ws, spec, Application.WorksheetFunction.Max(outRow, HEADER_ROW + 1), idCol
End Sub

' Centred, wrapped text in rows with room to breathe and a light grid, under a
' bold header row in white on dark blue. The rows left for new entries get the
' same alignment and height.
Private Sub LayOutTable(ByVal ws As Object, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim r As Long, fitted As Double
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow + ROOM_FOR_NEW_ROWS, lastCol))
        .HorizontalAlignment = XL_CENTER
        .VerticalAlignment = XL_CENTER
        .WrapText = True
    End With
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol))
        .Font.Bold = True
        .Font.Size = HEADER_FONT_SIZE
        .Font.Color = WHITE
        .Interior.Color = HEADER_FILL
    End With
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol)).Borders
        .LineStyle = XL_CONTINUOUS
        .Color = GRID_COLOUR
    End With
    For r = HEADER_ROW To lastRow
        ws.Rows(r).AutoFit
        fitted = ws.Rows(r).RowHeight + ROW_PADDING
        If fitted < MIN_ROW_HEIGHT Then fitted = MIN_ROW_HEIGHT
        If fitted > MAX_ROW_HEIGHT Then fitted = MAX_ROW_HEIGHT
        ws.Rows(r).RowHeight = fitted
    Next r
    ws.Range(ws.Rows(lastRow + 1), ws.Rows(lastRow + ROOM_FOR_NEW_ROWS)).RowHeight = MIN_ROW_HEIGHT
End Sub

' The note goes in one cell merged across the table's columns, in large type.
' Excel can't fit a merged cell's height to its text, so the row height comes
' from the note's length and the width of the visible columns.
Private Sub LayOutNote(ByVal ws As Object, ByVal lastCol As Long)
    Dim c As Long, width As Double, perLine As Long, lines As Long, para As Variant, height As Double
    If lastCol < 1 Then lastCol = 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .WrapText = True
        .HorizontalAlignment = XL_LEFT
        .VerticalAlignment = XL_CENTER
        .Font.Size = NOTE_FONT_SIZE
        .Interior.Color = NOTE_FILL
    End With
    On Error Resume Next        ' the bold heading line is a nicety
    ws.Cells(1, 1).Characters(1, InStr(CStr(ws.Cells(1, 1).Value), vbLf) - 1).Font.Bold = True
    On Error GoTo 0
    For c = 1 To lastCol
        If Not ws.Columns(c).Hidden Then width = width + ws.Columns(c).ColumnWidth
    Next c
    perLine = Int(width * 11 / NOTE_FONT_SIZE * 0.85)     ' column widths count 11-point characters
    If perLine < 20 Then perLine = 20
    For Each para In Split(CStr(ws.Cells(1, 1).Value), vbLf)
        lines = lines + 1 + Int(Len(para) / perLine)
    Next para
    height = lines * NOTE_FONT_SIZE * 1.4 + 12
    If height > MAX_ROW_HEIGHT Then height = MAX_ROW_HEIGHT
    ws.Rows(1).RowHeight = height
    ws.Rows(2).RowHeight = 8
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

' Returns False if Excel refuses the list (for example, longer than 255 characters).
Private Function AddDropdown(ByVal target As Object, ByVal list As String) As Boolean
    On Error Resume Next
    target.Validation.Delete
    Err.Clear
    target.Validation.Add Type:=XL_VALIDATE_LIST, AlertStyle:=1, Operator:=1, Formula1:=list
    AddDropdown = (Err.Number = 0)
    Err.Clear
End Function

Private Sub MakeTable(ByVal ws As Object, ByRef spec As TabSpec, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim t As Object
    On Error Resume Next        ' LibreOffice has no ListObjects; the sheet works without one
    Set t = ws.ListObjects.Add(1, ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol)), , 1)
    If Not t Is Nothing Then
        t.Name = "Collection_" & spec.Prefix
        t.ShowTableStyleRowStripes = False
        AddTableStyle ws.Parent
        t.TableStyle = TABLE_STYLE
    End If
    Err.Clear
End Sub

' A table style with the same grid and header colours as the cell formatting,
' and no stripes, so rows typed below the table look like the rest. (A table
' style can't set font size, so the header's size comes from cell formatting.)
Private Sub AddTableStyle(ByVal wb As Object)
    Const WHOLE_TABLE As Long = 0, HEADER_ROW_ELEMENT As Long = 1
    Dim s As Object, b As Variant
    On Error Resume Next
    Set s = wb.TableStyles(TABLE_STYLE)
    On Error GoTo 0
    If Not s Is Nothing Then Exit Sub
    Set s = wb.TableStyles.Add(TABLE_STYLE)
    For Each b In Array(7, 8, 9, 10, 11, 12)    ' xlEdgeLeft to xlInsideHorizontal
        With s.TableStyleElements(WHOLE_TABLE).Borders(b)
            .LineStyle = XL_CONTINUOUS
            .Color = GRID_COLOUR
        End With
    Next b
    With s.TableStyleElements(HEADER_ROW_ELEMENT)
        .Interior.Color = HEADER_FILL
        .Font.Bold = True
        .Font.Color = WHITE
    End With
End Sub

Private Function FindYourRecordsNote(ByRef spec As TabSpec) As String
    FindYourRecordsNote = "How to find your records" & vbLf & _
        "1. If Excel offers Sheet View (View > Sheet View > New), choose it first, " & _
        "so that your sorting and filtering don't change what colleagues see." & vbLf & _
        "2. Press Ctrl+F and search for your name, or use the filter arrow on the " & spec.NameColumn & " column." & vbLf & _
        "Something new: type it in a new row below the existing ones. " & _
        "Something changed, such as a status: update your existing row."
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
            If Not MatchesAny(src.Keys(c), spec.Drop) And src.Keys(c) <> NO_HEADER Then
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
    If Left(Norm(status), 7) = "dropped" Then
        PickProject = PICK_OUT
    ElseIf IsBlank(loaStart) Then
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
' Check Collection Workbook
' ===========================================================================

' Compares the Collection Workbook at collectionPath with the Main Database,
' using the ID lookup written when it was exported, and saves a highlighted
' copy in adminFolder. Returns the copy's path. The Collection Workbook itself
' is never opened: the copy is made first and only the copy is marked, so
' AutoSave on a synced SharePoint file can't carry the marks back.
Public Function RunCheckCollectionWorkbook(ByVal collectionPath As String, ByVal lookupPath As String, ByVal mainDatabase As Object, ByVal adminFolder As String) As String
    Dim cw As Object, copyPath As String, i As Long, j As Long
    LoadTabs
    ReadLookup lookupPath
    mReportCount = 0
    For i = 1 To 6
        For j = 1 To 6
            mCounts(i, j) = 0
        Next j
    Next i

    copyPath = adminFolder & BaseName(collectionPath) & " " & ChrW(8211) & " checked " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    EnsureNameFree copyPath
    CopyFile collectionPath, copyPath
    Set cw = Workbooks.Open(copyPath)
    For i = 1 To 6
        CheckTab cw, mainDatabase, i
    Next i
    ListDeleted
    WriteSummary cw, BaseName(collectionPath)
    SaveInPlace cw, copyPath
    RunCheckCollectionWorkbook = copyPath
End Function

Private Sub CheckTab(ByVal cw As Object, ByVal mainDatabase As Object, ByVal t As Long)
    Dim spec As TabSpec, ws As Object, md As SheetRows, rows As SheetRows
    spec = mTabs(t)
    On Error Resume Next
    Set ws = cw.Worksheets(spec.Title)
    On Error GoTo 0
    If ws Is Nothing Then
        ReportLine "Tab missing", spec.Title, "", "", ""
        Exit Sub
    End If
    md = ReadMainSheet(mainDatabase, spec)
    If Not md.Found Then
        ReportLine "Sheet or table not found in the Main Database", spec.Title, "", "", spec.SheetName
        Exit Sub
    End If
    rows = ReadCollectionSheet(ws)
    If Not rows.Found Then
        ReportLine "No Export ID column", spec.Title, "", "", ""
        Exit Sub
    End If

    ' Which Main Database column each Collection Workbook column is compared with (0: none).
    Dim exported() As Long, n As Long, mdCol() As Long, c As Long, k As Long, idCol As Long
    n = ExportColumns(spec, md, exported)
    ReDim mdCol(1 To rows.ColCount)
    idCol = FindColumn(rows, Norm(EXPORT_ID_HEADER))
    For c = 1 To rows.ColCount
        For k = 1 To n
            If md.Keys(exported(k)) = rows.Keys(c) Then mdCol(c) = exported(k)
        Next k
        ' Collection Workbooks exported before No. was left out still have it.
        If mdCol(c) = 0 And rows.Keys(c) = NO_HEADER Then mdCol(c) = FindColumn(md, NO_HEADER)
        If mdCol(c) = 0 And c <> idCol And rows.Keys(c) <> "" And rows.Keys(c) <> Norm(CHECK_NOTE_HEADER) Then
            mCounts(t, COUNT_EXTRA) = mCounts(t, COUNT_EXTRA) + 1
            ReportLine "Extra column", spec.Title, "", "", rows.Headers(c)
        End If
    Next c

    Dim noteCol As Long, r As Long, sheetRow As Long, id As String, li As Long, found As Long, diffs As Long, multiple As String
    noteCol = LastUsedColumn(ws) + 1
    ws.Cells(rows.HeaderRow, noteCol).Value = CHECK_NOTE_HEADER
    ws.Cells(rows.HeaderRow, noteCol).Font.Bold = True
    For r = 1 To rows.RowCount
        sheetRow = rows.RowNums(r)
        ListValuesOutside ws, spec, rows, r
        If IsRecord(rows, r, mdCol, idCol) Then
            id = Trim(CellText(rows.Values(r, idCol)))
            If id <> "" Then
                li = LookupIndex(id)
                found = 0
                multiple = ""
                If li > 0 Then
                    mIdSeen(li) = True
                    If mIds(li, 1) = spec.Title Then found = FindRecord(md, spec, mIds(li, 2), mIds(li, 4), mIds(li, 5), multiple)
                End If
                If found = 0 Then
                    If multiple = "" Then
                        Mark ws, rows, r, mdCol, noteCol, 0, "Can't find in Main Database"
                    Else
                        Mark ws, rows, r, mdCol, noteCol, 0, "Can't find in Main Database: more than one record matches (Main Database rows " & multiple & ")"
                    End If
                    Tally t, COUNT_CANT_FIND, "Can't find in Main Database", sheetRow, id, CellText(LookupTitle(li))
                Else
                    diffs = MarkDifferences(ws, rows, r, mdCol, md, found, False)
                    If diffs > 0 Then
                        Mark ws, rows, r, mdCol, noteCol, YELLOW, "Changed: " & RecordName(md, spec, found)
                        Tally t, COUNT_CHANGED, "Changed", sheetRow, id, RecordName(md, spec, found)
                    End If
                End If
            Else
                found = FindRecord(md, spec, RowField(rows, r, spec.NoHeader), RowField(rows, r, spec.Key1Header), RowField(rows, r, spec.Key2Header), multiple)
                If found = 0 And multiple <> "" Then
                    Mark ws, rows, r, mdCol, noteCol, 0, "Already recorded more than once (Main Database rows " & multiple & ")"
                    Tally t, COUNT_ALREADY, "Already recorded more than once", sheetRow, "", "Main Database rows " & multiple
                ElseIf found = 0 Then
                    Mark ws, rows, r, mdCol, noteCol, GREEN, "New"
                    Tally t, COUNT_NEW, "New", sheetRow, "", CellText(RowField(rows, r, spec.TitleHeader))
                Else
                    diffs = MarkDifferences(ws, rows, r, mdCol, md, found, True)
                    If diffs > 0 Then
                        Mark ws, rows, r, mdCol, noteCol, YELLOW, "Already recorded as " & RecordName(md, spec, found)
                    Else
                        Mark ws, rows, r, mdCol, noteCol, 0, "Already recorded as " & RecordName(md, spec, found)
                    End If
                    Tally t, COUNT_ALREADY, "Already recorded", sheetRow, "", RecordName(md, spec, found)
                End If
            End If
        End If
    Next r
End Sub

' A row is a record when it has an Export ID or a value in a compared column.
Private Function IsRecord(ByRef rows As SheetRows, ByVal r As Long, ByRef mdCol() As Long, ByVal idCol As Long) As Boolean
    Dim c As Long
    For c = 1 To rows.ColCount
        If mdCol(c) > 0 Or c = idCol Then
            If Not IsBlank(rows.Values(r, c)) Then
                IsRecord = True
                Exit Function
            End If
        End If
    Next c
End Function

Private Sub ListValuesOutside(ByVal ws As Object, ByRef spec As TabSpec, ByRef rows As SheetRows, ByVal r As Long)
    Dim c As Long
    For c = 1 To rows.ColCount
        If rows.Keys(c) = "" And Not IsBlank(rows.Values(r, c)) Then
            ReportLine "Value outside the exported columns", spec.Title, rows.RowNums(r), _
                ws.Cells(rows.RowNums(r), rows.Cols(c)).Address(False, False), CellText(rows.Values(r, c))
        End If
    Next c
End Sub

' Bolds each compared cell that differs from the Main Database record and puts
' the Main Database value in a comment on it. Returns how many differ.
' filledOnly skips cells left empty: a row typed again often holds only the
' fields that changed.
Private Function MarkDifferences(ByVal ws As Object, ByRef rows As SheetRows, ByVal r As Long, ByRef mdCol() As Long, ByRef md As SheetRows, ByVal found As Long, ByVal filledOnly As Boolean) As Long
    Dim c As Long, cell As Object, inScope As Boolean
    For c = 1 To rows.ColCount
        inScope = (mdCol(c) > 0)
        If inScope And filledOnly Then inScope = Not IsBlank(rows.Values(r, c))
        If inScope Then
            If Not SameValue(rows.Values(r, c), md.Values(found, mdCol(c))) Then
                Set cell = ws.Cells(rows.RowNums(r), rows.Cols(c))
                cell.Font.Bold = True
                cell.ClearComments
                cell.AddComment "Main Database: " & ShowValue(md.Values(found, mdCol(c)))
                MarkDifferences = MarkDifferences + 1
            End If
        End If
    Next c
End Function

' Fills a row's compared cells (unless fill is 0) and writes its note.
Private Sub Mark(ByVal ws As Object, ByRef rows As SheetRows, ByVal r As Long, ByRef mdCol() As Long, ByVal noteCol As Long, ByVal fill As Long, ByVal note As String)
    Dim c As Long
    If fill <> 0 Then
        For c = 1 To rows.ColCount
            If mdCol(c) > 0 Then ws.Cells(rows.RowNums(r), rows.Cols(c)).Interior.Color = fill
        Next c
    End If
    WriteValue ws.Cells(rows.RowNums(r), noteCol), note
End Sub

Private Sub Tally(ByVal t As Long, ByVal what As Long, ByVal entry As String, ByVal sheetRow As Variant, ByVal id As String, ByVal detail As String)
    mCounts(t, what) = mCounts(t, what) + 1
    ReportLine entry, mTabs(t).Title, sheetRow, id, detail
End Sub

' "No. 6, <title>", or the title alone when the record has no No.
Private Function RecordName(ByRef md As SheetRows, ByRef spec As TabSpec, ByVal found As Long) As String
    Dim no As Variant
    no = ColumnValue(md, found, spec.NoHeader)
    RecordName = CellText(ColumnValue(md, found, spec.TitleHeader))
    If Not IsBlank(no) Then RecordName = "No. " & CellText(no) & ", " & RecordName
End Function

Private Function RowField(ByRef rows As SheetRows, ByVal r As Long, ByVal pattern As String) As Variant
    RowField = ColumnValue(rows, r, pattern)
End Function

' The Main Database record a set of match key values points to: the first key,
' in the tab's order, that finds exactly one record. A key is skipped when its
' first field is empty. Returns 0 if no key finds exactly one; then, if a key
' found several, multiple lists their sheet rows ("5, 7").
Private Function FindRecord(ByRef md As SheetRows, ByRef spec As TabSpec, ByVal v0 As Variant, ByVal v1 As Variant, ByVal v2 As Variant, ByRef multiple As String) As Long
    Dim key As Variant, fields As Variant, values(0 To 2) As Variant, headers(0 To 2) As String
    Dim cols() As Long, f As Long, r As Long, hits As Long, hit As Long, usable As Boolean, rowList As String
    multiple = ""
    values(0) = v0: values(1) = v1: values(2) = v2
    headers(0) = spec.NoHeader: headers(1) = spec.Key1Header: headers(2) = spec.Key2Header
    For Each key In Split(spec.Keys, ";")
        fields = Split(key, "+")
        ReDim cols(0 To UBound(fields))
        usable = Not IsBlank(values(CLng(fields(0))))
        For f = 0 To UBound(fields)
            cols(f) = 0
            If headers(CLng(fields(f))) <> "" Then cols(f) = FindColumn(md, headers(CLng(fields(f))))
            If cols(f) = 0 Then usable = False
        Next f
        If usable Then
            hits = 0
            rowList = ""
            For r = 1 To md.RowCount
                If KeyMatches(md, r, cols, fields, values) Then
                    hits = hits + 1
                    hit = r
                    rowList = rowList & ", " & md.RowNums(r)
                End If
            Next r
            If hits > 1 And multiple = "" Then multiple = Mid(rowList, 3)
            If hits = 1 Then
                FindRecord = hit
                Exit Function
            End If
        End If
    Next key
End Function

Private Function KeyMatches(ByRef md As SheetRows, ByVal r As Long, ByRef cols() As Long, ByVal fields As Variant, ByRef values() As Variant) As Boolean
    Dim f As Long
    For f = 0 To UBound(fields)
        If Not SameValue(md.Values(r, cols(f)), values(CLng(fields(f)))) Then Exit Function
    Next f
    KeyMatches = True
End Function

Private Sub ListDeleted()
    Dim i As Long, t As Long
    For i = 1 To mIdCount
        If Not mIdSeen(i) Then
            For t = 1 To 6
                If mTabs(t).Title = mIds(i, 1) Then
                    Tally t, COUNT_DELETED, "Deleted", "", CStr(mIds(i, 0)), CellText(mIds(i, 3))
                End If
            Next t
        End If
    Next i
End Sub

Private Sub WriteSummary(ByVal cw As Object, ByVal name As String)
    Dim ws As Object, t As Long, k As Long, r As Long, c As Long, entry As Variant
    Set ws = cw.Worksheets.Add(Before:=cw.Worksheets(1))
    ws.Name = "Summary"
    ws.Range("A1").Value = "Check of " & name & " against the Main Database, " & Format(Now, "d mmm yyyy hh:mm")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:G3").Value = Array("Tab", "New", "Changed", "Already recorded", "Deleted", "Can't find in Main Database", "Extra columns")
    ws.Range("A3:G3").Font.Bold = True
    For t = 1 To 6
        ws.Cells(3 + t, 1).Value = mTabs(t).Title
        For k = 1 To 6
            ws.Cells(3 + t, 1 + k).Value = mCounts(t, k)
        Next k
    Next t
    ws.Range("A11").Value = "Details"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12:E12").Value = Array("What", "Tab", "Row", "Export ID / cell", "Detail")
    ws.Range("A12:E12").Font.Bold = True
    For r = 1 To mReportCount
        entry = mReport(r)
        For c = 0 To 4
            WriteValue ws.Cells(12 + r, c + 1), entry(c)
        Next c
    Next r
    ws.Columns("A:G").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 60
    ws.Activate
End Sub

Private Sub ReadLookup(ByVal lookupPath As String)
    Dim wb As Object, ws As Object, last As Long, r As Long, c As Long
    Set wb = Workbooks.Open(lookupPath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)
    last = LastFilledRow(ws, 1)
    mIdCount = Application.WorksheetFunction.Max(last - 1, 0)
    Dim ids() As Variant, seen() As Boolean
    ReDim ids(1 To Application.WorksheetFunction.Max(mIdCount, 1), 0 To 5)
    ReDim seen(1 To Application.WorksheetFunction.Max(mIdCount, 1))
    For r = 1 To mIdCount
        For c = 0 To 5
            ids(r, c) = ws.Cells(r + 1, c + 1).Value
        Next c
        ids(r, 0) = Trim(CellText(ids(r, 0)))
        ids(r, 1) = CellText(ids(r, 1))
    Next r
    mIds = ids
    mIdSeen = seen
    wb.Close False
End Sub

Private Function LookupIndex(ByVal id As String) As Long
    Dim i As Long
    For i = 1 To mIdCount
        If StrComp(mIds(i, 0), id, vbTextCompare) = 0 Then
            LookupIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function LookupTitle(ByVal li As Long) As Variant
    If li > 0 Then LookupTitle = mIds(li, 3)
End Function

' Reads a Collection Workbook tab: the header row is the one holding Export ID,
' and every column up to the last one used is read, headed or not.
Private Function ReadCollectionSheet(ByVal ws As Object) As SheetRows
    Dim headerRow As Long, firstCol As Long, lastCol As Long
    If Not ScanHeader(ws, Norm(EXPORT_ID_HEADER), headerRow, firstCol, lastCol) Then Exit Function
    ReadCollectionSheet = ReadBlock(ws, headerRow, 1, LastUsedColumn(ws))
End Function

' Copies a file, or saves a copy of it if it's open in Excel (FileCopy can't
' read a file Excel has open).
Private Sub CopyFile(ByVal fromPath As String, ByVal toPath As String)
    Dim wb As Object
    On Error Resume Next
    If FileExists(toPath) Then Kill toPath
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise MACRO_ERROR + 2, , "Close " & toPath & " and run the check again."
    End If
    FileCopy fromPath, toPath
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    For Each wb In Workbooks
        If StrComp(wb.FullName, fromPath, vbTextCompare) = 0 Then
            wb.SaveCopyAs toPath
            Exit Sub
        End If
    Next wb
    On Error GoTo 0
    Err.Raise MACRO_ERROR + 1, , "Can't copy " & fromPath & ". Close it, or pick a downloaded copy, and run the check again."
End Sub

' ===========================================================================
' Tabs
' ===========================================================================

Private Sub LoadTabs()
    With mTabs(1)
        .Title = "Projects": .NameColumn = "PI or Co-I/Collaborator": .Prefix = "PRJ"
        .SheetName = "DB_PrjGeneralInfo": .TableName = "tbl_PrjGeneralInfo": .Anchor = "project name"
        .Only = "grant ref no|project name|category|grant / collaborator partners|grant type|pi|" & _
                "co-i/collaborator|loa start date|loa end date|grant amount|grant amt (awarded to ttsh)|status"
        .Rule = RULE_PROJECT
        .NoHeader = "": .TitleHeader = "project name": .Key1Header = "grant ref no": .Key2Header = "project name"
        .Keys = "1;2"
    End With
    With mTabs(2)
        .Title = "Publications": .NameColumn = "TTSH Contributor": .Prefix = "PUB"
        .SheetName = "DB_Publications": .TableName = "tbl_Publication": .Anchor = "title of article"
        .Drop = "journal category & jif percentile listing*|jif percentile*"
        .Rule = RULE_DATE: .DateHeader = "date of publication"
        .NoHeader = "no.": .TitleHeader = "title of article": .Key1Header = "doi": .Key2Header = "title of article"
        .Keys = "0;1;2"
    End With
    With mTabs(3)
        .Title = "Poster & Oral Presentation": .NameColumn = "Full Name": .Prefix = "PRES"
        .SheetName = "DB_Poster& Oral Presentation": .TableName = "tbl_PosterOralPresentation"
        .Anchor = "title of presentation/topic"
        .Rule = RULE_DATE: .DateHeader = "date of presentation"
        .NoHeader = "no.": .TitleHeader = "title of presentation/topic": .Key1Header = "title of presentation/topic": .Key2Header = ""
        .Keys = "0;1"
    End With
    With mTabs(4)
        .Title = "Education": .NameColumn = "Trainer(s)/Speaker(s)": .Prefix = "EDU"
        .SheetName = "DB_Education": .TableName = "tbl_Education": .Anchor = "title of course*"
        .Rule = RULE_YEAR_THEN_DATE: .YearHeader = "year": .DateHeader = "date/year/period"
        .NoHeader = "": .TitleHeader = "title of course*": .Key1Header = "title of course*": .Key2Header = "year"
        .Keys = "1+2"
    End With
    With mTabs(5)
        .Title = "Media": .NameColumn = "Full name": .Prefix = "MED"
        .SheetName = "DB_Media": .TableName = "": .Anchor = "platform of media"
        .Rule = RULE_DATE: .DateHeader = "media publishing date"
        .NoHeader = "no.": .TitleHeader = "title of media published*": .Key1Header = "title of media published*": .Key2Header = "name of media published source*"
        .Keys = "1+2"
    End With
    With mTabs(6)
        .Title = "Talent Development": .NameColumn = "Name": .Prefix = "TD"
        .SheetName = "DB_Talent Develop": .TableName = "tbl_TalentDevelop": .Anchor = "award description"
        .Rule = RULE_YEAR: .YearHeader = "year"
        .NoHeader = "": .TitleHeader = "award description": .Key1Header = "name": .Key2Header = "award description"
        .Keys = "1+2"
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
        r = LastFilledRow(ws, c)
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

' The last row with a value in column c (0 if none). Scans the column's used
' cells as one array, from the bottom, rather than using End(xlUp), which
' LibreOffice works out from whichever window is active.
Private Function LastFilledRow(ByVal ws As Object, ByVal c As Long) As Long
    Dim bottom As Long, v As Variant, r As Long
    bottom = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If bottom = 1 Then
        If Not IsEmpty(ws.Cells(1, c).Value) Then LastFilledRow = 1
        Exit Function
    End If
    v = ws.Range(ws.Cells(1, c), ws.Cells(bottom, c)).Value
    For r = bottom To 1 Step -1
        If Not IsEmpty(v(r, 1)) Then
            LastFilledRow = r
            Exit Function
        End If
    Next r
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

' Whether two cell values are the same once formatting is set aside: spacing,
' line breaks and case; empty against "(blank)"; numbers compared as numbers
' and dates as dates, however each is formatted or stored.
Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim x As Double, y As Double, d1 As Date, d2 As Date
    If IsBlank(a) Or IsBlank(b) Then
        SameValue = (IsBlank(a) And IsBlank(b))
    ElseIf Norm(a) = Norm(b) Then
        SameValue = True
    ElseIf ReadNumber(a, x) And ReadNumber(b, y) Then
        SameValue = (Abs(x - y) < 0.000001)
    ElseIf ReadDate(a, d1) And ReadDate(b, d2) Then
        SameValue = (Int(CDbl(d1)) = Int(CDbl(d2)))
    End If
End Function

' A value as IREx Admin should see it in a comment.
Private Function ShowValue(ByVal v As Variant) As String
    If IsBlank(v) Then
        ShowValue = "(empty)"
    ElseIf VarType(v) = vbDate Then
        ShowValue = Format(v, "d mmm yyyy")
    Else
        ShowValue = CellText(v)
    End If
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

' The file name without its folder.
Private Function FileName(ByVal path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid(path, i, 1) = "\" Or Mid(path, i, 1) = "/" Then Exit For
    Next i
    FileName = Mid(path, i + 1)
End Function

' The file name without its folder or extension.
Private Function BaseName(ByVal path As String) As String
    BaseName = FileName(path)
    If InStrRev(BaseName, ".") > 0 Then BaseName = Left(BaseName, InStrRev(BaseName, ".") - 1)
End Function

' Raises an error if a workbook with path's file name is open, from any folder
' or from SharePoint: Excel can't save or open a second workbook of that name.
Private Sub EnsureNameFree(ByVal path As String)
    Dim wb As Object, name As String
    name = FileName(path)
    For Each wb In Workbooks
        If StrComp(wb.Name, name, vbTextCompare) = 0 Then
            Err.Raise MACRO_ERROR + 4, , name & " is open in Excel (" & wb.FullName & "). " & _
                "Excel can't have two workbooks of the same name open, so close it and run the macro again."
        End If
    Next wb
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    Err.Clear
End Function

'Saves an open workbook over its own file as .xlsx, and closes it.
Private Sub SaveInPlace(ByVal wb As Object, ByVal path As String)
    Dim alerts As Boolean
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=XL_WORKBOOK
    Application.DisplayAlerts = alerts
    wb.Close False
End Sub

' Saves a new workbook as .xlsx at path, replacing any file there, and closes it.
Private Sub SaveNew(ByVal wb As Object, ByVal path As String)
    Dim alerts As Boolean
    On Error Resume Next
    If FileExists(path) Then Kill path
    If Err.Number <> 0 Then
        On Error GoTo 0
        wb.Close False
        Err.Raise MACRO_ERROR + 3, , "Close " & path & " (it may be open on this PC) and run the macro again."
    End If
    On Error GoTo 0
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=XL_WORKBOOK
    Application.DisplayAlerts = alerts
    wb.Close False
End Sub
