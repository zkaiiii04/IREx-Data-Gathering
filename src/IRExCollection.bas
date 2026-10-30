Attribute VB_Name = "IRExCollection"
' IREx Collection Workbook macros, run from the Main Database.
Option Explicit

Private Const COLLECTION_PREFIX As String = "IREX_Rehab_Data_Collection_CY"
Private Const XL_WORKBOOK As Long = 51          ' xlOpenXMLWorkbook (.xlsx)
Private Const XL_SINGLE_SHEET As Long = -4167   ' xlWBATWorksheet

Private Type TabSpec
    Title As String
End Type

Private mTabs(1 To 6) As TabSpec

Public Function RunExportStaffCopy(ByVal year As Long, ByVal mainDatabase As Object, ByVal collectionFolder As String, ByVal adminFolder As String, Optional ByVal replaceExisting As Boolean = False) As String
    LoadTabs
    Dim path As String
    path = collectionFolder & COLLECTION_PREFIX & year & ".xlsx"

    Dim cw As Object, i As Long, ws As Object
    Set cw = Workbooks.Add(XL_SINGLE_SHEET)
    For i = 1 To 6
        If i = 1 Then
            Set ws = cw.Worksheets(1)
        Else
            Set ws = cw.Worksheets.Add(After:=cw.Worksheets(cw.Worksheets.Count))
        End If
        ws.Name = mTabs(i).Title
    Next i
    SaveAndClose cw, path
    RunExportStaffCopy = path
End Function

Private Sub LoadTabs()
    mTabs(1).Title = "Publications"
    mTabs(2).Title = "Poster & Oral Presentation"
    mTabs(3).Title = "Media"
    mTabs(4).Title = "Talent Development"
    mTabs(5).Title = "Education"
    mTabs(6).Title = "Projects"
End Sub

Private Sub SaveAndClose(ByVal wb As Object, ByVal path As String)
    Dim alerts As Boolean
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=XL_WORKBOOK
    Application.DisplayAlerts = alerts
    wb.Close False
End Sub
