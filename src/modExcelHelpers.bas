Attribute VB_Name = "modExcelHelpers"
Option Explicit

Public Function LastBusinessDay(ByVal d As Date) As Date
    Dim t As Date
    t = d
    If Weekday(t, vbMonday) = 6 Then t = t - 1
    If Weekday(t, vbMonday) = 7 Then t = t - 2
    LastBusinessDay = t
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then
        Set c = ws.Cells.Find(What:="*", LookIn:=xlValues, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    End If
    LastUsedRow = IIf(c Is Nothing, 1, c.Row)
End Function

Public Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then
        Set c = ws.Cells.Find(What:="*", LookIn:=xlValues, _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    End If
    LastUsedCol = IIf(c Is Nothing, 1, c.Column)
End Function

Public Function SheetExists(ByVal sheetName As String, ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function IsInArray(ByVal value As String, ByVal arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If StrComp(CStr(v), value, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next v
End Function
