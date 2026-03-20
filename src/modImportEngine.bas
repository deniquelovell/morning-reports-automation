Attribute VB_Name = "modImportEngine"
Option Explicit

Public Sub ClearAllWorkbookFilters(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next

    For Each ws In wb.Worksheets
        If ws.FilterMode Then ws.ShowAllData

        For Each lo In ws.ListObjects
            If Not lo Is Nothing Then
                If lo.ShowAutoFilter Then
                    If Not lo.AutoFilter Is Nothing Then
                        If lo.AutoFilter.FilterMode Then
                            lo.AutoFilter.ShowAllData
                        End If
                    End If
                End If
            End If
        Next lo
    Next ws

    On Error GoTo 0
End Sub

Public Sub PasteRangeFast(ByVal srcWS As Worksheet, _
                          ByVal nRows As Long, _
                          ByVal nCols As Long, _
                          ByVal destWS As Worksheet, _
                          ByVal dstRow As Long, _
                          ByVal dstCol As Long)

    Dim srcRng As Range
    Dim dstRng As Range

    Set srcRng = srcWS.Range(srcWS.Cells(1, 1), srcWS.Cells(nRows, nCols))
    Set dstRng = destWS.Range(destWS.Cells(dstRow, dstCol), _
                              destWS.Cells(dstRow + nRows - 1, dstCol + nCols - 1))

    dstRng.Value2 = srcRng.Value2
End Sub

Public Function PickSourceWorksheet(ByVal wb As Workbook, ByVal destName As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function

    Select Case LCase$(destName)
        Case "IMS Real Time Grid"
            Set PickSourceWorksheet = wb.Worksheets(1)

        Case "broker position summary", _
             "broker margin detail", _
             "broker debit credit interest", _
             "broker stock borrow"
            Set PickSourceWorksheet = wb.Worksheets(1)

        Case Else
            For Each ws In wb.Worksheets
                If LCase$(Trim$(ws.Name)) = "sheet1" Then
                    Set PickSourceWorksheet = ws
                    Exit Function
                End If
            Next ws

            Set PickSourceWorksheet = wb.Worksheets(1)
    End Select
End Function
