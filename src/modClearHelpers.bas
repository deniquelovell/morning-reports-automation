Attribute VB_Name = "modClearHelpers"
Option Explicit

Public Sub SafeClearFromAnchorFast(ByVal ws As Worksheet, _
                                   ByVal anchorRow As Long, _
                                   ByVal anchorCol As Long, _
                                   ByVal preserveFormatting As Boolean, _
                                   ByVal sheetName As String)

    Dim lastR As Long, lastC As Long
    Dim target As Range

    GetUsedBoundsFromAnchor ws, anchorRow, anchorCol, lastR, lastC

    If lastR < anchorRow Then lastR = anchorRow
    If lastC < anchorCol Then lastC = anchorCol

    Set target = ws.Range(ws.Cells(anchorRow, anchorCol), ws.Cells(lastR, lastC))

    If REMOVE_PICTURES_IN_CLEARED_AREAS Then DeletePicturesInRangeFast ws, target

    If preserveFormatting Then
        target.ClearContents
    Else
        target.Clear
    End If
End Sub

Public Sub SafeClearBelowAnchorFast(ByVal ws As Worksheet, _
                                    ByVal anchorRow As Long, _
                                    ByVal anchorCol As Long, _
                                    ByVal importedRows As Long, _
                                    ByVal preserveFormatting As Boolean, _
                                    ByVal sheetName As String)

    Dim startClearRow As Long, lastR As Long, lastC As Long
    Dim target As Range

    startClearRow = anchorRow + importedRows
    If startClearRow < anchorRow Then Exit Sub

    GetUsedBoundsFromAnchor ws, anchorRow, anchorCol, lastR, lastC
    If lastR = 0 Or lastC = 0 Then Exit Sub
    If startClearRow > lastR Then Exit Sub

    Set target = ws.Range(ws.Cells(startClearRow, anchorCol), ws.Cells(lastR, lastC))

    If REMOVE_PICTURES_IN_CLEARED_AREAS Then DeletePicturesInRangeFast ws, target

    If preserveFormatting Then
        target.ClearContents
    Else
        target.Clear
    End If
End Sub

Public Sub SafeClearRightOfAnchorFast(ByVal ws As Worksheet, _
                                      ByVal anchorRow As Long, _
                                      ByVal anchorCol As Long, _
                                      ByVal importedCols As Long, _
                                      ByVal preserveFormatting As Boolean, _
                                      ByVal sheetName As String)

    Dim startClearCol As Long, lastR As Long, lastC As Long
    Dim target As Range

    startClearCol = anchorCol + importedCols
    If startClearCol < anchorCol Then Exit Sub

    GetUsedBoundsFromAnchor ws, anchorRow, anchorCol, lastR, lastC
    If lastR = 0 Or lastC = 0 Then Exit Sub
    If startClearCol > lastC Then Exit Sub

    Set target = ws.Range(ws.Cells(anchorRow, startClearCol), ws.Cells(lastR, lastC))

    If REMOVE_PICTURES_IN_CLEARED_AREAS Then DeletePicturesInRangeFast ws, target

    If preserveFormatting Then
        target.ClearContents
    Else
        target.Clear
    End If
End Sub

Public Sub GetUsedBoundsFromAnchor(ByVal ws As Worksheet, _
                                   ByVal anchorRow As Long, _
                                   ByVal anchorCol As Long, _
                                   ByRef lastR As Long, _
                                   ByRef lastC As Long)

    Dim ur As Range

    lastR = 0
    lastC = 0

    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0

    If ur Is Nothing Then Exit Sub

    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1

    If lastR < anchorRow Then lastR = anchorRow
    If lastC < anchorCol Then lastC = anchorCol
End Sub

Public Sub DeletePicturesInRangeFast(ByVal ws As Worksheet, ByVal rng As Range)
    Dim i As Long
    Dim shp As Shape
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim tr As Long, tc As Long

    On Error Resume Next

    r1 = rng.Row
    c1 = rng.Column
    r2 = rng.Row + rng.Rows.Count - 1
    c2 = rng.Column + rng.Columns.Count - 1

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)

        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            tr = shp.TopLeftCell.Row
            tc = shp.TopLeftCell.Column

            If tr >= r1 And tr <= r2 And tc >= c1 And tc <= c2 Then
                shp.Delete
            Else
                tr = shp.BottomRightCell.Row
                tc = shp.BottomRightCell.Column
                If tr >= r1 And tr <= r2 And tc >= c1 And tc <= c2 Then shp.Delete
            End If
        End If
    Next i

    On Error GoTo 0
End Sub
