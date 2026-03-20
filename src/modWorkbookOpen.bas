Attribute VB_Name = "modWorkbookOpen"
Option Explicit

Public Function OpenReportWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    Dim head As String
    Dim oldAlerts As Boolean
    Dim tempPath As String
    Dim wbBefore As Long

    If FileLenSafe(fullPath) = 0 Then Exit Function

    On Error Resume Next
    Set wb = Workbooks.Open( _
        Filename:=fullPath, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        Local:=True, _
        CorruptLoad:=xlRepairFile)
    On Error GoTo 0

    If Not wb Is Nothing Then
        Set OpenReportWorkbook = wb
        Exit Function
    End If

    head = ReadFileHead(fullPath, 4096)

    If InStr(1, head, "<html", vbTextCompare) > 0 _
        Or InStr(1, head, "<table", vbTextCompare) > 0 _
        Or InStr(1, head, "<!doctype html", vbTextCompare) > 0 Then

        tempPath = CopyToTempWithExt(fullPath, ".html")
        If Len(tempPath) > 0 Then
            On Error Resume Next
            Set wb = Workbooks.Open(Filename:=tempPath, ReadOnly:=True)
            On Error GoTo 0

            If Not wb Is Nothing Then
                Set OpenReportWorkbook = wb
                Exit Function
            End If
        End If
    End If

    If Left$(head, 4) = "sep=" Or InStr(head, ",") > 0 Or InStr(head, vbTab) > 0 Then
        tempPath = CopyToTempWithExt(fullPath, ".csv")
        If Len(tempPath) > 0 Then
            wbBefore = Application.Workbooks.Count

            On Error Resume Next
            Application.Workbooks.OpenText _
                Filename:=tempPath, _
                DataType:=xlDelimited, _
                Comma:=True, _
                Tab:=True
            On Error GoTo 0

            If Application.Workbooks.Count = wbBefore + 1 Then
                Set wb = Application.Workbooks(Application.Workbooks.Count)
                If Not wb Is Nothing Then
                    Set OpenReportWorkbook = wb
                    Exit Function
                End If
            End If
        End If
    End If

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = True

    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True)
    On Error GoTo 0

    Application.DisplayAlerts = oldAlerts

    If Not wb Is Nothing Then Set OpenReportWorkbook = wb
End Function

Public Function FileLenSafe(ByVal p As String) As Long
    On Error Resume Next
    FileLenSafe = FileLen(p)
    If Err.Number <> 0 Then FileLenSafe = 0
    Err.Clear
    On Error GoTo 0
End Function

Public Function ReadFileHead(ByVal p As String, ByVal n As Long) As String
    Dim f As Integer
    Dim s As String

    On Error Resume Next
    f = FreeFile
    Open p For Input As #f
    s = Input$(n, #f)
    Close #f
    ReadFileHead = LCase$(s)
    On Error GoTo 0
End Function

Public Function CopyToTempWithExt(ByVal src As String, ByVal newExt As String) As String
    Dim fso As Object
    Dim tmpFolder As String
    Dim baseName As String
    Dim dst As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    tmpFolder = Environ$("TEMP")
    If Right$(tmpFolder, 1) <> "\" Then tmpFolder = tmpFolder & "\"

    baseName = fso.GetBaseName(src)
    dst = tmpFolder & baseName & "_" & Format(Now, "yymmdd_hhnnss") & newExt

    On Error Resume Next
    fso.CopyFile src, dst, True
    If Err.Number = 0 Then CopyToTempWithExt = dst
    Err.Clear
    On Error GoTo 0
End Function
