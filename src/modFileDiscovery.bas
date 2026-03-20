Attribute VB_Name = "modFileDiscovery"
Option Explicit

Public Function GetNewestFile_Robust_Dir(ByVal folderPath As String, _
                                         ByVal prefix As String, _
                                         ByVal searchDate As String) As String

    Dim pfxNorm As String, nmNorm As String
    Dim newest As String, newestDT As Date
    Dim f As String, cand() As String, k As Long, dt As Date

    pfxNorm = NormalizeName(prefix)

    If Len(searchDate) = 0 Then
        f = Dir$(folderPath & "*", vbNormal)
        Do While Len(f) > 0
            If ShouldConsiderFile(f) Then
                nmNorm = NormalizeName(f)
                If InStr(1, nmNorm, pfxNorm, vbTextCompare) > 0 Then
                    If SafeFileDateTime(folderPath & f, dt) Then
                        If dt > newestDT Then
                            newestDT = dt
                            newest = f
                        End If
                    End If
                End If
            End If
            f = Dir$
        Loop

        GetNewestFile_Robust_Dir = newest
        Exit Function
    End If

    cand = BuildDateCandidates(searchDate)

    f = Dir$(folderPath & "*", vbNormal)
    Do While Len(f) > 0
        If ShouldConsiderFile(f) Then
            nmNorm = NormalizeName(f)

            If InStr(1, nmNorm, pfxNorm, vbTextCompare) > 0 Then
                For k = LBound(cand) To UBound(cand)
                    If InStr(1, nmNorm, cand(k), vbTextCompare) > 0 Then
                        If SafeFileDateTime(folderPath & f, dt) Then
                            If dt > newestDT Then
                                newestDT = dt
                                newest = f
                            End If
                        End If
                        Exit For
                    End If
                Next k
            End If
        End If
        f = Dir$
    Loop

    GetNewestFile_Robust_Dir = newest
End Function

Private Function SafeFileDateTime(ByVal fullPath As String, ByRef outDT As Date) As Boolean
    On Error Resume Next
    outDT = FileDateTime(fullPath)
    SafeFileDateTime = (Err.Number = 0)
    Err.Clear
End Function

Private Function ShouldConsiderFile(ByVal fname As String) As Boolean
    Dim ext As String, dotPos As Long

    If Len(fname) = 0 Then Exit Function
    If Left$(fname, 2) = "~$" Then Exit Function

    dotPos = InStrRev(fname, ".")
    If dotPos = 0 Then Exit Function

    ext = LCase$(Mid$(fname, dotPos + 1))

    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls", "csv", "txt", "html", "htm"
            ShouldConsiderFile = True
    End Select
End Function

Private Function NormalizeName(ByVal s As String) As String
    s = LCase$(s)
    s = Replace$(s, "_", " ")
    s = Replace$(s, "-", " ")

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    NormalizeName = Trim$(s)
End Function

Private Function BuildDateCandidates(ByVal searchDate As String) As String()
    Dim out() As String, n As Long
    Dim dt As Date

    ReDim out(0 To 0)
    n = -1

    AddTok out, n, NormalizeDateToken(searchDate)

    On Error Resume Next
    dt = CDate(searchDate)

    If Err.Number = 0 Then
        AddTok out, n, NormalizeDateToken(Format(dt, "mm-dd-yyyy"))
        AddTok out, n, NormalizeDateToken(Format(dt, "yyyy-mm-dd"))
        AddTok out, n, NormalizeDateToken(Format(dt, "yyyymmdd"))
    End If

    BuildDateCandidates = out
End Function

Private Sub AddTok(ByRef arr() As String, ByRef n As Long, ByVal tok As String)
    If Len(tok) = 0 Then Exit Sub

    n = n + 1
    ReDim Preserve arr(0 To n)
    arr(n) = tok
End Sub

Private Function NormalizeDateToken(ByVal s As String) As String
    s = LCase$(s)
    s = Replace$(s, "-", " ")
    NormalizeDateToken = Trim$(s)
End Function
