Attribute VB_Name = "modPathHelpers"
Option Explicit

Public Function PathExists(ByVal p As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(p)
    If Err.Number = 0 Then PathExists = ((attr And vbDirectory) = vbDirectory)
    Err.Clear
    On Error GoTo 0
End Function

Public Function EnsureTrailingSlash(ByVal p As String) As String
    If Len(p) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\"
    End If
End Function

Public Function EndsWithDateFolder(ByVal p As String) As Boolean
    Dim t As String
    Dim i As Long

    t = Trim$(p)
    If Right$(t, 1) = "\" Then t = Left$(t, Len(t) - 1)

    i = InStrRev(t, "\")
    If i = 0 Then Exit Function

    t = Mid$(t, i + 1)
    EndsWithDateFolder = (Len(t) = 6 And IsNumeric(t))
End Function

Public Function ParentPath(ByVal p As String) As String
    Dim t As String
    Dim i As Long

    t = p
    If Right$(t, 1) = "\" Then t = Left$(t, Len(t) - 1)

    i = InStrRev(t, "\")
    If i = 0 Then Exit Function

    ParentPath = Left$(t, i)
End Function

Public Function HasYyMmDdSubfolder(ByVal rootPath As String) As Boolean
    Dim f As String, full As String

    f = Dir$(EnsureTrailingSlash(rootPath) & "*", vbDirectory)

    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            full = EnsureTrailingSlash(rootPath) & f
            On Error Resume Next
            If (GetAttr(full) And vbDirectory) <> 0 Then
                If Len(f) = 6 And IsNumeric(f) Then
                    HasYyMmDdSubfolder = True
                    On Error GoTo 0
                    Exit Function
                End If
            End If
            On Error GoTo 0
        End If
        f = Dir$
    Loop
End Function

Public Function FindMostRecentYyMmDdSubfolder(ByVal rootPath As String) As String
    Dim f As String, best As String, bestName As String, full As String

    f = Dir$(EnsureTrailingSlash(rootPath) & "*", vbDirectory)

    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            full = EnsureTrailingSlash(rootPath) & f
            On Error Resume Next
            If (GetAttr(full) And vbDirectory) <> 0 Then
                If Len(f) = 6 And IsNumeric(f) Then
                    If f > bestName Then
                        bestName = f
                        best = full
                    End If
                End If
            End If
            On Error GoTo 0
        End If
        f = Dir$
    Loop

    FindMostRecentYyMmDdSubfolder = best
End Function

Public Function ResolveDailyFolder(ByVal baseRoot As String, ByVal dayToken As String) As String
    Dim dayFolder As String
    Dim recentFolder As String
    Dim pickedParent As String

    dayFolder = EnsureTrailingSlash(baseRoot) & dayToken & "\"

    If PathExists(dayFolder) Then
        ResolveDailyFolder = dayFolder
        Exit Function
    End If

    If PROMPT_IF_TODAY_MISSING Then
        pickedParent = PromptForValidParent()
        If Len(pickedParent) = 0 Then Exit Function

        pickedParent = EnsureTrailingSlash(pickedParent)
        SetOrUpdateDefinedName "MorningReportsRoot", pickedParent

        dayFolder = pickedParent & dayToken & "\"
        If PathExists(dayFolder) Then
            ResolveDailyFolder = dayFolder
            Exit Function
        End If

        recentFolder = FindMostRecentYyMmDdSubfolder(pickedParent)
        If Len(recentFolder) > 0 Then
            ResolveDailyFolder = EnsureTrailingSlash(recentFolder)
            Exit Function
        End If
    Else
        recentFolder = FindMostRecentYyMmDdSubfolder(baseRoot)
        If Len(recentFolder) > 0 Then
            ResolveDailyFolder = EnsureTrailingSlash(recentFolder)
            Exit Function
        End If
    End If
End Function

Public Function GetDefinedNameText(ByVal nm As String) As String
    On Error Resume Next
    Dim v As Variant

    v = ThisWorkbook.Names(nm).RefersTo
    If Err.Number = 0 Then
        If Left$(CStr(v), 2) = "=""" And Right$(CStr(v), 1) = """" Then
            GetDefinedNameText = Mid$(CStr(v), 3, Len(CStr(v)) - 3)
        Else
            GetDefinedNameText = CStr(Evaluate(v))
        End If
    End If

    Err.Clear
    On Error GoTo 0
End Function

Public Sub SetOrUpdateDefinedName(ByVal nm As String, ByVal valueText As String)
    On Error Resume Next
    Dim nmObj As Name

    Set nmObj = ThisWorkbook.Names(nm)

    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=nm, RefersTo:="=""" & valueText & """"
    Else
        nmObj.RefersTo = "=""" & valueText & """"
    End If

    On Error GoTo 0
End Sub

Public Function ChooseRootFolder() As String
    Dim fd As Object
    Dim picked As String

    On Error Resume Next
    Set fd = Application.FileDialog(4)
    On Error GoTo 0

    If Not fd Is Nothing Then
        With fd
            .Title = "Select the parent folder that contains daily yymmdd subfolders"
            If .Show = -1 Then
                picked = .SelectedItems(1)
                picked = EnsureTrailingSlash(picked)
                If PathExists(picked) Then
                    ChooseRootFolder = picked
                    Exit Function
                End If
            End If
        End With
    End If

    picked = InputBox( _
        "Paste the parent folder path that contains the daily yymmdd subfolders:", _
        "Select Parent Folder" _
    )

    If Len(picked) = 0 Then Exit Function

    picked = EnsureTrailingSlash(picked)
    If PathExists(picked) Then
        ChooseRootFolder = picked
    Else
        MsgBox "That path does not exist. Please choose a valid folder.", vbCritical
    End If
End Function

Public Function PromptForValidParent() As String
    Dim p As String

    Do
        p = ChooseRootFolder()
        If Len(p) = 0 Then Exit Function

        If HasYyMmDdSubfolder(p) Then
            PromptForValidParent = EnsureTrailingSlash(p)
            Exit Function
        Else
            MsgBox "That folder has no yymmdd subfolders. Please choose the parent folder.", vbExclamation
        End If
    Loop
End Function
