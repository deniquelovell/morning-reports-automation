Attribute VB_Name = "modMain"
Option Explicit

Public Sub ImportMorningReports_OneDrive_Safe()

    Dim baseRoot As String, savedRoot As String
    Dim dayToken As String, dayFolder As String
    Dim todayDate As String, cobDate As String
    Dim reportList As Variant, i As Long
    Dim searchDate As String, fileName As String, fullPath As String
    Dim srcWB As Workbook, srcWS As Worksheet, destWS As Worksheet
    Dim pasteRows As Long, pasteCols As Long
    Dim isUrl As Boolean
    Dim calcState As XlCalculation, scrState As Boolean, evtState As Boolean, dispAlerts As Boolean
    Dim prefix As String, dstName As String, dateKind As String
    Dim dstRow As Long, dstCol As Long
    Dim preserveFmt As Boolean
    Dim isRequired As Boolean

    calcState = Application.Calculation
    scrState = Application.ScreenUpdating
    evtState = Application.EnableEvents
    dispAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CLEAN_FAIL

    ClearAllWorkbookFilters ThisWorkbook

    baseRoot = ThisWorkbook.Path
    savedRoot = GetDefinedNameText("MorningReportsRoot")
    If Len(savedRoot) > 0 Then baseRoot = CStr(savedRoot)

    If EndsWithDateFolder(baseRoot) Then baseRoot = ParentPath(baseRoot)

    isUrl = (InStr(1, baseRoot, "://", vbTextCompare) > 0 Or Left$(LCase$(baseRoot), 4) = "http")

    If isUrl Or Not PathExists(baseRoot) Or Not HasYyMmDdSubfolder(baseRoot) Then
        baseRoot = PromptForValidParent()
        If Len(baseRoot) = 0 Then
            MsgBox "No folder selected. Please sync OneDrive and try again.", vbCritical
            GoTo CLEAN_EXIT
        End If
        SetOrUpdateDefinedName "MorningReportsRoot", baseRoot
    End If

    baseRoot = EnsureTrailingSlash(baseRoot)

    dayToken = Format(Date, "yymmdd")
    todayDate = Format(Date, "m-d-yyyy")
    cobDate = Format(LastBusinessDay(Date - 1), "yymmdd")

    dayFolder = baseRoot & dayToken & "\"

    If Not PathExists(dayFolder) Then
        dayFolder = ResolveDailyFolder(baseRoot, dayToken)
        If Len(dayFolder) = 0 Then
            MsgBox "No valid daily report folder found.", vbCritical
            GoTo CLEAN_EXIT
        End If
    End If

    reportList = BuildReportList()

    For i = LBound(reportList) To UBound(reportList)

        prefix = CStr(reportList(i)(0))
        dstName = CStr(reportList(i)(1))
        dstRow = CLng(reportList(i)(2))
        dstCol = CLng(reportList(i)(3))
        dateKind = CStr(reportList(i)(4))
        isRequired = CBool(reportList(i)(5))

        searchDate = IIf(dateKind = "today", todayDate, IIf(dateKind = "cob", cobDate, ""))

        fileName = GetNewestFile_Robust_Dir(dayFolder, prefix, searchDate)

        If Len(fileName) = 0 Then
            If isRequired Then
                MsgBox "Missing required report: " & prefix, vbCritical
                GoTo CLEAN_EXIT
            Else
                GoTo NEXT_REPORT
            End If
        End If

        fullPath = dayFolder & fileName
        Set srcWB = OpenReportWorkbook(fullPath)

        If srcWB Is Nothing Then
            If isRequired Then
                MsgBox "Could not open report: " & fullPath, vbCritical
                GoTo CLEAN_EXIT
            Else
                GoTo NEXT_REPORT
            End If
        End If

        If Not SheetExists(dstName, ThisWorkbook) Then
            srcWB.Close False
            Set srcWB = Nothing
            MsgBox "Destination sheet not found: " & dstName, vbCritical
            GoTo CLEAN_EXIT
        End If

        Set destWS = ThisWorkbook.Sheets(dstName)
        preserveFmt = IsInArray(dstName, PreserveFormatSheets())

        Set srcWS = PickSourceWorksheet(srcWB, dstName)

        If Not srcWS Is Nothing Then
            pasteRows = LastUsedRow(srcWS)
            pasteCols = LastUsedCol(srcWS)

            SafeClearFromAnchorFast destWS, dstRow, dstCol, preserveFmt, dstName
            PasteRangeFast srcWS, pasteRows, pasteCols, destWS, dstRow, dstCol
            SafeClearBelowAnchorFast destWS, dstRow, dstCol, pasteRows, preserveFmt, dstName
            SafeClearRightOfAnchorFast destWS, dstRow, dstCol, pasteCols, preserveFmt, dstName
        End If

        srcWB.Close False
        Set srcWB = Nothing
        Set srcWS = Nothing

NEXT_REPORT:
    Next i

    MsgBox "Morning reports imported successfully.", vbInformation

CLEAN_EXIT:
    Application.Calculation = calcState
    If calcState = xlCalculationAutomatic Then Application.Calculate
    Application.ScreenUpdating = scrState
    Application.EnableEvents = evtState
    Application.DisplayAlerts = dispAlerts
    Application.StatusBar = False
    Exit Sub

CLEAN_FAIL:
    Application.Calculation = calcState
    If calcState = xlCalculationAutomatic Then Application.Calculate
    Application.ScreenUpdating = scrState
    Application.EnableEvents = evtState
    Application.DisplayAlerts = dispAlerts
    Application.StatusBar = False
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
