Attribute VB_Name = "modConfig"
Option Explicit

Public Const DEBUG_MATCH As Boolean = True
Public Const PROMPT_IF_TODAY_MISSING As Boolean = True
Public Const TREASURY_CREATE_IF_MISSING As Boolean = False
Public Const REMOVE_PICTURES_IN_CLEARED_AREAS As Boolean = True
Public Const TREASURY_DATE_FMT As String = "mm.dd.yy"

Public Function PreserveFormatSheets() As Variant
    PreserveFormatSheets = Array("Enfusion RTG", "Enfusion RTG Summary")
End Function

Public Function BuildReportList() As Variant
    BuildReportList = Array( _
        Array("Report Position Summary ", "Broker Position Report", 1, 1, "none", True), _
        Array("Report Portfolio Margin Detail ", "Broker Margin Report", 3, 3, "none", True), _
        Array("Report Debit-Credit Interest Accrual MTD ", "Broker Interest Report", 1, 1, "none", True), _
        Array("Report Rebate Detail ", "Broker Rebate Report", 3, 3, "none", True), _
        Array("IMS RTG COB", "IMS Real Time Grid", 1, 1, "cob", True) _
    )
End Function
