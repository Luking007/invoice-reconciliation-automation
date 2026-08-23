Attribute VB_Name = "mod_Config"
Option Explicit

'==========================================================
' mod_Config
' Loads every setting from the Config sheet into Public
' variables so any other module can use them directly,
' instead of re-reading the sheet every time.
'==========================================================

' ===== Aging thresholds (days) =====
Public AgingBucket1 As Integer
Public AgingBucket2 As Integer
Public AgingBucket3 As Integer

' ===== Currency & fiscal year =====
Public CurrencySymbol As String
Public FiscalYearStart As Date

' ===== File paths =====
Public RawDataFolderPath As String
Public InvoicesFileName As String
Public PaymentsFileName As String

' ===== Dashboard colors =====
Public ColorBackground As String
Public ReportingAsOfDate As Date
Public ColorRiskOverdue As String
Public ColorBorder As String


Sub LoadConfig()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    ' "ws" is just a nickname for the Config sheet for the rest
    ' of this Sub, so we don't have to type the full sheet
    ' reference eleven more times below.

    AgingBucket1 = ws.Range("B2").Value
    AgingBucket2 = ws.Range("B3").Value
    AgingBucket3 = ws.Range("B4").Value

    CurrencySymbol = ws.Range("B5").Value
    FiscalYearStart = ws.Range("B6").Value

    RawDataFolderPath = ws.Range("B7").Value
    InvoicesFileName = ws.Range("B8").Value
    PaymentsFileName = ws.Range("B9").Value

    ColorBackground = ws.Range("B10").Value
    ColorRiskOverdue = ws.Range("B11").Value
    ColorBorder = ws.Range("B12").Value
    ReportingAsOfDate = ws.Range("B14").Value

End Sub
