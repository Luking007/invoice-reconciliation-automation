Attribute VB_Name = "mod_Main"
Option Explicit

'==========================================================
' mod_Main
' One entry point, correct order: Import -> Reconcile ->
' DataQuality -> Dashboard -> PivotStaging -> card styling,
' then reset to the company-wide view.
'==========================================================


Sub RunFullRefresh()

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ImportAll
    ReconcileInvoices
    CheckDataQuality
    RefreshDashboard
    BuildPivotStaging
    FormatDashboardCards

    wsDash.Range("B8").Value = "All Customers"
    UpdateChartTitles "All Customers"

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Refresh complete." & vbNewLine & vbNewLine & _
           "Total Outstanding AR: " & Format(wsDash.Range("B2").Value, "$#,##0.00") & vbNewLine & _
           "Overdue 90+ Days: " & Format(wsDash.Range("B3").Value, "0.0%") & vbNewLine & _
           "Last Refreshed: " & Format(wsDash.Range("B7").Value, "yyyy-mm-dd hh:mm"), _
           vbInformation, "Invoice Reconciliation System"
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RunFullRefresh stopped partway through:" & vbNewLine & Err.Description & _
           vbNewLine & vbNewLine & "Check which sheet looks incomplete, fix that step, then run RunFullRefresh again from the top.", vbCritical
End Sub

