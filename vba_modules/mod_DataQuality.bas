Attribute VB_Name = "mod_DataQuality"
Option Explicit

'==========================================================
' mod_DataQuality
' Scans RawInvoices and RawPayments for problems with the
' data itself -- blanks, stray whitespace, unreadable dates,
' negative amounts. Separate from MatchLog on purpose: this
' checks the data's own health, not how the two files relate.
'==========================================================


Sub CheckDataQuality()

    Dim wsInv As Worksheet, wsPay As Worksheet, wsLog As Worksheet
    Dim lastInvRow As Long, lastPayRow As Long
    Dim i As Long, logRow As Long
    Dim invID As String, custName As String, region As String
    Dim invDateRaw As String, dueDateRaw As String, payDateRaw As String
    Dim amount As Double, payAmount As Double

    Set wsInv = ThisWorkbook.Sheets("RawInvoices")
    Set wsPay = ThisWorkbook.Sheets("RawPayments")
    Set wsLog = GetOrCreateSheet("DataQualityLog")

    ClearBelowHeader wsLog, 1
    wsLog.Range("A1:C1").Value = Array("Issue Type", "Row Reference", "Details")

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    lastInvRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    lastPayRow = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row
    logRow = 2

    ' ---- RawInvoices ----
    For i = 2 To lastInvRow
        invID = wsInv.Cells(i, 1).Value

        If Trim(invID) = "" Then
            wsLog.Cells(logRow, 1).Value = "Blank Invoice ID"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "Invoice ID is empty"
            logRow = logRow + 1
        ElseIf invID <> Trim(invID) Then
            wsLog.Cells(logRow, 1).Value = "Whitespace in Invoice ID"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "'" & invID & "' has leading/trailing spaces (cleaned automatically during reconciliation)"
            logRow = logRow + 1
        End If

        custName = wsInv.Cells(i, 2).Value
        If Trim(custName) = "" Then
            wsLog.Cells(logRow, 1).Value = "Blank Customer Name"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "Invoice " & invID & " has no customer name"
            logRow = logRow + 1
        End If

        region = wsInv.Cells(i, 6).Value
        If Trim(region) = "" Then
            wsLog.Cells(logRow, 1).Value = "Blank Region"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "Invoice " & invID & " has no region"
            logRow = logRow + 1
        End If

        invDateRaw = wsInv.Cells(i, 3).Value
        If ParseFlexibleDate(invDateRaw) = 0 Then
            wsLog.Cells(logRow, 1).Value = "Unparseable Invoice Date"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "'" & invDateRaw & "' could not be read as a date"
            logRow = logRow + 1
        End If

        dueDateRaw = wsInv.Cells(i, 4).Value
        If ParseFlexibleDate(dueDateRaw) = 0 Then
            wsLog.Cells(logRow, 1).Value = "Unparseable Due Date"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "'" & dueDateRaw & "' could not be read as a date"
            logRow = logRow + 1
        End If

        amount = wsInv.Cells(i, 5).Value
        If amount < 0 Then
            wsLog.Cells(logRow, 1).Value = "Negative Amount"
            wsLog.Cells(logRow, 2).Value = "RawInvoices row " & i
            wsLog.Cells(logRow, 3).Value = "Invoice " & invID & " has a negative amount: " & amount
            logRow = logRow + 1
        End If
    Next i

    ' ---- RawPayments ----
    For i = 2 To lastPayRow
        payDateRaw = wsPay.Cells(i, 3).Value
        If ParseFlexibleDate(payDateRaw) = 0 Then
            wsLog.Cells(logRow, 1).Value = "Unparseable Payment Date"
            wsLog.Cells(logRow, 2).Value = "RawPayments row " & i
            wsLog.Cells(logRow, 3).Value = "'" & payDateRaw & "' could not be read as a date"
            logRow = logRow + 1
        End If

        payAmount = wsPay.Cells(i, 4).Value
        If payAmount < 0 Then
            wsLog.Cells(logRow, 1).Value = "Negative Payment Amount"
            wsLog.Cells(logRow, 2).Value = "RawPayments row " & i
            wsLog.Cells(logRow, 3).Value = "Payment has a negative amount: " & payAmount
            logRow = logRow + 1
        End If
    Next i

    Application.ScreenUpdating = True
    Debug.Print "CheckDataQuality: " & (logRow - 2) & " issues logged to DataQualityLog."
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "CheckDataQuality failed: " & Err.Description, vbCritical
End Sub
