Attribute VB_Name = "mod_Reconcile"
Option Explicit

'==========================================================
' mod_Reconcile
' Matches RawInvoices against RawPayments using Dictionary
' lookups instead of nested loops. Reads both sheets into
' memory once (arrays), does all comparison in memory, then
' writes results back in one bulk operation — this is what
' keeps 3,000+ rows fast instead of taking minutes.
'==========================================================


Sub ReconcileInvoices()

    Dim wsInv As Worksheet, wsPay As Worksheet, wsRec As Worksheet, wsLog As Worksheet
    Dim invData As Variant, payData As Variant
    Dim dictInvoiceIdx As Object, dictPaymentTotal As Object
    Dim lastInvRow As Long, lastPayRow As Long
    Dim i As Long, logRow As Long, idx As Long
    Dim invID As String, payAmount As Double
    Dim key As Variant
    Dim outputArr() As Variant
    Dim custName As String, region As String
    Dim invDate As Date, dueDate As Date, amount As Double, totalPaid As Double
    Dim status As String, daysOverdue As Long, bucket As String

    LoadConfig

    Set wsInv = ThisWorkbook.Sheets("RawInvoices")
    Set wsPay = ThisWorkbook.Sheets("RawPayments")
    Set wsRec = GetOrCreateSheet("Reconciled")
    Set wsLog = GetOrCreateSheet("MatchLog")

    ClearBelowHeader wsRec, 1
    wsRec.Range("A1:I1").Value = Array("Invoice ID", "Customer Name", "Invoice Date", "Due Date", _
                                        "Amount", "Total Paid", "Status", "Aging Bucket", "Region")

    ClearBelowHeader wsLog, 1
    wsLog.Range("A1:C1").Value = Array("Issue Type", "Invoice ID", "Details")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    lastInvRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    lastPayRow = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row
    invData = wsInv.Range("A2:F" & lastInvRow).Value
    payData = wsPay.Range("A2:E" & lastPayRow).Value

    Set dictInvoiceIdx = CreateObject("Scripting.Dictionary")
    Set dictPaymentTotal = CreateObject("Scripting.Dictionary")
    logRow = 2

    ' ---- Pass 1: unique invoices, flag duplicates ----
    For i = 1 To UBound(invData, 1)
        invID = Trim(invData(i, 1))
        If invID <> "" Then
            If dictInvoiceIdx.Exists(invID) Then
                wsLog.Cells(logRow, 1).Value = "Duplicate Invoice ID"
                wsLog.Cells(logRow, 2).Value = invID
                wsLog.Cells(logRow, 3).Value = "Second entry for this ID ignored in Reconciled"
                logRow = logRow + 1
            Else
                dictInvoiceIdx.Add invID, i
            End If
        End If
    Next i

    ' ---- Pass 2: aggregate payments, flag orphans ----
    For i = 1 To UBound(payData, 1)
        invID = Trim(payData(i, 2))
        payAmount = payData(i, 4)
        If Not dictInvoiceIdx.Exists(invID) Then
            wsLog.Cells(logRow, 1).Value = "Orphaned Payment"
            wsLog.Cells(logRow, 2).Value = invID
            wsLog.Cells(logRow, 3).Value = "Payment of " & CurrencySymbol & Format(payAmount, "#,##0.00") & " matches no invoice"
            logRow = logRow + 1
        Else
            If dictPaymentTotal.Exists(invID) Then
                dictPaymentTotal(invID) = dictPaymentTotal(invID) + payAmount
            Else
                dictPaymentTotal.Add invID, payAmount
            End If
        End If
    Next i

    ' ---- Pass 3: reconcile each unique invoice, build output in memory ----
    ReDim outputArr(1 To dictInvoiceIdx.Count, 1 To 9)
    idx = 0
    For Each key In dictInvoiceIdx.Keys
        i = dictInvoiceIdx(key)
        idx = idx + 1

        custName = invData(i, 2)
        invDate = ParseFlexibleDate(invData(i, 3))
        dueDate = ParseFlexibleDate(invData(i, 4))
        amount = invData(i, 5)
        region = invData(i, 6)

        If dictPaymentTotal.Exists(key) Then
            totalPaid = dictPaymentTotal(key)
        Else
            totalPaid = 0
        End If
        
        If totalPaid <= 0.005 Then
            status = "Unpaid"
        ElseIf Abs(totalPaid - amount) <= 0.01 Then
            status = "Paid"
        ElseIf totalPaid < amount Then
            status = "Partial"
        Else
            status = "Overpaid"
        End If

        If status = "Unpaid" Or status = "Partial" Then
            daysOverdue = ReportingAsOfDate - dueDate
            If daysOverdue < 0 Then
                bucket = "Not Yet Due"
            ElseIf daysOverdue <= AgingBucket1 Then
                bucket = "0-" & AgingBucket1
            ElseIf daysOverdue <= AgingBucket2 Then
                bucket = (AgingBucket1 + 1) & "-" & AgingBucket2
            ElseIf daysOverdue <= AgingBucket3 Then
                bucket = (AgingBucket2 + 1) & "-" & AgingBucket3
            Else
                bucket = AgingBucket3 & "+"
            End If
        Else
            bucket = "N/A"
        End If

        outputArr(idx, 1) = key
        outputArr(idx, 2) = custName
        outputArr(idx, 3) = invDate
        outputArr(idx, 4) = dueDate
        outputArr(idx, 5) = amount
        outputArr(idx, 6) = totalPaid
        outputArr(idx, 7) = status
        outputArr(idx, 8) = bucket
        outputArr(idx, 9) = region
    Next key

    wsRec.Range(wsRec.Cells(2, 1), wsRec.Cells(1 + idx, 9)).Value = outputArr
    wsRec.Columns("C:D").NumberFormat = "yyyy-mm-dd"
    wsRec.Columns("E:F").NumberFormat = "#,##0.00"

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "ReconcileInvoices: " & idx & " unique invoices reconciled, " & (logRow - 2) & " issues logged to MatchLog."
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "ReconcileInvoices failed: " & Err.Description, vbCritical
End Sub

