Attribute VB_Name = "mod_Dashboard"
Option Explicit

'==========================================================
' mod_Dashboard (Part 1: calculations)
' Computes the 5 headline KPIs from Reconciled + RawPayments
' and writes raw values to the Dashboard sheet. Visual layout
' -- cards, charts, colors -- comes in the next step.
'==========================================================


Sub RefreshDashboard()

    Dim wsRec As Worksheet, wsPay As Worksheet, wsDash As Worksheet
    Dim lastRecRow As Long, lastPayRow As Long
    Dim i As Long
    Dim invID As String
    Dim amount As Double, totalPaid As Double, outstanding As Double
    Dim bucket As String, status As String
    Dim invDate As Date, payDate As Date, payAmount As Double

    Dim totalOutstanding As Double, outstanding90Plus As Double
    Dim overdueCount As Long
    Dim daysToPaySum As Double, paidCount As Long
    Dim collectedThisMonth As Double
    Dim pctOverdue90 As Double, avgDaysToPay As Double
    Dim dictLastPayDate As Object, dictValidIDs As Object

    LoadConfig

    Set wsRec = ThisWorkbook.Sheets("Reconciled")
    Set wsPay = ThisWorkbook.Sheets("RawPayments")
    Set wsDash = GetOrCreateSheet("Dashboard")

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    ' Real Invoice IDs first -- so an orphaned payment can never be
    ' counted as money collected against a real invoice.
    Set dictValidIDs = CreateObject("Scripting.Dictionary")
    lastRecRow = wsRec.Cells(wsRec.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRecRow
        dictValidIDs(Trim(wsRec.Cells(i, 1).Value)) = True
    Next i

    Set dictLastPayDate = CreateObject("Scripting.Dictionary")
    lastPayRow = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastPayRow
        invID = Trim(wsPay.Cells(i, 2).Value)
        If dictValidIDs.Exists(invID) Then
            payDate = ParseFlexibleDate(wsPay.Cells(i, 3).Value)
            payAmount = wsPay.Cells(i, 4).Value

            If payDate <> 0 Then
                If dictLastPayDate.Exists(invID) Then
                    If payDate > dictLastPayDate(invID) Then dictLastPayDate(invID) = payDate
                Else
                    dictLastPayDate.Add invID, payDate
                End If

                If Year(payDate) = Year(ReportingAsOfDate) And Month(payDate) = Month(ReportingAsOfDate) Then
                    collectedThisMonth = collectedThisMonth + payAmount
                End If
            End If
        End If
    Next i

    For i = 2 To lastRecRow
        invID = wsRec.Cells(i, 1).Value
        invDate = wsRec.Cells(i, 3).Value
        amount = wsRec.Cells(i, 5).Value
        totalPaid = wsRec.Cells(i, 6).Value
        status = wsRec.Cells(i, 7).Value
        bucket = wsRec.Cells(i, 8).Value

        outstanding = amount - totalPaid
        If outstanding < 0 Then outstanding = 0
        totalOutstanding = totalOutstanding + outstanding

        If bucket = "90+" Then outstanding90Plus = outstanding90Plus + outstanding
        If bucket <> "N/A" And bucket <> "Not Yet Due" Then overdueCount = overdueCount + 1

        If status = "Paid" And dictLastPayDate.Exists(invID) Then
            daysToPaySum = daysToPaySum + (dictLastPayDate(invID) - invDate)
            paidCount = paidCount + 1
        End If
    Next i

    If totalOutstanding > 0 Then pctOverdue90 = outstanding90Plus / totalOutstanding Else pctOverdue90 = 0
    If paidCount > 0 Then avgDaysToPay = daysToPaySum / paidCount Else avgDaysToPay = 0

    wsDash.Range("A1").Value = "Invoice Reconciliation Dashboard"
    wsDash.Range("A2").Value = "Total Outstanding AR": wsDash.Range("B2").Value = totalOutstanding
    wsDash.Range("A3").Value = "% Overdue 90+ Days": wsDash.Range("B3").Value = pctOverdue90
    wsDash.Range("A4").Value = "Avg Days to Pay": wsDash.Range("B4").Value = Round(avgDaysToPay, 1)
    wsDash.Range("A5").Value = "Collected This Month": wsDash.Range("B5").Value = collectedThisMonth
    wsDash.Range("A6").Value = "# Overdue Invoices": wsDash.Range("B6").Value = overdueCount
    wsDash.Range("A7").Value = "Last Refreshed": wsDash.Range("B7").Value = Now

    wsDash.Range("B2").NumberFormat = "#,##0.00"
    wsDash.Range("B3").NumberFormat = "0.0%"
    wsDash.Range("B5").NumberFormat = "#,##0.00"
    wsDash.Range("B7").NumberFormat = "yyyy-mm-dd hh:mm"
    
    Application.ScreenUpdating = True
    Debug.Print "RefreshDashboard: Outstanding=" & Format(totalOutstanding, "#,##0.00") & _
                ", %90+=" & Format(pctOverdue90, "0.0%") & _
                ", AvgDaysToPay=" & Round(avgDaysToPay, 1) & _
                ", CollectedThisMonth=" & Format(collectedThisMonth, "#,##0.00") & _
                ", OverdueCount=" & overdueCount
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "RefreshDashboard failed: " & Err.Description, vbCritical
End Sub

Sub BuildPivotStaging(Optional ByVal filterCustomer As String = "")

    Dim wsRec As Worksheet, wsPay As Worksheet, wsStg As Worksheet, wsDash As Worksheet
    Dim lastRecRow As Long, lastPayRow As Long
    Dim i As Long, a As Long, b As Long
    Dim key As Variant

    Dim dictBucketCount As Object, dictBucketOutstanding As Object
    Dim dictCustOutstanding As Object
    Dim dictStatusCount As Object, dictStatusAmount As Object
    Dim dictInvoicedByMonth As Object, dictCollectedByMonth As Object, allMonths As Object
    Dim dictValidIDs As Object, dictInvoiceCustomer As Object

    Dim custName As String, status As String, bucket As String, monthKey As String, invID As String
    Dim amount As Double, totalPaid As Double, outstanding As Double
    Dim invDate As Date, payDate As Date, payAmount As Double
    Dim tempS As String, tempD As Double
    Dim isFiltered As Boolean, includeRow As Boolean

    isFiltered = (filterCustomer <> "" And filterCustomer <> "All Customers")

    Set wsRec = ThisWorkbook.Sheets("Reconciled")
    Set wsPay = ThisWorkbook.Sheets("RawPayments")
    Set wsStg = GetOrCreateSheet("PivotStaging")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    wsStg.Cells.Clear

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Set dictBucketCount = CreateObject("Scripting.Dictionary")
    Set dictBucketOutstanding = CreateObject("Scripting.Dictionary")
    Set dictCustOutstanding = CreateObject("Scripting.Dictionary")
    Set dictStatusCount = CreateObject("Scripting.Dictionary")
    Set dictStatusAmount = CreateObject("Scripting.Dictionary")
    Set dictInvoicedByMonth = CreateObject("Scripting.Dictionary")
    Set dictCollectedByMonth = CreateObject("Scripting.Dictionary")
    Set dictValidIDs = CreateObject("Scripting.Dictionary")
    Set dictInvoiceCustomer = CreateObject("Scripting.Dictionary")

    lastRecRow = wsRec.Cells(wsRec.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRecRow
        dictValidIDs(Trim(wsRec.Cells(i, 1).Value)) = True
        dictInvoiceCustomer(Trim(wsRec.Cells(i, 1).Value)) = wsRec.Cells(i, 2).Value
    Next i

    For i = 2 To lastRecRow
        custName = wsRec.Cells(i, 2).Value
        invDate = wsRec.Cells(i, 3).Value
        amount = wsRec.Cells(i, 5).Value
        totalPaid = wsRec.Cells(i, 6).Value
        status = wsRec.Cells(i, 7).Value
        bucket = wsRec.Cells(i, 8).Value

        outstanding = amount - totalPaid
        If outstanding < 0 Then outstanding = 0

        ' Top 10 always uses every row, filter or not
        If Trim(custName) <> "" Then
            If dictCustOutstanding.Exists(custName) Then
                dictCustOutstanding(custName) = dictCustOutstanding(custName) + outstanding
            Else
                dictCustOutstanding.Add custName, outstanding
            End If
        End If

        includeRow = (Not isFiltered) Or (Trim(custName) = filterCustomer)
        If includeRow Then
            If bucket <> "N/A" Then
                If dictBucketCount.Exists(bucket) Then
                    dictBucketCount(bucket) = dictBucketCount(bucket) + 1
                    dictBucketOutstanding(bucket) = dictBucketOutstanding(bucket) + outstanding
                Else
                    dictBucketCount.Add bucket, 1
                    dictBucketOutstanding.Add bucket, outstanding
                End If
            End If

            If dictStatusCount.Exists(status) Then
                dictStatusCount(status) = dictStatusCount(status) + 1
                dictStatusAmount(status) = dictStatusAmount(status) + amount
            Else
                dictStatusCount.Add status, 1
                dictStatusAmount.Add status, amount
            End If

            monthKey = Format(invDate, "yyyy-mm")
            If dictInvoicedByMonth.Exists(monthKey) Then
                dictInvoicedByMonth(monthKey) = dictInvoicedByMonth(monthKey) + amount
            Else
                dictInvoicedByMonth.Add monthKey, amount
            End If
        End If
    Next i
    
    lastPayRow = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastPayRow
        invID = Trim(wsPay.Cells(i, 2).Value)
        If dictValidIDs.Exists(invID) Then
            includeRow = (Not isFiltered)
            If isFiltered Then
                If dictInvoiceCustomer.Exists(invID) Then
                    includeRow = (dictInvoiceCustomer(invID) = filterCustomer)
                End If
            End If
            If includeRow Then
                payDate = ParseFlexibleDate(wsPay.Cells(i, 3).Value)
                payAmount = wsPay.Cells(i, 4).Value
                If payDate <> 0 Then
                    monthKey = Format(payDate, "yyyy-mm")
                    If dictCollectedByMonth.Exists(monthKey) Then
                        dictCollectedByMonth(monthKey) = dictCollectedByMonth(monthKey) + payAmount
                    Else
                        dictCollectedByMonth.Add monthKey, payAmount
                    End If
                End If
            End If
        End If
    Next i

    Dim bucketOrder As Variant
    bucketOrder = Array("Not Yet Due", "0-30", "31-60", "61-90", "90+")
    wsStg.Range("A1:C1").Value = Array("Aging Bucket", "Count", "Outstanding")
    For i = 0 To UBound(bucketOrder)
        wsStg.Cells(i + 2, 1).Value = bucketOrder(i)
        If dictBucketCount.Exists(bucketOrder(i)) Then
            wsStg.Cells(i + 2, 2).Value = dictBucketCount(bucketOrder(i))
            wsStg.Cells(i + 2, 3).Value = dictBucketOutstanding(bucketOrder(i))
        Else
            wsStg.Cells(i + 2, 2).Value = 0: wsStg.Cells(i + 2, 3).Value = 0
        End If
    Next i
    wsStg.Range("C2:C6").NumberFormat = "#,##0.00"

    Set allMonths = CreateObject("Scripting.Dictionary")
    For Each key In dictInvoicedByMonth.Keys
        If Not allMonths.Exists(key) Then allMonths.Add key, True
    Next key
    For Each key In dictCollectedByMonth.Keys
        If Not allMonths.Exists(key) Then allMonths.Add key, True
    Next key

    Dim monthList() As String
    If allMonths.Count = 0 Then
        ReDim monthList(0 To 0): monthList(0) = "(no data)"
    Else
        ReDim monthList(0 To allMonths.Count - 1)
        i = 0
        For Each key In allMonths.Keys
            monthList(i) = key: i = i + 1
        Next key
        For a = 0 To UBound(monthList) - 1
            For b = a + 1 To UBound(monthList)
                If monthList(b) < monthList(a) Then
                    tempS = monthList(a): monthList(a) = monthList(b): monthList(b) = tempS
                End If
            Next b
        Next a
    End If

    wsStg.Range("E1:G1").Value = Array("Month", "Invoiced", "Collected")
    For i = 0 To UBound(monthList)
        wsStg.Cells(i + 2, 5).Value = monthList(i)
        If dictInvoicedByMonth.Exists(monthList(i)) Then wsStg.Cells(i + 2, 6).Value = dictInvoicedByMonth(monthList(i)) Else wsStg.Cells(i + 2, 6).Value = 0
        If dictCollectedByMonth.Exists(monthList(i)) Then wsStg.Cells(i + 2, 7).Value = dictCollectedByMonth(monthList(i)) Else wsStg.Cells(i + 2, 7).Value = 0
    Next i
    wsStg.Range("F2:G" & (UBound(monthList) + 2)).NumberFormat = "#,##0.00"

    Dim custKeys() As String, custVals() As Double, nCust As Long
    nCust = dictCustOutstanding.Count
    ReDim custKeys(0 To nCust - 1): ReDim custVals(0 To nCust - 1)
    i = 0
    For Each key In dictCustOutstanding.Keys
        custKeys(i) = key: custVals(i) = dictCustOutstanding(key): i = i + 1
    Next key
    For a = 0 To nCust - 2
        For b = a + 1 To nCust - 1
            If custVals(b) > custVals(a) Then
                tempD = custVals(a): custVals(a) = custVals(b): custVals(b) = tempD
                tempS = custKeys(a): custKeys(a) = custKeys(b): custKeys(b) = tempS
            End If
        Next b
    Next a

    wsStg.Range("I1:J1").Value = Array("Customer", "Outstanding")
    For i = 0 To 9
        wsStg.Cells(i + 2, 9).Value = custKeys(i)
        wsStg.Cells(i + 2, 10).Value = custVals(i)
    Next i
    wsStg.Range("J2:J11").NumberFormat = "#,##0.00"
    
    Dim statusOrder As Variant
    statusOrder = Array("Paid", "Partial", "Unpaid", "Overpaid")
    wsStg.Range("L1:N1").Value = Array("Status", "Count", "Amount")
    For i = 0 To UBound(statusOrder)
        wsStg.Cells(i + 2, 12).Value = statusOrder(i)
        If dictStatusCount.Exists(statusOrder(i)) Then
            wsStg.Cells(i + 2, 13).Value = dictStatusCount(statusOrder(i))
            wsStg.Cells(i + 2, 14).Value = dictStatusAmount(statusOrder(i))
        Else
            wsStg.Cells(i + 2, 13).Value = 0: wsStg.Cells(i + 2, 14).Value = 0
        End If
    Next i
    wsStg.Range("N2:N5").NumberFormat = "#,##0.00"

    ' Customer list for the filter dropdown -- alphabetical, separate copy
    wsStg.Range("P1").Value = "Customer List"
    wsStg.Range("P2").Value = "All Customers"
    Dim allCustKeys() As String
    ReDim allCustKeys(0 To nCust - 1)
    For i = 0 To nCust - 1
        allCustKeys(i) = custKeys(i)
    Next i
    For a = 0 To nCust - 2
        For b = a + 1 To nCust - 1
            If allCustKeys(b) < allCustKeys(a) Then
                tempS = allCustKeys(a): allCustKeys(a) = allCustKeys(b): allCustKeys(b) = tempS
            End If
        Next b
    Next a
    For i = 0 To nCust - 1
        wsStg.Cells(i + 3, 16).Value = allCustKeys(i)
    Next i

    wsDash.Range("A8").Value = "Filter by Customer:"
    If Trim(wsDash.Range("B8").Value) = "" Then wsDash.Range("B8").Value = "All Customers"
    With wsDash.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PivotStaging!$P$2:$P$" & (nCust + 2)
    End With

    Application.ScreenUpdating = True

    If isFiltered Then
        Dim filteredSum As Double
        filteredSum = 0
        For Each key In dictBucketOutstanding.Keys
            filteredSum = filteredSum + dictBucketOutstanding(key)
        Next key
        Debug.Print "BuildPivotStaging: filtered to '" & filterCustomer & "'."
        Debug.Print "  Cross-check, sum of bucket outstanding: " & Format(filteredSum, "#,##0.00")
    Else
        Debug.Print "BuildPivotStaging: 5 aging buckets, " & (UBound(monthList) + 1) & " months, top 10 customers, 4 statuses, " & nCust & " customers in filter list."
        Debug.Print "  Cross-check, sum of bucket outstanding: " & Format(dictBucketOutstanding("Not Yet Due") + dictBucketOutstanding("0-30") + dictBucketOutstanding("31-60") + dictBucketOutstanding("61-90") + dictBucketOutstanding("90+"), "#,##0.00")
    End If
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "BuildPivotStaging failed: " & Err.Description, vbCritical
End Sub

Sub UpdateChartTitles(ByVal selectedCustomer As String)
    Dim wsDash As Worksheet, co As ChartObject
    Dim baseTitle As String, suffix As String, currentTitle As String
    Dim dashPos As Long

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    If selectedCustomer = "" Or selectedCustomer = "All Customers" Then
        suffix = ""
    Else
        suffix = " -- " & selectedCustomer
    End If

    For Each co In wsDash.ChartObjects
        If co.Chart.HasTitle Then
            currentTitle = co.Chart.ChartTitle.Text
            dashPos = InStr(currentTitle, " -- ")
            If dashPos > 0 Then baseTitle = Left(currentTitle, dashPos - 1) Else baseTitle = currentTitle

            If InStr(baseTitle, "Top 10 Customers") = 0 Then
                co.Chart.ChartTitle.Text = baseTitle & suffix
            End If
        End If
    Next co
End Sub

Sub FormatDashboardCards()

    Dim wsDash As Worksheet
    Dim cardStarts As Variant, labels As Variant
    Dim i As Long, colStart As Long
    Dim lblCell As Range, valCell As Range

    LoadConfig
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    cardStarts = Array(1, 5, 9, 13, 17)   ' A, E, I, M, Q
    labels = Array("TOTAL OUTSTANDING AR", "% OVERDUE 90+ DAYS", "AVG DAYS TO PAY", "COLLECTED THIS MONTH", "# OVERDUE INVOICES")

    For i = 0 To 4
        colStart = cardStarts(i)
        Set lblCell = wsDash.Range(wsDash.Cells(10, colStart), wsDash.Cells(10, colStart + 2))
        Set valCell = wsDash.Range(wsDash.Cells(11, colStart), wsDash.Cells(13, colStart + 2))

        With lblCell
            .Merge
            .Value = labels(i)
            .Interior.Color = HexToRGB(ColorBorder)
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 9: .Font.Bold = True: .Font.Name = "Segoe UI"
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With

        With valCell
            .Merge
            .Interior.Color = HexToRGB(ColorBackground)
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 26: .Font.Bold = True: .Font.Name = "Segoe UI"
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With

        wsDash.Range(wsDash.Cells(10, colStart), wsDash.Cells(13, colStart + 2)).BorderAround _
            Weight:=xlMedium, Color:=HexToRGB(ColorRiskOverdue)
    Next i

    wsDash.Cells(11, 1).Formula = "=""$""&TEXT(B2,""#,##0"")"
    wsDash.Cells(11, 5).Formula = "=TEXT(B3,""0.0%"")"
    wsDash.Cells(11, 9).Formula = "=TEXT(B4,""0.0"")&"" days"""
    wsDash.Cells(11, 13).Formula = "=""$""&TEXT(B5,""#,##0"")"
    wsDash.Cells(11, 17).Formula = "=B6"

    If wsDash.Range("B3").Value > 0.25 Then
        wsDash.Range(wsDash.Cells(11, 5), wsDash.Cells(13, 7)).Interior.Color = HexToRGB(ColorRiskOverdue)
    End If

    wsDash.Cells(15, 1).Formula = "=""Last refreshed: ""&TEXT(B7,""yyyy-mm-dd hh:mm"")"
    wsDash.Cells(15, 1).Font.Size = 9: wsDash.Cells(15, 1).Font.Italic = True
    wsDash.Cells(15, 1).Font.Color = RGB(120, 120, 120)

    wsDash.Rows(10).RowHeight = 22
    wsDash.Rows(11).RowHeight = 20: wsDash.Rows(12).RowHeight = 20: wsDash.Rows(13).RowHeight = 20
    wsDash.Columns("A:S").ColumnWidth = 11
    wsDash.Columns("B").ColumnWidth = 20

    Debug.Print "FormatDashboardCards: 5 cards built. Card 2 risk-red triggered: " & (wsDash.Range("B3").Value > 0.25)

End Sub

