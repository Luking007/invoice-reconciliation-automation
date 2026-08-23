Attribute VB_Name = "mod_Import"
Option Explicit

'==========================================================
' mod_Import
' Reads invoices_raw.csv and payments_raw.csv as plain text.
' Reads the whole file as one block, then splits on line
' breaks manually (normalizing CRLF/LF/CR first) instead of
' using Line Input — works regardless of which OS wrote the file.
'==========================================================


Sub ImportInvoices()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileNum As Integer
    Dim wholeText As String
    Dim lines() As String
    Dim fields() As String
    Dim r As Long, i As Long

    LoadConfig
    filePath = RawDataFolderPath & InvoicesFileName

    If Dir(filePath) = "" Then
        MsgBox "Invoices file not found at:" & vbNewLine & filePath, vbCritical, "Import Failed"
        Exit Sub
    End If

    Set ws = GetOrCreateSheet("RawInvoices")
    ClearBelowHeader ws, 1
    ws.Columns("C:D").NumberFormat = "@"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    wholeText = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    wholeText = Replace(wholeText, vbCrLf, vbLf)
    wholeText = Replace(wholeText, vbCr, vbLf)
    lines = Split(wholeText, vbLf)

    r = 1
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            fields = Split(lines(i), ",")
            ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(fields) + 1)).Value = fields
            r = r + 1
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "ImportInvoices: " & (r - 2) & " data rows imported."
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If fileNum <> 0 Then Close #fileNum
    MsgBox "ImportInvoices failed: " & Err.Description, vbCritical
End Sub


Sub ImportPayments()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileNum As Integer
    Dim wholeText As String
    Dim lines() As String
    Dim fields() As String
    Dim r As Long, i As Long

    LoadConfig
    filePath = RawDataFolderPath & PaymentsFileName

    If Dir(filePath) = "" Then
        MsgBox "Payments file not found at:" & vbNewLine & filePath, vbCritical, "Import Failed"
        Exit Sub
    End If

    Set ws = GetOrCreateSheet("RawPayments")
    ClearBelowHeader ws, 1
    ws.Columns("C:C").NumberFormat = "@"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    wholeText = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    wholeText = Replace(wholeText, vbCrLf, vbLf)
    wholeText = Replace(wholeText, vbCr, vbLf)
    lines = Split(wholeText, vbLf)

    r = 1
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            fields = Split(lines(i), ",")
            ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(fields) + 1)).Value = fields
            r = r + 1
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "ImportPayments: " & (r - 2) & " data rows imported."
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If fileNum <> 0 Then Close #fileNum
    MsgBox "ImportPayments failed: " & Err.Description, vbCritical
End Sub


Sub ImportAll()
    ImportInvoices
    ImportPayments
    Debug.Print "ImportAll complete."
End Sub
