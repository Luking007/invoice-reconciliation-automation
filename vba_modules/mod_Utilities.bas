Attribute VB_Name = "mod_Utilities"
Option Explicit

'==========================================================
' mod_Utilities
' Reusable helper functions used by every other module in
' this project. Nothing in here is specific to invoices or
' payments on purpose — that's what makes it reusable.
'==========================================================


Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


Function GetOrCreateSheet(sheetName As String) As Worksheet
    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function


Sub ClearBelowHeader(ws As Worksheet, headerRow As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > headerRow Then
        ws.Rows(headerRow + 1 & ":" & lastRow).ClearContents
    End If
End Sub


Function ParseFlexibleDate(ByVal dateString As String) As Date
    On Error GoTo ParseError

    Dim cleanStr As String
    Dim y As Integer, m As Integer, d As Integer
    cleanStr = Trim(dateString)

    If Len(cleanStr) = 10 And Mid(cleanStr, 5, 1) = "-" Then
        ' Format: YYYY-MM-DD
        y = CInt(Left(cleanStr, 4))
        m = CInt(Mid(cleanStr, 6, 2))
        d = CInt(Mid(cleanStr, 9, 2))
    ElseIf Len(cleanStr) = 10 And Mid(cleanStr, 3, 1) = "/" Then
        ' Format: MM/DD/YYYY
        m = CInt(Left(cleanStr, 2))
        d = CInt(Mid(cleanStr, 4, 2))
        y = CInt(Right(cleanStr, 4))
    Else
        GoTo ParseError
    End If

    ParseFlexibleDate = DateSerial(y, m, d)
    Exit Function

ParseError:
    ParseFlexibleDate = 0   ' couldn't parse — caller checks for this
End Function

Function HexToRGB(ByVal hexColor As String) As Long
    Dim h As String
    h = Replace(hexColor, "#", "")
    HexToRGB = RGB(CLng("&H" & Mid(h, 1, 2)), CLng("&H" & Mid(h, 3, 2)), CLng("&H" & Mid(h, 5, 2)))
End Function
