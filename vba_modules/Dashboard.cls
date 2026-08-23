VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B8")) Is Nothing Then
        Application.EnableEvents = False
        On Error GoTo Cleanup

        Dim selectedCustomer As String
        selectedCustomer = Trim(Me.Range("B8").Value)

        If selectedCustomer = "" Or selectedCustomer = "All Customers" Then
            BuildPivotStaging
            UpdateChartTitles "All Customers"
        Else
            BuildPivotStaging selectedCustomer
            UpdateChartTitles selectedCustomer
        End If

Cleanup:
        Application.EnableEvents = True
    End If
End Sub
