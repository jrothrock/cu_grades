Attribute VB_Name = "Module1"
Option Explicit

Public Sub deleteSheet(shtName As String)
    Dim xWs As Worksheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.name = shtName Then
            xWs.Delete
        End If
    Next
End Sub
