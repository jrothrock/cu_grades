VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub findClass_Click()
    'is really a pre-initiation of the userform. Sets all the values on this listbox if there is a requirements and data worksheet
    Dim cell As Range, reqWks As Worksheet, i As Integer
    
    'So apparently an if statement will evaluate both sides, even if the first one fails. Interesting. If has to be broken up into two
    If Evaluate("ISREF('" & "Requirements" & "'!A1)") Then
        Set reqWks = ThisWorkbook.Worksheets("Requirements")
        If IsEmpty(ThisWorkbook.Worksheets("Requirements").Cells(1, Columns.Count).End(xlToLeft)) = False Then
            For i = 1 To ThisWorkbook.Worksheets("Requirements").Cells(1, Columns.Count).End(xlToLeft).Column
                classSearchForm.classListBox.AddItem (ThisWorkbook.Worksheets("Requirements").Cells(1, i).Value)
            Next i
            classSearchForm.classListBox.Selected(0) = True
        End If
    End If
    
    If Evaluate("ISREF('" & "Data" & "'!A1)") Then
        ThisWorkbook.Worksheets("Data").AutoFilter.ShowAllData
    End If
    
    classSearchForm.Show
End Sub

Private Sub Worksheet_Activate()
    'preserve watermarks'
    
    'ATTENTION: If you're reading this, please keep the watermarks. I could of just
    'not released this, or could have charged for it, and the minor advertising/watermarks
    'I believe are fair.
    
    ActiveSheet.Range("A1").Value = "Class Helper - Create A Schedule With The Highest GPA Producing Classes"
    ActiveSheet.Range("A2").Value = "Created By: Jack Rothrock"
End Sub


