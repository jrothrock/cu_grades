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
        Set reqWks = Application.Workbooks("Project.xlsm").Worksheets("Requirements")
        If IsEmpty(Application.Workbooks("Project.xlsm").Worksheets("Requirements").Range("L1")) = False Then
            For i = 1 To Application.Workbooks("Project.xlsm").Worksheets("Requirements").Range("A1:L1").Columns.Count
                UserForm1.ListBox1.AddItem (Application.Workbooks("Project.xlsm").Worksheets("Requirements").Cells(1, i).Value)
            Next i
            UserForm1.ListBox1.Selected(0) = True
        End If
    End If
    
    If Evaluate("ISREF('" & "Data" & "'!A1)") Then
        Application.Workbooks("Project.xlsm").Worksheets("Data").AutoFilter.ShowAllData
    End If
    
    UserForm1.Show
End Sub

Private Sub Worksheet_Activate()
    'preserve watermarks'
    
    'ATTENTION: If you're reading this, please keep the watermarks. I could of just
    'not released this, or could have charged for it, and the minor advertising/watermarks
    'I believe are fair.
    
    ActiveSheet.Range("A1").Value = "Class Helper - Create A Schedule With The Highest GPA Producing Classes"
    ActiveSheet.Range("A2").Value = "Created By: Jack Rothrock"
End Sub


