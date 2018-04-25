VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Find The Best Classes"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   0
   ClientWidth     =   6160
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'references required:
'Microsoft HTML Object Library
'Microsoft Internet Controls
'Microsoft VBScript Regular Expressions 1.0 & 5.5

Private Sub downloadFCQData_Click()
    Call mainGrades
    Application.Workbooks("Project.xlsm").Worksheets("Main").Activate
End Sub

Private Sub downloadRequirements_Click()
    Call Main
    Application.Workbooks("Project.xlsm").Worksheets("Main").Activate
End Sub

Private Sub Main()
    'this is the main function for downloading the requirements/core-classes
    Dim ws As Worksheet
    Call deleteSheet("Requirements")
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    ws.Name = "Requirements"
    ws.Activate
    Call getRequirements
    Call updateListBox
    downloadFCQData.Enabled = True
    If Evaluate("ISREF('" & "Data" & "'!A1)") Then
        findClass.Enabled = True
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    End If
End Sub

Private Sub updateListBox()
    'clears the listbox then updates it with all of the categories of classes found in the requirements worksheet
    ListBox1.Clear
    Dim rngRequirements As Range, cell As Range
    Set rngRequirements = Application.Workbooks("Project.xlsm").Worksheets("Requirements").Range("A1:L1")
    For Each cell In rngRequirements
        ListBox1.AddItem (cell.Value)
    Next cell
    ListBox1.Selected(0) = True
End Sub

Private Sub getRequirements()
    'loops through the links array and places the scraped data on the requirements worksheet. Sections where there are upper division and lower divison
    'classes will be split into seperate areas. This applies to sequenced and non sequenced too.
    Dim strUrl As String, isSequenced As Boolean, seqULLoop As Variant, activeSheetCounterSequenced As Integer, seqLILoop As Variant, activeSheetCounterUpper As Integer, topCounter As Integer, strParts As Variant, isMulti As Boolean, htwo As Variant, objHTwo As Object, objIe As InternetExplorer, objHtml As HTMLDocument, strHtml As String, objDivs As Variant, objAnchors As IHTMLElementCollection, intCounter As Integer, links As Variant, i As Integer, activeSheetCounter As Integer
    links = Array("https://www.colorado.edu/artsandsciences/human-diversity", "https://www.colorado.edu/artsandsciences/current-students/core-curriculum/ideals-and-values", "https://www.colorado.edu/artsandsciences/historical-context", "https://www.colorado.edu/artsandsciences/literature-and-arts", "https://www.colorado.edu/artsandsciences/united-states-context", "https://www.colorado.edu/artsandsciences/contemporary-societies", "https://www.colorado.edu/artsandsciences/natural-science", "https://www.colorado.edu/artsandsciences/written-communication")
    For i = 0 To (UBound(links) - LBound(links))
        activeSheetCounter = 0
        activeSheetCounterUpper = 0
        
        strUrl = links(i)
        Set objIe = New InternetExplorer
        objIe.Visible = False
        objIe.navigate strUrl
        While objIe.readyState <> READYSTATE_COMPLETE
            DoEvents
        Wend

        Set objHtml = New HTMLDocument
        Set objHtml = objIe.document
        Set objDivs = objIe.document.getElementsByTagName("LI")
        Set objHTwo = objIe.document.getElementsByTagName("H2")
        
        isMulti = False
        isSequenced = False
        
        For Each htwo In objHTwo
            If htwo.innerText = "Upper-Division Courses" Then
                isMulti = True
            ElseIf htwo.innerText = "Two-Semester Sequences" Then
                isSequenced = True
            End If
        Next htwo
        
        If isMulti Then
            Application.ActiveSheet.Cells(1, 1 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " L/D"
            Application.ActiveSheet.Cells(1, 2 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " U/D"
            topCounter = topCounter + 2
        ElseIf isSequenced Then
            Application.ActiveSheet.Cells(1, 1 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " Seq."
            Application.ActiveSheet.Cells(1, 2 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " Non-Seq."
            Application.ActiveSheet.Cells(1, 3 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " Labs"
            topCounter = topCounter + 3
        Else
            Application.ActiveSheet.Cells(1, 1 + topCounter).Value = objIe.document.getElementById("page-title").innerText
            topCounter = topCounter + 1
        End If
        Dim regEx As New RegExp, strPattern As String
        strPattern = "<strong>"
    
        With regEx
            .Global = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With
        
        
        If objDivs.Length > 0 And isSequenced = False Then
            For intCounter = 0 To objDivs.Length - 1
                If regEx.test(objDivs(intCounter).innerHTML) Then
                    If isMulti Then
                        strParts = Split(objDivs(intCounter).innerText, " ")
                        strParts = Split(strParts(1), "-")
                        If CInt(strParts(0)) >= 3000 Then
                            Application.ActiveSheet.Cells(2 + activeSheetCounterUpper, topCounter).Value = objDivs(intCounter).innerText
                            activeSheetCounterUpper = activeSheetCounterUpper + 1
                        Else
                            Application.ActiveSheet.Cells(2 + activeSheetCounter, topCounter - 1).Value = objDivs(intCounter).innerText
                            activeSheetCounter = activeSheetCounter + 1
                        End If
                    Else
                        Application.ActiveSheet.Cells(2 + activeSheetCounter, topCounter).Value = objDivs(intCounter).innerText
                        activeSheetCounter = activeSheetCounter + 1
                    End If
                End If
            Next intCounter
         ElseIf isSequenced = True Then
            'this feels really redundent. However, there's not real pattern in the LIs
            Set seqULLoop = objIe.document.getElementsByTagName("UL")(4).getElementsByTagName("LI")
            For Each seqLILoop In seqULLoop
                Application.ActiveSheet.Cells(2 + activeSheetCounterSequenced, topCounter - 2).Value = seqLILoop.innerText
                activeSheetCounterSequenced = activeSheetCounterSequenced + 1
            Next seqLILoop
                        
            activeSheetCounterSequenced = 0
            Set seqULLoop = objIe.document.getElementsByTagName("UL")(5).getElementsByTagName("LI")
            For Each seqLILoop In seqULLoop
                Application.ActiveSheet.Cells(2 + activeSheetCounterSequenced, topCounter - 1).Value = seqLILoop.innerText
                activeSheetCounterSequenced = activeSheetCounterSequenced + 1
            Next seqLILoop
                        
            activeSheetCounterSequenced = 0
            Set seqULLoop = objIe.document.getElementsByTagName("UL")(6).getElementsByTagName("LI")
            For Each seqLILoop In seqULLoop
                Application.ActiveSheet.Cells(2 + activeSheetCounterSequenced, topCounter).Value = seqLILoop.innerText
                activeSheetCounterSequenced = activeSheetCounterSequenced + 1
            Next seqLILoop
         End If
        Set objHtml = Nothing
        objIe.Quit
        Set objIe = Nothing
    Next i
    ActiveSheet.Range("A1:L1").Columns.AutoFit
    ActiveSheet.Range("A1:L1").Interior.Color = RGB(207, 216, 220)
End Sub

Public Sub mainGrades()
    'is the main activation area for downloading the fcq grades, this calls all of the other functions
    Dim targetFile As String, ws As Worksheet
    Call deleteSheet("Data")
    targetFile = zipTarget()
    Call DownloadFile(targetFile)
    Call writeData(targetFile)
    If Evaluate("ISREF('" & "Requirements" & "'!A1)") Then
        findClass.Enabled = True
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    End If
End Sub


Public Sub writeData(Target As String)
    'opens the FCQ data from the targetDestination and copys the entire Data worksheet the project workbook.
    Dim gradesWkb As Workbook, destWkb As Workbook, projectSheet, wrksheet As Worksheet
    Set gradesWkb = Workbooks.Open(Target, True, True)
    Set wrksheet = gradesWkb.Worksheets("Data")
    Set destWkb = Application.Workbooks("Project.xlsm")
    wrksheet.Copy After:=destWkb.Worksheets(1)
    gradesWkb.Close savechanges:=False
    Set gradesWkb = Nothing
End Sub

Public Function zipTarget()
    'create a temp spot to where the download will be placed - this area will be used later to open the file
    Dim targetFolder As Variant, targetFileTXT As Variant
    targetFolder = Environ("TEMP") & "\" & RandomString(8) & "\"
    MkDir targetFolder
    zipTarget = targetFolder & "cu_grades_data.csv"
End Function

Public Sub DownloadFile(targetDest)
    'downloads the FCQ file and places it in the targetDest created in zipTarget
    Dim myURL As String, oStream As Object
    myURL = "https://www.colorado.edu/fcq/content/file-grade-distribution"

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False, "username", "password"
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile targetDest, 1 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
End Sub


Private Function RandomString(cb As Integer) As String
    'create the random string used as part of the destination target -ie. /tmp/asd35a/test.zip
    Randomize
    Dim rgch As String, i As Integer
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Private Sub findClass_Click()
    Call findClassMain
End Sub

Private Sub findClassMain()
    'basically loop through all of the courses and add the subject and number to an array that'll be used to filter out the classes
    'all of the classes are taken from the specified column in the requirements worksheet
    Dim reqWks As Worksheet, isSequenced As Boolean, dataWks As Worksheet, years As Variant, cell As Variant, subjectParts As Variant, numbers As Variant, tmpClassInfoStr As String, classInfoStr As String, numberStr As String, numberParts As Variant, strPattern As String, regEx As New RegExp, columnNum As Integer, i As Integer, valueStr As String, valueParts As Variant, subjects As Variant, subjectStr As String, tmpNumber As String, tmpSubject As String
    Set reqWks = Application.Workbooks("Project.xlsm").Worksheets("Requirements")
    Set dataWks = Application.Workbooks("Project.xlsm").Worksheets("Data")
    columnNum = Application.WorksheetFunction.Match(ListBox1.Value, reqWks.Range("A1:L1"), 0)
    If ListBox1.Value = "Natural Science Seq." Then
        isSequenced = True
    End If
    For i = 0 To (Application.WorksheetFunction.CountA(reqWks.Columns(columnNum)) - 2)
        valueStr = reqWks.Cells(2 + i, columnNum).Value
        valueParts = Split(valueStr, " ")
        subjectStr = valueParts(0)
        
        strPattern = "/"
        
        With regEx
            .Global = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With
        
        numberParts = Split(valueParts(1), "-")
        numberStr = numberParts(0)
        
        If regEx.test(numberStr) Then
            numberParts = Split(numberStr, "/")
            numberStr = numberParts(0)
        End If
        
        If Not tmpNumber Like ("*" & numberStr & "*") Then
            tmpNumber = tmpNumber + numberStr + "|"
        End If
        
        If regEx.test(subjectStr) Then
            subjectParts = Split(subjectStr, "/")
            subjectStr = subjectParts(0)
        End If
        If Not tmpSubject Like ("*" & subjectStr & "*") Then
            tmpSubject = tmpSubject + subjectStr + "|"
        End If
        classInfoStr = classInfoStr + subjectStr + numberStr + "|"
        
        If isSequenced Then
            numberParts = Split(valueParts(3), "-")
            numberStr = numberParts(0)
            If Not tmpNumber Like ("*" & numberStr & "*") Then
                tmpNumber = tmpNumber + numberStr + "|"
            End If
            classInfoStr = classInfoStr + subjectStr + numberStr + "|"
        End If
    Next i
    tmpSubject = Left(tmpSubject, Len(tmpSubject) - 1)
    tmpNumber = Left(tmpNumber, Len(tmpNumber) - 1)
    classInfoStr = Left(classInfoStr, Len(classInfoStr) - 1)
    subjects = Split(tmpSubject, "|")
    numbers = Split(tmpNumber, "|")
    dataWks.Range("A4:F" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).AutoFilter field:=5, Criteria1:=(subjects), VisibleDropDown:=True, Operator:=xlFilterValues
    dataWks.Range("A4:F" & Application.WorksheetFunction.CountA(dataWks.Columns(6))).AutoFilter field:=6, Criteria1:=(numbers), VisibleDropDown:=True, Operator:=xlFilterValues
    If UserForm1.OptionButton1.Value = True Then
        dataWks.Range("A4:F" & Application.WorksheetFunction.CountA(dataWks.Columns(6))).AutoFilter field:=1, Criteria1:="2017*", VisibleDropDown:=True
    Else
        years = Split("20171|20161|20167|20177", "|")
        dataWks.Range("A4:F" & Application.WorksheetFunction.CountA(dataWks.Columns(6))).AutoFilter field:=1, Criteria1:=(years), VisibleDropDown:=True, Operator:=xlFilterValues
    End If
    
    'Have to loop through the cells again, as there may be some classes that were supposed to be filtered but weren't'
    For Each cell In dataWks.Range("A5:A" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible)
        tmpClassInfoStr = cell.Offset(0, 4).Value & cell.Offset(0, 5).Value
        If Not classInfoStr Like ("*" & tmpClassInfoStr & "*") Then
            cell.EntireRow.Hidden = True
        End If
    Next cell
    dataWks.Range("Q5").Sort key1:=dataWks.Range("Q5"), Order1:=xlDescending
    dataWks.Range("E5:G" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(237, 231, 246)
    dataWks.Range("Q5:Q" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(232, 234, 246)
    dataWks.Range("AE5:AE" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(232, 245, 233)
    dataWks.Activate
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'check to see if requirements and data worksheets exist. If so enable userform buttons'
    If Evaluate("ISREF('" & "Requirements" & "'!A1)") Then
        UserForm1.downloadFCQData.Enabled = True
        If Evaluate("ISREF('" & "Data" & "'!A1)") Then
            UserForm1.findClass.Enabled = True
            UserForm1.OptionButton1.Enabled = True
            UserForm1.OptionButton2.Enabled = True
        End If
    End If
End Sub
