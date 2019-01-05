VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} classSearchForm 
   Caption         =   "Find The Best Classes"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   6160
   OleObjectBlob   =   "classSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "classSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'references required:
'Microsoft HTML Object Library
'Microsoft Internet Controls
'Microsoft VBScript Regular Expressions 1.0 & 5.5

Private Sub UserForm_Initialize()
    'check to see if requirements and data worksheets exist. If so enable userform buttons'
    If Evaluate("ISREF('" & "Requirements" & "'!A1)") Then
        classSearchForm.downloadFCQData.Enabled = True
        If Evaluate("ISREF('" & "Data" & "'!A1)") Then
            classSearchForm.findClass.Enabled = True
            classSearchForm.oneYearButton.Enabled = True
            classSearchForm.multiYearButton.Enabled = True
        End If
    End If
    classSearchForm.oneYearButton.Caption = "Only " & (Year(Date) - 1)
    classSearchForm.multiYearButton.Caption = (Year(Date) - 1) & " & " & (Year(Date) - 2)
End Sub

Private Sub downloadRequirements_Click()
    Call mainRequirements
    ThisWorkbook.Worksheets("Main").Activate
End Sub

Private Sub mainRequirements()
    'this is the main function for downloading the requirements/core-classes
    Dim ws As Worksheet
    Call deleteSheet("Requirements")
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    ws.name = "Requirements"
    ws.Activate
    Call getRequirements
    Call updateListBox
    downloadFCQData.Enabled = True
    If Evaluate("ISREF('" & "Data" & "'!A1)") Then
        findClass.Enabled = True
        oneYearButton.Enabled = True
        multiYearButton.Enabled = True
    End If
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
            Application.ActiveSheet.Cells(1, 2 + topCounter).Value = objIe.document.getElementById("page-title").innerText & " Non Seq."
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
            'this feels really redundent. However, there's no real pattern in the LIs
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
    ActiveSheet.Range("A1:" & ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Address).Columns.AutoFit
    ActiveSheet.Range("A1:" & ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Address).Interior.Color = RGB(207, 216, 220)
End Sub

Private Sub updateListBox()
    'clears the listbox then updates it with all of the categories of classes found in the requirements worksheet
    classListBox.Clear
    Dim rngRequirements As Range, cell As Range, wkSheet
    Set wkSheet = ThisWorkbook.Worksheets("Requirements")
    Set rngRequirements = wkSheet.Range("A1:" & wkSheet.Cells(1, Columns.Count).End(xlToLeft).Address)
    For Each cell In rngRequirements
        classListBox.AddItem (cell.Value)
    Next cell
    classListBox.Selected(0) = True
End Sub

Private Sub downloadFCQData_Click()
    Call mainGrades
    ThisWorkbook.Worksheets("Main").Activate
End Sub

Private Sub mainGrades()
    'is the main activation area for downloading the fcq grades, this calls all of the other functions
    Dim targetFile As String, ws As Worksheet
    Call deleteSheet("Data")
    targetFile = zipTarget()
    Call DownloadFile(targetFile)
    Call writeData(targetFile)
    If Evaluate("ISREF('" & "Requirements" & "'!A1)") Then
        findClass.Enabled = True
        oneYearButton.Enabled = True
        multiYearButton.Enabled = True
    End If
End Sub

Private Function zipTarget()
    'create a temp spot to where the download will be placed - this area will be used later to open the file
    Dim targetFolder As Variant, targetFileTXT As Variant
    targetFolder = Environ("TEMP") & "\" & RandomString(8) & "\"
    MkDir targetFolder
    zipTarget = targetFolder & "cu_grades_data.csv"
End Function

Private Function RandomString(cb As Integer) As String
    'create the random string used as part of the destination target -ie. /tmp/asd35a/test.zip
    Dim rgch As String, i As Integer
    Randomize
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Private Sub DownloadFile(targetDest)
    'downloads the FCQ file and places it in the targetDest created in zipTarget
    Dim myURL As String, oStream As Object, WinHttpReq
    myURL = "https://www.colorado.edu/fcq/content/file-grade-distribution"

    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False, "username", "password"
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile targetDest, 1
        oStream.Close
    End If
End Sub

Private Sub writeData(Target As String)
    'opens the FCQ data from the targetDestination and copys the entire Data worksheet the project workbook.
    Dim gradesWkb As Workbook, projectSheet, wrksheet As Worksheet
    Set gradesWkb = Workbooks.Open(Target, True, True)
    Set wrksheet = gradesWkb.Worksheets("Data")
    wrksheet.Copy After:=ThisWorkbook.Worksheets(1)
    Call modifyData
    gradesWkb.Close savechanges:=False
    Set gradesWkb = Nothing
End Sub

Private Sub modifyData()
    'Move values around, deletes dumb data, and applies styles.
    Dim dataWks As Worksheet
    Set dataWks = ThisWorkbook.Worksheets("data")
    dataWks.Range("A3").Font.Color = RGB(0, 0, 0)
    dataWks.Range("A1, A2, A3").Value = ""
    dataWks.Range("A2, A3, A4").EntireRow.Insert
    dataWks.Range("A2, A3, A4, A5, A6").EntireRow.Insert
    dataWks.Range("A2, A3").EntireRow.Insert
    dataWks.Range("A2").UnMerge
    dataWks.Range("A1:G1").Style = ActiveWorkbook.Styles("Heading 2")
    dataWks.Range("A1:G1").Merge
    dataWks.Columns("L:L").UnMerge
    dataWks.Columns("L:L").Cut Destination:=dataWks.Columns("H:H")
    dataWks.Range("H14").Value = "Credit Hours"
    dataWks.Columns("M:M").UnMerge
    dataWks.Columns("M:M").Cut Destination:=dataWks.Columns("I:I")
    dataWks.Range("M14").Value = "N_ENROLL"
    dataWks.Columns("Q:V").UnMerge
    dataWks.Columns("Q:V").Cut Destination:=dataWks.Columns("J:O")
    dataWks.Columns("AE:AE").UnMerge
    dataWks.Columns("AE:AE").Cut Destination:=dataWks.Columns("P:P")
    dataWks.Columns("AM:AM").UnMerge
    dataWks.Columns("AM:AM").Cut Destination:=dataWks.Columns("Q:Q")
    dataWks.Columns("R:AZ").Delete
    dataWks.Range("A2:G2").Merge
    dataWks.Range("A3:G3").Merge
    dataWks.Range("A4:G4").Merge
    dataWks.Range("A5:G5").Merge
    dataWks.Range("A6:G6").Merge
    dataWks.Range("A7:G7").Merge
    dataWks.Range("A8:G8").Merge
    dataWks.Range("A9:G9").Merge
    dataWks.Range("A10:G10").Merge
    dataWks.Range("A11:G11").Merge
    dataWks.Range("A12:G12").Merge
    dataWks.Range("A13:G13").Merge
    dataWks.Range("A7:A10").Merge
    dataWks.Range("A7").Value = "While the above best class is a weighted average, it may not actually be the best. When looking at the data, classes that have very few enrollments (with super high GPAs) maybe meant for student TAs. A good example of this is the Health and Nutrition Class. Always check the number of enrolled."
    dataWks.Range("A7, A9").VerticalAlignment = xlTop
    dataWks.Range("A7, A9").WrapText = True
End Sub

Private Sub findClass_Click()
    Call findClassMain
End Sub

Private Sub findClassMain()
    'basically loop through all of the courses and add the subject and number to an array that'll be used to filter out the classes
    'all of the classes are taken from the specified column in the requirements worksheet
    Dim reqWks As Worksheet, classWks As Worksheet, personCount As Integer, tmpTotalWorkLoadClasses As Integer, tmpATotal As Double, tmpBTotal As Double, tmpCTotal As Double, tmpDTotal As Double, tmpFTotal As Double, totalGPA As Double, lowestClass As String, highestClass As String, lowestClassGPA As Double, highestClassGPA As Double, tmpTotalGPA As Double, tmpWeightedCount As Double, tmpTotalClasses As Integer, tmpTotalWorkLoad As Double, tmpCount As Integer, classGPA As Collection, teacherGPA As Collection, classGPAKeys() As String, tempClassGPAKeys As String, teacherGPAKeys() As String, tempTeacherGPAKeys As String
    Dim isSequenced As Boolean, dataWks As Worksheet, years As Variant, cell As Variant, subjectParts As Variant, numbers As Variant, tmpClassInfoStr As String, classInfoStr As String, numberStr As String, numberParts As Variant, strPattern As String, regEx As New RegExp, columnNum As Integer, i As Integer, valueStr As String, valueParts As Variant, subjects As Variant, subjectStr As String, tmpNumber As String, tmpSubject As String
    Dim highestTeacher As String, highestTeacherGPA As Double, lowestTeacher As String, lowestTeacherGPA As Double, teacherWks As Worksheet, tmpTeacherClasses As String, tmpWeightedTotalGPA As Double
    Set reqWks = ThisWorkbook.Worksheets("Requirements")
    Set dataWks = ThisWorkbook.Worksheets("Data")
    Set classGPA = New Collection
    Set teacherGPA = New Collection
    lowestClassGPA = 4
    columnNum = Application.WorksheetFunction.Match(classListBox.Value, reqWks.Range("A1:" & reqWks.Cells(1, Columns.Count).End(xlToLeft).Address), 0)
    If classListBox.Value = "Natural Science Seq." Then
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
    dataWks.Range("A12:F" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).AutoFilter field:=5, Criteria1:=(subjects), VisibleDropDown:=True, Operator:=xlFilterValues
    dataWks.Range("A12:F" & Application.WorksheetFunction.CountA(dataWks.Columns(6))).AutoFilter field:=6, Criteria1:=(numbers), VisibleDropDown:=True, Operator:=xlFilterValues
    years = IIf(classSearchForm.oneYearButton.Value, Split(Year(Date) - 1 & "1" & "|" & Year(Date) - 1 & "7", "|"), Split(Year(Date) - 1 & "1" & "|" & Year(Date) - 1 & "7" & "|" & Year(Date) - 2 & "1" & "|" & Year(Date) - 2 & "7", "|"))
    dataWks.Range("A12:F" & Application.WorksheetFunction.CountA(dataWks.Columns(6))).AutoFilter field:=1, Criteria1:=(years), VisibleDropDown:=True, Operator:=xlFilterValues
    
    'Have to loop through the cells again, as there may be some classes that were supposed to be filtered but weren't'
    For Each cell In dataWks.Range("A15:A" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible)
        tmpClassInfoStr = cell.Offset(0, 4).Value & cell.Offset(0, 5).Value
        If Not classInfoStr Like ("*" & tmpClassInfoStr & "*") Then
            cell.EntireRow.Hidden = True
        Else
            personCount = personCount + CInt(cell.Offset(0, 8).Value)
            totalGPA = totalGPA + CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)
            If cell.Offset(0, 4).Value <> "HONR" Then
                If HasKey(classGPA, cell.Offset(0, 6).Value) Then
                    tmpCount = classGPA.Item(cell.Offset(0, 6).Value + "_count")
                    tmpTotalGPA = classGPA.Item(cell.Offset(0, 6).Value + "_total")
                    tmpWeightedTotalGPA = classGPA.Item(cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA")
                    tmpWeightedCount = classGPA.Item(cell.Offset(0, 6).Value + "_tmpWeightedCount")
                    tmpTotalClasses = classGPA.Item(cell.Offset(0, 6).Value + "_classCount")
                    tmpTotalWorkLoad = classGPA.Item(cell.Offset(0, 6).Value + "_workloadTotal")
                    tmpTotalWorkLoadClasses = classGPA.Item(cell.Offset(0, 6).Value + "_workloadClasses")
                    tmpATotal = classGPA.Item(cell.Offset(0, 6).Value + "_aTotal")
                    tmpBTotal = classGPA.Item(cell.Offset(0, 6).Value + "_bTotal")
                    tmpCTotal = classGPA.Item(cell.Offset(0, 6).Value + "_cTotal")
                    tmpDTotal = classGPA.Item(cell.Offset(0, 6).Value + "_dTotal")
                    tmpFTotal = classGPA.Item(cell.Offset(0, 6).Value + "_fTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_count")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_total")
                    classGPA.Remove (cell.Offset(0, 6).Value)
                    classGPA.Remove (cell.Offset(0, 6).Value + "_classCount")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_tmpWeightedCount")
                    If IsEmpty(cell.Offset(0, 15)) <> True Then
                        classGPA.Remove (cell.Offset(0, 6).Value + "_workloadTotal")
                        classGPA.Remove (cell.Offset(0, 6).Value + "_workloadClasses")
                        classGPA.Remove (cell.Offset(0, 6).Value + "_workload")
                    End If
                    classGPA.Remove (cell.Offset(0, 6).Value + "_aTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_a")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_bTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_b")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_cTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_c")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_dTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_d")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_fTotal")
                    classGPA.Remove (cell.Offset(0, 6).Value + "_f")
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_count"), Item:=tmpCount + cell.Offset(0, 8).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_total"), Item:=tmpTotalGPA + CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)
                    classGPA.Add Key:=(cell.Offset(0, 6).Value), Item:=((tmpTotalGPA + (CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value))) / (tmpCount + cell.Offset(0, 8).Value))
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_classCount"), Item:=(tmpTotalClasses + 1)
                    If IsEmpty(cell.Offset(0, 15)) <> True Then
                        classGPA.Add Key:=(cell.Offset(0, 6).Value + "_workloadTotal"), Item:=tmpTotalWorkLoad + cell.Offset(0, 15).Value
                        classGPA.Add Key:=(cell.Offset(0, 6).Value + "_workloadClasses"), Item:=(tmpTotalWorkLoadClasses + 1)
                        classGPA.Add Key:=(cell.Offset(0, 6).Value + "_workload"), Item:=((tmpTotalWorkLoad + cell.Offset(0, 15).Value) / (tmpTotalWorkLoadClasses + 1))
                    End If
                    If cell.Value = Year(Date) - 2 & "1" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=tmpWeightedTotalGPA + (CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.6)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=tmpWeightedCount + cell.Offset(0, 8).Value * 0.6
                    ElseIf cell.Value = Year(Date) - 2 & "7" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=tmpWeightedTotalGPA + (CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.75)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=tmpWeightedCount + cell.Offset(0, 8).Value * 0.75
                    ElseIf cell.Value = Year(Date) - 1 & "1" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=tmpWeightedTotalGPA + (CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.9)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=tmpWeightedCount + cell.Offset(0, 8).Value * 0.9
                    ElseIf cell.Value = Year(Date) - 1 & "7" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=tmpWeightedTotalGPA + CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=tmpWeightedCount + cell.Offset(0, 8).Value
                    End If
                
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_aTotal"), Item:=tmpATotal + cell.Offset(0, 10).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_a"), Item:=((tmpATotal + cell.Offset(0, 10).Value) / (tmpTotalClasses + 1))
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_bTotal"), Item:=tmpBTotal + cell.Offset(0, 11).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_b"), Item:=((tmpBTotal + cell.Offset(0, 11).Value) / (tmpTotalClasses + 1))
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_cTotal"), Item:=tmpCTotal + cell.Offset(0, 12).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_c"), Item:=((tmpCTotal + cell.Offset(0, 12).Value) / (tmpTotalClasses + 1))
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_dTotal"), Item:=tmpDTotal + cell.Offset(0, 13).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_d"), Item:=((tmpDTotal + cell.Offset(0, 13).Value) / (tmpTotalClasses + 1))
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_fTotal"), Item:=tmpFTotal + cell.Offset(0, 14).Value
                    classGPA.Add Key:=(cell.Offset(0, 6).Value + "_f"), Item:=((tmpFTotal + cell.Offset(0, 14).Value) / (tmpTotalClasses + 1))
                Else
                    tempClassGPAKeys = tempClassGPAKeys + cell.Offset(0, 6).Value + "|"
                    classGPA.Add Item:=cell.Offset(0, 4).Value + " " + cell.Offset(0, 5).Value, Key:=cell.Offset(0, 6).Value + "_classCode"
                    classGPA.Add Item:=CDbl(cell.Offset(0, 8).Value), Key:=cell.Offset(0, 6).Value & "_count"
                    classGPA.Add Item:=(CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)), Key:=cell.Offset(0, 6).Value & "_total"
                    classGPA.Add Item:=((CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)) / CDbl(cell.Offset(0, 8).Value)), Key:=cell.Offset(0, 6).Value
                    classGPA.Add Item:=1, Key:=cell.Offset(0, 6).Value + "_classCount"
                    classGPA.Add Item:=cell.Offset(0, 15).Value, Key:=cell.Offset(0, 6).Value + "_workloadTotal"
                    classGPA.Add Item:=cell.Offset(0, 15).Value, Key:=cell.Offset(0, 6).Value + "_workload"
                    classGPA.Add Item:=IIf(IsEmpty(cell.Offset(0, 15)) = False, 1, 0), Key:=cell.Offset(0, 6).Value + "_workloadClasses"
                    classGPA.Add Item:=cell.Offset(0, 10).Value, Key:=cell.Offset(0, 6).Value + "_aTotal"
                    classGPA.Add Item:=cell.Offset(0, 10).Value, Key:=cell.Offset(0, 6).Value + "_a"
                    classGPA.Add Item:=cell.Offset(0, 11).Value, Key:=cell.Offset(0, 6).Value + "_bTotal"
                    classGPA.Add Item:=cell.Offset(0, 11).Value, Key:=cell.Offset(0, 6).Value + "_b"
                    classGPA.Add Item:=cell.Offset(0, 12).Value, Key:=cell.Offset(0, 6).Value + "_cTotal"
                    classGPA.Add Item:=cell.Offset(0, 12).Value, Key:=cell.Offset(0, 6).Value + "_c"
                    classGPA.Add Item:=cell.Offset(0, 13).Value, Key:=cell.Offset(0, 6).Value + "_dTotal"
                    classGPA.Add Item:=cell.Offset(0, 13).Value, Key:=cell.Offset(0, 6).Value + "_d"
                    classGPA.Add Item:=cell.Offset(0, 14).Value, Key:=cell.Offset(0, 6).Value + "_fTotal"
                    classGPA.Add Item:=cell.Offset(0, 14).Value, Key:=cell.Offset(0, 6).Value + "_f"
                    If cell.Value = Year(Date) - 2 & "1" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=(CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.6)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=cell.Offset(0, 8).Value * 0.6
                    ElseIf cell.Value = Year(Date) - 2 & "7" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=(CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.75)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=cell.Offset(0, 8).Value * 0.75
                    ElseIf cell.Value = Year(Date) - 1 & "1" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=(CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value) * 0.9)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=cell.Offset(0, 8).Value * 0.9
                    ElseIf cell.Value = Year(Date) - 1 & "7" Then
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedTotalGPA", Item:=CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)
                        classGPA.Add Key:=cell.Offset(0, 6).Value + "_tmpWeightedCount", Item:=cell.Offset(0, 8).Value
                    End If
                End If
                If HasKey(teacherGPA, cell.Offset(0, 16).Value) Then
                    tmpCount = teacherGPA.Item(cell.Offset(0, 16).Value + "_count")
                    tmpTeacherClasses = teacherGPA.Item(cell.Offset(0, 16).Value + "_classes")
                    tmpTotalGPA = teacherGPA.Item(cell.Offset(0, 16).Value + "_total")
                    tmpTotalClasses = teacherGPA.Item(cell.Offset(0, 16).Value + "_classCount")
                    tmpTotalWorkLoad = teacherGPA.Item(cell.Offset(0, 16).Value + "_workloadTotal")
                    tmpTotalWorkLoadClasses = teacherGPA.Item(cell.Offset(0, 16).Value + "_workloadClasses")
                    tmpATotal = teacherGPA.Item(cell.Offset(0, 16).Value + "_aTotal")
                    tmpBTotal = teacherGPA.Item(cell.Offset(0, 16).Value + "_bTotal")
                    tmpCTotal = teacherGPA.Item(cell.Offset(0, 16).Value + "_cTotal")
                    tmpDTotal = teacherGPA.Item(cell.Offset(0, 16).Value + "_dTotal")
                    tmpFTotal = teacherGPA.Item(cell.Offset(0, 16).Value + "_fTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_count")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_classes")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_total")
                    teacherGPA.Remove (cell.Offset(0, 16).Value)
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_classCount")
                    If IsEmpty(cell.Offset(0, 15)) <> True Then
                        teacherGPA.Remove (cell.Offset(0, 16).Value + "_workloadTotal")
                        teacherGPA.Remove (cell.Offset(0, 16).Value + "_workloadClasses")
                        teacherGPA.Remove (cell.Offset(0, 16).Value + "_workload")
                    End If
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_aTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_a")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_bTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_b")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_cTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_c")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_dTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_d")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_fTotal")
                    teacherGPA.Remove (cell.Offset(0, 16).Value + "_f")
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_count"), Item:=tmpCount + cell.Offset(0, 8).Value
                    If tmpTeacherClasses Like "*" & cell.Offset(0, 6).Value & "*" Then
                        teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_classes"), Item:=tmpTeacherClasses
                    Else
                        teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_classes"), Item:=tmpTeacherClasses & " -- " & cell.Offset(0, 6).Value
                    End If
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_total"), Item:=tmpTotalGPA + CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value), Item:=((tmpTotalGPA + (CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value))) / (tmpCount + cell.Offset(0, 8).Value))
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_classCount"), Item:=(tmpTotalClasses + 1)
                    If IsEmpty(cell.Offset(0, 15)) <> True Then
                        teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_workloadTotal"), Item:=tmpTotalWorkLoad + cell.Offset(0, 15).Value
                        teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_workloadClasses"), Item:=(tmpTotalWorkLoadClasses + 1)
                        teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_workload"), Item:=((tmpTotalWorkLoad + cell.Offset(0, 15).Value) / (tmpTotalWorkLoadClasses + 1))
                    End If
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_aTotal"), Item:=tmpATotal + cell.Offset(0, 10).Value
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_a"), Item:=((tmpATotal + cell.Offset(0, 10).Value) / (tmpTotalClasses + 1))
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_bTotal"), Item:=tmpBTotal + cell.Offset(0, 11).Value
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_b"), Item:=((tmpBTotal + cell.Offset(0, 11).Value) / (tmpTotalClasses + 1))
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_cTotal"), Item:=tmpCTotal + cell.Offset(0, 12).Value
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_c"), Item:=((tmpCTotal + cell.Offset(0, 12).Value) / (tmpTotalClasses + 1))
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_dTotal"), Item:=tmpDTotal + cell.Offset(0, 13).Value
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_d"), Item:=((tmpDTotal + cell.Offset(0, 13).Value) / (tmpTotalClasses + 1))
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_fTotal"), Item:=tmpFTotal + cell.Offset(0, 14).Value
                    teacherGPA.Add Key:=(cell.Offset(0, 16).Value + "_f"), Item:=((tmpFTotal + cell.Offset(0, 14).Value) / (tmpTotalClasses + 1))
                Else
                    Debug.Print (cell.Offset(0, 16).Value)
                    tempTeacherGPAKeys = tempTeacherGPAKeys + cell.Offset(0, 16).Value + "|"
                    teacherGPA.Add Item:=cell.Offset(0, 6), Key:=cell.Offset(0, 16).Value & "_classes"
                    teacherGPA.Add Item:=CDbl(cell.Offset(0, 8).Value), Key:=cell.Offset(0, 16).Value & "_count"
                    teacherGPA.Add Item:=(CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)), Key:=cell.Offset(0, 16).Value & "_total"
                    teacherGPA.Add Item:=((CDbl(cell.Offset(0, 8).Value) * CDbl(cell.Offset(0, 9).Value)) / CDbl(cell.Offset(0, 8).Value)), Key:=cell.Offset(0, 16).Value
                    teacherGPA.Add Item:=1, Key:=cell.Offset(0, 16).Value + "_classCount"
                    teacherGPA.Add Item:=cell.Offset(0, 15).Value, Key:=cell.Offset(0, 16).Value + "_workloadTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 15).Value, Key:=cell.Offset(0, 16).Value + "_workload"
                    teacherGPA.Add Item:=IIf(IsEmpty(cell.Offset(0, 15)) = False, 1, 0), Key:=cell.Offset(0, 16).Value + "_workloadClasses"
                    teacherGPA.Add Item:=cell.Offset(0, 10).Value, Key:=cell.Offset(0, 16).Value + "_aTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 10).Value, Key:=cell.Offset(0, 16).Value + "_a"
                    teacherGPA.Add Item:=cell.Offset(0, 11).Value, Key:=cell.Offset(0, 16).Value + "_bTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 11).Value, Key:=cell.Offset(0, 16).Value + "_b"
                    teacherGPA.Add Item:=cell.Offset(0, 12).Value, Key:=cell.Offset(0, 16).Value + "_cTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 12).Value, Key:=cell.Offset(0, 16).Value + "_c"
                    teacherGPA.Add Item:=cell.Offset(0, 13).Value, Key:=cell.Offset(0, 16).Value + "_dTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 13).Value, Key:=cell.Offset(0, 16).Value + "_d"
                    teacherGPA.Add Item:=cell.Offset(0, 14).Value, Key:=cell.Offset(0, 16).Value + "_fTotal"
                    teacherGPA.Add Item:=cell.Offset(0, 14).Value, Key:=cell.Offset(0, 16).Value + "_f"
                End If
            End If
        End If
    Next cell
    dataWks.Range("J15").Sort key1:=dataWks.Range("J15"), Order1:=xlDescending
    dataWks.Range("E15:G" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(237, 231, 246)
    dataWks.Range("P15:P" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(232, 234, 246)
    dataWks.Range("J15:J" & Application.WorksheetFunction.CountA(dataWks.Columns(5))).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(232, 245, 233)
    dataWks.Range("A1").Value = "Best Classes For " & classSearchForm.classListBox.Value
    dataWks.Range("A3").Value = "Average " & classSearchForm.classListBox.Value & " GPA: " & Round(CDbl(totalGPA / personCount), 2)
    tempClassGPAKeys = Left(tempClassGPAKeys, Len(tempClassGPAKeys) - 1)
    tempTeacherGPAKeys = Left(tempTeacherGPAKeys, Len(tempTeacherGPAKeys) - 1)
    classGPAKeys = Split(tempClassGPAKeys, "|")
    teacherGPAKeys = Split(tempTeacherGPAKeys, "|")
    Set classWks = createClassSheet()
    Set teacherWks = createTeacherSheet()
    classWks.Range("A1:G1").Merge
    classWks.Range("A1").Value = "Class Averages For " & classListBox.Value & " (Avg. Total GPA: " & Round(CDbl(totalGPA / personCount), 2) & ")"
    classWks.Range("A1").Font.Bold = True
    classWks.Range("A6").Value = "Class Code:"
    classWks.Range("B6").Value = "Class:"
    classWks.Range("C6").Value = "GPA:"
    classWks.Range("D6").Value = "GPA Trend:"
    classWks.Range("E6").Value = "N_Enroll:"
    classWks.Range("F6").Value = "N_Classes:"
    classWks.Range("G6").Value = "Workload avg:"
    classWks.Range("H6").Value = "PCT_A:"
    classWks.Range("I6").Value = "PCT_B:"
    classWks.Range("J6").Value = "PCT_C:"
    classWks.Range("K6").Value = "PCT_D:"
    classWks.Range("L6").Value = "PCT_F:"
    classWks.Range("A6:L6").Font.Bold = True
    For i = 0 To (UBound(classGPAKeys) - 1)
        If classGPA.Item(classGPAKeys(i)) > highestClassGPA Then
            highestClass = classGPAKeys(i)
            highestClassGPA = classGPA.Item(classGPAKeys(i))
        ElseIf classGPA.Item(classGPAKeys(i)) < lowestClassGPA And classGPA.Item(classGPAKeys(i)) > 0 Then
            lowestClass = classGPAKeys(i)
            lowestClassGPA = classGPA.Item(classGPAKeys(i))
        End If
        classWks.Cells(i + 7, 1).Value = classGPA.Item(classGPAKeys(i) & "_classCode")
        classWks.Cells(i + 7, 2).Value = classGPAKeys(i)
        classWks.Cells(i + 7, 3).Value = Round(classGPA.Item(classGPAKeys(i)), 2)
        If ((classGPA.Item(classGPAKeys(i) & "_tmpWeightedTotalGPA") / classGPA.Item(classGPAKeys(i) & "_tmpWeightedCount")) - Round(classGPA.Item(classGPAKeys(i)), 2)) > 0.015 Then
             classWks.Cells(i + 7, 4).Value = "Increasing"
        ElseIf ((classGPA.Item(classGPAKeys(i) & "_tmpWeightedTotalGPA") / classGPA.Item(classGPAKeys(i) & "_tmpWeightedCount")) - Round(classGPA.Item(classGPAKeys(i)), 2)) < -0.015 Then
            classWks.Cells(i + 7, 4).Value = "Decreasing"
        Else
            classWks.Cells(i + 7, 4).Value = "Similar"
            'classWks.Cells(i + 7, 4).Value = ((classGPA.Item(classGPAKeys(i) & "_tmpWeightedTotalGPA") / classGPA.Item(classGPAKeys(i) & "_tmpWeightedCount"))
            'classWks.Cells(i + 7, 4).Value = ((classGPA.Item(classGPAKeys(i) & "_tmpWeightedTotalGPA") / classGPA.Item(classGPAKeys(i) & "_tmpWeightedCount")) - Round(classGPA.Item(classGPAKeys(i)), 2))
        End If
        
        classWks.Cells(i + 7, 5).Value = classGPA.Item(classGPAKeys(i) & "_count")
        classWks.Cells(i + 7, 6).Value = classGPA.Item(classGPAKeys(i) & "_classCount")
        classWks.Cells(i + 7, 7).Value = Round(classGPA.Item(classGPAKeys(i) & "_workload"), 3)
        classWks.Cells(i + 7, 8).Value = Round(classGPA.Item(classGPAKeys(i) & "_a"), 3)
        classWks.Cells(i + 7, 9).Value = Round(classGPA.Item(classGPAKeys(i) & "_b"), 3)
        classWks.Cells(i + 7, 10).Value = Round(classGPA.Item(classGPAKeys(i) & "_c"), 3)
        classWks.Cells(i + 7, 11).Value = Round(classGPA.Item(classGPAKeys(i) & "_d"), 3)
        classWks.Cells(i + 7, 12).Value = Round(classGPA.Item(classGPAKeys(i) & "_f"), 3)
    Next i
    dataWks.Range("A4").Value = "Best Class: " & highestClass & " (avg. GPA: " & Round(highestClassGPA, 2) & ")"
    dataWks.Range("A5").Value = "Worst Class: " & lowestClass & " (avg. GPA: " & Round(lowestClassGPA, 2) & ")"
    dataWks.Range("A3").Interior.Color = RGB(255, 248, 225)
    dataWks.Range("A4").Interior.Color = RGB(232, 245, 233)
    dataWks.Range("A5").Interior.Color = RGB(251, 233, 231)
    classWks.Range("H7:L" & Application.WorksheetFunction.CountA(dataWks.Columns(5)) + 5).NumberFormat = "0.00%"
    classWks.Columns("A:B").AutoFit
    classWks.Columns("D:F").AutoFit
    classWks.Range("C7").Sort key1:=classWks.Range("C7"), Order1:=xlDescending
    classWks.Range("A7:A" & Application.WorksheetFunction.CountA(classWks.Columns(1)) + 4).Interior.Color = RGB(210, 235, 244)
    classWks.Range("B7:B" & Application.WorksheetFunction.CountA(classWks.Columns(1)) + 4).Interior.Color = RGB(237, 231, 246)
    classWks.Range("G7:G" & Application.WorksheetFunction.CountA(classWks.Columns(1)) + 4).Interior.Color = RGB(232, 234, 246)
    classWks.Range("C7:C" & Application.WorksheetFunction.CountA(classWks.Columns(1)) + 4).Interior.Color = RGB(232, 245, 233)
    classWks.Range("A6:L6").Interior.Color = RGB(255, 255, 102)
    classWks.Range("A3:A4").Merge
    classWks.Range("A3").Value = "SEE ""DATA"" FOR INDIVIDUAL CLASSES, TEACHERS, AND PERCETANGE OF As, Bs, and Cs."
    classWks.Range("A3").VerticalAlignment = xlTop
    classWks.Range("A3").WrapText = True
    classWks.Range("A3").Font.Size = 10
    classWks.Range("A3,A5").Font.Bold = True
    classWks.Range("A1").Style = ActiveWorkbook.Styles("Heading 2")
    classWks.Range("A5").Value = "Stats From " & IIf(classSearchForm.oneYearButton.Value, (Year(Date) - 1) & " Only", "Both " & (Year(Date) - 1) & " And " & (Year(Date) - 2))
    
    teacherWks.Range("A1:G1").Merge
    teacherWks.Range("A1").Value = "Teacher Averages For " & classListBox.Value & " (Avg. Total GPA: " & Round(CDbl(totalGPA / personCount), 2) & ")"
    teacherWks.Range("A1").Font.Bold = True
    teacherWks.Range("A6").Value = "Teacher:"
    teacherWks.Range("B6").Value = "Classes:"
    teacherWks.Range("C6").Value = "GPA:"
    teacherWks.Range("D6").Value = "N_Enroll:"
    teacherWks.Range("E6").Value = "N_Classes:"
    teacherWks.Range("F6").Value = "Workload avg:"
    teacherWks.Range("G6").Value = "PCT_A:"
    teacherWks.Range("H6").Value = "PCT_B:"
    teacherWks.Range("I6").Value = "PCT_C:"
    teacherWks.Range("J6").Value = "PCT_D:"
    teacherWks.Range("K6").Value = "PCT_F:"
    teacherWks.Range("L6:J6").Font.Bold = True
    For i = 0 To (UBound(teacherGPAKeys) - 1)
        If teacherGPA.Item(teacherGPAKeys(i)) > highestTeacherGPA Then
            highestTeacher = teacherGPAKeys(i)
            highestTeacherGPA = teacherGPA.Item(teacherGPAKeys(i))
        ElseIf teacherGPA.Item(teacherGPAKeys(i)) < lowestTeacherGPA And teacherGPA.Item(teacherGPAKeys(i)) > 0 Then
            lowestTeacher = teacherGPAKeys(i)
            lowestTeacherGPA = teacherGPA.Item(teacherGPAKeys(i))
        End If
        teacherWks.Cells(i + 7, 1).Value = teacherGPAKeys(i)
        teacherWks.Cells(i + 7, 2).Value = teacherGPA.Item(teacherGPAKeys(i) & "_classes")
        teacherWks.Cells(i + 7, 3).Value = Round(teacherGPA.Item(teacherGPAKeys(i)), 2)
        teacherWks.Cells(i + 7, 4).Value = teacherGPA.Item(teacherGPAKeys(i) & "_count")
        teacherWks.Cells(i + 7, 5).Value = teacherGPA.Item(teacherGPAKeys(i) & "_classCount")
        teacherWks.Cells(i + 7, 6).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_workload"), 3)
        teacherWks.Cells(i + 7, 7).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_a"), 3)
        teacherWks.Cells(i + 7, 8).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_b"), 3)
        teacherWks.Cells(i + 7, 9).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_c"), 3)
        teacherWks.Cells(i + 7, 10).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_d"), 3)
        teacherWks.Cells(i + 7, 11).Value = Round(teacherGPA.Item(teacherGPAKeys(i) & "_f"), 3)
    Next i
    teacherWks.Range("G7:K" & Application.WorksheetFunction.CountA(dataWks.Columns(5)) + 5).NumberFormat = "0.00%"
    teacherWks.Columns("A").AutoFit
    teacherWks.Columns("D:F").AutoFit
    teacherWks.Range("C7").Sort key1:=teacherWks.Range("C7"), Order1:=xlDescending
    teacherWks.Range("A7:A" & Application.WorksheetFunction.CountA(teacherWks.Columns(1)) + 4).Interior.Color = RGB(237, 231, 246)
    teacherWks.Range("F7:F" & Application.WorksheetFunction.CountA(teacherWks.Columns(1)) + 4).Interior.Color = RGB(232, 234, 246)
    teacherWks.Range("C7:C" & Application.WorksheetFunction.CountA(teacherWks.Columns(1)) + 4).Interior.Color = RGB(232, 245, 233)
    teacherWks.Range("A6:K6").Interior.Color = RGB(255, 255, 102)
    teacherWks.Range("A3:A4").Merge
    teacherWks.Range("A3").Value = "SEE ""DATA"" FOR INDIVIDUAL CLASSES, TEACHERS, AND PERCETANGE OF As, Bs, and Cs."
    teacherWks.Range("A3").VerticalAlignment = xlTop
    teacherWks.Range("A3").WrapText = True
    teacherWks.Range("A3").Font.Size = 10
    teacherWks.Range("A3,A5").Font.Bold = True
    teacherWks.Range("A1").Style = ActiveWorkbook.Styles("Heading 2")
    teacherWks.Range("A5").Value = "Stats From " & IIf(classSearchForm.oneYearButton.Value, (Year(Date) - 1) & " Only", "Both " & (Year(Date) - 1) & " And " & (Year(Date) - 2))
    classWks.Activate
    Unload Me
End Sub

Private Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (Err.Number = 0)
    Err.Clear
End Function

Private Function createClassSheet()
    Dim i As Integer, ws As Worksheet
    For i = 0 To classListBox.ListCount - 1
        If Evaluate("ISREF('" & Left("Class Avg. " & classListBox.List(i), 31) & "'!A1)") = True Then
            Call deleteSheet(Left("Class Avg. " & classListBox.List(i), 31))
        End If
    Next i
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count - 2))
    End With
    ws.name = Left("Class Avg. " & classListBox.Value, 31)
    Set createClassSheet = ws
End Function

Private Function createTeacherSheet()
    Dim i As Integer, ws As Worksheet
    For i = 0 To classListBox.ListCount - 1
        If Evaluate("ISREF('" & Left("Teacher Avg. " & classListBox.List(i), 31) & "'!A1)") = True Then
            Call deleteSheet(Left("Teacher Avg. " & classListBox.List(i), 31))
        End If
    Next i
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count - 2))
    End With
    ws.name = Left("Teacher Avg. " & classListBox.Value, 31)
    Set createTeacherSheet = ws
End Function

