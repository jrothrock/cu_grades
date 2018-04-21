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

Public Sub Main()
    'references required:
    'Microsoft HTML Object Library
    'Microsoft Internet Controls
    'Microsoft VBScript Regular Expressions 1.0 & 5.5
    Dim ws As Worksheet
    Call deleteSheet("Requirements")
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    ws.Name = "Requirements"
    ws.Activate
    Call getRequirements
End Sub

Sub getRequirements()

    Dim strUrl As String
    Dim objIe As InternetExplorer
    Dim objHtml As HTMLDocument
    Dim strHtml As String
    Dim objDivs As Variant
    Dim objAnchors As IHTMLElementCollection
    Dim intCounter As Integer
    Dim links As Variant
    Dim i As Integer, activeSheetCounter As Integer
    links = Array("https://www.colorado.edu/artsandsciences/human-diversity", "https://www.colorado.edu/artsandsciences/current-students/core-curriculum/ideals-and-values", "https://www.colorado.edu/artsandsciences/historical-context", "https://www.colorado.edu/artsandsciences/literature-and-arts", "https://www.colorado.edu/artsandsciences/united-states-context", "https://www.colorado.edu/artsandsciences/contemporary-societies", "https://www.colorado.edu/artsandsciences/natural-science")
    'set target to scrape
    Debug.Print "here"
    Debug.Print UBound(links) - LBound(links) + 1
    For i = 0 To (UBound(links) - LBound(links))
        'Debug.Print "here2"
        activeSheetCounter = 0
        strUrl = links(i)
        'get html from page
        Set objIe = New InternetExplorer
        objIe.Visible = False
        objIe.navigate strUrl
        While objIe.readyState <> READYSTATE_COMPLETE
            DoEvents
        Wend

        'assign html to DOM document
        Set objHtml = New HTMLDocument
        Set objHtml = objIe.document
        Set objDivs = objIe.document.getElementsByTagName("LI")
        
        Debug.Print objDivs.Length
        Application.ActiveSheet.Cells(1, 1 + i).Value = objIe.document.getElementById("page-title").innerText
        Dim regEx As New RegExp, strPattern As String
        strPattern = "<strong>"
    
        With regEx
            .Global = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With
    
        If objDivs.Length > 0 Then
            For intCounter = 0 To objDivs.Length - 1
                If regEx.test(objDivs(intCounter).innerHTML) Then
                    Application.ActiveSheet.Cells(2 + activeSheetCounter, 1 + i).Value = objDivs(intCounter).innerText
                    activeSheetCounter = activeSheetCounter + 1
                End If
            Next intCounter
         End If

        'clean up
        Set objHtml = Nothing
        objIe.Quit
        Set objIe = Nothing
    Next i
    ActiveSheet.Range("A1:H1").Columns.AutoFit
End Sub

Public Sub deleteSheet(shtName As String)
    Dim xWs As Worksheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name = shtName Then
            xWs.Delete
        End If
    Next
End Sub

Public Sub mainGrades()
    Dim targetFile As String, ws As Worksheet
    Call deleteSheet("Data")
    targetFile = zipTarget()
    Debug.Print "before"
    Call DownloadFile(targetFile)
    Call writeData(targetFile)
End Sub

Public Sub writeData(target As String)
    Dim gradesWkb As Workbook, destWkb As Workbook, projectSheet, wrksheet As Worksheet
   ' app.Visible = False
   Debug.Print "here"
    Set gradesWkb = Workbooks.Open(target, True, True)
    Set wrksheet = gradesWkb.Worksheets("Data")
    Set destWkb = Application.Workbooks("Project.xlsm")
    Debug.Print wrksheet.Name
    Debug.Print destWkb.Worksheets.Count
    '' One of the many mysteries of vba
    wrksheet.Copy After:=destWkb.Worksheets(1)
    gradesWkb.Close
    Set gradesWkb = Nothing
End Sub

Public Function zipTarget()
    Dim targetFolder As Variant, targetFileTXT As Variant
    targetFolder = Environ("TEMP") & "\" & RandomString(8) & "\"
    MkDir targetFolder
    zipTarget = targetFolder & "cu_grades_data.csv"
End Function

Public Sub DownloadFile(targetDest)
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
    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function
