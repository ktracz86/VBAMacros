Global FileName As Variant ' Holds the File Name
Global PathName As String ' Holds the Path Name
Global eCount As Integer  ' Holds the number of errors
Global logId As String
Global dataSheet As String
Global RequiredSheet As Excel.Worksheet
Global MasterWorkbook As Excel.Workbook
Global errorText As String




Sub XMLVerification()
    
    Set MasterWorkbook = ActiveWorkbook
    Set RequiredSheet = MasterWorkbook.Sheets("SelectRequired")
    
    GetXML
        If FileName = False Then
            Exit Sub
        End If
    CreateWorksheet
    Verification
    DacoChecker
    
    If eCount = 0 Then
        MsgBox ("There were " & recordCount & " records, and no errors were found")
    ElseIf eCount > 0 Then
        MsgBox ("There were " & recordCount & " records, and " & eCount & " errors were logged. Log file has been saved to " & PathName)
    End If

    If eCount > 0 Then
        SetLogID
        ExportLog
        ExportEmailText
        'SendEmail
    End If
    CleanUp

End Sub

Private Sub GetXML()

    PathName = ActiveWorkbook.Path


    ChDrive (Mid(PathName, 1, 1))
    ChDir (PathName)
    With Application
        ' Set File Name to selected File
        FileName = .GetOpenFilename("XML Files (*.xml), *.xml")
        
        ' Reset Start Drive/Path
        ChDrive (Left(.DefaultFilePath, 1))
        ChDir (.DefaultFilePath)
    End With
    
    ' Exit on Cancel
    If FileName = False Then
        MsgBox "No file was selected. Code can not continue."
        Exit Sub
    End If
    
    ' Open File
    Workbooks.OpenXML FileName, LoadOption:=xlXmlLoadImportToList

End Sub

Private Sub CreateWorksheet()
' Creates a new sheet called "ErrorLog".
    dataSheet = ActiveSheet.Name
    Sheets.Add.Name = "ErrorLog"
        
End Sub

Private Sub Verification()

Dim selectedCell As String
Dim recordCount As Integer
Dim id2 As Range
Dim dataRow As Integer
Dim i As Integer
Dim cbiValue As Variant
Dim titleValue As Variant
Dim authorValue As Variant
Dim reportDate As Variant

Worksheets("Sheet1").Select
ActiveSheet.Range("D8").Select
eCount = 0

' MsgBoxes have been commented out. They were used for testing purposes

recordCount = 0
Do While ActiveCell.Value <> Empty
    i = 0
    Set id2 = ActiveCell
    dataRowLength = Len(id2.Value)
    dataRow = CInt(Right(id2.Value, dataRowLength - 1)) + 1

    recordCount = recordCount + 1

    selectedCell = ActiveCell.Offset(0, 2).Value
    
    If Left(selectedCell, 2) = "0." Then 'If the selected cell is correspondence
        ' Find next id2
        Do
        i = i + 1
            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                Exit Do
            End If
        Loop
        'MsgBox("This is a DACO 0.*")
        
    ElseIf Left(selectedCell, 2) = "1." Then
        ' Find next id2
        Do
        i = i + 1
            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                Exit Do
            End If
        Loop
        'MsgBox ("This is a Label")
        
    Else
        'Do the Verification
        'Find CBI
        Do
        i = i + 1
            If Cells(ActiveCell.Row + i, ActiveCell.Column + 1).Value = "CBI_APPL_IND" Then
                Cells(ActiveCell.Row + i, ActiveCell.Column + 1).Select
                Exit Do
            End If
        Loop
        
        cbiValue = ActiveCell.Offset(0, 1).Value
        
        
        If IsEmpty(cbiValue) Then
            'MsgBox ("Missing CBI Information on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Missing CBI Information"
        End If
        
                
        'Find Title
        i = 0
        Do
        i = i + 1
            If Cells(ActiveCell.Row + i, ActiveCell.Column).Value = "DM_TITLE" Then
                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                Exit Do
            End If
        Loop
        titleValue = ActiveCell.Offset(0, 1).Value
        
        If IsEmpty(titleValue) Then
            'MsgBox ("Missing Title on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Missing Title. Title cannot be blank"
        ElseIf Len(titleValue) < 4 Then
            'MsgBox ("Invalid Title on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Title. Title should be descriptive."
        ElseIf InStr(1, titleValue, "not applicable", vbTextCompare) Then
            'MsgBox ("Invalid Title on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Title. Title cannot be 'Not Applicable'"
        End If
        
        
            
        'Find Author
        i = 0
        Do
        i = i + 1
            If Cells(ActiveCell.Row + i, ActiveCell.Column).Value = "DM_AUTHOR" Then
                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                Exit Do
            End If
        Loop
        
        authorValue = ActiveCell.Offset(0, 1).Value
        
        If IsEmpty(authorValue) Then
            'MsgBox ("Missing Author on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Missing Author."
        ElseIf Len(authorValue) < 4 Then
            'MsgBox ("Invalid Author on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Author."
        ElseIf InStr(1, authorValue, "not applicable", vbTextCompare) Then
            'MsgBox ("Invalid Author on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Author. Author cannot be 'Not Applicable'"
        End If
        
        
        
        
        'Find Report Date
        i = 0
        Do
        i = i + 1
            If Cells(ActiveCell.Row + i, ActiveCell.Column).Value = "DM_REPORT_DATE" Then
                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                Exit Do
            End If
        Loop
        
        reportDate = ActiveCell.Offset(0, 1).Value
        
        If IsEmpty(reportDate) Then
            'MsgBox ("Missing Report Date on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Missing Report Date"
        ElseIf Len(reportDate) < 4 Then
            'MsgBox ("Invalid Report Date on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Report Date. Date too short."
        ElseIf InStr(1, reportDate, "not applicable", vbTextCompare) Then
            'MsgBox ("Invalid Report Date on Row " & dataRow)
            eCount = eCount + 1
            Worksheets("ErrorLog").Cells(eCount, 1).Value = "Row " & dataRow & " - Invalid Report Date. Date cannot be 'Not Applicable'"
        End If
        
        
        'Find next id2
        Do
            i = i + 1
            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                Cells(id2.Row + i, id2.Column).Select
                Exit Do
            End If
        Loop
        'MsgBox ("Legit data")
        
    End If


    

Loop


End Sub
Public Sub DacoChecker()
Dim criteria As Integer
Dim req As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False

For criteria = 1 To 8
    Select Case criteria
        Case 1 'Search for Fee Form
            req = RequiredSheet.Cells(1, 2).Value
            If req = "NR" Then
                GoTo OutCase
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "Fee Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Fee Form might be required. Please check."
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "Fee Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Fee Form Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required Fee Form (DACO " & req & ") was not found."
            End If
                     
            

        Case 2 'Search for Cover Letter
            req = RequiredSheet.Cells(2, 2).Value
            If req = "NR" Then
                'Do nothing
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Or selectedCell = "0.8.1" Then
                        MsgBox "Cover Letter Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Cover Letter might be required. Please check."
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Or selectedCell = "0.8.1" Then
                        MsgBox "Cover Letter Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Cover Letter Not Found."
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required Cover Letter (DACO 0.8.1 or DACO " & req & ") was not found."
            End If
        
        Case 3 'Search for Application Form
            req = RequiredSheet.Cells(3, 2).Value
            If req = "NR" Then
                'Do nothing
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value

                    If selectedCell = "0.1.6000" Or selectedCell = "0.1.6006" Or selectedCell = "0.1.6008" Or selectedCell = "0.1.6012" Or selectedCell = "0.1.6110" Or selectedCell = "0.1.6117" Or selectedCell = "0.1.6301" Or selectedCell = req Then
                            MsgBox "Application Form Found"
                            GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Application Form might be required. Please check."

            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "Application Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Application Form Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required Application Form (DACO " & req & ") was not found."
            End If
        
        Case 4 'Search for Spec Form
            req = RequiredSheet.Cells(4, 2).Value
            If req = "NR" Then
                'Do nothing
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "Spec Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Spec Form might be required. Please check."
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "Spec Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "Spec Form Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required SPSF Form (DACO " & req & ") was not found."
            End If
        
        Case 5 'Search for LOC
            req = RequiredSheet.Cells(5, 2).Value
            If req = "NR" Then
                'Do nothing
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "LOC Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "LOC might be required. Please check."
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "LOC Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "LOC (DACO " & req & ") Not Found. Not required for TGAI, and not required if current LOC is less than 5 years. Please check."
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required Letter of Confirmation of Source of Supply (DACO " & req & ") was not found. Not required for TGAI or for EPs with LOC less than 5 years old. Please check."
            End If
        
        Case 6 'Search for English Label
            req = RequiredSheet.Cells(6, 2).Value
            If req = "NR" Then
                'Do nothing
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = "1.1.1" Then
                        MsgBox "English Label Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "English Label might be required"
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = "1.1.1" Then
                        MsgBox "English Label Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "English Label Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required English Label (DACO 1.1.1) was not found."
            End If
        
        Case 7 'Search for French Label
            req = RequiredSheet.Cells(6, 2).Value
            If req = "NR" Then
                GoTo OutCase
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = "1.1.2" Then
                        MsgBox "French Label Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "French Label might be required"
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = "1.1.2" Then
                        MsgBox "French Label Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "French Label Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required French Label (DACO 1.1.2) was not found."
            End If
        
        Case 8 'Search for New Use Form
            req = RequiredSheet.Cells(7, 2).Value
            If req = "NR" Then
                GoTo OutCase
            ElseIf req = "CR" Then
                'Check for all possible DACOs
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = "0.1.6023" Then
                        MsgBox "New Uses Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "New Uses Form might be required. Please check."
            Else
                'Check for req
                ActiveSheet.Range("D8").Select
                Do While ActiveCell.Value <> Empty
                    i = 0
                    Set id2 = ActiveCell
                    selectedCell = ActiveCell.Offset(0, 2).Value
                    If selectedCell = req Then
                        MsgBox "New Uses Form Found"
                        GoTo OutCase
                    Else 'look at next id2
                        Do
                        i = i + 1
                            If Cells(id2.Row + i, id2.Column).Value <> id2.Value Then
                                Cells(ActiveCell.Row + i, ActiveCell.Column).Select
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                MsgBox "New Uses Form Not Found"
                eCount = eCount + 1
                Worksheets("ErrorLog").Cells(eCount, 1).Value = "Required New Uses Form (DACO " & req & ") was not found."
            End If
    End Select
OutCase:
Next criteria
        
        
Application.DisplayAlerts = True
Application.ScreenUpdating = True



End Sub

Private Sub ExportLog()

Dim wb As Workbook
Dim WorkRng As Range
On Error Resume Next
Dim log As Range
Dim OutApp As Object
Dim OutMail As Object
Dim iFile As Integer




'Path was set manually here for debugging this private Sub.
'PathName = "C:\Users\KTRACZ\Desktop"

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Worksheets("ErrorLog").Activate
Worksheets("ErrorLog").Cells.Select
Set WorkRng = Application.Selection
errorText = WorkRng.Value
Set wb = Application.Workbooks.Add
WorkRng.Copy
wb.Worksheets(1).Paste
wb.SaveAs FileName:=PathName & "\ErrorLog " & logId & ".txt", FileFormat:=xlText
wb.Close

iFile = FreeFile

Open PathName & "\ErrorLog " & logId & ".txt" For Input As #iFile
errorText = Input(LOF(iFile), iFile)
Close #iFile

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
Private Sub SendEmail()


Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

With OutMail
    .To = "APPLICANT ADDRESS"
    .Subject = "Error Log"
    '.Body = "We found the following errors" & Chr(10) & Chr(10) & errorText 'THIS WAS THE OLD LINE
     .Body = "Your Category (INSERT - A,B or C) application for a new registration of (INSERT PRODUCT NAME) has been received by the PMRA. " _
        & vbCrLf & vbCrLf & _
        "In the initial screening of your submission package, it has been noted that the following form(s) was / were not included and is / are a requirement:" _
        & vbCrLf & "(INCLUDE ONLY THE FORMS MISSING)" _
        & vbCrLf & vbCrLf & errorText _
        & vbCrLf _
        & vbCrLf _
        & "As a result, your application package has been placed on hold under submission no. 201X-XXXX." _
        & vbCrLf & vbCrLf _
        & "In order for your submission to proceed to the next level, all of the above missing documents must be received. Please SUBMIT only the missing documents WITHIN 14 DAYS OF RECEIVING THIS EMAIL. Failure to do so will result in the rejection of your submission as per DIR2017-01 Revised Management of Submissions Policy, section 5.1.1." _
        & vbCrLf & vbCrLf _
        & "Updated Application and Fee forms are required to accompany all applications received on or after April 1, 2017 resulting from the coming into force of the Pest Control Products Fees and Charges Regulations. Additional information is available on the PMRA's cost recovery webpage: http://www.hc-sc.gc.ca/cps-spc/pest/registrant-titulaire/prod/cost-cout-eng.php" _
        & vbCrLf & vbCrLf _
        & "Please send the missing documents in PRZ format and reference submission number 201X-XXXX." _
        & vbCrLf & vbCrLf _
        & "If you have any questions regarding the above forms, please refer to the PMRA website: http://www.hc-sc.gc.ca/cps-spc/pest/registrant-titulaire/form/index-eng.php" _
        & vbCrLf & vbCrLf _
        & "Or you may contact the PMRA Information Service at pmra.inforserv@hc-sc.gc.ca" _
        & vbCrLf & vbCrLf _
        & "Thank you for your cooperation."
    '.Attachments.Add (PathName & "\ErrorLog " & logId & ".txt")
    .Display
End With

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub CleanUp()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Worksheets(dataSheet).Delete
RequiredSheet.Delete
ActiveWorkbook.Close

Application.DisplayAlerts = True
Application.ScreenUpdating = True

PathName = Empty
FileName = Empty
dataSheet = Empty
logId = Empty



End Sub

Private Sub SetLogID()

logId = Date & " " & Time
logId = Replace(logId, "-", "")
logId = Replace(logId, ":", "")
logId = Left(logId, Len(logId) - 3)

End Sub

Private Sub ExportEmailText()
Dim fso As Object
Set fso = CreateObject("Scripting.FilesystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(PathName & "\EmailText " & logId & ".txt", True, True)
'Set oFile = fso.CreateTextFile("C:\Users\KTRACZ\Desktop\XML Verification\XML Verification with DACO Check\\Test Changes\EmailText.txt", True, True) ' TO TEST
oFile.WriteLine "Your Category (INSERT - A,B or C) application for a new registration of (INSERT PRODUCT NAME) has been received by the PMRA. " _
& vbCrLf & vbCrLf & _
"In the initial screening of your submission package, it has been noted that the following form(s) was / were not included and is / are a requirement:" _
& vbCrLf & "(INCLUDE ONLY THE FORMS MISSING)" _
& vbCrLf & vbCrLf & errorText _
& vbCrLf _
& vbCrLf _
& "As a result, your application package has been placed on hold under submission no. 201X-XXXX." _
& vbCrLf & vbCrLf _
& "In order for your submission to proceed to the next level, all of the above missing documents must be received. Please SUBMIT only the missing documents WITHIN 14 DAYS OF RECEIVING THIS EMAIL. Failure to do so will result in the rejection of your submission as per DIR2017-01 Revised Management of Submissions Policy, section 5.1.1." _
& vbCrLf & vbCrLf _
& "Updated Application and Fee forms are required to accompany all applications received on or after April 1, 2017 resulting from the coming into force of the Pest Control Products Fees and Charges Regulations. Additional information is available on the PMRA's cost recovery webpage: http://www.hc-sc.gc.ca/cps-spc/pest/registrant-titulaire/prod/cost-cout-eng.php" _
& vbCrLf & vbCrLf _
& "Please send the missing documents in PRZ format and reference submission number 201X-XXXX." _
& vbCrLf & vbCrLf _
& "If you have any questions regarding the above forms, please refer to the PMRA website: http://www.hc-sc.gc.ca/cps-spc/pest/registrant-titulaire/form/index-eng.php" _
& vbCrLf & vbCrLf _
& "Or you may contact the PMRA Information Service at pmra.inforserv@hc-sc.gc.ca" _
& vbCrLf & vbCrLf _
& "Thank you for your cooperation."

oFile.Close
Set fso = Nothing
Set oFile = Nothing

End Sub

