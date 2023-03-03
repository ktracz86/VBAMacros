Option Explicit
Option Compare Text

Sub HyperlinkPMRA()
'This sub allows the user to choose a column of PMRA numbers and it will automatically turn them into links
 Dim RowCount, i, headers, s As Integer
 Dim link As String
 Dim c As Variant
 
 ' The prefix for the URL is "http://pmra-pw1.hc-sc.gc.ca:7777/ePRS/dox_web.v?p_ukid="&[PMRANUMBER]
 link = "http://pmra-pw1.hc-sc.gc.ca:7777/ePRS/dox_web.v?p_ukid="
 
 'User selects which column the PMRA number is
 c = InputBox("Which column contains the PMRA Numbers?", "Select Column", "Enter")
 
 'Asking officer if there's are headers. If there are, then row count will start at row 2. Otherwise, we'll start at Row 1.
 
 headers = MsgBox("Are there headers?", vbQuestion + vbYesNo, "Headers?")
 If headers = vbYes Then
    s = 2
 Else
    s = 1
 End If
 
 RowCount = WorksheetFunction.CountA(Range(c & ":" & c)) 'find the amount of rows in the sheet
 
 For i = s To RowCount
    If Range(c & i).Value > 1 Then
        Application.ActiveSheet.Hyperlinks.Add Anchor:=Range(c & i), Address:=link & Range(c & i).Value
    End If
Next i

End Sub

Sub CommaSeparate()
    Dim sourceColumn As Variant
    Dim i As Long
    Dim result As String
    Dim dest As Range
    Dim RowCount As Long
       
' User inputs the column here
    sourceColumn = InputBox("Which column to you want to use?", _
                    "Select Column", "Enter")
                    
' Goes down the colum and writes to the variable
    RowCount = WorksheetFunction.CountA(Range("A:A"))
    For i = 1 To RowCount
        Range(sourceColumn & i).Select
        result = result & Range(sourceColumn & i).Text + ","
    Next i
    
'Places variable into the clipboard
    Set dest = Application.InputBox("Select a cell to place result.", "Select destination", , , , , , 8)
    dest.Value = result
    'clip.SetText result
    'clip.PutInClipboard
    MsgBox "Done", vbInformation
    
    
'Selects the A1 cell to return to the top
    Range("A1").Select

End Sub

Sub HighlightDifferentRows()


    Dim firstValue As Range
    Dim secondValue As Range
    Dim userInput As String
    Dim userRangeStart As Variant
    Dim colour As Long
    
   
    userInput = InputBox("What column contains the unique values?", _
                            "Select Range")
                            
    userRangeStart = userInput & "2"
        
    Range(userRangeStart).Select
    Set firstValue = Range(userRangeStart)
 
    Do
    
    firstValue.Select
    Set secondValue = firstValue.Offset(1, 0)
    
    '*** Exit the loop if the second value is empty
    If IsEmpty(secondValue) Then Exit Do
    '***
    
    If firstValue = secondValue Then
        'sets the second cell row colour to match the first
        colour = firstValue.Interior.ColorIndex
        secondValue.Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Interior.ColorIndex = colour
        secondValue.Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.Interior.ColorIndex = colour
        
    ElseIf firstValue.Interior.ColorIndex = -4142 Then
        ' If the first cell has no fill, it puts a blue fill in the second cell
        secondValue.Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Interior.ColorIndex = 37
        secondValue.Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.Interior.ColorIndex = 37
        
    Else
        'If the first value is not blank, it sets the second value to be blank
        secondValue.Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Interior.ColorIndex = -4142
        secondValue.Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.Interior.ColorIndex = -4142
    End If

    Set firstValue = secondValue
    Loop
End Sub

Sub FindNeonic()



' User inputs the column which contains the three letter code for
' actives, and it highlights the cells that contain a neonic.

    Dim activesColumnInput As Variant
    Dim i As Long
       
' User inputs the column here
    activesColumnInput = InputBox("Which column contains the 3-letter active codes?", _
                    "Actives Column", "Enter")
                    
' Goes down the colum and check for a neonic, exits the loop
' if the cell is empty

    For i = 1 To Rows.Count
        Range(activesColumnInput & i).Select
        If InStr(1, ActiveCell, "THE") Then
            ActiveCell.Interior.Color = vbRed
        ElseIf InStr(1, ActiveCell, "COD") Then
            ActiveCell.Interior.Color = vbRed
        ElseIf InStr(1, ActiveCell, "IMI") Then
            ActiveCell.Interior.Color = vbRed
        ElseIf IsEmpty(ActiveCell) Then Exit For
        End If
    Next i
    
'Selects the A1 cell to return to the top
    Range("A1").Select
    
End Sub

Sub FormatTableStandard()
'
' FormatTableStandard Macro

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
      
    If MsgBox("Freeze Top Row?", vbYesNo, "Freezing Panes") = vbYes Then
        Range("A1").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        'Rows("1:1").Select
        ActiveWindow.FreezePanes = True
    Else
    End If
       
    
End Sub

Sub BulkReplace()
' x=1
'Loop
' Rx= wordToReplace
' Sx= replacement
'if wordToReplace = blank = exit loop
'Column D Replace wordtoReplace with replacement
'x = x+1

Dim sht As Worksheet
Dim fndList As Integer
Dim rplcList As Integer
Dim tbl As ListObject
Dim myArray As Variant
Dim TempArray As Variant
Dim x As Integer


'Create variable to point to your table
  Set tbl = Worksheets("Sheet1").ListObjects("Table1")
  

'Create an Array out of the Table's Data
  Set TempArray = tbl.DataBodyRange
  myArray = Application.Transpose(TempArray)
  
'Designate Columns for Find/Replace data
  fndList = 1
  rplcList = 2

'Loop through each item in Array lists
  For x = LBound(myArray, 1) To UBound(myArray, 2)
    'Loop through each worksheet in ActiveWorkbook (skip sheet with table in it)
      For Each sht In ActiveWorkbook.Worksheets
        If sht.Name <> tbl.Parent.Name Then
          
          sht.Cells.Replace What:=myArray(fndList, x), Replacement:=myArray(rplcList, x), _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
        
        End If
      Next sht
  Next x




End Sub
Sub GetFileNamesDate()
Dim xRow As Long
Dim xDirect$, xFname$, InitialFoldr$
Dim i As Integer
Dim cell As Range, fso As Object
Dim headers() As Variant


Sheets.Add After:=Sheets(Sheets.Count) 'Adds a new sheet

headers() = Array("Filename", "Date Last Modified", "Size KB") 'identifies the headers

Set fso = CreateObject("Scripting.FileSystemObject") 'creates a files system object that you need to get file attributes

InitialFoldr$ = "C:\"   'sets the first lookup folder a the C: Drive

With Application.FileDialog(msoFileDialogFolderPicker)   'Pops up the dialog to select a folder
.InitialFileName = Application.DefaultFilePath & "\"
.Title = "Please select a folder to list Files from"
.InitialFileName = InitialFoldr$
.Show

If .SelectedItems.Count <> 0 Then        'if a folder is selected, it writes each file name, date, and size to a row
    xDirect$ = .SelectedItems(1) & "\"
    xFname$ = Dir(xDirect$, 7)
    Do While xFname$ <> ""
        ActiveCell.Offset(xRow) = xFname$
        ActiveCell.Offset(xRow, 1).Value = Left(fso.getfile(xDirect$ & xFname$).DateLastModified, 10)
        ActiveCell.Offset(xRow, 2).Value = (fso.getfile(xDirect$ & xFname$).Size) / 1000
        xRow = xRow + 1
        xFname$ = Dir
    Loop
End If
End With

'Adds a new row and sets the headers
Rows("1:1").Select
Selection.Insert
i = 0
For i = 0 To 2
Cells(1, i + 1).Value = headers(i)
Next i

Set fso = Nothing
On Error GoTo 0
End Sub

Sub CreateHyperlinkLabelSearch()
 Dim RowCount, i As Integer
 Dim linkAddressStart, linkAddressEnd As String
 

RowCount = WorksheetFunction.CountA(Range("A:A")) 'find the amount of rows in the sheet
i = 2
linkAddressStart = "https://pr-rp.hc-sc.gc.ca/ls-re/result-eng.php?p_search_label="
linkAddressEnd = "&searchfield1=NONE&operator1=CONTAIN&criteria1=&logicfield1=AND&searchfield2=NONE&operator2=CONTAIN&criteria2=&logicfield2=AND&searchfield3=NONE&operator3=CONTAIN&criteria3=&logicfield3=AND&searchfield4=NONE&operator4=CONTAIN&criteria4=&logicfield4=AND&p_operatordate=%3D&p_criteriadate=&p_searchexpdate=EXP"

For i = 2 To RowCount
    If Len(Cells(i, 1).Value) > 1 Then
        'hyperlink it
        Application.ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 1), Address:=linkAddressStart & Cells(i, 1).Value & linkAddressEnd
    End If
Next i

End Sub

