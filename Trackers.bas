' -----------------------------------------------------------------------------
' Module:  Tracking_Sheets.bas
' Author:  Susan Oda (susan.oda@hc-sc.gc.ca)
' Date:    13-Oct-2011
' UPDATED: 31-Aug-2016 by Kamil Tracz
'          31-Oct-2019 by Kamil Tracz
'           - Removed the 600dpi Print Quality formatting
'           - Updated iCal feed location
'           - Added an additional exit in the DO Loop in InsertMonths Sub
'           - Updated and included flagging of Neonics and Cost Recovery submissions
'           - Added a Status bar to inform user what the script is doing
'
' Description:
'
' This creates the weekly tracking sheets used in the Submission Coordination
' Section as well as those sent to all divisions.  All code in this module
' should be commented on and made as clear as possible for the next person who
' may use the code.  The goal is for these sheets to be as close as possible
' to "one-click generation" to avoid confusion.
'
' Target audience:  Administrative Coordinators in the Submission Coordination
' Section who do not necessarily have any experience with computer code.
'
' -----------------------------------------------------------------------------
' To use this code, please do the following steps:
' 1.  Download the most recent snapshot from the e-PRS beta system in the
'     standard format (see written instructions)
' 2.  Open the previous week's report for the same category
' 3.  Run "TrackingSheets" and select the file with the new data when prompted.
'     At the end of this step, all your sheets should be generated.
' 4.  Manually remove highlighting in the Notes column for submissions which
'     have been updated recently (in the last month).
' 5.  If it is a Trackers Week, manually set the page breaks as appropriate.
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Global Variables
' -----------------------------------------------------------------------------

Global CategoryName As String ' Holds the Category Name (e.g., Category A)
Global SnapshotDate As Date   ' Holds the Snapshot Date (previous Sunday)
Global FileName As String ' Holds the File Name (of the file being generated)
Global PathName As String ' Holds the Path Name (of the file initially opened)

' -----------------------------------------------------------------------------
' Subroutines
' -----------------------------------------------------------------------------

Sub TrackingSheets()

    ' -------------------------------------------------------------------------
    ' TrackingSheets() sets the Global variables and then runs everything else.
    ' -------------------------------------------------------------------------

    Application.DisplayAlerts = False

    ' Gets the category name (sheet name) from the current sheet and sets as a
    ' Global variable.  Also gets the Path of the current file as a Global var.
    Application.StatusBar = "Setting File path, Category Name, and Snapshot Date"
    CategoryName = ActiveSheet.Name
    PathName = ActiveWorkbook.Path
    
    ' Finds the Snapshot date (the previous Sunday) and sets Global var.
    
    For i = 0 To 6
        If Weekday(Now) - i = 1 Then
            SnapshotDate = DateValue(Now) - i
        End If
    Next i
            
    ' This sets SnapshotDay as a string that contains a well-formatted day, in
    ' order to ensure that single-digit dates are preceded by a zero.
    
    If Day(SnapshotDate) < 10 Then
        SnapshotDay = "0" & Day(SnapshotDate)
    Else
        SnapshotDay = "" & Day(SnapshotDate)
    End If
    
    ' Saves the file under a new filename, as per naming convention, using the
    ' Snapshot date.  Sets the new filename as a Global variable.
    Application.StatusBar = "Creating new file for Standard Sheet"
    ActiveWorkbook.SaveAs FileName:=PathName & "\" & CategoryName & "_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate)
    FileName = ActiveWorkbook.Name
    
    ' Moves the data to the "Old List" and makes a new sheet for the new data.
        
    ActiveSheet.Name = "Old List"
    Sheets.Add.Name = CategoryName
    Sheets(CategoryName).Select
    
    ' Opens an Open File dialog for the user to select the new data (CSV file).
    Application.StatusBar = "Requesting new Data"
    OpenSingleFile "Select the file which contains the new snapshot data:", ".csv"
    
    ' Selects all the data and moves it over to the main working sheet. Closes
    ' the temporary file containing the snapshot data, without displaying a
    ' dialog box.
    Application.StatusBar = "Copying new data"
    Cells.Select
    Selection.Copy
    TempFileName = ActiveWorkbook.Name ' TempFileName contains the snapshot filename.
    Windows(FileName).Activate
    ActiveSheet.Paste
    'Application.DisplayAlerts = False ' Turns off the automatic dialog box.
    Windows(TempFileName).Close
    
    ' These subroutines do the main work of generating the sheets.  See the
    ' commented descriptions or the full descriptions with each module.
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Removing irrelevant submissions"
    DeleteUnwanted ' Removes submissions which are not meant to appear in the final sheets.
    Application.StatusBar = "Sorting List by E Deadline"
    SortList ' Sorts the list by Level E deadline (followed by current deadline, if E is blank).
    Application.ScreenUpdating = True
    Application.StatusBar = "Basic Formatting"
    BasicFormat ' Formats the sheet in some basic ways (font, etc).
    Application.StatusBar = "Inserting Row Numbers and Setting Column Width"
    InsertRowNumbers ' Inserts row numbers, numerically increasing by one.
    SetColumns ' Sets the column widths.
    Application.ScreenUpdating = False
    Application.StatusBar = "Retrieving notes from previous list"
    GetNotes ' Searches the old list for each submission and retrieves the Notes column.
    Application.ScreenUpdating = True
    Application.StatusBar = "Adding PassNums and DNotRequired"
    AddPass ' Adds which Pass a submission is on (e.g., C2 or C3)
    DNotRequired ' Adds (D not req.) to the section statuses if applicable.
    Application.StatusBar = "Formating for Printing"
    PrintFormat ' Formats the sheet for printing.
    Application.StatusBar = "Removing extra columns that are no longer needed"
    RemoveExtra ' Removes extra columns that are no longer needed.
    Application.ScreenUpdating = False
    Application.StatusBar = "Grouping related subs"
    RelatedSubs ' Formats related submissions to appear grouped.
    Application.StatusBar = "Setting Notes column on new submissions"
    NewSubs ' Sets the Notes column text on new submissions.
    
    ' This section adds any additional information or flags that are required,
    ' such as which submissions REMD is assisting on and which submissions
    ' are pilot subs.  This information must be retrieved from a separate sheet
    ' in order to be automated.
    Application.StatusBar = "Requesting Additional Flags"
    Sheets.Add.Name = "Flags"
    OpenSingleFile "Select the file which contains the extra flags (pilots and REMD):", ".xls"
    Application.StatusBar = "Retrieving Additional flags"
    Cells.Select
    Selection.Copy
    TempFileName = ActiveWorkbook.Name
    Windows(FileName).Activate
    ActiveSheet.Paste
    'Application.DisplayAlerts = False
    Windows(TempFileName).Close
    FlagSubs
    Application.StatusBar = "Flagging status changes"
    StatusDate ' Moves and formats Status Date column & flags recent changed statuses.
    
    ' Final format to fix the page breaks to not appear between related subs.
    Application.StatusBar = "Setting Page Breaks"
    Application.ScreenUpdating = True
    SetPageBreaks
    
    ' Save it again.
    Application.StatusBar = "Saving Standard Sheet"
    ActiveWorkbook.SaveAs FileName:=PathName & "\" & CategoryName & "_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate)
    
    ' Save it as the Trackers Club version.
    Application.StatusBar = "Creating Trackers Sheet"
    ActiveWorkbook.SaveAs FileName:=PathName & "\" & CategoryName & "_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate) & " - Trackers Club"
    
    ' Make the changes to generate the Trackers Club sheet.
    
    ActiveSheet.ResetAllPageBreaks
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    Highlight
    If CategoryName = "Cat B" Then
        Application.ScreenUpdating = False
        OtherACs
        Application.ScreenUpdating = True
    End If
    Application.StatusBar = "Inserting Month Breaks"
    InsertMonths
    Application.StatusBar = "Flagging Neonic Subs"
    NeonicFlags 'Goes down the column of actives and highlights in red all cells that contain THE, COD, or IMI.
    Application.StatusBar = "Flagging Cost Recovery"
    CostRecovery 'Windows Dialog box requests the NewSubs list and will highlight matches.
    
    Application.StatusBar = "Saving Trackers Sheet"
    ActiveWorkbook.SaveAs FileName:=PathName & "\" & CategoryName & "_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate) & " - Trackers Club"
    
    If CategoryName = "Cat B" Then
        Application.StatusBar = "Creating AC STL Sheet"
        ActiveWorkbook.SaveAs FileName:=PathName & "\Cat A and B_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate)
        NewFileName = ActiveWorkbook.Name
        OpenSingleFile "Select the file which contains the Cat A Trackers Club list:", ".xls"
        Cells.Select
        Range("A2:AD" & Range("B2").End(xlDown).Row).Select
        Selection.Copy
        TempFileName = ActiveWorkbook.Name
        Windows(NewFileName).Activate
        Rows("2:2").Select
        Selection.Insert Shift:=xlDown
        'Application.DisplayAlerts = False
        Windows(TempFileName).Close
        Application.StatusBar = "Highlighting Subs that may need status update"
        MayNeedUpdate
        Application.StatusBar = "Setting up AC STL Sheet"
        ActiveSheet.Name = "Cat A and B"
        With ActiveSheet.PageSetup
            ' Header information.
            .CenterHeader = _
            "&""Arial,Bold""Cat A and B @ Levels C,D,E,F,G,H,R with Division Status&""Arial,Regular""" & Chr(10) & "&8(sorted by Level E Deadline)"
        End With
        Application.StatusBar = "Saving AC STL Sheet"
        ActiveWorkbook.SaveAs FileName:=PathName & "\Cat A and B and L_" & MonthName(Month(SnapshotDate), True) & "_" & SnapshotDay & "_" & Year(SnapshotDate)
        Application.StatusBar = "Creating Calendar Feeds"
        iCalFeed
        Application.StatusBar = "DONE"
        Application.ScreenUpdating = True
        Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
        MsgBox ("Remember, you need to go through the Cat A and B sheet and manually remove the Notes highlighting from submissions which do not require an update.  Only submissions which are highlighted that have not been updated in a month or more should remain highlighted.  Also set the page breaks in the Trackers Club sheets if they are being sent out this week.")
    End If
    
    Cells(1, 1).Select
    MsgBox ("Code complete.")
    Application.StatusBar = False
    
    ' That's all, folks.  The sheet you end up with at the end of this should
    ' is currently the original format sheet and the Trackers Club lists.
    ' Manual work is still required for the list for ACs and STLs.
    
End Sub

Private Sub OtherACs()

    ' -------------------------------------------------------------------------
    ' OtherACs finds all submissions that aren't with SCS and
    ' moves them to a new tab, which it then formats for printing.
    ' -------------------------------------------------------------------------

    Dim Sentences
    Dim i As Long
    
    ' Create a variable that has all entries in the AC name column.
    Sentences = Range("L1", "L" & Range("A2").End(xlDown).Row)

    ' Add a new sheet and copy the headers from the original sheet.
    Sheets.Add.Name = "PPIP PSR and Screening"
    AustenRows = 2
    Sheets(CategoryName).Select
    Range("A1").EntireRow.Select
    Selection.Copy
    Sheets("PPIP PSR and Screening").Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    Sheets(CategoryName).Select
    
    ' Find the last row in the original sheet which contains data, then search
    ' through the list of AC names and copy any "Austen" submissions to the
    ' new sheet.  Clear the rows from the original sheet after.
    EndRow = Range("A2").End(xlDown).Row
    
    For i = 2 To EndRow Step 1
        iWordPos = 1
        If Sentences(i, 1) = "" Or InStr(LCase(Sentences(i, 1)), LCase("Orhiobhe")) Or InStr(LCase(Sentences(i, 1)), LCase("Mohamed")) Or InStr(LCase(Sentences(i, 1)), LCase("Silva, Minoli")) Or InStr(LCase(Sentences(i, 1)), LCase("Piva")) Or InStr(LCase(Sentences(i, 1)), LCase("Gulajska")) Or InStr(LCase(Sentences(i, 1)), LCase("Gardam")) Or InStr(LCase(Sentences(i, 1)), LCase("Dion")) Or InStr(LCase(Sentences(i, 1)), LCase("Mathew")) Or InStr(LCase(Sentences(i, 1)), LCase("Tracz")) Or InStr(LCase(Sentences(i, 1)), LCase("Cimicata")) Then
            iWordPos = 0
        End If
        If iWordPos = 1 Then
            Range("A" & i).EntireRow.Select
            Selection.Copy
            Sheets("PPIP PSR and Screening").Select
            Do Until Cells(AustenRows, 1) = ""
                AustenRows = AustenRows + 1
            Loop
            Cells(AustenRows, 1).Select
            ActiveSheet.Paste
            Sheets(CategoryName).Select
            Selection.Clear
        End If
    Next i
    
    ' Format the new sheet for printing.
    Sheets("PPIP PSR and Screening").Select
    SetColumns
    PrintFormat
    
    ' Delete all empty rows in the original sheet.
    Sheets(CategoryName).Select
    DeleteEmptyRows Range("A2:A" & EndRow)
    
End Sub
Private Sub DeleteEmptyRows(DeleteRange As Range)

    ' -------------------------------------------------------------------------
    ' DeleteEmptyRows scans through a sheet and deletes any blank rows.
    ' It uses an input range to know where to scan.
    ' -------------------------------------------------------------------------
    ' Example use:  DeleteEmptyRows Range("A1:U241")
    ' -------------------------------------------------------------------------
    
    Dim rCount As Long
    Dim r As Long
    Application.ScreenUpdating = False
    
    ' If the range is nothing, quit the sub
    If DeleteRange Is Nothing Then Exit Sub
    If DeleteRange.Areas.Count > 1 Then Exit Sub
    
    With DeleteRange
        rCount = .Rows.Count
        For r = rCount To 1 Step -1
            If Application.CountA(.Rows(r)) = 0 Then
                .Rows(r).EntireRow.Delete
            End If
        Next r
    End With
    
    Cells(1, 1).Select
    Application.ScreenUpdating = True
    
End Sub
Private Sub SetPageBreaks()

    ' -------------------------------------------------------------------------
    ' SetPageBreaks finds all the page breaks in the sheet and then checks if
    ' each one goes through a related set of submissions.  If it does, it moves
    ' the offending page break up and continues on its merry way.
    ' -------------------------------------------------------------------------
    
    ' Find the last row containing data.
    EndRow = Range("B2").End(xlDown).Row

    ' Scroll through the entire sheet.  This is needed due to oddities in Excel
    ' when it doesn't realize there is a page break without actually seeing it
    ' on screen.
    While ActiveWindow.ScrollRow < EndRow
        ActiveWindow.LargeScroll Down:=1
    Wend

    ' For each page break, check if it is in a merged cell.  If it is, go up
    ' from that cell until you find an unmerged one and place a page break
    
    For Each x In ActiveSheet.HPageBreaks
        Cells(x.Location.Row, 21).Select
        If Cells(x.Location.Row, 21) = "" And Cells(x.Location.Row, 4) <> "" Then
            Cells(x.Location.Row, 21).Select
            i = 0
            Do While Cells(x.Location.Row - i, 21) = "" And i < EndRow
                i = i + 1
            Loop
            ActiveSheet.HPageBreaks.Add Range("A" & (x.Location.Row - i))
        End If
        Cells(x.Location.Row, 21).Select
        If Cells(x.Location.Row - 1, 21) = "" And Cells(x.Location.Row - 1, 4) = "" Then
            ActiveSheet.HPageBreaks.Add Range("A" & (x.Location.Row - 1))
        End If
        
        Cells(1, 1).Select
        While ActiveWindow.ScrollRow < EndRow
            ActiveWindow.LargeScroll Down:=1
        Wend
    Next
    
    While ActiveWindow.ScrollRow > 1
        ActiveWindow.LargeScroll Up:=1
    Wend
    
    Cells(1, 1).Select
    
End Sub

Private Sub OpenSingleFile(OpenText As String, FilterType As String)

    ' Modified from http://www.tek-tips.com/faqs.cfm?fid=4114

    Dim Filter As String, Title As String
    Dim FilterIndex As Integer
    Dim FileName As Variant
    
    ' File filters
    Filter = "Excel Files (*.xls),*.xls," & _
        "Text Files (*.txt),*.txt," & _
        "CSV Files (*.csv),*.csv," & _
        "All Files (*.*),*.*"
    
    ' Sets the filter based on the input value
    If FilterType = ".xls" Then
        FilterIndex = 1
    ElseIf FilterType = ".txt" Then
        FilterIndex = 2
    ElseIf FilterType = ".csv" Then
        FilterIndex = 3
    Else
        FilterIndex = 4
    End If
    
    ' Set Dialog Caption
    Title = OpenText
    
    ' Select Start Drive & Path as the Directory the file is currently in.
    ChDrive (Mid(PathName, 1, 1))
    ChDir (PathName)
    With Application
        ' Set File Name to selected File
        FileName = .GetOpenFilename(Filter, FilterIndex, Title)
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
    Workbooks.Open FileName

End Sub

' Code to highlight "May need update" on New Format sheets

Private Sub MayNeedUpdate()

    Application.ScreenUpdating = False
    EndRow = Range("B2").End(xlDown).Row
    
    For i = 2 To EndRow
        If Cells(i, 2) <> "-" Then
            If CDate(Cells(i, 2)) <= SnapshotDate - 14 And Cells(i, 1) <> "" Then
                Cells(i, 21).Select
                Selection.Interior.ColorIndex = 36  ' Highlight it light yellow.
            End If
        End If
        If Cells(i, 10) <> "-" And Cells(i, 1) <> "" Then
            If CDate(Cells(i, 10)) <= SnapshotDate - 14 Then
                Cells(i, 21).Select
                Selection.Interior.ColorIndex = 36  ' Highlight it light yellow.
            End If
        End If
    Next i

End Sub

' Code to add [REMD] to statuses

Private Sub FlagSubs()

    Sheets("Flags").Select
    EndRow = Range("A2").End(xlDown).Row
    
    i = 0
    For i = 2 To EndRow
    
        SubN = Sheets("Flags").Cells(i, 1)
        flag = 0
        
        Sheets(CategoryName).Select
        
        For k = 2 To Range("A2").End(xlDown).Row
            If Cells(k, 4) = SubN Then
                flag = 1
                RowN = k
            End If
        Next k
        
        FlagColumn = Range("A2").End(xlToRight).Column + 1
        
        If flag = 1 Then
            If Sheets("Flags").Cells(i, 8) <> "" Then
                Cells(RowN, 4) = "" & Cells(RowN, 4) & Chr(10) & "[" & Sheets("Flags").Cells(i, 8) & "]"
                If Cells(RowN, FlagColumn) <> "" Then
                    Cells(RowN, FlagColumn) = Cells(RowN, FlagColumn) & ", pilot"
                Else
                    Cells(RowN, FlagColumn) = "pilot"
                End If
            Else
                For j = 15 To 20
                    If Cells(RowN, j) <> "  " Then
                        If Sheets("Flags").Cells(i, j - 13) <> "" Then
                            Cells(RowN, j) = "" & Cells(RowN, j) & Chr(10) & "[" & Sheets("Flags").Cells(i, j - 13) & "]"
                            If Cells(RowN, FlagColumn) <> "" Then
                                Cells(RowN, FlagColumn) = Cells(RowN, Range("A2").End(xlToRight).Column + 1) & ", REMD"
                            Else
                                Cells(RowN, FlagColumn) = "REMD"
                            End If
                        End If
                    End If
                Next j
            End If
        End If
        
    Next i
    
    Sheets(CategoryName).Select
    
    Cells(1, FlagColumn) = "Flags"
    Cells.Select
    Sheets("Flags").Select
    Cells.Select
    Selection.Clear
    Sheets("Flags").Delete
    
    Range("B1:W1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
End Sub

Private Sub DeleteUnwanted()

    ' -------------------------------------------------------------------------
    ' -------------------------------------------------------------------------
        
    ' Application.ScreenUpdating = False
        
    StringArray = Array("A  ", "B  ", "S  ", "I  ")
    
    i = 2
    
    EndRow = Range("A2").End(xlDown).Row
    Do While i <= EndRow
        flag = 0
        For j = 0 To 3
            If InStr(LCase(Cells(i, 13)), LCase(StringArray(j))) Then
                flag = 1
            End If
        Next j
        If flag = 1 Then
            Cells(i, 2).EntireRow.Delete
        Else
            i = i + 1
        End If
    Loop

End Sub
Private Sub SortList()

    ' Sort the entire list by E deadline
    Cells.Select
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("C2") _
        , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row
        
    ' Find where the E deadline goes late
    i = 2
    Do Until CDate(Cells(i, 1)) >= SnapshotDate Or Cells(i, 1) = "-" Or Cells(i, 1) = ""
        Cells(i, 1).Select
        Selection.Interior.ColorIndex = 6   ' Highlight yellow if older.
        i = i + 1
    Loop
    
    ' Note:  i = the row *after* the last late E sub.
    
    ' Find first NULL row ("-") with information.
    j = i
    Do Until Cells(j, 1) = "-" Or Cells(j, 1) = ""
        j = j + 1
    Loop

    ' Sort the NULL section by the Current Deadline.
    Range(j & ":" & EndRow).Select
    Selection.Sort Key1:=Range("I" & i), Order1:=xlAscending, Key2:=Range("C2") _
        , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
        
    For k = i To EndRow Step 1
        If Cells(k, 9) <> "-" Then
            If CDate(Cells(k, 9)) <= SnapshotDate Then
                Cells(k, 9).Select
                Selection.Interior.ColorIndex = 6
            End If
        End If
    Next k

End Sub
Private Sub BasicFormat()

    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row

    ' Full sheet - size change
    Cells.Select
    Selection.Font.Size = 6
    
    ' Add notes column
    Cells(1, 20).Select
    Selection.EntireColumn.Insert
    Cells(1, 20) = "Notes"
    
    ' All
    Range("A1", "T" & EndRow).Select
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Left align Sub type, Product name, Notes
    Range("B1", "B" & EndRow).Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    Range("G1", "G" & EndRow).Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    Range("T1", "T" & EndRow).Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    
    ' Header only
    Range("A1", "T1").Select
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
End Sub
Private Sub InsertRowNumbers()

    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row

    Cells(1, 1).Select
    Selection.EntireColumn.Insert
    
    Range("A1", "A" & EndRow).Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    Cells(2, 1) = 1
    Range("A2", "A" & EndRow).Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
    
End Sub
Private Sub SetColumns()

    ' -------------------------------------------------------------------------
    ' SetColumns sets the columns by pixel width to a standard set of widths
    ' for Tracker's reports, using SetColumnWidth(ColumnNumber, Pixels).
    ' -------------------------------------------------------------------------
    
    SetColumnWidth 1, 30
    SetColumnWidth 2, 55
    SetColumnWidth 3, 55
    SetColumnWidth 4, 55
    SetColumnWidth 5, 45
    SetColumnWidth 6, 40
    SetColumnWidth 7, 35
    SetColumnWidth 8, 180
    SetColumnWidth 9, 30
    SetColumnWidth 10, 60
    SetColumnWidth 11, 60
    SetColumnWidth 12, 50
    SetColumnWidth 13, 50
    SetColumnWidth 14, 68
    SetColumnWidth 15, 68
    SetColumnWidth 16, 68
    SetColumnWidth 17, 68
    SetColumnWidth 18, 68
    SetColumnWidth 19, 68
    SetColumnWidth 20, 68
    SetColumnWidth 21, 215
    
    Cells(1, 2) = "E Deadline yyyy-mm-dd"
    Cells(1, 4) = "Sub #"
    Cells(1, 5) = "Prod Type"
    Cells(1, 7) = "Actv List"
    Cells(1, 15) = "Occupational Status"
    Cells(1, 17) = "VRD Status"
    Cells(1, 18) = "Dietary Status"
    
    Rows("1:1").RowHeight = 18.75
    
    Range("B:B,J:K").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"
    
End Sub
Private Sub SetColumnWidth(ColumnNumber As Integer, Pixels As Integer)

    ' -------------------------------------------------------------------------
    ' SetColumnWidth is a quick converter and column width set function, which
    ' takes in the column number and the width in pixels, and sets it.  There
    ' is a conversion required because .ColumnWidth is set in points.
    ' -------------------------------------------------------------------------
    
    Columns(ColumnNumber).ColumnWidth = ((Pixels - 12) / 7) + 1
    
End Sub
Private Sub GetNotes()

    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row
    
    For i = 2 To EndRow
        Sheets(CategoryName).Select
        SubN = Cells(i, 4)
        j = 2
        flag = 0
        Sheets("Old List").Select
        OldListEndRow = Range("C2").End(xlDown).Row
        Do While flag = 0 And j <= OldListEndRow
            If InStr(LCase(Cells(j, 4)), SubN) Then
                flag = 1
            Else
                j = j + 1
            End If
        Loop
        If flag = 1 Then
            Sheets("Old List").Select
            Cells(j, 21).Copy
            Sheets(CategoryName).Select
            Cells(i, 21).Select
            ActiveSheet.Paste
        Else
            Sheets(CategoryName).Select
            Cells(i, 21) = "NEW"
            Cells(i, 21).Select
            Selection.Interior.ColorIndex = 4
        End If
    Next i
    
End Sub
Sub RelatedSubs()

    ' Find the last row
    EndRow = Range("B2").End(xlDown).Row
    
    For i = 3 To EndRow
        If Cells(i, 21) = "" Then
            Range("B" & i - 1, "T" & i).Select
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            For j = 12 To 13
                If Cells(i, j) = Cells(i - 1, j) Then
                    Cells(i, j).Select
                    Selection.Font.ColorIndex = 2
                    Selection.Borders(xlEdgeTop).LineStyle = xlNone
                End If
            Next j
            Range("U" & i - 1, "U" & i).Select
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        ElseIf Cells(i, 21) = "NEW" Then
            If Cells(i, 2) = Cells(i - 1, 2) And Cells(i, 12) = Cells(i - 1, 12) And Cells(i, 13) = Cells(i - 1, 13) Then
                If Cells(i, 7) = Cells(i - 1, 7) Or Cells(i, 9) = Cells(i - 1, 9) Then
                    Range("B" & i - 1, "T" & i).Select
                    With Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlHairline
                        .ColorIndex = xlAutomatic
                    End With
                    Range("L" & i, "M" & i).Select
                    Selection.Font.ColorIndex = 2
                    Selection.Borders(xlEdgeTop).LineStyle = xlNone
                    Range("U" & i - 1, "U" & i).Select
                    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                    Cells(i, 23) = "New"
                    Cells(i, 21) = ""
                    Cells(i, 21).Interior.ColorIndex = 2
                End If
            End If
        End If
    Next i
End Sub
Private Sub NewSubs()
    Dim CROlist(0 To 10) As String ' Remember to change this dim if you add/remove a name
    
    CROlist(0) = "Boucher, Alain"
    CROlist(1) = "Muir, Sean"
    CROlist(2) = "Murray, Patricia"
    CROlist(3) = "Pathy, Diane"
    CROlist(4) = "Silva, Minoli"
    CROlist(5) = "Smith, Charles"
    CROlist(6) = "Stiege, Stacie"
    CROlist(7) = "Stewart, Terri A"
    CROlist(8) = "Dang, Bing"
    CROlist(9) = "Girard, Stephanie"
    CROlist(10) = "Kavaslar, Nihan"

    Cells(1, 23) = "New"
    
    EndRow = Range("B2").End(xlDown).Row
    For i = 3 To EndRow
        If Cells(i, 21) = "NEW" Then
            Cells(i, 23) = "New"
            Cells(i, 21) = ""
            If Cells(i, 13) = "not required" Then
                Cells(i, 21) = "STL: AC"
            Else
                For j = LBound(CROlist) To UBound(CROlist)
                    If Cells(i, 13) = CROlist(j) Then
                        Cells(i, 21) = "STL: CRO"
                    End If
                Next j
                If Cells(i, 21) = "" Then
                    Cells(i, 21) = "STL: VRD"
                End If
            End If
            Cells(i, 21).Select
            With Selection
                .VerticalAlignment = xlTop
                .WrapText = True
                .Font.Size = 6
                .Font.Bold = True
                .Font.ColorIndex = Automatic
                .Interior.ColorIndex = 2
            End With
        End If
    Next i
End Sub
Private Sub PrintFormat()
    
    Application.ScreenUpdating = False
    
    ' Fit the row heights such that nothing is cut off.
    Range("A1:A" & Range("A2").End(xlDown).Row).EntireRow.AutoFit
    Cells(1, 1).Select
    
    Range("B1:T1").Select
    Selection.AutoFilter
    
    ' EndRow = Range("A2").End(xlDown).Row
    ' For i = 1 To EndRow
    '     If Cells(i, 1).RowHeight < 16.5 Then
    '         Cells(i, 1).RowHeight = 16.5
    '     End If
    ' Next i
    
    ' Set the view to Page Break preview and the zoom to 100%.
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
    
    SetHeader "" & CategoryName & "", "" & ActiveSheet.Name & "", "" & SnapshotDate & ""
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub SetHeader(CategoryName As String, SheetName As String, SnapshotDate As String)

    ' -------------------------------------------------------------------------
    ' SetHeader writes all the standard Tracker's text into the header and
    ' footer, as well as choosing the active print area, setting the margins,
    ' setting the print quality and making the sheets landscape orientation and
    ' legal sized.  It also makes sure the columns fit the width of one page.
    ' It requires an input of the snapshot date, to put in the header.
    ' -------------------------------------------------------------------------
    ' Example use:  SetHeader "Category A" "19-apr-2009"
    ' Example use:  SetHeader "" & SheetName & "", "" & SnapshotDate & ""
    ' -------------------------------------------------------------------------
    
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With
    
    ' Set print area.
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$" & Range("A2").End(xlDown).Row
    
    With ActiveSheet.PageSetup
        
        ' Header information.
        .LeftHeader = "&8Snapshot Date =" & SnapshotDate & " 9:00"
        .CenterHeader = _
        "&""Arial,Bold""" & CategoryName & " @ Levels C,D,E,F,G,H,R with Division Status&""Arial,Regular""" & Chr(10) & "&8(sorted by Level E Deadline)"
        .RightHeader = "Page &P of &N"
        
        ' Footer information.
        .CenterFooter = "Page &P of &N"
        .RightFooter = "Printed &D &T"
        
        ' Setting margins.
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.62992125984252)
        .BottomMargin = Application.InchesToPoints(0.511811023622047)
        .HeaderMargin = Application.InchesToPoints(0.236220472440945)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        
        ' Setting print information.
'        .PrintQuality = 600 (This is removed as it was triggering errors during the code. No required and can be set manually)
        .CenterHorizontally = True
        .Orientation = xlLandscape
        .PaperSize = xlPaperLegal
        .FitToPagesWide = 1
        
    End With
    
    ' Drag the vertical page break one over to account for the notes column
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    
End Sub
Private Sub AddPass()
    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row
    
    For i = 2 To EndRow
        If Cells(i, 22) <> 1 Then
            WithoutPass = Cells(i, 14)
            Cells(i, 14) = "" & WithoutPass & Chr(10) & " (Pass " & Cells(i, 22) & ")"
        End If
    Next i
End Sub
Private Sub DNotRequired()
    ' Find the last row
    EndRow = Range("A2").End(xlDown).Row
    For i = 2 To EndRow
        For j = 15 To 20
            If Cells(i, j) <> "  " Then
                If Cells(i, j + 8) = "Y" Then
                    Cells(i, j) = "" & Cells(i, j) & Chr(10) & "(D not req.)"
                End If
            End If
        Next j
    Next i
End Sub
Private Sub RemoveExtra()
    Range("V1:AC1").EntireColumn.Clear
    Sheets("Old List").Select
    Cells.Select
    Selection.Clear
    Sheets("Old List").Delete
End Sub
Private Sub StatusDate()
    lastcol = Range("B1").End(xlToRight).Column + 1
    Cells(1, lastcol) = "Status date changed last week"
    ' Move new columns
    Range("AD1", "AI1").EntireColumn.Select
    Selection.Cut
    Cells(1, lastcol + 1).Select
    ActiveSheet.Paste
    ' Fix formatting
    Range("V1", "AD1").Select
    Selection.Font.Bold = True
    Selection.Font.Size = 6
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    EndRow = Range("B2").End(xlDown).Row
    Range("V1", "AD" & EndRow).Select
    Selection.Font.Size = 6
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    Range("B1:AD1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ' Move the column
    ' Cells(1, 30).EntireColumn.Select
    ' Selection.Cut
    ' Cells(1, lastcol + 1).EntireColumn.Select
    ' ActiveSheet.Paste
    
    ' Flag changes
    For i = 2 To EndRow
        If CDate(Cells(i, lastcol + 1)) >= (SnapshotDate - 7) Then
            Cells(i, lastcol) = "Y"
        End If
    Next i
End Sub
Private Sub Highlight()
    EndRow = Range("B2").End(xlDown).Row
    For i = 2 To EndRow Step 1
        If Cells(i, 1) <> "" Then
            If InStr(LCase(Cells(i, 4)), LCase("pilot")) Then
                Cells(i, 4).Interior.ColorIndex = 37
            End If
            If Cells(i, 10) = "-" Or InStr(LCase(Cells(i, 3)), LCase("A.3.2")) Then
                Cells(i, 3).Interior.ColorIndex = 39  ' Highlight Joint Reviews purple.
            End If
            If Cells(i, 14) <> "" Then
                For j = 14 To 20 Step 1
                    If j = 14 Then
                        If InStr(LCase(Cells(i, j)), LCase("E  ")) Or InStr(LCase(Cells(i, j)), LCase("F  ")) Or InStr(LCase(Cells(i, j)), LCase("G  ")) Or InStr(LCase(Cells(i, j)), LCase("H  ")) Then
                            Cells(i, j).Interior.ColorIndex = 15   ' Highlight.
                        End If
                        If InStr(LCase(Cells(i, j)), LCase("ON HOLD")) Or InStr(LCase(Cells(i, j)), LCase("REJECTED")) Then
                            Cells(i, j).Interior.ColorIndex = 3    ' Highlight red.
                        End If
                    Else
                        If InStr(LCase(Cells(i, j)), LCase("STARTED")) Then
                            Cells(i, j).Interior.ColorIndex = 35   ' Highlight light green.
                        ElseIf InStr(LCase(Cells(i, j)), LCase("IN QUEUE")) Then
                            Cells(i, j).Interior.ColorIndex = 20   ' Highlight light blue.
                        ElseIf InStr(LCase(Cells(i, j)), LCase("ON HOLD")) Or InStr(LCase(Cells(i, j)), LCase("REJECTED")) Then
                            Cells(i, j).Interior.ColorIndex = 3    ' Highlight red.
                        ElseIf InStr(LCase(Cells(i, j)), LCase("COMPLETED")) Then
                            Cells(i, j).Interior.ColorIndex = 37   ' Highlight darker blue.
                        End If
                    End If
                Next j
            End If
        End If
    Next i
End Sub
Private Sub InsertMonths()
    
    Dim i As Integer
        
    For i = 0 To 6
        If Weekday(Now) - i = 1 Then
            SnapshotDate = DateValue(Now) - i
        End If
    Next i
        
    EndRow = Range("B2").End(xlDown).Row
    
    WriteBreaks 2, CDate(0), "Work on Hand"
    Range("A1:U1").Borders(xlEdgeBottom).LineStyle = xlNone
    Range("A2:U2").Borders(xlEdgeTop).LineStyle = xlContinuous
    
    i = 3
    Do While CDate(Cells(i, 2)) < SnapshotDate
        i = i + 1
    Loop
    
    If Month(CDate(Cells(i - 1, 2))) = Month(CDate(Cells(i, 2))) Then
        WriteBreaks i, SnapshotDate
    Else
        WriteBreaks i, CDate(Left(Cells(i, 2), 8) & "01")
    End If
    
    i = i + 1
    Do While i < EndRow
        If Cells(i, 2) <> "-" Then
            Do While Month(CDate(Cells(i, 2))) = Month(CDate(Cells(i - 1, 2)))
                i = i + 1
                If Cells(i, 2) = "-" Then
                    Exit Do
                ElseIf IsEmpty(Cells(i, 2)) Then 'Added this to allow another exit from the DO loop as it wasn't exiting in cases where date was "-".
                    Exit Do
                ElseIf Cells(i, 2) = "" Then 'Added as a replica of the IsEmpty exit above.
                    Exit Do
                End If
            Loop
        End If
        If Cells(i, 2) <> "-" Then
            WriteBreaks i, CDate(Left(Cells(i, 2), 8) & "01")
            i = i + 1
        Else
            WriteBreaks i, CDate((Cells(i - 1, 2) + 1)), "Levels F, G, H as well as C ON HOLD or PENDING FEES"
            Exit Do
        End If
    Loop
    
End Sub
Private Sub WriteBreaks(RowNum As Integer, EDeadline As Date, Optional OtherWording As String)
    
    ' -------------------------------------------------------------------------
    ' WriteBreaks() writes a single break (header row) into the Trackers sheet.
    ' It takes the row number, E deadline (to allow the row to be sorted) and
    ' optionally takes a string containing other wording if other text is
    ' required. It also formats itself.
    ' -------------------------------------------------------------------------
    
    ' Select the row that the header is being added above and insert the row.
    ' Select the print area section of the new break row.
    
    Cells(RowNum, 1).EntireRow.Select
    Selection.Insert Shift:=xlDown
    Range("A" & RowNum & ":U" & RowNum).Select
    
    ' Set the colours and borders on the break row (outside border only).
    
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Interior.ColorIndex = 15 ' 25% gray, #C0C0C0
    
    ' Set the formatting for the cell that will contain the break row text.
    
    Cells(RowNum, 11).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Font.Size = 10
        .Font.Bold = True
        .Font.ColorIndex = Automatic
    End With
    
    ' Insert the break row text.  It will either contain the optional other
    ' wording string, or the generic E deadline text using the snapshot date.
    
    If OtherWording <> "" Then
        Cells(RowNum, 11) = OtherWording
    Else
        Cells(RowNum, 11) = "E deadline after " & Format(EDeadline, "mmmm d, yyyy")
    End If
    
    ' Set the E deadline column to include the E deadline.  Hide it by setting
    ' the colour to be the same as the background.  Note that this is to allow
    ' the sheets to be sorted by E deadline without losing the location of the
    ' break rows.
    
    Cells(RowNum, 2) = EDeadline
    Cells(RowNum, 2).Select
    Selection.Font.ColorIndex = 15 ' 25% gray, #C0C0C0
    
    Range("A" & RowNum - 1 & ":U" & RowNum - 1).Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    
End Sub
Sub iCalFeed()

    ' Note, you need to have the tracking sheet open before running this.

    ' Directory for the iCal feeds
   ' PathName = "L:\RD\RSID\Sci Coord\HRP HC-6\Reporting HC-6-103\Trackers Reports\Calendar Feeds\"     - This folder has been moved over to the Y: drive.
   ' PathName = "\\Ncr-a-irbv1s\irbv1\HC\PMRA\RD\RSID\HEALTH RISK PROTECTION HC6\REPORTING\Science and Submission Management Section\Trackers Reports\Calendar Feeds\" This is the new path.
    PathName = "Y:\HC\PMRA\RD\RSID\HEALTH RISK PROTECTION HC6\REPORTING\Science and Submission Management Section\Trackers Reports\Calendar Feeds\" 'This is the simplified version of new path
    
    ' The names of the ACs that feeds will be generated for
    ACArray = Array("Mathew, Suzan", "Tracz, Kamil", "Orhiobhe, Adora", "Gardam, Kamerine", "Gulajska, Katarzyna", "Dion, Julie", "Ali, Sabrina", "Piva, Angela")
    
    For j = 0 To 7
        
        ' Where and what the iCal file will be named
        MyFile = PathName & ACArray(j) & " - E deadline iCal Feed.ics"
    
        ' Set and open file for output
        fnum = FreeFile()
        Open MyFile For Output As fnum
    
        ' iCal header info and a blank line
        Print #fnum, "BEGIN:VCALENDAR"
        Print #fnum, "VERSION:1.0"
        Write #fnum,
    
        ' Go through the tracking sheet line-by-line
        EndRow = Range("B2").End(xlDown).Row
        For i = 1 To EndRow
            ' If the AC is the one you want, and there is an E deadline
            If Cells(i, 12) = ACArray(j) And Cells(i, 2) <> "-" Then
                Print #fnum, "BEGIN:VEVENT"
                Print #fnum, "DTSTART;VALUE=DATE:" & Format(CDate(Cells(i, 2)), "yyyymmdd") ' Reformat the date to not have hyphens
                Print #fnum, "DTEND;VALUE=DATE:" & Format(CDate(Cells(i, 2)), "yyyymmdd")
                If Cells(i, 14) = "E  ON HOLD" Then
                    Print #fnum, "SUMMARY:" & Left(Cells(i, 4), 9) & " - E on hold" ' Marking E on holds
                ElseIf Cells(i, 10) = "-" Then
                    Print #fnum, "SUMMARY:" & Left(Cells(i, 4), 9) & " - JR" ' Marking Joint Reviews
                Else
                    Print #fnum, "SUMMARY:" & Left(Cells(i, 4), 9) ' Summary = what shows up on the calendar.  In this case, the sub. no. (truncated in case there is additional text in the cell)
                End If
                Print #fnum, "LOCATION:" & Cells(i, 8) ' Location = what shows up upon mouseover.  In this case, product name.
                Print #fnum, "END:VEVENT"
                Write #fnum,
            ElseIf Cells(i, 12) = ACArray(j) And Left(Cells(i, 14), 1) = "H" Then ' Adding H deadlines
                Print #fnum, "BEGIN:VEVENT"
                Print #fnum, "DTSTART;VALUE=DATE:" & Format(CDate(Cells(i, 10)), "yyyymmdd") ' Reformat the date to not have hyphens
                Print #fnum, "DTEND;VALUE=DATE:" & Format(CDate(Cells(i, 10)), "yyyymmdd")
                Print #fnum, "SUMMARY:" & Left(Cells(i, 4), 9) & " - H deadline" ' Marking H deadlines
                Print #fnum, "LOCATION:" & Cells(i, 8) ' Location = what shows up upon mouseover.  In this case, product name.
                Print #fnum, "END:VEVENT"
                Write #fnum,
            End If
        Next i
        
        ' Close the calendar
        Print #fnum, "END:VCALENDAR"
        Close #fnum
        
    Next j ' Go to the next AC
End Sub
Sub NeonicFlags()

                    
' Goes down the colum and check for a neonic, exits the loop if the cell is empty

EndRow = Range("B2").End(xlDown).Row

    For i = 1 To EndRow
        Range("G" & i).Select
        If InStr(1, ActiveCell, "THE") Then
            ActiveCell.Interior.Color = vbRed
        ElseIf InStr(1, ActiveCell, "COD") Then
            ActiveCell.Interior.Color = vbRed
        ElseIf InStr(1, ActiveCell, "IMI") Then
            ActiveCell.Interior.Color = vbRed
        'ElseIf IsEmpty(ActiveCell) Then Exit For
        End If
    Next i
    
'Selects the A1 cell to return to the top
    Range("A1").Select
    
End Sub

Private Sub CostRecovery()
'
' FlagCostRecovery Macro
'
' Windows prompt for a new file. This file will contain a list of subs under new Cost Recovery. Copy this list into a Sheet
' on the Trackers report.
' Compare each Submission number from D2 and onwards until there are 3 blanks in a row and see if there is a match
' to the Cost Recovery list. If there is, Flag it a colour.

Dim CRFileName As String
Dim CRPathName As String
Dim CRCategoryName As String

CRFileName = ActiveWorkbook.Name
CRPathName = ActiveWorkbook.Path
CRCategoryName = ActiveSheet.Name
        
    
'CostRecoveryCreateWorksheet
Sheets.Add.Name = "NewSubs"
        
OpenSingleFile "Select the file which contains the submissions received after April 1 2017:", ".csv"
Cells.Select
Selection.Copy
TempFileName = ActiveWorkbook.Name
Windows(CRFileName).Activate
ActiveSheet.Paste
Application.DisplayAlerts = False
Windows(TempFileName).Close
 
'CostRecoveryFlag

Sheets("NewSubs").Select
EndRow = Range("A2").End(xlDown).Row
i = 0
Application.ScreenUpdating = False 'Stops the screen from flickering and slowing down the script


For i = 2 To EndRow
    NewSub = Sheets("NewSubs").Cells(i, 1) ' Places the new submission into a placeholder
    Application.StatusBar = "Checking Submission " & i & " of " & EndRow 'Adds a Status Bar counter at the bottom of the screen.
        Sheets(CRCategoryName).Select
        k = 2
        blanks = 0
        

         Range("D2").Select
         Do 'This loop compares the submisison in the placeholder to each submission on the trackers list. If there's a match, it highlights it a colour
            If ActiveCell = "" Then
                blanks = blanks + 1
                ActiveCell.Offset(1, 0).Select
            ElseIf ActiveCell.Value = NewSub Then
                Selection.Interior.ColorIndex = 37
                blanks = 0
                'ActiveCell.Offset(1, 0).Select - Removing this code as it's not necessary to go to the next row if the sub is found.
                Exit Do 'added this Exit function as it doesn't need to keep searching for the sub once it has been found.
            Else
                ActiveCell.Offset(1, 0).Select
                blanks = 0
            End If
                         
        Loop Until blanks > 3
        
Next i
 
 
Sheets(CRCategoryName).Select
Sheets("NewSubs").Delete
 
 
 
End Sub
