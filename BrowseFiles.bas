Attribute VB_Name = "BROWSEFILES"
Sub SelectFile()

Dim DialogBox As FileDialog
Dim path As String

Set DialogBox = Application.FileDialog(msoFileDialogFilePicker)
DialogBox.Title = "Select file for " & FileType
DialogBox.Filters.Clear
DialogBox.Show

If DialogBox.SelectedItems.Count = 1 Then
path = DialogBox.SelectedItems(1)
End If

Sheets("Main").Range("InputDTRTemplate").Value = path
End Sub
Sub RunMyCodeNow()
    Dim CopyPivotLessGreaterThanData As String
    Dim CopyPivotLessGreaterThanData1 As String
    
    Dim Countemployees As Integer
    Dim CountTotalAgent As Integer
    Dim CopyTotalAgentbelow144 As String
    Dim thisarray As Variant
    Dim startNum, endNum As Integer
    
    Sheets("Attendance Breakdown").Select
    
    'Filter all agents and total hours below 144
    'Count total of emplyees
    Range("B28").Select
    Range(Selection, Selection.End(xlDown)).Select
    Countemployees = Selection.Cells.Count
    
    Rows("28:28").Select
    Selection.AutoFilter
    ActiveSheet.Range("$B$28:$FH$" & Countemployees + 1).AutoFilter Field:=4, Criteria1:="Agent"
    
    'Search Name and Copy employee data
    Range("A25").Select
     Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("B400").Select
    ActiveSheet.Paste
      
    'get the total count for agent
    DisCountforagentLesthan = Selection.Rows.Count - 1
    
    Rows("28:28").Select
    ActiveSheet.Range("$B$28:$FH$" & Countemployees + 1).AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues
    
    'count filtered data

    'Search Name and Copy employee data
     Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    'Select and count all agent
    Range(Selection, "B" & Countemployees + 28).Select
    'copy data that lessthan 143
    Application.CutCopyMode = False
    Selection.Copy
    
    'Paste and Select Range
    Range("A200").Select
    ActiveSheet.Paste
    
    'Copy and Select Range x get the total account
    Range("B201").Select
    Range(Selection, Selection.End(xlDown)).Select
 
    'get the total count for agent that lessthan 144
    totalCountforagentLesthan = Selection.Cells.Count
    Selection.Copy
    
    'Paste and Select Range
    Range("B16").Select
    ActiveSheet.Paste
    
    'Copy agent below 144
    Range("E201:E" & totalCountforagentLesthan + 200).Select
    Selection.Copy
    
    'Paste and Select Range
    Range("D16").Select
    ActiveSheet.Paste
    
   'Set yes/no in array
   thisarray = Range("E201:E" & totalCountforagentLesthan + 200).Value
   counter = 1                'looping structure to look at array
   While counter <= UBound(thisarray)
      'MsgBox thisarray(counter, 1)
      Range("E16:E" & totalCountforagentLesthan + 15) = "NO"
      counter = counter + 1
    Wend
    
    'Set value
    Range("B15") = "Agent"
    'Set the total of FTEs for agent
    Range("C15:C" & totalCountforagentLesthan + 14) = DisCountforagentLesthan - totalCountforagentLesthan
    'Set 1 to number of FTEs
    Range("C16:C" & totalCountforagentLesthan + 15) = 1
    'Set ~ # of days (if <144 hours or 18 days)
    Range("D15") = "~"
    'Set ~ # >= 144 hours or 18 days
    Range("E15") = "~"
    
    'Copy agent below 144
    Range("E400:FH" & totalCountforagentLesthan + 400).Select
    Selection.Copy
    'Paste and Select Range
    Range("A200").Select
    ActiveSheet.Paste
    
    'Clear data
    Range("B400").Select
    Selection.EntireRow.Delete

    'select and unfilter this row
    Rows("28:28").Select
    Selection.AutoFilter
    
    'Create Sr.agent Row
    Selection.AutoFilter
    ActiveSheet.Range("$B$28:$FH$" & Countemployees + 1).AutoFilter Field:=4, Criteria1:="Sr. Agent"
    
    'Search Name and Copy employee data
    Range("A25").Select
     Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("C400").Select
    ActiveSheet.Paste
      
    'get the total count for agent
    DisCountforSragentLesthan = Selection.Rows.Count - 1
    
    Rows("28:28").Select
    ActiveSheet.Range("$B$28:$FH$" & Countemployees + 1).AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues

    'Search Name and Copy employee data
     Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    'Select and count all sr agent
    Range(Selection, "B" & Countemployees + 28).Select
    'copy data that lessthan 143
    Application.CutCopyMode = False
    Selection.Copy
    
    'Paste and Select Range
    Range("A200").Select
    ActiveSheet.Paste
    
    'Copy and Select Range x get the total account
    Range("B201").Select
    'If Statement for 1 value
    Range("B199") = Application.WorksheetFunction.CountIf(Range("B201:B310"), "*")
    If Range("B199") = 1 Then
       totalCountforSragentLesthan = Selection.Cells.Count
       Selection.Copy
    End If
    
    If Range("B199") <> 1 Then
       totalCountforSragentLesthan = Selection.Cells.Count
       Selection.Copy
    End If
    
    'Subtract Agent to Sr Agent
    countforSub = 17 + totalCountforagentLesthan
    'Paste and Select Range
    Range("B" & countforSub).Select
    ActiveSheet.Paste

    'Copy agent below 144
    Range("E201:E" & totalCountforSragentLesthan + 200).Select
    Selection.Copy
    
    'Paste and Select Range
    Range("D" & countforSub).Select
    ActiveSheet.Paste
    
   'Set yes/no in array
   thisarray = Range("E201:E" & totalCountforSragentLesthan + 200).Value
   'If count not equal to 1
   If totalCountforSragentLesthan <> 1 Then
    counter = 1 'looping structure to look at array
    While counter <= UBound(thisarray)
       'MsgBox thisarray(counter, 1)
       Range("E" & countforSub & ":E" & totalCountforSragentLesthan + 17 + totalCountforagentLesthan) = "NO"
       counter = counter + 1
     Wend
    End If
    'If count is 1
    If totalCountforSragentLesthan = 1 Then
       Range("E" & countforSub & ":E" & totalCountforSragentLesthan + 16 + totalCountforagentLesthan) = "NO"
    End If
    
    'Set value
    Range("B" & countforSub - 1) = "Sr. Agent"
    'Set the total of FTEs for agent
    Range("C" & countforSub & ":C" & countforSub - 1) = DisCountforSragentLesthan - totalCountforSragentLesthan
    'Set 1 to number of FTEs
    Range("C" & countforSub & ":C" & countforSub + totalCountforSragentLesthan - 1) = 1
    'Set ~ # of days (if <144 hours or 18 days)
    Range("D" & countforSub - 1) = "~"
    'Set ~ # >= 144 hours or 18 days
    Range("E" & countforSub - 1) = "~"
    
    'Copy agent below 144
    Range("E400:FH" & totalCountforagentLesthan + 400).Select
    Selection.Copy
    'Paste and Select Range
    Range("A200").Select
    ActiveSheet.Paste
    
    'select and unfilter this row
    Rows("28:28").Select
    Selection.AutoFilter

    'Create border
     Range("B15:E" & countforSub).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
    Range("C15:E" & countforSub).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.SmallScroll Down:=-9
    
    'Bold Text
    Range("B15:E15").Select
    Selection.Font.Bold = True
    Range("B" & countforSub - 1 & ":E" & countforSub - 1).Select
    Selection.Font.Bold = True

    'remove data
    Rows("199:700").Select
    Selection.ClearContents
    Selection.Delete Shift:=xlUp

    'Create color coding
'FTEs with more than 144 hours or 18 days - Leadership and Support
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    LeadershipAndSupport = Selection.Cells.Count
    
    Rows("28:28").Select
    Selection.AutoFilter
    
    array1 = Range("C4").Value
    array2 = Range("C5").Value
    array3 = Range("C6").Value
    array4 = Range("C7").Value
    array5 = Range("C8").Value
    array6 = Range("C9").Value
    array7 = Range("C10").Value
    array8 = Range("C11").Value
    array9 = Range("C12").Value
    array10 = Range("C13").Value
   
    ActiveSheet.Range("$A$28:$FH$78").AutoFilter Field:=4, Criteria1:=Array(array1, array2, array3, array4, array5, array6, array7, array8, array9, array10), Operator:=xlFilterValues
    ActiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:=">=144", Operator:=xlFilterValues
    
    'Search Name and Copy employee data
    Range("A25").Select
    'search row
    Application.Goto Reference:="R28C2:R28C6"
    'Select and count all sr agent
    Range(Selection, Selection.End(xlDown)).Select
    
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'FTEs with more than 144 hours or 18 days - Agents
    'filter agent and more than 144 hrs
    ActiveSheet.Range("$A$28:$FH$78").AutoFilter Field:=4, Criteria1:=Array("Agent", "SR. Agent", array1, array2, array3, array4, array5, array6, array7, array8, array9, array10), Operator:=xlFilterValues
    ActiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues
    
    'Search Name and Copy employee data
    Range("A25").Select
    'search row
    Application.Goto Reference:="R28C2:R28C6"
    'Select and count all sr agent
    Range(Selection, Selection.End(xlDown)).Select

    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'unfilter
    'Selection.AutoFilter
     
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Application.Goto Reference:="R180C2"
    ActiveSheet.Paste
     
    Range("A25").Select
     Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Range(Selection, Selection.End(xlDown)).Select
    
     Range(Selection, Selection.End(xlDown)).Select
     countML = Selection.Cells.Count
    
    'ActiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues

startNum = 181
endNum = 179 + countML
PasteArea = 160
PasteArea1 = 197

Range("B181").Select
Do While startNum <= endNum
Do While PasteArea <= PasteArea1
    If Range("G" & startNum).Value = "ML" Then
        Range("G" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("L" & startNum).Value = "ML" Then
        Range("L" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("Q" & startNum).Value = "ML" Then
        Range("Q" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("V" & startNum).Value = "ML" Then
        Range("V" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("AA" & startNum).Value = "ML" Then
        Range("AA" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AF" & startNum).Value = "ML" Then
        Range("AF" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AK" & startNum).Value = "ML" Then
        Range("AK" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AP" & startNum).Value = "ML" Then
        Range("AP" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AU" & startNum).Value = "ML" Then
        Range("AU" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AZ" & startNum).Value = "ML" Then
        Range("AZ" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BE" & startNum).Value = "ML" Then
        Range("BE" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BJ" & startNum).Value = "ML" Then
        Range("BJ" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BO" & startNum).Value = "ML" Then
        Range("BO" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BT" & startNum).Value = "ML" Then
        Range("BT" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BY" & startNum).Value = "ML" Then
        Range("BY" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CD" & startNum).Value = "ML" Then
        Range("CD" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CI" & startNum).Value = "ML" Then
        Range("CI" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CN" & startNum).Value = "ML" Then
        Range("CN" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CS" & startNum).Value = "ML" Then
        Range("CS" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CX" & startNum).Value = "ML" Then
        Range("CX" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DC" & startNum).Value = "ML" Then
        Range("DC" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DH" & startNum).Value = "ML" Then
        Range("DH" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DM" & startNum).Value = "ML" Then
        Range("DM" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DR" & startNum).Value = "ML" Then
        Range("DR" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DW" & startNum).Value = "ML" Then
        Range("DW" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EB" & startNum).Value = "ML" Then
        Range("EB" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EG" & startNum).Value = "ML" Then
        Range("EG" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EL" & startNum).Value = "ML" Then
        Range("EL" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EQ" & startNum).Value = "ML" Then
        Range("EQ" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EV" & startNum).Value = "ML" Then
        Range("EV" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("FA" & startNum).Value = "ML" Then
        Range("FA" & startNum).Select
        Range("B" & startNum & ":F" & startNum).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    startNum = startNum + 1
    PasteArea = PasteArea + 1
Loop
Loop
   
    
''''''''''''''''''
    Rows("28:28").Select
    Selection.AutoFilter
    
   Marray1 = Range("B160").Value
   Marray2 = Range("B161").Value
   Marray3 = Range("B162").Value
   Marray4 = Range("B163").Value
   Marray5 = Range("B164").Value
   Marray6 = Range("B165").Value
   Marray7 = Range("B166").Value
   Marray8 = Range("B167").Value
   Marray9 = Range("B168").Value
   Marray10 = Range("B169").Value
   
    ActiveSheet.Range("$A$28:$FH$78").AutoFilter Field:=2, Criteria1:=Array(Marray1, Marray2, Marray3, Marray4, Marray5, Marray6, Marray7, Marray8, Marray9, Marray10), Operator:=xlFilterValues
    'ctiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:=">=144", Operator:=xlFilterValues
    
   Range("B160:FH230").Select
   Selection.EntireRow.Delete
   
     Cells.Find(What:="Days Equivalent", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
   
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
 'resigned
 'unfilter
 Rows("28:28").Select
 Selection.AutoFilter
 'filter
 Rows("28:28").Select
 Selection.AutoFilter
 ActiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues

   Range("A25").Select
    Cells.Find(What:="Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
 
    Selection.Copy
    Application.Goto Reference:="R180C2"
    ActiveSheet.Paste


     
    
   
    'ActiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:="<=143", Operator:=xlFilterValues

startNum2 = 181
endNum2 = 180 + countML
PasteArea2 = 160
PasteArea3 = 197

Range("B180").Select
Do While startNum2 <= endNum2
Do While PasteArea2 <= PasteArea3
    If Range("G" & startNum2).Value = "Resigned" Then
        Range("G" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("L" & startNum2).Value = "Resigned" Then
        Range("L" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("Q" & startNum2).Value = "Resigned" Then
        Range("Q" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("V" & startNum2).Value = "Resigned" Then
        Range("V" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("AA" & startNum2).Value = "Resigned" Then
        Range("AA" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AF" & startNum2).Value = "Resigned" Then
        Range("AF" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AK" & startNum2).Value = "Resigned" Then
        Range("AK" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AP" & startNum2).Value = "Resigned" Then
        Range("AP" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AU" & startNum2).Value = "Resigned" Then
        Range("AU" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("AZ" & startNum2).Value = "Resigned" Then
        Range("AZ" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BE" & startNum2).Value = "Resigned" Then
        Range("BE" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BJ" & startNum2).Value = "Resigned" Then
        Range("BJ" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BO" & startNum2).Value = "Resigned" Then
        Range("BO" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BT" & startNum2).Value = "Resigned" Then
        Range("BT" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("BY" & startNum2).Value = "Resigned" Then
        Range("BY" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CD" & startNum2).Value = "Resigned" Then
        Range("CD" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CI" & startNum2).Value = "Resigned" Then
        Range("CI" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CN" & startNum2).Value = "Resigned" Then
        Range("CN" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CS" & startNum2).Value = "Resigned" Then
        Range("CS" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("CX" & startNum2).Value = "Resigned" Then
        Range("CX" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DC" & startNum2).Value = "Resigned" Then
        Range("DC" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DH" & startNum2).Value = "Resigned" Then
        Range("DH" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DM" & startNum2).Value = "Resigned" Then
        Range("DM" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DR" & startNum2).Value = "Resigned" Then
        Range("DR" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
     If Range("DW" & startNum2).Value = "Resigned" Then
        Range("DW" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EB" & startNum2).Value = "Resigned" Then
        Range("EB" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EG" & startNum2).Value = "Resigned" Then
        Range("EG" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EL" & startNum2).Value = "Resigned" Then
        Range("EL" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EQ" & startNum2).Value = "Resigned" Then
        Range("EQ" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("EV" & startNum2).Value = "Resigned" Then
        Range("EV" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    If Range("FA" & startNum2).Value = "Resigned" Then
        Range("FA" & startNum2).Select
        Range("B" & startNum2 & ":F" & startNum2).Select
        Selection.Copy
        Range("B" & PasteArea).Select
        ActiveSheet.Paste
    Else
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    startNum2 = startNum2 + 1
    PasteArea2 = PasteArea2 + 1
Loop
Loop

''''''''''''''''''
   Rarray1 = Range("B198").Value
   Rarray2 = Range("B199").Value
   Rarray3 = Range("B200").Value
   Rarray4 = Range("B201").Value
   Rarray5 = Range("B202").Value
   Rarray6 = Range("B203").Value
   Rarray7 = Range("B204").Value
   Rarray8 = Range("B205").Value
   Rarray9 = Range("B206").Value
   Rarray10 = Range("B207").Value
   
   Rows("28:28").Select
   
    ActiveSheet.Range("$A$28:$FH$78").AutoFilter Field:=2, Criteria1:=Array(Rarray1, Rarray2, Rarray3, Rarray4, Rarray5, Rarray6, Rarray7, Rarray8, Rarray9, Rarray10), Operator:=xlFilterValues
    'ctiveSheet.Range("$B$28:$FH$78").AutoFilter Field:=5, Criteria1:=">=144", Operator:=xlFilterValues
    
   Range("B160:FH230").Select
   Selection.EntireRow.Delete
   
     Cells.Find(What:="Days Equivalent", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
   
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select


    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With





    
   'Display orig color
    Range("B28:F28").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
     
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
     Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
  
    Rows("28:28").Select
    Selection.AutoFilter

End Sub


