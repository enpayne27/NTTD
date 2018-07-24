Sub Export_Macro()
' File name: Export_Test Macro
' Author: Erin Payne
' Description: File for additional data export.
    
    'Sheet copying input from
    Dim inputSheet As String
    inputSheet = "FTE Input"
    
    'Sheet pasting input to
    Dim exportSheet As String
    exportSheet = "Blank Sheet 2"
    
    Sheets(exportSheet).Cells.Clear 'Clears any previous data inputted on exportSheet
    Rows.EntireRow.Hidden = False
    
    'Initializes all variables used from Name Manager
    HWInp_FirstRow = Int(Right(ActiveWorkbook.Names("HWInp_FirstRow"), 3))  'Takes row number from formula and converts to equivalent integer
    HWInp_MaxRow = Application.Evaluate("HWInp_MaxRow")                     'Takes total number of rows for given section
    OthInp_FirstRow = Int(Right(ActiveWorkbook.Names("OthInp_FirstRow"), 3))
    OthInp_MaxRow = Application.Evaluate("OthInp_MaxRow")
    SSInp_FirstRow = Int(Right(ActiveWorkbook.Names("SSInp_FirstRow"), 3))
    SSInp_MaxRow = Application.Evaluate("SSInp_MaxRow")
    TTInp_FirstRow = Int(Right(ActiveWorkbook.Names("TTInp_FirstRow"), 2))
    TTInp_MaxRow = Application.Evaluate("TTInp_MaxRow")
    TVLInp_FirstRow = Int(Right(ActiveWorkbook.Names("TVLInp_FirstRow"), 2))
    TVLInp_MaxRow = Application.Evaluate("TVLInp_MaxRow")
    termLength = Application.Evaluate("TermLength")
    
    'Row count of data being transferred
    Dim rowCount As Integer
    rowCount = TTInp_MaxRow 'Sets row size to TT row count for first copy
    
    'Cell row to start copy on import sheet
    Dim copyStart As Integer
    copyStart = TTInp_FirstRow
    
    'Cell row to start paste on export sheet
    Dim pasteLoc As Integer
    pasteLoc = 1
    
    'Category title of data inputted
    Dim category As String
    
    'TT FTE data copy
    Sheets(inputSheet).Activate                                           'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(rowCount + copyStart, termLength)).Copy          'Copies TT data for transfer
    Sheets(exportSheet).Activate                                          'initializes macro at "Blank Sheet 2" for pasting
    Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False                               'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Category"                           'Renames "Labor" heading to "Category"
    Rows(pasteLoc).Style = "Input Heading"                                'Adds heading format to first row
    Cells(pasteLoc + 1, 1).FormulaR1C1 = "TT FTE"                         'Renames "Labor" to "TT FTE"
    Cells(pasteLoc + 1, 1).AutoFill Destination:=Range(Cells(pasteLoc + 1, 1), _
        Cells(rowCount + 1, 1)), Type:=xlFillDefault
    
    copyStart = copyStart + 1 'Incremented by 1 to exclude heading line
    pasteLoc = rowCount + 2   'Incremented by 1 to exclude heading line and begin on new line
    
    'TT Base Labor Cost data copy
    category = "TT Base Labor Cost"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'TT Cost COLA data copy
    category = "TT Cost COLA"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'TT Cost Contingency data copy
    category = "TT Cost Contingency"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)
    

    'Sets row size of SS data
    rowCount = SSInp_MaxRow
    copyStart = SSInp_FirstRow + 1 'Incremented by 1 to exclude heading line

    'SS FTE data copy
    category = "SS FTE"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'SS Base Labor Cost data copy
    category = "SS Base Labor Cost"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'SS Cost COLA data copy
    category = "SS Cost COLA"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'SS Cost Contingency data copy
    category = "SS Cost Contingency"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)
    
    'Saves start of "Other Input" in exportSheet for billing info move
    Dim otherIn As Integer
    otherIn = pasteLoc
    
    'Sets parameters for "Other" data copy
    inputSheet = "Other Input"

    'Travel Cost data copy
    rowCount = TVLInp_MaxRow
    copyStart = TVLInp_FirstRow + 1
    category = "Travel Cost"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)

    'Other Cost data copy
    rowCount = OthInp_MaxRow
    copyStart = OthInp_FirstRow + 1
    category = "Other Cost"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)
    
    'HW/SW Cost data copy
    rowCount = HWInp_MaxRow
    copyStart = HWInp_FirstRow + 1
    category = "HW/SW Cost"
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)
    
    'Inserts "LOB" column
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "LOB"
    'Inserts "Shore" column
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").FormulaR1C1 = "Shore"
    
    'Moves Other Input billing information to column B for unity
    Sheets(exportSheet).Range(Cells(otherIn, 18), Cells(pasteLoc - 1, 18)).Cut Range(Cells(otherIn, 2), Cells(pasteLoc - 1, 2))
     Application.CutCopyMode = False
    
    Dim copyCol As Integer
    Dim pasteCol As Integer
    Dim RC As Integer
    
    'Sets parameters for TT "LOB" information
    inputSheet = "FTE Input"
    copyStart = 9
    curRow = 2
    copyCol = 164
    pasteCol = 5
    RC = TTInp_MaxRow
    'TT LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
    
    'Sets parameters for SS "LOB" information
    copyStart = 32
    RC = SSInp_MaxRow
    'SS LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
    
    'Moves Other Input LOB information to column E for unity
    Sheets(exportSheet).Range(Cells(otherIn, 10), Cells(pasteLoc - 1, 10)).Cut Range(Cells(otherIn, 5), Cells(pasteLoc - 1, 5))
    Application.CutCopyMode = False
    
    'Sets parameters for TT "Shore" information
    copyStart = 9
    curRow = 2
    copyCol = 163
    pasteCol = 8
    RC = TTInp_MaxRow
    'TT LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
    
    'Sets parameters for SS "Shore" information
    copyStart = 32
    RC = SSInp_MaxRow
    'SS LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)

    'Moves Other Input Shore information to column H for unity
    Sheets(exportSheet).Range(Cells(otherIn, 9), Cells(pasteLoc - 1, 9)).Cut Range(Cells(otherIn, 8), Cells(pasteLoc - 1, 8))
    Application.CutCopyMode = False
    
    'Hides blank rows of sheet
    Dim rng As Range
    For Each rng In Range(Cells(2, 3), Cells(pasteLoc - 1, 3))
        If rng.Value = "" Then
            rng.EntireRow.Hidden = True
        Else
            rng.EntireRow.Hidden = False
        End If
    Next rng
    
    'Autofits column size for legibility
    Columns("A:EH").EntireColumn.AutoFit
End Sub

Sub SetTopBorder(pasteLoc)
' File name: Set_Top_Border Macro
' Author: Erin Payne
' Description: Sets a thin top border to any cell.

    With Rows(pasteLoc).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
End Sub

Sub GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, category, termLength)
' File name: GetData
' Author: Erin Payne
' Description: Copys, pastes, and renames input data for export

    Sheets(inputSheet).Activate                                                'initializes macro at input sheet
    'Copies and pastes information to blank sheet
    Range(Cells(copyStart, 1), Cells(rowCount + copyStart - 1, termLength)).Copy           'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = category                                  'Renames cateogry title
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

End Sub

Sub GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
' File name: GetNewData
' Author: Erin Payne
' Description: Copys, pastes, and renames LOB and Shore data for export
    Dim i As Integer
    For i = 1 To 4
            Sheets(inputSheet).Activate
            Range(Cells(copyStart, copyCol), Cells(copyStart + RC - 1, copyCol)).Copy  'Copies data for transfer
            Sheets(exportSheet).Activate
            Range(Cells(curRow, pasteCol), Cells(curRow + RC - 1, pasteCol)).PasteSpecial _
                Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False                            'Pastes data to blank sheet
            curRow = curRow + RC                                                       'Increments current row
    Next i
End Sub

Sub GetCost()
'Calculates monthly costs
' Sub name: GetCost
' Author: Erin Payne
' Description: Calculates monthly costs
    Dim tthc As Range
    hc = TTInp_HC
    
    Dim rate As Range
    rate = TTInp_Modeled_Cost_Rates
    
    Dim hrs As Range
    hrs = TTInp_Mthly_Cost_Hrs
    
    Dim ans As Range 'Need to be an integer?
    ans = hc * rate * hrs
    
    Sheets(exportSheet).Range(Cells(pasteRow, pasteCol), Cells(pasteRow + rowCount, pasteCol + colCount)).Value = ans
    'Range().Style = "Currency"
    
End Sub
