Sub Export_Macro()
' File name: Export_Test Macro
' Author: Erin Payne
' Description: File for additional data export.

    Rows.EntireRow.Hidden = False '******FOR TESTING ONLY*************
    
    Dim rowCount As Integer 'Row count of TT data
    rowCount = 0
    
    Dim copyStart As Integer 'Cell to start copy
    copyStart = 8
    
    Dim pasteLoc As Integer 'Cell to start paste
    
    Dim curRow As Integer 'Current cell of data being copied
    curRow = 9
    
    Dim inputSheet As String 'Sheet copying input from
    inputSheet = "FTE Input"
    
    Dim exportSheet As String 'Sheet pasting input to
    exportSheet = "Blank Sheet 2"
    
    'Counts row size of TT data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim ttRC As Integer
    ttRC = rowCount
    
    'TT FTE data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets(inputSheet).Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets(exportSheet).Activate                                               'initializes macro at "Blank Sheet 2" for pasting
    Range("A1").PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False                                    'Pastes data to blank sheet
    Range("A1").FormulaR1C1 = "Category"                                       'Renames "Labor" heading to "Category"
    Rows(1).Style = "Input Heading"                                            'Adds heading format to first row
    Range("A2").FormulaR1C1 = "TT FTE"                                         'Renames "Labor" to "TT FTE"
    Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(rowCount + 1, 1)), Type:=xlFillDefault
    copyStart = copyStart + 1
    pasteLoc = rowCount + 2
    
    'TT Base Labor Cost data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Base Labor Cost"                      'Renames "Labor" to "TT Base Labor Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

    'TT Cost COLA data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Cost COLA"                            'Renames "Labor" to "TT Cost COLA"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

    'TT Cost Contingency data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Cost Contingency"                     'Renames "Labor" to "TT Cost Contingency"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
    
    
    'Sets parameters for SS data copy
    copyStart = 32
    curRow = 32
    rowCount = 0

    'Counts row size of SS data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim ssRC As Integer
    ssRC = rowCount

    'SS FTE data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS FTE"                                  'Renames "Labor" to "SS FTE"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

    'SS Base Labor Cost data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Base Labor Cost"                       'Renames "Labor" to "SS Base Labor Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

    'SS Cost COLA data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Cost COLA"                             'Renames "Labor" to "SS Cost COLA"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

    'SS Cost Contingency data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Cost Contingency"                      'Renames "Labor" to "SS Cost Contingency"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
    'Saves start of "Other Input" for billing info move
    Dim otherIn As Integer
    otherIn = pasteLoc
    
    
    'Sets parameters for Travel data copy
    inputSheet = "Other Input"
    copyStart = 8
    curRow = 8
    rowCount = 0

    'Counts row size of Travel data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim travelRC As Integer
    travelRC = rowCount
    
    'Travel Cost data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Travel Cost"                                'Renames "Travel" to "Travel Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
    
    'Sets parameters for Other data copy
    copyStart = 21
    curRow = 21
    rowCount = 0

    'Counts row size of Other data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim otherRC As Integer
    otherRC = rowCount
    
    'Other Cost data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Other Cost"                              'Renames "Other" to "Other Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
    
    'Sets parameters for HW/SW data copy
    copyStart = 34
    curRow = 34
    rowCount = 0

    'Counts row size of HW/SW data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim hwswRC As Integer
    hwswRC = rowCount
    
    'HW/SW Cost data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "HW/SW Cost"                              'Renames "HW/SW" to "HW/SW Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
    
    'Inserts "LOB" column
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "LOB"
    'Inserts "Shore" column
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").FormulaR1C1 = "Shore"
    
    'Moves Other Input billing information to column B for unity
    Sheets(exportSheet).Range(Cells(otherIn, 18), Cells(pasteLoc - 1, 18)).Cut Range(Cells(otherIn, 2), Cells(pasteLoc - 1, 2))
    Application.CutCopyMode = False
    
    'Sets parameters for TT "LOB" information
    inputSheet = "FTE Input"
    copyStart = 9
    curRow = 2
    'TT LOB data copy
    Dim i As Integer
    For i = 1 To 4
        Sheets(inputSheet).Activate
        Range(Cells(copyStart, 164), Cells(copyStart + ttRC - 1, 164)).Copy  'Copies data for transfer
        Sheets(exportSheet).Activate
        Range(Cells(curRow, 5), Cells(curRow + ttRC - 1, 5)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                      'Pastes data to blank sheet
        curRow = curRow + ttRC                                               'Increments current row
    Next i
    
    'Sets parameters for SS "LOB" information
    copyStart = 32
    'SS LOB data copy
    For i = 1 To 4
        Sheets(inputSheet).Activate
        Range(Cells(copyStart, 164), Cells(copyStart + ssRC - 1, 164)).Copy  'Copies data for transfer
        Sheets(exportSheet).Activate
        Range(Cells(curRow, 5), Cells(curRow + ssRC - 1, 5)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                      'Pastes data to blank sheet
        curRow = curRow + ssRC                                               'Increments current row
    Next i
    
    'Moves Other Input LOB information to column E for unity
    Sheets(exportSheet).Range(Cells(otherIn, 10), Cells(pasteLoc - 1, 10)).Cut Range(Cells(otherIn, 5), Cells(pasteLoc - 1, 5))
    Application.CutCopyMode = False
    
    'Sets parameters for TT "Shore" information
    copyStart = 9
    curRow = 2
    'TT LOB data copy
    For i = 1 To 4
        Sheets(inputSheet).Activate
        Range(Cells(copyStart, 163), Cells(copyStart + ttRC - 1, 163)).Copy 'Copies data for transfer
        Sheets(exportSheet).Activate
        Range(Cells(curRow, 8), Cells(curRow + ttRC - 1, 8)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                     'Pastes data to blank sheet
        curRow = curRow + ttRC                                              'Increments current row
    Next i
    
    'Sets parameters for SS "Shore" information
    copyStart = 32
    'SS LOB data copy
    For i = 1 To 4
        Sheets(inputSheet).Activate
        Range(Cells(copyStart, 163), Cells(copyStart + ssRC - 1, 163)).Copy  'Copies data for transfer
        Sheets(exportSheet).Activate
        Range(Cells(curRow, 8), Cells(curRow + ssRC - 1, 8)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                      'Pastes data to blank sheet
        curRow = curRow + ssRC                                               'Increments current row
    Next i

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
    Columns("E:T").EntireColumn.AutoFit
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

Sub GetRowCount(inputSheet, curRow, rowCount)
' File name: GetRowCount
' Author: Erin Payne
' Description: Gets row count for any range of rows.

    While Sheets(inputSheet).Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
End Sub