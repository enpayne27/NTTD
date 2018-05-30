Sub Test_Macro()
' File name: Test_Macro Macro
' Author: Erin Payne
' Description: Final test file for additional data export.

    Dim rowCount As Integer 'Row count of TT data
    rowCount = 0
    
    Dim copyStart As Integer 'Cell to start copy
    copyStart = 8
    
    Dim pasteLoc As Integer 'Cell to start paste
    
    Dim curRow As Integer 'Current cell of data being copied
    curRow = 9
    
    'Counts row size of TT data
    While Sheets("FTE Input").Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
    MsgBox ("TT Rows = " & rowCount)   'Displays row count
    
    'TT FTE data copy
    Sheets("FTE Input").Activate                                    'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate                                'initializes macro at "Blank Sheet 2" for pasting
    Sheets("Blank Sheet 2").Range("A1").PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Range("A1").FormulaR1C1 = "Category"                            'Renames "Labor" heading to "Category"
    Rows(1).Style = "Input Heading"                                 'Adds heading format to first row
    Range("A2").FormulaR1C1 = "TT FTE"                              'Renames "Labor" to "TT FTE"
    Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(rowCount + 1, 1)), Type:=xlFillDefault
    copyStart = copyStart + 1
    pasteLoc = rowCount + 2
    
    'TT Base Labor Cost data copy
    Sheets("FTE Input").Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                             'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Base Labor Cost"                       'Renames "Labor" to "TT Base Labor Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount

    'TT Cost COLA data copy
    Sheets("FTE Input").Activate                                                 'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy  'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                              'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Cost COLA"                              'Renames "Labor" to "TT Cost COLA"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount

    'TT Cost Contingency data copy
    Sheets("FTE Input").Activate                                                 'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy  'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                              'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "TT Cost Contingency"                       'Renames "Labor" to "TT Cost Contingency"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount
    
    
    'Sets parameters for SS data copy
    copyStart = 32
    curRow = 32
    rowCount = 0

    'Counts row size of SS data
    While Sheets("FTE Input").Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
    MsgBox ("SS Rows = " & rowCount)   'Displays row count

    'SS FTE data copy
    Sheets("FTE Input").Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate                                            'initializes macro at "Blank Sheet 2" for pasting
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                             'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS FTE"                                   'Renames "Labor" to "SS FTE"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount

    'SS Base Labor Cost data copy
    Sheets("FTE Input").Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                             'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Base Labor Cost"                       'Renames "Labor" to "SS Base Labor Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount

    'SS Cost COLA data copy
    Sheets("FTE Input").Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                             'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Cost COLA"                             'Renames "Labor" to "SS Cost COLA"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount

    'SS Cost Contingency data copy
    Sheets("FTE Input").Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("FTE Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies TT data for transfer
    Sheets("Blank Sheet 2").Activate
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                             'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "SS Cost Contingency"                      'Renames "Labor" to "SS Cost Contingency"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount
    'Saves start of "Other Input" for billing info move
    Dim otherIn As Integer
    otherIn = pasteLoc
    
    
    'Sets parameters for Travel data copy
    copyStart = 8
    curRow = 8
    rowCount = 0

    'Counts row size of Travel data
    While Sheets("Other Input").Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
    MsgBox ("Travel Rows = " & rowCount)   'Displays row count
    
    'Travel Cost data copy
    Sheets("Other Input").Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Other Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies Travel data for transfer
    Sheets("Blank Sheet 2").Activate                                              'initializes macro at "Blank Sheet 2" for pasting
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                               'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Travel Cost"                                'Renames "Travel" to "Travel Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount
    
    'Sets parameters for Other data copy
    copyStart = 21
    curRow = 21
    rowCount = 0

    'Counts row size of Other data
    While Sheets("Other Input").Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
    MsgBox ("Other Rows = " & rowCount)   'Displays row count
    
    'Other Cost data copy
    Sheets("Other Input").Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Other Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies Other data for transfer
    Sheets("Blank Sheet 2").Activate                                              'initializes macro at "Blank Sheet 2" for pasting
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                               'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Other Cost"                                 'Renames "Other" to "Other Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount
    
    'Sets parameters for HW/SW data copy
    copyStart = 34
    curRow = 34
    rowCount = 0

    'Counts row size of HW/SW data
    While Sheets("Other Input").Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
    MsgBox ("HW/SW Rows = " & rowCount)   'Displays row count
    
    'HW/SW Cost data copy
    Sheets("Other Input").Activate                                                'initializes macro at "Other Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Other Input").Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy 'Copies HW/SW data for transfer
    Sheets("Blank Sheet 2").Activate                                              'initializes macro at "Blank Sheet 2" for pasting
    Sheets("Blank Sheet 2").Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False                               'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "HW/SW Cost"                                 'Renames "HW/SW" to "HW/SW Cost"
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    pasteLoc = pasteLoc + rowCount
    
    'Inserts "LOB" column
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "LOB"
    'Inserts "Shore" column
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").FormulaR1C1 = "Shore"
    
    'Moves Other Input billing information to column B for unity
    Sheets("Blank Sheet 2").Range(Cells(otherIn, 18), Cells(pasteLoc - 1, 18)).Cut Range(Cells(otherIn, 2), Cells(pasteLoc - 1, 2))
    Application.CutCopyMode = False
End Sub
