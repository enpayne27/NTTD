Sub Export_Test()
' File name: Export_Test Macro
' Author: Erin Payne
' Description: Final test file for additional data export.

    Rows.EntireRow.Hidden = False '******FOR TESTING ONLY*************
    
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
    Dim ttRC As Integer
    ttRC = rowCount
    'MsgBox ("TT Rows = " & rowCount)   'Displays row count
    
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
    Dim ssRC As Integer
    ssRC = rowCount
    'MsgBox ("SS Rows = " & rowCount)   'Displays row count

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
    Dim travelRC As Integer
    travelRC = rowCount
    'MsgBox ("Travel Rows = " & rowCount)   'Displays row count
    
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
    Dim otherRC As Integer
    otherRC = rowCount
    'MsgBox ("Other Rows = " & rowCount)   'Displays row count
    
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
    Dim hwswRC As Integer
    hwswRC = rowCount
    'MsgBox ("HW/SW Rows = " & rowCount)   'Displays row count
    
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
    
    'Sets parameters for TT "LOB" information
    copyStart = 9
    curRow = 2
    'TT LOB data copy
    Dim i As Integer
    For i = 1 To 4
        Sheets("FTE Input").Activate
        Sheets("FTE Input").Range(Cells(copyStart, 164), Cells(copyStart + ttRC - 1, 164)).Copy 'Copies data for transfer
        Sheets("Blank Sheet 2").Activate
        Sheets("Blank Sheet 2").Range(Cells(curRow, 5), Cells(curRow + ttRC - 1, 5)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                                         'Pastes data to blank sheet
        curRow = curRow + ttRC                                                                  'Increments current row
    Next i
    
    'Sets parameters for SS "LOB" information
    copyStart = 32
    'SS LOB data copy
    For i = 1 To 4
        Sheets("FTE Input").Activate
        Sheets("FTE Input").Range(Cells(copyStart, 164), Cells(copyStart + ssRC - 1, 164)).Copy 'Copies data for transfer
        Sheets("Blank Sheet 2").Activate
        Sheets("Blank Sheet 2").Range(Cells(curRow, 5), Cells(curRow + ssRC - 1, 5)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                                         'Pastes data to blank sheet
        curRow = curRow + ssRC                                                                  'Increments current row
    Next i
    
    'Moves Other Input LOB information to column E for unity
    Sheets("Blank Sheet 2").Range(Cells(otherIn, 10), Cells(pasteLoc - 1, 10)).Cut Range(Cells(otherIn, 5), Cells(pasteLoc - 1, 5))
    Application.CutCopyMode = False
    
    'Sets parameters for TT "Shore" information
    copyStart = 9
    curRow = 2
    'TT LOB data copy
    For i = 1 To 4
        Sheets("FTE Input").Activate
        Sheets("FTE Input").Range(Cells(copyStart, 163), Cells(copyStart + ttRC - 1, 163)).Copy 'Copies data for transfer
        Sheets("Blank Sheet 2").Activate
        Sheets("Blank Sheet 2").Range(Cells(curRow, 8), Cells(curRow + ttRC - 1, 8)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                                         'Pastes data to blank sheet
        curRow = curRow + ttRC                                                                  'Increments current row
    Next i
    
    'Sets parameters for SS "Shore" information
    copyStart = 32
    'SS LOB data copy
    For i = 1 To 4
        Sheets("FTE Input").Activate
        Sheets("FTE Input").Range(Cells(copyStart, 163), Cells(copyStart + ssRC - 1, 163)).Copy 'Copies data for transfer
        Sheets("Blank Sheet 2").Activate
        Sheets("Blank Sheet 2").Range(Cells(curRow, 8), Cells(curRow + ssRC - 1, 8)).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False                                         'Pastes data to blank sheet
        curRow = curRow + ssRC                                                                  'Increments current row
    Next i

    'Moves Other Input Shore information to column H for unity
    Sheets("Blank Sheet 2").Range(Cells(otherIn, 9), Cells(pasteLoc - 1, 9)).Cut Range(Cells(otherIn, 8), Cells(pasteLoc - 1, 8))
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
End Sub