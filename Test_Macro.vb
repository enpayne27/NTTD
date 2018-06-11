Sub Export_Macro()
' File name: Export_Test Macro
' Author: Erin Payne
' Description: File for additional data export.

    Sheets("Blank Sheet 2").Range("A1:EH200").Clear '******FOR TESTING ONLY*************
    
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
    
    Dim category As String 'Category title of data inputted
    
    'Counts row size of TT data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim ttRC As Integer 'TT row count
    ttRC = rowCount
    
    'TT FTE data copy
    Sheets(inputSheet).Activate                                                'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies TT data for transfer
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
    category = "TT Base Labor Cost"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)

    'TT Cost COLA data copy
    category = "TT Cost COLA"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)

    'TT Cost Contingency data copy
    category = "TT Cost Contingency"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
    
    
    'Sets parameters for SS data copy
    copyStart = 32
    curRow = 32
    rowCount = 0

    'Counts row size of SS data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim ssRC As Integer 'SS row count
    ssRC = rowCount

    'SS FTE data copy
    category = "SS FTE"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)

    'SS Base Labor Cost data copy
    category = "SS Base Labor Cost"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)

    'SS Cost COLA data copy
    category = "SS Cost COLA"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)

    'SS Cost Contingency data copy
    category = "SS Cost Contingency"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
    
    'Saves start of "Other Input" for billing info move
    Dim otherIn As Integer 'Starting position of "Other Input"
    otherIn = pasteLoc
    
    'Sets parameters for Travel data copy
    inputSheet = "Other Input"
    copyStart = 8
    curRow = 8
    rowCount = 0

    'Counts row size of Travel data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim travelRC As Integer 'Travel row count
    travelRC = rowCount
    
    'Travel Cost data copy
    category = "Travel Cost"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
    
    'Sets parameters for Other data copy
    copyStart = 21
    curRow = 21
    rowCount = 0

    'Counts row size of Other data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim otherRC As Integer 'Other row count
    otherRC = rowCount
    
    'Other Cost data copy
    category = "Other Cost"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
    
    'Sets parameters for HW/SW data copy
    copyStart = 34
    curRow = 34
    rowCount = 0

    'Counts row size of HW/SW data
    Call GetRowCount(inputSheet, curRow, rowCount)
    Dim hwswRC As Integer 'HW/SW row count
    hwswRC = rowCount
    
    'HW/SW Cost data copy
    category = "HW/SW Cost"
    Call GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
    
    'Inserts "LOB" column
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "LOB"
    'Inserts "Shore" column
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").FormulaR1C1 = "Shore"
    
    'Moves Other Input billing information to column B for unity
    Sheets(exportSheet).Range(Cells(otherIn, 18), Cells(pasteLoc - 1, 18)).Cut Range(Cells(otherIn, 2), Cells(pasteLoc - 1, 2))
    Application.CutCopyMode = False
    
    Dim copyCol As Integer 'Column letter for data copy
    Dim pasteCol As Integer 'Column letter for data paste
    Dim RC As Integer 'Row count of data being transferred
    
    'Sets parameters for TT "LOB" information
    inputSheet = "FTE Input"
    copyStart = 9
    curRow = 2
    copyCol = 164
    pasteCol = 5
    RC = ttRC
    'TT LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
    
    'Sets parameters for SS "LOB" information
    copyStart = 32
    RC = ssRC
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
    RC = ttRC
    'TT LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
    
    'Sets parameters for SS "Shore" information
    copyStart = 32
    RC = ssRC
    'SS LOB data copy
    Call GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)

    'Moves Other Input Shore information to column H for unity
    Sheets(exportSheet).Range(Cells(otherIn, 9), Cells(pasteLoc - 1, 9)).Cut Range(Cells(otherIn, 8), Cells(pasteLoc - 1, 8))
    Application.CutCopyMode = False
    
    'Calculates monthly costs
    copyStart = 9
    curRow = 2
    RC = ttRC + curRow - 1
    Call GetCost(inputSheet, exportSheet, copyStart, curRow, RC)

    copyStart = 9
    curRow = curRow + ttRC
    RC = ttRC + curRow - 1
    Call GetCost(inputSheet, exportSheet, copyStart, curRow, RC)

'    copyStart = 9
'    curRow = curRow + ttRC
'    RC = ttRC + curRow - 1
'    Call GetCost(inputSheet, exportSheet, copyStart, curRow, RC)
'
'    copyStart = 9
'    curRow = 2
'    Call GetCost(inputSheet, exportSheet, copyStart, curRow, RC)

    
'    'Hides blank rows of sheet
'    Dim rng As Range
'    For Each rng In Range(Cells(2, 3), Cells(pasteLoc - 1, 3))
'        If rng.Value = "" Then
'            rng.EntireRow.Hidden = True
'        Else
'            rng.EntireRow.Hidden = False
'        End If
'    Next rng

    'Autofits column size for legibility
    Columns("E:EH").EntireColumn.AutoFit
End Sub

Sub SetTopBorder(pasteLoc)
' Sub name: SetTopBorder Macro
' Author: Erin Payne
' Description: Sets a thin top border to any cell.

    With Rows(pasteLoc).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
End Sub

Sub GetRowCount(inputSheet, curRow, rowCount)
' Sub name: GetRowCount
' Author: Erin Payne
' Description: Gets row count for any range of rows.

    While Sheets(inputSheet).Cells(curRow, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        curRow = curRow + 1
    Wend
End Sub

Sub GetData(inputSheet, exportSheet, copyStart, curRow, pasteLoc, rowCount, category)
' Sub name: GetData
' Author: Erin Payne
' Desctiption: Copys, pastes, and renames input data for export

    Sheets(inputSheet).Activate                                                'initializes macro at input sheet
    'Copies and pastes information to blank sheet
    Range(Cells(copyStart, 1), Cells(curRow - 1, 138)).Copy                    'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                 'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = category                                  'Renames cateogry title
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount

End Sub

Sub GetNewData(inputSheet, exportSheet, copyStart, RC, copyCol, pasteCol, curRow)
' Sub name: GetNewData
' Author: Erin Payne
' Desctiption: Copys, pastes, and renames LOB and Shore data for export

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

Sub GetCost(inputSheet, exportSheet, copyStart, curRow, RC)
'Calculates monthly costs
' Sub name: GetCost
' Author: Erin Payne
' Desctiption: Calculates monthly costs

    Dim ans As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim i As Integer
    Dim j As Integer
    For i = 21 To 26 'Columns '*************Change month end value************
        For j = curRow To RC 'Rows
            If Sheets(exportSheet).Cells(j, i).Value <> "" Then
                x = Sheets(inputSheet).Cells(copyStart, 19).Value 'S
                y = Sheets(inputSheet).Cells(copyStart, 153).Value 'EW
                z = Sheets(inputSheet).Cells(copyStart, 166).Value 'FJ
                ans = x * y
                ans = ans * z
                Sheets(exportSheet).Cells(j, i).Value = ans
                Cells(j, i).Style = "Currency"
                copyStart = copyStart + 1
            Else 'If blank
                copyStart = copyStart + 1
            End If
        Next j
        copyStart = 9
    Next i
End Sub