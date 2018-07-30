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
    SSInp_LOB = Application.Evaluate("SSInp_LOB")
    SSInp_MaxRow = Application.Evaluate("SSInp_MaxRow")
    SSInp_Shore = Application.Evaluate("SSInp_Shore")
    
    TTInp_FirstRow = Int(Right(ActiveWorkbook.Names("TTInp_FirstRow"), 2))
    TTInp_LOB = Application.Evaluate("TTInp_LOB")
    TTInp_MaxRow = Application.Evaluate("TTInp_MaxRow")
    TTInp_Shore = Application.Evaluate("TTInp_Shore")
    
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
    Sheets(inputSheet).Activate                                                    'initializes macro at "FTE Input" sheet
    'Copies and pastes TT FTE information to blank sheet
    Range(Cells(copyStart, 1), Cells(rowCount + copyStart, 21 + termLength)).Copy      'Copies TT data for transfer
    Sheets(exportSheet).Activate                                                   'initializes macro at "Blank Sheet 2" for pasting
    Cells(pasteLoc, 1).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False                                        'Pastes data to blank sheet
    Cells(pasteLoc, 1).FormulaR1C1 = "Category"                                    'Renames "Labor" heading to "Category"
    Rows(pasteLoc).Style = "Input Heading"                                         'Adds heading format to first row
    Cells(pasteLoc + 1, 1).FormulaR1C1 = "TT FTE"                                  'Renames "Labor" to "TT FTE"
    Cells(pasteLoc + 1, 1).AutoFill Destination:=Range(Cells(pasteLoc + 1, 1), _
        Cells(rowCount + 1, 1)), Type:=xlFillDefault
    
    copyStart = copyStart + 1 'Incremented by 1 to exclude heading line
    pasteLoc = rowCount + 2   'Incremented by 1 to exclude heading line and begin on new line
    
    'TT Base Labor Cost data copy
    Dim TT_Cost_Base As Integer 'Starting row number of TT Base Labor Cost section
    TT_Cost_Base = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "TT Base Labor Cost", termLength)
    '*Call cost calculation here

    'TT Cost COLA data copy
    Dim TT_Cost_COLA As Integer 'Starting row number of TT Cost COLA section
    TT_Cost_COLA = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "TT Cost COLA", termLength)
    
    'TT Cost Contingency data copy
    Dim TT_Cost_Cont As Integer 'Starting row number of TT Cost Contingency section
    TT_Cost_Cont = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "TT Cost Contingency", termLength)
 
    'Sets parameters for SS data copy
    rowCount = SSInp_MaxRow        'Sets row size of SS data
    copyStart = SSInp_FirstRow + 1 'Incremented by 1 to exclude heading line

    'SS FTE data copy
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "SS FTE", termLength)

    'SS Base Labor Cost data copy
    Dim SS_Cost_Base As Integer 'Starting row number of SS Base Labor Cost section
    SS_Cost_Base = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "SS Base Labor Cost", termLength)

    'SS Cost COLA data copy
    Dim SS_Cost_COLA As Integer 'Starting row number of SS Cost COLA section
    SS_Cost_COLA = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "SS Cost COLA", termLength)

    'SS Cost Contingency data copy
    Dim SS_Cost_Cont As Integer 'Starting row number of SS Cost Contingency section
    SS_Cost_Cont = pasteLoc
    Call GetData(inputSheet, exportSheet, copyStart, pasteLoc, rowCount, "SS Cost Contingency", termLength)
    
    
    'Saves start of "Other Input" in exportSheet for billing info move
    Dim pasteRow As Integer
    pasteRow = pasteLoc
    
    'Sets parameters for "Other" data copy
    inputSheet = "Other Input"

    'Travel Cost data copy
    Call GetData(inputSheet, exportSheet, TVLInp_FirstRow + 1, pasteLoc, TVLInp_MaxRow, "Travel Cost", termLength)

    'Other Cost data copy
    Call GetData(inputSheet, exportSheet, OthInp_FirstRow + 1, pasteLoc, OthInp_MaxRow, "Other Cost", termLength)
    
    'HW/SW Cost data copy
    Call GetData(inputSheet, exportSheet, HWInp_FirstRow + 1, pasteLoc, HWInp_MaxRow, "HW/SW Cost", termLength)
    
    
    'Inserts "LOB" column
    Dim LOBCat As Integer 'Column number for LOB category in exportSheet
    LOBCat = 5            'Column "E"
    Columns(LOBCat).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, LOBCat).FormulaR1C1 = "LOB"
    
    'Inserts "Shore" column
    Dim ShoreCat As Integer 'Column number for Shore category in exportSheet
    ShoreCat = 8            'Column "H"
    Columns(ShoreCat).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, ShoreCat).FormulaR1C1 = "Shore"
    
    'Moves Other Input billing information from R to column B for unity
    Dim BillCat As Integer 'Column number for Billing category in exportSheet
    BillCat = 2
    copyStart = 18
    Sheets(exportSheet).Range(Cells(pasteRow, copyStart), Cells(pasteLoc - 1, copyStart)).Cut Range(Cells(pasteRow, BillCat), Cells(pasteLoc - 1, BillCat))
    Application.CutCopyMode = False
    

    'Sets parameters for LOB data copy
    Dim pasteCol As Integer
    pasteRow = 2
    pasteCol = LOBCat
    
    'TT LOB data copy
    Dim i As Integer
     For i = 1 To 4
        Call GetNewData(pasteRow, pasteCol, TTInp_LOB, TTInp_MaxRow)
    Next i
    
    'SS LOB data copy
    For i = 1 To 4
        Call GetNewData(pasteRow, pasteCol, SSInp_LOB, SSInp_MaxRow)
    Next i
    
    'Moves Other Input LOB information from column J to column E for unity
    copyStart = 10
    Sheets(exportSheet).Range(Cells(pasteRow, copyStart), Cells(pasteLoc - 1, copyStart)).Cut Range(Cells(pasteRow, LOBCat), Cells(pasteLoc - 1, LOBCat))
    Application.CutCopyMode = False
    
    
    'Sets parameters for Shore information
    pasteRow = 2
    pasteCol = ShoreCat
    
    'TT Shore data copy
    For i = 1 To 4
        Call GetNewData(pasteRow, pasteCol, TTInp_Shore, TTInp_MaxRow)
    Next i
    
    'SS LOB data copy
    For i = 1 To 4
        Call GetNewData(pasteRow, pasteCol, SSInp_Shore, SSInp_MaxRow)
    Next i
    
    'Moves Other Input Shore information to column H for unity
    copyStart = 9
    Sheets(exportSheet).Range(Cells(pasteRow, copyStart), Cells(pasteLoc - 1, copyStart)).Cut Range(Cells(pasteRow, ShoreCat), Cells(pasteLoc - 1, ShoreCat))
    Application.CutCopyMode = False
    
    
    'Imports HC data for COLA and Contingency subsections
    inputSheet = "FTE Input"
    pasteCol = 21
    
    'TT HC data copy
    Sheets(inputSheet).Range("TTInp_HC_ExpCOLA").Copy
    Sheets(exportSheet).Range(Cells(TT_Cost_COLA, pasteCol), Cells(TT_Cost_COLA + TTInp_MaxRow, pasteCol + termLength - 1)).PasteSpecial xlPasteValues 'TT COLA HC data copy
    Sheets(inputSheet).Range("TTInp_HC_Cont").Copy
    Sheets(exportSheet).Range(Cells(TT_Cost_Cont, pasteCol), Cells(TT_Cost_Cont + TTInp_MaxRow, pasteCol + termLength - 1)).PasteSpecial xlPasteValues 'TT Contingency HC data copy
    
    'SS HC data copy
    Sheets(inputSheet).Range("SSInp_HC_ExpCOLA").Copy
    Sheets(exportSheet).Range(Cells(SS_Cost_COLA, pasteCol), Cells(SS_Cost_COLA + SSInp_MaxRow, pasteCol + termLength - 1)).PasteSpecial xlPasteValues 'SS COLA HC data copy
    Sheets(inputSheet).Range("SSInp_HC_Cont").Copy
    Sheets(exportSheet).Range(Cells(SS_Cost_Cont, pasteCol), Cells(SS_Cost_Cont + SSInp_MaxRow, pasteCol + termLength - 1)).PasteSpecial xlPasteValues 'SS Contingency HC data copy


    'TT cost calculation
    importSheet = "FTE Input"
    rowCount = TTInp_MaxRow
    copyRow = 9                 'Begins calculations at row 9 on import sheet
    pasteRow = 2 + rowCount     'Begins calculation paste after first set of TT FTE data
    pasteCol = 21               'Begins calculation paste at column 21 on export sheet
    hrsCol = 166                'Column including monthly cost hours data
    rateCol = 153               'Column including cost rate data
    
    Call GetCost(importSheet, exportSheet, copyRow, 19, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'TT Base Labor cost calculation
    Call GetCost(importSheet, exportSheet, copyRow, 315, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'TT COLA cost calculation
    Call GetCost(importSheet, exportSheet, copyRow, 436, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'TT Contingency calculation
    
    'SS cost calculation
    copyRow = 31
    rowCount = SSInp_MaxRow
    pasteRow = pasteRow + rowCount
    Call GetCost(importSheet, exportSheet, copyRow, 19, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'SS Base Labor cost calculation
    Call GetCost(importSheet, exportSheet, copyRow, 315, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'SS COLA cost calculation
    Call GetCost(importSheet, exportSheet, copyRow, 436, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength) 'SS Contingency calculation

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
' Sub name: GetData
' Author: Erin Payne
' Description: Copies, pastes, and renames input data for export.

    Sheets(inputSheet).Activate                                                  'initializes macro at input sheet
    'Copies and pastes information to blank sheet
    Range(Cells(copyStart, 1), Cells(rowCount + copyStart - 1, 21 + termLength)).Copy 'Copies data for transfer
    Sheets(exportSheet).Activate
    Cells(pasteLoc, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False                   'Pastes data to export sheet
    Cells(pasteLoc, 1).FormulaR1C1 = category                                    'Renames category title
    Cells(pasteLoc, 1).AutoFill Destination:=Range(Cells(pasteLoc, 1), Cells(pasteLoc + rowCount - 1, 1)), Type:=xlFillDefault
    Call SetTopBorder(pasteLoc)
    pasteLoc = pasteLoc + rowCount
End Sub

Sub GetNewData(pasteRow, pasteCol, pasteData, rowCount)
' Sub name: GetNewData
' Author: Erin Payne
' Description: Copies, pastes, and renames LOB and Shore data for export.

    'Pastes data to blank sheet
    Range(Cells(pasteRow, pasteCol), Cells(pasteRow + rowCount - 1, pasteCol)).FormulaR1C1 = pasteData
    pasteRow = pasteRow + rowCount
End Sub

Sub GetCost(importSheet, exportSheet, copyRow, copyCol, pasteRow, pasteCol, hrsCol, rateCol, rowCount, termLength)
' Sub name: GetCost
' Author: Erin Payne
' Description: Calculates monthly costs.
    
    Dim hc As Long 'Head count to be calculated
    Dim rate As Long 'Cost rate to be calculated
    Dim hrs As Long 'Cost hours to be calculated

    'Navigates through rows
    For i = copyRow To copyRow + rowCount - 1
        hrs = Sheets(importSheet).Cells(i, hrsCol).Value
        rate = Sheets(importSheet).Cells(i, rateCol).Value
        cost = rate * hrs
        
        'Navigates through columns
        For j = copyCol To copyCol + termLength - 1
            hc = Sheets(exportSheet).Cells(pasteRow, pasteCol).Value
            ans = hc * cost
            Sheets(exportSheet).Cells(pasteRow, pasteCol).Value = ans
            pasteCol = pasteCol + 1
        Next j
        pasteRow = pasteRow + 1
        pasteCol = 21
    Next i
End Sub