Sub Test_Macro()
' File name: Test_Macro Macro
' Author: Erin Payne
' Description: Final test file for additional data export.
    
    Dim rowCount As Integer
    rowCount = 0
    
    Dim n As Integer
    n = 9
    
    While Sheets("FTE Input").Cells(n, 1).Value <> "DO NOT DELETE THIS ROW!!!"
        rowCount = rowCount + 1
        MsgBox ("Rows = " & rowCount)
        n = n + 1
    Wend
    
    'TT FTE data copy
    Sheets("FTE Input").Range("A8:EH19").Copy         'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A1").FormulaR1C1 = "Category"              'Renames "Labor" heading to "Category"
    Rows(1).Style = "Input Heading"                   'Adds heading format to first row
    Range("A2").FormulaR1C1 = "TT FTE"                'Renames "Labor" to "TT FTE"
    Range("A2").AutoFill Destination:=Range("A2:A12"), Type:=xlFillDefault                         'TO DO: Change fixed range
    
    'TT Base Labor Cost data copy
    Sheets("FTE Input").Range("A9:EH19").Copy         'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A13").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A13").FormulaR1C1 = "TT Base Labor Cost"   'Renames "Labor" to "TT Base Labor Cost"
    Range("A13").AutoFill Destination:=Range("A13:A23"), Type:=xlFillDefault                       'TO DO: Change fixed range
    
    'TT Cost COLA data copy
    Sheets("FTE Input").Range("A9:EH19").Copy         'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A24").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A24").FormulaR1C1 = "TT Cost COLA"         'Renames "Labor" to "TT Cost COLA"
    Range("A24").AutoFill Destination:=Range("A24:A34"), Type:=xlFillDefault                        'TO DO: Change fixed range
    
    'TT Cost Contingency data copy
    Sheets("FTE Input").Range("A9:EH19").Copy         'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A35").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A35").FormulaR1C1 = "TT Cost Contingency"  'Renames "Labor" to "TT Cost Contingency"
    Range("A35").AutoFill Destination:=Range("A35:A45"), Type:=xlFillDefault                        'TO DO: Change fixed range
    
    'SS FTE data copy
    Sheets("FTE Input").Range("A30:EH39").Copy        'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A46").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A46").FormulaR1C1 = "SS FTE"               'Renames "Labor" to "SS FTE"
    Range("A46").AutoFill Destination:=Range("A46:A55"), Type:=xlFillDefault                        'TO DO: Change fixed range
    
    'SS Base Labor Cost data copy
    Sheets("FTE Input").Range("A30:EH39").Copy        'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A56").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A56").FormulaR1C1 = "SS Base Labor Cost"   'Renames "Labor" to "SS Base Labor Cost"
    Range("A56").AutoFill Destination:=Range("A56:A65"), Type:=xlFillDefault                        'TO DO: Change fixed range
    
    'SS Cost COLA data copy
    Sheets("FTE Input").Range("A30:EH39").Copy        'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A66").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A66").FormulaR1C1 = "SS Cost COLA"         'Renames "Labor" to "SS Cost COLA"
    Range("A66").AutoFill Destination:=Range("A66:A75"), Type:=xlFillDefault                        'TO DO: Change fixed range
    
    'SS Cost Contingency data copy
    Sheets("FTE Input").Range("A30:EH39").Copy         'initializes macro at "FTE Input" sheet       TO DO: Change fixed range
    'Copies and pastes TT FTE information to blank sheet
    Sheets("Blank Sheet 2").Range("A76").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("A76").FormulaR1C1 = "SS Cost Contingency"   'Renames "Labor" to "SS Cost Contingency"
    Range("A76").AutoFill Destination:=Range("A76:A85"), Type:=xlFillDefault                        'TO DO: Change fixed range
End Sub
