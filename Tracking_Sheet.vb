Sub AddEntry()
'File: AddEntry Macro 
'Author: Erin Payne 
'Description: For automating hyperlinking to template tabs.

    Dim inpYear As Integer
    inp = InputBox("In what year would you like to add your entry?")
    inpYear = CInt(inp)
    inpName = InputBox("What is the title of your entry?")
    
    Dim i As Integer
    i = 2
    Do While CInt(Cells(2, i).Value) <> inpYear
        i = i + 1
    Loop
    i = i + 1
    Cells(1, i).EntireColumn.Insert
    Cells(3, i).Value = inpName
    
    documentName = InputBox("What is the name of the document for the entry you would like to add?" & vbNewLine & "Please type exactly as titled.")
    Dim tabName As String
    
    For j = 4 To 21
        tabName = Cells(j, 1).Value
        Dim inpUrl As String
        inpUrl = "//ls.nttdata.com/ds/fin-FinancialAnalysis/CommercialPricing/Team%20Documents/Contractual_Terms_Database/" & inpYear & "/NMP.xlsx#'" & tabName & "'!A1"
        
        If WorksheetFunction.CountA(Range("A3:K10")) = 0 Then
            .Hyperlinks.Add Anchor:=Cells(j, i), _
                Address:=inpUrl, _
                TextToDisplay:="X"
            'Cells(j, i).Hyperlink(inpUrl).TextToDisplay = "X"
        Else
            .Hyperlinks.Add Anchor:=Cells(j, i), _
                Address:=inpUrl, _
                TextToDisplay:="N/A"
            'Cells(j, i).Hyperlink(inpUrl).TextToDisplay = "N/A"
        End If
    Next j
End Sub
