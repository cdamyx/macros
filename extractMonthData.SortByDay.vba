Sub promptForMonth(x)
    
    x = InputBox("Enter Month as 1 - 12", "Month")

End Sub

Sub findColumn(ws, col, str)

    col = Application.WorksheetFunction.Match(str, ws.Range("1:1"), 0)

End Sub

Sub convertColToLetter(col, ltrCol)

    ltrCol = Split(Cells(1, col).Address, "$")(1)

End Sub

Sub filterColByMo(ws, col, mo)

    ws.Cells(col).EntireColumn.AutoFilter Field:=1, Criteria1:=mo, Operator:=11

End Sub

Sub createSheet(wksht, sheetName)

    Sheets.Add After:=wksht
    ActiveSheet.Name = sheetName

End Sub

Sub copyData(copyFrom, pasteTo)
    
    copyFrom.UsedRange.Copy
    pasteTo.Paste

End Sub

Sub extractDay(x, ltrCol)
    
    Do While (Range(ltrCol & x).Value <> "")
        divided = Split(Range(ltrCol & x).Value, "/")
        Range("Z" & x).Value = divided(1)
        x = x + 1
    Loop
    x = 2

End Sub

Sub sortWS(wksht)

    wksht.Range("A:Z").Sort Key1:=Range("Z:Z"), Order1:=xlAscending, Header:=xlYes
    Columns("Z").EntireColumn.Delete

End Sub

Sub formatting(wksht, col, col2)

    wksht.Range(col + ":" + col2).WrapText = False
    wksht.Columns(col + ":" + col2).AutoFit

End Sub

Sub createCopy(wkshtOrig, wksht)
    
    Set wkshtOrig = Worksheets("Birthday")
    createSheet wkshtOrig, "Copy"
    Worksheets("Copy").Activate
    Set wksht = Worksheets("Copy")
    copyData wkshtOrig, wksht
    formatting wksht, "A", "Z"

End Sub

Sub replaceRehire(wksht, j, ltrColR, ltrColH)
    Do While (wksht.Range("A" & j).Value <> "")
            If (wksht.Range(ltrColR & j).Value <> "") Then
                'put current value of hire date into a comment
                wksht.Range(ltrColH & j).AddComment ("Original Hire Date: " + wksht.Range(ltrColH & j).Text)
                'replace hire date with rehire date
                wksht.Range(ltrColH & j).Value = wksht.Range(ltrColR & j).Value
            End If
            j = j + 1
    Loop
End Sub
Sub filterCopySort(wksht, wksht2, col, arr, usrIn, count, ltrCol)

    'turn off filter first to remove filter feature currently enabled on any column
    wksht.AutoFilterMode = False
    'make new worksheet active for extract/sort later
    wksht2.Activate
    'filter main ws by dates occuring in selected month
    filterColByMo wksht, col, arr(usrIn)
    'Copy the filtered data over to new sheet
    copyData wksht, wksht2
    'take day out of date and place in col Z
    extractDay count, ltrCol
    'sort on col Z, then delete
    sortWS wksht2
    'formatting each row to be short(wraptext), and each column to be wide enough to display all text(autofit)
    formatting wksht2, "A", "Z"
    'clear any filtered rows on main worksheet
    wksht.ShowAllData

End Sub

Sub main()
    Dim i As Integer
    Dim j As Integer
    Dim wsOriginal As Worksheet
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim userInput As Integer
    Dim colBirthDate As Integer
    Dim colHireDate As Integer
    Dim colRehireDate As Integer
    Dim colLetterBirth As String
    Dim colLetterHire As String
    Dim colLetterRehire As String
    Dim months(1 To 12) As Integer
    
    'array to change month to filter criteria value used by VBA
    months(1) = 21
    months(2) = 22
    months(3) = 23
    months(4) = 24
    months(5) = 25
    months(6) = 26
    months(7) = 27
    months(8) = 28
    months(9) = 29
    months(10) = 30
    months(11) = 31
    months(12) = 32
    
    'values for iteration later
    i = 2
    j = 2
    
    'create and activate "Copy", copy data over to "Copy" to run calcs/edits
    createCopy wsOriginal, ws
    
    'find column of each type of date
    findColumn ws, colBirthDate, "Birth Date"
    findColumn ws, colHireDate, "Hire Date"
    findColumn ws, colRehireDate, "Rehire Date"
    
    'convert column number to column letter
    convertColToLetter colBirthDate, colLetterBirth
    convertColToLetter colHireDate, colLetterHire
    convertColToLetter colRehireDate, colLetterRehire
    
    'copy any rehire dates over to hire date column and add comment with original hire date
    replaceRehire ws, j, colLetterRehire, colLetterHire
    
    'create and assign other 2 worksheets
    createSheet ws, "Hire Date"
    createSheet ws, "Birth Date"
    Set ws2 = Worksheets("Birth Date")
    Set ws3 = Worksheets("Hire Date")
    
    'get month from user
    promptForMonth userInput
    
    'birth date sheet specific
    filterCopySort ws, ws2, colBirthDate, months, userInput, i, colLetterBirth
    
    'hire date sheet specific
    filterCopySort ws, ws3, colHireDate, months, userInput, i, colLetterHire
    
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    wsOriginal.Activate

End Sub
