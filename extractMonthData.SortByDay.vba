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

    'Cells(col).EntireColumn: how to select a whole column by number, i.e., col = 10, so select entire column J
    ws.Cells(col).EntireColumn.AutoFilter Field:=1, Criteria1:=mo, Operator:=11
    
    'Excel has built in constants that can be used as filter criteria
    ' Value Constant
    ' 1     xlFilterToday
    ' 2     xlFilterYesterday
    ' 3     xlFilterTomorrow
    ' 4     xlFilterThisWeek
    ' 5     xlFilterLastWeek
    ' 6     xlFilterNextWeek
    ' 7     xlFilterThisMonth
    ' 8     xlFilterLastMonth
    ' 9     xlFilterNextMonth
    ' 10    xlFilterThisQuarter
    ' 11    xlFilterLastQuarter
    ' 12    xlFilterNextQuarter
    ' 13    xlFilterThisYear
    ' 14    xlFilterLastYear
    ' 15    xlFilterNextYear
    ' 16    xlFilterYearToDate
    ' 17    xlFilterAllDatesInPeriodQuarter1
    ' 18    xlFilterAllDatesInPeriodQuarter2
    ' 19    xlFilterAllDatesInPeriodQuarter3
    ' 20    xlFilterAllDatesInPeriodQuarter4
    ' 21    xlFilterAllDatesInPeriodJanuary
    ' 22    xlFilterAllDatesInPeriodFebruary
    ' 23    xlFilterAllDatesInPeriodMarch
    ' 24    xlFilterAllDatesInPeriodApril
    ' 25    xlFilterAllDatesInPeriodMay
    ' 26    xlFilterAllDatesInPeriodJune
    ' 27    xlFilterAllDatesInPeriodJuly
    ' 28    xlFilterAllDatesInPeriodAugust
    ' 29    xlFilterAllDatesInPeriodSeptember
    ' 30    xlFilterAllDatesInPeriodOctober
    ' 31    xlFilterAllDatesInPeriodNovember
    ' 32    xlFilterAllDatesInPeriodDecember
    
    'also, Operator:=11 represents xlFilterDynamic. Just keep it as 11. https://docs.microsoft.com/en-us/office/vba/api/excel.xlautofilteroperator

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
    wksht2.Range("A:Z").WrapText = False
    wksht2.Columns("A:Z").AutoFit
    
    'make main worksheet visible
    'wksht.Activate
    
    'clear any filtered rows on main worksheet
    wksht.ShowAllData

End Sub

Sub main()
    Dim i As Integer
    Dim j As Integer
    Dim wsOriginal As Worksheet
    Dim ws As Worksheet
    Dim wsCopy As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
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
    
    'switch to this array and use loops, when finished
    'colNames(0) = "Birth Date"
    'colNames(1) = "Hire Date"
    'colNames(2) = "Rehire Date"
    
    'set i for extract day later
    i = 2
    j = 2
    
    'create and activate "Copy", copy data over to "Copy" to run calcs/edits
    Set wsOriginal = Worksheets("Birthday")
    createSheet wsOriginal, "Copy"
    Worksheets("Copy").Activate
    Set ws = Worksheets("Copy")
    copyData wsOriginal, ws
    ws.Range("A:Z").WrapText = False
    ws.Columns("A:Z").AutoFit
    
    'find column of each type of date
    findColumn ws, colBirthDate, "Birth Date"
    findColumn ws, colHireDate, "Hire Date"
    findColumn ws, colRehireDate, "Rehire Date"
    
    'convert column number to column letter
    convertColToLetter colBirthDate, colLetterBirth
    convertColToLetter colHireDate, colLetterHire
    convertColToLetter colRehireDate, colLetterRehire
    
    'before we really do much of anything, go ahead and copy any rehire dates over to hire date column, highlight, and add comment with original hire date
    Do While (ws.Range("A" & j).Value <> "")
        If (ws.Range(colLetterRehire & j).Value <> "") Then
            'put current value of hire date into a comment
            ws.Range(colLetterHire & j).AddComment ("Original Hire Date: " + ws.Range(colLetterHire & j).Text)
            
            'replace hire date with rehire date and highlight yellow
            ws.Range(colLetterHire & j).Value = ws.Range(colLetterRehire & j).Value
            ws.Range(colLetterHire & j).Interior.ColorIndex = 6
            
        End If
        j = j + 1
    Loop
    
    
    
    'createSheet ws, "Rehire Date"
    createSheet ws, "Hire Date"
    createSheet ws, "Birth Date"
    
    Set ws2 = Worksheets("Birth Date")
    Set ws3 = Worksheets("Hire Date")
    'Set ws4 = Worksheets("Rehire Date")
    
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
