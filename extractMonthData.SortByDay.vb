Sub promptForMonth(x)
    
    x = InputBox("Enter Month as 1 - 12", "Month")

End Sub

Sub findColumn(ws, col, str)

    col = Application.WorksheetFunction.Match(str, ws.Range("1:1"), 0)

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
    ' 22    xlFilterAllDatesInPeriodFebruray <-February is misspelled
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

Sub createSheet(ws, sheetName)

    Sheets.Add After:=ws
    
    ActiveSheet.Name = sheetName


End Sub

Sub copyData(copyFrom, pasteTo)
    
    copyFrom.UsedRange.Copy
    pasteTo.Paste

End Sub

Sub main()
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    Dim userInput As Integer
    Dim colBirthDate As Integer
    Dim colHireDate As Integer
    Dim colRehireDate As Integer
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
    
    
    'might need to change this. could do Worksheets(1), or activate/activesheet?
    Set ws = Worksheets("Birthday")
    
    'createSheet ws, "Rehire Date"
    createSheet ws, "Hire Date"
    createSheet ws, "Birth Date"
    
    Set ws2 = Worksheets("Birth Date")
    Set ws3 = Worksheets("Hire Date")
    'Set ws4 = Worksheets("Rehire Date")
    
    'get month from user
    promptForMonth userInput

    'find column of each type of date
    findColumn ws, colBirthDate, "Birth Date"
    findColumn ws, colHireDate, "Hire Date"
    findColumn ws, colRehireDate, "Rehire Date"
    
    filterColByMo ws, colBirthDate, months(userInput)
    
    copyData ws, ws2
    
    'MsgBox ("Birth " + CStr(colBirthDate) + ", " + "Hire " + CStr(colHireDate) + ", " + "Rehire " + CStr(colRehireDate))
    
    'Note for later: explicitly select sheet, then manipulate with ActiveSheet: do Worksheets(1).Activate to make sheet 1 _
    the selected sheet, then use ActiveSheet with various commands to manipulate

End Sub
