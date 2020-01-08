Sub pullDayAndSort()
    Dim i As Integer
    
    i = 1
    
    'sort entire spreadsheet based on date column, to remove blanks at end
    Range("A:A").Sort Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlNo
    
    'might need to add header on day column here
    
    'extract day from entire date, put in column B
    Do While (Range("A" & i).Value <> "")
        divided = Split(Range("A" & i).Value, "/")
        Range("B" & i).Value = divided(1)
        
        'if we need Mo/Day in a column, uncomment below
        'Range("C" & i).Value = divided(0) + "/" + divided(1)
        
        i = i + 1
    Loop
    
    'sort both columns based on column B
    Range("A:C").Sort Key1:=Range("B:B"), Order1:=xlAscending, Header:=xlNo
    
    'delete unnecessary days column B
    Columns(2).EntireColumn.Delete
    
End Sub
