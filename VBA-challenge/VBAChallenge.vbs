Sub VBA_Homework():
   
        Sheets.Add.Name = "Master_Sheet"
        Sheets("Master_Sheet").Move Before:=Sheets(A)
        set as new_sheet = Worksheets("Master_Sheet")
        
    
  
        For Each ws In Worksheets
                lastRow= new_sheet.Cells(Rows.Count, "A").End(xlUP).Row + 1
                lastRowTicker= ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
'I know this doesn't run because of this line, but I couldn't figure out the range 
                new_sheet.Range("A" & lastRow & ":A" & ((lastRowTicker -1)+lastRow)).Value = ws.Range("A2:A" & (lastRowTicker + 1)).Value

        Next ws

        For Each ws in Worksheets
                lastRow= new_sheet.Cells(Rows.Count, "A").End(xlUP).Row + 1
                lastRowTicker= ws.Cells(Rows.Count, "A").End(xlUp).Row - 1  
                close=(lastRowTicker - 1, "F")
                open=(lastRowTicker - 1, "C")
                yearly=(close-open).Value
End Sub
