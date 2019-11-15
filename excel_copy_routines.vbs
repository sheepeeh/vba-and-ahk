Sub copyContent()

    Dim Nrow As Long, Nsheet As Long
    Dim i As Long
    Dim LastRow As Long
    

    Worksheets.Add(before:=Worksheets(1)).Name = "All Data"

    Nrow = 1    'row to copy
    Nsheet = Worksheets.Count  'the count AFTER adding the destination worksheet
    
   
    
    For i = 4 To Nsheet
        LastRow = Worksheets(1).UsedRange.Rows(Worksheets(1).UsedRange.Rows.Count).Row
        Worksheets(i).UsedRange.Rows.Copy Destination:=Worksheets(1).Cells(LastRow + 1, "A")
    Next i

End Sub


Sub copyRow()

    Dim Nrow As Long, Nsheet As Long
    Dim i As Long
    Dim rowContent

    Worksheets.Add(before:=Worksheets(1)).Name = "All Rows"

    Nrow = 1    'row to copy
    Nsheet = Worksheets.Count  'the count AFTER adding the destination worksheet

    For i = 2 To Nsheet
        Worksheets(i).Cells(Nrow, "A").EntireRow.Copy Destination:=Worksheets(1).Cells(i - 1, "A")
    Next i

End Sub


Sub countContent()

    Dim Nrow As Long, Nsheet As Long
    Dim i As Long
    Dim LastRow As Long
    Dim RowCount As Long
    Dim sheetName As String
    

   
    Nsheet = Worksheets.Count  'the count AFTER adding the destination worksheet
    
   
    
    For i = 3 To Nsheet
        sheetName = Worksheets(i).Name
        RowCount = Worksheets(i).UsedRange.Rows.Count
        LastRow = Worksheets(1).UsedRange.Rows(Worksheets(1).UsedRange.Rows.Count).Row
        Worksheets(1).Range("B" & i - 1).Value = RowCount
        Worksheets(1).Range("A" & i - 1).Value = sheetName
        
    Next i

End Sub



