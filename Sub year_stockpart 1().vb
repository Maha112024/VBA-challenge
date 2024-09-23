Sub year_stock()

    Dim i As Integer
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim worksheetname As String

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        ' Get the name of the current worksheet
        worksheetname = ws.Name
        
        ' Find the last row in column A
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Insert a new column at the beginning (A)
        ws.Columns(1).Insert Shift:=xlToRight
        
        ' Set header for the new column
        ws.Cells(1, 1).Value = "Quarter"
        
        ' Fill the new column with the worksheet name from row 2 to the last row
        ws.Range("A2:A" & lastrow).Value = worksheetname
    Next ws

End Sub