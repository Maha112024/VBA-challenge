Sub year_part2()
    ' Create variables
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim lastrowquarter As Long
    Dim combinedsheet As Worksheet

    ' Add a new sheet and name it "Combined_data"
    Set combinedsheet = Sheets.Add
    combinedsheet.Name = "Combined_data"
    
    ' Move the combined sheet to the first position
    combinedsheet.Move Before:=Sheets(1)

    ' Loop through all sheets
    For Each ws In Worksheets
        ' Exclude the combined sheet from processing
        If ws.Name <> "Combined_data" Then
            ' Find the last row of the combined sheet after each paste
            lastrow = combinedsheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
            
            ' Find the last row of each worksheet (subtract 1 to ignore header)
            lastrowquarter = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
            
            ' Check if there are rows to copy
            If lastrowquarter > 1 Then
                ' Copy the contents of each quarter sheet into the combined sheet
                combinedsheet.Range("A" & lastrow & ":H" & (lastrow + lastrowquarter - 2)).Value = ws.Range("A2:H" & (lastrowquarter + 1)).Value
            End If
        End If
    Next ws

    ' Copy the headers from the first worksheet (excluding the combined sheet)
    If Worksheets.Count > 1 Then
        combinedsheet.Range("A1:H1").Value = Worksheets(2).Range("A1:H1").Value
    End If

    ' Autofit to display data
    combinedsheet.Columns("A:H").AutoFit

End Sub