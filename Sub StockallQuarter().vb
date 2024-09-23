Sub StockallQuarter()

    ' Create variables
    Dim i As Long
    Dim lastrow As Long
    Dim ticker As String
    Dim total_stock As Double
    Dim Summary_Table_Row As Long
    Dim Q As Worksheet
    Dim open_qrt As Double
    Dim close_qrt As Double
    Dim qrt_change As Double
    Dim Price_change As Double
    Dim maxvalue As Double
    Dim minvalue As Double
    Dim maxticker As String
    Dim minticker As String
    Dim maxvol As Double
    Dim maxvolticker As String
    Dim sheetNames As Variant
    Dim summaryRowOffset As Long

    ' List of worksheets to loop through
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    ' Loop through each sheet
    For Each sheetName In sheetNames
        Set Q = Worksheets(sheetName)
        total_stock = 0
        Summary_Table_Row = 2 ' Start summary table from row 2

        lastrow = Q.Cells(Rows.Count, "A").End(xlUp).Row ' Find last row in the current sheet

        ' Loop through rows in the column
        For i = 2 To lastrow
            ' Check if we are at the last row or if the next ticker is different
            If i = lastrow Or Q.Cells(i + 1, 2).Value <> Q.Cells(i, 2).Value Then
                ticker = Q.Cells(i, 2).Value ' Current ticker
                close_qrt = Q.Cells(i, 7).Value ' Get the closing price (Column G)

                ' Set the opening price for the first occurrence of the ticker
                open_qrt = Q.Cells(i - (Application.CountIf(Q.Range("B2:B" & i), ticker) - 1), 4).Value ' Column D
                
                ' Print the Ticker in the Summary Table
                Q.Range("K" & Summary_Table_Row).Value = ticker

                ' Sum up stocks for this ticker
                total_stock = total_stock + Q.Cells(i, 8).Value ' Assuming stock value is in column H (8)

                ' Print the total stock amount in the summary table
                Q.Range("N" & Summary_Table_Row).Value = total_stock
                
                ' Calculate the quarterly change
                qrt_change = close_qrt - open_qrt ' Closing - Opening
                
                If (qrt_change <= 0) Then
                    Q.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                Else
                    Q.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                End If
                
                ' Calculate price change, check for division by zero
                If open_qrt <> 0 Then
                    Price_change = (close_qrt - open_qrt) / open_qrt
                Else
                    Price_change = 0 ' Handle potential division by zero
                End If
                
                  If (Price_change <= 0) Then
                    Q.Range("M" & Summary_Table_Row).Interior.ColorIndex = 9 ' Red
                Else
                    Q.Range("M" & Summary_Table_Row).Interior.ColorIndex = 10 ' Green
                End If
                
                ' Output values in the summary table
                Q.Range("O" & Summary_Table_Row).Value = open_qrt
                Q.Range("P" & Summary_Table_Row).Value = close_qrt
                Q.Range("L" & Summary_Table_Row).Value = qrt_change
                
                ' Format Price_change as a percentage
                Q.Range("M" & Summary_Table_Row).Value = Price_change
                Q.Range("M" & Summary_Table_Row).NumberFormat = "0.00%" ' Formats the cell as a percentage

                ' Move to the next row in the summary table
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset total_stock for the next ticker
                total_stock = 0
                
            Else
                ' Add to the total stock for the current ticker
                total_stock = total_stock + Q.Cells(i, 8).Value
            End If
        Next i

        ' Calculate max and min values from the summary table
        maxvalue = Application.WorksheetFunction.Max(Q.Range("M2:M" & Summary_Table_Row - 1)) ' Quarterly change
        minvalue = Application.WorksheetFunction.Min(Q.Range("M2:M" & Summary_Table_Row - 1))
        maxvol = Application.WorksheetFunction.Max(Q.Range("N2:N" & Summary_Table_Row - 1)) ' Max volume

        ' Find ticker associated with max value
        For i = 2 To Summary_Table_Row - 1
            If Q.Range("M" & i).Value = maxvalue Then
                maxticker = Q.Range("K" & i).Value
                Exit For
            End If
        Next i

        ' Find ticker associated with min value
        For i = 2 To Summary_Table_Row - 1
            If Q.Range("M" & i).Value = minvalue Then
                minticker = Q.Range("K" & i).Value
                Exit For
            End If
        Next i

        ' Find ticker associated with max volume
        For i = 2 To Summary_Table_Row - 1
            If Q.Range("N" & i).Value = maxvol Then
                maxvolticker = Q.Range("K" & i).Value
                Exit For
            End If
        Next i

        ' Output the max and min values and their associated tickers in summary
        Q.Range("T2").Value = maxvalue
        Q.Range("T2").NumberFormat = "0.00%" ' Formats the cell as a percentage
        Q.Range("S2").Value = maxticker
        
        Q.Range("T3").Value = minvalue
        Q.Range("T3").NumberFormat = "0.00%" ' Formats the cell as a percentage
        Q.Range("S3").Value = minticker

        Q.Range("T4").Value = maxvol
        Q.Range("S4").Value = maxvolticker ' Output max volume ticker
    Next sheetName

End Sub