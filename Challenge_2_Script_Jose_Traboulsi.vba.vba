Sub Ticker_Exercise_AllSheets()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim i As Double
    Dim Summary_Table_Row As Double
    Dim Price_Start As Double
    Dim LastRow As Double
    Dim Stock_Volume As Double
    Dim MaxValue As Double
    Dim Max_Ticker As String
    Dim MinValue As Double
    Dim Min_Ticker As String
    Dim MaxVolume As Double
    Dim MaxVolume_Ticker As String
    Dim LastRow_Summary As Double

    For Each ws In Worksheets
    
        Summary_Table_Row = 2
        Price_Start = 2
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Stock_Volume = 0

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("I1:L1").Columns.AutoFit

       'Find Tickers

        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i, 1).Value

                ' Find price change
                ws.Cells(Summary_Table_Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(Price_Start, 3).Value
                ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"

                ' Find % Change
                ws.Cells(Summary_Table_Row, 11).Value = (ws.Cells(Summary_Table_Row, 10).Value / ws.Cells(Price_Start, 3).Value)
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"

                ' Conditional formatting
                If ws.Cells(Summary_Table_Row, 10).Value > 0 And ws.Cells(Summary_Table_Row, 11).Value > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                    
                ElseIf ws.Cells(Summary_Table_Row, 10).Value < 0 And ws.Cells(Summary_Table_Row, 11).Value < 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                End If

                ' Find Total stock volume and place it in the summary table, reset the counter
               
                For j = Price_Start To i
                
                    Stock_Volume = Stock_Volume + ws.Cells(j, 7).Value
               
                Next j

                ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume

                Stock_Volume = 0


                ' Adjust Summary Table Row and Price Start Row for next Ticker values
                
                Summary_Table_Row = Summary_Table_Row + 1
                Price_Start = i + 1
            
        End If
        Next i

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        LastRow_Summary = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

        ' Find greatest % increase, decrease and respective tickers and place these in the new summary table
        MaxValue = ws.Range("K2").Value
        For p = 3 To LastRow_Summary
            If ws.Cells(p, 11).Value > MaxValue Then
                MaxValue = ws.Cells(p, 11).Value
                Max_Ticker = ws.Cells(p, 9).Value
            End If
        Next p
        ws.Range("P2").Value = Max_Ticker
        ws.Range("Q2").Value = MaxValue
        ws.Range("Q2").NumberFormat = "0.00%"

        MinValue = ws.Range("K2").Value
        For q = 3 To LastRow_Summary
            If ws.Cells(q, 11).Value < MinValue Then
                MinValue = ws.Cells(q, 11).Value
                Min_Ticker = ws.Cells(q, 9).Value
            End If
        Next q
        ws.Range("P3").Value = Min_Ticker
        ws.Range("Q3").Value = MinValue
        ws.Range("Q3").NumberFormat = "0.00%"

        MaxVolume = ws.Range("L2").Value
        For Z = 3 To LastRow_Summary
            If ws.Cells(Z, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(Z, 12).Value
                MaxVolume_Ticker = ws.Cells(Z, 9).Value
            End If
        Next Z
        ws.Range("P4").Value = MaxVolume_Ticker
        ws.Range("Q4").Value = MaxVolume

        ws.Range("O:Q").Columns.AutoFit
     
    Next ws
    
End Sub
