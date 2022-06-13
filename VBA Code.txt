Sub StockMarketAnalysis():

    ' Loop / Iterate Through All Worksheets
    For Each ws In Worksheets

            ' Defining the Column Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        
            ' Formatting the Column/Row Headers
            ws.Range("I1:L1").Font.Bold = True
            ws.Range("I1:L1").Font.ColorIndex = 2
            ws.Range("I1:L1").Interior.ColorIndex = 16
            
            ws.Range("P1:Q1").Font.Bold = True
            ws.Range("P1:Q1").Font.ColorIndex = 2
            ws.Range("P1:Q1").Interior.ColorIndex = 16
            
            ws.Range("O2:O4").Font.Italic = True
            ws.Range("O2:O4").Font.ColorIndex = 1
            ws.Range("O2:O4").Interior.ColorIndex = 15
            
            ws.Range("A1:G1").Font.Bold = True
            ws.Range("A1:G1").Font.ColorIndex = 1
            ws.Range("A1:G1").Interior.ColorIndex = 36
            
            ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
            ' Format Table Columns To Auto Fit
            ws.Columns("I:Q").AutoFit
            
            

        ' Declaring the Variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalVolume As Double
        TotalVolume = 0
        
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        
        Dim PreviousValue As Long
        PreviousAmount = 2
        
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        ' Find the Last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

    ' Finding the total volume for each Ticker range
    TotalTickerVolume = TotalTickerVolume + ws.Cells(i, "G").Value
            
            ' Check If We Are Still Within The Same Ticker Name If It Is Not...
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then


                ' Set Ticker Name
                TickerName = ws.Cells(i, "A").Value
                ' Print The Ticker Name In The Summary Table
                ws.Range("I" & SummaryTableRow).Value = TickerName
                
                ' Print The Ticker Total Volume Amount To The Summary Table
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                ' Reset Ticker Total
                TotalTickerVolume = 0


        ' Set Yearly Open, Yearly Close and Yearly Change Name
        
                'Calculates the sum of <open> for that specific TICKER and then adds it to the previous yearly open amount
                YearlyOpen = ws.Range("C" & PreviousAmount)
                
                'Calculates the Yearly Close for the Ticker by looping through 'i'
                YearlyClose = ws.Range("F" & i)
                
                'Calculation for Yearly Change
                YearlyChange = YearlyClose - YearlyOpen
                
                'Prints the Yearly Change in the Summary table
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                ' Determine Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    
                    'Calculates the sum of <open> for that specific TICKER and then adds it to the previous value
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    
                    'Calculates the Percentage Change for the summary table
                    PercentChange = YearlyChange / YearlyOpen
                    
                    End If
                ' Format Double To Include % Symbol And Two Decimal Places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                    
                
                ' Conditional Formatting Highlight if value is equal or greater than 0 result will be Positive = Green
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            
                Else
                
                '' Conditional Formatting Highlight if value is less than 0 result will be Negative = Red
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
                End If
            
                ' Add One To The Summary Table Row - moves to the next ticker range.
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            
            Next i
            
            
    'Bonus challenge
    
            ' Calculation for Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
            ' Start Loop For Final Results
            For i = 2 To LastRow
        
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        
    Next ws

End Sub

