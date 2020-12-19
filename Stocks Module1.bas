Attribute VB_Name = "Module1"
Sub Stocks()
    
    'define variables
        Dim LastRow As Long
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As LongLong
        Dim StockVolume As LongLong
        Dim SummaryTableRow As Integer
        Dim LastRow_SummaryTable As Integer
        Dim OpenPrice As Double
        Dim ClosePrice As Double
    
    'for loop to iterate through each sheet in the workbook
    For Each ws In Worksheets
           
                       
        'add headers to the summary table
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Stock Volume"
        
        'autofit the summary table to fit the headers and values
        ws.Columns("I:L").AutoFit
                       
        'find the last row in each sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                     
        'establish first row of summary table
        SummaryTableRow = 2
        
        'establish the first open price
        OpenPrice = ws.Cells(2, 3).Value
        
        'iterate through each row in the given worksheet
        For Row = 2 To LastRow
                
                'find the ticker changes and put the ticker value into the summary table
                If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
                    Ticker = ws.Cells(Row, 1).Value
                    ws.Range("I" & SummaryTableRow) = Ticker
                    
                    'add the last like ticker's stock volume to total ticker stock volume
                    StockVolume = StockVolume + Cells(Row, 7).Value
                    
                    'put total stock volume into the summary table
                    ws.Range("L" & SummaryTableRow).Value = StockVolume
                    
                    'collect information about close price
                    ClosePrice = ws.Cells(Row, 6).Value
                                       
                    'to calculate the yearly change in stock price
                    YearlyChange = ClosePrice - OpenPrice
                    
                    'put yearly change in the summary table
                    ws.Range("J" & SummaryTableRow) = YearlyChange
                    
                    'to calculate the yearly % change using an if statement to avoid division by zero
                    If OpenPrice > 0 Then
                        PercentageChange = (ClosePrice - OpenPrice) / OpenPrice
                    Else
                        PercentageChange = 0
                    End If
                    
                    'print yearly % change in summary table
                    ws.Range("K" & SummaryTableRow).Value = PercentageChange
                    
                    'format % change as %
                    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                    
                    'move down 1 row on the summary table
                    SummaryTableRow = SummaryTableRow + 1
                    
                    'set StockVolume to 0 when the ticker changes
                    StockVolume = 0
                    
                    'reset the open price when the ticker changes
                    OpenPrice = ws.Cells(Row + 1, 3)
                              
                'add incremental ticker stock volume to total ticker stock volume when ticker not changing
                Else
                    StockVolume = StockVolume + ws.Cells(Row, 7).Value
                End If
        
        Next Row
            
    
    
    'to perform conditional formatting in the yearly change column that will turn negative values red and positive values green
    'find last row of summary table
        LastRow_SummaryTable = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    
    'reestablish SummaryTableRow as 2
        SummaryTableRow = 2
        
        
    'format stock volume with hundreds separated by commas
        ws.Range("L2:L" & LastRow_SummaryTable).NumberFormat = "0,000"
        
    'loop through each yearly change value in the summary table and format positive green and negative red
        For i = SummaryTableRow To LastRow_SummaryTable
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    
        'Bonus: create an additional summary table that shows greatest % increase, greatest % decrease, and greatest total volume
        
        'establish the row labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'establish the column headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'find the max % increase, min % increase, greatest volume in the summary table
        Dim MaxPercentage As Double
        Dim MinPercentage As Double
        Dim GreatestVolume As Double
        
        MaxPercentage = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow_SummaryTable))
        MinPercentage = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow_SummaryTable))
        GreatestVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow_SummaryTable))
        
        
        
        'loop through each row in the % change column and print the max % and corresponding ticker in the second summary table and format accordingly
        For j = SummaryTableRow To LastRow_SummaryTable
            If ws.Cells(j, 11).Value = MaxPercentage Then
                ws.Range("Q2").Value = ws.Cells(j, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = ws.Cells(j, 9).Value
            End If
            
            If ws.Cells(j, 11).Value = MinPercentage Then
                ws.Range("Q3").Value = ws.Cells(j, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Cells(j, 9).Value
            End If
                    
            If ws.Cells(j, 12).Value = GreatestVolume Then
                ws.Range("Q4").Value = ws.Cells(j, 12).Value
                ws.Range("Q4").NumberFormat = "0,000"
                ws.Range("P4").Value = ws.Cells(j, 9).Value
            End If
        
        Next j
            
        
        'autofit the second summary table
        ws.Columns("O:Q").AutoFit
    
    Next ws


End Sub


