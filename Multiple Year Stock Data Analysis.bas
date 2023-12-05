Attribute VB_Name = "Module1"
Sub stockAnalysis():

    Dim total As LongLong 'total stock volume
    Dim row As Long 'loop contorl variable goes through Rows
    Dim rowCount As Long 'variable hold number of rows in a sheet
    Dim yearlyChange As Double 'variable that holds the yearly change for each stock in a sheet
    Dim percentChange As Double 'variabel that holds the percent change for each stock in a sheet
    Dim summaryTableRow As Long 'variable holds the rows of the summary table row
    Dim stockStartRow As Long 'variable that holds the start of a stock's row in sheet
    Dim startValue As Long ' start value for a stock

    'loop through all of the worksheets
    For Each ws In Worksheets
    
        'set the title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
    
        'initialize the values
        summaryTableRow = 0
        total = 0 'total stock volume for a stock starts at 0
        yearlyChange = 0 'yearly change starts at 0
        stockStartRow = 2 'first stock in the sheet is going to start at 2
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        For row = 2 To rowCount
        
            'check to see if there are changes in column A
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                'calculate the total one last time
                total = total + ws.Cells(row, 7).Value
                
                'check to see if the value of the total volume is 0
                If total = 0 Then
                    'print the results in columns I,J,K,L
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                    
                Else
                    'find the first non-zero starting value
                    If ws.Cells(stockStartRow, 3).Value = 0 Then
                        For findValue = stockStartRow To row
                    
                        'check to see if the next (or next) value does not equal to 0
                            If ws.Cells(findValue, 3).Value <> 0 Then
                            stockStartRow = findValue
                            'once we ahve a non-zero value, break out of the loop
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    'Calculate the yearly change (difference in last close - first open)
                    yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(stockStartRow, 3).Value)
                    
                    'Calculate the percent change
                    percentChange = yearlyChange / ws.Cells(stockStartRow, 3).Value
                
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                    ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00" ' formats column
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
                    
                    If yearlyChange > 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 'green
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 'red
                    Else
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    
                    End If
                    
                        
                        
                End If
                
                'reset the value of the total
                total = 0
                'reset the value of yearly change
                yearlyChange = 0
                'move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
                
            ' if the ticker is the same
            Else
                total = total + ws.Cells(row, 7).Value 'Grabs value from the 7th column
                
            End If
        
        
        Next row
        
        ' after looping through the rows, find the max and min and place them in Q2, Q3, and Q4
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
        ws.Range("Q4").NumberFormat = "#,###"
        
        'do matching in order to match the ticker names with the values
        'tell the row in the summary table where the ticker matches the greatest increase
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        
         'tell the row in the summary table where the ticker matches the greatest decrease
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
        
        volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        ws.Range("P4").Value = ws.Cells(volNumber + 1, 9)
        
        'AutoFit the columns
        
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub
