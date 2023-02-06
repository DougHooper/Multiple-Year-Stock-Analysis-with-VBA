Attribute VB_Name = "Module1"
Sub VBA_challenge()

    'create loops through all the stocks for one year and outputs the following informaiton
    ' ticker symbol
    'yearly change from opening price at the beginning of a givenyear to the closing price at the end of that year
    'Percentage change from the opening price at the beginning og a given year ro the closing price at the en dof that year
    'The total stock volume
    
    
    
    'object to find the last row of data
    lastRow = Cells(Rows.count, 1).End(xlUp).row
    
    ' declare a variable to hold the count (accumulator)
    Dim counter As Integer
    
    'check on the ticker symbol
    Dim tickersymbol As String
    
    'variable to hold the totals stock volume
    Dim TotalVolume As Double
    TotalVolume = 0 'starts the total at 0
    
    'variable to hold the rows in the total columns
    Dim stockrow As Integer
    stockrow = 2
    
    'variable to hold open value
    Dim start As Double
    
    
    'variable to hold close value
    Dim ending As Double
     
    'variable for yearly change
    Dim yearlychange As Double
    
    'variable to hold percent change
    Dim percent As Double
    percent = 0
    
    'declare variable to hold the row
      Dim row As Long
    
    'loop through the rows and check the changes in the credit cards
    For row = 2 To lastRow

    
        'check the changes in the credit cards
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
            
            'if the ticker symbol changes
            
            ' set the name ticker column name
            tickersymbol = Cells(row, 1).Value 'grabs the value from Ticker column before the change
            
            'add to the stock total
            TotalVolume = TotalVolume + Cells(row, 7).Value 'grabs the value from volume column before the change
            
            'display the ticker sybmol on the total columns
            Cells(stockrow, 9).Value = tickersymbol
            
            'display the total stock column on the total column
            Cells(stockrow, 12).Value = TotalVolume
            
            'set the open stock amount
            start = Cells(row, 3).Value
            
            'set the close stock amount
            ending = Cells(row, 6).Value
            
            
            ' add 1 to the stock rows to go to the next row
            stockrow = stockrow + 1
            
            'reset the credit card total for the next credit card
            TotalVolume = 0
            
            
            
            'calcualte yearly change
            
            
            
            'calculate percent change
            
        Else
        
            'if there is no change in the ticker symbols, keep adding to the total
            TotalVolume = TotalVolume + Cells(row, 7).Value   'Grabs the value from column G
            
            
        End If
    
    
    Next row
    
    End Sub
    
