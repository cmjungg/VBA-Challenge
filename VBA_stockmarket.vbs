Attribute VB_Name = "Module1"
'## Instructions

'Create a script that loops through all the stocks for one year and outputs the following information:

'  * The ticker symbol.

'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'       yearly change divided by initial stock

'  * The total stock volume of the stock.

' **Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'## Bonus

'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

'## VBA Code

Sub Stock_market()

    'Loop through all sheets
    Dim ws As Worksheet
    
    Dim lastrow As Double
    Dim lastrow_results As Integer
    Dim ticket_symbol As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_v As LongLong
    Dim intial_stock As Double
    Dim end_stock As Double
    Dim next_row As Double
    Dim GPI As Double
    Dim GPD As Double
    Dim GTV As Double
          
    lastws = Sheets.Count
    For k = 1 To lastws
    Sheets(k).Activate
    
        

    
    
    
        'Insert headers
        
                Range("I1").Value = "Ticker"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Total Stock Volume"
                Cells(2, 14).Value = "Greatest % Increase"
                Cells(3, 14).Value = "Greatest % Decrease"
                Cells(4, 14).Value = "Greatest Total Volume"
                Cells(1, 15).Value = "Ticker"
                Cells(1, 16).Value = "Value"
                
                'Initialize Total Stock Volume
                
                total_stock_v = 0
                
                'For each Stock
                
                lastrow = Cells(Rows.Count, 1).End(xlUp).Row
                nextrow = 2
                initial_stock = 2
    
        For i = 2 To lastrow
        
            'Check if still in the same ticker
        
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            'Add the ticker
                  
            Cells(nextrow, 9).Value = Cells(initial_stock, 1).Value
            
            'note the end of the stock
            'error
            
            end_stock = i
        
            'Add Total Stock Volume
            
            total_stock_v = total_stock_v + Cells(i, 7).Value
            Cells(nextrow, 12).Value = total_stock_v
            
            'Add Yearly Change and Percent Change
            
            yearly_change = Cells(end_stock, 6).Value - Cells(initial_stock, 3).Value
            Cells(nextrow, 10) = yearly_change
            Cells(nextrow, 10).NumberFormat = "0.00"
            
            percent_change = yearly_change / Cells(initial_stock, 3).Value
            Cells(nextrow, 11).Value = percent_change
            Cells(nextrow, 11).Style = "Percent"
            Cells(nextrow, 11).NumberFormat = "0.00%"
            
            'move to the next row for results
            
            nextrow = nextrow + 1
            
            'Reset Total Stock Volume
            
            total_stock_v = 0
            
            'Update intial stock
            initial_stock = i + 1
            
            Else
                
                'Add Total Stock Volume
                total_stock_v = total_stock_v + Cells(i, 7).Value
                
                
            End If
            
            
         Next i
       
     '##BONUS
       
                lastrow_results = Cells(Rows.Count, 10).End(xlUp).Row
                GPI = 0
                GPD = 0
                GTV = 0
    
        For j = 2 To lastrow_results
        
           'Add color for visualization
            If Cells(j, 10) < 0 Then
            
            Cells(j, 10).Interior.ColorIndex = 3
            
            ElseIf Cells(j, 10) > 0 Then
            
            Cells(j, 10).Interior.ColorIndex = 4
            
            End If
            
            'Find MAX, MIN and Total Volume MAX
            
            If Cells(j, 11) > GPI Then
            GPI = Cells(j, 11).Value
            rowinc = j
            
            End If
            
            If Cells(j, 11) < GPD Then
            GPD = Cells(j, 11).Value
            rowdec = j
            
            End If
            
            If Cells(j, 12) > GTV Then
            GTV = Cells(j, 12).Value
            rowTV = j
            
            End If
            
            Next j
        
        'Print Values
        
             Cells(2, 16).Value = GPI
              Cells(2, 15).Value = Cells(rowinc, 9).Value
                      Cells(2, 16).Style = "Percent"
                        Cells(2, 16).NumberFormat = "0.00%"
              Cells(3, 16).Value = GPD
              Cells(3, 15).Value = Cells(rowdec, 9).Value
                      Cells(3, 16).Style = "Percent"
                        Cells(3, 16).NumberFormat = "0.00%"
              Cells(4, 16).Value = GTV
              Cells(4, 15).Value = Cells(rowTV, 9).Value
    
            'Column alingment
            Columns("I:P").EntireColumn.AutoFit
    
    
    Next k


End Sub

