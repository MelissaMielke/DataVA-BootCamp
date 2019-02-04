Sub Stock_mcm()

    'Set variable for worksheets to cycle through
    Dim ws As Worksheet

    'Start looping through all worksheets
    For Each ws In Worksheets
    
         'Define column headers for the summary data table
         ws.Cells(1, 9).Value = "Ticker Symbol"
         ws.Cells(1, 10).Value = "Yearly Change"
         ws.Cells(1, 11).Value = "Percent Change"
         ws.Cells(1, 12).Value = "Total Stock Volume"
    
    	 'Set variable for ticker symbol
         Dim ticker As String
    
   		 'Set variables to determine yearly change
   		 Dim y_open As Double
    	 Dim y_close As Double
    	 Dim y_change As Double
    
    	 'Set variable for percent change
    	 Dim y_percent As Double
    
    	 'Set for total stock volume
    	 Dim volume As Double
   
    	 'Initiate default values for variables (if not found)
    	 y_open = 0
    	 y_close = 0
    	 y_change = 0
         y_percent = 0
   		 volume = 0
    
   		 'Set variable to track row count in summary table; starting row 2 after headers
   		 Dim r_count As Long
    	 r_count = 2
    
   		 'Set variable for total rows to loop through
   		 Dim lastrow As Long
   		 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
   		 'Loop through all ticker symbols to calculate summary table
   		 For r = 2 To lastrow
    
        	'Conditional IF to check ticker symbol is still remaining constant
       	 	 If ws.Cells(r, 1).Value <> ws.Cells(r - 1, 1).Value Then
        		 'Pull in year open price
        		 y_open = ws.Cells(r, 3).Value
         	 End If
        
        	 'Calculate total stock volume; move to summary table
       		 volume = volume + ws.Cells(r, 7).Value
        
        	 'Conditional IF to check ticker symbol is still remaining constant
       	 	 If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
            	'Move ticker symbol to summary table
            	 ws.Cells(r_count, 9).Value = ws.Cells(r, 1).Value		 	 
			 End If
                
             'Move total stock volume to summary table
             ws.Cells(r_count, 12).Value = volume

             'Pull in year end price
             y_close = ws.Cells(r, 6).Value
                
             'Calculate price change for year; move to summary table
             y_change = y_close - y_open
             ws.Cells(r_count, 10).Value = y_change
            
             'Calculate percent change for year; move to summary table; format as a percent

             'Conditional IF year close is 0, would error for '0' denominator
             If y_close = 0 Then
                y_percent = 0
                ws.Cells(r_count, 11).Value = y_percent
                ws.Cells(r_count, 11).NumberFormat = "0.0%"
                
             'Conditional ELSEIF year open is 0, would be infinite percent; declare new variable
             ElseIf y_open = 0 Then
                 Dim y_percent_new As String
                 y_percent_new = "New"
                 ws.Cells(r_count, 11).Value = y_percent
                
             'Conditional ELSE able to calculate a percent change
             Else
                 y_percent = y_change / y_open
                 ws.Cells(r_count, 11).Value = y_percent
                 ws.Cells(r_count, 11).NumberFormat = "0.0%"
             End If
            
             'Conditional IF the change over year was positive format green
             If y_change >= 0 Then
                ws.Cells(r_count, 10).Interior.ColorIndex = 4
             'Conditional IF the change over year was negative format red
             Else
                 ws.Cells(r_count, 10).Interior.ColorIndex = 3
             End If
        
             'Move on to next empty row in summary table for next ticker symbol
             r_count = r_count + 1
        
             'Reset all variables before next loop
             volume = 0
             y_open = 0
             y_close = 0
             y_change = 0
             y_percent = 0
        
    	 Next r
    
   		 'Define column headers for the best/worst summary data table
   		 ws.Cells(1, 16).Value = "Ticker Symbol"
  		 ws.Cells(1, 17).Value = "Greatest Value"
    	 ws.Cells(2, 15).Value = "Greatest % Increase"
   		 ws.Cells(3, 15).Value = "Greatest % Decrease"
    	 ws.Cells(4, 15).Value = "Greatest Total Volume"

    	 'Assign lastrow to count the number of ticker symbol rows in the summary table
    	 lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

   		 'Set variables for ticker & value for greatest % increase, % decrease, total volume
    	 Dim tick_ginc As String
   		 Dim val_ginc As Double
    	 Dim tick_gdec As String
    	 Dim val_gdec As Double
    	 Dim tick_gvol As String
    	 Dim val_gvol As Double

         'Initiate values for greatest increase, decrease, and total volume to first fow value
         val_ginc = ws.Cells(2, 11).Value
         val_gdec = ws.Cells(2, 11).Value
         val_gvol = ws.Cells(2, 12).Value

         'Loop to search through summary table
         For i = 2 To lastrow

             'Conditional IF to determine ticker stock with greatest increase
             If ws.Cells(i, 11).Value > val_ginc Then
                val_ginc = ws.Cells(i, 11).Value
                tick_ginc = ws.Cells(i, 9).Value
             End If

             'Conditional IF to determine ticker stock with greatest decrease
             If ws.Cells(i, 11).Value < val_gdec Then
                val_gdec = ws.Cells(i, 11).Value
                tick_gdec = ws.Cells(i, 9).Value
             End If

             'Conditional to determine stock with the greatest total volume
             If ws.Cells(i, 12).Value > val_gvol Then
                 val_gvol = ws.Cells(i, 12).Value
                 tick_gvol = ws.Cells(i, 9).Value
             End If

         Next i

         'Move values for greatest increase, decrease, and total volume to performance summary table
         ws.Cells(2, 16).Value = tick_ginc
         ws.Cells(2, 17).Value = val_ginc
         ws.Cells(2, 17).NumberFormat = "0.0%"
         ws.Cells(3, 16).Value = tick_gdec
         ws.Cells(3, 17).Value = val_gdec
         ws.Cells(3, 17).NumberFormat = "0.0%"
         ws.Cells(4, 16).Value = tick_gvol
         ws.Cells(4, 17).Value = val_gvol

         'Fit summary tables to column width of contents
         ws.Columns("I:L").EntireColumn.AutoFit
         ws.Columns("O:Q").EntireColumn.AutoFit

     Next ws

End Sub
