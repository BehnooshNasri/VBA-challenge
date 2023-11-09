Sub VBA_Assignment():
 
For Each ws In Worksheets
SheetName = ws.Name
MsgBox SheetName


    Dim ticker_count As Double
    Dim vol_count As Double
    Dim earliest_date As Double
    Dim latest_date As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim current_date As Double
  
   
    earliest_date = 20231109
    latest_date = 20131109
    vol_count = 0
    ticker_count = 2
    yearly_change = 0
    percent_change = 0
    last_row = ws.Range("A1").End(xlDown).Row
    summary_last_row = ws.Range("K1").End(xlDown).Row
    
         
    'insert the new columns

   
    ws.Range("K1").EntireColumn.Insert
    ws.Range("K1").EntireColumn.Insert
    ws.Range("K1").EntireColumn.Insert
    ws.Range("K1").EntireColumn.Insert

   
    'create the new columns headers

   
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Voulme"
    ws.Range("M1:M" & summary_last_row).NumberFormat = "0.00%"
    

   
    'getting the closing and opening price and the percentage

 

    For i = 2 To last_row
        
        'Value2 function I got from ChatGPT and it corrected the bug
        
        
       current_date = ws.Cells(i, 2).Value2
     
     
        If current_date <= earliest_date Then
            earliest_date = current_date
            open_price = ws.Cells(i, 3).Value
           

        End If

   
        If current_date > latest_date Then
            latest_date = current_date
            close_price = ws.Cells(i, 6).Value
         
        End If
      
    'create the loop to count the same ticker, volume, yearly change and percentage
   
          
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            vol_count = vol_count + ws.Cells(i, 7).Value

       
        Else

            vol_count = vol_count + ws.Cells(i, 7).Value
            yearly_change = (close_price - open_price)
            percent_change = (yearly_change / open_price)

           
            ws.Cells(ticker_count, 11).Value = ws.Cells(i, 1).Value
            ws.Cells(ticker_count, 14).Value = vol_count
            ws.Cells(ticker_count, 12).Value = yearly_change
            ws.Cells(ticker_count, 13).Value = percent_change
         
            vol_count = 0
            yearly_change = 0
            percent_change = 0
            ticker_count = ticker_count + 1
            earliest_date = 20231109
            latest_date = 20131109
   
            
        End If

       Next i
   
    For i = 2 To summary_last_row
  
        If ws.Cells(i, 12).Value < 0 Then
          
            ws.Cells(i, 12).Interior.ColorIndex = 3
       
        ElseIf ws.Cells(i, 12).Value > 0 Then
        
            ws.Cells(i, 12).Interior.ColorIndex = 4

        End If
     

    Next i
    
    
   'insert the greatest increase, decrease and the greatest volume
   
   
    ws.Range("R1").EntireColumn.Insert
    ws.Range("R1").EntireColumn.Insert
    
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"
    ws.Range("Q2").Value = "Greatest % increase"
    ws.Range("Q3").Value = "Greatest % decrease"
    ws.Range("Q4").Value = "Greatest total volume"
    
    'I got these functions below with a lot of help from AskBCS
    
    
    ws.Range("S2") = "%" & Application.WorksheetFunction.Max(ws.Range("M2:M" & last_row)) * 100
    value_increase_ticker = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("M2:M" & last_row)), ws.Range("M2:M" & last_row), 0)
    ws.Range("R2").Value = ws.Cells(value_increase_ticker + 1, 11)
    
    ws.Range("S3") = "%" & Application.WorksheetFunction.Min(ws.Range("M2:M" & last_row)) * 100
    value_decrease_ticker = WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("M2:M" & last_row)), ws.Range("M2:M" & last_row), 0)
    ws.Range("R3").Value = ws.Cells(value_decrease_ticker + 1, 11)
    
    ws.Range("S4") = Application.WorksheetFunction.Max(ws.Range("N2:N" & last_row))
    value_increase_total = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("N2:N" & last_row)), ws.Range("N2:N" & last_row), 0)
    ws.Range("R4").Value = ws.Cells(value_increase_total + 1, 11)
    

Next ws
    
End Sub

