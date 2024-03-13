Attribute VB_Name = "Module1"
Sub Part1_stocks_totalVol()

'assigning variables:
    
    
    'Loop through all sheets
Dim ws As Worksheet
For Each ws In Sheets


    ' name of ticker
    Dim tickername As String
    
    ' where to count total volume of each ticker
    Dim ticker_total As LongPtr
    

    
    'Set the count of total to zero
        ticker_total = 0

    'where to put the total count for each ticker
    Dim summary_table_row As Long
    
    'In which row to put each ticker
    summary_table_row = 2
    
    
'Look for the individual ticker name and sum total vol
'Determine the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

 For i = 2 To lastrow
 
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'get the name
    tickername = ws.Cells(i, 1).Value
    
    'label summary_table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percent change"
    ws.Range("L1").Value = "Total Stock Vol"
    
        
    
    'put the total in the internal counter
    ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    'put the total vol of that specific row in the summary
    ws.Range("L" & summary_table_row).Value = ticker_total
     
     'put the ticker name in the summary
    ws.Range("I" & summary_table_row).Value = tickername
    
    'go to the next cell to include a new ticker name
    summary_table_row = summary_table_row + 1
    
    'reset the total for the next ticker
    ticker_total = 0
    
    'if the cell inmediately following a row is the same name
    Else
    
    'keep adding the total for each ticker
    ticker_total = ticker_total + Cells(i, 7).Value
       

        
    End If
    
    Next i
    
    Next ws

End Sub


'-------------------



