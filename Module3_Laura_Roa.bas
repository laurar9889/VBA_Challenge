Attribute VB_Name = "Module3"
Sub Part3_the_greatest()

 'Loop through all sheets
Dim ws As Worksheet
For Each ws In Sheets



    'Putting titles in second summary
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Vol"

    ' Assigning variables
    Dim percent_change As Double
    Dim ticker_total As LongPtr
   
    'Result1 for greatest % increase
        Dim result_incr As Double
      
      'Result2 for greatest % decrease
    Dim result_decr As Double
     
    ' Result3 is for total Vol
   Dim result_vol As LongPtr
    

    'Netting them to zero
    percent_change = 0
    ticker_total = 0
    result1 = 0
    result2 = 0
    result3 = 0

    'Determine the last row
    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
      
    'setting the column rage that needs to get the information of percentage from.
        ' function found in excel mojo website
            Set Rng1 = ws.Range("k2:k" & lastrow)
      
     'setting the column rage that needs to get the information of volume from.
        Set Rng2 = ws.Range("L2:L" & lastrow)

          'to calculate max %
    maxvalue1 = ws.Application.WorksheetFunction.Max(Rng1)
    
              
    ' to calculate min %
    minvalue2 = ws.Application.WorksheetFunction.Min(Rng1)
    
        ' to calculate max vol
    maxvalue3 = ws.Application.WorksheetFunction.Max(Rng2)

        
   'place then in the right cell at the right format
    result_incr = maxvalue1
    result_decr = minvalue2
    result_vol = maxvalue3

    ws.Range("Q2").Value = result_incr
    ws.Range("Q3").Value = result_decr
    ws.Range("Q4").Value = result_vol

  'put the  variance in percentage of min and max value in summary
    ws.Range("Q2").Value = Format(result_incr, "#.##%")
    ws.Range("Q3").Value = Format(result_decr, "#.##%")
    
    'find ticker for each item
For i = 2 To lastrow


If ws.Cells(i, 11).Value = maxvalue1 Then ws.Range("p2").Value = ws.Cells(i, 9).Value

If ws.Cells(i, 11).Value = minvalue2 Then ws.Range("p3").Value = ws.Cells(i, 9).Value

If ws.Cells(i, 12).Value = maxvalue3 Then ws.Range("p4").Value = ws.Cells(i, 9).Value

Next i

Next ws

End Sub



