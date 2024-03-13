Attribute VB_Name = "Module2"
Sub Part2_yearly_change()

'Loop through all sheets
Dim ws As Worksheet
For Each ws In Sheets


'create variable to hold worksheet name which is the year
Dim wsname As String
'Get sheet name
wsname = ws.Name


'assign variables:
Dim open_val As Double
Dim close_val As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim start_date As String
Dim end_date As String

start_date = "0102"
end_date = "1231"



open_val = 0
close_val = 0

'Determine the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     'where to put the total count for each ticker
    Dim Run_second_summary_table_row As Double
    
    summary_table_row = 2
    
 For i = 2 To lastrow
    
     
    'Look for opening value
    
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 2).Value = wsname & start_date Then

open_val = ws.Cells(i, 3).Value

   'put the open value of that specific ticker in the summary
    'Range("J" & summary_table_row).Value = open_val
    

ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 2).Value = wsname & end_date Then
close_val = ws.Cells(i, 6).Value

   'put the open value of that specific ticker in the summary
    'Range("k" & summary_table_row).Value = close_val
    
    'calculate variance from opening to close
    yearly_change = close_val - open_val
    
    'calculate variance % from opening to close
    percent_change = yearly_change / open_val
                  
    
       'put the yearly change value of that specific ticker in the summary
    ws.Range("J" & summary_table_row).Value = yearly_change
    
     'put the  variance in % of that specific ticker in the summary
    ws.Range("K" & summary_table_row).Value = Format(percent_change, "#.##%")
    
    
    'move to the next row available
       summary_table_row = summary_table_row + 1

End If

Next i

 'If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 2).Value = "20200102" Then
ws.Cells(i, 3).Value = open_val
 
 'Determine the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     For i = 2 To lastrow
     
 'Determine if the value should be green for greater tahn 0 or red
 If ws.Cells(i, 10).Value > 0 Then
 ws.Cells(i, 10).Interior.ColorIndex = 4
 Else
 ws.Cells(i, 10).Interior.ColorIndex = 3
 
 End If
 
  If ws.Cells(i, 11).Value > 0 Then
 ws.Cells(i, 11).Interior.ColorIndex = 4
 Else
 ws.Cells(i, 11).Interior.ColorIndex = 3
 
 End If
 
Next i

Next ws

End Sub





