Attribute VB_Name = "Module1"
'Reference- Worked with my tutor, study group, Slack Overflow, "The SpreadsheetGuru" and lots of Google searches to fill in where I didn't know how to write the code.
'Why didn't we just use pivot tables for this, would have been so much faster

'Create a script that will loop through all the stocks for one year and output the following information.
'Assignment
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

'CHALLENGES
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.


Sub Formatdata():
'loop through all tabs
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
'Set which worksheet is first
    Set starting_ws = ActiveSheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate


'Variables
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Percent_Change As Double
    Dim Volume As Double
    Dim Lastrow As Long
    Dim Row As Double
    Dim column As Integer
    
'Format Data for grid
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly_Change"
Range("k1").Value = "Percent_Change"
Range("l1").Value = "Total_Stock_Value"
Rows("1:1").Font.Bold = True

'Start Volume at 0
    Volume = 0
'Set output to second row, to start in second row and not on headers
    Row = 2
    column = 1
    
 'Variable i- like an abreviation
    Dim i As Long
 
 'What is my opening Price?  (2 rows down, and Column 1+2- starts us at C2)
    Open_Price = ws.Cells(2, column + 2).Value
 
 'Tell is where the last row is- use xlUp trick
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 'Need to loop through ticker
     For i = 2 To Lastrow
 
 ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
 'Give Ticker a name for reference
    Ticker_Name = ws.Cells(i, column).Value
    ws.Cells(Row, column + 8).Value = Ticker_Name
    
 'Set Closing Price
    Close_Price = ws.Cells(i, column + 5).Value
    
 'Need to know Yearly change, but no column of data for it
    Yearly_Change = Close_Price - Open_Price
 'Where is yearly change going? In the table summary
    ws.Cells(Row, column + 9).Value = Yearly_Change
    
 'Need to know the % change (use If, ElseIf and Else here) End if
    If (Open_Price = 0 And Close_Price = 0) Then
    Percent_Change = 0
    ElseIf (Open_Price = 0 And Close_Price <> 0) Then
        Percent_Change = 1
     Else
        Percent_Change = Yearly_Change / Open_Price
        ws.Cells(Row, column + 10).Value = Percent_Change
        ws.Cells(Row, column + 10).NumberFormat = "0.00%"
        
    End If
    
  'Add in the Total Volume
    Volume = Volume + ws.Cells(i, column + 6).Value
    ws.Cells(Row, column + 11).Value = Volume
    
  'Add onto the Summary Table
    Row = Row + 1
    
  'Need to reset the Open Price to start (C2)
    Open_Price = ws.Cells(i + 1, column + 2)
    
  'Need to reset the Volume to start at 0
    Volume = 0
    
  'Need to know if the cells have the same ticker (A and G)
    Else
    Volume = Volume + ws.Cells(i, column + 6).Value
    
    'end statement
    End If
    
    Next i
    
   'Determine the Last Row of the Yearly Change
     YCLastRow = ws.Cells(Rows.Count, column + 8).End(xlUp).Row
     
    'Make the Cell Colors (see Saturday class, how we did this)
     For j = 2 To YCLastRow
        If (ws.Cells(j, column + 9).Value > 0 Or ws.Cells(j, column + 9).Value = 0) Then
            ws.Cells(j, column + 9).Interior.ColorIndex = 10
        ElseIf ws.Cells(j, column + 9).Value < 0 Then
            ws.Cells(j, column + 9).Interior.ColorIndex = 3
        End If
        
'Challenge
    Next j
    
'Set up Greatest% Increase, Decrease and Total Volume (where are they going to go- make a table)
    ws.Cells(2, column + 14).Value = "Greatest % Increase"
    ws.Cells(3, column + 14).Value = "Greatest % Decrease"
    ws.Cells(4, column + 14).Value = "Greatest Total Volume"
    ws.Cells(1, column + 15).Value = "Ticker"
    ws.Cells(1, column + 16).Value = "Value"
    
'Look at each row and find greatest value and its Ticker (Q is just like i and j- another abreviation)(use if and elseIf statements here)
    For Q = 2 To YCLastRow
    If ws.Cells(Q, column + 10).Value = Application.WorksheetFunction.Max(Range("K2:K" & YCLastRow)) Then
        ws.Cells(2, column + 15).Value = ws.Cells(Q, column + 8).Value
        ws.Cells(2, column + 16).Value = ws.Cells(Q, column + 10).Value
        ws.Cells(2, column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Q, column + 10).Value = Application.WorksheetFunction.Min(Range("K2:K" & YCLastRow)) Then
        ws.Cells(3, column + 15).Value = ws.Cells(Q, column + 8).Value
        ws.Cells(3, column + 16).Value = ws.Cells(Q, column + 10).Value
        ws.Cells(3, column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Q, column + 11).Value = Application.WorksheetFunction.Max(Range("L2:L" & YCLastRow)) Then
        ws.Cells(4, column + 15).Value = ws.Cells(Q, column + 8).Value
        ws.Cells(4, column + 16).Value = ws.Cells(Q, column + 11).Value
    End If
    
    Next Q
    
'Set Cell A1 to each sheet to ticker
    ws.Cells(1, 1) = "<ticker>"
    Next
'Tell it which was the first sheet
    starting_ws.Activate

End Sub

