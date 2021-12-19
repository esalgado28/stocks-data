Attribute VB_Name = "Module1"
Sub stocks_all_years()
'loops through all worksheets, running the code in each

Dim ws As Worksheet

For Each ws In Worksheets
    ws.Select
    Call stocks
    Call summary_table
Next

End Sub

'runs through the daily data and summarizes each stock
Sub stocks()

'set the headings for the table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Columns(10).AutoFit                     'auto column width to look nice
Cells(1, 11).Value = "Percent Change"
Columns(11).AutoFit
Cells(1, 12).Value = "Total Stock Volume"
Columns(12).AutoFit

'initialize variables we need for table
Dim ticker As String
ticker = Cells(2, 1).Value

Dim initial_price As Double
initial_price = Cells(2, 3).Value

Dim final_price As Double
final_price = Cells(2, 6)

Dim total_volume As Double
total_volume = Cells(2, 7)

'this is the starting table row
Dim table_row As Integer
table_row = 2

'find the last row so we know where to stop the for loop
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through every row in the raw data
For i = 3 To lastrow + 1

    'if ticker doesn't match, we are at new stock or past the end
    'calculate desired values and add row to table
    If (Cells(i, 1) <> ticker) Then
    
        Cells(table_row, 9).Value = ticker
        Cells(table_row, 10).Value = final_price - initial_price
        
        'format color to indicate wether price increased or decreased
        If (Cells(table_row, 10).Value > 0) Then
            Cells(table_row, 10).Interior.ColorIndex = 4    'green
            
        ElseIf (Cells(table_row, 10).Value < 0) Then
            Cells(table_row, 10).Interior.ColorIndex = 3    'red
            
        Else                                                'if =0
            Cells(table_row, 10).Interior.ColorIndex = 16   'gray
        
        End If
        
        If (initial_price = 0) Then             'added to avoid /0 error
                Cells(table_row, 11).Value = 0  'while not 0 mathematically, may avoid strange behavior
            
        Else
            'calculate percent change
            Cells(table_row, 11).Value = (final_price - initial_price) / initial_price
            
        End If
        
        Cells(table_row, 12).Value = total_volume
        
        'update/reset variables for next stock
        ticker = Cells(i, 1).Value
        initial_price = Cells(i, 3).Value
        final_price = Cells(i, 6).Value
        total_volume = Cells(i, 7).Value
        
        'move to next row in table
        table_row = table_row + 1
    
    'if ticker is the same, update final price and add to total volume
    Else
        final_price = Cells(i, 6).Value
        total_volume = total_volume + Cells(i, 7).Value
        
    End If

Next i

'set number format to percentages for %change column
Columns(11).NumberFormat = "0.00%"

End Sub

'this creates the summary table based off result of stocks()
'finds stocks that saw largest increase in value, largest decrease in value, and the stock traded the most
Sub summary_table()

'initialize needed variables
Dim greatest_increase, greatest_decrease, greatest_volume As Double
Dim lucky_stock, poor_stock, popular_stock As String
greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

'find last row in table created from stocks()
lastrow = Cells(Rows.Count, 9).End(xlUp).Row

'loop through every row in table
For i = 2 To lastrow
    
    'update if current volume is larger than greatest volume found so far
    If (Cells(i, 12).Value > greatest_volume) Then
        greatest_volume = Cells(i, 12).Value
        popular_stock = Cells(i, 9).Value
    
    End If
        
    'update if current % increase is larger than greatest increase found so far
    If (Cells(i, 11).Value > greatest_increase) Then
        greatest_increase = Cells(i, 11).Value
        lucky_stock = Cells(i, 9).Value
        
    'update if current % decrease is larger (more negative) than largest found so far
    ElseIf (Cells(i, 11).Value < greatest_decrease) Then
        greatest_decrease = Cells(i, 11).Value
        poor_stock = Cells(i, 9).Value
        
    End If

Next i
    
'create table with values found
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Cells(2, 15).Value = "Greatest % Increase"
Cells(2, 16).Value = lucky_stock
Cells(2, 17).Value = greatest_increase

Cells(3, 15).Value = "Greatest % Decrease"
Cells(3, 16).Value = poor_stock
Cells(3, 17).Value = greatest_decrease

Cells(4, 15).Value = "Greatest Total Volume"
Cells(4, 16).Value = popular_stock
Cells(4, 17).Value = greatest_volume

'some finishing touches
Columns(15).AutoFit
Range("Q2:Q3").NumberFormat = "0.00%"
Columns(17).AutoFit

End Sub
