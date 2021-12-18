Attribute VB_Name = "Module1"
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
        
        If (initial_price = 0) Then         'added to avoid /0 error
            Cells(table_row, 11).Value = 0  'while not mathematically 0, may avoid strange behavior later
            
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
    
    'if ticker is the same, update final price and add volume to total volume
    Else
        final_price = Cells(i, 6).Value
        total_volume = total_volume + Cells(i, 7).Value
        
    End If

Next i

'set the number format to percentages for % change column
Columns(11).NumberFormat = "0.00%"

End Sub
