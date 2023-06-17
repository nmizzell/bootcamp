Attribute VB_Name = "Module1"
Sub stocks_on_all_sheets()

For Each ws In ThisWorkbook.Sheets
    
    ws.Activate
    
    Call stocks
    
    Call greatest
    
Next

End Sub


Sub stocks()
'define variables
Dim output_index As Integer
Dim ticker As String
Dim yearly_change As Double
    Dim yearly_beginning_price As Double
    Dim yearly_ending_price As Double
Dim percent_change As Double
Dim total_stock_volume As Double


'create headers for output data
Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'set output index to 2. This is the first row that the unique stock ticker data will be stored
output_index = 2

'set first start variables
yearly_beginning_price = Cells(2, 3)


For i = 2 To Range("A:A").End(xlDown).Row Step 1

    total_stock_volume = total_stock_volume + Cells(i, 7)
    
    'each time the next item in the list changes, then we have to do the following:
    If Cells(i + 1, 1) <> Cells(i, 1) Or IsEmpty(Cells(i + 1, 1)) Then
        'output total stock volume
        Cells(output_index, 12) = total_stock_volume
        
        'format total stock volume
        Cells(output_index, 12).NumberFormat = "#,###"
        
        'reset total stock volume to zero
        total_stock_volume = 0
        
        'get the yearly ending price
        yearly_ending_price = Cells(i, 3)
        
        'compute the change from the beginning price
        yearly_change = yearly_ending_price - yearly_beginning_price
        
        'output the yearly change
        Cells(output_index, 10) = yearly_change
        
        'format the yearly change as red or green depending on pos or neg
        If yearly_change >= 0 Then
            Cells(output_index, 10).Interior.Color = RGB(0, 255, 0)
        Else:
            Cells(output_index, 10).Interior.Color = RGB(255, 0, 0)
        End If
             
        'compute the percent change and output it
        Cells(output_index, 11) = (yearly_ending_price - yearly_beginning_price) / yearly_beginning_price
        
        'format the percent change as a percentage
        Cells(output_index, 11).NumberFormat = "0.00%"
        
        'get the next yearly beginning price
        yearly_beginning_price = Cells(i + 1, 3)
        

        
        'ouput the ticker
        Cells(output_index, 9) = Cells(i, 1)
    
        'lastly, add one to the output_index
        output_index = output_index + 1
    
    Else: Cells(1, 1) = Cells(1, 1)
    
    End If
    
Next i


End Sub

Sub greatest()

Dim greatest_increase As Double
Dim greatest_increase_ticker As String
Dim greatest_decrease As Double
Dim greatest_decrease_ticker As String
Dim greatest_volume As Double
Dim greatest_volume_ticker As String

'set default greates to the first value
greatest_increase = Cells(2, 11)
greatest_increase_ticker = Cells(2, 9)

greatest_decrease = Cells(2, 11)
greatest_decrease_ticker = Cells(2, 9)

greatest_volume = Cells(2, 12)
greatest_volume_ticker = Cells(2, 9)

For i = 2 To Range("I:I").End(xlDown).Row Step 1
    
    'check if this stock is the greatest increase
    If Cells(i, 11) > greatest_increase Then
        greatest_increase = Cells(i, 11)
        greatest_increase_ticker = Cells(i, 9)
    End If
    
    'check if the stock is the greatest decrease
    If Cells(i, 11) < greatest_decrease Then
        greatest_decrease = Cells(i, 11)
        greatest_decrease_ticker = Cells(i, 9)
    End If
    
    'check if the stock is the greatest volume
    If Cells(i, 12) > greatest_volume Then
        greatest_volume = Cells(i, 12)
        greatest_volume_ticker = Cells(i, 9)
    End If
    
Next i

'output greatest
'create headers
Range("o1:q1") = Array("", "Ticker", "Value")
'Range("o2:o4") = Array("Greatest PCT Increase", "Greatest PCT Decrease", "Greatest Total Volume")

Range("o2") = "Greatest PCT Increase"
Range("o3") = "Greatest PCT Decrease"
Range("o4") = "Greatest Volume"


'populate values
Range("p2:q2") = Array(greatest_increase_ticker, greatest_increase)
Range("p3:q3") = Array(greatest_decrease_ticker, greatest_decrease)
Range("p4:q4") = Array(greatest_volume_ticker, greatest_volume)

'format values
Range("q2:q3").NumberFormat = "0.00%"
Range("q4").NumberFormat = "#,###"


End Sub
