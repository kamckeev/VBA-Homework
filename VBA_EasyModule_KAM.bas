Attribute VB_Name = "Stockmarket_VBA"
Sub total_volume()
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
'Your result should look as follows (note: all solution images are for 2015 data).

'Defining the variables and the type they are:

Dim i As Double
Dim ticker_name As String
Dim ticker_value As Double
Dim LastRow As Double
Dim summaryrow As Double
Dim ticker_totalvalue As Double

'Defining measures that you will use in the formulas

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
summaryrow = 2

'Initial variables
ticker_name = Cells(2, 1)
ticker_toalvalue = 0
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Stock Volume"

'Writing a loop
For i = 2 To LastRow
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
        ticker_value = Cells(i, 7).Value
        ticker_totalvalue = ticker_totalvalue + ticker_value
       
        
    Else
    'Adding the voulme of the previous cells to the next cell that meets the criteria
    
        ticker_name = Cells(i, 1).Value
        ticker_totalvalue = ticker_totalvalue + Cells(i, 7).Value
        Range("J" & summaryrow).Value = ticker_name
        Range("K" & summaryrow).Value = ticker_totalvalue
        summaryrow = summaryrow + 1
        ticker_totalvalue = 0
        
    End If
Next i


End Sub


Sub All_Sheets()
    For Each ws In Worksheets
        ws.Activate
        Call total_volume
    Next ws
End Sub
