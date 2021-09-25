Attribute VB_Name = "Module1"
Sub RunAllSheets()
    Dim xSheets As Worksheet
    Application.ScreenUpdating = False
    For Each xSheets In Worksheets
        xSheets.Select
        Call StockOutput
        Call Cal_Greatest
    Next
    Application.ScreenUpdating = True
End Sub




Sub StockOutput()

outRow = 2
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"


Total_stock_Vol = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    Cur_Ticker = Cells(i, 1).Value

    Next_Ticker = Cells((i + 1), 1).Value
    
    Prev_Ticker = Cells((i - 1), 1).Value
    
    volume = Cells(i, 7).Value
    
    If (Cur_Ticker <> Prev_Ticker) Then

        Total_stock_Vol = Total_stock_Vol + volume

        Stockopen_value = Cells(i, 3).Value

        Cells(outRow, 9).Value = Cur_Ticker

    End If

    
    If (Cur_Ticker = Next_Ticker) Then

    Total_stock_Vol = Total_stock_Vol + volume

    Else

        StockClose_value = Cells(i, 6).Value
        Yearly_Change = StockClose_value - Stockopen_value

        
       If (Stockopen_value = 0) Then
        
           Percentage_change = 0
            
        
        Else
            Percentage_change = (StockClose_value - Stockopen_value) / Stockopen_value
        End If
        
        Total_stock_Vol = Total_stock_Vol + volume


        Cells(outRow, 9).Value = Cur_Ticker
        Cells(outRow, 10).Value = Yearly_Change
        If Yearly_Change < 0 Then
            Cells(outRow, 10).Interior.Color = RGB(255, 0, 0)
        Else
            Cells(outRow, 10).Interior.Color = RGB(0, 255, 0)
        End If
        
        Cells(outRow, 11).Value = Percentage_change
        Cells(outRow, 11).NumberFormat = "0.00%"
        Cells(outRow, 12).Value = Total_stock_Vol
        
        Total_stock_Vol = 0
        outRow = outRow + 1

    End If

Next i



End Sub

Sub Cal_Greatest()

outRow = 2
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

lastrow = Cells(Rows.Count, 11).End(xlUp).Row

greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

For i = 2 To lastrow
    cur_increase = Cells(i, 11).Value
    If cur_increase > greatest_increase Then
        greatest_increase = cur_increase
        greatest_Ticker = Cells(i, 9).Value
    End If
Next i
    
    Cells(2, 17).Value = greatest_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(2, 16).Value = greatest_Ticker
    
For i = 2 To lastrow
    cur_decrease = Cells(i, 11).Value
    If cur_decrease < greatest_decrease Then
        greatest_decrease = cur_decrease
        greatest_Ticker = Cells(i, 9).Value
    End If
Next i
    
    Cells(3, 17).Value = greatest_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = greatest_Ticker
    
For i = 2 To lastrow
    cur_increase = Cells(i, 12).Value
    If cur_increase > greatest_volume Then
        greatest_volume = cur_increase
        greatest_Ticker = Cells(i, 9).Value
    End If
Next i
    
    Cells(4, 17).Value = greatest_volume
    Cells(4, 16).Value = greatest_Ticker
        
End Sub

