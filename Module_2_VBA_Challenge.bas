Attribute VB_Name = "Module1"
Sub Stocks()

'defining everything

    Dim ticker As String
    Dim yearly_open As Double
    Dim yearly_close As Double
    Dim percent_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String

    ticker_table_row = 2
    total_stock_vol = 0

'starting

    For i = 2 To 22771
    ticker = Cells(i, 1)
    
    total_stock_vol = total_stock_vol + Cells(i, 7).Value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    yearly_close = Cells(i, 6).Value
    yearly_change = yearly_close - yearly_open
    Range("J" & ticker_table_row).Value = yearly_change
                
' yearly_change colours for increase and decrease

    Dim changeCell As Range
    Set changeCell = Range("J" & ticker_table_row)

' red for decrease and green for increase

    If yearly_change > 0 Then
    changeCell.Interior.Color = RGB(0, 255, 0)
    ElseIf yearly_change < 0 Then
    changeCell.Interior.Color = RGB(255, 0, 0)
    End If
            
' percent change = yearly change / yearly open
' since yearly open is 0 we switch to 1 (math)

    If yearly_open = 0 Then
    yearly_open = 1
    End If
            
    percent_change = yearly_change / yearly_open
    Range("K" & ticker_table_row).Value = percent_change
        
' ticker = ticker + Cells(i, 9).Value

    Range("I" & ticker_table_row).Value = ticker
    Range("L" & ticker_table_row).Value = total_stock_vol
    total_stock_vol = 0
    ticker_table_row = ticker_table_row + 1
    End If
        
' yearly change = close - open

    If Cells(i - 1, 1) <> Cells(i, 1) Then
    yearly_open = Cells(i, 3)
    End If
        
'greatest increase, greatest decrease, greatest stock vol

    If percent_change > greatest_increase Then
    greatest_increase = percent_change
    greatest_increase_ticker = ticker
    End If

    If percent_change < greatest_decrease Then
    greatest_decrease = percent_change
    greatest_decrease_ticker = ticker
    End If

    If total_stock_vol > greatest_volume Then
    greatest_volume = total_stock_vol
    greatest_volume_ticker = ticker
    End If
         
    Next i
    
'defining where the calculatons for the increases and decreases will go including the ticker etc... titles
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("N2").Value = "Greatest % Increase"
    Range("O2").Value = "Greatest % Decrease"
    Range("P2").Value = "Greatest Total Volume"

    Range("N3").Value = greatest_increase
    Range("O3").Value = greatest_decrease
    Range("P3").Value = greatest_volume

    Range("N4").Value = greatest_increase_ticker
    Range("O4").Value = greatest_decrease_ticker
    Range("P4").Value = greatest_volume_ticker
    
End Sub

