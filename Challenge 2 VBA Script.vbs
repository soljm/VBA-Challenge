Sub Challenge2()
For Each ws In Worksheets
'Rename columns and rows as needed
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greates Total Volume"

'Setting variables for holding ticker and total stock volume
Dim Ticker As String
Dim Stock As Double
Stock = 0

'Set up summary
Dim Summary As Integer
Summary = 2

'find last row
Dim LR As String
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

'define open and close variables
Dim Open_value As Double
Dim Close_value As Double
Dim Yearly_change As Double
Dim Percent_change As Double

Open_value = ws.Cells(2, 3).Value

'loop through ticker
For i = 2 To LR
    'check to see if ticker is the same as the row before
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Close_value = ws.Cells(i, 6).Value
        'calculate yearly change
        Yearly_change = Close_value - Open_value
        ws.Range("J" & Summary).Value = Yearly_change
        ws.Range("J" & Summary).NumberFormat = "0.00"
        'colour cell
        If ws.Range("J" & Summary).Value < 0 Then
            ws.Range("J" & Summary).Interior.Color = RGB(255, 171, 171)
        Else
            ws.Range("J" & Summary).Interior.Color = RGB(171, 255, 188)
        End If
        'calculate percent change
        Percent_change = Yearly_change / Open_value
        ws.Range("K" & Summary).Value = Percent_change
        ws.Range("K" & Summary).NumberFormat = "0.00%"
        'set new open value
        Open_value = ws.Cells(i + 1, 3).Value
        Ticker = ws.Cells(i, 1).Value
        'adding the stock volume
        Stock = Stock + ws.Cells(i, 7).Value
        'print ticker and total in summary table
        ws.Range("I" & Summary).Value = Ticker
        ws.Range("L" & Summary).Value = Stock
        'go down row in summary
        Summary = Summary + 1
        'reset stock
        Stock = 0
    'if the ticker is the same as the row before
    Else
        Stock = Stock + ws.Cells(i, 7).Value
    End If
Next i

'Bonus
Dim Max_Percent, Min_Percent, Total_Volume As Double
Dim Max_Ticker, Min_Ticker, Total_Ticker As String

Max_Percent = ws.Range("K2").Value
Max_Ticker = ws.Range("I2").Value
Min_Percent = ws.Range("K2").Value
Min_Ticker = ws.Range("I2").Value
Total_Volume = ws.Range("L2").Value
Total_Ticker = ws.Range("I2").Value

Dim LRSummary As String
LRSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To LRSummary
    'finding greatest % increase
    If ws.Cells(j, 11).Value > Max_Percent Then
        Max_Percent = ws.Cells(j, 11).Value
        Max_Ticker = ws.Cells(j, 9).Value
    End If
    'adding values to cells
    ws.Range("O2").Value = Max_Ticker
    ws.Range("P2").Value = Max_Percent
    ws.Range("P2:P3").NumberFormat = "0.00%"
    'finding greatest % decrease
    If ws.Cells(j, 11).Value < Min_Percent Then
        Min_Percent = ws.Cells(j, 11).Value
        Min_Ticker = ws.Cells(j, 9).Value
    End If
    'adding values to cell
    ws.Range("O3").Value = Min_Ticker
    ws.Range("P3").Value = Min_Percent
    'finding greatest total volume
    If ws.Cells(j, 12).Value > Total_Volume Then
        Total_Volume = ws.Cells(j, 12).Value
        Total_Ticker = ws.Cells(j, 9).Value
    End If
    'adding values to cells
    ws.Range("O4").Value = Total_Ticker
    ws.Range("P4").Value = Total_Volume
Next j
'autofit summary columns
ws.Columns("I:P").AutoFit
Next ws
End Sub

