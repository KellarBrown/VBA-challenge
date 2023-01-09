Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()
Dim ticker As String
Dim ticker2 As String
Dim ticker3 As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim tickerCount As Long
Dim fOpenLocation As Long
Dim lCloseLocation As Long
Dim yearlyChange As Double
Dim count As Long
Dim count2 As Long
Dim totalVol As LongLong
Dim maxPercentIncrease As Double
Dim maxPercentDecrease As Double
Dim maxTotalVol As LongLong
tickerCount = 1
ticker = ""
ticker2 = Cells(2, 1).Value
ticker3 = Cells(2, 1).Value
count = 1
count2 = 1
fOpenLocation = 2
totalVol = 0
maxTotalVol = 0
maxPercentIncrease = 0
maxPercentDecrease = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest%Increase"
Cells(3, 15).Value = "Greatest%Decrease"""
Cells(4, 15).Value = "Greatest Total Volume"

For i = 2 To Cells(Rows.count, 1).End(xlUp).Row
    If ticker <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    tickerCount = tickerCount + 1
    Cells(tickerCount, 9).Value = ticker
    End If
Next i
For j = 2 To Cells(Rows.count, 1).End(xlUp).Row + 1
    If ticker2 <> Cells(j, 1).Value And j <> Cells(Rows.count, 1).End(xlUp).Row Then
    ticker2 = Cells(j, 1).Value
    count = count + 1
    lCloseLocation = j - 1
    yearlyChange = Cells(lCloseLocation, 6).Value - Cells(fOpenLocation, 3).Value
    Cells(count, 10).Value = yearlyChange
    Cells(count, 11).Value = Format(yearlyChange / Cells(fOpenLocation, 3).Value, "#.00%")
    fOpenLocation = j
    End If
Next j
For k = 2 To Cells(Rows.count, 1).End(xlUp).Row + 1
    If ticker3 <> Cells(k, 1).Value Then
    ticker3 = Cells(k, 1).Value
    count2 = count2 + 1
    Cells(count2, 12).Value = totalVol
    totalVol = 0
    End If
    totalVol = totalVol + Val(Cells(k, 7).Value)
Next k
For m = 2 To Cells(Rows.count, 10).End(xlUp).Row
    If Cells(m, 11).Value > maxPercentIncrease Then
    maxPercentIncrease = Cells(m, 11).Value
    Cells(2, 16).Value = Cells(m, 9).Value
    ElseIf Cells(m, 11).Value < maxPercentDecrease Then
    maxPercentDecrease = Cells(m, 11).Value
    Cells(3, 16).Value = Cells(m, 9).Value
    End If
    If Cells(m, 12).Value > maxTotalVol Then
    maxTotalVol = Cells(m, 12).Value
    Cells(4, 16).Value = Cells(m, 9).Value
    End If
Next m
Cells(2, 17).Value = Format(maxPercentIncrease, "0.00%")
Cells(3, 17).Value = Format(maxPercentDecrease, "0.00%")
Cells(4, 17).Value = maxTotalVol

End Sub
    
