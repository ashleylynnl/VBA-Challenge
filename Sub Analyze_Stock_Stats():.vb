Sub Analyze_Stock_Stats():

    ' Set initial variables
    Dim Ticker As String
    Dim NumberTickers As Integer
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentIncreaseTicker As String
    Dim GreatestPercentDecrease As Double
    Dim GreatestPercentDecreaseTicker As String
    Dim GreatestStockVolume As Double
    Dim GreatestStockVolumeTicker As String

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets

    ws.Activate

    ' Find the last row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Header Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set baseline variables for each worksheet
    NumberTickers = 0
    Ticker = ""
    YearlyChange = 0
    OpeningPrice = 0
    PercentChange = 0
    TotalStockVolume = 0
    
    ' Loop through the list of tickers
    For i = 2 To LastRowState

        ' Get name of the ticker symbol
        Ticker = Cells(i, 1).Value
        
        ' Get the start of the year opening price for the ticker
        If OpeningPrice = 0 Then
            OpeningPrice = Cells(i, 3).Value
        End If
        
        ' Add up the total stock volume values for a ticker
        TotalStockVolume = tTotalStockVolume + Cells(i, 7).Value
        
        ' Run this if we get to a different ticker in the list
        If Cells(i + 1, 1).Value <> Ticker Then
        
            ' Increment the number of tickers when we get to a different ticker in the list.
            NumberTickers = NumberTickers + 1
            Cells(NumberTickers + 1, 9) = Ticker
            
            ' Get the end of the year closing price for ticker
            ClosingPrice = Cells(i, 6)
            
            ' Calculate Yearly Change
            YearlyChange = ClosingPrice - OpeningPrice
            
            ' Add Yearly Change value to the appropriate cell
            Cells(NumberTickers + 1, 10).Value = YearlyChange
            
            ' If Yearly Change value is greater than 0, color cell green.
            If YearlyChange > 0 Then
                Cells(NumberTickers + 1, 10).Interior.ColorIndex = 4
            ' If Yearly Change value is less than 0, color cell red.
            ElseIf YearlyChange < 0 Then
                Cells(NumberTickers + 1, 10).Interior.ColorIndex = 3
            ' If Yearly Change value is 0, color cell yellow.
            Else
                Cells(NumberTickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change value for ticker.
            If OpeningPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = (YearlyChange / OpeningPrice)
            End If
            
            
            ' Format the percent_change value as a percent.
            Cells(NumberTickers + 1, 11).Value = Format(PercentChange, "Percent")
        
            ' Set opening price back to 0 when we get to a different ticker in the list.
            OpeningPrice = 0
            
            ' Add total stock volume value to the appropriate cell in each worksheet.
            Cells(NumberTickers + 1, 12).Value = TotalStockVolume
            
            ' Set total stock volume back to 0 when we get to a different ticker in the list.
            TotalStockVolume = 0
        End If
        
    Next i
    
    ' Display greatest percent increase, greatest percent decrease, and greatest total volume for each year
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Get the last row
    LastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    GreatestPercentIncrease = Cells(2, 11).Value
    GreatestPercentIncreaseTicker = Cells(2, 9).Value
    GreatestPercentDecrease = Cells(2, 11).Value
    GreatestPercentDecreaseTicker = Cells(2, 9).Value
    GreatestStockVolume = Cells(2, 12).Value
    GreatestStockVolumeTicker = Cells(2, 9).Value
    
    
    ' Loop through the list of tickers
    For i = 2 To LastRowState
    
        ' Find ticker with the Greatest Percent Increase
        If Cells(i, 11).Value > GreatestPercentIncrease Then
            GreatestPercentIncrease = Cells(i, 11).Value
            GreatestPercentIncreaseTicker = Cells(i, 9).Value
        End If
        
        ' Find ticker with the Greatest Percent Decrease
        If Cells(i, 11).Value < GreatestPercentDecrease Then
            GreatestPercentDecrease = Cells(i, 11).Value
            GreatestPercentDecreaseTicker = Cells(i, 9).Value
        End If
        
        ' Find ticker with Greatest Stock Volume
        If Cells(i, 12).Value > GreatestStockVolume Then
            GreatestStockVolume = Cells(i, 12).Value
            GreatestStockVolumeTicker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet
    Range("P2").Value = Format(GreatestPercentIncreaseTicker, "Percent")
    Range("Q2").Value = Format(GreatestPercentIncrease, "Percent")
    Range("P3").Value = Format(GreatestPercentDecreaseTicker, "Percent")
    Range("Q3").Value = Format(GreatestPercentDecrease, "Percent")
    Range("P4").Value = GreatestStockVolumeTicker
    Range("Q4").Value = GreatestStockVolume
    
Next ws


End Sub

