# Stocks Analysis

### Challenge 2

Steve wanted to do more research for his parents and was gathering data for the entire stock market in the past 2 years. We needed to improve and refactor the initial code because of new codes that needed to be integrated into our Analysis. The first thing I focused on was creating 2-Dimensional Array As Varyant type because the final number of tickers was not known. I found out that if in this last part of the refactor in order to make the formula run more efficiently, we had to integrate the Thicker of opening price, closing price, volumes, and shares. For that reason, we had to bring the variable tickerIndex so that we could then use it with For Loop and be able to exchange in between stock tickers and put it in the Refactor. In this manner we can run the analysis effectively and improve it from the last version. 


    Sub Challenge()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("Challenge 2").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
    'Variant type
        Dim ticker() As Variant

        Worksheets(yearValue).Activate

    'get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'set tickerIndex to zero
        tickerIndex = 0
   
    'subsequently ReDim in the FOR loop
        ReDim ticker(4, tickerIndex)
   
    'initial value for the ticker
        ticker(1, tickerIndex) = Cells(2, 1).Value
   
        For I = 2 To RowCount
        If ticker(1, tickerIndex) = Cells(I, 1).Value Then
            If I = 2 Then
                ticker(2, tickerIndex) = Cells(I, 3).Value
            End If
            'Set the closing price
                ticker(3, tickerIndex) = Cells(I, 6).Value
            'Increment Volume
                ticker(4, tickerIndex) = ticker(4, tickerIndex) + Cells(I, 8).Value
        Else
            tickerIndex = tickerIndex + 1
            ReDim Preserve ticker(4, tickerIndex)
            ticker(1, tickerIndex) = Cells(I, 1).Value
            ticker(2, tickerIndex) = Cells(I, 3).Value
            ticker(3, tickerIndex) = Cells(I, 6).Value
            'Volume
                ticker(4, tickerIndex) = ticker(4, tickerIndex) + Cells(I, 8).Value
        End If
    Next I
    
    'Write outputs to spreadsheet
        Worksheets("Challenge 2").Activate
        For I = 0 To tickerIndex
        Cells(4 + I, 1).Value = ticker(1, I)
        Cells(4 + I, 2).Value = ticker(4, I)
        Cells(4 + I, 3).Value = ticker(3, I) / ticker(2, I) - 1
    Next I
    
 

    'Formatting
    Worksheets("Challenge 2").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    For I = dataRowStart To dataRowEnd

        If Cells(I, 3) > 0 Then

            Cells(I, 3).Interior.Color = vbGreen

        Else

            Cells(I, 3).Interior.Color = vbRed

        End If

    Next I
    


    End Sub


![stocks1](https://github.com/lrovira/-stocks-analysis-/blob/master/stocks1.png)

