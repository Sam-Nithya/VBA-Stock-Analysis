Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Integer
    Dim OriginalDate As String
    Dim FormattedDate As Date
    Dim FirstOpenPrice As Boolean

    ' Variables for tracking the greatest values
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim CurrentPercentChange As Double

    ' Initialize tracking variables
    GreatestIncrease = -9999999
    GreatestDecrease = 9999999
    GreatestVolume = 0
    GreatestIncreaseTicker = ""
    GreatestDecreaseTicker = ""
    GreatestVolumeTicker = ""

    ' Loop through all worksheets
    For Each ws In Worksheets
        ws.Activate
        
        ' Determine the last row of data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the summary row counter
        SummaryRow = 2
        
        ' Set up headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize TotalVolume and FirstOpenPrice flag
        TotalVolume = 0
        FirstOpenPrice = True
        
        ' Loop through each row of stock data
        For i = 2 To LastRow
            ' Convert date format from YYYYMMDD to MM/DD/YYYY in the date column (assuming column B contains dates)
            OriginalDate = ws.Cells(i, 2).Value
            If IsNumeric(OriginalDate) And Len(OriginalDate) = 8 Then
                FormattedDate = DateSerial(Left(OriginalDate, 4), Mid(OriginalDate, 5, 2), Right(OriginalDate, 2))
                ws.Cells(i, 2).Value = FormattedDate
            End If
            ' Apply date formatting
            ws.Cells(i, 2).NumberFormat = "M/D/YYYY"
            
            ' Check if we are on the first occurrence of the ticker for capturing the opening price
            If FirstOpenPrice Then
                OpenPrice = ws.Cells(i, 3).Value ' Open price from the first occurrence
                FirstOpenPrice = False
            End If
            
            ' Check if the current ticker is the last entry of that ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Capture the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                ' Add the last day's volume to the total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                ' Get the closing price from the last occurrence of the ticker
                ClosePrice = ws.Cells(i, 6).Value
                ' Calculate the quarterly change
                QuarterlyChange = ClosePrice - OpenPrice
                ' Calculate the percent change
                If OpenPrice <> 0 Then
                    PercentChange = QuarterlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If

                ' Populate the summary table with calculated values
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Apply number formatting to Total Stock Volume
                ws.Cells(SummaryRow, 12).NumberFormat = "#,##0" ' Format as number with thousand separators

                ' Apply number formatting to Percent Change
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%" ' Format as percentage with two decimal places

                ' Apply conditional formatting for Quarterly Change only
                If QuarterlyChange >= 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If

                ' Update the greatest values
                CurrentPercentChange = PercentChange * 100 ' Convert to percentage for comparison

                If CurrentPercentChange > GreatestIncrease Then
                    GreatestIncrease = CurrentPercentChange
                    GreatestIncreaseTicker = Ticker
                End If

                If CurrentPercentChange < GreatestDecrease Then
                    GreatestDecrease = CurrentPercentChange
                    GreatestDecreaseTicker = Ticker
                End If

                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If

                ' Reset the total volume counter and FirstOpenPrice flag
                TotalVolume = 0
                FirstOpenPrice = True
                ' Move to the next summary row
                SummaryRow = SummaryRow + 1

            Else
                ' Accumulate volume if the ticker is the same
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output the greatest results in columns after column 14
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Add headers for Ticker and Value
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Display tickers near the greatest results
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        
        ' Display the corresponding values in the Value column
        ws.Cells(2, 17).Value = GreatestIncrease / 100 ' Convert to percentage
        ws.Cells(3, 17).Value = GreatestDecrease / 100 ' Convert to percentage
        ws.Cells(4, 17).Value = GreatestVolume
        
        ' Apply number formatting to the Value column
        ws.Cells(2, 17).NumberFormat = "0.00%" ' Format as percentage with two decimal places
        ws.Cells(3, 17).NumberFormat = "0.00%" ' Format as percentage with two decimal places
        ws.Cells(4, 17).NumberFormat = "#,##0" ' Format as number with thousand separators

    Next ws

End Sub

