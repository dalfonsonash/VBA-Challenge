'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
Sub stock_analysis()
    
    'Set ws as a worksheet object variable.
    Dim headers() As Variant
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    For Each ws In ThisWorkbook.Worksheets
    
    'Set headers.
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

    Next ws
        ' Declare variables
        Dim ticker As String
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_volume As Double
        Dim last_row As Long
        Dim summary_row As Long
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String
    
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate

        ' Set initial values for variables
        summary_row = 2
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0

        ' Loop through all stocks
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Check if we are still on the same ticker or if we have moved on to a new one
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set ticker variable
                ticker = ws.Cells(i, 1).Value

                ' Set opening and closing prices
                opening_price = ws.Cells(i - 11, 3).Value
                closing_price = ws.Cells(i, 6).Value

                ' Calculate yearly change and percent change
                yearly_change = closing_price - opening_price
                percent_change = yearly_change / opening_price

                ' Add total volume
                total_volume = Application.Sum(Range(Cells(i - 11, 7), Cells(i, 7)))

                ' Print results to summary table
                ws.Cells(summary_row, 9).Value = ticker
                ws.Cells(summary_row, 10).Value = yearly_change

                ' Apply conditional formatting to yearly change column
                If yearly_change > 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3 ' Red
                End If

                ws.Cells(summary_row, 11).Value = percent_change
                ws.Cells(summary_row, 12).Value = total_volume

                ' Find greatest increase, greatest decrease, and greatest volume
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
                End If

                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
                End If

                If total_volume > greatest_volume Then
                    greatest_volume = total_volume
                    greatest_volume_ticker = ticker
                End If

                ' Move to next row in summary table
                summary_row = summary_row + 1

            End If

        Next i
    
        ' Print results for greatest increase, greatest decrease, and greatest volume
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(2, 17).Value = greatest_increase
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(3, 17).Value = greatest_decrease
        ws.Cells(4, 16).Value = greatest_volume_ticker
        ws.Cells(4, 17).Value = greatest_volume
        ws.Columns("A:S").AutoFit
    Next ws

End Sub

