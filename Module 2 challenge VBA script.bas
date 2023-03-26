Attribute VB_Name = "Module2"
'Create a script that loops through all the stocks for one year
  'and outputs the following information:

  ' The ticker symbol.

  ' Yearly change from opening price at the beginning of a given year
  'to the closing price at the end of that year.

  ' The percent change from opening price at the beginning of a given year
  ' to the closing price at the end of that year.

  ' The total stock volume of the stock.'

Sub StockDataLoop()

    'Set ws as a worksheet object variable.
    Dim headers() As Variant
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    'Set header Info
    headers() = Array("ticker ", "date ", "open", "high", "low", "close", "vol", "", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
    
    For Each ws In wb.Sheets
        With ws
        .Rows(1).Value = ""
            For i = LBound(headers()) To UBound(headers())
            .Cells(1, 1 + i).Value = headers(i)
                     
           Next i
            .Rows(1).Font.Bold = True
            .Rows(1).VerticalAlignment = xlCenter
           End With
              
         'Loop through each column in the worksheet.
            For Each col In ws.UsedRange.Columns
            
            ' Autofit the column width based on the content.
            
            col.EntireColumn.AutoFit
            
        Next col
            
    Next ws
       
        'Loop through all worksheets in the workbook.
        
       For Each ws In Worksheets
        'Set the variables for calculations
        Dim Ticker_Name As String
        Ticker_Name = ""
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Beg_Price As Double
        Beg_Price = 0
        Dim End_Price As Double
        End_Price = 0
        Dim Yearly_Price_Change As Double
        Yearly_Price_Change = 0
        Dim Yearly_Price_Change_Percent As Double
        Yearly_Price_Change_Percent = 0
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = ""
        Dim Min_Ticker_Name As String
        Min_Ticker_Name = ""
        Dim Max_Percent As Double
        Max_Percent = 0
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker_Name As String
        Max_Volume_Ticker_Name = ""
        Dim Max_Volume As Double
        Max_Volume = 0
        
        ' Set location for Summary Table Row.
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set row count for the entire workbook.
        Dim lastRow As Long
        
        
        'Loop through all sheets to find last non-empty cell.
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Initial value of beginning stock value for the first ticker of main worksheet.
        Beg_Price = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
        
          'Check if still on same ticker.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Starting point for ticker name.
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Price calculations.
                End_Price = ws.Cells(i, 6).Value
                Yearly_Price_Change = End_Price - Beg_Price
                
                'Condition for a zero value.
                If Beg_Price <> 0 Then
                    Yearly_Price_Change_Percent = (Yearly_Price_Change / Beg_Price) * 100
                    
                End If
                 
                ' Add total volume to ticker name.
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
                
                ' Print ticker name to column I in Summary Table Row.
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' Print yearly price change in colum J.
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
                
                'Color fill the yearly price chang: green is positive, red is negative.
                If (Yearly_Price_Change > 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    ElseIf (Yearly_Price_Change <= 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                'Print the remaining outputs to appropriate Summary Table Row.
                
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Price_Change_Percent) & "%")
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                'Add to summary table row count.
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Get the next beginning price and perform calculations.
                Beg_Price = ws.Cells(i + 1, 3).Value
                
            If (Yearly_Price_Change_Percent > Max_Percent) Then
                Max_Percent = Yearly_Price_Change_Percent
                Max_Ticker_Name = Ticker_Name
                
                ElseIf (Yearly_Price_Change_Percent < Min_Percent) Then
                Min_Percent = Yearly_Price_Change_Percent
                Min_Ticker_Name = Ticker_Name
                
            End If
            
            If (Total_Ticker_Volume > Max_Volume) Then
                Max_Volume = Total_Ticker_Volume
                Max_Volume_Ticker_Name = Ticker_Name
            End If
                
            'Reset values
            
            Yearly_Price_Change_Percent = 0
            Total_Ticker_Volume = 0
            
        Else
        
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    
            'Print values in assigned cells
            ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
            ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
            ws.Range("P2").Value = Max_Ticker_Name
            ws.Range("P3").Value = Min_Ticker_Name
            ws.Range("P4").Value = Max_Volume_Ticker_Name
            ws.Range("Q4").Value = Max_Volume
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
    Next ws
    
End Sub
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

 0 , 1 7 6 5 0 9 6 9 , 1 9 2 0 0 0 8 7 , 6 1 3 7 4 3 5 , 2 5 0 3 6 3 1 1 , 1 7 5 7 3 6 4 3 , 6 1 1 2 0 3 3 , 7 8 7 6 6 2 1 , 3 7 3 3 2 9 4 7 , 3 7 8 8 9 3 6 6 , 1 9 4 3 7 7 1 8 , 6 3 6 6 2 2 4 , 1 7 1 3 3 8 3 1 , 7 8 7 6 6 2 2 , 2 0 9 9 8 1 5 7 , 5 7 3 8 9 9 2 7 1 , 4 0 9 2 3 3 4 8 , 7 9 9 6 8 0 7 , 6 5 7 5 0 7 1 , 5 4 0 8 4 4 4 4 4 , 7 8 7 6 6 2 3 , 6 3 6 6 2 9 8 , 2 0 8 3 3 9 5 1 , 6 1 6 2 3 8 2 , 1 6 8 5 9 3 6 3 , 1 8 1 4 7 4 6 2 , 1 9 1 8 2 1 4 8 , 2 0 5 1 0 0 2 6 , 1 9 9 3 3 2 6 1 , 8 9 8 8 2 9 3 , 9 1