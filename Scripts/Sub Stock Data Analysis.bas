Attribute VB_Name = "Module1"
Sub stock_data_analysis()

For Each WS In Worksheets
    
    WS.Activate

    'Adding header columns
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    
    'Additional Summary Table
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    'Variable for Last_Row
    Dim Last_Row As Long

    'Find the last non-blank cell in column A
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

    'Setting Constants
    Column = 1
    Row = 2

    'Setting Initial Conditions
    Total_Stock_Volume = 0
    Open_Price = Cells(2, Column + 2).Value
    Greatest_Percent_Increase = 0
    Greatest_Percent_Decrease = 0
    Greatest_Total_Volume = 0
    
    'Looping through stock data
    For I = 2 To Last_Row

        'Aggregating total volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, Column + 6).Value
        Cells(Row, "L").Value = Total_Stock_Volume
        
        'Conditions for ticker change and calculations
        If Cells(I + 1, Column).Value <> Cells(I, Column).Value Then
        
            'Set Ticker Name and add to Ticker column
            Ticker_Name = Cells(I, Column).Value
            Cells(Row, "I").Value = Ticker_Name
        
            'Setting Close Price
            Close_Price = Cells(I, Column + 5).Value
            'Calculating Yearly Change and adding to column
            Yearly_Change = Close_Price - Open_Price
            Cells(Row, "J") = Yearly_Change
        
            'Calculating Percent Change, adding to column, and changing number formatting
            Percent_Change = (Close_Price - Open_Price) / Open_Price
            Cells(Row, "K") = Percent_Change
            Cells(Row, "K").NumberFormat = "0.00%"
        
            'Gathering the greatest percent increase
            If Percent_Change > 0 And Percent_Change > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = Cells(Row, "K").Value
                Greatest_Percent_Change = Percent_Change
                Cells(2, "Q").Value = Greatest_Percent_Change
                Cells(2, "Q").NumberFormat = "0.00%"
                Cells(2, "P").Value = Cells(Row, "I")
            End If
            
            'Gathering the greatest percent decrease
            If Percent_Change < 0 And Percent_Change < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Cells(Row, "K").Value
                Greatest_Percent_Change = Percent_Change
                Cells(3, "Q").Value = Greatest_Percent_Change
                Cells(3, "Q").NumberFormat = "0.00%"
                Cells(3, "P").Value = Cells(Row, "I")
            End If
            
            'Gathering the greatest total volume
            If Total_Stock_Volume > 0 And Total_Stock_Volume > Greatest_Total_Volume Then
                Greatest_Total_Volume = Cells(Row, "L").Value
                Greatest_Total_Change = Total_Stock_Volume
                Cells(4, "Q").Value = Greatest_Total_Change
                Cells(4, "P").Value = Cells(Row, "I")
            End If
            
            'Implementing conditional formatting
            If Yearly_Change > 0 Then
                Cells(Row, "J").Interior.ColorIndex = 4
            ElseIf Yearly_Change < 0 Then
                Cells(Row, "J").Interior.ColorIndex = 3
            Else
                Cells(Row, "J").Interior.ColorIndex = 6
            End If
            
            'Increasing row value
            Row = Row + 1
            
            'Resetting values
            Total_Stock_Volume = 0
            Open_Price = Cells(I + 1, Column + 2).Value
            
        End If
    
    Next
    
Next WS

End Sub
