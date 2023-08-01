Attribute VB_Name = "Module1"
Sub Test()
    
    'Create varible to look through worksheets
    Dim ws          As Worksheet
    
    'Loop through each sheet in the workbook
    For Each ws In Worksheets
        ws.Activate
        
        ' Create variable to check the extent of the data
        Dim Row     As Long
        
        ' Check and store the extent of the data
        Row = Range("A2").End(xlDown).Row
        
        'Input to keep track of position in output table
        Dim Output_table As Integer
        Output_table = 2
        
        'Create variable to store Ticker
        Dim Ticker  As String
        
        'Create variable to store Yearly change
        Dim Yearly_change As Double
        
        'Create variable to store Percentage Change
        Dim Percent_Change As Double
        
        ' Create varible for the opening price
        Dim Opening As Double
        Opening = Cells(2, 3).Value
        
        ' Create varible for the closing price
        Dim Closing As Double
        
        'Create varible to store price change
        Dim Price_Change As Double
        Price_Change = 0
        
        'Create variable to store Total Stock Volume
        Dim Total_volume As Variant
        Total_volume = 0
        
        'Create a varible for the Percentage change
        Dim Percentage_change As Double
        Percentage_change = 0
        
        ' Create headings for the output table which contains the summary parameters
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percentage Change"
        Cells(1, 12) = "Total Stock Volume"
        
        'Create Headers for advanced table
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest total volume"
        
        ' Create variable for Greatest % increase
        Dim Max     As Double
        Max = 0
        
        ' Create variable for Greatest % increase ticker value
        Dim Max_ticker As String
        
        ' Create variable for Greatest % decrease
        Dim Min     As Double
        Min = 0
        
        ' Create variable for Greatest % decrease ticker value
        Dim Min_ticker As String
        
        'Create variable for Grestest Total Volume
        Dim Max_volume As Double
        Greatest_Total_vol = 0
        
        'Create variable for Grestest Total Volume ticker value
        Dim Max_vol_ticker As String
        
        'Create Loop to evaluate when the Ticker changes
        For i = 2 To Row
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Set the Ticker name
                Ticker = Cells(i, 1).Value
                
                'Set the closing price
                Closing = Cells(i, 6).Value
                
                'Calculate price change
                Price_Change = Closing - Opening
                
                'Add stock to the total volume
                Total_volume = Total_volume + Cells(i, 7).Value
                
                'Determine the Percentage change
                If Opening <> 0 Then
                    Percentage_change = (Price_Change / Opening)
                End If
                
                'Print Ticker name to the Output table
                Range("I" & Output_table).Value = Ticker
                
                'Print Yearly Change
                Range("J" & Output_table).Value = Price_Change
                
                'Print Percent_change
                Range("K" & Output_table).Value = Percentage_change
                
                'Print Total Sotck Volume
                Range("L" & Output_table).Value = Total_volume
                
                'Format percentage change column
                Range("K" & Output_table).NumberFormat = "0.00%"
                
                'Format the Price Change with green (positive) and red (negative)
                If Price_Change > 0 Then
                    Range("J" & Output_table).Interior.ColorIndex = 4
                Else
                    Range("J" & Output_table).Interior.ColorIndex = 3
                End If
                
                'Reset values for the next ticker
                Price_Change = 0
                Opening = Cells(i + 1, 3).Value
                Closing = 0
                
                'Set the values for the Advance table
                'Determine the greatest increase and decrease for the tickers
                If Percentage_change > Max Then
                    Max = Percentage_change
                    Max_ticker = Ticker
                ElseIf Percentage_change < Min Then
                    Min = Percentage_change
                    Min_ticker = Ticker
                    
                End If
                
                'Determine the greastest volume
                If Total_volume > Max_vol Then
                    Max_vol = Total_volume
                    Max_vol_ticker = Ticker
                End If
                
                'Reset Total Stock Volume for next ticker
                Total_volume = 0
                
                'Add one row to the Output table
                Output_table = Output_table + 1
                
            Else
                
                'Add stock to the total volume
                Total_volume = Total_volume + Cells(i, 7).Value
                
            End If
        Next i
        
        'Print values into the advanced table
        Cells(2, 16).Value = Max_ticker
        Cells(2, 17).Value = Max
        Cells(3, 16).Value = Min_ticker
        Cells(3, 17).Value = Min
        Cells(4, 16).Value = Max_vol_ticker
        Cells(4, 17).Value = Max_vol
        
        'Format the advanced table percent values
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        
        ' Resize columns to ensure all the data and headings fit into the columns
        ws.Cells.Columns.AutoFit
        
    Next ws
    
End Sub
