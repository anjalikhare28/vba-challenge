Attribute VB_Name = "Module3"
Sub Stock_Summary()
Attribute Stock_Summary.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim sheet As Worksheet
    'Set sheet = Sheets("A")
    
    For Each sheet In Worksheets
        
        'Create column lables
        sheet.Cells(1, 9).Value = "Ticker"
        sheet.Cells(1, 10).Value = "Yearly Change"
        sheet.Cells(1, 11).Value = "Percent Change"
        sheet.Cells(1, 12).Value = "Total Stock Volume"
    
        'Declare variables and set values
        Dim ticker As String
        Dim volume As Double
        volume = 0
        Dim row_count As Long
        row_count = 2
        Dim opening_price As Double
        opening_price = 0
        Dim closing_price As Double
        closing_price = 0
        Dim price_change As Double
        price_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim last_row As Long
        lastrow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' loop from the dataset
        For i = 2 To lastrow
            
            ' pick open price
            If sheet.Cells(i, 1).Value <> sheet.Cells(i - 1, 1).Value Then
               opening_price = sheet.Cells(i, 3).Value
            End If
            ' calculate total volume
            volume = volume + sheet.Cells(i, 7)
            
            ' check if ticker is changing
            If sheet.Cells(i, 1).Value <> sheet.Cells(i + 1, 1).Value Then
                
                ' pick closing price
                closing_price = sheet.Cells(i, 6).Value

                ' calculate price change
                price_change = closing_price - opening_price
                
                ' calculate percent change
                If opening_price = 0 And closing_price = 0 Then
                    percent_change = 0
                ElseIf opening_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = price_change / opening_price
                End If
                
                ' Move values to second table
                sheet.Cells(row_count, 9).Value = sheet.Cells(i, 1).Value
                sheet.Cells(row_count, 12).Value = volume
                sheet.Cells(row_count, 10).Value = price_change
                sheet.Cells(row_count, 11).Value = percent_change
                                
                'format positive or negative change & percentage
                If price_change >= 0 Then
                    sheet.Cells(row_count, 10).Interior.ColorIndex = 4
                Else
                    sheet.Cells(row_count, 10).Interior.ColorIndex = 3
                End If
                
                sheet.Cells(row_count, 11).NumberFormat = "0.00%"
                
                'Pick next empty row
                row_count = row_count + 1

                'Reset local variables
                volume = 0
                opening_price = 0
                closing_price = 0
                price_change = 0
                percent_change = 0
                
            End If
        Next i
        
        ' Bonus, create third table
        sheet.Cells(2, 15).Value = "Greatest % Increase"
        sheet.Cells(3, 15).Value = "Greatest % Decrease"
        sheet.Cells(4, 15).Value = "Greatest Total Volume"
        sheet.Cells(1, 16).Value = "Ticker"
        sheet.Cells(1, 17).Value = "Value"
        
        ' last row of second dataset
        lastrow = sheet.Cells(Rows.Count, 9).End(xlUp).Row

        ' declare variables for various performance
        Dim best_ticker As String
        Dim best_price As Double
        Dim worst_ticker As String
        Dim worst_price As Double
        Dim largest_ticker As String
        Dim largest_volume As Double
        
        ' Set all variables to first ticker
        best_price = sheet.Cells(2, 11).Value
        worst_price = sheet.Cells(2, 11).Value
        largest_volume = sheet.Cells(2, 12).Value

        ' loop through second dataset
        For i = 2 To lastrow
                    
            ' best ticker
            If sheet.Cells(i, 11).Value > best_price Then
                best_price = sheet.Cells(i, 11).Value
                best_ticker = sheet.Cells(i, 9).Value
            End If

            ' worst ticker
            If sheet.Cells(i, 11).Value < worst_price Then
                worst_price = sheet.Cells(i, 11).Value
                worst_ticker = sheet.Cells(i, 9).Value
            End If

            ' largest ticker
            If sheet.Cells(i, 12).Value > largest_volume Then
                largest_volume = sheet.Cells(i, 12).Value
                largest_ticker = sheet.Cells(i, 9).Value
            End If
        Next i
        
        ' move values to third table
        sheet.Cells(2, 16).Value = best_ticker
        sheet.Cells(2, 17).Value = best_price
        sheet.Cells(2, 17).NumberFormat = "0.00%"
        sheet.Cells(3, 16).Value = worst_ticker
        sheet.Cells(3, 17).Value = worst_price
        sheet.Cells(3, 17).NumberFormat = "0.00%"
        sheet.Cells(4, 16).Value = largest_ticker
        sheet.Cells(4, 17).Value = largest_volume
        
        sheet.Columns("I:L").EntireColumn.AutoFit
        sheet.Columns("O:Q").EntireColumn.AutoFit
    Next sheet
    
End Sub
