# vba-challenge
VBA Module 2 Challenge

Sub Success2()

Dim Ticker As String
Dim opening_price As Double
Dim closing_price As Double
Dim year_change As Double
Dim stock_vol As Double
Dim percent_change As Double
Dim first_row As Integer
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    first_row = 2
    input_row = 1
    stock_vol = 0
    
    'last row of current stock
    
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To last_row
       
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Get Ticker
            
            Ticker = ws.Cells(i, 1).Value
            
            'Prepare for next stock
            
            input_row = input_row + 1
            
            'Opening and closing values
            
            opening_price = ws.Cells(input_row, 3).Value
            closing_price = ws.Cells(i, 6).Value
            
            'Total Stock Volume
            
            For j = input_row To i
            
                stock_vol = stock_vol + ws.Cells(j, 7).Value
                
            Next j
            
            If opening_price = 0 Then
            
                percent_change = closing_price
                
            Else
            
                year_change = closing_price - opening_price
                
                percent_change = year_change / opening_price
                
            End If
         
            'Ticker, yearly change and percent change
            
            ws.Cells(first_row, 9).Value = Ticker
            ws.Cells(first_row, 10).Value = year_change
            ws.Cells(first_row, 11).Value = percent_change
            
        
            ws.Cells(first_row, 11).NumberFormat = "0.00%"
            ws.Cells(first_row, 12).Value = stock_vol
            
            first_row = first_row + 1
            
            stock_vol = 0
            year_change = 0
            percent_change = 0
            
            input_row = i
        
        End If
    
    Next i
    
    '% increase, % decrease and greatest total volume
    
    'Percent change
    
    last_row_k = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    'Define variables for summary table value

    greatest_increase = 0
    greatest_decrease = 0
    greatest_total_volume = 0
    
        For k = 3 To last_row_k

            last_k = k - 1
                        
            'current row for percentage
            
            current_k = ws.Cells(k, 11).Value
            
            'previous row for percentage
            
            previous_k = ws.Cells(last_k, 11).Value
            
            'Greatest total volume row
            
            volume = ws.Cells(k, 12).Value
            
            'previous greatest volume row
            
            previous_vol = ws.Cells(last_k, 12).Value
            
            '% increase
            
            If greatest_increase > current_k And greatest_increase > previous_k Then
                
                greatest_increase = greatest_increase
                
            ElseIf current_k > greatest_increase And current_k > previous_k Then
                
                greatest_increase = current_k
                
                percent_increase = ws.Cells(k, 9).Value
                
            ElseIf previous_k > greatest_increase And previous_k > current_k Then
            
                greatest_increase = previous_k
                
                percent_increase = ws.Cells(last_k, 9).Value
                
            End If
                
            '% decrease
            
            If greatest_decrease < current_k And greatest_decrease < previous_k Then
                
                greatest_decrease = greatest_decrease
    
            ElseIf current_k < greatest_increase And current_k < previous_k Then
                
                greatest_decrease = current_k
                           
                percent_decrease = ws.Cells(k, 9).Value
                
            ElseIf previous_k < greatest_increase And previous_k < current_k Then
            
                greatest_decrease = previous_k

                percent_decrease = ws.Cells(last_k, 9).Value
                
            End If
            
           'Greatest total volume
           
            If greatest_total_volume > volume And greatest_total_volume > previous_vol Then
            
                greatest_total_volume = greatest_total_volume
            
            ElseIf volume > greatest_total_volume And volume > previous_vol Then
            
                greatest_total_volume = volume
                
                greatest_vol = ws.Cells(k, 9).Value
                
            ElseIf previous_vol > greatest_total_volume And previous_vol > volume Then
                
                greatest_total_volume = previous_vol
                
                greatest_vol = ws.Cells(last_k, 9).Value
                
            End If
            
        Next k
        
    'Print greatest increase, greatest decrease, and greatest volume on each worksheet
    
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Get values
    
    ws.Range("O2").Value = percent_increase
    ws.Range("O3").Value = percent_decrease
    ws.Range("O4").Value = greatest_vol
    ws.Range("P2").Value = greatest_increase
    ws.Range("P3").Value = greatest_decrease
    ws.Range("P4").Value = greatest_total_volume
    
    'Greatest % increase and decrease in percentage format
    
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


    'Higlight Yearly Change column based on numerical value
    'Red if <0, green if >=0

    last_row_j = ws.Cells(Rows.Count, "J").End(xlUp).Row
    

        For j = 2 To last_row_j
            
            If ws.Cells(j, 10) > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
    
'Go to next worksheet

Next ws

End Sub
