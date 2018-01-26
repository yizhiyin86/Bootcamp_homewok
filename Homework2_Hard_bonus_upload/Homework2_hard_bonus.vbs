Sub Homework2_hard_bonus()
    ' ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
            'Dim LastRow
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Fill in header for output table
            header_name = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
            ws.Range("I1:L1").Value = header_name
            
            '____________________________________
            'PartI to fill in ticker name and Total stock volume
            '____________________________________
            
                'keep track of the output row for name and total
                name_row = 2
                'Assign the initial value of total_volume as 0
                total_volume = 0
                'Enter the loop
                For i = 2 To LastRow
                    'identify rows of the cell ticker
                    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                    
                        'sum the first and the second to the last volume for each ticker
                        total_volume = total_volume + ws.Cells(i, 7).Value
                        
                    'identify the last volume of each ticker
                    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    
                        'add the last volume of each ticker and after this
                        'the total_volume is the final total_volume for each item
                        total_volume = total_volume + ws.Cells(i, 7).Value
                        
                        'Assign the ticker name into a cell under colume I
                        ws.Range("I" & name_row).Value = ws.Cells(i, 1).Value
                        
                        'Assign the final total_volume into a cell under colume L
                        ws.Range("L" & name_row).Value = total_volume
                        
                        'Clear the previous total_volume to 0 for next ticker
                        total_volume = 0
                        
                        'add one row for the name_row so next name and volume of the
                        'next ticker can be added under the previous one
                        name_row = name_row + 1
                    End If
                    Next i
                '____________________________________
            'PartII to fill in yearly price change percentage 
            '____________________________________
                
                'Keep track of the output row for yearly change and percentage change
                price_row = 2
                
                'enter the loop again by using j
                For j = 2 To LastRow
                    'Find the yearly open price for each ticker
                    If ws.Cells(j - 1, 1).Value <> ws.Cells(j, 1).Value Then
                        open_price = ws.Cells(j, 3).Value
                    'Find the yearly close price for each ticker
                    ElseIf ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1).Value Then
                        close_price = ws.Cells(j, 6).Value
                        'Define yearly_change
                        yearly_change = close_price - open_price
                    
                        'Put yearly_change into a cell under column J
                        ws.Range("J" & price_row).Value = yearly_change
                        
                        '________________
                        'conditional formating
                        '_________________
                    
                        'fill in green background if the value is positive
                        If ws.Range("J" & price_row).Value > 0 Then
                            ws.Range("J" & price_row).Interior.Color = vbGreen
                            
                        'fill in red background if the value is negative
                        ElseIf ws.Range("J" & price_row).Value < 0 Then
                            ws.Range("J" & price_row).Interior.Color = vbRed
                        
                        End If
                    'Exclude ticker with value of 0 (one ticker has a value of 0)
                        If open_price = 0 Then
                        change_percentage = 0
             
                        'Define yearly_percentage change
                        ElseIf open_price <> 0 Then
                        change_percentage = (close_price - open_price) / open_price
                        End If    
                    
                    'Put change price into a cell under column K and change the style as percentage
                    ws.Range("K" & price_row).Value = change_percentage
                    ws.Range("K" & price_row).NumberFormat = "0.00%"
                    
                    'add one row to the row counter
                    price_row = price_row + 1
                    
                    End If
                    
                Next j
                
                
            '____________________________________
            'PartIII to calculate the greatest value
            '____________________________________
            
            'Dim the LastRow for the calculated area
            Table_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            'compare value by row and keep the bigger and smaller one
            the_bigger = ws.Cells(2, 11).Value
            the_bigger_ticker = ws.Cells(2, 9).Value
            the_smaller = ws.Cells(2, 11).Value
            the_smaller_ticker = ws.Cells(2, 9).Value
            the_bigger_volume = ws.Cells(2, 12).Value
            the_bigger_volume_ticker = ws.Cells(2, 9).Value
            
            'Enter the loop
            For m = 3 To Table_LastRow
                If ws.Cells(m, 11).Value > the_bigger Then
                    the_bigger = ws.Cells(m, 11).Value
                    the_bigger_ticker = ws.Cells(m, 9).Value
                End If

                If ws.Cells(m, 11).Value < the_smaller Then
                    the_smaller = ws.Cells(m, 11).Value
                    the_smaller_ticker = ws.Cells(m, 9).Value
                End if

                If ws.Cells(m, 12).Value > the_bigger_volume Then
                    the_bigger_volume = ws.Cells(m, 12).Value
                    the_bigger_volume_ticker = ws.Cells(m, 9).Value
                End If
                
            Next m
            ws.Range("P1:Q1").Value = Array("Ticker", "Value")
            ws.Range("O2:Q2").Value = Array("Greatest % increase", the_bigger_ticker, the_bigger)
            ws.Range("O3:Q3").Value = Array("Greatest % decrease", the_smaller_ticker, the_smaller)
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("O4:Q4").Value = Array("Greatest Total Volume", the_bigger_volume_ticker, the_bigger_volume)
     Next ws

    MsgBox ("Homework Complete")       
            
End Sub



